"""FormWriter — preflight + enter edit mode + batch write + save + verify.

Handles one PBF template workbook at a time. Config-driven via YAML.
Full live closed-loop verified 2026-11-11 on YCCBW028_其他成本编制表_CBG:
  enter_edit_mode → batch_write → Plan_Save → Refresh_Sheet → verified persisted in BPC PRD.
"""
from __future__ import annotations

import logging
import math
import re
import time
from dataclasses import dataclass, field
from typing import Optional

from .cancel import raise_if_cancelled
from .excel import ExcelEngine, ExcelError

log = logging.getLogger(__name__)


def _is_empty(v) -> bool:
    """值是否为空（None / 空串 / NaN）——空值不应写入单元格（否则会把 NaN 写进 Excel）。"""
    if v is None:
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    try:
        return math.isnan(v)
    except (TypeError, ValueError):
        return False


def _same_value(a, b) -> bool:
    """判断两个单元格值是否等价（数值按容差比较，文本按去空白比较）。

    空值与 0 视为等价：BPC 里空单元格语义上就是 0，写 0 覆盖空白不应该被当成
    "有变化"，否则批量清 0 时 AO 提示"没有要保存的数据"会被误判为失败。
    """
    if _is_empty(a) and _is_empty(b):
        return True

    def as_num(v):
        if _is_empty(v):
            return 0.0
        try:
            return float(v)
        except (TypeError, ValueError):
            return None

    fa, fb = as_num(a), as_num(b)
    if fa is not None and fb is not None:
        return abs(fa - fb) <= max(1e-6, abs(fb) * 1e-9)
    if _is_empty(a) != _is_empty(b):
        return False
    return str(a).strip() == str(b).strip()


@dataclass
class PreflightResult:
    ok: bool
    checks: list[tuple[str, bool, str]] = field(default_factory=list)
    matched_rows: list[dict] = field(default_factory=list)
    skipped: list[str] = field(default_factory=list)
    # 表单空白时需要在数据区新增行：先写维度列，再写月度数据。
    key_updates: dict[str, object] = field(default_factory=dict)
    new_rows: list[dict] = field(default_factory=list)


class FormWriter:
    def __init__(self, engine: ExcelEngine, form_config: dict) -> None:
        self.engine = engine
        self.cfg = form_config
        # 空白表单新增行的状态（每次 preflight 重置）
        self._key_updates: dict[str, object] = {}
        self._new_rows: list[dict] = []
        self._blank_rows: Optional[list[int]] = None
        self._last_data_row: Optional[int] = None
        # 不同匹配列组合把同一个键指向不同行（键冲突，必须中止）
        self._key_conflicts: list[tuple[str, int, int, str]] = []
        # measure_blocks：按表头动态识别出来的 {度量: {期间: 列}}（每次 preflight 重置）
        self._measure_map: Optional[dict[str, dict[str, str]]] = None

    # ---------- workbook / sheet ----------

    def _resolve(self, wait_open: bool = True):
        pattern = self.cfg["workbook_pattern"]
        wb = self.engine.get_workbook(pattern, wait_sec=30) if wait_open else self.engine.find_workbook(pattern)
        sheet = self.engine.open_sheet(wb, self.cfg["sheet"])
        return wb, sheet

    # ---------- edit-mode switching ----------

    @staticmethod
    def _cell_is_editable(sheet, cell: str) -> bool:
        """按填充色判断单元格是否可录入。

        AO 表单里蓝色填充 = 只读，白色/浅米色 = 可编辑。
        Interior.Color 返回 BGR 整数。
        """
        try:
            c = int(sheet.Range(cell).Interior.Color)
        except Exception:
            return True
        r = c & 0xFF
        g = (c >> 8) & 0xFF
        b = (c >> 16) & 0xFF
        # 蓝色系且偏深 → 只读
        return not (b > r + 20 and b > 120)

    _MODE_LABEL = "当前模式"
    _MODE_QUERY_WORDS = ("查询", "只读")
    _MODE_EDIT_WORDS = ("编制", "编辑")

    def _read_mode(self, sheet) -> tuple[str, str]:
        """读取页面顶部的「当前模式」显示，返回 (模式文本, 来源单元格)。

        优先用配置的 mode_display_cell（如价格表的 AM51）；否则扫描表头上方
        55 行的文本，找「当前模式:」并取其右侧单元格。
        """
        cell = self.cfg.get("mode_display_cell")
        if cell:
            try:
                return str(self.engine.read(sheet, cell) or "").strip(), cell
            except ExcelError:
                return "", cell

        rows = int(self.cfg.get("mode_scan_rows")
                   or max(int(self.cfg.get("header_row") or 0) + 5, 55))
        try:
            block = sheet.Range(f"A1:AZ{rows}").Value2
        except Exception:
            return "", ""
        if not isinstance(block, tuple):
            block = ((block,),)
        for i, row_vals in enumerate(block, start=1):
            if not isinstance(row_vals, tuple):
                row_vals = (row_vals,)
            for j, val in enumerate(row_vals, start=1):
                if not (isinstance(val, str) and self._MODE_LABEL in val):
                    continue
                after = val.split(self._MODE_LABEL, 1)[1].lstrip(":： ")
                if after.strip():
                    return after.strip(), f"{self._col_letter(j)}{i}"
                if j < len(row_vals):
                    nxt = row_vals[j]
                    if isinstance(nxt, str) and nxt.strip():
                        return nxt.strip(), f"{self._col_letter(j + 1)}{i}"
                try:
                    nxt = sheet.Cells(i, j + 1).Value2
                    return str(nxt or "").strip(), f"{self._col_letter(j + 1)}{i}"
                except Exception:
                    return "", ""
        return "", ""

    @classmethod
    def _mode_state(cls, text: str) -> str:
        """把模式显示文本归为 query / edit / 空。"""
        t = str(text or "")
        if any(w in t for w in cls._MODE_QUERY_WORDS):
            return "query"
        if any(w in t for w in cls._MODE_EDIT_WORDS):
            return "edit"
        return ""

    def enter_edit_mode(self, sample_cell: Optional[str] = None,
                        sample_cells: Optional[list[str]] = None) -> None:
        """确保表单处于编制态，必要时调用 Switch_mode。

        判定顺序：
        1. 页面「当前模式」显示为编制/编辑 → 直接跳过（此时再调 Switch_mode
           会反向切回只读）。
        2. 显示为查询/只读，且目标单元格里确实有蓝色（只读）格 → 调切换宏。
           查询态下 AO 只读格是蓝底，写入不会被识别，Plan_Save 会弹
           「没有要保存的数据」（2026-11-23 实测）。
        3. 显示为查询但目标格全是白底（如客户费用表）→ AO 允许直接写入，跳过。
        4. 页面没有模式显示 → 按填充色兜底（全白即跳过）。
        切换后用「当前模式」显示轮询确认，仍未进入编制态则报错，避免拿旧数据
        继续跑后续计算。
        """
        wb, sheet = self._resolve(wait_open=False)

        ms = self.cfg.get("mode_switch") or {}
        macro_name = ms.get("macro") or self.cfg.get("macros", {}).get("switch_mode")
        mode_cell = self.cfg.get("mode_cell")
        label_cell = self.cfg.get("mode_label_cell")
        force = ms.get("force_set", False)

        cells = [c for c in (sample_cells or ([sample_cell] if sample_cell else [])) if c]
        probe = cells[:60]
        blue = sum(0 if self._cell_is_editable(sheet, c) else 1 for c in probe)
        mode_text, mode_src = self._read_mode(sheet)
        state = self._mode_state(mode_text)

        if state == "edit":
            log.info(f"当前模式={mode_text}（{mode_src}），已是编制态，跳过模式切换")
            return

        if state == "query":
            if not probe or blue == 0:
                log.info(f"当前模式={mode_text}（{mode_src}），目标格均为白底可录入，"
                         f"按 AO 行为直接写入，跳过模式切换")
                return
            if not macro_name:
                raise ExcelError(
                    f"表单处于查询模式（{mode_src}={mode_text}），{blue}/{len(probe)} 个"
                    f"目标格为蓝色只读，且配置未提供切换宏")
            log.info(f"当前模式=查询，{blue}/{len(probe)} 个目标格为蓝色只读，"
                     f"运行 {macro_name} 切换到编制态")
            try:
                self.engine.run_macro(macro_name, wb)
            except ExcelError as exc:
                raise ExcelError(f"{macro_name} 执行失败，无法进入编制态: {exc}")
            interval = float(self.cfg.get("mode_poll_interval_sec", 1.0))
            attempts = int(self.cfg.get("mode_poll_attempts", 15))
            for _ in range(attempts):
                time.sleep(interval)
                mode_text, mode_src = self._read_mode(sheet)
                if self._mode_state(mode_text) == "edit":
                    log.info(f"模式已切换：{mode_src}={mode_text}")
                    return
            raise ExcelError(
                f"{macro_name} 后模式仍为「{mode_text}」（{mode_src}），未能进入编制态；"
                f"此时写入不会被保存")

        # 页面无模式显示 → 按填充色兜底
        if probe and blue == 0 and not mode_cell:
            log.info(f"目标单元格均非蓝色且无模式显示，跳过模式切换")
            return

        # 既无切换宏、也无模式单元格 → 本表无需切换（如纯查询表）
        if not macro_name and not mode_cell:
            log.info("本表无模式切换，跳过 enter_edit_mode")
            return

        if macro_name:
            log.info(f"切换模式: 运行 {macro_name} on {wb.Name}")
            try:
                self.engine.run_macro(macro_name, wb)
            except ExcelError as exc:
                log.warning(f"{macro_name} 执行失败: {exc} — 仍尝试强制写模式格")

        if force and mode_cell:
            # 同时写模式格与标签格，AO 读的是模式格
            try:
                sheet.Range(mode_cell).Value2 = "编制"
            except Exception as exc:
                log.warning(f"force-set {mode_cell}: {exc}")
            if label_cell:
                try:
                    sheet.Range(label_cell).Value2 = "编制"
                except Exception:
                    pass

        time.sleep(3.0)
        log.info("模式切换完成，进入可编辑状态")

    # ---------- preflight ----------

    def _key_match_patterns(self) -> list[list[str]]:
        """组合键的匹配列组合。

        第一组是 `key_columns`；`key_match_columns` 可再给若干组同长度的列组合，
        让同一个数据键能匹配到不同显示列。例如价格表的数据源「产品」可能是
        产品成员（AD）或客户型号（AF），两种都要能命中已有行。
        """
        key_cols = self.cfg.get("key_columns") or []
        if not key_cols:
            return []
        primary = [k["column"] for k in key_cols]
        patterns = [primary]
        for pat in self.cfg.get("key_match_columns") or []:
            pat = [str(c).strip().upper() for c in pat]
            if len(pat) != len(primary):
                raise ExcelError(
                    f"key_match_columns 列数({len(pat)})与 key_columns({len(primary)})不一致：{pat}")
            if pat not in patterns:
                patterns.append(pat)
        return patterns

    def _build_row_index(self, sheet) -> dict[str, int]:
        """扫描数据区，返回 {组合键: 行号}。simple 用单科目列，keyed 用组合键。"""
        cfg = self.cfg
        input_rows = cfg.get("input_rows", [1, 100])
        start, end = input_rows[0], input_rows[-1]
        skip_rows = set(cfg.get("skip_rows", []))
        index: dict[str, int] = {}
        self._key_conflicts = []

        key_cols = cfg.get("key_columns")
        if key_cols:
            patterns = self._key_match_patterns()
            for r in range(start, end + 1):
                if r in skip_rows:
                    continue
                # 主组合键优先；别名字段只补空缺，避免覆盖主键命中的行。
                for i, cols in enumerate(patterns):
                    vals = [self.engine.read(sheet, f"{c}{r}") for c in cols]
                    key = "|".join("" if v is None else str(v).strip() for v in vals)
                    if not key.strip("|"):
                        continue
                    if key in index:
                        if index[key] != r:
                            self._key_conflicts.append(
                                (key, index[key], r, "+".join(cols)))
                        continue
                    index[key] = r
            return index

        item_cols = [cfg.get("item_column", "P")] + cfg.get("item_alternate_columns", [])
        for r in range(start, end + 1):
            if r in skip_rows:
                continue
            for col in item_cols:
                v = self.engine.read(sheet, f"{col}{r}")
                if v:
                    key = str(v).strip()
                    if key and key not in index:
                        index[key] = r
        return index

    def _data_key(self, row: dict) -> str:
        key_cols = self.cfg.get("key_columns")
        if key_cols:
            return "|".join(self._dim(row, k["name"]) for k in key_cols)
        if "item" in row:
            return self._dim(row, "item")
        # simple 表：数据首列即科目名
        return self._dim(row, next(iter(row))) if row else ""

    @staticmethod
    def _dim(row: dict, name: str) -> str:
        """维度值转字符串；NaN/None/空串统一为空（不能把 'nan' 当成员）。"""
        v = row.get(name)
        return "" if _is_empty(v) else str(v).strip()

    def _key_columns(self) -> list[str]:
        key_cols = self.cfg.get("key_columns")
        if key_cols:
            return [k["column"] for k in key_cols]
        return [self.cfg.get("item_column", "P")] + list(self.cfg.get("item_alternate_columns", []))

    def _blank_check_columns(self) -> list[str]:
        """判断一行是否完全空白时看的列。

        组合键表必须所有维度列都为空才算空行；单科目表只看科目列——
        AA/AB 可能放分组标题，不能因此把可录入行排除掉。
        """
        if self.cfg.get("key_columns"):
            return self._key_columns()
        return [self.cfg.get("item_column", "P")]

    def _format_sig(self, sheet, row: int) -> Optional[tuple[str, ...]]:
        """数据行的格式指纹（维度列 + 第一个度量列）。

        AO 查询结果行有模板定义的数值格式；查询区之后残留的模板行通常是
        文本格式（`@`）。拿不到格式（例如测试桩/无 COM）时返回 None，
        表示无法判断，沿用旧行为。
        """
        cells = [f"{c}{row}" for c in self._key_columns()]
        measures = self._measure_map or self.cfg.get("measures") or {}
        for cols in measures.values():
            if cols:
                first = next(iter(cols.values()))
                cells.append(f"{first}{row}")
                break
        if len(cells) == len(self._key_columns()):
            month_cols = self.cfg.get("month_columns") or {}
            if month_cols:
                cells.append(f"{next(iter(month_cols.values()))}{row}")
        try:
            return tuple(str(sheet.Range(c).NumberFormat) for c in cells)
        except Exception:
            return None

    def _blank_row_pool(self, sheet) -> list[int]:
        if self._blank_rows is not None:
            return self._blank_rows
        append = bool(self.cfg.get("append_rows"))
        if self.cfg.get("allow_new_rows", True) is False and not append:
            self._blank_rows = []
            log.info("本表配置禁止自动新增行（allow_new_rows: false）")
            return []
        start, end = (self.cfg.get("input_rows") or [1, 100])[:2]
        skip = set(self.cfg.get("skip_rows", []))
        cols = self._blank_check_columns()
        rows: list[int] = []
        last_data_row: Optional[int] = None
        dropped = 0
        for r in range(start, end + 1):
            if r in skip:
                continue
            if not all(_is_empty(self.engine.read(sheet, f"{c}{r}")) for c in cols):
                last_data_row = r
                continue
            if append:
                # append_rows：只在表尾追加，不占用查询区中间的空行
                continue
            # 新增行必须在 AO 数据区格式内：以上方最近一条数据行为基准，
            # 格式不一致的是查询区之外的模板残留行。
            if last_data_row is not None:
                ref_sig = self._format_sig(sheet, last_data_row)
                sig = self._format_sig(sheet, r)
                if ref_sig is not None and sig is not None and sig != ref_sig:
                    dropped += 1
                    continue
            rows.append(r)
        self._last_data_row = last_data_row
        if append:
            # 2026-11-24 实测：在 AO 查询结果最后一行下方写入新维度组合，
            # Plan_Save 保存后关闭重开仍在（原以为会被查询丢弃，实际可以）。
            # 空白表（一条数据行都没有，如 TCL商用客户费用）没有锚点行，
            # 从首个可写行开始追加——表头下方就是 AO 数据区首行。
            # 连续空行都可以追加（多条新键各占一行），遇到宏参数行
            # （skip_rows）或已有数据行即停，不能跨过去写。
            row = (last_data_row + 1) if last_data_row is not None else start
            first = row
            while row <= end:
                if row in skip or not all(
                        _is_empty(self.engine.read(sheet, f"{c}{row}")) for c in cols):
                    break
                rows.append(row)
                row += 1
            if rows:
                whence = f"最后数据行 {last_data_row}" if last_data_row is not None else "空白表"
                log.info(f"表尾追加行 {first}-{rows[-1]}（{whence}）")
            else:
                log.warning(f"无可追加行：起点 {first} 已被占用、是参数行或超出可写范围")
        self._blank_rows = rows
        extra = f"，跳过格式不一致的模板残留行 {dropped} 行" if dropped else ""
        log.info(f"表单空白可新增行 {len(rows)} 行（{start}-{end}）{extra}")
        return rows

    def _row_for_key(self, sheet, row_index: dict[str, int], key: str,
                     key_values: Optional[dict[str, str]] = None,
                     extra_values: Optional[dict[str, str]] = None) -> Optional[int]:
        """已有组合键返回原行；否则取一条空白行，登记维度写入（先维度后数据）。"""
        if key in row_index:
            return row_index[key]
        if not key:
            return None
        rows = self._blank_row_pool(sheet)
        if not rows:
            return None
        row = rows.pop(0)
        values = dict(key_values or {})
        for col, value in (extra_values or {}).items():
            values.setdefault(col, value)
        for col, value in values.items():
            if _is_empty(value):
                continue
            self._key_updates[f"{col}{row}"] = value
        row_index[key] = row
        self._new_rows.append({"row": row, "key": key, "key_values": dict(values)})
        log.info(f"新增行 {row}: {key}")
        return row

    def _copied_dims(self, sheet, row: dict) -> tuple[dict[str, str], str]:
        """新行缺少的维度（如产品型号/客户型号）从同产品的既有行复制。

        价格表只匹配 组织/渠道/品牌/产品 四维；AO 保存新组合时 AE 产品型号、
        AF 客户型号、AG 产品型号状态 也必须有值（手工新增行时用户就是这么
        填的）。这些是产品属性，与渠道无关，可从同产品的任意既有行复制。
        未配置时返回 ({}, "")；找不到可复制的行时返回 ("", 原因)。
        """
        spec = self.cfg.get("new_row_copy") or {}
        cols = [str(c).strip().upper() for c in (spec.get("columns") or [])]
        match_col = str(spec.get("match_column") or "").strip().upper()
        if not cols or not match_col:
            return {}, ""
        names = {str(k.get("column", "")).strip().upper(): k.get("name")
                 for k in self.cfg.get("key_columns") or []}
        value = self._dim(row, names.get(match_col) or "")
        if not value:
            return {}, f"缺少可复制维度 {match_col} 的值"
        start, end = (self.cfg.get("input_rows") or [1, 100])[:2]
        skip = set(self.cfg.get("skip_rows", []))
        for r in range(start, end + 1):
            if r in skip:
                continue
            got = str(self.engine.read(sheet, f"{match_col}{r}") or "").strip()
            if got != value:
                continue
            vals = {c: str(self.engine.read(sheet, f"{c}{r}") or "").strip()
                    for c in cols}
            missing = [c for c, v in vals.items() if not v]
            if missing:
                log.debug(f"同产品行 {r} 的 {missing} 为空，继续找其它行")
                continue
            log.info(f"新增行维度复制自第 {r} 行（{match_col}={value}）：{vals}")
            return vals, ""
        return {}, f"无同产品行可复制维度 {'/'.join(cols)}"

    # ---------- 期间覆盖 ----------

    def _data_periods(self, data: list[dict]) -> set[str]:
        """数据里出现的期间。

        measure_blocks 的期间在行里（period_field），其余布局在列名上（如 "2027.03"）。
        """
        if self.cfg.get("layout") == "measure_blocks":
            field = self.cfg.get("period_field", "期间")
            return {str(r.get(field, "")).strip() for r in data if str(r.get(field, "")).strip()}
        periods: set[str] = set()
        for row in data:
            for key in row:
                if isinstance(key, str) and re.fullmatch(r"\d{4}\.\d{2}", key.strip()):
                    periods.add(key.strip())
        return periods

    def _configured_periods(self) -> set[str]:
        """表单配置里声明了列的期间（measure_blocks 优先用表头动态识别的结果）。"""
        layout = self.cfg.get("layout", "simple")
        if layout == "measure_blocks":
            source = self._measure_map or (self.cfg.get("measures") or {})
            out: set[str] = set()
            for cols in source.values():
                out.update((cols or {}).keys())
            return out
        if layout == "matrix":
            return set((self.cfg.get("period_blocks") or {}).keys())
        return set((self.cfg.get("month_columns") or {}).keys())

    def preflight(self, data: list[dict]) -> PreflightResult:
        """校验数据能否落到表单：匹配行、校验版本与模式。"""
        layout = self.cfg.get("layout", "simple")

        wb, sheet = self._resolve(wait_open=True)
        result = PreflightResult(ok=False)
        self._key_updates = {}
        self._new_rows = []
        self._blank_rows = None
        self._last_data_row = None
        self._key_conflicts = []
        self._measure_map = None
        result.checks.append(("workbook", True, f"{wb.Name}"))

        # 模式检查 —— 只报告；真正的切换在写值前按「当前模式 + 目标格颜色」执行，
        # 因为查询态下部分表单（客户费用）仍可直接写入，不能一概切换/判失败。
        mode_cell = self.cfg.get("mode_cell")
        mode_text, mode_src = self._read_mode(sheet)
        state = self._mode_state(mode_text)
        if state:
            note = "编制态" if state == "edit" else "查询态（写入前自动判断是否切换）"
            result.checks.append(("编制模式", True, f"{mode_src}={mode_text}（{note}）"))
        elif mode_cell:
            mode_val = self.engine.read(sheet, mode_cell) or ""
            ok = "编制" in str(mode_val)
            if not ok:
                log.warning(f"Mode={mode_val}, 自动切编制模式")
                try:
                    self.enter_edit_mode()
                    mode_val = self.engine.read(sheet, mode_cell) or ""
                    ok = "编制" in str(mode_val)
                except ExcelError as exc:
                    log.error(f"enter_edit_mode 失败: {exc}")
            result.checks.append(("编制模式", ok, f"{mode_cell}={mode_val}"))
        else:
            result.checks.append(("编制模式", True, "(本表无模式显示)"))

        # 版本检查
        version_cell = self.cfg.get("version_cell")
        expected_version = self.cfg.get("expected_version")
        if version_cell and expected_version:
            actual = self.engine.read(sheet, version_cell) or ""
            ok = expected_version in str(actual)
            result.checks.append(("版本", ok, f"{version_cell}={actual}, 期望 {expected_version}"))
        elif version_cell:
            actual = self.engine.read(sheet, version_cell) or ""
            result.checks.append(("版本", True, f"{version_cell}={actual}"))

        row_index = self._build_row_index(sheet)
        key_cols = self.cfg.get("key_columns")
        aliases = self.cfg.get("item_aliases", {})

        if layout == "matrix":
            matched, skipped = self._cells_matrix(sheet, row_index, data)
        elif layout == "measure_blocks":
            matched, skipped = self._cells_measures(sheet, row_index, data)
        else:
            matched, skipped = self._cells_simple_keyed(sheet, row_index, data, key_cols, aliases)

        result.matched_rows = matched
        result.skipped = skipped
        result.key_updates = dict(self._key_updates)
        result.new_rows = list(self._new_rows)
        detail = f"key={'+'.join(k['name'] for k in key_cols)}" if key_cols else "item_column"
        result.checks.append(("行匹配", True, f"{detail}, 命中 {len(matched)} 格"))
        if self._key_conflicts:
            sample = "；".join(
                f"{k} → 行{a}/{b}（{cols}）" for k, a, b, cols in self._key_conflicts[:3])
            result.checks.append((
                "键冲突", False,
                f"{len(self._key_conflicts)} 个键在不同匹配列组合下指向不同行：{sample}"))
        if result.new_rows:
            result.checks.append((
                "新增行", True,
                f"空白表单新增 {len(result.new_rows)} 行，先写 {len(result.key_updates)} 个维度格再写数据"))

        # 期间覆盖：数据里出现的期间必须都有对应列，否则会被静默丢弃
        data_periods = self._data_periods(data)
        configured = self._configured_periods()
        missing_periods = sorted(p for p in data_periods if p not in configured)
        result.checks.append((
            "期间覆盖", not missing_periods,
            f"{sorted(data_periods)} 均有对应列" if not missing_periods
            else f"配置缺少期间 {missing_periods}，这部分数据不会被写入"))

        # 单元格唯一：两行数据落到同一格会互相覆盖
        cell_owner: dict[str, str] = {}
        dups: list[str] = []
        for r in matched:
            cell = r["cell"]
            if cell in cell_owner:
                dups.append(f"{cell}({cell_owner[cell]}→{r['item']})")
            else:
                cell_owner[cell] = r["item"]
        result.checks.append((
            "单元格唯一", not dups,
            f"{len(cell_owner)} 格无重复" if not dups else f"重复写入 {dups[:3]}"))

        result.ok = (all(pass_ for _, pass_, _ in result.checks)
                     and len(skipped) == 0 and len(matched) > 0)
        return result

    # ---------- 各布局的单元格定位 ----------

    def _cells_simple_keyed(self, sheet, row_index, data, key_cols, aliases) -> tuple[list, list]:
        month_cols = self.cfg.get("month_columns", {})
        item_col = self.cfg.get("item_column", "P")
        matched, skipped = [], []
        for row in data:
            raw_key = self._data_key(row)
            label = raw_key.replace("|", " / ") or str(row.get("item", ""))
            key = raw_key
            target_row = row_index.get(key)

            # simple 表允许通过别名兜底：数据里是"售后服务费-维修费"，表单行是"维修费"
            if target_row is None and not key_cols:
                canonical = None
                for cand, al in aliases.items():
                    if raw_key in al or raw_key == cand:
                        canonical = cand
                        break
                if canonical:
                    key = canonical
                    target_row = row_index.get(canonical)

            if target_row is None:
                if key_cols:
                    key_values = {k["column"]: self._dim(row, k["name"]) for k in key_cols}
                    if any(not v for v in key_values.values()):
                        skipped.append(f"{label} [维度缺失]")
                        continue
                else:
                    if not key:
                        skipped.append(label)
                        continue
                    key_values = {item_col: key}
                target_row = self._row_for_key(sheet, row_index, key, key_values)
                if target_row is None:
                    skipped.append(f"{label} [无空白行]")
                    continue
                label = key.replace("|", " / ")

            for month, col_letter in month_cols.items():
                if month not in row:
                    continue
                if _is_empty(row[month]):
                    continue
                cell = f"{col_letter}{target_row}"
                matched.append({
                    "item": label, "month": month, "cell": cell,
                    "new_value": row[month],
                    "old_value": self.engine.read(sheet, cell),
                })
        return matched, skipped

    @staticmethod
    def _col_letter(n: int) -> str:
        s = ""
        while n > 0:
            n, rem = divmod(n - 1, 26)
            s = chr(65 + rem) + s
        return s

    def _detect_measure_columns(self, sheet) -> Optional[dict[str, dict[str, str]]]:
        """按表头动态识别度量块的 (度量 × 期间 → 列) 映射。

        背景：AO 模板的度量块宽度随版本期间窗口变化（实测 2026.11 版是 7 列，
        11-11 录制时是 6 列）。硬编码列会在块宽变化时整体错月——销量块错 1 列、
        加权CT 错 2 列、销售收入错 3 列。因此以 `measure_name_row`（度量名）+
        `header_row`（期间）两行表头为准，每次打开表单重新识别。
        """
        cfg = self.cfg
        try:
            header_row = int(cfg["header_row"])
        except (KeyError, TypeError, ValueError):
            return None
        name_row = int(cfg.get("measure_name_row") or header_row - 1)
        config_measures = cfg.get("measures") or {}
        if not config_measures:
            return None

        try:
            used = sheet.UsedRange
            first_col = int(getattr(used, "Column", 1) or 1)
            count = int(getattr(getattr(used, "Columns", None), "Count", 0) or 0)
        except Exception:
            first_col, count = 1, 0
        last_col = first_col + count - 1 if count else first_col + 99
        last_col = min(last_col, first_col + 199)

        def norm(v) -> str:
            return re.sub(r"\s+", "", str(v or ""))

        groups: dict[str, dict[str, str]] = {}
        for n in range(first_col, last_col + 1):
            col = self._col_letter(n)
            name = norm(self.engine.read(sheet, f"{col}{name_row}"))
            period = str(self.engine.read(sheet, f"{col}{header_row}") or "").strip()
            if not name or not re.fullmatch(r"\d{4}\.\d{2}", period):
                continue
            groups.setdefault(name, {})[period] = col

        mapping: dict[str, dict[str, str]] = {}
        for measure in config_measures:
            key = norm(measure)
            for header, cols in groups.items():
                if header == key or header.startswith(key) or key.startswith(header):
                    mapping[measure] = cols
                    break
        # 识别失败/不完整绝不静默回退配置列：配置列是按模板实测写死的，
        # 模板块宽度随版本期间窗口变（6 列 ↔ 7 列），识别失败往往意味着
        # 模板和配置对不上，硬写配置列会把数字写进错误的度量/期间。
        if not mapping:
            raise ExcelError(
                f"度量块表头识别失败：{name_row}/{header_row} 行没有找到"
                "「度量名 + YYYY.MM 期间」表头（配置的 measures 一个都对不上）。"
                "可能是表单尚未加载完成，或 header_row/measure_name_row 配置有误；"
                "为避免按旧列号写错列，本次中止。请重新打开表单后重试，"
                "或核对表单配置")
        missing = [m for m in config_measures if m not in mapping]
        if missing:
            raise ExcelError(
                f"度量块表头识别不完整：缺少 {missing} 的表头。模板度量块与配置的 "
                "measures 对不上，为避免静默写错列，本次中止。请核对表单配置")
        sizes = {m: len(cols) for m, cols in mapping.items()}
        if len(set(sizes.values())) > 1:
            raise ExcelError(
                f"度量块表头识别异常：各度量的期间数不一致（{sizes}），"
                "表头可能有合并单元格或空列。为避免写错列，本次中止。请核对表头")
        owners: dict[str, str] = {}
        for m, cols in mapping.items():
            for period, col in cols.items():
                if col in owners and owners[col] != m:
                    raise ExcelError(
                        f"度量块表头识别异常：列 {col} 同时识别给 "
                        f"{owners[col]} 和 {m}，表头可能有重复期间。"
                        "为避免写错列，本次中止。请核对表头")
                owners[col] = m
        log.info("度量块表头识别: " + "；".join(
            f"{m}={sorted(cols)}" for m, cols in mapping.items()))
        return mapping

    def _cells_measures(self, sheet, row_index, data) -> tuple[list, list]:
        """度量块布局：一行 = 一个维度组合 + 一个期间 + 多个度量值。

        期间取数据行的 `period_field`（默认"期间"）列，定位到该度量对应的列。
        列映射按表头动态识别（见 `_detect_measure_columns`）；识别失败/不完整
        会直接报错而不是回退配置列——模板块宽会变，硬写配置列可能写错列。
        """
        if self._measure_map is None:
            self._measure_map = self._detect_measure_columns(sheet)
        measures = self._measure_map or self.cfg.get("measures", {})
        alias_map = self.cfg.get("measure_aliases", {})
        period_field = self.cfg.get("period_field", "期间")
        default_period = self.cfg.get("default_period")
        key_cols = self.cfg.get("key_columns") or []
        matched, skipped = [], []
        for row in data:
            key = self._data_key(row)
            label = key.replace("|", " / ") or str(row.get("item", ""))
            target_row = row_index.get(key)
            if target_row is None:
                key_values = {k["column"]: self._dim(row, k["name"]) for k in key_cols}
                if not key or any(not v for v in key_values.values()):
                    skipped.append(f"{label} [维度缺失]")
                    continue
                copy_values, copy_err = self._copied_dims(sheet, row)
                if copy_err:
                    skipped.append(f"{label} [{copy_err}]")
                    continue
                target_row = self._row_for_key(sheet, row_index, key, key_values,
                                               extra_values=copy_values)
                if target_row is None:
                    skipped.append(f"{label} [无空白行]")
                    continue
                label = key.replace("|", " / ")

            period = str(row.get(period_field, "")).strip() or default_period
            if not period:
                skipped.append(f"{label} [缺期间列 '{period_field}']")
                continue

            for measure, cols in measures.items():
                # 数据列名可能是度量别名（如"销售单价（含税）"）
                data_name = measure
                if measure not in row:
                    data_name = None
                    for cand in alias_map.get(measure, []):
                        if cand in row:
                            data_name = cand
                            break
                    if data_name is None:
                        continue
                if _is_empty(row[data_name]):
                    continue
                col_letter = cols.get(period)
                if not col_letter:
                    continue
                cell = f"{col_letter}{target_row}"
                matched.append({
                    "item": f"{label} [{measure}]", "month": period, "cell": cell,
                    "new_value": row[data_name],
                    "old_value": self.engine.read(sheet, cell),
                })
        return matched, skipped

    def _cells_matrix(self, sheet, row_index, data) -> tuple[list, list]:
        """矩阵布局：科目 × (期间块 × 区域)。列 = 期间起始列 + 区域内偏移。"""
        blocks = self.cfg.get("period_blocks", {})
        region_order = self.cfg.get("region_order", [])
        region_names = self.cfg.get("region_names", {})
        region_field = self.cfg.get("region_field", "区域")

        def col_num(letter: str) -> int:
            n = 0
            for ch in letter.upper():
                n = n * 26 + (ord(ch) - 64)
            return n

        def col_letter(n: int) -> str:
            s = ""
            while n > 0:
                n, rem = divmod(n - 1, 26)
                s = chr(65 + rem) + s
            return s

        def region_index(v: str) -> int | None:
            if v in region_order:
                return region_order.index(v)
            for code, name in region_names.items():
                if name == v:
                    return region_order.index(code) if code in region_order else None
            return None

        item_col = self.cfg.get("item_column", "AA")
        matched, skipped = [], []
        for row in data:
            key = self._data_key(row)
            label = key.replace("|", " / ") or str(row.get("item", ""))
            target_row = row_index.get(key)
            if target_row is None:
                if not key:
                    skipped.append(label)
                    continue
                target_row = self._row_for_key(sheet, row_index, key, {item_col: key})
                if target_row is None:
                    skipped.append(f"{label} [无空白行]")
                    continue

            region = str(row.get(region_field, "")).strip()
            idx = region_index(region)
            if idx is None:
                skipped.append(f"{label} [区域={region}]")
                continue

            for period, start in blocks.items():
                if period not in row:
                    continue
                if _is_empty(row[period]):
                    continue
                cell = f"{col_letter(col_num(start) + idx)}{target_row}"
                matched.append({
                    "item": f"{label} [{region}]", "month": period, "cell": cell,
                    "new_value": row[period],
                    "old_value": self.engine.read(sheet, cell),
                })
        return matched, skipped

    # ---------- 锁定检测 ----------

    @staticmethod
    def _looks_lock(text: str) -> bool:
        text = str(text or "")
        return bool(
            re.search(r"被.{0,40}锁定", text)
            or re.search(r"locked\s+by", text, re.I)
            or ("表单" in text and "锁定" in text)
            or ("表单" in text and "正在编辑" in text)
        )

    def _table_looks_readonly(self, sheet) -> bool:
        """抽查数据区填充色：整片蓝色（只读）才认为可能是锁定/查询态。"""
        start, end = (self.cfg.get("input_rows") or [1, 100])[:2]
        key_cols = self._key_columns()
        layout = self.cfg.get("layout", "simple")
        if layout == "measure_blocks":
            source = self._measure_map or (self.cfg.get("measures") or {})
            data_cols = [next(iter(cols.values())) for cols in source.values() if cols]
        elif layout == "matrix":
            data_cols = list((self.cfg.get("period_blocks") or {}).values())
        else:
            data_cols = list((self.cfg.get("month_columns") or {}).values())
        # 查询态下各数据列底色不一（价格表：单价白底、销量/CT 蓝底），只抽
        # 前两列会把可编辑误判成锁定；抽全数据列（封顶 8 个）才能区分，
        # 真锁定时整片蓝，多抽几个不影响判定。
        cols = list(dict.fromkeys((key_cols[:2] + data_cols)))[:8]
        if not cols:
            return False
        samples = [f"{c}{r}" for r in range(start, min(start + 4, end + 1)) for c in cols]
        if not samples:
            return False
        return all(not self._cell_is_editable(sheet, addr) for addr in samples)

    def detect_lock(self) -> Optional[str]:
        """检测表单是否被其他用户锁定。

        线索有三处：Excel 状态栏左下角提示、表单上的提示形状/文本框、
        以及数据区全部变蓝（只读）。命中任一处就返回提示文本，由编排层中止流程。
        """
        try:
            _wb, sheet = self._resolve(wait_open=False)
        except ExcelError:
            return None

        app = getattr(self.engine, "_excel", None)
        try:
            status = getattr(app, "StatusBar", None) if app is not None else None
        except Exception:
            status = None
        if isinstance(status, str) and self._looks_lock(status):
            return status.strip()

        try:
            for shape in sheet.Shapes:
                text = ""
                try:
                    if shape.TextFrame.HasText:
                        text = str(shape.TextFrame.Characters().Text)
                except Exception:
                    pass
                if text and self._looks_lock(text):
                    return text.strip()
        except Exception:
            pass

        if not self._table_looks_readonly(sheet):
            return None

        # 数据区整片蓝色时再扫单元格文本（提示可能写在页面左下角某格）
        try:
            vals = sheet.UsedRange.Value2
            if isinstance(vals, tuple):
                for row_vals in vals:
                    if not isinstance(row_vals, tuple):
                        row_vals = (row_vals,)
                    for v in row_vals:
                        if isinstance(v, str) and self._looks_lock(v):
                            return v.strip()
            elif isinstance(vals, str) and self._looks_lock(vals):
                return vals.strip()
        except Exception:
            pass
        return "表单数据区全部为只读（蓝色），疑似被其他用户锁定（未能读到用户名）"

    # ---------- write cycle ----------

    def write_and_save(self, preflight: PreflightResult, verify_after: bool = True) -> dict:
        """执行 enter_edit_mode → write → Plan_Save → (可选)验证。"""
        if not preflight.ok:
            raise ExcelError(f"Preflight failed (or no values to write)")

        raise_if_cancelled()               # 安全点：还没动任何单元格
        wb, sheet = self._resolve(wait_open=False)
        mac = self.cfg.get("macros", {})
        result = {"written": 0, "saved": False, "verified": False, "failures": []}

        # 先算好要写的单元格；拿多个目标格做颜色采样——价格表查询态下
        # 单价列是白底、销量/CT 列才是蓝底，只看第一个格会误判为可编辑。
        updates = {row["cell"]: row["new_value"] for row in preflight.matched_rows}
        sample_cells = list(updates)
        sample_cell = sample_cells[0] if sample_cells else None

        # Ensure edit mode — 按「当前模式」显示 + 采样格颜色决定是否切换
        try:
            self.enter_edit_mode(sample_cell=sample_cell, sample_cells=sample_cells)
        except ExcelError as exc:
            result["failures"].append(f"edit_mode: {exc}")
            return result

        # 1. write —— 空白表单必须先写维度（组织/渠道/品牌/产品…），
        # AO 会按这些维度校验后面写入的月度数据。顺序反了可能被丢弃或串行。
        key_updates = dict(getattr(preflight, "key_updates", {}) or {})
        if key_updates:
            raise_if_cancelled()           # 安全点：写入前最后检查
            log.info(f"先写维度 {len(key_updates)} 格（新增行）")
            try:
                self.engine.write_batch(sheet, key_updates)
                result["keys_written"] = len(key_updates)
            except ExcelError as exc:
                result["failures"].append(f"write_dimensions: {exc}")
                return result
            time.sleep(2.0)                      # 让 AO 完成维度校验再写数值

        raise_if_cancelled()               # 安全点：数值写入前最后检查
        log.info(f"Writing {len(updates)} cells into {wb.Name}")
        try:
            self.engine.write_batch(sheet, updates)
            result["written"] = len(updates)
        except ExcelError as exc:
            result["failures"].append(f"write_batch: {exc}")
            return result

        # 2. save
        # 取消发生在这里 → 不保存，BPC 里仍是旧数据；上层会关闭工作簿丢弃本地改动
        raise_if_cancelled()
        result["cells"] = {row["cell"]: row["new_value"] for row in preflight.matched_rows}
        save_macro = mac.get("save")
        try:
            self.engine.save(wb, macro_hint=save_macro)
            result["saved"] = True
            unchanged = any("没有要保存" in d for d in getattr(self.engine, "last_dialogs", []))
            # 有新值却被告知"没有要保存的数据"：说明本次写入没被接受（常见于未进入
            # 编制态、单元格只读）。不能当成"数据未变化"放过，否则后续计算会基于
            # 旧数据一路跑下去，最终输出"看似成功"的错误结果。
            changed = [r for r in preflight.matched_rows
                       if not _same_value(r.get("old_value"), r.get("new_value"))]
            if unchanged and changed:
                sample = changed[0]
                result["saved"] = False
                result["failures"].append(
                    f"AO 提示「没有要保存的数据」，但本次有 {len(changed)} 格新值与旧值不同"
                    f"（示例 {sample['cell']}: {sample.get('old_value')!r} → "
                    f"{sample.get('new_value')!r}），疑似未进入编制态或写入未生效")
                log.error(result["failures"][-1])
                return result
            if unchanged:
                # 第二次跑同一期间时，写入值与已保存值一致，Plan_Save 会弹这个提示
                result["unchanged"] = True
                log.warning("AO 提示「没有要保存的数据」：本次写入与已保存值一致，跳过保存")
            else:
                log.info("Plan_Save executed")
        except ExcelError as exc:
            result["failures"].append(f"save: {exc}")
            return result

        # 3. 验证
        # 注意：AO 的 Refresh_Sheet 会关闭当前工作簿，导致后续读取拿到失效对象。
        # 因此默认不在原地刷新，落库验证交给上层「关闭不保存 → 重开 → 读值」。
        if verify_after and self.cfg.get("verify_mode") == "refresh":
            time.sleep(2.0)
            refresh_macro = mac.get("refresh")
            try:
                self.engine.refresh(wb, macro_hint=refresh_macro)
                time.sleep(3.0)
            except ExcelError as exc:
                result["failures"].append(f"refresh: {exc}")

            all_match = True
            for row in preflight.matched_rows:
                actual = self.engine.read(sheet, row["cell"])
                try:
                    diff = abs(float(actual or 0) - float(row["new_value"]))
                    ok = diff < 1e-2
                except (TypeError, ValueError):
                    ok = str(actual) == str(row["new_value"])
                if not ok:
                    all_match = False
                    result["failures"].append(
                        f"{row['item']} {row['month']} {row['cell']}: 写入 {row['new_value']}, 刷新后={actual}")
            result["verified"] = all_match

        return result
