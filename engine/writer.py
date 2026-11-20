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


class FormWriter:
    def __init__(self, engine: ExcelEngine, form_config: dict) -> None:
        self.engine = engine
        self.cfg = form_config

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

    def enter_edit_mode(self, sample_cell: Optional[str] = None) -> None:
        """切换到编制（可编辑）状态。

        优先看目标单元格填充色：
        - 已经是白色/浅色 → 本表可直接录入，不要再调 Switch_mode（会对 AO 状态造成干扰）
        - 蓝色只读 → 调用 Switch_mode 宏（等同点击"切换模式"按钮）
        """
        wb, sheet = self._resolve(wait_open=False)

        ms = self.cfg.get("mode_switch") or {}
        macro_name = ms.get("macro") or self.cfg.get("macros", {}).get("switch_mode")
        mode_cell = self.cfg.get("mode_cell")
        label_cell = self.cfg.get("mode_label_cell")
        force = ms.get("force_set", False)

        # 填充色为白不代表 AO 处于编制模式（实测：其他成本表在"查询"模式下输入区也是白底，
        # 此时 Plan_Save 会弹"没有要保存的资源"并阻塞）。因此只要有 mode_cell 就必须切。
        if sample_cell and self._cell_is_editable(sheet, sample_cell) and not mode_cell:
            log.info(f"{sample_cell} 已是可编辑(非蓝色)，且本表无模式格，跳过模式切换")
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

    def _build_row_index(self, sheet) -> dict[str, int]:
        """扫描数据区，返回 {组合键: 行号}。simple 用单科目列，keyed 用组合键。"""
        cfg = self.cfg
        input_rows = cfg.get("input_rows", [1, 100])
        start, end = input_rows[0], input_rows[-1]
        skip_rows = set(cfg.get("skip_rows", []))
        index: dict[str, int] = {}

        key_cols = cfg.get("key_columns")
        if key_cols:
            cols = [k["column"] for k in key_cols]
            for r in range(start, end + 1):
                if r in skip_rows:
                    continue
                vals = [self.engine.read(sheet, f"{c}{r}") for c in cols]
                key = "|".join("" if v is None else str(v).strip() for v in vals)
                if key.strip("|") and key not in index:
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
            return "|".join(str(row.get(k["name"], "")).strip() for k in key_cols)
        if "item" in row:
            return str(row["item"]).strip()
        # simple 表：数据首列即科目名
        return str(next(iter(row.values()))).strip() if row else ""

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
        """表单配置里声明了列的期间。"""
        layout = self.cfg.get("layout", "simple")
        if layout == "measure_blocks":
            out: set[str] = set()
            for cols in (self.cfg.get("measures") or {}).values():
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
        result.checks.append(("workbook", True, f"{wb.Name}"))

        # 模式检查
        mode_cell = self.cfg.get("mode_cell")
        if mode_cell:
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
            result.checks.append(("编制模式", True, "(本表无模式切换)"))

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
        detail = f"key={'+'.join(k['name'] for k in key_cols)}" if key_cols else "item_column"
        result.checks.append(("行匹配", True, f"{detail}, 命中 {len(matched)} 格"))

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
        matched, skipped = [], []
        for row in data:
            key = self._data_key(row)
            label = key.replace("|", " / ") or str(row.get("item", ""))
            if key and key not in row_index:
                # simple 表允许通过别名兜底
                resolved = None
                if not key_cols:
                    for canonical, al in aliases.items():
                        if (key in al or key == canonical) and canonical in row_index:
                            resolved = canonical
                            break
                if resolved is None:
                    skipped.append(label)
                    continue
                key = resolved
                label = resolved

            target_row = row_index[key]
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

    def _cells_measures(self, sheet, row_index, data) -> tuple[list, list]:
        """度量块布局：一行 = 一个维度组合 + 一个期间 + 多个度量值。

        期间取数据行的 `period_field`（默认"期间"）列，定位到该度量对应的列。
        """
        measures = self.cfg.get("measures", {})
        alias_map = self.cfg.get("measure_aliases", {})
        period_field = self.cfg.get("period_field", "期间")
        default_period = self.cfg.get("default_period")
        matched, skipped = [], []
        for row in data:
            key = self._data_key(row)
            label = key.replace("|", " / ") or str(row.get("item", ""))
            if key not in row_index:
                skipped.append(label)
                continue
            target_row = row_index[key]

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

        matched, skipped = [], []
        for row in data:
            key = self._data_key(row)
            label = key.replace("|", " / ") or str(row.get("item", ""))
            if key not in row_index:
                skipped.append(label)
                continue
            target_row = row_index[key]

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

    # ---------- AO 状态校验 ----------

    def _wait_pbf_ready(self, timeout: float = 20.0) -> None:
        """等待 PBF 表单初始化结束（Application.EnableEvents 恢复为 True）。

        PBF 模板的回调都是"先关事件、做完再开"（导航表 Callback_AfterRedisplay、
        表单 Callback_AfterRedisplay / Refresh_Sheet / Switch_mode / Copy_Range）。
        外部写入如果落在这个窗口里，Excel 不触发 SheetChange，AO 收不到变更，
        Plan_Save 就会立刻弹「没有要保存的数据」（2026-11-20 批量清除复现）。
        """
        deadline = time.time() + timeout
        while time.time() < deadline:
            if self.engine.events_enabled() is not False:
                return
            time.sleep(0.5)
        log.warning(f"等待 Excel 事件恢复超时（{timeout:.0f}s）")
        if self.engine.ensure_events():
            log.warning("已强制开启 Application.EnableEvents（PBF 宏未恢复事件）")

    def _plan_data_dirty(self) -> Optional[bool]:
        """AO 是否记录了未保存的计划数据变更；读取失败返回 None。"""
        try:
            return bool(self.engine.run_macro("SAPGetProperty", None, "HasChangedPlanData"))
        except ExcelError as exc:
            log.debug(f"读取 HasChangedPlanData 失败: {exc}")
            return None

    # ---------- write cycle ----------

    def write_and_save(self, preflight: PreflightResult, verify_after: bool = True) -> dict:
        """执行 enter_edit_mode → write → Plan_Save → (可选)验证。"""
        if not preflight.ok:
            raise ExcelError(f"Preflight failed (or no values to write)")

        wb, sheet = self._resolve(wait_open=False)
        mac = self.cfg.get("macros", {})
        result = {"written": 0, "saved": False, "verified": False, "failures": []}

        # 先算好要写的单元格，用它作为"是否已可编辑"的采样点
        updates = {row["cell"]: row["new_value"] for row in preflight.matched_rows}
        sample_cell = next(iter(updates), None)

        # Ensure edit mode — 已是可编辑状态则不切模式
        try:
            self.enter_edit_mode(sample_cell=sample_cell)
        except ExcelError as exc:
            result["failures"].append(f"edit_mode: {exc}")
            return result

        changed = [r for r in preflight.matched_rows
                   if not _same_value(r.get("old_value"), r.get("new_value"))]

        # 1. write
        self._wait_pbf_ready()
        if self.engine.ensure_events():
            log.warning("写入前 Excel 事件处于关闭状态（PBF 宏未恢复），已强制开启")
        before_dirty = self._plan_data_dirty()
        log.info(f"Writing {len(updates)} cells into {wb.Name}")
        try:
            self.engine.write_batch(sheet, updates)
            result["written"] = len(updates)
        except ExcelError as exc:
            result["failures"].append(f"write_batch: {exc}")
            return result

        # 写入后立刻确认 AO 收到了变更：事件被关掉或表单还在初始化时，单元格虽然
        # 写进去了，但 AO 的计划缓冲区没有更新，后面 Plan_Save 必然弹"没有要保存"。
        dirty = self._plan_data_dirty()
        log.info(f"写入后 HasChangedPlanData={dirty}（写入前={before_dirty}）")
        if changed and dirty is False:
            log.warning("写入后 HasChangedPlanData=False：AO 未记录本次变更，重写一次")
            time.sleep(1.0)
            if self.engine.ensure_events():
                log.warning("重写前 Excel 事件处于关闭状态，已强制开启")
            try:
                self.engine.write_batch(sheet, updates)
            except ExcelError as exc:
                result["failures"].append(f"write_batch(重试): {exc}")
                return result
            dirty = self._plan_data_dirty()
        if changed and dirty is False:
            back = sum(1 for cell, want in updates.items()
                       if not _same_value(self.engine.read(sheet, cell), want))
            result["saved"] = False
            result["failures"].append(
                "写入后 AO 未记录任何变更（HasChangedPlanData=False），重写一次仍无效；"
                f"回读 {len(updates)} 格中有 {back} 格与目标值不同。未调用保存")
            log.error(result["failures"][-1])
            return result
        if before_dirty:
            log.info("写入前 AO 已有未保存变更，HasChangedPlanData 保持 True")

        # 2. save
        result["cells"] = {row["cell"]: row["new_value"] for row in preflight.matched_rows}
        save_macro = mac.get("save")
        try:
            self.engine.save(wb, macro_hint=save_macro)
            result["saved"] = True
            unchanged = any("没有要保存" in d for d in getattr(self.engine, "last_dialogs", []))
            # 有新值却被告知"没有要保存的数据"：说明本次写入没被接受（常见于未进入
            # 编制态、单元格只读）。不能当成"数据未变化"放过，否则后续计算会基于
            # 旧数据一路跑下去，最终输出"看似成功"的错误结果。
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

        # 3. 验证（默认关闭）
        # 流程里落库验证统一由编排层做：保存后原地 Refresh_Sheet 读回（默认），
        # 或配置成「关闭不保存 → 重开 → 读值」。这里保留给 verify_mode=refresh 的独立调用。
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
