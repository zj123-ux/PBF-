"""离线配置/数据预校验（不连 Excel、不改任何状态）。

run-flow 在启动 Excel 之前调用这里，把「配置漏字段」「数据期间在表单里没有列」
「数据文件缺失/为空」「导航表路径不存在」等问题提前暴露，避免跑了一半才失败。
"""
from __future__ import annotations

import csv
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

PERIOD_RE = re.compile(r"^\d{4}\.\d{2}$")
COL_RE = re.compile(r"^[A-Z]{1,3}$")
LAYOUTS = {"simple", "keyed", "matrix", "measure_blocks"}


@dataclass(frozen=True)
class Issue:
    level: str          # "error" | "warning" | "info"
    message: str

    def __str__(self) -> str:
        return f"[{self.level}] {self.message}"


def _err(msg: str) -> Issue:
    return Issue("error", msg)


def _warn(msg: str) -> Issue:
    return Issue("warning", msg)


def _info(msg: str) -> Issue:
    return Issue("info", msg)


# ---------- 期间 ----------

def configured_periods(cfg: dict) -> set[str]:
    """表单配置里声明了列的期间。"""
    layout = cfg.get("layout", "simple")
    if layout == "measure_blocks":
        out: set[str] = set()
        for cols in (cfg.get("measures") or {}).values():
            out.update((cols or {}).keys())
        return {str(p).strip() for p in out}
    if layout == "matrix":
        return {str(p).strip() for p in (cfg.get("period_blocks") or {})}
    return {str(p).strip() for p in (cfg.get("month_columns") or {})}


def data_periods(rows: Iterable[dict], cfg: dict) -> set[str]:
    """数据里实际出现的期间（宽表看列名，measure_blocks 看行内期间字段）。"""
    rows = list(rows)
    if cfg.get("layout") == "measure_blocks":
        field = str(cfg.get("period_field", "期间"))
        return {str(r.get(field, "")).strip() for r in rows
                if str(r.get(field, "")).strip()}
    out: set[str] = set()
    for row in rows:
        for key in row:
            if isinstance(key, str) and PERIOD_RE.match(key.strip()):
                out.add(key.strip())
    return out


# ---------- 数据文件 ----------

def read_rows(path: Path) -> list[dict]:
    """读取 CSV（自动尝试 utf-8-sig / gbk）或 Excel，返回 dict 行。"""
    suffix = path.suffix.lower()
    if suffix == ".csv":
        last: Exception | None = None
        for enc in ("utf-8-sig", "gbk"):
            try:
                with open(path, newline="", encoding=enc) as f:
                    return list(csv.DictReader(f))
            except UnicodeDecodeError as exc:
                last = exc
        raise ValueError(f"CSV 编码无法识别（尝试 utf-8-sig / gbk）：{last}")
    if suffix in (".xlsx", ".xls", ".xlsm"):
        import pandas as pd

        return pd.read_excel(path).to_dict(orient="records")
    raise ValueError(f"不支持的数据文件扩展名：{suffix}")


def _required_data_columns(cfg: dict) -> list[tuple[str, list[str]]]:
    """[(错误信息里的列名, 可接受的列名列表)]。

    measure_blocks 的数据列可能使用度量别名（如“销售单价（含税）”），
    只要命中 canonical 或任一别名即可。
    `computed_measures` 里声明的度量（如 销售收入，由 Excel 内 Execute_PS
    计算得到、本就不来自 CSV）不算必填列——与 writer 缺列即跳过的行为一致。
    """
    layout = cfg.get("layout", "simple")
    req: list[tuple[str, list[str]]] = []
    for key in cfg.get("key_columns") or []:
        name = str(key.get("name", "")).strip()
        if name:
            req.append((name, [name]))
    if layout == "measure_blocks":
        alias_map = cfg.get("measure_aliases") or {}
        computed = {str(m) for m in (cfg.get("computed_measures") or [])}
        for measure in (cfg.get("measures") or {}):
            if str(measure) in computed:
                continue
            names = [str(measure)] + [str(a) for a in (alias_map.get(measure) or [])]
            req.append((str(measure), names))
        field = str(cfg.get("period_field", "期间"))
        if all(field != label for label, _ in req):
            req.append((field, [field]))
    return req


def _row_key(row: dict, cfg: dict) -> str:
    key_cols = cfg.get("key_columns")
    if key_cols:
        return "|".join(str(row.get(k.get("name", ""), "")).strip() for k in key_cols)
    if "item" in row:
        return str(row["item"]).strip()
    return str(next(iter(row.values()))).strip() if row else ""


def check_data_file(path: Path, cfg: dict, target_period: str = "") -> tuple[list[Issue], list[dict]]:
    """校验单个数据文件，返回 (issues, rows)。"""
    issues: list[Issue] = []
    if not path.exists():
        return [_err(f"数据文件不存在：{path}")], []
    try:
        rows = read_rows(path)
    except Exception as exc:
        return [_err(f"数据文件读取失败 {path.name}：{exc}")], []
    if not rows:
        return [_err(f"数据文件为空：{path.name}")], rows

    headers = set(rows[0].keys())
    missing = [label for label, accepted in _required_data_columns(cfg)
               if not any(name in headers for name in accepted)]
    if missing:
        issues.append(_err(f"{path.name} 缺少列 {missing}（表单 {cfg.get('form_id', '?')} 需要）"))

    actual = data_periods(rows, cfg)
    configured = configured_periods(cfg)
    uncovered = sorted(p for p in actual if p not in configured)
    if uncovered:
        issues.append(_err(
            f"{path.name} 的期间 {uncovered} 在表单配置里没有对应列"
            f"（已配置：{sorted(configured)}）"))
    if target_period and actual and target_period not in actual:
        # 滚动预测常态：后端月份还没填数，导出时全空行被跳过，目标期间自然缺席
        issues.append(_info(
            f"{path.name} 不包含目标期间 {target_period}（实际：{sorted(actual)}）"))

    # 同一维度键 + 期间重复 → 会写到同一格，后写覆盖先写
    seen: dict[str, int] = {}
    dup: list[str] = []
    long_layout = cfg.get("layout") == "measure_blocks"
    field = str(cfg.get("period_field", "期间"))
    for row in rows:
        key = _row_key(row, cfg)
        if long_layout:
            key = f"{key}|{str(row.get(field, '')).strip()}"
        if not key.strip("|"):
            continue
        seen[key] = seen.get(key, 0) + 1
        if seen[key] == 2:
            dup.append(key.replace("|", " / "))
    if dup:
        issues.append(_err(f"{path.name} 有重复维度键（会互相覆盖）：{dup[:5]}"))
    return issues, rows


# ---------- 表单配置 ----------

def check_form_config(form_id: str, cfg: dict) -> list[Issue]:
    issues: list[Issue] = []
    layout = cfg.get("layout", "simple")
    if layout not in LAYOUTS:
        issues.append(_err(f"{form_id}: 未知 layout={layout!r}（可选 {sorted(LAYOUTS)}）"))
    for key in ("workbook_pattern", "sheet", "nav_cell"):
        if not cfg.get(key):
            issues.append(_err(f"{form_id}: 缺少 {key}"))
    if cfg.get("form_id") and cfg["form_id"] != form_id:
        issues.append(_warn(f"{form_id}: 配置内 form_id={cfg['form_id']!r} 与文件名不一致"))

    rows = cfg.get("input_rows")
    if not (isinstance(rows, list) and len(rows) == 2
            and all(isinstance(r, int) and not isinstance(r, bool) for r in rows)
            and rows[0] <= rows[1]):
        issues.append(_err(f"{form_id}: input_rows 应为 [起始行, 结束行] 且起始<=结束"))

    periods = configured_periods(cfg)
    if not periods:
        issues.append(_err(f"{form_id}: 没有配置任何期间列（{layout}）"))
    bad_periods = [p for p in periods if p and not (PERIOD_RE.match(p) or re.fullmatch(r"\d{4}", p))]
    if bad_periods:
        issues.append(_err(f"{form_id}: 期间键格式异常 {bad_periods[:5]}（应为 YYYY.MM 或 YYYY）"))

    key_cols = cfg.get("key_columns") or []
    if layout in ("keyed", "measure_blocks"):
        if not key_cols:
            issues.append(_err(f"{form_id}: {layout} 布局缺少 key_columns"))
        for k in key_cols:
            col = str(k.get("column", "")).strip().upper()
            if not COL_RE.match(col):
                issues.append(_err(f"{form_id}: key_columns 列号异常 {k!r}"))
            if not str(k.get("name", "")).strip():
                issues.append(_err(f"{form_id}: key_columns 缺少 name {k!r}"))
    if layout == "simple" and not cfg.get("item_column"):
        issues.append(_err(f"{form_id}: simple 布局缺少 item_column"))
    if layout == "measure_blocks" and not cfg.get("measures"):
        issues.append(_err(f"{form_id}: measure_blocks 布局缺少 measures"))
    unknown_computed = [str(m) for m in (cfg.get("computed_measures") or [])
                        if str(m) not in (cfg.get("measures") or {})]
    if unknown_computed:
        issues.append(_err(
            f"{form_id}: computed_measures {unknown_computed} 不在 measures 中"))

    # 期间列号唯一性（不同期间共用一列会互相覆盖）
    for label, mapping in (("month_columns", cfg.get("month_columns")),
                           ("period_blocks", cfg.get("period_blocks"))):
        if not mapping:
            continue
        cols = [str(v).strip().upper() for v in mapping.values()]
        dup_cols = sorted({c for c in cols if cols.count(c) > 1})
        if dup_cols:
            issues.append(_err(f"{form_id}: {label} 列号重复 {dup_cols}"))
    return issues


# ---------- 全局 ----------

def resolve_from_step(flow_cfg: dict, raw: object) -> int:
    """把 `--from-step` 的原始值（1 起序号或表单 ID）解析成 1 起的步骤序号。

    argparse 不限制类型（既可能是 "4" 也可能是表单 ID），而 check_flow 按
    数字比较，必须在进入校验前统一解析；None/空表示从头开始。
    """
    steps = flow_cfg.get("steps") or []
    key = str(raw).strip() if raw not in (None, "") else ""
    if not key:
        return 1
    if key.isdigit():
        idx = int(key)
        if not 1 <= idx <= len(steps):
            raise ValueError(f"--from-step {key} 超出范围（共 {len(steps)} 步）")
        return idx
    names = [str(step.get("form", "")) for step in steps]
    if key not in names:
        raise ValueError(f"--from-step 未找到表单 '{key}'，可选：{names}")
    return names.index(key) + 1


def check_flow(flow_cfg: dict, form_cfgs: dict[str, dict], data_dir: Path,
               check_data: bool = True, from_step: int = 1) -> list[Issue]:
    """校验 flow.yaml + 各步骤表单/数据文件。

    `from_step`（1 起）与 run-flow --from-step 对齐：前置步骤本次不跑，
    其数据文件校验一并跳过（"只跑后半段"的断点续跑不再被前置文件拦截）。
    """
    issues: list[Issue] = []

    # 直接传字符串（CLI 原始值）时也按序号/表单 ID 解析，避免 int >= str
    if isinstance(from_step, str):
        try:
            from_step = resolve_from_step(flow_cfg, from_step)
        except ValueError as exc:
            issues.append(_err(str(exc)))
            from_step = 1

    nav = flow_cfg.get("navigator") or {}
    source = str(nav.get("source") or "bw").lower()
    for key in ("workbook_pattern", "sheet", "version_cell", "refresh_flow_macro",
                "open_button_cell", "open_wait_sec"):
        if nav.get(key) in (None, ""):
            issues.append(_err(f"flow.yaml navigator 缺少 {key}"))
    if source == "file":
        launch = nav.get("launch_file")
        if not launch:
            issues.append(_err("flow.yaml navigator.source=file 但未配置 launch_file"))
        elif not Path(str(launch)).exists():
            issues.append(_err(f"导航表不存在：{launch}（source=file）"))
    elif source == "bw":
        if not nav.get("bw_connection"):
            issues.append(_warn("flow.yaml navigator.source=bw 但未配置 bw_connection，使用默认连接"))

    pkg = flow_cfg.get("package") or {}
    for key in ("workbook_pattern", "sheet", "execute_macro", "refresh_macro",
                "status_cell", "max_wait_sec"):
        if pkg.get(key) in (None, ""):
            issues.append(_err(f"flow.yaml package 缺少 {key}"))

    verify = flow_cfg.get("verify") or {}
    for key in ("nav_cell", "form_row", "sheet", "header_row", "indicator_rows"):
        if verify.get(key) in (None, ""):
            issues.append(_err(f"flow.yaml verify 缺少 {key}"))
    if verify and not verify.get("indicator_rows"):
        issues.append(_err("flow.yaml verify.indicator_rows 不能为空"))

    steps = flow_cfg.get("steps") or []
    if not steps:
        issues.append(_err("flow.yaml steps 为空"))
    if not data_dir.exists():
        issues.append(_err(f"数据目录不存在：{data_dir}"))
        check_data = False
    targets = re.findall(r"\d{4}\.\d{2}", data_dir.name)
    target = max(targets) if targets else ""

    for i, step in enumerate(steps, start=1):
        form_id = str(step.get("form", "")).strip()
        if not form_id:
            issues.append(_err(f"steps[{i}] 缺少 form"))
            continue
        if form_id not in form_cfgs:
            issues.append(_err(f"steps[{i}] 的表单配置不存在：{form_id}"))
            continue
        if not step.get("nav_cell"):
            issues.append(_err(f"steps[{i}]（{form_id}）缺少 nav_cell"))
        if not step.get("data"):
            issues.append(_err(f"steps[{i}]（{form_id}）缺少 data"))
            continue
        if check_data and i >= from_step:
            path = data_dir / step["data"]
            file_issues, _ = check_data_file(path, form_cfgs[form_id], target_period=target)
            issues.extend(Issue(x.level, f"{form_id}: {x.message}") for x in file_issues)
        after_close = step.get("after_close") or {}
        if after_close and not after_close.get("nav_cell"):
            issues.append(_err(f"steps[{i}]（{form_id}）after_close 缺少 nav_cell"))
        if after_close.get("packages") and not all(
                isinstance(r, int) and r > 0 for r in after_close["packages"]):
            issues.append(_err(f"steps[{i}]（{form_id}）after_close.packages 应为行号列表"))
    return issues
