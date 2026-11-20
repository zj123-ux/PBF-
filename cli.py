#!/usr/bin/env python
"""pbf-uploader CLI — load data file → preflight → write → save → verify.

Usage:
  python cli.py preflight --form E61_CBG_其他成本编制 --data upload.csv
  python cli.py apply     --form E61_CBG_其他成本编制 --data upload.csv
  python cli.py switch-version C03202611V1
  python cli.py run-package
  python cli.py read --form E61_CBG_其他成本编制 --cells AP66,AQ66,AS61
  python cli.py open-form --form-id E61_CBG_其他成本编制
  python cli.py scan --form E61_CBG_客户费用编制 --open --nav-cell AC57

  # 从模板导出 1–6 个期间 / 批量清除（期间可单值/列表/区间）
  python cli.py export --template 滚动预测模板表.xlsx --periods 2027.01-2027.04
  python cli.py clear  --periods 2027.01-2027.04 --run

  # 全流程编排（价格销量 → 客户费用 → HQ毛上+2运行包 → 其他成本 → 费用损益+1运行包 → 核对）
  python cli.py run-flow
  python cli.py run-flow --skip-fill      # 跳过填表，只跑计算/运行包/核对
  python cli.py run-flow --only-verify    # 只确认版本 + 读损益指标
  python cli.py run-flow --from-step 4    # 从第 4 步（其他成本）开始，前面 3 步跳过
  python cli.py run-flow --unit 乐华·电视 # 指定业务单元（品牌×品类），覆盖导航行
  python cli.py units                      # 列出/校验品牌×品类配置
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import yaml

from engine import (ExcelEngine, ExcelError, FormWriter, Navigator, RunPackageRunner,
                    FormReader, FormScanner, ensure_excel_ready, run_flow)
from engine.business_units import (UnitError, apply_unit, check_units,
                                   csv_dimension_warnings, default_unit)
from engine.clear import clear_data
from engine.config_check import (Issue, check_flow, check_form_config,
                                 resolve_from_step)
from engine.periods import parse_periods
from engine.template_export import export_template
from data_dirs import (latest_data_dir as find_latest_data_dir,
                       list_data_dirs as find_data_dirs,
                       unrecognized_data_dirs as find_unrecognized_data_dirs)
from paths import app_root

ROOT = app_root()
CONFIG_DIR = ROOT / "config" / "forms"
LOG_DIR = ROOT / "logs"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)
log = logging.getLogger("cli")


# ---------- helpers ----------

def load_yaml(path: Path) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def load_form_config(form_id: str) -> dict:
    candidates = sorted(CONFIG_DIR.glob(f"{form_id}*.yaml"))
    if not candidates:
        raise SystemExit(f"No config found for form_id='{form_id}' under {CONFIG_DIR}")
    if len(candidates) > 1:
        # 精确同名优先；否则说明 form_id 有歧义，继续取第一个会静默用错配置
        exact = CONFIG_DIR / f"{form_id}.yaml"
        if exact.exists():
            return load_yaml(exact)
        names = ", ".join(p.name for p in candidates)
        raise SystemExit(
            f"form_id='{form_id}' 匹配到多个配置文件：{names}。"
            "请改用精确的 form_id，或重命名多余配置。")
    return load_yaml(candidates[0])


def load_dimensions() -> dict:
    return load_yaml(ROOT / "config" / "dimensions.yaml")


def load_units_config() -> dict:
    path = ROOT / "config" / "business_units.yaml"
    return load_yaml(path) if path.exists() else {}


def load_data(path: Path) -> list[dict]:
    """Load CSV/Excel into list of dict rows（保留原始列名，供组合键匹配）。"""
    suffix = path.suffix.lower()
    if suffix in (".csv",):
        df = pd.read_csv(path)
    elif suffix in (".xlsx", ".xls"):
        df = pd.read_excel(path)
    else:
        raise SystemExit(f"Unsupported data file: {suffix}. Use .csv / .xlsx")
    return df.to_dict(orient="records")


_audit_pruned = False


def _audit_retention_days() -> int:
    """审计日志保留天数（天），<=0 表示永不清理。

    审计日志是财务操作留痕，默认不清理；只有显式设置
    PBF_AUDIT_RETENTION_DAYS（正整数）才会清理过期文件。
    非法值（非数字 / 0 / 负数）一律告警并按"不清理"处理，
    不能因为环境变量写错就让所有命令崩在 import 期。
    """
    raw = (os.environ.get("PBF_AUDIT_RETENTION_DAYS") or "").strip()
    if not raw:
        return 0
    try:
        days = int(raw)
    except ValueError:
        log.warning("PBF_AUDIT_RETENTION_DAYS=%r 不是数字，忽略（审计日志不清理）", raw)
        return 0
    if days <= 0:
        log.warning("PBF_AUDIT_RETENTION_DAYS=%d 非正数，忽略（审计日志不清理）", days)
        return 0
    return days


def _prune_audit_logs(log_dir: Path = LOG_DIR) -> None:
    """删除超过保留期的审计日志（每次进程只扫描一次；未配置保留期时不清理）。"""
    global _audit_pruned
    if _audit_pruned:
        return
    _audit_pruned = True
    max_age_days = _audit_retention_days()
    if max_age_days <= 0:
        return
    cutoff = time.time() - max_age_days * 86400
    removed = 0
    for f in log_dir.glob("*.jsonl"):
        try:
            if f.stat().st_mtime < cutoff:
                f.unlink()
                removed += 1
        except OSError:
            pass
    if removed:
        print(f"[审计] 已按 PBF_AUDIT_RETENTION_DAYS={max_age_days} "
              f"清理 {removed} 个过期审计日志文件")


def audit_log(action: str, payload: dict) -> None:
    LOG_DIR.mkdir(exist_ok=True)
    _prune_audit_logs()
    fn = LOG_DIR / f"{datetime.now():%Y%m%d}.jsonl"
    record = {"ts": datetime.now().isoformat(), "action": action, **payload}
    with open(fn, "a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False, default=str) + "\n")


def print_issues(issues: list[Issue]) -> tuple[int, int, int]:
    """打印校验结果，返回 (错误数, 告警数, 提示数)。"""
    errors = [i for i in issues if i.level == "error"]
    warnings = [i for i in issues if i.level == "warning"]
    infos = [i for i in issues if i.level == "info"]
    for i in errors:
        print(f"  ✗ {i.message}")
    for i in warnings:
        print(f"  ! {i.message}")
    for i in infos:
        print(f"  · {i.message}")
    return len(errors), len(warnings), len(infos)


def validate_before_run(flow_cfg: dict, form_cfgs: dict[str, dict], data_dir: Path,
                        check_data: bool = True, from_step: int = 1) -> None:
    """启动 Excel 前的离线校验；有错误直接中止。"""
    issues = check_flow(flow_cfg, form_cfgs, data_dir, check_data=check_data,
                        from_step=from_step)
    if not issues:
        return
    print("\n[配置校验]")
    errors, warnings, infos = print_issues(issues)
    if errors:
        print(f"\n配置校验未通过（{errors} 个错误，{warnings} 个告警），已中止；未启动 Excel。")
        sys.exit(1)
    extra = f"，{infos} 条提示" if infos else ""
    print(f"\n配置校验通过（{warnings} 个告警{extra}）")


# ---------- commands ----------

def cmd_preflight(args) -> None:
    engine = ExcelEngine()
    cfg = load_form_config(args.form)
    writer = FormWriter(engine, cfg)
    data = load_data(Path(args.data))

    print(f"\n=== Preflight: form={args.form}, workbook_pattern={cfg['workbook_pattern']} ===")
    result = writer.preflight(data)

    # Print checks
    print(f"\n[Checks] ok={result.ok}")
    for label, pass_, detail in result.checks:
        mark = "✓" if pass_ else "✗"
        print(f"  {mark} {label}: {detail}")

    # Matched rows
    print(f"\n[Matched] {len(result.matched_rows)} cells will be written:")
    header = f"  {'item':<20} {'month':<10} {'cell':<10} {'old':<14} {'new':<14}"
    print(header)
    print("  " + "-" * (len(header) - 2))
    for row in result.matched_rows:
        old = row["old_value"] if row["old_value"] not in (None, "") else "(空)"
        print(f"  {row['item']:<20} {row['month']:<10} {row['cell']:<10} {str(old):<14} {str(row['new_value']):<14}")

    if result.skipped:
        print(f"\n[Skipped — not found in form] {len(result.skipped)} items:")
        for name in result.skipped:
            print(f"  ✗ {name}")

    audit_log("preflight", {"form": args.form, "matched": len(result.matched_rows),
                            "skipped": result.skipped})

    if not result.ok:
        sys.exit(1)
    print(f"\nPreflight PASSED. Run 'apply' to execute write+save.")


def cmd_apply(args) -> None:
    engine = ExcelEngine()
    cfg = load_form_config(args.form)
    writer = FormWriter(engine, cfg)
    data = load_data(Path(args.data))

    print(f"\n=== Apply: form={args.form} ===")

    # --fresh：先关掉旧实例再重新打开，避免 AO 内部状态被历史操作污染
    if args.fresh and cfg.get("nav_cell"):
        nav = Navigator(engine, load_dimensions()["navigator"])
        try:
            wb_old = engine.find_workbook(cfg["workbook_pattern"])
            engine.close_workbook(wb_old, save=False)
            print("[Fresh] 已关闭旧实例")
            time.sleep(2)
        except ExcelError:
            pass
        print(f"[Fresh] 重新打开 {cfg['nav_cell']}")
        nav.open_form(cfg["nav_cell"])
        engine.get_workbook(cfg["workbook_pattern"], wait_sec=60)

    pf = writer.preflight(data)
    if not pf.ok:
        print("Preflight FAILED — aborting. Run preflight to see details.")
        audit_log("apply", {"form": args.form, "status": "preflight_failed",
                            "checks": [(l, p, d) for l, p, d in pf.checks]})
        sys.exit(1)

    result = writer.write_and_save(pf, verify_after=False)

    print(f"\n[Result] written={result['written']}, saved={result['saved']}")

    # 落库验证：按 AO 实际流程「关闭工作簿(不保存) → 重新打开 → 读值」
    if not args.no_verify and result["saved"] and result.get("cells"):
        nav_cell = cfg.get("nav_cell")
        if not nav_cell:
            print("[Verify] 未配置 nav_cell，跳过落库验证")
        else:
            print("[Verify] 关闭工作簿(不保存) → 重开 → 读值")
            wb = engine.find_workbook(cfg["workbook_pattern"])
            engine.close_workbook(wb, save=False)
            time.sleep(5)
            nav = Navigator(engine, load_dimensions()["navigator"])
            nav.open_form(nav_cell)
            wb2 = engine.get_workbook(cfg["workbook_pattern"], wait_sec=60)
            sheet2 = engine.open_sheet(wb2, cfg["sheet"])
            ok_all = True
            for cell, expected in result["cells"].items():
                actual = engine.read(sheet2, cell)
                try:
                    diff = abs(float(actual or 0) - float(expected))
                    ok = diff < 1e-2
                except (TypeError, ValueError):
                    ok = str(actual) == str(expected)
                if not ok:
                    ok_all = False
                    result["failures"].append(f"{cell}: 期望 {expected}, 重开后={actual}")
            result["verified"] = ok_all

    if result.get("failures"):
        print("[Failures]")
        for f in result["failures"]:
            print(f"  ✗ {f}")
    elif result.get("verified"):
        print("所有单元格重开后仍在 BPC 中。✓")

    audit_log("apply", {"form": args.form, "written": result["written"],
                        "saved": result["saved"], "verified": result.get("verified"),
                        "failures": result["failures"]})

    if result.get("failures") or not result.get("saved"):
        sys.exit(1)


def cmd_switch_version(args) -> None:
    # 切版本要在导航表上操作，先把导航表打开（等价于双击进入 PBF），再连 COM
    flow_cfg = load_yaml(ROOT / "config" / "flow.yaml")
    engine = ensure_excel_ready(flow_cfg)
    dims = load_dimensions()
    member_key = args.member_key
    display = dims["versions"].get(member_key, args.display_name or member_key)
    nav_cfg = dims["navigator"]
    nav = Navigator(engine, nav_cfg)
    nav.switch_version(member_key, display)
    audit_log("switch_version", {"member_key": member_key, "display": display})


def cmd_run_package(args) -> None:
    engine = ExcelEngine()
    dims = load_dimensions()
    runner = RunPackageRunner(engine, dims["runner"])
    print("Executing run package, waiting up to 20s...")
    result = runner.execute_and_wait(wait_seconds=args.wait)
    print(f"\n[Run] status={result.status}, message={result.message}")
    audit_log("run_package", {"status": result.status, "message": result.message})


def cmd_read(args) -> None:
    engine = ExcelEngine()
    cfg = load_form_config(args.form)
    reader = FormReader(engine)
    cells = [c.strip() for c in args.cells.split(",")]
    values = reader.read_cells(cfg["workbook_pattern"], cfg["sheet"], cells)
    print(f"\n[Read] form={args.form}")
    for c, v in values.items():
        print(f"  {c} = {v}")


def cmd_open_form(args) -> None:
    engine = ExcelEngine()
    dims = load_dimensions()
    nav = Navigator(engine, dims["navigator"])
    # workflow_row_col — accept user-specified cell or look up in form config
    cell = args.cell or load_form_config(args.form).get("nav_cell")
    if not cell:
        sys.exit("Need --cell AA57 (workflow row cell in navigator) or set nav_cell in form yaml")
    new_wb = nav.open_form(cell)
    print(f"Opened: {new_wb}")
    audit_log("open_form", {"form": args.form, "cell": cell, "new_wb": new_wb})


def cmd_scan(args) -> None:
    """扫描已打开（或自动打开）的编制表，推导 YAML 配置。"""
    engine = ExcelEngine()

    cfg: dict = {}
    if args.form:
        cfg = load_form_config(args.form)
    pattern = args.pattern or cfg.get("workbook_pattern")
    sheet_name = args.sheet or cfg.get("sheet")
    if not pattern or not sheet_name:
        sys.exit("--form 需要已有配置，或显式给 --pattern 和 --sheet")

    # 工作簿未打开且给了导航行 → 自动打开
    try:
        wb = engine.find_workbook(pattern)
    except ExcelError:
        nav_cell = args.nav_cell or cfg.get("nav_cell")
        if not (args.open and nav_cell):
            sys.exit(f"工作簿未打开: {pattern}\n加 --open --nav-cell AC65 可自动打开")
        dims = load_dimensions()
        nav = Navigator(engine, dims["navigator"])
        print(f"通过导航表打开: {nav_cell}")
        nav.open_form(nav_cell)
        wb = engine.get_workbook(pattern, wait_sec=30)

    sheet = engine.open_sheet(wb, sheet_name)
    info = FormScanner(engine).scan(wb, sheet)

    print(f"\n=== Scan: {info['workbook']} / {info['sheet']} ===")
    print(f"表头行: {info['header_row']}")
    print(f"科目列: {info['item_column']}  候选: {info['item_column_candidates']}")
    print(f"数据行: {info['input_rows']}")
    print(f"模式格: {info['mode_cells']}  标签格: {info['mode_label_cell']}")
    print(f"版本格: {info['version_cell']}")
    print(f"\n期间列 ({len(info['month_columns'])}):")
    for m, c in info["month_columns"].items():
        print(f"  {m} → {c}")

    if args.write:
        cfg.setdefault("form_id", args.form or Path(pattern).stem)
        cfg.setdefault("workbook_pattern", pattern)
        cfg.setdefault("sheet", sheet_name)
        cfg["month_columns"] = info["month_columns"]
        cfg["item_column"] = info["item_column"]
        cfg["input_rows"] = info["input_rows"]
        if info["mode_cells"]:
            cfg["mode_cell"] = info["mode_cells"][0]["cell"]
            if len(info["mode_cells"]) > 1:
                cfg["mode_label_cell"] = info["mode_cells"][1]["cell"]
        if info["mode_label_cell"]:
            cfg["mode_label_cell"] = info["mode_label_cell"]
        if info["version_cell"]:
            cfg["version_cell"] = info["version_cell"]
        out = CONFIG_DIR / f"{cfg['form_id']}.yaml"
        with open(out, "w", encoding="utf-8") as f:
            yaml.safe_dump(cfg, f, allow_unicode=True, sort_keys=False)
        print(f"\n已写入: {out}")

    audit_log("scan", {"form": args.form or pattern, "item_column": info["item_column"],
                      "months": len(info["month_columns"]), "rows": info["input_rows"]})


def latest_data_dir() -> Path:
    """默认数据目录：按期间取最新（排除 *_clear）；没有则报错提示先导出数据。"""
    latest = find_latest_data_dir(ROOT)
    if latest is None:
        raise SystemExit(
            "未找到任何 data_* 数据目录。请先导出上传数据：\n"
            "  python cli.py export --template 滚动预测模板表.xlsx --periods 2027.01-2027.04\n"
            "或改用 --data-dir 指定已有数据目录。")
    return latest


def existing_data_dirs() -> list[Path]:
    """所有已存在的上传数据目录（按期间升序，排除 *_clear 清除目录）。"""
    return find_data_dirs(ROOT)


def warn_unrecognized_data_dirs() -> None:
    """data_ 前缀但不符合命名规则的目录不参与数据源选择，遇到时提示一句。"""
    ignored = find_unrecognized_data_dirs(ROOT)
    if ignored:
        names = "、".join(p.name for p in ignored)
        print(f"[提示] 以下 data_ 开头的目录不符合命名规则 data_YYYY.MM"
              f"（或 data_YYYY.MM_YYYY.MM），未当作数据源：{names}")


def default_export_dir(periods: list[str]) -> Path:
    """导出目录：单期间 data_YYYY.MM，多期间 data_<首>_<末>。"""
    first, last = periods[0], periods[-1]
    name = f"data_{first}" if first == last else f"data_{first}_{last}"
    return ROOT / name


def clear_out_dir(periods: list[str]) -> Path:
    """清除目录：data_YYYY.MM_clear 或 data_<首>_<末>_clear。"""
    first, last = periods[0], periods[-1]
    rng = first if first == last else f"{first}_{last}"
    return ROOT / f"data_{rng}_clear"


def cmd_export(args) -> None:
    """从《滚动预测模板表.xlsx》导出 1–6 个期间的上传数据。"""
    periods = parse_periods(args.periods)
    template = Path(args.template)
    out_dir = Path(args.out_dir) if args.out_dir else default_export_dir(periods)

    print(f"\n=== export ===")
    print(f"模板: {template}")
    print(f"期间: {', '.join(periods)}")
    print(f"输出: {out_dir}")
    rows = export_template(template, periods, out_dir, zero=args.zero)
    print("\n[导出结果]")
    for name, n in rows.items():
        print(f"  {name}: {n} 行")
    for warning in csv_dimension_warnings(out_dir):
        print(f"[数据告警] {warning}")
    if args.zero:
        print("\n（zero 模式：所有数值已填 0）")
    audit_log("export", {"template": str(template), "periods": periods,
                         "out_dir": str(out_dir), "zero": args.zero, "rows": rows})


def cmd_clear(args) -> None:
    """批量清除：把既有 data_* 目录里填过的数字清成 0，期间可自由选择。"""
    warn_unrecognized_data_dirs()
    periods = parse_periods(args.periods)
    data_dirs = [Path(d) for d in args.data_dirs] if args.data_dirs else existing_data_dirs()
    out_dir = Path(args.out_dir) if getattr(args, "out_dir", None) else clear_out_dir(periods)

    print(f"\n=== clear ===")
    print(f"清除期间: {', '.join(periods)}")
    print(f"数据来源: {', '.join(d.name for d in data_dirs)}")
    print(f"输出: {out_dir}")
    try:
        rows = clear_data(data_dirs, periods, out_dir)
    except ValueError as exc:
        print(f"\n[清除中止] {exc}")
        raise SystemExit(1)
    print("\n[清除结果]（全 0）")
    for name, n in rows.items():
        print(f"  {name}: {n} 行")
    audit_log("clear", {"periods": periods, "data_dirs": [str(d) for d in data_dirs],
                        "out_dir": str(out_dir), "rows": rows})

    if args.run:
        print("\n[清除] 接着运行 run-flow 填 0…")
        from types import SimpleNamespace
        rf_args = SimpleNamespace(
            data_dir=str(out_dir), skip_fill=False, only_verify=False,
            from_step=None, version=None, preflight_only=False,
            version_key=None, unit=getattr(args, "unit", None),
            report_file=getattr(args, "report_file", None))
        cmd_run_flow(rf_args)


def cmd_validate(args) -> None:
    """离线校验：flow.yaml / 表单配置 / 数据文件 / 导航表路径，不连接 Excel。"""
    flow_cfg = load_yaml(ROOT / "config" / "flow.yaml")
    issues: list[Issue] = []

    form_cfgs: dict[str, dict] = {}
    for path in sorted(CONFIG_DIR.glob("*.yaml")):
        try:
            cfg = load_yaml(path)
        except Exception as exc:
            issues.append(Issue("error", f"{path.name} 读取失败：{exc}"))
            continue
        form_id = str(cfg.get("form_id") or path.stem)
        form_cfgs[form_id] = cfg
        issues.extend(check_form_config(form_id, cfg))

    check_data = not bool(getattr(args, "config_only", False))
    if args.data_dir:
        data_dir = Path(args.data_dir)
    elif check_data:
        try:
            data_dir = latest_data_dir()
        except SystemExit as exc:
            issues.append(Issue("error", str(exc)))
            data_dir = ROOT
    else:
        data_dir = ROOT

    issues.extend(check_flow(flow_cfg, form_cfgs, data_dir, check_data=check_data))

    print("\n[配置校验]")
    if not issues:
        print("  全部通过。")
        return
    errors, warnings, infos = print_issues(issues)
    extra = f"，{infos} 条提示" if infos else ""
    print(f"\n结果：{errors} 个错误，{warnings} 个告警{extra}")
    if errors:
        sys.exit(1)


def cmd_units(args) -> None:
    """列出并校验业务单元（品牌×品类）配置。"""
    units_cfg = load_units_config()
    flow_cfg = load_yaml(ROOT / "config" / "flow.yaml")
    if not units_cfg.get("units"):
        print("未配置业务单元：config/business_units.yaml")
        return
    print(f"\n品牌: {', '.join(units_cfg.get('品牌') or []) or '(空)'}")
    print(f"品类: {', '.join(units_cfg.get('品类') or []) or '(空)'}")
    print("\n业务单元:")
    for name, unit in units_cfg["units"].items():
        unit = unit or {}
        mark = "✓ 启用" if unit.get("启用") else "  未配置"
        print(f"  {mark}  {name}  （{unit.get('品牌', '')}·{unit.get('品类', '')}）")
    issues = check_units(units_cfg, flow_cfg)
    print(f"\n校验: {'通过' if not issues else f'{len(issues)} 项问题'}")
    for issue in issues:
        print(f"  {issue}")


def cmd_run_flow(args) -> None:
    """按业务顺序跑完整 PBF 滚动预测流程。"""
    warn_unrecognized_data_dirs()
    base_flow = load_yaml(ROOT / "config" / "flow.yaml")
    units_cfg = load_units_config()
    requested = (getattr(args, "unit", None) or "").strip()
    # 与 webapp 的 _resolve_unit 统一：未指定 --unit 时默认第一个启用单元
    unit_name = requested or default_unit(units_cfg) or ""
    if unit_name:
        try:
            flow_cfg = apply_unit(base_flow, units_cfg, unit_name)
        except UnitError as exc:
            raise SystemExit(f"业务单元配置错误: {exc}")
        unit = flow_cfg["_unit"]
        note = "" if requested else "（未指定 --unit，默认第一个启用单元）"
        print(f"业务单元: {unit_name}（{unit['brand']}·{unit['category']}）{note}")
    elif units_cfg.get("units"):
        # 配置了单元却没有任何启用：拒绝静默跑 flow.yaml 写死行（与 webapp 行为一致）
        raise SystemExit(
            "业务单元配置错误: 没有启用的业务单元（至少一个单元 启用: true）；"
            "如不需要按单元运行，请移除 config/business_units.yaml 的 units 配置")
    else:
        flow_cfg = base_flow
        print("业务单元: 未配置业务单元，沿用 flow.yaml 的默认导航行")
    if units_cfg.get("units"):
        # 只校验本次要跑的单元，其余单元（占位/未启用）的问题不阻塞本次运行
        for issue in check_units(units_cfg, base_flow, only=unit_name or None):
            if issue.startswith("错误:"):
                raise SystemExit(f"业务单元配置错误: {issue}")
            print(f"[业务单元] {issue}")

    form_cfgs = {s["form"]: load_form_config(s["form"]) for s in flow_cfg["steps"]}
    data_dir = Path(args.data_dir) if args.data_dir else latest_data_dir()

    print(f"\n=== run-flow ===")
    print(f"数据目录: {data_dir}")
    for warning in csv_dimension_warnings(data_dir):
        print(f"[数据告警] {warning}")
    targets = [f"{y}.{m}" for y, m in re.findall(r"(\d{4})\.(\d{2})", data_dir.name)]
    if targets:
        print(f"目标期间: {max(targets)}")
    preflight_only = bool(getattr(args, "preflight_only", False))
    print(f"模式: {'预检干跑（不写入）' if preflight_only else ('跳填表' if args.skip_fill else '填表')}"
          f" | 仅核对: {args.only_verify}"
          + (f" | 从步骤 {args.from_step} 开始" if args.from_step else "") + "\n")

    # --from-step 既接受序号也接受表单 ID；校验按数字比较，先统一解析
    raw_from_step = getattr(args, "from_step", None)
    try:
        from_step = resolve_from_step(flow_cfg, raw_from_step)
    except ValueError as exc:
        raise SystemExit(str(exc))

    # 先做离线校验：配置/数据有问题时不要启动 Excel、更不要跑到一半才失败
    validate_before_run(flow_cfg, form_cfgs, data_dir,
                        check_data=not (args.skip_fill or args.only_verify),
                        from_step=from_step)

    engine = ensure_excel_ready(flow_cfg)

    report = run_flow(engine, flow_cfg, form_cfgs, data_dir,
                      skip_fill=args.skip_fill, only_verify=args.only_verify,
                      expected_version=args.version,
                      from_step=from_step if raw_from_step else None,
                      preflight_only=preflight_only,
                      switch_version_key=getattr(args, "version_key", None))
    if unit_name:
        report["unit"] = unit_name

    if report.get("login"):
        print("\n[AO 登录] 已就绪")
    print(f"\n版本: {report['version']}")

    if report.get("preflights"):
        print("\n[预检（只读，未写入任何数据）]")
        for p in report["preflights"]:
            mark = "✓" if p["ok"] else "✗"
            extra = f"，未匹配 {len(p['skipped'])} 项" if p["skipped"] else ""
            print(f"  {mark} {p['index']}. {p['form']}: 命中 {p['matched']} 格{extra}")
            for label, ok, detail in p.get("checks", []):
                if not ok:
                    print(f"      ✗ {label}: {detail}")
            for s in (p.get("skipped") or [])[:5]:
                print(f"      ✗ 未匹配: {s}")
            if len(p.get("skipped") or []) > 5:
                print(f"      … 另有 {len(p['skipped']) - 5} 项未匹配")
        all_ok = all(p["ok"] for p in report["preflights"])
        print(f"\n预检结果: {'全部通过' if all_ok else '存在失败项，请先修正再正式运行'}")
        if not all_ok:
            audit_log("run-flow", report)
            sys.exit(1)
        audit_log("run-flow", report)
        return

    if report["steps"]:
        print("\n[填表 / 计算]")
        for s in report["steps"]:
            if "written" in s:
                if s.get("unchanged"):
                    mark = "✓ 数据未变化，无需保存"
                elif s.get("verified") is True:
                    mark = "✓ 已落库校验"
                elif s.get("verified") is False:
                    mark = "✗ 落库校验失败"
                elif s.get("saved"):
                    mark = "✓ 已保存（未校验）"
                else:
                    mark = "⚠ 待人工保存"
                print(f"  {s['form']}: 写入 {s['written']} 格  {mark}")
                for failure in s.get("failures") or []:
                    print(f"      ✗ {failure}")
            elif "calc" in s:
                print(f"  {s['form']}: 执行 {s['calc']} ✓")

    if report["packages"]:
        print("\n[运行包]")
        for p in report["packages"]:
            mark = "✓" if p["ok"] else "✗"
            print(f"  {mark} {p['status']} 时长={p['duration']}  {p['package']}")

    if report["indicators"]:
        print("\n[损益核对]")
        periods = list(next(iter(report["indicators"].values())).keys())
        header = "  " + "指标".ljust(16) + "".join(str(p).rjust(20) for p in periods)
        print(header)
        print("  " + "-" * (len(header) - 2))
        for name, vals in report["indicators"].items():
            row = "  " + name.ljust(16)
            for p in periods:
                v = vals.get(p)
                row += ("" if v is None else f"{v:,.4f}".rstrip("0").rstrip(".")).rjust(20)
            print(row)

    audit_log("run-flow", report)

    # 供 Web 界面读取结构化结果（version/steps/packages/indicators）
    if getattr(args, "report_file", None):
        rp = Path(args.report_file)
        rp.parent.mkdir(parents=True, exist_ok=True)
        with open(rp, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, default=str)

    print("\n完成。")


# ---------- main ----------

def _init_console() -> None:
    """打包成 exe 后 stdout 编码可能与控制台代码页不一致（中文变乱码），统一成控制台代码页。

    仅当输出是真实控制台（isatty）时才改。Web 界面把 stdout 重定向到管道，父进程按
    UTF-8 读日志；这里若再改成 cp936 会写出乱码。管道场景沿用 PYTHONIOENCODING。
    """
    try:
        if not (hasattr(sys.stdout, "isatty") and sys.stdout.isatty()):
            return
        import ctypes

        cp = ctypes.windll.kernel32.GetConsoleOutputCP()
        enc = f"cp{cp}" if cp else "utf-8"
        for stream in (sys.stdout, sys.stderr):
            if hasattr(stream, "reconfigure"):
                stream.reconfigure(encoding=enc, errors="replace")
    except Exception:
        pass


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="pbf-uploader", description="Auto-update PBF forecast via Excel COM")
    sub = p.add_subparsers(dest="cmd", required=True)

    pp = sub.add_parser("preflight", help="Check data against form without writing")
    pp.add_argument("--form", required=True, help="Form id, e.g. E61_CBG_其他成本编制")
    pp.add_argument("--data", required=True, help="CSV/XLSX data file")

    ap = sub.add_parser("apply", help="Preflight + write + save + verify")
    ap.add_argument("--form", required=True)
    ap.add_argument("--data", required=True)
    ap.add_argument("--no-verify", action="store_true", help="Skip post-close-reopen verification")
    ap.add_argument("--fresh", action="store_true", help="Close old instance and reopen before writing")

    sp = sub.add_parser("switch-version", help="Change forecast version")
    sp.add_argument("member_key", help="BPC member key, e.g. C03202611V1")
    sp.add_argument("--display-name", help="Optional display label")

    rp = sub.add_parser("run-package", help="Execute BPC run package and poll")
    rp.add_argument("--wait", type=int, default=20, help="Seconds to wait")

    rr = sub.add_parser("read", help="Read cell values from an open form")
    rr.add_argument("--form", required=True)
    rr.add_argument("--cells", required=True, help="Comma-separated cells like AP66,AQ66")

    op = sub.add_parser("open-form", help="Open a child form via navigator")
    op.add_argument("--form", required=True)
    op.add_argument("--cell", help="Workflow row cell in navigator, e.g. AA57")

    sc = sub.add_parser("scan", help="Scan form layout and derive YAML config")
    sc.add_argument("--form", help="Existing form id (reuse its pattern/sheet)")
    sc.add_argument("--pattern", help="Workbook name pattern, e.g. 'YCSRW017_*'")
    sc.add_argument("--sheet", help="Worksheet name")
    sc.add_argument("--open", action="store_true", help="Auto-open via navigator if not open")
    sc.add_argument("--nav-cell", help="Navigator cell, e.g. AC65")
    sc.add_argument("--write", action="store_true", help="Write derived config back to YAML")

    vd = sub.add_parser("validate", help="离线校验配置与数据（不连接 Excel）")
    vd.add_argument("--data-dir", help="数据目录，默认取最新的 data_YYYY.MM")
    vd.add_argument("--config-only", action="store_true", help="只校验配置，不检查数据文件")

    rf = sub.add_parser("run-flow", help="按业务顺序跑完整流程（填表→计算→运行包→核对）")
    rf.add_argument("--data-dir", help="数据目录，默认取最新的 data_YYYY.MM")
    rf.add_argument("--skip-fill", action="store_true", help="跳过填表，只跑计算/运行包/核对")
    rf.add_argument("--only-verify", action="store_true", help="只做版本确认+损益核对")
    rf.add_argument("--from-step",
                    help="从第几步开始（1 开始的序号，或表单 ID），之前的步骤整步跳过")
    rf.add_argument("--version", help="期望版本子串，如 '滚动预测2026年11月月初版'")
    rf.add_argument("--version-key",
                    help="运行前先切到该版本成员（如 C03202611V1），随后自动刷新")
    rf.add_argument("--preflight-only", action="store_true",
                    help="干跑：只打开各表单做行匹配/期间覆盖预检，不写入任何数据")
    rf.add_argument("--report-file", help="把最终 report 写成 JSON 到指定路径（供 Web 界面读取）")
    rf.add_argument("--unit",
                    help="业务单元（品牌×品类），如 乐华·电视；覆盖 flow.yaml 的 nav_cell")

    ex = sub.add_parser("export", help="从滚动预测模板表导出 1–6 个期间的上传数据")
    ex.add_argument("--template", required=True, help="滚动预测模板表 .xlsx 路径")
    ex.add_argument("--periods", required=True,
                    help="期间，支持单值/列表/区间，如 2027.01 或 2027.01,2027.03 或 2027.01-2027.04")
    ex.add_argument("--out-dir", help="输出目录，默认 data_<首>_<末>")
    ex.add_argument("--zero", action="store_true", help="所有数值填 0（导出清零数据）")

    cl = sub.add_parser("clear", help="批量清除：把既有 data_* 目录里填过的数字清成 0")
    cl.add_argument("--periods", required=True, help="要清除的期间，支持单值/列表/区间")
    cl.add_argument("--data-dirs", nargs="*", help="数据来源目录，默认全部既有 data_*")
    cl.add_argument("--out-dir", help="输出目录，默认 data_<首>_<末>_clear")
    cl.add_argument("--run", action="store_true", help="生成后接着运行 run-flow 填 0")
    cl.add_argument("--unit", help="业务单元（品牌×品类），配合 --run 使用")
    cl.add_argument("--report-file", help="run-flow 结果 JSON 路径（配合 --run）")

    un = sub.add_parser("units", help="列出并校验业务单元（品牌×品类）配置")

    return p


def main() -> None:
    _init_console()
    parser = build_parser()
    argv = sys.argv[1:]
    double_clicked = not argv
    if double_clicked:
        argv = ["run-flow"]            # 双击 exe 时默认跑全流程
    args = parser.parse_args(argv)
    try:
        {"preflight": cmd_preflight,
         "apply": cmd_apply,
         "switch-version": cmd_switch_version,
         "run-package": cmd_run_package,
         "read": cmd_read,
         "open-form": cmd_open_form,
         "scan": cmd_scan,
         "export": cmd_export,
         "clear": cmd_clear,
         "validate": cmd_validate,
         "units": cmd_units,
         "run-flow": cmd_run_flow}[args.cmd](args)
    except ExcelError as exc:
        print(f"\n[ExcelError] {exc}", file=sys.stderr)
        sys.exit(2)
    except KeyboardInterrupt:
        print("\nInterrupted.", file=sys.stderr)
        sys.exit(130)
    except Exception as exc:                  # COM 断开 / Excel 崩溃 / 配置错误等
        if os.environ.get("PBF_DEBUG"):
            raise
        print(f"\n[错误] {exc}", file=sys.stderr)
        print("若是 Excel 被关闭或崩溃：重开导航表后重试，或用 --from-step N 续跑。",
              file=sys.stderr)
        sys.exit(2)
    finally:
        if double_clicked:             # 双击运行的窗口不要一闪而过
            try:
                input("\n按回车键退出…")
            except Exception:
                pass


if __name__ == "__main__":
    main()
