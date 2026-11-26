"""业务单元（品牌 × 品类）配置校验与流程覆盖。

`config/business_units.yaml` 为每个业务单元按顺序给出各编制表的 `nav_cell`。
`apply_unit()` 把这些 `nav_cell` 覆盖到 `flow.yaml` 的 steps 上，让「品牌 × 品类」
选择真正决定打开导航表的哪一行（而不是永远跑 flow.yaml 里写死的乐华行）。

纯逻辑模块：不依赖 pywin32 / COM，也不直接读 YAML（由调用方传入已解析的 dict），
便于单元测试。
"""
from __future__ import annotations

import copy
import csv
import re
from pathlib import Path

DIM_COLUMNS = ("品牌", "品类")


class UnitError(ValueError):
    """业务单元配置错误（未知/未启用/步骤不对应等）。"""


class NavRowError(ValueError):
    """按 AA+AB+AC 标签定位导航行失败（零命中 / 多行命中且配置行不在其中）。"""


def _norm(s: object) -> str:
    """表单名归一化：去掉空白和下划线后做包含匹配。"""
    return re.sub(r"[\s_]+", "", str(s or ""))


def units(units_cfg: dict) -> dict[str, dict]:
    return dict(units_cfg.get("units") or {})


def unit_names(units_cfg: dict) -> list[str]:
    return list(units(units_cfg))


def enabled_unit_names(units_cfg: dict) -> list[str]:
    return [name for name, u in units(units_cfg).items() if u.get("启用")]


def default_unit(units_cfg: dict) -> str | None:
    """默认业务单元：第一个启用的单元。"""
    names = enabled_unit_names(units_cfg)
    return names[0] if names else None


def get_unit(units_cfg: dict, name: str) -> dict:
    """取一个已启用的单元；未知或未启用抛 UnitError。"""
    all_units = units(units_cfg)
    if name not in all_units:
        raise UnitError(f"未知业务单元 '{name}'，可选：{list(all_units)}")
    unit = all_units[name] or {}
    if not unit.get("启用"):
        raise UnitError(
            f"业务单元 '{name}' 未启用（config/business_units.yaml 中 启用: false）")
    return unit


def _skip_hints(unit: dict) -> list[str]:
    """单元要跳过的流程步骤（表单名关键字，如 [HQ毛上调整]）。"""
    raw = unit.get("跳过步骤") or []
    if not isinstance(raw, list):
        raise UnitError("跳过步骤 必须是列表（如 [HQ毛上调整]）")
    return [str(x).strip() for x in raw if str(x).strip()]


def _apply_skips(flow_steps: list[dict], unit: dict,
                 name: str) -> tuple[list[dict], list[dict]]:
    """按单元 `跳过步骤` 从流程里去掉整步，返回 (剩余步骤, 被跳过的步骤)。

    跳过的步骤连同它的 after_save 计算宏、after_close 运行包一起不执行
    （与 `run-flow --from-step` 的语义一致）。关键字必须唯一命中一个步骤，
    否则视为配置错误——猜错一步会漏填/漏算整张表。
    返回的步骤是同一批 dict 对象，调用方修改它们会改到入参 flow，因此
    apply_unit 必须传入 deepcopy 后的 steps。
    """
    hints = _skip_hints(unit)
    remaining = list(flow_steps)
    skipped: list[dict] = []
    for hint in hints:
        h = _norm(hint)
        hits = [s for s in remaining
                if h and (h in _norm(s.get("form", ""))
                          or _norm(s.get("form", "")) in h)]
        if len(hits) != 1:
            names = [s.get("form") for s in remaining]
            raise UnitError(
                f"业务单元 '{name}' 的 跳过步骤 '{hint}' 匹配到 {len(hits)} 个流程步骤"
                f"（应恰好 1 个）：{names}")
        remaining.remove(hits[0])
        skipped.append(hits[0])
    return remaining, skipped


def _pair_steps(flow_steps: list[dict], unit: dict, name: str) -> list[tuple[dict, dict]]:
    """按顺序把单元 steps 与 flow steps 配对并校验。

    顺序是权威的（YAML 注释约定 steps 与 flow.yaml 顺序一一对应）；同时用
    `表单` 名称做一次包含匹配，防止配置顺序被改乱后静默覆盖到错误的表。
    `flow_steps` 应是**应用 `跳过步骤` 之后**的步骤列表。
    """
    unit_steps = unit.get("steps") or []
    if len(unit_steps) != len(flow_steps):
        raise UnitError(
            f"业务单元 '{name}' 的 steps 数量 {len(unit_steps)} 与 "
            f"（应用跳过步骤后）流程的 {len(flow_steps)} 不一致")
    pairs: list[tuple[dict, dict]] = []
    for i, (flow_step, unit_step) in enumerate(zip(flow_steps, unit_steps), start=1):
        hint = _norm(unit_step.get("表单", ""))
        form = _norm(flow_step.get("form", ""))
        if hint and hint not in form:
            raise UnitError(
                f"业务单元 '{name}' 第 {i} 步表单 '{unit_step.get('表单')}' 与 "
                f"flow.yaml 的 '{flow_step.get('form')}' 不对应（steps 必须按顺序一一对应）")
        nav_cell = str(unit_step.get("nav_cell") or "").strip()
        if not nav_cell:
            raise UnitError(f"业务单元 '{name}' 第 {i} 步缺少 nav_cell")
        for key in ("nav_expect", "after_close_nav_expect"):
            expect = unit_step.get(key)
            if expect is not None and (
                    not isinstance(expect, list)
                    or not any(str(x).strip() for x in expect)):
                raise UnitError(
                    f"业务单元 '{name}' 第 {i} 步 {key} 必须是非空列表"
                    "（导航行 AA+AB+AC 标签必须包含的关键字，如 [乐华, 智屏]）")
        if unit_step.get("after_close_nav_expect") and not flow_step.get("after_close"):
            raise UnitError(
                f"业务单元 '{name}' 第 {i} 步配置了 after_close_nav_expect，"
                "但 flow.yaml 该步骤没有 after_close")
        pairs.append((flow_step, unit_step))
    return pairs


def _expect_list(value) -> list[str]:
    """归一化标签关键字列表：去空白、丢空项。"""
    if not isinstance(value, list):
        return []
    return [str(x).strip() for x in value if str(x).strip()]


# 导航表身份三列：AA(业务模块) / AB(流程步骤) / AC(业务单元)。
# 人工判断"该选哪一行"就是看这三列；批量扫描也固定读这三列。
NAV_LABEL_COLUMNS = ("AA", "AB", "AC")


def nav_rows_from_grid(values, start_row: int = 1) -> list[dict]:
    """把导航表 AA:AC 的批量读取值整理成 [{row, aa, ab, ac, label}]。

    values 是 Excel ``Range.Value2`` 的形状：二维 tuple，单行/单格时会退化成
    一维或标量。全空行直接跳过（扫描范围通常大于实际使用区域）。
    注意：合并单元格只有左上角有值，批量读取的其余格子是 None；标签匹配
    主要依赖 AB/AC（AA 可能是跨多行合并的组标题）。
    """
    if not isinstance(values, (tuple, list)):
        values = [[values]]
    elif values and not any(isinstance(v, (tuple, list)) for v in values):
        values = [values]
    rows: list[dict] = []
    for offset, raw in enumerate(values):
        if not isinstance(raw, (tuple, list)):
            raw = [raw]
        cells = ["" if v is None else str(v).strip() for v in raw]
        cells += [""] * (len(NAV_LABEL_COLUMNS) - len(cells))
        aa, ab, ac = cells[:len(NAV_LABEL_COLUMNS)]
        if not (aa or ab or ac):
            continue
        label = " | ".join(p for p in dict.fromkeys((aa, ab, ac)) if p)
        rows.append({"row": start_row + offset, "aa": aa, "ab": ab, "ac": ac,
                     "label": label})
    return rows


def match_nav_rows(rows: list[dict], keywords) -> list[dict]:
    """返回标签包含全部关键字的行。未配置关键字时不匹配任何行。"""
    expect = _expect_list(keywords)
    if not expect:
        return []
    return [r for r in rows if all(k in r["label"] for k in expect)]


def resolve_nav_row(rows: list[dict], nav_cell: str,
                    keywords) -> tuple[str, str]:
    """按标签内容决定实际选中的导航格，返回 ``(cell, 模式)``。

    模式：
    - ``fixed``     未配置关键字，沿用配置格（调用方仍会做单格标签校验）；
    - ``unique``    唯一命中：用命中行（换账号后行号整体移动也能对上）；
    - ``ambiguous`` 多行命中且配置行在其中：沿用配置格并告警（与 #1 之前行为一致）。

    零命中、或多行命中且配置行不在其中，抛 :class:`NavRowError`——此时无法
    确定哪一行是对的，必须零副作用中止，绝不猜。
    """
    m = re.match(r"([A-Za-z]+)", str(nav_cell or "").strip())
    col = m.group(1).upper() if m else "AC"
    cfg_row = int("".join(ch for ch in str(nav_cell or "") if ch.isdigit()) or 0)
    expect = _expect_list(keywords)
    if not expect:
        return nav_cell, "fixed"

    hits = match_nav_rows(rows, expect)
    if len(hits) == 1:
        return f"{col}{hits[0]['row']}", "unique"
    if cfg_row and cfg_row in [h["row"] for h in hits]:
        return nav_cell, "ambiguous"

    detail = "；".join(f"AC{r['row']}「{r['label']}」" for r in hits[:8])
    span = f"{rows[0]['row']}~{rows[-1]['row']}" if rows else "（空）"
    if not hits:
        raise NavRowError(
            f"扫描范围第 {span} 行没有包含 {expect} 的标签行（AA+AB+AC 合并文本）")
    raise NavRowError(
        f"标签包含 {expect} 的行有 {len(hits)} 个（{detail}），"
        f"配置行 {nav_cell} 不在其中，无法确定该用哪一行")


def apply_unit(flow_cfg: dict, units_cfg: dict, name: str) -> dict:
    """返回覆盖了单元 nav_cell 的 flow 配置副本（不修改入参）。

    覆盖范围：
    - 单元 `跳过步骤` 命中的流程步骤整步移除（如 TCL商用没有 HQ 渠道毛上调整）；
    - 每个 step 的 `nav_cell`（必填）；
    - step 可选的 `after_close_nav_cell` → `after_close.nav_cell`；
    - 单元可选的 `verify_nav_cell` / `verify_form_row` → `verify.*`。
    """
    unit = get_unit(units_cfg, name)
    flow = copy.deepcopy(flow_cfg)
    flow_steps, skipped = _apply_skips(flow.get("steps") or [], unit, name)
    flow["steps"] = flow_steps

    for flow_step, unit_step in _pair_steps(flow_steps, unit, name):
        flow_step["nav_cell"] = str(unit_step["nav_cell"]).strip()
        expect = _expect_list(unit_step.get("nav_expect"))
        if expect:
            # 该步导航行标签必须额外包含的关键字（如 智屏/价格）；
            # 未配置时沿用单元级 nav_expect。
            flow_step["nav_expect"] = expect
        after = flow_step.get("after_close")
        if after:
            if unit_step.get("after_close_nav_cell"):
                after["nav_cell"] = str(unit_step["after_close_nav_cell"]).strip()
            after_expect = _expect_list(unit_step.get("after_close_nav_expect"))
            if after_expect:
                after["nav_expect"] = after_expect

    verify = flow.get("verify") or {}
    if unit.get("verify_nav_cell"):
        verify["nav_cell"] = str(unit["verify_nav_cell"]).strip()
        flow["verify"] = verify
    if unit.get("verify_form_row") is not None:
        verify["form_row"] = int(unit["verify_form_row"])
        flow["verify"] = verify
    verify_expect = _expect_list(unit.get("verify_nav_expect"))
    if verify_expect:
        verify["nav_expect"] = verify_expect
        flow["verify"] = verify

    flow["_unit"] = {
        "name": name,
        "brand": str(unit.get("品牌", "")),
        "category": str(unit.get("品类", "")),
        # 导航行标签必须包含的关键字。换账号后导航表按权限重新生成，
        # 行号可能整体移动；选中后先核对标签，避免写进别的业务单元。
        "nav_expect": _expect_list(unit.get("nav_expect")),
        # 被单元配置整步移除的流程（表单 ID）；报告与日志用。
        "skipped_steps": [s.get("form") for s in skipped],
    }
    return flow


def apply_form_overrides(form_cfgs: dict[str, dict], units_cfg: dict,
                         name: str) -> dict[str, dict]:
    """把单元级 `表单覆盖` 应用到表单配置副本（不修改入参）。

    目前支持 ``跳过度量``：把某个度量从 `measures` 里移除，等价于
    「该单元这张表不编制这一块」。writer 的表头识别、写入和 CSV 列校验
    都只看 `measures`，所以移除后既不会要求 CSV 提供该列，也不会写它。
    典型场景：TCL商用 价格&销量&加权CT 表不编制 CT 成本（加权CT）。
    """
    unit = get_unit(units_cfg, name)
    overrides = unit.get("表单覆盖") or {}
    if not isinstance(overrides, dict):
        raise UnitError(
            f"业务单元 '{name}' 的 表单覆盖 必须是映射（表单ID → 覆盖项）")
    result = dict(form_cfgs)
    for key, spec in overrides.items():
        if not isinstance(spec, dict):
            raise UnitError(
                f"业务单元 '{name}' 的表单覆盖 '{key}' 必须是映射（如 {{跳过度量: [加权CT]}}）")
        form_id = str(key).strip()
        if form_id not in result:
            hint = _norm(form_id)
            matched = [fid for fid in result
                       if hint and (hint in _norm(fid) or _norm(fid) in hint)]
            if len(matched) != 1:
                raise UnitError(
                    f"业务单元 '{name}' 的表单覆盖 '{form_id}' 匹配到 "
                    f"{len(matched)} 个表单（应恰好 1 个）：{list(result)}")
            form_id = matched[0]
        cfg = copy.deepcopy(result[form_id])
        raw_skip = spec.get("跳过度量") or []
        if not isinstance(raw_skip, list):
            raise UnitError(
                f"业务单元 '{name}' 的表单覆盖 '{form_id}' 的 跳过度量 必须是列表")
        skip = [str(m).strip() for m in raw_skip if str(m).strip()]
        measures = cfg.get("measures") or {}
        unknown = [m for m in skip if m not in measures]
        if unknown:
            raise UnitError(
                f"业务单元 '{name}' 的表单覆盖 '{form_id}' 的 跳过度量 {unknown} "
                f"不在 measures 中：{list(measures)}")
        for measure in skip:
            measures.pop(measure, None)
        if skip:
            cfg["computed_measures"] = [
                m for m in (cfg.get("computed_measures") or []) if m not in skip]
        result[form_id] = cfg
    return result


def check_units(units_cfg: dict, flow_cfg: dict, only: str | None = None) -> list[str]:
    """离线校验业务单元配置，返回人类可读的问题列表（空列表 = 通过）。

    级别前缀：`错误:` 会阻止运行，`告警:` 仅提示。
    `only` 指定单元名时只校验该单元——按 `--unit` 运行时，其他单元
    （占位/未启用/缺品牌品类）的问题与本次运行无关，不应阻塞。
    """
    issues: list[str] = []
    brands = units_cfg.get("品牌") or []
    categories = units_cfg.get("品类") or []
    all_units = units(units_cfg)

    if not all_units:
        issues.append("错误: 未配置任何业务单元")
        return issues
    if only is not None:
        if only not in all_units:
            issues.append(f"错误: 未知业务单元 '{only}'，可选：{list(all_units)}")
            return issues
        targets = {only: all_units[only]}
    else:
        if not enabled_unit_names(units_cfg):
            issues.append("错误: 没有启用的业务单元（至少一个单元 启用: true）")
        targets = all_units

    for name, unit in targets.items():
        unit = unit or {}
        brand, category = str(unit.get("品牌", "")), str(unit.get("品类", ""))
        if not brand or not category:
            issues.append(f"错误: 单元 '{name}' 缺少 品牌/品类")
        else:
            if brands and brand not in brands:
                issues.append(f"告警: 单元 '{name}' 的品牌 '{brand}' 不在 品牌 列表中")
            if categories and category not in categories:
                issues.append(f"告警: 单元 '{name}' 的品类 '{category}' 不在 品类 列表中")

        if not unit.get("启用"):
            continue
        if not _expect_list(unit.get("nav_expect")):
            issues.append(
                f"错误: 单元 '{name}' 未配置 nav_expect（导航行标签关键字，如 [乐华]）。"
                "换账号后导航表行号可能整体移动，没有标签校验会存在写错业务单元的风险")
        verify_expect = unit.get("verify_nav_expect")
        if verify_expect is None:
            issues.append(
                f"告警: 单元 '{name}' 未配置 verify_nav_expect，损益核对行"
                "将只按 nav_expect 校验（可能有同一单元多行重名，建议配置）")
        elif not _expect_list(verify_expect):
            issues.append(
                f"错误: 单元 '{name}' 的 verify_nav_expect 必须是非空列表")
        overrides = unit.get("表单覆盖")
        if overrides is not None:
            if not isinstance(overrides, dict):
                issues.append(f"错误: 单元 '{name}' 的 表单覆盖 必须是映射")
            else:
                for fid, spec in overrides.items():
                    if not isinstance(spec, dict):
                        issues.append(
                            f"错误: 单元 '{name}' 的表单覆盖 '{fid}' 必须是映射")
                        continue
                    skip = spec.get("跳过度量")
                    if skip is not None and (
                            not isinstance(skip, list)
                            or not all(str(m).strip() for m in skip)):
                        issues.append(
                            f"错误: 单元 '{name}' 的表单覆盖 '{fid}' 的 跳过度量 "
                            "必须是非空字符串列表")
        try:
            steps, _skipped = _apply_skips(flow_cfg.get("steps") or [], unit, name)
            pairs = _pair_steps(steps, unit, name)
        except UnitError as exc:
            issues.append(f"错误: {exc}")
            continue
        cells = [str(us["nav_cell"]).strip() for _, us in pairs]
        dup = sorted({c for c in cells if cells.count(c) > 1})
        if dup:
            issues.append(f"错误: 单元 '{name}' 的 nav_cell 重复：{dup}")
    return issues


def _read_csv(path: Path) -> list[list[str]]:
    for encoding in ("utf-8-sig", "gbk"):
        try:
            with open(path, "r", encoding=encoding, newline="") as f:
                return list(csv.reader(f))
        except UnicodeDecodeError:
            continue
    raise OSError(f"无法识别 CSV 编码: {path}")


def csv_dimension_warnings(out_dir: Path | str, dims: tuple[str, ...] = DIM_COLUMNS) -> list[str]:
    """检查导出目录里的 CSV：品牌/品类 列存在但整列为空时给出告警。

    「品牌品类为空」意味着表单的组合键匹配会整行 skipped，提前提示比跑到
    Excel 里才发现更省时间。
    """
    warnings: list[str] = []
    for path in sorted(Path(out_dir).glob("*.csv")):
        try:
            rows = _read_csv(path)
        except OSError as exc:
            warnings.append(f"{path.name}: 读取失败（{exc}）")
            continue
        if len(rows) < 2:
            continue
        header = [str(h).strip() for h in rows[0]]
        for dim in dims:
            if dim not in header:
                continue
            idx = header.index(dim)
            values = [row[idx] for row in rows[1:] if idx < len(row)]
            if values and not any(str(v).strip() for v in values):
                warnings.append(f"{path.name}: '{dim}' 列全部为空（请检查模板表对应列）")
    return warnings
