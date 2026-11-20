"""批量清除 — 把既有 data_* 目录里填过的数字清成 0，期间可自由选择。

「填过的项」= 所有既有 data_* 目录里出现过的维度键并集（不关心期间列，只看维度）。
生成全 0 的上传 CSV，供 run-flow 复用填 0。
"""
from __future__ import annotations

import csv
from pathlib import Path

import pandas as pd

# 5 张上传 CSV 的结构（与 engine/template_export.py 的输出字段一致）
# (布局, 维度列, 度量列) — 长表期间在行字段"期间"，宽表期间是列名
FORM_SCHEMA: dict[str, tuple[str, list[str], list[str]]] = {
    "价格销量CT.csv": ("长表", ["组织", "渠道", "品牌", "产品"],
                       ["销售单价（含税）", "销量(出货口径)", "加权CT"]),
    "其他成本.csv": ("宽表", ["item"], []),
    "客户费用.csv": ("宽表", ["渠道", "科目"], []),
    "HQ毛上调整.csv": ("宽表", ["科目", "区域"], []),
    "费用及损益其他项.csv": ("宽表", ["组织", "科目", "产品", "区域", "渠道", "品牌"], []),
}


def _collect_keys(data_dirs: list[Path], csv_name: str) -> list[dict]:
    """读取所有 data_* 目录里该 CSV，返回去重后的维度键列表（保持首见顺序）。"""
    _, dims, _ = FORM_SCHEMA[csv_name]
    seen: dict[tuple, dict] = {}
    for d in data_dirs:
        path = Path(d) / csv_name
        if not path.exists():
            continue
        try:
            df = pd.read_csv(path)
        except Exception:
            continue
        for _, row in df.iterrows():
            vals = []
            for c in dims:
                v = row.get(c)
                if pd.isna(v):
                    v = ""
                vals.append("" if v is None else str(v).strip())
            key = tuple(vals)
            if key not in seen and any(key):
                seen[key] = {c: vals[i] for i, c in enumerate(dims)}
    return list(seen.values())


def clear_data(data_dirs: list[Path], periods: list[str], out_dir: Path,
               require_all: bool = False) -> dict[str, int]:
    """对所选期间生成全 0 CSV，返回 {文件名: 行数}。

    输出目录里上一轮的 CSV 会先删除（本次生成 0 行的表不应残留旧文件）；
    若输出目录与任一数据来源目录相同则直接拒绝，避免覆盖原始上传数据。

    键收集在任何文件被改动之前完成：数据来源里找不到已填项时直接抛错，
    此时输出目录还保持原样，不会被清空一半。

    require_all=True（配合 `clear --run`）时，任一表单在来源里没有可清除
    的行就中止：run-flow 需要 5 份非空数据文件，先生成残缺目录再报成功
    只会让后面的流程必然失败。
    """
    out_resolved = Path(out_dir).resolve()
    for d in data_dirs:
        if Path(d).resolve() == out_resolved:
            raise ValueError(f"输出目录不能与数据来源目录相同：{d}（会覆盖原始上传数据）")

    collected: dict[str, list[dict]] = {
        csv_name: _collect_keys(data_dirs, csv_name) for csv_name in FORM_SCHEMA}
    if not any(collected.values()):
        raise ValueError(
            "未在数据来源目录中找到可清除的已填项；"
            "请检查 --data-dirs（默认取全部 data_* 目录）")
    missing = [name for name, keys in collected.items() if not keys]
    if require_all and missing:
        raise ValueError(
            f"以下表单在数据来源中没有可清除的行，无法生成完整的清除目录："
            f"{'、'.join(missing)}；请检查来源数据，或去掉 --run 只做部分清除")

    out_dir.mkdir(parents=True, exist_ok=True)
    for name in FORM_SCHEMA:
        try:
            (out_dir / name).unlink(missing_ok=True)
        except OSError:
            pass

    result: dict[str, int] = {}
    for csv_name, (layout, dims, measures) in FORM_SCHEMA.items():
        keys = collected[csv_name]
        rows: list[dict] = []
        if layout == "长表":
            for k in keys:
                for p in periods:
                    row = dict(k)
                    row["期间"] = p
                    for m in measures:
                        row[m] = 0
                    rows.append(row)
        else:
            for k in keys:
                row = dict(k)
                for p in periods:
                    row[p] = 0
                rows.append(row)
        if rows:
            with open(out_dir / csv_name, "w", encoding="utf-8-sig", newline="") as f:
                w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
                w.writeheader()
                w.writerows(rows)
            result[csv_name] = len(rows)
    return result
