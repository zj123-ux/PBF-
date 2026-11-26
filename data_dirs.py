"""数据目录选择 — CLI / Web 共用的“最新数据目录”判定。

历史问题：
- 按目录名字典序取最后一个：data_*_clear 排在 data_* 之后，会把批量清除目录
  当成最新数据源（清除后直接跑 run-flow 就会把全 0 写进 BPC）；
  跨年时 data_2027.03 也会排在 data_2027.10_2027.12 之后。
- 只认 data_YYYY.MM / data_YYYY.MM_YYYY.MM，其它同前缀目录（备份、临时目录）不参与。

目录名规则（2026-11-25 起支持按业务单元分目录）：
  data_<首>[_<末>][_<业务单元>]
  例：data_2026.11_2026.12、data_2026.11_2026.12_TCL·TCL商用
  乐华和 TCL商用 同期上传时各用各的目录，互不覆盖。

单元后缀只有在调用方传入已知单元名（units）时才被认作数据目录；不传 units
时按旧规则只认无后缀目录，避免把 data_2027.01_bak 这类备份目录当数据源。

规则：排除 *_clear，按目录名中的最大期间排序（而不是字符串排序）。
"""
from __future__ import annotations

import re
from collections.abc import Iterable
from pathlib import Path

# data_2027.03 / data_2027.01_2027.04 / data_2027.01_2027.04_TCL·TCL商用
_DATA_DIR_RE = re.compile(r"^data_(\d{4}\.\d{2})(?:_(\d{4}\.\d{2}))?(?:_(.+))?$")
_PERIOD_RE = re.compile(r"^(\d{4})\.(\d{2})$")
# Windows 文件名非法字符；单元名来自配置（如 TCL·TCL商用），正常不会出现
_INVALID_NAME_CHARS = re.compile(r'[<>:"/\\|?*]')


def _valid_period(period: str) -> bool:
    """月份必须在 01–12；data_2027.13 这类目录不能当数据源。"""
    m = _PERIOD_RE.match(period or "")
    return bool(m) and 1 <= int(m.group(2)) <= 12


def _period_index(period: str) -> int:
    """'2027.03' → 月序号（便于比较先后）。与 engine.periods.period_index 等价，
    这里独立实现是为了让本模块不依赖 engine 包（测试环境无需 pywin32）。"""
    m = _PERIOD_RE.match(period or "")
    if not m or not _valid_period(period):
        raise ValueError(f"非法期间: {period!r}")
    return int(m.group(1)) * 12 + int(m.group(2)) - 1


def sanitize_unit_name(unit: str) -> str:
    """业务单元名 → 可用作目录后缀的片段（只处理文件名非法字符）。"""
    return _INVALID_NAME_CHARS.sub("_", (unit or "").strip())


def data_dir_name(periods: list[str], unit: str | None = None) -> str:
    """期间 + 业务单元 → 数据目录名。

    单期间 data_2026.11[_单元]，多期间 data_2026.11_2026.12[_单元]。
    """
    first, last = periods[0], periods[-1]
    rng = first if first == last else f"{first}_{last}"
    name = f"data_{rng}"
    unit = sanitize_unit_name(unit or "")
    return f"{name}_{unit}" if unit else name


def data_dir_unit(name: str) -> str | None:
    """目录名里的业务单元后缀；无后缀返回 None（*_clear 目录不算数据目录）。"""
    if name.endswith("_clear"):
        return None
    m = _DATA_DIR_RE.match(name)
    return m.group(3) if m else None


def data_dir_periods(name: str, units: Iterable[str] | None = None) -> list[str]:
    """目录名 → 期间列表；不是数据目录名或月份非法时返回空列表。

    units 给定时，带单元后缀的目录只有后缀在 units 里才算合规；
    units 为 None 时按旧规则只认无后缀目录（备份目录不能被误当数据源）。
    """
    if name.endswith("_clear"):
        return []
    m = _DATA_DIR_RE.match(name)
    if not m:
        return []
    periods = [p for p in m.groups()[:2] if p]
    if not all(_valid_period(p) for p in periods):
        return []
    unit = m.group(3)
    if unit is not None:
        if not units or unit not in set(units):
            return []
    return periods


def data_dir_target_period(name: str, units: Iterable[str] | None = None) -> str | None:
    """目录名对应的目标期间（多期间取最大）；非法目录名返回 None。"""
    periods = data_dir_periods(name, units)
    return max(periods, key=_period_index) if periods else None


def is_data_dir(path: Path, units: Iterable[str] | None = None) -> bool:
    """是否是上传数据目录（排除 *_clear 与不合规名字）。"""
    path = Path(path)
    return (path.is_dir()
            and not path.name.endswith("_clear")
            and bool(data_dir_periods(path.name, units)))


def list_data_dirs(root: Path, units: Iterable[str] | None = None) -> list[Path]:
    """root 下所有数据目录，按目标期间升序（同期间按名字排序）。"""
    names = list(units) if units else None
    dirs = [p for p in Path(root).iterdir() if is_data_dir(p, names)]
    return sorted(dirs, key=lambda p: (_period_index(data_dir_target_period(p.name, names)),
                                       p.name))


def latest_data_dir(root: Path, units: Iterable[str] | None = None) -> Path | None:
    """最新数据目录；没有则返回 None。"""
    dirs = list_data_dirs(root, units)
    return dirs[-1] if dirs else None


def latest_data_dir_for(root: Path, unit: str, units: Iterable[str] | None = None) -> Path | None:
    """指定业务单元的最新数据目录；该单元还没有专用目录时返回 None。"""
    unit = sanitize_unit_name(unit)
    if not unit:
        return None
    dirs = [p for p in list_data_dirs(root, units) if data_dir_unit(p.name) == unit]
    return dirs[-1] if dirs else None


def unrecognized_data_dirs(root: Path, units: Iterable[str] | None = None) -> list[Path]:
    """data_ 前缀但不符合命名规则的目录（旧版不补零的 data_2027.1、备份、临时目录等）。

    这些目录不参与数据源选择与“最新目录”判定；列出来供 CLI 提示，
    避免用户以为数据丢了。*_clear 是批量清除的约定目录名，不算在内。
    """
    names = list(units) if units else None
    out: list[Path] = []
    for p in Path(root).iterdir():
        if (p.is_dir() and p.name.lower().startswith("data_")
                and not p.name.endswith("_clear")
                and not data_dir_periods(p.name, names)):
            out.append(p)
    return sorted(out, key=lambda p: p.name)
