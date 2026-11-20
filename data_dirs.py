"""数据目录选择 — CLI / Web 共用的“最新数据目录”判定。

历史问题：
- 按目录名字典序取最后一个：data_*_clear 排在 data_* 之后，会把批量清除目录
  当成最新数据源（清除后直接跑 run-flow 就会把全 0 写进 BPC）；
  跨年时 data_2027.03 也会排在 data_2027.10_2027.12 之后。
- 只认 data_YYYY.MM / data_YYYY.MM_YYYY.MM，其它同前缀目录（备份、临时目录）不参与。

规则：排除 *_clear，按目录名中的最大期间排序（而不是字符串排序）。
"""
from __future__ import annotations

import re
from pathlib import Path

# data_2027.03 / data_2027.01_2027.04
_DATA_DIR_RE = re.compile(r"^data_(\d{4}\.\d{2})(?:_(\d{4}\.\d{2}))?$")
_PERIOD_RE = re.compile(r"^(\d{4})\.(\d{2})$")


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


def data_dir_periods(name: str) -> list[str]:
    """目录名 → 期间列表；不是数据目录名或月份非法时返回空列表。"""
    m = _DATA_DIR_RE.match(name)
    if not m:
        return []
    periods = [p for p in m.groups() if p]
    return periods if all(_valid_period(p) for p in periods) else []


def data_dir_target_period(name: str) -> str | None:
    """目录名对应的目标期间（多期间取最大）；非法目录名返回 None。"""
    periods = data_dir_periods(name)
    return max(periods, key=_period_index) if periods else None


def is_data_dir(path: Path) -> bool:
    """是否是上传数据目录（排除 *_clear 与不合规名字）。"""
    path = Path(path)
    return (path.is_dir()
            and not path.name.endswith("_clear")
            and bool(data_dir_periods(path.name)))


def list_data_dirs(root: Path) -> list[Path]:
    """root 下所有数据目录，按目标期间升序（同期间按名字排序）。"""
    dirs = [p for p in Path(root).iterdir() if is_data_dir(p)]
    return sorted(dirs, key=lambda p: (_period_index(data_dir_target_period(p.name)), p.name))


def latest_data_dir(root: Path) -> Path | None:
    """最新数据目录；没有则返回 None。"""
    dirs = list_data_dirs(root)
    return dirs[-1] if dirs else None


def unrecognized_data_dirs(root: Path) -> list[Path]:
    """data_ 前缀但不符合命名规则的目录（旧版不补零的 data_2027.1、备份、临时目录等）。

    这些目录不参与数据源选择与“最新目录”判定；列出来供 CLI 提示，
    避免用户以为数据丢了。*_clear 是批量清除的约定目录名，不算在内。
    """
    out: list[Path] = []
    for p in Path(root).iterdir():
        if (p.is_dir() and p.name.lower().startswith("data_")
                and not p.name.endswith("_clear")
                and not data_dir_periods(p.name)):
            out.append(p)
    return sorted(out, key=lambda p: p.name)
