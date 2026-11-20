"""期间表头解析（纯逻辑，不依赖 Excel COM，便于单元测试）。

AO 表头单元格可能是文本（"2027.03"）、数字（2027.03）、日期（2027-03-01）
或纯年份（"2026" / 2026）。统一归一化成后续匹配用的键。
"""
from __future__ import annotations

import math
import re
from datetime import date, datetime
from typing import Any

# 字符串形式的期间: "2026.11" / "2026-11" / "2026年11月"
MONTH_STR = re.compile(r"^(\d{4})\s*[.\-/年]\s*(\d{1,2})\s*月?$")


def month_candidate(v: Any) -> tuple[int, int] | None:
    """把一个单元格值解析成 (年, 月候选)。返回 None 表示不是期间。"""
    if v is None:
        return None
    if isinstance(v, (datetime, date)):
        return (v.year, v.month)
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        f = float(v)
        # NaN/inf（公式异常、加载中的单元格）不能 int()，直接视为非期间
        if not math.isfinite(f):
            return None
        # 纯年份整数（如 2027）不是期间
        if f == int(f) and 1900 < f < 2100:
            return None
        y = int(f)
        if not (2000 <= y <= 2099):
            return None
        m = int(round((f - y) * 100))
        if 1 <= m <= 12:
            return (y, m)
        return None
    s = str(v).strip()
    mt = MONTH_STR.match(s)
    if mt:
        y, m = int(mt.group(1)), int(mt.group(2))
        if 2000 <= y <= 2099 and 1 <= m <= 12:
            return (y, m)
    return None


def _ambiguous_month(v: Any) -> bool:
    """值是否属于 1 月/10 月两义的表示。

    两义的只有：浮点（2027.1）与未补零文本（"2027.1" / "2027-1"）。
    明确写法——零填充文本 "2027.01"、"2027年1月"、日期——不存在歧义，
    损益表头出现重复的 1 月块时不能被改判成 10 月。
    """
    if isinstance(v, bool):
        return False
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, str):
        return bool(re.match(r"^\d{4}\s*[.\-/]\s*\d$", v.strip()))
    return False


def normalize_period_header(v: Any, prev: tuple[int, int] | None = None
                            ) -> tuple[str, tuple[int, int] | None]:
    """单元格值 → (表头键, 更新后的 prev)。

    - 日期 / 数字 / 文本月份 → 'YYYY.MM'
    - 纯年份 → 'YYYY'
    - 其它 → 原文去空白

    prev 用于消歧：浮点 2027.1 既可能是 1 月也可能是 10 月，若按 1 月会
    相对上一个期间倒退，则判为 10 月（与 FormScanner 的规则一致）。
    消歧只作用于两义表示（浮点 / 未补零文本）；"2027.01"、日期是明确的
    1 月，重复的 1 月块照原样返回。
    """
    if v is None:
        return "", prev
    cand = month_candidate(v)
    if cand:
        y, m = cand
        if (m == 1 and prev is not None and (y, 1) <= prev
                and _ambiguous_month(v)):
            m = 10
        return f"{y}.{m:02d}", (y, m)
    if isinstance(v, bool):
        return str(v).strip(), prev
    if isinstance(v, (int, float)):
        f = float(v)
        if f.is_integer() and 1900 < f < 2100:
            return str(int(f)), prev
    return str(v).strip(), prev
