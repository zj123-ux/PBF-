"""FormScanner — 扫描已打开的 PBF 编制表结构，推导 YAML 配置片段。

自动识别:
- month_columns: 表头行所有期间列（全期间，2026.01 ~ 2027.12）
- item_column:   科目名称所在列（文本最密集的列）
- input_rows:    数据行范围
- mode_cell:     值为"编制"/"查询"的模式单元格
- version_cell:  含"滚动预测"/"预算"字样的版本单元格

用法: FormScanner(engine).scan(wb, sheet)
"""
from __future__ import annotations

import re
from datetime import date, datetime
from typing import Any

from .excel import ExcelEngine
from .header_periods import MONTH_STR as _MONTH_STR
from .header_periods import _ambiguous_month
from .header_periods import month_candidate as _month_candidate


def col_letter(n: int) -> str:
    """列号 → 字母 (1→A, 29→AC, 62→BJ)."""
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def _is_text(v: Any) -> bool:
    """判断是否为科目名这类文本（非纯数字/非日期）。"""
    if v is None:
        return False
    if isinstance(v, (int, float, bool, datetime, date)):
        return False
    s = str(v).strip()
    if not s:
        return False
    if _MONTH_STR.match(s):
        return False
    # 纯数字或纯符号不算科目
    if re.fullmatch(r"[\d.,%\-\s]+", s):
        return False
    return True


class FormScanner:
    def __init__(self, engine: ExcelEngine, max_row: int = 500, max_col: int = 80) -> None:
        self.engine = engine
        self.max_row = max_row
        self.max_col = max_col

    # ---------- 底层: 一次性读整块区域 ----------

    def _read_block(self, sheet) -> list[list[Any]]:
        ur = sheet.UsedRange
        last_row = min(ur.Row + ur.Rows.Count - 1, self.max_row)
        last_col = min(ur.Column + ur.Columns.Count - 1, self.max_col)
        last_row = max(last_row, 1)
        last_col = max(last_col, 1)
        rng = sheet.Range(sheet.Cells(1, 1), sheet.Cells(last_row, last_col))
        val = rng.Value2
        if not isinstance(val, tuple):
            return [[val]]
        return [list(r) if isinstance(r, tuple) else [r] for r in val]

    # ---------- 表头: 找月份行 ----------

    def _find_month_header(self, block: list[list[Any]]) -> tuple[int, dict[str, str]]:
        """返回 (表头行号(1-based), {期间: 列字母})。选期间数最多的行。"""
        best_row, best_map = 0, {}
        for i, row in enumerate(block, start=1):
            raw: list[tuple[int, Any, int, int]] = []
            for j, v in enumerate(row, start=1):
                cand = _month_candidate(v)
                if cand:
                    raw.append((j, v, cand[0], cand[1]))
            if len(raw) < 3:
                continue
            # 1 月与 10 月存在浮点歧义（2026.1 可能是 1 月或 10 月），
            # 用"期间必须递增"消除：若按 1 月会倒退，则判为 10 月。
            # 与 header_periods 保持一致：只对两义表示消歧，零填充的
            # "2027.01"、日期这类明确写法不能被改判。
            resolved: dict[str, str] = {}
            prev = (0, 0)
            for col, v, y, m in raw:
                chosen = m
                if m == 1 and (y, 1) <= prev and _ambiguous_month(v):
                    chosen = 10
                resolved[f"{y}.{chosen:02d}"] = col_letter(col)
                prev = (y, chosen)
            if len(resolved) > len(best_map):
                best_row, best_map = i, resolved
        return best_row, best_map

    # ---------- 科目列 ----------

    def _find_item_column(self, block: list[list[Any]], header_row: int,
                          month_cols: dict[str, str]) -> tuple[str, dict[str, int]]:
        """返回 (科目列字母, {列字母: 文本数量})。"""
        month_letters = set(month_cols.values())
        ncols = max((len(r) for r in block), default=0)
        counts: dict[str, int] = {}
        for j in range(1, ncols + 1):
            letter = col_letter(j)
            if letter in month_letters:
                continue
            cnt = 0
            for row in block[header_row:]:
                if j - 1 < len(row) and _is_text(row[j - 1]):
                    cnt += 1
            if cnt:
                counts[letter] = cnt
        if not counts:
            return "", {}
        # 文本最多的列即科目列
        best = max(counts, key=lambda k: counts[k])
        return best, dict(sorted(counts.items(), key=lambda kv: -kv[1])[:6])

    # ---------- 数据行范围 ----------

    def _find_input_rows(self, block: list[list[Any]], header_row: int,
                         item_col: str) -> tuple[int, int]:
        if not item_col:
            return (header_row + 1, len(block))
        j = 0
        for k in range(1, self.max_col + 1):
            if col_letter(k) == item_col:
                j = k
                break
        first = last = 0
        for i in range(header_row + 1, len(block) + 1):
            row = block[i - 1]
            if j - 1 < len(row) and _is_text(row[j - 1]):
                if not first:
                    first = i
                last = i
        return (first or header_row + 1, last or len(block))

    # ---------- 模式区 / 版本区 ----------

    def _find_mode_cells(self, block: list[list[Any]]) -> dict[str, Any]:
        mode_cells = []
        label_cell = None
        for i, row in enumerate(block[:60], start=1):
            for j, v in enumerate(row[:20], start=1):
                s = str(v).strip() if v is not None else ""
                if s in ("编制", "查询"):
                    mode_cells.append({"cell": f"{col_letter(j)}{i}", "value": s})
                elif "当前模式" in s or ("模式" in s and len(s) <= 8):
                    label_cell = f"{col_letter(j)}{i}"
        return {"mode_cells": mode_cells, "mode_label_cell": label_cell}

    def _find_version_cell(self, block: list[list[Any]]) -> str | None:
        for i, row in enumerate(block[:60], start=1):
            for j, v in enumerate(row[:20], start=1):
                s = str(v).strip() if v is not None else ""
                if ("滚动预测" in s or "年度预算" in s) and len(s) <= 30:
                    return f"{col_letter(j)}{i}"
        return None

    # ---------- 主入口 ----------

    def scan(self, wb, sheet) -> dict[str, Any]:
        block = self._read_block(sheet)
        header_row, month_cols = self._find_month_header(block)
        item_col, item_candidates = self._find_item_column(block, header_row, month_cols)
        row_start, row_end = self._find_input_rows(block, header_row, item_col)
        mode_info = self._find_mode_cells(block)
        version_cell = self._find_version_cell(block)

        # 排除期间的汇总列（如 AR57=2027 这种纯年份列已在解析时过滤）
        months = dict(sorted(month_cols.items()))

        return {
            "workbook": wb.Name,
            "sheet": sheet.Name,
            "header_row": header_row,
            "month_columns": months,
            "item_column": item_col,
            "item_column_candidates": item_candidates,
            "input_rows": [row_start, row_end],
            "mode_cells": mode_info["mode_cells"],
            "mode_label_cell": mode_info["mode_label_cell"],
            "version_cell": version_cell,
        }
