"""写入后 AO 变更跟踪校验：HasChangedPlanData=False 时不得调用保存。

对应 2026-11-20 批量清除故障：603 格写入后 Plan_Save 弹「没有要保存的数据」。
根因是 PBF 的 VBA 回调临时关闭 Application.EnableEvents，写入落在该窗口里，
Excel 不触发 SheetChange，AO 收不到变更。这里用假 ExcelEngine 锁定行为：

1. 写入后 HCPD=True                → 正常保存
2. 写入后 HCPD=False，重写后 True   → 保存一次（重写救回）
3. 写入后 HCPD=False，重写仍 False  → 判失败，绝不调用保存
"""
from __future__ import annotations

import sys
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401

from engine.writer import FormWriter, PreflightResult


class _FakeInterior:
    def __init__(self, color: int = 16777215) -> None:
        self.Color = color


class _FakeRange:
    def __init__(self, sheet: "_FakeSheet", cell: str) -> None:
        self._sheet = sheet
        self._cell = cell.upper()
        self.Interior = _FakeInterior()

    @property
    def Value2(self):
        return self._sheet.cells.get(self._cell)

    @Value2.setter
    def Value2(self, value):
        self._sheet.cells[self._cell] = value
        self._sheet.writes.append(self._cell)


class _FakeSheet:
    Name = "Sheet1"

    def __init__(self) -> None:
        self.cells = {"AJ57": 2599.0, "AN57": 597.0}
        self.writes: list[str] = []

    def Range(self, cell: str) -> _FakeRange:  # noqa: N802 - 模仿 COM 接口
        return _FakeRange(self, cell)


class _FakeEngine:
    """只实现 FormWriter 用到的方法；HCPD 值按脚本依次返回。"""

    def __init__(self, dirty_sequence: list) -> None:
        self.sheet = _FakeSheet()
        self.wb = mock.Mock(Name="fake.xlsb")
        self.dirty_sequence = list(dirty_sequence)
        self.saved = 0
        self.last_dialogs: list[str] = []
        self.events_on = True

    # workbook / sheet
    def find_workbook(self, pattern: str):
        return self.wb

    def get_workbook(self, pattern: str, wait_sec: float = 20.0):
        return self.wb

    def open_sheet(self, wb, sheet_name: str):
        return self.sheet

    # cells
    def read(self, sheet, cell: str):
        return sheet.cells.get(cell.upper())

    def write_batch(self, sheet, updates: dict) -> None:
        for cell, value in updates.items():
            sheet.Range(cell).Value2 = value

    # AO 属性
    def run_macro(self, macro_name: str, wb=None, *args):
        assert macro_name == "SAPGetProperty"
        return self.dirty_sequence.pop(0) if self.dirty_sequence else False

    # 事件
    def events_enabled(self):
        return self.events_on

    def ensure_events(self) -> bool:
        was_off = not self.events_on
        self.events_on = True
        return was_off

    def save(self, wb, macro_hint=None):
        self.saved += 1


CFG = {
    "workbook_pattern": "fake*",
    "sheet": "Sheet1",
    "mode_cell": None,
    "macros": {"save": "Plan_Save"},
}


def _preflight() -> PreflightResult:
    return PreflightResult(
        ok=True,
        matched_rows=[
            {"item": "A", "month": "2027.01", "cell": "AJ57",
             "old_value": 2599.0, "new_value": 0},
            {"item": "A", "month": "2027.01", "cell": "AN57",
             "old_value": 597.0, "new_value": 0},
        ],
    )


class WriterTrackingTests(unittest.TestCase):
    def _run(self, dirty_sequence: list):
        engine = _FakeEngine(dirty_sequence)
        writer = FormWriter(engine, CFG)
        with mock.patch("engine.writer.time.sleep"):
            result = writer.write_and_save(_preflight(), verify_after=False)
        return engine, result

    def test_saves_when_ao_tracks_changes(self) -> None:
        # before=False, after=True
        engine, result = self._run([False, True])
        self.assertTrue(result["saved"])
        self.assertEqual(engine.saved, 1)
        self.assertEqual(engine.sheet.writes.count("AJ57"), 1)

    def test_rewrites_once_when_ao_missed_changes(self) -> None:
        # before=False, after写=False, 重写后=True
        engine, result = self._run([False, False, True])
        self.assertTrue(result["saved"])
        self.assertEqual(engine.saved, 1)
        self.assertEqual(engine.sheet.writes.count("AJ57"), 2)

    def test_fails_without_save_when_changes_never_tracked(self) -> None:
        # before=False, after写=False, 重写后仍=False
        engine, result = self._run([False, False, False])
        self.assertFalse(result["saved"])
        self.assertEqual(engine.saved, 0)
        self.assertIn("HasChangedPlanData=False", result["failures"][0])
        # 重写发生时单元格已被写入两次
        self.assertEqual(engine.sheet.writes.count("AJ57"), 2)


if __name__ == "__main__":
    unittest.main()
