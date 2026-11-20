"""编排层的两个行为：打开表单卡死 → 可恢复错误；落库校验 → 原地刷新不关表。

对应 2026-11-20 批量清除：反复点 AT52 打不开表单时只会在同一个 Excel 上重试，
每轮卡 5 分钟以上；以及每张表保存后都要"关闭→重开→读回"，反复开关同一张表。
"""
from __future__ import annotations

import sys
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401

from engine.excel import ExcelError, RecoverableError
from engine.orchestrator import PbfFlow


class _FakeRange:
    def __init__(self) -> None:
        self.Interior = mock.Mock(Color=16777215)

    @property
    def Value2(self):
        return None

    @Value2.setter
    def Value2(self, value):
        pass


class _FakeNavSheet:
    Name = "nav"

    def __init__(self) -> None:
        self.selected: list[str] = []

    def Activate(self) -> None:  # noqa: N802
        pass

    def Range(self, cell: str) -> _FakeRange:  # noqa: N802
        self.selected.append(cell)

        class _R(_FakeRange):
            def Select(self) -> None:  # noqa: N802
                pass

        return _R()


class _FakeWorkbook:
    Name = "form.xlsb"

    def Activate(self) -> None:  # noqa: N802
        pass


class _FakeEngine:
    def __init__(self, cells: dict | None = None) -> None:
        self.nav_sheet = _FakeNavSheet()
        self.cells = cells or {}
        self.macros: list[str] = []
        self.names: list[str] = ["nav.xlsb"]
        self.wb = _FakeWorkbook()

    @staticmethod
    def _retry(fn, attempts: int = 6, delay: float = 2.0):
        return fn()

    def get_workbook(self, pattern: str, wait_sec: float = 20.0):
        return self.wb

    def open_sheet(self, wb, sheet_name: str):
        return self.nav_sheet

    def workbook_names(self) -> list[str]:
        return list(self.names)

    def find_workbook(self, pattern: str):
        raise AssertionError("find_workbook 不应在未打开时被调用")

    def run_macro(self, macro_name: str, wb=None, *args):
        self.macros.append(macro_name)
        return None

    def read(self, sheet, cell: str):
        return self.cells.get(cell.upper())


FLOW_CFG = {
    "navigator": {
        "workbook_pattern": "nav*",
        "sheet": "nav",
        "open_button_cell": "AT52",
        "open_wait_sec": 0,
    },
    "package": {"workbook_pattern": "pkg*", "sheet": "pkg"},
    "verify_writeback": {"enabled": True, "mode": "refresh", "tolerance": 0.01},
    "recovery": {"enabled": True, "max_retries": 1},
    "verify": {},
}

FORMS = {
    "F1": {
        "workbook_pattern": "form*",
        "sheet": "Sheet1",
        "macros": {"save": "Plan_Save", "refresh": "Refresh_Sheet"},
    },
}


def _flow(engine: _FakeEngine) -> PbfFlow:
    return PbfFlow(engine, FLOW_CFG, FORMS, Path("data_2027.01"))


class OpenTimeoutTests(unittest.TestCase):
    def test_open_timeout_raises_recoverable(self) -> None:
        engine = _FakeEngine()
        flow = _flow(engine)
        with mock.patch("engine.orchestrator.time.sleep"):
            with self.assertRaises(RecoverableError):
                flow.open_via_nav("AC65", "AT52", "form", wait_sec=0, pattern="form*")


class VerifyRefreshedTests(unittest.TestCase):
    def test_refreshes_in_place_without_reopening(self) -> None:
        engine = _FakeEngine(cells={"AJ57": 0, "AN57": 0})
        flow = _flow(engine)
        with mock.patch("engine.orchestrator.time.sleep"):
            fails = flow.verify_refreshed("F1", {"AJ57": 0, "AN57": 0})
        self.assertEqual(fails, [])
        # 只跑了一次 Refresh_Sheet，没有关闭/重开工作簿
        self.assertEqual(engine.macros, ["Refresh_Sheet"])

    def test_reports_mismatch_after_refresh(self) -> None:
        engine = _FakeEngine(cells={"AJ57": 2599.0})
        flow = _flow(engine)
        with mock.patch("engine.orchestrator.time.sleep"):
            fails = flow.verify_refreshed("F1", {"AJ57": 0})
        self.assertEqual(len(fails), 1)
        self.assertIn("刷新后", fails[0])


class IndicatorGuardTests(unittest.TestCase):
    """损益核对的全 0 判定：普通填数算失败，批量清除目录（*_clear）属预期。"""

    ALL_ZERO = {"销量": {"2027.02": 0}, "净收入": {"2027.02": None},
                "税后考核净利润": {"2027.02": 0}}

    def _flow(self, data_dir: str) -> PbfFlow:
        return PbfFlow(_FakeEngine(), FLOW_CFG, FORMS, Path(data_dir))

    def test_all_zero_rejected_for_normal_dir(self) -> None:
        with self.assertRaises(ExcelError):
            self._flow("data_2027.01_2027.02")._check_indicators(self.ALL_ZERO, "2027.02")

    def test_all_zero_allowed_for_clear_dir(self) -> None:
        # 不抛异常即为通过
        self._flow("data_2027.01_2027.02_clear")._check_indicators(self.ALL_ZERO, "2027.02")


if __name__ == "__main__":
    unittest.main()
