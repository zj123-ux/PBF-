"""填表扫描与损益核对使用同一套 1 月/10 月消歧规则。"""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401  （确保仓库根在 sys.path，并 stub win32com）

from engine.scanner import FormScanner


class MonthHeaderTest(unittest.TestCase):
    def _scan(self, row: list) -> dict[str, str]:
        _, cols = FormScanner(None)._find_month_header([row])
        return cols

    def test_padded_january_is_not_reclassified(self) -> None:
        # "2027.01" 是明确的 1 月：即使它出现在 2027.06 之后，也不能改判 10 月
        cols = self._scan(["组织", "2027.04", "2027.06", "2027.01"])
        self.assertIn("2027.01", cols)
        self.assertNotIn("2027.10", cols)

    def test_ambiguous_float_january_becomes_october(self) -> None:
        # 浮点 2027.1 两义：序列里倒退时按 10 月（与 header_periods 一致）
        cols = self._scan(["组织", "2027.04", "2027.06", 2027.1])
        self.assertIn("2027.10", cols)
        self.assertNotIn("2027.01", cols)


if __name__ == "__main__":
    unittest.main()
