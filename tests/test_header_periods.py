"""期间表头归一化：文本 / 数字 / 日期 / 纯年份。"""
import importlib.util
import unittest
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def _load(name: str):
    spec = importlib.util.spec_from_file_location(name, ROOT / "engine" / f"{name}.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


hp = _load("header_periods")


class NormalizePeriodHeaderTest(unittest.TestCase):
    def test_text_period(self):
        key, prev = hp.normalize_period_header("2027.03")
        self.assertEqual(key, "2027.03")
        self.assertEqual(prev, (2027, 3))

    def test_chinese_period(self):
        key, _ = hp.normalize_period_header("2027年3月")
        self.assertEqual(key, "2027.03")

    def test_datetime_period(self):
        key, prev = hp.normalize_period_header(datetime(2027, 3, 1))
        self.assertEqual(key, "2027.03")
        self.assertEqual(prev, (2027, 3))

    def test_numeric_period(self):
        key, prev = hp.normalize_period_header(2027.03)
        self.assertEqual(key, "2027.03")
        self.assertEqual(prev, (2027, 3))

    def test_numeric_december(self):
        key, _ = hp.normalize_period_header(2026.12)
        self.assertEqual(key, "2026.12")

    def test_year_int(self):
        key, prev = hp.normalize_period_header(2026)
        self.assertEqual(key, "2026")
        self.assertIsNone(prev)

    def test_year_float(self):
        key, _ = hp.normalize_period_header(2026.0)
        self.assertEqual(key, "2026")

    def test_none(self):
        self.assertEqual(hp.normalize_period_header(None), ("", None))

    def test_nan_and_inf_do_not_raise(self):
        # Excel 异常公式可能返回 NaN/inf：不能被 int() 抛错中断整轮核对
        for value in (float("nan"), float("inf"), float("-inf")):
            self.assertIsNone(hp.month_candidate(value))
            key, prev = hp.normalize_period_header(value, (2027, 6))
            self.assertEqual(prev, (2027, 6))
            self.assertEqual(key, str(value).strip())

    def test_empty_keeps_prev(self):
        self.assertEqual(hp.normalize_period_header("", (2027, 1)), ("", (2027, 1)))

    def test_float_january_is_disambiguated_by_sequence(self):
        # 序列里第一次出现 2027.01 按 1 月；再次出现且不递增时按 10 月（与 FormScanner 一致）
        key, prev = hp.normalize_period_header(2027.01)
        self.assertEqual(key, "2027.01")
        key, prev = hp.normalize_period_header(2027.01, prev)
        self.assertEqual(key, "2027.10")

    def test_unpadded_text_january_is_disambiguated_by_sequence(self):
        # 未补零文本 "2027.1" 同样两义：序列不递增时按 10 月
        key, prev = hp.normalize_period_header("2027.1")
        self.assertEqual(key, "2027.01")
        key, _ = hp.normalize_period_header("2027.1", prev)
        self.assertEqual(key, "2027.10")

    def test_padded_text_duplicate_january_stays_january(self):
        # "2027.01" 是明确的 1 月：重复的月份块不能被改判成 10 月
        key, prev = hp.normalize_period_header("2027.01")
        self.assertEqual(key, "2027.01")
        key, _ = hp.normalize_period_header("2027.01", prev)
        self.assertEqual(key, "2027.01")

    def test_datetime_duplicate_january_stays_january(self):
        # 日期同样是明确月份，不受消歧影响
        key, prev = hp.normalize_period_header(datetime(2027, 1, 1))
        self.assertEqual(key, "2027.01")
        key, _ = hp.normalize_period_header(datetime(2027, 1, 1), prev)
        self.assertEqual(key, "2027.01")


if __name__ == "__main__":
    unittest.main()
