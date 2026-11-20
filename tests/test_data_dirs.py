"""数据目录选择规则（P0-1）。"""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401  （确保仓库根在 sys.path）

from data_dirs import (data_dir_periods, data_dir_target_period, is_data_dir,
                       latest_data_dir, list_data_dirs, unrecognized_data_dirs)


class DataDirTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def _mkdir(self, *names: str) -> None:
        for name in names:
            (self.root / name).mkdir()

    def test_period_parsing(self) -> None:
        self.assertEqual(data_dir_periods("data_2027.03"), ["2027.03"])
        self.assertEqual(data_dir_periods("data_2027.01_2027.04"),
                         ["2027.01", "2027.04"])
        self.assertEqual(data_dir_target_period("data_2027.01_2027.04"), "2027.04")
        self.assertEqual(data_dir_periods("data_2027.01_2027.04_clear"), [])
        self.assertEqual(data_dir_periods("data_backup"), [])
        self.assertEqual(data_dir_periods("logs"), [])

    def test_is_data_dir_excludes_clear_and_other_dirs(self) -> None:
        self._mkdir("data_2027.03", "data_2027.03_clear", "data_2027.01_2027.04",
                    "data_tmp", "notes")
        self.assertTrue(is_data_dir(self.root / "data_2027.03"))
        self.assertFalse(is_data_dir(self.root / "data_2027.03_clear"))
        self.assertFalse(is_data_dir(self.root / "data_tmp"))
        self.assertFalse(is_data_dir(self.root / "notes"))

    def test_list_sorts_by_period_not_by_string(self) -> None:
        # 字符串序会把 data_2027.03 排在 data_2027.10_2027.12 之后；
        # 期间序应把 2027.10_2027.12 排在最后。
        self._mkdir("data_2027.03", "data_2027.10_2027.12", "data_2027.01_2027.04")
        names = [p.name for p in list_data_dirs(self.root)]
        self.assertEqual(names,
                         ["data_2027.03", "data_2027.01_2027.04", "data_2027.10_2027.12"])

    def test_latest_ignores_clear_dir(self) -> None:
        self._mkdir("data_2027.01_2027.04", "data_2027.01_2027.04_clear")
        self.assertEqual(latest_data_dir(self.root).name, "data_2027.01_2027.04")

    def test_latest_none_when_no_data_dir(self) -> None:
        self._mkdir("data_2027.01_2027.04_clear", "logs")
        self.assertIsNone(latest_data_dir(self.root))

    def test_invalid_month_dirs_are_ignored(self) -> None:
        # 月份必须 01–12：data_2027.13 / data_2027.00 不能当数据源
        self._mkdir("data_2027.11", "data_2027.13", "data_2027.00",
                    "data_2027.11_2027.13")
        self.assertEqual(data_dir_periods("data_2027.13"), [])
        self.assertEqual(data_dir_periods("data_2027.00"), [])
        self.assertEqual(data_dir_periods("data_2027.11_2027.13"), [])
        self.assertEqual(latest_data_dir(self.root).name, "data_2027.11")
        names = [p.name for p in unrecognized_data_dirs(self.root)]
        self.assertEqual(names, ["data_2027.00", "data_2027.11_2027.13", "data_2027.13"])

    def test_unrecognized_data_dirs(self) -> None:
        # data_ 前缀但不符合命名规则的目录（不补零、备份、临时目录）单独列出，
        # 供 CLI 提示；合规目录与 *_clear / 非 data_ 目录不算。
        self._mkdir("data_2027.03", "data_2027.1", "data_2027.01_bak",
                    "data_tmp", "data_2027.01_2027.04_clear", "notes")
        names = [p.name for p in unrecognized_data_dirs(self.root)]
        self.assertEqual(names, ["data_2027.01_bak", "data_2027.1", "data_tmp"])


if __name__ == "__main__":
    unittest.main()
