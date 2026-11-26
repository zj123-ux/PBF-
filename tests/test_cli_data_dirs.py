"""CLI 数据目录选择：按业务单元分目录，避免 TCL商用 / 乐华 互相覆盖。

覆盖 data_dir_for_unit（run-flow 默认目录）与 clear_source_dirs（批量清除来源）
的三种情形：有专用目录、只有唯一的旧版目录、旧版目录多份且无法判断。
"""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401  （确保仓库根在 sys.path、补齐 COM 占位模块）


def _units_cfg() -> dict:
    return {"units": {
        "乐华·电视": {"启用": True, "品牌": "乐华", "品类": "电视"},
        "TCL·TCL商用": {"启用": True, "品牌": "TCL", "品类": "TCL商用"},
    }}


class CliDataDirTests(unittest.TestCase):
    def setUp(self) -> None:
        import cli
        self.cli = cli
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self._old_root = cli.ROOT
        self._old_units = cli.load_units_config
        cli.ROOT = self.root
        cli.load_units_config = _units_cfg

    def tearDown(self) -> None:
        self.cli.ROOT = self._old_root
        self.cli.load_units_config = self._old_units
        self._tmp.cleanup()

    def _mkdir(self, *names: str) -> None:
        for name in names:
            (self.root / name).mkdir()

    def test_export_and_clear_dir_names_carry_unit(self) -> None:
        self.assertEqual(self.cli.default_export_dir(["2026.11", "2026.12"], "乐华·电视").name,
                         "data_2026.11_2026.12_乐华·电视")
        self.assertEqual(self.cli.default_export_dir(["2026.11"]).name, "data_2026.11")
        self.assertEqual(
            self.cli.clear_out_dir(["2026.11", "2026.12"], "TCL·TCL商用").name,
            "data_2026.11_2026.12_TCL·TCL商用_clear")

    def test_run_flow_prefers_unit_dir(self) -> None:
        self._mkdir("data_2026.11_2026.12_乐华·电视",
                    "data_2026.11_2026.12_TCL·TCL商用", "data_2026.11_2026.12")
        self.assertEqual(self.cli.data_dir_for_unit("TCL·TCL商用").name,
                         "data_2026.11_2026.12_TCL·TCL商用")
        self.assertEqual(self.cli.data_dir_for_unit("乐华·电视").name,
                         "data_2026.11_2026.12_乐华·电视")

    def test_run_flow_falls_back_to_single_legacy_dir(self) -> None:
        # 还没有单元专用目录时，唯一的旧版目录可回退（当前 2026.11 旧数据就是这种）
        self._mkdir("data_2026.11_2026.12")
        self.assertEqual(self.cli.data_dir_for_unit("TCL·TCL商用").name,
                         "data_2026.11_2026.12")

    def test_run_flow_refuses_to_guess_between_legacy_dirs(self) -> None:
        self._mkdir("data_2026.11", "data_2026.12")
        with self.assertRaises(SystemExit) as ctx:
            self.cli.data_dir_for_unit("TCL·TCL商用")
        self.assertIn("多个旧版目录", str(ctx.exception))

    def test_run_flow_errors_when_no_dir_at_all(self) -> None:
        with self.assertRaises(SystemExit):
            self.cli.data_dir_for_unit("TCL·TCL商用")

    def test_clear_sources_filtered_by_unit(self) -> None:
        self._mkdir("data_2026.11_2026.12_乐华·电视",
                    "data_2026.11_2026.12_TCL·TCL商用",
                    "data_2026.11_2026.12")
        names = [p.name for p in self.cli.clear_source_dirs("TCL·TCL商用")]
        self.assertEqual(names, ["data_2026.11_2026.12_TCL·TCL商用"])

    def test_clear_sources_fall_back_to_single_legacy(self) -> None:
        self._mkdir("data_2026.11_2026.12")
        names = [p.name for p in self.cli.clear_source_dirs("乐华·电视")]
        self.assertEqual(names, ["data_2026.11_2026.12"])

    def test_clear_sources_no_unit_returns_all(self) -> None:
        self._mkdir("data_2026.11_2026.12_乐华·电视",
                    "data_2026.11_2026.12_TCL·TCL商用")
        self.assertEqual(len(self.cli.clear_source_dirs("")), 2)


if __name__ == "__main__":
    unittest.main()
