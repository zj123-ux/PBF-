"""模板导出：0 行也要写表头，且必须清掉上一轮残留的 CSV（P1-8）。"""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401

from openpyxl import Workbook

from engine.template_export import EXPORT_FILES, EmptyExportError, export_template


def _make_template(path: Path) -> None:
    """最小可用的《滚动预测模板表》结构（periods 只放 2027.01）。"""
    wb = Workbook()

    ws = wb.active
    ws.title = "价格&销量&加权CT"
    ws.append(["", "", "", "", "单价", "销量", "CT"])
    ws.append(["组织", "渠道", "品牌", "产品", "2027.01", "2027.01", "2027.01"])
    ws.append(["TCL", "电商", "乐华", "产品A", 100, 10, 1000])

    ws = wb.create_sheet("客户费用")
    ws.append([""])
    ws.append(["渠道", "科目", "2027.01"])
    ws.append(["电商", "客户费用-返利", 5])

    ws = wb.create_sheet("HQ毛上调整")
    ws.append(["", "", "2027.01"])
    ws.append(["", ""])
    ws.append(["", "中国营销本部"])
    ws.append(["科目A", "", 9])

    ws = wb.create_sheet("其他成本")
    ws.append([""])
    ws.append(["", "", "item", "2027.01"])
    ws.append(["", "", "其他成本A", 7])

    ws = wb.create_sheet("变动费用")
    ws.append([""])
    ws.append(["组织", "科目", "产品", "区域", "渠道", "品牌", "2027.01"])
    ws.append(["TCL", "费用A", "产品A", "中国", "电商", "乐华", 11])

    wb.save(path)


class ExportTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        base = Path(self._tmp.name)
        self.template = base / "模板.xlsx"
        self.out = base / "data_2027.01"
        self.out.mkdir()
        _make_template(self.template)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def _write_stale(self) -> None:
        for name in EXPORT_FILES:
            (self.out / name).write_text("OLD\n", encoding="utf-8")

    def test_exports_all_five_files_with_rows(self) -> None:
        self._write_stale()
        rows = export_template(self.template, ["2027.01"], self.out)
        self.assertEqual(set(rows), set(EXPORT_FILES))
        self.assertEqual(rows["价格销量CT.csv"], 1)
        self.assertEqual(rows["客户费用.csv"], 1)
        self.assertEqual(rows["HQ毛上调整.csv"], 1)
        self.assertEqual(rows["其他成本.csv"], 1)
        self.assertEqual(rows["费用及损益其他项.csv"], 1)
        for name in EXPORT_FILES:
            text = (self.out / name).read_text(encoding="utf-8-sig")
            self.assertNotIn("OLD", text, f"{name} 仍残留旧内容")
            self.assertTrue(text.strip(), f"{name} 为空")

    def test_zero_rows_still_writes_headers_and_removes_stale(self) -> None:
        self._write_stale()
        rows = export_template(self.template, ["2027.02"], self.out)
        self.assertEqual(sum(rows.values()), 0)
        # 长表（价格销量）期间在行字段里，表头只有“期间”；其余宽表期间是列名
        wide = {"客户费用.csv", "其他成本.csv", "HQ毛上调整.csv", "费用及损益其他项.csv"}
        for name in EXPORT_FILES:
            self.assertTrue((self.out / name).exists(), f"{name} 未生成")
            text = (self.out / name).read_text(encoding="utf-8-sig")
            self.assertNotIn("OLD", text, f"{name} 仍残留旧内容")
            header = text.splitlines()[0]
            if name in wide:
                self.assertIn("2027.02", header, f"{name} 表头缺期间列")
            else:
                self.assertIn("期间", header, f"{name} 表头缺期间字段")

    def test_failed_export_keeps_previous_output(self) -> None:
        # 导出先写临时目录，全部成功后才替换：模板读不了时不破坏上一轮 CSV
        self._write_stale()
        with self.assertRaises(FileNotFoundError):
            export_template(self.out.parent / "不存在的模板.xlsx", ["2027.01"], self.out)
        for name in EXPORT_FILES:
            self.assertEqual((self.out / name).read_text(encoding="utf-8"), "OLD\n",
                             f"{name} 在导出失败后被清掉了")

    def test_min_rows_zero_keeps_previous_output(self) -> None:
        # Web 上传 0 行模板：必须在替换目录之前失败，旧数据保持原样
        self._write_stale()
        with self.assertRaises(EmptyExportError):
            export_template(self.template, ["2027.02"], self.out, min_rows=1)
        for name in EXPORT_FILES:
            self.assertEqual((self.out / name).read_text(encoding="utf-8"), "OLD\n",
                             f"{name} 在 0 行导出后被清掉了")

    def test_staging_and_backup_dirs_are_cleaned_up(self) -> None:
        self._write_stale()
        export_template(self.template, ["2027.01"], self.out)
        leftovers = [p.name for p in self.out.parent.iterdir()
                     if p.name.startswith(f".{self.out.name}.")]
        self.assertEqual(leftovers, [])


if __name__ == "__main__":
    unittest.main()
