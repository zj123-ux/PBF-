"""Web 上传：0 行模板不得先覆盖上一份正确数据再报错。"""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401  （确保仓库根在 sys.path）
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from engine.template_export import EXPORT_FILES
from test_template_export import _make_template

try:
    import webapp.app as wa
    HAS_FLASK = True
except ImportError:                     # 轻量环境未装 flask：跳过，CI 装了会跑
    wa = None
    HAS_FLASK = False


@unittest.skipUnless(HAS_FLASK, "需要 flask（pip install -r requirements.txt）")
class UploadTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self._old_root = wa.APP_ROOT
        self._old_web_run = wa.WEB_RUN_DIR
        wa.APP_ROOT = self.root
        wa.WEB_RUN_DIR = self.root / "logs" / "web"
        # 隔离真实 business_units.yaml：默认按“未配置业务单元”走旧版目录名
        self._old_units_file = wa.BUSINESS_UNITS_FILE
        wa.BUSINESS_UNITS_FILE = self.root / "config" / "business_units.yaml"
        self.template = self.root / "模板.xlsx"
        _make_template(self.template)
        self.data_dir = self.root / "data_2027.02"
        self.data_dir.mkdir()
        for name in EXPORT_FILES:
            (self.data_dir / name).write_text("OLD\n", encoding="utf-8")

    def tearDown(self) -> None:
        wa.APP_ROOT = self._old_root
        wa.WEB_RUN_DIR = self._old_web_run
        wa.BUSINESS_UNITS_FILE = self._old_units_file
        self._tmp.cleanup()

    def _upload(self, period: str, unit: str = ""):
        client = wa.app.test_client()
        with open(self.template, "rb") as f:
            return client.post(
                "/api/upload",
                data={"file": (f, "模板.xlsx"), "period": period, "unit": unit},
                content_type="multipart/form-data",
            )

    def _write_units_config(self, *names: str) -> None:
        lines = ["units:"]
        for name in names:
            lines += [f"  {name}:", "    启用: true",
                      f"    品牌: {name.split('·')[0]}",
                      f"    品类: {name.split('·')[1]}"]
        path = wa.BUSINESS_UNITS_FILE
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("\n".join(lines) + "\n", encoding="utf-8")

    def test_zero_row_upload_returns_400_and_keeps_previous_data(self) -> None:
        # 模板只有 2027.01 的数据，按 2027.02 导出为 0 行
        resp = self._upload("2027.02")
        self.assertEqual(resp.status_code, 400)
        self.assertIn("行数为 0", resp.get_json()["error"])
        for name in EXPORT_FILES:
            self.assertEqual((self.data_dir / name).read_text(encoding="utf-8"), "OLD\n",
                             f"{name} 在 0 行上传失败后被覆盖了")

    def test_successful_upload_replaces_data(self) -> None:
        resp = self._upload("2027.01")
        self.assertEqual(resp.status_code, 200, resp.get_json())
        data = resp.get_json()
        self.assertGreater(data["total"], 0)
        self.assertEqual(data["data_dir"], "data_2027.01")
        for name in EXPORT_FILES:
            text = (self.root / "data_2027.01" / name).read_text(encoding="utf-8-sig")
            self.assertNotIn("OLD", text)

    def test_upload_with_unit_uses_separate_dir(self) -> None:
        # 乐华 / TCL商用 同期间上传各自落在 data_<期间>_<单元>，互不覆盖
        self._write_units_config("乐华·电视", "TCL·TCL商用")
        resp = self._upload("2027.01", unit="乐华·电视")
        self.assertEqual(resp.status_code, 200, resp.get_json())
        data = resp.get_json()
        self.assertEqual(data["data_dir"], "data_2027.01_乐华·电视")
        self.assertEqual(data["unit"], "乐华·电视")
        for name in EXPORT_FILES:
            self.assertTrue((self.root / "data_2027.01_乐华·电视" / name).exists())
        self.assertFalse((self.root / "data_2027.01").exists())

        resp2 = self._upload("2027.01", unit="TCL·TCL商用")
        self.assertEqual(resp2.get_json()["data_dir"], "data_2027.01_TCL·TCL商用")
        # 第二次上传不能动乐华的目录
        self.assertTrue((self.root / "data_2027.01_乐华·电视" / EXPORT_FILES[0]).exists())

    def test_upload_with_unknown_unit_returns_400(self) -> None:
        self._write_units_config("乐华·电视")
        resp = self._upload("2027.01", unit="不存在的单元")
        self.assertEqual(resp.status_code, 400)
        self.assertIn("未知业务单元", resp.get_json()["error"])


if __name__ == "__main__":
    unittest.main()
