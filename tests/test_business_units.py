"""业务单元（品牌 × 品类）配置：覆盖、校验与导出空列告警。"""
from __future__ import annotations

import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def _load(name: str, filename: str):
    path = ROOT / "engine" / filename
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


bu = _load("business_units_under_test", "business_units.py")


def _flow() -> dict:
    return {
        "steps": [
            {"form": "E61_CBG_价格销量CT编制", "nav_cell": "AC55"},
            {"form": "HQ_毛上调整", "nav_cell": "AC59",
             "after_close": {"nav_cell": "AC62", "packages": [52, 53]}},
            {"form": "费用及损益其他项编制_非海外", "nav_cell": "AC65",
             "after_close": {"nav_cell": "AC70", "packages": [52]}},
        ],
        "verify": {"nav_cell": "AC70", "form_row": 54},
    }


def _units() -> dict:
    return {
        "品牌": ["TCL", "雷鸟", "乐华"],
        "品类": ["电视", "空调"],
        "units": {
            "乐华·电视": {
                "启用": True, "品牌": "乐华", "品类": "电视",
                "nav_expect": ["乐华"],
                "verify_nav_expect": ["乐华", "损益"],
                "verify_nav_cell": "B70", "verify_form_row": 55,
                "steps": [
                    {"表单": "价格销量CT", "nav_cell": "B55",
                     "nav_expect": ["乐华", "价格"]},
                    {"表单": "HQ毛上调整", "nav_cell": "B59",
                     "after_close_nav_cell": "B62",
                     "after_close_nav_expect": ["乐华", "毛利"]},
                    {"表单": "费用及损益", "nav_cell": "B65", "after_close_nav_cell": "B70"},
                ],
            },
            "TCL·电视": {"启用": False, "品牌": "TCL", "品类": "电视"},
            "TCL·空调": {"启用": False, "品牌": "TCL", "品类": "空调"},
        },
    }


class TestLookup(unittest.TestCase):
    def test_enabled_and_default(self):
        cfg = _units()
        self.assertEqual(bu.enabled_unit_names(cfg), ["乐华·电视"])
        self.assertEqual(bu.default_unit(cfg), "乐华·电视")
        self.assertEqual(bu.unit_names(cfg),
                         ["乐华·电视", "TCL·电视", "TCL·空调"])

    def test_no_enabled_unit(self):
        cfg = _units()
        for u in cfg["units"].values():
            u["启用"] = False
        self.assertIsNone(bu.default_unit(cfg))

    def test_get_unknown_unit(self):
        with self.assertRaises(bu.UnitError):
            bu.get_unit(_units(), "雷鸟·电视")

    def test_get_disabled_unit(self):
        with self.assertRaises(bu.UnitError):
            bu.get_unit(_units(), "TCL·电视")


class TestApplyUnit(unittest.TestCase):
    def test_overrides_nav_cells(self):
        flow = bu.apply_unit(_flow(), _units(), "乐华·电视")
        self.assertEqual([s["nav_cell"] for s in flow["steps"]], ["B55", "B59", "B65"])
        self.assertEqual(flow["steps"][1]["after_close"]["nav_cell"], "B62")
        self.assertEqual(flow["steps"][2]["after_close"]["nav_cell"], "B70")
        # packages 不被单元配置破坏
        self.assertEqual(flow["steps"][1]["after_close"]["packages"], [52, 53])
        self.assertEqual(flow["steps"][1]["after_close"]["nav_expect"], ["乐华", "毛利"])
        self.assertEqual(flow["verify"]["nav_cell"], "B70")
        self.assertEqual(flow["verify"]["form_row"], 55)
        self.assertEqual(flow["verify"]["nav_expect"], ["乐华", "损益"])
        self.assertEqual(flow["_unit"],
                         {"name": "乐华·电视", "brand": "乐华", "category": "电视",
                          "nav_expect": ["乐华"], "skipped_steps": []})
        # steps 级 nav_expect 覆盖到 flow step；没配置的步骤不凭空多出字段
        self.assertEqual(flow["steps"][0]["nav_expect"], ["乐华", "价格"])
        self.assertNotIn("nav_expect", flow["steps"][1])

    def test_does_not_mutate_inputs(self):
        base = _flow()
        bu.apply_unit(base, _units(), "乐华·电视")
        self.assertEqual(base["steps"][0]["nav_cell"], "AC55")
        self.assertEqual(base["verify"]["nav_cell"], "AC70")
        self.assertNotIn("_unit", base)
        self.assertNotIn("nav_expect", base["steps"][0])
        self.assertNotIn("nav_expect", base["verify"])
        self.assertNotIn("nav_expect", base["steps"][1]["after_close"])

    def test_without_optional_overrides(self):
        units = _units()
        unit = units["units"]["乐华·电视"]
        unit.pop("verify_nav_cell")
        unit.pop("verify_form_row")
        unit.pop("verify_nav_expect")
        for step in unit["steps"]:
            step.pop("after_close_nav_cell", None)
            step.pop("after_close_nav_expect", None)
        flow = bu.apply_unit(_flow(), units, "乐华·电视")
        self.assertEqual(flow["steps"][1]["after_close"]["nav_cell"], "AC62")
        self.assertEqual(flow["verify"]["nav_cell"], "AC70")
        self.assertNotIn("nav_expect", flow["verify"])
        self.assertNotIn("nav_expect", flow["steps"][1]["after_close"])

    def test_step_count_mismatch(self):
        units = _units()
        units["units"]["乐华·电视"]["steps"].pop()
        with self.assertRaisesRegex(bu.UnitError, "steps 数量"):
            bu.apply_unit(_flow(), units, "乐华·电视")

    def test_step_form_mismatch(self):
        units = _units()
        units["units"]["乐华·电视"]["steps"][1]["表单"] = "客户费用"
        with self.assertRaisesRegex(bu.UnitError, "不对应"):
            bu.apply_unit(_flow(), units, "乐华·电视")

    def test_missing_nav_cell(self):
        units = _units()
        units["units"]["乐华·电视"]["steps"][0]["nav_cell"] = ""
        with self.assertRaisesRegex(bu.UnitError, "nav_cell"):
            bu.apply_unit(_flow(), units, "乐华·电视")

    def test_empty_step_nav_expect_rejected(self):
        units = _units()
        units["units"]["乐华·电视"]["steps"][0]["nav_expect"] = []
        with self.assertRaisesRegex(bu.UnitError, "nav_expect"):
            bu.apply_unit(_flow(), units, "乐华·电视")

    def test_empty_after_close_expect_rejected(self):
        units = _units()
        units["units"]["乐华·电视"]["steps"][1]["after_close_nav_expect"] = []
        with self.assertRaisesRegex(bu.UnitError, "after_close_nav_expect"):
            bu.apply_unit(_flow(), units, "乐华·电视")

    def test_after_close_expect_without_after_close_rejected(self):
        units = _units()
        units["units"]["乐华·电视"]["steps"][0]["after_close_nav_expect"] = ["乐华"]
        with self.assertRaisesRegex(bu.UnitError, "没有 after_close"):
            bu.apply_unit(_flow(), units, "乐华·电视")

    def test_disabled_unit_rejected(self):
        with self.assertRaises(bu.UnitError):
            bu.apply_unit(_flow(), _units(), "TCL·电视")


class TestSkipSteps(unittest.TestCase):
    """单元级 跳过步骤：该单元流程里没有的步骤整步移除。"""

    def _units_with_skip(self, skip=("HQ_毛上调整",)):
        units = _units()
        unit = units["units"]["乐华·电视"]
        unit["跳过步骤"] = list(skip)
        drop = {h.replace("_", "") for h in skip}
        unit["steps"] = [
            s for s in unit["steps"] if s["表单"].replace("_", "") not in drop]
        return units

    def test_skip_removes_flow_step(self):
        flow = bu.apply_unit(_flow(), self._units_with_skip(), "乐华·电视")
        self.assertEqual([s["form"] for s in flow["steps"]],
                         ["E61_CBG_价格销量CT编制", "费用及损益其他项编制_非海外"])
        # 剩余步骤仍按顺序覆盖 nav_cell
        self.assertEqual(flow["steps"][0]["nav_cell"], "B55")
        self.assertEqual(flow["steps"][1]["nav_cell"], "B65")
        self.assertEqual(flow["steps"][1]["after_close"]["nav_cell"], "B70")
        self.assertEqual(flow["_unit"]["skipped_steps"], ["HQ_毛上调整"])

    def test_skip_does_not_mutate_input(self):
        base = _flow()
        bu.apply_unit(base, self._units_with_skip(), "乐华·电视")
        self.assertEqual(len(base["steps"]), 3)
        self.assertNotIn("skipped_steps", base.get("_unit") or {})

    def test_unknown_skip_hint_rejected(self):
        units = _units()
        units["units"]["乐华·电视"]["跳过步骤"] = ["不存在的步骤"]
        with self.assertRaisesRegex(bu.UnitError, "跳过步骤"):
            bu.apply_unit(_flow(), units, "乐华·电视")

    def test_ambiguous_skip_hint_rejected(self):
        units = _units()
        units["units"]["乐华·电视"]["跳过步骤"] = ["编制"]   # 命中 2 个步骤
        with self.assertRaisesRegex(bu.UnitError, "应恰好 1 个"):
            bu.apply_unit(_flow(), units, "乐华·电视")

    def test_non_list_skip_rejected(self):
        units = _units()
        units["units"]["乐华·电视"]["跳过步骤"] = "HQ毛上调整"
        with self.assertRaisesRegex(bu.UnitError, "列表"):
            bu.apply_unit(_flow(), units, "乐华·电视")

    def test_skip_used_by_check_units(self):
        issues = bu.check_units(self._units_with_skip(), _flow(), only="乐华·电视")
        self.assertEqual(issues, [])


class TestApplyFormOverrides(unittest.TestCase):
    """单元级 表单覆盖：跳过度量（该单元这张表不编制的度量块）。"""

    def _forms(self):
        return {
            "E61_CBG_价格销量CT编制": {
                "measures": {
                    "销售单价": {"2026.11": "AH"},
                    "销量": {"2026.11": "AO"},
                    "加权CT": {"2026.11": "AV"},
                    "销售收入": {"2026.11": "BC"},
                },
                "computed_measures": ["销售收入"],
            },
            "E61_CBG_其他成本编制": {"measures": {"其他": {}}},
        }

    def _units_with_override(self, skip=None):
        units = _units()
        units["units"]["乐华·电视"]["表单覆盖"] = {
            "E61_CBG_价格销量CT编制": {"跳过度量": ["加权CT"] if skip is None else skip}}
        return units

    def test_skips_measure(self):
        forms = bu.apply_form_overrides(self._forms(), self._units_with_override(),
                                        "乐华·电视")
        measures = forms["E61_CBG_价格销量CT编制"]["measures"]
        self.assertNotIn("加权CT", measures)
        self.assertEqual(set(measures), {"销售单价", "销量", "销售收入"})
        self.assertEqual(forms["E61_CBG_价格销量CT编制"]["computed_measures"],
                         ["销售收入"])

    def test_does_not_mutate_input(self):
        base = self._forms()
        bu.apply_form_overrides(base, self._units_with_override(), "乐华·电视")
        self.assertIn("加权CT", base["E61_CBG_价格销量CT编制"]["measures"])

    def test_unknown_measure_rejected(self):
        with self.assertRaisesRegex(bu.UnitError, "跳过度量"):
            bu.apply_form_overrides(self._forms(), self._units_with_override(["CT"]),
                                    "乐华·电视")

    def test_unknown_form_rejected(self):
        units = _units()
        units["units"]["乐华·电视"]["表单覆盖"] = {
            "不存在的表单": {"跳过度量": ["加权CT"]}}
        with self.assertRaisesRegex(bu.UnitError, "表单覆盖"):
            bu.apply_form_overrides(self._forms(), units, "乐华·电视")

    def test_empty_skip_is_noop(self):
        forms = bu.apply_form_overrides(self._forms(), self._units_with_override([]),
                                        "乐华·电视")
        self.assertIn("加权CT", forms["E61_CBG_价格销量CT编制"]["measures"])

    def test_bad_override_type_reported_by_check_units(self):
        units = _units()
        units["units"]["乐华·电视"]["表单覆盖"] = ["加权CT"]
        issues = bu.check_units(units, _flow(), only="乐华·电视")
        self.assertTrue(any("表单覆盖" in i for i in issues))


class TestCheckUnits(unittest.TestCase):
    def test_valid_config_passes(self):
        self.assertEqual(bu.check_units(_units(), _flow()), [])

    def test_missing_brand_or_category(self):
        units = _units()
        units["units"]["TCL·电视"].pop("品类")
        issues = bu.check_units(units, _flow())
        self.assertTrue(any("缺少 品牌/品类" in i for i in issues))

    def test_enabled_unit_without_nav_expect_is_error(self):
        # 换账号后导航行会移动；没有标签关键字就失去了唯一的安全网，必须阻塞
        units = _units()
        units["units"]["乐华·电视"].pop("nav_expect")
        issues = bu.check_units(units, _flow())
        self.assertTrue(any(i.startswith("错误:") and "nav_expect" in i
                            for i in issues))

    def test_disabled_unit_without_nav_expect_is_ok(self):
        # 未启用单元不参与运行，不要求配置标签关键字
        issues = bu.check_units(_units(), _flow())
        self.assertEqual(issues, [])

    def test_missing_verify_nav_expect_warns(self):
        units = _units()
        units["units"]["乐华·电视"].pop("verify_nav_expect")
        issues = bu.check_units(units, _flow())
        self.assertTrue(any(i.startswith("告警:") and "verify_nav_expect" in i
                            for i in issues))
        self.assertFalse(any(i.startswith("错误:") and "verify_nav_expect" in i
                             for i in issues))

    def test_empty_verify_nav_expect_is_error(self):
        units = _units()
        units["units"]["乐华·电视"]["verify_nav_expect"] = []
        issues = bu.check_units(units, _flow())
        self.assertTrue(any(i.startswith("错误:") and "verify_nav_expect" in i
                            for i in issues))

    def test_brand_not_in_list(self):
        units = _units()
        units["units"]["TCL·空调"]["品牌"] = "索尼"
        issues = bu.check_units(units, _flow())
        self.assertTrue(any("不在 品牌 列表中" in i for i in issues))

    def test_enabled_unit_with_bad_steps(self):
        units = _units()
        units["units"]["乐华·电视"]["steps"] = units["units"]["乐华·电视"]["steps"][:1]
        issues = bu.check_units(units, _flow())
        self.assertTrue(any(i.startswith("错误:") and "steps 数量" in i for i in issues))

    def test_duplicate_nav_cell(self):
        units = _units()
        steps = units["units"]["乐华·电视"]["steps"]
        steps[1]["nav_cell"] = steps[0]["nav_cell"]
        issues = bu.check_units(units, _flow())
        self.assertTrue(any("nav_cell 重复" in i for i in issues))

    def test_no_units(self):
        issues = bu.check_units({}, _flow())
        self.assertTrue(any("未配置任何业务单元" in i for i in issues))


class TestCheckUnitsOnly(unittest.TestCase):
    """按 --unit 运行时只校验被选中的单元，其余单元的问题不阻塞。"""

    def test_only_ignores_other_units(self):
        units = _units()
        # 未启用的占位单元缺品类：全量校验报错，only 校验不应报
        units["units"]["TCL·电视"].pop("品类")
        self.assertTrue(any("缺少 品牌/品类" in i
                            for i in bu.check_units(units, _flow())))
        self.assertEqual(bu.check_units(units, _flow(), only="乐华·电视"), [])

    def test_only_reports_selected_unit_issues(self):
        units = _units()
        steps = units["units"]["乐华·电视"]["steps"]
        steps[1]["nav_cell"] = steps[0]["nav_cell"]
        issues = bu.check_units(units, _flow(), only="乐华·电视")
        self.assertTrue(any("nav_cell 重复" in i for i in issues))

    def test_only_reports_selected_unit_missing_dims(self):
        units = _units()
        units["units"]["乐华·电视"].pop("品类")
        issues = bu.check_units(units, _flow(), only="乐华·电视")
        self.assertTrue(any("缺少 品牌/品类" in i for i in issues))

    def test_only_unknown_unit(self):
        issues = bu.check_units(_units(), _flow(), only="雷鸟·电视")
        self.assertTrue(any("未知业务单元" in i for i in issues))

    def test_only_skips_global_no_enabled_check(self):
        # only 指定的单元本身是启用的，全局“没有启用单元”检查不参与
        units = _units()
        for name, u in units["units"].items():
            if name != "乐华·电视":
                u["启用"] = False
        self.assertEqual(bu.check_units(units, _flow(), only="乐华·电视"), [])


class TestNavLocate(unittest.TestCase):
    """按 AA+AB+AC 标签内容定位导航行（换账号后行号整体移动也能对上）。"""

    def _rows(self):
        # 模拟 Range.Value2 的二维读法：AA / AB / AC 三列；行 81 全空被跳过
        return bu.nav_rows_from_grid((
            ("收入", "CBG总部编制价格&销量&开单扣率", "CBG-智屏-价格&销量-乐华"),
            (None, None, None),
            ("变动费用", "编制毛利之下费用", "CBG-智屏-乐华品牌"),
            ("变动费用", "查询毛利之下费用", "CBG-智屏-乐华品牌"),
        ), start_row=80)

    def test_grid_rows_and_labels(self):
        rows = self._rows()
        self.assertEqual([r["row"] for r in rows], [80, 82, 83])
        self.assertTrue(rows[0]["label"].startswith("收入 | CBG总部编制价格"))
        self.assertEqual(rows[1]["ac"], "CBG-智屏-乐华品牌")

    def test_single_row_grid_normalized(self):
        rows = bu.nav_rows_from_grid(("收入", "AB文本", "AC文本"), start_row=7)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["row"], 7)

    def test_unique_match_relocates(self):
        rows = self._rows()
        # 配置 AC82（编制行），内容唯一 → 用命中行（这里恰好同格）
        cell, mode = bu.resolve_nav_row(rows, "AC82", ["乐华", "编制毛利之下费用"])
        self.assertEqual((cell, mode), ("AC82", "unique"))
        # 配置行已过时（AC99），仍然按内容找到 82 → 唯一命中
        cell, mode = bu.resolve_nav_row(rows, "AC99", ["乐华", "编制毛利之下费用"])
        self.assertEqual((cell, mode), ("AC82", "unique"))

    def test_same_ac_distinguished_by_ab_keyword(self):
        rows = self._rows()
        cell, mode = bu.resolve_nav_row(rows, "AC83", ["乐华", "编制毛利之下费用"])
        self.assertEqual((cell, mode), ("AC82", "unique"))

    def test_ambiguous_keeps_configured_row(self):
        rows = self._rows()
        cell, mode = bu.resolve_nav_row(rows, "AC83", ["乐华"])
        self.assertEqual((cell, mode), ("AC83", "ambiguous"))

    def test_zero_hit_raises(self):
        with self.assertRaises(bu.NavRowError) as ctx:
            bu.resolve_nav_row(self._rows(), "AC82", ["乐华", "不存在的流程"])
        self.assertIn("标签", str(ctx.exception))
        self.assertIn("乐华", str(ctx.exception))

    def test_multiple_hits_without_configured_raises(self):
        with self.assertRaises(bu.NavRowError) as ctx:
            bu.resolve_nav_row(self._rows(), "AC99", ["乐华"])
        self.assertIn("不在其中", str(ctx.exception))
        self.assertIn("AC82", str(ctx.exception))

    def test_no_keywords_uses_fixed_cell(self):
        cell, mode = bu.resolve_nav_row(self._rows(), "AC82", [])
        self.assertEqual((cell, mode), ("AC82", "fixed"))


class TestCsvDimensionWarnings(unittest.TestCase):
    def _write(self, directory: Path, name: str, text: str) -> None:
        (directory / name).write_text(text, encoding="utf-8-sig")

    def test_empty_brand_and_category(self):
        with tempfile.TemporaryDirectory() as td:
            d = Path(td)
            self._write(d, "价格销量CT.csv",
                        "组织,渠道,品牌,产品,期间,销量\n组织A,渠道A,,产品A,2027.01,10\n")
            warnings = bu.csv_dimension_warnings(d)
            self.assertEqual(len(warnings), 1)
            self.assertIn("品牌", warnings[0])
            self.assertIn("全部为空", warnings[0])

    def test_filled_dimensions_no_warning(self):
        with tempfile.TemporaryDirectory() as td:
            d = Path(td)
            self._write(d, "a.csv",
                        "品牌,品类,期间,值\n乐华,电视,2027.01,1\n")
            self.assertEqual(bu.csv_dimension_warnings(d), [])

    def test_missing_column_no_warning(self):
        with tempfile.TemporaryDirectory() as td:
            d = Path(td)
            self._write(d, "b.csv", "科目,期间,值\nA,2027.01,1\n")
            self.assertEqual(bu.csv_dimension_warnings(d), [])

    def test_gbk_csv(self):
        with tempfile.TemporaryDirectory() as td:
            d = Path(td)
            (d / "c.csv").write_bytes("品牌,期间\n,2027.01\n".encode("gbk"))
            warnings = bu.csv_dimension_warnings(d)
            self.assertEqual(len(warnings), 1)


if __name__ == "__main__":
    unittest.main()
