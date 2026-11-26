"""导航表步骤面板刷新检测。

背景（2026-11-23 现场）：网页清 0 运行到最后一步时，AO 会话已失效，
选中 AC70 后右侧「步骤事项」面板仍停在上一张表的清单，AT52 点下去
毫无反应，重试 3 次后报「未打开新工作簿」。修复后必须先确认面板
真的刷新、按钮对应行类型正确，会话失效则抛 SessionExpiredError
（由 run_with_recovery 强制重启 Excel 并重登）。
"""
from __future__ import annotations

import fnmatch
import re
import sys
import time
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401

from engine.excel import ExcelError
from engine.orchestrator import (
    PbfFlow, SessionExpiredError, _package_items, is_session_expired)


class FakeRange:
    def __init__(self, sheet: "FakeSheet", addr: str) -> None:
        self.sheet = sheet
        self.addr = addr

    def Select(self) -> None:      # noqa: N802 - 模仿 Excel COM 命名
        self.sheet.select(self.addr)

    @property
    def Value2(self):              # noqa: N802
        return self.sheet.value(self.addr)


class FakeWorkbook:
    def __init__(self, name: str) -> None:
        self.Name = name

    def Activate(self) -> None:     # noqa: N802
        return None


class FakePkgSheet:
    """运行包执行清单：AB73 步骤名 / AB60 版本 / AO73 状态 / AW73 时长。"""

    def __init__(self, cells: dict) -> None:
        self.cells = dict(cells)

    def value(self, addr: str):
        return self.cells.get(addr)


class FakeSheet:
    """导航表：面板按选中业务单元静态写入，模拟真实 VBA 行为。"""

    def __init__(self, engine: "FakeEngine", panels: dict[str, dict[int, tuple]],
                 refresh_on_select: bool = True) -> None:
        self.engine = engine
        self.panels = panels
        self.refresh_on_select = refresh_on_select
        self.panel: dict[int, list] = {}
        self.version = "滚动预测2026年11月财总版"
        self.selection_log: list[str] = []
        self.btn = "AT52"
        # 导航表 Worksheet_SelectionChange 在处理选中行时会写 B4（本次选择行数），
        # 写在 AO 调用之前；handler_runs=False 模拟选中未触发 VBA（会话/窗口异常）。
        self.b4 = None
        self.handler_runs = True
        # 真实 AO 打开的工作簿名带时间戳（…(17-55-41).xlsb），每次都不一样；
        # 若这里重复追加同名，"新增工作簿"按集合差集判断会永远为空。
        self.open_count = 0
        self.clicks_blocked = False
        # AA(业务模块) / AB(流程步骤) / AC(业务单元) 三列标签（现场实测原文）
        self.labels: dict[str, str] = {
            "AA55": "收入", "AB55": "CBG总部编制价格&销量&开单扣率",
            "AC55": "CBG-智屏-价格&销量-乐华",
            "AA57": "收入", "AB57": "CBG总部编制客户费用流程",
            "AC57": "CBG-乐华客户费用-乐华",
            "AA62": "收入", "AB62": "毛利（CT）之上查询_CBG_查询流程",
            "AC62": "毛利之上查询-CBG-乐华",
            "AA65": "变动费用", "AB65": "编制毛利之下费用",
            "AC65": "CBG-智屏-乐华品牌",
            "AA67": "变动费用", "AB67": "查询毛利之下费用",
            "AC67": "CBG-智屏-乐华品牌",
            "AA70": "损益", "AB70": "CBG损益查询流程",
            "AC70": "CBG-乐华",
        }

    # COM 形状
    def Activate(self) -> None:     # noqa: N802
        return None

    def Range(self, addr: str) -> FakeRange:   # noqa: N802
        return FakeRange(self, addr)

    # 行为
    def select(self, addr: str) -> None:
        self.selection_log.append(addr)
        if addr == "A1":
            return
        if not self.handler_runs:
            return
        if addr in self.panels:
            self.b4 = int(addr[2:])
            if self.refresh_on_select:
                self.panel = {r: list(v) for r, v in self.panels[addr].items()}
            return
        if (addr == self.btn and not self.clicks_blocked
                and self.panel.get(52, [None] * 5)[4] == "打开"):
            self.open_count += 1
            self.engine.names.append(
                f"PBF Pro运行包执行及日志查看清单 (9-{self.open_count:02d}).xlsb")

    def value(self, addr: str):
        if addr == "AP48:AU60":
            return tuple(
                tuple(self.panel.get(r, [None] * 6)[:6]) for r in range(48, 61)
            )
        m = re.fullmatch(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", addr)
        if m:
            # 批量扫描 AA:AC：返回二维 tuple（未配置的格子为 None），
            # 与 Excel Range.Value2 形状一致
            c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
            cols = [c1] if c1 == c2 else ["AA", "AB", "AC"]
            return tuple(
                tuple(self.value(f"{col}{r}") for col in cols)
                for r in range(r1, r2 + 1)
            )
        if addr == "AR45":
            return self.version
        if addr == "B4":
            return self.b4
        if addr.startswith("AT"):
            r = int(addr[2:])
            row = self.panel.get(r)
            return row[4] if row else None
        if addr.startswith("AR"):
            r = int(addr[2:])
            row = self.panel.get(r)
            return row[2] if row else None
        if addr.startswith("AP"):
            r = int(addr[2:])
            row = self.panel.get(r)
            return row[0] if row else None
        if addr in self.labels:
            return self.labels[addr]
        return None


class FakeEngine:
    def __init__(self, sheet: FakeSheet, wb: FakeWorkbook) -> None:
        self.sheet = sheet
        self.wb = wb
        self.names = [wb.Name]
        self.pid = 4321
        self.macro_error: Exception | None = None
        self.pkg_sheet: FakePkgSheet | None = None
        self.closed: list[str] = []
        self.macros_run: list[str] = []

    def _retry(self, fn, attempts: int = 1, delay: float = 0.0):
        return fn()

    def _find_wb(self, pattern: str) -> FakeWorkbook:
        for name in self.names:
            if fnmatch.fnmatchcase(name, pattern):
                if "运行包" in name:
                    # 运行包清单每次打开都是新副本（名字带时间戳）
                    return FakeWorkbook(name)
                return self.wb
        raise ExcelError(f"没找到匹配 {pattern} 的工作簿")

    def get_workbook(self, pattern: str, wait_sec: int = 0) -> FakeWorkbook:
        return self._find_wb(pattern)

    def find_workbook(self, pattern: str) -> FakeWorkbook:
        return self._find_wb(pattern)

    def open_sheet(self, wb, name: str):
        if "运行包" in getattr(wb, "Name", "") and self.pkg_sheet is not None:
            return self.pkg_sheet
        return self.sheet

    def workbook_names(self) -> list[str]:
        return list(self.names)

    def close_workbook(self, wb, save: bool = False) -> None:
        self.closed.append(wb.Name)
        self.names = [n for n in self.names if n != wb.Name]

    def read(self, sheet, cell: str):
        return sheet.value(cell)

    def select(self, sheet: FakeSheet, cell: str) -> None:
        sheet.select(cell)

    def run_macro(self, name: str, wb=None, *args):
        self.macros_run.append(name)
        if self.macro_error is not None:
            raise self.macro_error
        return None


def make_flow(refresh_on_select: bool = True, unit_expect=None, user_key=None):
    panels = {
        "AC62": {
            52: ["PS_经营模型_收入_预测_滚动预测收入计算", "ZPS_1", "运行包", "收入计算", "打开"],
            53: ["PS_经营模型_收入_预测_计算到销售组织的成本", "ZPS_2", "运行包", "成本计算", "打开"],
        },
        "AC70": {
            52: ["YC_CBG损益查询表", "ZWB_1", "表单", "损益查询", "打开"],
            53: ["PS_计算损益_滚动预测", "ZPS_9", "运行包", "计算滚动预测损益", "打开"],
        },
        "AC55": {
            52: ["YCSRW026_价格&销量&加权CT编制表", "ZWB_X", "表单", "价格销量", "打开"],
        },
    }
    wb = FakeWorkbook("PBF Pro - 全面预算系统流程导航.xlsb")
    sheet = FakeSheet(None, panels, refresh_on_select=refresh_on_select)
    engine = FakeEngine(sheet, wb)
    sheet.engine = engine
    # 初始面板：上一张表的清单，AT52 仍是"打开"——最迷惑人的状态
    sheet.panel = {52: ["上一张表", "ZWB_PREV", "表单", "上一张表", "打开"]}

    cfg = {
        "navigator": {
            "workbook_pattern": "PBF Pro*导航*",
            "sheet": "nav",
            "open_button_cell": "AT52",
            "refresh_flow_macro": "Refresh_Status",
            "open_wait_sec": 0.2,
            "panel_range": "AP48:AU60",
            "panel_wait_sec": 0.2,
            "version_cell": "AR45",
        },
        "package": {
            "workbook_pattern": "PBF Pro运行包执行及日志查看清单*",
            "sheet": "运行包执行及日志查看清单",
            "execute_macro": "Trigger_PS",
            "refresh_macro": "Refresh_Sheet",
            "version_cell": "AB60",
            "status_cell": "AO73",
            "step_cell": "AB73",
            "duration_cell": "AW73",
            "poll_interval_sec": 0,
            "max_wait_sec": 1,
        },
        "verify_writeback": {},
    }
    if unit_expect is not None:
        cfg["_unit"] = {
            "name": "乐华·电视", "brand": "乐华", "category": "电视",
            "nav_expect": list(unit_expect),
        }
    flow = PbfFlow(engine, cfg, {}, Path("data_2027.02_clear"), user_key=user_key)
    return flow, engine, sheet


class PanelRefreshTests(unittest.TestCase):
    def test_stale_panel_raises_session_expired(self) -> None:
        flow, engine, sheet = make_flow(refresh_on_select=False)

        def dead(attempts=0, delay=0.0):
            raise SessionExpiredError("刷新流程失败：AO/BPC 会话可能已过期")

        flow.refresh_flow_retry = dead
        with self.assertRaises(SessionExpiredError):
            flow.open_via_nav("AC70", open_cell="AT52", expect="package")
        # 必须先点过 A1 强制 SelectionChange（重复选中同一格不会触发事件）
        self.assertEqual(sheet.selection_log[:2], ["A1", "AC70"])

    def test_panel_refresh_opens_package(self) -> None:
        flow, engine, sheet = make_flow()
        name = flow.open_via_nav("AC62", open_cell="AT52", expect="package")
        self.assertIn("运行包", name)
        self.assertIn(flow._panel_cache_key("AC62"), flow._panel_cache)

    def test_wrong_row_type_is_rejected(self) -> None:
        # AC70 的面板里 AT52 是表单；运行包实际在 AT53。行号配错必须报错，
        # 不能"打开"了表单还以为在跑运行包。
        flow, _, _ = make_flow()
        with self.assertRaises(ExcelError) as ctx:
            flow.open_via_nav("AC70", open_cell="AT52", expect="package")
        self.assertIn("行类型", str(ctx.exception))
        self.assertIn("表单", str(ctx.exception))

    def test_form_open_still_works(self) -> None:
        flow, engine, sheet = make_flow()
        sheet.panels["AC55"][52][4] = "打开"
        name = flow.open_via_nav("AC55", expect="form")
        self.assertIn("运行包", name)   # fake 打开后统一返回运行包清单名

    def test_refresh_flow_maps_0x800A9C68(self) -> None:
        flow, engine, _ = make_flow()
        engine.macro_error = ExcelError(
            "macro 'x'!Refresh_Status: (-2147352567, '发生意外。', (0, None, None, None, 0, -2146788248), None)")
        with self.assertRaises(SessionExpiredError):
            flow.refresh_flow()


class PanelCacheTests(unittest.TestCase):
    def test_reopening_same_nav_cell_not_judged_stale(self) -> None:
        """同一单元格重开（落库校验）时面板内容与上次一致，应命中缓存。"""
        flow, engine, sheet = make_flow()
        flow.open_via_nav("AC55", expect="form")
        # 再选同一行：A1 不改变面板 → 选中 AC55 后面板内容与 before 相同
        name = flow.open_via_nav("AC55", expect="form")
        self.assertIn("运行包", name)

    def test_panel_already_on_target_hits_persisted_cache(self) -> None:
        """上次运行把面板停在目标行（如手工探测后），新进程不应误判会话失效。

        现场 2026-11-23：清 0 续跑时面板恰好停在 AC70，选中前后内容一致，
        被误判为"面板未刷新"而强制重启 Excel——本次修复要杜绝这种误判。
        """
        flow1, _, sheet1 = make_flow()
        flow1.engine.select(sheet1, "AC62")
        fingerprint = flow1._panel_fingerprint()

        flow2, _, sheet2 = make_flow(refresh_on_select=False)
        # 新进程启动时面板已经停在 AC62 上
        sheet2.panel = {r: list(v) for r, v in sheet2.panels["AC62"].items()}
        flow2._panel_cache[flow2._panel_cache_key("AC62")] = fingerprint

        name = flow2.open_via_nav("AC62", expect="package")
        self.assertIn("运行包", name)

    def test_panel_already_on_target_without_cache_accepted_by_row_type(self) -> None:
        """冷启动（无缓存）+ 面板停在目标行：刷新成功过一轮后应按行类型放行。

        原缺陷：新进程没有持久化缓存，面板恰好停在目标行、内容不再变化，
        第 2 轮重试会把"没变化"判成会话失效，白重启一次 Excel 并重新登录。
        现在第 1 轮仍交给刷新流程分辨，刷新成功过（会话活着）之后，
        按钮行行类型与期望一致就接受静止面板。
        """
        flow, _, sheet = make_flow(refresh_on_select=False)
        # 面板本来就停在 AC62 的正确内容上，且没有持久化缓存
        sheet.panel = {r: list(v) for r, v in sheet.panels["AC62"].items()}

        name = flow.open_via_nav("AC62", open_cell="AT52", expect="package")
        self.assertIn("运行包", name)
        # 静止面板被采用后也应落库，下次可直接命中缓存
        self.assertIn(flow._panel_cache_key("AC62"), flow._panel_cache)

    def test_static_panel_accepted_for_any_expect_when_b4_matches(self) -> None:
        """两行共用同一流程模板（商用/乐华均为 CBG损益查询流程）时面板内容相同。

        expect="any"（损益核对）没有期望行类型，原逻辑永远不接受静止面板，
        会把"面板内容本来就一样"误报成会话失效。B4 记录本次选中行 == 目标行
        时按真实行类型接受。
        """
        flow, _, sheet = make_flow(refresh_on_select=False)
        sheet.panel = {r: list(v) for r, v in sheet.panels["AC70"].items()}
        name = flow.open_via_nav("AC70", open_cell="AT52", expect="any")
        self.assertIn("运行包", name)
        self.assertIn(flow._panel_cache_key("AC70"), flow._panel_cache)

    def test_static_panel_rejected_when_b4_points_to_another_row(self) -> None:
        """面板没变化且 B4 不是目标行：选中未生效/会话异常，不能静态放行。"""
        flow, _, sheet = make_flow(refresh_on_select=False)
        sheet.panel = {r: list(v) for r, v in sheet.panels["AC70"].items()}
        sheet.handler_runs = False     # 选中没有触发导航表 VBA，B4 停在旧行
        sheet.b4 = 52
        with self.assertRaises(SessionExpiredError):
            flow.open_via_nav("AC70", open_cell="AT52", expect="any")

    def test_stale_static_panel_with_wrong_row_type_still_rejected(self) -> None:
        """静止面板若停在别的行上，行类型对不上仍要拦下，不能被静态放行绕过。"""
        flow, _, sheet = make_flow(refresh_on_select=False)
        # 面板停在 AC70 的内容：AT52 行类型是「表单」，与期望的「运行包」不符
        sheet.panel = {r: list(v) for r, v in sheet.panels["AC70"].items()}
        with self.assertRaises(SessionExpiredError):
            flow.open_via_nav("AC62", open_cell="AT52", expect="package")

    def test_cached_panel_click_no_response_raises_session_expired(self) -> None:
        """面板是"已知正确内容"但点击毫无反应：会话大概率已死，必须走恢复。

        与原缺陷不同的是：这里不是面板停在别人身上（那会用行类型校验拦下），
        而是面板内容正确、按钮也是"打开"，点下去却什么都没发生。
        """
        flow1, _, sheet1 = make_flow()
        flow1.engine.select(sheet1, "AC62")
        fingerprint = flow1._panel_fingerprint()

        flow2, _, sheet2 = make_flow(refresh_on_select=False)
        sheet2.panel = {r: list(v) for r, v in sheet2.panels["AC62"].items()}
        sheet2.clicks_blocked = True
        flow2._panel_cache[flow2._panel_cache_key("AC62")] = fingerprint
        flow2.nav_cfg["open_wait_sec"] = 0.01

        with self.assertRaises(SessionExpiredError):
            flow2.open_via_nav("AC62", expect="package")

    def test_cache_persisted_across_processes(self) -> None:
        import json
        import tempfile

        with tempfile.TemporaryDirectory() as td:
            cache_file = Path(td) / "nav_panel_cache.json"
            flow, _, _ = make_flow()
            flow.persist_panel_cache = True
            flow._panel_cache_file = cache_file
            flow.open_via_nav("AC62", expect="package")

            key = flow._panel_cache_key("AC62")
            data = json.loads(cache_file.read_text(encoding="utf-8"))
            self.assertIn(key, data)

            flow2, _, _ = make_flow()
            flow2._panel_cache_file = cache_file
            flow2._load_panel_cache()
            self.assertEqual(flow2._panel_cache.get(key), data[key])


class NavLabelGuardTests(unittest.TestCase):
    """换账号后导航表按权限重新生成，行号可能整体移动。

    固定 nav_cell 一旦指到别的业务单元，仅凭行类型/工作簿名无法识别
    （同名表单的工作簿 pattern 相同）。因此选中前必须核对 AC 列标签，
    不匹配时零副作用中止（不点 A1、不触发 SelectionChange）。
    """

    def test_wrong_label_aborts_without_selecting(self):
        flow, engine, sheet = make_flow(unit_expect=["乐华"])
        sheet.labels["AC55"] = "CBG-智屏-价格&销量-TCL"
        with self.assertRaises(ExcelError) as ctx:
            flow.open_via_nav("AC55", expect="form")
        self.assertIn("标签", str(ctx.exception))
        self.assertIn("乐华", str(ctx.exception))
        self.assertEqual(sheet.selection_log, [])
        self.assertEqual(engine.names, ["PBF Pro - 全面预算系统流程导航.xlsb"])

    def test_correct_label_passes(self):
        flow, _, sheet = make_flow(unit_expect=["乐华", "智屏"])
        name = flow.open_via_nav("AC55", expect="form")
        self.assertIn("运行包", name)
        self.assertEqual(sheet.selection_log[:2], ["A1", "AC55"])

    def test_empty_label_rejected(self):
        flow, _, sheet = make_flow(unit_expect=["乐华"])
        sheet.labels["AC55"] = ""
        with self.assertRaisesRegex(ExcelError, "标签"):
            flow.open_via_nav("AC55", expect="form")
        self.assertEqual(sheet.selection_log, [])

    def test_step_level_expect_overrides_unit_default(self):
        flow, _, sheet = make_flow(unit_expect=["乐华"])
        # 单元级只有「乐华」能过；步骤级追加「智屏」，AC57 标签不含 → 拦截
        with self.assertRaisesRegex(ExcelError, "智屏"):
            flow.open_via_nav("AC57", expect="form",
                              expect_labels=["乐华", "智屏"])
        self.assertEqual(sheet.selection_log, [])

    def test_after_close_uses_unit_level_default(self):
        flow, _, _ = make_flow(unit_expect=["乐华"])
        # AC62 标签不含「智屏」，after_close 走单元级关键字，不应被步骤级规则误伤
        name = flow.open_via_nav("AC62", expect="package")
        self.assertIn("运行包", name)

    def test_ab_label_distinguishes_same_ac(self):
        """AC65/AC67 的 AC 都是「CBG-智屏-乐华品牌」，只能靠 AB 的编制/查询区分。"""
        flow, _, sheet = make_flow(unit_expect=["乐华"])
        expect = ["乐华", "编制毛利之下费用"]
        flow._check_nav_label(sheet, "AC65", expect)          # 编制行通过
        with self.assertRaisesRegex(ExcelError, "编制毛利之下费用"):
            flow._check_nav_label(sheet, "AC67", expect)      # 查询行被拦

    def test_verify_ab_keyword_used(self):
        """核对行 AC70 需同时命中 AB 的「CBG损益查询流程」。"""
        flow, _, sheet = make_flow(unit_expect=["乐华"])
        flow._check_nav_label(sheet, "AC70", ["乐华", "CBG损益查询流程"])
        with self.assertRaisesRegex(ExcelError, "CBG损益查询流程"):
            flow._check_nav_label(sheet, "AC62", ["乐华", "CBG损益查询流程"])


class NavLocateTests(unittest.TestCase):
    """按 AA+AB+AC 内容定位（#4）：换账号行号整体移动时也能打开正确表单。"""

    def test_unique_label_relocates_moved_row(self):
        flow, engine, sheet = make_flow(unit_expect=["乐华"])
        # 新账号权限不同：AC55 整行下移到 AC58，配置仍是 AC55
        for col in ("AA", "AB", "AC"):
            sheet.labels[f"{col}58"] = sheet.labels.pop(f"{col}55")
        sheet.panels["AC58"] = sheet.panels["AC55"]

        name = flow.open_via_nav(
            "AC55", expect="form",
            expect_labels=["乐华", "智屏", "编制价格"])
        self.assertIn("运行包", name)
        # 必须选中按内容找到的 AC58，而不是配置的 AC55
        self.assertEqual(sheet.selection_log[:2], ["A1", "AC58"])

    def test_unknown_label_aborts_without_selecting(self):
        flow, engine, sheet = make_flow(unit_expect=["乐华"])
        with self.assertRaises(ExcelError) as ctx:
            flow.open_via_nav("AC55", expect="form",
                              expect_labels=["乐华", "不存在的流程"])
        self.assertIn("按标签定位失败", str(ctx.exception))
        self.assertEqual(sheet.selection_log, [])
        self.assertEqual(engine.names, ["PBF Pro - 全面预算系统流程导航.xlsb"])

    def test_query_row_relocates_to_edit_row(self):
        """配置写成 AC67（查询）但关键字是「编制毛利之下费用」：按内容定位到 AC65。"""
        flow, _, sheet = make_flow(unit_expect=["乐华"])
        cell = flow._resolve_nav_cell(
            sheet, "AC67", ["乐华", "编制毛利之下费用"])
        self.assertEqual(cell, "AC65")

    def test_query_keywords_locate_query_row(self):
        """反向：关键字「查询毛利之下费用」唯一命中 AC67，与配置行无关。"""
        flow, _, sheet = make_flow(unit_expect=["乐华"])
        cell = flow._resolve_nav_cell(
            sheet, "AC65", ["乐华", "查询毛利之下费用"])
        self.assertEqual(cell, "AC67")

    def test_locate_disabled_uses_fixed_cell(self):
        flow, _, sheet = make_flow(unit_expect=["乐华", "编制毛利之下费用"])
        flow.nav_cfg["locate_by_label"] = False
        name = flow.open_via_nav("AC65", expect="form")
        self.assertIn("运行包", name)
        self.assertEqual(sheet.selection_log[:2], ["A1", "AC65"])


class PanelCacheAccountTests(unittest.TestCase):
    def test_cache_key_separated_by_account(self):
        a, _, _ = make_flow(user_key="alice")
        a2, _, _ = make_flow(user_key="alice")
        b, _, _ = make_flow(user_key="bob")
        ka = a._panel_cache_key("AC62")
        self.assertNotEqual(ka, b._panel_cache_key("AC62"))
        self.assertEqual(ka, a2._panel_cache_key("AC62"))
        self.assertIn("alice", ka)


class RunPackageTests(unittest.TestCase):
    """运行包 AB73 步骤名校验：面板停在别的同类型行上时绝不执行错包。

    评审 2026-11-24：accept_static 区分不了「面板停在目标行」和「面板停在
    另一个同类型行」，点下去会开到错误的运行包上下文。清单工作簿名不随包变，
    只能靠 AB73 步骤名与 flow.yaml 配置的期望名核对。
    注意：AB73 存的是面板 AP 列的模块名（PS_...），不是 AS 列的说明文字——
    期望名必须取 AB73 实际值的独特子串（评审 #14 用运行日志实测核出过）。
    """

    # 现场实测 AB73 原值（logs/20261123.jsonl，评审 #14 提供）
    AB73_INCOME = "PS_经营模型_收入_预测_滚动预测收入计算_价量一起编制"
    AB73_COST = "PS_经营模型_收入_预测_计算到销售组织的成本"
    AB73_PL = "PS_计算损益_滚动预测"

    PKG_CELLS = {
        "AB60": "滚动预测2026年11月财总版",
        "AO73": "已完成",
        "AW73": "0:01:23",
    }

    def _pkg_sheet(self, ab73: str) -> FakePkgSheet:
        return FakePkgSheet({**self.PKG_CELLS, "AB73": ab73})

    def test_step_name_match_runs_package(self) -> None:
        # AC62 R52：期望名取 AB73 的独特子串（不是 AS 列说明「收入计算」）
        flow, engine, _ = make_flow()
        engine.pkg_sheet = self._pkg_sheet(self.AB73_INCOME)
        r = flow.run_package("AC62", 52, expected_name="滚动预测收入计算")
        self.assertTrue(r["ok"], r)
        self.assertEqual(r["package"], self.AB73_INCOME)
        self.assertIn("Trigger_PS", engine.macros_run)
        self.assertIn("Refresh_Sheet", engine.macros_run)
        # 跑完关闭清单
        self.assertEqual(engine.closed, ["PBF Pro运行包执行及日志查看清单 (9-01).xlsb"])

    def test_step_name_match_ps_full_name(self) -> None:
        # AC70 的运行包行：期望名直接用 AB73 原值
        flow, engine, sheet = make_flow()
        sheet.btn = "AT53"          # fake 的 AC70 面板里运行包在第 53 行
        engine.pkg_sheet = self._pkg_sheet(self.AB73_PL)
        r = flow.run_package("AC70", 53, expected_name="PS_计算损益_滚动预测")
        self.assertTrue(r["ok"], r)
        self.assertEqual(r["package"], self.AB73_PL)

    def test_step_name_mismatch_raises_without_executing(self) -> None:
        # 现场会踩的形态：面板停在 R52（收入计算），却按 R53 的期望名核对
        flow, engine, _ = make_flow()
        engine.pkg_sheet = self._pkg_sheet(self.AB73_INCOME)
        with self.assertRaises(ExcelError) as ctx:
            flow.run_package("AC62", 52, expected_name="计算到销售组织的成本")
        self.assertIn("步骤名不符", str(ctx.exception))
        self.assertIn("计算到销售组织的成本", str(ctx.exception))
        # 核对不过就中止：Execute 宏一个都不能跑
        self.assertNotIn("Trigger_PS", engine.macros_run)
        self.assertEqual(engine.closed, ["PBF Pro运行包执行及日志查看清单 (9-01).xlsb"])

    def test_legacy_list_packages_skip_verification(self) -> None:
        # 旧行号列表写法（无期望包名）：不校验，行为同前
        flow, engine, _ = make_flow()
        engine.pkg_sheet = self._pkg_sheet(self.AB73_INCOME)
        r = flow.run_package("AC62", 52)
        self.assertTrue(r["ok"], r)

    def test_empty_name_in_dict_mapping_raises_early(self) -> None:
        # 评审建议 3：映射写法里包名为空会悄悄关掉 AB73 校验，跑之前就报错
        with self.assertRaises(ExcelError) as ctx:
            _package_items({52: ""})
        self.assertIn("包名为空", str(ctx.exception))
        with self.assertRaises(ExcelError):
            _package_items({52: None})


class SessionExpiredDetectionTests(unittest.TestCase):
    """is_session_expired：按文本特征码识别宏里直接抛出的会话失效。

    run_with_recovery 只对 SessionExpiredError 和「Excel 断开」做恢复；
    trigger_calc/run_package 的宏不走 refresh_flow 的翻译，0x800A9C68
    会裸抛上来——恢复前必须先把这类错误认出来。
    """

    def test_macro_session_codes_are_detected(self) -> None:
        self.assertTrue(is_session_expired(ExcelError(
            "macro 'x'!Execute: (-2147352567, '发生意外。', "
            "(0, None, None, None, 0, -2146788248), None)")))
        self.assertTrue(is_session_expired(Exception("0x800A9C68")))
        self.assertTrue(is_session_expired(
            SessionExpiredError("面板未刷新 (0x800A9C68)")))

    def test_generic_wrapper_codes_are_not_detected(self) -> None:
        # 「发生意外」/-2147352567 是通用 DISP_E_EXCEPTION 包装：
        # 版本与期间不匹配等普通错误也长这样，不能据此强制重启 Excel
        self.assertFalse(is_session_expired(ExcelError("发生意外。")))
        self.assertFalse(is_session_expired(ExcelError(
            "(-2147352567, '发生意外。', (0, None, None, None, 0, -2147352567), None)")))
        self.assertFalse(is_session_expired(ExcelError("未打开新工作簿")))
        self.assertFalse(is_session_expired(None))

    def test_session_expired_error_instance_without_code(self) -> None:
        # 纯文本的 SessionExpiredError 靠 run_with_recovery 的 isinstance
        # 分支识别，is_session_expired 只管文本特征——两者是互补关系
        self.assertFalse(is_session_expired(SessionExpiredError("面板一直没有刷新")))


if __name__ == "__main__":
    unittest.main()
