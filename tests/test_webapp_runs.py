"""Web 运行注册表：并发抢占、释放、淘汰。"""
import unittest
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from webapp.runs import RunRegistry, can_force_release, should_hold_slot


class ShouldHoldSlotTest(unittest.TestCase):
    def test_alive_proc_holds(self):
        # 子进程没被确认终止：无论如何都不放行，避免两个流程操作同一个 Excel
        self.assertTrue(should_hold_slot({}, proc_alive=True, excel_running=False))
        self.assertTrue(should_hold_slot({}, proc_alive=True, excel_running=True))

    def test_cancelled_holds_until_excel_exits(self):
        # Excel 由 COM 启动、不在 taskkill /T 的进程树内：Excel 没退出前占位
        state = {"cancelled": True}
        self.assertTrue(should_hold_slot(state, proc_alive=False, excel_running=True))
        self.assertFalse(should_hold_slot(state, proc_alive=False, excel_running=False))

    def test_timed_out_holds_until_excel_exits(self):
        state = {"timed_out": True}
        self.assertTrue(should_hold_slot(state, proc_alive=False, excel_running=True))
        self.assertFalse(should_hold_slot(state, proc_alive=False, excel_running=False))

    def test_normal_completion_releases(self):
        self.assertFalse(should_hold_slot({}, proc_alive=False, excel_running=True))


class CanForceReleaseTest(unittest.TestCase):
    def test_stuck_finished_dead_proc_can_release(self):
        self.assertTrue(can_force_release({"done": True, "stuck": True},
                                          proc_alive=False))

    def test_live_proc_cannot_release(self):
        self.assertFalse(can_force_release({"done": True, "stuck": True},
                                           proc_alive=True))

    def test_not_stuck_cannot_release(self):
        self.assertFalse(can_force_release({"done": True, "stuck": False},
                                           proc_alive=False))

    def test_not_done_cannot_release(self):
        self.assertFalse(can_force_release({"done": False, "stuck": True},
                                           proc_alive=False))


class RunRegistryTest(unittest.TestCase):
    def test_begin_is_exclusive(self):
        reg = RunRegistry()
        self.assertTrue(reg.begin("a", {"done": False}))
        self.assertFalse(reg.begin("b", {"done": False}))
        self.assertEqual(reg.current, "a")
        self.assertIsNone(reg.get("b"))

    def test_end_releases_slot(self):
        reg = RunRegistry()
        reg.begin("a", {"done": False})
        reg.end("a")
        self.assertIsNone(reg.current)
        self.assertTrue(reg.begin("b", {"done": False}))
        self.assertEqual(reg.current, "b")

    def test_end_only_releases_matching_run(self):
        reg = RunRegistry()
        reg.begin("a", {"done": False})
        reg.end("other")
        self.assertEqual(reg.current, "a")

    def test_prune_keeps_recent_finished_runs(self):
        reg = RunRegistry(max_kept=3)
        for i in range(5):
            rid = f"r{i}"
            reg.begin(rid, {"done": True})
            reg.end(rid)
        # 旧的已结束运行被淘汰，只保留最近 3 个
        self.assertIsNone(reg.get("r0"))
        self.assertIsNone(reg.get("r1"))
        self.assertIsNotNone(reg.get("r4"))

    def test_prune_never_drops_current_run(self):
        reg = RunRegistry(max_kept=1)
        reg.begin("running", {"done": False})
        reg.begin("done", {"done": True})  # 抢占失败，不注册
        self.assertIsNotNone(reg.get("running"))
        self.assertIsNone(reg.get("done"))


if __name__ == "__main__":
    unittest.main()
