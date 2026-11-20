"""卡住的占位运行：人工释放接口只放行已结束且 stuck 的状态。"""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

try:
    import webapp.app as wa
    HAS_FLASK = True
except ImportError:                     # 轻量环境未装 flask：跳过，CI 装了会跑
    wa = None
    HAS_FLASK = False


@unittest.skipUnless(HAS_FLASK, "需要 flask（pip install -r requirements.txt）")
class ReleaseEndpointTest(unittest.TestCase):
    RUN_ID = "test-release"

    def setUp(self) -> None:
        if wa.REGISTRY.current is not None:
            wa.REGISTRY.end(wa.REGISTRY.current)

    def tearDown(self) -> None:
        wa.REGISTRY.end(self.RUN_ID)

    def _post(self):
        return wa.app.test_client().post("/api/run/release")

    def test_releases_stuck_finished_run(self) -> None:
        wa.REGISTRY.begin(self.RUN_ID, {"done": True, "stuck": True, "proc": None})
        resp = self._post()
        self.assertEqual(resp.status_code, 200, resp.get_json())
        self.assertIsNone(wa.REGISTRY.current)

    def test_refuses_when_child_process_still_alive(self) -> None:
        class LiveProc:
            def poll(self):
                return None

        wa.REGISTRY.begin(self.RUN_ID,
                          {"done": True, "stuck": True, "proc": LiveProc()})
        resp = self._post()
        self.assertEqual(resp.status_code, 409)
        self.assertEqual(wa.REGISTRY.current, self.RUN_ID)

    def test_refuses_running_or_not_stuck(self) -> None:
        wa.REGISTRY.begin(self.RUN_ID, {"done": True, "stuck": False, "proc": None})
        resp = self._post()
        self.assertEqual(resp.status_code, 409)
        self.assertEqual(wa.REGISTRY.current, self.RUN_ID)

    def test_refuses_when_no_run(self) -> None:
        resp = self._post()
        self.assertEqual(resp.status_code, 409)


if __name__ == "__main__":
    unittest.main()
