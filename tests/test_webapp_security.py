"""Web 请求来源校验：防跨站 POST 与 DNS rebinding。"""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from webapp.security import origin_allowed


class OriginAllowedTest(unittest.TestCase):
    def test_no_origin_allows_scripts(self):
        self.assertTrue(origin_allowed(None, "127.0.0.1:8765"))

    def test_local_origin_and_host_allowed(self):
        self.assertTrue(origin_allowed("http://127.0.0.1:8765", "127.0.0.1:8765"))
        self.assertTrue(origin_allowed("http://localhost:8765", "localhost:8765"))

    def test_foreign_origin_rejected(self):
        self.assertFalse(origin_allowed("http://evil.example", "127.0.0.1:8765"))

    def test_dns_rebinding_rejected(self):
        # 攻击者域名解析到 127.0.0.1 时 Origin 与 Host 相同，也必须拒绝
        self.assertFalse(origin_allowed("http://evil.example:8765", "evil.example:8765"))

    def test_cross_site_fetch_metadata_rejected(self):
        self.assertFalse(origin_allowed(None, "127.0.0.1:8765", "cross-site"))
        self.assertFalse(origin_allowed("http://127.0.0.1:8765", "127.0.0.1:8765",
                                        "cross-site"))

    def test_non_local_host_rejected_even_with_local_origin(self):
        self.assertFalse(origin_allowed("http://127.0.0.1:8765", "192.168.1.20:8765"))


if __name__ == "__main__":
    unittest.main()
