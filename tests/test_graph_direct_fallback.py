import unittest

import register_outlook_standalone as standalone


class ProxyParseTests(unittest.TestCase):
    """socks5h 代理 URL 必须被正确识别，不能退化成 http://socks5h://... 畸形 URL。"""

    def test_parse_proxy_handles_socks5h(self):
        p = standalone.BitBrowserClient._parse_proxy("socks5h://u:p@1.2.3.4:1080")
        self.assertIsNotNone(p)
        self.assertEqual(p["type"], "socks5")
        self.assertEqual(p["host"], "1.2.3.4")
        self.assertEqual(p["port"], "1080")
        self.assertEqual(p["username"], "u")
        self.assertEqual(p["password"], "p")

    def test_proxy_for_requests_preserves_socks5h(self):
        proxies = standalone._proxy_for_requests("socks5h://u:p@1.2.3.4:1080")
        self.assertIsNotNone(proxies)
        self.assertEqual(proxies["https"], "socks5h://u:p@1.2.3.4:1080")
        self.assertEqual(proxies["http"], "socks5h://u:p@1.2.3.4:1080")
        # 回归保护：不能退化成 http://socks5h://... 这种畸形 URL
        self.assertFalse(proxies["https"].startswith("http://socks5h"))


if __name__ == "__main__":
    unittest.main()
