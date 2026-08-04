import unittest
from unittest.mock import patch

import register_outlook_standalone as standalone


PROXY = "socks5://u:p@1.2.3.4:1080"


def _proxies_for(s):
    return {"https": s, "http": s} if s else None


class GraphDirectFallbackTests(unittest.TestCase):
    """代理授权全失败后应回退直连 GRAPH_DIRECT_FALLBACK_ATTEMPTS 次；直连模式不触发回退。"""

    def _patches(self, get_token_fn):
        # extract_graph_token_http 内部 `from extract_graph_tokens import get_graph_token`
        return (
            patch("extract_graph_tokens.get_graph_token", side_effect=get_token_fn),
            patch.object(standalone, "_proxy_for_requests", side_effect=_proxies_for),
            patch.object(standalone.time, "sleep"),
        )

    def test_proxy_mode_falls_back_to_direct(self):
        calls = []

        def fake_get(email, password, idx, proxies=None):
            calls.append(proxies)
            return None  # 始终失败

        p1, p2, p3 = self._patches(fake_get)
        with p1, p2, p3:
            result = standalone.extract_graph_token_http("a@b.com", "p", 1, attempts=3, proxy_str=PROXY)

        # 3 次代理 + 2 次直连回退
        self.assertEqual(len(calls), 5)
        self.assertEqual(calls[:3], [_proxies_for(PROXY)] * 3)
        self.assertEqual(calls[3:], [None, None])
        self.assertIsNone(result)

    def test_proxy_mode_succeeds_no_fallback(self):
        calls = []

        def fake_get(email, password, idx, proxies=None):
            calls.append(proxies)
            return {"refresh_token": "rt", "client_id": "cid"}

        p1, p2, p3 = self._patches(fake_get)
        with p1, p2, p3:
            result = standalone.extract_graph_token_http("a@b.com", "p", 1, attempts=3, proxy_str=PROXY)

        # 第 1 次代理就成功，不回退
        self.assertEqual(len(calls), 1)
        self.assertEqual(calls[0], _proxies_for(PROXY))
        self.assertEqual(result, {"refresh_token": "rt", "client_id": "cid"})

    def test_direct_mode_no_fallback(self):
        calls = []

        def fake_get(email, password, idx, proxies=None):
            calls.append(proxies)
            return None

        p1, p2, p3 = self._patches(fake_get)
        with p1, p2, p3:
            result = standalone.extract_graph_token_http("a@b.com", "p", 1, attempts=3, proxy_str=None)

        # 直连模式：只 attempts(3) 次，不触发回退
        self.assertEqual(len(calls), 3)
        self.assertEqual(calls, [None, None, None])
        self.assertIsNone(result)


if __name__ == "__main__":
    unittest.main()
