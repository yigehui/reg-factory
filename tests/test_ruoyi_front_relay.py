import asyncio
import os
import unittest
from types import SimpleNamespace
from unittest.mock import patch

import common.ruyi as ruyi
from common.ruyi.chain import (
    ChainRelayHub,
    parse_front_proxy,
    front_relay_url,
    front_proxy_raw,
    ensure_front_relay,
)

import register_outlook_ruoyi as register_mod
import unlock_outlook as unlock_mod


PROXY = "socks5://u:p@1.2.3.4:1080"
FRONT = "socks5://127.0.0.1:10808"


def _reset_singleton():
    """清掉 chain 模块进程级单例,避免测试间串扰(端口映射按 front 实例缓存)。"""
    import common.ruyi.chain as chain
    with chain._HUB_LOCK:
        hub = chain._HUB
        chain._HUB = None
    if hub is not None:
        try:
            hub.stop()
        except Exception:
            pass


class ParseFrontProxyTests(unittest.TestCase):
    def test_parse_socks5_with_auth(self):
        p = parse_front_proxy("socks5://user:pass@1.2.3.4:10808")
        self.assertEqual(p["scheme"], "socks5")
        self.assertEqual(p["host"], "1.2.3.4")
        self.assertEqual(p["port"], 10808)
        self.assertEqual(p["username"], "user")
        self.assertEqual(p["password"], "pass")

    def test_parse_http_defaults_scheme_socks5_when_missing(self):
        p = parse_front_proxy("127.0.0.1:7897")
        self.assertEqual(p["scheme"], "socks5")
        self.assertEqual(p["port"], 7897)

    def test_parse_invalid_raises(self):
        for bad in ("", "   ", "http://", "socks5://:0"):
            with self.assertRaises(ValueError):
                parse_front_proxy(bad)


class FrontProxyRawTests(unittest.TestCase):
    def test_value_takes_priority(self):
        self.assertEqual(front_proxy_raw(" socks5://1.1.1.1:1 "), "socks5://1.1.1.1:1")

    def test_env_fallback(self):
        with patch.dict(os.environ, {"LAUNCH_FRONT_PROXY": FRONT}):
            self.assertEqual(front_proxy_raw(""), FRONT)
            self.assertEqual(front_proxy_raw(None), FRONT)

    def test_empty_when_unset(self):
        with patch.dict(os.environ, {"LAUNCH_FRONT_PROXY": ""}, clear=False):
            os.environ.pop("LAUNCH_FRONT_PROXY", None)
            self.assertEqual(front_proxy_raw(None), "")


class FrontRelayUrlTests(unittest.TestCase):
    def setUp(self):
        _reset_singleton()

    def tearDown(self):
        _reset_singleton()

    def test_front_empty_returns_upstream_unchanged(self):
        self.assertEqual(front_relay_url(PROXY, front=""), PROXY)
        self.assertEqual(front_relay_url(PROXY, front=None), PROXY)

    def test_front_set_maps_to_local_relay(self):
        url = front_relay_url(PROXY, front=FRONT, tag="[t]")
        self.assertTrue(url.startswith("socks5://127.0.0.1:"))
        self.assertNotIn(":", url[len("socks5://127.0.0.1:"):])  # 无 :u:p 后缀
        port = int(url.rsplit(":", 1)[1])
        self.assertGreater(port, 0)

    def test_front_set_per_tab_appends_credentials(self):
        url = front_relay_url(PROXY, front=FRONT, per_tab=True, tag="[t]")
        self.assertTrue(url.startswith("socks5://127.0.0.1:"))
        self.assertTrue(url.endswith(":u:p"), url)

    def test_same_upstream_maps_to_same_port(self):
        u1 = front_relay_url(PROXY, front=FRONT, tag="[t]")
        u2 = front_relay_url(PROXY, front=FRONT, tag="[t]")
        self.assertEqual(u1, u2)

    def test_invalid_upstream_raises(self):
        with self.assertRaises(ValueError):
            front_relay_url("socks5://", front=FRONT, tag="[t]")

    def test_invalid_front_raises(self):
        with self.assertRaises(ValueError):
            front_relay_url(PROXY, front="socks5://", tag="[t]")

    def test_relay_accepts_connection(self):
        """端到端:映射出的本地端口确实是监听中的 SOCKS5(可建 TCP 连接)。"""
        import socket
        url = front_relay_url(PROXY, front=FRONT, tag="[t]")
        port = int(url.rsplit(":", 1)[1])
        s = socket.create_connection(("127.0.0.1", port), timeout=3)
        try:
            s.sendall(b"\x05\x01\x00")  # SOCKS5 greet, no-auth
            resp = s.recv(2)
            self.assertEqual(resp, b"\x05\x00")
        finally:
            s.close()


class ChainRelayHubConcurrencyTests(unittest.TestCase):
    def test_port_for_concurrent_same_upstream_one_port(self):
        _reset_singleton()
        hub = ChainRelayHub(parse_front_proxy(FRONT))
        hub.start()
        try:
            up = parse_front_proxy(PROXY)
            import threading
            results = []

            def worker():
                results.append(hub.port_for(up))

            threads = [threading.Thread(target=worker) for _ in range(4)]
            for t in threads:
                t.start()
            for t in threads:
                t.join(15)
            self.assertEqual(len(set(results)), 1, results)
        finally:
            hub.stop()


class RegisterFrontProxyTests(unittest.TestCase):
    def test_argparse_has_front_proxy_default_env(self):
        import register_outlook_ruoyi as mod
        src_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        # argparse 在 main() 内,直接检查源码中的定义(默认值读 env)
        with open(os.path.join(src_root, "register_outlook_ruoyi.py"), encoding="utf-8") as f:
            src = f.read()
        self.assertIn('"--front-proxy"', src)
        self.assertIn('os.environ.get("LAUNCH_FRONT_PROXY"', src)

    def test_resolve_graph_auth_proxy_front_on_maps_relay(self):
        """front 开 + reg_proxy → 走本地中继 socks5h://127.0.0.1:{port}。"""
        _reset_singleton()
        args = SimpleNamespace(graph_auth_use_reg_proxy=True, front_proxy=FRONT)
        try:
            url = register_mod._resolve_graph_auth_proxy(args, PROXY, "[#1]")
            self.assertIsNotNone(url)
            self.assertTrue(url.startswith("socks5h://127.0.0.1:"))
            port = int(url.rsplit(":", 1)[1])
            self.assertGreater(port, 0)
        finally:
            _reset_singleton()

    def test_resolve_graph_auth_proxy_front_off_original_behavior(self):
        args = SimpleNamespace(graph_auth_use_reg_proxy=True, front_proxy="")
        url = register_mod._resolve_graph_auth_proxy(args, PROXY, "[#1]")
        self.assertIsNotNone(url)
        self.assertIn("1.2.3.4:1080", url)
        self.assertTrue(url.startswith("socks5h://"))


class UnlockFrontProxyTests(unittest.TestCase):
    def test_argparse_has_front_proxy_default_env(self):
        with open(os.path.join(
                os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                "unlock_outlook.py"), encoding="utf-8") as f:
            src = f.read()
        self.assertIn('"--front-proxy"', src)
        self.assertIn('os.environ.get("LAUNCH_FRONT_PROXY"', src)

    def test_launch_firefox_signature_has_front_proxy(self):
        import inspect
        sig = inspect.signature(unlock_mod.launch_firefox)
        self.assertIn("front_proxy", sig.parameters)
        self.assertEqual(sig.parameters["front_proxy"].default, "")


if __name__ == "__main__":
    unittest.main()
