import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import register_outlook_ruoyi as mod


class RuoyiProxyUnifiedTests(unittest.TestCase):
    def test_invalid_outlook_domain_falls_back_to_outlook_com(self):
        self.assertEqual(mod._email_domain("user@outlook.co"), "outlook.com")
        self.assertEqual(mod._email_domain("user@hotmail.com"), "hotmail.com")

    def test_parse_proxy_lines_strips_utf8_bom(self):
        parsed = mod._parse_proxy_lines(["\ufeff1.2.3.4:1080"], "http://test/proxies.txt")

        self.assertEqual(parsed, ["socks5://1.2.3.4:1080"])

    def test_parse_proxy_lines_default_scheme_is_http(self):
        parsed = mod._parse_proxy_lines(["107.151.249.13:21429"], "http://test/proxies.txt")

        self.assertEqual(parsed, ["socks5://107.151.249.13:21429"])

    def test_parse_proxy_lines_preserves_explicit_scheme(self):
        parsed = mod._parse_proxy_lines(["socks5://107.151.249.13:21429"], "http://test/proxies.txt")

        self.assertEqual(parsed, ["socks5://107.151.249.13:21429"])

    def test_load_proxy_list_uses_http_source_when_selected(self):
        args = SimpleNamespace(
            proxy_source="http",
            proxy_url="http://127.0.0.1:8787/proxies",
            proxy_file="proxies.txt",
        )
        with (
            patch.object(mod, "parse_proxy_pool") as parse_proxy_pool,
            patch.object(mod, "fetch_proxy_list_http", return_value=["api1"]) as fetch_proxy_list_http,
        ):
            merged = mod.load_proxy_list(args)

        self.assertEqual(merged, ["api1"])
        parse_proxy_pool.assert_not_called()
        fetch_proxy_list_http.assert_called_once_with("http://127.0.0.1:8787/proxies")

    def test_load_proxy_list_uses_file_source_by_default(self):
        args = SimpleNamespace(
            proxy_source="file",
            proxy_file="proxies.txt",
        )
        with (
            patch.object(mod, "parse_proxy_pool", return_value=["file1"]) as parse_proxy_pool,
            patch.object(mod, "fetch_proxy_list_http") as fetch_proxy_list_http,
        ):
            merged = mod.load_proxy_list(args)

        self.assertEqual(merged, ["file1"])
        parse_proxy_pool.assert_called_once_with("proxies.txt")
        fetch_proxy_list_http.assert_not_called()

    def test_load_proxy_list_falls_back_to_http_when_file_empty(self):
        args = SimpleNamespace(
            proxy_source="file",
            proxy_file="missing.txt",
            proxy_url="http://127.0.0.1:8787/proxies",
        )
        with (
            patch.object(mod, "parse_proxy_pool", return_value=[]) as parse_proxy_pool,
            patch.object(mod, "fetch_proxy_list_http", return_value=["api1"]) as fetch_proxy_list_http,
        ):
            merged = mod.load_proxy_list(args)

        self.assertEqual(merged, ["api1"])
        parse_proxy_pool.assert_called_once_with("missing.txt")
        fetch_proxy_list_http.assert_called_once_with("http://127.0.0.1:8787/proxies")

    def test_fetch_proxy_list_http_rejects_non_http_scheme(self):
        with self.assertRaisesRegex(RuntimeError, "仅支持 http/https"):
            mod.fetch_proxy_list_http("file:///tmp/proxies.txt")

    def test_pick_user_agent_stays_unique_for_first_ten_slots(self):
        with patch.dict(mod.os.environ, {"OUTLOOK_RUOYI_UA_POOL": ""}, clear=False):
            picked = [mod._pick_user_agent(i) for i in range(1, 11)]

        self.assertEqual(len(set(picked)), 10)

    def test_consumable_pool_keeps_same_exit_ip_until_take_time(self):
        pool = mod.ConsumableProxyPool(SimpleNamespace(proxy_file="", proxy_source="file", proxy_url=""))

        with (
            patch.object(mod, "load_proxy_list", return_value=["a:1:u:p", "b:2:u:p", "c:3:u:p"]),
            patch.object(mod, "_proxy_exit_key", side_effect=["ip:1.1.1.1", "ip:1.1.1.1", "ip:2.2.2.2"]),
        ):
            pool.start()

        self.assertEqual(pool.remaining(), 3)
        self.assertEqual(pool.stats()["active_exit_keys"], 0)

    def test_consumable_pool_avoids_active_exit_ip_reuse(self):
        pool = mod.ConsumableProxyPool(SimpleNamespace(proxy_file="", proxy_source="file", proxy_url=""))
        pool._list = ["proxy1", "proxy2", "proxy3"]

        def fake_exit_key(proxy, timeout=mod.PROXY_IDENTITY_TIMEOUT, use_cache=True):
            return {"proxy1": "ip:1.1.1.1", "proxy2": "ip:1.1.1.1", "proxy3": "ip:2.2.2.2"}[proxy]

        with (
            patch.object(mod, "_proxy_exit_key", side_effect=fake_exit_key),
            patch.object(mod.random, "randrange", side_effect=[0, 0, 0, 0]),
        ):
            first = pool.take()
            second = pool.take()
            pool.release("proxy1")
            third = pool.take()

        self.assertEqual(first, ["proxy1"])
        self.assertEqual(second, ["proxy3"])
        self.assertEqual(third, ["proxy2"])


    def test_cleanup_ruoyi_profile_root_removes_cached_items_only(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "slot_01").mkdir()
            (root / "slot_01" / "prefs.js").write_text("x", encoding="utf-8")
            (root / "slot_02").mkdir()
            (root / "singleton.lock").write_text("lock", encoding="utf-8")

            cleaned = mod._cleanup_ruoyi_profile_root(str(root))

            self.assertEqual(cleaned, 3)
            self.assertTrue(root.exists())
            self.assertEqual(list(root.iterdir()), [])

    def test_ruoyi_profile_dir_uses_unique_attempt_dir_per_slot(self):
        with tempfile.TemporaryDirectory() as tmp, patch.object(mod, "RUOYI_PROFILE_ROOT", tmp):
            opts = SimpleNamespace(ruoyi_slot=2, concurrency=4)

            first = Path(mod._ruoyi_profile_dir(opts, 1))
            self.assertTrue(first.exists())
            second = Path(mod._ruoyi_profile_dir(opts, 2))

            self.assertNotEqual(first, second)
            self.assertEqual(first.parent.name, "slot_02")
            self.assertEqual(second.parent.name, "slot_02")
            self.assertTrue(first.exists())
            self.assertTrue(second.exists())

    def test_ruoyi_profile_dir_does_not_delete_existing_slot_runs(self):
        with tempfile.TemporaryDirectory() as tmp, patch.object(mod, "RUOYI_PROFILE_ROOT", tmp):
            slot = Path(tmp) / "slot_01"
            slot.mkdir()
            existing = slot / "run_active"
            existing.mkdir()
            (existing / "prefs.js").write_text("x", encoding="utf-8")

            created = Path(mod._ruoyi_profile_dir(SimpleNamespace(ruoyi_slot=1, concurrency=1), 7))

            self.assertTrue(existing.exists())
            self.assertTrue(created.exists())
            self.assertEqual(created.parent, slot)

    def test_ruoyi_should_block_resource_request_uses_sec_fetch_dest(self):
        req = SimpleNamespace(
            url="https://example.com/assets/logo",
            headers={"Sec-Fetch-Dest": "image"},
        )

        self.assertTrue(mod._ruoyi_should_block_resource_request(req))

    def test_ruoyi_should_block_resource_request_uses_accept_and_suffix(self):
        req = SimpleNamespace(
            url="https://example.com/static/font.woff2?v=1",
            headers={"Accept": "*/*"},
        )

        self.assertTrue(mod._ruoyi_should_block_resource_request(req))

    def test_ruoyi_should_block_resource_request_skips_allow_hosts(self):
        req = SimpleNamespace(
            url="https://client.px-cloud.net/assets/hold.png",
            headers={"Sec-Fetch-Dest": "image"},
        )

        self.assertFalse(mod._ruoyi_should_block_resource_request(req))


if __name__ == "__main__":
    unittest.main()
