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

    def test_load_proxy_list_merges_file_and_aimili_list(self):
        args = SimpleNamespace(
            proxy_file="proxies.txt",
            aimili_url="http://127.0.0.1:8787",
            aimili_token="tok",
        )
        with (
            patch.object(mod, "parse_proxy_pool", return_value=["file1", "file2"]) as parse_proxy_pool,
            patch.object(mod, "fetch_aimili_proxy_list", return_value=["api1"]) as fetch_aimili_proxy_list,
        ):
            merged = mod.load_proxy_list(args)

        self.assertEqual(merged, ["file1", "file2", "api1"])
        parse_proxy_pool.assert_called_once_with("proxies.txt")
        fetch_aimili_proxy_list.assert_called_once_with("http://127.0.0.1:8787", "tok")

    def test_load_proxy_list_uses_file_only_when_aimili_missing(self):
        args = SimpleNamespace(
            proxy_file="proxies.txt",
            aimili_url="",
            aimili_token="",
        )
        with (
            patch.object(mod, "parse_proxy_pool", return_value=["file1"]) as parse_proxy_pool,
            patch.object(mod, "fetch_aimili_proxy_list") as fetch_aimili_proxy_list,
        ):
            merged = mod.load_proxy_list(args)

        self.assertEqual(merged, ["file1"])
        parse_proxy_pool.assert_called_once_with("proxies.txt")
        fetch_aimili_proxy_list.assert_not_called()

    def test_pick_user_agent_stays_unique_for_first_ten_slots(self):
        with patch.dict(mod.os.environ, {"OUTLOOK_RUOYI_UA_POOL": ""}, clear=False):
            picked = [mod._pick_user_agent(i) for i in range(1, 11)]

        self.assertEqual(len(set(picked)), 10)

    def test_consumable_pool_keeps_same_exit_ip_until_take_time(self):
        pool = mod.ConsumableProxyPool(SimpleNamespace(proxy_file="", aimili_url="", aimili_token=""))

        with (
            patch.object(mod, "load_proxy_list", return_value=["a:1:u:p", "b:2:u:p", "c:3:u:p"]),
            patch.object(mod, "_proxy_exit_key", side_effect=["ip:1.1.1.1", "ip:1.1.1.1", "ip:2.2.2.2"]),
        ):
            pool.start()

        self.assertEqual(pool.remaining(), 3)
        self.assertEqual(pool.stats()["active_exit_keys"], 0)

    def test_consumable_pool_avoids_active_exit_ip_reuse(self):
        pool = mod.ConsumableProxyPool(SimpleNamespace(proxy_file="", aimili_url="", aimili_token=""))
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


if __name__ == "__main__":
    unittest.main()
