import unittest
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


if __name__ == "__main__":
    unittest.main()
