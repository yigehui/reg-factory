import asyncio
import unittest
from types import SimpleNamespace
from unittest.mock import patch

import register_outlook_ruoyi as mod


PROXY = "socks5://u:p@1.2.3.4:1080"


class GraphAuthRegProxyTests(unittest.IsolatedAsyncioTestCase):
    def test_resolve_returns_none_when_option_off(self):
        args = SimpleNamespace(graph_auth_use_reg_proxy=False)
        self.assertIsNone(mod._resolve_graph_auth_proxy(args, PROXY, "[#1]"))

    def test_resolve_returns_none_when_no_proxy(self):
        args = SimpleNamespace(graph_auth_use_reg_proxy=True)
        self.assertIsNone(mod._resolve_graph_auth_proxy(args, "", "[#1]"))
        self.assertIsNone(mod._resolve_graph_auth_proxy(args, None, "[#1]"))

    def test_resolve_returns_canonical_url_when_option_on(self):
        args = SimpleNamespace(graph_auth_use_reg_proxy=True)
        url = mod._resolve_graph_auth_proxy(args, PROXY, "[#1]")
        self.assertIsNotNone(url)
        self.assertIn("1.2.3.4:1080", url)
        self.assertTrue(url.startswith("socks5h://"))

    async def test_run_one_direct_deferred_carries_reg_proxy(self):
        args = SimpleNamespace(
            timeout=300, live_file="x", token_file="y", graph_auth_use_reg_proxy=True
        )
        helpers = SimpleNamespace(extract_graph_token_http=lambda *a, **k: None)
        save_lock = asyncio.Lock()

        def fake_register(opts, _pool, _idx):
            return "a@outlook.com", "Pass1!", "", {}

        with (
            patch.object(mod, "select_proxy_for_account", return_value=[PROXY]),
            patch.object(mod, "_probe_proxy_before_browser", return_value=True, create=True),
            patch.object(mod, "register_outlook", side_effect=fake_register),
            patch.object(mod, "release_proxy_for_account"),
        ):
            result = await mod._run_one_direct(
                args, helpers, proxy_pool=[], idx=1, total=1, save_lock=save_lock,
                consumable_pool=None, defer_graph_auth=True,
            )

        self.assertEqual(result["status"], "registered")
        self.assertEqual(result["reg_proxy"], PROXY)


if __name__ == "__main__":
    unittest.main()
