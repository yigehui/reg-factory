import asyncio
import time
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

import register_outlook_ruoyi as mod


class RuoyiProxyPrecheckTests(unittest.IsolatedAsyncioTestCase):
    def test_proxy_precheck_uses_20_second_default_timeout(self):
        session = Mock()
        session.get.return_value = SimpleNamespace(status_code=200)

        with (
            patch.object(mod, "_proxy_for_ip_lookup", return_value={"http": "socks5h://127.0.0.1:1080"}),
            patch.object(mod.requests, "Session", return_value=session),
        ):
            ok = mod._probe_proxy_before_browser(["127.0.0.1:1080"], "[#1][ruoyi]")

        self.assertTrue(ok)
        self.assertEqual(session.get.call_args.kwargs["timeout"], 20.0)

    async def test_run_one_direct_skips_browser_when_proxy_precheck_fails(self):
        args = SimpleNamespace(live_file="emails.txt", token_file="tokens.json")
        helpers = SimpleNamespace(extract_graph_token_http=lambda *a, **k: None)
        save_lock = asyncio.Lock()

        with (
            patch.object(mod, "select_proxy_for_account", return_value=["127.0.0.1:1080"]),
            patch.object(mod, "_probe_proxy_before_browser", return_value=False, create=True),
            patch.object(mod, "register_outlook", return_value=(None, None, "should_not_run")) as register_outlook,
        ):
            status, elapsed, px_metrics = await mod._run_one_direct(
                args,
                helpers,
                proxy_pool=[],
                idx=1,
                total=1,
                save_lock=save_lock,
                consumable_pool=None,
            )

        self.assertEqual(status, "fail")
        self.assertGreaterEqual(elapsed, 0.0)
        self.assertEqual(px_metrics["max_presses"], 0)
        self.assertAlmostEqual(px_metrics["px_elapsed"], 0.0)
        register_outlook.assert_not_called()

    async def test_run_one_direct_passes_remaining_budget_to_register(self):
        args = SimpleNamespace(timeout=150, live_file="emails.txt", token_file="tokens.json")
        helpers = SimpleNamespace(extract_graph_token_http=lambda *a, **k: None)
        save_lock = asyncio.Lock()
        seen_timeouts = []

        def fake_register(opts, _proxy_pool, _idx):
            seen_timeouts.append(opts.timeout)
            return None, None, "timeout", {}

        with (
            patch.object(mod, "select_proxy_for_account", return_value=[]),
            patch.object(mod, "register_outlook", side_effect=fake_register),
            patch.object(mod.time, "perf_counter", side_effect=[100.0, 220.0, 221.0]),
        ):
            status, elapsed, _px_metrics = await mod._run_one_direct(
                args,
                helpers,
                proxy_pool=[],
                idx=1,
                total=1,
                save_lock=save_lock,
                consumable_pool=None,
            )

        self.assertEqual(status, "fail")
        self.assertEqual(seen_timeouts, [30.0])
        self.assertEqual(elapsed, 121.0)


    def test_px_detail_includes_registration_elapsed(self):
        metrics = mod._normalize_px_metrics(3, {"idx": 3, "max_presses": 2, "px_elapsed": 5.0, "reg_elapsed": 42.25})

        lines = mod._format_batch_summary_lines(1, 0, 0, 1, 42.25, 42.25, [metrics])

        self.assertIn("PX_DETAIL: #3 max_presses 2 | px_elapsed 5.00s | reg_elapsed 42.25s", lines)

    def test_handle_consent_uses_remaining_deadline_for_locator_timeout(self):
        page = SimpleNamespace(url="https://login.live.com/privacynotice")
        timeouts = []

        def fake_click_any(_page, _locators, timeout=0):
            timeouts.append(timeout)
            return None

        with (
            patch.object(mod, "_body_text", return_value="consent"),
            patch.object(mod, "_on_signup_form", return_value=False),
            patch.object(mod, "_click_any", side_effect=fake_click_any),
            patch.object(mod, "_shot"),
            patch.object(mod.time, "sleep"),
        ):
            deadline = time.time() + 0.05
            mod._handle_consent(page, "[#1][ruoyi]", 1, deadline=deadline)

        self.assertTrue(timeouts)
        self.assertLessEqual(max(timeouts), 0.05)


if __name__ == "__main__":
    unittest.main()
