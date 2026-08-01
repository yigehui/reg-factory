import unittest
from types import SimpleNamespace
from unittest.mock import patch

import register_outlook_ruoyi as mod


class CaptchaDetectionTests(unittest.TestCase):
    def test_generic_please_wait_without_px_context_is_not_validating(self):
        with (
            patch.object(mod, "_px_captcha_completed_wait", return_value=False),
            patch.object(mod, "_body_text", return_value="Loading... Please wait while we continue."),
            patch.object(mod, "_context_has_iframe_hint", return_value=False),
        ):
            self.assertFalse(mod._captcha_is_validating(object()))


class WaitStateTimeoutTests(unittest.TestCase):
    def test_wait_state_times_out_at_deadline(self):
        self.assertTrue(mod._wait_state_timed_out(100.0, now=120.0, timeout=20.0))

    def test_wait_state_ignores_missing_start(self):
        self.assertFalse(mod._wait_state_timed_out(None, now=120.0, timeout=20.0))

    def test_post_press_reappear_wait_still_runs_after_last_allowed_press(self):
        self.assertTrue(mod._should_enter_post_press_reappear_wait(True, 3, 3))

    def test_post_press_reappear_wait_skips_when_not_awaiting_result(self):
        self.assertFalse(mod._should_enter_post_press_reappear_wait(False, 3, 3))


class CaptchaSuccessUrlTests(unittest.TestCase):
    def test_captcha_url_change_treats_signup_transition_as_success(self):
        self.assertTrue(
            mod._captcha_signup_url_changed(
                "https://signup.live.com/signup?lic=1",
                "https://signup.live.com/?lic=1&uaid=next-step",
            )
        )

    def test_captcha_url_change_ignores_same_signup_url(self):
        self.assertFalse(
            mod._captcha_signup_url_changed(
                "https://signup.live.com/signup?lic=1",
                "https://signup.live.com/signup?lic=1",
            )
        )


class UserAgentPoolTests(unittest.TestCase):
    def test_default_ua_pool_has_ten_entries(self):
        with patch.dict(mod.os.environ, {"OUTLOOK_RUOYI_UA_POOL": ""}, clear=False):
            self.assertEqual(10, len(mod._load_ua_pool()))


class LoadingTimeoutFallbackTests(unittest.TestCase):
    def test_loading_timeout_graph_fallback_saves_graph_on_success(self):
        opts = SimpleNamespace()
        helpers = SimpleNamespace(
            extract_graph_token_http=lambda email, password, idx, retries, proxy: {
                "refresh_token": "rt",
                "client_id": "cid",
            }
        )

        self.assertTrue(
            mod._loading_timeout_graph_fallback(
                helpers,
                opts,
                "foo@outlook.com",
                "Pass1!",
                9,
                "[#9][ruoyi]",
            )
        )
        self.assertEqual("rt", opts._ruoyi_graph_fallback["refresh_token"])

    def test_loading_timeout_graph_fallback_fails_without_refresh_token(self):
        opts = SimpleNamespace()
        helpers = SimpleNamespace(
            extract_graph_token_http=lambda email, password, idx, retries, proxy: {
                "refresh_token": "",
            }
        )

        self.assertFalse(
            mod._loading_timeout_graph_fallback(
                helpers,
                opts,
                "foo@outlook.com",
                "Pass1!",
                9,
                "[#9][ruoyi]",
            )
        )
        self.assertFalse(hasattr(opts, "_ruoyi_graph_fallback"))


class MicrosoftLoadingGuardTests(unittest.TestCase):
    def test_loading_pauses_max_press_deadline_and_starts_loading_timer(self):
        press_wait, loading_wait, timed_out = mod._update_loading_wait_state(
            100.0,
            None,
            loading=True,
            now=102.0,
            timeout=20.0,
        )
        self.assertIsNone(press_wait)
        self.assertEqual(102.0, loading_wait)
        self.assertFalse(timed_out)

    def test_loading_keeps_existing_loading_timer_until_timeout(self):
        press_wait, loading_wait, timed_out = mod._update_loading_wait_state(
            None,
            102.0,
            loading=True,
            now=115.0,
            timeout=20.0,
        )
        self.assertIsNone(press_wait)
        self.assertEqual(102.0, loading_wait)
        self.assertFalse(timed_out)

    def test_loading_times_out_with_independent_deadline(self):
        press_wait, loading_wait, timed_out = mod._update_loading_wait_state(
            None,
            102.0,
            loading=True,
            now=123.0,
            timeout=20.0,
        )
        self.assertIsNone(press_wait)
        self.assertEqual(102.0, loading_wait)
        self.assertTrue(timed_out)


class SubmitClickLoggingTests(unittest.TestCase):
    def test_same_submit_selector_is_logged_once_until_state_resets(self):
        logs = []

        with (
            patch.object(mod, "_submit_wait"),
            patch.object(
                mod,
                "_click_any",
                side_effect=[
                    'css:button[type="submit"]',
                    'css:button[type="submit"]',
                    'css:button[type="submit"]',
                ],
            ),
            patch.object(mod, "log", side_effect=lambda msg, *_args: logs.append(msg)),
        ):
            submitted, last_hit = mod._try_submit(object(), "[#13][ruoyi]", None)
            self.assertTrue(submitted)
            self.assertEqual('css:button[type="submit"]', last_hit)
            self.assertEqual(1, len(logs))

            submitted, last_hit = mod._try_submit(object(), "[#13][ruoyi]", last_hit)
            self.assertTrue(submitted)
            self.assertEqual('css:button[type="submit"]', last_hit)
            self.assertEqual(1, len(logs))

            last_hit, _ = mod._update_submit_wait_state(last_hit, visible=True, now=200.0)
            self.assertIsNone(last_hit)

            submitted, last_hit = mod._try_submit(object(), "[#13][ruoyi]", last_hit)
            self.assertTrue(submitted)
            self.assertEqual(2, len(logs))


if __name__ == "__main__":
    unittest.main()
