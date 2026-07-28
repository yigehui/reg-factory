import unittest
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


class PostSignupSuccessDetectionTests(unittest.TestCase):
    def test_proofs_add_frame_counts_as_registration_complete(self):
        frame = type("Frame", (), {"url": "https://account.live.com/proofs/Add?mkt=zh-TW"})()
        page = type(
            "Page",
            (),
            {
                "url": "https://signup.live.com/signup?lic=1",
                "get_all_frames": lambda self: [frame],
            },
        )()

        self.assertTrue(mod._registration_completed(page))

    def test_any_url_change_after_captcha_counts_as_registration_complete(self):
        frame = type("Frame", (), {"url": "https://example.com/post-signup"})()
        page = type(
            "Page",
            (),
            {
                "url": "https://signup.live.com/signup?lic=1",
                "get_all_frames": lambda self: [frame],
            },
        )()

        self.assertTrue(mod._registration_completed(page, after_captcha=True))


class WaitStateTimeoutTests(unittest.TestCase):
    def test_wait_state_times_out_at_deadline(self):
        self.assertTrue(mod._wait_state_timed_out(100.0, now=120.0, timeout=20.0))

    def test_wait_state_ignores_missing_start(self):
        self.assertFalse(mod._wait_state_timed_out(None, now=120.0, timeout=20.0))


class MicrosoftLoadingGuardTests(unittest.TestCase):
    def test_microsoft_loading_timeout_defaults_to_30_seconds(self):
        self.assertEqual(30.0, mod.MICROSOFT_LOADING_TIMEOUT)

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
