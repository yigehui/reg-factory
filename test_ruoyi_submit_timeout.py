import unittest

import register_outlook_ruoyi as mod


class SubmitTimeoutTests(unittest.TestCase):
    def test_first_submit_starts_wait_window(self):
        started_at, timed_out = mod._update_submit_wait_state(
            None,
            submitted=True,
            now=100.0,
        )

        self.assertEqual(started_at, 100.0)
        self.assertFalse(timed_out)

    def test_repeat_submit_does_not_reset_wait_window(self):
        started_at, timed_out = mod._update_submit_wait_state(
            100.0,
            submitted=True,
            now=110.0,
        )

        self.assertEqual(started_at, 100.0)
        self.assertFalse(timed_out)

    def test_no_transition_for_15s_times_out(self):
        started_at, timed_out = mod._update_submit_wait_state(
            100.0,
            now=115.0,
        )

        self.assertEqual(started_at, 100.0)
        self.assertTrue(timed_out)

    def test_captcha_or_loading_transition_clears_wait_window(self):
        for kwargs in (
            {"visible": True},
            {"validating": True},
            {"loading": True},
        ):
            with self.subTest(kwargs=kwargs):
                started_at, timed_out = mod._update_submit_wait_state(
                    100.0,
                    now=200.0,
                    **kwargs,
                )

                self.assertIsNone(started_at)
                self.assertFalse(timed_out)


if __name__ == "__main__":
    unittest.main()
