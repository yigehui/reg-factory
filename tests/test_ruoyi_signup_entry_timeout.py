import unittest
from unittest.mock import patch

import register_outlook_ruoyi as mod


class EnsureSignupEntryTests(unittest.TestCase):
    def test_returns_signup_step_when_form_enters_in_time(self):
        with (
            patch.object(mod, "_handle_consent") as handle_consent,
            patch.object(mod, "_wait_signup_step", return_value="email"),
            patch.object(mod, "_detect_signup_step", return_value="email"),
            patch.object(mod, "log"),
            patch.object(mod, "_shot"),
            patch.object(mod.time, "time", side_effect=[100.0, 100.0]),
        ):
            step = mod._ensure_signup_entry(object(), "[#1][ruoyi]", 1, timeout=20)

        self.assertEqual(step, "email")
        handle_consent.assert_called_once()

    def test_returns_empty_when_form_never_enters_before_timeout(self):
        with (
            patch.object(mod, "_handle_consent") as handle_consent,
            patch.object(mod, "_wait_signup_step", return_value="unknown") as wait_signup_step,
            patch.object(mod, "_detect_signup_step", return_value="unknown"),
            patch.object(mod, "log"),
            patch.object(mod, "_shot") as shot,
            patch.object(mod.time, "time", side_effect=[100.0, 121.0]),
        ):
            step = mod._ensure_signup_entry(object(), "[#1][ruoyi]", 1, timeout=20)

        self.assertEqual(step, "")
        handle_consent.assert_called_once()
        shot.assert_called_once()
        self.assertEqual(wait_signup_step.call_args.kwargs["max_wait"], 0.0)


if __name__ == "__main__":
    unittest.main()
