import unittest
from unittest.mock import patch

import register_outlook_ruoyi as mod


class PxPressScreenshotTests(unittest.TestCase):
    def test_px_failure_shots_require_toggle(self):
        with patch.object(mod, "_env_bool", return_value=False):
            self.assertFalse(mod._should_save_failure_shot("press_fail"))
            self.assertFalse(mod._should_save_failure_shot("captcha_no_target"))
            self.assertFalse(mod._should_save_failure_shot("timeout"))

        with patch.object(mod, "_env_bool", return_value=True):
            self.assertTrue(mod._should_save_failure_shot("press_fail"))
            self.assertTrue(mod._should_save_failure_shot("captcha_no_target"))
            self.assertTrue(mod._should_save_failure_shot("timeout"))

    def test_all_failure_shots_require_toggle(self):
        with patch.object(mod, "_env_bool", return_value=False):
            self.assertFalse(mod._should_save_failure_shot("email_fail"))
            self.assertFalse(mod._should_save_failure_shot("blocked"))

        with patch.object(mod, "_env_bool", return_value=True):
            self.assertTrue(mod._should_save_failure_shot("email_fail"))
            self.assertTrue(mod._should_save_failure_shot("blocked"))

    def test_disabled_does_not_save_screenshots(self):
        with (
            patch.object(mod, "_perform_hold", return_value=True) as perform_hold,
            patch.object(mod, "_save_screenshot") as save_screenshot,
        ):
            ok = mod._perform_hold_with_px_screenshots(
                "page", "ctx", "target", 1, 3, "[#1][ruoyi]", enabled=False
            )

        self.assertTrue(ok)
        perform_hold.assert_called_once_with("page", "ctx", "target", 1, 3, "[#1][ruoyi]")
        save_screenshot.assert_not_called()

    def test_enabled_success_saves_only_last_before_after_names(self):
        with (
            patch.object(mod, "_perform_hold", return_value=True),
            patch.object(mod, "_save_screenshot") as save_screenshot,
        ):
            ok = mod._perform_hold_with_px_screenshots(
                "page", "ctx", "target", 1, 3, "[#1][ruoyi]", enabled=True
            )

        self.assertTrue(ok)
        self.assertEqual(
            save_screenshot.call_args_list,
            [
                unittest.mock.call("page", "before_press_last", 1, "[#1][ruoyi]"),
                unittest.mock.call("page", "after_press_last", 1, "[#1][ruoyi]"),
            ],
        )

    def test_enabled_failure_still_saves_last_before_after_names(self):
        with (
            patch.object(mod, "_perform_hold", return_value=False),
            patch.object(mod, "_save_screenshot") as save_screenshot,
        ):
            ok = mod._perform_hold_with_px_screenshots(
                "page", "ctx", "target", 1, 4, "[#1][ruoyi]", enabled=True
            )

        self.assertFalse(ok)
        self.assertEqual(
            save_screenshot.call_args_list,
            [
                unittest.mock.call("page", "before_press_last", 1, "[#1][ruoyi]"),
                unittest.mock.call("page", "after_press_last", 1, "[#1][ruoyi]"),
            ],
        )


if __name__ == "__main__":
    unittest.main()
