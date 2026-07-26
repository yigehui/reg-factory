import inspect
import time
import unittest
from unittest.mock import patch

import register_outlook_ruoyi as mod


class BirthdayEntryStepTests(unittest.TestCase):
    def test_resolve_birthday_entry_step_returns_birthday_immediately_when_controls_visible(self):
        with (
            patch.object(mod, "_is_name_page", return_value=False),
            patch.object(mod, "_is_birthday_page", return_value=True),
            patch.object(mod, "_wait_signup_step") as wait_signup_step,
        ):
            step = mod._resolve_birthday_entry_step(object(), max_wait=5.0)

        self.assertEqual(step, "birthday")
        wait_signup_step.assert_not_called()

    def test_resolve_birthday_entry_step_returns_name_immediately(self):
        with (
            patch.object(mod, "_is_name_page", return_value=True),
            patch.object(mod, "_is_birthday_page", return_value=False),
            patch.object(mod, "_wait_signup_step") as wait_signup_step,
        ):
            step = mod._resolve_birthday_entry_step(object(), max_wait=5.0)

        self.assertEqual(step, "name")
        wait_signup_step.assert_not_called()


class BrowserQuitTimeoutTests(unittest.TestCase):
    def test_quit_browser_page_falls_back_to_close_when_quit_hangs(self):
        class _FakeBrowser:
            def __init__(self):
                self.closed = 0

            def quit(self):
                time.sleep(0.2)

            def close(self):
                self.closed += 1

        browser = _FakeBrowser()
        start = time.perf_counter()
        ok = mod._quit_browser_page(browser, timeout=0.01)
        elapsed = time.perf_counter() - start

        self.assertFalse(ok)
        self.assertEqual(browser.closed, 1)
        self.assertLess(elapsed, 0.15)


class LeftSignupLogTests(unittest.TestCase):
    def test_register_outlook_uses_microsoft_loading_wording_for_left_signup_log(self):
        source = inspect.getsource(mod.register_outlook)

        self.assertIn("Microsoft Loading page, keep waiting for redirect", source)
        self.assertNotIn("left signup ->", source)


if __name__ == "__main__":
    unittest.main()
