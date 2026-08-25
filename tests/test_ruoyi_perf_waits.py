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




if __name__ == "__main__":
    unittest.main()
