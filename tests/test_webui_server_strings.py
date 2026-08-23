import unittest
from pathlib import Path

import webui.server as mod


class WebuiServerStringTests(unittest.TestCase):
    def test_user_facing_errors_are_not_placeholder_question_marks(self):
        text = Path(mod.__file__).read_text(encoding="utf-8")
        self.assertNotIn("????", text)
        self.assertNotIn("?????????", text)


if __name__ == "__main__":
    unittest.main()
