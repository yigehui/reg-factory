import io
import re
import unittest
from contextlib import redirect_stdout

import unlock_outlook as mod


class LogFormatTests(unittest.TestCase):
    def test_log_has_timestamp_level_message(self):
        buf = io.StringIO()
        with redirect_stdout(buf):
            mod.log("step login: 3.2s", "INFO")
        out = buf.getvalue().strip()
        self.assertRegex(
            out,
            r"^\[\d{2}:\d{2}:\d{2}\] \[INFO\] step login: 3\.2s$",
        )

    def test_log_default_level_is_info(self):
        buf = io.StringIO()
        with redirect_stdout(buf):
            mod.log("hello")
        out = buf.getvalue().strip()
        self.assertRegex(out, r"^\[\d{2}:\d{2}:\d{2}\] \[INFO\] hello$")

    def test_log_warn_level(self):
        buf = io.StringIO()
        with redirect_stdout(buf):
            mod.log("凭证错误,跳过", "WARN")
        out = buf.getvalue().strip()
        self.assertRegex(out, r"^\[\d{2}:\d{2}:\d{2}\] \[WARN\] 凭证错误,跳过$")


if __name__ == "__main__":
    unittest.main()
