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


class SummaryFormatTests(unittest.TestCase):
    def test_save_results_prints_summary_lines(self):
        """save_results 末尾打印 SUMMARY / SUMMARY_TIME,前缀与注册一致(webui 解析依赖)。"""
        results = [
            ("a@outlook.com", "p1", "a@outlook.com----p1", "unlocked"),
            ("b@outlook.com", "p2", "b@outlook.com----p2", "needs_phone"),
            ("c@outlook.com", "p3", "c@outlook.com----p3", "timeout"),
        ]
        buf = io.StringIO()
        with redirect_stdout(buf):
            mod.save_results(results, "20260826_120000")
        out = buf.getvalue()
        self.assertIn("SUMMARY: unlocked 1 | needs_phone 1 | failed 1 | total 3", out)
        self.assertRegex(out, r"SUMMARY_TIME: total_elapsed [\d.]+s")


if __name__ == "__main__":
    unittest.main()
