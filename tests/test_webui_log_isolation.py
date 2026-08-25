import unittest
from pathlib import Path


APP_JS = Path("webui/static/app.js").read_text(encoding="utf-8")


class WebuiLogIsolationTests(unittest.TestCase):
    def test_app_js_uses_script_logs_state(self):
        """前端按 script_id 聚合日志,不再单全局 curRun 覆盖。"""
        self.assertIn("scriptLogs", APP_JS)
        self.assertIn("curLogScript", APP_JS)

    def test_run_script_does_not_blank_log_area(self):
        """runScript 不再用 log.textContent='' 清空日志区(防回退到覆盖行为)。"""
        self.assertNotIn("log.textContent=''", APP_JS)
        self.assertNotIn("log.textContent = ''", APP_JS)

    def test_select_script_switches_log_view(self):
        """selectScript 切菜单时把日志区切到该脚本。"""
        self.assertIn("curLogScript", APP_JS)


if __name__ == "__main__":
    unittest.main()
