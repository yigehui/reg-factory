import unittest
from pathlib import Path


class WebuiIndexViewsTests(unittest.TestCase):
    def test_index_contains_task_and_account_views(self):
        text = Path("webui/static/index.html").read_text(encoding="utf-8")
        self.assertIn('id="view-tasks"', text)
        self.assertIn('id="view-accounts"', text)

    def test_index_contains_account_controls_and_modals(self):
        text = Path("webui/static/index.html").read_text(encoding="utf-8")
        self.assertIn('id="accounts-import-modal"', text)
        self.assertIn('id="accounts-import-file"', text)
        self.assertIn('id="accounts-email-filter"', text)
        self.assertIn('id="accounts-status-filter"', text)
        self.assertIn('id="btn-query-accounts"', text)
        self.assertIn('id="accounts-page-size"', text)
        self.assertIn('id="accounts-pagination"', text)
        self.assertIn('id="tasks-page-size"', text)
        self.assertIn('id="tasks-pagination"', text)
        self.assertIn('id="task-detail-accounts-pagination"', text)
        self.assertIn('id="accounts-edit-modal"', text)
        self.assertIn('id="btn-save-edit-accounts"', text)
        self.assertIn('id="task-log-modal"', text)
        self.assertIn('id="task-log-content"', text)
        self.assertNotIn('id="task-detail-log"', text)


if __name__ == "__main__":
    unittest.main()
