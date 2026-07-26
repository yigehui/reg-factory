import unittest
from contextlib import nullcontext
from unittest.mock import mock_open
from unittest.mock import patch

import register_outlook_ruoyi as ruoyi


class RegisterOutlookRuoyiTaskStoreTests(unittest.TestCase):
    def test_save_direct_result_records_webui_account(self):
        graph = {"refresh_token": "rt-1", "client_id": "cid-1"}

        with (
            patch.object(ruoyi, "OUTPUT_DIR", "."),
            patch.object(ruoyi, "_interprocess_lock", side_effect=lambda *_args, **_kwargs: nullcontext()),
            patch("builtins.open", mock_open(), create=True),
            patch.object(ruoyi.os.path, "isfile", return_value=False),
            patch.object(ruoyi, "append_graph_account_to_emails_pool") as append_pool,
            patch.object(ruoyi, "_record_webui_account") as record_webui_account,
        ):
            ruoyi._save_direct_result("a@outlook.com", "Pass123!", graph, "live.txt", "")

        append_pool.assert_called_once_with("a@outlook.com", "Pass123!", graph)
        record_webui_account.assert_called_once_with("a@outlook.com", "Pass123!", graph, "ok")

    def test_save_no_graph_result_records_webui_account(self):
        with (
            patch.object(ruoyi, "append_account_to_email_nograph", return_value=True) as append_nograph,
            patch.object(ruoyi, "_record_webui_account") as record_webui_account,
        ):
            ruoyi._save_no_graph_result("a@outlook.com", "Pass123!")

        append_nograph.assert_called_once_with("a@outlook.com", "Pass123!")
        record_webui_account.assert_called_once_with("a@outlook.com", "Pass123!", None, "no_graph")


if __name__ == "__main__":
    unittest.main()
