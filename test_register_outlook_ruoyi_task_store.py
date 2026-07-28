import unittest
from contextlib import nullcontext
from unittest.mock import mock_open
from unittest.mock import patch

import register_outlook_ruoyi as ruoyi
import task_store


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
            ruoyi._save_direct_result(
                "a@outlook.com",
                "Pass123!",
                graph,
                "live.txt",
                "",
                register_ip="1.2.3.4",
                register_region="Japan",
            )

        append_pool.assert_called_once_with("a@outlook.com", "Pass123!", graph)
        record_webui_account.assert_called_once_with(
            "a@outlook.com",
            "Pass123!",
            graph,
            "ok",
            register_ip="1.2.3.4",
            register_region="Japan",
        )

    def test_save_no_graph_result_records_webui_account(self):
        with (
            patch.object(ruoyi, "append_account_to_email_nograph", return_value=True) as append_nograph,
            patch.object(ruoyi, "_record_webui_account") as record_webui_account,
        ):
            ruoyi._save_no_graph_result(
                "a@outlook.com",
                "Pass123!",
                register_ip="1.2.3.4",
                register_region="Japan",
            )

        append_nograph.assert_called_once_with("a@outlook.com", "Pass123!")
        record_webui_account.assert_called_once_with(
            "a@outlook.com",
            "Pass123!",
            None,
            "no_graph",
            register_ip="1.2.3.4",
            register_region="Japan",
        )

    def test_record_webui_account_saves_register_ip_and_region(self):
        captured = {}

        def _capture(**kwargs):
            captured.update(kwargs)

        with (
            patch.dict(
                ruoyi.os.environ,
                {
                    "WEBUI_TASK_RUN_ID": "run-1",
                    "WEBUI_MYSQL_HOST": "127.0.0.1",
                    "WEBUI_MYSQL_USER": "root",
                    "WEBUI_MYSQL_DATABASE": "db1",
                },
                clear=False,
            ),
            patch.object(task_store, "add_task_account", side_effect=_capture),
        ):
            ruoyi._record_webui_account(
                "a@outlook.com",
                "Pass123!",
                {"refresh_token": "rt-1", "client_id": "cid-1"},
                "ok",
                register_ip="1.2.3.4",
                register_region="Japan",
            )

        self.assertEqual("1.2.3.4", captured["register_ip"])
        self.assertEqual("Japan", captured["register_region"])


if __name__ == "__main__":
    unittest.main()
