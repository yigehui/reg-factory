import unittest
from unittest.mock import patch

import outlook_reg_loop as loop


class OutlookRegLoopTaskStoreTests(unittest.TestCase):
    def test_persist_success_account_records_webui_account(self):
        graph = {"refresh_token": "rt-1", "client_id": "cid-1"}
        record = {"email": "a@outlook.com"}

        with (
            patch.object(loop, "write_record", return_value="rec.json") as write_record,
            patch.object(loop, "append_to_emails_pool") as append_to_emails_pool,
            patch.object(loop, "_record_webui_account") as record_webui_account,
        ):
            fname = loop._persist_success_account(
                email="a@outlook.com",
                password="Pass123!",
                graph=graph,
                record=record,
            )

        self.assertEqual("rec.json", fname)
        write_record.assert_called_once_with(record)
        append_to_emails_pool.assert_called_once_with("a@outlook.com", "Pass123!")
        record_webui_account.assert_called_once_with("a@outlook.com", "Pass123!", graph, "ok")


if __name__ == "__main__":
    unittest.main()
