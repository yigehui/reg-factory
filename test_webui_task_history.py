import tempfile
import unittest
from pathlib import Path
from unittest.mock import AsyncMock, patch

from fastapi.testclient import TestClient

import webui.server as mod


class _FakeTaskStore:
    def __init__(self, run=None, accounts=None, paged=None, runs_paged=None):
        self._run = run or {}
        self._accounts = accounts or []
        self._paged = paged
        self._runs_paged = runs_paged
        self.last_list_kwargs = None
        self.last_run_list_kwargs = None
        self.last_update = None
        self.last_delete_task_id = None
        self.last_finish_kwargs = None

    def get_task_run(self, task_id):
        item = dict(self._run)
        item.setdefault("id", task_id)
        return item

    def list_task_runs(self, **kwargs):
        self.last_run_list_kwargs = dict(kwargs)
        if self._runs_paged is not None:
            return dict(self._runs_paged)
        return {"items": [], "total": 0, "page": 1, "page_size": 20}

    def list_accounts(self, **kwargs):
        self.last_list_kwargs = dict(kwargs)
        if self._paged is not None:
            return dict(self._paged)
        return list(self._accounts)

    def update_account(self, account_id, fields):
        self.last_update = {"account_id": account_id, "fields": dict(fields)}
        return {"id": account_id, **fields}

    def delete_task_run(self, task_id):
        self.last_delete_task_id = task_id
        return True

    def finish_task_run(self, **kwargs):
        self.last_finish_kwargs = dict(kwargs)
        return True


class WebuiTaskHistoryTests(unittest.TestCase):
    def setUp(self):
        self.client = TestClient(mod.app)
        self.runs_backup = mod.RUNS
        mod.RUNS = {}
        self.task_store_backup = getattr(mod, "task_store", None)

    def tearDown(self):
        mod.RUNS = self.runs_backup
        if self.task_store_backup is None and hasattr(mod, "task_store"):
            delattr(mod, "task_store")
        else:
            mod.task_store = self.task_store_backup

    def test_api_run_rejects_when_another_run_active(self):
        mod.RUNS["r1"] = {"done": False}
        with patch.object(mod.asyncio, "create_subprocess_exec", AsyncMock()) as create_proc:
            response = self.client.post(
                "/api/run",
                json={"script": "register_outlook_ruoyi", "args": {"count": 1}},
            )

        self.assertEqual(409, response.status_code)
        self.assertIn("\u5df2\u6709\u8fd0\u884c\u4e2d\u4efb\u52a1", response.json()["error"])
        create_proc.assert_not_awaited()

    def test_append_run_line_persists_log_file(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "r1.txt"
            rec = {"lines": [], "log_file": str(path)}

            mod._append_run_line(rec, "hello")
            mod._append_run_line(rec, "world")

            self.assertEqual("hello\nworld\n", path.read_text(encoding="utf-8"))

    def test_task_run_log_endpoint_returns_log_text(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "r1.txt"
            path.write_text("a\nb\n", encoding="utf-8")
            mod.task_store = _FakeTaskStore(run={"log_file_path": str(path)})

            response = self.client.get("/api/task-runs/7/log")

        self.assertEqual(200, response.status_code)
        self.assertEqual("a\nb\n", response.text)

    def test_task_runs_list_supports_pagination(self):
        store = _FakeTaskStore(
            runs_paged={
                "items": [{"id": 7, "script_title": "ruoyi"}],
                "total": 32,
                "page": 2,
                "page_size": 20,
            }
        )
        mod.task_store = store

        response = self.client.get("/api/task-runs?page=2&page_size=20")

        self.assertEqual(200, response.status_code)
        payload = response.json()
        self.assertEqual(32, payload["total"])
        self.assertEqual(2, payload["page"])
        self.assertEqual(20, payload["page_size"])
        self.assertEqual(2, store.last_run_list_kwargs["page"])
        self.assertEqual(20, store.last_run_list_kwargs["page_size"])

    def test_task_run_detail_no_longer_embeds_accounts(self):
        mod.task_store = _FakeTaskStore(run={"id": 9, "script_title": "ruoyi"})

        response = self.client.get("/api/task-runs/9")

        self.assertEqual(200, response.status_code)
        payload = response.json()
        self.assertEqual(9, payload["id"])
        self.assertNotIn("accounts", payload)

    def test_task_run_accounts_endpoint_supports_pagination(self):
        store = _FakeTaskStore(
            paged={
                "items": [{"id": 1, "email": "abc@outlook.com"}],
                "total": 41,
                "page": 3,
                "page_size": 20,
                "sort_by": "created_at",
                "sort_dir": "desc",
            }
        )
        mod.task_store = store

        response = self.client.get("/api/task-runs/7/accounts?page=3&page_size=20")

        self.assertEqual(200, response.status_code)
        payload = response.json()
        self.assertEqual(41, payload["total"])
        self.assertEqual(3, payload["page"])
        self.assertEqual(20, payload["page_size"])
        self.assertEqual(7, store.last_list_kwargs["task_run_id"])
        self.assertEqual(3, store.last_list_kwargs["page"])
        self.assertEqual(20, store.last_list_kwargs["page_size"])

    def test_task_run_delete_endpoint_calls_store(self):
        store = _FakeTaskStore()
        mod.task_store = store

        response = self.client.delete("/api/task-runs/12")

        self.assertEqual(200, response.status_code)
        self.assertTrue(response.json()["ok"])
        self.assertEqual(12, store.last_delete_task_id)

    def test_build_run_id_is_unique(self):
        a = mod._build_run_id()
        b = mod._build_run_id()
        self.assertTrue(a.startswith("r"))
        self.assertTrue(b.startswith("r"))
        self.assertNotEqual(a, b)

    def test_finish_task_run_deletes_empty_record(self):
        store = _FakeTaskStore()
        mod.task_store = store
        rec = {"store_task_row": 88, "run_id": "r-x", "status": "fail", "ended_ts": 1, "duration_seconds": 2}

        mod._maybe_finish_task_run(rec, {"success": 0, "no_graph": 0, "fail": 3, "total": 3})

        self.assertEqual(88, store.last_delete_task_id)
        self.assertIsNone(store.last_finish_kwargs)

    def test_finish_task_run_keeps_record_when_has_no_graph(self):
        store = _FakeTaskStore()
        mod.task_store = store
        rec = {"store_task_row": 77, "run_id": "r-y", "status": "success", "ended_ts": 1, "duration_seconds": 2}

        mod._maybe_finish_task_run(rec, {"success": 0, "no_graph": 1, "fail": 2, "total": 3, "avg_success_elapsed": 80})

        self.assertIsNone(store.last_delete_task_id)
        self.assertEqual("r-y", store.last_finish_kwargs["run_id"])
        self.assertEqual(1, store.last_finish_kwargs["no_graph_count"])

    def test_accounts_export_returns_current_txt_format(self):
        mod.task_store = _FakeTaskStore(
            accounts=[
                {
                    "id": 9,
                    "email": "a@outlook.com",
                    "password": "Pass123!",
                    "refresh_token": "rt-1",
                    "client_id": "cid-1",
                }
            ]
        )

        response = self.client.get("/api/accounts/export")

        self.assertEqual(200, response.status_code)
        self.assertIn("a@outlook.com----Pass123!----rt-1----cid-1", response.text)

    def test_accounts_export_supports_selected_ids(self):
        store = _FakeTaskStore(
            accounts=[
                {"id": 9, "email": "a@outlook.com", "password": "Pass123!", "refresh_token": "rt-1", "client_id": "cid-1"},
            ]
        )
        mod.task_store = store

        response = self.client.get("/api/accounts/export?ids=9,10")

        self.assertEqual(200, response.status_code)
        self.assertEqual([9, 10], store.last_list_kwargs["ids"])

    def test_accounts_list_supports_pagination_and_filters(self):
        store = _FakeTaskStore(
            paged={
                "items": [{"id": 1, "email": "abc@outlook.com"}],
                "total": 41,
                "page": 2,
                "page_size": 50,
                "sort_by": "created_at",
                "sort_dir": "asc",
            }
        )
        mod.task_store = store

        response = self.client.get(
            "/api/accounts?task_run_id=7&email=abc&status=no_graph&page=2&page_size=50&sort_dir=asc"
        )

        self.assertEqual(200, response.status_code)
        payload = response.json()
        self.assertEqual(41, payload["total"])
        self.assertEqual(2, payload["page"])
        self.assertEqual(50, payload["page_size"])
        self.assertTrue(payload["enabled"])
        self.assertEqual(7, store.last_list_kwargs["task_run_id"])
        self.assertEqual("abc", store.last_list_kwargs["email"])
        self.assertEqual("no_graph", store.last_list_kwargs["status"])
        self.assertEqual(2, store.last_list_kwargs["page"])
        self.assertEqual(50, store.last_list_kwargs["page_size"])
        self.assertEqual("asc", store.last_list_kwargs["sort_dir"])

    def test_accounts_update_endpoint_passes_editable_fields(self):
        store = _FakeTaskStore()
        mod.task_store = store

        response = self.client.put(
            "/api/accounts/9",
            json={
                "email": "new@outlook.com",
                "password": "Pass456!",
                "client_id": "cid-9",
                "refresh_token": "rt-9",
            },
        )

        self.assertEqual(200, response.status_code)
        self.assertEqual(9, store.last_update["account_id"])
        self.assertEqual("new@outlook.com", store.last_update["fields"]["email"])
        self.assertEqual("Pass456!", store.last_update["fields"]["password"])
        self.assertEqual("cid-9", store.last_update["fields"]["client_id"])
        self.assertEqual("rt-9", store.last_update["fields"]["refresh_token"])


if __name__ == "__main__":
    unittest.main()
