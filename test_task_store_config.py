import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import task_store


class TaskStoreConfigTests(unittest.TestCase):
    def test_is_configured_reads_values_from_env_file(self):
        with tempfile.TemporaryDirectory() as td:
            env_path = Path(td) / ".env"
            env_path.write_text(
                "\n".join(
                    [
                        "WEBUI_MYSQL_HOST=127.0.0.1",
                        "WEBUI_MYSQL_PORT=3306",
                        "WEBUI_MYSQL_USER=u",
                        "WEBUI_MYSQL_PASSWORD=p",
                        "WEBUI_MYSQL_DATABASE=db",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            with patch.object(task_store, "_ENV_PATH", env_path):
                with patch.dict(os.environ, {}, clear=True):
                    self.assertTrue(task_store.is_configured())
                    cfg = task_store._cfg()
        self.assertEqual("127.0.0.1", cfg["host"])
        self.assertEqual("u", cfg["user"])
        self.assertEqual("db", cfg["database"])


if __name__ == "__main__":
    unittest.main()
