# -*- coding: utf-8 -*-
"""Outlook 解锁 Web UI schema/cmd 对齐测试。

验证：
1. webui/scripts.py 的 unlock_outlook schema 暴露了与 unlock_outlook.py CLI 对齐的字段。
2. webui/server.py _build_cmd 按代理来源(file/aimili-list)正确拼出 CLI 参数，
   隐藏字段(proxy-source=aimili-list 时的 --proxy-file)即使被前端误传也会被后端过滤。
"""
import unittest

from webui import scripts as scripts_mod
from webui.server import _build_cmd


def _schema(sid):
    for s in scripts_mod.SCRIPTS:
        if s["id"] == sid:
            return s
    raise AssertionError(f"script {sid} not found")


class UnlockOutlookSchemaTests(unittest.TestCase):
    def setUp(self):
        self.s = _schema("unlock_outlook")
        self.by_flag = {a["flag"]: a for a in self.s["args"]}

    def test_proxy_source_choices_are_file_and_aimili_list_only(self):
        spec = self.by_flag["--proxy-source"]
        self.assertEqual(spec["type"], "choice")
        values = []
        for c in spec["choices"]:
            values.append(c["value"] if isinstance(c, dict) else c)
        self.assertEqual(values, ["file", "aimili-list"])

    def test_aimili_token_marked_secret(self):
        self.assertTrue(self.by_flag["--aimili-token"].get("secret"))

    def test_proxy_file_default_matches_cli(self):
        self.assertEqual(self.by_flag["--proxy-file"]["default"], "proxies_outlook.txt")

    def test_all_expected_flags_present(self):
        for f in ["--input", "--proxy-source", "--proxy-file", "--aimili-url",
                  "--aimili-token", "--concurrency", "--headless", "--max-press",
                  "--timeout", "--log-level"]:
            self.assertIn(f, self.by_flag, f"missing {f}")

    def test_visible_if_wired_for_proxy_fields(self):
        self.assertEqual(self.by_flag["--proxy-file"].get("visible_if"),
                         {"flag": "--proxy-source", "equals": "file"})
        self.assertEqual(self.by_flag["--aimili-url"].get("visible_if"),
                         {"flag": "--proxy-source", "equals": "aimili-list"})
        self.assertEqual(self.by_flag["--aimili-token"].get("visible_if"),
                         {"flag": "--proxy-source", "equals": "aimili-list"})


class UnlockBuildCmdTests(unittest.TestCase):
    def test_file_source_emits_proxy_file_not_aimili(self):
        cmd = _build_cmd(_schema("unlock_outlook"), {
            "--proxy-source": "file",
            "--proxy-file": "proxies_outlook.txt",
            "--concurrency": 2,
            "--headless": True,
            "--max-press": 5,
            "--timeout": 300,
            "--log-level": "INFO",
        })
        s = " ".join(cmd)
        self.assertIn("--proxy-source file", s)
        self.assertIn("--proxy-file proxies_outlook.txt", s)
        self.assertIn("--concurrency 2", s)
        self.assertIn("--headless", s)
        self.assertIn("--max-press 5", s)
        self.assertIn("--timeout 300", s)
        self.assertIn("--log-level INFO", s)
        self.assertNotIn("--aimili-url", s)
        self.assertNotIn("--aimili-token", s)

    def test_aimili_source_emits_aimili_not_proxy_file(self):
        cmd = _build_cmd(_schema("unlock_outlook"), {
            "--proxy-source": "aimili-list",
            "--aimili-url": "http://host:8787",
            "--aimili-token": "SECRET",
            "--concurrency": 1,
        })
        s = " ".join(cmd)
        self.assertIn("--proxy-source aimili-list", s)
        self.assertIn("--aimili-url http://host:8787", s)
        self.assertIn("--aimili-token SECRET", s)
        self.assertNotIn("--proxy-file", s)

    def test_headless_false_omits_flag(self):
        cmd = _build_cmd(_schema("unlock_outlook"), {
            "--proxy-source": "file",
            "--headless": False,
        })
        self.assertNotIn("--headless", cmd)


if __name__ == "__main__":
    unittest.main()
