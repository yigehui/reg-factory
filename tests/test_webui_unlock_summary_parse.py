import unittest

import webui.server as server


class UnlockSummaryParseTests(unittest.TestCase):
    def test_extract_unlock_summary_three_buckets(self):
        """unlock 的 unlocked/needs_phone/failed SUMMARY 能被 webui 解析,
        映射 success=unlocked / fail=failed / no_graph=needs_phone。"""
        lines = [
            "[12:00:00] [INFO] SUMMARY: unlocked 2 | needs_phone 1 | failed 3 | total 6",
            "[12:00:00] [INFO] SUMMARY_TIME: total_elapsed 12.34s",
        ]
        counts = server._extract_run_counts(lines)
        self.assertIsNotNone(counts, "unlock SUMMARY 应被 webui 解析,不应返回 None")
        self.assertEqual(counts["success"], 2)
        self.assertEqual(counts["fail"], 3)
        self.assertEqual(counts["no_graph"], 1)
        self.assertEqual(counts["total"], 6)
        self.assertEqual(counts["total_elapsed"], 12.34)

    def test_extract_unlock_summary_no_needs_phone(self):
        """无 needs_phone 时 no_graph=0。"""
        lines = [
            "SUMMARY: unlocked 5 | needs_phone 0 | failed 1 | total 6",
            "SUMMARY_TIME: total_elapsed 8.0s",
        ]
        counts = server._extract_run_counts(lines)
        self.assertIsNotNone(counts)
        self.assertEqual(counts["success"], 5)
        self.assertEqual(counts["fail"], 1)
        self.assertEqual(counts["no_graph"], 0)
        self.assertEqual(counts["total"], 6)

    def test_extract_unlock_summary_time_without_avg(self):
        """unlock SUMMARY_TIME 只有 total_elapsed(无 avg_success_elapsed),也应解析。
        avg_success_elapsed 允许为 None。"""
        lines = [
            "SUMMARY: unlocked 1 | needs_phone 0 | failed 0 | total 1",
            "SUMMARY_TIME: total_elapsed 0.00s",
        ]
        counts = server._extract_run_counts(lines)
        self.assertIsNotNone(counts)
        self.assertEqual(counts["total_elapsed"], 0.0)
        # avg_success_elapsed 缺失时不应阻断解析;可为 None 或 0,不强制
        # (只要 counts 非 None 且 total_elapsed 正确即可)


if __name__ == "__main__":
    unittest.main()
