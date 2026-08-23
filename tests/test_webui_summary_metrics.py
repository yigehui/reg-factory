import unittest

import webui.server as server


class WebuiSummaryMetricsTests(unittest.TestCase):
    def test_extract_run_counts_parses_time_metrics(self):
        counts = server._extract_run_counts(
            [
                "DONE: success=4/18 fail=14 no_graph=0",
                "SUMMARY: success 4 | fail 14 | no_graph 0 | total 18",
                "SUMMARY_TIME: total_elapsed 196.35s | avg_success_elapsed 68.80s",
            ]
        )

        self.assertEqual(counts["success"], 4)
        self.assertEqual(counts["fail"], 14)
        self.assertEqual(counts["no_graph"], 0)
        self.assertEqual(counts["total"], 18)
        self.assertAlmostEqual(counts["total_elapsed"], 196.35)
        self.assertAlmostEqual(counts["avg_success_elapsed"], 68.80)

    def test_format_webui_summary_includes_time_metrics(self):
        line = server._format_webui_run_summary(
            {
                "success": 4,
                "fail": 14,
                "no_graph": 0,
                "total": 18,
                "total_elapsed": 196.35,
                "avg_success_elapsed": 68.80,
            }
        )

        self.assertIn("success=4", line)
        self.assertIn("fail=14", line)
        self.assertIn("196.35s", line)
        self.assertIn("68.80s", line)


if __name__ == "__main__":
    unittest.main()
