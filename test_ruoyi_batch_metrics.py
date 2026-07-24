import unittest
from types import SimpleNamespace
from unittest.mock import patch

import register_outlook_ruoyi as mod


class BatchMetricsTests(unittest.TestCase):
    def test_format_batch_summary_lines_include_time_and_px_metrics(self):
        lines = mod._format_batch_summary_lines(
            4,
            0,
            14,
            18,
            196.35,
            68.80,
            [
                {"idx": 1, "max_presses": 2, "px_elapsed": 55.57},
                {"idx": 2, "max_presses": 1, "px_elapsed": 33.86},
            ],
        )

        self.assertIn("success=4/18", lines[0])
        self.assertIn("196.35s", lines[2])
        self.assertIn("PX_SUMMARY", lines[3])
        self.assertIn("max_presses 2", lines[3])
        self.assertIn("PX_DETAIL: #1", lines[4])
        self.assertIn("55.57s", lines[4])

    def test_summary_ignores_failed_accounts_in_average(self):
        summary = mod._summarize_batch_metrics(
            ["ok", "fail", "no_graph", "fail"],
            [12.0, 99.0, 18.0, 77.0],
            total_elapsed=50.0,
            px_stats=[
                {"idx": 1, "max_presses": 2, "px_elapsed": 55.0},
                {"idx": 2, "max_presses": 0, "px_elapsed": 0.0},
                {"idx": 3, "max_presses": 1, "px_elapsed": 33.0},
                {"idx": 4, "max_presses": 0, "px_elapsed": 0.0},
            ],
        )

        self.assertEqual(summary["success_count"], 2)
        self.assertEqual(summary["fail_count"], 2)
        self.assertAlmostEqual(summary["total_elapsed"], 50.0)
        self.assertAlmostEqual(summary["avg_success_elapsed"], 15.0)
        self.assertEqual(summary["max_px_presses"], 2)
        self.assertAlmostEqual(summary["avg_px_elapsed"], 44.0)

    def test_summary_ignores_failed_accounts_in_px_stats(self):
        summary = mod._summarize_batch_metrics(
            ["ok", "fail", "no_graph", "fail"],
            [12.0, 99.0, 18.0, 77.0],
            total_elapsed=50.0,
            px_stats=[
                {"idx": 1, "max_presses": 2, "px_elapsed": 55.0},
                {"idx": 2, "max_presses": 9, "px_elapsed": 88.0},
                {"idx": 3, "max_presses": 1, "px_elapsed": 33.0},
                {"idx": 4, "max_presses": 7, "px_elapsed": 66.0},
            ],
        )

        self.assertEqual(summary["max_px_presses"], 2)
        self.assertAlmostEqual(summary["avg_px_elapsed"], 44.0)
        self.assertEqual(
            summary["px_stats"],
            [
                {"idx": 1, "max_presses": 2, "px_elapsed": 55.0},
                {"idx": 3, "max_presses": 1, "px_elapsed": 33.0},
            ],
        )

    def test_summary_returns_zero_average_without_success(self):
        summary = mod._summarize_batch_metrics(
            ["fail", "fail"],
            [11.0, 22.0],
            total_elapsed=40.0,
        )

        self.assertEqual(summary["success_count"], 0)
        self.assertAlmostEqual(summary["avg_success_elapsed"], 0.0)
        self.assertEqual(summary["max_px_presses"], 0)
        self.assertAlmostEqual(summary["avg_px_elapsed"], 0.0)

    def test_run_direct_batch_returns_six_values_with_px_stats(self):
        responses = [
            ("ok", 12.0, {"idx": 1, "max_presses": 2, "px_elapsed": 33.0}),
            ("fail", 99.0, {"idx": 2, "max_presses": 9, "px_elapsed": 88.0}),
        ]

        async def fake_run_one_direct(*args, **kwargs):
            return responses.pop(0)

        class FakePool:
            def stats(self):
                return {"source": "file+aimili-list", "remaining": 1}

            def remaining(self):
                return 0

            def stop(self):
                return None

        args = SimpleNamespace(count=2, concurrency=1, launch_stagger=0)
        with (
            patch.object(mod, "_run_one_direct", side_effect=fake_run_one_direct),
            patch.object(mod, "_load_ua_pool", return_value=["ua"]),
            patch.object(mod, "set_consumable_proxy_pool"),
            patch.object(mod.random, "uniform", return_value=0.0),
        ):
            result = mod.asyncio.run(mod._run_direct_batch(args, object(), FakePool()))

        self.assertEqual(len(result), 6)
        self.assertEqual(result[:3], (1, 0, 1))
        self.assertGreaterEqual(result[3], 0.0)
        self.assertEqual(result[4], 12.0)
        self.assertEqual(result[5], [{"idx": 1, "max_presses": 2, "px_elapsed": 33.0}])


if __name__ == "__main__":
    unittest.main()
