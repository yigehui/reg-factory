import signal
import subprocess
import unittest
from unittest.mock import patch

import process_utils as mod


class ProcessUtilsTests(unittest.TestCase):
    def test_child_creationflags_use_new_process_group_on_windows(self):
        with patch.object(mod.sys, "platform", "win32"):
            self.assertEqual(
                mod.child_creationflags(),
                getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0),
            )

    def test_stop_process_gracefully_uses_ctrl_break_on_windows(self):
        proc = unittest.mock.Mock()
        proc.poll.return_value = None

        with patch.object(mod.sys, "platform", "win32"):
            mod.stop_process_gracefully(proc, timeout=7)

        proc.send_signal.assert_called_once_with(signal.CTRL_BREAK_EVENT)
        proc.wait.assert_called_once_with(timeout=7)
        proc.terminate.assert_not_called()

    def test_stop_process_gracefully_falls_back_to_taskkill_tree(self):
        proc = unittest.mock.Mock()
        proc.poll.return_value = None
        proc.send_signal.side_effect = OSError("no console")

        with patch.object(mod.sys, "platform", "win32"):
            mod.stop_process_gracefully(proc, timeout=9)

        proc.terminate.assert_not_called()
        proc.wait.assert_called_with(timeout=9)

    def test_stop_process_gracefully_taskkills_process_tree_on_windows_fallback(self):
        proc = unittest.mock.Mock()
        proc.poll.return_value = None
        proc.pid = 4321
        proc.send_signal.side_effect = OSError("no console")

        with (
            patch.object(mod.sys, "platform", "win32"),
            patch.object(mod.subprocess, "run") as run,
        ):
            mod.stop_process_gracefully(proc, timeout=9)

        run.assert_called_once_with(
            ["taskkill", "/PID", "4321", "/T", "/F"],
            check=False,
            capture_output=True,
            timeout=9,
        )

    def test_stop_process_gracefully_taskkills_tree_when_ctrl_break_wait_times_out(self):
        proc = unittest.mock.Mock()
        proc.poll.return_value = None
        proc.pid = 5678
        proc.wait.side_effect = [subprocess.TimeoutExpired(cmd="x", timeout=7), None]

        with (
            patch.object(mod.sys, "platform", "win32"),
            patch.object(mod.subprocess, "run") as run,
        ):
            mod.stop_process_gracefully(proc, timeout=7)

        proc.send_signal.assert_called_once_with(signal.CTRL_BREAK_EVENT)
        run.assert_called_once_with(
            ["taskkill", "/PID", "5678", "/T", "/F"],
            check=False,
            capture_output=True,
            timeout=7,
        )


if __name__ == "__main__":
    unittest.main()
