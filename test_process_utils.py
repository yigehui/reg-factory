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

    def test_stop_process_gracefully_falls_back_to_terminate(self):
        proc = unittest.mock.Mock()
        proc.poll.return_value = None
        proc.send_signal.side_effect = OSError("no console")

        with patch.object(mod.sys, "platform", "win32"):
            mod.stop_process_gracefully(proc, timeout=9)

        proc.terminate.assert_called_once()
        proc.wait.assert_called_with(timeout=9)


if __name__ == "__main__":
    unittest.main()
