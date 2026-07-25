import signal
import subprocess
import sys


def child_creationflags():
    if sys.platform == "win32":
        return getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    return 0


def _taskkill_process_tree(proc, timeout):
    pid = getattr(proc, "pid", None)
    if not pid:
        return False
    subprocess.run(
        ["taskkill", "/PID", str(pid), "/T", "/F"],
        check=False,
        capture_output=True,
        timeout=timeout,
    )
    proc.wait(timeout=timeout)
    return True


def stop_process_gracefully(proc, timeout=15):
    if proc is None or proc.poll() is not None:
        return
    if sys.platform == "win32" and hasattr(signal, "CTRL_BREAK_EVENT"):
        try:
            proc.send_signal(signal.CTRL_BREAK_EVENT)
            proc.wait(timeout=timeout)
            return
        except Exception:
            pass
        try:
            if _taskkill_process_tree(proc, timeout):
                return
        except Exception:
            pass
    proc.terminate()
    proc.wait(timeout=timeout)
