import signal
import subprocess
import sys


def child_creationflags():
    if sys.platform == "win32":
        return getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    return 0


def stop_process_gracefully(proc, timeout=15):
    if proc is None or proc.poll() is not None:
        return
    waited = False
    if sys.platform == "win32" and hasattr(signal, "CTRL_BREAK_EVENT"):
        try:
            proc.send_signal(signal.CTRL_BREAK_EVENT)
            proc.wait(timeout=timeout)
            waited = True
        except Exception:
            pass
    if waited:
        return
    proc.terminate()
    proc.wait(timeout=timeout)
