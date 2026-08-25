"""G 组:firefox 内核路径解析 + 进程树 kill。

去 Administrator 硬编码,改用 LOCALAPPDATA env。RUOYI_FIREFOX_PATH 用 PEP 562
__getattr__ 延迟求值(避免 import 即 glob),get_firefox_path() 带 _state 缓存。
"""

import os
import subprocess
import sys
import time

from ._logging import log
from . import _state


def _localappdata_root():
    """ruyipage browsers 安装根,基于 LOCALAPPDATA(不锁 Administrator)。"""
    return os.path.join(os.environ.get("LOCALAPPDATA", ""), "ruyipage", "browsers")


DEFAULT_RUOYI_FIREFOX = os.path.join(
    os.environ.get("LOCALAPPDATA", ""),
    "ruyipage", "browsers",
    "firefox-151.0a1-151-ruyi-win64", "firefox", "firefox.exe",
)


def _resolve_ruoyi_firefox_path(env_value):
    """动态解析 ruyipage Firefox 内核路径,不锁定具体版本。

    优先级:
      1. 环境变量 RUOYI_FIREFOX_PATH(显式指定,优先尊重)
      2. `ruyipage path` 命令输出(官方管理,装哪个版本指哪个)
      3. browsers 目录下 glob 到的 firefox-*/firefox/firefox.exe(取最新一个)
      4. DEFAULT_RUOYI_FIREFOX 硬编码默认(151 回退,兼容旧安装)
    """
    candidate = str(env_value or "").strip()
    if candidate and os.path.isfile(candidate):
        return candidate

    try:
        out = subprocess.check_output(
            [sys.executable, "-m", "ruyipage", "path"],
            stderr=subprocess.DEVNULL,
            timeout=10,
        )
        line = out.decode("utf-8", "ignore").strip().splitlines()
        if line:
            p = line[-1].strip()
            if p and os.path.isfile(p):
                return p
    except Exception:
        pass

    try:
        root = _localappdata_root()
        if os.path.isdir(root):
            found = []
            for name in os.listdir(root):
                exe = os.path.join(root, name, "firefox", "firefox.exe")
                if os.path.isfile(exe):
                    found.append(exe)
            if found:
                found.sort(reverse=True)
                return found[0]
    except Exception:
        pass

    return DEFAULT_RUOYI_FIREFOX


def get_firefox_path():
    """延迟求值 + 缓存。首次调用解析,后续直接返回。"""
    if _state._FIREFOX_PATH_CACHE is None:
        _state._FIREFOX_PATH_CACHE = _resolve_ruoyi_firefox_path(os.environ.get("RUOYI_FIREFOX_PATH"))
    return _state._FIREFOX_PATH_CACHE


def set_firefox_path(path):
    """显式注入路径(供测试/调用方覆盖)。"""
    _state._FIREFOX_PATH_CACHE = str(path or "").strip() or None
    return _state._FIREFOX_PATH_CACHE


# PEP 562:兼容旧调用方直接读 module.RUOYI_FIREFOX_PATH,触发延迟求值。
def __getattr__(name):
    if name == "RUOYI_FIREFOX_PATH":
        return get_firefox_path()
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


def _ruoyi_firefox_running_count_by_path(firefox_path=None):
    """按指定 firefox_path 统计当前残留 firefox.exe 进程数(按 ExecutablePath 过滤)。"""
    fp = firefox_path or get_firefox_path()
    if not fp or not os.path.isfile(fp):
        return 0
    norm = os.path.normpath(fp)
    target = norm.replace("'", "''")
    ps = (
        "$ErrorActionPreference='SilentlyContinue';"
        "$t='%s'.ToLower();"
        "$p=Get-Process firefox -ErrorAction SilentlyContinue;"
        "if($p){$k=$p|Where-Object{$_.Path -and $_.Path.ToLower() -eq $t};"
        "@($k).Count}else{0}"
    ) % target
    try:
        out = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                             capture_output=True, text=True, timeout=10)
        n = (out.stdout or "").strip()
        if n.isdigit():
            return int(n)
    except Exception:
        pass
    return 0


def _force_kill_ruoyi_firefox(log_fn=None, firefox_path=None):
    """强杀所有以指定 path 启动的 firefox.exe 进程(quit 超时残留兜底)。

    按 ExecutablePath 过滤,不影响用户自己的 Firefox。仅在批次间隙(所有 slot 空闲)调用,
    避免误杀同批正在运行的实例。杀完轮询等待进程真正消失(释放 XPCOM/profile 文件句柄),
    否则下批次启动 firefox 会撞文件锁 -> "Couldn't load XPCOM"。返回杀掉的进程数。
    """
    fp = firefox_path or get_firefox_path()
    if not fp or not os.path.isfile(fp):
        return 0
    lf = log_fn or log
    norm = os.path.normpath(fp)
    target = norm.replace("'", "''")
    ps = (
        "$ErrorActionPreference='SilentlyContinue';"
        "$t='%s'.ToLower();"
        "$p=Get-Process firefox -ErrorAction SilentlyContinue;"
        "if($p){$k=$p|Where-Object{$_.Path -and $_.Path.ToLower() -eq $t};"
        "if($k){$n=@($k).Count; $k|Stop-Process -Force; Write-Output $n}}"
    ) % target
    killed = 0
    try:
        out = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                             capture_output=True, text=True, timeout=15)
        n = (out.stdout or "").strip()
        if n and n.isdigit():
            killed = int(n)
            lf(f"ruoyi 残留 Firefox 强杀: {killed} 个进程(按路径 {norm})", "OK")
    except Exception as exc:
        lf(f"ruoyi 残留 Firefox 强杀失败: {type(exc).__name__}: {exc}", "WARN")

    # Stop-Process 返回后 Windows 不一定立即释放 firefox 的 DLL/XPCOM 文件句柄,
    # 紧接着启动下批次会撞文件锁。这里轮询等到进程对象真正消失再返回。
    wait_total = max(0.0, float(_state.RUOYI_FIREFOX_FORCEKILL_WAIT or 0.0))
    if wait_total > 0:
        deadline = time.time() + wait_total
        while time.time() < deadline:
            if _ruoyi_firefox_running_count_by_path(fp) <= 0:
                break
            time.sleep(0.2)
        still = _ruoyi_firefox_running_count_by_path(fp)
        if still > 0:
            lf(f"ruoyi Firefox 强杀后仍有 {still} 个进程未退出(等 {wait_total:.1f}s 超时)", "WARN")
    return killed


def _kill_ruoyi_firefox_by_profile(profile_dir, timeout=8.0, log_fn=None):
    """按 profile_dir 精准杀 firefox 进程树(主进程 + content/gpu 子进程),并等待全部退出。

    ruyipage 的 quit() 只 terminate() 主进程,不杀子进程树;firefox 多进程子进程变孤儿后
    继续占用 XPCOM 组件文件和 profile 目录,导致下个 run 启动撞文件锁 -> "Couldn't load XPCOM"。
    本函数按命令行 --profile <profile_dir> 匹配所有相关 firefox.exe(子进程命令行同样带该参数),
    taskkill /F 逐个杀,再轮询确认退出。按 profile_dir 唯一匹配(mkdtemp 唯一路径),不误杀其他 slot。

    返回 (killed, all_gone)。"""
    import json

    lf = log_fn or log
    path = os.path.abspath(str(profile_dir or "").strip())
    if not path or not os.path.isabs(path):
        return 0, True
    needle = os.path.normpath(path).lower()

    ps_list = (
        "$ErrorActionPreference='SilentlyContinue';"
        "Get-CimInstance Win32_Process | "
        "Where-Object{$_.Name -eq 'firefox.exe'} | "
        "Select-Object ProcessId,CommandLine | ConvertTo-Json -Compress"
    )

    def _match_pids():
        try:
            out = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_list],
                                 capture_output=True, text=True, timeout=15)
            raw = (out.stdout or "").strip()
        except Exception:
            return []
        try:
            data = json.loads(raw) if raw else []
        except Exception:
            return []
        if isinstance(data, dict):
            data = [data]
        pids = []
        for item in data:
            try:
                pid = int(item.get("ProcessId") or 0)
            except (TypeError, ValueError):
                pid = 0
            if pid <= 0:
                continue
            cmd = str(item.get("CommandLine") or "").lower()
            if needle in cmd:
                pids.append(pid)
        return pids

    pids = _match_pids()
    killed = 0
    for pid in pids:
        try:
            subprocess.run(["taskkill", "/F", "/PID", str(pid)],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=5)
            killed += 1
        except Exception:
            pass

    # 轮询等待该 profile 关联的 firefox 全部退出,确保 XPCOM/profile 文件句柄释放。
    all_gone = True
    wait_total = max(0.0, float(timeout or 0.0))
    if wait_total > 0:
        deadline = time.time() + wait_total
        while time.time() < deadline:
            if not _match_pids():
                break
            all_gone = False
            time.sleep(0.2)
        if _match_pids():
            all_gone = False

    if killed:
        lf(f"ruoyi 按 profile 清 firefox 进程树: 杀 {killed} 个(run={os.path.basename(path)})", "OK")
    if not all_gone:
        lf(f"ruoyi 按 profile 清 firefox 仍有残留未退出(等 {wait_total:.1f}s 超时)", "WARN")
    return killed, all_gone
