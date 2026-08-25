"""E 组:profile 目录创建 + 清理。

围栏重构:去 register 的 ROOT 依赖。原 root in {ROOT, os.sep, parent} 改为
"root 不能是系统根/父根 + 不在 protect_paths 中"。protect_paths 供调用方(register)
注入业务 ROOT,新包默认空,零业务依赖。
"""

import os
import shutil
import tempfile
import time

from ._logging import log
from . import _state


def get_profile_root():
    """返回 profile 根目录:_PROFILE_ROOT 注入值 或 默认 cwd/profiles_ruoyi_tmp。"""
    if _state._PROFILE_ROOT:
        return _state._PROFILE_ROOT
    return os.path.join(os.getcwd(), "profiles_ruoyi_tmp")


def set_profile_root(path):
    """供调用方注入业务 profile 根目录。"""
    _state._PROFILE_ROOT = str(path or "").strip() or None
    return _state._PROFILE_ROOT


def _fence_root(root, protect_paths=None):
    """围栏:root 不能是系统根/父根 + 不能等于 protect_paths 中任一路径。

    返回 True 表示安全可清理,False 表示应拒绝(防误删)。"""
    abspath = os.path.abspath(str(root or "").strip())
    if not abspath:
        return False
    sys_root = os.path.abspath(os.sep)
    parent_root = os.path.abspath(os.path.join(abspath, os.pardir))
    if abspath == sys_root or abspath == parent_root:
        return False
    if protect_paths:
        for p in protect_paths:
            if p and os.path.abspath(str(p)) == abspath:
                return False
    return True


def _ruoyi_profile_dir(opts, idx, profile_root=None):
    root = os.path.abspath(str(profile_root or get_profile_root() or "").strip())
    os.makedirs(root, exist_ok=True)
    _cleanup_stale_ruoyi_profile_root(root)

    try:
        slot = max(1, int(getattr(opts, "ruoyi_slot", None) or 0))
    except Exception:
        slot = 0
    if not slot:
        try:
            slots = max(1, int(getattr(opts, "concurrency", 1) or 1))
        except Exception:
            slots = 1
        try:
            slot = ((max(1, int(idx or 1)) - 1) % slots) + 1
        except Exception:
            slot = 1

    slot_root = os.path.join(root, f"slot_{slot:02d}")
    os.makedirs(slot_root, exist_ok=True)

    try:
        run_idx = max(1, int(idx or 1))
    except Exception:
        run_idx = 1

    profile_dir = tempfile.mkdtemp(prefix=f"run_{run_idx:04d}_", dir=slot_root)
    with _state._ACTIVE_RUOYI_PROFILE_DIRS_LOCK:
        _state._ACTIVE_RUOYI_PROFILE_DIRS.add(os.path.abspath(profile_dir))
    return profile_dir


def _cleanup_ruoyi_profile_root(profile_root=None, protect_paths=None):
    root = os.path.abspath(str(profile_root or get_profile_root() or "").strip())
    if not _fence_root(root, protect_paths):
        return 0
    if not os.path.isdir(root):
        return 0
    cleaned = 0
    for name in os.listdir(root):
        path = os.path.join(root, name)
        try:
            if os.path.isdir(path):
                shutil.rmtree(path)
            else:
                os.remove(path)
            cleaned += 1
        except Exception as exc:
            log(f"清理 ruoyi profile 缓存失败: {path}: {type(exc).__name__}: {exc}", "WARN")
    return cleaned


def _cleanup_stale_ruoyi_profile_root(profile_root=None, max_age_sec=None, protect_paths=None):
    root = os.path.abspath(str(profile_root or get_profile_root() or "").strip())
    if not root or not os.path.isdir(root):
        return 0
    now = time.time()
    stale_sec = float(max_age_sec or _state.RUOYI_PROFILE_STALE_SEC or 0.0)
    cleaned = 0

    paths = []
    for name in os.listdir(root):
        path = os.path.join(root, name)
        if os.path.isdir(path) and name.startswith("run_"):
            paths.append(path)
        elif os.path.isdir(path) and name.startswith("slot_"):
            for child in os.listdir(path):
                child_path = os.path.join(path, child)
                if os.path.isdir(child_path) and child.startswith("run_"):
                    paths.append(child_path)

    for path in paths:
        try:
            with _state._ACTIVE_RUOYI_PROFILE_DIRS_LOCK:
                if os.path.abspath(path) in _state._ACTIVE_RUOYI_PROFILE_DIRS:
                    continue
            age_sec = max(0.0, now - os.path.getmtime(path))
            if stale_sec > 0 and age_sec < stale_sec:
                continue
            shutil.rmtree(path)
            cleaned += 1
        except Exception as exc:
            log(f"清理过期 ruoyi tmp 失败: {path}: {type(exc).__name__}: {exc}", "WARN")
    return cleaned


def _cleanup_ruoyi_slot_root(slot_root, profile_root=None):
    root = os.path.abspath(str(profile_root or get_profile_root() or "").strip())
    path = os.path.abspath(str(slot_root or "").strip())
    if not root or not path or not os.path.isdir(path):
        return 0
    try:
        if os.path.commonpath([root, path]) != root:
            return 0
    except Exception:
        return 0
    if not os.path.basename(path).startswith("slot_"):
        return 0
    cleaned = 0
    for name in os.listdir(path):
        child = os.path.join(path, name)
        try:
            if os.path.isdir(child):
                shutil.rmtree(child)
            else:
                os.remove(child)
            cleaned += 1
        except Exception:
            pass
    return cleaned


def _cleanup_ruoyi_run_profile_dir(profile_dir, profile_root=None):
    root = os.path.abspath(str(profile_root or get_profile_root() or "").strip())
    path = os.path.abspath(str(profile_dir or "").strip())
    if not root or not path or not os.path.isdir(path):
        return False
    try:
        if os.path.commonpath([root, path]) != root:
            return False
    except Exception:
        return False
    if not os.path.basename(path).startswith("run_"):
        return False

    with _state._ACTIVE_RUOYI_PROFILE_DIRS_LOCK:
        _state._ACTIVE_RUOYI_PROFILE_DIRS.discard(path)

    retries = max(1, int(_state.RUOYI_PROFILE_CLEANUP_RETRIES or 1))
    delay = max(0.0, float(_state.RUOYI_PROFILE_CLEANUP_RETRY_DELAY or 0.0))
    last_exc = None
    for attempt in range(retries):
        try:
            shutil.rmtree(path)
            return True
        except Exception as exc:
            last_exc = exc
            if attempt + 1 >= retries:
                break
            if delay > 0:
                time.sleep(delay)
    if last_exc is not None:
        log(f"清理 ruoyi tmp profile 失败: {path}: {type(last_exc).__name__}: {last_exc}", "WARN")
    return False
