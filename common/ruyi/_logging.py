"""骨架 log:级别阈值 + 时间戳,无业务关键词过滤。

新包内部子模块 + unlock/launch 等用本 log。register 保留自己的业务版 log
(含 _should_keep_prod_log/_should_demote_to_debug 的 50+ Outlook 关键词过滤),
不 re-export 本模块的 log/set_log_level,避免覆盖业务版。

通过 register_log_filter(keep_fn, demote_fn) 可注入业务过滤钩子,默认 no-op,
保证新包零业务依赖同时允许调用方定制过滤。
"""

import os
import threading
from datetime import datetime

LOG_LEVELS = {
    "DEBUG": 10,
    "INFO": 20,
    "OK": 20,
    "STEP": 20,
    "WARN": 30,
    "PROD": 35,
    "ERR": 40,
}

LOG_LEVEL = "INFO"

_log_lock = threading.Lock()

# 可注入过滤钩子,默认 no-op(全保留/不降级)
_keep_fn = None
_demote_fn = None


def register_log_filter(keep_fn=None, demote_fn=None):
    """注入业务过滤钩子(供 register 注入 _should_keep_prod_log/_should_demote_to_debug)。

    keep_fn(msg, rendered_level) -> bool  PROD 模式下是否保留(返回 False 则丢弃)。
    demote_fn(msg, rendered_level) -> str|None  非 PROD 模式下降级到的级别(None=不降级)。
    """
    global _keep_fn, _demote_fn
    _keep_fn = keep_fn
    _demote_fn = demote_fn


def _normalize_log_level(value, default="INFO"):
    raw = str(value or default).strip().upper()
    aliases = {
        "TRACE": "DEBUG",
        "DBG": "DEBUG",
        "WARNING": "WARN",
        "PRODUCTION": "PROD",
        "ERROR": "ERR",
        "SUCCESS": "OK",
    }
    raw = aliases.get(raw, raw)
    return raw if raw in LOG_LEVELS else default


def _log_level_value(value):
    return LOG_LEVELS.get(_normalize_log_level(value), LOG_LEVELS["INFO"])


def set_log_level(value):
    """设 LOG_LEVEL + 写中性 env RUOYI_LOG_LEVEL。"""
    global LOG_LEVEL
    LOG_LEVEL = _normalize_log_level(value)
    os.environ["RUOYI_LOG_LEVEL"] = LOG_LEVEL
    return LOG_LEVEL


def log(msg, level="INFO"):
    """骨架 log:级别阈值 + 时间戳 print。

    若注入了业务过滤钩子,走业务过滤;否则只做级别阈值比较。
    """
    rendered = str(level or "INFO").strip().upper() or "INFO"

    if LOG_LEVEL == "PROD":
        if _keep_fn is not None:
            if not _keep_fn(msg, rendered):
                return
        elif _log_level_value(rendered) < _log_level_value("PROD"):
            return
    else:
        if _demote_fn is not None:
            effective = _demote_fn(msg, rendered) or rendered
        else:
            effective = rendered
        if _log_level_value(effective) < _log_level_value(LOG_LEVEL):
            return

    with _log_lock:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] [{rendered}] {msg}", flush=True)


def debug(msg):
    log(msg, "DEBUG")
