#!/usr/bin/env python3

# -*- coding: utf-8 -*-

"""

Outlook 自注册养号(ruoyi) —— 完整链路版



目标：

  - 使用 ruyipage 定制 Firefox + per-tab SOCKS5 代理池完成 Outlook 自注册

  - 产出可登录账号，并在直跑模式下补抽 Graph refresh_token

  - 可被 outlook_reg_loop.py 作为完整后端调用



说明：

  - 与原 BitBrowser/Playwright 链路相比，ruoyi 走 Firefox BiDi。

  - 本文件保留独立直跑能力，同时暴露 register_outlook() 给 loop 调用。

"""



from __future__ import annotations



import argparse

import asyncio

import atexit

from contextlib import contextmanager

import importlib.util

import json

import os

import random

import re

import requests

import shutil

import subprocess

import signal

import sys

import tempfile

import threading

import time

import urllib.parse

import urllib.request

from datetime import datetime

from types import SimpleNamespace

from urllib.parse import quote

from config import _load_dotenv

from common.notify import send_tg_message

from common import ruyi as _ruyi_pkg

from common import ruyi_browser as _rb


_load_dotenv()

# --- ruyi 公共能力:自包含包 common.ruyi(通用层)。register 做 env 兼容层 + re-export。 ---
# 新包只认中性 env 名(RUOYI_*),这里把旧名(OUTLOOK_*)提升/回落到中性名,保持本文件行为不变。
# 业务版 log(_should_keep_prod_log/_should_demote_to_debug)留 register,不 re-export 新包 log。




DEFAULT_RUOYI_FIREFOX = _ruyi_pkg.DEFAULT_RUOYI_FIREFOX
_resolve_ruoyi_firefox_path = _ruyi_pkg._resolve_ruoyi_firefox_path
get_firefox_path = _ruyi_pkg.get_firefox_path
# RUOYI_FIREFOX_PATH 延迟求值(新包 PEP 562 __getattr__);保留模块级名字供旧引用读
RUOYI_FIREFOX_PATH = _ruyi_pkg.RUOYI_FIREFOX_PATH


ROOT = os.path.dirname(os.path.abspath(__file__))

ENGINE_NAME = "ruoyi"

# 代理池常量:旧名优先,回落新包中性默认。
# 同时把旧名值提升到中性名(setdefault),新包 ConsumableProxyPool.from_args 在 args 空时
# 回落读中性名(RUOYI_PROXY_FILE 等),从而旧名配置仍能被新包池解析到。
PROXY_FILE = os.environ.get("OUTLOOK_PROXY_FILE", _ruyi_pkg.PROXY_FILE)
os.environ.setdefault("RUOYI_PROXY_FILE", os.environ.get("OUTLOOK_PROXY_FILE", ""))

RUOYI_PROXY_SOURCE = os.environ.get("OUTLOOK_RUOYI_PROXY_SOURCE", _ruyi_pkg.RUOYI_PROXY_SOURCE)
os.environ.setdefault("RUOYI_PROXY_SOURCE", os.environ.get("OUTLOOK_RUOYI_PROXY_SOURCE", ""))

PROXY_URL = os.environ.get("OUTLOOK_PROXY_URL", _ruyi_pkg.PROXY_URL)
os.environ.setdefault("RUOYI_PROXY_URL", os.environ.get("OUTLOOK_PROXY_URL", ""))

SCREENSHOT_DIR = os.path.join(ROOT, "screenshots_ruoyi")

HAR_DIR = os.path.join(ROOT, "har_ruoyi")

OUTPUT_DIR = os.path.join(ROOT, "outlook_accounts")

RUOYI_PROFILE_ROOT = os.environ.get("OUTLOOK_RUOYI_PROFILE_ROOT", os.path.join(ROOT, "profiles_ruoyi_tmp"))

# 注入业务 profile root 到新包(围栏 protect_paths 在调用时传 ROOT)
_ruyi_pkg.set_profile_root(RUOYI_PROFILE_ROOT)

RUOYI_PROFILE_STALE_SEC = float(os.environ.get("OUTLOOK_RUOYI_PROFILE_STALE_SEC", _ruyi_pkg._state.RUOYI_PROFILE_STALE_SEC) or _ruyi_pkg._state.RUOYI_PROFILE_STALE_SEC)
_ruyi_pkg._state.RUOYI_PROFILE_STALE_SEC = RUOYI_PROFILE_STALE_SEC


# no_graph 账号累计文件，与 EMAILS_POOL 同级(ROOT 外层)，保持与成功号桥接 emails.txt 对称

EMAIL_NOGRAPH = os.path.join(ROOT, "email_nograph.txt")

EMAILS_POOL = os.path.join(ROOT, "emails.txt")

# 仅注册(跳过 Graph 授权)账号累计文件，与 EMAILS_POOL 同级

EMAIL_REG = os.path.join(ROOT, "email_reg.txt")

SIGNUP_URL = "https://signup.live.com/signup?lic=1"

# warmup 首页:与 signup 同 eTLD+1(login.live.com),落地即种 PX/MS 挑战 cookie
WARMUP_HOME_URL = "https://login.live.com/"


def _ruoyi_profile_dir(opts, idx):
    """包装新包 _ruoyi_profile_dir:读本模块 RUOYI_PROFILE_ROOT(profile_root 注入),
    保留以便测试 patch.object(mod, 'RUOYI_PROFILE_ROOT', ...) 仍生效。"""
    return _ruyi_pkg._ruoyi_profile_dir(opts, idx, profile_root=RUOYI_PROFILE_ROOT)

# B 组状态 + IP 端点:迁入新包 _state,re-export 保持模块级名字
IP_INFO_ENDPOINTS = _ruyi_pkg._state.IP_INFO_ENDPOINTS
_CURRENT_IP_INFO = _ruyi_pkg._state._CURRENT_IP_INFO
_CURRENT_IP_INFO_LOCK = _ruyi_pkg._state._CURRENT_IP_INFO_LOCK

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

ALLOWED_EMAIL_SUFFIXES = ("outlook.com", "hotmail.com")



REGISTER_TIMEOUT = 300

SIGNUP_ENTRY_TIMEOUT = 20

SUBMIT_RESULT_TIMEOUT = 15

# submit 卡死硬超时:持续点中 submit 按钮(页面没跳 captcha/loading/blocked)但 N 秒内 URL 仍无变化,
# 判定卡死(代理抖动/出口IP风控导致提交无响应),fail 换号换代理,不死等到 REGISTER_TIMEOUT。
SUBMIT_STUCK_NOCHANGE_TIMEOUT = float(os.environ.get("OUTLOOK_RUOYI_SUBMIT_STUCK_TIMEOUT", "25") or "25")

# captcha 主循环无进展总超时:进入注册后等待/按压期间,连续 N 秒既没成功按压、没跳转、也没出 captcha 进展,
# 判死循环卡死(覆盖 submit 后 loading 不结束、点不中 submit 静默空转等所有场景),早退 fail 换号换代理。
NO_PROGRESS_TIMEOUT = float(os.environ.get("OUTLOOK_RUOYI_NO_PROGRESS_TIMEOUT", "60") or "60")

# 邮箱页 stuck 换号上限:连续 stuck 换到第 N 个号仍过不去就早退 fail,避免 5 轮 × 8s 累加到 60s。
EMAIL_STUCK_MAX_ROTATE = int(os.environ.get("OUTLOOK_RUOYI_EMAIL_STUCK_MAX_ROTATE", "2") or "2")

PROXY_PRECHECK_TIMEOUT = float(os.environ.get("OUTLOOK_RUOYI_PROXY_PRECHECK_TIMEOUT", _ruyi_pkg._state.PROXY_PRECHECK_TIMEOUT) or _ruyi_pkg._state.PROXY_PRECHECK_TIMEOUT)
_ruyi_pkg._state.PROXY_PRECHECK_TIMEOUT = PROXY_PRECHECK_TIMEOUT

PROXY_IDENTITY_TIMEOUT = float(os.environ.get("OUTLOOK_RUOYI_PROXY_IDENTITY_TIMEOUT", _ruyi_pkg._state.PROXY_IDENTITY_TIMEOUT) or _ruyi_pkg._state.PROXY_IDENTITY_TIMEOUT)
_ruyi_pkg._state.PROXY_IDENTITY_TIMEOUT = PROXY_IDENTITY_TIMEOUT

PROXY_IDENTITY_CACHE_TTL = float(os.environ.get("OUTLOOK_RUOYI_PROXY_IDENTITY_CACHE_TTL", _ruyi_pkg._state.PROXY_IDENTITY_CACHE_TTL) or _ruyi_pkg._state.PROXY_IDENTITY_CACHE_TTL)
_ruyi_pkg._state.PROXY_IDENTITY_CACHE_TTL = PROXY_IDENTITY_CACHE_TTL

# 代理预检业务域名:注入到新包 _state(新包默认空),所有内部探测调用自动生效
PROXY_PRECHECK_TARGETS = (
    "https://signup.live.com/",
    "https://login.live.com/",
)
_ruyi_pkg._state.PROXY_PRECHECK_TARGETS = PROXY_PRECHECK_TARGETS
PROXY_PRECHECK_URL = PROXY_PRECHECK_TARGETS[0]


VERIFY_AFTER_REGISTER = True

# Wait after captcha becomes actionable before the first/normal press.

INITIAL_PRESS_DELAY = 5

# Once the hold starts, begin checking the PX hold label after 5s and stop early

# when it flips to display:none instead of blindly waiting to the end.

PX_HOLD_EARLY_RELEASE_AFTER = 5.0

PX_HOLD_EARLY_RELEASE_INTERVAL = 0.5

PX_HOLD_SECONDS_MIN = 10.5

PX_HOLD_SECONDS_MAX = 11.0

# Give up after max_press if no redirect happens within this many seconds.

POST_MAX_PRESS_WAIT = 5

# After a press, wait this long before retrying even if loading is not seen.

POST_PRESS_LOADING_CHECK = 8

# Max time to stay in one captcha wait state (reappear/validating/unclear).

CAPTCHA_STATE_TIMEOUT = float(os.environ.get("OUTLOOK_RUOYI_CAPTCHA_STATE_TIMEOUT", "20") or "20")

# 真按钮异步渲染等待上限：PX iframe 已出现但 "Press and hold" 按钮还没渲染时，等这么久。

BUTTON_RENDER_TIMEOUT = float(os.environ.get("OUTLOOK_RUOYI_BUTTON_RENDER_TIMEOUT", "30") or "30")

BIRTHDAY_ENTRY_TIMEOUT = float(os.environ.get("OUTLOOK_RUOYI_BDAY_ENTRY_TIMEOUT", "2.5") or "2.5")

BIRTHDAY_SUBMIT_TIMEOUT = float(os.environ.get("OUTLOOK_RUOYI_BDAY_SUBMIT_TIMEOUT", "2.0") or "2.0")

BROWSER_QUIT_TIMEOUT = float(os.environ.get("OUTLOOK_RUOYI_BROWSER_QUIT_TIMEOUT", _ruyi_pkg._state.BROWSER_QUIT_TIMEOUT) or _ruyi_pkg._state.BROWSER_QUIT_TIMEOUT)
_ruyi_pkg._state.BROWSER_QUIT_TIMEOUT = BROWSER_QUIT_TIMEOUT

RUOYI_PROFILE_CLEANUP_RETRIES = int(os.environ.get("OUTLOOK_RUOYI_PROFILE_CLEANUP_RETRIES", _ruyi_pkg._state.RUOYI_PROFILE_CLEANUP_RETRIES) or _ruyi_pkg._state.RUOYI_PROFILE_CLEANUP_RETRIES)
_ruyi_pkg._state.RUOYI_PROFILE_CLEANUP_RETRIES = RUOYI_PROFILE_CLEANUP_RETRIES

RUOYI_PROFILE_CLEANUP_RETRY_DELAY = float(os.environ.get("OUTLOOK_RUOYI_PROFILE_CLEANUP_RETRY_DELAY", _ruyi_pkg._state.RUOYI_PROFILE_CLEANUP_RETRY_DELAY) or _ruyi_pkg._state.RUOYI_PROFILE_CLEANUP_RETRY_DELAY)
_ruyi_pkg._state.RUOYI_PROFILE_CLEANUP_RETRY_DELAY = RUOYI_PROFILE_CLEANUP_RETRY_DELAY

# Retry gap after a failed challenge before the next press.

POST_PRESS_RETRY_GAP_MIN = 2.0

POST_PRESS_RETRY_GAP_MAX = 4.0

# Slow down page start; actual submit now waits once with a random 0-max delay.

PAGE_START_DELAY = float(os.environ.get("OUTLOOK_RUOYI_PAGE_START_DELAY", "2") or "2")

SUBMIT_DELAY = float(os.environ.get("OUTLOOK_RUOYI_SUBMIT_DELAY", "0.5") or "0.5")

# UA 池常量 + HEADLESS_*:re-export 自新包(单一来源,中性 env 名通过兼容层提升)
# 旧 OUTLOOK_RUOYI_HEADLESS_* 提升到中性 RUOYI_HEADLESS_*,新包读中性名时拿到旧值
os.environ.setdefault("RUOYI_HEADLESS_UA", os.environ.get("OUTLOOK_RUOYI_HEADLESS_UA", ""))
os.environ.setdefault("RUOYI_HEADLESS_WIDTH", os.environ.get("OUTLOOK_RUOYI_HEADLESS_WIDTH", ""))
os.environ.setdefault("RUOYI_HEADLESS_HEIGHT", os.environ.get("OUTLOOK_RUOYI_HEADLESS_HEIGHT", ""))
HEADLESS_WINDOW_WIDTH = _ruyi_pkg.HEADLESS_WINDOW_WIDTH
HEADLESS_WINDOW_HEIGHT = _ruyi_pkg.HEADLESS_WINDOW_HEIGHT
_DEFAULT_UA_POOL = _ruyi_pkg._DEFAULT_UA_POOL
HEADLESS_USER_AGENT = _ruyi_pkg.HEADLESS_USER_AGENT

# 并发启动错峰：同一批拿到 slot 后，相邻两个浏览器启动至少间隔这么多秒。

# 4 并发默认 10s → 约 0/10/20/30s 错峰拉满，避免四窗同时砸 signup。

LAUNCH_STAGGER_SECONDS = float(os.environ.get("OUTLOOK_RUOYI_LAUNCH_STAGGER", "10") or "10")

# firefox 退出/强杀等待超时:读旧名,回落新包中性默认,同步写回 _state
RUOYI_FIREFOX_EXIT_WAIT = float(os.environ.get("OUTLOOK_RUOYI_FIREFOX_EXIT_WAIT", _ruyi_pkg._state.RUOYI_FIREFOX_EXIT_WAIT) or _ruyi_pkg._state.RUOYI_FIREFOX_EXIT_WAIT)
_ruyi_pkg._state.RUOYI_FIREFOX_EXIT_WAIT = RUOYI_FIREFOX_EXIT_WAIT

RUOYI_FIREFOX_FORCEKILL_WAIT = float(os.environ.get("OUTLOOK_RUOYI_FIREFOX_FORCEKILL_WAIT", _ruyi_pkg._state.RUOYI_FIREFOX_FORCEKILL_WAIT) or _ruyi_pkg._state.RUOYI_FIREFOX_FORCEKILL_WAIT)
_ruyi_pkg._state.RUOYI_FIREFOX_FORCEKILL_WAIT = RUOYI_FIREFOX_FORCEKILL_WAIT

# UA 池旧名提升到中性名,新包 _load_ua_pool/_pick_user_agent 读中性名
os.environ.setdefault("RUOYI_UA_POOL", os.environ.get("OUTLOOK_RUOYI_UA_POOL", ""))

_HELPERS = None

# C/F/E 全局状态:re-export 自新包 _state(单一源,跨模块同步)
_UA_RR_LOCK = _ruyi_pkg._state._UA_RR_LOCK
_UA_RR_IDX = _ruyi_pkg._state._UA_RR_IDX
_ACTIVE_BROWSER_PAGES = _ruyi_pkg._state._ACTIVE_BROWSER_PAGES
_ACTIVE_BROWSER_LOCK = _ruyi_pkg._state._ACTIVE_BROWSER_LOCK
_SHUTDOWN_HANDLERS_INSTALLED = _ruyi_pkg._state._SHUTDOWN_HANDLERS_INSTALLED
_ACTIVE_RUOYI_PROFILE_DIRS = _ruyi_pkg._state._ACTIVE_RUOYI_PROFILE_DIRS
_ACTIVE_RUOYI_PROFILE_DIRS_LOCK = _ruyi_pkg._state._ACTIVE_RUOYI_PROFILE_DIRS_LOCK






# _load_ua_pool / _track_browser_page / _untrack_browser_page / _close_tracked_browser_pages:
# re-export 自新包(单一源)。新包 _load_ua_pool 读中性名 RUOYI_UA_POOL,
# 上方 setdefault 已把 OUTLOOK_RUOYI_UA_POOL 提升到中性名,行为等价。
_load_ua_pool = _ruyi_pkg._load_ua_pool
_track_browser_page = _ruyi_pkg._track_browser_page
_untrack_browser_page = _ruyi_pkg._untrack_browser_page
_close_tracked_browser_pages = _ruyi_pkg._close_tracked_browser_pages

# E 组 profile 清理:re-export 自新包(单一源,围栏重构去 ROOT 依赖)。
# 调用方传 profile_root 或走 _state._PROFILE_ROOT(上方 set_profile_root 已注入业务根)。
# _cleanup_stale_ruoyi_profile_root 的 max_age_sec 默认读 _state.RUOYI_PROFILE_STALE_SEC,
# 兼容层已把旧名值同步写回 _state,行为等价。
_cleanup_ruoyi_profile_root = _ruyi_pkg._cleanup_ruoyi_profile_root
_cleanup_stale_ruoyi_profile_root = _ruyi_pkg._cleanup_stale_ruoyi_profile_root
_cleanup_ruoyi_slot_root = _ruyi_pkg._cleanup_ruoyi_slot_root
_cleanup_ruoyi_run_profile_dir = _ruyi_pkg._cleanup_ruoyi_run_profile_dir

# G 组 firefox 进程树 kill:re-export 自新包(单一源,firefox_path 参数化,去 Administrator 硬编码)。
# _ruoyi_firefox_running_count_by_path / _force_kill_ruoyi_firefox / _kill_ruoyi_firefox_by_profile
# 默认 firefox_path=get_firefox_path()(PEP 562 延迟求值,读 RUOYI_FIREFOX_PATH)。
_ruyi_firefox_running_count_by_path = _ruyi_pkg._ruoyi_firefox_running_count_by_path
_force_kill_ruoyi_firefox = _ruyi_pkg._force_kill_ruoyi_firefox
_kill_ruoyi_firefox_by_profile = _ruyi_pkg._kill_ruoyi_firefox_by_profile

# F 组 page 生命周期/quit/shutdown + UA 选取:re-export 自新包(单一源)。
# _quit_browser_page 默认 timeout 读 _state.BROWSER_QUIT_TIMEOUT(兼容层已同步)。
# _install_shutdown_handlers 写 _state._SHUTDOWN_HANDLERS_INSTALLED(全局态单源)。
# _pick_user_agent 读中性名 RUOYI_UA_POOL(上方 setdefault 已提升旧名),_UA_RR_IDX/LOCK 走 _state。
_quit_browser_page = _ruyi_pkg._quit_browser_page
_install_shutdown_handlers = _ruyi_pkg._install_shutdown_handlers
_pick_user_agent = _ruyi_pkg._pick_user_agent

# A 组纯函数:re-export 自新包(单一源,零业务依赖)。
_mask_ua = _ruyi_pkg._mask_ua
_env_bool = _ruyi_pkg._env_bool

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





def set_log_level(value):

    global LOG_LEVEL

    LOG_LEVEL = _normalize_log_level(value)

    os.environ["OUTLOOK_LOG_LEVEL"] = LOG_LEVEL

    return LOG_LEVEL





def _log_level_value(value):

    return LOG_LEVELS.get(_normalize_log_level(value), LOG_LEVELS["INFO"])





def _should_keep_prod_log(msg, level):

    if _normalize_log_level(level) == "ERR":

        return True

    low = str(msg or "").strip().lower()

    if not low:

        return False

    if low.startswith((

        "开始:",

        "proxy list ready:",

        "代理 list 就绪:",

        "========== 注册 #",

        "#",

        "done:",

        "summary:",

        "summary_time:",

        "px_summary:",

        "px_detail:",

    )):

        if low.startswith("#") and " 代理 -> " not in str(msg or ""):

            return any(token in low for token in (" ok: ", "结果:", "result:", "授权结果:"))

        return True

    keep_tokens = (

        " ok: ",

        "结果:",

        "result:",

        "授权结果:",

        # TG 通知成败日志：PROD 下也保留，避免发送失败被静默吞掉无法诊断

        "tg 通知",

        "send_tg_message",

        "tg 文本",

    )

    return any(token in low for token in keep_tokens)





def _should_demote_to_debug(msg, level):

    normalized = _normalize_log_level(level)

    if normalized not in {"INFO", "OK", "STEP"}:

        return False

    low = str(msg or "").lower()

    if "press #" in low:

        return False

    debug_patterns = (

        "账号格式:",

        "代理池准备完毕:",

        # 注意：开始/完成汇总、结果/授权结果 不能降级，WebUI 靠这些行抽 success/fail/未授权

        "email_nograph:",

        "screenshot:",

        "screenshot empty",

        "页面已保存",

        "har saved",

        "挂载 ",

        "启动 ruyipage firefox:",

        "使用当前打开页面承载注册页",

        "已关闭 firefox 启动默认空白页",

        "ruoyi headless options applied",

        "ruoyi headless emulation applied",

        "ruoyi headless page patches applied",

        "step open_signup:",

        "step wait_loading:",

        "step handle_consent:",

        "step confirm_before_register:",

        "step post_signup_cleanup:",

        "step verify_registered_outlook:",

        "browser model:",

        "current ip:",

        "filled email:",

        "filled prefix",

        "密码已填",

        "生日页 select 数=",

        "无 select，用 combobox",

        "combo[",

        "month=ok",

        "month=fail",

        "day=ok",

        "day=fail",

        "年份=",

        "年份(js)=",

        "name page start wait",

        "name(generic):",

        "name:",

        "name enter",

        "checked terms",

        "checked required checkbox",

        "timings:",

        "graph proxy ->",

        "waiting for captcha reappear",

        "captcha still validating",

        "waiting for post-captcha redirect",

        "microsoft loading still active",

        "post-captcha state unclear",

        "captcha visible, wait ",

        "focused page before captcha press",

        "submit random wait",

        " next: ",

        "challenge failed",

        "batch:",

        "ua pool pick",

        "browser ua override",

        "launch stagger wait",

        # DOM 步骤探测/轮询细节：默认 INFO 刷屏，降到 DEBUG

        "dom step=",

        "email submit outcome=",

        "stuck overridden",

        "pending resolved",

        "suggestion submit outcome=",

        "email step already advanced",

        "password wait",

        "after password",

        "birthday enter",

        "birthday left",

        "birthday controls not present",

        "still on birthday after submit",

    )

    return any(pat in low for pat in debug_patterns)





def log(msg, level="INFO"):

    rendered = str(level or "INFO").strip().upper() or "INFO"

    if LOG_LEVEL == "PROD":

        if not _should_keep_prod_log(msg, rendered):

            return

    else:

        effective = "DEBUG" if _should_demote_to_debug(msg, rendered) else rendered

        if _log_level_value(effective) < _log_level_value(LOG_LEVEL):

            return

    print(f"[{datetime.now().strftime('%H:%M:%S')}] [{rendered}] {msg}", flush=True)





def debug(msg):

    log(msg, "DEBUG")





def _timed_step(tag, name, fn, *args, detail=None, **kwargs):

    started = time.perf_counter()

    result = fn(*args, **kwargs)

    elapsed = time.perf_counter() - started

    extra = ""

    if callable(detail):

        try:

            extra = detail(result) or ""

        except Exception:

            extra = ""

    elif detail:

        extra = str(detail)

    suffix = f" {extra}" if extra else ""

    log(f"  {tag} step {name}: {elapsed:.2f}s{suffix}", "INFO")

    return result, elapsed





def _summarize_batch_metrics(results, elapsed_list, total_elapsed, px_stats=None):

    success_statuses = {"ok", "no_graph", "reg_only", True}

    fail_statuses = {"fail", False}

    success_elapsed = [

        float(elapsed)

        for status, elapsed in zip(results or [], elapsed_list or [])

        if status in success_statuses

    ]

    success_count = len(success_elapsed)

    fail_count = sum(1 for status in (results or []) if status in fail_statuses)

    if results:

        paired_px_stats = list(zip(results or [], px_stats or []))

        success_px_stats = [

            stat for status, stat in paired_px_stats

            if status in success_statuses

        ]

    else:

        success_px_stats = list(px_stats or [])

    active_px_stats = [

        stat for stat in success_px_stats

        if float((stat or {}).get("px_elapsed") or 0.0) > 0.0 or int((stat or {}).get("max_presses") or 0) > 0

    ]

    return {

        "total_elapsed": float(total_elapsed or 0.0),

        "success_count": success_count,

        "fail_count": fail_count,

        "avg_success_elapsed": (sum(success_elapsed) / success_count) if success_count else 0.0,

        "max_px_presses": max((int((stat or {}).get("max_presses") or 0) for stat in success_px_stats), default=0),

        "avg_px_elapsed": (

            sum(float((stat or {}).get("px_elapsed") or 0.0) for stat in active_px_stats) / len(active_px_stats)

        ) if active_px_stats else 0.0,

        "px_stats": success_px_stats,

    }





def _format_batch_summary_lines(ok, no_graph, failed, total, total_elapsed, avg_success_elapsed, px_stats=None):

    summary = _summarize_batch_metrics([], [], total_elapsed, px_stats=px_stats)

    lines = [

        f"DONE: success={ok}/{total} fail={failed} no_graph={no_graph}",

        f"SUMMARY: success {ok} | fail {failed} | no_graph {no_graph} | total {total}",

        f"SUMMARY_TIME: total_elapsed {float(total_elapsed or 0.0):.2f}s | avg_success_elapsed {float(avg_success_elapsed or 0.0):.2f}s",

    ]

    if px_stats:

        lines.append(

            f"PX_SUMMARY: max_presses {int(summary['max_px_presses'] or 0)} | avg_px_elapsed {float(summary['avg_px_elapsed'] or 0.0):.2f}s"

        )

        for stat in sorted((px_stats or []), key=lambda item: int((item or {}).get("idx") or 0)):

            lines.append(

                f"PX_DETAIL: #{int((stat or {}).get('idx') or 0)} max_presses {int((stat or {}).get('max_presses') or 0)} | px_elapsed {float((stat or {}).get('px_elapsed') or 0.0):.2f}s | reg_elapsed {float((stat or {}).get('reg_elapsed') or 0.0):.2f}s"

            )

    return lines


# send_tg_message / Telegram 通知已抽到 common/notify.py，供多流程复用。


def _count_account_lines(path):
    """统计账号文件非空非注释行数(emails.txt / email_nograph.txt)。"""
    if not path or not os.path.isfile(path):
        return 0
    try:
        n = 0
        with open(path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#"):
                    n += 1
        return n
    except Exception:
        return 0


def _notify_tg(args, summary_lines, batch_no=None):

    """把批次汇总按日志同款前缀渲染后发 TG；未配置 token/chat_id 则跳过。"""

    bot_token = getattr(args, "tg_bot_token", "") or os.environ.get("TG_BOT_TOKEN", "")

    chat_id = getattr(args, "tg_chat_id", "") or os.environ.get("TG_CHAT_ID", "")

    proxy = getattr(args, "tg_proxy", "") or os.environ.get("TG_PROXY", "")

    if not bot_token or not chat_id or not summary_lines:

        return

    # TG 通知只发标题 + 摘要纯内容(去掉 SUMMARY:/SUMMARY_TIME: 前缀和日志前缀)，
    # 过滤 DONE/PX_SUMMARY/PX_DETAIL(明细在日志里)，末尾追加 email/no_graph 累计总数。

    head_lines = [ln for ln in summary_lines if str(ln).startswith("SUMMARY")]
    lines = [ln.split(":", 1)[1].strip() if ":" in ln else ln for ln in head_lines]

    if batch_no is not None:
        lines.insert(0, f"🔁 ruoyi 养号 第 {batch_no} 批完成")
    else:
        lines.insert(0, "🔁 ruoyi 养号 注册完成")

    lines.append(f"email 总数: {_count_account_lines(EMAILS_POOL)}")
    lines.append(f"no_graph 总数: {_count_account_lines(EMAIL_NOGRAPH)}")

    send_tg_message("\n".join(lines), bot_token, chat_id,
                    proxy=proxy or None, log_fn=log)





def _normalize_px_metrics(idx, metrics=None):

    data = dict(metrics or {})

    return {

        "idx": int(data.get("idx") or idx or 0),

        "max_presses": int(data.get("max_presses") or 0),

        "px_elapsed": float(data.get("px_elapsed") or 0.0),
        "reg_elapsed": float(data.get("reg_elapsed") or 0.0),

    }





set_log_level(os.environ.get("OUTLOOK_LOG_LEVEL", "INFO"))





# A 组纯函数 _browser_model_name:re-export 自新包(单一源)。
_browser_model_name = _ruyi_pkg._browser_model_name

def _load_helpers():

    global _HELPERS

    if _HELPERS is not None:

        return _HELPERS

    path = os.path.join(ROOT, "register_outlook_standalone.py")

    spec = importlib.util.spec_from_file_location("_outlook_standalone_helpers", path)

    mod = importlib.util.module_from_spec(spec)

    spec.loader.exec_module(mod)

    _HELPERS = mod

    return mod





# A/D/B 组代理工具 + 代理池 + IP/代理探测:re-export 自新包(单一源,零业务依赖)。
# 池单例 _PROXY_LIST 经 _state 读写(get_consumable_proxy_pool/set_consumable_proxy_pool),
# 兼容旧名 get_session_proxy_runtime/set_session_proxy_runtime 同源。
# _probe_proxy_identity / _proxy_exit_key 默认 timeout 读 _state.PROXY_IDENTITY_TIMEOUT(已同步)。
# load_proxy_batch = load_proxy_list 兼容旧名(新包已别名)。
_strip_proxy_scheme = _ruyi_pkg._strip_proxy_scheme
_proxy_url_to_ruoyi = _ruyi_pkg._proxy_url_to_ruoyi
_parse_ruoyi_proxy = _ruyi_pkg._parse_ruoyi_proxy
mask_ruoyi_proxy = _ruyi_pkg.mask_ruoyi_proxy
_proxy_host_port_key = _ruyi_pkg._proxy_host_port_key
_proxy_identity_cache_get = _ruyi_pkg._proxy_identity_cache_get
_proxy_identity_cache_put = _ruyi_pkg._proxy_identity_cache_put
_coerce_float_or_none = _ruyi_pkg._coerce_float_or_none
_geo_timezone_from_payload = _ruyi_pkg._geo_timezone_from_payload
_probe_proxy_identity = _ruyi_pkg._probe_proxy_identity
_proxy_exit_key = _ruyi_pkg._proxy_exit_key
_parse_proxy_lines = _ruyi_pkg._parse_proxy_lines
parse_proxy_pool = _ruyi_pkg.parse_proxy_pool
fetch_proxy_list_http = _ruyi_pkg.fetch_proxy_list_http
_proxy_source_label = _ruyi_pkg._proxy_source_label
load_proxy_list = _ruyi_pkg.load_proxy_list
load_proxy_batch = _ruyi_pkg.load_proxy_batch
build_proxy_source = _ruyi_pkg.build_proxy_source
ConsumableProxyPool = _ruyi_pkg.ConsumableProxyPool
get_consumable_proxy_pool = _ruyi_pkg.get_consumable_proxy_pool
set_consumable_proxy_pool = _ruyi_pkg.set_consumable_proxy_pool
get_session_proxy_runtime = _ruyi_pkg.get_session_proxy_runtime
set_session_proxy_runtime = _ruyi_pkg.set_session_proxy_runtime
SessionProxyRuntime = _ruyi_pkg.SessionProxyRuntime
select_proxy_for_account = _ruyi_pkg.select_proxy_for_account
release_proxy_for_account = _ruyi_pkg.release_proxy_for_account

@contextmanager

def _interprocess_lock(target_path):

    lock_path = f"{target_path}.lock"

    os.makedirs(os.path.dirname(os.path.abspath(lock_path)), exist_ok=True)

    fh = open(lock_path, "a+b")

    try:

        fh.seek(0, os.SEEK_END)

        if fh.tell() == 0:

            fh.write(b"0")

            fh.flush()

        fh.seek(0)

        if os.name == "nt":

            import msvcrt



            while True:

                try:

                    msvcrt.locking(fh.fileno(), msvcrt.LK_LOCK, 1)

                    break

                except OSError:

                    time.sleep(0.05)

        else:

            import fcntl



            fcntl.flock(fh.fileno(), fcntl.LOCK_EX)

        yield

    finally:

        try:

            fh.seek(0)

            if os.name == "nt":

                import msvcrt



                msvcrt.locking(fh.fileno(), msvcrt.LK_UNLCK, 1)

            else:

                import fcntl



                fcntl.flock(fh.fileno(), fcntl.LOCK_UN)

        finally:

            fh.close()



def append_graph_account_to_emails_pool(email, password, graph):

    token = (graph or {}).get("refresh_token") or ""

    client_id = (graph or {}).get("client_id") or ""

    if not token:

        log(f"emails.txt skip {email}: no graph refresh_token", "WARN")

        return False

    try:

        existing = set()

        with _interprocess_lock(EMAILS_POOL):

            if os.path.isfile(EMAILS_POOL):

                with open(EMAILS_POOL, encoding="utf-8") as f:

                    for line in f:

                        line = line.strip()

                        if line and not line.startswith("#"):

                            existing.add(line.split("----")[0].strip().lower())

            if email.lower() in existing:

                return True

            with open(EMAILS_POOL, "a", encoding="utf-8") as f:

                f.write(f"{email}----{password}----{client_id}----{token}----{graph.get('cf_address', '')}----{graph.get('cf_password') or ''}\n")

        log(f"emails.txt += {email} (token=yes)", "OK")

        return True

    except Exception as exc:

        log(f"append_graph_account_to_emails_pool failed: {type(exc).__name__}: {exc}", "WARN")

        return False





def append_account_to_email_nograph(email, password):

    if not email or not password:

        return False

    try:

        os.makedirs(OUTPUT_DIR, exist_ok=True)

        existing = set()

        with _interprocess_lock(EMAIL_NOGRAPH):

            if os.path.isfile(EMAIL_NOGRAPH):

                with open(EMAIL_NOGRAPH, encoding="utf-8") as f:

                    for line in f:

                        line = line.strip()

                        if line and not line.startswith("#"):

                            existing.add(line.split("----")[0].strip().lower())

            if email.lower() in existing:

                return True

            with open(EMAIL_NOGRAPH, "a", encoding="utf-8") as f:

                f.write(f"{email}----{password}\n")

        log(f"email_nograph += {email}", "OK")

        return True

    except Exception as exc:

        log(f"append_account_to_email_nograph failed: {type(exc).__name__}: {exc}", "WARN")

        return False



def append_account_to_email_reg(email, password):

    """仅注册(跳过 Graph 授权)账号追加到 email_reg.txt，格式 email----password。"""

    if not email or not password:

        return False

    try:

        existing = set()

        with _interprocess_lock(EMAIL_REG):

            if os.path.isfile(EMAIL_REG):

                with open(EMAIL_REG, encoding="utf-8") as f:

                    for line in f:

                        line = line.strip()

                        if line and not line.startswith("#"):

                            existing.add(line.split("----")[0].strip().lower())

            if email.lower() in existing:

                return True

            with open(EMAIL_REG, "a", encoding="utf-8") as f:

                f.write(f"{email}----{password}\n")

        log(f"email_reg += {email}", "OK")

        return True

    except Exception as exc:

        log(f"append_account_to_email_reg failed: {type(exc).__name__}: {exc}", "WARN")

        return False



def _append_nograph_account(email, password, nograph_file):

    """no_graph 账号写入 per-run accounts_ruoyi_nograph_{ts}.txt，与成功号 live_file 对称。"""

    if not email or not password or not nograph_file:

        return False

    try:

        os.makedirs(os.path.dirname(nograph_file) or OUTPUT_DIR, exist_ok=True)

        with _interprocess_lock(nograph_file):

            with open(nograph_file, "a", encoding="utf-8") as f:

                f.write(f"{email}----{password}\n")

        return True

    except Exception as exc:

        log(f"_append_nograph_account failed: {type(exc).__name__}: {exc}", "WARN")

        return False



def _shot(page, name, idx):

    if _skip_duplicate_failure_screenshot(name, idx):

        return None

    failure_prefixes = (

        "blocked",

        "error",

        "timeout",

        "press_fail",

        "captcha_no_target",

        "email_fail",

        "email_input_fail",

        "email_empty_value",

        "email_exc",

        "email_stuck",

        "no_email",

        "pwd_fail",

        "bday_fail",

        "name_fail",

    )

    if not _should_save_failure_shot(name):

        return None

    os.makedirs(SCREENSHOT_DIR, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]

    base = os.path.join(SCREENSHOT_DIR, f"ruoyi_{idx}_{name}_{ts}")

    path = f"{base}.png"

    try:

        page.screenshot(path=path, full_page=True)

        if os.path.isfile(path) and os.path.getsize(path) > 0:

            log(f"截图已保存: {path}", "OK")

        else:

            log(f"截图为空或缺失: {path}", "WARN")

    except Exception as e:

        log(f"截图保存失败 {path}: {e}", "WARN")

    try:

        state_path = f"{base}.html"

        url = getattr(page, "url", "") or ""

        title = getattr(page, "title", "") or ""

        try:

            html = page.run_js_loaded("return document.documentElement ? document.documentElement.outerHTML : '';") or ""

        except Exception as exc:

            html = f"<!-- main html capture failed: {type(exc).__name__}: {exc} -->"

        try:

            text = page.run_js_loaded("return document.documentElement ? document.documentElement.innerText : '';") or ""

        except Exception as exc:

            text = f"text capture failed: {type(exc).__name__}: {exc}"

        frame_parts = []

        try:

            frames = page.get_all_frames() or []

        except Exception:

            frames = []

        for fi, frame in enumerate(frames, 1):

            frame_url = getattr(frame, "url", "") or ""

            try:

                frame_text = frame.run_js_loaded(

                    "return document.documentElement ? document.documentElement.innerText : '';"

                ) or ""

            except Exception as exc:

                frame_text = f"frame text capture failed: {type(exc).__name__}: {exc}"

            try:

                frame_html = frame.run_js_loaded(

                    "return document.documentElement ? document.documentElement.outerHTML : '';"

                ) or ""

            except Exception as exc:

                frame_html = f"<!-- frame html capture failed: {type(exc).__name__}: {exc} -->"

            frame_parts.append(

                "\n\n"

                f"<!-- frame #{fi} url={frame_url!r} -->\n"

                "<pre data-codex-frame-text>\n"

                + frame_text.replace("</pre>", "<\\/pre>")

                + "\n</pre>\n"

                + frame_html

            )

        state = (

            "<!doctype html>\n"

            "<meta charset=\"utf-8\">\n"

            f"<!-- captured_at={ts} idx={idx} name={name!r} url={url!r} title={title!r} -->\n"

            "<pre data-codex-page-text>\n"

            + text.replace("</pre>", "<\\/pre>")

            + "\n</pre>\n"

            + html

            + "".join(frame_parts)

        )

        with open(state_path, "w", encoding="utf-8") as f:

            f.write(state)

        log(f"页面已保存: {state_path}", "OK")

    except Exception as e:

        log(f"页面保存失败 {base}.html: {e}", "WARN")

    return path





_PRESS_FAIL_SCREENSHOT_DEDUP_SEC = 8.0
_RECENT_SCREENSHOT_KEYS = {}
_RECENT_SCREENSHOT_LOCK = threading.Lock()


def _skip_duplicate_failure_screenshot(name, idx):

    if not str(name or "").startswith("press_fail"):
        return False
    key = (int(idx or 0), "press_fail")
    now = time.time()
    with _RECENT_SCREENSHOT_LOCK:
        last = float(_RECENT_SCREENSHOT_KEYS.get(key) or 0.0)
        if now - last < _PRESS_FAIL_SCREENSHOT_DEDUP_SEC:
            return True
        _RECENT_SCREENSHOT_KEYS[key] = now
    return False


def _save_screenshot(page, name, idx, tag=None):

    if _skip_duplicate_failure_screenshot(name, idx):

        return None

    os.makedirs(SCREENSHOT_DIR, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]

    path = os.path.join(SCREENSHOT_DIR, f"ruoyi_{idx}_{name}_{ts}.png")

    try:

        page.screenshot(path=path, full_page=True)

        if os.path.isfile(path) and os.path.getsize(path) > 0:

            prefix = f"  {tag} " if tag else ""

            log(f"{prefix}screenshot: {path}", "OK")

        else:

            prefix = f"  {tag} " if tag else ""

            log(f"{prefix}screenshot empty/missing: {path}", "WARN")

    except Exception as e:

        prefix = f"  {tag} " if tag else ""

        log(f"{prefix}screenshot failed {path}: {e}", "WARN")

    return path





class _RuoyiHarCollector:

    """Collect ruyipage listener packets and save a compact HAR-like JSON on failure."""



    def __init__(self, page, tag, idx, email_getter=None):

        self.page = page

        self.tag = tag

        self.idx = idx

        self.email_getter = email_getter or (lambda: "")

        self.entries = []

        self._stop = threading.Event()

        self._thread = None

        self.started = False



    def start(self):

        capture = getattr(self.page, "capture", None)

        if capture is None:

            log(f"  {self.tag} HAR capture API unavailable", "WARN")

            return False

        try:

            capture.start(targets=True, collect_bodies=True)

            self.started = True

        except Exception as exc:

            log(f"  {self.tag} HAR capture start failed: {type(exc).__name__}: {exc}", "WARN")

            return False



        def _worker():

            while not self._stop.is_set():

                try:

                    packet = capture.wait(timeout=0.5)

                except Exception:

                    continue

                if packet is not None:

                    try:

                        self.entries.append(self._packet_to_entry(packet))

                    except Exception as exc:

                        self.entries.append(self._fallback_entry(packet, exc))



        self._thread = threading.Thread(target=_worker, name=f"ruoyi-har-{self.idx}", daemon=True)

        self._thread.start()

        log(f"  {self.tag} HAR capture started")

        return True



    def stop(self):

        self._stop.set()

        if self._thread is not None:

            try:

                self._thread.join(timeout=1.5)

            except Exception:

                pass

        try:

            if self.started and getattr(self.page, "capture", None) is not None:

                self.page.capture.stop()

        except Exception:

            pass



    @staticmethod

    def _headers_to_list(headers):

        if not isinstance(headers, dict):

            return []

        return [{"name": str(k), "value": str(v)} for k, v in headers.items()]



    @staticmethod

    def _safe_started_datetime(value):

        try:

            ts = float(value or 0)

            if ts > 100000000000:

                ts = ts / 1000.0

            if ts <= 0 or ts > 4102444800:

                ts = time.time()

            return datetime.fromtimestamp(ts).isoformat()

        except Exception:

            return datetime.now().isoformat()



    def _fallback_entry(self, packet, exc):

        url = ""

        method = ""

        status = 0

        try:

            url = getattr(packet, "url", "") or ""

            method = getattr(packet, "method", "") or ""

            status = int(getattr(packet, "response_status", 0) or getattr(packet, "status", 0) or 0)

        except Exception:

            pass

        return {

            "startedDateTime": datetime.now().isoformat(),

            "time": 0,

            "request": {

                "method": method,

                "url": url,

                "httpVersion": "HTTP/2",

                "headers": [],

                "queryString": [],

                "cookies": [],

                "headersSize": -1,

                "bodySize": -1,

            },

            "response": {

                "status": status,

                "statusText": "collectorError",

                "httpVersion": "HTTP/2",

                "headers": [],

                "cookies": [],

                "content": {"size": -1, "mimeType": ""},

                "redirectURL": "",

                "headersSize": -1,

                "bodySize": -1,

            },

            "cache": {},

            "timings": {"send": 0, "wait": 0, "receive": 0},

            "_collectorError": f"{type(exc).__name__}: {exc}",

        }



    def _packet_to_entry(self, packet):

        url = getattr(packet, "url", "") or ""

        method = getattr(packet, "method", "") or ""

        status = int(getattr(packet, "response_status", 0) or 0)

        req_headers = getattr(packet, "request_headers", None) or {}

        resp_headers = getattr(packet, "response_headers", None) or {}

        event_type = getattr(packet, "event_type", "") or ""

        is_failed = bool(getattr(packet, "is_failed", False))

        req_body = None

        try:

            req_body = packet.request_body

        except Exception:

            req_body = None

        resp_body = None

        try:

            resp_body = packet.response_body

        except Exception:

            resp_body = None

        mime_type = ""

        if isinstance(resp_headers, dict):

            mime_type = str(resp_headers.get("content-type", ""))

        req_body_size = len(req_body) if req_body else 0

        resp_body_size = len(resp_body) if resp_body else 0

        return {

            "startedDateTime": self._safe_started_datetime(getattr(packet, "timestamp", 0)),

            "time": 0,

            "request": {

                "method": method,

                "url": url,

                "httpVersion": "HTTP/2",

                "headers": self._headers_to_list(req_headers),

                "queryString": [],

                "cookies": [],

                "headersSize": -1,

                "bodySize": req_body_size,

                "postData": {"mimeType": str(req_headers.get("content-type", "")) if isinstance(req_headers, dict) else "", "text": req_body} if req_body else None,

            },

            "response": {

                "status": status,

                "statusText": "fetchError" if is_failed else event_type,

                "httpVersion": "HTTP/2",

                "headers": self._headers_to_list(resp_headers),

                "cookies": [],

                "content": {

                    "size": resp_body_size,

                    "mimeType": mime_type,

                    "text": resp_body,

                },

                "redirectURL": "",

                "headersSize": -1,

                "bodySize": resp_body_size,

            },

            "cache": {},

            "timings": {"send": 0, "wait": 0, "receive": 0},

            "_eventType": event_type,

            "_requestId": getattr(packet, "request_id", "") or "",

        }



    def save(self, reason="failure"):

        self.stop()

        os.makedirs(HAR_DIR, exist_ok=True)

        ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]

        try:

            email = (self.email_getter() or "").replace("@", "_at_")

        except Exception:

            email = ""

        suffix = f"_{email}" if email else ""

        path = os.path.join(HAR_DIR, f"ruoyi_{self.idx}_{reason}{suffix}_{ts}.har")

        data = {

            "log": {

                "version": "1.2",

                "creator": {"name": "register_outlook_ruoyi", "version": "har-like-listener"},

                "pages": [],

                "entries": self.entries,

            },

            "_meta": {

                "reason": reason,

                "entry_count": len(self.entries),

                "saved_at": datetime.now().isoformat(),

            },

        }

        with open(path, "w", encoding="utf-8") as f:

            json.dump(data, f, ensure_ascii=False, indent=2)

        log(f"  {self.tag} HAR saved: {path} ({len(self.entries)} entries)", "OK")

        return path





def _ele(page, locator, timeout=None):

    from ruyipage import NoneElement



    try:

        el = page.ele(locator, timeout=timeout)

        if el is None or isinstance(el, NoneElement):

            return None

        return el

    except Exception:

        return None





def _should_save_failure_shot(name):

    failure_prefixes = (

        "blocked",

        "error",

        "timeout",

        "submit_timeout",

        "press_fail",

        "captcha_no_target",

        "email_fail",

        "email_input_fail",

        "email_empty_value",

        "email_exc",

        "email_stuck",

        "no_email",

        "pwd_fail",

        "bday_fail",

        "name_fail",

    )

    raw = str(name or "")

    if not any(raw.startswith(prefix) for prefix in failure_prefixes):

        return False

    return _env_bool("OUTLOOK_PX_PRESS_SCREENSHOTS", False)





def _update_submit_wait_state(

    started_at,

    *,

    submitted=False,

    transitioned=False,

    visible=False,

    validating=False,

    loading=False,

    now=None,

    timeout=SUBMIT_RESULT_TIMEOUT,

):

    if now is None:

        now = time.time()

    if transitioned or visible or validating or loading:

        return None, False

    if submitted:

        return (started_at if started_at is not None else now), False

    if started_at is None:

        return None, False

    return started_at, (now - started_at) >= timeout





def _wait_state_timed_out(started_at, *, now=None, timeout=CAPTCHA_STATE_TIMEOUT):

    if started_at is None:

        return False

    if now is None:

        now = time.time()

    return (now - started_at) >= timeout


def _should_enter_post_press_reappear_wait(awaiting_reappear, press_count, max_press):

    if not awaiting_reappear:

        return False

    # max_press 只限制“不能再按下一次”，不影响“当前这一次按完后等待校验结果”。

    return True


def _update_loading_wait_state(

    press_wait_started,

    loading_wait_started,

    *,

    loading=False,

    now=None,

    timeout=CAPTCHA_STATE_TIMEOUT,

):

    if now is None:

        now = time.time()

    if not loading:

        return press_wait_started, None, False

    loading_wait_started = loading_wait_started if loading_wait_started is not None else now

    return None, loading_wait_started, (now - loading_wait_started) >= timeout





def _eles(page, locator, timeout=None):

    try:

        items = page.eles(locator, timeout=timeout)

        return list(items or [])

    except Exception:

        return []





def _body_text(page):

    for script in (

        "return document.body ? document.body.innerText : '';",

        "return document.documentElement ? document.documentElement.innerText : '';",

    ):

        try:

            return page.run_js_loaded(script) or ""

        except Exception:

            continue

    return ""





def _context_text(ctx):

    for script in (

        "return document.body ? document.body.innerText : '';",

        "return document.documentElement ? document.documentElement.innerText : '';",

    ):

        try:

            return ctx.run_js_loaded(script) or ""

        except Exception:

            continue

    return ""





# F 组 _try_option_call:re-export 自新包(单一源,纯函数)。register 内不再直接调用,
# 仅为 unlock 等历史引用保留模块级名字。
_try_option_call = _ruyi_pkg._try_option_call

# --- headless/anti-detect launch stack: re-exported from common.ruyi (single source) ---
_apply_ruoyi_quiet_prefs = _ruyi_pkg._apply_ruoyi_quiet_prefs
_apply_ruoyi_headless_options = _ruyi_pkg._apply_ruoyi_headless_options
_apply_ruoyi_browser_ua = _ruyi_pkg._apply_ruoyi_browser_ua
_apply_ruoyi_headless_emulation = _ruyi_pkg._apply_ruoyi_headless_emulation
_build_headless_patch_js = _ruyi_pkg._build_headless_patch_js
RUOYI_HEADLESS_PATCH_JS = _ruyi_pkg.RUOYI_HEADLESS_PATCH_JS
_ensure_ruoyi_headless_preload = _ruyi_pkg._ensure_ruoyi_headless_preload
_apply_ruoyi_headless_page_patches = _ruyi_pkg._apply_ruoyi_headless_page_patches




def _on_signup_form(url):

    low = (url or "").lower()

    return "signup.live.com" in low and "privacynotice" not in low


def _captcha_signup_url_changed(signup_url, current_url):

    before = str(signup_url or "").strip().lower()

    now = str(current_url or "").strip().lower()

    if not before or not now or before == now:

        return False

    return _on_signup_form(before) and any(host in now for host in ("live.com", "microsoft.com", "outlook"))







def _scroll_into_view(el):

    """把元素滚进视口，避免 ruyipage 报『无法获取元素可点击坐标』。



    注意：element.run_js 走 BiDi callFunction，script 必须是函数声明

    （function(){...} / ()=>{...}），裸 return / this.xxx 会直接抛 JS 错。

    """

    if el is None:

        return False

    try:

        el.run_js(

            "function(){ if(this&&this.scrollIntoView) this.scrollIntoView({block:'center',inline:'nearest'}); }"

        )

        return True

    except Exception:

        pass

    try:

        # 某些版本用 page 上下文执行

        owner = getattr(el, "owner", None) or getattr(el, "page", None)

        if owner is not None:

            owner.run_js_loaded(

                "function(el){ if(el&&el.scrollIntoView) el.scrollIntoView({block:'center',inline:'nearest'}); return true; }",

                el,

            )

            return True

    except Exception:

        pass

    return False





def _read_input_value(el):

    """读 input 当前 value。优先 ruyipage 原生属性，避免错误的 run_js 形态导致假空。"""

    if el is None:

        return ""

    # 1) 原生 .value property（内部是 (el) => el.value）

    try:

        val = getattr(el, "value", None)

        if val is not None and str(val).strip():

            return str(val).strip()

    except Exception:

        pass

    # 2) HTML attribute

    try:

        val = el.attr("value")

        if val is not None and str(val).strip():

            return str(val).strip()

    except Exception:

        pass

    # 3) 合法 functionDeclaration 形式

    try:

        val = el.run_js("function(){ return (this && this.value != null) ? String(this.value) : ''; }")

        if val is not None and str(val).strip():

            return str(val).strip()

    except Exception:

        pass

    return ""





def _probe_email_value_on_page(page):

    """元素句柄失效时，从页面直接查邮箱框 value（page.run_js 支持 return 包装）。"""

    try:

        val = page.run_js_loaded(

            """

return (() => {

  const sels = [

    'input[type="email"]',

    'input[name="email"]',

    'input[name="MemberName"]',

    '#MemberName',

    '#usernameInput',

    'input[name="Username"]',

    'input[aria-label*="email" i]',

    'input[aria-label="New email"]',

  ];

  for (const s of sels) {

    const el = document.querySelector(s);

    if (el && el.offsetParent !== null) {

      const v = (el.value || el.getAttribute('value') || '').trim();

      if (v) return v;

    }

  }

  return '';

})();

            """

        )

        return str(val or "").strip()

    except Exception:

        return ""





def _domain_control(page):

    """返回 (kind, el)：kind 为 select / fluent / None。



    经典页：select#LiveDomainBoxList

    Fluent UI：button#domainDropdownId（只填本地前缀，域名默认 @outlook.com）

    Fluent 为主路径，先查 dropdown，避免经典 select 空等 0.5s。

    """

    btn = _ele(

        page,

        'css:#domainDropdownId, button[name="domainDropdownName"], button[aria-label*="domain" i], '

        'button[aria-label*="Email domain" i]',

        timeout=0.2,

    )

    if btn is not None:

        return "fluent", btn

    dd = _ele(

        page,

        'css:select[id="LiveDomainBoxList"], select[name="LiveDomainBoxList"], #LiveDomainBoxList',

        timeout=0.2,

    )

    if dd is not None:

        return "select", dd

    return None, None





def _safe_input(el, value, clear=True):

    """填写输入框：先滚入视口；click 坐标失败时改用 JS 赋值，避免二次填邮箱把整窗打崩。"""

    if el is None:

        return False

    text = "" if value is None else str(value)

    _scroll_into_view(el)

    try:

        el.input(text, clear=clear)

        return True

    except Exception as exc1:

        try:

            el.input(text, clear=False)

            return True

        except Exception:

            pass

        try:

            # BiDi callFunction：必须 function 声明，参数走形参，不能用 arguments[]

            ok = el.run_js(

                """

function(v, doClear) {

  const text = String(v ?? '');

  const el = this;

  el.focus();

  const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set

    || Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value')?.set;

  if (doClear) {

    if (setter) setter.call(el, ''); else el.value = '';

    el.dispatchEvent(new Event('input', {bubbles: true}));

  }

  if (setter) setter.call(el, text); else el.value = text;

  el.dispatchEvent(new Event('input', {bubbles: true}));

  el.dispatchEvent(new Event('change', {bubbles: true}));

  return true;

}

                """,

                text,

                bool(clear),

            )

            return bool(ok)

        except Exception as exc2:

            log(

                f"  safe_input 失败: {type(exc1).__name__}: {exc1} / "

                f"{type(exc2).__name__}: {exc2}",

                "WARN",

            )

            return False





def _safe_click(el):

    if el is None:

        return False

    _scroll_into_view(el)

    try:

        el.click_self()

        return True

    except Exception:

        try:

            el.click_self(by_js=True)

            return True

        except Exception:

            try:

                el.run_js("function(){ this.click(); return true; }")

                return True

            except Exception:

                return False





def _click_any(page, locators, timeout=2):

    for sel in locators:

        el = _ele(page, sel, timeout=timeout)

        if el is None:

            continue

        if _safe_click(el):

            return sel

    return None





def _page_start_wait(tag, page_name):

    if PAGE_START_DELAY > 0:

        log(f"  {tag} {page_name} page start wait {PAGE_START_DELAY:g}s")

        time.sleep(PAGE_START_DELAY)





def _submit_wait(tag, phase="submit"):

    if SUBMIT_DELAY > 0:

        delay = random.uniform(0, SUBMIT_DELAY)

        log(f"  {tag} {phase} random wait {delay:.2f}s")

        if delay > 0:

            time.sleep(delay)

        return delay

    return 0





def _click_next(page, tag, wait_before=True, wait_after=False, timeout=1.2):

    if wait_before:

        _submit_wait(tag, "submit")

    sels = [

        'css:button[data-testid="primaryButton"]',

        'css:button[type="submit"]',

        'css:input[type="submit"]',

        '#iSignupAction',

        'css:button[id="iSignupAction"]',

        'text:Next',

        'text:下一步',

        'text:Suivant',

        '#iNext',

    ]

    hit = _click_any(page, sels, timeout=timeout)

    triggered = bool(hit)

    if not hit:

        try:

            page.actions.press("\ue007").perform()

            triggered = True

        except Exception:

            pass

    if wait_after:

        time.sleep(0.15)

    log(f"  {tag} next: {hit or '(Enter)'}")

    return bool(hit)





def _email_input_selector():

    return (

        'css:input[type="email"], input[name="email"], input[name="MemberName"], '

        '#MemberName, #usernameInput, input[name="Username"], '

        'input[aria-label*="new email" i], input[aria-label*="email" i], '

        'input[aria-label="Email"], input[placeholder*="email" i]'

    )





def _password_input_selector():

    # Fluent: floatingLabelInput* type=password, 常无 name=Password

    return (

        'css:input[type="password"], input[name="Password"], input[name="passwd"], '

        '#PasswordInput, input[id*="Password" i], input[autocomplete="new-password"], '

        'input[aria-label*="Password" i], input[aria-label*="password" i], '

        'input[placeholder*="Password" i], input[placeholder*="password" i]'

    )





def _name_input_selector():

    """Fluent UI 用 firstNameInput/lastNameInput；经典页用 FirstName/LastName。"""

    return (

        'css:input[name="firstNameInput"], input[name="lastNameInput"], '

        '#firstNameInput, #lastNameInput, '

        'input[name="FirstName"], input[name="LastName"], #FirstName, #LastName, '

        'input[id*="firstName" i], input[id*="lastName" i], '

        'input[aria-label*="First name" i], input[aria-label*="Last name" i], '

        'input[aria-label*="first name" i], input[aria-label*="last name" i], '

        'input[aria-label*="名" i], input[aria-label*="姓" i]'

    )





def _name_input_present(page, timeout=0.3):

    """真实姓名输入框（不是页面文案里带 name 字样）。"""

    if _ele(page, _name_input_selector(), timeout=timeout) is not None:

        return True

    # Fluent 文案兜底：标题已是 Add your name，且有两个 text input

    try:

        title = page.run_js_loaded(

            "return (document.querySelector('[data-testid=\"title\"]')||{}).textContent||document.title||'';"

        ) or ""

    except Exception:

        title = ""

    low = str(title).lower()

    if any(k in low for k in ("add your name", "your name", "添加你的姓名", "你的姓名")):

        texts = _eles(page, 'css:input[type="text"]:not([type="hidden"])', timeout=0.2)

        if len(texts) >= 2:

            return True

    return False





def _birthday_control_present(page, timeout=0.3):

    """经典 select + Fluent combobox / year input。"""

    if (

        _ele(

            page,

            'css:select[name="BirthMonth"], select[name="BirthDay"], select[name="BirthYear"], '

            'input[name="BirthYear"], #BirthYear, #BirthYearInput, input[name="BirthYearInput"], '

            'button[name="BirthMonth"], button[name="BirthDay"], [data-testid*="birth" i], '

            'input[aria-label*="Birth year" i], input[aria-label*="year" i], '

            'button[aria-label*="Month" i], button[aria-label*="Day" i], '

            'button[name*="Birth" i], [role="combobox"][aria-label*="Month" i], '

            '[role="combobox"][aria-label*="Day" i]',

            timeout=timeout,

        )

        is not None

    ):

        return True

    # Fluent 生日页常有 Month/Day combobox + Year 文本框，无 name=Birth*

    try:

        title = page.run_js_loaded(

            "return (document.querySelector('[data-testid=\"title\"]')||{}).textContent||document.title||'';"

        ) or ""

    except Exception:

        title = ""

    low = str(title).lower()

    if any(k in low for k in ("birthday", "birth date", "出生", "date of birth", "what's your date")):

        combos = _eles(page, 'css:button[role="combobox"], [role="combobox"]', timeout=0.2)

        if len(combos) >= 2:

            return True

    return False





def _password_input_present(page, timeout=0.3):

    if _ele(page, _password_input_selector(), timeout=timeout) is not None:

        return True

    # Fluent 过渡/慢渲染：标题已是 Create your password 且邮箱框消失，也算密码步

    try:

        title = page.run_js_loaded(

            "return (document.querySelector('[data-testid=\"title\"]')||{}).textContent||document.title||'';"

        ) or ""

    except Exception:

        title = ""

    low = str(title).lower()

    if any(

        k in low

        for k in (

            "create your password",

            "create a password",

            "your password",

            "choose a password",

            "创建密码",

            "设置密码",

        )

    ) and not _email_input_present(page, timeout=0.1):

        return True

    return False





def _is_password_page(page):

    # 事实：密码输入框在场；标题兜底防过渡帧漏检。

    return _password_input_present(page, timeout=0.4)





def _is_birthday_page(page):

    # 事实：生日控件在场。邮箱/密码框在场时否。

    if _email_input_present(page, timeout=0.15) or _password_input_present(page, timeout=0.15):

        return False

    return _birthday_control_present(page, timeout=0.3)





def _is_name_page(page):

    # 事实：姓名输入框在场。禁止裸词 name / 定时推断。

    if _email_input_present(page, timeout=0.15) or _password_input_present(page, timeout=0.15):

        return False

    return _name_input_present(page, timeout=0.3)





def _resolve_birthday_entry_step(page, max_wait=BIRTHDAY_ENTRY_TIMEOUT):

    if _is_name_page(page):

        return "name"

    if _is_birthday_page(page):

        return "birthday"

    return _wait_signup_step(

        page,

        want=("birthday", "name", "password"),

        max_wait=max_wait,

        poll=0.15,

    )





def _wait_after_birthday_submit_step(page, max_wait=BIRTHDAY_SUBMIT_TIMEOUT):

    if _is_name_page(page):

        return "name"

    if not _is_birthday_page(page):

        step = _detect_signup_step(page)

        if step and step != "birthday":

            return step

    return _wait_signup_step(

        page,

        want=("name", "password", "email", "email_taken", "blocked", "problem"),

        leave=("birthday",),

        max_wait=max_wait,

        poll=0.15,

    )





def _detect_signup_step(page):

    """唯一事实源：当前注册处于哪一步。只看 DOM 控件/错误态，不看计时。



    返回:

      password | birthday | name | email_taken | email_format | email |

      blocked | problem | unknown

    优先级：真实下一步控件 > 邮箱错误态 > 邮箱表单 > 全局错误页。

    """

    # 1) 下一步控件（硬事实）

    if _password_input_present(page, timeout=0.25):

        return "password"

    if _name_input_present(page, timeout=0.2):

        return "name"

    if _birthday_control_present(page, timeout=0.2):

        return "birthday"



    # 2) 邮箱表单事实

    has_email = _email_input_present(page, timeout=0.25)

    err = _email_error_kind(page) if has_email else ""

    if has_email:

        if err == "taken" or _email_suggestion_candidates(page):

            return "email_taken"

        if err == "format":

            return "email_format"

        return "email"



    # 无邮箱框但还有 taken 建议 chip

    if _email_suggestion_candidates(page):

        return "email_taken"



    # 3) 全局错误页（仍是事实文案，不是计时）

    txt = _body_text(page)

    low = (txt or "").lower()

    if "account creation has been blocked" in low or "unusual activity" in low:

        return "blocked"

    if "we ran into a problem" in low or "please try again" in low:

        return "problem"



    return "unknown"





def _email_step_succeeded(page):

    """邮箱步真正成功：已出现密码/生日/姓名真实控件。"""

    return _detect_signup_step(page) in ("password", "birthday", "name")





def _wait_signup_step(page, want=None, leave=None, max_wait=8.0, poll=0.2):

    """轮询 DOM 事实直到命中 want / 离开 leave，或达到最长等待。



    max_wait 只是防死等上限，判断永远来自 _detect_signup_step。

    返回最终 step 字符串。

    """

    want_set = None

    if want is not None:

        if isinstance(want, (list, tuple, set)):

            want_set = set(want)

        else:

            want_set = {want}

    leave_set = None

    if leave is not None:

        if isinstance(leave, (list, tuple, set)):

            leave_set = set(leave)

        else:

            leave_set = {leave}



    deadline = time.time() + max(0.0, float(max_wait or 0))

    last = _detect_signup_step(page)

    while True:

        step = _detect_signup_step(page)

        last = step

        if want_set is not None and step in want_set:

            return step

        if leave_set is not None and step not in leave_set:

            return step

        if time.time() >= deadline:

            return last

        time.sleep(max(0.05, float(poll or 0.2)))





def _extract_email_from_text(text):



    raw = str(text or "")

    if not raw:

        return ""

    matches = re.findall(r"[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}", raw, flags=re.I)

    return matches[0].strip() if matches else ""





def _current_signup_identity_email(page):

    scripts = [

        """

const sels = ['#identityBadge', '[data-testid="identityBanner"]', '#bannerText', '#displayName'];

const vals = [];

for (const sel of sels) {

  const el = document.querySelector(sel);

  if (!el) continue;

  vals.push(el.getAttribute('aria-label') || '');

  vals.push(el.textContent || '');

  vals.push(el.innerText || '');

}

return vals.join('\\n');

        """,

        """

return document.documentElement ? document.documentElement.innerText : '';

        """,

    ]

    for script in scripts:

        try:

            raw = page.run_js_loaded(script) or ""

        except Exception:

            raw = ""

        email = _extract_email_from_text(raw)

        if email:

            return email

    return ""





def _resolve_submitted_email(page, fallback_email="", timeout=0.0):

    deadline = time.time() + max(0.0, float(timeout or 0.0))

    fallback = str(fallback_email or "").strip()

    while True:

        actual = _current_signup_identity_email(page)

        if actual:

            return actual

        if time.time() >= deadline:

            return fallback

        time.sleep(0.2)





def _email_input_present(page, timeout=0.3):

    return _ele(page, _email_input_selector(), timeout=timeout) is not None





def _email_suggestion_candidates(page, preferred_domain=None):

    """Collect Microsoft signup alternate-username chips.



    New Fluent UI no longer uses [data-testid="suggestions"]; taken usernames show as

    fui-InteractionTag chips (aria-label = bare prefix). Support both.

    """

    domain = (preferred_domain or "outlook.com").strip().lstrip("@").lower() or "outlook.com"

    values = []

    try:

        values = page.run_js_loaded(

            """

const out = [];

const push = (v) => {

  const t = String(v || '').trim();

  if (t) out.push(t);

};

const roots = [

  document.querySelector('[data-testid="suggestions"]'),

  document.querySelector('[class*="suggestions" i]'),

  document.body,

].filter(Boolean);

const seenNodes = new Set();

for (const root of roots) {

  const nodes = root.querySelectorAll(

    '[id^="fui-InteractionTag"], [class*="InteractionTag"], [data-testid="suggestions"] [aria-label], ' +

    '[data-testid="suggestions"] button, [role="listbox"] [role="option"], [role="listbox"] button'

  );

  for (const n of nodes) {

    if (seenNodes.has(n)) continue;

    seenNodes.add(n);

    push(n.getAttribute('aria-label'));

    push(n.innerText || n.textContent);

  }

}

return out;

            """

        ) or []

    except Exception:

        values = []

    candidates = []

    seen = set()

    for raw in values:

        text = str(raw or "").strip()

        if not text:

            continue

        # Skip full-sentence noise from broader DOM scans.

        if len(text) > 64 or " " in text:

            email = _extract_email_from_text(text)

            if not email:

                continue

        else:

            email = _extract_email_from_text(text)

            if not email:

                prefix = re.sub(r"\s+", "", text)

                if prefix and re.fullmatch(r"[A-Z0-9._%+\-]+", prefix, flags=re.I):

                    email = f"{prefix}@{domain}"

        if not email:

            continue

        low = email.lower()

        if low in seen:

            continue

        seen.add(low)

        candidates.append(email)

    return candidates





def _click_email_suggestion(page, preferred_domain=None, preferred_email=None):

    """Click a taken-page suggestion chip. Returns chosen email or ''."""

    domain = (preferred_domain or "outlook.com").strip().lstrip("@").lower() or "outlook.com"

    preferred = str(preferred_email or "").strip()

    preferred_prefix = preferred.split("@", 1)[0].strip() if preferred else ""

    try:

        chosen = page.run_js_loaded(

            f"""

const preferredPrefix = {json.dumps(preferred_prefix)};

const domain = {json.dumps(domain)};

const nodes = [...document.querySelectorAll(

  '[id^="fui-InteractionTag"], [class*="InteractionTag"], [data-testid="suggestions"] button, ' +

  '[data-testid="suggestions"] [aria-label], [role="listbox"] [role="option"], [role="listbox"] button'

)];

const parse = (n) => {{

  const raw = (n.getAttribute('aria-label') || n.innerText || n.textContent || '').trim();

  if (!raw || raw.length > 64 || /\\s/.test(raw)) {{

    const m = raw.match(/[A-Z0-9._%+\\-]+@[A-Z0-9.\\-]+\\.[A-Z]{{2,}}/i);

    return m ? m[0] : '';

  }}

  if (raw.includes('@')) return raw;

  if (/^[A-Z0-9._%+\\-]+$/i.test(raw)) return raw + '@' + domain;

  return '';

}};

let target = null;

let email = '';

if (preferredPrefix) {{

  for (const n of nodes) {{

    const e = parse(n);

    if (e && e.split('@')[0].toLowerCase() === preferredPrefix.toLowerCase()) {{

      target = n; email = e; break;

    }}

  }}

}}

if (!target) {{

  for (const n of nodes) {{

    const e = parse(n);

    if (e) {{ target = n; email = e; break; }}

  }}

}}

if (!target || !email) return '';

const clickable = target.closest('button') || target.querySelector('button') || target;

clickable.click();

return email;

            """

        ) or ""

    except Exception:

        chosen = ""

    return str(chosen or "").strip()





def _is_email_page(page):

    if _email_input_present(page, timeout=0.3):

        return True

    if _email_suggestion_candidates(page):

        return True

    txt = _body_text(page)

    low = txt.lower()

    if "enter your new email address" in low or "enter your email address" in low:

        return True

    if "new email" in low:

        return True

    if "that username is already taken" in low or "that email is already taken" in low:

        return True

    if "already taken" in low or "available options" in low:

        return True

    return False





def _wait_email_submit_outcome(page, timeout=5.0):

    """提交邮箱后等结果：完全由 _detect_signup_step 事实驱动。



    关键：过渡帧 step=unknown 绝不能当 stuck。

    leave 含 unknown，继续等到 password/taken/blocked 等硬事实。

    """

    step = _wait_signup_step(

        page,

        want=(

            "password",

            "birthday",

            "name",

            "email_taken",

            "email_format",

            "blocked",

            "problem",

        ),

        # email 仍在表单；unknown=路由动画，继续等

        leave=("email", "unknown"),

        max_wait=timeout,

        poll=0.12,

    )

    # 超时后最终再读一次；密码框优先（防 leave 漏检）

    if step in ("email", "unknown") or not step:

        if _password_input_present(page, timeout=0.35):

            step = "password"

        elif _birthday_control_present(page, timeout=0.2):

            step = "birthday"

        elif _name_input_present(page, timeout=0.2):

            step = "name"

        else:

            step = _detect_signup_step(page)

    if step == "email":

        step = _detect_signup_step(page)

    mapping = {

        "password": "password",

        "birthday": "advanced",

        "name": "advanced",

        "email_taken": "taken",

        "email_format": "format",

        "email": "stuck",

        "blocked": "blocked",

        "problem": "problem",

        # unknown 仍不应当 stuck：外层会再探测

        "unknown": "pending",

    }

    return mapping.get(step, "stuck")





# F 组 _all_contexts:re-export 自新包(单一源)。register 3 处调用仍按原名用。
_all_contexts = _ruyi_pkg._all_contexts

# --- resource blocking: re-exported from common.ruyi (single source) ---
_RUOYI_RESOURCE_BLOCK_KINDS = _ruyi_pkg._RUOYI_RESOURCE_BLOCK_KINDS
_RUOYI_RESOURCE_ALLOW_HOST_HINTS = _ruyi_pkg._RUOYI_RESOURCE_ALLOW_HOST_HINTS
_RUOYI_TELEMETRY_HOSTS = _ruyi_pkg._RUOYI_TELEMETRY_HOSTS
_RUOYI_RESOURCE_BLOCK_EXTS = _ruyi_pkg._RUOYI_RESOURCE_BLOCK_EXTS
_ruoyi_should_block_resource_request = _ruyi_pkg._ruoyi_should_block_resource_request
_start_ruoyi_resource_blocking = _ruyi_pkg._start_ruoyi_resource_blocking



def _stop_ruoyi_resource_blocking(page):

    if page is None:

        return False

    interceptor = getattr(page, "intercept", None)

    if interceptor is None or not callable(getattr(interceptor, "stop", None)):

        return False

    if not getattr(page, "_ruoyi_resource_blocking_active", False):

        return False

    try:

        interceptor.stop()

    except Exception:

        return False

    setattr(page, "_ruoyi_resource_blocking_active", False)

    return True



def _click_post_signup(page, tag):

    hit = _click_any(

        page,

        [

            'text:OK',

            'text:Accept',

            'text:Continue',

            'text:Next',

            'text:I agree',

            'text:Got it',

            'text:Agree',

            'text:同意',

            'text:继续',

            'text:下一步',

            'css:input[type="submit"]',

            'css:button[type="submit"]',

            '#idBtn_Accept',

            '#iNext',

            '#acceptButton',

        ],

        timeout=2,

    )

    if hit:

        log(f"  {tag} post-click: {hit}")

    return bool(hit)





def _handle_consent(page, tag, idx, deadline=None):

    for attempt in range(5):

        remaining = None if deadline is None else (deadline - time.time())

        if remaining is not None and remaining <= 0:

            return

        txt = _body_text(page)

        low = txt.lower()

        url = page.url.lower()

        if _on_signup_form(url) and not (

            any(kw in txt for kw in ["同意并继续", "个人数据", "数据导出"])

            or any(

                kw in low

                for kw in [

                    "agree and continue",

                    "consent",

                    "data export",

                    "accepter et continuer",

                    "consentement",

                ]

            )

            or "privacynotice" in url

        ):

            return

        log(f"  {tag} consent/privacy 第{attempt + 1}轮，点同意…")

        locator_timeout = 0.25

        if remaining is not None:

            locator_timeout = max(0.01, min(locator_timeout, remaining))

        _click_any(

            page,

            [

                'text:同意并继续',

                'text:同意',

                'text:Agree and continue',

                'text:Accept',

                'text:Continue',

                'text:OK',

                'text:Accepter et continuer',

                'text:Accepter',

                'text:Continuer',

                'text:Suivant',

                'css:input[type="submit"]',

                'css:button[type="submit"]',

                '#iNext',

                '#iAgree',

                '#acceptButton',

            ],

            timeout=locator_timeout,

        )

        sleep_for = 3.0

        if deadline is not None:

            sleep_for = min(sleep_for, max(0.0, deadline - time.time()))

        if sleep_for <= 0:

            return

        time.sleep(sleep_for)

        _shot(page, f"after_consent_{attempt}", idx)





def _ensure_signup_entry(page, tag, idx, timeout=SIGNUP_ENTRY_TIMEOUT):

    wanted_steps = (

        "email",

        "email_taken",

        "email_format",

        "password",

        "birthday",

        "name",

        "blocked",

        "problem",

    )

    deadline = time.time() + max(0.0, float(timeout or 0.0))

    _handle_consent(page, tag, idx, deadline=deadline)

    remaining = max(0.0, deadline - time.time())

    step = _wait_signup_step(page, want=wanted_steps, max_wait=remaining, poll=0.15)

    if step not in wanted_steps:

        step = _detect_signup_step(page)

    if step in wanted_steps:

        log(f"  {tag} signup entry DOM step={step}")

        return step

    current_url = str(getattr(page, "url", "") or "")

    log(f"  {tag} signup entry timeout {int(timeout or 0)}s url={current_url[:120]!r}", "WARN")

    _shot(page, "signup_entry_timeout", idx)

    return ""





def _email_error_kind(page):

    """返回 'taken' / 'format' / ''。"""

    txt = _body_text(page)

    lower = (txt or "").lower()

    # EN: "That username is already taken. Try another one or use one of these available options."

    taken_en = (

        "already taken" in lower

        or "username is already" in lower

        or "email is already" in lower

        or "email address is already" in lower

        or ("already" in lower and "taken" in lower)

        or ("already" in lower and "email" in lower and "account" in lower)

        or "available options" in lower

        or "try another one" in lower

        or "someone already has this" in lower

        or "not available" in lower

    )

    if taken_en:

        return "taken"

    # 中文占用提示

    if any(k in txt for k in ["已被使用", "不可用", "已被占用", "已经有人", "换一个", "已被注册", "已经存在"]):

        if any(k in lower for k in ["email", "address", "microsoft", "outlook", "帐户", "账户", "账号", "用户名"]):

            return "taken"

        if any(k in txt for k in ["电子邮件", "邮箱", "用户名"]):

            return "taken"

        # 短错误条只有“已被使用”时也按占用处理

        if any(k in txt for k in ["已被使用", "已被占用", "已被注册"]):

            return "taken"

    if any(k in lower for k in ["needs to start", "in the format", "enter a valid", "use letters", "invalid email"]):

        return "format"

    if any(k in txt for k in ["格式", "无效", "请输入有效"]):

        return "format"

    return ""





def _normalize_email_domain(domain):

    raw = str(domain or "").strip().lstrip("@").lower()

    return raw if raw in ALLOWED_EMAIL_SUFFIXES else "outlook.com"





def _normalize_email_suffixes(email_suffixes):

    if isinstance(email_suffixes, (list, tuple, set)):

        items = [str(x).strip() for x in email_suffixes if str(x).strip()]

    else:

        items = [s.strip() for s in re.split(r"[,;\s]+", str(email_suffixes or "")) if s.strip()]

    out = []

    seen = set()

    for item in items:

        normalized = _normalize_email_domain(item)

        if normalized in seen:

            continue

        seen.add(normalized)

        out.append(normalized)

    return ",".join(out or ["outlook.com"])





def _email_domain(email):

    if not email or "@" not in str(email):

        return "outlook.com"

    return _normalize_email_domain(str(email).rsplit("@", 1)[-1])





def _random_digits(count):

    return "".join(str(random.randint(0, 9)) for _ in range(max(1, int(count or 1))))





def _append_random_digits_email(email=None, prefix=None, preferred_domain=None, digits=1):

    domain = (preferred_domain or _email_domain(email)).strip().lstrip("@").lower() or "outlook.com"

    base_prefix = str(prefix or "").strip()

    if not base_prefix and email and "@" in str(email):

        base_prefix = str(email).split("@", 1)[0].strip()

    if not base_prefix:

        return _new_email_candidate("taken", domain)

    new_prefix = f"{base_prefix}{_random_digits(digits)}"

    return f"{new_prefix}@{domain}", new_prefix





def _new_email_candidate(

    kind="taken",

    preferred_domain=None,

    current_email=None,

    current_prefix=None,

    extra_digits=1,

):

    """Generate a replacement email while reusing helper format/suffix settings."""

    domain = (preferred_domain or "").strip().lstrip("@").lower()

    if kind == "taken" and (current_prefix or current_email):

        return _append_random_digits_email(current_email, current_prefix, domain, extra_digits)

    helpers = _load_helpers()

    try:

        email, _password, prefix = helpers.generate_email_password()

        if domain:

            prefix = (prefix or email.split("@", 1)[0]).strip()

            email = f"{prefix}@{domain}"

        else:

            domain = _email_domain(email)

            prefix = prefix or email.split("@", 1)[0]

        if kind == "taken":

            prefix = f"{prefix}{random.randint(10, 99)}"

            email = f"{prefix}@{domain}"

        return email, prefix

    except Exception:

        import string



        n = 11 if kind == "taken" else 9

        prefix = random.choice(string.ascii_lowercase) + "".join(

            random.choices(string.ascii_lowercase + string.digits, k=n)

        )

        domain = domain or "outlook.com"

        return f"{prefix}@{domain}", prefix





def _select_domain(dd, domain, tag):

    """选择 LiveDomain 下拉：优先 outlook.com / hotmail.com。"""

    domain = (domain or "outlook.com").strip().lstrip("@").lower()

    candidates = [domain]

    if domain == "hotmail.com":

        candidates += ["hotmail.com", "Hotmail.com"]

    else:

        candidates += ["outlook.com", "Outlook.com"]

    _scroll_into_view(dd)

    for cand in candidates:

        try:

            dd.select.by_value(cand)

            return cand

        except Exception:

            pass

        try:

            dd.select.by_text(cand)

            return cand

        except Exception:

            pass

    try:

        if _safe_input(dd, domain, clear=True):

            return domain

    except Exception:

        pass

    log(f"  {tag} 域名下拉选择失败，目标={domain}", "WARN")

    return None





def _still_on_email_or_taken(page):

    """True when DOM 事实仍停在邮箱步（含占用/格式错误）。"""

    return _detect_signup_step(page) in ("email", "email_taken", "email_format")





def _rotate_taken_email(cur_email, cur_prefix, cur_domain, taken_retry_count, tag, reason="taken"):

    """Pick next email after a taken/stuck: suggestion already handled by caller."""

    # 1st: append 3 digits; 2nd+: append more / full regenerate after 3 tries

    if taken_retry_count >= 3:

        cur_email, cur_prefix = _new_email_candidate("format", cur_domain)  # full new prefix

        cur_domain = _email_domain(cur_email)

        log(f"  {tag} email {reason}, full regenerate: {cur_email}", "WARN")

        return cur_email, cur_prefix, cur_domain

    extra_digits = 3 if taken_retry_count <= 1 else 2

    cur_email, cur_prefix = _new_email_candidate(

        "taken", cur_domain, cur_email, cur_prefix, extra_digits

    )

    cur_domain = _email_domain(cur_email)

    log(

        f"  {tag} email {reason}, append {extra_digits} digit(s), retry: {cur_email}",

        "WARN",

    )

    return cur_email, cur_prefix, cur_domain





def _fill_email(page, email, prefix, tag, idx):

    """Fill email step. Only returns after password input (or later) is visible.



    Hard rule: if email input is present, ALWAYS fill+submit. Never early-return

    on name/birthday text heuristics (they false-positive on "username already taken").

    """

    cur_email, cur_prefix = email, prefix

    cur_domain = _email_domain(cur_email)

    taken_retry_count = 0

    stuck_rotate_count = 0  # 真 stuck(邮箱页提交无响应,代理抖动/出口IP风控)换号计数,超限早退

    for attempt in range(12):

        try:

            try:

                page.run_js_loaded("window.scrollTo(0, 0); return true;")

            except Exception:

                pass

            # 仅当密码框（或后续真实控件）已出现，才允许跳过填写

            if _email_step_succeeded(page) and not _email_input_present(page, timeout=0.2):

                log(f"  {tag} email step already advanced (pre-check)")

                return _resolve_submitted_email(page, cur_email, timeout=0.3)



            sel = _email_input_selector()

            email_el = _ele(page, sel, timeout=2)

            if email_el is None:

                if _email_step_succeeded(page):

                    log(f"  {tag} email step already advanced (no input)")

                    return _resolve_submitted_email(page, cur_email, timeout=0.3)

                if _email_error_kind(page) == "taken" or _email_suggestion_candidates(page, cur_domain):

                    taken_retry_count += 1

                    cur_email, cur_prefix, cur_domain = _rotate_taken_email(

                        cur_email, cur_prefix, cur_domain, taken_retry_count, tag, "taken(no-input)"

                    )

                    time.sleep(0.25)

                    continue

                log(f"  {tag} email input not found", "ERR")

                _shot(page, "no_email", idx)

                return None



            # ---- 强制填写：只要邮箱框在，就必须 input ----

            domain_kind, dd = _domain_control(page)

            # Fluent 域名下拉（button#domainDropdownId）= 只填本地前缀；

            # 经典 select#LiveDomainBoxList 也只填前缀；否则写完整邮箱。

            value = cur_prefix if domain_kind in ("select", "fluent") else cur_email

            filled_ok = _safe_input(email_el, value, clear=True)

            if not filled_ok:

                log(f"  {tag} email input failed, retrying", "WARN")

                _shot(page, "email_input_fail", idx)

                email_el = _ele(page, sel, timeout=2)

                filled_ok = email_el is not None and _safe_input(email_el, value, clear=True)

            if not filled_ok:

                taken_retry_count += 1

                cur_email, cur_prefix, cur_domain = _rotate_taken_email(

                    cur_email, cur_prefix, cur_domain, taken_retry_count, tag, "input-fail"

                )

                continue



            # 读回 value，确认真的写进去了（用原生 .value / attr，不瞎用裸 return run_js）

            actual_val = _read_input_value(email_el)

            if not actual_val:

                # 再硬写一次：合法 functionDeclaration

                try:

                    email_el.run_js(

                        """

function(v) {

  const text = String(v || '');

  const el = this;

  el.focus();

  const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set;

  if (setter) setter.call(el, text); else el.value = text;

  el.dispatchEvent(new Event('input', {bubbles:true}));

  el.dispatchEvent(new Event('change', {bubbles:true}));

  return el.value;

}

                        """,

                        value,

                    )

                except Exception:

                    pass

                actual_val = _read_input_value(email_el)

            if not actual_val:

                # 元素句柄可能 stale，从页面直接探测

                actual_val = _probe_email_value_on_page(page)

            if not actual_val:

                log(f"  {tag} email value still empty after fill, retry", "WARN")

                _shot(page, "email_empty_value", idx)

                time.sleep(0.25)

                continue



            if domain_kind == "select" and dd is not None:

                chosen = _select_domain(dd, cur_domain, tag) or cur_domain

                cur_domain = chosen

                cur_email = f"{cur_prefix}@{cur_domain}"

                log(f"  {tag} filled prefix: {cur_prefix}@{cur_domain} (value={actual_val!r})")

            elif domain_kind == "fluent":

                # Fluent 下拉默认 @outlook.com；读按钮文案校正

                try:

                    domain_txt = (

                        str(getattr(dd, "text", "") or "")

                        or str(dd.attr("value") or "")

                        or ""

                    ).strip().lstrip("@").lower()

                    if domain_txt and "." in domain_txt:

                        cur_domain = domain_txt.split()[0]

                except Exception:

                    pass

                cur_email = f"{cur_prefix}@{cur_domain}"

                log(f"  {tag} filled prefix(fluent): {cur_email} (value={actual_val!r})")

            else:

                # 全邮箱写入时，以输入值为准

                if "@" in actual_val:

                    cur_email = actual_val

                    cur_prefix = cur_email.split("@", 1)[0]

                    cur_domain = _email_domain(cur_email)

                log(f"  {tag} filled email: {cur_email} (value={actual_val!r})")



            # 填完立刻提交：不再固定睡 1s / 随机 SUBMIT_DELAY

            _click_next(page, tag, wait_before=False, wait_after=False, timeout=1.0)



            outcome = _wait_email_submit_outcome(page, timeout=5.0)

            # 最终硬事实：密码/生日/姓名框在 → 成功，覆盖任何 stuck/pending 误判

            if _password_input_present(page, timeout=0.25):

                outcome = "password"

            elif _birthday_control_present(page, timeout=0.2) or _name_input_present(page, timeout=0.2):

                outcome = "advanced"

            # 成功硬闸：必须真正进入密码/后续页

            if outcome in ("password", "advanced") and not _email_step_succeeded(page):

                # 再给一次短等（过渡帧）

                step2 = _wait_signup_step(

                    page,

                    want=("password", "birthday", "name"),

                    max_wait=1.5,

                    poll=0.1,

                )

                if step2 in ("password", "birthday", "name"):

                    outcome = "password" if step2 == "password" else "advanced"

                elif _password_input_present(page, timeout=0.3):

                    outcome = "password"

                else:

                    log(f"  {tag} outcome={outcome} but email step not really done, override", "WARN")

                    outcome = _email_error_kind(page) or (

                        "taken" if _email_suggestion_candidates(page, cur_domain) else "stuck"

                    )

            # 邮箱框还在且不是已成功后续步，才降为 stuck/taken

            if (

                _email_input_present(page, timeout=0.25)

                and outcome in ("password", "advanced", "", "pending")

                and not _password_input_present(page, timeout=0.15)

            ):

                kind = _email_error_kind(page)

                outcome = kind or ("taken" if _email_suggestion_candidates(page, cur_domain) else "stuck")

            log(f"  {tag} email submit outcome={outcome or 'empty'} attempt={attempt + 1}")



            if outcome in ("password", "advanced") and (

                _email_step_succeeded(page) or _password_input_present(page, timeout=0.2)

            ):

                _shot(page, "after_email", idx)

                return _resolve_submitted_email(page, cur_email, timeout=0.3)



            if outcome == "taken" or (

                not outcome and (_email_error_kind(page) == "taken" or _email_suggestion_candidates(page, cur_domain))

            ):

                taken_retry_count += 1

                suggestions = _email_suggestion_candidates(page, cur_domain)

                if suggestions and taken_retry_count <= 4:

                    pick = suggestions[0]

                    clicked = _click_email_suggestion(page, cur_domain, pick)

                    if clicked:

                        cur_email = clicked

                        cur_prefix = cur_email.split("@", 1)[0]

                        cur_domain = _email_domain(cur_email)

                        log(f"  {tag} email taken, clicked suggestion: {cur_email}", "WARN")

                        time.sleep(0.25)

                        _click_next(page, tag, wait_before=False, wait_after=False, timeout=1.0)

                        outcome2 = _wait_email_submit_outcome(page, timeout=5.0)

                        if outcome2 in ("password", "advanced") and _email_step_succeeded(page):

                            log(f"  {tag} suggestion submit outcome={outcome2}")

                            _shot(page, "after_email", idx)

                            return _resolve_submitted_email(page, cur_email, timeout=0.3)

                        log(f"  {tag} suggestion submit outcome={outcome2 or 'empty'}", "WARN")

                    else:

                        sug_prefix = pick.split("@", 1)[0]

                        cur_email = f"{sug_prefix}@{cur_domain}"

                        cur_prefix = sug_prefix

                        log(f"  {tag} email taken, reuse suggestion prefix: {cur_email}", "WARN")

                        time.sleep(0.2)

                        continue

                cur_email, cur_prefix, cur_domain = _rotate_taken_email(

                    cur_email, cur_prefix, cur_domain, taken_retry_count, tag, "taken"

                )

                time.sleep(0.25)

                continue



            if outcome == "format":

                cur_email, cur_prefix = _new_email_candidate("format", cur_domain)

                cur_domain = _email_domain(cur_email)

                log(f"  {tag} format error, retry: {cur_email}", "WARN")

                time.sleep(0.25)

                continue



            # stuck 前最后一次硬探测：密码页已出就直接成功，禁止误换号

            if _password_input_present(page, timeout=0.4) or _email_step_succeeded(page):

                log(f"  {tag} stuck overridden: already advanced (password/next step)", "WARN")

                _shot(page, "after_email", idx)

                return _resolve_submitted_email(page, cur_email, timeout=0.3)

            if outcome == "pending":

                # 过渡中，再等一小会儿而不是立刻换号

                step3 = _wait_signup_step(

                    page,

                    want=("password", "birthday", "name", "email_taken", "email_format", "blocked"),

                    leave=("email", "unknown"),

                    max_wait=2.0,

                    poll=0.12,

                )

                if step3 in ("password", "birthday", "name") or _password_input_present(page, timeout=0.3):

                    log(f"  {tag} pending resolved → {step3 or 'password'}")

                    return _resolve_submitted_email(page, cur_email, timeout=0.3)

                if step3 == "email_taken" or _email_error_kind(page) == "taken":

                    taken_retry_count += 1

                    cur_email, cur_prefix, cur_domain = _rotate_taken_email(

                        cur_email, cur_prefix, cur_domain, taken_retry_count, tag, "taken"

                    )

                    time.sleep(0.25)

                    continue

                if step3 == "email_format":

                    cur_email, cur_prefix = _new_email_candidate("format", cur_domain)

                    cur_domain = _email_domain(cur_email)

                    log(f"  {tag} format error, retry: {cur_email}", "WARN")

                    time.sleep(0.25)

                    continue



            # 真 stuck：邮箱框仍在 / 无后续控件

            log(f"  {tag} still on email page after submit (outcome={outcome or 'empty'}), retry", "WARN")

            _shot(page, "email_stuck", idx)

            stuck_rotate_count += 1

            # 邮箱页连续 stuck 换号超限:多半代理抖动/出口IP风控导致提交无响应,
            # 再换号也是一样卡,早退 fail 让上层换代理,别 5 轮 ×8s 累加到 60s。
            if stuck_rotate_count > EMAIL_STUCK_MAX_ROTATE:
                log(f"  {tag} email page stuck {stuck_rotate_count}x rotate, give up (proxy jitter)", "WARN")
                _shot(page, "email_stuck_giveup", idx)
                return None

            taken_retry_count += 1

            cur_email, cur_prefix, cur_domain = _rotate_taken_email(

                cur_email, cur_prefix, cur_domain, taken_retry_count, tag, "stuck"

            )

            time.sleep(0.3)

            continue

        except Exception as exc:

            log(f"  {tag} email step exception ({type(exc).__name__}: {exc}), retrying", "WARN")

            _shot(page, "email_exc", idx)

            taken_retry_count += 1

            cur_email, cur_prefix, cur_domain = _rotate_taken_email(

                cur_email, cur_prefix, cur_domain, taken_retry_count, tag, "exc"

            )

            time.sleep(0.4)

            continue

    log(f"  {tag} email step failed after retries", "ERR")

    _shot(page, "email_fail", idx)

    return None





def _fill_password(page, password, tag, idx):

    """Fill password. Returns True / False / 'email_taken'.



    步骤判定只认 _detect_signup_step 事实；等待只是轮询直到事实变化。

    """

    step = _detect_signup_step(page)

    if step in ("email", "email_taken", "email_format"):

        log(f"  {tag} password step but DOM step={step} → retry email", "WARN")

        _shot(page, "email_still_on_pwd", idx)

        return "email_taken"



    # 还没见到密码框：轮询事实，直到 password / 回邮箱 / 超时

    if step != "password":

        step = _wait_signup_step(

            page,

            want=("password", "email", "email_taken", "email_format", "birthday", "name", "blocked", "problem"),

            max_wait=5.0,

            poll=0.15,

        )

        log(f"  {tag} password wait DOM step={step}")

        if step in ("email", "email_taken", "email_format"):

            log(f"  {tag} DOM back to {step} when expecting password → retry email", "WARN")

            _shot(page, "email_taken_as_pwd", idx)

            return "email_taken"

        if step != "password":

            # 已经到生日/姓名 = 密码步被跳过，也算过

            if step in ("birthday", "name"):

                log(f"  {tag} password skipped, already at {step}")

                return True

            log(f"  {tag} 密码框未找到 (DOM step={step})", "ERR")

            _shot(page, "pwd_fail", idx)

            return False



    pwd_el = _ele(page, _password_input_selector(), timeout=1)

    if pwd_el is None:

        step = _detect_signup_step(page)

        if step in ("email", "email_taken", "email_format"):

            log(f"  {tag} 密码框未找到（DOM step={step}）→ retry email", "WARN")

            _shot(page, "email_taken_as_pwd", idx)

            return "email_taken"

        log(f"  {tag} 密码框未找到 (DOM step={step})", "ERR")

        _shot(page, "pwd_fail", idx)

        return False

    if not _safe_input(pwd_el, password, clear=True):

        log(f"  {tag} 密码输入失败", "ERR")

        _shot(page, "pwd_fail", idx)

        return False

    log(f"  {tag} 密码已填")

    _click_next(page, tag, wait_before=False, wait_after=False, timeout=1.0)

    # 提交后看事实，不靠固定 sleep 当成功

    after = _wait_signup_step(

        page,

        leave=("password",),

        max_wait=4.0,

        poll=0.15,

    )

    log(f"  {tag} after password DOM step={after}")

    _shot(page, "after_pwd", idx)

    return True





def _fill_birthday(page, year, month, day, tag, idx):

    """填写生日页：Month/Day 用 combobox，Year 用独立 input。



    历史问题：

      1) 年份回退误点 combos[-1]（实际是 Day），导致 Day 被填成年份，或 Day 被选两遍

      2) text:{day} 模糊匹配，day=1 会命中 10/11/12 等

    """

    month_words = [

        "",

        "January",

        "February",

        "March",

        "April",

        "May",

        "June",

        "July",

        "August",

        "September",

        "October",

        "November",

        "December",

    ]

    month_words_cn = [

        "",

        "1月",

        "2月",

        "3月",

        "4月",

        "5月",

        "6月",

        "7月",

        "8月",

        "9月",

        "10月",

        "11月",

        "12月",

    ]

    day_str = str(day)

    year_str = str(year)

    month_candidates = [

        month_words[month],

        month_words_cn[month],

        str(month),

        f"{month:02d}",

    ]



    def _select_options_text(ele):

        try:

            opts = ele.eles("css:option", timeout=1) or []

            parts = []

            for opt in opts[:40]:

                txt = ((getattr(opt, "text", "") or "") + " " + (opt.attr("value") or "")).strip()

                if txt:

                    parts.append(txt.lower())

            return " ".join(parts)

        except Exception:

            return ""



    def _select_label(ele):

        chunks = []

        for attr in ("aria-label", "id", "name"):

            try:

                chunks.append(ele.attr(attr) or "")

            except Exception:

                pass

        try:

            chunks.append(ele.text or "")

        except Exception:

            pass

        return " ".join(chunks).lower()



    def _pick_select(ele, value, texts=()):

        candidates = [str(value), *[str(t) for t in texts if t]]

        for candidate in candidates:

            try:

                ele.select.by_value(candidate)

                return True

            except Exception:

                pass

            try:

                ele.select.by_text(candidate)

                return True

            except Exception:

                pass

        try:

            ele.input(str(value))

            return True

        except Exception:

            return False



    def _combo_meta(ele):

        try:

            aria = (ele.attr("aria-label") or "").strip()

        except Exception:

            aria = ""

        try:

            eid = (ele.attr("id") or "").strip()

        except Exception:

            eid = ""

        try:

            text = (ele.text or "").strip()

        except Exception:

            text = ""

        label = f"{aria} {eid} {text}".lower()

        return {"ele": ele, "aria": aria, "id": eid, "text": text, "label": label}



    def _classify_combo(meta):

        label = meta["label"]

        text = (meta["text"] or "").strip()

        text_low = text.lower()

        eid = meta["id"].lower()



        is_country = any(k in label for k in ["country", "region", "国家", "地区", "pays", "país"])

        is_month = (

            any(k in label for k in ["month", "月", "mois", "mes", "monat", "mês"])

            or any(k in eid for k in ["month", "birthmonth"])

            or text_low in ("month", "月", "月份", "mois", "mes")

        )

        is_day = (

            any(k in label for k in ["day", "日", "jour", "día", "tag", "dia"])

            or any(k in eid for k in ["day", "birthday"])

            or text_low in ("day", "日", "jour", "día")

        )

        is_year = any(k in label for k in ["year", "年", "année", "año", "ano"])



        # birthmonth / birthday 同时含 day/month 子串时按更具体的字段判定

        if is_month and is_day:

            if "month" in eid or "month" in label:

                is_day = False

            elif "day" in eid or "birthday" in eid or "day" in label:

                is_month = False



        # 已选中值：1月 / January / 15 / 15日

        if not is_month and not is_day and not is_country and not is_year:

            if text.endswith("月") and len(text) <= 4:

                is_month = True

            elif text.endswith("日") and len(text) <= 4 and text[:-1].isdigit():

                is_day = True

            elif text.isdigit() and 1 <= int(text) <= 31:

                is_day = True

            elif text_low in {w.lower() for w in month_words[1:]}:

                is_month = True



        if is_country:

            return "country"

        if is_year:

            return "year"

        if is_month:

            return "month"

        if is_day:

            return "day"

        return "unknown"



    def _press_escape():

        try:

            page.actions.press("\ue00c").perform()  # Escape

            time.sleep(0.12)

        except Exception:

            pass



    def _click_option_exact(candidates):

        """精确匹配 option 文本，避免 text:1 命中 10/11/12。"""

        wants = [str(c).strip() for c in candidates if c is not None and str(c).strip()]

        if not wants:

            return False

        try:

            ok = page.run_js_loaded(

                """

const wants = arguments[0].map(String);

const opts = [...document.querySelectorAll('[role="option"], option')];

for (const want of wants) {

  const hit = opts.find(o => {

    const t = (o.textContent || '').trim();

    const v = (o.getAttribute('value') || '').trim();

    return t === want || v === want;

  });

  if (hit) {

    hit.scrollIntoView({block: 'nearest'});

    hit.click();

    return true;

  }

}

return false;

                """,

                wants,

            )

            if ok:

                return True

        except Exception:

            pass

        # 兜底：只点可见 option，且文本完全相等

        for want in wants:

            for opt in _eles(page, 'css:[role="option"]', timeout=1):

                try:

                    txt = (opt.text or "").strip()

                except Exception:

                    continue

                if txt == want:

                    try:

                        opt.click_self()

                        return True

                    except Exception:

                        try:

                            opt.click_self(by_js=True)

                            return True

                        except Exception:

                            pass

        return False



    def _open_and_pick(combo, candidates, typed_fallback=None):

        try:

            combo.click_self()

        except Exception:

            try:

                combo.click_self(by_js=True)

            except Exception:

                return False

        time.sleep(0.15)

        if _click_option_exact(candidates):

            time.sleep(0.1)

            _press_escape()

            return True

        if typed_fallback is not None:

            try:

                page.actions.type(str(typed_fallback)).press("\ue007").perform()

                time.sleep(0.1)

                _press_escape()

                return True

            except Exception:

                pass

        _press_escape()

        return False



    def _combo_shows_value(meta, expected_tokens):

        try:

            text = (meta["ele"].text or "").strip()

        except Exception:

            text = meta.get("text") or ""

        low = text.lower()

        for tok in expected_tokens:

            t = str(tok).strip().lower()

            if t and t in low:

                return True

        # 已不是占位文案则视为已选

        placeholders = {"month", "day", "月", "日", "jour", "mois", "día", ""}

        return bool(text) and low not in placeholders



    def _fill_year_input():

        yr = None

        for sel in [

            'css:#BirthYearInput, input[name="BirthYear"], input[name="BirthYearInput"]',

            'css:input[aria-label*="Birth year" i], input[aria-label*="year" i], input[aria-label*="年"]',

            'css:input[id*="BirthYear" i], input[id*="Year" i]',

            'css:input[placeholder*="Year" i], input[placeholder*="year"], input[placeholder*="年"]',

            'css:input[type="number"], input[type="text"][inputmode="numeric"]',

        ]:

            yr = _ele(page, sel, timeout=0.35)

            if yr is not None:

                break

        if yr is not None:

            try:

                yr.input(year_str, clear=True)

                log(f"  {tag} 年份={year_str}")

                return True

            except Exception:

                pass

        try:

            ok = page.run_js_loaded(

                """

const y = String(arguments[0]);

const inputs = [...document.querySelectorAll('input')];

const target = inputs.find(el =>

  /year|年|année/i.test(el.placeholder || '') ||

  /year|birthyear/i.test(el.id || '') ||

  /year|birthyear/i.test(el.name || '') ||

  /year|年/i.test(el.getAttribute('aria-label') || '')

);

if (!target) return false;

target.focus();

const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set;

if (setter) setter.call(target, y); else target.value = y;

target.dispatchEvent(new Event('input', {bubbles: true}));

target.dispatchEvent(new Event('change', {bubbles: true}));

return true;

                """,

                year_str,

            )

            if ok:

                log(f"  {tag} 年份(JS)={year_str}")

                return True

        except Exception:

            pass

        log(f"  {tag} 年份输入未找到", "WARN")

        return False



    def _apply_birthday_form():

        selects = _eles(page, "css:select", timeout=2)

        select_count = len(selects)

        log(f"  {tag} 生日页 select 数={select_count}")

        if select_count >= 3:

            metas = []

            for pos, ele in enumerate(selects):

                label = _select_label(ele)

                options_text = _select_options_text(ele)

                metas.append({"pos": pos, "ele": ele, "label": label, "options": options_text})



            def _find_idx(pred):

                for meta in metas:

                    if pred(meta):

                        return meta["pos"]

                return None



            country_idx = _find_idx(

                lambda m: any(k in m["label"] for k in ["country", "region", "国家", "地区", "pays"])

                or any(k in m["options"] for k in ["united states", " usa ", "美国"])

            )

            month_idx = _find_idx(

                lambda m: any(k in m["label"] for k in ["month", "月", "mois"])

                or any(word.lower() in m["options"] for word in month_words[1:])

            )

            day_idx = _find_idx(

                lambda m: (

                    any(k in m["label"] for k in ["day", "日", "jour"])

                    and "month" not in m["label"]

                )

                or (

                    # option 是 1..31 数字列表

                    all(str(n) in m["options"].split() for n in (1, 15, 28))

                    and not any(word.lower() in m["options"] for word in month_words[1:])

                )

            )

            year_idx = _find_idx(

                lambda m: any(k in m["label"] for k in ["year", "年", "année"])

                or str(year) in m["options"]

            )



            used = {x for x in (country_idx, month_idx, day_idx, year_idx) if x is not None}

            remain = [meta["pos"] for meta in metas if meta["pos"] not in used]



            # 不切换国家/地区：保留页面按 IP 预填的默认值；识别到 country 只占位，不改

            if month_idx is None and remain:

                month_idx = remain.pop(0)

            if day_idx is None and remain:

                day_idx = remain.pop(0)

            if year_idx is None and remain:

                year_idx = remain.pop(-1)



            if month_idx is not None:

                _pick_select(selects[month_idx], month, (month_words[month], month_words_cn[month]))

            if day_idx is not None:

                _pick_select(selects[day_idx], day, (f"{day}日", day_str, f"{day:02d}"))

            if year_idx is not None:

                _pick_select(selects[year_idx], year)

            else:

                _fill_year_input()

            log(

                f"  {tag} 生日(select): country=keep(idx={country_idx}) "

                f"month={month_idx} day={day_idx} year={year_idx}"

            )

            return



        if select_count >= 2:

            _pick_select(selects[0], month, (month_words[month], month_words_cn[month]))

            _pick_select(selects[1], day, (f"{day}日", day_str))

            _fill_year_input()

            return



        log(f"  {tag} 无 select，用 combobox…")

        raw_combos = _eles(page, 'css:button[role="combobox"], [role="combobox"]', timeout=2)

        metas = [_combo_meta(c) for c in raw_combos]

        roles = []

        for i, meta in enumerate(metas):

            role = _classify_combo(meta)

            roles.append(role)

            log(f"  {tag} combo[{i}]: role={role} text={meta['text']!r} id={meta['id']!r}")



        # 位置兜底：常见顺序 Country, Month, Day（Year 是 input，不在 combobox 里）

        unknown_idxs = [i for i, r in enumerate(roles) if r == "unknown"]

        if "month" not in roles and unknown_idxs:

            roles[unknown_idxs.pop(0)] = "month"

        if "day" not in roles and unknown_idxs:

            roles[unknown_idxs.pop(0)] = "day"



        month_idx = next((i for i, r in enumerate(roles) if r == "month"), None)

        day_idx = next((i for i, r in enumerate(roles) if r == "day"), None)



        month_filled = False

        day_filled = False



        if month_idx is not None:

            month_filled = _open_and_pick(

                metas[month_idx]["ele"],

                month_candidates,

                typed_fallback=str(month),

            )

            if month_filled:

                # 刷新 text 再校验

                metas[month_idx] = _combo_meta(metas[month_idx]["ele"])

                month_filled = _combo_shows_value(metas[month_idx], month_candidates)

            log(f"  {tag} month={'ok' if month_filled else 'FAIL'} idx={month_idx} val={month}")



        if day_idx is not None:

            day_candidates = [day_str, f"{day}日", f"{day:02d}"]

            day_filled = _open_and_pick(

                metas[day_idx]["ele"],

                day_candidates,

                typed_fallback=day_str,

            )

            if day_filled:

                metas[day_idx] = _combo_meta(metas[day_idx]["ele"])

                day_filled = _combo_shows_value(metas[day_idx], day_candidates)

            log(f"  {tag} day={'ok' if day_filled else 'FAIL'} idx={day_idx} val={day}")

        else:

            log(f"  {tag} 未识别到 Day combobox, roles={roles}", "WARN")



        # 年份只走独立 input，绝不点 combos[-1]（那是 Day）

        year_ok = _fill_year_input()

        if not month_filled or not day_filled or not year_ok:

            log(

                f"  {tag} 生日字段状态: month={month_filled} day={day_filled} year={year_ok}",

                "WARN",

            )



    def _wait_after_birthday_submit(max_wait=BIRTHDAY_SUBMIT_TIMEOUT):

        # 事实轮询：离开 birthday 控件，或进入 name。填完立刻提交，这里只等页面切换。

        return _wait_after_birthday_submit_step(page, max_wait=max_wait)



    def _submit_birthday_form():

        # 填完直接点 Next，不再随机等 0~SUBMIT_DELAY

        _click_next(page, tag, wait_before=False, wait_after=False)

        step = _wait_after_birthday_submit(BIRTHDAY_SUBMIT_TIMEOUT)

        if step == "birthday" or _is_birthday_page(page):

            log(f"  {tag} still on birthday after submit (DOM step={step}), refill year and retry", "WARN")

            _fill_year_input()

            _click_next(page, tag, wait_before=False, wait_after=False)

            _wait_after_birthday_submit(BIRTHDAY_SUBMIT_TIMEOUT)



    # 等生日控件出现（事实），不是 sleep 再猜文案

    step = _resolve_birthday_entry_step(page, max_wait=BIRTHDAY_ENTRY_TIMEOUT)

    log(f"  {tag} birthday enter DOM step={step}")

    if step == "name":

        log(f"  {tag} already on name page, skip birthday")

        return True

    if step != "birthday" and not _is_birthday_page(page):

        log(f"  {tag} birthday controls not present (DOM step={step})", "WARN")

        # 没控件就别瞎填

        if step not in ("birthday",):

            _shot(page, "bday_fail", idx)

            return False

    _shot(page, "bday_page", idx)



    for attempt in range(2):

        if attempt > 0:

            log(f"  {tag} 生日提交后仍停在 birthday，重选月/日/年再提交", "WARN")

        _apply_birthday_form()

        _submit_birthday_form()

        step = _detect_signup_step(page)

        if step != "birthday" and not _is_birthday_page(page):

            log(f"  {tag} birthday left → DOM step={step}")

            _shot(page, "after_bday", idx)

            return True



    _shot(page, "after_bday", idx)

    step = _detect_signup_step(page)

    if step == "birthday" or _is_birthday_page(page):

        log(f"  {tag} 生日页仍未通过 (DOM step={step})", "WARN")

        _shot(page, "bday_fail", idx)

        return False

    return True





def _fill_name_and_terms(page, first, last, prefix, tag, idx):

    # 等姓名输入框出现（事实），不用固定 sleep 猜页面

    step = _wait_signup_step(page, want=("name",), max_wait=6.0, poll=0.15)

    if step != "name" and not _is_name_page(page):

        log(f"  {tag} name page not detected (DOM step={step}), skip name fill", "WARN")

        _shot(page, "name_fail", idx)

        return False

    log(f"  {tag} name enter DOM step={step}")



    # Fluent: firstNameInput / lastNameInput；经典: FirstName / LastName

    first_sel = (

        'css:#firstNameInput, input[name="firstNameInput"], input[name="FirstName"], #FirstName, '

        'input[id*="firstName" i], input[aria-label*="First name" i], input[aria-label*="first name" i], '

        'input[aria-label*="名" i]'

    )

    last_sel = (

        'css:#lastNameInput, input[name="lastNameInput"], input[name="LastName"], #LastName, '

        'input[id*="lastName" i], input[aria-label*="Last name" i], input[aria-label*="last name" i], '

        'input[aria-label*="姓" i]'

    )



    fe = le = None

    text_inputs = []

    field_deadline = time.time() + 3

    while time.time() < field_deadline:

        fe = _ele(page, first_sel, timeout=0.4)

        le = _ele(page, last_sel, timeout=0.4)

        if fe is not None and le is not None:

            break

        text_inputs = _eles(page, 'css:input[type="text"]', timeout=0.3)

        if len(text_inputs) >= 2:

            break

        time.sleep(0.15)



    filled = False

    if fe is not None:

        ok_f = _safe_input(fe, first, clear=True)

        ok_l = True

        if le is not None:

            ok_l = _safe_input(le, last, clear=True)

        if ok_f and ok_l:

            log(f"  {tag} name: {first} {last}")

            filled = True

        else:

            log(f"  {tag} name input partial fail first={ok_f} last={ok_l}", "WARN")

    if not filled and len(text_inputs) >= 2:

        # Fluent 顺序通常 First, Last

        _safe_input(text_inputs[0], first, clear=True)

        _safe_input(text_inputs[1], last, clear=True)

        log(f"  {tag} name(generic): {first} {last}")

        filled = True

    if not filled:

        uname_sel = (

            'css:input[id*="displayName" i], input[id*="gamertag" i], '

            'input[name*="displayName" i], input[aria-label*="gamertag" i]'

        )

        ue = _ele(page, uname_sel, timeout=0.8)

        txt = _body_text(page)

        low = (txt or "").lower()

        if ue is not None and any(k in low for k in ["gamertag", "display name", "pseudo", "surnom"]):

            username = prefix[:8] + str(random.randint(100, 999))

            _safe_input(ue, username, clear=True)

            log(f"  {tag} username: {username}")

            _click_next(page, tag, wait_before=False, wait_after=False)

            return True

        log(f"  {tag} name inputs not found", "WARN")

        _shot(page, "name_fail", idx)

        return False



    for cb in _eles(

        page,

        'css:input[type="checkbox"][required], [role="checkbox"][aria-required="true"]',

        timeout=0.3,

    ):

        try:

            # is_checked 是方法不是属性:getattr 拿到的是绑定方法, bool() 恒 True
            # 会导致「已勾选」误判而从不点击。必须实际调用 is_checked()
            _is_checked_fn = getattr(cb, "is_checked", None)
            checked = bool(_is_checked_fn()) if callable(_is_checked_fn) else False

        except Exception:

            checked = False

        if not checked:

            _safe_click(cb)

            log(f"  {tag} checked required checkbox")

    # Add your name 页营销订阅 checkbox(marketingOptIn):用户要求提交时不勾选。
    # Fluent UI Checkbox 结构: <span class="fui-Checkbox"><input id=marketingOptIn
    #   data-testid=marketingOptIn><div class="fui-Checkbox__indicator"/><label/></span>
    # 真实勾选态以 input.checked 为准(Fluent 用原生 input 作受控源),span aria-checked 兜底。
    # 取消优先点 label(Fluent 监听 label click 切换内部状态,最可靠),回读确认,JS 兜底。
    mkt = _ele(

        page,

        'css:[data-testid="marketingOptIn"], #marketingOptIn, input[name="marketingOptIn"], '

        'input[id*="marketingOptIn" i], [role="checkbox"][aria-label*="marketing" i]',

        timeout=0.4,

    )

    if mkt is not None:

        # marketingOptIn 取消勾选(Fluent UI Checkbox,真实 HTML 实证):
        # 用户给的两态 HTML 对比证明:<input> 在勾选/未勾选时完全一样(都无 checked、无 aria-checked 属性),
        # input.checked 不反映真实勾选态!唯一可靠信号是 .fui-Checkbox__indicator 内有没有 <svg> 勾选图标:
        #   未勾选 -> indicator 空 div;勾选 -> indicator 内有 <svg><path/></svg>。
        # 所以读 input.checked 恒 False(误判没勾->不点取消->提交时仍勾着)。改读 indicator 内 svg。
        # 点击:真实 BiDi click_self 点 <label>/<input> 能 toggle 翻转(合成 JS click 无效),indicator/span 无效。
        # toggle 语义:仅在判定已勾(svg 存在)时点一次翻回未勾;未勾时绝不点。
        def _mkt_is_checked():
            try:

                return bool(mkt.run_js(r"""function(){
  const inp = this;
  if (!inp) return false;
  const box = inp.closest('.fui-Checkbox') || inp.parentElement;
  if (!box) return false;
  const ind = box.querySelector('.fui-Checkbox__indicator');
  if (!ind) return false;
  // 勾选态:indicator 内渲染了 svg 勾选图标;未勾选时 indicator 为空 div
  return !!ind.querySelector('svg, path, [data-icon]');
}"""))

            except Exception:

                return False

        def _click_mkt_real():
            # 真实 BiDi 点击:优先 label(实测翻转可靠),回退 input
            for sel in ('css:label[for="marketingOptIn"]', 'css:#marketingOptIn'):

                try:

                    cand = _ele(page, sel, timeout=0.4)

                    if cand is None:

                        continue

                    cand.click_self()

                    return True

                except Exception:

                    continue

            return False

        # 确保未勾:最多点 3 次。每次真实点击后等 svg 重新渲染再读,
        # 只要读到未勾(indicator 无 svg)就停;点不翻或元素丢失则停,避免空转。
        _final_checked = _mkt_is_checked()
        for _attempt in range(3):
            if not _final_checked:
                break
            if not _click_mkt_real():
                break
            time.sleep(0.5)  # 等 React 重渲染 indicator(svg 增删)
            _final_checked = _mkt_is_checked()

    _click_next(page, tag, wait_before=False, wait_after=False)

    _shot(page, "after_name", idx)

    return True





def _find_hold_target(ctx):
    """??? browsing context ?? Press-and-hold ???"""
    hold_patterns = [
        r"press\s*(?:and|&)?\s*hold",
        r"long\s*press",
        "長押し",
        "按住|长按",
        r"appuyer\s*et\s*maintenir",
        r"\bhalten\b",
    ]
    challenge_patterns = [
        r"challenge|human|verify|verification",
        "チャレンジ|ヒューマン|検証|確認",
        "验证|驗證|人机|人機",
        "défi|vérifi",
        "prüfung|verifiz",
    ]
    done_patterns = [
        r"challenge\s+completed",
        r"completed,\s*please\s+wait",
        r"completed",
        r"已完成",
    ]
    script = (r"""
return (() => {
  const HOLD_PATTERNS = %s.map((src) => new RegExp(src, 'i'));
  const CHALLENGE_PATTERNS = %s.map((src) => new RegExp(src, 'i'));
  const DONE_PATTERNS = %s.map((src) => new RegExp(src, 'i'));
  const SHORT_LABEL_MAX = 32;
  const norm = (v) => String(v || '').replace(/\s+/g, ' ').trim();
  const matchesAny = (text, patterns) => {
    const s = norm(text);
    return !!s && patterns.some(re => re.test(s));
  };
  const textOf = (el) => norm([
    el?.getAttribute?.('aria-label') || '',
    el?.innerText || '',
    el?.textContent || '',
    el?.value || ''
  ].join(' '));
  const visible = (el) => {
    if (!el) return false;
    const s = getComputedStyle(el);
    if (s.display === 'none' || s.visibility === 'hidden' || Number(s.opacity) === 0) return false;
    const r = el.getBoundingClientRect();
    return r.width > 20 && r.height > 10 && r.bottom > 0 && r.right > 0;
  };
  const buttonish = (el) => {
    const tag = (el.tagName || '').toLowerCase();
    if (tag === 'button' || tag === 'input') return true;
    const role = (el.getAttribute('role') || '').toLowerCase();
    return role === 'button';
  };
  const pack = (el, quality, source, labelEl = null) => {
    const r = el.getBoundingClientRect();
    return {
      x: Math.round(r.x + r.width / 2),
      y: Math.round(r.y + r.height / 2),
      left: Math.round(r.x),
      top: Math.round(r.y),
      width: Math.round(r.width),
      height: Math.round(r.height),
      text: norm(el.innerText || el.textContent || el.value || '').slice(0, 80),
      id: el.id || '',
      tag: (el.tagName || '').toLowerCase(),
      role: el.getAttribute('role') || '',
      ariaLabel: norm(el.getAttribute('aria-label') || ''),
      quality,
      source: source || '',
      holdLabelId: labelEl && labelEl.id ? labelEl.id : ''
    };
  };
  const candidates = [];
  const seen = new Set();
  const isDone = (el) => {
    // PX 已完成(Human Challenge completed, please wait)：含 completed 文字的元素绝不当按压目标，
    // 否则完成态残留按钮会被反复按压直到 max_press 误判失败。
    const s = norm([el?.innerText || '', el?.textContent || '', el?.getAttribute?.('aria-label') || ''].join(' '));
    return !!s && DONE_PATTERNS.some(re => re.test(s));
  };
  const addCandidate = (el, quality, source, labelEl = null) => {
    if (!el || seen.has(el) || !visible(el) || !buttonish(el) || isDone(el)) return;
    seen.add(el);
    candidates.push(pack(el, quality, source, labelEl));
  };

  const px = document.querySelector('#px-captcha');
  if (px && visible(px) && buttonish(px)) addCandidate(px, 0, '#px-captcha');

  for (const el of document.querySelectorAll('[role="button"][aria-label], button[aria-label], a[role="button"][aria-label], input[aria-label]')) {
    const aria = norm(el.getAttribute('aria-label'));
    if (!aria) continue;
    const hasHold = matchesAny(aria, HOLD_PATTERNS);
    const hasChallenge = matchesAny(aria, CHALLENGE_PATTERNS);
    if (hasHold && hasChallenge) addCandidate(el, 1, 'aria-label:hold+challenge');
    else if (hasHold) addCandidate(el, 2, 'aria-label:hold');
  }

  for (const node of document.querySelectorAll('p, span, div')) {
    if (!visible(node)) continue;
    const txt = norm(node.innerText || node.textContent || '');
    if (!txt || txt.length > SHORT_LABEL_MAX || !matchesAny(txt, HOLD_PATTERNS)) continue;
    let cur = node;
    while (cur && cur !== document.body && cur !== document.documentElement) {
      if (buttonish(cur) && visible(cur)) {
        addCandidate(cur, 3, 'label->ancestor-button', node);
        break;
      }
      cur = cur.parentElement;
    }
  }

  for (const el of document.querySelectorAll('[role="button"], button, a[role="button"], input[type="button"], input[type="submit"]')) {
    if (matchesAny(textOf(el), HOLD_PATTERNS)) addCandidate(el, 5, 'button-text-hold');
  }

  if (!candidates.length) return null;
  candidates.sort((a, b) =>
    (a.quality - b.quality) ||
    ((a.width * a.height) - (b.width * b.height)) ||
    ((b.ariaLabel || '').length - (a.ariaLabel || '').length)
  );
  return candidates[0];
})()
""" % (json.dumps(hold_patterns, ensure_ascii=True), json.dumps(challenge_patterns, ensure_ascii=True), json.dumps(done_patterns, ensure_ascii=True)))
    try:
        target = ctx.run_js_loaded(script)
    except Exception:
        return None
    if not isinstance(target, dict):
        return None
    if target.get("width", 0) < 20 or target.get("height", 0) < 10:
        return None
    return target


def _find_hsprotect_iframe_box(page):

    """Return visible hsprotect iframe box in page coordinates as fallback."""

    script = r"""

return (() => {

  const frames = [...document.querySelectorAll('iframe[src*="hsprotect.net"], iframe[src*="arkose"], iframe[src*="funcaptcha"]')];

  const out = [];

  for (const f of frames) {

    const r = f.getBoundingClientRect();

    const s = getComputedStyle(f);

    if (s.display === 'none' || s.visibility === 'hidden') continue;

    if (r.width < 50 || r.height < 30) continue;

    out.push({

      x: Math.round(r.x + r.width / 2),

      y: Math.round(r.y + r.height * 0.55),

      left: Math.round(r.x),

      top: Math.round(r.y),

      width: Math.round(r.width),

      height: Math.round(r.height),

      text: 'iframe-box',

      id: f.id || '',

      tag: 'iframe',

      quality: 3,

      source: 'iframe-box',

      cx_ratio: 0.5,

      cy_ratio: 0.55

    });

  }

  if (!out.length) return null;

  out.sort((a, b) => (a.width * a.height) - (b.width * b.height));

  return out[0];

})()

"""

    try:

        box = page.run_js_loaded(script)

    except Exception:

        return None

    if not isinstance(box, dict):

        return None

    try:

        left = float(box.get("left", 0))

        top = float(box.get("top", 0))

        w = float(box.get("width", 0))

        h = float(box.get("height", 0))

        box["x"] = int(left + w * 0.5)

        box["y"] = int(top + h * 0.55)

        box["left"] = int(left)

        box["top"] = int(top)

    except Exception:

        pass

    return box





def _context_has_iframe_hint(page):

    try:

        script = """

return !!document.querySelector('iframe[src*="hsprotect.net"], iframe[src*="arkose"], iframe[src*="funcaptcha"]');

"""

        return bool(page.run_js_loaded(script))

    except Exception:

        return False





def _captcha_visible(page):

    for ctx in _all_contexts(page):

        t = _find_hold_target(ctx)

        if t and int(t.get("quality", 9)) <= 5:

            return True

    if _context_has_iframe_hint(page):

        return True

    low = _body_text(page).lower()

    return any(

        kw in low

        for kw in [

            "press and hold",

            "verify you're human",

            "captcha",

            "perimeterx",

            "appuyer et maintenir",

            "按住",

            "长按",

        ]

    )





def _px_captcha_completed_wait(page):

    """Detect PX loading by scanning #px-captcha subtree for the completed-wait phrase."""

    phrase = "human challenge completed, please wait"

    script = r"""

return (() => {

  const phrase = 'human challenge completed, please wait';

  const seen = new Set();

  const parts = [];

  const add = (v) => {

    if (v === undefined || v === null) return;

    const s = String(v).trim();

    if (s) parts.push(s);

  };

  const walk = (node, depth = 0) => {

    if (!node || depth > 10 || seen.has(node)) return;

    seen.add(node);

    try {

      add(node.innerText);

      add(node.textContent);

      if (node.getAttribute) {

        for (const name of ['aria-label', 'aria-live', 'title', 'role', 'data-testid']) {

          add(node.getAttribute(name));

        }

      }

      if (node.shadowRoot) walk(node.shadowRoot, depth + 1);

      if (node.tagName && node.tagName.toLowerCase() === 'iframe') {

        try {

          if (node.contentDocument) walk(node.contentDocument, depth + 1);

        } catch (e) {}

      }

      const children = node.children || node.childNodes || [];

      for (const child of children) walk(child, depth + 1);

    } catch (e) {}

  };

  const root = document.querySelector('#px-captcha');

  if (root) walk(root);

  // Some ruyipage frame contexts are already inside the nested #px-captcha iframe;

  // in that case #px-captcha is in the parent frame, so check this frame body too.

  if (!root) walk(document.body || document.documentElement);

  const text = parts.join('\n').toLowerCase();

  return text.includes(phrase);

})()

"""

    for ctx in _all_contexts(page):

        try:

            if bool(ctx.run_js_loaded(script)):

                return True

        except Exception:

            pass

    return False





def _resolve_hold_label_for_target(ctx, target):
    if not isinstance(target, dict):
        return None
    hold_patterns = [
        r"press\s*(?:and|&)?\s*hold",
        r"long\s*press",
        "長押し",
        "按住|长按",
        r"appuyer\s*et\s*maintenir",
        r"\bhalten\b",
    ]
    target_id = str(target.get("id") or "").strip()
    left = float(target.get("left", 0) or 0)
    top = float(target.get("top", 0) or 0)
    width = float(target.get("width", 0) or 0)
    height = float(target.get("height", 0) or 0)
    center_x = float(target.get("x", 0) or 0) or (left + width / 2.0)
    center_y = float(target.get("y", 0) or 0) or (top + height / 2.0)
    script = (r"""
return (() => {
  const targetId = %s;
  const cx = %r;
  const cy = %r;
  const HOLD_PATTERNS = %s.map((src) => new RegExp(src, 'i'));
  const SHORT_LABEL_MAX = 32;
  const norm = (v) => String(v || '').replace(/\s+/g, ' ').trim().toLowerCase();
  const textOf = (el) => [el.textContent || '', el.innerText || '', el.getAttribute('aria-label') || ''].join(' ');
  const matchesHold = (text) => {
    const s = norm(text);
    return !!s && s.length <= SHORT_LABEL_MAX && HOLD_PATTERNS.some(re => re.test(s));
  };
  const visible = (el) => {
    if (!el) return false;
    const s = getComputedStyle(el);
    if (s.display === 'none' || s.visibility === 'hidden' || Number(s.opacity) === 0) return false;
    const r = el.getBoundingClientRect();
    return r.width > 0 && r.height > 0 && r.bottom > 0 && r.right > 0;
  };
  const hits = [];
  const pushHit = (el) => {
    if (!visible(el)) return;
    const raw = textOf(el);
    if (!matchesHold(raw)) return;
    const r = el.getBoundingClientRect();
    const px = r.x + r.width / 2;
    const py = r.y + r.height / 2;
    hits.push({
      id: el.id || '',
      text: raw.trim().slice(0, 80),
      display: getComputedStyle(el).display || '',
      tag: (el.tagName || '').toLowerCase(),
      distance: Math.abs(px - cx) + Math.abs(py - cy),
    });
  };
  if (targetId) {
    const exactTarget = document.getElementById(targetId);
    if (exactTarget) {
      for (const el of exactTarget.querySelectorAll('p, span, div')) pushHit(el);
    }
  }
  for (const el of document.querySelectorAll('p, span, div')) pushHit(el);
  if (!hits.length) return null;
  hits.sort((a, b) => (a.distance - b.distance) || ((b.id ? 1 : 0) - (a.id ? 1 : 0)));
  return hits[0];
})()
""" % (json.dumps(target_id), center_x, center_y, json.dumps(hold_patterns, ensure_ascii=True)))
    try:
        state = ctx.run_js_loaded(script)
    except Exception:
        return None
    return state if isinstance(state, dict) else None


def _px_hold_instruction_state(ctx, hold_p_id=None):
    hold_patterns = [
        r"press\s*(?:and|&)?\s*hold",
        r"long\s*press",
        "長押し",
        "按住|长按",
        r"appuyer\s*et\s*maintenir",
        r"\bhalten\b",
    ]
    target_id = str(hold_p_id or "").strip()
    script = (r"""
return (() => {
  const targetId = %s;
  const HOLD_PATTERNS = %s.map((src) => new RegExp(src, 'i'));
  const SHORT_LABEL_MAX = 32;
  const norm = (v) => String(v || '').replace(/\s+/g, ' ').trim().toLowerCase();
  const textOf = (el) => [el.textContent || '', el.innerText || '', el.getAttribute('aria-label') || ''].join(' ');
  const matchesHold = (text) => {
    const s = norm(text);
    return !!s && s.length <= SHORT_LABEL_MAX && HOLD_PATTERNS.some(re => re.test(s));
  };
  const pack = (el) => {
    const raw = textOf(el);
    const s = getComputedStyle(el);
    return {
      id: el.id || '',
      text: raw.trim().slice(0, 80),
      display: s.display || '',
      visibility: s.visibility || '',
      opacity: s.opacity || '',
      hidden: s.display === 'none'
    };
  };
  if (targetId) {
    const exact = document.getElementById(targetId);
    if (!exact) return null;
    if (!matchesHold(textOf(exact))) return null;
    return pack(exact);
  }
  const hits = [];
  for (const el of document.querySelectorAll('p, span, div')) {
    if (!matchesHold(textOf(el))) continue;
    hits.push(pack(el));
  }
  if (!hits.length) return null;
  return hits.find(item => item.hidden) || hits.find(item => item.id) || hits[0];
})()
""" % (json.dumps(target_id), json.dumps(hold_patterns, ensure_ascii=True)))
    try:
        state = ctx.run_js_loaded(script)
    except Exception:
        return None
    return state if isinstance(state, dict) else None


def _captcha_is_validating(page):

    if _px_captcha_completed_wait(page):

        return True

    iframe_hint = _context_has_iframe_hint(page)

    if iframe_hint and not _find_hold_context(page)[1]:

        return True

    if not iframe_hint:

        return False

    low = _body_text(page).lower()

    if any(

        kw in low

        for kw in [

            "verifying",

            "verification",

            "checking",

            "loading",

            "please wait",

            "just a moment",

            "???",

            "???",

            "???",

            "???",

            "???",

            "???",

            "???",

            "v?rification",

            "chargement",

            "veuillez patienter",

        ]

    ):

        return True

    return False


_PX_VALIDATING_TEXT_KWS = (
    "verifying",
    "verification",
    "checking",
    "loading",
    "please wait",
    "just a moment",
    "v?rification",
    "chargement",
    "veuillez patienter",
)


_PX_CAPTCHA_TEXT_KWS = (
    "press and hold",
    "verify you're human",
    "captcha",
    "perimeterx",
    "appuyer et maintenir",
    "按住",
    "长按",
)


def _px_probe(page, min_quality=5):
    """合并探测：一次拿 visible/validating/hold_ctx/hold_target，等价于
    _captcha_visible + _captcha_is_validating + _find_hold_context 但 JS 往返
    从 3N 降到 N（主循环热路径，多线程下是主要拖慢来源）。

    关键：iframe-box(只拿到 PX 容器中心、真按钮还没渲染)绝不能当可按压目标——
    直接按中心会选中 iframe 内容导致人机挑战不加载/卡死。这种情况只标记
    captcha_hint=True 让主循环等真按钮渲染，不按、不提交。"""

    hold_ctx, hold_target = _find_hold_context(page, min_quality=min_quality)

    target_quality = _target_quality(hold_target)
    is_iframe_box = isinstance(hold_target, dict) and hold_target.get("source") == "iframe-box"
    # 真按钮(非 iframe-box)且质量达标才算可按；iframe-box 只当"PX 在加载"信号
    target_actionable = (
        hold_target is not None
        and not is_iframe_box
        and target_quality <= min_quality
    )
    captcha_hint = bool(is_iframe_box)

    iframe_hint = _context_has_iframe_hint(page)

    visible = bool(target_actionable or iframe_hint or captcha_hint)

    if not visible:
        low = _body_text(page).lower()
        visible = any(kw in low for kw in _PX_CAPTCHA_TEXT_KWS)
        if visible:
            captcha_hint = True

    validating = False
    if target_actionable:
        validating = False
    elif _px_captcha_completed_wait(page):
        validating = True
    elif iframe_hint:
        # 复刻原 _captcha_is_validating 语义：iframe 在但拿不到任何 hold_target 才算 validating；
        # 拿到 iframe-box(非 None)不算 validating，只算"等真按钮渲染"。
        if hold_target is None:
            validating = True
        else:
            low = _body_text(page).lower()
            validating = any(kw in low for kw in _PX_VALIDATING_TEXT_KWS)

    return visible, validating, hold_ctx, hold_target, captcha_hint






def _microsoft_loading_page(page):

    """Detect the Microsoft full-page Loading spinner after captcha/signup submit.



    PX iframe can remain in the DOM while Microsoft is already moving to the

    next step. Treat this page as navigation/loading, not captcha failure.

    """

    txt = (_body_text(page) or "").strip()

    if not txt:

        return False

    low = txt.lower()

    first = low.splitlines()[0].strip() if low.splitlines() else low

    if first not in ("loading", "loading...") and not first.startswith("loading"):

        return False

    if "press and hold" in low or "human challenge completed" in low:

        return False

    # The saved failure page only contains Loading + Microsoft footer links.

    if "privacy and cookies" in low or "terms of use" in low or len(low) < 400:

        return True

    return False





def _target_quality(target):

    try:

        return int((target or {}).get("quality", 9))

    except Exception:

        return 9





def _find_hold_context(page, min_quality=5):

    """Find best press-and-hold target; prefer hsprotect iframe button."""

    hits = []

    for ctx in _all_contexts(page):

        target = _find_hold_target(ctx)

        if not target:

            continue

        q = _target_quality(target)

        if q > min_quality:

            continue

        url = (getattr(ctx, "url", "") or "").lower()

        frame_bonus = 0

        if any(k in url for k in ["hsprotect.net", "arkose", "funcaptcha"]):

            frame_bonus = -10

        elif ctx is not page and url:

            frame_bonus = -3

        rank = q + frame_bonus

        hits.append((rank, q, url, ctx, target))



    if hits:

        hits.sort(key=lambda item: (item[0], item[1], item[2]))

        return hits[0][3], hits[0][4]



    box = _find_hsprotect_iframe_box(page)

    if box is not None:

        return page, box

    return None, None





def _try_submit(page, tag, last_logged_hit=None):

    _submit_wait(tag, "submit")

    hit = _click_any(page, ['#iSignupAction', 'css:input[type="submit"]', 'css:button[type="submit"]'], timeout=1)

    if hit and hit != last_logged_hit:

        log(f"  {tag} submit click: {hit}")

    return bool(hit), (hit if hit else last_logged_hit)





def _maybe_skip_passkey(page, tag):

    low_url = page.url.lower()

    if "fido" not in low_url and "passkey" not in low_url:

        return False

    if _click_any(

        page,

        ['text:Skip', 'text:No thanks', 'text:Cancel', 'text:跳过', '#skipBtn', 'css:button[type="submit"]'],

        timeout=2,

    ):

        log(f"  {tag} 跳过 passkey/FIDO")

        time.sleep(3)

        return True

    return False





def _post_signup_cleanup(page, tag, idx):

    for retry in range(12):

        current_url = page.url.lower()

        if not _on_signup_form(current_url) and "privacynotice" not in current_url:

            return True

        if _maybe_skip_passkey(page, tag):

            continue

        _click_post_signup(page, tag)

        if retry % 3 == 0:

            _shot(page, f"post_cleanup_{retry}", idx)

        time.sleep(3)

    return not _on_signup_form(page.url.lower())






def _ruoyi_point_inside_target(target, ratio_x=0.5, ratio_y=0.55, jitter_x=0.0, jitter_y=0.0):

    base_x = float((target or {}).get("x", 0) or 0)

    base_y = float((target or {}).get("y", 0) or 0)

    left = float((target or {}).get("left", base_x) or base_x)

    top = float((target or {}).get("top", base_y) or base_y)

    width = float((target or {}).get("width", 0) or 0)

    height = float((target or {}).get("height", 0) or 0)

    if width <= 1.0 and height <= 1.0 and (base_x or base_y):

        return (

            int(round(base_x + random.uniform(-jitter_x, jitter_x))),

            int(round(base_y + random.uniform(-jitter_y, jitter_y))),

        )

    width = max(1.0, width)

    height = max(1.0, height)

    ratio_x = min(0.9, max(0.1, float(ratio_x)))

    ratio_y = min(0.9, max(0.1, float(ratio_y)))

    return (

        int(round(left + (width * ratio_x) + random.uniform(-jitter_x, jitter_x))),

        int(round(top + (height * ratio_y) + random.uniform(-jitter_y, jitter_y))),

    )



def _perform_hold(page, ctx, target, idx, press_count, tag):

    """Press-and-hold with early release when the linked Press and hold p hides itself."""

    hold_sec = random.uniform(PX_HOLD_SECONDS_MIN, PX_HOLD_SECONDS_MAX)

    motion_profile = getattr(page, "_ruoyi_px_motion_profile", None) or {}

    jitter_x = float(motion_profile.get("jitter_x", 1.8))

    jitter_y = float(motion_profile.get("jitter_y", 1.2))

    cx, cy = _ruoyi_point_inside_target(
        target,
        ratio_x=motion_profile.get("press_ratio_x", 0.5),
        ratio_y=motion_profile.get("press_ratio_y", 0.55),
        jitter_x=jitter_x,
        jitter_y=jitter_y,
    )

    settle_x, settle_y = _ruoyi_point_inside_target(
        target,
        ratio_x=motion_profile.get("settle_ratio_x", 0.48),
        ratio_y=motion_profile.get("settle_ratio_y", 0.57),
        jitter_x=float(motion_profile.get("settle_jitter_x", 0.8)),
        jitter_y=float(motion_profile.get("settle_jitter_y", 0.8)),
    )

    micro_x, micro_y = _ruoyi_point_inside_target(
        target,
        ratio_x=motion_profile.get("micro_ratio_x", motion_profile.get("press_ratio_x", 0.5)),
        ratio_y=motion_profile.get("micro_ratio_y", motion_profile.get("press_ratio_y", 0.55)),
        jitter_x=float(motion_profile.get("micro_jitter_x", 0.45)),
        jitter_y=float(motion_profile.get("micro_jitter_y", 0.45)),
    )

    lead_x = int(round(settle_x + float(motion_profile.get("lead_dx", -10.0)) + random.uniform(-1.2, 1.2)))

    lead_y = int(round(settle_y + float(motion_profile.get("lead_dy", -4.0)) + random.uniform(-1.0, 1.0)))

    move_ms = int(min(420, max(200, hold_sec * float(motion_profile.get("move_ratio", 0.03)) * 1000.0)))

    lead_ms = int(move_ms * float(motion_profile.get("lead_ratio", 0.48)))

    lead_ms = min(max(60, lead_ms), move_ms - 120)

    settle_ms = int(move_ms * float(motion_profile.get("settle_ratio", 0.24)))

    settle_ms = min(max(45, settle_ms), move_ms - lead_ms - 70)

    micro_ms = int(move_ms * float(motion_profile.get("micro_ratio", 0.14)))

    micro_ms = min(max(25, micro_ms), move_ms - lead_ms - settle_ms - 35)

    final_ms = move_ms - lead_ms - settle_ms - micro_ms

    pre_hover_sec = min(0.22, max(0.04, hold_sec * float(motion_profile.get("hover_ratio", 0.01))))

    state_ctx = ctx

    hold_label = _resolve_hold_label_for_target(state_ctx, target) or {}

    hold_p_id = str(hold_label.get("id") or "").strip()

    log(

        f"  {tag} press #{press_count}: ({cx},{cy}) hold={hold_sec:.1f}s path={move_ms}ms"

        + (f" text={target.get('text', '')[:30]!r}" if target.get("text") else "")

        + (f" hold_p_id={hold_p_id}" if hold_p_id else "")

    )

    actions = ctx.actions

    hold_started = None

    release_state = None

    try:

        actions.move_to({"x": lead_x, "y": lead_y}, duration=lead_ms)

        actions.move_to({"x": settle_x, "y": settle_y}, duration=settle_ms)

        actions.move_to({"x": micro_x, "y": micro_y}, duration=micro_ms)

        actions.move_to({"x": cx, "y": cy}, duration=final_ms)

        if pre_hover_sec > 0:

            actions.wait(pre_hover_sec)

        actions.hold().perform()

        hold_started = time.time()

        next_check = hold_started + PX_HOLD_EARLY_RELEASE_AFTER

        while True:

            now = time.time()

            elapsed = now - hold_started

            if elapsed >= hold_sec:

                break

            if now >= next_check:

                release_state = _px_hold_instruction_state(state_ctx, hold_p_id)

                if release_state is None:

                    refreshed_ctx, refreshed_target = _find_hold_context(page)

                    if refreshed_ctx is None:

                        refreshed_ctx, refreshed_target = state_ctx, target

                    refreshed_label = _resolve_hold_label_for_target(refreshed_ctx, refreshed_target) or {}

                    refreshed_id = str(refreshed_label.get("id") or "").strip()

                    if refreshed_id and refreshed_id != hold_p_id:

                        state_ctx = refreshed_ctx

                        hold_p_id = refreshed_id

                        log(f"  {tag} press #{press_count} hold label id refreshed -> {refreshed_id}")

                    release_state = _px_hold_instruction_state(state_ctx, hold_p_id)

                if release_state and str(release_state.get("display") or "").strip().lower() == "none":

                    break

                next_check += PX_HOLD_EARLY_RELEASE_INTERVAL

                continue

            sleep_for = min(0.1, max(0.0, next_check - now), hold_sec - elapsed)

            time.sleep(sleep_for if sleep_for > 0 else 0.05)

        actions.release().perform()

        held = time.time() - hold_started

        if release_state and str(release_state.get("display") or "").strip().lower() == "none":

            label_id = str(release_state.get("id") or hold_p_id or "").strip()

            label_text = str(release_state.get("text") or "").strip()

            extra = f" id={label_id}" if label_id else ""

            extra += " display=none"

            if label_text:

                extra += f" text={label_text[:20]!r}"

            log(f"  {tag} press #{press_count} released early after {held:.1f}s because hold label hid{extra}")

        return held

    except Exception as exc:

        try:

            actions.release_all()

        except Exception:

            pass

        log(f"  {tag} hold failed: {type(exc).__name__}: {exc}", "WARN")

        return False



def _perform_hold_with_px_screenshots(page, ctx, target, idx, press_count, tag, enabled=False):

    ok = _perform_hold(page, ctx, target, idx, press_count, tag)

    if enabled and not ok:

        _save_screenshot(page, f"press_fail_last_{press_count}", idx, tag)

    return ok





def _new_ruoyi_px_motion_profile():
    return {
        "press_ratio_x": random.uniform(0.42, 0.58),
        "press_ratio_y": random.uniform(0.48, 0.64),
        "settle_ratio_x": random.uniform(0.44, 0.56),
        "settle_ratio_y": random.uniform(0.50, 0.62),
        "micro_ratio_x": random.uniform(0.45, 0.57),
        "micro_ratio_y": random.uniform(0.49, 0.63),
        "jitter_x": random.uniform(1.0, 2.8),
        "jitter_y": random.uniform(0.8, 2.0),
        "settle_jitter_x": random.uniform(0.4, 1.0),
        "settle_jitter_y": random.uniform(0.3, 0.9),
        "micro_jitter_x": random.uniform(0.2, 0.6),
        "micro_jitter_y": random.uniform(0.2, 0.6),
        "lead_dx": random.choice((-1, 1)) * random.uniform(8.0, 16.0),
        "lead_dy": random.uniform(-6.0, 6.0),
        "move_ratio": random.uniform(0.024, 0.036),
        "lead_ratio": random.uniform(0.42, 0.56),
        "settle_ratio": random.uniform(0.20, 0.28),
        "micro_ratio": random.uniform(0.10, 0.18),
        "hover_ratio": random.uniform(0.008, 0.016),
    }



def _wait_before_next_captcha_press(tag, reason="challenge failed"):

    delay = random.uniform(POST_PRESS_RETRY_GAP_MIN, POST_PRESS_RETRY_GAP_MAX)

    log(f"  {tag} {reason}, retry press in {delay:.2f}s")

    time.sleep(delay)

    return delay





# B 组 IP/代理探测尾段:re-export 自新包(单一源)。
# _proxy_for_ip_lookup / _log_current_ip:纯逻辑,日志走新包骨架 log。
_proxy_for_ip_lookup = _ruyi_pkg._proxy_for_ip_lookup
_log_current_ip = _ruyi_pkg._log_current_ip

# _apply_ruoyi_proxy_geo_emulation 已移除:ruyipage 155 内核在 privileged scope 拒绝
# emulation.setTimezoneOverride/setGeolocationOverride,且反复打 IP 查询服务只为设
# timezone/geo,对注册/解锁成功率无增益。保留 _probe_proxy_identity/_log_current_ip
# 用于出口 IP 记录与去重。


# B 组代理预检:re-export 自新包(单一源)。targets 默认读 _state.PROXY_PRECHECK_TARGETS,
# 兼容层已注入业务 signup.live.com targets;空 targets 时新包跳过预检返回 True。
# timeout 默认读 _state.PROXY_PRECHECK_TIMEOUT(已同步业务值)。
_probe_proxy_before_browser = _ruyi_pkg._probe_proxy_before_browser
_probe_proxy_targets = _ruyi_pkg._probe_proxy_targets

def _log_current_ip_async(proxy_pool, tag):

    pool_snapshot = list(proxy_pool or [])



    def _worker():

        try:

            _log_current_ip(pool_snapshot, tag)

        except Exception as exc:

            log(f"  {tag} async IP probe failed: {type(exc).__name__}: {exc}", "WARN")



    th = threading.Thread(target=_worker, name=f"ruoyi-ip-probe-{tag}", daemon=True)

    th.start()

    return th





def _apply_account_options(opts=None):

    """把 CLI/opts 的账号格式、后缀配置落到当前线程 helper 上下文。"""

    helpers = _load_helpers()

    opts = opts or SimpleNamespace()



    email_suffixes = (

        getattr(opts, "email_suffixes", None)

        or os.environ.get("OUTLOOK_ACCOUNT_SUFFIXES")

        or os.environ.get("OUTLOOK_EMAIL_SUFFIXES")

        or ""

    )

    if email_suffixes:

        email_suffixes = _normalize_email_suffixes(email_suffixes)

        os.environ["OUTLOOK_ACCOUNT_SUFFIXES"] = str(email_suffixes).strip()



    custom = (

        getattr(opts, "account_format", None)

        or os.environ.get("OUTLOOK_ACCOUNT_FORMAT")

        or ""

    )

    mode = (

        getattr(opts, "account_format_mode", None)

        or os.environ.get("OUTLOOK_ACCOUNT_FORMAT_MODE")

        or ("custom" if os.environ.get("OUTLOOK_ACCOUNT_FORMAT") else "name")

    )

    # 用户填了指定格式模板就按模板走,无视 mode 预设(WebUI 默认 name,易漏切到 custom)

    if str(custom).strip() and str(mode).strip().lower() != "custom":

        mode = "custom"



    password_format = (

        getattr(opts, "password_format", None)

        or os.environ.get("OUTLOOK_PASSWORD_FORMAT")

        or ""

    )

    set_fn = getattr(helpers, "set_account_generation_options", None)

    resolve_format_fn = getattr(helpers, "resolve_account_format", None)

    resolved_format = ""

    if callable(resolve_format_fn):

        # 邮箱前缀格式统一由 ACCOUNT_FORMAT_PRESETS/resolve_account_format 决定:
        # random→纯随机, custom→用户填的固定内容(不展开模板,见 _generate_prefix 的 literal 分支)。
        # 旧 name/name_digits 模式已并入 custom,预设统一 {letters:7}{digits:6}(约 8e15 种,taken≈0)。
        resolved_format = resolve_format_fn(mode, custom)

    elif custom and str(mode).lower() == "custom":

        resolved_format = str(custom)

    elif str(mode).lower() == "random":

        resolved_format = ""

    if callable(set_fn):

        set_fn(

            email_suffixes=email_suffixes,

            account_format_mode=mode,

            account_format=custom,

            password_format=password_format,

        )

    else:

        if email_suffixes:

            os.environ["OUTLOOK_ACCOUNT_SUFFIXES"] = str(email_suffixes).strip()

        os.environ["OUTLOOK_ACCOUNT_FORMAT_MODE"] = str(mode)

        apply_fn = getattr(helpers, "apply_account_format_mode", None)

        if callable(apply_fn):

            apply_fn(mode, custom)

        elif resolved_format:

            os.environ["OUTLOOK_ACCOUNT_FORMAT"] = resolved_format

        else:

            os.environ.pop("OUTLOOK_ACCOUNT_FORMAT", None)

        if password_format:

            os.environ["OUTLOOK_PASSWORD_FORMAT"] = str(password_format)

        else:

            os.environ.pop("OUTLOOK_PASSWORD_FORMAT", None)



    log(

        f"账号格式: mode={str(mode)} "

        f"format={resolved_format!r} "

        f"suffixes={str(email_suffixes or 'outlook.com')!r} "

        f"pwd_format={str(password_format or '')!r}"

    )





def register_outlook(opts, proxy_pool, idx):

    from ruyipage import FirefoxOptions, FirefoxPage



    _install_shutdown_handlers()

    helpers = _load_helpers()

    set_log_level(getattr(opts, "log_level", None) or os.environ.get("OUTLOOK_LOG_LEVEL", "INFO"))

    _apply_account_options(opts)

    generate_email_password = helpers.generate_email_password

    generate_birthday = helpers.generate_birthday

    generate_name = helpers.generate_name

    verify_registered_outlook = helpers.verify_registered_outlook



    tag = f"[#{idx}][ruoyi]"

    raw_timeout = getattr(opts, "timeout", None)
    if raw_timeout in (None, ""):
        raw_timeout = REGISTER_TIMEOUT
    try:
        timeout = max(0.0, float(raw_timeout))
    except Exception:
        timeout = float(REGISTER_TIMEOUT)

    max_press = int(getattr(opts, "max_press", os.environ.get("OUTLOOK_REG_MAX_PRESS", "5")) or 5)

    need_verify = not bool(getattr(opts, "no_verify", False))

    confirm_before_register = bool(getattr(opts, "confirm_before_register", False))

    is_headless = bool(getattr(opts, "headless", False))

    capture_har = bool(getattr(opts, "har", False)) or _env_bool("OUTLOOK_RUOYI_HAR", False)

    px_press_screenshots = bool(getattr(opts, "px_press_screenshots", False)) or _env_bool(

        "OUTLOOK_PX_PRESS_SCREENSHOTS", False

    )

    step_timings = {}
    register_started = time.perf_counter()

    user_agent = _pick_user_agent(idx)

    log(f"  {tag} ua pool pick -> {_mask_ua(user_agent)} ({user_agent[:72]}...)")



    tb = FirefoxOptions()

    tb.set_browser_path(RUOYI_FIREFOX_PATH)
    profile_dir = _ruoyi_profile_dir(opts, idx)
    tb.set_profile(profile_dir)

    selected_browser_proxy = str(proxy_pool[0] or "").strip() if proxy_pool else ""

    if selected_browser_proxy:

        tb.set_proxy(selected_browser_proxy)

        # log(f"挂载浏览器代理: {mask_ruoyi_proxy(selected_browser_proxy)}")

    else:

        log("没挂代理——直接本机出口", "WARN")

    # 有头/无头都写 UA，避免 4 并发全是同一条默认 UA

    _apply_ruoyi_browser_ua(tb, tag, user_agent)

    _apply_ruoyi_quiet_prefs(tb, tag)

    if is_headless:

        _apply_ruoyi_headless_options(tb, tag, user_agent=user_agent)

        tb.headless(True)

    # 直接以注册页 URL 启动 Firefox(命令行末尾带 URL),省掉「先开空白页再 page.get 导航」的等待。
    # 默认开启;OUTLOOK_RUOYI_DIRECT_SIGNUP=0 关闭走原 about:blank 启动+导航流程。
    _direct_signup = _env_bool("OUTLOOK_RUOYI_DIRECT_SIGNUP", True)

    # 首屏 warmup:先落地 login.live.com 首页种下第一方挑战 cookie 再进注册页,
    # 免"零 cookie 新浏览器直达注册页"触发挑战拦截。OUTLOOK_RUOYI_WARMUP=0 关闭。
    _warmup_enabled = _env_bool("OUTLOOK_RUOYI_WARMUP", True)

    _warmup_seed_enabled = _env_bool("OUTLOOK_RUOYI_WARMUP_SEED", True)

    _launch_url = SIGNUP_URL if _direct_signup else ""

    if _warmup_enabled:

        # warmup 时启动 URL 换成首页(注册页导航延后到 warmup 之后)
        _launch_url = WARMUP_HOME_URL

    if _launch_url:

        try:

            tb.set_argument(_launch_url)

        except Exception as _e:

            log(f"  {tag} set_argument({_launch_url}) 失败,回退空白页启动: {_e}", "WARN")

            _direct_signup = False

            _launch_url = ""

            _warmup_enabled = False



    log(

        f"启动 ruyipage Firefox: model={_browser_model_name(RUOYI_FIREFOX_PATH)} "

        f"headless={is_headless} path={RUOYI_FIREFOX_PATH} profile={profile_dir} ua={_mask_ua(user_agent)}",

        "INFO",

    )

    browser_page = None

    page = None

    email = password = None

    success = False

    har_collector = None

    failure_reason = "failure"

    # 三元返回的原因盒：blocked / failure / exception_*；成功为空串

    fail_reason_box = ["failure"]
    px_hold_elapsed = 0.0



    def _current_px_metrics():

        px_elapsed = px_hold_elapsed or step_timings.get("captcha")

        if px_elapsed is None and captcha_started is not None and (had_captcha or press_count > 0):

            px_elapsed = max(0.0, time.perf_counter() - captcha_started)

        return {

            "idx": idx,

            "max_presses": int(press_count or 0),

            "px_elapsed": float(px_elapsed or 0.0),
            "reg_elapsed": max(0.0, time.perf_counter() - register_started),

        }



    def _finish(email_out=None, password_out=None, reason="failure"):

        """?????(email, password, fail_reason, px_metrics)?"""

        nonlocal success, failure_reason

        px_metrics = _current_px_metrics()

        if email_out and password_out:

            success = True

            failure_reason = "success"

            fail_reason_box[0] = ""

            return email_out, password_out, "", px_metrics

        failure_reason = reason or failure_reason or "failure"

        fail_reason_box[0] = failure_reason

        return None, None, failure_reason, px_metrics

    deadline = time.time() + timeout

    press_wait_started = None

    had_captcha = False

    gone_rounds = 0

    press_count = 0

    captcha_started = None

    captcha_signup_url = ""

    no_target_rounds = 0

    initial_press_wait_started = None

    button_wait_started = None

    validation_wait_started = None

    microsoft_loading_wait_started = None

    submit_wait_started = None
    last_submit_click_hit = None

    # 按压成功后：等 loading/消失，再等 captcha 重新出现后才允许下一次按压

    awaiting_reappear = False

    post_press_saw_gap = False

    post_press_started_at = None

    headless_patch_logged = False

    signup_opened = False



    try:

        browser_page = FirefoxPage(tb)

        _track_browser_page(browser_page)

        page = browser_page
        setattr(page, "_ruoyi_px_motion_profile", _new_ruoyi_px_motion_profile())



        if proxy_pool:

            log(f"  {tag} 使用当前打开页面承载注册页，不再新建 container tab")

        if _direct_signup and not _warmup_enabled:

            # 直接以 SIGNUP_URL 启动:Firefox 启动即注册页,跳过关空白页+page.get 导航
            # warmup 模式启动 URL 是首页,注册页导航延后,warmup 之后走 open_signup

            signup_opened = True

        else:

            try:

                browser_page.close_other_tabs(page)

                log(f"  {tag} 已关闭 Firefox 启动默认空白页，仅保留当前注册页")

            except Exception as close_exc:

                log(f"  {tag} 关闭默认空白页失败: {type(close_exc).__name__}: {close_exc}", "WARN")



        if bool(getattr(opts, "block_resources", False)):

            _start_ruoyi_resource_blocking(page, tag)

        _apply_ruoyi_headless_page_patches(page, tag, log_once=True, user_agent=user_agent)

        headless_patch_logged = True

        if _warmup_enabled:

            # 首屏预热:首页停留种第一方 cookie + 回灌同代理同UA种子(挑战 cookie 复用)
            _warmup_seed_list = []

            if _warmup_seed_enabled and selected_browser_proxy:

                _warmup_seed_list = _ruyi_pkg.load_seed(
                    "outlook", selected_browser_proxy, user_agent
                )

                if _warmup_seed_list:

                    log(f"  {tag} 命中种子池: {len(_warmup_seed_list)} 条挑战 cookie 待回灌")
                else:

                    log(f"  {tag} 种子池未命中(新代理/UA 或已过期),本轮当探路")

            _, step_timings["warmup"] = _timed_step(
                tag, "warmup", _ruyi_pkg.warmup, page, tag,
                home_url=WARMUP_HOME_URL, seed_cookies=_warmup_seed_list,
            )

        if capture_har:

            har_collector = _RuoyiHarCollector(page, tag, idx, email_getter=lambda: email or "")

            har_collector.start()

        if not signup_opened:

            _, step_timings["open_signup"] = _timed_step(tag, "open_signup", page.get, SIGNUP_URL)

        _, step_timings["wait_loading"] = _timed_step(tag, "wait_loading", page.wait_loading, 20)

        _apply_ruoyi_headless_page_patches(page, tag, log_once=False, user_agent=user_agent)

        try:

            browser_ua = page.run_js_loaded("return navigator.userAgent") or ""

        except Exception:

            browser_ua = ""

        log(

            f"  {tag} browser model: {_browser_model_name(RUOYI_FIREFOX_PATH)} "

            f"ua={browser_ua[:120]!r} url={page.url}",

            "INFO",

        )

        _log_current_ip_async(proxy_pool, tag)

        _shot(page, "start", idx)



        if confirm_before_register:

            _, step_timings["confirm_before_register"] = _timed_step(tag, "confirm_before_register", _click_post_signup, page, tag)

            time.sleep(3)

        signup_step, step_timings["enter_signup"] = _timed_step(

            tag,

            "enter_signup",

            _ensure_signup_entry,

            page,

            tag,

            idx,

            SIGNUP_ENTRY_TIMEOUT,

            detail=lambda result: f"step={result or 'timeout'}",

        )

        if signup_step == "blocked":

            return _finish(reason="blocked")

        if signup_step == "problem":

            return _finish(reason="problem")

        if not signup_step:

            return _finish(reason="signup_entry_timeout")



        email, password, prefix = generate_email_password()

        log(f"  {tag} 将注册: {email} / {password}")



        # 邮箱占用可在本号内多轮换号重试；密码步若撞回 taken 也回填邮箱，不直接整号 FAIL。

        ok_password = False

        for email_round in range(5):

            ok_email, step_timings[f"fill_email_r{email_round}"] = _timed_step(

                tag,

                "fill_email",

                _fill_email,

                page,

                email,

                prefix,

                tag,

                idx,

                detail=lambda result: f"email={result}" if result else "",

            )

            if not ok_email:

                if email_round < 4 and _still_on_email_or_taken(page):

                    email, password, prefix = generate_email_password()

                    log(f"  {tag} email step failed but still on form, new candidate: {email}", "WARN")

                    continue

                return _finish(reason="failure")

            email = ok_email



            # 二次闸：邮箱占用未清干净 → 换号重填，不进密码步

            if _still_on_email_or_taken(page):

                log(f"  {tag} email step returned but still on email/taken, rotate", "WARN")

                _shot(page, "email_taken_after_fill", idx)

                email, password, prefix = generate_email_password()

                prefix = email.split("@", 1)[0]

                log(f"  {tag} new email candidate: {email}")

                continue



            ok_password, step_timings[f"fill_password_r{email_round}"] = _timed_step(

                tag, "fill_password", _fill_password, page, password, tag, idx

            )

            if ok_password is True:

                break

            if ok_password == "email_taken":

                # 密码步发现仍在占用页 → 换号回邮箱步

                email, password, prefix = generate_email_password()

                prefix = email.split("@", 1)[0]

                log(f"  {tag} password saw email-taken, retry with: {email}", "WARN")

                continue

            # 真密码失败

            return _finish(reason="failure")

        else:

            log(f"  {tag} email/password rounds exhausted", "ERR")

            return _finish(reason="failure")

        if not ok_password:

            return _finish(reason="failure")



        year, month, day = generate_birthday()

        ok_birthday, step_timings["fill_birthday"] = _timed_step(

            tag, "fill_birthday", _fill_birthday, page, year, month, day, tag, idx

        )

        if not ok_birthday:

            return _finish(reason="failure")



        first, last = generate_name()

        ok_name, step_timings["fill_name_and_terms"] = _timed_step(

            tag, "fill_name_and_terms", _fill_name_and_terms, page, first, last, prefix, tag, idx

        )

        if not ok_name:

            return _finish(reason="failure")



        captcha_started = time.perf_counter()

        # submit 卡死追踪:记录"开始能点 submit 后页面仍无进展"的起点与当时 URL。
        # 若持续点中 submit 但 SUBMIT_STUCK_NOCHANGE_TIMEOUT 内 URL 仍未变、且未出 captcha/loading,
        # 判定卡死(代理抖动/出口IP风控致提交无响应),早退 fail 换号,不死等到 deadline。
        submit_first_click_at = None
        submit_first_click_url = None

        # 无进展总超时:任何有意义进展(按压、captcha 出现、URL 变化)更新此时间戳,
        # 超过 NO_PROGRESS_TIMEOUT 无进展判死循环卡死,早退 fail(覆盖 submit 后 loading 不结束等场景)。
        last_progress_at = time.time()
        last_progress_url = page.url.lower()

        while time.time() < deadline:

            _apply_ruoyi_headless_page_patches(

                page, tag, log_once=(not headless_patch_logged), user_agent=user_agent

            )

            headless_patch_logged = True

            submitted = False

            current_url = page.url.lower()

            body = _body_text(page)

            low = body.lower()

            # 无进展超时:URL 变化(任何跳转/loading)算进展重新计时;否则累计无进展,
            # 超过 NO_PROGRESS_TIMEOUT 判死循环卡死(覆盖 submit 后 loading 不结束、点不中 submit 空转等)。
            if current_url != last_progress_url:
                last_progress_url = current_url
                last_progress_at = time.time()
            else:
                no_progress = time.time() - last_progress_at
                if no_progress >= NO_PROGRESS_TIMEOUT and not had_captcha:
                    log(f"  {tag} no progress for {int(no_progress)}s on signup, give up", "WARN")
                    _shot(page, "no_progress_timeout", idx)
                    return _finish(reason="timeout")



            if not _on_signup_form(current_url) and any(

                h in current_url

                for h in [

                    "privacynotice",

                    "account.microsoft.com",

                    "account.live.com",

                    "outlook.live.com",

                    "outlook.office",

                    "login.live.com/oauth20",

                ]

            ):

                log(f"  {tag} captcha passed, left signup -> {current_url[:70]}")

                break

            if "outlook" in current_url and "signup" not in current_url and "login" not in current_url:

                log(f"  {tag} registration complete!")

                break

            if any(kw in low for kw in ["welcome", "inbox", "account has been created"]):

                log(f"  {tag} registration complete!")

                break

            if "signup" not in current_url and "live.com" in current_url:

                log(f"  {tag} left signup: {current_url[:60]}")

                break



            if any(

                kw in low

                for kw in [

                    "账户创建已被阻止",

                    "已被阻止",

                    "阻止创建",

                    "account creation has been blocked",

                    "has been blocked",

                    "account has been suspended",

                    "cr?ation de compte a ?t? bloqu?e",

                    "bloqu?e",

                    "unusual activity",

                    "异常活动",

                    "activit? inhabituelle",

                ]

            ):

                log(f"  {tag} BLOCKED: account creation blocked by Microsoft", "WARN")

                _shot(page, "blocked", idx)

                return _finish(reason="blocked")

            if had_captcha and _captcha_signup_url_changed(captcha_signup_url, current_url):

                log(f"  {tag} captcha passed, signup url changed -> {current_url[:70]}")

                break



            if _maybe_skip_passkey(page, tag):

                submit_wait_started = None
                last_submit_click_hit = None
                submit_first_click_at = None
                submit_first_click_url = None

                continue

            if "privacynotice" in current_url:

                submit_wait_started = None
                last_submit_click_hit = None
                submit_first_click_at = None
                submit_first_click_url = None

                _click_post_signup(page, tag)

                time.sleep(3)

                continue



            visible, validating, hold_ctx, hold_target, captcha_hint = _px_probe(page, min_quality=5)

            # iframe-box 兜底(PX 容器在、真按钮未渲染)不算可按目标，只在 captcha_hint 里用
            is_iframe_box_target = isinstance(hold_target, dict) and hold_target.get("source") == "iframe-box"

            actionable = bool(

                visible

                and (not validating)

                and hold_target is not None

                and (not is_iframe_box_target)

                and _target_quality(hold_target) <= 5

            )

            # PX 已出现(iframe-box 或文本命中)但真按钮还没渲染时，先标记 had_captcha，
            # 阻止下面的 _try_submit 在按钮渲染前抢点 submit。
            if captcha_hint and not had_captcha:
                had_captcha = True
                last_progress_at = time.time()
                if not captcha_signup_url:
                    captcha_signup_url = current_url

            loading_page = _microsoft_loading_page(page)
            loop_now = time.time()
            press_wait_started, microsoft_loading_wait_started, loading_timed_out = _update_loading_wait_state(
                press_wait_started,
                microsoft_loading_wait_started,
                loading=loading_page,
                now=loop_now,
                timeout=CAPTCHA_STATE_TIMEOUT,
            )

            submit_wait_started, _ = _update_submit_wait_state(

                submit_wait_started,

                transitioned=had_captcha,

                visible=visible,

                validating=validating,

                loading=loading_page,

            )
            if submit_wait_started is None and (had_captcha or visible or validating or loading_page):
                last_submit_click_hit = None

            if loading_page:

                if loading_timed_out:

                    waited = int(loop_now - (microsoft_loading_wait_started or loop_now))

                    log(f"  {tag} Microsoft Loading stuck for {waited}s, give up", "WARN")

                    _shot(page, "timeout_loading", idx)

                    return _finish(reason="timeout")

                if microsoft_loading_wait_started is None:

                    microsoft_loading_wait_started = loop_now

                    log(f"  {tag} Microsoft Loading page, keep waiting for redirect")

                elif int(loop_now - microsoft_loading_wait_started) % 15 < 3:

                    waited = int(loop_now - microsoft_loading_wait_started)

                    log(f"  {tag} Microsoft Loading still active, waited {waited}s")

                time.sleep(1)

                continue

            microsoft_loading_wait_started = None



            if _should_enter_post_press_reappear_wait(awaiting_reappear, press_count, max_press):

                gap_waited = time.time() - (post_press_started_at or time.time())

                if not actionable:

                    if _wait_state_timed_out(post_press_started_at, timeout=CAPTCHA_STATE_TIMEOUT):

                        log(

                            f"  {tag} captcha did not reappear/change after press for {int(gap_waited)}s, give up",

                            "WARN",

                        )

                        _shot(page, "timeout_captcha_reappear", idx)

                        return _finish(reason="timeout")

                    if not post_press_saw_gap:

                        post_press_saw_gap = True

                        if validating:

                            log(f"  {tag} press entered loading/validation, wait for captcha to reappear")

                        else:

                            log(f"  {tag} captcha not actionable after press, waiting for reappear")

                    elif gone_rounds % 10 == 0 and gone_rounds > 0:

                        waited = int(gap_waited)

                        log(f"  {tag} waiting for captcha reappear, {waited}s")

                    gone_rounds += 1

                    time.sleep(1.5)

                    continue

                if post_press_saw_gap:

                    awaiting_reappear = False

                    post_press_saw_gap = False

                    post_press_started_at = None

                    initial_press_wait_started = time.time() - INITIAL_PRESS_DELAY

                    gone_rounds = 0

                    _wait_before_next_captcha_press(tag, "challenge failed and captcha reappeared")

                    continue

                if gap_waited < POST_PRESS_LOADING_CHECK:

                    time.sleep(0.5)

                    continue

                awaiting_reappear = False

                post_press_saw_gap = False

                post_press_started_at = None

                initial_press_wait_started = time.time() - INITIAL_PRESS_DELAY

                gone_rounds = 0

                _wait_before_next_captcha_press(

                    tag,

                    f"challenge failed with no loading after {POST_PRESS_LOADING_CHECK}s",

                )

                continue



            if had_captcha and (not visible or validating or not actionable):

                if press_count >= max_press:

                    if press_wait_started is None:

                        press_wait_started = time.time()

                    waited = time.time() - press_wait_started

                    if waited >= POST_MAX_PRESS_WAIT:

                        log(f"  {tag} no redirect after {max_press} presses and {POST_MAX_PRESS_WAIT}s, give up", "WARN")

                        _shot(page, "press_fail", idx)

                        return _finish(reason="failure")

                    time.sleep(1)

                    continue

                gone_rounds += 1

                if validation_wait_started is None:

                    validation_wait_started = time.time()

                waited = time.time() - validation_wait_started

                if _wait_state_timed_out(validation_wait_started, timeout=CAPTCHA_STATE_TIMEOUT):

                    log(f"  {tag} captcha post-press state stuck for {int(waited)}s, give up", "WARN")

                    _shot(page, "timeout_captcha_state", idx)

                    return _finish(reason="timeout")

                if gone_rounds == 1:

                    if validating and visible:

                        log(f"  {tag} still validating after press, keep waiting")

                    else:

                        log(f"  {tag} captcha disappeared/loading, wait for redirect")

                elif gone_rounds % 10 == 0:

                    waited = int(time.time() - validation_wait_started)

                    if validating and visible:

                        log(f"  {tag} captcha still validating, waited {waited}s")

                    else:

                        log(f"  {tag} waiting for post-captcha redirect, {waited}s")

                time.sleep(1)

                continue



            if visible and actionable:

                validation_wait_started = None

                button_wait_started = None

                if not captcha_signup_url:

                    captcha_signup_url = current_url

                had_captcha = True

                last_progress_at = time.time()  # captcha 出现算进展

                gone_rounds = 0

                if press_count < max_press:

                    if initial_press_wait_started is None:

                        initial_press_wait_started = time.time()

                        log(f"  {tag} captcha visible, wait {INITIAL_PRESS_DELAY}s before press")

                        time.sleep(INITIAL_PRESS_DELAY)

                        continue

                    if time.time() - initial_press_wait_started < INITIAL_PRESS_DELAY:

                        time.sleep(1)

                        continue

                    ctx, target = hold_ctx, hold_target

                    if ctx is None or target is None or _target_quality(target) > 5:

                        if had_captcha and press_count > 0:

                            log(f"  {tag} captcha target not actionable, keep waiting")

                            awaiting_reappear = True

                            post_press_saw_gap = True

                            post_press_started_at = post_press_started_at or time.time()

                            time.sleep(1)

                            continue

                        no_target_rounds += 1

                        log(f"  {tag} captcha visible but no hold target", "WARN")

                        if no_target_rounds >= 5:

                            log(f"  {tag} captcha target missing for multiple rounds, give up", "WARN")

                            _shot(page, "captcha_no_target", idx)

                            return _finish(reason="failure")

                    else:

                        no_target_rounds = 0

                        press_count += 1

                        hold_elapsed = _perform_hold_with_px_screenshots(

                            page, ctx, target, idx, press_count, tag, enabled=px_press_screenshots

                        )

                        if hold_elapsed:

                            px_hold_elapsed += float(hold_elapsed)

                            validation_wait_started = time.time()

                            awaiting_reappear = True

                            post_press_saw_gap = False

                            post_press_started_at = time.time()

                            last_progress_at = time.time()  # 成功按压算进展

                            if press_count >= max_press:

                                press_wait_started = time.time()

                            continue

                        awaiting_reappear = False

                        post_press_saw_gap = False

                        post_press_started_at = None

                        initial_press_wait_started = time.time() - INITIAL_PRESS_DELAY

                        _wait_before_next_captcha_press(tag, "hold action failed")

                        continue

                if press_count >= max_press:

                    if press_wait_started is None:

                        press_wait_started = time.time()

                        log(f"  {tag} max press {max_press} reached, wait up to {POST_MAX_PRESS_WAIT}s")

                    waited = time.time() - press_wait_started

                    if waited >= POST_MAX_PRESS_WAIT:

                        log(f"  {tag} no redirect after {max_press} presses and {POST_MAX_PRESS_WAIT}s, give up", "WARN")

                        _shot(page, "press_fail", idx)

                        return _finish(reason="failure")

                    time.sleep(1)

                    continue

            else:

                if had_captcha:

                    # PX 已出现(iframe-box/文本命中)但真按钮还没渲染——等按钮，不按、不提交。

                    # 给较长宽限(按钮异步渲染最多 ~20s)，避免被下面的 validation_wait(20s)误杀。

                    if captcha_hint and not actionable and press_count == 0:

                        if button_wait_started is None:

                            button_wait_started = time.time()

                            log(f"  {tag} PX challenge loading, waiting for hold button to render")

                        elif _wait_state_timed_out(button_wait_started, timeout=BUTTON_RENDER_TIMEOUT):

                            waited = int(time.time() - button_wait_started)

                            log(f"  {tag} hold button did not render for {waited}s, give up", "WARN")

                            _shot(page, "timeout_button_render", idx)

                            return _finish(reason="timeout")

                        time.sleep(1)

                        continue

                    if press_count >= max_press:

                        if press_wait_started is None:

                            press_wait_started = time.time()

                        waited = time.time() - press_wait_started

                        if waited >= POST_MAX_PRESS_WAIT:

                            log(f"  {tag} no redirect after {max_press} presses and {POST_MAX_PRESS_WAIT}s, give up", "WARN")

                            _shot(page, "press_fail", idx)

                            return _finish(reason="failure")

                        time.sleep(1)

                        continue

                    if awaiting_reappear:

                        post_press_saw_gap = True

                        if _wait_state_timed_out(post_press_started_at, timeout=CAPTCHA_STATE_TIMEOUT):

                            waited = time.time() - (post_press_started_at or time.time())

                            log(

                                f"  {tag} captcha reappear wait stuck for {int(waited)}s, give up",

                                "WARN",

                            )

                            _shot(page, "timeout_captcha_reappear", idx)

                            return _finish(reason="timeout")

                        time.sleep(1)

                        continue

                    if validation_wait_started is None:

                        validation_wait_started = time.time()

                        log(f"  {tag} post-captcha state unclear, keep waiting")

                    elif _wait_state_timed_out(validation_wait_started, timeout=CAPTCHA_STATE_TIMEOUT):

                        waited = time.time() - validation_wait_started

                        log(f"  {tag} post-captcha state unclear for {int(waited)}s, give up", "WARN")

                        _shot(page, "timeout_captcha_state", idx)

                        return _finish(reason="timeout")

                    time.sleep(1)

                    continue

                validation_wait_started = None

                loop_now = time.time()

                submit_wait_started, submit_timed_out = _update_submit_wait_state(

                    submit_wait_started,

                    now=loop_now,

                )

                if submit_timed_out:

                    waited = int(loop_now - submit_wait_started)

                    log(f"  {tag} submit stuck for {waited}s without state change, give up", "WARN")

                    _shot(page, "submit_timeout", idx)

                    return _finish(reason="submit_timeout")

                submitted, last_submit_click_hit = _try_submit(page, tag, last_submit_click_hit)

                submit_wait_started, _ = _update_submit_wait_state(

                    submit_wait_started,

                    submitted=submitted,

                    now=time.time(),

                )

                # submit 卡死硬超时:点中 submit 后页面一直无进展(URL 没变、没出 captcha/loading/blocked),
                # 超过 SUBMIT_STUCK_NOCHANGE_TIMEOUT 就判卡死,早退 fail 换号。
                # URL 变化以新 URL 重新计时(防止页面在 signup 内中途跳转后卡死漏判)。
                # 注:若 URL 变成跳出 signup/loading,会先被上面状态分支 break 接住,这里只兜底卡死态。
                if submitted:
                    if submit_first_click_at is None or submit_first_click_url != current_url:
                        submit_first_click_at = time.time()
                        submit_first_click_url = current_url
                    elif not visible and not loading_page and not validating:
                        submit_nochange = time.time() - submit_first_click_at
                        if submit_nochange >= SUBMIT_STUCK_NOCHANGE_TIMEOUT:
                            log(f"  {tag} submit stuck for {int(submit_nochange)}s with no URL/captcha change, give up", "WARN")
                            _shot(page, "submit_stuck_nochange", idx)
                            return _finish(reason="submit_timeout")
                else:
                    submit_first_click_at = None
                    submit_first_click_url = None



            if int(time.time()) % 15 < 3:

                _shot(page, f"wait_{int(time.time() % 1000)}", idx)

            if not submitted:

                time.sleep(2)

        else:

            log(f"  {tag} captcha timeout", "WARN")

            _shot(page, "timeout", idx)

            return _finish(reason="timeout")

        step_timings["captcha"] = time.perf_counter() - captcha_started

        log(f"  {tag} step captcha: {step_timings['captcha']:.2f}s presses={press_count}", "INFO")



        _, step_timings["post_signup_cleanup"] = _timed_step(tag, "post_signup_cleanup", _post_signup_cleanup, page, tag, idx)

        verify_ok = True

        if need_verify:

            verify_ok, step_timings["verify_registered_outlook"] = _timed_step(

                tag, "verify_registered_outlook", verify_registered_outlook, email, password, tag

            )

        if need_verify and not verify_ok:

            log(f"  {tag} verification failed, discarding account", "WARN")

            return _finish(reason="verify_fail")



        timings_summary = ", ".join(f"{k}={v:.2f}s" for k, v in step_timings.items())

        log(f"  {tag} timings: {timings_summary}", "INFO")

        log(f"  {tag} OK: {email} / {password}", "OK")

        return _finish(email, password)

    except Exception as e:

        failure_reason = f"exception_{type(e).__name__}"

        log(f"  {tag} FAILED: {type(e).__name__}: {e}", "ERR")

        try:

            if page is not None:

                _shot(page, "error", idx)

        except Exception:

            pass

        return _finish(reason=failure_reason)

    finally:

        _stop_ruoyi_resource_blocking(page)

        if har_collector is not None:

            if capture_har:

                try:

                    har_collector.save(reason="success" if success else failure_reason)

                except Exception as exc:

                    log(f"  {tag} HAR save failed: {type(exc).__name__}: {exc}", "WARN")

            else:

                har_collector.stop()

        # 注册成功才入种子池:挑战 cookie 是"放行态"的凭证,失败实例的种子多半已烧穿,
        # 灌给下轮只会带指纹矛盾。成功 = PX 已放行,此时采集最干净。
        if success and page is not None and _warmup_seed_enabled:

            try:

                _ruyi_pkg.save_seed(
                    page, "outlook", selected_browser_proxy, user_agent, tag
                )

            except Exception as exc:

                log(f"  {tag} 种子入池失败: {type(exc).__name__}: {exc}", "WARN")

        if page is not None and browser_page is not None and page is not browser_page:

            try:

                page.close()

            except Exception:

                pass

        try:

            if browser_page is not None:

                _quit_browser_page(browser_page, tag, timeout=BROWSER_QUIT_TIMEOUT)

        except Exception:

            pass

        _untrack_browser_page(browser_page)

        # quit() 只 terminate 主进程,子进程树残留会占用 XPCOM/profile 文件锁,

        # 导致下个 run 启动撞 "Couldn't load XPCOM"。删 profile 前先清进程树并等退出。

        try:

            _kill_ruoyi_firefox_by_profile(

                profile_dir,

                timeout=RUOYI_FIREFOX_EXIT_WAIT,

                log_fn=lambda m, s: log(f"  {tag} {m}", s),

            )

        except Exception as exc:

            log(f"  {tag} 清 firefox 进程树失败: {type(exc).__name__}: {exc}", "WARN")

        _cleanup_ruoyi_run_profile_dir(profile_dir)

        clear_fn = getattr(helpers, "clear_account_generation_options", None)

        if callable(clear_fn):

            try:

                clear_fn()

            except Exception:

                pass





def _save_direct_result(email, password, graph, live_file, token_file):

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with _interprocess_lock(live_file):

        with open(live_file, "a", encoding="utf-8") as f:

            f.write(f"{email}----{password}----{graph.get('client_id', '')}----{graph['refresh_token']}----{graph.get('cf_address', '')}----{graph.get('cf_password') or ''}\n")

    if token_file:

        with _interprocess_lock(token_file):

            tokens = []

            if os.path.isfile(token_file):

                try:

                    with open(token_file, encoding="utf-8") as f:

                        tokens = json.load(f)

                except Exception:

                    tokens = []

            tokens.append(

                {

                    "email": email,

                    "password": password,

                    "client_id": graph.get("client_id"),

                    "refresh_token": graph.get("refresh_token"),

                    "secondary_email": graph.get("cf_address"),

                    "secondary_password": graph.get("cf_password"),

                }

            )

            with open(token_file, "w", encoding="utf-8") as f:

                json.dump(tokens, f, ensure_ascii=False, indent=2)

    append_graph_account_to_emails_pool(email, password, graph)





def _resolve_graph_auth_proxy(args_or_opts, reg_proxy, tag):

    """授权是否复用注册代理。

    返回传给 extract_graph_token_http 的 proxy_str（已归一化为 scheme://user:pass@host:port）。

    未开启选项或无代理时返回 None（直连，保持原行为）。
    """

    if not getattr(args_or_opts, "graph_auth_use_reg_proxy", False):

        return None

    if not reg_proxy:

        return None

    proxies = _proxy_for_ip_lookup([reg_proxy], tag)

    if not proxies:

        return None

    return proxies.get("https") or proxies.get("http")




def _build_bind_secondary(email, idx, tag, reg_proxy=None):
    """注册后授权前,给微软号建 cf 辅助邮箱,构造透传给 get_graph_token 的 bind_secondary。
    8月起微软对新号收紧:无辅助邮箱 Graph 授权一律 access_denied,必须先绑。
    复用 bind_secondary_email_http 的逻辑:create_or_get_address 建每个号专属 cf 邮箱
    (ms-<前缀>@<CF_MAIL_DOMAIN>),cf worker 走 CF_MAIL_PROXY 代理收信(MS 端点仍直连)。
    失败返回 None(降级走原 Skip 路径,不阻断授权,记 no_graph 让上层分类)。"""
    try:
        from common import cloudflare_mail as cm
    except Exception as e:
        log(f"  {tag} bind_secondary: import cloudflare_mail 失败: {type(e).__name__}: {e}", "WARN")
        return None
    try:
        cf_proxy = os.environ.get("CF_MAIL_PROXY") or None
        cm.set_proxy(cf_proxy)
        d = cm.create_or_get_address(email)
        bs = {
            "cf_address": d["address"],
            "cf_jwt": d.get("jwt"),
            "use_admin": d.get("use_admin", False),
            "cm": cm,
            "cf_password": d.get("password"),  # 地址密码(ENABLE_ADDRESS_PASSWORD 开启才有;None=未开启),透传给落盘第6字段
        }
        if bs["cf_password"]:
            log(f"  {tag} bind_secondary: cf 邮箱密码已生成 {d['address']} / {bs['cf_password']}")
        log(f"  {tag} bind_secondary: cf 辅助邮箱就绪 {d['address']} (use_admin={d.get('use_admin', False)})")
        return bs
    except Exception as e:
        log(f"  {tag} bind_secondary: 建 cf 辅助邮箱失败: {type(e).__name__}: {e}", "WARN")
        return None




async def _finish_direct_graph_auth(args, helpers, email, password, idx, save_lock, started, px_metrics, graph=None, reg_proxy=None):

    tag = f"#{idx}"

    if getattr(args, "skip_graph_auth", False):

        async with save_lock:

            await asyncio.to_thread(append_account_to_email_reg, email, password)

        total_elapsed = time.perf_counter() - started

        log(f"{tag} 跳过 Graph 授权(仅注册);saved to email_reg: {email}", "OK")

        log(f"{tag} 授权结果: SKIP(reg_only)", "OK")

        log(f"{tag} 结果: OK(reg_only) {email} total={total_elapsed:.2f}s", "OK")

        return "reg_only", total_elapsed, px_metrics

    if graph and graph.get("refresh_token"):

        log(f"{tag} Graph token extracting…", "INFO")

        log(f"{tag} graph fallback -> reuse", "INFO")

    else:

        log(f"{tag} Graph token extracting…", "INFO")

        auth_proxy_str = _resolve_graph_auth_proxy(args, reg_proxy, f"[{tag}]")

        if auth_proxy_str:

            log(f"{tag} graph proxy -> reg {mask_ruoyi_proxy(reg_proxy)}", "INFO")

        else:

            log(f"{tag} graph proxy -> direct", "INFO")

        # 8月起微软对新号收紧:无辅助邮箱一律 access_denied。注册后授权必须先绑 cf 辅助邮箱。
        # 这里复用 bind_secondary_email_http 的逻辑:create_or_get_address 建每个号专属 cf 邮箱,
        # 透传给 extract_graph_token_http -> get_graph_token 的 bind_secondary,proofs/Add 真绑再拿 RT。
        bind_secondary = _build_bind_secondary(email, idx, tag, reg_proxy)

        graph = await asyncio.to_thread(
            helpers.extract_graph_token_http, email, password, idx, 3, auth_proxy_str, bind_secondary
        )

    if not graph or not graph.get("refresh_token"):

        async with save_lock:

            await asyncio.to_thread(append_account_to_email_nograph, email, password)

            nograph_file = getattr(args, "nograph_file", None)

            if nograph_file:

                await asyncio.to_thread(_append_nograph_account, email, password, nograph_file)

        total_elapsed = time.perf_counter() - started

        log(f"{tag} 授权结果: FAIL(no_graph)", "WARN")

        log(f"{tag} registered but graph RT missing; saved to email_nograph: {email}", "WARN")

        log(f"{tag} 结果: OK(no_graph) {email} total={total_elapsed:.2f}s", "OK")

        return "no_graph", total_elapsed, px_metrics

    async with save_lock:

        await asyncio.to_thread(_save_direct_result, email, password, graph, args.live_file, args.token_file)

    total_elapsed = time.perf_counter() - started

    log(f"{tag} 授权结果: OK", "OK")

    log(f"{tag} 结果: OK {email} total={total_elapsed:.2f}s", "OK")

    return "ok", total_elapsed, px_metrics



async def _run_one_direct(args, helpers, proxy_pool, idx, total, save_lock, consumable_pool=None, ruoyi_slot=None, defer_graph_auth=False):

    """单号执行。返回 'ok' | 'no_graph' | 'fail'。"""

    tag = f"#{idx}"

    started = time.perf_counter()
    try:
        total_budget = max(0.0, float(getattr(args, "timeout", REGISTER_TIMEOUT) or REGISTER_TIMEOUT))
    except Exception:
        total_budget = float(REGISTER_TIMEOUT)
    overall_deadline = started + total_budget

    log(f"========== 注册 {tag}/{total} ==========")

    pool = consumable_pool if consumable_pool is not None else get_consumable_proxy_pool()

    selected_pool = select_proxy_for_account(proxy_pool, runtime=pool)

    selected_proxy = selected_pool[0] if selected_pool else None

    if selected_pool:

        log(f"{tag} 代理 -> {mask_ruoyi_proxy(selected_proxy)}")

        if isinstance(pool, ConsumableProxyPool):

            log(f"{tag} proxy list remaining={pool.remaining()}", "INFO")

    else:

        log(f"{tag} 无可用代理", "WARN")



    px_metrics = _normalize_px_metrics(idx)
    if selected_pool:
        precheck_timeout = min(PROXY_PRECHECK_TIMEOUT, max(0.1, overall_deadline - time.perf_counter()))
        if not await asyncio.to_thread(_probe_proxy_before_browser, selected_pool, f"[{tag}][ruoyi]", precheck_timeout):
            total_elapsed = time.perf_counter() - started
            log(f"{tag} result: FAIL(proxy_precheck_failed) total={total_elapsed:.2f}s", "WARN")
            px_metrics["reg_elapsed"] = total_elapsed
            release_proxy_for_account(selected_proxy, runtime=pool)
            return "fail", total_elapsed, px_metrics

    email = password = None

    fail_reason = "failure"

    fallback_graph = None

    try:

        remaining = max(0.0, overall_deadline - time.perf_counter())
        if remaining <= 0:
            raise TimeoutError(f"account timeout before browser ({total_budget:g}s)")
        run_args = SimpleNamespace(**vars(args))
        run_args.timeout = remaining
        if ruoyi_slot is not None:
            run_args.ruoyi_slot = ruoyi_slot
        result = await asyncio.to_thread(register_outlook, run_args, selected_pool, idx)

        if isinstance(result, tuple) and len(result) >= 4:

            email, password, fail_reason = result[0], result[1], (result[2] or "")

            px_metrics = _normalize_px_metrics(idx, result[3])

        elif isinstance(result, tuple) and len(result) >= 3:

            email, password, fail_reason = result[0], result[1], (result[2] or "")

        elif isinstance(result, tuple) and len(result) >= 2:

            email, password = result[0], result[1]

            fail_reason = "" if email else "failure"

        else:

            email = password = None

            fail_reason = "failure"

    except Exception as exc:

        log(f"{tag} register task raised {type(exc).__name__}: {exc}", "ERR")

        fail_reason = f"exception_{type(exc).__name__}"

    if not email:

        total_elapsed = time.perf_counter() - started

        log(f"{tag} 结果: FAIL({fail_reason or 'failure'}) total={total_elapsed:.2f}s", "WARN")

        release_proxy_for_account(selected_proxy, runtime=pool)

        return "fail", total_elapsed, px_metrics



    release_proxy_for_account(selected_proxy, runtime=pool)

    if defer_graph_auth:

        return {

            "status": "registered",

            "email": email,

            "password": password,

            "started": started,

            "px_metrics": px_metrics,

            "graph": fallback_graph,

            "reg_proxy": selected_proxy,

        }

    return await _finish_direct_graph_auth(args, helpers, email, password, idx, save_lock, started, px_metrics, graph=fallback_graph, reg_proxy=selected_proxy)





async def _run_direct_batch(args, helpers, consumable_pool):

    """返回 (ok, no_graph, fail, total_elapsed, avg_success_elapsed) 五元组。"""

    count = max(0, int(args.count or 0))

    concurrency = max(1, int(args.concurrency or 1))

    sem = asyncio.Semaphore(concurrency)

    save_lock = asyncio.Lock()
    slot_queue = asyncio.Queue()
    for slot_id in range(1, concurrency + 1):
        slot_queue.put_nowait(slot_id)

    batch_started = time.perf_counter()

    # 任务级可消耗代理池：start 已在外层完成，这里只绑定全局，结束 stop

    if not isinstance(consumable_pool, ConsumableProxyPool):

        consumable_pool = ConsumableProxyPool.from_args(args).start()

    set_consumable_proxy_pool(consumable_pool)

    st0 = consumable_pool.stats()

    log(f"proxy list ready: source={st0['source']} size={st0['remaining']}", "INFO")

    # 全局启动闸：保证任意两个 Firefox 启动至少错开 LAUNCH_STAGGER_SECONDS。

    # 4 并发时首波约 0/10/20/30s 依次点火，之后谁先腾 slot 谁按闸排队。

    launch_gate = asyncio.Lock()

    next_launch_at = [0.0]

    stagger = float(getattr(args, "launch_stagger", None) or LAUNCH_STAGGER_SECONDS or 0.0)

    if stagger < 0:

        stagger = 0.0

    ua_pool = _load_ua_pool()

    log(

        f"batch: count={count} concurrency={concurrency} "

        f"launch_stagger={stagger:g}s ua_pool={len(ua_pool)}"

    )



    async def runner(i):

        async with sem:
            slot_id = await slot_queue.get()
            try:

                # 启动错峰：拿到并发 slot 后还要等全局 launch_gate

                async with launch_gate:

                    now = time.monotonic()

                    wait = max(0.0, next_launch_at[0] - now)

                    if wait > 0:

                        log(f"#{i + 1} launch stagger wait {wait:.1f}s")

                        await asyncio.sleep(wait)

                    # 轻微抖动，避免整秒对齐

                    jitter = random.uniform(0.0, min(2.0, max(0.3, stagger * 0.15))) if stagger > 0 else random.uniform(0.2, 1.0)

                    if jitter > 0:

                        await asyncio.sleep(jitter)

                    next_launch_at[0] = time.monotonic() + stagger

                reg_result = await _run_one_direct(
                    args,
                    helpers,
                    consumable_pool,
                    i + 1,
                    count,
                    save_lock,
                    consumable_pool,
                    ruoyi_slot=slot_id,
                    defer_graph_auth=True,
                )
            finally:
                slot_queue.put_nowait(slot_id)

        if isinstance(reg_result, dict) and reg_result.get("status") == "registered":

            return await _finish_direct_graph_auth(
                args,
                helpers,
                reg_result["email"],
                reg_result["password"],
                i + 1,
                save_lock,
                reg_result["started"],
                reg_result["px_metrics"],
                graph=reg_result.get("graph"),
                reg_proxy=reg_result.get("reg_proxy"),
            )

        return reg_result



    try:

        results = await asyncio.gather(*(runner(i) for i in range(count)))

    finally:

        log(f"proxy list end: remaining={consumable_pool.remaining()}", "INFO")

        consumable_pool.stop()

        set_consumable_proxy_pool(None)

    statuses = [r[0] if isinstance(r, tuple) else r for r in results]

    elapsed_list = [r[1] if isinstance(r, tuple) and len(r) > 1 else 0.0 for r in results]

    px_stats = [_normalize_px_metrics(i + 1, r[2] if isinstance(r, tuple) and len(r) > 2 else None) for i, r in enumerate(results)]

    ok = sum(1 for r in statuses if r in ("ok", "reg_only"))

    reg_only = sum(1 for r in statuses if r == "reg_only")

    no_graph = sum(1 for r in statuses if r == "no_graph")

    fail = sum(1 for r in statuses if r == "fail")

    # 兼容旧布尔返回

    fail += sum(1 for r in statuses if r is False)

    ok += sum(1 for r in statuses if r is True)

    summary = _summarize_batch_metrics(statuses, elapsed_list, time.perf_counter() - batch_started, px_stats=px_stats)

    if reg_only:

        log(f"REG_ONLY: {reg_only} -> {EMAIL_REG}", "OK")

    return ok, no_graph, fail, summary["total_elapsed"], summary["avg_success_elapsed"], summary["px_stats"]





class _TeeStdout:
    """把 stdout 同时写到原 stdout 和日志文件；控制台/前端(SSE 读子进程 stdout)行为不变。"""

    def __init__(self, original, file_handle):
        self._orig = original
        self._file = file_handle
        self._lock = threading.Lock()
        self.encoding = getattr(original, "encoding", "utf-8") or "utf-8"
        self.errors = getattr(original, "errors", "replace")

    def write(self, s):
        try:
            self._orig.write(s)
        except Exception:
            pass
        if s:
            with self._lock:
                try:
                    self._file.write(s)
                    self._file.flush()
                except Exception:
                    pass
        return len(s) if s else 0

    def flush(self):
        try:
            self._orig.flush()
        except Exception:
            pass
        with self._lock:
            try:
                self._file.flush()
            except Exception:
                pass

    def isatty(self):
        try:
            return self._orig.isatty()
        except Exception:
            return False

    def fileno(self):
        return self._orig.fileno()

    def reconfigure(self, **kwargs):
        try:
            self._orig.reconfigure(**kwargs)
        except Exception:
            pass

    def __getattr__(self, name):
        return getattr(self._orig, name)


def _install_run_log_tee():
    """脚本启动时把本次运行日志 tee 到 logs/<YYYYMMDD_HHMMSS>.log。
    控制台与前端 UI 输出不变；失败仅告警，不阻断主流程。返回日志文件路径或 None。"""
    try:
        log_dir = os.path.join(ROOT, "logs")
        os.makedirs(log_dir, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = os.path.join(log_dir, f"{ts}.log")
        fh = open(path, "w", encoding="utf-8", buffering=1)
        sys.stdout = _TeeStdout(sys.stdout, fh)
        atexit.register(lambda: fh.close() if not fh.closed else None)
        sys.stdout.write(f"===== 本次运行日志 {ts} =====\n")
        return path
    except Exception as exc:
        try:
            print(f"_install_run_log_tee failed: {type(exc).__name__}: {exc}", file=sys.stderr)
        except Exception:
            pass
        return None


def main():

    _install_run_log_tee()

    ap = argparse.ArgumentParser(description="Outlook 自注册养号(ruoyi) — 完整链路版")

    ap.add_argument("--proxy-file", "-p", default=PROXY_FILE, help=f"代理池文件(默认 {PROXY_FILE})")

    ap.add_argument("--proxy-source", default=RUOYI_PROXY_SOURCE,

                    choices=["file", "http"],

                    help="ruoyi 代理来源：file=本地代理文件；http=HTTP GET 拉取 txt 代理列表。启动装 list，注册随机取删，空则重载")

    ap.add_argument("--proxy-url", default=PROXY_URL,

                    help="HTTP GET 代理列表地址，返回 txt；每行一个代理")

    ap.add_argument("--count", "-n", type=int, default=1, help="注册次数(默认 1)")

    ap.add_argument("--concurrency", "-c", type=int, default=1, help="并发注册数(默认 1)")

    ap.add_argument(

        "--launch-stagger",

        type=float,

        default=LAUNCH_STAGGER_SECONDS,

        help=f"并发启动错峰秒数(默认 {LAUNCH_STAGGER_SECONDS:g}；0=关闭；4并发建议 8~15)",

    )

    ap.add_argument(

        "--ua-pool",

        default=os.environ.get("OUTLOOK_RUOYI_UA_POOL", ""),

        help="UA 池，| 或换行分隔；空=内置 6 条 Firefox 轮询",

    )

    ap.add_argument("--headless", action="store_true", help="无头模式")

    ap.add_argument("--block-resources", action="store_true",

                    default=_env_bool("OUTLOOK_RUOYI_BLOCK_RESOURCES", False),

                    help="屏蔽 image/font/media 资源请求(遥测域名无论是否勾选都会屏蔽)")

    ap.add_argument("--timeout", "-t", type=int, default=REGISTER_TIMEOUT, help="单号超时(秒)")

    ap.add_argument("--max-press", default=os.environ.get("OUTLOOK_REG_MAX_PRESS", "5"), help="按住次数上限")

    ap.add_argument("--no-verify", action="store_true", help="跳过 Outlook 登录校验")

    ap.add_argument("--graph-auth-use-reg-proxy", action="store_true",

                    default=_env_bool("OUTLOOK_RUOYI_GRAPH_AUTH_REG_PROXY", False),

                    help="Graph 授权复用注册代理；默认关闭(直连授权)，开启后按当前账号注册代理走授权")

    ap.add_argument("--skip-graph-auth", action="store_true",

                    default=_env_bool("OUTLOOK_RUOYI_SKIP_GRAPH_AUTH", False),

                    help="只注册不授权 Graph；注册成功的号追加到 email_reg.txt(默认关闭，保持注册后授权)")

    ap.add_argument("--confirm-before-register", action="store_true", help="页面打开后先尝试点确认")

    ap.add_argument("--px-press-screenshots", action=argparse.BooleanOptionalAction,

                    default=_env_bool("OUTLOOK_PX_PRESS_SCREENSHOTS", False),

                    help="保存 ruoyi PX/失败相关截图")

    ap.add_argument("--har", action="store_true", default=_env_bool("OUTLOOK_RUOYI_HAR", False),

                    help="保存完整链路 HAR；开启后成功/失败都会保存，默认关闭")

    ap.add_argument("--log-level", default=os.environ.get("OUTLOOK_LOG_LEVEL", "INFO"),

                    choices=["DEBUG", "INFO", "WARN", "PROD", "ERR"],

                    help="????")

    ap.add_argument(

        "--email-suffixes",

        default=os.environ.get("OUTLOOK_ACCOUNT_SUFFIXES") or os.environ.get("OUTLOOK_EMAIL_SUFFIXES") or "outlook.com",

        help="邮箱后缀池，逗号/空格分隔，如 outlook.com,hotmail.com",

    )

    ap.add_argument(

        "--account-format-mode",

        default=(

            os.environ.get("OUTLOOK_ACCOUNT_FORMAT_MODE")

            or "custom"

        ),

        choices=["random", "custom"],

        help="账号格式：random=纯随机,custom=用 --account-format 模板",

    )

    ap.add_argument(

        "--account-format",

        default=os.environ.get("OUTLOOK_ACCOUNT_FORMAT", "") or "{letters:7}{digits:6}",

        help="邮箱前缀模板。{letters:7}{digits:6}=7随机字母+6数字(默认,约8e15种,taken几乎为0);custom 模式下填固定内容(如 myname123)则该内容直接做邮箱前缀不展开模板",

    )

    ap.add_argument(

        "--password-format",

        default=os.environ.get("OUTLOOK_PASSWORD_FORMAT", ""),

        help="密码模板，如 Aa1!{rand:12}；留空用默认随机",

    )

    ap.add_argument("--loop", action="store_true",

                    help="循环养号：跑完一批后等待 --loop-interval 秒继续下一批，直到被停止(Ctrl-C 或 webui 停止)")

    ap.add_argument("--loop-interval", type=int,

                    default=int(os.environ.get("OUTLOOK_LOOP_INTERVAL", "300") or "300"),

                    help="循环养号两批之间的等待秒数(默认 300=5 分钟)")

    ap.add_argument("--tg-bot-token", default=os.environ.get("TG_BOT_TOKEN", ""),

                    help="Telegram bot token，配置后每批完成把汇总发到 TG；留空不发")

    ap.add_argument("--tg-chat-id", default=os.environ.get("TG_CHAT_ID", ""),

                    help="Telegram chat id，与 --tg-bot-token 配合使用")

    ap.add_argument("--tg-proxy", default=os.environ.get("TG_PROXY", ""),

                    help="Telegram 走的 HTTP 代理，如 http://127.0.0.1:7897；国内网络必填，否则 api.telegram.org 直连不通")

    args = ap.parse_args()

    set_log_level(args.log_level)

    # CLI 覆盖环境，保证 _load_ua_pool / batch stagger 读到最新值

    if getattr(args, "ua_pool", None):

        os.environ["OUTLOOK_RUOYI_UA_POOL"] = str(args.ua_pool)

    try:

        stagger_val = float(getattr(args, "launch_stagger", None) or LAUNCH_STAGGER_SECONDS or 0.0)

    except Exception:

        stagger_val = float(LAUNCH_STAGGER_SECONDS or 0.0)

    if stagger_val < 0:

        stagger_val = 0.0

    args.launch_stagger = stagger_val

    os.environ["OUTLOOK_RUOYI_LAUNCH_STAGGER"] = str(stagger_val)



    if not os.path.isfile(RUOYI_FIREFOX_PATH):

        log(f"定制 Firefox 内核不存在: {RUOYI_FIREFOX_PATH}", "ERR")

        log("请先运行: .venv\\Scripts\\python.exe -m ruyipage install", "ERR")

        sys.exit(1)



    helpers = _load_helpers()

    _apply_account_options(args)



    consumable_pool = ConsumableProxyPool.from_args(args).start()

    st = consumable_pool.stats()

    log(f"代理 list 就绪: source={st['source']} size={st['remaining']}")



    loop_mode = bool(getattr(args, "loop", False))

    loop_interval = max(0, int(getattr(args, "loop_interval", 300) or 0))

    if loop_mode:

        log(f"循环养号模式: count={args.count} 间隔 {loop_interval}s，Ctrl-C 或 webui 停止可中断", "OK")

        batch_no = 0

        try:

            while True:

                batch_no += 1

                log(f"===== 循环第 {batch_no} 批开始 =====", "OK")

                summary_lines = _run_one_batch(args, helpers, consumable_pool, stagger_val)

                _notify_tg(args, summary_lines, batch_no=batch_no)

                log(f"第 {batch_no} 批完成，等待 {loop_interval}s 后继续下一批…", "OK")

                if loop_interval > 0:

                    time.sleep(loop_interval)

        except KeyboardInterrupt:

            log(f"循环养号已停止(共 {batch_no} 批)", "OK")

    else:

        summary_lines = _run_one_batch(args, helpers, consumable_pool, stagger_val)

        _notify_tg(args, summary_lines)



def _run_one_batch(args, helpers, consumable_pool, stagger_val):

    """跑一批注册并打印汇总；返回汇总行列表(供 TG 通知复用)。"""

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    _force_kill_ruoyi_firefox()  # 批次开始前强杀上批残留 Firefox(防 XPCOM 加载失败)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    live_file = os.path.join(OUTPUT_DIR, f"accounts_ruoyi_{ts}.txt")

    token_file = os.path.join(OUTPUT_DIR, f"graph_tokens_ruoyi_{ts}.json")

    nograph_file = os.path.join(OUTPUT_DIR, f"accounts_ruoyi_nograph_{ts}.txt")

    args.live_file = live_file

    args.token_file = token_file

    args.nograph_file = nograph_file

    log(

        f"开始: count={args.count} concurrency={max(1, int(args.concurrency or 1))} "

        f"launch_stagger={stagger_val:g}s ua_pool={len(_load_ua_pool())} timeout={args.timeout}s"

    )

    ok, no_graph, failed, batch_total_elapsed, avg_success_elapsed, px_stats = asyncio.run(_run_direct_batch(args, helpers, consumable_pool))

    total = max(0, int(args.count or 0))

    # 多行汇总：WebUI 正则 + 人眼可读中文都覆盖

    summary_lines = _format_batch_summary_lines(ok, no_graph, failed, total, batch_total_elapsed, avg_success_elapsed, px_stats)

    for line in summary_lines:

        log(line, "OK")

    if os.path.isfile(live_file):

        log(f"账号输出: {live_file}")

    if os.path.isfile(token_file):

        log(f"Token 输出: {token_file}")

    if os.path.isfile(nograph_file):

        log(f"未授权账号输出: {nograph_file}")

    if no_graph and os.path.isfile(EMAIL_NOGRAPH):

        log(f"未授权输出: {EMAIL_NOGRAPH}")

    log(f"email_nograph: {EMAIL_NOGRAPH}")

    if os.path.isfile(EMAIL_REG):

        log(f"仅注册(跳过授权)输出: {EMAIL_REG}")

    log(f"email_reg: {EMAIL_REG}")

    _force_kill_ruoyi_firefox()  # 批次结束强杀本批残留 Firefox(防累积导致 XPCOM)

    return summary_lines



if __name__ == "__main__":

    main()

