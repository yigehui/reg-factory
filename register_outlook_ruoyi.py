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

from contextlib import contextmanager

import importlib.util

import json

import os

import random

import re

import requests

import signal

import sys

import threading

import time

import urllib.parse

import urllib.request

from datetime import datetime

from types import SimpleNamespace

from urllib.parse import quote



DEFAULT_RUOYI_FIREFOX = (

    r"C:\Users\Administrator\AppData\Local\ruyipage\browsers"

    r"\firefox-151.0a1-151-ruyi-win64\firefox\firefox.exe"

)

RUOYI_FIREFOX_PATH = os.environ.get("RUOYI_FIREFOX_PATH", DEFAULT_RUOYI_FIREFOX)



ROOT = os.path.dirname(os.path.abspath(__file__))

ENGINE_NAME = "ruoyi"

PROXY_FILE = os.environ.get("OUTLOOK_PROXY_FILE", "proxies_outlook.txt")

RUOYI_PROXY_SOURCE = os.environ.get("OUTLOOK_RUOYI_PROXY_SOURCE", "file")

AIMILI_POOL_URL = os.environ.get("OUTLOOK_AIMILI_POOL_URL") or os.environ.get("OUTLOOK_AIMILI_POOL_BASE_URL", "")

AIMILI_POOL_TOKEN = os.environ.get("OUTLOOK_AIMILI_POOL_TOKEN", "")

SCREENSHOT_DIR = os.path.join(ROOT, "screenshots_ruoyi")

HAR_DIR = os.path.join(ROOT, "har_ruoyi")

OUTPUT_DIR = os.path.join(ROOT, "outlook_accounts")

EMAIL_NOGRAPH = os.path.join(OUTPUT_DIR, "email_nograph.txt")

EMAILS_POOL = os.path.join(ROOT, "emails.txt")

SIGNUP_URL = "https://signup.live.com/signup?lic=1"

IP_INFO_ENDPOINTS = [

    ("ipwhois", "https://ipwho.is/"),

]

_CURRENT_IP_INFO = {}

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

PROXY_PRECHECK_TIMEOUT = 30

PROXY_PRECHECK_URL = SIGNUP_URL

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
MICROSOFT_LOADING_TIMEOUT = float(os.environ.get("OUTLOOK_RUOYI_MICROSOFT_LOADING_TIMEOUT", "30") or "30")

BIRTHDAY_ENTRY_TIMEOUT = float(os.environ.get("OUTLOOK_RUOYI_BDAY_ENTRY_TIMEOUT", "2.5") or "2.5")

BIRTHDAY_SUBMIT_TIMEOUT = float(os.environ.get("OUTLOOK_RUOYI_BDAY_SUBMIT_TIMEOUT", "2.0") or "2.0")

BROWSER_QUIT_TIMEOUT = float(os.environ.get("OUTLOOK_RUOYI_BROWSER_QUIT_TIMEOUT", "5") or "5")

# Retry gap after a failed challenge before the next press.

POST_PRESS_RETRY_GAP_MIN = 2.0

POST_PRESS_RETRY_GAP_MAX = 4.0

# Slow down page start; actual submit now waits once with a random 0-max delay.

PAGE_START_DELAY = float(os.environ.get("OUTLOOK_RUOYI_PAGE_START_DELAY", "2") or "2")

SUBMIT_DELAY = float(os.environ.get("OUTLOOK_RUOYI_SUBMIT_DELAY", "0.5") or "0.5")

HEADLESS_WINDOW_WIDTH = int(os.environ.get("OUTLOOK_RUOYI_HEADLESS_WIDTH", "1280") or "1280")

HEADLESS_WINDOW_HEIGHT = int(os.environ.get("OUTLOOK_RUOYI_HEADLESS_HEIGHT", "800") or "800")

# 并发启动错峰：同一批拿到 slot 后，相邻两个浏览器启动至少间隔这么多秒。

# 4 并发默认 10s → 约 0/10/20/30s 错峰拉满，避免四窗同时砸 signup。

LAUNCH_STAGGER_SECONDS = float(os.environ.get("OUTLOOK_RUOYI_LAUNCH_STAGGER", "10") or "10")

# 内置 Firefox UA 池（Win10 x64，版本轮换）。可用 OUTLOOK_RUOYI_UA_POOL 覆盖（| 或换行分隔）。

_DEFAULT_UA_POOL = (

    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:151.0) Gecko/20100101 Firefox/151.0",

    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:150.0) Gecko/20100101 Firefox/150.0",

    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:149.0) Gecko/20100101 Firefox/149.0",

    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:148.0) Gecko/20100101 Firefox/148.0",

    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:147.0) Gecko/20100101 Firefox/147.0",

    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:146.0) Gecko/20100101 Firefox/146.0",

)

HEADLESS_USER_AGENT = os.environ.get("OUTLOOK_RUOYI_HEADLESS_UA", _DEFAULT_UA_POOL[0])

_UA_RR_LOCK = threading.Lock()

_UA_RR_IDX = 0

_HELPERS = None

_ACTIVE_BROWSER_PAGES = {}

_ACTIVE_BROWSER_LOCK = threading.Lock()

_SHUTDOWN_HANDLERS_INSTALLED = False





def _load_ua_pool():

    """UA 池：环境变量 OUTLOOK_RUOYI_UA_POOL 优先，否则内置 6 条 Firefox。"""

    raw = str(os.environ.get("OUTLOOK_RUOYI_UA_POOL", "") or "").strip()

    pool = []

    if raw:

        for part in re.split(r"[\n|]+", raw):

            ua = part.strip()

            if ua:

                pool.append(ua)

    if not pool:

        # 单条 HEADLESS_UA 放队首，再拼内置池去重

        seed = str(HEADLESS_USER_AGENT or "").strip()

        seen = set()

        for ua in (([seed] if seed else []) + list(_DEFAULT_UA_POOL)):

            if ua and ua not in seen:

                seen.add(ua)

                pool.append(ua)

    return pool or list(_DEFAULT_UA_POOL)





def _track_browser_page(page):

    if page is None:

        return

    with _ACTIVE_BROWSER_LOCK:

        _ACTIVE_BROWSER_PAGES[id(page)] = page





def _untrack_browser_page(page):

    if page is None:

        return

    with _ACTIVE_BROWSER_LOCK:

        _ACTIVE_BROWSER_PAGES.pop(id(page), None)





def _close_tracked_browser_pages():

    with _ACTIVE_BROWSER_LOCK:

        pages = list(_ACTIVE_BROWSER_PAGES.values())

        _ACTIVE_BROWSER_PAGES.clear()

    for browser_page in pages:

        try:

            browser_page.quit()

            continue

        except Exception:

            pass

        try:

            browser_page.close()

        except Exception:

            pass





def _quit_browser_page(browser_page, tag="", timeout=BROWSER_QUIT_TIMEOUT):

    if browser_page is None:

        return True

    done = threading.Event()



    def _worker():

        try:

            browser_page.quit()

        except Exception:

            pass

        finally:

            done.set()



    thread = threading.Thread(target=_worker, daemon=True)

    thread.start()

    if done.wait(max(0.01, float(timeout or 0.0))):

        return True

    try:

        browser_page.close()

    except Exception:

        pass

    if tag:

        log(f"  {tag} browser quit timed out after {float(timeout):.1f}s; continue closing", "WARN")

    return False





def _install_shutdown_handlers():

    global _SHUTDOWN_HANDLERS_INSTALLED

    if _SHUTDOWN_HANDLERS_INSTALLED:

        return



    def _handle_shutdown(signum, _frame):

        signame = str(signum)

        try:

            signame = signal.Signals(signum).name

        except Exception:

            pass

        try:

            log(f"收到 {signame}，先关闭 ruyi 浏览器再退出", "WARN")

        except Exception:

            pass

        _close_tracked_browser_pages()

        raise SystemExit(0)



    for sig_name in ("SIGINT", "SIGTERM", "SIGBREAK"):

        sig = getattr(signal, sig_name, None)

        if sig is None:

            continue

        try:

            signal.signal(sig, _handle_shutdown)

        except Exception:

            continue

    _SHUTDOWN_HANDLERS_INSTALLED = True





def _pick_user_agent(idx=None):

    """按任务序号轮询 UA；idx 为空时线程安全全局自增。"""

    pool = _load_ua_pool()

    if not pool:

        return HEADLESS_USER_AGENT

    if idx is not None:

        try:

            n = int(idx)

        except Exception:

            n = 1

        return pool[(max(1, n) - 1) % len(pool)]

    global _UA_RR_IDX

    with _UA_RR_LOCK:

        ua = pool[_UA_RR_IDX % len(pool)]

        _UA_RR_IDX += 1

        return ua





def _mask_ua(ua):

    raw = str(ua or "")

    m = re.search(r"Firefox/([\d.]+)", raw)

    ver = m.group(1) if m else "?"

    return f"Firefox/{ver}"





def _env_bool(name, default=False):

    value = os.environ.get(name)

    if value is None:

        return bool(default)

    return str(value).strip().lower() not in ("0", "false", "no", "off", "")





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

    success_statuses = {"ok", "no_graph", True}

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





def _normalize_px_metrics(idx, metrics=None):

    data = dict(metrics or {})

    return {

        "idx": int(data.get("idx") or idx or 0),

        "max_presses": int(data.get("max_presses") or 0),

        "px_elapsed": float(data.get("px_elapsed") or 0.0),
        "reg_elapsed": float(data.get("reg_elapsed") or 0.0),

    }





set_log_level(os.environ.get("OUTLOOK_LOG_LEVEL", "INFO"))





def _browser_model_name(browser_path):

    parts = [p for p in os.path.normpath(str(browser_path or "")).split(os.sep) if p]

    for part in reversed(parts):

        low = part.lower()

        if low.startswith("firefox-") or "ruyi" in low:

            return part

    return os.path.basename(str(browser_path or "")) or "unknown"





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





def _strip_proxy_scheme(value):

    s = str(value or "").strip()

    for pfx in ("socks5h://", "socks5://", "socks4://", "http://", "https://"):

        if s.lower().startswith(pfx):

            return s[len(pfx):]

    return s





def _proxy_url_to_ruoyi(value):

    """Normalize proxy URL / user:pass@host:port / host:port to ruyipage format."""

    raw = str(value or "").strip()

    if not raw:

        return ""

    if "://" in raw:

        try:

            parsed = urllib.parse.urlsplit(raw)

            host = parsed.hostname or ""

            port = parsed.port or 0

            user = urllib.parse.unquote(parsed.username or "")

            pwd = urllib.parse.unquote(parsed.password or "")

            if host and port:

                return f"{host}:{port}:{user}:{pwd}" if (user or pwd) else f"{host}:{port}"

        except Exception:

            pass

    s = _strip_proxy_scheme(raw).replace(",", "@", 1) if "@" not in raw and "," in raw else _strip_proxy_scheme(raw)

    if "@" in s:

        auth, hostport = s.rsplit("@", 1)

        if ":" in auth and ":" in hostport:

            user, pwd = auth.split(":", 1)

            host, port = hostport.rsplit(":", 1)

            return f"{host}:{port}:{user}:{pwd}"

    parts = s.split(":")

    if len(parts) == 4:

        return s

    if len(parts) == 2 and parts[1].isdigit():

        return s

    return ""





def _parse_ruoyi_proxy(proxy_str):

    s = str(proxy_str or "").strip()

    if not s:

        return None

    parts = s.split(":")

    if len(parts) >= 4:

        host, port = parts[0], parts[1]

        user = parts[2]

        pwd = ":".join(parts[3:])

        return {"host": host, "port": port, "username": user, "password": pwd}

    if len(parts) == 2 and parts[1].isdigit():

        return {"host": parts[0], "port": parts[1], "username": "", "password": ""}

    normalized = _proxy_url_to_ruoyi(s)

    if normalized and normalized != s:

        return _parse_ruoyi_proxy(normalized)

    return None





def mask_ruoyi_proxy(proxy_str):

    p = _parse_ruoyi_proxy(proxy_str)

    if not p:

        return "***" if proxy_str else "noproxy"

    auth = f"{p['username'][:8]}...@" if p.get("username") else ""

    return f"socks5://{auth}{p['host']}:{p['port']}"





def parse_proxy_pool(path):

    """读本地代理文件，支持 URL、user:pass@host:port、host:port。"""

    if not path:

        return []

    if not os.path.isfile(path):

        log(f"代理池文件不存在: {path}", "WARN")

        return []

    out = []

    with open(path, "r", encoding="utf-8") as f:

        for ln in f:

            ln = ln.strip()

            if not ln or ln.startswith("#"):

                continue

            normalized = _proxy_url_to_ruoyi(ln)

            if not normalized:

                log(f"跳过非法代理行: {ln[:80]}", "WARN")

                continue

            out.append(normalized)

    return out





def _aimili_endpoint(pool_url, source):

    raw = str(pool_url or "").strip().rstrip("/")

    if not raw:

        raise RuntimeError("OUTLOOK_AIMILI_POOL_URL/--aimili-url 为空")

    source = str(source or "aimili-list").replace("-", "_")

    parsed = urllib.parse.urlsplit(raw)

    path = parsed.path.rstrip("/")

    if path.startswith("/api/pool"):

        if source == "aimili_random" and path.endswith("/proxies"):

            path = path + "/random"

        elif source == "aimili_list" and path.endswith("/proxies/random"):

            path = path[: -len("/random")]

        return urllib.parse.urlunsplit((parsed.scheme, parsed.netloc, path, parsed.query, parsed.fragment))

    suffix = "/api/pool/proxies/random" if source == "aimili_random" else "/api/pool/proxies"

    return raw + suffix





def _aimili_api_json(pool_url, token, source, timeout=12):

    url = _aimili_endpoint(pool_url, source)

    req = urllib.request.Request(url, headers={"Accept": "application/json"})

    if token:

        req.add_header("Authorization", f"Bearer {token}")

    with urllib.request.urlopen(req, timeout=timeout) as resp:

        data = resp.read().decode("utf-8", errors="replace")

    return json.loads(data)





def _aimili_proxy_entry_to_ruoyi(item):

    if not isinstance(item, dict):

        return ""

    for key in ("socks5", "http"):

        value = item.get(key)

        if value:

            normalized = _proxy_url_to_ruoyi(value)

            if normalized:

                return normalized

    host = item.get("host") or item.get("public_host")

    port = item.get("port")

    if host and port:

        user = item.get("username") or ""

        pwd = item.get("password") or ""

        return f"{host}:{port}:{user}:{pwd}" if (user or pwd) else f"{host}:{port}"

    return ""





def fetch_aimili_proxy_list(pool_url, token=""):

    data = _aimili_api_json(pool_url, token, "aimili_list")

    proxies = data.get("proxies") if isinstance(data, dict) else None

    out = []

    for item in proxies or []:

        proxy = _aimili_proxy_entry_to_ruoyi(item)

        if proxy:

            out.append(proxy)

    return out





def fetch_aimili_proxy_random(pool_url, token=""):

    """Aimili random 接口：返回 0~1 条。"""

    data = _aimili_api_json(pool_url, token, "aimili_random")

    item = data.get("proxy") if isinstance(data, dict) else None

    proxy = _aimili_proxy_entry_to_ruoyi(item)

    if not proxy:

        raise RuntimeError(f"Aimili random API 未返回可用代理: {data}")

    return [proxy]





def _proxy_source_label(args):

    has_file = bool(str(getattr(args, "proxy_file", "") or PROXY_FILE).strip())

    has_aimili = bool(str(getattr(args, "aimili_url", "") or getattr(args, "aimili_base_url", "") or AIMILI_POOL_URL).strip())

    if has_file and has_aimili:

        return "file+aimili-list"

    if has_aimili:

        return "aimili-list"

    if has_file:

        return "file"

    return "empty"





def load_proxy_list(args):

    """统一代理 list：本地文件 + Aimili 列表接口一起加载。"""

    out = parse_proxy_pool(getattr(args, "proxy_file", "") or PROXY_FILE)

    pool_url = getattr(args, "aimili_url", "") or getattr(args, "aimili_base_url", "") or AIMILI_POOL_URL

    token = getattr(args, "aimili_token", "") or AIMILI_POOL_TOKEN

    if str(pool_url or "").strip():

        out.extend(fetch_aimili_proxy_list(pool_url, token))

    return out





# 兼容旧名

load_proxy_batch = load_proxy_list





def build_proxy_source(args):

    """兼容旧调用：返回当前来源的一批代理 list。"""

    return load_proxy_list(args)





class ConsumableProxyPool:

    """任务级代理 list（就一个 list，不是消息队列）。



    启动 load → 注册 pop 一条 → list 空了再 load → 任务停 clear。

    加锁只是为了并发注册不抢同一条。

    """



    def __init__(self, source_args):

        self.source_args = source_args

        self.source = _proxy_source_label(source_args)

        self._lock = threading.Lock()

        self._list = []



    @classmethod

    def from_args(cls, args):

        source_args = SimpleNamespace(

            proxy_file=getattr(args, "proxy_file", "") or PROXY_FILE,

            aimili_url=getattr(args, "aimili_url", "") or getattr(args, "aimili_base_url", "") or AIMILI_POOL_URL,

            aimili_token=getattr(args, "aimili_token", "") or AIMILI_POOL_TOKEN,

        )

        return cls(source_args)



    def __bool__(self):

        with self._lock:

            return bool(self._list)



    def remaining(self):

        with self._lock:

            return len(self._list)



    def stats(self):

        with self._lock:

            return {"remaining": len(self._list), "source": self.source}



    def _load_locked(self):

        """文件 load / 接口重调，结果塞进 self._list。"""

        try:

            batch = load_proxy_list(self.source_args) or []

        except Exception as exc:

            log(f"proxy load failed ({self.source}): {type(exc).__name__}: {exc}", "WARN")

            batch = []

        # 同批 host:port 去重

        seen = set()

        fresh = []

        for p in batch:

            p = str(p or "").strip()

            if not p:

                continue

            parsed = _parse_ruoyi_proxy(p)

            key = f"{parsed['host']}:{parsed['port']}" if parsed and parsed.get("host") and parsed.get("port") else p

            if key in seen:

                continue

            seen.add(key)

            fresh.append(p)

        self._list.extend(fresh)

        log(

            f"proxy load source={self.source} got={len(fresh)} list={len(self._list)}",

            "INFO" if fresh else "WARN",

        )

        return len(fresh)



    def start(self):

        """任务启动：初始化 list。"""

        with self._lock:

            self._list = []

            self._load_locked()

            if not self._list:

                log(f"proxy list empty after start (source={self.source})", "WARN")

            return self



    def reload(self):

        """list 空时重载：文件再读 / 接口再调。"""

        with self._lock:

            return self._load_locked()



    def take(self):

        """按当前 list size 随机下标取一条并删除；list 空则重新 load。"""

        with self._lock:

            if not self._list:

                self._load_locked()

            if not self._list:

                return []

            idx = random.randrange(len(self._list))

            proxy = self._list.pop(idx)

            log(f"proxy take -> {mask_ruoyi_proxy(proxy)} remaining={len(self._list)}", "DEBUG")

            return [proxy]



    def stop(self):

        """任务结束：销毁 list。"""

        with self._lock:

            n = len(self._list)

            self._list = []

            log(f"proxy list destroyed (cleared {n})", "INFO")





_PROXY_LIST = None





def get_consumable_proxy_pool():

    return _PROXY_LIST





def set_consumable_proxy_pool(pool):

    """任务启动绑定 list，结束传 None。"""

    global _PROXY_LIST

    _PROXY_LIST = pool

    return _PROXY_LIST





# 兼容旧名

def get_session_proxy_runtime():

    return _PROXY_LIST





def set_session_proxy_runtime(runtime):

    return set_consumable_proxy_pool(runtime)





SessionProxyRuntime = ConsumableProxyPool





def select_proxy_for_account(proxy_pool=None, runtime=None):

    """注册前从 list 取一条（取后删除）。"""

    pool = runtime if runtime is not None else None

    if pool is None and isinstance(proxy_pool, ConsumableProxyPool):

        pool = proxy_pool

    if pool is None:

        pool = _PROXY_LIST

    if isinstance(pool, ConsumableProxyPool):

        return pool.take()

    if not proxy_pool:

        return []

    items = list(proxy_pool)

    if not items:

        return []

    return [items.pop(random.randrange(len(items)))]





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

                f.write(f"{email}----{password}----{token}----{client_id}\n")

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





def _webui_task_store_enabled():

    return bool(

        os.environ.get("WEBUI_TASK_RUN_ID")

        and os.environ.get("WEBUI_MYSQL_HOST")

        and os.environ.get("WEBUI_MYSQL_USER")

        and os.environ.get("WEBUI_MYSQL_DATABASE")

    )



def _record_webui_account(email, password, graph=None, status="ok"):

    if not _webui_task_store_enabled() or not email:

        return False

    try:

        import task_store as _task_store

        _task_store.add_task_account(

            run_id=os.environ.get("WEBUI_TASK_RUN_ID") or "",

            email=email,

            password=password or "",

            client_id=(graph or {}).get("client_id") or "",

            refresh_token=(graph or {}).get("refresh_token") or "",

            generated_at=datetime.now().isoformat(),

            register_ip="",

            register_region="",

            status=status,

            source="runtime",

        )

        return True

    except Exception as exc:

        log(f"task_store account save failed: {type(exc).__name__}: {exc}", "WARN")

        return False



def _save_no_graph_result(email, password):

    append_account_to_email_nograph(email, password)

    _record_webui_account(email, password, None, "no_graph")



def _shot(page, name, idx):

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





def _save_screenshot(page, name, idx, tag=None):

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

        listen = getattr(self.page, "listen", None)

        if listen is None:

            log(f"  {self.tag} HAR listen API unavailable", "WARN")

            return False

        try:

            listen.start(targets=True, collect_response=False)

            self.started = True

        except Exception as exc:

            log(f"  {self.tag} HAR listen start failed: {type(exc).__name__}: {exc}", "WARN")

            return False



        def _worker():

            while not self._stop.is_set():

                try:

                    packet = listen.wait(timeout=0.5)

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

            if self.started and getattr(self.page, "listen", None) is not None:

                self.page.listen.stop()

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

            status = int(getattr(packet, "status", 0) or 0)

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

        req = getattr(packet, "request", None) or {}

        resp = getattr(packet, "response", None) or {}

        url = getattr(packet, "url", "") or req.get("url") or ""

        method = getattr(packet, "method", "") or req.get("method") or ""

        status = int(getattr(packet, "status", 0) or resp.get("status", 0) or 0)

        headers = getattr(packet, "headers", None) or resp.get("headers") or {}

        event_type = getattr(packet, "event_type", "") or ""

        return {

            "startedDateTime": self._safe_started_datetime(getattr(packet, "timestamp", 0)),

            "time": 0,

            "request": {

                "method": method,

                "url": url,

                "httpVersion": "HTTP/2",

                "headers": self._headers_to_list(req.get("headers") or {}),

                "queryString": [],

                "cookies": [],

                "headersSize": -1,

                "bodySize": -1,

            },

            "response": {

                "status": status,

                "statusText": event_type,

                "httpVersion": "HTTP/2",

                "headers": self._headers_to_list(headers),

                "cookies": [],

                "content": {

                    "size": -1,

                    "mimeType": str(headers.get("content-type", "")) if isinstance(headers, dict) else "",

                },

                "redirectURL": "",

                "headersSize": -1,

                "bodySize": -1,

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





def _try_option_call(obj, names, *args, **kwargs):

    for name in names:

        fn = getattr(obj, name, None)

        if not callable(fn):

            continue

        try:

            fn(*args, **kwargs)

            return True, name

        except TypeError:

            continue

        except Exception:

            return False, name

    return False, ""





def _apply_ruoyi_headless_options(tb, tag, user_agent=None):

    """Best-effort Firefox headless tuning with ruyipage-native APIs."""

    ua = str(user_agent or HEADLESS_USER_AGENT or "").strip() or HEADLESS_USER_AGENT

    applied = []

    try:

        if callable(getattr(tb, "set_window_size", None)):

            tb.set_window_size(HEADLESS_WINDOW_WIDTH, HEADLESS_WINDOW_HEIGHT)

            applied.append(f"set_window_size({HEADLESS_WINDOW_WIDTH}x{HEADLESS_WINDOW_HEIGHT})")

    except Exception as exc:

        log(f"  {tag} set_window_size failed: {type(exc).__name__}: {exc}", "WARN")



    arg_pairs = (

        ("--width", str(HEADLESS_WINDOW_WIDTH)),

        ("--height", str(HEADLESS_WINDOW_HEIGHT)),

        ("--window-size", f"{HEADLESS_WINDOW_WIDTH},{HEADLESS_WINDOW_HEIGHT}"),

        ("--lang", "en-US"),

    )

    for arg, value in arg_pairs:

        ok, method = _try_option_call(tb, ("set_argument", "add_argument", "set_arg", "add_arg"), arg, value)

        if not ok:

            ok, method = _try_option_call(tb, ("set_argument", "add_argument", "set_arg", "add_arg"), f"{arg}={value}")

        if ok:

            applied.append(f"{method}({arg})")



    prefs = {

        "general.useragent.override": ua,

        "intl.accept_languages": "en-US,en",

        "dom.webdriver.enabled": False,

        "useAutomationExtension": False,

        "media.navigator.enabled": True,

        "webgl.disabled": False,

        "webgl.force-enabled": True,

        "webgl.enable-webgl2": True,

        "webgl.msaa-force": True,

        "webgl.disable-fail-if-major-performance-caveat": True,

        "webgl.min_capability_mode": False,

        "webgl.out-of-process": True,

        "webgl.angle.force-d3d11": True,

        "webgl.dxgl.enabled": True,

        "layers.acceleration.force-enabled": True,

        "layers.acceleration.disabled": False,

        "gfx.webrender.all": True,

        "gfx.webrender.enabled": True,

        "gfx.webrender.software": True,

        "gfx.webrender.force-disabled": False,

        "gfx.canvas.accelerated": True,

        "media.hardware-video-decoding.enabled": True,

        "privacy.resistFingerprinting": False,

    }

    for key, value in prefs.items():

        ok, method = _try_option_call(

            tb,

            ("set_preference", "set_pref", "set_prefs", "set_option"),

            key,

            value,

        )

        if ok:

            applied.append(f"{method}({key})")

    if applied:

        log(f"  {tag} ruoyi headless options applied: {', '.join(applied[:6])} ua={_mask_ua(ua)}")

    else:

        log(f"  {tag} ruoyi headless options: no compatible option API found", "WARN")





def _apply_ruoyi_browser_ua(tb, tag, user_agent):

    """有头/无头都尽量写 UA override，保证并发实例 UA 不一致。"""

    ua = str(user_agent or "").strip()

    if not ua:

        return False

    ok, method = _try_option_call(

        tb,

        ("set_preference", "set_pref", "set_prefs", "set_option"),

        "general.useragent.override",

        ua,

    )

    if ok:

        log(f"  {tag} browser ua override via {method}: {_mask_ua(ua)}")

        return True

    # 部分版本可能有 set_user_agent

    ok2, method2 = _try_option_call(tb, ("set_user_agent", "user_agent", "set_ua"), ua)

    if ok2:

        log(f"  {tag} browser ua override via {method2}: {_mask_ua(ua)}")

        return True

    return False





def _apply_ruoyi_headless_emulation(page, tag=None, log_once=False, user_agent=None):

    emu = getattr(page, "emulation", None)

    if emu is None:

        if log_once:

            log(f"  {tag} ruoyi headless emulation API not available", "WARN")

        return 0

    ua = str(user_agent or HEADLESS_USER_AGENT or "").strip() or HEADLESS_USER_AGENT

    applied = 0

    calls = (

        ("set_locale", ("en-US",)),

        ("set_screen_size", (1920, 1080)),

        ("set_device_scale_factor", (1,)),

        ("set_user_agent", (ua,)),

        ("set_extra_headers", ({"Accept-Language": "en-US,en;q=0.9"},)),

    )

    for name, args in calls:

        fn = getattr(emu, name, None)

        if not callable(fn):

            continue

        try:

            fn(*args)

            applied += 1

        except Exception:

            continue

    if log_once:

        level = "OK" if applied else "WARN"

        log(f"  {tag} ruoyi headless emulation applied: {applied} ua={_mask_ua(ua)}", level)

    return applied





def _build_headless_patch_js(user_agent=None):

    ua = str(user_agent or HEADLESS_USER_AGENT or "").strip() or HEADLESS_USER_AGENT

    ua_js = json.dumps(ua)

    # appVersion 常见形态：去掉 "Mozilla/" 前缀

    app_ver = ua[8:] if ua.startswith("Mozilla/") else ua

    app_ver_js = json.dumps(app_ver)

    return f"""

try {{

  Object.defineProperty(document, 'hidden', {{get: () => false, configurable: true}});

  Object.defineProperty(document, 'visibilityState', {{get: () => 'visible', configurable: true}});

}} catch(e) {{}}

try {{ document.hasFocus = function(){{ return true; }}; }} catch(e) {{}}

try {{

  Object.defineProperty(navigator, 'webdriver', {{get: () => undefined, configurable: true}});

}} catch(e) {{}}

try {{

  Object.defineProperty(navigator, 'userAgent', {{get: () => {ua_js}, configurable: true}});

  Object.defineProperty(navigator, 'appVersion', {{get: () => {app_ver_js}, configurable: true}});

}} catch(e) {{}}

try {{

  Object.defineProperty(navigator, 'languages', {{get: () => ['en-US', 'en'], configurable: true}});

}} catch(e) {{}}

try {{

  if (!navigator.deviceMemory) Object.defineProperty(navigator, 'deviceMemory', {{get: () => 8, configurable: true}});

  if (!navigator.hardwareConcurrency || navigator.hardwareConcurrency < 2) {{

    Object.defineProperty(navigator, 'hardwareConcurrency', {{get: () => 8, configurable: true}});

  }}

}} catch(e) {{}}

try {{

  Object.defineProperty(screen, 'width', {{get: () => 1920, configurable: true}});

  Object.defineProperty(screen, 'height', {{get: () => 1080, configurable: true}});

  Object.defineProperty(screen, 'availWidth', {{get: () => 1920, configurable: true}});

  Object.defineProperty(screen, 'availHeight', {{get: () => 1040, configurable: true}});

  Object.defineProperty(screen, 'colorDepth', {{get: () => 24, configurable: true}});

  Object.defineProperty(screen, 'pixelDepth', {{get: () => 24, configurable: true}});

}} catch(e) {{}}

try {{

  if (!window.outerWidth || window.outerWidth < 100) {{

    Object.defineProperty(window, 'outerWidth', {{get: () => {HEADLESS_WINDOW_WIDTH} + 16, configurable: true}});

  }}

  if (!window.outerHeight || window.outerHeight < 100) {{

    Object.defineProperty(window, 'outerHeight', {{get: () => {HEADLESS_WINDOW_HEIGHT} + 88, configurable: true}});

  }}

}} catch(e) {{}}

try {{

  if (!navigator.plugins || navigator.plugins.length === 0) {{

    Object.defineProperty(navigator, 'plugins', {{get: () => [1, 2, 3], configurable: true}});

  }}

}} catch(e) {{}}

try {{

  const proto = HTMLCanvasElement && HTMLCanvasElement.prototype;

  const origGetContext = proto && proto.getContext;

  if (origGetContext && !proto.__ruoyiWebglShim) {{

    Object.defineProperty(proto, '__ruoyiWebglShim', {{value: true}});

    const fakeWebgl = (canvas) => {{

      const dbg = {{UNMASKED_VENDOR_WEBGL: 37445, UNMASKED_RENDERER_WEBGL: 37446}};

      const values = {{

        7936: 'WebKit',

        7937: 'WebKit WebGL',

        7938: 'WebGL 1.0',

        35724: 'WebGL GLSL ES 1.0',

        37445: 'Google Inc. (NVIDIA)',

        37446: 'ANGLE (NVIDIA, NVIDIA GeForce RTX 3060 Direct3D11 vs_5_0 ps_5_0, D3D11)',

        3379: 16384,

        3386: new Int32Array([16384, 16384]),

        3410: 8,

        3411: 8,

        3412: 8,

        3413: 8,

        3414: 24,

        3415: 8,

      }};

      const base = {{

        canvas,

        VENDOR: 7936,

        RENDERER: 7937,

        VERSION: 7938,

        SHADING_LANGUAGE_VERSION: 35724,

        drawingBufferWidth: canvas.width || 300,

        drawingBufferHeight: canvas.height || 150,

        getExtension: (name) => String(name || '').toUpperCase() === 'WEBGL_DEBUG_RENDERER_INFO' ? dbg : null,

        getSupportedExtensions: () => ['WEBGL_debug_renderer_info', 'OES_texture_float', 'OES_standard_derivatives'],

        getParameter: (p) => Object.prototype.hasOwnProperty.call(values, p) ? values[p] : 0,

        isContextLost: () => false,

      }};

      return new Proxy(base, {{

        get(target, prop) {{

          if (prop in target) return target[prop];

          if (typeof prop === 'string' && /^[A-Z0-9_]+$/.test(prop)) return 0;

          return function () {{ return 0; }};

        }}

      }});

    }};

    proto.getContext = function(type, attrs) {{

      const t = String(type || '').toLowerCase();

      const ctx = origGetContext.call(this, type, attrs);

      if (ctx) return ctx;

      if (t === 'webgl' || t === 'experimental-webgl' || t === 'webgl2') return fakeWebgl(this);

      return ctx;

    }};

  }}

}} catch(e) {{}}

try {{

  const gp = WebGLRenderingContext && WebGLRenderingContext.prototype.getParameter;

  if (gp && !WebGLRenderingContext.prototype.__ruoyiHeadlessPatched) {{

    Object.defineProperty(WebGLRenderingContext.prototype, '__ruoyiHeadlessPatched', {{value: true}});

    WebGLRenderingContext.prototype.getParameter = function(p) {{

      if (p === 37445) return 'NVIDIA Corporation';

      if (p === 37446) return 'NVIDIA GeForce GTX 750 Ti/PCIe/SSE2';

      return gp.call(this, p);

    }};

  }}

}} catch(e) {{}}

try {{

  delete window.__playwright;

  delete window.__pwInitScripts;

  for (const k of Object.getOwnPropertyNames(window)) {{

    if (/^cdc_|webdriver|selenium|driver/i.test(k)) {{

      try {{ delete window[k]; }} catch(e) {{}}

    }}

  }}

}} catch(e) {{}}

return true;

"""





# 兼容旧引用（无 UA 参数时的默认 patch）

RUOYI_HEADLESS_PATCH_JS = _build_headless_patch_js(HEADLESS_USER_AGENT)





def _apply_ruoyi_headless_page_patches(page, tag=None, log_once=False, user_agent=None):

    ua = str(user_agent or HEADLESS_USER_AGENT or "").strip() or HEADLESS_USER_AGENT

    _apply_ruoyi_headless_emulation(page, tag, log_once=False, user_agent=ua)

    patch_js = _build_headless_patch_js(ua)

    ok_count = 0

    for ctx in _all_contexts(page):

        try:

            if ctx.run_js_loaded(patch_js):

                ok_count += 1

        except Exception:

            continue

    if log_once:

        level = "DEBUG"

        log(f"  {tag} ruoyi headless page patches applied to {ok_count} context(s) ua={_mask_ua(ua)}", level)

    return ok_count





def _on_signup_form(url):

    low = (url or "").lower()

    return "signup.live.com" in low and "privacynotice" not in low







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





def _all_contexts(page):



    contexts = [page]

    try:

        contexts.extend(page.get_all_frames() or [])

    except Exception:

        pass

    return contexts





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

            time.sleep(0.2)

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

        time.sleep(0.25)

        if _click_option_exact(candidates):

            time.sleep(0.15)

            _press_escape()

            return True

        if typed_fallback is not None:

            try:

                page.actions.type(str(typed_fallback)).press("\ue007").perform()

                time.sleep(0.15)

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

            time.sleep(0.1)



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

            time.sleep(0.1)

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

            checked = bool(getattr(cb, "is_checked", False))

        except Exception:

            checked = False

        if not checked:

            _safe_click(cb)

            log(f"  {tag} checked required checkbox")

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
    script = (r"""
return (() => {
  const HOLD_PATTERNS = %s.map((src) => new RegExp(src, 'i'));
  const CHALLENGE_PATTERNS = %s.map((src) => new RegExp(src, 'i'));
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
  const collectCandidates = (root) => {
    const candidates = [];
    const seen = new Set();
    const addCandidate = (el, quality, source, labelEl = null) => {
      if (!el || seen.has(el) || !visible(el) || !buttonish(el)) return;
      seen.add(el);
      candidates.push(pack(el, quality, source, labelEl));
    };

    if (root && root.id === 'px-captcha' && visible(root) && buttonish(root)) {
      addCandidate(root, 0, '#px-captcha');
    }

    for (const el of root.querySelectorAll('[role="button"][aria-label], button[aria-label], a[role="button"][aria-label], input[aria-label]')) {
      const aria = norm(el.getAttribute('aria-label'));
      if (!aria) continue;
      const hasHold = matchesAny(aria, HOLD_PATTERNS);
      const hasChallenge = matchesAny(aria, CHALLENGE_PATTERNS);
      if (hasHold && hasChallenge) addCandidate(el, 1, 'aria-label:hold+challenge');
      else if (hasHold) addCandidate(el, 2, 'aria-label:hold');
    }

    for (const node of root.querySelectorAll('p, span, div')) {
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

    for (const el of root.querySelectorAll('[role="button"], button, a[role="button"], input[type="button"], input[type="submit"]')) {
      if (matchesAny(textOf(el), HOLD_PATTERNS)) addCandidate(el, 5, 'button-text-hold');
    }

    if (!candidates.length) return null;
    candidates.sort((a, b) =>
      (a.quality - b.quality) ||
      ((a.width * a.height) - (b.width * b.height)) ||
      ((b.ariaLabel || '').length - (a.ariaLabel || '').length)
    );
    return candidates[0];
  };

  const px = document.querySelector('#px-captcha');
  if (px) {
    const scoped = collectCandidates(px);
    if (scoped) return scoped;
  }
  return collectCandidates(document);
})()
""" % (json.dumps(hold_patterns, ensure_ascii=True), json.dumps(challenge_patterns, ensure_ascii=True)))
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





def _focus_page_before_captcha_press(page, tag):

    """Click the main page before each captcha hold so PX receives focus."""

    try:

        page.run_js_loaded(

            """

try { window.focus(); } catch (e) {}

try { document.body && document.body.focus && document.body.focus(); } catch (e) {}

return true;

"""

        )

    except Exception:

        pass

    try:

        # Real pointer click on a safe main-page point; avoid the captcha iframe itself.

        page.actions.move_to({"x": 18, "y": 18}, duration=random.randint(120, 260)).hold().wait(

            random.uniform(0.05, 0.12)

        ).release().perform()

        log(f"  {tag} focused page before captcha press")

        time.sleep(0.25)

        return True

    except Exception as exc:

        log(f"  {tag} focus click before captcha press failed: {type(exc).__name__}: {exc}", "WARN")

        return False





def _perform_hold(page, ctx, target, idx, press_count, tag):

    """Press-and-hold with early release when the linked Press and hold p hides itself."""

    cx = int(target.get("x", 0) + random.uniform(-3, 3))

    cy = int(target.get("y", 0) + random.uniform(-2, 2))

    hold_sec = random.uniform(PX_HOLD_SECONDS_MIN, PX_HOLD_SECONDS_MAX)

    state_ctx = ctx

    hold_label = _resolve_hold_label_for_target(state_ctx, target) or {}

    hold_p_id = str(hold_label.get("id") or "").strip()

    log(

        f"  {tag} press #{press_count}: ({cx},{cy}) hold={hold_sec:.1f}s"

        + (f" text={target.get('text', '')[:30]!r}" if target.get("text") else "")

        + (f" hold_p_id={hold_p_id}" if hold_p_id else "")

    )

    actions = ctx.actions

    hold_started = None

    release_state = None

    extra_release_wait = 0.0

    try:

        actions.move_to({"x": cx, "y": cy}, duration=random.randint(250, 550)).hold().perform()

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

                    extra_release_wait = random.uniform(0.5, 1.5)

                    time.sleep(extra_release_wait)

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

            extra = ""

            if extra_release_wait > 0:

                extra += f" extra wait {extra_release_wait:.2f}s"

            if label_text:

                extra += f" text={label_text[:20]!r}"

            log(f"  {tag} press #{press_count} released early after {held:.1f}s because hold label hid{extra}")

        return held

    except Exception as exc:

        if hold_started is not None:

            try:

                actions.release_all()

            except Exception:

                pass

        log(f"  {tag} hold failed: {type(exc).__name__}: {exc}", "WARN")

        return False



def _perform_hold_with_px_screenshots(page, ctx, target, idx, press_count, tag, enabled=False):

    if enabled:

        _save_screenshot(page, "before_press_last", idx, tag)

    ok = _perform_hold(page, ctx, target, idx, press_count, tag)

    if enabled:

        _save_screenshot(page, "after_press_last", idx, tag)

    return ok





def _wait_before_next_captcha_press(tag, reason="challenge failed"):

    delay = random.uniform(POST_PRESS_RETRY_GAP_MIN, POST_PRESS_RETRY_GAP_MAX)

    log(f"  {tag} {reason}, retry press in {delay:.2f}s")

    time.sleep(delay)

    return delay





def _proxy_for_ip_lookup(proxy_pool, tag):

    if not proxy_pool:

        return None

    proxy_str = str(proxy_pool[0] or "").strip()

    p = _parse_ruoyi_proxy(proxy_str)

    if not p:

        log(f"  {tag} IP probe proxy format invalid: {proxy_str[:80]!r}", "WARN")

        return None

    user = quote(p.get("username") or "", safe="")

    pwd = quote(p.get("password") or "", safe="")

    auth = f"{user}:{pwd}@" if (user or pwd) else ""

    proxy_url = f"socks5h://{auth}{p['host']}:{p['port']}"

    return {"http": proxy_url, "https": proxy_url}



def _log_current_ip(proxy_pool, tag):

    proxies = _proxy_for_ip_lookup(proxy_pool, tag)

    session = requests.Session()

    session.trust_env = False

    for source_name, ip_endpoint in IP_INFO_ENDPOINTS:

        try:

            resp = session.get(

                ip_endpoint,

                headers={"Accept": "application/json, text/plain;q=0.9, */*;q=0.8"},

                proxies=proxies,

                timeout=15,

            )

            resp.raise_for_status()

            data = resp.json()

            ip = str(data.get("ip") or data.get("query") or "").strip()

            country = str(

                data.get("country_name")

                or data.get("country")

                or data.get("countryCode")

                or data.get("country_code")

                or ""

            ).strip()

            if ip:

                country_text = country or "UNKNOWN"

                log(f"  {tag} current IP: {ip!r} country: {country_text!r} via {source_name}")

                return

            log(f"  {tag} IP parse failed({source_name}): {resp.text[:120]!r}", "WARN")

        except Exception as e:

            log(f"  {tag} IP probe failed({source_name}): {e}", "WARN")

    log(f"  {tag} current IP/country probe failed", "WARN")





def _probe_proxy_before_browser(proxy_pool, tag, timeout=PROXY_PRECHECK_TIMEOUT):

    proxies = _proxy_for_ip_lookup(proxy_pool, tag)

    if not proxies:

        return True

    session = requests.Session()

    session.trust_env = False

    try:

        resp = session.get(

            PROXY_PRECHECK_URL,

            headers={"User-Agent": "Mozilla/5.0", "Accept": "text/html,application/xhtml+xml"},

            proxies=proxies,

            timeout=max(0.1, float(timeout or PROXY_PRECHECK_TIMEOUT)),

            allow_redirects=False,

            stream=True,

        )

        status = getattr(resp, "status_code", 0)

        log(f"  {tag} proxy precheck ok: status={status} url={PROXY_PRECHECK_URL}", "INFO")

        return True

    except Exception as exc:

        log(f"  {tag} proxy precheck failed: {type(exc).__name__}: {exc}", "WARN")

        return False





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



    mode = (

        getattr(opts, "account_format_mode", None)

        or os.environ.get("OUTLOOK_ACCOUNT_FORMAT_MODE")

        or ("custom" if os.environ.get("OUTLOOK_ACCOUNT_FORMAT") else "name")

    )

    custom = (

        getattr(opts, "account_format", None)

        or os.environ.get("OUTLOOK_ACCOUNT_FORMAT")

        or ""

    )

    password_format = (

        getattr(opts, "password_format", None)

        or os.environ.get("OUTLOOK_PASSWORD_FORMAT")

        or ""

    )

    set_fn = getattr(helpers, "set_account_generation_options", None)

    resolve_format_fn = getattr(helpers, "resolve_account_format", None)

    resolved_format = ""

    if callable(resolve_format_fn):

        resolved_format = resolve_format_fn(mode, custom)

    elif custom and str(mode).lower() == "custom":

        resolved_format = str(custom)

    elif str(mode).lower() == "name":

        resolved_format = "{first}_{last}"

    elif str(mode).lower() == "name_digits":

        resolved_format = "{first}_{last}{digits:3}"

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

    if proxy_pool:

        tb.set_per_tab_proxies(proxy_pool, exhausted="wrap")

        log(f"挂载 {len(proxy_pool)} 条 SOCKS5 代理到 per-tab 池(wrap 轮换)")

    else:

        log("没挂代理——直接本机出口", "WARN")

    # 有头/无头都写 UA，避免 4 并发全是同一条默认 UA

    _apply_ruoyi_browser_ua(tb, tag, user_agent)

    if is_headless:

        _apply_ruoyi_headless_options(tb, tag, user_agent=user_agent)

        tb.headless(True)



    log(

        f"启动 ruyipage Firefox: model={_browser_model_name(RUOYI_FIREFOX_PATH)} "

        f"headless={is_headless} path={RUOYI_FIREFOX_PATH} ua={_mask_ua(user_agent)}",

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

    no_target_rounds = 0

    initial_press_wait_started = None

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



        if proxy_pool:

            log(f"  {tag} 使用当前打开页面承载注册页，不再新建 container tab")

        try:

            browser_page.close_other_tabs(page)

            log(f"  {tag} 已关闭 Firefox 启动默认空白页，仅保留当前注册页")

        except Exception as close_exc:

            log(f"  {tag} 关闭默认空白页失败: {type(close_exc).__name__}: {close_exc}", "WARN")



        if is_headless and not signup_opened:

            _apply_ruoyi_headless_page_patches(page, tag, log_once=True, user_agent=user_agent)

            headless_patch_logged = True

        if capture_har:

            har_collector = _RuoyiHarCollector(page, tag, idx, email_getter=lambda: email or "")

            har_collector.start()

        if not signup_opened:

            _, step_timings["open_signup"] = _timed_step(tag, "open_signup", page.get, SIGNUP_URL)

        _, step_timings["wait_loading"] = _timed_step(tag, "wait_loading", page.wait_loading, 20)

        if is_headless:

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

        log(f"  {tag} 将注册: {email}")



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

        while time.time() < deadline:

            if is_headless:

                _apply_ruoyi_headless_page_patches(

                    page, tag, log_once=(not headless_patch_logged), user_agent=user_agent

                )

                headless_patch_logged = True

            submitted = False

            current_url = page.url.lower()

            body = _body_text(page)

            low = body.lower()



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

                log(f"  {tag} Microsoft Loading page, keep waiting for redirect")

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



            if _maybe_skip_passkey(page, tag):

                submit_wait_started = None
                last_submit_click_hit = None

                continue

            if "privacynotice" in current_url:

                submit_wait_started = None
                last_submit_click_hit = None

                _click_post_signup(page, tag)

                time.sleep(3)

                continue



            visible = _captcha_visible(page)

            validating = _captcha_is_validating(page)

            hold_ctx, hold_target = _find_hold_context(page, min_quality=5)

            actionable = bool(

                visible

                and (not validating)

                and hold_target is not None

                and _target_quality(hold_target) <= 5

            )

            loading_page = _microsoft_loading_page(page)
            loop_now = time.time()
            press_wait_started, microsoft_loading_wait_started, loading_timed_out = _update_loading_wait_state(
                press_wait_started,
                microsoft_loading_wait_started,
                loading=loading_page,
                now=loop_now,
                timeout=MICROSOFT_LOADING_TIMEOUT,
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

                time.sleep(3)

                continue

            microsoft_loading_wait_started = None



            if awaiting_reappear and press_count < max_press:

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

                time.sleep(3)

                continue



            if visible and actionable:

                validation_wait_started = None

                had_captcha = True

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

                            time.sleep(3)

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

                        _focus_page_before_captcha_press(page, tag)

                        hold_elapsed = _perform_hold_with_px_screenshots(

                            page, ctx, target, idx, press_count, tag, enabled=px_press_screenshots

                        )

                        if hold_elapsed:

                            px_hold_elapsed += float(hold_elapsed)

                            validation_wait_started = time.time()

                            awaiting_reappear = True

                            post_press_saw_gap = False

                            post_press_started_at = time.time()

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

                        time.sleep(2)

                        continue

                    if validation_wait_started is None:

                        validation_wait_started = time.time()

                        log(f"  {tag} post-captcha state unclear, keep waiting")

                    elif _wait_state_timed_out(validation_wait_started, timeout=CAPTCHA_STATE_TIMEOUT):

                        waited = time.time() - validation_wait_started

                        log(f"  {tag} post-captcha state unclear for {int(waited)}s, give up", "WARN")

                        _shot(page, "timeout_captcha_state", idx)

                        return _finish(reason="timeout")

                    time.sleep(3)

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



            if int(time.time()) % 15 < 3:

                _shot(page, f"wait_{int(time.time() % 1000)}", idx)

            if not submitted:

                time.sleep(3)

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

        if har_collector is not None:

            if capture_har:

                try:

                    har_collector.save(reason="success" if success else failure_reason)

                except Exception as exc:

                    log(f"  {tag} HAR save failed: {type(exc).__name__}: {exc}", "WARN")

            else:

                har_collector.stop()

        if page is not None and browser_page is not None and page is not browser_page:

            try:

                page.close()

            except Exception:

                pass

        try:

            if browser_page is not None:

                _quit_browser_page(browser_page, tag, timeout=min(BROWSER_QUIT_TIMEOUT, max(0.01, deadline - time.time())))

        except Exception:

            pass

        _untrack_browser_page(browser_page)

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

            f.write(f"{email}----{password}----{graph['refresh_token']}----{graph.get('client_id', '')}\n")

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

                    "refresh_token": graph.get("refresh_token"),

                    "client_id": graph.get("client_id"),

                }

            )

            with open(token_file, "w", encoding="utf-8") as f:

                json.dump(tokens, f, ensure_ascii=False, indent=2)

    append_graph_account_to_emails_pool(email, password, graph)

    _record_webui_account(email, password, graph, "ok")





async def _run_one_direct(args, helpers, proxy_pool, idx, total, save_lock, consumable_pool=None):

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
            return "fail", total_elapsed, px_metrics

    email = password = None

    fail_reason = "failure"

    try:

        remaining = max(0.0, overall_deadline - time.perf_counter())
        if remaining <= 0:
            raise TimeoutError(f"account timeout before browser ({total_budget:g}s)")
        run_args = SimpleNamespace(**vars(args))
        run_args.timeout = remaining
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

        return "fail", total_elapsed, px_metrics



    log(f"{tag} Graph token extracting…", "INFO")

    # Graph 授权强制直连：不挂注册代理，extract_graph_token_http 传 proxy_str=None。

    log(f"{tag} graph proxy -> direct", "INFO")

    graph = await asyncio.to_thread(helpers.extract_graph_token_http, email, password, idx, 3, None)

    if not graph or not graph.get("refresh_token"):

        async with save_lock:

            await asyncio.to_thread(_save_no_graph_result, email, password)

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





async def _run_direct_batch(args, helpers, consumable_pool):

    """返回 (ok, no_graph, fail, total_elapsed, avg_success_elapsed) 五元组。"""

    count = max(0, int(args.count or 0))

    concurrency = max(1, int(args.concurrency or 1))

    sem = asyncio.Semaphore(concurrency)

    save_lock = asyncio.Lock()

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

            return await _run_one_direct(args, helpers, consumable_pool, i + 1, count, save_lock, consumable_pool)



    try:

        results = await asyncio.gather(*(runner(i) for i in range(count)))

    finally:

        log(f"proxy list end: remaining={consumable_pool.remaining()}", "INFO")

        consumable_pool.stop()

        set_consumable_proxy_pool(None)

    statuses = [r[0] if isinstance(r, tuple) else r for r in results]

    elapsed_list = [r[1] if isinstance(r, tuple) and len(r) > 1 else 0.0 for r in results]

    px_stats = [_normalize_px_metrics(i + 1, r[2] if isinstance(r, tuple) and len(r) > 2 else None) for i, r in enumerate(results)]

    ok = sum(1 for r in statuses if r == "ok")

    no_graph = sum(1 for r in statuses if r == "no_graph")

    fail = sum(1 for r in statuses if r == "fail")

    # 兼容旧布尔返回

    fail += sum(1 for r in statuses if r is False)

    ok += sum(1 for r in statuses if r is True)

    summary = _summarize_batch_metrics(statuses, elapsed_list, time.perf_counter() - batch_started, px_stats=px_stats)

    return ok, no_graph, fail, summary["total_elapsed"], summary["avg_success_elapsed"], summary["px_stats"]





def main():

    ap = argparse.ArgumentParser(description="Outlook 自注册养号(ruoyi) — 完整链路版")

    ap.add_argument("--proxy-file", "-p", default=PROXY_FILE, help=f"代理池文件(默认 {PROXY_FILE})")

    ap.add_argument("--proxy-source", default=RUOYI_PROXY_SOURCE,

                    choices=["file", "aimili-random", "aimili-list"],

                    help="ruoyi 代理来源：file=本地文件；aimili-random=Aimili 随机接口；aimili-list=Aimili 列表。启动装 list，注册取删，空则重载")

    ap.add_argument("--aimili-url", default=AIMILI_POOL_URL,

                    help="AimiliVPN URL：可填管理端根地址 http://host:8787，也可直接填 /api/pool/proxies 或 /api/pool/proxies/random 完整地址")

    ap.add_argument("--aimili-token", default=AIMILI_POOL_TOKEN,

                    help="AimiliVPN 代理池 API Token")

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

    ap.add_argument("--timeout", "-t", type=int, default=REGISTER_TIMEOUT, help="单号超时(秒)")

    ap.add_argument("--max-press", default=os.environ.get("OUTLOOK_REG_MAX_PRESS", "5"), help="按住次数上限")

    ap.add_argument("--no-verify", action="store_true", help="跳过 Outlook 登录校验")

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

            or ("custom" if os.environ.get("OUTLOOK_ACCOUNT_FORMAT") else "name")

        ),

        choices=["random", "name", "name_digits", "custom"],

        help="账号格式：random/name/name_digits/custom",

    )

    ap.add_argument(

        "--account-format",

        default=os.environ.get("OUTLOOK_ACCOUNT_FORMAT", ""),

        help="指定格式模板，如 {first}.{last}{digits:3}；mode=custom 时生效",

    )

    ap.add_argument(

        "--password-format",

        default=os.environ.get("OUTLOOK_PASSWORD_FORMAT", ""),

        help="密码模板，如 Aa1!{rand:12}；留空用默认随机",

    )

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



    os.makedirs(OUTPUT_DIR, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    live_file = os.path.join(OUTPUT_DIR, f"accounts_ruoyi_{ts}.txt")

    token_file = os.path.join(OUTPUT_DIR, f"graph_tokens_ruoyi_{ts}.json")

    args.live_file = live_file

    args.token_file = token_file



    log(

        f"开始: count={args.count} concurrency={max(1, int(args.concurrency or 1))} "

        f"launch_stagger={stagger_val:g}s ua_pool={len(_load_ua_pool())} timeout={args.timeout}s"

    )

    ok, no_graph, failed, batch_total_elapsed, avg_success_elapsed, px_stats = asyncio.run(_run_direct_batch(args, helpers, consumable_pool))

    total = max(0, int(args.count or 0))

    # 多行汇总：WebUI 正则 + 人眼可读中文都覆盖

    for line in _format_batch_summary_lines(ok, no_graph, failed, total, batch_total_elapsed, avg_success_elapsed, px_stats):

        log(line, "OK")

    if os.path.isfile(live_file):

        log(f"账号输出: {live_file}")

    if os.path.isfile(token_file):

        log(f"Token 输出: {token_file}")

    if no_graph and os.path.isfile(EMAIL_NOGRAPH):

        log(f"未授权输出: {EMAIL_NOGRAPH}")



    log(f"email_nograph: {EMAIL_NOGRAPH}")



if __name__ == "__main__":

    main()

