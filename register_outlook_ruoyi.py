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
import importlib.util
import json
import os
import random
import requests
import sys
import threading
import time
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
SCREENSHOT_DIR = os.path.join(ROOT, "screenshots_ruoyi")
HAR_DIR = os.path.join(ROOT, "har_ruoyi")
OUTPUT_DIR = os.path.join(ROOT, "outlook_accounts")
EMAILS_POOL = os.path.join(ROOT, "emails.txt")
SIGNUP_URL = "https://signup.live.com/signup?lic=1"
IP_INFO_ENDPOINTS = [
    ("ipwhois", "https://ipwho.is/"),
]
_CURRENT_IP_INFO = {}

REGISTER_TIMEOUT = 300
VERIFY_AFTER_REGISTER = True
# Wait before each captcha press.
INITIAL_PRESS_DELAY = 8
# Give up after max_press if no redirect happens within this many seconds.
POST_MAX_PRESS_WAIT = 5
# After a press, wait this long before retrying even if loading is not seen.
POST_PRESS_LOADING_CHECK = 8
# Short retry gap after a press that did not trigger loading.
POST_PRESS_RETRY_GAP = 0
# Slow down page start and form submit; headless can outrun Microsoft/PX state setup.
PAGE_START_DELAY = float(os.environ.get("OUTLOOK_RUOYI_PAGE_START_DELAY", "2") or "2")
SUBMIT_DELAY = float(os.environ.get("OUTLOOK_RUOYI_SUBMIT_DELAY", "2") or "2")
HEADLESS_WINDOW_WIDTH = int(os.environ.get("OUTLOOK_RUOYI_HEADLESS_WIDTH", "1280") or "1280")
HEADLESS_WINDOW_HEIGHT = int(os.environ.get("OUTLOOK_RUOYI_HEADLESS_HEIGHT", "800") or "800")
HEADLESS_USER_AGENT = os.environ.get(
    "OUTLOOK_RUOYI_HEADLESS_UA",
    (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:151.0) "
        "Gecko/20100101 Firefox/151.0"
    ),
)
_HELPERS = None


def _env_bool(name, default=False):
    value = os.environ.get(name)
    if value is None:
        return bool(default)
    return str(value).strip().lower() not in ("0", "false", "no", "off", "")


def log(msg, level="INFO"):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] [{level}] {msg}", flush=True)


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


def parse_proxy_pool(path):
    """读 user:pass@host:port 文件，转成 ruyipage 的 socks5 per-tab 格式。"""
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
            for pfx in ("socks5://", "http://", "https://"):
                if ln.lower().startswith(pfx):
                    ln = ln[len(pfx):]
                    break
            if "@" not in ln:
                log(f"跳过非法代理行(无 @): {ln[:60]}", "WARN")
                continue
            auth, hostport = ln.rsplit("@", 1)
            if ":" not in auth or ":" not in hostport:
                log(f"跳过非法代理行(缺 user/pass 或 host/port): {ln[:60]}", "WARN")
                continue
            user, pwd = auth.split(":", 1)
            host, port = hostport.split(":", 1)
            out.append(f"{host}:{port}:{user}:{pwd}")
    return out


def select_proxy_for_account(proxy_pool):
    """每个账号开跑前随机挑一个代理。"""
    if not proxy_pool:
        return []
    return [random.choice(proxy_pool)]


def append_graph_account_to_emails_pool(email, password, graph):
    token = (graph or {}).get("refresh_token") or ""
    client_id = (graph or {}).get("client_id") or ""
    if not token:
        log(f"emails.txt skip {email}: no graph refresh_token", "WARN")
        return False
    try:
        existing = set()
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


def _shot(page, name, idx):
    failure_prefixes = (
        "blocked",
        "error",
        "timeout",
        "press_fail",
        "captcha_no_target",
        "email_fail",
        "email_input_fail",
        "email_exc",
        "no_email",
        "pwd_fail",
        "bday_fail",
        "name_fail",
    )
    if not any(name.startswith(prefix) for prefix in failure_prefixes):
        return None
    os.makedirs(SCREENSHOT_DIR, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
    base = os.path.join(SCREENSHOT_DIR, f"ruoyi_{idx}_{name}_{ts}")
    path = f"{base}.png"
    try:
        page.screenshot(path=path, full_page=True)
        if os.path.isfile(path) and os.path.getsize(path) > 0:
            log(f"??: {path}", "OK")
        else:
            log(f"????????: {path}", "WARN")
    except Exception as e:
        log(f"???? {path}: {e}", "WARN")
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
        log(f"??: {state_path}", "OK")
    except Exception as e:
        log(f"?????? {base}.html: {e}", "WARN")
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


def _apply_ruoyi_headless_options(tb, tag):
    """Best-effort Firefox headless tuning with ruyipage-native APIs."""
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
        "general.useragent.override": HEADLESS_USER_AGENT,
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
        log(f"  {tag} ruoyi headless options applied: {', '.join(applied[:6])}")
    else:
        log(f"  {tag} ruoyi headless options: no compatible option API found", "WARN")


def _apply_ruoyi_headless_emulation(page, tag=None, log_once=False):
    emu = getattr(page, "emulation", None)
    if emu is None:
        if log_once:
            log(f"  {tag} ruoyi headless emulation API not available", "WARN")
        return 0
    applied = 0
    calls = (
        ("set_locale", ("en-US",)),
        ("set_screen_size", (1920, 1080)),
        ("set_device_scale_factor", (1,)),
        ("set_user_agent", (HEADLESS_USER_AGENT,)),
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
        log(f"  {tag} ruoyi headless emulation applied: {applied}", level)
    return applied


RUOYI_HEADLESS_PATCH_JS = f"""
try {{
  Object.defineProperty(document, 'hidden', {{get: () => false, configurable: true}});
  Object.defineProperty(document, 'visibilityState', {{get: () => 'visible', configurable: true}});
}} catch(e) {{}}
try {{ document.hasFocus = function(){{ return true; }}; }} catch(e) {{}}
try {{
  Object.defineProperty(navigator, 'webdriver', {{get: () => undefined, configurable: true}});
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


def _apply_ruoyi_headless_page_patches(page, tag=None, log_once=False):
    _apply_ruoyi_headless_emulation(page, tag, log_once=False)
    ok_count = 0
    for ctx in _all_contexts(page):
        try:
            if ctx.run_js_loaded(RUOYI_HEADLESS_PATCH_JS):
                ok_count += 1
        except Exception:
            continue
    if log_once:
        level = "OK" if ok_count else "WARN"
        log(f"  {tag} ruoyi headless page patches applied to {ok_count} context(s)", level)
    return ok_count


def _on_signup_form(url):
    low = (url or "").lower()
    return "signup.live.com" in low and "privacynotice" not in low



def _scroll_into_view(el):
    """把元素滚进视口，避免 ruyipage 报『无法获取元素可点击坐标』。"""
    if el is None:
        return False
    try:
        el.run_js("this.scrollIntoView({block:'center', inline:'nearest'});")
        return True
    except Exception:
        pass
    try:
        # 某些版本用 page 上下文执行
        owner = getattr(el, "owner", None) or getattr(el, "page", None)
        if owner is not None:
            owner.run_js_loaded(
                "const el=arguments[0]; if(el&&el.scrollIntoView) el.scrollIntoView({block:'center',inline:'nearest'}); return true;",
                el,
            )
            return True
    except Exception:
        pass
    return False


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
            ok = el.run_js(
                """
const v = String(arguments[0] ?? '');
const doClear = !!arguments[1];
const el = this;
el.focus();
const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set
  || Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value')?.set;
if (doClear) {
  if (setter) setter.call(el, ''); else el.value = '';
  el.dispatchEvent(new Event('input', {bubbles: true}));
}
if (setter) setter.call(el, v); else el.value = v;
el.dispatchEvent(new Event('input', {bubbles: true}));
el.dispatchEvent(new Event('change', {bubbles: true}));
return true;
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
                el.run_js("this.click(); return true;")
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


def _submit_wait(tag):
    if SUBMIT_DELAY > 0:
        log(f"  {tag} submit wait {SUBMIT_DELAY:g}s")
        time.sleep(SUBMIT_DELAY)


def _click_next(page, tag):
    sels = [
        'css:input[type="submit"]',
        'css:button[type="submit"]',
        '#iSignupAction',
        'css:button[id="iSignupAction"]',
        'text:Next',
        'text:下一步',
        'text:Suivant',
        '#iNext',
    ]
    hit = _click_any(page, sels, timeout=2)
    if not hit:
        try:
            page.actions.press("\ue007").perform()
        except Exception:
            pass
    log(f"  {tag} next: {hit or '(Enter)'}")
    return bool(hit)


def _is_birthday_page(page):
    txt = _body_text(page)
    low = txt.lower()
    if "add some details" in low and ("birthdate" in low or "month" in low or "year" in low):
        return True
    if any(k in txt for k in ["请输入你的出生日期", "输入你的出生日期", "输入出生日期"]):
        return True
    if any(k in low for k in ["enter your birthdate", "birthdate", "country/region"]):
        return True
    return False


def _is_name_page(page):
    txt = _body_text(page)
    low = txt.lower()
    return any(
        k in low
        for k in [
            "first name",
            "last name",
            "what's your name",
            "your name",
            "name",
            "pr?nom",
            "nom de famille",
        ]
    ) or any(k in txt for k in ["??", "??", "??", "?", "?"])


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


def _handle_consent(page, tag, idx):
    for attempt in range(5):
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
            timeout=2,
        )
        time.sleep(3)
        _shot(page, f"after_consent_{attempt}", idx)


def _email_error_kind(page):
    """返回 'taken' / 'format' / ''。"""
    lower = _body_text(page).lower()
    if ("already" in lower and "email" in lower) or "taken" in lower:
        return "taken"
    # 中文占用提示
    txt = _body_text(page)
    if any(k in txt for k in ["已被使用", "不可用", "已被占用", "已经有人", "换一个"]):
        if any(k in lower for k in ["email", "address", "microsoft", "outlook", "帐户", "账户", "账号"]):
            return "taken"
        if any(k in txt for k in ["电子邮件", "邮箱", "用户名"]):
            return "taken"
    if any(k in lower for k in ["needs to start", "in the format", "enter a valid", "use letters"]):
        return "format"
    if any(k in txt for k in ["格式", "无效", "请输入有效"]):
        return "format"
    return ""


def _email_domain(email):
    if not email or "@" not in str(email):
        return "outlook.com"
    return str(email).rsplit("@", 1)[-1].strip().lower() or "outlook.com"


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


def _fill_email(page, email, prefix, tag, idx):
    cur_email, cur_prefix = email, prefix
    cur_domain = _email_domain(cur_email)
    taken_retry_count = 0
    for attempt in range(5):
        try:
            try:
                page.run_js_loaded("window.scrollTo(0, 0); return true;")
            except Exception:
                pass
            time.sleep(2)

            sel = (
                'css:input[type="email"], input[name="MemberName"], #MemberName, '
                '#usernameInput, input[name="Username"]'
            )
            email_el = _ele(page, sel, timeout=4)
            if email_el is None:
                log(f"  {tag} email input not found", "ERR")
                _shot(page, "no_email", idx)
                return None
            domain_sel = (
                'css:select[id="LiveDomainBoxList"], select[name="LiveDomainBoxList"], #LiveDomainBoxList'
            )
            dd = _ele(page, domain_sel, timeout=1)
            value = cur_prefix if dd is not None else cur_email
            if not _safe_input(email_el, value, clear=True):
                log(f"  {tag} email input failed, retrying", "WARN")
                _shot(page, "email_input_fail", idx)
                email_el = _ele(page, sel, timeout=2)
                if email_el is None or not _safe_input(email_el, value, clear=True):
                    cur_email, cur_prefix = _new_email_candidate("taken", cur_domain)
                    cur_domain = _email_domain(cur_email)
                    log(f"  {tag} email input still failed, retry: {cur_email}", "WARN")
                    continue

            if dd is not None:
                chosen = _select_domain(dd, cur_domain, tag) or cur_domain
                cur_domain = chosen
                cur_email = f"{cur_prefix}@{cur_domain}"
                log(f"  {tag} filled prefix: {cur_prefix}@{cur_domain}")
            else:
                log(f"  {tag} filled email: {cur_email}")
            time.sleep(2)
            _click_next(page, tag)
            time.sleep(3)

            kind = _email_error_kind(page)
            if kind == "taken":
                taken_retry_count += 1
                extra_digits = 3 if taken_retry_count == 1 else 1
                cur_email, cur_prefix = _new_email_candidate(
                    "taken", cur_domain, cur_email, cur_prefix, extra_digits
                )
                cur_domain = _email_domain(cur_email)
                log(
                    f"  {tag} email taken, append {extra_digits} digit(s), retry: {cur_email}",
                    "WARN",
                )
                time.sleep(3)
                continue
            if kind == "format":
                cur_email, cur_prefix = _new_email_candidate("format", cur_domain)
                cur_domain = _email_domain(cur_email)
                log(f"  {tag} format error, retry: {cur_email}", "WARN")
                time.sleep(3)
                continue
            _shot(page, "after_email", idx)
            return cur_email
        except Exception as exc:
            log(f"  {tag} email step exception ({type(exc).__name__}: {exc}), retrying", "WARN")
            _shot(page, "email_exc", idx)
            cur_email, cur_prefix = _new_email_candidate("taken", cur_domain)
            cur_domain = _email_domain(cur_email)
            time.sleep(2)
            continue
    log(f"  {tag} email step failed", "ERR")
    _shot(page, "email_fail", idx)
    return None


def _fill_password(page, password, tag, idx):
    sel = (
        'css:input[type="password"], input[name="Password"], '
        '#PasswordInput, input[name="passwd"]'
    )
    pwd_el = None
    for _ in range(10):
        pwd_el = _ele(page, sel, timeout=1)
        if pwd_el is not None:
            break
        time.sleep(1)
    if pwd_el is None:
        log(f"  {tag} 密码框未找到", "ERR")
        _shot(page, "pwd_fail", idx)
        return False
    if not _safe_input(pwd_el, password, clear=True):
        log(f"  {tag} 密码输入失败", "ERR")
        _shot(page, "pwd_fail", idx)
        return False
    log(f"  {tag} 密码已填")
    time.sleep(1)
    _click_next(page, tag)
    time.sleep(3)
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
        time.sleep(0.6)
        if _click_option_exact(candidates):
            time.sleep(0.4)
            _press_escape()
            return True
        if typed_fallback is not None:
            try:
                page.actions.type(str(typed_fallback)).press("\ue007").perform()
                time.sleep(0.4)
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
            'css:input[name="BirthYear"]',
            'css:input[aria-label*="Birth year" i]',
            'css:input[aria-label*="year" i]',
            'css:input[aria-label*="年"]',
            'css:#BirthYearInput',
            'css:input[id*="BirthYear" i]',
            'css:input[id*="Year" i]',
            'css:input[placeholder*="Year" i]',
            'css:input[placeholder*="year"]',
            'css:input[placeholder*="年"]',
            'css:input[type="number"]',
            'css:input[type="text"][inputmode="numeric"]',
        ]:
            yr = _ele(page, sel, timeout=1)
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
            time.sleep(0.3)

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
            time.sleep(0.3)
        else:
            log(f"  {tag} 未识别到 Day combobox, roles={roles}", "WARN")

        # 年份只走独立 input，绝不点 combos[-1]（那是 Day）
        year_ok = _fill_year_input()
        if not month_filled or not day_filled or not year_ok:
            log(
                f"  {tag} 生日字段状态: month={month_filled} day={day_filled} year={year_ok}",
                "WARN",
            )

    def _wait_after_birthday_submit(max_wait=3.0):
        deadline = time.time() + max_wait
        while time.time() < deadline:
            if _is_name_page(page):
                return "name"
            if not _is_birthday_page(page):
                return "left_birthday"
            time.sleep(0.2)
        return "timeout"

    def _submit_birthday_form():
        time.sleep(0.2)
        _click_next(page, tag)
        _wait_after_birthday_submit(3.0)
        if _is_birthday_page(page):
            low = _body_text(page).lower()
            if "birthdate" in low or "enter your birthdate" in low or "birth" in low:
                log(f"  {tag} still on birthday page after submit, refill year and retry", "WARN")
                _fill_year_input()
                time.sleep(0.2)
                _click_next(page, tag)
                _wait_after_birthday_submit(3.0)

    time.sleep(2)
    for _ in range(10):
        txt = _body_text(page)
        low = txt.lower()
        if any(k in low for k in ["birth", "country", "region", "naissance", "pays", "région", "details"]) or any(
            k in txt for k in ["出生", "国家", "地区", "年份", "详细信息"]
        ):
            break
        time.sleep(1)
    _shot(page, "bday_page", idx)

    for attempt in range(2):
        if attempt > 0:
            log(f"  {tag} 生日提交后仍报错，重选一次月/日/年再提交", "WARN")
            time.sleep(2)
        _apply_birthday_form()
        _submit_birthday_form()
        if not _is_birthday_page(page):
            _shot(page, "after_bday", idx)
            return True

    _shot(page, "after_bday", idx)
    if _is_birthday_page(page):
        low = _body_text(page).lower()
        log(f"  {tag} 生日页仍未通过: {low[:120]!r}", "WARN")
        _shot(page, "bday_fail", idx)
        return False
    return True


def _fill_name_and_terms(page, first, last, prefix, tag, idx):
    # If the name page is already visible, fill immediately; otherwise wait briefly
    # for post-birthday navigation. Avoid fixed sleeps before typing names.
    deadline = time.time() + 6
    while time.time() < deadline:
        if _is_name_page(page):
            break
        time.sleep(0.2)
    if not _is_name_page(page):
        log(f"  {tag} name page not detected, skip name fill", "WARN")
        _shot(page, "name_fail", idx)
        return False
    _page_start_wait(tag, "name")

    first_sel = (
        'css:input[name="FirstName"], #FirstName, input[name="firstNameInput"], '
        '#firstNameInput, css:input[aria-label*="first" i], '
        'css:input[aria-label*="pr?nom" i], css:input[aria-label*="?" i]'
    )
    last_sel = (
        'css:input[name="LastName"], #LastName, input[name="lastNameInput"], '
        '#lastNameInput, css:input[aria-label*="last" i], '
        'css:input[aria-label*="nom de famille" i], css:input[aria-label*="?" i]'
    )

    fe = le = None
    text_inputs = []
    field_deadline = time.time() + 4
    while time.time() < field_deadline:
        fe = _ele(page, first_sel, timeout=0.5)
        le = _ele(page, last_sel, timeout=0.5)
        if fe is not None or le is not None:
            break
        text_inputs = _eles(page, 'css:input[type="text"]', timeout=0.5)
        if len(text_inputs) >= 2:
            break
        time.sleep(0.2)

    if fe is not None:
        if le is not None:
            le.input(last, clear=True)
        fe.input(first, clear=True)
        log(f"  {tag} name: {first} {last}")
    elif len(text_inputs) >= 2:
        text_inputs[0].input(last, clear=True)
        text_inputs[1].input(first, clear=True)
        log(f"  {tag} name(generic): {first} {last}")
    else:
        # Only handle username/gamertag page after first/last fields are absent.
        uname_sel = (
            'css:input[id*="displayName" i], input[id*="gamertag" i], '
            'input[name*="displayName" i], input[aria-label*="gamertag" i]'
        )
        ue = _ele(page, uname_sel, timeout=1)
        txt = _body_text(page)
        low = txt.lower()
        if ue is not None and (
            any(k in low for k in ["gamertag", "display name", "pseudo", "surnom"])
            or any(k in txt for k in ["???", "????", "????"])
        ):
            username = prefix[:8] + str(random.randint(100, 999))
            ue.input(username, clear=True)
            log(f"  {tag} username: {username}")
            _submit_wait(tag)
            _click_next(page, tag)
            time.sleep(3)
            return True
        log(f"  {tag} name inputs not found", "WARN")
        _shot(page, "name_fail", idx)
        return False

    cb = _ele(page, 'css:input[type="checkbox"], [role="checkbox"]', timeout=1)
    if cb is not None:
        try:
            if not cb.is_checked:
                cb.click_self()
                log(f"  {tag} checked terms")
        except Exception:
            try:
                cb.click_self(by_js=True)
            except Exception:
                pass
    _submit_wait(tag)
    _click_next(page, tag)
    time.sleep(3)
    _shot(page, "after_name", idx)
    return True


def _find_hold_target(ctx):
    """在单个 browsing context 内找 Press-and-hold 按钮。

    返回带 quality 分的 dict：
      0 = #px-captcha 真按钮（最佳）
      1 = 小尺寸 button/role=button
      2 = 其它可点元素
      9 = 过大容器/说明文案（应尽量不用）
    第一次按压常失败的根因：iframe 未就绪时主文档命中大 div「Press and hold the button.」说明文字。
    """
    script = r"""
return (() => {
  const PRESS_RE = /press\s*and\s*hold|appuyer\s*et\s*maintenir|按住|长按|halten/i;
  const visible = (el) => {
    if (!el) return false;
    const s = getComputedStyle(el);
    if (s.display === 'none' || s.visibility === 'hidden' || Number(s.opacity) === 0) return false;
    const r = el.getBoundingClientRect();
    return r.width > 20 && r.height > 10 && r.bottom > 0 && r.right > 0;
  };
  const pack = (el, quality, source) => {
    const r = el.getBoundingClientRect();
    const text = (el.innerText || el.textContent || el.value || '').trim().slice(0, 80);
    return {
      x: Math.round(r.x + r.width / 2),
      y: Math.round(r.y + r.height / 2),
      left: Math.round(r.x),
      top: Math.round(r.y),
      width: Math.round(r.width),
      height: Math.round(r.height),
      text,
      id: el.id || '',
      tag: (el.tagName || '').toLowerCase(),
      quality,
      source: source || ''
    };
  };
  const buttonish = (el) => {
    const tag = (el.tagName || '').toLowerCase();
    if (tag === 'button' || tag === 'input') return true;
    const role = (el.getAttribute('role') || '').toLowerCase();
    return role === 'button';
  };
  const score = (el) => {
    const id = (el.id || '').toLowerCase();
    const r = el.getBoundingClientRect();
    const text = (el.innerText || el.textContent || el.value || '').trim();
    if (id === 'px-captcha') return 0;
    // 真按钮通常是矮扁：高 20-80、宽 80-420
    const sizeOk = r.height >= 18 && r.height <= 90 && r.width >= 60 && r.width <= 480;
    if (buttonish(el) && sizeOk && PRESS_RE.test(text)) return 1;
    if (buttonish(el) && sizeOk) return 1;
    if (sizeOk && PRESS_RE.test(text) && text.length <= 40) return 2;
    // 过大容器/整段说明文（含 robot 图+文案）质量很差
    if (r.height > 120 || r.width > 520 || text.length > 60) return 9;
    if (PRESS_RE.test(text)) return 5;
    return 8;
  };

  const candidates = [];
  // 1) 最高优先：真按钮 #px-captcha
  const px = document.querySelector('#px-captcha');
  if (px && visible(px)) candidates.push(pack(px, 0, '#px-captcha'));

  // 2) button / role=button / input
  for (const sel of ['button', '[role="button"]', 'input[type="button"]', 'input[type="submit"]']) {
    for (const el of document.querySelectorAll(sel)) {
      if (!visible(el)) continue;
      const text = (el.innerText || el.textContent || el.value || '').trim();
      const id = (el.id || '').toLowerCase();
      if (id === 'px-captcha' || PRESS_RE.test(text)) {
        candidates.push(pack(el, score(el), sel));
      }
    }
  }

  // 3) 小范围 div（排除超大说明区）
  for (const el of document.querySelectorAll('div, span, a')) {
    if (!visible(el)) continue;
    const text = (el.innerText || el.textContent || '').trim();
    if (!PRESS_RE.test(text)) continue;
    // 只要短文案节点，避免整卡容器
    if (text.length > 48) continue;
    const r = el.getBoundingClientRect();
    if (r.height > 100 || r.width > 500) continue;
    candidates.push(pack(el, score(el), el.tagName.toLowerCase()));
  }

  if (!candidates.length) return null;
  candidates.sort((a, b) => (a.quality - b.quality) || (a.width * a.height - b.width * b.height));
  return candidates[0];
})()
"""
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


def _captcha_is_validating(page):
    if _px_captcha_completed_wait(page):
        return True
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
    if _context_has_iframe_hint(page) and not _find_hold_context(page)[1]:
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


def _try_submit(page, tag):
    hit = _click_any(page, ['#iSignupAction', 'css:input[type="submit"]', 'css:button[type="submit"]'], timeout=1)
    if hit:
        log(f"  {tag} 提交推动: {hit}")
        time.sleep(2)
    return bool(hit)


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
    """Original hold chain: move_to -> hold -> wait -> release."""
    cx = int(target.get("x", 0) + random.uniform(-3, 3))
    cy = int(target.get("y", 0) + random.uniform(-2, 2))
    hold_sec = random.uniform(10.0, 13.0)
    log(
        f"  {tag} press #{press_count}: ({cx},{cy}) hold={hold_sec:.1f}s"
        + (f" text={target.get('text', '')[:30]!r}" if target.get("text") else "")
    )
    try:
        ctx.actions.move_to(
            {"x": cx, "y": cy}, duration=random.randint(250, 550)
        ).hold().wait(hold_sec).release().perform()
        time.sleep(random.uniform(2, 4))
        return True
    except Exception as exc:
        log(f"  {tag} hold failed: {type(exc).__name__}: {exc}", "WARN")
        return False


def _proxy_for_ip_lookup(proxy_pool, tag):
    if not proxy_pool:
        return None
    proxy_str = str(proxy_pool[0] or "").strip()
    if not proxy_str:
        return None
    try:
        host, port, user, pwd = proxy_str.split(":", 3)
    except Exception:
        log(f"  {tag} IP probe proxy format invalid: {proxy_str[:80]!r}", "WARN")
        return None
    auth = f"{quote(user, safe='')}:{quote(pwd, safe='')}"
    proxy_url = f"socks5h://{auth}@{host}:{port}"
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
    """把 CLI/opts 的账号格式、后缀配置落到 helpers 环境变量。"""
    helpers = _load_helpers()
    opts = opts or SimpleNamespace()

    email_suffixes = (
        getattr(opts, "email_suffixes", None)
        or os.environ.get("OUTLOOK_ACCOUNT_SUFFIXES")
        or os.environ.get("OUTLOOK_EMAIL_SUFFIXES")
        or ""
    )
    if email_suffixes:
        # 规范化：支持 multi 列表/逗号串
        if isinstance(email_suffixes, (list, tuple, set)):
            email_suffixes = ",".join(str(x) for x in email_suffixes if str(x).strip())
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
    os.environ["OUTLOOK_ACCOUNT_FORMAT_MODE"] = str(mode)
    apply_fn = getattr(helpers, "apply_account_format_mode", None)
    if callable(apply_fn):
        apply_fn(mode, custom)
    elif custom and str(mode).lower() == "custom":
        os.environ["OUTLOOK_ACCOUNT_FORMAT"] = str(custom)

    password_format = (
        getattr(opts, "password_format", None)
        or os.environ.get("OUTLOOK_PASSWORD_FORMAT")
        or ""
    )
    if password_format:
        os.environ["OUTLOOK_PASSWORD_FORMAT"] = str(password_format)

    log(
        f"账号格式: mode={os.environ.get('OUTLOOK_ACCOUNT_FORMAT_MODE', mode)} "
        f"format={os.environ.get('OUTLOOK_ACCOUNT_FORMAT', '')!r} "
        f"suffixes={os.environ.get('OUTLOOK_ACCOUNT_SUFFIXES', 'outlook.com')!r} "
        f"pwd_format={os.environ.get('OUTLOOK_PASSWORD_FORMAT', '')!r}"
    )


def register_outlook(opts, proxy_pool, idx):
    from ruyipage import FirefoxOptions, FirefoxPage

    helpers = _load_helpers()
    _apply_account_options(opts)
    generate_email_password = helpers.generate_email_password
    generate_birthday = helpers.generate_birthday
    generate_name = helpers.generate_name
    verify_registered_outlook = helpers.verify_registered_outlook

    tag = f"[#{idx}][ruoyi]"
    timeout = int(getattr(opts, "timeout", REGISTER_TIMEOUT) or REGISTER_TIMEOUT)
    max_press = int(getattr(opts, "max_press", os.environ.get("OUTLOOK_REG_MAX_PRESS", "5")) or 5)
    need_verify = not bool(getattr(opts, "no_verify", False))
    confirm_before_register = bool(getattr(opts, "confirm_before_register", False))
    is_headless = bool(getattr(opts, "headless", False))

    tb = FirefoxOptions()
    tb.set_browser_path(RUOYI_FIREFOX_PATH)
    if proxy_pool:
        tb.set_per_tab_proxies(proxy_pool, exhausted="wrap")
        log(f"挂载 {len(proxy_pool)} 条 SOCKS5 代理到 per-tab 池(wrap 轮换)")
    else:
        log("没挂代理——直接本机出口", "WARN")
    if is_headless:
        _apply_ruoyi_headless_options(tb, tag)
        tb.headless(True)

    log(f"启动 ruyipage Firefox: {RUOYI_FIREFOX_PATH}")
    browser_page = None
    page = None
    email = password = None
    success = False
    har_collector = None
    failure_reason = "failure"
    deadline = time.time() + timeout
    press_wait_started = None
    had_captcha = False
    gone_rounds = 0
    press_count = 0
    no_target_rounds = 0
    initial_press_wait_started = None
    validation_wait_started = None
    microsoft_loading_wait_started = None
    # 按压成功后：等 loading/消失，再等 captcha 重新出现后才允许下一次按压
    awaiting_reappear = False
    post_press_saw_gap = False
    post_press_started_at = None
    headless_patch_logged = False
    signup_opened = False

    try:
        browser_page = FirefoxPage(tb)
        page = browser_page

        if proxy_pool:
            try:
                if is_headless:
                    page = browser_page.new_container_tab()
                    log(f"  {tag} 使用 Firefox container tab 承载注册页，以匹配 ruyipage per-tab 代理分配")
                else:
                    page = browser_page.new_container_tab(url=SIGNUP_URL)
                    signup_opened = True
                    log(f"  {tag} 使用 Firefox container tab 直接打开注册页，以匹配 ruyipage per-tab 代理分配")
                try:
                    browser_page.close_other_tabs(page)
                    log(f"  {tag} 已关闭 Firefox 启动默认空白页，仅保留注册页")
                except Exception as close_exc:
                    log(f"  {tag} 关闭默认空白页失败: {type(close_exc).__name__}: {close_exc}", "WARN")
            except Exception as exc:
                log(f"  {tag} 创建 container tab 失败，停止本次注册避免默认 tab 误走非 per-tab 代理: {exc}", "ERR")
                return None, None

        if is_headless and not signup_opened:
            _apply_ruoyi_headless_page_patches(page, tag, log_once=True)
            headless_patch_logged = True
        if is_headless:
            har_collector = _RuoyiHarCollector(page, tag, idx, email_getter=lambda: email or "")
            har_collector.start()
        if not signup_opened:
            page.get(SIGNUP_URL)
        page.wait_loading(20)
        if is_headless:
            _apply_ruoyi_headless_page_patches(page, tag, log_once=False)
        log(f"页面 title={page.title!r} url={page.url}")
        _log_current_ip_async(proxy_pool, tag)
        _shot(page, "start", idx)

        if confirm_before_register:
            _click_post_signup(page, tag)
            time.sleep(3)
        _handle_consent(page, tag, idx)

        email, password, prefix = generate_email_password()
        log(f"  {tag} 将注册: {email}")
        ok_email = _fill_email(page, email, prefix, tag, idx)
        if not ok_email:
            return None, None
        email = ok_email

        if not _fill_password(page, password, tag, idx):
            return None, None

        year, month, day = generate_birthday()
        if not _fill_birthday(page, year, month, day, tag, idx):
            return None, None

        first, last = generate_name()
        if not _fill_name_and_terms(page, first, last, prefix, tag, idx):
            return None, None

        while time.time() < deadline:
            if is_headless:
                _apply_ruoyi_headless_page_patches(page, tag, log_once=(not headless_patch_logged))
                headless_patch_logged = True
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
                return None, None

            if _maybe_skip_passkey(page, tag):
                continue
            if "privacynotice" in current_url:
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
            if _microsoft_loading_page(page):
                if microsoft_loading_wait_started is None:
                    microsoft_loading_wait_started = time.time()
                    log(f"  {tag} Microsoft Loading page, keep waiting for redirect")
                elif int(time.time() - microsoft_loading_wait_started) % 15 < 3:
                    waited = int(time.time() - microsoft_loading_wait_started)
                    log(f"  {tag} Microsoft Loading still active, waited {waited}s")
                time.sleep(3)
                continue
            microsoft_loading_wait_started = None

            if awaiting_reappear and press_count < max_press:
                gap_waited = time.time() - (post_press_started_at or time.time())
                if not actionable:
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
                    initial_press_wait_started = time.time()
                    gone_rounds = 0
                    log(f"  {tag} captcha reappeared, wait {INITIAL_PRESS_DELAY}s before next press")
                    time.sleep(1)
                    continue
                if gap_waited < POST_PRESS_LOADING_CHECK:
                    time.sleep(0.5)
                    continue
                awaiting_reappear = False
                post_press_saw_gap = False
                post_press_started_at = None
                initial_press_wait_started = time.time() - INITIAL_PRESS_DELAY
                gone_rounds = 0
                if POST_PRESS_RETRY_GAP > 0:
                    log(f"  {tag} no loading after {POST_PRESS_LOADING_CHECK}s, retry in {POST_PRESS_RETRY_GAP}s")
                    time.sleep(POST_PRESS_RETRY_GAP)
                else:
                    log(f"  {tag} no loading after {POST_PRESS_LOADING_CHECK}s, retry now")
                continue

            if had_captcha and (not visible or validating or not actionable):
                if press_count >= max_press:
                    if press_wait_started is None:
                        press_wait_started = time.time()
                    waited = time.time() - press_wait_started
                    if waited >= POST_MAX_PRESS_WAIT:
                        log(f"  {tag} no redirect after {max_press} presses and {POST_MAX_PRESS_WAIT}s, give up", "WARN")
                        _shot(page, "press_fail", idx)
                        return None, None
                    time.sleep(1)
                    continue
                gone_rounds += 1
                if validation_wait_started is None:
                    validation_wait_started = time.time()
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
                            return None, None
                    else:
                        no_target_rounds = 0
                        press_count += 1
                        _focus_page_before_captcha_press(page, tag)
                        if _perform_hold(page, ctx, target, idx, press_count, tag):
                            validation_wait_started = time.time()
                            awaiting_reappear = True
                            post_press_saw_gap = False
                            post_press_started_at = time.time()
                            if press_count >= max_press:
                                press_wait_started = time.time()
                            continue
                        awaiting_reappear = True
                        post_press_saw_gap = False
                        post_press_started_at = time.time()
                if press_count >= max_press:
                    if press_wait_started is None:
                        press_wait_started = time.time()
                        log(f"  {tag} max press {max_press} reached, wait up to {POST_MAX_PRESS_WAIT}s")
                    waited = time.time() - press_wait_started
                    if waited >= POST_MAX_PRESS_WAIT:
                        log(f"  {tag} no redirect after {max_press} presses and {POST_MAX_PRESS_WAIT}s, give up", "WARN")
                        _shot(page, "press_fail", idx)
                        return None, None
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
                            return None, None
                        time.sleep(1)
                        continue
                    if awaiting_reappear:
                        post_press_saw_gap = True
                        time.sleep(2)
                        continue
                    if validation_wait_started is None:
                        validation_wait_started = time.time()
                        log(f"  {tag} post-captcha state unclear, keep waiting")
                    time.sleep(3)
                    continue
                validation_wait_started = None
                _try_submit(page, tag)

            if int(time.time()) % 15 < 3:
                _shot(page, f"wait_{int(time.time() % 1000)}", idx)
            time.sleep(3)
        else:
            log(f"  {tag} captcha timeout", "WARN")
            _shot(page, "timeout", idx)
            return None, None

        _post_signup_cleanup(page, tag, idx)
        if need_verify and not verify_registered_outlook(email, password, tag):
            log(f"  {tag} verification failed, discarding account", "WARN")
            return None, None

        log(f"  {tag} OK: {email} / {password}", "OK")
        success = True
        return email, password
    except Exception as e:
        failure_reason = f"exception_{type(e).__name__}"
        log(f"  {tag} FAILED: {type(e).__name__}: {e}", "ERR")
        try:
            if page is not None:
                _shot(page, "error", idx)
        except Exception:
            pass
        return None, None
    finally:
        if har_collector is not None:
            if is_headless and not success:
                try:
                    har_collector.save(reason=failure_reason)
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
                browser_page.quit()
        except Exception:
            pass


def _save_direct_result(email, password, graph, live_file, token_file):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(live_file, "a", encoding="utf-8") as f:
        f.write(f"{email}----{password}----{graph['refresh_token']}----{graph.get('client_id', '')}\n")
    if token_file:
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


def main():
    ap = argparse.ArgumentParser(description="Outlook 自注册养号(ruoyi) — 完整链路版")
    ap.add_argument("--proxy-file", "-p", default=PROXY_FILE, help=f"代理池文件(默认 {PROXY_FILE})")
    ap.add_argument("--count", "-n", type=int, default=1, help="注册次数(默认 1)")
    ap.add_argument("--headless", action="store_true", help="无头模式")
    ap.add_argument("--timeout", "-t", type=int, default=REGISTER_TIMEOUT, help="单号超时(秒)")
    ap.add_argument("--max-press", default=os.environ.get("OUTLOOK_REG_MAX_PRESS", "5"), help="按住次数上限")
    ap.add_argument("--no-verify", action="store_true", help="跳过 Outlook 登录校验")
    ap.add_argument("--confirm-before-register", action="store_true", help="页面打开后先尝试点确认")
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

    if not os.path.isfile(RUOYI_FIREFOX_PATH):
        log(f"定制 Firefox 内核不存在: {RUOYI_FIREFOX_PATH}", "ERR")
        log("请先运行: .venv\\Scripts\\python.exe -m ruyipage install", "ERR")
        sys.exit(1)

    helpers = _load_helpers()
    _apply_account_options(args)
    try:
        proxy_env = helpers.ensure_clash_proxy_env()
        if proxy_env:
            log(f"proxy env ready: {proxy_env}")
    except Exception:
        pass

    proxy_pool = parse_proxy_pool(args.proxy_file) if args.proxy_file else []
    log(f"代理池准备完毕: {len(proxy_pool)} 条")

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    live_file = os.path.join(OUTPUT_DIR, f"accounts_ruoyi_{ts}.txt")
    token_file = os.path.join(OUTPUT_DIR, f"graph_tokens_ruoyi_{ts}.json")

    ok = 0
    for i in range(args.count):
        log(f"========== 注册 #{i + 1}/{args.count} ==========")
        selected_pool = select_proxy_for_account(proxy_pool)
        if selected_pool:
            masked = selected_pool[0].split(":")
            masked_proxy = f"{masked[0]}:{masked[1]}:{masked[2]}:***" if len(masked) == 4 else "***"
            log(f"#{i + 1} 随机代理 -> {masked_proxy}")
        email, password = register_outlook(args, selected_pool, i + 1)
        if not email:
            log(f"#{i + 1} 结果: FAIL", "WARN")
            continue
        log(f"#{i + 1} Graph token extracting…")
        graph = helpers.extract_graph_token_http(email, password, i + 1)
        if not graph or not graph.get("refresh_token"):
            log(f"#{i + 1} registered but graph RT missing; not saved: {email}", "WARN")
            continue
        _save_direct_result(email, password, graph, live_file, token_file)
        ok += 1
        log(f"#{i + 1} 结果: OK {email}", "OK")

    log(f"完成: success={ok}/{args.count}")
    if ok:
        log(f"账号输出: {live_file}")
        log(f"Token 输出: {token_file}")


if __name__ == "__main__":
    main()
