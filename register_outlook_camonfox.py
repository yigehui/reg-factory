#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook 自注册养号(camonfox) —— 完整链路版

目标：
  - 使用 camoufox.sync_api 原生 Camoufox 完成 Outlook 自注册
  - 产出可登录账号，并在直跑模式下补抽 Graph refresh_token
  - 可被 outlook_reg_loop.py 作为完整后端调用

说明：
  - 与 ruoyi 版本同链路，但浏览器后端改为原生 Camoufox，不再回退 ruyipage+exe。
  - 本文件保留独立直跑能力，同时暴露 register_outlook() 给 loop 调用。
"""

from __future__ import annotations

import argparse
import select
import importlib.util
import json
import os
import random
import requests
import socket
import struct
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from urllib.parse import quote, unquote, urlparse

ROOT = os.path.dirname(os.path.abspath(__file__))
ENGINE_NAME = "camonfox"
PROXY_FILE = os.environ.get("OUTLOOK_PROXY_FILE", "proxies_outlook.txt")
SCREENSHOT_DIR = os.path.join(ROOT, "screenshots_camonfox")
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
HEADLESS_WINDOW_WIDTH = int(os.environ.get("OUTLOOK_CAMONFOX_HEADLESS_WIDTH", "1280") or "1280")
HEADLESS_WINDOW_HEIGHT = int(os.environ.get("OUTLOOK_CAMONFOX_HEADLESS_HEIGHT", "800") or "800")
HEADLESS_USER_AGENT = os.environ.get(
    "OUTLOOK_CAMONFOX_HEADLESS_UA",
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


def _normalize_proxy_entry(raw, source="proxy"):
    """Normalize one proxy entry.

    支持：
      - http://user:pass@host:port
      - socks5://user:pass@host:port
      - user:pass@host:port
      - host:port:user:pass
    """
    ln = str(raw or "").strip()
    if not ln or ln.startswith("#"):
        return ""
    explicit_scheme = ""
    for pfx in ("socks5://", "http://", "https://"):
        if ln.lower().startswith(pfx):
            explicit_scheme = pfx[:-3]
            ln = ln[len(pfx):]
            break
    if "@" in ln:
        auth, hostport = ln.rsplit("@", 1)
        if ":" not in auth or ":" not in hostport:
            log(f"跳过非法代理({source}, 缺 user/pass 或 host/port): {ln[:60]}", "WARN")
            return ""
        user, pwd = auth.split(":", 1)
        host, port = hostport.rsplit(":", 1)
        if explicit_scheme:
            return f"{explicit_scheme}://{quote(user, safe='')}:{quote(pwd, safe='')}@{host}:{port}"
        return f"{host}:{port}:{user}:{pwd}"
    parts = ln.split(":", 3)
    if len(parts) == 4:
        host, port, user, pwd = parts
        if explicit_scheme:
            return f"{explicit_scheme}://{quote(user, safe='')}:{quote(pwd, safe='')}@{host}:{port}"
        return f"{host}:{port}:{user}:{pwd}"
    if len(parts) == 2:
        host, port = parts
        if explicit_scheme:
            return f"{explicit_scheme}://{host}:{port}"
        return f"{host}:{port}"
    log(f"跳过非法代理({source}): {ln[:60]}", "WARN")
    return ""


def parse_proxy_value(value):
    item = _normalize_proxy_entry(value, source="proxy-ip")
    return [item] if item else []


def parse_proxy_pool(path):
    """读代理池文件，转成本项目统一格式；显式 http/socks5 协议会保留。"""
    if not path:
        return []
    if not os.path.isfile(path):
        log(f"代理池文件不存在: {path}", "WARN")
        return []
    out = []
    with open(path, "r", encoding="utf-8") as f:
        for ln in f:
            item = _normalize_proxy_entry(ln, source=os.path.basename(path))
            if item:
                out.append(item)
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
    base = os.path.join(SCREENSHOT_DIR, f"camonfox_{idx}_{name}_{ts}")
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
    path = os.path.join(SCREENSHOT_DIR, f"camonfox_{idx}_{name}_{ts}.png")
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


def _ele(page, locator, timeout=None):
    try:
        el = page.ele(locator, timeout=timeout)
        if el is None:
            return None
        if el.__class__.__name__ == "NoneElement":
            return None
        if not bool(el):
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




# ---------------------------------------------------------------------------
# Native Camoufox backend (camoufox.sync_api)
# ---------------------------------------------------------------------------

class CamoufoxNotAvailable(RuntimeError):
    pass


def _mask_proxy_raw(proxy):
    raw = str(proxy or "").strip()
    if not raw:
        return ""
    if "://" in raw:
        try:
            p = urlparse(raw)
            auth = "***@" if p.username else ""
            return f"{p.scheme}://{auth}{p.hostname}:{p.port or ''}"
        except Exception:
            return "***"
    parts = raw.split(":")
    if len(parts) >= 4:
        return f"{parts[0]}:{parts[1]}:{parts[2]}:***"
    return raw[:80]




class _Socks5AuthForwarder:
    """Local no-auth SOCKS5 -> upstream SOCKS5 with username/password.

    Playwright Firefox/Camoufox cannot launch with authenticated SOCKS5
    directly. This shim lets Camoufox use socks5://127.0.0.1:port without
    auth, while the shim authenticates to the real upstream proxy.
    """

    def __init__(self, host, port, username="", password="", tag=""):
        self.host = str(host)
        self.port = int(port)
        self.username = str(username or "")
        self.password = str(password or "")
        self.tag = tag or ""
        self._stop = threading.Event()
        self._sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self._sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self._sock.bind(("127.0.0.1", 0))
        self._sock.listen(64)
        self._sock.settimeout(0.5)
        self.local_host, self.local_port = self._sock.getsockname()
        self._thread = threading.Thread(target=self._serve, name="camonfox-socks5-forwarder", daemon=True)
        self._thread.start()

    @property
    def server(self):
        return f"socks5://{self.local_host}:{self.local_port}"

    def close(self):
        self._stop.set()
        try:
            self._sock.close()
        except Exception:
            pass
        try:
            self._thread.join(timeout=2)
        except Exception:
            pass

    def _serve(self):
        while not self._stop.is_set():
            try:
                client, _addr = self._sock.accept()
            except socket.timeout:
                continue
            except OSError:
                break
            except Exception:
                continue
            threading.Thread(target=self._handle, args=(client,), daemon=True).start()

    @staticmethod
    def _recv_exact(sock, n):
        data = b""
        while len(data) < n:
            chunk = sock.recv(n - len(data))
            if not chunk:
                raise ConnectionError("short read")
            data += chunk
        return data

    def _handle(self, client):
        upstream = None
        try:
            client.settimeout(20)
            head = self._recv_exact(client, 2)
            if head[0] != 5:
                return
            n_methods = head[1]
            if n_methods:
                self._recv_exact(client, n_methods)
            client.sendall(b"\x05\x00")

            req = self._recv_exact(client, 4)
            ver, cmd, _rsv, atyp = req
            if ver != 5 or cmd != 1:
                client.sendall(b"\x05\x07\x00\x01\x00\x00\x00\x00\x00\x00")
                return
            if atyp == 1:
                dst_host = socket.inet_ntoa(self._recv_exact(client, 4))
            elif atyp == 3:
                ln = self._recv_exact(client, 1)[0]
                dst_host = self._recv_exact(client, ln).decode("idna", errors="ignore")
            elif atyp == 4:
                dst_host = socket.inet_ntop(socket.AF_INET6, self._recv_exact(client, 16))
            else:
                client.sendall(b"\x05\x08\x00\x01\x00\x00\x00\x00\x00\x00")
                return
            dst_port = struct.unpack("!H", self._recv_exact(client, 2))[0]

            import socks

            upstream = socks.socksocket()
            upstream.set_proxy(
                socks.SOCKS5,
                self.host,
                self.port,
                rdns=True,
                username=self.username or None,
                password=self.password or None,
            )
            upstream.settimeout(30)
            upstream.connect((dst_host, dst_port))
            client.sendall(b"\x05\x00\x00\x01\x00\x00\x00\x00\x00\x00")
            client.settimeout(None)
            upstream.settimeout(None)
            self._relay(client, upstream)
        except Exception:
            try:
                client.sendall(b"\x05\x01\x00\x01\x00\x00\x00\x00\x00\x00")
            except Exception:
                pass
        finally:
            for s in (client, upstream):
                try:
                    if s is not None:
                        s.close()
                except Exception:
                    pass

    def _relay(self, a, b):
        sockets = [a, b]
        while not self._stop.is_set():
            try:
                readable, _, _ = select.select(sockets, [], [], 60)
            except Exception:
                break
            if not readable:
                continue
            for src in readable:
                dst = b if src is a else a
                try:
                    data = src.recv(65536)
                    if not data:
                        return
                    dst.sendall(data)
                except Exception:
                    return


def _parse_proxy_parts(proxy):
    raw = str(proxy or "").strip()
    if not raw:
        return None
    explicit_scheme = ""
    if "://" in raw:
        parsed = urlparse(raw)
        if not parsed.hostname:
            return None
        return {
            "scheme": parsed.scheme or "socks5",
            "host": parsed.hostname,
            "port": parsed.port or (443 if parsed.scheme == "https" else 1080),
            "username": unquote(parsed.username or ""),
            "password": unquote(parsed.password or ""),
            "explicit_scheme": True,
        }
    parts = raw.split(":", 3)
    if len(parts) == 4 and "@" not in raw:
        host, port, user, pwd = parts
        return {"scheme": "socks5", "host": host, "port": int(port), "username": user, "password": pwd, "explicit_scheme": False}
    if "@" in raw:
        auth, hostport = raw.rsplit("@", 1)
        user, pwd = (auth.split(":", 1) + [""])[:2]
        host, port = hostport.rsplit(":", 1)
        return {"scheme": "socks5", "host": host, "port": int(port), "username": user, "password": pwd, "explicit_scheme": False}
    if ":" in raw:
        host, port = raw.rsplit(":", 1)
        return {"scheme": "socks5", "host": host, "port": int(port), "username": "", "password": "", "explicit_scheme": False}
    return None


def _proxy_to_camoufox_with_forwarder(proxy, tag=""):
    info = _parse_proxy_parts(proxy)
    if not info:
        return None, None
    scheme = (info.get("scheme") or "socks5").lower()
    user = info.get("username") or ""
    pwd = info.get("password") or ""
    if scheme.startswith("socks") and (user or pwd):
        fw = _Socks5AuthForwarder(info["host"], info["port"], user, pwd, tag=tag)
        log(f"  {tag} SOCKS5 auth proxy bridged locally: {fw.server} -> {info['host']}:{info['port']}:{user}:***")
        return {"server": fw.server}, fw
    cfg = {"server": f"{scheme}://{info['host']}:{info['port']}"}
    if user:
        cfg["username"] = user
    if pwd:
        cfg["password"] = pwd
    return cfg, None
def _proxy_to_camoufox(proxy):
    """Convert this repo's proxy line to Camoufox/Playwright proxy config."""
    raw = str(proxy or "").strip()
    if not raw:
        return None
    if "://" not in raw:
        parts = raw.split(":", 3)
        if len(parts) == 4 and "@" not in raw:
            host, port, user, pwd = parts
            return {
                "server": f"socks5://{host}:{port}",
                "username": user,
                "password": pwd,
            }
        # user:pass@host:port / host:port both default to SOCKS5 in this flow.
        raw = "socks5://" + raw
    parsed = urlparse(raw)
    if not parsed.hostname:
        return None
    scheme = parsed.scheme or "socks5"
    port = parsed.port or (443 if scheme == "https" else 1080)
    cfg = {"server": f"{scheme}://{parsed.hostname}:{port}"}
    if parsed.username:
        cfg["username"] = unquote(parsed.username)
    if parsed.password:
        cfg["password"] = unquote(parsed.password)
    return cfg


def _proxy_to_requests_url(proxy):
    raw = str(proxy or "").strip()
    if not raw:
        return ""
    if "://" in raw:
        return raw
    parts = raw.split(":", 3)
    if len(parts) == 4 and "@" not in raw:
        host, port, user, pwd = parts
        return f"socks5h://{quote(user, safe='')}:{quote(pwd, safe='')}@{host}:{port}"
    return "socks5h://" + raw


def _fetch_exit_ip_for_camoufox(proxy, timeout=8.0):
    url = _proxy_to_requests_url(proxy)
    if not url:
        return ""
    session = requests.Session()
    session.trust_env = False
    try:
        resp = session.get("https://api.ipify.org", proxies={"http": url, "https": url}, timeout=timeout)
        ip = (resp.text or "").strip()
        if ip and all(ch.isdigit() or ch in ".:abcdefABCDEF" for ch in ip) and 3 <= len(ip) <= 45:
            return ip
    except Exception:
        return ""
    return ""


def _resolve_native_camoufox_exe():
    """Best-effort path lookup only for logging; launch still goes through camoufox.sync_api."""
    name = "camoufox.exe" if sys.platform.startswith("win") else "camoufox"
    found = []
    try:
        from camoufox import pkgman
        try:
            root = pkgman.camoufox_path(download_if_missing=False)
            p = Path(root) / name if root else None
            if p:
                found.append(Path(p))
        except Exception:
            pass
        for attr in ("INSTALL_DIR",):
            val = getattr(pkgman, attr, None)
            if val:
                root = Path(val)
                if root.exists():
                    found.extend(root.rglob(name))
    except Exception:
        pass
    local = os.environ.get("LOCALAPPDATA") or ""
    if local:
        root = Path(local) / "camoufox"
        if root.exists():
            try:
                found.extend(root.rglob(name))
            except Exception:
                pass
    uniq = []
    seen = set()
    for p in found:
        try:
            if not p.is_file():
                continue
            key = str(p).lower()
            if key in seen:
                continue
            seen.add(key)
            uniq.append(p)
        except Exception:
            continue
    if not uniq:
        return ""
    try:
        return str(max(uniq, key=lambda p: p.stat().st_mtime))
    except Exception:
        return str(uniq[0])


def _fetch_camoufox_binary():
    log("first Camoufox use: running python -m camoufox fetch ...")
    proc = subprocess.Popen(
        [sys.executable, "-m", "camoufox", "fetch"],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
    )
    lines = []
    deadline = time.time() + 60 * 30
    last_ping = 0.0
    try:
        while True:
            if time.time() > deadline:
                proc.kill()
                raise CamoufoxNotAvailable("camoufox fetch timeout")
            line = proc.stdout.readline() if proc.stdout else ""
            if line:
                msg = line.rstrip()
                lines.append(msg)
                if msg and ("download" in msg.lower() or "fetch" in msg.lower() or "%" in msg or "error" in msg.lower()):
                    log(f"Camoufox fetch: {msg[:180]}")
            elif proc.poll() is not None:
                break
            else:
                now = time.time()
                if now - last_ping > 15:
                    log("Camoufox fetch still running, please wait ...")
                    last_ping = now
                time.sleep(0.5)
        code = proc.wait(timeout=10)
    except Exception:
        try:
            proc.kill()
        except Exception:
            pass
        raise
    if code != 0:
        tail = "\n".join(lines[-20:])
        raise CamoufoxNotAvailable(f"camoufox fetch failed(code={code}): {tail[:500]}")


def _ensure_native_camoufox_ready():
    try:
        from camoufox.sync_api import Camoufox  # noqa: F401
    except ImportError as exc:
        raise CamoufoxNotAvailable('camoufox is not installed; run: pip install "camoufox[geoip]"') from exc
    exe = _resolve_native_camoufox_exe()
    if exe:
        return exe
    _fetch_camoufox_binary()
    exe = _resolve_native_camoufox_exe()
    if not exe:
        raise CamoufoxNotAvailable("Camoufox browser binary not found; run: python -m camoufox fetch")
    return exe


def _pw_key(key):
    mapping = {
        "\ue007": "Enter",
        "\ue00c": "Escape",
        "\ue003": "Backspace",
        "\ue004": "Tab",
    }
    return mapping.get(key, key)


def _dp_selector_to_css(raw):
    raw = str(raw or "").strip()
    if raw.startswith("css:"):
        return raw[4:].strip()
    if raw.startswith("@") and "=" in raw:
        key, val = raw[1:].split("=", 1)
        return f"[{key.strip()}=\"{val.strip().strip(chr(34)).strip(chr(39))}\"]"
    return raw


class _NativeSelectProxy:
    def __init__(self, element):
        self._element = element

    def by_value(self, value):
        return self._element._locator.select_option(value=str(value), timeout=2500)

    def by_text(self, text):
        return self._element._locator.select_option(label=str(text), timeout=2500)


class _NativeElement:
    def __init__(self, ctx, locator):
        self._ctx = ctx
        self._locator = locator.first
        self.select = _NativeSelectProxy(self)

    def _handle(self):
        return self._locator.element_handle(timeout=3000)

    def click_self(self, by_js=False):
        if by_js:
            return self.run_js("this.click(); return true;")
        self._locator.scroll_into_view_if_needed(timeout=3000)
        self._locator.click(timeout=5000, force=True)
        return True

    def input(self, text, clear=True):
        val = "" if text is None else str(text)
        self._locator.scroll_into_view_if_needed(timeout=3000)
        if clear:
            try:
                self._locator.fill(val, timeout=5000, force=True)
                return True
            except Exception:
                pass
        self._locator.click(timeout=3000, force=True)
        if clear:
            try:
                self._ctx._page.keyboard.press("Control+A")
                self._ctx._page.keyboard.press("Backspace")
            except Exception:
                pass
        self._locator.type(val, timeout=8000, delay=random.randint(20, 55))
        return True

    def run_js(self, script, *args):
        js_args = [a._handle() if isinstance(a, _NativeElement) else a for a in args]
        src = str(script or "")
        wrapped = "(el, argv) => { return (function () {\n" + src + "\n}).apply(el, argv); }"
        return self._locator.evaluate(wrapped, js_args, timeout=8000)

    def attr(self, name):
        try:
            return self._locator.get_attribute(str(name), timeout=2500) or ""
        except Exception:
            return ""

    @property
    def text(self):
        try:
            return self._locator.inner_text(timeout=2500) or ""
        except Exception:
            try:
                return self._locator.text_content(timeout=2500) or ""
            except Exception:
                return ""

    @property
    def is_checked(self):
        try:
            return bool(self._locator.is_checked(timeout=2500))
        except Exception:
            return False

    def eles(self, selector, timeout=None):
        css = _dp_selector_to_css(selector)
        loc = self._locator.locator(css)
        try:
            if timeout is not None:
                loc.first.wait_for(state="attached", timeout=max(1, int(float(timeout) * 1000)))
        except Exception:
            return []
        try:
            return [_NativeElement(self._ctx, loc.nth(i)) for i in range(loc.count())]
        except Exception:
            return []


class _NativeActions:
    def __init__(self, ctx):
        self._ctx = ctx
        self._ops = []

    def move_to(self, target, duration=0):
        self._ops.append(("move", target, duration))
        return self

    def hold(self):
        self._ops.append(("down",))
        return self

    def wait(self, seconds):
        self._ops.append(("wait", float(seconds or 0)))
        return self

    def release(self):
        self._ops.append(("up",))
        return self

    def press(self, key):
        self._ops.append(("press", key))
        return self

    def type(self, text):
        self._ops.append(("type", str(text or "")))
        return self

    def perform(self):
        page = self._ctx._page
        for op in self._ops:
            if op[0] == "move":
                target, duration = op[1], op[2]
                if isinstance(target, dict):
                    x, y = float(target.get("x", 0)), float(target.get("y", 0))
                else:
                    x, y = target
                x, y = self._ctx._to_page_xy(x, y)
                steps = max(1, min(25, int(float(duration or 0) / 35)))
                page.mouse.move(x, y, steps=steps)
            elif op[0] == "down":
                page.mouse.down()
            elif op[0] == "up":
                page.mouse.up()
            elif op[0] == "wait":
                time.sleep(max(0, op[1]))
            elif op[0] == "press":
                page.keyboard.press(_pw_key(op[1]))
            elif op[0] == "type":
                page.keyboard.type(op[1], delay=random.randint(20, 55))
        self._ops = []
        return True


class _NativeContext:
    def __init__(self, browser, page, ctx):
        self._browser = browser
        self._page = page
        self._ctx = ctx

    @property
    def url(self):
        try:
            return self._ctx.url or ""
        except Exception:
            return ""

    @property
    def actions(self):
        return _NativeActions(self)

    def _to_page_xy(self, x, y):
        return x, y

    def _locator(self, selector):
        raw = str(selector or "").strip()
        if raw.startswith("text:"):
            text = raw[5:].strip()
            try:
                return self._ctx.get_by_text(text, exact=False)
            except Exception:
                return self._ctx.locator(f"text={text}")
        return self._ctx.locator(_dp_selector_to_css(raw))

    def ele(self, selector, timeout=None):
        loc = self._locator(selector).first
        try:
            ms = max(1, int(float(timeout if timeout is not None else 0.5) * 1000))
            loc.wait_for(state="attached", timeout=ms)
            return _NativeElement(self, loc)
        except Exception:
            return None

    def eles(self, selector, timeout=None):
        loc = self._locator(selector)
        try:
            if timeout is not None:
                loc.first.wait_for(state="attached", timeout=max(1, int(float(timeout) * 1000)))
        except Exception:
            return []
        try:
            return [_NativeElement(self, loc.nth(i)) for i in range(loc.count())]
        except Exception:
            return []

    def run_js_loaded(self, script, *args):
        src = str(script or "")
        js_args = [a._handle() if isinstance(a, _NativeElement) else a for a in args]
        last_exc = None
        if js_args:
            wrapped = "(argv) => { return (function () {\n" + src + "\n}).apply(null, argv); }"
            return self._ctx.evaluate(wrapped, js_args)
        for code in (
            "() => {\n" + src + "\n}",
            "(function(){\n" + src + "\n})()",
            src,
        ):
            try:
                return self._ctx.evaluate(code)
            except Exception as exc:
                last_exc = exc
        raise last_exc or RuntimeError("evaluate failed")

    def get_all_frames(self):
        return []


class _NativeFrame(_NativeContext):
    def __init__(self, browser, page, frame):
        super().__init__(browser, page, frame)
        self._frame = frame

    def _to_page_xy(self, x, y):
        try:
            box = self._frame.frame_element().bounding_box()
            if box:
                return float(box.get("x", 0)) + x, float(box.get("y", 0)) + y
        except Exception:
            pass
        return x, y

    def get_all_frames(self):
        out = []
        try:
            for fr in self._frame.child_frames:
                out.append(_NativeFrame(self._browser, self._page, fr))
        except Exception:
            pass
        return out


class _NativePage(_NativeContext):
    def __init__(self, browser, page):
        super().__init__(browser, page, page)

    @property
    def title(self):
        try:
            return self._page.title()
        except Exception:
            return ""

    def get(self, url):
        self._page.goto(url, wait_until="domcontentloaded", timeout=60000)
        return self

    def wait_loading(self, timeout=20):
        ms = max(1, int(float(timeout or 20) * 1000))
        try:
            self._page.wait_for_load_state("domcontentloaded", timeout=ms)
        except Exception:
            pass
        try:
            self._page.wait_for_load_state("load", timeout=min(ms, 10000))
        except Exception:
            pass

    def screenshot(self, path, full_page=True):
        self._page.screenshot(path=path, full_page=bool(full_page), timeout=30000)

    def close(self):
        try:
            self._page.close()
        except Exception:
            pass

    def get_all_frames(self):
        out = []
        try:
            for fr in self._page.frames:
                if fr is self._page.main_frame:
                    continue
                out.append(_NativeFrame(self._browser, self._page, fr))
        except Exception:
            pass
        return out


class _NativeCamoufoxBrowser:
    def __init__(self, proxy_pool=None, headless=True, tag=""):
        _ensure_native_camoufox_ready()
        from camoufox.sync_api import Camoufox
        import warnings

        self._cm = None
        self._browser = None
        self._contexts = []
        self._pages = []
        self._proxy_forwarder = None
        proxy_raw = (proxy_pool or [""])[0] if proxy_pool else ""
        proxy_cfg = _proxy_to_camoufox(proxy_raw)
        launch_kwargs = {
            "headless": bool(headless),
            "humanize": True,
            "geoip": False,
        }
        try:
            from camoufox import DefaultAddons
            launch_kwargs["exclude_addons"] = [DefaultAddons.UBO]
        except Exception:
            pass
        if sys.platform.startswith("win"):
            launch_kwargs["os"] = "windows"
        if proxy_cfg:
            launch_kwargs["proxy"] = proxy_cfg
            if _env_bool("OUTLOOK_CAMONFOX_GEOIP", True):
                exit_ip = _fetch_exit_ip_for_camoufox(proxy_raw)
                if exit_ip:
                    launch_kwargs["geoip"] = exit_ip
                    log(f"  {tag} Camoufox geoip aligned to exit IP: {exit_ip}")
                    try:
                        from camoufox.geolocation import get_geolocation
                        geo = get_geolocation(exit_ip)
                        region = getattr(getattr(geo, "locale", None), "region", "") or ""
                        if region:
                            _CURRENT_IP_INFO.update({"ip": exit_ip, "country_code": str(region).upper()})
                            log(f"  {tag} Camoufox geoip country: {str(region).upper()}")
                    except Exception as exc:
                        log(f"  {tag} Camoufox geoip country lookup failed: {type(exc).__name__}: {exc}", "WARN")
                else:
                    launch_kwargs["i_know_what_im_doing"] = True
            else:
                launch_kwargs["i_know_what_im_doing"] = True
                log(f"  {tag} Camoufox geoip disabled")
        locale_pref = (os.environ.get("OUTLOOK_CAMONFOX_LOCALE") or "en-US").strip()
        if locale_pref and locale_pref.lower() not in ("auto", "ip", "geoip"):
            launch_kwargs["locale"] = locale_pref
            log(f"  {tag} Camoufox locale forced: {locale_pref}")
        else:
            launch_kwargs["locale"] = "en-US"
            log(f"  {tag} Camoufox locale: en-US")
        exe = _resolve_native_camoufox_exe()
        log(f"launch native Camoufox: {exe or '(camoufox package default)'} headless={bool(headless)} proxy={_mask_proxy_raw(proxy_raw) or 'DIRECT'}")
        last_exc = None
        for attempt in range(2):
            try:
                with warnings.catch_warnings():
                    warnings.filterwarnings("ignore", category=UserWarning)
                    log(f"  {tag} entering Camoufox context attempt {attempt + 1}")
                    self._cm = Camoufox(**launch_kwargs)
                    self._browser = self._cm.__enter__()
                    log(f"  {tag} Camoufox context ready")
                break
            except TypeError as exc:
                last_exc = exc
                launch_kwargs.pop("os", None)
                launch_kwargs.pop("i_know_what_im_doing", None)
            except Exception as exc:
                last_exc = exc
                msg = str(exc or "").lower()
                if any(k in msg for k in ("geoip", "geolite", "mmdb")):
                    launch_kwargs["geoip"] = False
                    launch_kwargs["i_know_what_im_doing"] = True
                    continue
                raise
        if self._browser is None:
            raise CamoufoxNotAvailable(f"Camoufox launch failed: {last_exc}")
        self._pages.append(self._new_page())

    def _new_page(self, url=""):
        try:
            raw = self._browser.new_page(no_viewport=True)
        except TypeError:
            try:
                ctx = self._browser.new_context(no_viewport=True)
                self._contexts.append(ctx)
                raw = ctx.new_page()
            except Exception:
                raw = self._browser.new_page()
        except Exception:
            raw = self._browser.new_page()
        page = _NativePage(self, raw)
        if url:
            page.get(url)
        return page

    def new_container_tab(self, url="about:blank"):
        page = self._new_page(url)
        self._pages.append(page)
        return page

    def get_tab(self, index=0):
        if not self._pages:
            self._pages.append(self._new_page())
        try:
            return self._pages[index]
        except Exception:
            return self._pages[0]

    def quit(self):
        for page in list(self._pages):
            try:
                page.close()
            except Exception:
                pass
        self._pages.clear()
        try:
            if self._cm is not None:
                self._cm.__exit__(None, None, None)
        except Exception:
            try:
                if self._browser is not None:
                    self._browser.close()
            except Exception:
                pass
        try:
            if self._proxy_forwarder is not None:
                self._proxy_forwarder.close()
        except Exception:
            pass

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


def _apply_camonfox_headless_options(tb, tag):
    """兼容旧参数对象的无头调优；原生 Camoufox 启动路径不再调用。"""
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
        log(f"  {tag} camonfox headless options applied: {', '.join(applied[:6])}")
    else:
        log(f"  {tag} camonfox headless options: no compatible option API found", "WARN")


def _apply_camonfox_headless_emulation(page, tag=None, log_once=False):
    emu = getattr(page, "emulation", None)
    if emu is None:
        if log_once:
            log(f"  {tag} camonfox headless emulation API not available", "WARN")
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
        log(f"  {tag} camonfox headless emulation applied: {applied}", level)
    return applied


CAMONFOX_HEADLESS_PATCH_JS = f"""
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
  const gp = WebGLRenderingContext && WebGLRenderingContext.prototype.getParameter;
  if (gp && !WebGLRenderingContext.prototype.__camonfoxHeadlessPatched) {{
    Object.defineProperty(WebGLRenderingContext.prototype, '__camonfoxHeadlessPatched', {{value: true}});
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


def _apply_camonfox_headless_page_patches(page, tag=None, log_once=False):
    _apply_camonfox_headless_emulation(page, tag, log_once=False)
    ok_count = 0
    for ctx in _all_contexts(page):
        try:
            if ctx.run_js_loaded(CAMONFOX_HEADLESS_PATCH_JS):
                ok_count += 1
        except Exception:
            continue
    if log_once:
        level = "OK" if ok_count else "WARN"
        log(f"  {tag} camonfox headless page patches applied to {ok_count} context(s)", level)
    return ok_count


def _on_signup_form(url):
    low = (url or "").lower()
    return "signup.live.com" in low and "privacynotice" not in low



def _scroll_into_view(el):
    """把元素滚进视口，避免点击坐标不可用。"""
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


def _click_visible_button_text(page, labels, tag=None):
    labels = [str(x or "").strip() for x in labels if str(x or "").strip()]
    if not labels:
        return ""
    script = r"""
const labels = arguments[0].map(s => String(s || '').trim()).filter(Boolean);
const norm = s => String(s || '').replace(/\s+/g, ' ').trim().toLowerCase();
const wants = labels.map(norm);
const visible = el => {
  if (!el) return false;
  const st = getComputedStyle(el);
  const r = el.getBoundingClientRect();
  return st.display !== 'none' && st.visibility !== 'hidden' && Number(st.opacity || 1) > 0
    && r.width > 5 && r.height > 5 && r.bottom > 0 && r.right > 0;
};
const textOf = el => (
  el.innerText || el.textContent || el.value || el.getAttribute('aria-label') || el.title || ''
);
const nodes = [
  ...document.querySelectorAll('button, input[type="button"], input[type="submit"], [role="button"], a')
].filter(visible);
let best = null;
for (const el of nodes) {
  const text = norm(textOf(el));
  if (!text) continue;
  let rank = 999;
  for (const want of wants) {
    if (text === want) rank = Math.min(rank, 0);
    else if (text.includes(want) && text.length <= Math.max(30, want.length + 20)) rank = Math.min(rank, 1);
  }
  if (rank < 999) {
    const r = el.getBoundingClientRect();
    const area = r.width * r.height;
    if (!best || rank < best.rank || (rank === best.rank && area < best.area)) {
      best = {el, rank, area, text: textOf(el)};
    }
  }
}
if (!best) return '';
best.el.scrollIntoView({block:'center', inline:'center'});
try { best.el.focus(); } catch(e) {}
best.el.click();
return String(best.text || '').trim();
"""
    for ctx in _all_contexts(page):
        try:
            hit = ctx.run_js_loaded(script, labels)
            if hit:
                return hit
        except Exception:
            continue
    return ""


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
    try:
        if page.run_js_loaded(
            """
return (() => {
  const sels = [
    'input[name="FirstName"]', '#FirstName', 'input[name="LastName"]', '#LastName',
    'input[name="firstNameInput"]', '#firstNameInput',
    'input[name="lastNameInput"]', '#lastNameInput'
  ];
  if (sels.some(s => document.querySelector(s))) return true;
  const visibleTextInputs = [...document.querySelectorAll('input[type="text"], input:not([type])')]
    .filter(el => {
      const r = el.getBoundingClientRect();
      const st = getComputedStyle(el);
      return r.width > 40 && r.height > 10 && st.display !== 'none' && st.visibility !== 'hidden';
    });
  if (visibleTextInputs.length >= 2) {
    const pageText = (document.body && document.body.innerText || '').toLowerCase();
    if (!/birth|birthday|country|region|password|email|captcha/.test(pageText)) return true;
  }
  return false;
})()
            """
        ):
            return True
    except Exception:
        pass
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
            "nombre",
            "apellido",
            "vorname",
            "nachname",
            "nome",
            "cognome",
            "氏名",
            "名前",
            "姓",
            "名",
            "성",
            "이름",
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
    labels = [
        'OK',
        'Accept',
        'Continue',
        'Next',
        'I agree',
        'Got it',
        'Agree',
        '同意',
        '继续',
        '下一步',
    ]
    js_hit = _click_visible_button_text(page, labels, tag=tag)
    if js_hit:
        log(f"  {tag} post-click: button:{js_hit}")
        return True
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




_COUNTRY_NAME_BY_CODE = {
    "US": ("United States", "United States of America", "USA", "??"),
    "JP": ("Japan", "??"),
    "GB": ("United Kingdom", "Great Britain", "??"),
    "CA": ("Canada", "???"),
    "AU": ("Australia", "????"),
    "FR": ("France", "??"),
    "DE": ("Germany", "Deutschland", "??"),
    "IT": ("Italy", "Italia", "???"),
    "ES": ("Spain", "Espa?a", "???"),
    "NL": ("Netherlands", "Holland", "??"),
    "SG": ("Singapore", "???"),
    "HK": ("Hong Kong", "??"),
    "TW": ("Taiwan", "??"),
    "KR": ("South Korea", "Korea", "??"),
    "TH": ("Thailand", "??"),
    "VN": ("Vietnam", "??"),
    "PH": ("Philippines", "???"),
    "ID": ("Indonesia", "?????"),
    "MY": ("Malaysia", "????"),
    "IN": ("India", "??"),
    "BR": ("Brazil", "Brasil", "??"),
    "MX": ("Mexico", "M?xico", "???"),
    "PL": ("Poland", "Polska", "??"),
    "TR": ("Turkey", "T?rkiye", "???"),
    "AE": ("United Arab Emirates", "UAE", "???"),
}


def _country_code_from_name(name):
    raw = str(name or "").strip()
    if not raw:
        return ""
    up = raw.upper()
    if len(up) == 2 and up.isalpha():
        return up
    low = raw.lower()
    for code, names in _COUNTRY_NAME_BY_CODE.items():
        if low == code.lower() or any(low == str(n).lower() for n in names):
            return code
    return ""


def _target_country_code():
    forced = (os.environ.get("OUTLOOK_CAMONFOX_COUNTRY") or os.environ.get("OUTLOOK_COUNTRY") or "").strip()
    if forced and forced.lower() not in ("auto", "ip", "geoip"):
        code = _country_code_from_name(forced) or forced[:2].upper()
        return code if len(code) == 2 else ""
    for key in ("country_code", "countryCode", "country"):
        code = _country_code_from_name(_CURRENT_IP_INFO.get(key, ""))
        if code:
            return code
    return ""


def _country_select_candidates(code):
    code = _country_code_from_name(code)
    if not code:
        return []
    out = [code, code.lower()]
    out.extend(_COUNTRY_NAME_BY_CODE.get(code, ()))
    # preserve order, unique
    seen = set()
    uniq = []
    for item in out:
        key = str(item).strip().lower()
        if key and key not in seen:
            seen.add(key)
            uniq.append(item)
    return uniq
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


def _clear_email_input(el):
    if el is None:
        return False
    _scroll_into_view(el)
    scripts = [
        """
const el = this;
el.focus();
const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set
  || Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value')?.set;
if (setter) setter.call(el, ''); else el.value = '';
el.dispatchEvent(new Event('input', {bubbles: true}));
el.dispatchEvent(new Event('change', {bubbles: true}));
return true;
        """,
    ]
    for script in scripts:
        try:
            if el.run_js(script):
                return True
        except Exception:
            pass
    try:
        el.input("", clear=True)
        return True
    except Exception:
        return False


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
            # Camoufox/Firefox 下邮箱占用后，页面有时仍是“前缀输入 + 固定域名”
            # 结构，但 select 一瞬间不可见/不可取；此时如果二次填入完整
            # xxx@outlook.com 会触发格式错误，导致无法继续提交。
            prefix_only = dd is not None or taken_retry_count > 0
            value = cur_prefix if prefix_only else cur_email
            if taken_retry_count > 0:
                if _clear_email_input(email_el):
                    log(f"  {tag} email retry input cleared before refill")
                else:
                    log(f"  {tag} email retry input clear failed, continue refill", "WARN")
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
            elif prefix_only:
                cur_email = f"{cur_prefix}@{cur_domain}"
                log(f"  {tag} filled prefix(no domain select): {cur_prefix}@{cur_domain}")
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
            # 点提交后如果仍停留在邮箱输入页，但没有明确 taken/format 文案，
            # 也按“邮箱提交失败”重试。Camoufox 下按钮点击/输入框残留时常见。
            still_email_el = _ele(page, sel, timeout=1)
            pwd_probe = _ele(
                page,
                'css:input[type="password"], input[name="Password"], #PasswordInput, input[name="passwd"]',
                timeout=1,
            )
            if still_email_el is not None and pwd_probe is None:
                taken_retry_count += 1
                extra_digits = 3 if taken_retry_count == 1 else 1
                cur_email, cur_prefix = _new_email_candidate(
                    "taken", cur_domain, cur_email, cur_prefix, extra_digits
                )
                cur_domain = _email_domain(cur_email)
                log(
                    f"  {tag} email submit stayed on email page, append {extra_digits} digit(s), retry: {cur_email}",
                    "WARN",
                )
                try:
                    _clear_email_input(still_email_el)
                except Exception:
                    pass
                time.sleep(2)
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

            # 国家/地区不切换：保留 Microsoft 页面默认值。
            target_country = ""
            country_picked = False
            if False and country_idx is not None and target_country:
                country_picked = bool(
                    _pick_select(
                        selects[country_idx],
                        target_country,
                        _country_select_candidates(target_country),
                    )
                )
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
                f"  {tag} 生日(select): country={target_country or 'keep'}"
                f"({'ok' if country_picked else 'keep/fail'} idx={country_idx}) "
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
    first_sel = (
        'css:input[name="FirstName"], #FirstName, input[name="firstNameInput"], '
        '#firstNameInput, css:input[aria-label*="first" i], '
        'css:input[aria-label*="pr?nom" i], css:input[aria-label*="nombre" i], '
        'css:input[aria-label*="vorname" i], css:input[aria-label*="nome" i], '
        'css:input[aria-label*="名前" i], css:input[aria-label*="名" i], '
        'css:input[aria-label*="?" i]'
    )
    last_sel = (
        'css:input[name="LastName"], #LastName, input[name="lastNameInput"], '
        '#lastNameInput, css:input[aria-label*="last" i], '
        'css:input[aria-label*="nom de famille" i], css:input[aria-label*="apellido" i], '
        'css:input[aria-label*="nachname" i], css:input[aria-label*="cognome" i], '
        'css:input[aria-label*="姓" i], css:input[aria-label*="?" i]'
    )

    fe = le = None
    text_inputs = []
    field_deadline = time.time() + 12
    while time.time() < field_deadline:
        fe = _ele(page, first_sel, timeout=0.5)
        le = _ele(page, last_sel, timeout=0.5)
        if fe is not None or le is not None:
            break
        text_inputs = _eles(page, 'css:input[type="text"]', timeout=0.5)
        if len(text_inputs) >= 2:
            break
        time.sleep(0.2)

    if fe is None and le is None and len(text_inputs) < 2 and not _is_name_page(page):
        log(f"  {tag} name page not detected, skip name fill url={getattr(page, 'url', '')}", "WARN")
        _shot(page, "name_fail", idx)
        return False

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
            time.sleep(1)
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
    time.sleep(1)
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
  // Some frame contexts are already inside the nested #px-captcha iframe;
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
        _shot(page, f"hold_{press_count}", idx)
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
    proxy_url = _proxy_to_requests_url(proxy_str)
    if not proxy_url:
        return None
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
                country_code = str(data.get("country_code") or data.get("countryCode") or "").strip().upper()
                _CURRENT_IP_INFO.update(
                    {
                        "ip": ip,
                        "country": country,
                        "country_code": country_code or _country_code_from_name(country),
                    }
                )
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

    th = threading.Thread(target=_worker, name=f"camonfox-ip-probe-{tag}", daemon=True)
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
    helpers = _load_helpers()
    _apply_account_options(opts)
    generate_email_password = helpers.generate_email_password
    generate_birthday = helpers.generate_birthday
    generate_name = helpers.generate_name
    verify_registered_outlook = helpers.verify_registered_outlook

    tag = f"[#{idx}][camonfox]"
    timeout = int(getattr(opts, "timeout", REGISTER_TIMEOUT) or REGISTER_TIMEOUT)
    max_press = int(getattr(opts, "max_press", os.environ.get("OUTLOOK_REG_MAX_PRESS", "5")) or 5)
    need_verify = not bool(getattr(opts, "no_verify", False))
    confirm_before_register = bool(getattr(opts, "confirm_before_register", False))
    is_headless = bool(getattr(opts, "headless", False))
    px_press_screenshots = bool(getattr(opts, "px_press_screenshots", False)) or _env_bool(
        "OUTLOOK_PX_PRESS_SCREENSHOTS", False
    )
    if proxy_pool:
        log(f"原生 Camoufox 使用本次账号代理: {_mask_proxy_raw(proxy_pool[0])}")
    else:
        log("没挂代理——直接本机出口", "WARN")

    browser_page = None
    page = None
    email = password = None
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

    try:
        browser_page = _NativeCamoufoxBrowser(proxy_pool=proxy_pool, headless=is_headless, tag=tag)
        page = browser_page.get_tab(0)

        if is_headless:
            _apply_camonfox_headless_page_patches(page, tag, log_once=True)
            headless_patch_logged = True
        page.get(SIGNUP_URL)
        page.wait_loading(20)
        if is_headless:
            _apply_camonfox_headless_page_patches(page, tag, log_once=False)
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
                _apply_camonfox_headless_page_patches(page, tag, log_once=(not headless_patch_logged))
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
                        if px_press_screenshots:
                            _save_screenshot(page, f"before_press_{press_count}", idx, tag)
                        if _perform_hold(page, ctx, target, idx, press_count, tag):
                            if px_press_screenshots:
                                _save_screenshot(page, f"after_press_{press_count}", idx, tag)
                            validation_wait_started = time.time()
                            awaiting_reappear = True
                            post_press_saw_gap = False
                            post_press_started_at = time.time()
                            if press_count >= max_press:
                                press_wait_started = time.time()
                            continue
                        if px_press_screenshots:
                            _save_screenshot(page, f"after_press_failed_{press_count}", idx, tag)
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
        return email, password
    except Exception as e:
        log(f"  {tag} FAILED: {type(e).__name__}: {e}", "ERR")
        try:
            if page is not None:
                _shot(page, "error", idx)
        except Exception:
            pass
        return None, None
    finally:
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
    ap = argparse.ArgumentParser(description="Outlook 自注册养号(camonfox) — 原生 Camoufox 版")
    ap.add_argument(
        "--proxy-source",
        choices=["file", "ip"],
        default=os.environ.get("OUTLOOK_CAMONFOX_PROXY_SOURCE", "file"),
        help="代理来源：file=代理文件，ip=单条代理地址",
    )
    ap.add_argument("--proxy-file", "-p", default=PROXY_FILE, help=f"代理池文件(默认 {PROXY_FILE})")
    ap.add_argument(
        "--proxy-ip",
        default=os.environ.get("OUTLOOK_CAMONFOX_PROXY_IP", ""),
        help="单条代理地址，如 http://user:pass@host:port 或 host:port:user:pass",
    )
    ap.add_argument("--count", "-n", type=int, default=1, help="注册次数(默认 1)")
    ap.add_argument("--headless", action="store_true", help="无头模式")
    ap.add_argument("--timeout", "-t", type=int, default=REGISTER_TIMEOUT, help="单号超时(秒)")
    ap.add_argument("--max-press", default=os.environ.get("OUTLOOK_REG_MAX_PRESS", "5"), help="按住次数上限")
    ap.add_argument("--no-verify", action="store_true", help="跳过 Outlook 登录校验")
    ap.add_argument("--confirm-before-register", action="store_true", help="页面打开后先尝试点确认")
    ap.add_argument("--px-press-screenshots", action=argparse.BooleanOptionalAction,
                    default=_env_bool("OUTLOOK_PX_PRESS_SCREENSHOTS", False),
                    help="保存 PX 按压前/按压后截图")
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

    _ensure_native_camoufox_ready()

    helpers = _load_helpers()
    _apply_account_options(args)
    try:
        proxy_env = helpers.ensure_clash_proxy_env()
        if proxy_env:
            log(f"proxy env ready: {proxy_env}")
    except Exception:
        pass

    if str(args.proxy_source or "file").lower() == "ip":
        proxy_pool = parse_proxy_value(args.proxy_ip)
        log(f"单条代理准备完毕: {len(proxy_pool)} 条")
    else:
        proxy_pool = parse_proxy_pool(args.proxy_file) if args.proxy_file else []
        log(f"代理池准备完毕: {len(proxy_pool)} 条")

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    live_file = os.path.join(OUTPUT_DIR, f"accounts_camonfox_{ts}.txt")
    token_file = os.path.join(OUTPUT_DIR, f"graph_tokens_camonfox_{ts}.json")

    ok = 0
    for i in range(args.count):
        log(f"========== 注册 #{i + 1}/{args.count} ==========")
        selected_pool = select_proxy_for_account(proxy_pool)
        if selected_pool:
            masked_proxy = _mask_proxy_raw(selected_pool[0])
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
