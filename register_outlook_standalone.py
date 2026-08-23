# -*- coding: utf-8 -*-
"""
Standalone Outlook Email Registration Script
Uses BitBrowser + Playwright + Proxy to register Outlook accounts
Independent from the main register.py — only registers Outlook accounts

Usage:
  python register_outlook_standalone.py --count 10
  python register_outlook_standalone.py --count 5 --concurrency 2
  python register_outlook_standalone.py --proxy-file proxies.txt
"""

import argparse
import asyncio
from contextlib import contextmanager
import json
import math
import os
import queue
import random
import re
import string
import sys
import threading
import time
from datetime import datetime, timezone
from urllib.parse import parse_qs, urlsplit

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    # pytest 下 stdin 是 DontReadFromInput(无 reconfigure),兼容之
    try:
        sys.stdin.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

import requests
from playwright.async_api import async_playwright
try:
    from playwright_stealth import Stealth as _StealthCls
    _HAS_STEALTH = True
    _stealth_obj = _StealthCls()
except ImportError:
    _HAS_STEALTH = False
    _stealth_obj = None

try:
    from check_outlook_status import check_account_api
except Exception:
    check_account_api = None

# ======================== Configuration ========================

# 导入 config 以触发 .env 加载（密钥来自 .env / 真实环境变量）。
try:
    import config  # noqa: F401
except Exception:
    pass

# BitBrowser local API
BITBROWSER_API = os.environ.get("BITBROWSER_API", "http://127.0.0.1:54345")


def _fingerprint_provider():
    return (
        os.environ.get("FINGERPRINT_BROWSER")
        or os.environ.get("BROWSER_PROVIDER")
        or "bitbrowser"
    ).strip().lower()

# Output
OUTPUT_DIR = "outlook_accounts"
SCREENSHOT_DIR = "screenshots_outlook"

# Registration timeout per account (seconds)
REGISTER_TIMEOUT = 300
VERIFY_AFTER_REGISTER = True
ACCOUNT_QUEUE = None
_ACCOUNT_CONTEXT = threading.local()
_ACCOUNT_OPTIONS_CONTEXT = threading.local()


def verify_registered_outlook(email, password, tag=""):
    """Verify the saved password can actually log in before exporting the account."""
    if not VERIFY_AFTER_REGISTER:
        return True
    if check_account_api is None:
        print(f"  {tag} verify skipped: check_outlook_status unavailable")
        return True
    result = check_account_api(email, password)
    status = result.get("status")
    code = result.get("code") or ""
    msg = result.get("message") or ""
    print(f"  {tag} post-register verify: {status} {code} {msg[:80]}")
    return status == "ok"

# Default proxies (user:pass@host:port)
# 住宅代理账密池来自环境变量 OUTLOOK_PROXIES（多个用换行或逗号分隔），默认空。
# 也可用 --proxy-file 指定文件；两者都为空时不走代理。
def _load_default_proxies():
    raw = os.environ.get("OUTLOOK_PROXIES", "")
    if not raw:
        return []
    parts = [p.strip() for p in raw.replace(",", "\n").splitlines()]
    return [p for p in parts if p and not p.startswith("#")]


DEFAULT_PROXIES = _load_default_proxies()


# ======================== BitBrowser API ========================

class BitBrowserClient:
    """BitBrowser local API client with proxy support"""

    def __new__(cls, api_base=None):
        if cls is BitBrowserClient and _fingerprint_provider() in {"adspower", "ads_power", "ads"}:
            from bitbrowser import BitBrowser
            return BitBrowser(api_base=api_base)
        return super().__new__(cls)

    def __init__(self, api_base=None):
        self.api_base = api_base or BITBROWSER_API

    def _post(self, path, data=None):
        url = f"{self.api_base}{path}"
        resp = requests.post(url, json=data or {}, timeout=120)
        resp.raise_for_status()
        result = resp.json()
        if not result.get("success"):
            raise Exception(f"BitBrowser API error: {result.get('msg', 'unknown')}")
        return result

    def create_browser(self, name="outlook_reg", proxy_str=None):
        """Create a new browser profile with optional proxy.
        proxy_str format: user:pass@host:port
        """
        data = {
            "name": name,
            "remark": "outlook standalone registration",
            "proxyMethod": 2,  # custom proxy
            "browserFingerPrint": {
                "coreVersion": "130",
            },
        }

        if proxy_str:
            parsed = self._parse_proxy(proxy_str)
            if parsed:
                data["proxyType"] = parsed.get("type", "http")
                data["host"] = parsed["host"]
                data["port"] = parsed["port"]
                if parsed.get("username"):
                    data["proxyUserName"] = parsed["username"]
                if parsed.get("password"):
                    data["proxyPassword"] = parsed["password"]
                print(f"  proxy: [{data['proxyType']}] {parsed['host']}:{parsed['port']} (user={parsed.get('username', 'none')[:20]}...)")
            else:
                data["proxyType"] = "noproxy"
                print(f"  proxy: invalid format, using noproxy")
        else:
            data["proxyType"] = "noproxy"

        result = self._post("/browser/update", data)
        profile_id = result["data"]["id"]
        print(f"  browser created: {name} (ID: {profile_id})")
        return profile_id

    def open_browser(self, profile_id):
        """Open browser window, returns WebSocket debug URL"""
        result = self._post("/browser/open", {"id": profile_id})
        return result["data"]

    def close_browser(self, profile_id):
        """Close browser window"""
        try:
            self._post("/browser/close", {"id": profile_id})
        except Exception:
            pass

    def delete_browser(self, profile_id):
        """Delete browser profile"""
        try:
            self._post("/browser/delete", {"id": profile_id})
        except Exception:
            pass

    def cleanup_browsers(self, keep=0):
        """Delete all browser profiles (release quota)"""
        result = self._post("/browser/list", {"page": 0, "pageSize": 200})
        browsers = result["data"]["list"]
        if not browsers:
            return 0
        browsers.sort(key=lambda b: b.get("seq", 0), reverse=True)
        to_delete = browsers[keep:]
        deleted = 0
        for b in to_delete:
            try:
                self.close_browser(b["id"])
            except Exception:
                pass
            time.sleep(1)
            try:
                self.delete_browser(b["id"])
                deleted += 1
            except Exception:
                pass
        print(f"  cleanup: deleted {deleted}/{len(to_delete)} browsers")
        return deleted

    @staticmethod
    def _parse_proxy(proxy_str):
        """Parse proxy string into dict.
        Supported formats:
          socks5h://user:pass@host:port
          socks5h://host:port
          socks5://user:pass@host:port
          socks5://host:port
          user:pass@host:port          (defaults to http)
          host:port                    (defaults to http)
        """
        # Strip protocol prefix
        proxy_type = "http"
        lower = proxy_str.lower()
        if lower.startswith("socks5h://"):
            proxy_type = "socks5"
            proxy_str = proxy_str[len("socks5h://"):]
        elif lower.startswith("socks5://"):
            proxy_type = "socks5"
            proxy_str = proxy_str[len("socks5://"):]
        elif lower.startswith("http://"):
            proxy_str = proxy_str[len("http://"):]
        elif lower.startswith("https://"):
            proxy_str = proxy_str[len("https://"):]

        # Handle comma-separated format: user:pass,host:port
        proxy_str = proxy_str.replace(",", "@", 1) if "@" not in proxy_str and "," in proxy_str else proxy_str

        match = re.match(r'^(.+):(.+)@(.+):(\d+)$', proxy_str)
        if match:
            return {
                "type": proxy_type,
                "username": match.group(1),
                "password": match.group(2),
                "host": match.group(3),
                "port": match.group(4),
            }
        match2 = re.match(r'^(.+):(\d+)$', proxy_str)
        if match2:
            return {
                "type": proxy_type,
                "host": match2.group(1),
                "port": match2.group(2),
            }
        return None


# ======================== Helper Functions ========================

def generate_birthday():
    """Generate a random birthday (25-40 years old)"""
    current_year = datetime.now().year
    year = random.randint(current_year - 40, current_year - 25)
    month = random.randint(1, 12)
    if month in (1, 3, 5, 7, 8, 10, 12):
        max_day = 31
    elif month in (4, 6, 9, 11):
        max_day = 30
    else:
        max_day = 28
    day = random.randint(1, max_day)
    return year, month, day


def generate_name():
    """Generate a random English name"""
    first_names = [
        "James", "John", "Robert", "Michael", "David", "William", "Richard", "Joseph",
        "Thomas", "Charles", "Mary", "Patricia", "Jennifer", "Linda", "Barbara",
        "Elizabeth", "Susan", "Jessica", "Sarah", "Karen", "Emily", "Emma", "Olivia",
        "Daniel", "Matthew", "Anthony", "Mark", "Steven", "Andrew", "Brian",
    ]
    last_names = [
        "Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller",
        "Davis", "Rodriguez", "Martinez", "Anderson", "Taylor", "Thomas", "Moore",
        "Jackson", "Martin", "Lee", "Thompson", "White", "Harris", "Clark",
    ]
    return random.choice(first_names), random.choice(last_names)


def _slug_name(value):
    return re.sub(r"[^a-z0-9]", "", str(value or "").lower())


def _current_account_names(create=False):
    first = getattr(_ACCOUNT_CONTEXT, "first_name", "")
    last = getattr(_ACCOUNT_CONTEXT, "last_name", "")
    if create and (not first or not last):
        first, last = generate_name()
        _ACCOUNT_CONTEXT.first_name = first
        _ACCOUNT_CONTEXT.last_name = last
    return first, last


def _consume_account_names():
    first, last = _current_account_names(create=True)
    _ACCOUNT_CONTEXT.first_name = ""
    _ACCOUNT_CONTEXT.last_name = ""
    return first, last


def _random_chars(chars, length):
    return "".join(random.choices(chars, k=max(0, int(length))))


def _expand_random_template(template, default_generator):
    """Expand lightweight random placeholders used by account/password formats."""
    if not template:
        return default_generator()

    def repl(match):
        kind = (match.group(1) or "rand").lower()
        length = int(match.group(2) or "8")
        if kind in {"digit", "digits", "num", "number"}:
            chars = string.digits
        elif kind in {"letter", "letters", "lower"}:
            chars = string.ascii_lowercase
        elif kind == "upper":
            chars = string.ascii_uppercase
        else:
            chars = string.ascii_letters + string.digits
        return _random_chars(chars, length)

    # Supported: {rand:12}, {letters:8}, {digits:4}, {upper:2}
    value = re.sub(r"\{(rand|random|letters?|lower|upper|digits?|num|number):(\d+)\}", repl, template)
    first, last = _current_account_names(create=bool(re.search(r"\{(?:first|last|name)", value, re.I)))
    first_slug = _slug_name(first)
    last_slug = _slug_name(last)
    name_values = {
        "{first}": first_slug,
        "{firstname}": first_slug,
        "{last}": last_slug,
        "{lastname}": last_slug,
        "{first_initial}": first_slug[:1],
        "{last_initial}": last_slug[:1],
        "{fi}": first_slug[:1],
        "{li}": last_slug[:1],
        "{name}": f"{first_slug}{last_slug}",
        "{first.raw}": first,
        "{last.raw}": last,
    }
    for key, replacement in name_values.items():
        value = value.replace(key, replacement)
    # Convenience aliases without length.
    value = value.replace("{rand}", _random_chars(string.ascii_letters + string.digits, 8))
    value = value.replace("{lower}", _random_chars(string.ascii_lowercase, 8))
    value = value.replace("{digits}", _random_chars(string.digits, 4))
    return value


def _email_suffixes():
    raw = getattr(_ACCOUNT_OPTIONS_CONTEXT, "email_suffixes", None)
    if raw is None:
        raw = (
            os.environ.get("OUTLOOK_ACCOUNT_SUFFIXES")
            or os.environ.get("OUTLOOK_EMAIL_SUFFIXES")
            or "outlook.com"
        )
    suffixes = [s.strip().lstrip("@") for s in re.split(r"[,;\s]+", raw) if s.strip()]
    return suffixes or ["outlook.com"]


def _generate_prefix():
    fmt = getattr(_ACCOUNT_OPTIONS_CONTEXT, "account_format", None)
    is_literal = getattr(_ACCOUNT_OPTIONS_CONTEXT, "account_format_literal", False)
    if fmt is None:
        fmt = os.environ.get("OUTLOOK_ACCOUNT_FORMAT", "").strip()
        is_literal = is_literal or os.environ.get("OUTLOOK_ACCOUNT_FORMAT_LITERAL", "").strip() in ("1", "true", "yes", "on")
    if fmt:
        # custom 指定格式:固定内容,不做 {占位符} 展开(用户填什么邮箱前缀就是什么)。
        if is_literal:
            return fmt
        return _expand_random_template(fmt, lambda: "")
    return random.choice(string.ascii_lowercase) + _random_chars(string.ascii_lowercase + string.digits, 11)


def _generate_password():
    fmt = getattr(_ACCOUNT_OPTIONS_CONTEXT, "password_format", None)
    if fmt is None:
        fmt = os.environ.get("OUTLOOK_PASSWORD_FORMAT", "").strip()
    if fmt:
        return _expand_random_template(fmt, lambda: "")
    return "Aa1!" + _random_chars(string.ascii_letters + string.digits, 12)


def _normalize_account_spec(line):
    """Parse exe-like account formats: account / account@suffix / account----password."""
    raw = (line or "").strip()
    if not raw or raw.startswith("#"):
        return None
    if raw.count("----") > 1:
        raise ValueError("account line can contain at most one ---- separator")
    account, password = (raw.split("----", 1) + [""])[:2] if "----" in raw else (raw, "")
    account = account.strip()
    password = password.strip()
    if not account:
        return None
    if "@" in account:
        email = account
        prefix = account.split("@", 1)[0]
    else:
        prefix = account
        email = f"{account}@{random.choice(_email_suffixes())}"
    return email, password or _generate_password(), prefix


def _next_account_spec():
    global ACCOUNT_QUEUE
    if ACCOUNT_QUEUE is None:
        return None
    try:
        return ACCOUNT_QUEUE.get_nowait()
    except queue.Empty:
        return None


def generate_email_password():
    """Generate or pop an Outlook email/password pair."""
    queued = _next_account_spec()
    if queued:
        return queued
    prefix = _generate_prefix()
    email = f"{prefix}@{random.choice(_email_suffixes())}"
    password = _generate_password()
    return email, password, prefix


def _random_digits(count):
    return "".join(random.choice(string.digits) for _ in range(max(1, int(count or 1))))


def _email_domain(email):
    if not email or "@" not in str(email):
        return random.choice(_email_suffixes())
    return str(email).rsplit("@", 1)[-1].strip().lstrip("@") or random.choice(_email_suffixes())


def _append_random_digits_email(email, prefix=None, digits=1, domain=None):
    domain = (domain or _email_domain(email)).strip().lstrip("@") or random.choice(_email_suffixes())
    base_prefix = str(prefix or "").strip()
    if not base_prefix and email and "@" in str(email):
        base_prefix = str(email).split("@", 1)[0].strip()
    if not base_prefix:
        base_prefix = _generate_prefix()
    new_prefix = f"{base_prefix}{_random_digits(digits)}"
    return f"{new_prefix}@{domain}", new_prefix


def load_account_queue(path):
    q = queue.Queue()
    if not path:
        return q
    loaded = 0
    with open(path, "r", encoding="utf-8") as f:
        for lineno, line in enumerate(f, 1):
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            try:
                item = _normalize_account_spec(line)
            except ValueError as exc:
                print(f"  account file line {lineno} skipped: {exc}")
                continue
            if item:
                q.put(item)
                loaded += 1
    print(f"  loaded account specs: {loaded} from {path}")
    return q


ACCOUNT_FORMAT_PRESETS = {
    "random": "",
    # 名字库仅 30x21=630 组合,Outlook 常见英文名几乎全被占,裸名必 taken 每次重试浪费 3-5s。
    # name/name_digits 统一用纯随机字母+数字({letters:7}{digits:6}=36^7×10^6≈8e15 种),taken 概率压到 0。
    # 姓名页仍用干净 first/last(独立 generate_name),不受影响。WebUI 已只暴露 custom。
    "name": "{letters:7}{digits:6}",
    "name_digits": "{letters:7}{digits:6}",
}


def resolve_account_format(mode, custom_format=""):
    mode = (mode or "").strip().lower()
    if mode == "custom":
        return (custom_format or "").strip()
    return ACCOUNT_FORMAT_PRESETS.get(mode, "")


def apply_account_format_mode(mode, custom_format=""):
    value = resolve_account_format(mode, custom_format)
    if value:
        os.environ["OUTLOOK_ACCOUNT_FORMAT"] = value
    else:
        os.environ.pop("OUTLOOK_ACCOUNT_FORMAT", None)
    # custom + 纯固定内容(不含 {占位符}) → 原样不展开;其它(含模板占位符) → 正常展开
    mode_lower = str(mode or "").strip().lower()
    raw_fmt = str(custom_format or "").strip()
    is_literal = (mode_lower == "custom") and bool(raw_fmt) and ("{" not in raw_fmt)
    if is_literal:
        os.environ["OUTLOOK_ACCOUNT_FORMAT_LITERAL"] = "1"
    else:
        os.environ.pop("OUTLOOK_ACCOUNT_FORMAT_LITERAL", None)


def set_account_generation_options(email_suffixes=None, account_format_mode=None, account_format="", password_format=None):
    normalized_suffixes = None
    if email_suffixes is not None:
        if isinstance(email_suffixes, (list, tuple, set)):
            normalized_suffixes = ",".join(str(x).strip() for x in email_suffixes if str(x).strip())
        else:
            normalized_suffixes = str(email_suffixes or "").strip()
        normalized_suffixes = normalized_suffixes or None
    setattr(_ACCOUNT_OPTIONS_CONTEXT, "email_suffixes", normalized_suffixes)
    # custom(指定格式)模式下,account_format 若是纯固定内容(不含 {占位符})则原样做邮箱前缀,不展开;
    # 若含 {letters:7}{digits:6} 这类占位符则正常展开。其他模式(name 等)走预设模板,正常展开。
    mode_lower = str(account_format_mode or "").strip().lower()
    raw_fmt = str(account_format or "").strip()
    is_literal = (mode_lower == "custom") and bool(raw_fmt) and ("{" not in raw_fmt)
    setattr(_ACCOUNT_OPTIONS_CONTEXT, "account_format", resolve_account_format(account_format_mode, account_format) or "")
    setattr(_ACCOUNT_OPTIONS_CONTEXT, "account_format_literal", is_literal)
    normalized_password_format = None if password_format is None else str(password_format or "").strip()
    setattr(_ACCOUNT_OPTIONS_CONTEXT, "password_format", normalized_password_format)


def clear_account_generation_options():
    for name in ("email_suffixes", "account_format", "account_format_literal", "password_format"):
        if hasattr(_ACCOUNT_OPTIONS_CONTEXT, name):
            delattr(_ACCOUNT_OPTIONS_CONTEXT, name)


# ======================== Graph API Token ========================

# 代理授权全失败后，回退直连的重试次数（短退避）；直连模式不触发，仍按 attempts 重试。
GRAPH_DIRECT_FALLBACK_ATTEMPTS = 3


def extract_graph_token_http(email, password, idx=0, attempts=3, proxy_str=None, bind_secondary=None):
    """Extract Graph refresh_token through the shared pure-HTTP OAuth flow.

    proxy_str 为空 -> 直连（proxies=None + trust_env=False），重试 attempts 次。
    proxy_str 非空 -> 走该代理重试 attempts 次；全部失败后回退直连再重试
    GRAPH_DIRECT_FALLBACK_ATTEMPTS 次（短退避），避免代理故障导致拿不到 token。

    bind_secondary: 传入则 proofs/Add 真绑 cf 辅助邮箱(而非 Skip),
      结构见 extract_graph_tokens.get_graph_token 的 bind_secondary。
      8月起微软对新号收紧,无辅助邮箱一律 access_denied,注册后授权必须传这个。
    不传则保持原 Skip 行为(向后兼容)。
    """
    try:
        from extract_graph_tokens import get_graph_token
    except Exception as exc:
        print(f"  [#{idx}] [graph] import error: {exc}")
        return None

    proxies = _proxy_for_requests(proxy_str) if proxy_str else None
    if proxies:
        print(f"  [#{idx}] [graph] use reg proxy for auth")
    else:
        print(f"  [#{idx}] [graph] direct (no proxy)")

    def _run_attempts(count, label, current_proxies, backoff_base):
        for attempt in range(count):
            try:
                print(f"  [#{idx}] [graph] attempt {attempt + 1}/{count} proxy={label}")
                # trust_env=False 已在 get_graph_token 内保证；proxies=None 即直连。
                res = get_graph_token(email, password, idx, proxies=current_proxies, bind_secondary=bind_secondary)
            except Exception as exc:
                print(f"  [#{idx}] [graph] attempt {attempt + 1}/{count} error: {exc}")
                res = None
            if res and res.get("refresh_token"):
                return {
                    "refresh_token": res["refresh_token"],
                    "client_id": res.get("client_id") or "",
                    "cf_address": res.get("cf_address") or "",
                    "cf_password": res.get("cf_password") or "",
                }
            if attempt < count - 1:
                print(f"  [#{idx}] [graph] no refresh_token on attempt {attempt + 1}/{count}; retry...")
                time.sleep(backoff_base * (attempt + 1))
        return None

    # 直连模式：只走直连 attempts 次，行为与原先一致（不触发回退）。
    if not proxies:
        return _run_attempts(attempts, "direct", None, 3)

    # 代理模式：先走代理 attempts 次，全失败再回退直连 GRAPH_DIRECT_FALLBACK_ATTEMPTS 次（短退避）。
    result = _run_attempts(attempts, "reg", proxies, 3)
    if result:
        return result
    print(
        f"  [#{idx}] [graph] reg proxy exhausted after {attempts} attempts, "
        f"fallback to direct {GRAPH_DIRECT_FALLBACK_ATTEMPTS}x"
    )
    return _run_attempts(GRAPH_DIRECT_FALLBACK_ATTEMPTS, "direct", None, 2)


# ======================== Outlook Registration ========================


def _env_truthy(name, default="0"):
    return (os.environ.get(name, default) or "").strip().lower() in {"1", "true", "yes", "on"}


async def _maybe_confirm_before_register(page, tag, captcha_early_abort=False):
    """Auto-click a confirmation/consent gate shown before the signup form."""
    if captcha_early_abort or not _env_truthy("OUTLOOK_CONFIRM_BEFORE_REGISTER"):
        return
    try:
        title = await page.title()
    except Exception:
        title = ""
    print(f"  {tag} signup page opened: {page.url}")
    if title:
        print(f"  {tag} page title: {title[:100]}")
    selectors = [
        'button:has-text("确认")', 'button:has-text("确定")',
        'button:has-text("同意")', 'button:has-text("接受")',
        'button:has-text("OK")', 'button:has-text("Ok")',
        'button:has-text("Confirm")', 'button:has-text("Accept")',
        'button:has-text("Agree")', 'button:has-text("Agree and continue")',
        'button:has-text("同意して続行")', 'button:has-text("確認")',
        'button:has-text("確定")',
        'input[type="submit"][value*="确认"]', 'input[type="submit"][value*="确定"]',
        'input[type="submit"][value*="同意"]', 'input[type="submit"][value*="接受"]',
        'input[type="submit"][value*="OK"]', 'input[type="submit"][value*="Confirm"]',
        'input[type="submit"][value*="Accept"]', 'input[type="submit"][value*="Agree"]',
        'a:has-text("确认")', 'a:has-text("确定")', 'a:has-text("同意")',
        'a:has-text("OK")', 'a:has-text("Confirm")', 'a:has-text("Accept")',
    ]
    for _ in range(3):
        for sel in selectors:
            try:
                btn = page.locator(sel).first
                if await btn.count() > 0 and await btn.is_visible(timeout=800):
                    await btn.click(timeout=3000)
                    print(f"  {tag} auto-confirm clicked: {sel}")
                    await asyncio.sleep(2)
                    return
            except Exception:
                pass
        await asyncio.sleep(1)
    print(f"  {tag} auto-confirm: no confirmation button found")
    return


async def register_outlook(page, context, idx=0, captcha_early_abort=False):
    """
    Register a new Outlook email account.
    Returns (email, password) on success, (None, None) on failure.

    captcha_early_abort: when True (headless mode), abort immediately after captcha
    solvers fail so the caller can fall back faster. When False (browser/BitBrowser
    mode), keep the loop running — PX presses sometimes pass after 10–30 s naturally.
    """
    os.makedirs(SCREENSHOT_DIR, exist_ok=True)
    tag = f"[#{idx}]"
    px_press_screenshots = _env_truthy("OUTLOOK_PX_PRESS_SCREENSHOTS", "0")

    try:
        print(f"  {tag} navigating to signup page...")
        await page.goto("https://signup.live.com/signup?lic=1", timeout=60000, wait_until="domcontentloaded")
        await asyncio.sleep(3)
        await page.screenshot(path=f"{SCREENSHOT_DIR}/outlook_{idx}_start.png")
        await _maybe_confirm_before_register(page, tag, captcha_early_abort)

        # Handle privacy/consent pages (Chinese "个人数据导出许可", "同意并继续", etc.)
        for _consent_try in range(5):
            page_text = await page.evaluate("() => document.body.innerText")
            current_url = page.url.lower()
            # Check if on a consent/privacy page (not the actual signup form)
            # Only trigger for actual privacy/consent standalone pages, not signup pages with footer links
            is_signup_form = "signup.live.com" in current_url and "privacynotice" not in current_url
            if not is_signup_form and (
                any(kw in page_text for kw in ["同意并继续", "个人数据", "数据导出"]) or \
                any(kw in page_text.lower() for kw in [
                    "agree and continue", "consent", "data export",
                    "accepter et continuer", "consentement",
                ]) or "privacynotice" in current_url
            ):
                print(f"  {tag} privacy/consent page detected, clicking accept...")
                clicked = False
                # Try various accept buttons
                for sel in [
                    'button:has-text("同意并继续")', 'input[value="同意并继续"]',
                    'button:has-text("同意")', 'a:has-text("同意并继续")',
                    'button:has-text("Agree and continue")', 'button:has-text("Accept")',
                    'button:has-text("Continue")', 'button:has-text("OK")',
                    'button:has-text("Accepter et continuer")', 'button:has-text("Accepter")',
                    'button:has-text("Continuer")', 'button:has-text("Suivant")',
                    'input[type="submit"]', 'button[type="submit"]',
                    '#iNext', '#iAgree', '#acceptButton',
                ]:
                    btn = page.locator(sel).first
                    if await btn.count() > 0:
                        try:
                            await btn.click(timeout=5000)
                            print(f"  {tag} clicked consent: {sel}")
                            clicked = True
                            break
                        except Exception:
                            pass
                if not clicked:
                    # Fallback: click any visible button
                    try:
                        await page.evaluate("""() => {
                            const btns = document.querySelectorAll('button, input[type="submit"], a.btn');
                            for (const b of btns) {
                                if (b.offsetParent !== null && b.textContent.length < 30) {
                                    b.click(); return true;
                                }
                            }
                            return false;
                        }""")
                        print(f"  {tag} JS-clicked consent button")
                    except Exception:
                        pass
                await asyncio.sleep(3)
                await page.screenshot(path=f"{SCREENSHOT_DIR}/outlook_{idx}_after_consent_{_consent_try}.png")
            else:
                break

        # Generate email and password
        email, password, prefix = generate_email_password()
        print(f"  {tag} registering: {email}")

        # Step 1: Enter email
        email_ok = False
        taken_retry_count = 0
        current_domain = _email_domain(email)
        for retry in range(5):
            email_input = page.locator(
                'input[type="email"], input[name="MemberName"], input[id="MemberName"], '
                'input[id="usernameInput"], input[name="Username"]'
            ).first
            if await email_input.count() == 0:
                print(f"  {tag} email input not found")
                await page.screenshot(path=f"{SCREENSHOT_DIR}/outlook_{idx}_no_email.png")
                return None, None

            domain_dropdown = page.locator(
                'select[id="LiveDomainBoxList"], select[name="LiveDomainBoxList"], #LiveDomainBoxList'
            ).first
            has_domain_dropdown = await domain_dropdown.count() > 0

            await email_input.fill("")
            await asyncio.sleep(0.3)
            if has_domain_dropdown:
                await email_input.fill(prefix)
                try:
                    await domain_dropdown.select_option("outlook.com")
                    current_domain = "outlook.com"
                    email = f"{prefix}@{current_domain}"
                except Exception:
                    pass
                print(f"  {tag} filled prefix: {prefix} (dropdown)")
            else:
                await email_input.fill(email)
                print(f"  {tag} filled email: {email}")

            await asyncio.sleep(0.5)
            for sel in ['input[type="submit"]', 'button[type="submit"]', '#iSignupAction', 'button[id="iSignupAction"]']:
                btn = page.locator(sel).first
                if await btn.count() > 0:
                    await btn.click(timeout=3000)
                    break
            await asyncio.sleep(3)

            page_text = await page.evaluate("() => document.body.innerText")
            page_lower = page_text.lower()

            if ("already" in page_lower and "email" in page_lower) or "taken" in page_lower:
                taken_retry_count += 1
                extra_digits = 3 if taken_retry_count == 1 else 1
                email, prefix = _append_random_digits_email(
                    email, prefix, extra_digits, current_domain
                )
                current_domain = _email_domain(email)
                print(f"  {tag} email taken, append {extra_digits} digit(s), retry: {email}")
                continue

            if "needs to start" in page_lower or "in the format" in page_lower or "enter a valid" in page_lower or "use letters" in page_lower:
                prefix = random.choice(string.ascii_lowercase) + "".join(
                    random.choices(string.ascii_lowercase + string.digits, k=9)
                )
                email = f"{prefix}@outlook.com"
                print(f"  {tag} format error, retry: {email}")
                continue

            email_ok = True
            break

        if not email_ok:
            print(f"  {tag} all email attempts failed")
            await page.screenshot(path=f"{SCREENSHOT_DIR}/outlook_{idx}_email_fail.png")
            return None, None

        await page.screenshot(path=f"{SCREENSHOT_DIR}/outlook_{idx}_after_email.png")

        # Step 2: Enter password
        await asyncio.sleep(2)
        pwd_input = None
        for _ in range(10):
            pwd_input = page.locator(
                'input[type="password"], input[name="Password"], '
                'input[id="PasswordInput"], input[name="passwd"]'
            ).first
            if await pwd_input.count() > 0:
                break
            await asyncio.sleep(1)

        if pwd_input and await pwd_input.count() > 0:
            await pwd_input.fill(password)
            print(f"  {tag} password filled")
            await asyncio.sleep(0.5)

            clicked_next = False
            for sel in ['#iSignupAction', 'input[type="submit"]', 'button[type="submit"]',
                        'button:has-text("Next")', 'button:has-text("next")',
                        'button:has-text("下一步")', 'button:has-text("Suivant")']:
                btn = page.locator(sel).first
                if await btn.count() > 0:
                    try:
                        await btn.click(timeout=3000)
                        clicked_next = True
                        break
                    except Exception:
                        pass
            if not clicked_next:
                await page.keyboard.press("Enter")

            await asyncio.sleep(3)
            await page.screenshot(path=f"{SCREENSHOT_DIR}/outlook_{idx}_after_pwd.png")
        else:
            print(f"  {tag} password input not found")
            return None, None

        # Step 3: Country + Birthday
        year, month, day = generate_birthday()
        await asyncio.sleep(2)

        # Wait for birthday page to load (CN/EN/FR)
        for _ in range(10):
            page_text = await page.evaluate("() => document.body.innerText")
            if any(kw in page_text.lower() for kw in [
                "birth", "country", "region",
                "naissance", "pays", "région", "détails",
            ]) or any(kw in page_text for kw in ["出生", "国家", "地区", "年份", "详细信息"]):
                break
            await asyncio.sleep(1)

        await page.screenshot(path=f"{SCREENSHOT_DIR}/outlook_{idx}_bday_page.png")

        # Debug: dump all form elements
        form_debug = await page.evaluate("""() => {
            const els = document.querySelectorAll('input, select, button[role="combobox"], [role="combobox"], [role="listbox"]');
            return Array.from(els).filter(e => e.offsetParent !== null).map(e => ({
                tag: e.tagName, id: e.id, name: e.name, type: e.type || '',
                role: e.getAttribute('role') || '',
                ariaLabel: e.getAttribute('aria-label') || '',
                text: e.textContent ? e.textContent.trim().substring(0, 30) : '',
                placeholder: e.placeholder || '',
            }));
        }""")
        print(f"  {tag} form elements: {json.dumps(form_debug, ensure_ascii=False)[:600]}")

        all_selects = page.locator('select')
        select_count = await all_selects.count()
        print(f"  {tag} found {select_count} select elements")

        if select_count >= 2:
            # Traditional <select> dropdowns
            if select_count >= 3:
                try:
                    await all_selects.nth(0).select_option("US")
                    print(f"  {tag} country: US")
                except Exception:
                    try:
                        await all_selects.nth(0).select_option(index=1)
                    except Exception:
                        pass
                await all_selects.nth(1).select_option(str(month))
                await all_selects.nth(2).select_option(str(day))
            else:
                await all_selects.nth(0).select_option(str(month))
                await all_selects.nth(1).select_option(str(day))

            year_input = page.locator(
                'input[id*="Year"], input[id*="year"], input[name*="Year"], '
                'input[name*="year"], input[type="text"]'
            ).first
            if await year_input.count() > 0:
                await year_input.fill(str(year))
        else:
            # New UI with combobox/dropdown (Chinese or English)
            print(f"  {tag} no <select>, trying new UI (combobox)...")

            month_names_en = ["", "January", "February", "March", "April", "May", "June",
                              "July", "August", "September", "October", "November", "December"]
            # Chinese month names: 1月, 2月, ... 12月
            month_names_cn = ["", "1月", "2月", "3月", "4月", "5月", "6月",
                              "7月", "8月", "9月", "10月", "11月", "12月"]
            # French month names
            month_names_fr = ["", "janvier", "février", "mars", "avril", "mai", "juin",
                              "juillet", "août", "septembre", "octobre", "novembre", "décembre"]

            # Find all visible comboboxes
            combos = page.locator('button[role="combobox"], [role="combobox"]')
            combo_count = await combos.count()
            print(f"  {tag} found {combo_count} comboboxes")

            # Strategy: identify combos by their text/aria-label/position
            # Typically order is: Country, Month, Day (country may already be set)
            month_filled = False
            day_filled = False

            for ci in range(combo_count):
                combo = combos.nth(ci)
                try:
                    box = await combo.bounding_box()
                    if not box or box['width'] < 10:
                        continue
                    combo_text = (await combo.text_content() or "").strip()
                    combo_label = (await combo.get_attribute("aria-label") or "").lower()
                    combo_id = (await combo.get_attribute("id") or "").lower()
                    info = f"text='{combo_text}' label='{combo_label}' id='{combo_id}'"
                    print(f"  {tag} combo[{ci}]: {info}")

                    # Multi-language detection: EN/CN/FR/ES/DE/PT
                    is_month = any(kw in combo_label for kw in ["month", "月", "mois", "mes", "monat", "mês"]) or \
                               any(kw in combo_id for kw in ["month", "birthmonth"]) or \
                               combo_text in ["月", "Month", "月份", "Mois", "Mes"]
                    is_day = any(kw in combo_label for kw in ["day", "日", "jour", "día", "tag", "dia"]) or \
                             any(kw in combo_id for kw in ["day", "birthday"]) or \
                             combo_text in ["日", "Day", "Jour", "Día"]

                    # Disambiguate: if id contains both "day" and "month" substrings, use the more specific match
                    if is_month and is_day:
                        # Prefer the specific keyword: "birthdaydropdown" → day, "birthmonthdropdown" → month
                        if "month" in combo_id:
                            is_day = False
                        elif "day" in combo_id:
                            is_month = False

                    # If text contains "月" or "日" at end, it's already showing a value
                    if not is_month and not is_day:
                        if combo_text.endswith("月") and len(combo_text) <= 3:
                            is_month = True
                        elif combo_text.endswith("日") and len(combo_text) <= 4:
                            is_day = True

                    if is_month and not month_filled:
                        await combo.click(force=True)
                        await asyncio.sleep(1)
                        # Try month option (Chinese → English → French → number)
                        month_opt = page.locator(f'[role="option"]:has-text("{month_names_cn[month]}")').first
                        if await month_opt.count() == 0:
                            month_opt = page.locator(f'[role="option"]:has-text("{month_names_en[month]}")').first
                        if await month_opt.count() == 0:
                            month_opt = page.locator(f'[role="option"]:has-text("{month_names_fr[month]}")').first
                        if await month_opt.count() == 0:
                            month_opt = page.locator(f'[role="option"]:has-text("{month}")').first
                        if await month_opt.count() > 0:
                            await month_opt.click()
                            month_filled = True
                            print(f"  {tag} month: {month}")
                        else:
                            await page.keyboard.type(str(month))
                            await asyncio.sleep(0.3)
                            await page.keyboard.press("Enter")
                            month_filled = True
                        await asyncio.sleep(1)

                    elif is_day and not day_filled:
                        await combo.click(force=True)
                        await asyncio.sleep(1)
                        # Try day option: exact match first to avoid "1" matching "10","11"...
                        day_str = str(day)
                        # Try exact match via all options
                        day_opt = None
                        try:
                            all_opts = page.locator('[role="option"]')
                            opt_count = await all_opts.count()
                            for oi in range(opt_count):
                                opt_text = (await all_opts.nth(oi).text_content() or "").strip()
                                if opt_text == day_str or opt_text == f"{day}日":
                                    day_opt = all_opts.nth(oi)
                                    break
                        except Exception:
                            pass
                        if not day_opt:
                            day_opt = page.locator(f'[role="option"]:has-text("{day}日")').first
                        if not day_opt or await day_opt.count() == 0:
                            day_opt = page.locator(f'[role="option"]:has-text("{day_str}")').first
                        if await day_opt.count() > 0:
                            await day_opt.click()
                            day_filled = True
                            print(f"  {tag} day: {day}")
                        else:
                            await page.keyboard.type(str(day))
                            await asyncio.sleep(0.3)
                            await page.keyboard.press("Enter")
                            day_filled = True
                        await asyncio.sleep(1)
                except Exception as e:
                    print(f"  {tag} combo[{ci}] error: {e}")

            if not month_filled or not day_filled:
                print(f"  {tag} WARNING: month_filled={month_filled}, day_filled={day_filled}")

            # Year input (text field)
            year_input = page.locator(
                '#BirthYearInput, [aria-label*="year" i], [aria-label*="年" i], '
                '[id*="Year" i], [id*="year" i], [placeholder*="年" i], '
                'input[type="text"][inputmode="numeric"], input[type="number"]'
            ).first
            # Fallback: find the text input that's NOT already filled
            if await year_input.count() == 0:
                all_text = page.locator('input[type="text"]')
                for ti in range(await all_text.count()):
                    inp = all_text.nth(ti)
                    val = await inp.input_value()
                    if not val:  # empty text input = likely year
                        year_input = inp
                        break
            if await year_input.count() > 0:
                await year_input.fill(str(year))
                print(f"  {tag} year: {year}")

        await asyncio.sleep(0.5)
        for sel in ['input[type="submit"]', 'button[type="submit"]', '#iSignupAction',
                    'button[id="iSignupAction"]', 'button:has-text("下一步")',
                    'button:has-text("Next")', 'button:has-text("next")',
                    'button:has-text("Suivant")']:
            btn = page.locator(sel).first
            if await btn.count() > 0:
                await btn.click(timeout=3000)
                print(f"  {tag} clicked next (bday): {sel}")
                break
        await asyncio.sleep(3)
        await page.screenshot(path=f"{SCREENSHOT_DIR}/outlook_{idx}_after_bday.png")

        # Step 4: Username/Gamertag (Chinese: 游戏标签/用户名)
        await asyncio.sleep(2)
        username_input = page.locator(
            'input[id*="displayName"], input[id*="gamertag"], input[name*="displayName"], '
            'input[placeholder*="name"], input[type="text"]'
        ).first
        if await username_input.count() > 0:
            page_text = await page.evaluate("() => document.body.innerText")
            if any(kw in page_text.lower() for kw in ["name", "gamertag", "nom", "pseudo", "surnom"]) or \
               any(kw in page_text for kw in ["用户名", "游戏标签", "显示名称"]):
                username = prefix[:8] + str(random.randint(100, 999))
                await username_input.fill(username)
                print(f"  {tag} username: {username}")
                await asyncio.sleep(0.5)
                for sel in ['input[type="submit"]', 'button[type="submit"]', '#iSignupAction',
                            'button:has-text("下一步")', 'button:has-text("Next")',
                            'button:has-text("Suivant")']:
                    btn = page.locator(sel).first
                    if await btn.count() > 0:
                        await btn.click(timeout=3000)
                        break
                await asyncio.sleep(3)

        # Step 5: First/Last Name + checkbox (Chinese: 姓/名)
        first_name, last_name = _consume_account_names()
        await asyncio.sleep(2)

        for _ in range(10):
            fname_input = page.locator(
                'input[name="FirstName"], input[id="FirstName"], input[name="firstNameInput"], '
                'input[id="firstNameInput"], input[aria-label*="first" i], input[placeholder*="first" i], '
                'input[aria-label*="名" i], input[placeholder*="名" i], '
                'input[aria-label*="prénom" i], input[placeholder*="prénom" i]'
            ).first
            lname_input = page.locator(
                'input[name="LastName"], input[id="LastName"], input[name="lastNameInput"], '
                'input[id="lastNameInput"], input[aria-label*="last" i], input[aria-label*="surname" i], '
                'input[placeholder*="last" i], input[aria-label*="姓" i], input[placeholder*="姓" i], '
                'input[aria-label*="nom de famille" i], input[placeholder*="nom de famille" i]'
            ).first
            if await fname_input.count() > 0 or await lname_input.count() > 0:
                break
            all_text_inputs = page.locator('input[type="text"]')
            if await all_text_inputs.count() >= 2:
                break
            await asyncio.sleep(1)

        if await fname_input.count() > 0:
            if await lname_input.count() > 0:
                await lname_input.fill(last_name)
            await fname_input.fill(first_name)
            print(f"  {tag} name: {first_name} {last_name}")
        else:
            all_text_inputs = page.locator('input[type="text"]')
            count = await all_text_inputs.count()
            if count >= 2:
                await all_text_inputs.nth(0).fill(last_name)
                await all_text_inputs.nth(1).fill(first_name)
                print(f"  {tag} name (generic): {first_name} {last_name}")

        checkbox = page.locator('input[type="checkbox"], [role="checkbox"]').first
        if await checkbox.count() > 0:
            try:
                checked = await checkbox.is_checked()
            except Exception:
                checked = False
            if not checked:
                await checkbox.click(force=True)
                print(f"  {tag} checkbox checked")

        await asyncio.sleep(0.5)
        for sel in ['input[type="submit"]', 'button[type="submit"]', '#iSignupAction',
                    'button[id="iSignupAction"]', 'button:has-text("Next")',
                    'button:has-text("下一步")', 'button:has-text("Suivant")']:
            btn = page.locator(sel).first
            if await btn.count() > 0:
                await btn.click(timeout=3000)
                print(f"  {tag} clicked next (name): {sel}")
                break
        await asyncio.sleep(3)
        await page.screenshot(path=f"{SCREENSHOT_DIR}/outlook_{idx}_after_name.png")

        # Step 6: CAPTCHA handling
        print(f"  {tag} checking for captcha...")
        await asyncio.sleep(3)

        arkose_solved = False
        press_count = 0
        # headless: 5 presses max (abort quickly on fail)
        # browser: 15 presses max (keep trying; PX sometimes passes after retries)
        max_press = 5 if captcha_early_abort else 15
        # Allow caller to cap presses tighter via env (e.g. bs_register_step1
        # sets this to 3 so failed PX checks fail-fast and we move on to
        # lqqq/backup instead of burning ~3 min per dud signup).
        _env_max_press = os.environ.get("OUTLOOK_REG_MAX_PRESS", "").strip()
        if _env_max_press.isdigit():
            max_press = min(max_press, int(_env_max_press))
        no_btn_rounds = 0
        had_captcha = False          # 是否真的出现过 captcha（避免一上来误判已通过）
        gone_rounds = 0              # captcha 消失后连续多少轮仍停在 signup（等跳转）

        async def _captcha_visible():
            """页面上是否还有【可交互】的 PerimeterX 按住验证（按住按钮 / hsprotect iframe）。
            captcha 通过后会变成 Loading 转圈、这些元素消失 -> 返回 False。"""
            try:
                for sel in ['button:has-text("Press and hold")', 'button:has-text("Appuyer et maintenir")',
                            'button:has-text("按住")', 'button:has-text("长按")',
                            'button:has-text("Halten")', '#px-captcha']:
                    el = page.locator(sel).first
                    if await el.count() > 0:
                        b = await el.bounding_box()
                        if b and b['width'] > 30:
                            return True
                ifr = page.locator('iframe[src*="hsprotect.net"], iframe[src*="arkose"], iframe[src*="funcaptcha"]')
                for hi in range(await ifr.count()):
                    b = await ifr.nth(hi).bounding_box()
                    if b and b['width'] > 50 and b['height'] > 30:
                        return True
            except Exception:
                pass
            return False

        # headless: 90 s captcha window; browser: 240 s (multiple press rounds)
        _captcha_rounds = 30 if captcha_early_abort else 80
        # When max_press is small, shrink the wait loop too — otherwise we'd
        # exhaust presses then idle for the remaining captcha window.
        # Rough budget: ~10s per press cycle. +20s slack for first solver call.
        _capped_rounds = max(8, max_press * 4 + 8)
        _captcha_rounds = min(_captcha_rounds, _capped_rounds)

        for wait_round in range(_captcha_rounds):
            try:
                page_text = (await page.evaluate("() => document.body.innerText")).lower()
                current_url = page.url.lower()
            except Exception:
                await asyncio.sleep(3)
                try:
                    page_text = (await page.evaluate("() => document.body.innerText")).lower()
                    current_url = page.url.lower()
                except Exception:
                    current_url = page.url.lower()
                    if "signup" not in current_url:
                        break
                    continue

            # —— captcha 通过判定（精确按 host）——
            # 坑：captcha 过后页面跳到 privacynotice.account.microsoft.com/notice?ru=...，
            # 其 ru= 参数里带 "signup" 字样，旧的裸 "signup" 子串判定 -> 误以为还在 signup
            # -> 一直 retrying presses 直到超时。改为按真实 host 判断是否已离开 signup 表单。
            on_signup_form = ("signup.live.com" in current_url) and ("privacynotice" not in current_url)
            if not on_signup_form and any(h in current_url for h in [
                    "privacynotice", "account.microsoft.com", "account.live.com",
                    "outlook.live.com", "outlook.office", "login.live.com/oauth20"]):
                print(f"  {tag} captcha passed, left signup -> {current_url[:70]}")
                break

            # Success checks
            if "outlook" in current_url and "signup" not in current_url and "login" not in current_url:
                print(f"  {tag} registration complete!")
                break
            if "welcome" in page_text or "inbox" in page_text or "account has been created" in page_text:
                print(f"  {tag} registration complete!")
                break
            if "signup" not in current_url and "live.com" in current_url:
                print(f"  {tag} left signup: {current_url[:60]}")
                break

            # captcha 消失判定：过验证后页面变 "Loading..." 转圈、按住按钮/iframe 消失，
            # 但 URL 可能还没跳转（异步）。此时不该再按压/超时 —— 标记已过，进入等跳转模式。
            if had_captcha:
                if await _captcha_visible():
                    gone_rounds = 0
                else:
                    gone_rounds += 1
                    if gone_rounds == 1:
                        print(f"  {tag} captcha 元素已消失（验证通过/Loading），等待页面跳转…")
                    # 轻推一下提交按钮，催收尾
                    if gone_rounds % 3 == 0:
                        for sel in ['#iSignupAction', 'input[type="submit"]', 'button[type="submit"]']:
                            try:
                                b = page.locator(sel).first
                                if await b.count() > 0 and await b.is_visible():
                                    await b.click(timeout=3000); break
                            except Exception:
                                pass
                    # 等够 ~20 轮(≈60s)仍没跳转，去 post-captcha 收尾兜底（不再傻等超时）
                    if gone_rounds >= 20:
                        print(f"  {tag} captcha 已过但久未跳转，进入收尾流程")
                        break
                    await asyncio.sleep(3)
                    continue

            # Account blocked detection (CN/EN/FR)
            if any(kw in page_text for kw in [
                "帐户创建已被阻止", "已被阻止", "阻止创建",
                "account creation has been blocked", "has been blocked", "account has been suspended",
                "création de compte a été bloquée", "a été bloquée", "bloquée",
                "unusual activity", "异常活动", "activité inhabituelle",
            ]):
                print(f"  {tag} BLOCKED: account creation blocked by Microsoft")
                await page.screenshot(path=f"{SCREENSHOT_DIR}/outlook_{idx}_blocked.png")
                return None, None

            # FIDO/passkey - skip
            if "fido" in current_url or "passkey" in current_url:
                for sel in ['a:has-text("Skip")', 'button:has-text("Skip")', 'a:has-text("No thanks")',
                            'button:has-text("No thanks")', 'button:has-text("Cancel")', '#skipBtn']:
                    btn = page.locator(sel).first
                    if await btn.count() > 0:
                        try:
                            await btn.click(timeout=3000)
                            break
                        except Exception:
                            pass
                else:
                    try:
                        await page.evaluate("""() => {
                            for (const l of document.querySelectorAll('a, button')) {
                                const t = l.textContent.toLowerCase();
                                if (t.includes('skip') || t.includes('no thanks') || t.includes('cancel')) { l.click(); return; }
                            }
                        }""")
                    except Exception:
                        pass
                await asyncio.sleep(3)
                continue

            # Privacy notice
            if "privacynotice" in current_url:
                await asyncio.sleep(2)
                for label in ['OK', 'Accept', 'Continue', 'Next', 'I agree', 'Got it']:
                    btn = page.locator(f'button:has-text("{label}"), input[value="{label}"], a:has-text("{label}")').first
                    if await btn.count() > 0:
                        try:
                            await btn.click(timeout=3000)
                            break
                        except Exception:
                            pass
                await asyncio.sleep(3)
                continue

            # PerimeterX press-and-hold
            if press_count < max_press:
                pressed = False
                target_box = None

                # 定位「按住」按钮。诊断已确认：按钮是可见 hsprotect iframe 内的 #px-captcha
                # 元素(box 如 y485~527,height 42)，page.locator 穿不进跨域 iframe，必须遍历
                # page.frames 在 frame 内取 #px-captcha 的真实坐标。优先用它(box_is_button=True)，
                # 拿不到才退回整个 iframe 框按比例。
                box_is_button = False
                # 1) frame 内真按钮 #px-captcha（取可见的那个 frame：width>0）
                for f in page.frames:
                    if f == page.main_frame or 'hsprotect.net' not in (f.url or ''):
                        continue
                    try:
                        px = f.locator('#px-captcha').first
                        if await px.count() > 0:
                            b = await px.bounding_box()
                            if b and b['width'] > 30 and b['height'] > 8:
                                target_box = b; box_is_button = True
                                break
                    except Exception:
                        pass
                # 2) 退回整个可见 hsprotect iframe 框
                if not target_box:
                    try:
                        hs = page.locator('iframe[src*="hsprotect.net"]')
                        for hi in range(await hs.count()):
                            b = await hs.nth(hi).bounding_box()
                            if b and b['width'] > 50 and b['height'] > 30:
                                target_box = b
                                break
                    except Exception:
                        pass

                if target_box and target_box['width'] > 30 and target_box['height'] >= 8:
                    press_count += 1
                    pressed = True
                    had_captcha = True   # 出现过 captcha，供「消失=已通过」判定使用
                    bx, by, bw, bh = target_box['x'], target_box['y'], target_box['width'], target_box['height']
                    if box_is_button:
                        # target_box 就是真按钮 #px-captcha：按其中心 + 小随机抖动
                        cx = bx + bw * random.uniform(0.40, 0.60)
                        cy = by + bh * random.uniform(0.40, 0.60)
                    else:
                        # 退回整个 iframe 框：按钮在中部窄带（实测 0.48-0.62 命中）
                        cx = bx + bw * random.uniform(0.42, 0.58)
                        cy = by + bh * random.uniform(0.48, 0.62)
                    print(f"  {tag} press #{press_count}: ({cx:.0f},{cy:.0f}){' [btn]' if box_is_button else ' [box]'}")
                    if px_press_screenshots:
                        try:
                            await page.screenshot(path=f"{SCREENSHOT_DIR}/outlook_{idx}_before_press_{press_count}.png")
                        except Exception:
                            pass

                    # Bezier mouse movement
                    sx, sy = random.uniform(200, 800), random.uniform(200, 400)
                    await page.mouse.move(sx, sy)
                    await asyncio.sleep(random.uniform(0.3, 0.8))
                    steps = random.randint(15, 30)
                    ctrl_x = (sx + cx) / 2 + random.uniform(-100, 100)
                    ctrl_y = (sy + cy) / 2 + random.uniform(-80, 80)
                    for step in range(1, steps + 1):
                        t = step / steps
                        mx = (1 - t) ** 2 * sx + 2 * (1 - t) * t * ctrl_x + t ** 2 * cx + random.uniform(-1.5, 1.5)
                        my = (1 - t) ** 2 * sy + 2 * (1 - t) * t * ctrl_y + t ** 2 * cy + random.uniform(-1.5, 1.5)
                        await page.mouse.move(mx, my)
                        await asyncio.sleep(random.uniform(0.005, 0.025))

                    await asyncio.sleep(random.uniform(0.1, 0.3))
                    await page.mouse.down()

                    # 拟人按住：关键是按住到进度条走满(captcha 消失)才松手 —— 不能固定短时长，
                    # 太短(3-5s)进度没满就松手 => 不过。策略：持续按住并每 ~0.5s 检测 captcha 是否
                    # 消失(进度满/通过)，一消失立刻松手；设一个较长上限(~14s)兜底防止卡死。
                    max_hold = random.uniform(9.0, 11.0)
                    hold_start = asyncio.get_event_loop().time()
                    drift_phase = random.uniform(0, 2 * math.pi)
                    drift_freq = random.uniform(1.5, 2.8)
                    last_chk = 0.0
                    passed_in_hold = False
                    while True:
                        elapsed = asyncio.get_event_loop().time() - hold_start
                        if elapsed >= max_hold:
                            break
                        ph = drift_phase + elapsed * drift_freq
                        await page.mouse.move(cx + 1.0 * math.sin(ph), cy + 0.6 * math.cos(ph * 1.3))
                        # 进度满后 captcha 消失就松手（但至少按住 1.5s，避免误判刚加载的空隙）
                        if elapsed > 1.5 and elapsed - last_chk > 0.5:
                            last_chk = elapsed
                            try:
                                if not await _captcha_visible():
                                    passed_in_hold = True
                                    break
                            except Exception:
                                pass
                        await asyncio.sleep(random.uniform(0.03, 0.08))

                    await asyncio.sleep(random.uniform(0.05, 0.18))
                    await page.mouse.up()
                    held = asyncio.get_event_loop().time() - hold_start
                    print(f"  {tag} held {held:.1f}s{' (passed)' if passed_in_hold else ''}")
                    await asyncio.sleep(random.uniform(2, 4))

                    try:
                        await page.screenshot(path=f"{SCREENSHOT_DIR}/outlook_{idx}_hold_{press_count}.png")
                    except Exception:
                        pass
                    if px_press_screenshots:
                        try:
                            await page.screenshot(path=f"{SCREENSHOT_DIR}/outlook_{idx}_after_press_{press_count}.png")
                        except Exception:
                            pass
                else:
                    no_btn_rounds += 1
                    # Scan frames for clickable buttons
                    try:
                        for f in page.frames:
                            if f == page.main_frame:
                                continue
                            frame_url = f.url.lower()
                            if frame_url == "about:blank" or "cfp.microsoft.com" in frame_url:
                                continue
                            try:
                                btns = f.locator('button, [role="button"], input[type="button"], input[type="submit"]')
                                for bi in range(await btns.count()):
                                    box = await btns.nth(bi).bounding_box()
                                    if box and box['width'] > 30 and box['height'] > 20:
                                        x = box['x'] + box['width'] / 2
                                        y = box['y'] + box['height'] / 2
                                        press_count += 1
                                        await page.mouse.move(x, y)
                                        await asyncio.sleep(0.3)
                                        await page.mouse.down()
                                        await asyncio.sleep(18)
                                        await page.mouse.up()
                                        pressed = True
                                        await asyncio.sleep(5)
                                        break
                                if pressed:
                                    break
                            except Exception:
                                continue
                    except Exception:
                        pass

                if pressed:
                    no_btn_rounds = 0

            # Main page captcha buttons
            if press_count < max_press and no_btn_rounds >= 3:
                try:
                    main_btns = page.locator('#hipTemplateContainer button, #HipPaneForm button, [id*="hip"] button')
                    for bi in range(await main_btns.count()):
                        box = await main_btns.nth(bi).bounding_box()
                        if box and box['width'] > 20:
                            press_count += 1
                            await main_btns.nth(bi).click(timeout=3000)
                            await asyncio.sleep(5)
                            no_btn_rounds = 0
                            break
                except Exception:
                    pass

            # Try submit
            if no_btn_rounds >= 8 and no_btn_rounds % 8 == 0:
                try:
                    for sel in ['#iSignupAction', 'input[type="submit"]', 'button[type="submit"]']:
                        submit = page.locator(sel).first
                        if await submit.count() > 0 and await submit.is_visible():
                            await submit.click(timeout=3000)
                            await asyncio.sleep(5)
                            break
                except Exception:
                    pass

            # 按满次数：两个打码器(capsolver-px/ezcaptcha-px)对 MS 这个 PerimeterX
            # 按住验证都没用(类型不支持/解不出)，已移除。按满后给一个短观察窗等跳转：
            # 若手动按住其实已过，循环顶部「captcha 消失=已通过」会接管收尾；若仍可见(没过)，
            # 观察窗内不再按压、等满 ~24s 就快速放弃，不空等到 captcha timeout。
            if press_count >= max_press and not arkose_solved:
                arkose_solved = True
                arkose_wait_start = wait_round
                print(f"  {tag} 按满 {max_press} 次，停止按压，等待页面跳转")
            if arkose_solved and had_captcha:
                # 仍能看到 captcha = 没过；给 8 轮(~24s)缓冲后快速放弃
                if await _captcha_visible():
                    if wait_round - arkose_wait_start >= 8:
                        print(f"  {tag} 按满仍未通过，快速放弃本号")
                        await page.screenshot(path=f"{SCREENSHOT_DIR}/outlook_{idx}_press_fail.png")
                        return None, None

            if wait_round % 5 == 0:
                await page.screenshot(path=f"{SCREENSHOT_DIR}/outlook_{idx}_wait_{wait_round}.png")
                print(f"  {tag} waiting... ({wait_round * 3}s)")

            await asyncio.sleep(3)
        else:
            print(f"  {tag} captcha timeout")
            await page.screenshot(path=f"{SCREENSHOT_DIR}/outlook_{idx}_timeout.png")
            return None, None

        # Post-captcha pages
        for retry in range(10):
            current_url = page.url.lower()
            # 精确判断：已离开 signup 表单且不在隐私声明页 = 收尾完成（同样避免 ru= 里
            # 的 "signup" 子串误伤）。
            on_signup_form = ("signup.live.com" in current_url) and ("privacynotice" not in current_url)
            if not on_signup_form and "privacynotice" not in current_url:
                break
            for label in ['OK', 'Accept', 'Continue', 'Next', 'I agree', 'Got it', 'Agree']:
                btn = page.locator(f'button:has-text("{label}"), input[value="{label}"], a:has-text("{label}")').first
                if await btn.count() > 0:
                    try:
                        await btn.click(timeout=3000)
                        break
                    except Exception:
                        pass
            await asyncio.sleep(3)

        if not verify_registered_outlook(email, password, tag):
            print(f"  {tag} verification failed, discarding account")
            return None, None

        print(f"  {tag} OK: {email} / {password}")
        return email, password

    except Exception as e:
        print(f"  {tag} FAILED: {e}")
        try:
            await page.screenshot(path=f"{SCREENSHOT_DIR}/outlook_{idx}_error.png")
        except Exception:
            pass
        return None, None


# ======================== Protocol Mode (pure HTTP) ========================

def _proxy_for_requests(proxy_str):
    """Convert proxy string to requests proxies dict."""
    if not proxy_str:
        return None
    raw = str(proxy_str or "").strip()
    if "://" not in raw:
        parts = raw.split(":")
        if len(parts) >= 4 and str(parts[1]).isdigit():
            host = parts[0]
            port = parts[1]
            username = parts[2]
            password = ":".join(parts[3:])
            auth = f"{username}:{password}@" if username else ""
            url = f"socks5h://{auth}{host}:{port}"
            return {"http": url, "https": url}
        if len(parts) == 2 and str(parts[1]).isdigit():
            url = f"socks5h://{parts[0]}:{parts[1]}"
            return {"http": url, "https": url}
    p = BitBrowserClient._parse_proxy(raw)
    if not p:
        return None
    auth = f"{p['username']}:{p['password']}@" if p.get("username") else ""
    scheme = p.get("type", "http")
    if scheme == "socks5":
        scheme = "socks5h"
    url = f"{scheme}://{auth}{p['host']}:{p['port']}"
    return {"http": url, "https": url}


@contextmanager
def _proxy_env_context(proxy_str):
    proxy_env_keys = ("HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy")
    old = {key: os.environ.get(key) for key in proxy_env_keys}
    try:
        for key in proxy_env_keys:
            os.environ.pop(key, None)
        proxies = _proxy_for_requests(proxy_str) or {}
        proxy_url = proxies.get("https") or proxies.get("http") or ""
        if proxy_url:
            os.environ["HTTP_PROXY"] = os.environ["HTTPS_PROXY"] = proxy_url
            os.environ["http_proxy"] = os.environ["https_proxy"] = proxy_url
        yield
    finally:
        for key, value in old.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value


def _proxy_for_playwright(proxy_str):
    """Convert proxy string to Playwright proxy dict."""
    if not proxy_str:
        return None
    p = BitBrowserClient._parse_proxy(proxy_str)
    if not p:
        return None
    result = {"server": f"{p.get('type', 'http')}://{p['host']}:{p['port']}"}
    if p.get("username"):
        result["username"] = p["username"]
        result["password"] = p["password"]
    return result
