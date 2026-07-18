"""Standalone Outlook registration loop. Continuously registers fresh
outlook accounts via BitBrowser + standalone register_outlook script, and
writes each success to _data_bundle/_outlook_pool/ as one JSON file per
record (email + password + session cookies).

The Replit batch (_batch_register.py / bs_register_step1.py) consumes these
via the `pool` email source — fully decoupled, so a slow self-reg attempt
never blocks the Replit signup pipeline.

Usage:
  python outlook_reg_loop.py                       # loop forever
  python outlook_reg_loop.py --count 20            # 20 attempts then exit
  python outlook_reg_loop.py --target-pool 10      # stop refilling once pool >= 10
  python outlook_reg_loop.py --max-press 5         # OUTLOOK_REG_MAX_PRESS
  python outlook_reg_loop.py --sleep 5             # gap between attempts (s)

Standalone mode uses a random proxy from --proxy-file / OUTLOOK_PROXIES and
puts it directly into the BitBrowser profile. Set SELF_REG_SCRIPT_PATH to
override standalone script location.
"""
from __future__ import annotations

import argparse
import asyncio
import json
import os
import random
import re
import sys
import time
import importlib.util
import urllib.request
from contextlib import contextmanager
from datetime import datetime
from types import SimpleNamespace

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

ARTIFACT_DIR = os.path.dirname(os.path.abspath(__file__))
POOL_DIR = os.path.join(ARTIFACT_DIR, "_outlook_pool")
# 账号注册侧消费的池（common/emails.next_email 读取），格式 email----password----token----clientid
EMAILS_POOL = os.path.join(ARTIFACT_DIR, "emails.txt")
STANDALONE_PATH = os.environ.get(
    "SELF_REG_SCRIPT_PATH",
    os.path.join(ARTIFACT_DIR, "register_outlook_standalone.py"),
)
RUOYI_PATH = os.environ.get(
    "SELF_REG_RUOYI_PATH",
    os.path.join(ARTIFACT_DIR, "register_outlook_ruoyi.py"),
)
CAMONFOX_PATH = os.environ.get(
    "SELF_REG_CAMONFOX_PATH",
    os.path.join(ARTIFACT_DIR, "register_outlook_camonfox.py"),
)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def log(msg, level="INFO"):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] [{level}] {msg}", flush=True)


def _env_bool(name, default=False):
    value = os.environ.get(name)
    if value is None:
        return bool(default)
    return str(value).strip().lower() not in ("0", "false", "no", "off", "")


_PROXY_ENV_KEYS = ("HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy")


def _looks_like_local_proxy(value):
    v = (value or "").strip().lower()
    if not v:
        return False
    return ("127.0.0.1" in v or "localhost" in v) and any(
        port in v for port in (":7890", ":7897", ":7898")
    )


def clear_inherited_local_proxy_env():
    """养号循环不用本机转发代理，避免覆盖账号专用代理。"""
    removed = []
    for key in _PROXY_ENV_KEYS:
        value = os.environ.get(key)
        if _looks_like_local_proxy(value):
            os.environ.pop(key, None)
            removed.append(key)
    if removed:
        log(f"ignored inherited local proxy env: {', '.join(sorted(set(removed)))}")


def parse_proxy(proxy_str):
    """Parse user:pass@host:port / host:port / socks5://... into BitBrowser fields."""
    if not proxy_str:
        return None
    proxy_type = "http"
    s = str(proxy_str).strip()
    for prefix in ("socks5://", "socks4://", "http://", "https://"):
        if s.lower().startswith(prefix):
            proxy_type = prefix.split("://", 1)[0]
            s = s[len(prefix):]
            break
    s = s.replace(",", "@", 1) if "@" not in s and "," in s else s
    m = re.match(r"^(.+):(.+)@(.+):(\d+)$", s)
    if m:
        return {
            "type": proxy_type,
            "username": m.group(1),
            "password": m.group(2),
            "host": m.group(3),
            "port": m.group(4),
        }
    m = re.match(r"^(.+):(\d+)$", s)
    if m:
        return {"type": proxy_type, "host": m.group(1), "port": m.group(2)}
    return None


def mask_proxy(proxy_str):
    p = parse_proxy(proxy_str)
    if not p:
        return "***" if proxy_str else "noproxy"
    user = p.get("username")
    auth = f"{user[:8]}...@" if user else ""
    return f"{p.get('type', 'http')}://{auth}{p['host']}:{p['port']}"


def proxy_url_for_requests(proxy_str):
    p = parse_proxy(proxy_str)
    if not p:
        return ""
    auth = f"{p['username']}:{p['password']}@" if p.get("username") else ""
    return f"{p.get('type', 'http')}://{auth}{p['host']}:{p['port']}"


@contextmanager
def account_proxy_env(proxy_str):
    """Temporarily route trust_env=True HTTP helpers through the account proxy."""
    old = {k: os.environ.get(k) for k in _PROXY_ENV_KEYS}
    try:
        for key in _PROXY_ENV_KEYS:
            os.environ.pop(key, None)
        url = proxy_url_for_requests(proxy_str)
        if url:
            os.environ["HTTP_PROXY"] = os.environ["HTTPS_PROXY"] = url
            os.environ["http_proxy"] = os.environ["https_proxy"] = url
        yield
    finally:
        for key, value in old.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value


def load_proxy_pool(proxy_file="", mod=None):
    pool = []
    if proxy_file and os.path.isfile(proxy_file):
        with open(proxy_file, "r", encoding="utf-8") as f:
            pool.extend(
                line.strip() for line in f
                if line.strip() and not line.lstrip().startswith("#")
            )
    elif proxy_file:
        log(f"proxy file not found: {proxy_file}", "WARN")
    if not pool and mod is not None:
        pool.extend(list(getattr(mod, "DEFAULT_PROXIES", []) or []))
    return pool


def load_standalone():
    if not os.path.isfile(STANDALONE_PATH):
        log(f"standalone not found at {STANDALONE_PATH}", "ERR")
        sys.exit(1)
    spec = importlib.util.spec_from_file_location("_self_reg_standalone", STANDALONE_PATH)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    log(f"loaded standalone from {STANDALONE_PATH}")
    return m


def load_ruoyi():
    if not os.path.isfile(RUOYI_PATH):
        log(f"ruoyi backend not found at {RUOYI_PATH}", "ERR")
        sys.exit(1)
    spec = importlib.util.spec_from_file_location("_self_reg_ruoyi", RUOYI_PATH)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    log(f"loaded ruoyi from {RUOYI_PATH}")
    return m


def load_camonfox():
    if not os.path.isfile(CAMONFOX_PATH):
        log(f"camonfox backend not found at {CAMONFOX_PATH}", "ERR")
        sys.exit(1)
    spec = importlib.util.spec_from_file_location("_self_reg_camonfox", CAMONFOX_PATH)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    log(f"loaded camonfox from {CAMONFOX_PATH}")
    return m


def load_engine(engine):
    if engine == "ruoyi":
        return load_ruoyi()
    if engine == "camonfox":
        return load_camonfox()
    return load_standalone()


BB_API = os.environ.get("BITBROWSER_API", "http://127.0.0.1:54345")
# Match bs_register_step1 — user's BitBrowser has Chromium 146 not 130.
BB_CORE_VERSION = os.environ.get("BB_CORE_VERSION", "146")


def _fingerprint_provider():
    return (
        os.environ.get("FINGERPRINT_BROWSER")
        or os.environ.get("BROWSER_PROVIDER")
        or "bitbrowser"
    ).strip().lower()


def _bb_call(path, body):
    data = json.dumps(body).encode()
    req = urllib.request.Request(
        f"{BB_API}{path}", data=data, method="POST",
        headers={"Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.loads(r.read())


def apply_bitbrowser_proxy(data, proxy_str):
    p = parse_proxy(proxy_str)
    if not p:
        data["proxyType"] = "noproxy"
        return False
    data["proxyType"] = p.get("type", "http")
    data["host"] = p["host"]
    data["port"] = p["port"]
    if p.get("username"):
        data["proxyUserName"] = p["username"]
    if p.get("password"):
        data["proxyPassword"] = p["password"]
    return True


def bb_create_for_outlook_reg(name, proxy_str=None):
    """Mirror bs_register_step1.bb_create_ephemeral so we share the working
    fingerprint config and the installed Chromium version, while putting the
    selected account proxy directly into the BitBrowser profile."""
    if _fingerprint_provider() in {"adspower", "ads_power", "ads"}:
        from bitbrowser import BitBrowser
        kwargs = {
            "remark": "outlook reg loop auto-deleted after use",
            "platform": "https://outlook.live.com",
            "platformIcon": "outlook.live.com",
            "proxyMethod": 2,
            "browserFingerPrint": {
                "ostype": "PC",
                "os": "Win32",
                "coreVersion": BB_CORE_VERSION,
                "isIpCreateTimeZone": True,
                "isIpCreateLanguage": True,
                "isIpCreateDisplayLanguage": True,
                "isIpCreatePosition": True,
                "isIpCountry": True,
            },
        }
        apply_bitbrowser_proxy(kwargs, proxy_str)
        return BitBrowser().create_browser(name=name, **kwargs)
    body = {
        "name": name,
        "remark": "outlook reg loop — auto-deleted after use",
        "platform": "https://outlook.live.com",
        "platformIcon": "outlook.live.com",
        "proxyMethod": 2,
        "browserFingerPrint": {
            "ostype": "PC",
            "os": "Win32",
            "coreVersion": BB_CORE_VERSION,
            "isIpCreateTimeZone": True,
            "isIpCreateLanguage": True,
            "isIpCreateDisplayLanguage": True,
            "isIpCreatePosition": True,
            "isIpCountry": True,
        },
    }
    apply_bitbrowser_proxy(body, proxy_str)
    log(f"BitBrowser proxy: {mask_proxy(proxy_str)}")
    r = _bb_call("/browser/update", body)
    if not r.get("success"):
        raise RuntimeError(f"/browser/update failed: {r}")
    data = r.get("data") or {}
    pid = data.get("id") or data.get("browserId")
    if not pid:
        raise RuntimeError(f"/browser/update returned no id: {data}")
    return pid


def count_pool():
    if not os.path.isdir(POOL_DIR):
        return 0
    try:
        return sum(1 for f in os.listdir(POOL_DIR) if f.endswith(".json"))
    except Exception:
        return 0


def extract_graph_for_account(email, password, attempts=3, proxy_str=None):
    """Return Graph token data for a freshly registered Outlook account."""
    try:
        from extract_graph_tokens import get_graph_token
        for attempt in range(attempts):
            with account_proxy_env(proxy_str):
                res = get_graph_token(email, password)
            if res and res.get("refresh_token"):
                graph = {
                    "refresh_token": res["refresh_token"],
                    "client_id": res.get("client_id") or "",
                }
                log(f"graph token extracted for {email}", "OK")
                return graph
            if attempt < attempts - 1:
                log(f"graph token attempt {attempt + 1}/{attempts} failed, retry: {email}", "WARN")
                time.sleep(3 * (attempt + 1))
        log(f"graph token missing after {attempts} attempts: {email}", "WARN")
    except Exception as exc:
        log(f"graph token extraction error: {type(exc).__name__}: {exc}", "WARN")
    return None


def append_graph_account_to_emails_pool(email, password, graph):
    """Append only Graph-ready accounts to emails.txt."""
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
        log(f"append_to_emails_pool failed: {type(exc).__name__}: {exc}", "WARN")
        return False


def append_to_emails_pool(email, password):
    """把成功号桥接进 emails.txt 池，供账号注册侧 common/emails.next_email 消费。
    注册成功后立即用纯 HTTP OAuth 抽 Graph refresh_token（extract_graph_tokens.get_graph_token），
    写真 token/client_id —— 之后 ChatGPT 取码全走 Graph API，免浏览器登录/取码。
    抽取失败（偶发风控/网络）才回退占位符 fresh，消费侧届时退化到浏览器取码。"""
    token = client_id = "fresh"
    graph = globals().pop("_CURRENT_GRAPH_ACCOUNT", None)
    if graph is not None:
        return append_graph_account_to_emails_pool(email, password, graph)
    try:
        from extract_graph_tokens import get_graph_token
        # 抽取经代理偶发 TLS 抖动(SSLEOFError)，单试一次一抖就回退 fresh、白丢 token 快路；
        # 这里重试 3 次(短退避)，绝大多数抖动二/三次就过。
        res = None
        for _try in range(3):
            res = get_graph_token(email, password)
            if res and res.get("refresh_token"):
                break
            if _try < 2:
                log(f"graph token 抽取第{_try+1}次未成，重试: {email}", "WARN")
                time.sleep(3 * (_try + 1))
        if res and res.get("refresh_token"):
            token = res["refresh_token"]
            client_id = res.get("client_id") or "fresh"
            log(f"graph token extracted for {email}", "OK")
        else:
            log(f"graph token 抽取失败(3 次)，回退 fresh: {email}", "WARN")
    except Exception as e:
        log(f"graph token 抽取异常，回退 fresh: {type(e).__name__}: {e}", "WARN")
    try:
        existing = set()
        if os.path.isfile(EMAILS_POOL):
            with open(EMAILS_POOL, encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#"):
                        existing.add(line.split("----")[0].strip().lower())
        if email.lower() in existing:
            return
        with open(EMAILS_POOL, "a", encoding="utf-8") as f:
            f.write(f"{email}----{password}----{token}----{client_id}\n")
        log(f"emails.txt += {email} (token={'yes' if token != 'fresh' else 'fresh'})", "OK")
    except Exception as e:
        log(f"append_to_emails_pool failed: {type(e).__name__}: {e}", "WARN")


def write_record(record):
    os.makedirs(POOL_DIR, exist_ok=True)
    safe = record["email"].replace("@", "_at_").replace("/", "_")
    fname = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:18] + f"_{safe}.json"
    tmp = os.path.join(POOL_DIR, fname + ".tmp")
    dst = os.path.join(POOL_DIR, fname)
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(record, f, ensure_ascii=False, indent=2)
        os.rename(tmp, dst)
    except Exception as e:
        log(f"write_record FAILED: {type(e).__name__}: {e}  (tmp={tmp})", "ERR")
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass
        raise
    # Verify it actually landed.
    if os.path.isfile(dst):
        sz = os.path.getsize(dst)
        log(f"write_record OK: {dst}  ({sz} bytes)", "OK")
    else:
        log(f"write_record sus: {dst} missing right after rename!", "ERR")
    return fname


async def _run_outlook_on_ctx(mod, ctx, idx):
    """Scrub residual state -> 新页注册 -> 导出 outlook 相关 cookie。"""
    # Scrub Chromium residual state so signup.live.com doesn't see a
    # stale identity from a previous session.
    try:
        await ctx.clear_cookies()
        for _pg in ctx.pages:
            try:
                c = await ctx.new_cdp_session(_pg)
                await c.send("Network.clearBrowserCookies")
                await c.send("Network.clearBrowserCache")
                try: await c.detach()
                except Exception: pass
                break
            except Exception:
                pass
    except Exception:
        pass
    page = await ctx.new_page()
    email, password = await mod.register_outlook(page, ctx, idx)
    cookies = []
    if email:
        try:
            all_cookies = await ctx.cookies()
            keep_domains = (
                "outlook.", "live.com", "microsoftonline.",
                "microsoft.com", "office.com", ".office365.",
                "msn.com", "bing.com", "mail.live.com",
            )
            cookies = [
                c for c in all_cookies
                if any(d in (c.get("domain") or "") for d in keep_domains)
            ]
        except Exception as e:
            log(f"cookie export failed: {e}", "WARN")
    return email, password, cookies


async def one_attempt_standalone(mod, proxy_str, idx):
    """Mirrors bs_register_step1.fetch_email_from_self_register's inline
    flow, but doesn't carry the breaker state — we're a dedicated loop and
    want to keep trying."""
    profile_id = None
    bb = mod.BitBrowserClient()
    try:
        ts = datetime.now().strftime("%m%d_%H%M%S")
        for _r in range(5):
            try:
                # Use our own create that picks coreVersion=146 (matches the
                # BitBrowser install on this machine). Standalone's hardcoded
                # 130 makes BB return 502.
                profile_id = bb_create_for_outlook_reg(f"outlook_loop_{ts}_{idx}", proxy_str=proxy_str)
                break
            except Exception as e:
                m = str(e)
                if "最大" in m or "超过" in m:
                    log("BitBrowser quota — cleanup_browsers(keep=2)", "WARN")
                    try: bb.cleanup_browsers(keep=2)
                    except Exception: pass
                    await asyncio.sleep(3)
                    continue
                if _r >= 4:
                    raise
                log(f"create_browser err (try {_r+1}/5): {m[:200]}", "WARN")
                await asyncio.sleep(3 + _r)
        if not profile_id:
            return None, None, []
        info = bb.open_browser(profile_id)
        ws = info.get("ws", "")
        if not ws:
            return None, None, []
        from playwright.async_api import async_playwright as _apw
        async with _apw() as p:
            browser = await p.chromium.connect_over_cdp(ws)
            ctx = browser.contexts[0] if browser.contexts else await browser.new_context()
            email, password, cookies = await _run_outlook_on_ctx(mod, ctx, idx)
        return email, password, cookies
    finally:
        if profile_id:
            try:
                bb.close_browser(profile_id)
                await asyncio.sleep(2)
                bb.delete_browser(profile_id)
            except Exception:
                pass


def _one_attempt_ruoyi(
    mod,
    proxy_file,
    idx,
    timeout,
    max_press,
    confirm_before_register,
    headless,
    email_suffixes="",
    account_format_mode="name",
    account_format="",
    password_format="",
):
    proxy_path = proxy_file or getattr(mod, "PROXY_FILE", "")
    proxy_pool = mod.parse_proxy_pool(proxy_path) if proxy_path else []
    selected_pool = []
    if proxy_pool:
        select_proxy = getattr(mod, "select_proxy_for_account", None)
        if callable(select_proxy):
            selected_pool = select_proxy(proxy_pool)
        else:
            selected_pool = [random.choice(proxy_pool)]
        masked = selected_pool[0].split(":")
        masked_proxy = f"{masked[0]}:{masked[1]}:{masked[2]}:***" if len(masked) == 4 else "***"
        log(f"{getattr(mod, 'ENGINE_NAME', 'ruoyi')} attempt #{idx} proxy -> {masked_proxy}")
    opts = SimpleNamespace(
        headless=headless,
        no_verify=False,
        timeout=timeout,
        max_press=max_press,
        confirm_before_register=confirm_before_register,
        proxy_file=proxy_path,
        email_suffixes=email_suffixes,
        account_format_mode=account_format_mode,
        account_format=account_format,
        password_format=password_format,
    )
    email, password = mod.register_outlook(opts, selected_pool, idx)
    return email, password, []


async def one_attempt(engine, mod, proxy_str, idx, args):
    if engine in ("ruoyi", "camonfox"):
        return await asyncio.to_thread(
            _one_attempt_ruoyi,
            mod,
            args.proxy_file,
            idx,
            args.timeout,
            args.max_press,
            args.confirm_before_register,
            args.headless,
            getattr(args, "email_suffixes", "") or "",
            getattr(args, "account_format_mode", "name") or "name",
            getattr(args, "account_format", "") or "",
            getattr(args, "password_format", "") or "",
        )
    return await one_attempt_standalone(mod, proxy_str, idx)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--engine", choices=["standalone", "ruoyi", "camonfox"],
                    default=os.environ.get("OUTLOOK_REG_ENGINE", "ruoyi"),
                    help="Outlook 自注册后端；ruoyi=Firefox BiDi，camonfox=原生 Camoufox，standalone=BitBrowser/Playwright")
    ap.add_argument("--count", type=int, default=0,
                    help="run this many attempts then exit (0 = loop forever)")
    ap.add_argument("--target-pool", type=int, default=0,
                    help="stop registering once pool dir has this many records "
                         "(0 = no cap; producer always runs)")
    ap.add_argument("--max-press", default="3",
                    help="OUTLOOK_REG_MAX_PRESS — captcha press-and-hold cap")
    ap.add_argument("--confirm-before-register", action="store_true",
                    help="auto-click confirmation on the signup page before filling")
    ap.add_argument("--headless", action="store_true",
                    help="仅 ruoyi 后端：以无头模式启动 Firefox")
    ap.add_argument("--timeout", type=int, default=180,
                    help="hard cap per attempt (seconds)")
    ap.add_argument("--proxy-file", default=os.environ.get("OUTLOOK_PROXY_FILE", "proxies_outlook.txt"),
                    help="代理池文件(每行 user:pass@host:port)；standalone/BitBrowser 每次随机取一个")
    ap.add_argument("--email-suffixes",
                    default=os.environ.get("OUTLOOK_ACCOUNT_SUFFIXES") or os.environ.get("OUTLOOK_EMAIL_SUFFIXES") or "outlook.com",
                    help="邮箱后缀池，逗号/空格分隔，如 outlook.com,hotmail.com")
    ap.add_argument("--account-format-mode",
                    default=(os.environ.get("OUTLOOK_ACCOUNT_FORMAT_MODE")
                             or ("custom" if os.environ.get("OUTLOOK_ACCOUNT_FORMAT") else "name")),
                    choices=["random", "name", "name_digits", "custom"],
                    help="账号格式预设：random/name/name_digits/custom")
    ap.add_argument("--account-format", default=os.environ.get("OUTLOOK_ACCOUNT_FORMAT", ""),
                    help="指定格式模板，如 {first}.{last}{digits:3}")
    ap.add_argument("--password-format", default=os.environ.get("OUTLOOK_PASSWORD_FORMAT", ""),
                    help="密码模板，如 Aa1!{rand:12}")
    ap.add_argument("--sleep", type=int, default=5,
                    help="seconds between attempts (after fail or success)")
    ap.add_argument("--sleep-when-full", type=int, default=60,
                    help="seconds to sleep when pool is at target")
    args = ap.parse_args()

    os.environ.setdefault("OUTLOOK_REG_MAX_PRESS", args.max_press)
    if args.confirm_before_register:
        os.environ["OUTLOOK_CONFIRM_BEFORE_REGISTER"] = "1"
    if args.proxy_file:
        os.environ["OUTLOOK_PROXY_FILE"] = args.proxy_file
    if args.email_suffixes:
        os.environ["OUTLOOK_ACCOUNT_SUFFIXES"] = args.email_suffixes
    os.environ["OUTLOOK_ACCOUNT_FORMAT_MODE"] = args.account_format_mode
    if args.account_format:
        os.environ["OUTLOOK_ACCOUNT_FORMAT"] = args.account_format
    if args.password_format:
        os.environ["OUTLOOK_PASSWORD_FORMAT"] = args.password_format
    if sys.platform == "win32":
        try:
            asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
        except Exception:
            pass

    clear_inherited_local_proxy_env()
    mod = load_engine(args.engine)
    proxy_pool = []
    if args.engine == "standalone":
        proxy_pool = load_proxy_pool(args.proxy_file, mod=mod)
        if proxy_pool:
            log(f"standalone BitBrowser random proxy pool: {len(proxy_pool)} from {args.proxy_file}")
        else:
            log("standalone proxy pool empty — BitBrowser will run noproxy", "WARN")
    elif args.proxy_file:
        log(f"ruoyi proxy file: {args.proxy_file}")

    log(f"pool dir: {POOL_DIR}")
    os.makedirs(POOL_DIR, exist_ok=True)
    log(f"current pool size: {count_pool()}")

    n = 0
    succ = 0
    failed = 0
    while True:
        n += 1
        if args.count and n > args.count:
            log(f"reached --count {args.count}, exit (success={succ}, fail={failed})")
            break
        ps = count_pool()
        if args.target_pool and ps >= args.target_pool:
            log(f"pool at target ({ps}/{args.target_pool}) — sleep {args.sleep_when_full}s")
            time.sleep(args.sleep_when_full)
            continue
        selected_proxy = random.choice(proxy_pool) if args.engine == "standalone" and proxy_pool else None
        if args.engine == "standalone":
            log(f"attempt #{n} BitBrowser proxy -> {mask_proxy(selected_proxy)}")
        log(f"=== attempt #{n}  (pool={ps}, succ={succ}, fail={failed}) ===")
        t0 = time.time()
        email = password = None
        cookies = []
        try:
            email, password, cookies = asyncio.run(
                asyncio.wait_for(one_attempt(args.engine, mod, selected_proxy, n, args), timeout=args.timeout)
            )
        except Exception as e:
            log(f"attempt raised {type(e).__name__}: {str(e)[:200]}", "WARN")
        elapsed = time.time() - t0
        if email and password:
            graph = extract_graph_for_account(email, password, proxy_str=selected_proxy)
            if not graph or not graph.get("refresh_token"):
                failed += 1
                log(f"registered but graph RT missing; not saved: {email}", "WARN")
                time.sleep(args.sleep)
                continue
            fname = write_record({
                "email": email,
                "password": password,
                "refresh_token": graph["refresh_token"],
                "client_id": graph.get("client_id") or "",
                "graph": graph,
                "outlook_cookies": cookies,
                "source": "self-loop",
                "ts": datetime.now().isoformat(),
            })
            globals()["_CURRENT_GRAPH_ACCOUNT"] = graph
            append_to_emails_pool(email, password)   # 桥接进账号注册池
            succ += 1
            log(f"OK in {elapsed:.1f}s: {email} -> {fname} (pool now {count_pool()})", "OK")
        else:
            failed += 1
            log(f"FAIL in {elapsed:.1f}s (success rate {succ}/{n} = {100*succ/n:.0f}%)", "WARN")
        time.sleep(args.sleep)


if __name__ == "__main__":
    main()
