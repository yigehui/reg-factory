# -*- coding: utf-8 -*-
"""
Extract Microsoft Graph API refresh tokens from Outlook accounts.
Uses pure requests to simulate OAuth2 authorization code flow (no browser needed).

Output format: email----password----refresh_token----client_id

Usage:
  python extract_graph_tokens.py outlook_accounts/accounts_20260413_043056.txt
  python extract_graph_tokens.py --email user@outlook.com --password pass123
"""

import argparse
import json
import os
import re
import sys
import urllib.parse
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stdin.reconfigure(encoding="utf-8")

import requests


def _load_dotenv_if_present():
    """CLI 直接跑时也吃项目根 .env（不覆盖已有环境变量）。"""
    try:
        root = os.path.dirname(os.path.abspath(__file__))
        path = os.path.join(root, ".env")
        if not os.path.isfile(path):
            return
        with open(path, encoding="utf-8") as f:
            for line in f:
                s = line.strip()
                if not s or s.startswith("#") or "=" not in s:
                    continue
                k, _, v = s.partition("=")
                k = k.strip()
                v = v.strip().strip('"').strip("'")
                if k and k not in os.environ:
                    os.environ[k] = v
    except Exception:
        pass


_load_dotenv_if_present()

# Thunderbird client — public, supports personal accounts.
# 用 Graph Mail.Read 资源域：下游 common/mailbox.get_code_by_token 走 Graph REST
# (/me/mailFolders/.../messages) 取码，必须拿 graph.microsoft.com 资源的 refresh_token；
# 之前用 outlook.office.com(IMAP) 资源的 token 无法换 Graph token，取码必失败。
CLIENT_ID = "9e5f94bc-e8a4-4e73-b8be-63364c29d753"
REDIRECT_URI = "http://localhost"
SCOPE = "offline_access https://graph.microsoft.com/Mail.Read"
OUTPUT_DIR = "outlook_accounts"
LOG_LEVELS = {"DEBUG": 10, "INFO": 20, "WARN": 30, "ERR": 40}


def _normalize_log_level(value, default="INFO"):
    raw = str(value or default).strip().upper()
    aliases = {"TRACE": "DEBUG", "WARNING": "WARN", "ERROR": "ERR"}
    raw = aliases.get(raw, raw)
    return raw if raw in LOG_LEVELS else default


def _log_level_value(value):
    return LOG_LEVELS.get(_normalize_log_level(value), LOG_LEVELS["INFO"])


def _should_demote_graph_log(msg, level):
    normalized = _normalize_log_level(level)
    if normalized != "INFO":
        return False
    low = str(msg or "").lower()
    debug_patterns = (
        "proxy=",
        "fetching auth page",
        "auth page ",
        "submitting credentials",
        "submit credentials ",
        "auto-submit intermediate",
        "got auth code!",
        "accepting consent/update",
        "skipping proofs/add",
        "submitting consent",
        "exchanging code for tokens",
        "token exchange ",
    )
    return any(pat in low for pat in debug_patterns)


def _graph_log(tag, msg, level="INFO"):
    normalized = _normalize_log_level(level)
    effective = "DEBUG" if _should_demote_graph_log(msg, normalized) else normalized
    current = _normalize_log_level(os.environ.get("OUTLOOK_LOG_LEVEL", "INFO"))
    if _log_level_value(effective) < _log_level_value(current):
        return
    print(f"[{datetime.now().strftime('%H:%M:%S')}] [{normalized}] {tag} [graph] {msg}")


def _mask_proxy_url(proxy_url):
    if not proxy_url:
        return "direct"
    parsed = urllib.parse.urlsplit(str(proxy_url))
    host = parsed.hostname or ""
    port = parsed.port or ""
    user = urllib.parse.unquote(parsed.username or "")
    auth = f"{user[:6]}...@" if user else ""
    scheme = parsed.scheme or "http"
    return f"{scheme}://{auth}{host}:{port}" if port else f"{scheme}://{auth}{host}"


def _save_graph_debug_html(email, idx, reason, html):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    safe = str(email or f"idx_{idx}").replace("@", "_").replace("/", "_")
    ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
    path = os.path.join(OUTPUT_DIR, f"graph_debug_{idx}_{reason}_{safe}_{ts}.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html or "")
    return path


def _extract_login_error(html):
    text = str(html or "")
    code = ""
    err = ""
    for pattern in (
        r'"sErrorCode":"([^"]+)"',
        r"<!-- HR=([A-F0-9]+) -->",
    ):
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            code = match.group(1)
            break
    err_match = re.search(r'"sErrTxt":"([^"]+)"', text)
    if err_match:
        err = err_match.group(1)
    err = re.sub(r"<[^>]+>", " ", err)
    err = err.replace("\\u0027", "'").replace("\\/", "/").replace('\\"', '"')
    err = re.sub(r"\s+", " ", err).strip()
    return code, err


def _probe_proxy_exit_ip(proxies):
    if not proxies:
        return ""
    session = requests.Session()
    session.trust_env = False
    session.proxies.update(proxies)
    try:
        resp = session.get("https://api.ipify.org?format=json", timeout=20)
        resp.raise_for_status()
        data = resp.json()
        return str(data.get("ip") or "").strip()
    except Exception:
        return ""


def _extract_flow_token(text):
    raw = str(text or "")
    patterns = (
        r'sFTTag.*?value=\\+"([^"\\]+)',
        r'sFTTag.*?value=["\']([^"\']+)["\']',
        r'name=\\"PPFT\\"[^>]*value=\\"([^"\\]+)\\"',
        r'"sFT"\s*:\s*"([^"]+)"',
        r'name="PPFT"[^>]*value="([^"]+)"',
    )
    for pattern in patterns:
        try:
            match = re.search(pattern, raw, re.DOTALL)
        except re.error as exc:
            raise RuntimeError(f"bad flow-token regex {pattern!r}: {exc}") from exc
        if match:
            return match.group(1)
    return ""


def get_graph_token(email, password, idx=0, proxies=None):
    """Get refresh_token via pure HTTP OAuth flow (no browser)."""
    tag = f"[#{idx}]"
    session = requests.Session()
    session.trust_env = False
    if proxies:
        session.proxies.update(proxies)
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36",
    })

    try:
        proxy_url = (proxies or {}).get("https") or (proxies or {}).get("http") or ""
        proxy_exit_ip = _probe_proxy_exit_ip(proxies)
        _graph_log(tag, f"proxy={_mask_proxy_url(proxy_url)} trust_env=False exit_ip={proxy_exit_ip or 'unknown'}")

        auth_url = (
            f"https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize"
            f"?client_id={CLIENT_ID}"
            f"&response_type=code"
            f"&redirect_uri={urllib.parse.quote(REDIRECT_URI, safe='')}"
            f"&scope={urllib.parse.quote(SCOPE)}"
            f"&response_mode=query"
        )
        _graph_log(tag, f"{email} - fetching auth page...")
        t_auth = datetime.now()
        resp = session.get(auth_url, timeout=30, allow_redirects=True)
        _graph_log(tag, f"auth page {(datetime.now() - t_auth).total_seconds():.2f}s status={resp.status_code} url={resp.url[:120]}")

        text = resp.text
        flow_token = _extract_flow_token(text)

        post_url = ""
        urlpost_match = re.search(r'"urlPost"\s*:\s*"([^"]+)"', text)
        if urlpost_match:
            post_url = urlpost_match.group(1).replace("\u0026", "&")

        ctx = ""
        sctx_match = re.search(r'"sCtx"\s*:\s*"([^"]+)"', text)
        if sctx_match:
            ctx = sctx_match.group(1)

        if not flow_token:
            debug_path = _save_graph_debug_html(email, idx, "no_flow_token", text)
            _graph_log(tag, f"FAIL: no flow token found debug={debug_path}", "WARN")
            return None

        if not post_url:
            post_url = "https://login.live.com/ppsecure/post.srf"

        _graph_log(tag, "submitting credentials...")
        login_data = {
            "login": email,
            "loginfmt": email,
            "passwd": password,
            "PPFT": flow_token,
            "ctx": ctx,
            "type": "11",
            "LoginOptions": "3",
            "i13": "0",
            "CookieDisclosure": "0",
            "IsFidoSupported": "0",
            "isSignupPost": "0",
            "i19": "16393",
        }

        t_submit = datetime.now()
        resp2 = session.post(post_url, data=login_data, timeout=30, allow_redirects=True)
        _graph_log(tag, f"submit credentials {(datetime.now() - t_submit).total_seconds():.2f}s status={resp2.status_code} url={getattr(resp2, 'url', '')[:120]}")

        for _ in range(5):
            _html = resp2.text or ''
            if ('DoSubmit' in _html or ('fmHF' in _html and 'onload' in _html)) and 'action=' in _html:
                _m = re.search(r'action="([^"]+)"', _html)
                if _m:
                    _fa = _m.group(1).replace('&amp;', '&')
                    _hid = re.findall(r'<input[^>]*name="([^"]*)"[^>]*value="([^"]*)"', _html)
                    _fd = {n: v for n, v in _hid}
                    resp2 = session.post(_fa, data=_fd, timeout=30, allow_redirects=True)
                    _graph_log(tag, f"auto-submit intermediate -> {getattr(resp2, 'url', '')[:120]}", "DEBUG")
                    continue
            break

        auth_code = None
        for _step in range(15):
            while resp2.status_code in (301, 302, 303, 307):
                loc = resp2.headers.get("Location", "")
                if "localhost" in loc and "code=" in loc:
                    resp2 = type('R', (), {'url': loc, 'text': '', 'status_code': 200})()
                    break
                if "localhost" in loc and "error" in loc:
                    resp2 = type('R', (), {'url': loc, 'text': '', 'status_code': 200})()
                    break
                resp2 = session.get(loc, timeout=30, allow_redirects=False)

            url = resp2.url
            text = resp2.text if hasattr(resp2, 'text') and resp2.text else ''

            if "localhost" in url and "code=" in url:
                parsed = urllib.parse.urlparse(url)
                params = urllib.parse.parse_qs(parsed.query)
                auth_code = params.get("code", [None])[0]
                if auth_code:
                    _graph_log(tag, "got auth code!")
                    break

            if "localhost" in url and "error" in url:
                parsed = urllib.parse.parse_qs(urllib.parse.urlparse(url).query)
                err = parsed.get("error_description", parsed.get("error", ["?"]))[0]
                _graph_log(tag, f"OAuth error: {err[:100]}", "WARN")
                return None

            if "Consent/Update" in url or "Consent/update" in url:
                m_sd = re.search(r'ServerData\s*=\s*(\{.*?\});', text, re.DOTALL)
                if m_sd:
                    sd = json.loads(m_sd.group(1))
                    form_data_consent = {
                        'ucaction': 'Yes',
                        'client_id': sd.get('sClientId', ''),
                        'scope': sd.get('sRawInputScopes', ''),
                        'cscope': sd.get('sRawInputGrantedScopes', ''),
                        'canary': sd.get('sCanary', ''),
                    }
                    _graph_log(tag, "accepting Consent/Update...")
                    resp2 = session.post(url, data=form_data_consent, timeout=30, allow_redirects=False)
                    continue
                debug_path = _save_graph_debug_html(email, idx, "consent_update", text)
                _graph_log(tag, f"FAIL: Consent/Update with no ServerData debug={debug_path}", "WARN")
                return None

            if "proofs/Add" in url or "proofs/add" in url:
                form_match2 = re.search(r'<form[^>]*action="([^"]+)"[^>]*>(.*?)</form>', text, re.DOTALL | re.IGNORECASE)
                if form_match2:
                    form_action2 = form_match2.group(1).replace("&amp;", "&")
                    form_body2 = form_match2.group(2)
                    hidden2 = re.findall(r'<input[^>]*name="([^"]*)"[^>]*value="([^"]*)"', form_body2)
                    form_data2 = {n: v for n, v in hidden2}
                    form_data2["action"] = "Skip"
                    if not form_action2.startswith("http"):
                        base2 = urllib.parse.urlparse(url)
                        form_action2 = f"{base2.scheme}://{base2.netloc}{form_action2}"
                    _graph_log(tag, f"skipping proofs/Add (action=Skip) -> {form_action2[:80]}...")
                    resp2 = session.post(form_action2, data=form_data2, timeout=30, allow_redirects=False)
                    continue
                debug_path = _save_graph_debug_html(email, idx, "proofs_add", text)
                _graph_log(tag, f"FAIL: proofs/Add with no form debug={debug_path}", "WARN")
                return None

            form_match = re.search(r'<form[^>]*action="([^"]+)"[^>]*>(.*?)</form>', text, re.DOTALL | re.IGNORECASE)
            if form_match:
                form_action = form_match.group(1).replace("&amp;", "&")
                form_body = form_match.group(2)
                hidden = re.findall(r'<input[^>]*name="([^"]*)"[^>]*value="([^"]*)"', form_body)
                form_data = {name: val for name, val in hidden}

                if "consent" in form_action.lower() or "consent" in url.lower():
                    form_data["ucaccept"] = "Yes"
                    _graph_log(tag, "submitting consent...", "DEBUG")

                if not form_action.startswith("http"):
                    base = urllib.parse.urlparse(url)
                    form_action = f"{base.scheme}://{base.netloc}{form_action}"

                resp2 = session.post(form_action, data=form_data, timeout=30, allow_redirects=False)
                while resp2.status_code in (301, 302, 303, 307):
                    loc = resp2.headers.get("Location", "")
                    if "localhost" in loc:
                        resp2 = type('R', (), {'url': loc, 'text': '', 'status_code': 200})()
                        break
                    elif loc:
                        resp2 = session.get(loc, timeout=30, allow_redirects=False)
                    else:
                        break
                continue

            debug_path = _save_graph_debug_html(email, idx, "stuck", text)
            err_code, err_text = _extract_login_error(text)
            plain = re.sub(r"\s+", " ", (text or "")).strip()
            # 纯文本错误页（如 "Bad user credential or too many signin attempts..."）
            if not err_text and plain and len(plain) < 300 and "<html" not in plain.lower():
                err_text = plain
            detail = f" code={err_code}" if err_code else ""
            if err_text:
                detail += f" err={err_text[:200]!r}"
            low = (err_text or plain or "").lower()
            if "bad user credential" in low or "too many signin" in low:
                _graph_log(
                    tag,
                    f"FAIL: bad_cred_or_rate_limit at {url[:100]} "
                    f"(status={resp2.status_code}){detail} debug={debug_path}",
                    "WARN",
                )
            else:
                _graph_log(
                    tag,
                    f"FAIL: stuck at {url[:100]} (status={resp2.status_code}){detail} debug={debug_path}",
                    "WARN",
                )
            return None

        if not auth_code:
            _graph_log(tag, "FAIL: no auth code extracted", "WARN")
            return None

        _graph_log(tag, "exchanging code for tokens...")
        t_token = datetime.now()
        token_resp = session.post(
            "https://login.microsoftonline.com/consumers/oauth2/v2.0/token",
            data={
                "client_id": CLIENT_ID,
                "grant_type": "authorization_code",
                "code": auth_code,
                "redirect_uri": REDIRECT_URI,
                "scope": SCOPE,
            },
            timeout=30,
        )
        try:
            token_data = token_resp.json()
        except Exception:
            token_data = {"error": "non_json", "error_description": (token_resp.text or "")[:200]}

        if "access_token" in token_data:
            rt = token_data.get("refresh_token", "")
            _graph_log(
                tag,
                f"token exchange {(datetime.now() - t_token).total_seconds():.2f}s "
                f"refresh_token={'yes' if rt else 'no'}",
            )
            return {
                "email": email,
                "password": password,
                "refresh_token": rt,
                "client_id": CLIENT_ID,
            }
        else:
            err = token_data.get("error_description", token_data.get("error", "?"))
            _graph_log(tag, f"token error: {str(err)[:150]}", "WARN")
            return None

    except Exception as e:
        _graph_log(tag, f"error: {type(e).__name__}: {e}", "ERR")
        return None


def main():
    parser = argparse.ArgumentParser(description="Extract Graph API tokens")
    parser.add_argument("accounts_file", nargs="?")
    parser.add_argument("--email", "-e", type=str)
    parser.add_argument("--password", "-p", type=str)
    parser.add_argument("--concurrency", "-c", type=int, default=5)
    parser.add_argument("--log-level", default=os.environ.get("OUTLOOK_LOG_LEVEL", "INFO"),
                        choices=["DEBUG", "INFO", "WARN", "ERR"])
    args = parser.parse_args()
    os.environ["OUTLOOK_LOG_LEVEL"] = _normalize_log_level(args.log_level)

    accounts = []
    if args.email and args.password:
        accounts.append((args.email, args.password))
    elif args.accounts_file:
        with open(args.accounts_file, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#"):
                    parts = line.split("----")
                    if len(parts) >= 2:
                        accounts.append((parts[0], parts[1]))
    else:
        # Auto-scan: load all unlocked accounts from unlock_results/, skip already extracted
        unlock_dir = "unlock_results"

        # Collect emails that already have tokens
        token_emails = set()
        if os.path.isdir(OUTPUT_DIR):
            for tf in sorted(os.listdir(OUTPUT_DIR)):
                if tf.startswith("graph_tokens_") and tf.endswith(".txt"):
                    with open(os.path.join(OUTPUT_DIR, tf), "r", encoding="utf-8") as tf_f:
                        for line in tf_f:
                            parts = line.strip().split("----")
                            if parts and parts[0]:
                                token_emails.add(parts[0].lower())

        # Collect all unlocked accounts, deduplicate by email
        seen_emails: set = set()
        if os.path.isdir(unlock_dir):
            for uf in sorted(os.listdir(unlock_dir)):
                if uf.startswith("unlocked_clean_") and uf.endswith(".txt"):
                    with open(os.path.join(unlock_dir, uf), "r", encoding="utf-8") as uf_f:
                        for line in uf_f:
                            line = line.strip()
                            if not line or line.startswith("#"):
                                continue
                            parts = line.split("----")
                            if len(parts) >= 2:
                                email_lc = parts[0].lower()
                                if email_lc not in seen_emails and email_lc not in token_emails:
                                    accounts.append((parts[0], parts[1]))
                                    seen_emails.add(email_lc)

        if token_emails:
            print(f"  Skipping {len(token_emails)} already-extracted accounts")
        print(f"  Auto-loaded {len(accounts)} new accounts from {unlock_dir}/")

    if not accounts:
        print("  No accounts to process.")
        return

    print("=" * 60)
    print(f"  Graph API Token Extraction (pure HTTP)")
    print(f"  accounts={len(accounts)}  concurrency={args.concurrency}")
    print(f"  client_id={CLIENT_ID}")
    print("=" * 60)

    results = []
    with ThreadPoolExecutor(max_workers=args.concurrency) as pool:
        futures = {pool.submit(get_graph_token, e, p, i + 1): (e, p) for i, (e, p) in enumerate(accounts)}
        for future in as_completed(futures):
            result = future.result()
            if result:
                results.append(result)

    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print(f"\n{'=' * 60}")
    print(f"  RESULTS: {len(results)}/{len(accounts)} tokens extracted")
    print("=" * 60)

    if results:
        out_file = os.path.join(OUTPUT_DIR, f"graph_tokens_{ts}.txt")
        with open(out_file, "w", encoding="utf-8") as f:
            for r in results:
                f.write(f"{r['email']}----{r['password']}----{r.get('refresh_token','')}----{CLIENT_ID}\n")
        print(f"  Saved to: {out_file}")

        for r in results:
            rt = r.get("refresh_token", "")
            print(f"  [OK] {r['email']}  rt={rt[:50]}...")

    print("=" * 60)


if __name__ == "__main__":
    main()
