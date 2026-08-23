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
import threading
import urllib.parse
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    # pytest 下 stdin 是 DontReadFromInput(无 reconfigure),兼容之
    try:
        sys.stdin.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

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

# 失败网页保存开关:默认不保存(全量跑几千个 graph_debug_*.html 纯占盘)。
# 开启方式:env OUTLOOK_SAVE_DEBUG_HTML=1/true,或调用方 set_save_debug_html(True)。
SAVE_DEBUG_HTML = str(os.environ.get("OUTLOOK_SAVE_DEBUG_HTML", "")).strip().lower() in {"1", "true", "yes", "on"}

# 每线程失败原因槽:get_graph_token 失败时记下 abuse/noexist/pwd_incorrect/...,
# 调用方(如 auth_nograph_to_all)不落盘也能分类。threading.local 保证并发不串号。
_tls = threading.local()


def set_save_debug_html(enabled):
    """运行时开关:授权失败是否保存网页到 outlook_accounts/。"""
    global SAVE_DEBUG_HTML
    SAVE_DEBUG_HTML = bool(enabled)


def _set_thread_fail_reason(reason):
    _tls.fail_reason = str(reason or "")


def get_thread_fail_reason():
    """读本线程最近一次 get_graph_token 的失败原因;没记过返回 ''(不抛)。"""
    return str(getattr(_tls, "fail_reason", "") or "")


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
    if not SAVE_DEBUG_HTML:
        return ""  # 开关关:不落盘不建目录,日志侧 debug= 为空即"未保存"
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


def _parse_proof_add_form(text, url):
    """从 proofs/Add 真表单页解析出 form action + hidden 字段(含 canary)。
    返回 (form_action, form_data) 或 (None, None)。"""
    form_match = re.search(r'<form[^>]*action="([^"]+)"[^>]*>(.*?)</form>', text, re.DOTALL | re.IGNORECASE)
    if not form_match:
        return None, None
    form_action = form_match.group(1).replace("&amp;", "&")
    form_body = form_match.group(2)
    hidden = re.findall(r'<input[^>]*name="([^"]*)"[^>]*value="([^"]*)"', form_body)
    form_data = {n: v for n, v in hidden}
    if not form_action.startswith("http"):
        base = urllib.parse.urlparse(url)
        form_action = f"{base.scheme}://{base.netloc}{form_action}"
    return form_action, form_data


def _parse_proof_verify_form(text, url):
    """从 proofs/Verify 页解析 frmVerifyProof 的 action(含 epid)+ hidden(canary/action=VerifyProof)。
    返回 (form_action, form_data) 或 (None, None)。epid 在 action URL 里,一起带回去。"""
    # 只取 frmVerifyProof 这个 form(页面还有个 frmSubmitSLT 干扰,且 slt 为空不提交)
    form_match = re.search(
        r'<form[^>]*(?:id|name)="frmVerifyProof"[^>]*action="([^"]+)"[^>]*>(.*?)</form>',
        text, re.DOTALL | re.IGNORECASE)
    if not form_match:
        # 退而求其次:含 Verify 的 form
        form_match = re.search(r'<form[^>]*action="([^"]*proofs/Verify[^"]*)"[^>]*>(.*?)</form>',
                               text, re.DOTALL | re.IGNORECASE)
        if not form_match:
            return None, None
    form_action = form_match.group(1).replace("&amp;", "&")
    form_body = form_match.group(2)
    hidden = re.findall(r'<input[^>]*name="([^"]*)"[^>]*value="([^"]*)"', form_body)
    form_data = {n: v for n, v in hidden}
    if not form_action.startswith("http"):
        base = urllib.parse.urlparse(url)
        form_action = f"{base.scheme}://{base.netloc}{form_action}"
    return form_action, form_data


def bind_proof_in_session(session, html, url, cf_address, cf_jwt=None, use_admin=False, idx=0,
                          cm_module=None, max_wait=180, poll=3):
    """在一个已登录的 requests.Session 里,把 proofs/Add 真绑成 cf 辅助邮箱。
    步骤:AddProof(填 EmailAddress)→ 微软发码到 cf → 收码 → VerifyProof(填 iOttText)。
    成功后返回下一个响应 resp(通常落在 Consent 或已登录页),调用方继续跟 oauth。
    失败返回 None。

    cm_module: common.cloudflare_mail 模块(避免循环 import,由调用方传入)。
    """
    tag = f"[#{idx}]"
    if cm_module is None:
        try:
            from common import cloudflare_mail as cm_module
        except Exception as e:
            _graph_log(tag, f"bind_proof: cloudflare_mail 不可用: {e}", "ERR")
            return None
    cm = cm_module

    # 1) proofs/Add:解析 form + 填 EmailAddress
    form_action, form_data = _parse_proof_add_form(html, url)
    if not form_action:
        debug_path = _save_graph_debug_html(cf_address or "bind", idx, "proofs_add_no_form", html)
        _graph_log(tag, f"bind: proofs/Add 无 form debug={debug_path}", "WARN")
        return None
    if not cf_address:
        _graph_log(tag, "bind: 缺 cf 辅助邮箱地址", "ERR")
        return None
    # AddProof 提交字段:canary(已有)+ action=AddProof + EmailAddress + iProofOptions=Email
    form_data["action"] = "AddProof"
    form_data["EmailAddress"] = cf_address
    form_data.setdefault("iProofOptions", "Email")
    # 提交前记 cf 收件箱基线 id(微软发码很快,基线必须在提交前取)
    base_last_id = 0
    try:
        if use_admin:
            raws = cm.fetch_admin_mails(cf_address, limit=20)
            base_last_id = max([m.get("id", 0) for m in raws] or [0])
        else:
            mails = cm.fetch_parsed_mails(cf_jwt, limit=20)
            base_last_id = max([m.get("id", 0) for m in mails] or [0])
    except Exception as e:
        _graph_log(tag, f"bind: 取 cf 基线失败(继续): {e}", "WARN")
    _graph_log(tag, f"bind: 提交 AddProof email={cf_address} base_id={base_last_id}")

    resp = session.post(form_action, data=form_data, timeout=30, allow_redirects=True)
    vurl = getattr(resp, "url", "") or ""
    vtext = resp.text or ""
    _graph_log(tag, f"bind: AddProof -> {resp.status_code} url={vurl[:80]}")

    # 2) 跟重定向到 proofs/Verify(可能要追一跳)
    for _ in range(5):
        if "proofs/verify" in (vurl or "").lower():
            break
        # 有时落在中间页(Consent 不应在这阶段;proofs/Add 重复说明没绑成功)
        if "consent" in (vurl or "").lower() or "localhost" in (vurl or ""):
            break
        fm = re.search(r'<form[^>]*action="([^"]+)"', vtext, re.IGNORECASE)
        if fm and ("DoSubmit" in vtext or "fmHF" in vtext):
            fa = fm.group(1).replace("&amp;", "&")
            if not fa.startswith("http"):
                base = urllib.parse.urlparse(vurl)
                fa = f"{base.scheme}://{base.netloc}{fa}"
            hid = re.findall(r'<input[^>]*name="([^"]*)"[^>]*value="([^"]*)"', vtext)
            resp = session.post(fa, data={n: v for n, v in hid}, timeout=30, allow_redirects=True)
            vurl = getattr(resp, "url", "") or ""; vtext = resp.text or ""
            continue
        break

    if "proofs/verify" not in (vurl or "").lower():
        debug_path = _save_graph_debug_html(cf_address or "bind", idx, "proofs_verify_not_reached", vtext)
        _graph_log(tag, f"bind: 未到 proofs/Verify (url={vurl[:80]}) debug={debug_path}", "WARN")
        return None

    # 3) 解析 Verify form(canary/epid 在 action URL)+ 收码
    vaction, vdata = _parse_proof_verify_form(vtext, vurl)
    if not vaction:
        debug_path = _save_graph_debug_html(cf_address or "bind", idx, "proofs_verify_no_form", vtext)
        _graph_log(tag, f"bind: proofs/Verify 无 form debug={debug_path}", "WARN")
        return None
    vdata["action"] = "VerifyProof"
    code = cm.wait_for_code(cf_jwt, received_after_id=base_last_id, max_wait=max_wait,
                            poll=poll, use_admin=use_admin, address=cf_address)
    if not code:
        # 兜底:基线后没新码,可能是重试同一 proof(微软限频不发新码),
        # 但收件箱里基线那封码对当前 pending proof 仍有效 —— 取最新一封匹配码试。
        code = _fetch_latest_code_fallback(cm, cf_address, cf_jwt, use_admin)
        if not code:
            _graph_log(tag, "bind: cf 取码超时(且无兜底可用码)", "WARN")
            return None
        _graph_log(tag, f"bind: 用兜底最新码 {code}(重试场景基线码仍有效)")
    else:
        _graph_log(tag, f"bind: 取到验证码 {code}")
    vdata["iOttText"] = code
    # allow_redirects=False:绑完后重定向链会一路跟到 http://localhost/?code=...
    # 若 allow_redirects=True,requests 会真去连本地 80 端口 → ConnectionError。
    # 让主循环手动 follow Location(它有 localhost code 拦截),避免本地连不上炸掉。
    resp = session.post(vaction, data=vdata, timeout=30, allow_redirects=False)
    _graph_log(tag, f"bind: VerifyProof -> {resp.status_code} url={getattr(resp,'url','')[:80]}")
    return resp


def _fetch_latest_code_fallback(cm, address, jwt, use_admin):
    """wait_for_code 超时兜底:直接取收件箱里最新一封匹配的微软安全码(忽略基线)。
    用于重试同一 pending proof 时微软限频不发新码、但旧码仍有效的场景。"""
    try:
        if use_admin:
            raws = cm.fetch_admin_mails(address, limit=10)
            mails = [cm.parse_admin_mail(m) for m in raws]
        else:
            mails = cm.fetch_parsed_mails(jwt, limit=10)
        for m in mails:
            code = cm.extract_code(m.get("text") or m.get("html") or "") or cm.extract_code(m.get("subject"))
            if code:
                return code
    except Exception:
        pass
    return None


def get_graph_token(email, password, idx=0, proxies=None, bind_secondary=None):
    """Get refresh_token via pure HTTP OAuth flow (no browser).

    bind_secondary: 可选 dict,传入则 proofs/Add 真绑辅助邮箱(而非 Skip):
      {"cf_address": "...", "cf_jwt": "..."/None, "use_admin": bool, "cm": cloudflare_mail模块}
    不传则保持原 Skip 行为(向后兼容,不影响注册链路)。

    失败时返回 None,并把原因记到线程槽(get_thread_fail_reason() 可读:
    abuse/noexist/pwd_incorrect/rate_limited/... ),供调用方不落盘也能分类。"""

    """Get refresh_token via pure HTTP OAuth flow (no browser)."""
    tag = f"[#{idx}]"
    _set_thread_fail_reason("")  # 清旧值,失败前不残留上一号的原因
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
            _set_thread_fail_reason("no_flow_token")
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
        # allow_redirects=False:已绑邮箱的号登录后会一路 302 到 http://localhost/?code=...,
        # 若 allow_redirects=True,requests 会真去连本地 80 → ConnectionError。手动 follow 让
        # 主循环的 redirect follower 拦截 localhost code。
        resp2 = session.post(post_url, data=login_data, timeout=30, allow_redirects=False)
        _graph_log(tag, f"submit credentials {(datetime.now() - t_submit).total_seconds():.2f}s status={resp2.status_code} url={getattr(resp2, 'url', '')[:120]}")

        for _ in range(5):
            # 先 follow 302 到下个页面(非 localhost),再判 DoSubmit 跳板
            while resp2.status_code in (301, 302, 303, 307):
                loc = resp2.headers.get("Location", "")
                if "code=" in loc or "error=" in loc:
                    break  # localhost code/error,留给主循环处理
                if not loc:
                    break
                if not loc.startswith("http"):
                    base = urllib.parse.urlparse(getattr(resp2, "url", "") or post_url)
                    loc = f"{base.scheme}://{base.netloc}{loc if loc.startswith('/') else '/'+loc}"
                try:
                    resp2 = session.get(loc, timeout=30, allow_redirects=False)
                except Exception:
                    if "code=" in loc:
                        break
                    raise
            if "code=" in (resp2.headers.get("Location", "") if resp2.status_code in (301, 302, 303, 307) else ""):
                break  # 已到 localhost code 重定向,跳出跳板循环交给主循环
            _html = resp2.text or ''
            if ('DoSubmit' in _html or ('fmHF' in _html and 'onload' in _html)) and 'action=' in _html:
                _m = re.search(r'action="([^"]+)"', _html)
                if _m:
                    _fa = _m.group(1).replace('&amp;', '&')
                    _hid = re.findall(r'<input[^>]*name="([^"]*)"[^>]*value="([^"]*)"', _html)
                    _fd = {n: v for n, v in _hid}
                    resp2 = session.post(_fa, data=_fd, timeout=30, allow_redirects=False)
                    _graph_log(tag, f"auto-submit intermediate -> {getattr(resp2, 'url', '')[:120]}", "DEBUG")
                    continue
            break

        auth_code = None

        def _is_code_redirect(loc):
            """Location 是 localhost?code= 重定向(授权成功)。
            可能绝对(http://localhost/?code=)或相对(/?code=、?code=)。"""
            return "code=" in loc and ("localhost" in loc or loc.lstrip().startswith("/?")
                    or loc.lstrip().startswith("?"))

        def _is_error_redirect(loc):
            return "error=" in loc and ("localhost" in loc or loc.lstrip().startswith("/?")
                    or loc.lstrip().startswith("?"))

        def _abs_localhost(loc):
            if not loc.startswith("http") and (loc.startswith("/?") or loc.startswith("?")):
                return f"http://localhost/{loc.lstrip('/')}"
            return loc

        for _step in range(15):
            while resp2.status_code in (301, 302, 303, 307):
                loc = resp2.headers.get("Location", "")
                _graph_log(tag, f"redirect status={resp2.status_code} loc={loc[:90]}", "DEBUG")
                if _is_code_redirect(loc):
                    resp2 = type('R', (), {'url': _abs_localhost(loc), 'text': '', 'status_code': 200})()
                    break
                if _is_error_redirect(loc):
                    resp2 = type('R', (), {'url': _abs_localhost(loc), 'text': '', 'status_code': 200})()
                    break
                try:
                    resp2 = session.get(loc, timeout=30, allow_redirects=False)
                except Exception:
                    # 兜底:连 localhost:80 被拒等——loc 带 code= 即视为授权成功,提取 code。
                    if "code=" in loc:
                        resp2 = type('R', (), {'url': _abs_localhost(loc), 'text': '', 'status_code': 200})()
                        break
                    raise

            url = resp2.url
            text = resp2.text if hasattr(resp2, 'text') and resp2.text else ''

            if "code=" in url and ("localhost" in url or url.startswith("http://localhost")):
                parsed = urllib.parse.urlparse(url)
                params = urllib.parse.parse_qs(parsed.query)
                auth_code = params.get("code", [None])[0]
                if auth_code:
                    _graph_log(tag, "got auth code!")
                    break

            if "error=" in url and ("localhost" in url or url.startswith("http://localhost")):
                parsed = urllib.parse.parse_qs(urllib.parse.urlparse(url).query)
                err = parsed.get("error_description", parsed.get("error", ["?"]))[0]
                _graph_log(tag, f"OAuth error: {err[:100]}", "WARN")
                _set_thread_fail_reason("oauth_error")
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
                _set_thread_fail_reason("consent_update")
                return None

            if "proofs/Add" in url or "proofs/add" in url:
                # 优先路径:真绑辅助邮箱(若调用方传了 bind_secondary)
                if bind_secondary:
                    bs = bind_secondary
                    bresp = bind_proof_in_session(
                        session, text, url,
                        cf_address=bs.get("cf_address"),
                        cf_jwt=bs.get("cf_jwt"),
                        use_admin=bs.get("use_admin", False),
                        idx=idx,
                        cm_module=bs.get("cm"),
                    )
                    if bresp is None:
                        _graph_log(tag, "FAIL: 辅助邮箱绑定失败", "WARN")
                        _set_thread_fail_reason("bind_fail")
                        return None
                    # 绑定成功后,继续跟 oauth(把 resp2 换成绑定后的响应,循环继续追 code/consent)
                    resp2 = bresp
                    continue
                # 兼容路径:不绑,只 Skip(原行为)
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
                _set_thread_fail_reason("proofs_add")
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
            url_low = (url or "").lower()
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
            # 记线程内失败原因(调用方不落盘也能分类;口径同 auth_nograph_to_all._classify_failure)
            if "/abuse" in url_low:
                _set_thread_fail_reason("abuse")
            elif "doesn't exist" in low or "does not exist" in low:
                _set_thread_fail_reason("noexist")
            elif "too many requests" in low or "too many signin" in low or "bad user credential" in low:
                _set_thread_fail_reason("rate_limited")
            elif str(err_code) == "80041012" or "password is incorrect" in low:
                _set_thread_fail_reason("pwd_incorrect")
            else:
                _set_thread_fail_reason("unknown")
            return None

        if not auth_code:
            _graph_log(tag, "FAIL: no auth code extracted", "WARN")
            _set_thread_fail_reason("no_auth_code")
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
            _set_thread_fail_reason("token_error")
            return None

    except Exception as e:
        _graph_log(tag, f"error: {type(e).__name__}: {e}", "ERR")
        _set_thread_fail_reason("exc")
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
