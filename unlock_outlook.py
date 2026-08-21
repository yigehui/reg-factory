# -*- coding: utf-8 -*-
"""
Outlook Account Batch Unlock Script  --  ruyipage Firefox edition

业务逻辑(状态机 / 登录 / PX 按压 / 解锁流程 / 并发 / 文件 IO)照搬原 unlock_outlook.py,
浏览器层从 BitBrowser + Playwright 换成 ruyipage Firefox BiDi,复用 register_outlook_ruoyi 的:
  - 浏览器启动栈(FirefoxOptions + set_per_tab_proxies + FirefoxPage)
  - 代理池(ConsumableProxyPool: file / http / aimili-list / aimili-random,per-tab socks5)
  - PX press-and-hold(_find_hold_context + _perform_hold_with_px_screenshots,ctx.actions 链)
  - 页面 helper(_safe_input / _click_next / _body_text / _maybe_skip_passkey / ...)

代理格式对外 user:pass@host:port(与 outlook_reg_loop --proxy-file 一致),
ruyi 内部归一化为 host:port:user:pwd 喂给 ruyipage set_per_tab_proxies(它只认这个)。

Usage:
  python unlock_outlook.py --input outlook_accounts/accounts_xxx.txt
  python unlock_outlook.py --input emails_locked.txt --concurrency 2
  python unlock_outlook.py --proxy-file proxies_outlook.txt --proxy-source aimili-list --aimili-url http://host:8787 --aimili-token XXX
  python unlock_outlook.py                         (auto-scan all accounts, skip unlocked)

Input file format (---- separated, one per line):
  email----password
  email----password----any_extra_fields...

Output (unlock_results/):
  unlocked_*.txt          successfully unlocked
  needs_phone_*.txt       requires SMS - cannot auto-unlock
  failed_*.txt            failed / timeout
"""

import argparse, asyncio, importlib.util, os, signal, sys, time
from datetime import datetime
from types import SimpleNamespace

# 顶部加载 .env 凭据(真实环境变量优先),保持仓库内无明文凭据
try:
    from config import EZCAPTCHA_API_KEY as _EZCAPTCHA_KEY, EZCAPTCHA_API_BASE as _EZCAPTCHA_BASE
except Exception:
    _EZCAPTCHA_KEY = os.environ.get("EZCAPTCHA_API_KEY", "")
    _EZCAPTCHA_BASE = os.environ.get("EZCAPTCHA_API_BASE", "https://api.ez-captcha.com")

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stdin.reconfigure(encoding="utf-8")

import requests

# ── 复用 ruyi 模块(register_outlook_ruoyi)─────────────────────────────
_RUOYI_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "register_outlook_ruoyi.py")
_spec = importlib.util.spec_from_file_location("_unlock_ruoyi", _RUOYI_PATH)
_ruoyi = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_ruoyi)

from ruyipage import FirefoxOptions, FirefoxPage

# ── Config ───────────────────────────────────────────────────────────
OUTPUT_DIR      = "unlock_results"
SCREENSHOT_DIR  = "screenshots_unlock"
UNLOCK_TIMEOUT  = 300   # seconds per account
EZCAPTCHA_KEY   = _EZCAPTCHA_KEY
EZCAPTCHA_BASE  = _EZCAPTCHA_BASE
LOGIN_URL       = "https://login.live.com/login.srf"
PX_APP_ID       = "PXzC5j78di"
DEFAULT_MAX_PRESS = 5
LOADING_WAIT_TIMEOUT = 60   # PX 挑战后微软 Loading 转圈页等待上限(秒),超时 give up
MAX_PROXY_RETRY = 3   # 代理无法访问微软时,换节点重开浏览器重试上限

# ── 从 ruyi 复用的符号 ────────────────────────────────────────────────
RUOYI_FIREFOX_PATH            = _ruoyi.RUOYI_FIREFOX_PATH
ConsumableProxyPool           = _ruoyi.ConsumableProxyPool
select_proxy_for_account      = _ruoyi.select_proxy_for_account
release_proxy_for_account     = _ruoyi.release_proxy_for_account
set_consumable_proxy_pool     = _ruoyi.set_consumable_proxy_pool
mask_ruoyi_proxy              = _ruoyi.mask_ruoyi_proxy
set_log_level                 = _ruoyi.set_log_level

_install_shutdown_handlers    = _ruoyi._install_shutdown_handlers
_track_browser_page           = _ruoyi._track_browser_page
_untrack_browser_page         = _ruoyi._untrack_browser_page
_quit_browser_page            = _ruoyi._quit_browser_page
_force_kill_ruoyi_firefox     = _ruoyi._force_kill_ruoyi_firefox
_pick_user_agent              = _ruoyi._pick_user_agent
_browser_model_name           = _ruoyi._browser_model_name
_ruoyi_profile_dir            = _ruoyi._ruoyi_profile_dir
_apply_ruoyi_browser_ua       = _ruoyi._apply_ruoyi_browser_ua
_apply_ruoyi_headless_options = _ruoyi._apply_ruoyi_headless_options
_log_current_ip               = _ruoyi._log_current_ip
_probe_proxy_before_browser   = _ruoyi._probe_proxy_before_browser

# PX 按压(直接复用 ruyi 成熟实现)
_find_hold_context            = _ruoyi._find_hold_context
_perform_hold_with_px_screenshots = _ruoyi._perform_hold_with_px_screenshots
_new_ruoyi_px_motion_profile  = _ruoyi._new_ruoyi_px_motion_profile
_wait_before_next_captcha_press = _ruoyi._wait_before_next_captcha_press
_maybe_skip_passkey           = _ruoyi._maybe_skip_passkey

# 页面操作 helper
_body_text                    = _ruoyi._body_text
_safe_input                   = _ruoyi._safe_input
_click_any                    = _ruoyi._click_any
_click_next                   = _ruoyi._click_next
_ele                          = _ruoyi._ele


# ── EZCaptcha PX API (fallback, 与原版一致)──────────────────────────
def solve_px(page_url, app_id=PX_APP_ID, max_wait=90):
    try:
        resp = requests.post(f"{EZCAPTCHA_BASE}/createTask", json={
            "clientKey": EZCAPTCHA_KEY,
            "task": {"type": "PerimeterX", "websiteURL": page_url, "websiteKey": app_id}
        }, timeout=30)
        d = resp.json()
        if d.get("errorId", 1) != 0:
            print(f"    [px] error: {d.get('errorDescription', d)}")
            return None
        tid = d["taskId"]
        print(f"    [px] task {tid}")
        start = time.time()
        while time.time() - start < max_wait:
            time.sleep(5)
            r2 = requests.post(f"{EZCAPTCHA_BASE}/getTaskResult",
                               json={"clientKey": EZCAPTCHA_KEY, "taskId": tid},
                               timeout=30).json()
            if r2.get("status") == "ready":
                sol = r2.get("solution", {})
                print(f"    [px] solved! keys={list(sol.keys())}")
                return sol
            if r2.get("status") == "failed":
                print("    [px] failed"); return None
        print("    [px] timeout"); return None
    except Exception as e:
        print(f"    [px] error: {e}"); return None


def inject_px_solution(page, sol):
    """把 EZCaptcha 返回的 PX cookie/token 注入页面(ruyipage 版)。"""
    try:
        for key in ['_pxCaptcha', '_px3', '_px2', '_pxhd', '_pxvid', '_pxde']:
            val = sol.get(key)
            if val:
                page.run_js_loaded(
                    f'document.cookie="{key}={val};domain=.live.com;path=/";'
                )
        tok = sol.get("token") or sol.get("uuid")
        if tok:
            page.run_js_loaded(
                'var h=document.querySelector(\'input[name="_pxCaptcha"]\');'
                f'if(h){{h.value="{tok}";}}'
            )
    except Exception as e:
        print(f"    inject px error: {e}")


# 解锁后过渡页(如 account.live.com/Abuse?id=389)的"继续"按钮选择器
# 参考 register_outlook_ruoyi 继续页处理:文案按钮优先,再 fallback submit
_CONTINUE_SELECTORS = [
    'text:Continue', 'text:OK', 'text:Next', 'text:Got it',
    'text:Accept', 'text:Agree', 'text:I agree', 'text:Yes',
    'text:继续', 'text:下一步', 'text:同意',
    'css:#idSIButton9', 'css:button[type="submit"]', 'css:input[type="submit"]',
]

# Something went wrong 等错误页的"重试"按钮选择器(大小写/中英文变体)
_RETRY_SELECTORS = [
    'text:Try again', 'text:Try Again', 'text:try again',
    'text:重试', 'text:再试一次', 'text:再试',
]


# ── Page state classifier (套原 classify 文案到 ruyipage)─────────────
def classify(page):
    url = page.url or ""
    # 遍历所有 context(主文档 + iframe)取 body 文本 + h1/h2 标题,合并匹配
    # 成功文案(Your account has been unblocked)可能在 iframe 里,主文档拿不到
    text_parts = []
    heading_parts = []
    for ctx in _ruoyi._all_contexts(page):
        try:
            text_parts.append(_ruoyi._context_text(ctx) or "")
        except Exception:
            pass
        try:
            heading_parts.append(ctx.run_js_loaded(
                "return Array.from(document.querySelectorAll('h1,h2')).map(e=>e.innerText||'').join('\\n');") or "")
        except Exception:
            pass
    t = ("\n".join(text_parts) + "\n" + "\n".join(heading_parts)).lower()
    u = (url or "").lower()
    # PX 人工挑战优先:Abuse?id=389 等页可能内嵌 PX,要先按压而非点继续/跳过
    if "let's prove you're human" in t or "press and hold" in t: return "px_challenge"
    # 解锁成功文案优先于 URL 判定(无论哪个页面/iframe 出现都算成功)
    if "account has been unblocked" in t:                  return "logged_in"
    if any(x in t for x in ["stay signed in", "保持登录"]):  return "logged_in"
    # 微软 Loading 转圈页(PX 按压后等解锁结果):第一行 Loading + 页脚,排除 PX 文案
    if _ruoyi._microsoft_loading_page(page):               return "loading"
    if "account.live.com/abuse" in u:                       return "abuse"
    if "account.microsoft.com" in u and "unlock" not in u: return "logged_in"
    if "account.live.com" in u and "proofs" in u:          return "logged_in"
    if "fido/create" in u or "fido/update" in u:           return "fido_setup"
    if "setting up your passkey" in t or "passkey" in t:   return "fido_setup"
    if any(x in t for x in ["your account has been locked", "we've locked",
                              "locked for your protection", "帐户已锁定"]): return "locked"
    if any(x in t for x in ["enter the code", "we texted", "we sent", "verification code",
                              "验证码", "短信"]): return "sms_verify"
    if any(x in t for x in ["verify your identity", "unusual activity"]): return "verify_needed"
    if "something went wrong" in t: return "error_page"
    if "chrome-error://" in u or "about:neterror" in u: return "net_error"
    if "enter your password" in t:  return "login_form"
    if any(x in t for x in ["email or phone", "sign in", "enter your email"]): return "email_form"
    return "unknown"


def snap(page, tag, name, idx):
    try:
        path = f"{SCREENSHOT_DIR}/{tag}_{name}.png"
        page.screenshot(path=path, full_page=True)
    except Exception:
        pass
    state = classify(page)
    print(f"    [{name}] {state}  {(page.url or '')[:60]}")
    return state


def clear_live_cookies(page):
    """账号间隔离:清 live.com 域 cookie,避免上个账号 session 残留。"""
    try:
        page.run_js_loaded(
            "document.cookie.split(';').forEach(function(c){"
            "var k=(c.split('=')[0]||'').trim(); if(!k) return;"
            "['','.live.com','login.live.com','account.live.com'].forEach(function(d){"
            "document.cookie=k+'=;expires=Thu, 01 Jan 1970 00:00:00 GMT;domain='+d+';path=/';"
            "});}); return true;"
        )
    except Exception:
        pass


def _fill_login_email(page, email, tag):
    try:
        el = _ele(page, 'css:input[name="loginfmt"]', timeout=8) \
             or _ele(page, 'css:input[type="email"]', timeout=2)
        if el:
            _safe_input(el, email)
            _click_next(page, tag)
    except Exception as e:
        print(f"    fill(email) error: {e}")


def _fill_login_password(page, password, tag):
    try:
        el = _ele(page, 'css:input[type="password"]', timeout=8)
        if el:
            _safe_input(el, password)
            _click_next(page, tag)
    except Exception as e:
        print(f"    fill(pwd) error: {e}")


def _page_go_back(page):
    for fn in ("back", "go_back"):
        m = getattr(page, fn, None)
        if callable(m):
            try:
                m(); return
            except Exception:
                pass
    try:
        page.run_js_loaded("history.back();")
    except Exception:
        pass


# ── per-tab 代理归一化 ─────────────────────────────────────────────
def _per_tab_format(proxy_str):
    """归一化为 ruyipage set_per_tab_proxies 要求的 scheme://host:port:user:pass。
    支持多种输入格式:
      user:pass@host:port | host:port:user:pass
      socks5://user:pass@host:port | http://user:pass@host:port
      socks5://host:port:user:pass (已合规,原样返回)
      host:port (无认证,per-tab 不支持,抛错跳过)
    per-tab 内核按 userContextId 认证,必须带账号密码,且 user/pass 不能含冒号。
    """
    s = str(proxy_str or "").strip()
    if not s:
        raise ValueError("空代理")
    # 已是 ruyipage 私有格式 scheme://host:port:user:pass(:// 后无 @) -> 原样
    if "://" in s and "@" not in s.split("://", 1)[1]:
        return s
    parsed = _ruoyi._parse_ruoyi_proxy(s)
    if not parsed or not parsed.get("host") or not parsed.get("port"):
        raise ValueError(f"无法解析代理: {proxy_str}")
    scheme = (parsed.get("scheme") or "socks5").lower()
    user = parsed.get("username") or ""
    pwd = parsed.get("password") or ""
    if not user or not pwd:
        raise ValueError(f"per-tab 代理必须带账号密码: {proxy_str}")
    if ":" in user or ":" in pwd:
        raise ValueError(f"per-tab 代理 user/pass 不能含冒号: {proxy_str}")
    return f"{scheme}://{parsed['host']}:{parsed['port']}:{user}:{pwd}"


# ── 浏览器启动 (参考 ruyi register_outlook :8844-8985, 精简)─────────
def launch_firefox(proxy_pool, idx, headless, concurrency, tag):
    _install_shutdown_handlers()
    user_agent = _pick_user_agent(idx)
    opts = SimpleNamespace(concurrency=concurrency, ruoyi_slot=None)

    tb = FirefoxOptions()
    tb.set_browser_path(RUOYI_FIREFOX_PATH)
    tb.set_profile(_ruoyi_profile_dir(opts, idx))

    if proxy_pool:
        per_tab = []
        for p in proxy_pool:
            try:
                per_tab.append(_per_tab_format(p))
            except ValueError as exc:
                print(f"  {tag} 跳过不合规 per-tab 代理: {exc}", file=sys.stderr)
        if per_tab:
            tb.set_per_tab_proxies(per_tab, exhausted="wrap")
            print(f"  {tag} 挂载 {len(per_tab)}/{len(proxy_pool)} 条代理到 per-tab 池(wrap)")
        else:
            print(f"  {tag} 代理池归一化后为空 -- 本机直连", file=sys.stderr)
    else:
        print(f"  {tag} 没挂代理 -- 直接本机出口", file=sys.stderr)

    _apply_ruoyi_browser_ua(tb, tag, user_agent)
    if headless:
        _apply_ruoyi_headless_options(tb, tag, user_agent=user_agent)
        tb.headless(True)

    print(f"  {tag} 启动 ruyipage Firefox: model={_browser_model_name(RUOYI_FIREFOX_PATH)} "
          f"headless={headless} ua={user_agent[:60]}...")

    page = FirefoxPage(tb)
    _track_browser_page(page)
    setattr(page, "_ruoyi_px_motion_profile", _new_ruoyi_px_motion_profile())

    try:
        page.close_other_tabs(page)
    except Exception:
        pass

    if proxy_pool:
        try: _log_current_ip(proxy_pool, tag)
        except Exception: pass

    return page


def close_firefox(page):
    """关闭单个账号浏览器:quit(force=True) 真正杀 firefox.exe 进程。

    ruyipage 的 page.close() 只关当前标签页(源码 _pages/firefox_page.py 注释
    明写"关闭当前标签页"),不杀进程;若用它,每做完一个账号 firefox 窗口残留,
    concurrency=N 时实际窗口数会随账号数累积远超 N。必须用 quit(force=True)。

    切勿在此调 _force_kill_ruoyi_firefox():它按 RUOYI_FIREFOX_PATH 全局强杀
    所有 ruyi firefox,会误杀同批其他正在运行的 worker 实例。全局强杀只在
    批次边界(所有 worker 空闲)调用,见 run() 末尾。
    """
    if page is None:
        return
    try:
        _quit_browser_page(page, tag="unlock")
    except Exception:
        pass
    _untrack_browser_page(page)


def _load_login_page(page):
    """加载 LOGIN_URL 并等渲染;wait_loading 超时/异常返回 False。

    返回 False 视为连接问题(代理不稳/目标站加载超时),计入 net_err_count,
    累计达阈值返回 proxy_dead -> discard 节点换新重试,避免空转空跑。
    """
    try:
        page.get(LOGIN_URL)
    except Exception:
        return False
    try:
        page.wait_loading(20)
        return True
    except Exception:
        return False


# ── Core unlock logic (照搬原 unlock_account, 改 ruyipage API)────────
def unlock_account(page, email, password, tag, idx, max_press=DEFAULT_MAX_PRESS, timeout=UNLOCK_TIMEOUT):
    os.makedirs(SCREENSHOT_DIR, exist_ok=True)
    deadline = time.time() + timeout

    # ── Step 1: Login ────────────────────────────────────────────────
    net_err_count = 0
    if not _load_login_page(page):
        net_err_count += 1
    time.sleep(2)

    for i in range(20):
        if time.time() > deadline: return "timeout"
        state = snap(page, tag, f"L{i:02d}", idx)
        if state == "abuse":
            _click_any(page, _CONTINUE_SELECTORS, timeout=2)
            time.sleep(3); continue
        if state == "logged_in":  return "already_ok"
        if state == "sms_verify": return "needs_phone"
        if state == "fido_setup":
            _maybe_skip_passkey(page, tag); time.sleep(4)
            return "unlocked"
        if state in ("locked", "px_challenge"): break
        if state == "net_error":
            net_err_count += 1
            if net_err_count >= 3: return "proxy_dead"
            if not _load_login_page(page):
                net_err_count += 1
                if net_err_count >= 3: return "proxy_dead"
            time.sleep(3); continue
        if state == "error_page":
            tried = _click_any(page, _RETRY_SELECTORS, timeout=2)
            if not tried:
                _page_go_back(page)
            time.sleep(5 if tried else 3)
            continue
        if state == "email_form":
            _fill_login_email(page, email, tag)
            time.sleep(3); continue
        if state == "login_form":
            _fill_login_password(page, password, tag)
            time.sleep(3); continue
        if not _click_any(page, ['css:#idSIButton9', 'css:button[type="submit"]',
                                  'css:input[type="submit"]'], timeout=2):
            time.sleep(2)

    state = snap(page, tag, "L_final", idx)
    if state == "logged_in":  return "already_ok"
    if state == "sms_verify": return "needs_phone"
    if state == "fido_setup":
        _maybe_skip_passkey(page, tag); time.sleep(4)
        return "unlocked"

    # ── Step 2: PX press-and-hold + unlock flow ──────────────────────
    press_count   = 0
    no_btn_rounds = 0
    px_api_tried  = False
    net_err_count = 0
    abuse_rounds  = 0
    loading_wait_started = None

    for i in range(60):
        if time.time() > deadline: return "timeout"
        state = snap(page, tag, f"U{i:02d}", idx)

        # 离开 Loading 页重置计时(loading -> 别的状态 -> 再 loading 重新计时)
        if state != "loading" and loading_wait_started is not None:
            loading_wait_started = None

        if state == "loading":
            # 微软 Loading 转圈页(PX 按压后等解锁结果):等它转完出 unblocked/跳转
            # 参考注册 _update_loading_wait_state + loading_timed_out -> return timeout
            if loading_wait_started is None:
                loading_wait_started = time.time()
                print(f"    [{tag}] Microsoft Loading, waiting for redirect...")
            elif time.time() - loading_wait_started > LOADING_WAIT_TIMEOUT:
                print(f"    [{tag}] Loading stuck {int(time.time()-loading_wait_started)}s, give up", file=sys.stderr)
                return "failed_loading_timeout"
            time.sleep(3); continue

        if state == "abuse":
            # Abuse?id=389 中间过渡页:点英文"继续"尝试跳 logged_in
            # 跳到 logged_in 才算成功;连点 5 轮没跳走 -> failed_abuse,不误判成功
            if abuse_rounds >= 5:
                return "failed_abuse"
            abuse_rounds += 1
            _click_any(page, _CONTINUE_SELECTORS, timeout=2)
            time.sleep(3); continue
        if state == "logged_in":  return "unlocked"
        if state == "sms_verify": return "needs_phone"
        if state == "fido_setup":
            _maybe_skip_passkey(page, tag); time.sleep(4)
            return "unlocked"
        if state == "error_page":
            tried = _click_any(page, _RETRY_SELECTORS, timeout=2)
            if not tried:
                _page_go_back(page)
            time.sleep(5 if tried else 3)
            continue
        if state == "net_error":
            net_err_count += 1
            if net_err_count >= 5: return "proxy_dead"
            if not _load_login_page(page):
                net_err_count += 1
                if net_err_count >= 5: return "proxy_dead"
            time.sleep(3); continue

        if state == "locked":
            _click_any(page, ['css:button[type="submit"]', 'css:#idSIButton9',
                              'css:input[type="submit"]'], timeout=2)
            time.sleep(6); continue

        if state == "px_challenge":
            if press_count < max_press:
                # 等待找到真正按压按钮(#px-captcha 等)再按,不用 iframe-box fallback 急按
                ctx = None; target = None
                for c in _ruoyi._all_contexts(page):
                    try:
                        t = _ruoyi._find_hold_target(c)
                    except Exception:
                        t = None
                    if t and _ruoyi._target_quality(t) <= 5:
                        ctx, target = c, t
                        break
                if target and ctx:
                    press_count += 1
                    held = _perform_hold_with_px_screenshots(
                        page, ctx, target, idx, press_count, tag, enabled=True)
                    print(f"    held {held} (#{press_count})")
                    _wait_before_next_captcha_press(tag, reason="unlock press")
                    no_btn_rounds = 0
                else:
                    no_btn_rounds += 1
                    print(f"    no hold target (round {no_btn_rounds})")
                    time.sleep(3)
            elif not px_api_tried:
                px_api_tried = True
                print("    fallback: EZCaptcha PX API...")
                sol = solve_px(page_url=page.url)
                if sol:
                    inject_px_solution(page, sol)
                    try: page.get(page.url)
                    except Exception: pass
                    time.sleep(5)
                else:
                    print("    PX API failed - giving up"); break
            else:
                print("    all PX attempts exhausted"); break
            continue

        # Generic next/submit for intermediate steps
        if not _click_any(page, ['css:button[type="submit"]', 'css:#idSIButton9',
                                  'css:input[type="submit"]'], timeout=2):
            time.sleep(3)

    state = snap(page, tag, "U_final", idx)
    if state == "logged_in":  return "unlocked"
    if state == "abuse":      return "failed_abuse"
    if state == "sms_verify": return "needs_phone"
    if state == "fido_setup":
        _maybe_skip_passkey(page, tag); time.sleep(4)
        return "unlocked"
    return f"failed_{state}"


# ── Worker ────────────────────────────────────────────────────────────
async def _attempt_account(pool, worker_id, args, concurrency, tag, email, password):
    """take 代理 -> 开浏览器 -> 解锁 -> 关浏览器。返回 outcome。
    proxy_dead 时调用方 discard 失效代理后重试(换节点)。"""
    selected_pool = []
    reused = False
    if pool is not None:
        try:
            selected_pool = select_proxy_for_account(pool, runtime=pool) or []
        except Exception:
            selected_pool = []
        if not selected_pool and hasattr(pool, "borrow_reuse"):
            selected_pool = pool.borrow_reuse() or []
            reused = bool(selected_pool)
        masked = mask_ruoyi_proxy(selected_pool[0]) if selected_pool else "noproxy"
        rem = pool.remaining() if hasattr(pool, "remaining") else "?"
        print(f"[worker-{worker_id}] proxy -> {masked} remaining={rem}{' [reused]' if reused else ''}")

    # 代理预检(开浏览器前):走代理打 login.live.com / signup.live.com,
    # 不通直接 proxy_dead -> 末尾 discard 节点换新重试,不浪费浏览器窗口。
    # 与 ruyi 注册 _probe_proxy_before_browser 一致(打目标站本身,非 IP 查询服务)。
    if selected_pool:
        try:
            precheck_ok = await asyncio.to_thread(
                _probe_proxy_before_browser, selected_pool, f"[{tag}][precheck]")
        except Exception as exc:
            print(f"[worker-{worker_id}] proxy precheck error: {exc}", file=sys.stderr)
            precheck_ok = False
        if not precheck_ok:
            print(f"[worker-{worker_id}] proxy precheck failed (login.live.com 不可达) -> proxy_dead", file=sys.stderr)
            # discard 废代理(非 reused),让 worker 重试 take 换新节点;reused 不动(共享出口)
            if pool is not None and not reused and hasattr(pool, "discard"):
                try: pool.discard(selected_pool[0])
                except Exception: pass
            return "proxy_dead"

    page = None
    try:
        page = await asyncio.to_thread(
            launch_firefox, selected_pool, worker_id, args.headless, concurrency, tag)
    except Exception as e:
        print(f"[worker-{worker_id}] launch error: {e}")
        # launch 失败多半代理问题:discard 失效代理,return proxy_dead 让 worker 换节点重试(不当放弃)
        if pool is not None and selected_pool and not reused:
            if hasattr(pool, "discard"):
                try: pool.discard(selected_pool[0])
                except Exception: pass
            else:
                try: release_proxy_for_account(selected_pool[0], runtime=pool)
                except Exception: pass
        return "proxy_dead"

    try:
        await asyncio.to_thread(clear_live_cookies, page)
        outcome = await asyncio.to_thread(
            unlock_account, page, email, password, tag, worker_id,
            args.max_press, args.timeout)
    except Exception as e:
        outcome = f"error: {str(e)[:80]}"
    finally:
        await asyncio.to_thread(close_firefox, page)

    # proxy_dead: 代理失效,discard 永久移除(让重试 take 换新节点);其他结果正常 release
    if pool is not None and selected_pool and not reused:
        if outcome == "proxy_dead" and hasattr(pool, "discard"):
            try: pool.discard(selected_pool[0])
            except Exception: pass
        else:
            try: release_proxy_for_account(selected_pool[0], runtime=pool)
            except Exception: pass
    return outcome


async def worker(accounts, pool, worker_id, results, sem, args, concurrency):
    async with sem:
        tag = f"w{worker_id}"
        for email, password, raw_line in accounts:
            print(f"\n[worker-{worker_id}] {email}")
            outcome = await _attempt_account(pool, worker_id, args, concurrency, tag, email, password)
            # 代理无法访问微软 -> discard 失效节点 -> 换新节点重开浏览器重试
            retry = 0
            while outcome == "proxy_dead" and retry < MAX_PROXY_RETRY:
                retry += 1
                print(f"[worker-{worker_id}] 代理无法访问微软,换节点重试 {retry}/{MAX_PROXY_RETRY}", file=sys.stderr)
                outcome = await _attempt_account(pool, worker_id, args, concurrency, tag, email, password)
            print(f"[worker-{worker_id}] {email} => {outcome}")
            results.append((email, password, raw_line, outcome))


# ── File I/O (照搬原版)───────────────────────────────────────────────
def load_accounts(path):
    accounts = []
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"): continue
            parts = line.split("----")
            if len(parts) >= 2:
                accounts.append((parts[0].strip(), parts[1].strip(), line))
            else:
                print(f"[warn] skip: {line[:60]}")
    return accounts

def scan_all_accounts():
    """Scan outlook_accounts/ for all registered accounts, skip those already unlocked."""
    reg_dir = "outlook_accounts"
    unlock_dir = "unlock_results"

    unlocked_emails = set()
    if os.path.isdir(unlock_dir):
        for uf in os.listdir(unlock_dir):
            if uf.startswith("unlocked_clean_") and uf.endswith(".txt"):
                with open(os.path.join(unlock_dir, uf), "r", encoding="utf-8") as f:
                    for line in f:
                        parts = line.strip().split("----")
                        if parts and parts[0]:
                            unlocked_emails.add(parts[0].lower())

    seen = set()
    accounts = []
    if os.path.isdir(reg_dir):
        for af in sorted(os.listdir(reg_dir)):
            if af.startswith("accounts_") and af.endswith(".txt"):
                with open(os.path.join(reg_dir, af), "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if not line or line.startswith("#"): continue
                        parts = line.split("----")
                        if len(parts) >= 2:
                            email_lc = parts[0].lower()
                            if email_lc not in seen and email_lc not in unlocked_emails:
                                accounts.append((parts[0].strip(), parts[1].strip(), line))
                                seen.add(email_lc)

    if unlocked_emails:
        print(f"[auto] Skipping {len(unlocked_emails)} already-unlocked accounts")
    print(f"[auto] {len(accounts)} new accounts to unlock from {reg_dir}/")
    return accounts

def save_results(results, ts):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    unlocked   = [r for r in results if r[3] in ("unlocked", "already_ok")]
    needs_ph   = [r for r in results if r[3] == "needs_phone"]
    failed     = [r for r in results if r[3] not in ("unlocked", "already_ok", "needs_phone")]

    def write(name, rows):
        p = os.path.join(OUTPUT_DIR, f"{name}_{ts}.txt")
        with open(p, "w", encoding="utf-8") as f:
            for email, password, raw, outcome in rows:
                f.write(f"{raw}----{outcome}\n")
        print(f"  {name:<22s} {len(rows):4d}  -> {p}")

    print(f"\n{'='*55}")
    write("unlocked", unlocked)
    write("needs_phone", needs_ph)
    write("failed", failed)
    print(f"{'─'*55}")
    print(f"  Total     : {len(results)}")
    print(f"  Unlocked  : {len(unlocked)}")
    print(f"  NeedsPhone: {len(needs_ph)}")
    print(f"  Failed    : {len(failed)}")
    print(f"{'='*55}")

    ok_path = os.path.join(OUTPUT_DIR, f"unlocked_clean_{ts}.txt")
    with open(ok_path, "w", encoding="utf-8") as f:
        for email, password, _, _ in unlocked:
            f.write(f"{email}----{password}\n")
    if unlocked:
        print(f"\n  Clean unlocked list: {ok_path}")


# ── 代理池构建 (参考 outlook_reg_loop._one_attempt_ruoyi)────────────
def build_pool(args):
    source_args = SimpleNamespace(
        proxy_file=args.proxy_file or "",
        proxy_source=args.proxy_source,
        proxy_url=getattr(args, "proxy_url", "") or "",
        aimili_url=args.aimili_url,
        aimili_token=args.aimili_token,
    )
    pool = ConsumableProxyPool.from_args(source_args).start()
    set_consumable_proxy_pool(pool)
    st = pool.stats() if hasattr(pool, "stats") else {}
    print(f"[pool] source={st.get('source', args.proxy_source)} size={st.get('remaining', '?')}")
    return pool


# ── Main ──────────────────────────────────────────────────────────────
def find_latest_input():
    """Auto-find most recent accounts file to unlock."""
    for d, pat in [
        ("check_results",   "locked_for_unlock_"),
        ("outlook_accounts","accounts_"),
    ]:
        if not os.path.isdir(d): continue
        files = sorted(
            [f for f in os.listdir(d) if f.startswith(pat) and f.endswith(".txt")],
            reverse=True
        )
        if files:
            return os.path.join(d, files[0])
    return None

async def run(accounts_or_file, args, pool):
    if isinstance(accounts_or_file, str):
        accounts = load_accounts(accounts_or_file)
        label = accounts_or_file
    else:
        accounts = accounts_or_file
        label = f"(auto-scanned, {len(accounts)} accounts)"

    if not accounts:
        print("[error] no accounts found"); return

    concurrency = args.concurrency
    print(f"Input     : {label}")
    print(f"Accounts  : {len(accounts)}")
    print(f"Concurrency: {concurrency}")

    results = []
    sem     = asyncio.Semaphore(concurrency)
    chunks  = [[] for _ in range(concurrency)]
    for i, acc in enumerate(accounts):
        chunks[i % concurrency].append(acc)

    await asyncio.gather(*[
        worker(chunks[i], pool, i, results, sem, args, concurrency)
        for i in range(concurrency)
        if chunks[i]
    ])

    # 批次结束:所有 worker 已空闲,强杀残留 ruyi Firefox(quit 超时兜底,防窗口累积/XPCOM)
    try:
        _force_kill_ruoyi_firefox()
    except Exception:
        pass

    save_results(results, datetime.now().strftime("%Y%m%d_%H%M%S"))

def _install_force_shutdown():
    """Ctrl-C 直接强杀:不等浏览器优雅关闭。
    解锁卡在被锁/PX 页时,ruyi 的 _close_tracked_browser_pages 同步关 Firefox 会挂,
    导致优雅 handler 卡死、Ctrl-C 无响应。这里直接 os._exit。"""
    _ruoyi._SHUTDOWN_HANDLERS_INSTALLED = True  # 阻止 launch_firefox 内 _install_shutdown_handlers 装优雅 handler

    def _handler(signum, _frame):
        print("\n[shutdown] 收到中断,强杀退出(os._exit)", file=sys.stderr)
        os._exit(0)

    for sig_name in ("SIGINT", "SIGBREAK", "SIGTERM"):
        sig = getattr(signal, sig_name, None)
        if sig is not None:
            try:
                signal.signal(sig, _handler)
            except Exception:
                pass


def main():
    parser = argparse.ArgumentParser(
        description="Batch Outlook Account Unlock (ruyipage Firefox edition)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python unlock_outlook.py --input outlook_accounts/accounts_20260414_124527.txt
  python unlock_outlook.py --input emails_locked.txt --concurrency 2
  python unlock_outlook.py --proxy-file proxies_outlook.txt --proxy-source aimili-list --aimili-url http://host:8787 --aimili-token XXX
  python unlock_outlook.py                          (auto-scan all accounts, skip unlocked)
""")
    parser.add_argument("--input", "-i", default=None,
        help="Input file (email----password per line). "
             "Auto-scans outlook_accounts/ and skips already-unlocked if omitted.")
    parser.add_argument("--proxy-file", "-p",
        default=os.environ.get("OUTLOOK_PROXY_FILE", "proxies_outlook.txt"),
        help="Proxy list file (one user:pass@host:port per line)")
    parser.add_argument("--proxy-source",
        default=os.environ.get("OUTLOOK_RUOYI_PROXY_SOURCE", "file"),
        choices=["file", "http", "aimili-list", "aimili-random"],
        help="file/http/aimili-list/aimili-random (file=本地文件;http=HTTP GET 拉 txt 列表)")
    parser.add_argument("--proxy-url",
        default=os.environ.get("OUTLOOK_PROXY_URL", ""),
        help="HTTP GET 代理列表地址(返回 txt,每行一条;配合 --proxy-source=http)")
    parser.add_argument("--aimili-url",
        default=(os.environ.get("OUTLOOK_AIMILI_POOL_URL")
                 or os.environ.get("OUTLOOK_AIMILI_POOL_BASE_URL", "")),
        help="AimiliVPN URL: 根地址或 /api/pool/proxies(/random) 完整地址")
    parser.add_argument("--aimili-token",
        default=os.environ.get("OUTLOOK_AIMILI_POOL_TOKEN", ""),
        help="AimiliVPN 代理池 API Token")
    parser.add_argument("--concurrency", "-c", type=int, default=1,
        help="Parallel workers (default: 1)")
    parser.add_argument("--headless", action="store_true",
        help="无头模式启动 Firefox")
    parser.add_argument("--max-press", type=int,
        default=int(os.environ.get("OUTLOOK_REG_MAX_PRESS", str(DEFAULT_MAX_PRESS))),
        help="PX press-and-hold cap (default: 5)")
    parser.add_argument("--timeout", type=int, default=UNLOCK_TIMEOUT,
        help="hard cap per account (seconds)")
    parser.add_argument("--log-level",
        default=os.environ.get("OUTLOOK_LOG_LEVEL", "INFO"),
        choices=["DEBUG", "INFO", "WARN", "PROD", "ERR"],
        help="log verbosity")
    args = parser.parse_args()

    set_log_level(args.log_level)
    _install_force_shutdown()
    os.environ["OUTLOOK_PROXY_FILE"] = args.proxy_file
    os.environ["OUTLOOK_RUOYI_PROXY_SOURCE"] = args.proxy_source
    if args.aimili_url:   os.environ["OUTLOOK_AIMILI_POOL_URL"] = args.aimili_url
    if args.aimili_token: os.environ["OUTLOOK_AIMILI_POOL_TOKEN"] = args.aimili_token
    if getattr(args, "proxy_url", ""): os.environ["OUTLOOK_PROXY_URL"] = args.proxy_url
    os.environ["OUTLOOK_REG_MAX_PRESS"] = str(args.max_press)

    if args.input:
        if not os.path.exists(args.input):
            print(f"[error] file not found: {args.input}")
            sys.exit(1)
        accounts_or_file = args.input
    else:
        accounts_or_file = scan_all_accounts()
        if not accounts_or_file:
            print("[info] No new accounts to unlock.")
            sys.exit(0)

    pool = build_pool(args)
    asyncio.run(run(accounts_or_file, args, pool))

if __name__ == "__main__":
    main()
