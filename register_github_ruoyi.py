#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GitHub 自动注册(ruoyi) —— 复用 ruyi 公共组件

对齐 register_outlook_ruoyi.py 骨架,把 GitHub 注册从 BitBrowser/Playwright 迁移到
ruyipage(Firefox BiDi)+ common/ruyi 公共包:
  - 代理池 ConsumableProxyPool + select_proxy_for_account
  - profile _ruoyi_profile_dir / _cleanup_ruoyi_run_profile_dir
  - firefox get_firefox_path / _kill_ruoyi_firefox_by_profile
  - launch FirefoxOptions/FirefoxPage + after_launch 反检测
  - warmup 首屏预热(落地 github.com 种匿名 cookie)
  - probe _log_current_ip_async
  - ua _pick_user_agent

GitHub 专属(公共组件不覆盖,保留自 register_github.py):
  - 单页表单 email/password/username + 国家下拉
  - Create account 触发 Arkose FunCaptcha
  - Arkose token 回灌(postMessage captcha-complete)
  - 邮件 launch code 取信(复用 _outlook_pool 邮箱池,浏览器登录 Outlook)
  - cookie 落盘(适配 ruyipage get_cookies)

用法:
    python register_github_ruoyi.py                # 默认:填到验证步停下
    python register_github_ruoyi.py --auto          # 试走完整流程
    python register_github_ruoyi.py --headless      # 无头
    python register_github_ruoyi.py --proxy-file proxies_github.txt
    python register_github_ruoyi.py --email a@b.com --password xxx   # 指定邮箱
"""

from __future__ import annotations

import argparse
import atexit
import glob
import json
import os
import random
import re
import requests
import string
import sys
import threading
import time
from datetime import datetime
from types import SimpleNamespace

from config import _load_dotenv

from common import ruyi as _ruyi_pkg

from ruyipage import NoneElement

_load_dotenv()

# ── env 兼容层:GitHub 也认 OUTLOOK_* 名(同一套代理池),提升到中性名 ──
DEFAULT_RUOYI_FIREFOX = _ruyi_pkg.DEFAULT_RUOYI_FIREFOX
_resolve_ruoyi_firefox_path = _ruyi_pkg._resolve_ruoyi_firefox_path
get_firefox_path = _ruyi_pkg.get_firefox_path
RUOYI_FIREFOX_PATH = _ruyi_pkg.RUOYI_FIREFOX_PATH

ROOT = os.path.dirname(os.path.abspath(__file__))

# 代理池:优先 GITHUB_* 名,回落 OUTLOOK_*(同池),再中性名。
PROXY_FILE = os.environ.get(
    "GITHUB_PROXY_FILE",
    os.environ.get("OUTLOOK_PROXY_FILE", _ruyi_pkg.PROXY_FILE),
)
os.environ.setdefault(
    "RUOYI_PROXY_FILE",
    os.environ.get("GITHUB_PROXY_FILE", os.environ.get("OUTLOOK_PROXY_FILE", "")),
)
RUOYI_PROXY_SOURCE = os.environ.get(
    "GITHUB_RUOYI_PROXY_SOURCE",
    os.environ.get("OUTLOOK_RUOYI_PROXY_SOURCE", _ruyi_pkg.RUOYI_PROXY_SOURCE),
)
os.environ.setdefault(
    "RUOYI_PROXY_SOURCE",
    os.environ.get("GITHUB_RUOYI_PROXY_SOURCE", os.environ.get("OUTLOOK_RUOYI_PROXY_SOURCE", "")),
)
PROXY_URL = os.environ.get(
    "GITHUB_PROXY_URL",
    os.environ.get("OUTLOOK_PROXY_URL", _ruyi_pkg.PROXY_URL),
)
os.environ.setdefault(
    "RUOYI_PROXY_URL",
    os.environ.get("GITHUB_PROXY_URL", os.environ.get("OUTLOOK_PROXY_URL", "")),
)

SCREENSHOT_DIR = os.path.join(ROOT, "screenshots_github_ruoyi")
OUTPUT_DIR = os.path.join(ROOT, "github_accounts")
RUOYI_PROFILE_ROOT = os.environ.get(
    "GITHUB_RUOYI_PROFILE_ROOT",
    os.environ.get("OUTLOOK_RUOYI_PROFILE_ROOT", os.path.join(ROOT, "profiles_ruoyi_tmp")),
)
_ruyi_pkg.set_profile_root(RUOYI_PROFILE_ROOT)

# ── GitHub 常量 ──
PLATFORM = "github"
SIGNUP_URL = "https://github.com/signup"
WARMUP_HOME_URL = "https://github.com/"     # warmup 落地:种第一方匿名 cookie
KEY_COOKIES = ["user_session", "__Host-user_session_same_site", "_gh_sess"]
POOL_DIR = "_outlook_pool"
REGISTER_TIMEOUT = 600

# GitHub 验证 = Arkose Labs FunCaptcha(实测固定参数)
ARKOSE_PUBLIC_KEY = "747B83EC-2CA3-43AD-A7DF-701F286FBABA"
ARKOSE_API_SUBDOMAIN = "github-api.arkoselabs.com"

# GitHub 发件人 / launch code 特征
GH_SENDER = ("github.com", "noreply@github.com", "notifications@github.com")
GH_SUBJECT = ("launch code", "github", "verify", "verification", "code")

# 打码 key(config .env)
try:
    from config import (CAPSOLVER_API_KEY, EZCAPTCHA_API_KEY, EZCAPTCHA_API_BASE,
                        YESCAPTCHA_API_KEY, YESCAPTCHA_API_BASE)
except Exception:
    CAPSOLVER_API_KEY = ""
    EZCAPTCHA_API_KEY = ""
    EZCAPTCHA_API_BASE = "https://api.ez-captcha.com"
    YESCAPTCHA_API_KEY = ""
    YESCAPTCHA_API_BASE = "https://api.yescaptcha.com"


# ruyipage page/frame run_js 返回已解析的 Python 原生值(不是 ScriptResult)。
# 自判定:带 args -> callFunction(function 声明);`return ` 开头 -> callFunction;
# 否则 evaluate-as-expression。element 一律走 el.run_js(function 声明)。
def _js(page_or_frame, script, *args, loaded=True):
    fn = "run_js_loaded" if loaded else "run_js"
    runner = getattr(page_or_frame, fn, None) or page_or_frame.run_js
    try:
        return runner(script, *args)
    except Exception:
        return None


def _set_input_value(el, value, clear=True):
    """对 FirefoxElement 用 BiDi-compatible function 声明填 React 受控输入。
    参数走形参(v, doClear),不依赖 this 绑定的 selector。测试过 ruyi callFunction。"""
    if el is None:
        return False
    try:
        ok = el.run_js(
            """
function(v, doClear) {
  const text = String(v ?? '');
  const e = this;
  e.focus();
  const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set
    || Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value')?.set;
  if (doClear) {
    if (setter) setter.call(e, ''); else e.value = '';
    e.dispatchEvent(new Event('input', {bubbles: true}));
  }
  if (setter) setter.call(e, text); else e.value = text;
  e.dispatchEvent(new Event('input', {bubbles: true}));
  e.dispatchEvent(new Event('change', {bubbles: true}));
  return true;
}
            """,
            value,
            bool(clear),
        )
        return bool(ok)
    except Exception:
        return False


def _shot(page, name, idx):
    try:
        import os
        os.makedirs(SCREENSHOT_DIR, exist_ok=True)
        fname = f"{SCREENSHOT_DIR}/{idx or 0:04d}_{name}.png"
        page.screenshot(path=fname)
    except Exception as e:
        print(f"  [shot] {name} err: {str(e)[:60]}")


def _dump_interstitial_html(page, name, idx):
    """DataDome interstitial 落点时把顶层 + DD iframe HTML 落盘,便于离线分析。"""
    import os
    try:
        os.makedirs(SCREENSHOT_DIR, exist_ok=True)
        parts = []
        try:
            parts.append(f"<!-- TOP-LEVEL url={_page_url(page)} -->\n"
                         + (page.html or "<empty>"))
        except Exception as e:
            parts.append(f"<!-- top html err: {type(e).__name__}: {str(e)[:80]} -->")
        fr = _datadome_frame(page)
        if fr is not None:
            try:
                parts.append("<!-- DATADOME IFRAME -->\n" + (fr.html or "<empty>"))
            except Exception as e:
                parts.append(f"<!-- dd frame html err: {type(e).__name__}: {str(e)[:80]} -->")
        else:
            parts.append("<!-- datadome frame NOT FOUND -->")
        fname = f"{SCREENSHOT_DIR}/{idx or 0:04d}_{name}.html"
        with open(fname, "w", encoding="utf-8") as f:
            f.write("\n\n".join(parts))
        print(f"  [dump] interstitial html saved: {fname}")
    except Exception as e:
        print(f"  [dump] {name} err: {str(e)[:60]}")


# ── GitHub 表单工具 ──────────────────────────────────────────────

def _fill(page, selector, value, label="", settle=0.5):
    """填单个字段(GitHub 受控输入)。优先 el.input() 原生填,兜底 run_js setter。
    value: str 或 (label->str) 映射省略,直接传值。"""
    from ruyipage import NoneElement
    try:
        el = page.ele(selector)
        if el is None or isinstance(el, NoneElement):
            if label:
                print(f"  [form] {label}= MISSING ({selector})")
            return False
    except Exception:
        if label:
            print(f"  [form] {label}= MISSING ({selector})")
        return False
    ok = False
    try:
        el.input(str(value))
        ok = True
    except Exception:
        ok = _set_input_value(el, str(value))
    time.sleep(settle)
    if label:
        print(f"  [form] {label}={value} -> {'OK' if ok else 'FAILED'}")
    return ok


def _body_text(page):
    try:
        return _js(page, "return document.body ? document.body.innerText.slice(0, 2000) : '';") or ""
    except Exception:
        return ""


def _page_url(page):
    try:
        return str(getattr(page, "url", "") or "")
    except Exception:
        return ""


def _random_digits(n):
    return "".join(random.choices(string.digits, k=n))


def rand_password():
    # GitHub 要求 >=15 位,或 >=8 位含数字+小写。给足 16 位混合最稳。
    return "Gh1!" + "".join(random.choices(string.ascii_letters + string.digits, k=14))


def rand_username():
    adj = random.choice(["cool", "fast", "blue", "red", "neo", "sky", "dev", "byte", "code", "pixel"])
    noun = random.choice(["fox", "wolf", "cat", "owl", "bear", "hawk", "lion", "frog", "deer", "crab"])
    return f"{adj}{noun}{random.randint(1000, 9999)}"


def load_pool_accounts():
    """读 _outlook_pool/*.json -> [(email, password)] 最新优先。"""
    files = sorted(glob.glob(os.path.join(POOL_DIR, "*.json")), reverse=True)
    out = []
    for f in files:
        try:
            d = json.load(open(f, encoding="utf-8"))
            email, pw = d.get("email"), d.get("password")
            if email and pw:
                out.append((email, pw))
        except Exception:
            continue
    return out


def _click_any_frame(page, frame_urls, text, max_wait=8, _exact=False):
    """跨 frame 点文本元素(Arkose 的 Visual puzzle 按钮在最深处 game frame)。
    用 get_all_frames 递归拿所有后代 frame。返回命中 frame 或 None。"""
    deadline = time.time() + max_wait
    while time.time() < deadline:
        frames = page.get_all_frames() if hasattr(page, "get_all_frames") else (page.get_frames() or [])
        for fr in frames:
            u = str(getattr(fr, "url", "") or "")
            if not any(k in u for k in frame_urls):
                continue
            try:
                els = fr.eles("text=%s" % text) if hasattr(fr, "eles") else []
                if els:
                    els[0].click()
                    return fr
            except Exception:
                pass
        time.sleep(1.5)
    return None


def _screenshot(page, name):
    _shot(page, name, getattr(page, "_idx", 1))


# ── GitHub 验证(Arkose FunCaptcha)- 适配 ruyipage 同步 BiDi ──────
# 旧 Playwright 实现: page.frames 递归 + await el.get_by_text + await el.evaluate。
# ruyipage: get_all_frames()(递归所有后代 frame) + frame.run_js(return 表达式) +
#           frame.eles("text=...") 跨 frame 文本定位。

def _arkose_frames(page, visible_only=False):
    """递归拿所有 arkose/funcaptcha 相关 frame。

    visible_only=True 时只要**可见**的 frame:GitHub signup 页常驻隐藏的
    octocaptcha iframe(0x0/v-hidden,后台收 token)不算——按可见性查顶层 DOM,
    frame 对象本身在 BiDi 里查不到几何,所以从父页 iframe 元素反查。"""
    try:
        frames = page.get_all_frames() if hasattr(page, "get_all_frames") else (page.get_frames() or [])
        hits = []
        for fr in frames or []:
            u = str(getattr(fr, "url", "") or "")
            if not any(k in u for k in ("octocaptcha", "arkose", "funcaptcha")):
                continue
            if visible_only:
                vis = _js(page, """return (() => {
                    for (const f of document.querySelectorAll('iframe')) {
                        if (!(f.src || '').includes('octocaptcha') && !(f.src || '').includes('arkose')
                            && !(f.src || '').includes('funcaptcha')) continue;
                        const r = f.getBoundingClientRect();
                        const cs = getComputedStyle(f);
                        if (r.width > 0 && r.height > 0 && cs.visibility !== 'hidden' && cs.display !== 'none')
                            return true;
                    }
                    return false;
                })();""", loaded=False)
                if not vis:
                    continue
            hits.append(fr)
        return hits
    except Exception:
        return []


def click_visual_puzzle(page, max_wait=50):
    """点 octocaptcha 深处的 'Visual puzzle' 按钮。
    关键：这一步触发 loadFunCaptchaV2 —— 建立 Arkose onCompleted 回调(postMessage
    captcha-complete)+ 创建 #funcaptcha 元素(带 data-target-origin / data-data-exchange-payload)。
    实测坑：Create account 后 Arkose 先跑 ~16s proof-of-work,之后才出选择页,所以要轮询久。"""
    deadline = time.time() + max_wait
    while time.time() < deadline:
        for fr in _arkose_frames(page):
            try:
                els = fr.eles("text=Visual puzzle") if hasattr(fr, "eles") else []
                if els:
                    els[0].click()
                    print("  [arkose] clicked 'Visual puzzle'")
                    return True
            except Exception:
                pass
        time.sleep(3)
    print("  [arkose] 'Visual puzzle' 没等到(可能直接进拼图)")
    return False


def _extract_arkose_blob(page):
    """取 #funcaptcha 的 data-data-exchange-payload(打码必须带上,否则 token 不被 GitHub 接受)。"""
    for fr in _arkose_frames(page):
        try:
            b = _js(fr,
                    "return (() => { const el = document.querySelector('#funcaptcha'); "
                    "return el ? (el.getAttribute('data-data-exchange-payload') || '') : ''; })();",
                    loaded=False)
            if b and str(b).strip():
                return str(b).strip()
        except Exception:
            pass
    # 顶层父页也找一次(GitHub 有时把 blob 放父 DOM)
    try:
        b = _js(page,
                "return (() => { const el = document.querySelector('#funcaptcha'); "
                "return el ? (el.getAttribute('data-data-exchange-payload') || '') : ''; })();",
                loaded=False)
        if b and str(b).strip():
            return str(b).strip()
    except Exception:
        pass
    return None


def _inject_arkose_token(page, token):
    """把打码 token 回灌: 在 octocaptcha frame 里向父页 postMessage captcha-complete。
    target_origin 取 #funcaptcha 的 data-target-origin,缺省 https://github.com。"""
    injected = False
    frames = _arkose_frames(page)
    for fr in frames:
        if "octocaptcha" not in str(getattr(fr, "url", "") or ""):
            continue
        try:
            origin = _js(fr,
                         "return (() => { const el = document.querySelector('#funcaptcha'); "
                         "return el ? (el.getAttribute('data-target-origin') || '') : ''; })();",
                         loaded=False)
            origin = str(origin or "") or "https://github.com"
            _js(fr,
                "function(tok, org) { parent.postMessage({event:'captcha-complete', sessionToken:tok}, org || '*'); return true; }",
                token, origin,
                loaded=False)
            print(f"  [arkose] posted captcha-complete (origin={origin})")
            injected = True
            break
        except Exception:
            continue
    if not injected:
        # 兜底: 向父页及所有 iframe 广播
        try:
            _js(page,
                "function(tok) { const m={event:'captcha-complete', sessionToken:tok}; "
                "window.postMessage(m,'*'); document.querySelectorAll('iframe').forEach(f=>{"
                "try{f.contentWindow.postMessage(m,'*');}catch(e){}}); return true; }",
                token, loaded=False)
            print("  [arkose] fallback: broadcast captcha-complete")
            return True
        except Exception as e:
            print(f"  [arkose] fallback inject error: {str(e)[:80]}")
    return injected


def solve_arkose(page, max_wait=200):
    """拿 FunCaptcha token 并回灌给 GitHub。返回是否注入成功。优先 YesCaptcha。"""
    if not (YESCAPTCHA_API_KEY or CAPSOLVER_API_KEY or EZCAPTCHA_API_KEY):
        print("  [arkose] 无打码 key(YESCAPTCHA/CAPSOLVER/EZCAPTCHA_API_KEY),跳过自动解码")
        return False
    click_visual_puzzle(page)
    time.sleep(3)
    blob = _extract_arkose_blob(page)
    print(f"  [arkose] data-exchange blob: {'got len='+str(len(blob)) if blob else 'NONE (token 可能不被接受)'}")
    print(f"  [arkose] solving FunCaptcha (pk={ARKOSE_PUBLIC_KEY})...")
    token = _solve_funcaptcha_yescaptcha(ARKOSE_PUBLIC_KEY, SIGNUP_URL, ARKOSE_API_SUBDOMAIN, blob, max_wait)
    if not token:
        token = _solve_funcaptcha_capsolver(ARKOSE_PUBLIC_KEY, SIGNUP_URL, ARKOSE_API_SUBDOMAIN, blob, max_wait)
    if not token:
        token = _solve_funcaptcha_ezcaptcha(ARKOSE_PUBLIC_KEY, SIGNUP_URL, ARKOSE_API_SUBDOMAIN, blob, max_wait)
    if not token:
        print("  [arkose] 打码失败")
        return False
    return _inject_arkose_token(page, token)


def _solve_funcaptcha_yescaptcha(public_key, page_url, subdomain, blob=None, max_wait=200):
    """YesCaptcha 解 Arkose FunCaptcha → token。API 与 CapSolver 兼容(type=FunCaptchaTaskProxyless)。"""
    if not YESCAPTCHA_API_KEY:
        return None
    try:
        task = {
            "type": "FunCaptchaTaskProxyless",
            "websiteURL": page_url,
            "websitePublicKey": public_key,
            "funcaptchaApiJSSubdomain": f"https://{subdomain}",
        }
        if blob:
            task["data"] = json.dumps({"blob": blob})
        resp = requests.post(f"{YESCAPTCHA_API_BASE}/createTask",
                             json={"clientKey": YESCAPTCHA_API_KEY, "task": task}, timeout=30)
        data = resp.json()
        if data.get("errorId", 1) != 0:
            print(f"  [yescaptcha] create error: {data.get('errorDescription', data)}")
            return None
        task_id = data["taskId"]
        print(f"  [yescaptcha] funcaptcha task: {task_id}")
        start = time.time()
        while time.time() - start < max_wait:
            time.sleep(6)
            r = requests.post(f"{YESCAPTCHA_API_BASE}/getTaskResult",
                              json={"clientKey": YESCAPTCHA_API_KEY, "taskId": task_id}, timeout=30).json()
            st = r.get("status")
            if st == "ready":
                tok = r.get("solution", {}).get("token") or r.get("solution", {}).get("gRecaptchaResponse")
                print(f"  [yescaptcha] solved (token len={len(tok or '')})")
                return tok
            if st == "failed" or r.get("errorId"):
                print(f"  [yescaptcha] failed: {r.get('errorDescription', '')}")
                return None
        print("  [yescaptcha] timeout")
        return None
    except Exception as e:
        print(f"  [yescaptcha] error: {str(e)[:80]}")
        return None


def _solve_funcaptcha_capsolver(public_key, page_url, subdomain, blob=None, max_wait=180):
    if not CAPSOLVER_API_KEY:
        return None
    try:
        task = {
            "type": "FunCaptchaTaskProxyLess",
            "websiteURL": page_url,
            "websitePublicKey": public_key,
            "funcaptchaApiJSSubdomain": f"https://{subdomain}",
        }
        if blob:
            task["data"] = json.dumps({"blob": blob})
        resp = requests.post("https://api.capsolver.com/createTask",
                             json={"clientKey": CAPSOLVER_API_KEY, "task": task}, timeout=30)
        data = resp.json()
        if data.get("errorId", 1) != 0:
            print(f"  [capsolver] create error: {data.get('errorDescription', data)}")
            return None
        task_id = data["taskId"]
        print(f"  [capsolver] funcaptcha task: {task_id}")
        start = time.time()
        while time.time() - start < max_wait:
            time.sleep(6)
            r = requests.post("https://api.capsolver.com/getTaskResult",
                              json={"clientKey": CAPSOLVER_API_KEY, "taskId": task_id}, timeout=30).json()
            st = r.get("status")
            if st == "ready":
                tok = r.get("solution", {}).get("token")
                print(f"  [capsolver] solved (token len={len(tok or '')})")
                return tok
            if st == "failed" or r.get("errorId"):
                print(f"  [capsolver] failed: {r.get('errorDescription', '')}")
                return None
        print("  [capsolver] timeout")
        return None
    except Exception as e:
        print(f"  [capsolver] error: {str(e)[:80]}")
        return None


def _solve_funcaptcha_ezcaptcha(public_key, page_url, subdomain, blob=None, max_wait=180):
    if not EZCAPTCHA_API_KEY:
        return None
    try:
        task = {
            "type": "FunCaptchaTaskProxyless",
            "websiteURL": page_url,
            "websitePublicKey": public_key,
            "funcaptchaApiJSSubdomain": f"https://{subdomain}",
        }
        if blob:
            task["data"] = json.dumps({"blob": blob})
        resp = requests.post(f"{EZCAPTCHA_API_BASE}/createTask", json={
            "clientKey": EZCAPTCHA_API_KEY,
            "task": task,
        }, timeout=30)
        data = resp.json()
        if data.get("errorId", 1) != 0:
            print(f"  [ezcaptcha] create error: {data.get('errorDescription', data)}")
            return None
        task_id = data["taskId"]
        print(f"  [ezcaptcha] funcaptcha task: {task_id}")
        start = time.time()
        while time.time() - start < max_wait:
            time.sleep(6)
            r = requests.post(f"{EZCAPTCHA_API_BASE}/getTaskResult",
                              json={"clientKey": EZCAPTCHA_API_KEY, "taskId": task_id}, timeout=30).json()
            st = r.get("status")
            if st == "ready":
                tok = r.get("solution", {}).get("token")
                print(f"  [ezcaptcha] solved (token len={len(tok or '')})")
                return tok
            if st == "failed" or r.get("errorId"):
                print(f"  [ezcaptcha] failed: {r.get('errorDescription', '')}")
                return None
        print("  [ezcaptcha] timeout")
        return None
    except Exception as e:
        print(f"  [ezcaptcha] error: {str(e)[:80]}")
        return None


def _detect_captcha(page, max_wait=20):
    """轮询 octocaptcha/arkose frame **可见地**出现。

    关键:GitHub signup 页常驻预埋 octocaptcha iframe(0x0 + v-hidden,后台收集
    token 用,实测 2026-09)。按"URL 命中即算"必然误报——提交状态机 6 轮死循环的
    根因。必须 frame 尺寸 >0 且 visibility 非 hidden 才算真弹出验证视图。"""
    deadline = time.time() + max_wait
    while time.time() < deadline:
        try:
            if _arkose_frames(page, visible_only=True):
                return True
            if _js(page, """return (() => {
                const f = document.querySelector('iframe[src*="octocaptcha"], iframe[src*="arkose"]');
                if (!f) return false;
                const r = f.getBoundingClientRect();
                const cs = getComputedStyle(f);
                return r.width > 0 && r.height > 0 && cs.visibility !== 'hidden' && cs.display !== 'none';
            })();""", loaded=False):
                return True
        except Exception:
            pass
        time.sleep(1.5)
    return False


# ── DataDome 滑块(github.com/signup 门前挑战)──────────────────
# GitHub 把 /signup 直连导航挡在 DataDome 后面(geo.captcha-delivery.com iframe):
# - 直连 /signup → "Verification Required / Slide right to secure your access"
# - 首页点 Sign up 链接 → 多数情况直接放行(人类路径)
# 滑块 handle: iframe 内 .captcha__btn,轨道 .captcha__slider; BiDi performActions
# 的 viewport 坐标可以穿透 iframe(浏览器级输入)。

def _datadome_frame(page):
    """找 geo.captcha-delivery.com 的 frame。

    两级查找:
    1) 常规: page.get_all_frames()(browsingContext.getTree 以 page 为根)
    2) 兜底: 浏览器全局 tree(root=None)里找 captcha-delivery.com context,
       手动包成 FirefoxFrame。实测 interstitial 整页导航后(2026-09)page 为根的
       tree 经常不返回 children,全局 tree 能看到。"""
    try:
        for fr in (page.get_all_frames() or []):
            if "captcha-delivery.com" in str(getattr(fr, "url", "") or ""):
                return fr
    except Exception:
        pass
    try:
        from ruyipage._bidi import browsing_context as _bc
        from ruyipage._pages.firefox_frame import FirefoxFrame
        result = _bc.get_tree(page._driver._browser_driver, root=None)
        contexts = result.get("contexts", [])

        def find_dd(nodes):
            for n in nodes:
                u = str(n.get("url", "") or "")
                if "captcha-delivery.com" in u:
                    return n["context"], u
                hit = find_dd(n.get("children") or [])
                if hit:
                    return hit
            return None

        hit = find_dd(contexts)
        if hit:
            ctx_id, u = hit
            print(f"  [datadome] 全局 tree 找到 DD frame(ctx={ctx_id})")
            return FirefoxFrame(page._browser, ctx_id, page)
    except Exception:
        pass
    return None


def _datadome_state(page):
    """DataDome iframe 状态: ('none'|'restricted'|'blocked'|'slider'|'passed', slider_info|None)。
    blocked = IP 信誉极差时 DataDome 给的 audio/blocked 变体(没有滑块,拖不动)。"""
    fr = _datadome_frame(page)
    if fr is None:
        return ("none", None)
    raw = _js(fr, r"""return (() => {
        const out = {restricted: false, blocked: false, slider: null, track: null};
        const bodyTxt = document.body ? document.body.innerText : '';
        if (/temporarily restricted/i.test(bodyTxt)) out.restricted = true;
        // "audio verification" 不能算 blocked——无障碍提示文案("use the audio verification
        // instead")也含这句(2026-09 实测)。只认真正的 blocked 硬文案。
        if (/you have been blocked/i.test(document.title || '')
            || /you have been blocked\.?$/im.test(bodyTxt)) out.blocked = true;
        // handle/track selector:旧版 .captcha__btn/.captcha__slide;
        // 2026-09 新版 .slider/.sliderContainer(见 dd_interstitial_r1.html dump)
        const sEl = document.querySelector('.captcha__btn, .slider:not(.sliderMask):not(.sliderbg), [role="slider"]');
        const tEl = document.querySelector('.captcha__slider, .sliderContainer, .captcha__slide');
        for (const [el, key] of [[sEl, 'slider'], [tEl, 'track']]) {
            if (!el) continue;
            const r = el.getBoundingClientRect();
            if (r.width < 5 || r.height < 5) continue;
            out[key] = {
                cls: String(el.className || '').slice(0, 100),
                x: Math.round(r.x), y: Math.round(r.y),
                w: Math.round(r.width), h: Math.round(r.height)
            };
        }
        return JSON.stringify(out);
    })();""", loaded=False)
    try:
        data = json.loads(raw) if isinstance(raw, str) else None
    except Exception:
        data = None
    if data is None:
        return ("none", None)
    if data.get("restricted"):
        return ("restricted", None)
    # blocked 判定让位给 slider:2026-09 interstitial 页面文案残留 "blocked"/"audio"
    # 字样但滑块照样渲染(实测截图 0001_05b_landed_r1.png),有滑块先拖,拖不动再说 blocked
    if data.get("slider"):
        return ("slider", data)
    if data.get("blocked"):
        return ("blocked", None)
    return ("none", data)


def _os_drag_available():
    """已废弃:曾试验 pyautogui 真鼠标拖拽,被否——会干扰真实桌面。
    保留空实现只为兼容可能的外部引用,一律走 BiDi(page.actions)路径。"""
    return False


_OS_DRAG_CACHE = False


def _os_drag_flag():
    return False


def _bidi_slider_drag(page, sx, sy, ex, ey, duration=None, frame=None):
    """用 Actions 拟人轨迹(human_move bezier/windmouse)拖滑块。

    与 Outlook PX 按住同源(BiDi performActions,isTrusted=true,实测合成
    事件 screenX=clientX+窗口偏移与真鼠标同构)。全程不动真实鼠标。

    frame 关键性(2026-09 fast20 探针实测): DataDome iframe 是跨域 OOPIF
    (独立进程),从顶层 page.actions 派发 input.performActions 事件**不进
    iframe**(探针 total=0,连按 7 次滑块纹丝不动)。必须对 iframe 自己的
    browsingContext 派发(frame.actions,context=DD frame id),坐标用
    iframe 内部坐标(不加顶层偏移)。"""
    import random
    import random as _rnd
    target = frame if frame is not None else page
    actions = target.actions
    # 接近段:拟人 bezier 移到把手(PX lead/settle 同思路)
    print(f"    [slider] step1 鼠标移向把手 ({int(sx)},{int(sy)}) ctx={'iframe' if frame is not None else 'top'}")
    actions.human_move({"x": int(sx), "y": int(sy)})
    actions.wait(_rnd.uniform(0.15, 0.4))
    # 按下后按住拖:用内置 drag_to(按下→分段移动→释放,单次 perform)
    print(f"    [slider] step2 按下左键并拖动 ({int(sx)},{int(sy)}) -> ({int(ex)},{int(ey)}) 距离={int(ex-sx)}px")
    actions.drag_to({"x": int(sx), "y": int(sy)}, {"x": int(ex), "y": int(ey)},
                    duration=_rnd.randint(800, 1400), steps=_rnd.randint(25, 40))
    print("    [slider] step3 松开左键")
    actions.perform()


_DD_PROBE_JS = r"""return (() => {
    if (window.__ddprobe) return 'already';
    window.__ddprobe = [];
    const rec = (e) => {
        window.__ddprobe.push({
            type: e.type,
            x: Math.round(e.clientX), y: Math.round(e.clientY),
            trusted: e.isTrusted,
            tgt: (e.target.className || e.target.tagName || '').toString().slice(0, 30),
            btns: e.buttons
        });
    };
    for (const t of ['mousedown','mouseup','mousemove'])
        document.addEventListener(t, rec, true);
    const h = document.querySelector('.slider:not(.sliderMask):not(.sliderbg), .captcha__btn');
    if (h) for (const t of ['mousedown','mousemove','mouseup'])
        h.addEventListener(t, rec, true);
    return 'ok handle=' + (h ? 'yes' : 'missing');
})();"""

_DD_PROBE_READ_JS = r"""return (() => {
    const ev = window.__ddprobe || [];
    return JSON.stringify({total: ev.length,
        first: ev.slice(0, 6), last: ev.slice(-6)});
})();"""


def _datadome_frame_offset(page):
    """DD iframe 在顶层 viewport 的偏移。

    实测 2026-09 interstitial(fast18 截图 0001_05b_landed_r1.png): 顶部
    "Please enable JS and disable any ad blocker" 行占 ~18px,iframe y 偏移
    非 0 —— 把手 iframe 内量得 y=287,viewport 实际 ~305。不加偏移
    pointerDown 按在把手上沿外,滑块永远不动(fast18 连按 7 次全空按)。"""
    raw = _js(page, r"""return (() => {
        let el = document.querySelector('iframe[src*="captcha-delivery"]');
        if (!el) {
            for (const f of document.querySelectorAll('iframe')) {
                const r = f.getBoundingClientRect();
                if (r.width > 200 && r.height > 100) { el = f; break; }
            }
        }
        if (!el) return null;
        const r = el.getBoundingClientRect();
        return JSON.stringify({x: Math.round(r.x), y: Math.round(r.y)});
    })();""", loaded=False)
    try:
        d = json.loads(raw) if isinstance(raw, str) else None
        if d and isinstance(d, dict):
            return d
    except Exception:
        pass
    return {"x": 0, "y": 0}


def _solve_datadome_slider(page, max_wait=30, interstitial=False):
    """拖 DataDome 滑块到轨道右端。成功(挑战 frame 消失/放行)返回 True。

    interstitial=True: 整页 interstitial 模式(实测 2026-09)。整页导航后 BiDi
    get_all_frames() 常常拿不到 captcha-delivery.com frame(延迟数秒~不复现),
    "none" 不能当放行——要用顶层 interstitial 标志(noscript 文案/RETRY 按钮)
    区分"挑战还在"与"真通过"。frame 迟迟不出现时点 RETRY 重出滑块。"""
    deadline = time.time() + max_wait
    attempts = 0
    retried = False
    while time.time() < deadline:
        state, s = _datadome_state(page)
        if state == "slider" and s and s.get("slider"):
            attempts += 1
            h = s["slider"]
            tr = s.get("track") or {}
            # 事件直接派发进 iframe context:坐标用 iframe 内部坐标(不加顶层偏移)
            _drag_fr = _datadome_frame(page)
            off = _datadome_frame_offset(page) or {"x": 0, "y": 0}
            hx = h["x"] + h["w"] // 2
            hy = h["y"] + h["h"] // 2
            # 用户指定: 长按滑块拖 240px 松开
            end_x = hx + 240
            print(f"  [datadome] === 滑块尝试 #{attempts} (ctx={'iframe' if _drag_fr is not None else 'top'}) ===")
            print(f"    [datadome] iframe 偏移 ({off['x']},{off['y']}), 把手 iframe 内坐标 ({hx},{hy}), 拖动 240px -> ({end_x},{hy})")
            # 探针:挂进 DD frame,拖完回读事件流(诊断 BiDi 事件是否真进 iframe)
            _fr = _datadome_frame(page)
            if _fr is not None and attempts == 1:
                try:
                    _pr = _js(_fr, _DD_PROBE_JS, loaded=False)
                    print(f"    [datadome] 探针安装: {_pr}")
                except Exception as _e:
                    print(f"    [datadome] 探针安装失败: {type(_e).__name__}")
            try:
                _bidi_slider_drag(page, hx, hy, end_x, hy, frame=_drag_fr)
            except Exception as e:
                print(f"    [datadome] drag err: {type(e).__name__}: {str(e)[:60]}")
            if _fr is not None:
                try:
                    _raw = _js(_fr, _DD_PROBE_READ_JS, loaded=False)
                    print(f"    [datadome] 探针事件流: {str(_raw)[:400]}")
                except Exception:
                    print("    [datadome] 探针回读失败(frame 已断)")
            print("    [datadome] 等待服务器判定 (6s)...")
            time.sleep(6)
        elif state == "restricted":
            print("  [datadome] restricted(出口被硬限,滑块都没有)")
            return False
        elif state == "blocked":
            print("  [datadome] blocked/audio 变体(IP 信誉差,不给滑块,拖不动)")
            return False
        elif state in ("none", "passed"):
            if not interstitial:
                # iframe 版挑战:frame 消失 = 放行
                print(f"  [datadome] 挑战消失/放行(拖了 {attempts} 次)")
                return True
            # 整页 interstitial:frame 还没进 BiDi 刲表或真放行了,看顶层标志
            body = ""
            try:
                body = _body_text(page).lower()
            except Exception:
                pass
            # 拖动过 = 挑战若消失即放行。fast22 实测:滑块通过后 GitHub 把你
            # 带回 /signup 表单页(URL 不变),之前要求 URL 离开 /signup 导致
            # 通过了还判超时,傻等 60s+乱点 RETRY。改为:只要拖过且 interstitial
            # 标志(顶层 noscript 文案)消失 = 放行,URL 是否回表单不管。
            if attempts > 0 and "please enable js and disable any ad blocker" not in body:
                print(f"  [datadome] interstitial 放行(拖了 {attempts} 次,回注册表单页)")
                return True
            if ("please enable js and disable any ad blocker" not in body
                    and not _retry_button_present(page)
                    and _page_url(page).split("?")[0] != "https://github.com/signup"):
                print(f"  [datadome] interstitial 放行(拖了 {attempts} 次)")
                return True
            # 诊断: none 时 frame 到底找没找到、run_js 返回了什么(只打一次)
            if attempts == 0 and not retried and time.time() > deadline - max_wait + 8:
                _diag_fr = _datadome_frame(page)
                _diag_raw = None
                if _diag_fr is not None:
                    try:
                        _diag_raw = _js(_diag_fr, "return document.title + '|' + (document.body ? document.body.innerText.slice(0, 60) : 'nobody');", loaded=False)
                    except Exception as e:
                        _diag_raw = f"<err {type(e).__name__}>"
                print(f"  [datadome][diag] frame={_diag_fr is not None} raw={str(_diag_raw)[:80]}")
            # 挑战还在:frame 迟迟不出现 -> 点 RETRY 重出滑块(只点一次)
            if not retried and time.time() > deadline - max_wait + 8 and _retry_button_present(page):
                _hit = _click_try_again(page)
                print(f"  [datadome] interstitial frame 未现,点 RETRY: {_hit or 'miss'}")
                retried = True
                time.sleep(5)
            time.sleep(2)
        else:
            time.sleep(2)
    print("  [datadome] 超时未通过")
    return False


def _goto_signup_via_home(page, max_wait=30):
    """人类路径:github.com 首页 -> 点右上角 Sign up。多数情况绕过 DataDome 门前挑战。
    返回 True=已到注册表单。"""
    try:
        cur = str(getattr(page, "url", "") or "")
        if "github.com" not in cur:
            page.get(WARMUP_HOME_URL)
            try:
                page.wait_loading(15)
            except Exception:
                pass
            time.sleep(2)
    except Exception as e:
        print(f"  [nav] home goto err: {str(e)[:60]}")
    deadline = time.time() + max_wait
    email_sel = "css:input#email"
    while time.time() < deadline:
        # 已在注册表单?
        try:
            el = page.ele(email_sel, timeout=2)
            if el and not isinstance(el, NoneElement):
                return True
        except Exception:
            pass
        # 点 Sign up 链接。ruyipage locator 对 a[href*='signup'] 命中不稳,
        # 直接 run_js 找到链接后 el.click() 最可靠(isTrusted 无所谓,导航点击)。
        clicked = False
        try:
            r = _js(page, """return (() => {
                const a = document.querySelector('a[href*="signup"]');
                if (!a) return 'miss';
                a.click();
                return a.getAttribute('href');
            })();""", loaded=False)
            if r and r != "miss":
                clicked = True
                print(f"  [nav] js-clicked Sign up link ({str(r)[:60]})")
        except Exception as e:
            print(f"  [nav] js click err: {str(e)[:60]}")
        if not clicked:
            # 兜底:ruyipage 原生点击(xpath 不带前导斜杠匹配)
            try:
                a = page.ele("xpath://a[contains(@href,'signup')]", timeout=2)
                if a and not isinstance(a, NoneElement):
                    a.click()
                    clicked = True
                    print("  [nav] clicked Sign up link (ruyipage xpath)")
            except Exception:
                pass
        if clicked:
            try:
                page.wait_loading(15)
            except Exception:
                pass
            # 点击后原地长轮询:表单出现=成功;DataDome 挑战交给调用方
            poll_deadline = time.time() + 25
            while time.time() < poll_deadline:
                try:
                    el = page.ele(email_sel, timeout=3)
                    if el and not isinstance(el, NoneElement):
                        return True
                except Exception:
                    pass
                st, _ = _datadome_state(page)
                if st in ("slider", "restricted"):
                    print(f"  [nav] 点击后进入 DataDome(state={st})")
                    return False
                time.sleep(1.5)
            return False
        time.sleep(1.5)
    try:
        body = _body_text(page).replace("\n", " | ")[:150]
        print(f"  [nav] Sign up 链接没点上,首页状态: url={_page_url(page)} body={body}")
    except Exception:
        print("  [nav] Sign up 链接没点上")
    # 诊断:DOM 里到底有没有 signup 链接
    try:
        diag = _js(page, """return (() => {
            const links = Array.from(document.querySelectorAll('a[href*="signup"]'));
            return JSON.stringify({n: links.length,
                hrefs: links.slice(0, 5).map(a => a.getAttribute('href') + ' | ' + a.innerText.slice(0, 30))});
        })();""", loaded=False)
        print(f"  [nav] signup links diag: {str(diag)[:300]}")
    except Exception as e:
        print(f"  [nav] diag err: {str(e)[:60]}")
    return False


def _click_label_button(ctx, label):
    """跨标签/角色点文本按钮。ruyipage 不支持 :has-text(),用 xpath contains 定位。
    返回元素或 None。覆盖 button/a/span[role=button]/input[value]。"""
    from ruyipage import NoneElement as _NE
    escaped = label.replace('"', '&quot;').replace("'", "&apos;")
    sel = ('xpath://*[self::button or self::a or @role="button" or self::span '
           'or self::input][contains(., "%s")]' % escaped)
    try:
        el = ctx.ele(sel)
        if el and not isinstance(el, _NE):
            return el
    except Exception:
        pass
    # 兜底: text= 简写(精确文本)
    try:
        el = ctx.ele("text=%s" % label)
        if el and not isinstance(el, _NE):
            return el
    except Exception:
        pass
    return None


def _click_create_account(page, max_wait=15, verbose=False):
    """点 'Create account' —— 只认 form 里 type=submit 的那个按钮。
    每次尝试打印按钮真实状态(found/disabled/visible),确认点击真落上去,
    不再用 text= 模糊匹配(会点到非 submit 元素,表现为"点了但没反应")。"""
    deadline = time.time() + max_wait
    _saw_disabled = False
    while time.time() < deadline:
        try:
            btns = page.eles('css:button[type="submit"]') or []
            for b in btns:
                try:
                    raw = b.run_js("""function(){
                        const r = this.getBoundingClientRect();
                        const cs = getComputedStyle(this);
                        return JSON.stringify({
                            text: (this.textContent || '').trim().slice(0, 40),
                            disabled: this.hasAttribute('disabled') || this.getAttribute('aria-disabled') === 'true',
                            visible: r.width > 0 && r.height > 0 && cs.display !== 'none' && cs.visibility !== 'hidden'
                        });
                    }""")
                    meta = {}
                    if raw:
                        try:
                            meta = json.loads(raw)
                        except Exception:
                            meta = {}
                    if not meta.get("text") or "create account" not in meta["text"].lower():
                        continue
                    if meta.get("disabled"):
                        _saw_disabled = True
                        if verbose:
                            print("  [form] 'Create account'(type=submit) disabled,等字段校验通过...")
                        break  # 找到 submit 但 disabled:不点(点了也没用),等下一轮
                    if not meta.get("visible"):
                        continue
                    b.click()
                    print(f"  [form] clicked 'Create account' (type=submit, text={meta['text']!r})")
                    return True
                except Exception:
                    continue
        except Exception:
            pass
        time.sleep(1)
    if _saw_disabled:
        print("  [form] 'Create account'(type=submit) 一直 disabled，某字段校验没过")
    else:
        print("  [form] 没找到 enabled 的 type=submit 'Create account' 按钮")
    return False


_SWW_JS = """return (() => {
    const errs = document.querySelectorAll('[class*="error"], [role="alert"], .flash-error, [class*="flash"]');
    for (const e of errs) {
        const t = (e.innerText || '').trim();
        if (/something went wrong/i.test(t)) return true;
    }
    return false;
})();"""

_RETRY_JS = """return (() => {
    const pats = ['try again', 'retry', '重试', '再试'];
    const btns = Array.from(document.querySelectorAll(
        'button, a, input[type=submit], input[type=button], input[type=image], [role=button]'));
    for (const b of btns) {
        const tx = (b.textContent || b.innerText || b.value || b.getAttribute('aria-label') || '').trim().toLowerCase();
        if (!tx) continue;
        for (const p of pats) {
            if (tx.indexOf(p) >= 0) { b.click(); return tx.slice(0, 40); }
        }
    }
    return '';
})();"""

# octocaptcha spinner 出现 = 验证流程真在跑(对比"iframe 常驻"可靠)
_SPINNER_JS = """return (() => {
    const sp = document.querySelector('.js-octocaptcha-spinner');
    if (!sp) return false;
    const r = sp.getBoundingClientRect();
    const cs = getComputedStyle(sp);
    return r.width > 0 && r.height > 0 && cs.display !== 'none';
})();"""


def _interstitial_detected(page):
    """DataDome 整页 interstitial 三层识别(实测 2026-09):
    1) iframe 版挑战(get_all_frames 里 captcha-delivery.com)
    2) 顶层 noscript 文案 "Please enable JS and disable any ad blocker"
       —— interstitial 顶层 body 只有这句(挑战内容全在 iframe 里),
       正常 GitHub 页面不会以它为 body 主体
    3) 顶层可点的 RETRY/Try again 按钮(滑块失败后的橙色按钮)
    返回 True=是 interstitial 页。"""
    try:
        if _datadome_state(page)[0] in ("slider", "restricted", "blocked"):
            return True
        body = _body_text(page).lower()
        if "please enable js and disable any ad blocker" in body:
            return True
        if _retry_button_present(page):
            return True
    except Exception:
        pass
    return False


_RETRY_SCAN_JS = """return (() => {
    const pats = ['try again', 'retry', '重试', '再试'];
    const btns = Array.from(document.querySelectorAll(
        'button, a, input[type=submit], input[type=button], input[type=image], [role=button]'));
    for (const b of btns) {
        const tx = (b.textContent || b.innerText || b.value || '').trim().toLowerCase();
        for (const p of pats) {
            if (tx && tx.indexOf(p) >= 0) {
                const r = b.getBoundingClientRect();
                if (r.width > 0 && r.height > 0) return true;
            }
        }
    }
    return false;
})();"""


def _retry_button_present(page):
    """有没有可点的 RETRY/Try again 按钮(不点,只查)。
    顶层没有时查 DataDome iframe(interstitial 的 RETRY 在 geo.captcha-delivery.com
    iframe 里,顶层 DOM 扫不到——实测 2026-09)。"""
    try:
        if bool(_js(page, _RETRY_SCAN_JS, loaded=False)):
            return True
    except Exception:
        pass
    try:
        fr = _datadome_frame(page)
        if fr is not None:
            return bool(_js(fr, _RETRY_SCAN_JS, loaded=False))
    except Exception:
        pass
    return False


def _sww_present(page):
    return bool(_js(page, _SWW_JS, loaded=False))


def _click_try_again(page):
    """点页面上的 Try again/重试(unlock 侧同款病灶)。返回点到的文本或 ''。
    顶层 miss 后同样落到 DataDome iframe 里点(interstitial RETRY 在 iframe 内)。"""
    for target in [page, _datadome_frame(page)]:
        if target is None:
            continue
        try:
            hit = str(_js(target, _RETRY_JS, loaded=False) or "")
            if hit:
                return hit
        except Exception:
            pass
    return ""


def _spinner_visible(page):
    return bool(_js(page, _SPINNER_JS, loaded=False))


def _trigger_verify(page, url_before=None, max_clicks=3):
    """用户流程: 点 submit -> 页面没跳转就再点,循环 3 次。
    跳转(url 变化)或滑块出现 = 本阶段结束,交给外层状态机。
    返回 'captcha'|'slider'|'restricted'|'spinner'|'navigated'|'none'。"""
    clicked = 0
    for attempt in range(max_clicks):
        if _click_create_account(page, max_wait=15, verbose=True):
            clicked += 1
        # 等 3s 看反应: 跳转/滑块/spinner/验证 任一出现即停
        deadline = time.time() + 3
        while time.time() < deadline:
            time.sleep(0.5)
            if url_before is not None and _page_url(page) != url_before:
                print(f"  [verify] 页面跳转 after {clicked} click(s)")
                return "navigated"
            if _detect_captcha(page, max_wait=0.5):
                print(f"  [verify] captcha after {clicked} click(s)")
                return "captcha"
            if _spinner_visible(page):
                print(f"  [verify] octocaptcha spinner after {clicked} click(s)")
                return "spinner"
            st, _ = _datadome_state(page)
            if st in ("slider", "restricted"):
                print(f"  [verify] datadome {st} after {clicked} click(s)")
                return st
        if clicked < max_clicks:
            print(f"  [verify] 没跳转,再点 ({clicked}/{max_clicks})")
            time.sleep(1)
    print(f"  [verify] submit loop done: actual_clicked={clicked}/{max_clicks}")
    return "none"


def _select_country(page, country="United States of America"):
    """选 Country/Region 下拉(GitHub 自定义 button[role=combobox] + listbox)。
    用 css: 明确前缀 + [role=combobox] 定位;ruyipage 不支持 Playwright 的 :has-text()。
    fast24 提速: 已经选对就直接跳过(重填轮次不必再点开下拉);固定 sleep 全部压到 0.3s。"""
    try:
        opener = page.ele('css:button[role="combobox"], [role="combobox"]')
        if not opener or isinstance(opener, NoneElement):
            return False
        # 已选对(重填轮次表单会保留旧值) → 不开下拉,省 2s+
        cur = _js(page, """return (() => {
            const b = document.querySelector('button[role="combobox"]');
            return b ? (b.textContent || '').trim() : '';
        })();""")
        if cur and country.lower() in cur.lower():
            return True
        opener.click()
        time.sleep(0.3)
        # 过滤输入(假如下拉有)
        filt_sel = 'css:input[placeholder*="Filter" i], input[aria-label*="Filter" i]'
        filt = page.ele(filt_sel)
        if filt and not isinstance(filt, NoneElement):
            _fill(page, filt_sel, country[:12], settle=0.3)
        # 用 xpath 精确定位 country 文本项(比 :has-text 稳)
        item = page.ele('xpath://*[self::button or self::li or self::div][@role="option" and contains(., "%s")]' % country)
        if not item or isinstance(item, NoneElement):
            item = page.ele('xpath://*[contains(., "%s")][@role="option"]' % country)
        if item and not isinstance(item, NoneElement):
            item.click()
            print(f"  [form] country selected: {country}")
            time.sleep(0.3)
            return True
    except Exception as e:
        print(f"  [form] select_country failed: {str(e)[:70]}")
    return False


def _uncheck_marketing(page):
    try:
        cb = page.ele("input#user_signup\\[marketing_consent\\], input[name='user_signup[marketing_consent]']")
        if cb and not isinstance(cb, NoneElement) and cb.run_js("function(){ return !!this.checked; }"):
            # 解除勾选
            _js(page, """const el = document.querySelector("input#user_signup\\\\[marketing_consent\\\\], input[name='user_signup[marketing_consent]']");
                if (el && el.checked) { el.click(); }""")
    except Exception:
        pass


def _get_code_api(email, refresh_token, client_id, max_wait=120, received_after=None):
    """纯 API 取 GitHub launch code(有 refresh_token 时)。
    received_after: 只收该时刻后的邮件(fast22 教训: 旧轮次的码 615757 被当成本轮的)。"""
    from common.mailbox import get_code_by_token
    code = get_code_by_token(
        email, refresh_token, client_id=client_id,
        sender_contains=GH_SENDER, subject_contains=GH_SUBJECT,
        code_regex=r"\b(\d{6,8})\b", max_wait=max_wait,
        received_after=received_after,
    )
    return code


def _outlook_fetch_code_tab(page, email, password, max_wait=180):
    """在独立 container tab 登录 Outlook 取 GitHub launch code。
    设计:新建 container tab(独立 userContext),避免污染 GitHub 主会话的
    登录态/cookie。复用 common/mailbox 的 Playwright 逻辑太重,这里精简适配。"""
    try:
        tab = page.new_container_tab(url="https://login.live.com/")
        time.sleep(4)
        # email
        ei = tab.ele("input[type='email'], input[name='loginfmt']")
        if ei and not isinstance(ei, NoneElement):
            _fill(tab, "input[type='email'], input[name='loginfmt']", email, "outlook-email", settle=0.6)
            # ruyipage 无 press():点 Next/#iNext 提交
            nxt = _click_label_button(tab, "Next") or tab.ele("css:#iNext") or tab.ele("css:input[type='submit']")
            if nxt and not isinstance(nxt, NoneElement):
                nxt.click()
            time.sleep(4)
        # password
        pi = tab.ele("input[type='password'], input[name='passwd']")
        if pi and not isinstance(pi, NoneElement):
            _fill(tab, "input[type='password'], input[name='passwd']", password, "outlook-pass", settle=0.6)
            nxt = _click_label_button(tab, "Sign in") or tab.ele("css:#idSIButton9") or tab.ele("css:input[type='submit']")
            if nxt and not isinstance(nxt, NoneElement):
                nxt.click()
            time.sleep(6)
        # 过隐私/passkey 中间页
        from ruyipage import NoneElement as _NE
        for _ in range(6):
            time.sleep(2)
            cur_url = str(getattr(tab, "url", "") or "")
            body = _body_text(tab).lower()[:400]
            clicked = False
            for label in ["Accept and continue", "Accept", "Agree and continue", "I agree", "Agree",
                          "接受并继续", "接受", "同意并继续", "同意", "Continue", "继续", "繼續", "Next", "OK",
                          "Yes", "是", "确认", "確認"]:
                b = _click_label_button(tab, label)
                if b:
                    try:
                        b.click(); clicked = True; break
                    except Exception:
                        pass
            if not clicked:
                for label in ["Skip", "跳过", "跳過", "Not now", "稍后", "稍後", "Maybe later", "暂时跳过"]:
                    b = _click_label_button(tab, label)
                    if b:
                        try:
                            b.click(); clicked = True; break
                        except Exception:
                            pass
            if "outlook" in cur_url or "mail.live" in cur_url:
                break
        # 进收件箱
        tab.get("https://outlook.live.com/mail/0/")
        time.sleep(6)
        # 轮询扫 launch code
        pat = re.compile(r"\b(\d{6,8})\b")
        start = time.time()
        while time.time() - start < max_wait:
            # 扫收件箱/垃圾邮件列表项
            found = _js(tab, """const items = [...document.querySelectorAll('[role="option"]')].slice(0, 8);
                const hits = [];
                for (const it of items) {
                    const t = it.textContent || '';
                    if (/github|launch code|verify|verification/i.test(t)) hits.push(t);
                }
                return hits.join('\\n');""")
            if found:
                m = pat.search(str(found))
                if m:
                    print(f"  [mail-tab] code from list: {m.group(1)}")
                    return m.group(1)
            time.sleep(6)
        return None
    except Exception as e:
        print(f"  [mail-tab] error: {str(e)[:80]}")
        return None
    finally:
        try:
            tab.close()
        except Exception:
            pass


def _get_gh_code(page, email, password, pool_entry, max_wait=180, received_after=None):
    """取 GitHub launch code,按可用源优先:
    broker / 纯 API(refresh_token) / container tab 浏览器. 返回 code 或 None."""
    # 1) broker
    if os.environ.get("MAILBOX_BROKER"):
        import asyncio
        from common.mailbox import fetch_from_broker
        try:
            return asyncio.run(fetch_from_broker(email, password, GH_SENDER, GH_SUBJECT,
                                                 r"\b(\d{6,8})\b", "code", max_wait))
        except Exception as e:
            print(f"  [broker] run error: {str(e)[:60]}")
    # 2) 纯 API
    if pool_entry and pool_entry.get("refresh_token"):
        code = _get_code_api(email, pool_entry["refresh_token"], pool_entry.get("client_id") or "",
                             max_wait, received_after=received_after)
        if code:
            return code
    # 3) container tab 浏览器
    return _outlook_fetch_code_tab(page, email, password, max_wait)


def _dump_state(page, tag=""):
    try:
        print(f"  --- state {tag} ---")
        print(f"  url: {_page_url(page)}")
        body = _body_text(page).replace("\n", " | ")[:200]
        print(f"  body: {body}")
        from ruyipage import NoneElement
        inputs = page.eles("input")
        for i in range(min(len(inputs or []), 8)):
            try:
                el = inputs[i]
                print(f"    input[{i}] type={el.attr('type')} name={el.attr('name')} id={el.attr('id')} autocomplete={el.attr('autocomplete')}")
            except Exception:
                pass
        _shot(page, tag or "state", getattr(page, "_idx", 1))
    except Exception as e:
        print(f"  dump_state error: {e}")


def _save_github_cookies(page, email, gh_password):
    """保存 GitHub 登录 cookies(适配 ruyipage get_cookies)。返回 key_cookie 或 None。"""
    import os, json
    from datetime import datetime
    pdir = os.path.join("cookies", PLATFORM)
    os.makedirs(pdir, exist_ok=True)
    cookies = page.get_cookies(True) or []
    key_val = None
    found_names = []
    serializable = []
    for c in cookies:
        raw = getattr(c, "raw", None)
        if not isinstance(raw, dict):
            raw = dict(getattr(c, "__dict__", {}) or {})
        name = str(raw.get("name") or "")
        if name in KEY_COOKIES:
            found_names.append(name)
            if key_val is None:
                key_val = raw.get("value")
        # 转成可序列化 dict
        serializable.append({"name": name, "value": raw.get("value", ""),
                             "domain": raw.get("domain", ""), "path": raw.get("path", "")})
    print(f"  [github] {len(cookies)} cookies, key cookies: {found_names}")
    # 防假阳性(fast22): _gh_sess 匿名会话也有 → 只有登录态才写 accounts.txt。
    # 判定: 有 'logged_in' 或 'user_session' cookie(GitHub 登录后才发)。
    cookie_names = {str(c_raw.get("name") or "") for c_raw in
                    [dict(getattr(c, "raw", None) or getattr(c, "__dict__", {}) or {}) for c in cookies]}
    if "logged_in" not in cookie_names and "user_session" not in cookie_names:
        print("  [github] 无 logged_in/user_session cookie —— 非登录态,不保存账号")
        return None
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    full_path = os.path.join(pdir, f"full_ruoyi_{ts}.json")
    with open(full_path, "w", encoding="utf-8") as f:
        json.dump(serializable, f, indent=2, ensure_ascii=False)
    print(f"  [github] full cookies saved: {full_path}")
    if key_val and email:
        acc_path = os.path.join(pdir, "accounts.txt")
        with open(acc_path, "a", encoding="utf-8") as f:
            f.write(f"{email}|{gh_password}|{key_val}\n")
        # 也记到 github_accounts/accounts.txt(与 outlook 目录对齐)
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        with open(os.path.join(OUTPUT_DIR, "accounts.txt"), "a", encoding="utf-8") as f:
            f.write(f"{email}|{gh_password}|{key_val}\n")
        print(f"  [github] account saved: {acc_path}")
    return key_val
# ── 主注册流程 ───────────────────────────────────────────────────

def register_github(opts, proxy_pool, idx, runtime=None):
    from ruyipage import FirefoxOptions, FirefoxPage, NoneElement

    _ruyi_pkg._install_shutdown_handlers()

    tag = f"[#{idx}][gh-ruoyi]"
    raw_timeout = getattr(opts, "timeout", None) or REGISTER_TIMEOUT
    try:
        timeout = max(0.0, float(raw_timeout))
    except Exception:
        timeout = float(REGISTER_TIMEOUT)

    is_headless = bool(getattr(opts, "headless", False))
    auto = bool(getattr(opts, "auto", False))
    keep = bool(getattr(opts, "keep", True))
    capture_har = False

    # UA 池
    user_agent = _ruyi_pkg._pick_user_agent(idx)
    print(f"  {tag} ua pool pick -> {_ruyi_pkg._mask_ua(user_agent)}")

    # FirefoxOptions 装配(复刻 outlook_ruoyi)
    tb = FirefoxOptions()
    tb.set_browser_path(RUOYI_FIREFOX_PATH)
    profile_dir = _ruyi_pkg._ruoyi_profile_dir(opts, idx)
    tb.set_profile(profile_dir)

    selected_browser_proxy = str(proxy_pool[0] or "").strip() if proxy_pool else ""
    _front = _ruyi_pkg.front_proxy_raw(getattr(opts, "front_proxy", None))
    if selected_browser_proxy:
        tb.set_proxy(_ruyi_pkg.front_relay_url(selected_browser_proxy, front=_front, tag=tag))
    else:
        print(f"  {tag} 没挂代理——直接本机出口", "WARN" if hasattr(print, '_warn') else "")

    _ruyi_pkg._apply_ruoyi_browser_ua(tb, tag, user_agent)
    _ruyi_pkg._apply_ruoyi_quiet_prefs(tb, tag)
    if is_headless:
        _ruyi_pkg._apply_ruoyi_headless_options(tb, tag, user_agent=user_agent)
        tb.headless(True)

    # warmup:启动 URL 一律用首页 github.com(种匿名挑战 cookie)。
    # 用户要求:不许直连 /signup,必须 首页 -> 点 Sign up(人类路径)。
    _warmup = True
    _launch_url = WARMUP_HOME_URL
    try:
        tb.set_argument(_launch_url)
    except Exception as _e:
        print(f"  {tag} set_argument failed: {_e}")
        _launch_url = ""

    print(f"  {tag} 启动 ruyipage Firefox: headless={is_headless} path={RUOYI_FIREFOX_PATH} profile={profile_dir} proxy={_ruyi_pkg.mask_ruoyi_proxy(selected_browser_proxy) if selected_browser_proxy else 'none'}")

    browser_page = None
    page = None
    email = password = None
    gh_password = ""
    success = False
    failure_reason = "failure"

    try:
        browser_page = FirefoxPage(tb)
        _ruyi_pkg._track_browser_page(browser_page)
        page = browser_page
        setattr(page, "_idx", idx)

        # geo 固定 US
        _ruyi_pkg._apply_ruoyi_geo_emulation(page, None, tag)

        # 资源拦截 image/font/media
        _ruyi_pkg._start_ruoyi_resource_blocking(page, tag)

        # 反检测注入(hardwareConcurrency=16)
        _ruyi_pkg._apply_ruoyi_headless_page_patches(page, tag, log_once=True, user_agent=user_agent)

        # warmup 首屏(已以首页启动则停留+种 cookie;否则自动导航首页)
        if _warmup:
            _ruyi_pkg.warmup(page, tag, home_url=WARMUP_HOME_URL, seed_cookies=None)

        # 导航到注册页。用户要求:只走 首页 -> 点 Sign up,不许直连 /signup。
        # 点失败就重试点击(回首页再点),不回退直连。
        if _warmup:
            if not _goto_signup_via_home(page):
                print("  [nav] 首页 Sign up 没到表单,回首页再点一次(不直连 /signup)")
                try:
                    page.get(WARMUP_HOME_URL)
                    try:
                        page.wait_loading(15)
                    except Exception:
                        pass
                    time.sleep(2)
                except Exception as _e:
                    print(f"  {tag} goto home failed: {_e}")
                if not _goto_signup_via_home(page, max_wait=40):
                    print("  [nav] 二次点击 Sign up 仍没到表单")
        else:
            # 理论不可达(_warmup 恒 True),保底也走人类路径
            if not _goto_signup_via_home(page):
                print("  [nav] 首页 Sign up 没到表单")
        try:
            page.wait_loading(20)
        except Exception:
            pass
        _ruyi_pkg._apply_ruoyi_headless_page_patches(page, tag, log_once=False, user_agent=user_agent)

        # 打印出口 IP(走代理,用请求探测)
        try:
            from common.ruyi.probe import _probe_proxy_identity
            ident = _probe_proxy_identity(proxy_pool, timeout=15) if proxy_pool else None
            if ident and ident.get("ip"):
                print(f"  {tag} current IP: {ident['ip']!r} country: {ident.get('country')!r}")
        except Exception:
            pass

        _shot(page, "01_after_load", idx)

        # ===== DataDome 门前挑战处理 =====
        # 直连 /signup 常被 DataDome 挡住(滑块/restricted)。挑战 iframe 出现慢,轮询等。
        # 顺序:没到表单 -> 回首页点 Sign up(人类路径,实测多数直过)
        #       -> 仍拦住就拖滑块 -> restricted 则换出口才解
        _has_email = False
        try:
            _em = page.ele("css:input#email", timeout=3)
            _has_email = bool(_em and not isinstance(_em, NoneElement))
        except Exception:
            pass
        if not _has_email:
            # 等挑战 iframe 出现(实测直连 DataDome 要 ~16s 才渲染,给 25s)
            _dd_state = "none"
            _dd_deadline = time.time() + 25
            while time.time() < _dd_deadline:
                _dd_state, _ = _datadome_state(page)
                if _dd_state != "none":
                    break
                try:
                    _em = page.ele("css:input#email", timeout=2)
                    if _em and not isinstance(_em, NoneElement):
                        _has_email = True
                        break
                except Exception:
                    pass
                time.sleep(1.5)
            if _has_email:
                print("  [nav] /signup 直接到表单(无挑战)")
            elif _dd_state != "none":
                print(f"  [datadome] /signup 被门前挑战拦截(state={_dd_state}),改走首页 Sign up 人类路径")
                _shot(page, "01b_datadome_gate", idx)
                if _goto_signup_via_home(page):
                    print("  [nav] 人类路径直接到注册表单")
                else:
                    # 点击后可能出滑块,拖它
                    _st, _ = _datadome_state(page)
                    if _st == "slider":
                        if _solve_datadome_slider(page, max_wait=40):
                            _shot(page, "01c_datadome_passed", idx)
                        else:
                            _shot(page, "01c_datadome_failed", idx)
                    elif _st == "restricted":
                        print("  [datadome] restricted:该出口 IP 被硬限,需换 proxy")
                        failure_reason = "datadome_restricted"
                        _shot(page, "01c_datadome_restricted", idx)
                        return (email, gh_password, False, failure_reason)
            else:
                print("  [nav] 无表单也无 DataDome 挑战,页面异常")
                _dump_state(page, "02_no_email")
                failure_reason = "no_email_input"

        # ===== 单页表单 =====
        email_sel = "css:input#email"
        if not page.ele(email_sel) or isinstance(page.ele(email_sel), NoneElement):
            print("  email input not found — GitHub 布局可能变了")
            _dump_state(page, "02_no_email")
            failure_reason = "no_email_input"
        else:
            # 邮箱:优先 opts.email,否则从池取
            if getattr(opts, "email", None):
                email = opts.email
                password = getattr(opts, "password", None) or ""
            else:
                accs = load_pool_accounts()
                if not accs:
                    print(f"  没有可用邮箱:{POOL_DIR} 为空(用 --email 指定)")
                    failure_reason = "no_pool"
                else:
                    email, password = random.choice(accs)
            if email:
                _fill(page, email_sel, email, "email")
                # 密码一次性生成并全程复用(表单填的 = 保存 cookie 登录用的)
                gh_password = gh_password or rand_password()
                _fill(page, "css:input#password", gh_password, "password")
                # 用户名(异步校验)
                username = rand_username()
                for _ in range(3):
                    _fill(page, "css:input#login", username, "username")
                    time.sleep(2.5)
                    body = _body_text(page).lower()
                    if any(k in body for k in ["unavailable", "already taken", "not available", "is already"]):
                        username = rand_username()
                        print(f"  [2] username taken, retry -> {username}")
                        continue
                    break
                _shot(page, "03_form_filled", idx)

                # 国家下拉
                _select_country(page, "United States of America")
                _uncheck_marketing(page)
                _shot(page, "04_before_submit", idx)

                # 等 Arkose enforcement 初始化
                print("  [3.5] settling for Arkose enforcement to init...")
                time.sleep(10)

                # ===== 用户流程状态机 =====
                # STEP A: 填表(上面已完成) -> 点 submit×3(没跳转就重点)
                # STEP B: 进滑块页 -> 找滑块 -> 长按左键拖 240px -> 松开
                # STEP C: 页面跳回填表页 -> 重填信息 -> 再 submit×3
                # STEP D: 验证码 -> 创建成功 -> 保存账号
                print("  [4] ===== STEP A: 点 submit(循环3次,没跳转就重点) =====")
                _otp_reached = False
                _signup_started_at = time.time()  # 验证码只收这之后的(防旧码)
                for _round in range(1, 7):
                    url_before = _page_url(page)
                    page._prev_submit_url = url_before
                    _click_result = _trigger_verify(page, url_before=url_before, max_clicks=3)
                    # 提交后观察落点(最长 25s)
                    _submit_deadline = time.time() + 25
                    _submit_state = "unknown"
                    while time.time() < _submit_deadline:
                        time.sleep(2)
                        _u = _page_url(page)
                        if _u != url_before:
                            try:
                                page.wait_loading(12)
                            except Exception:
                                pass
                            _submit_state = "navigated"
                            break
                        try:
                            _em2 = page.ele("css:input#email", timeout=1)
                            if _em2 and not isinstance(_em2, NoneElement):
                                _submit_state = "form_again"
                                break
                        except Exception:
                            pass
                        _st, _ = _datadome_state(page)
                        if _st in ("slider", "restricted"):
                            _submit_state = f"datadome_{_st}"
                            break
                        if _spinner_visible(page):
                            _submit_state = "spinner"
                            break
                        if _detect_captcha(page, max_wait=0.5):
                            _submit_state = "captcha"
                            break
                    print(f"  [4.{_round}] --- 落点判定: state={_submit_state} ---")
                    _shot(page, f"05_round{_round}", idx)

                    if _submit_state == "form_again":
                        # 页面没跳转,回在表单页 —— 用户流程: 重填信息,继续下一轮 submit×3
                        print(f"  [4.{_round}] STEP-C: 页面回到填表页,重填信息")
                        # 先看 SWW 横幅:点 Try again 重试 token 流(不重填,横幅说明
                        # 前端没发提交请求——octocaptcha token 流失败,与 unlock 侧同病灶)
                        if _sww_present(page):
                            _hit = _click_try_again(page)
                            print(f"  [4.{_round}] SWW 横幅, Try again={_hit or '没找到'}")
                            if _hit:
                                # 点到重试按钮:token 流重启,表单还留着,直接进下一轮
                                time.sleep(4)
                                continue
                            # 没有重试按钮:表单多半已被清空(fast23 实测 4.2-4.6 全空表单
                            # submit),必须落到下面的重填路径,否则空表单 submit 永远没码
                            print(f"  [4.{_round}] Try again 没找到,落到重填路径")
                        # 打回表单:先查是不是用户名被占用(提交后报 "Username xxx is not
                        # available" 这类校验),是则换名;表单可能被清空,重填全部
                        _body = _body_text(page)
                        _new_user = False
                        if re.search(r"username\s+\S+\s+is not available", _body, re.I) or \
                                any(k in _body.lower() for k in ["already taken", "unavailable", "not available", "is already", "username exists"]):
                            # 从报错里抠出被占的名字,换一个新名(加随机数字后缀降低再撞率)
                            _m = re.search(r"[Uu]sername\s+(\S+)\s+is not available", _body)
                            _old = _m.group(1) if _m else username
                            username = rand_username()
                            _new_user = True
                            print(f"  [4.{_round}] 用户名校验不过: '{_old}' 被占 -> 换新名 '{username}'")
                        print(f"  [4.{_round}] 重填表单 (new_username={_new_user}): email/password/username/country")
                        _fill(page, "css:input#email", email, "email")
                        _fill(page, "css:input#password", gh_password, "password")
                        if _new_user:
                            _fill(page, "css:input#login", username, "username")
                        else:
                            _fill(page, "css:input#login", username, "username")
                        _select_country(page, "United States of America")
                        _uncheck_marketing(page)
                        _shot(page, f"05r{_round}_refilled", idx)
                        print(f"  [4.{_round}] 重填完成,进下一轮 submit×3")
                        continue  # 重新进下一轮 submit

                    elif _submit_state.startswith("datadome_"):
                        if _submit_state == "datadome_restricted":
                            print("  [datadome] restricted:换出口才解")
                            failure_reason = "datadome_restricted"
                            _shot(page, "05_restricted", idx)
                            _round = 99  # 跳出
                            break
                        print("  ===== STEP B: 进入滑块页,找滑块位置,长按拖 240px =====")
                        if _solve_datadome_slider(page, max_wait=60, interstitial=True):
                            _shot(page, f"05r{_round}_slider_passed", idx)
                            # 滑块过了后表单被清空要求重填(实测)——回到表单分支
                            print(f"  [4.{_round}] 滑块通过,等待跳回填表页...")
                            time.sleep(3)
                            continue
                        else:
                            _shot(page, f"05r{_round}_slider_failed", idx)
                            continue

                    elif _submit_state == "captcha":
                        print("  ===== STEP-D: Arkose/octocaptcha 验证出现,打码 =====")
                        _shot(page, "06_CAPTCHA", idx)
                        if auto:
                            solved = solve_arkose(page, max_wait=200)
                            print(f"  [5] 验证结果: {'通过(注入token)' if solved else '打码失败'}")
                            if solved:
                                time.sleep(6)
                                _shot(page, "06b_after_solve", idx)
                                continue
                            _shot(page, "06b_solve_failed", idx)
                            continue
                        else:
                            print("  [explore] 停在验证步,窗口保留。")
                            failure_reason = "captcha_reached"
                            break

                    elif _submit_state == "spinner":
                        # octocaptcha spinner 在转 = 验证流程真在跑,等它出结果
                        # (token 到位→自动提交;失败→SWW 横幅;成功→跳转)
                        print("  [verify] octocaptcha spinner 转起,等结果(最长 40s)")
                        _sp_deadline = time.time() + 40
                        while time.time() < _sp_deadline:
                            time.sleep(2)
                            if not _spinner_visible(page):
                                break
                            if _sww_present(page):
                                break
                        if _sww_present(page):
                            print("  [verify] spinner 后出 SWW,下一轮 Try again")
                            continue
                        if not _spinner_visible(page):
                            # spinner 停了:看是跳转了还是回表单
                            _u2 = _page_url(page)
                            if _u2 != url_before:
                                _submit_state = "navigated"
                                _body = _body_text(page)
                                print(f"  [4.{_round}] landed: {_u2}")
                                if any(k in _body.lower() for k in ["verification code", "verify your", "enter the code", "check your email", "one-time code", "otp"]):
                                    _otp_reached = True
                                    print(f"  [4.{_round}] 到达接码页!")
                                    break
                                continue
                            print("  [verify] spinner 停但没跳转,回表单重提")
                            continue
                        print("  [verify] spinner 超时,回表单重提")
                        continue

                    elif _submit_state == "navigated":
                        # 整页跳转:可能是接码/验证邮箱页(成功路径)、DataDome interstitial
                        # (顶层整页挑战,实测 2026-09:滑块→失败变 RETRY)、或回表单
                        _body = _body_text(page)
                        print(f"  [4.{_round}] landed: {_page_url(page)}")
                        print(f"  [4.{_round}] body: {_body.replace(chr(10), ' | ')[:200]}")
                        _shot(page, f"05b_landed_r{_round}", idx)
                        # fast24 实测: account_verifications 页 body 是营销文案,
                        # 关键词匹配不到 — URL 本身就是判据
                        if "account_verifications" in _page_url(page) or \
                                any(k in _body.lower() for k in ["verification code", "verify your", "enter the code", "check your email", "one-time code", "otp"]):
                            _otp_reached = True
                            print(f"  ===== [4.{_round}] 到达接码页! STEP-D: 输入验证码 =====")
                            break
                        # DataDome interstitial: 整页挑战(三层识别,见 _interstitial_detected)
                        if _interstitial_detected(page):
                            print(f"  ===== [4.{_round}] STEP B: DataDome interstitial(整页挑战),找滑块拖 240px =====")
                            _dump_interstitial_html(page, f"dd_interstitial_r{_round}", idx)
                            if not _solve_datadome_slider(page, max_wait=60, interstitial=True):
                                # 拖失败/超时:挑战按钮变 RETRY,点它重出滑块
                                _retry_hit = _click_try_again(page)
                                print(f"  [4.{_round}] interstitial RETRY click: {_retry_hit or 'miss'}")
                                time.sleep(5)
                            continue
                        # 别的页(如 splash):看有没有表单,没有就继续轮询一轮
                        continue

                    else:
                        # unknown:超时啥也没发生。页面可能卡在 DataDome interstitial
                        # 的 RETRY 态(实测:滑块失败后橙色 RETRY,不点就永远停这)。
                        if _interstitial_detected(page):
                            print(f"  [4.{_round}] unknown 落点是 DataDome interstitial,处理")
                            _dump_interstitial_html(page, f"dd_interstitial_unknown_r{_round}", idx)
                            if not _solve_datadome_slider(page, max_wait=60, interstitial=True):
                                _retry_hit = _click_try_again(page)
                                print(f"  [4.{_round}] interstitial RETRY click: {_retry_hit or 'miss'}")
                                time.sleep(5)
                            continue
                        _retry_hit = _click_try_again(page)
                        if _retry_hit:
                            print(f"  [4.{_round}] unknown: 点到重试按钮 '{_retry_hit}'")
                            time.sleep(4)
                        continue

                if _round >= 99:
                    return (email, gh_password, False, failure_reason)
                if not _otp_reached and auto:
                    # 兜底:没明确到接码页,按旧逻辑探一次验证码
                    has_captcha = _detect_captcha(page)
                    if not has_captcha:
                        print("  [5] no captcha detected at this point")
                        _shot(page, "06_no_captcha", idx)

                if auto:
                    # ===== STEP-D: 输入验证码 =====
                    print("  ===== STEP-D: 取验证码并填入 =====")
                    time.sleep(4)
                    pool_entry = {}
                    try:
                        with open(next(f for f in sorted(glob.glob(os.path.join(POOL_DIR, "*.json")), reverse=True) if email in open(f, encoding='utf-8').read()), encoding="utf-8") as _f:
                            pool_entry = json.load(_f)
                    except Exception:
                        pass
                    code = _get_gh_code(page, email, password, pool_entry, max_wait=180,
                                        received_after=_signup_started_at)
                    if code:
                        print(f"  [STEP-D] got launch code: {code}")
                        _fill(page, "css:input[autocomplete='one-time-code']", code, "otp")
                        print("  [STEP-D] 验证码已填入,等待创建完成...")
                        time.sleep(4)
                        _shot(page, "09_after_code", idx)
                    else:
                        print("  [STEP-D] FAIL: 没收到验证码")

                    # ===== STEP-E: 确认创建成功,保存账号 =====
                    print("  ===== STEP-E: 确认创建成功,保存账号信息 =====")

                    # 跳主页确认落域。判定要防假阳性:没到接码页 + 没拿到 code 时,
                    # _gh_sess 只是匿名会话 cookie(实测 2026-09:整个 submit 循环卡
                    # DataDome interstitial,照样能存 cookie,但账号根本没建)。
                    try:
                        page.get("https://github.com/")
                        time.sleep(4)
                    except Exception:
                        pass
                    _shot(page, "10_final", idx)
                    # 登录态硬判定: logged_in cookie(GitHub 登录成功才种)
                    _has_login_cookie = False
                    try:
                        _lc = page.run_js("return document.cookie.indexOf('logged_in=') >= 0;")
                        _has_login_cookie = bool(_lc)
                    except Exception:
                        pass
                    if not _otp_reached and not code and not _has_login_cookie:
                        _final_body = _body_text(page).lower()
                        _logged_in = any(k in _final_body for k in ["create repository", "new repository", "your repositories", "start a project"]) or _has_login_cookie
                        if _logged_in:
                            print("  [STEP-E] 主页已登录态(免接码路径)")
                        else:
                            print("  [STEP-E] FAIL: 没到接码页也没拿到 code,主页非登录态 —— 判失败")
                            _save_github_cookies(page, email, gh_password)
                            failure_reason = failure_reason or "no_otp_no_login"
                            return (email, gh_password, False, failure_reason)

                    # 保存 cookie
                    key_val = _save_github_cookies(page, email, gh_password)
                    if key_val:
                        print("  ===== [STEP-E] OK: 账号创建成功,账号信息已保存 =====")
                        success = True
                        failure_reason = ""
                        return (email, gh_password, success, "success")
                    print("  [STEP-E] FAIL: 没拿到 session cookie")
                    failure_reason = "no_session_cookie"
                else:
                    failure_reason = "form_done"

    except Exception as e:
        import traceback
        print(f"  ERROR: {e}")
        traceback.print_exc()
        failure_reason = f"exception:{type(e).__name__}"
    finally:
        # 清理:归还代理池(避免烧穿 credits)+ 关浏览器 + 杀 firefox 进程树 + 删 profile
        try:
            if runtime is not None and proxy_pool:
                _ruyi_pkg.release_proxy_for_account(proxy_pool[0], runtime=runtime)
        except Exception:
            pass
        if browser_page is not None:
            try:
                _ruyi_pkg._quit_browser_page(browser_page, tag)
            except Exception:
                pass
        try:
            _ruyi_pkg._kill_ruoyi_firefox_by_profile(profile_dir, timeout=8.0)
        except Exception as e:
            print(f"  {tag} 清 firefox 进程树失败: {e}")
        try:
            _ruyi_pkg._cleanup_ruoyi_run_profile_dir(profile_dir)
        except Exception:
            pass
    return (email or "", gh_password, success, failure_reason)


# ── main ─────────────────────────────────────────────────────

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
    from common.ruyi import ConsumableProxyPool

    parser = argparse.ArgumentParser(description="GitHub Auto Register (ruoyi)")
    parser.add_argument("--email", default=None)
    parser.add_argument("--password", default=None)
    parser.add_argument("--auto", action="store_true")
    parser.add_argument("--no-keep", action="store_true")
    parser.add_argument("--headless", action="store_true")
    parser.add_argument("--count", "-n", type=int, default=1)
    parser.add_argument("--timeout", "-t", type=int, default=600)
    parser.add_argument("--proxy-file", default=None)
    parser.add_argument("--proxy-source", default=None)
    parser.add_argument("--proxy-url", default=None)
    parser.add_argument("--proxy", default=None, help="单条代理(不走池)")
    parser.add_argument("--front-proxy", default=None, help="前置代理(代理链)")
    parser.add_argument("--concurrency", "-c", type=int, default=1)
    parser.add_argument("--slot", type=int, default=0)
    args = parser.parse_args()

    global REGISTER_TIMEOUT
    REGISTER_TIMEOUT = args.timeout

    # 打造代理池
    proxy_pool = []
    pool = None
    if args.proxy:
        proxy_pool = [args.proxy]
    else:
        # ConsumableProxyPool:从 file/http 载入
        pdata = SimpleNamespace(
            proxy_file=args.proxy_file or PROXY_FILE,
            proxy_source=args.proxy_source or RUOYI_PROXY_SOURCE,
            proxy_url=args.proxy_url or PROXY_URL,
        )
        pool = ConsumableProxyPool(pdata)
        pool.start()
        picked = _ruyi_pkg.select_proxy_for_account(None, runtime=pool)
        proxy_pool = list(picked)
        print(f"  proxy pool take -> {_ruyi_pkg.mask_ruoyi_proxy(proxy_pool[0]) if proxy_pool else 'none'} remaining={pool.remaining()}")

    print("=" * 56)
    print(f"  GitHub Auto Register (ruoyi) auto={args.auto} headless={args.headless} "
          f"keep={not args.no_keep} count={args.count}")
    print("=" * 56)

    results = []
    for i in range(1, args.count + 1):
        print(f"\n----- 尝试 {i}/{args.count} -----")
        res = register_github(args, proxy_pool, i, runtime=pool)
        results.append(res)

    # 汇总
    ok = [r for r in results if r and r[2]]
    print(f"\n{'='*56}")
    print(f"  total={len(results)} ok={len(ok)} ")
    for r in results:
        if r:
            print(f"    email={r[0]} ok={r[2]} reason={r[3]}")
    print(f"{'='*56}")


if __name__ == "__main__":
    main()
