"""F 组(page 生命周期/quit/shutdown)+ 合并 common/ruyi_browser(启动栈+资源拦截)。

本模块合并原 common/ruyi_browser.py 全部内容 + register 的 F 组 page 生命周期函数。
不 import ruyipage(tb/page 全部 duck-type via getattr + _try_option_call)。
"""

import os
import re
import json
import signal
import threading
import urllib.parse
from datetime import datetime

from ._logging import log
from . import _state
from .ua import HEADLESS_USER_AGENT, HEADLESS_WINDOW_WIDTH, HEADLESS_WINDOW_HEIGHT, _DEFAULT_UA_POOL
from ._utils import _env_bool, _mask_ua


# ── 小工具(自包含复制,duck-type ruyipage) ───────────────────────

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


def _all_contexts(page):
    """page + 所有 iframe frame,用于在每个 context 注入反检测 JS。"""
    contexts = [page]
    try:
        contexts.extend(page.get_all_frames() or [])
    except Exception:
        pass
    return contexts


# ── F 组:page 生命周期 ────────────────────────────────────────────

def _track_browser_page(page):
    if page is None:
        return
    with _state._ACTIVE_BROWSER_LOCK:
        _state._ACTIVE_BROWSER_PAGES[id(page)] = page


def _untrack_browser_page(page):
    if page is None:
        return
    with _state._ACTIVE_BROWSER_LOCK:
        _state._ACTIVE_BROWSER_PAGES.pop(id(page), None)


def _close_tracked_browser_pages():
    with _state._ACTIVE_BROWSER_LOCK:
        pages = list(_state._ACTIVE_BROWSER_PAGES.values())
        _state._ACTIVE_BROWSER_PAGES.clear()
    for browser_page in pages:
        try:
            browser_page.quit(timeout=max(2.0, _state.BROWSER_QUIT_TIMEOUT), force=True)
            continue
        except Exception:
            pass
        try:
            browser_page.close()
        except Exception:
            pass


def _quit_browser_page(browser_page, tag="", timeout=None):
    if timeout is None:
        timeout = _state.BROWSER_QUIT_TIMEOUT
    if browser_page is None:
        return True
    done = threading.Event()

    def _worker():
        try:
            wait_timeout = max(2.0, float(timeout or 0.0))
            browser_page.quit(timeout=wait_timeout, force=True)
        except Exception:
            pass
        finally:
            done.set()

    thread = threading.Thread(target=_worker, daemon=True)
    thread.start()
    wait_timeout = max(2.0, float(timeout or 0.0))
    if done.wait(wait_timeout + 0.5):
        return True
    try:
        browser_page.close()
    except Exception:
        pass
    if tag:
        log(f"  {tag} browser quit timed out after {wait_timeout:.1f}s; continue closing", "WARN")
    return False


def _install_shutdown_handlers():
    if _state._SHUTDOWN_HANDLERS_INSTALLED:
        return

    def _handle_shutdown(signum, _frame):
        signame = str(signum)
        try:
            signame = signal.Signals(signum).name
        except Exception:
            pass
        try:
            log(f"收到 {signame},先关闭 ruyi 浏览器再退出", "WARN")
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
    _state._SHUTDOWN_HANDLERS_INSTALLED = True


def set_shutdown_handlers_installed(value=True):
    """显式标记 shutdown handlers 已装(unlock 等外部流程用)。"""
    _state._SHUTDOWN_HANDLERS_INSTALLED = bool(value)


# ── tb 阶段:启动前配 FirefoxOptions ─────────────────────────────

def _apply_ruoyi_browser_ua(tb, tag, user_agent):
    """有头/无头都尽量写 UA override,保证并发实例 UA 不一致。"""
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
    ok2, method2 = _try_option_call(tb, ("set_user_agent", "user_agent", "set_ua"), ua)
    if ok2:
        log(f"  {tag} browser ua override via {method2}: {_mask_ua(ua)}")
        return True
    return False


def _apply_ruoyi_quiet_prefs(tb, tag=""):
    """有头/无头都关掉会弹窗的 firefox 行为:会话恢复、崩溃报告、退出警告、默认浏览器检查。

    XPCOM 启动失败由进程树清理治本,这里灭其他边缘弹窗。另含降资源 pref:限制 content 进程数 +
    关磁盘缓存(临时 profile 一次性,缓存无意义且增 IO),多并发下每个 Firefox 实例少起若干子进程/少写盘,
    显著降内存与磁盘开销,对注册/PX 流程零影响。"""
    prefs = {
        "browser.sessionstore.enabled": False,
        "browser.sessionstore.resume_from_crash": False,
        "browser.warnOnQuit": False,
        "browser.shell.checkDefaultBrowser": False,
        "toolkit.crashreporter.enabled": False,
        "dom.ipc.crashreporter.enabled": False,
        # ---- 降资源(并发下每个实例少起子进程/少写盘) ----
        "dom.ipc.processCount": 1,
        "browser.cache.disk.enable": False,
        "media.hardware-video-decoding.enabled": False,
        "plugins.update.notifyUser": False,
    }
    applied = []
    for key, value in prefs.items():
        ok, method = _try_option_call(
            tb,
            ("set_preference", "set_pref", "set_prefs", "set_option"),
            key,
            value,
        )
        if ok:
            applied.append(key)
    return bool(applied)


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


def build_browser_options(tb, *, tag, user_agent, is_headless, direct_url=None):
    """配 FirefoxOptions:UA + quiet_prefs + (无头)headless_options + (可选)set_argument(url)。

    代理/profile 由调用方先行设置(tb.set_browser_path / set_profile / set_per_tab_proxies),
    本函数只负责 UA + prefs + headless + 直接启动 URL。
    """
    _apply_ruoyi_browser_ua(tb, tag, user_agent)
    _apply_ruoyi_quiet_prefs(tb, tag)
    if is_headless:
        _apply_ruoyi_headless_options(tb, tag, user_agent=user_agent)
        tb.headless(True)
    if direct_url:
        try:
            tb.set_argument(direct_url)
        except Exception as exc:
            log(f"  {tag} set_argument({direct_url}) 失败,回退空白页启动: {exc}", "WARN")


# ── page 阶段:启动后反检测注入 + 资源拦截 ───────────────────────

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
    """反检测 patch JS:visibility/webdriver/UA/languages/screen/hardwareConcurrency 等。

    hardwareConcurrency=16 是 PX 放行关键(本机 20 触发二次审查)。"""
    ua = str(user_agent or HEADLESS_USER_AGENT or "").strip() or HEADLESS_USER_AGENT
    ua_js = json.dumps(ua)
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
  delete window.__playwright;
  delete window.__pwInitScripts;
  for (const k of Object.getOwnPropertyNames(window)) {{
    if (/^cdc_|webdriver|selenium|driver/i.test(k)) {{
      try {{ delete window[k]; }} catch(e) {{}}
    }}
  }}
}} catch(e) {{}}
// hardwareConcurrency 伪装:本机 20 线程少见,PX 会据此触发 px-cloud.net/v
// 二次审查导致按压不放行;改 16 对齐正常机器(实测 PX 秒过)
try {{
  Object.defineProperty(navigator, 'hardwareConcurrency', {{get: () => 16, configurable: true}});
}} catch(e) {{}}
return true;
"""


# 兼容旧引用(无 UA 参数时的默认 patch)
RUOYI_HEADLESS_PATCH_JS = _build_headless_patch_js(HEADLESS_USER_AGENT)


def _ensure_ruoyi_headless_preload(page, tag=None, log_once=False, user_agent=None):
    add_script = getattr(page, "add_preload_script", None)
    if not callable(add_script):
        # preload 是 patch_js 的兜底注入渠道;主路径 run_js_loaded 已生效,
        # preload 缺失不影响,降级为 DEBUG 不刷 WARN。
        if log_once:
            log(f"  {tag} ruoyi headless preload API not available(patches via run_js)", "DEBUG")
        return False
    ua = str(user_agent or HEADLESS_USER_AGENT or "").strip() or HEADLESS_USER_AGENT
    if getattr(page, "_ruoyi_headless_preload_ready", False) and getattr(page, "_ruoyi_headless_preload_ua", "") == ua:
        return True
    preload_js = f"() => {{{_build_headless_patch_js(ua)}}}"
    try:
        add_script(preload_js)
        setattr(page, "_ruoyi_headless_preload_ready", True)
        setattr(page, "_ruoyi_headless_preload_ua", ua)
        if log_once:
            log(f"  {tag} ruoyi headless preload ready ua={_mask_ua(ua)}")
        return True
    except Exception as exc:
        # 155 内核下 add_preload_script 走 BiDi 需 --remote-allow-system-access,未加会抛
        # "System access is required"。但 patch_js 主路径(run_js_loaded IIFE 注入)已生效,
        # preload 非必需;降级为 DEBUG 静默处理,不刷 WARN。
        if log_once:
            log(f"  {tag} ruoyi headless preload skipped({type(exc).__name__}); patches via run_js", "DEBUG")
        return False


def _apply_ruoyi_headless_page_patches(page, tag=None, log_once=False, user_agent=None, wait_loaded=True):
    """启动后/导航后/每轮主循环注入反检测:emulation + preload + patch_js(hardwareConcurrency=16)。

    有头/无头都调。PX 放行依赖 hardwareConcurrency=16。
    wait_loaded:True(默认)用 run_js_loaded 等 doc_loaded 再注入;
      False(unlock 首屏用)直接 run_js 立即注入 —— patch 只 define 属性,不需等加载,
      避免登录页 doc_loaded 卡 10s 阻塞输入邮箱。"""
    ua = str(user_agent or HEADLESS_USER_AGENT or "").strip() or HEADLESS_USER_AGENT
    _apply_ruoyi_headless_emulation(page, tag, log_once=False, user_agent=ua)
    _ensure_ruoyi_headless_preload(page, tag, log_once=log_once, user_agent=ua)
    patch_js = _build_headless_patch_js(ua)
    # patch_js 以 return 结尾(函数体),需包成 IIFE 才能 run_js 直接执行
    inject_js = f"(() => {{ {patch_js} }})()"
    runner = "run_js_loaded" if wait_loaded else "run_js"
    ok_count = 0
    for ctx in _all_contexts(page):
        try:
            fn = getattr(ctx, runner, None) or ctx.run_js_loaded
            if fn(inject_js):
                ok_count += 1
        except Exception:
            continue
    if log_once:
        log(f"  {tag} ruoyi headless page patches applied to {ok_count} context(s) ua={_mask_ua(ua)}", "DEBUG")
    return ok_count


# ── 资源拦截 ─────────────────────────────────────────────────────

_RUOYI_RESOURCE_BLOCK_KINDS = ("image", "font", "media")

_RUOYI_RESOURCE_ALLOW_HOST_HINTS = (
    "fpt.live.com",
    "hsprotect.net",
    "px-cloud.net",
    "px-cdn.net",
    "client.px-cloud.net",
)

# 纯遥测/分析域名黑名单(HAR 实测):OneCollector 遥测等,与注册功能/PX 放行无关,始终屏蔽减负。
# 不含 hsprotect/px-cloud(PX 行为遥测,放行判定需要)与 logincdn(注册表单 JS + 按钮图)。
_RUOYI_TELEMETRY_HOSTS = (
    "browser.events.data.microsoft.com",  # OneCollector 遥测(9 个空响应)
    "arc.msn.com",                         # 微软分析
    "vortex.data.microsoft.com",           # 微软遥测
    "telemetry.microsoft.com",
    "events.data.microsoft.com",
)

_RUOYI_RESOURCE_BLOCK_EXTS = (
    ".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".ico", ".bmp", ".avif",
    ".woff", ".woff2", ".ttf", ".otf", ".eot",
    ".mp4", ".webm", ".mp3", ".wav", ".ogg", ".m4a",
)


def _ruoyi_should_block_resource_request(req):
    url = str(getattr(req, "url", "") or "")
    low = url.lower()
    if any(host in low for host in _RUOYI_RESOURCE_ALLOW_HOST_HINTS):
        return False
    # 遥测黑名单:纯遥测域名与注册/PX 无关,始终屏蔽(不依赖 block_resources)。
    try:
        host = (urllib.parse.urlsplit(low).netloc or "").lower()
        if host and any(host == th or host.endswith("." + th) for th in _RUOYI_TELEMETRY_HOSTS):
            return True
    except Exception:
        pass
    headers = getattr(req, "headers", None) or {}
    dest = str(headers.get("Sec-Fetch-Dest") or headers.get("sec-fetch-dest") or "").strip().lower()
    if dest in _RUOYI_RESOURCE_BLOCK_KINDS:
        return True
    accept = str(headers.get("Accept") or headers.get("accept") or "").strip().lower()
    if "image/" in accept or "font/" in accept or "audio/" in accept or "video/" in accept:
        return True
    path = urllib.parse.urlsplit(low).path or ""
    return any(path.endswith(ext) for ext in _RUOYI_RESOURCE_BLOCK_EXTS)


def _start_ruoyi_resource_blocking(page, tag=None):
    if page is None:
        return False
    interceptor = getattr(page, "intercept", None)
    if interceptor is None or not callable(getattr(interceptor, "start_requests", None)):
        return False
    if getattr(page, "_ruoyi_resource_blocking_active", False):
        return True

    def _handler(req):
        # 整体包 try/except: fail()/continue_request() 对"已不存在"的请求会抛
        # no such request(浏览器内部已取消/重定向的请求),被 ruyipage 拦截层
        # 捕获后会打 "拦截回调异常" warning。这类异常无害(请求本就要丢弃/已结束),
        # 静默吞掉,避免刷屏且不阻断其他请求的处理。
        try:
            if _ruoyi_should_block_resource_request(req):
                req.fail()
                return
            req.continue_request()
        except Exception:
            pass

    try:
        interceptor.start_requests(_handler)
        setattr(page, "_ruoyi_resource_blocking_active", True)
        if tag:
            log(f"  {tag} ruoyi resource blocking active: image/font/media")
        return True
    except Exception as exc:
        if tag:
            log(f"  {tag} ruoyi resource blocking start failed: {type(exc).__name__}: {exc}", "WARN")
        return False


def after_launch(page, *, tag, user_agent, block_resources=False, wait_loaded=True):
    """启动后:资源拦截(可选)+ 反检测注入(必做,含 hardwareConcurrency=16)。

    wait_loaded:True(默认)patch 等 doc_loaded;False(unlock 首屏)立即注入,
      避免 doc_loaded 卡 10s 阻塞输入邮箱。"""
    if block_resources:
        _start_ruoyi_resource_blocking(page, tag)
    _apply_ruoyi_headless_page_patches(page, tag, log_once=True, user_agent=user_agent,
                                       wait_loaded=wait_loaded)
