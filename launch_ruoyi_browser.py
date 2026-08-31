#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
launch_ruoyi_browser.py —— 手动唤起 ruyipage 定制 Firefox（调试/挂机用）

用途：
  - 命令行或 WebUI 唤起一台/多台 ruyipage 定制 Firefox，便于人工调试、登录、挂机。
  - 单代理：--proxy socks5://user:pass@host:port —— 整个浏览器走该代理（FirefoxPage 级 set_proxy）。
  - 代理列表轮询：--proxy-file / --proxy-url —— 用 FirefoxOptions.set_per_tab_proxies，
    每个 container tab 分配一条不同代理（ruyipage 内核按 userContextId 轮换代理）。
  - 代理链（前置代理）：--front-proxy socks5://127.0.0.1:10808 —— 当上游代理需“外网 IP 才能连”
    （如 1024proxy 这类做 IP 白名单的机房代理）时，本地起一个 SOCKS5 中继，
    Firefox -> 本地中继 -> 前置(10808) -> 上游代理 -> 目标站点。
  - 默认 keep 浏览器不关（--no-keep 才自动关），Ctrl-C 优雅退出。

示例：
  # 单代理，打开 outlook 收件箱
  python launch_ruoyi_browser.py --url https://outlook.live.com --proxy socks5://u:p@1.2.3.4:1080

  # 1024proxy 这类需外网 IP 的代理：经本地 10808 再连上游
  python launch_ruoyi_browser.py --proxy http://user:pass@us.1024proxy.io:3000 --front-proxy socks5://127.0.0.1:10808

  # 10 条代理各开一个 tab，轮询出口
  python launch_ruoyi_browser.py --proxy-file proxies_outlook.txt --tabs 10 --url https://ipinfo.io

  # HTTP 拉取代理列表后轮询
  python launch_ruoyi_browser.py --proxy-source http --proxy-url https://example.com/list.txt --tabs 5

  # 无代理直连
  python launch_ruoyi_browser.py --url https://example.com
"""
from __future__ import annotations

import argparse
import os
import sys
import time

# 项目根 = 本文件所在目录
ROOT = os.path.dirname(os.path.abspath(__file__))


def _load_dotenv():
    """吃项目根 .env（注册脚本同款），不覆盖已有环境变量。"""
    try:
        path = os.path.join(ROOT, ".env")
        if not os.path.isfile(path):
            return
        for line in open(path, encoding="utf-8"):
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


_load_dotenv()

# 复用 ruoyi 注册脚本里成熟的工具：代理归一化、profile 目录、UA、headless 调优、shutdown 处理。
# 这些都是纯函数/小工具，import 不触发注册流程；register_outlook() 不会被调用。
# FirefoxOptions/FirefoxPage 直接从 ruyipage 导入——register_outlook_ruoyi 是在函数内
# 局部 import 的，模块级未暴露这两个符号。
import register_outlook_ruoyi as rr  # noqa: E402
from ruyipage import FirefoxOptions, FirefoxPage  # noqa: E402

from common.ruyi import (  # noqa: E402
    RUOYI_FIREFOX_PATH,
    _browser_model_name,
    _env_bool,
    _mask_ua,
    _proxy_url_to_ruoyi,
    mask_ruoyi_proxy,
    parse_proxy_pool,
    fetch_proxy_list_http,
)


def _eprint(*a, **k):
    print(*a, **k, file=sys.stderr)


def _log(msg, level="INFO"):
    tag = level or "INFO"
    print(f"[{_now()}] [{tag}] {msg}")


def _now():
    try:
        from datetime import datetime
        return datetime.now().strftime("%H:%M:%S")
    except Exception:
        return ""


# ============================================================ 代理链中继
# 实现已抽到 common/ruyi/chain.py（register/unlock 同款复用），此处仅导入。
from common.ruyi import ChainRelayHub, parse_front_proxy as _parse_proxy_url  # noqa: E402,F401


# ============================================================ 代理收集
def _load_proxies(args):
    """按 --proxy / --proxy-file / --proxy-source+--proxy-url 收上游代理列表。
    返回 ruyipage 归一化后的代理字符串列表（已去空/去非法）。"""
    explicit = str(getattr(args, "proxy", "") or "").strip()
    if explicit:
        norm = _proxy_url_to_ruoyi(explicit)
        if not norm:
            raise SystemExit(f"[ERR] 非法代理 --proxy: {explicit}")
        return [norm]

    want_rotate = bool(getattr(args, "proxy_file", "") or getattr(args, "proxy_url", "")
                       or getattr(args, "proxy_source", "") == "http")
    if not want_rotate:
        return []  # 无代理直连

    source = str(getattr(args, "proxy_source", "") or "file").strip().lower()
    if source == "http":
        url = str(getattr(args, "proxy_url", "") or "").strip()
        if not url:
            raise SystemExit("[ERR] --proxy-source=http 必须配 --proxy-url")
        return fetch_proxy_list_http(url)

    proxy_file = getattr(args, "proxy_file", "") or "proxies_outlook.txt"
    batch = parse_proxy_pool(proxy_file)
    if not batch and str(getattr(args, "proxy_url", "") or "").strip():
        _log(f"本地代理文件 {proxy_file} 为空，回退 --proxy-url HTTP 列表", "WARN")
        try:
            batch = fetch_proxy_list_http(str(args.proxy_url).strip())
        except Exception as exc:
            _log(f"回退 HTTP 代理列表失败: {type(exc).__name__}: {exc}", "WARN")
            batch = []
    return batch


def _per_tab_format(proxy_str):
    """set_per_tab_proxies 要求 host:port:user:pass 或 socks5://host:port:user:pass。
    把 ruoyi 归一化的 socks5://user:pass@host:port 转成 socks5://host:port:user:pass。"""
    parsed = rr._parse_ruoyi_proxy(proxy_str)
    if not parsed or not parsed.get("host") or not parsed.get("port"):
        raise ValueError(f"无法解析代理: {proxy_str}")
    scheme = (parsed.get("scheme") or "socks5").lower()
    user = parsed.get("username") or ""
    pwd = parsed.get("password") or ""
    if not user or not pwd:
        raise ValueError(
            f"per-tab 代理必须带账号密码(内核按 userContextId 认证): {proxy_str}"
        )
    if ":" in user or ":" in pwd:
        raise ValueError(f"per-tab 代理 user/pass 不能含冒号: {proxy_str}")
    return f"{scheme}://{parsed['host']}:{parsed['port']}:{user}:{pwd}"


def _probe_exit_ip(page, timeout=20):
    """best-effort 探测当前 tab 出口 IP（经代理出网）。失败返回空串。"""
    try:
        page.get("https://ipinfo.io/json", timeout=timeout)
    except Exception:
        return ""
    try:
        body = page.run_js_loaded("return document.body.innerText") or ""
    except Exception:
        return ""
    import json as _json
    try:
        d = _json.loads(body.strip())
        ip = d.get("ip") or ""
        country = d.get("country") or ""
        city = d.get("city") or ""
        return f"{ip} {country}/{city}".strip()
    except Exception:
        import re as _re
        m = _re.search(r"\b\d{1,3}(?:\.\d{1,3}){3}\b", body)
        return m.group(0) if m else ""


def _open_target_url(page, url):
    """打开目标 URL；url 为空则停在 about:blank（便于人工输入）。"""
    if not url:
        return
    try:
        page.get(url)
        _log(f"已打开: {url}", "OK")
    except Exception as exc:
        _log(f"打开 {url} 失败(可手动在地址栏输入): {type(exc).__name__}: {exc}", "WARN")


# ============================================================ 构造 FirefoxOptions
def build_options(args, *, single_proxy, per_tab_proxies, user_agent):
    """构造 FirefoxOptions。
       single_proxy: 整浏览器走该代理(字符串或 None)。
       per_tab_proxies: per-tab 轮询列表(已是 socks5://host:port:u:p 形式)或 []。
       两者都空 = 直连。"""
    tb = FirefoxOptions()
    tb.set_browser_path(RUOYI_FIREFOX_PATH)

    profile_dir = rr._ruoyi_profile_dir(args, idx=int(getattr(args, "tab_idx_base", 1) or 1))
    tb.set_profile(profile_dir)
    _log(f"profile: {profile_dir}")

    if single_proxy:
        tb.set_proxy(single_proxy)
        _log(f"单代理(整浏览器): {mask_ruoyi_proxy(single_proxy)}", "OK")
    elif per_tab_proxies:
        exhausted = str(getattr(args, "proxy_exhausted", "wrap") or "wrap").strip().lower()
        tb.set_per_tab_proxies(list(per_tab_proxies), exhausted=exhausted)
        _log(f"per-tab 代理轮询: {len(per_tab_proxies)} 条 exhausted={exhausted}", "OK")
    else:
        _log("无代理 —— 直接本机出口", "WARN")

    if user_agent:
        try:
            rr._apply_ruoyi_browser_ua(tb, "[launch]", user_agent)
        except Exception as exc:
            _log(f"UA override 失败(忽略): {type(exc).__name__}: {exc}", "WARN")

    if getattr(args, "headless", False):
        try:
            rr._apply_ruoyi_headless_options(tb, "[launch]", user_agent=user_agent)
            tb.headless(True)
            _log("headless 模式", "OK")
        except Exception as exc:
            _log(f"headless 调优失败(忽略): {type(exc).__name__}: {exc}", "WARN")

    return tb, profile_dir


def main():
    ap = argparse.ArgumentParser(
        description="唤起 ruyipage 定制 Firefox（单代理 / per-tab 代理列表轮询 / 代理链）"
    )
    ap.add_argument("--url", default="",
                    help="打开的页面 URL(留空=about:blank，便于人工输入)")
    # 代理：--proxy 单条优先；否则走列表轮询
    ap.add_argument("--proxy", default=os.environ.get("LAUNCH_UPSTREAM_PROXY", ""),
                    help="上游代理(整浏览器走它)：socks5://user:pass@host:port 或 http://user:pass@host:port；"
                         "留空读 .env LAUNCH_UPSTREAM_PROXY")
    ap.add_argument("--front-proxy",
                    default=os.environ.get("LAUNCH_FRONT_PROXY", ""),
                    help="前置代理(本机可达)：socks5://127.0.0.1:10808 或 http://127.0.0.1:7897。"
                         "上游代理需外网 IP 才能连时用它链一跳；留空读 .env LAUNCH_FRONT_PROXY")
    ap.add_argument("--proxy-source", default="file", choices=["file", "http"],
                    help="列表来源：file=本地文件；http=HTTP GET 拉 txt 列表")
    ap.add_argument("--proxy-url", default=os.environ.get("OUTLOOK_PROXY_URL", ""),
                    help="HTTP GET 代理列表地址(配合 --proxy-source=http)")
    ap.add_argument("--proxy-file", default=os.environ.get("OUTLOOK_PROXY_FILE", "proxies_outlook.txt"),
                    help="代理池文件(每行一条；与 --proxy-source=file 配合)")
    ap.add_argument("--proxy-exhausted", default="wrap",
                    choices=["wrap", "direct", "none", "stop"],
                    help="per-tab 代理耗尽策略：wrap 轮回复用 / direct 回退直连 / none 不回退 / stop 停止开新 tab")
    ap.add_argument("--tabs", type=int, default=1,
                    help="per-tab 模式下要开的 container tab 数(单代理模式忽略；默认 1)")
    ap.add_argument("--probe", action="store_true",
                    help="每个 tab 打开后探测出口 IP(ipinfo.io)并打印(便于核对代理生效)")
    ap.add_argument("--headless", action="store_true", help="无头模式(调试一般不用)")
    ap.add_argument("--no-keep", action="store_true",
                    help="脚本退出时自动关浏览器(默认 keep：退出后浏览器仍开着，便于人工接管)")
    ap.add_argument("--log-level", default=os.environ.get("OUTLOOK_LOG_LEVEL", "INFO"),
                    choices=["DEBUG", "INFO", "WARN", "PROD", "ERR"],
                    help="日志等级")
    args = ap.parse_args()

    rr.set_log_level(args.log_level)
    rr._install_shutdown_handlers()

    if not os.path.isfile(RUOYI_FIREFOX_PATH):
        _eprint(f"[ERR] 定制 Firefox 内核不存在: {RUOYI_FIREFOX_PATH}")
        _eprint("请先运行: .venv\\Scripts\\python.exe -m ruyipage install")
        sys.exit(1)

    upstream_proxies = _load_proxies(args)
    _log(f"上游代理收集完成: {len(upstream_proxies)} 条" + (
        f" 来源={'http' if args.proxy_source == 'http' and not args.proxy else 'file/explicit'}"
        if upstream_proxies else " (直连)"))

    # 前置代理 -> 本地 SOCKS5 中继链
    hub = None
    front_raw = str(getattr(args, "front_proxy", "") or "").strip()
    if front_raw and upstream_proxies:
        try:
            front = _parse_proxy_url(front_raw)
        except ValueError as exc:
            _eprint(f"[ERR] 非法 --front-proxy: {exc}")
            sys.exit(1)
        hub = ChainRelayHub(front)
        hub.start()
        _log(f"代理链中继已启动，前置代理: {mask_ruoyi_proxy(front_raw)}", "OK")
        # 每条上游代理 -> 一个本地中继端口
        single_proxy = None
        per_tab_proxies = []
        for p in upstream_proxies:
            up = _parse_proxy_url(p)
            port = hub.port_for(up)
            _log(f"  中继 {up['host']}:{up['port']} -> 127.0.0.1:{port}", "OK")
            if len(upstream_proxies) == 1:
                # 单代理：Firefox 走本地中继(SOCKS5 无认证)
                single_proxy = f"socks5://127.0.0.1:{port}"
            else:
                # per-tab：ruyipage 内核要求每条带账号密码(无冒号)，中继对任意账号放行
                per_tab_proxies.append(f"socks5://127.0.0.1:{port}:u:p")
        if not single_proxy and not per_tab_proxies:
            _eprint("[ERR] 代理链归一化后为空")
            sys.exit(1)
    elif front_raw and not upstream_proxies:
        _log("--front-proxy 已配但无上游代理，前置代理忽略", "WARN")
        single_proxy = None
        per_tab_proxies = []
    else:
        # 无前置代理：直接用上游
        if len(upstream_proxies) == 1:
            single_proxy = upstream_proxies[0]
            per_tab_proxies = []
        elif len(upstream_proxies) > 1:
            single_proxy = None
            per_tab_proxies = []
            for p in upstream_proxies:
                try:
                    per_tab_proxies.append(_per_tab_format(p))
                except ValueError as exc:
                    _log(f"跳过不合规 per-tab 代理: {exc}", "WARN")
            if not per_tab_proxies:
                _eprint("[ERR] per-tab 代理列表归一化后为空(检查格式/账号密码)")
                sys.exit(1)
        else:
            single_proxy = None
            per_tab_proxies = []

    # per-tab 模式按 --tabs 截断/补齐；单代理/直连固定 1 个 tab
    rotate_mode = bool(per_tab_proxies)
    tabs_wanted = max(1, int(args.tabs or 1)) if rotate_mode else 1
    if rotate_mode and tabs_wanted > len(per_tab_proxies):
        _log(f"--tabs {tabs_wanted} > 代理数 {len(per_tab_proxies)}，按 exhausted={args.proxy_exhausted} 复用/回退",
             "WARN")

    user_agent = rr._pick_user_agent(1)
    _log(f"ua pool pick -> {_mask_ua(user_agent)}")

    tb, profile_dir = build_options(args, single_proxy=single_proxy,
                                    per_tab_proxies=per_tab_proxies, user_agent=user_agent)

    _log(
        f"启动 ruyipage Firefox: model={_browser_model_name(RUOYI_FIREFOX_PATH)} "
        f"headless={bool(args.headless)} path={RUOYI_FIREFOX_PATH}",
        "INFO",
    )

    browser_page = None
    try:
        browser_page = FirefoxPage(tb)
        rr._track_browser_page(browser_page)
        try:
            browser_page.close_other_tabs(browser_page)
        except Exception as exc:
            _log(f"关默认空白页失败(忽略): {type(exc).__name__}: {exc}", "WARN")

        target_url = str(args.url or "").strip() or None

        first = browser_page
        _open_target_url(first, target_url)
        if args.probe:
            ip = _probe_exit_ip(first)
            _log(f"tab#1 出口: {ip or '(探测失败)'}", "OK" if ip else "WARN")

        extra_tabs = []
        for i in range(2, tabs_wanted + 1):
            try:
                tab = browser_page.new_container_tab(url=target_url, background=False)
                extra_tabs.append(tab)
                _log(f"已开 container tab#{i}/{tabs_wanted}", "OK")
                if args.probe:
                    ip = _probe_exit_ip(tab)
                    _log(f"tab#{i} 出口: {ip or '(探测失败)'}", "OK" if ip else "WARN")
            except Exception as exc:
                _log(f"开 container tab#{i} 失败: {type(exc).__name__}: {exc}", "WARN")
                if str(args.proxy_exhausted).lower() == "stop":
                    _log("exhausted=stop，停止开新 tab", "WARN")
                    break

        _log(f"就绪：{1 + len(extra_tabs)} 个 tab 已打开，浏览器保持运行(keep={not args.no_keep})", "OK")
        _log("提示：Ctrl-C 优雅退出；--no-keep 时退出会关浏览器", "INFO")

        if not args.no_keep:
            try:
                while True:
                    time.sleep(3600)
            except KeyboardInterrupt:
                _log("收到 Ctrl-C，准备退出", "WARN")
        else:
            _log("--no-keep：浏览器将在脚本退出时关闭", "INFO")

    except KeyboardInterrupt:
        _log("收到 Ctrl-C，准备退出", "WARN")
    except Exception as exc:
        _eprint(f"[ERR] 启动失败: {type(exc).__name__}: {exc}")
        raise
    finally:
        if browser_page is not None and getattr(args, "no_keep", False):
            try:
                browser_page.quit(timeout=5)
                _log("浏览器已关闭", "OK")
            except Exception as exc:
                _log(f"关闭浏览器失败(忽略): {type(exc).__name__}: {exc}", "WARN")
        try:
            rr._cleanup_ruoyi_run_profile_dir(profile_dir)
        except Exception:
            pass
        if hub is not None:
            hub.stop()


def _build_per_tab(upstream_proxies):
    """无前置代理时的 per-tab 归一化（保留旧路径）。"""
    out = []
    for p in upstream_proxies:
        try:
            out.append(_per_tab_format(p))
        except ValueError as exc:
            _log(f"跳过不合规 per-tab 代理: {exc}", "WARN")
    return out


if __name__ == "__main__":
    main()


def _build_per_tab(upstream_proxies):
    """无前置代理时的 per-tab 归一化（保留旧路径）。"""
    out = []
    for p in upstream_proxies:
        try:
            out.append(_per_tab_format(p))
        except ValueError as exc:
            _log(f"跳过不合规 per-tab 代理: {exc}", "WARN")
    return out
