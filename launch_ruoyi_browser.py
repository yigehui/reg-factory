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
import asyncio
import base64
import functools
import os
import socket
import struct
import sys
import threading
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
# Firefox 不能链两段 SOCKS5，故本地起一个 SOCKS5 中继：
#   Firefox --SOCKS5(无认证/任意账号)--> 127.0.0.1:LOCAL
#   中继 --(前置代理 front)--> 上游代理 upstream --(目标 host:port)
# front / upstream 均支持 socks5(可带账号) 与 http(CONNECT+Basic) 两种协议。


def _parse_proxy_url(url):
    p = rr._parse_ruoyi_proxy(url)
    if not p or not p.get("host") or not p.get("port"):
        raise ValueError(f"非法代理 URL: {url}")
    p = dict(p)
    p["scheme"] = (p.get("scheme") or "socks5").lower()
    p["port"] = int(p["port"])
    return p


async def _socks5_connect(reader, writer, host, port, auth, log_tag):
    """在已连到某 SOCKS5 代理的 socket 上，请求 CONNECT 到 host:port。auth=(user,pass)|None。"""
    if auth:
        writer.write(b"\x05\x01\x02")  # VER NMETHODS METHODS=[user/pass]
        await writer.drain()
        m = await reader.readexactly(2)
        if m[1] != 0x02:
            raise RuntimeError(f"[{log_tag}] 前置/上游 SOCKS5 不支持账号认证")
        user, pwd = (auth[0] or "").encode(), (auth[1] or "").encode()
        writer.write(b"\x01" + bytes([len(user)]) + user + bytes([len(pwd)]) + pwd)
        await writer.drain()
        r = await reader.readexactly(2)
        if r[1] != 0x00:
            raise RuntimeError(f"[{log_tag}] SOCKS5 账号认证失败(rep={r[1]})")
    else:
        writer.write(b"\x05\x01\x00")  # 无认证
        await writer.drain()
        m = await reader.readexactly(2)
        if m[1] != 0x00:
            raise RuntimeError(f"[{log_tag}] SOCKS5 拒绝无认证(method={m[1]})")
    hb = host.encode()
    writer.write(b"\x05\x01\x00\x03" + bytes([len(hb)]) + hb + struct.pack(">H", port))
    await writer.drain()
    rep = await reader.readexactly(4)  # VER REP RSV ATYP
    if rep[1] != 0x00:
        raise RuntimeError(f"[{log_tag}] SOCKS5 CONNECT 失败 rep={rep[1]}")
    atyp = rep[3]
    if atyp == 0x01:
        await reader.readexactly(4)
    elif atyp == 0x03:
        ln = (await reader.readexactly(1))[0]
        await reader.readexactly(ln)
    elif atyp == 0x04:
        await reader.readexactly(16)
    else:
        raise RuntimeError(f"[{log_tag}] SOCKS5 未知 ATYP={atyp}")
    await reader.readexactly(2)  # BND.PORT


async def _http_connect(reader, writer, host, port, auth, log_tag):
    """在已连到某 HTTP 代理的 socket 上，发 CONNECT。auth=(user,pass)|None。"""
    lines = [f"CONNECT {host}:{port} HTTP/1.1", f"Host: {host}:{port}"]
    if auth:
        cred = base64.b64encode(f"{auth[0]}:{auth[1]}".encode()).decode()
        lines.append(f"Proxy-Authorization: Basic {cred}")
    lines += ["", ""]
    writer.write("\r\n".join(lines).encode())
    await writer.drain()
    buf = b""
    while b"\r\n\r\n" not in buf:
        chunk = await reader.read(4096)
        if not chunk:
            raise RuntimeError(f"[{log_tag}] HTTP 代理早断(无响应头)")
        buf += chunk
    status_line = buf.split(b"\r\n", 1)[0].decode("latin1", "replace")
    parts = status_line.split()
    if len(parts) < 2 or parts[1] != "200":
        raise RuntimeError(f"[{log_tag}] HTTP CONNECT 失败: {status_line[:120]}")


async def _tunnel_to(reader, writer, scheme, host, port, auth, log_tag):
    if scheme.startswith("socks"):
        await _socks5_connect(reader, writer, host, port, auth, log_tag)
    else:
        await _http_connect(reader, writer, host, port, auth, log_tag)


async def _open_chain(target_host, target_port, front, upstream):
    """front -> upstream -> target，返回 (reader, writer) 已贯通到 target。"""
    reader, writer = await asyncio.open_connection(front["host"], front["port"])
    front_auth = (front.get("username"), front.get("password")) if front.get("username") else None
    # 1) front 把我们送到 upstream 主机
    await _tunnel_to(reader, writer, front.get("scheme", "socks5"),
                     upstream["host"], upstream["port"], front_auth, "front")
    # 2) upstream 把我们送到目标；socks5/http 自动兜底
    up_auth = (upstream.get("username"), upstream.get("password")) if upstream.get("username") else None
    up_scheme = upstream.get("scheme", "socks5")
    try:
        await _tunnel_to(reader, writer, up_scheme, target_host, target_port, up_auth, "upstream")
    except Exception as exc:
        # 用户填的 scheme 可能与上游实际协议不符（如 1024proxy:3000 实为 HTTP 却填 socks5），
        # 换另一种协议重试一次（需重开整条隧道，因为 socket 已被半截握手污染）。
        alt = "http" if up_scheme.startswith("socks") else "socks5"
        _log(f"上游 {up_scheme} 连接失败({type(exc).__name__}: {str(exc)[:80]})，回退 {alt} 重试", "WARN")
        try:
            writer.close()
        except Exception:
            pass
        reader, writer = await asyncio.open_connection(front["host"], front["port"])
        await _tunnel_to(reader, writer, front.get("scheme", "socks5"),
                         upstream["host"], upstream["port"], front_auth, "front")
        await _tunnel_to(reader, writer, alt, target_host, target_port, up_auth, "upstream")
    return reader, writer


async def _pipe(src_reader, dst_writer):
    try:
        while True:
            data = await src_reader.read(65536)
            if not data:
                break
            dst_writer.write(data)
            await dst_writer.drain()
    except Exception:
        pass
    finally:
        try:
            dst_writer.close()
        except Exception:
            pass


async def _handle_client(client_reader, client_writer, front, upstream):
    """对 Firefox 暴露 SOCKS5（接受无认证与 user/pass 两种方法，均放行）。"""
    tag = f"{upstream['host']}:{upstream['port']}"
    try:
        # 1) SOCKS5 握手
        ver = await client_reader.readexactly(2)  # 05 NMETHODS
        nmethods = ver[1]
        methods = await client_reader.readexactly(nmethods)
        if 0x00 in methods:
            client_writer.write(b"\x05\x00")
            await client_writer.drain()
        elif 0x02 in methods:
            client_writer.write(b"\x05\x02")
            await client_writer.drain()
            hdr = await client_reader.readexactly(2)  # 01 ULEN
            ulen = hdr[1]
            await client_reader.readexactly(ulen)
            plen = (await client_reader.readexactly(1))[0]
            await client_reader.readexactly(plen)
            client_writer.write(b"\x01\x00")  # 认证通过(任意账号都放行)
            await client_writer.drain()
        else:
            client_writer.write(b"\x05\xff")
            await client_writer.drain()
            return
        # 2) 读 CONNECT 请求
        req = await client_reader.readexactly(4)  # VER CMD RSV ATYP
        if req[1] != 0x01:  # 只支持 CONNECT
            client_writer.write(b"\x05\x07\x00\x01\x00\x00\x00\x00\x00\x00")
            await client_writer.drain()
            return
        atyp = req[3]
        if atyp == 0x01:
            target_host = socket.inet_ntoa(await client_reader.readexactly(4))
        elif atyp == 0x03:
            ln = (await client_reader.readexactly(1))[0]
            target_host = (await client_reader.readexactly(ln)).decode("utf-8", "replace")
        elif atyp == 0x04:
            target_host = socket.inet_ntop(socket.AF_INET6, await client_reader.readexactly(16))
        else:
            client_writer.write(b"\x05\x08\x00\x01\x00\x00\x00\x00\x00\x00")
            await client_writer.drain()
            return
        target_port = struct.unpack(">H", await client_reader.readexactly(2))[0]
        # 3) 开链
        try:
            up_reader, up_writer = await _open_chain(target_host, target_port, front, upstream)
        except Exception as exc:
            _log(f"[relay {tag}] 链路建立失败 {target_host}:{target_port} -> {type(exc).__name__}: {str(exc)[:100]}", "WARN")
            client_writer.write(b"\x05\x05\x00\x01\x00\x00\x00\x00\x00\x00")  # connection refused
            await client_writer.drain()
            return
        client_writer.write(b"\x05\x00\x00\x01\x00\x00\x00\x00\x00\x00")  # success
        await client_writer.drain()
        _log(f"[relay {tag}] {target_host}:{target_port} 链路就绪", "DEBUG")
        # 4) 双向透传
        await asyncio.gather(
            _pipe(client_reader, up_writer),
            _pipe(up_reader, client_writer),
        )
    except asyncio.IncompleteReadError:
        pass
    except Exception as exc:
        _log(f"[relay {tag}] 处理异常: {type(exc).__name__}: {str(exc)[:80]}", "WARN")
    finally:
        try:
            client_writer.close()
        except Exception:
            pass


class ChainRelayHub:
    """按上游代理去复用本地 SOCKS5 中继端口；单事件循环跑在后台线程。"""

    def __init__(self, front):
        self.front = front
        self.loop = None
        self.thread = None
        self._servers = []
        self._port_by_key = {}

    def start(self):
        self.loop = asyncio.new_event_loop()

        def _run():
            asyncio.set_event_loop(self.loop)
            self.loop.run_forever()

        self.thread = threading.Thread(target=_run, name="chain-relay", daemon=True)
        self.thread.start()

    def port_for(self, upstream):
        key = f"{upstream.get('host')}:{upstream.get('port')}"
        if key in self._port_by_key:
            return self._port_by_key[key]
        fut = asyncio.run_coroutine_threadsafe(self._start_server(upstream), self.loop)
        port = fut.result(timeout=15)
        self._port_by_key[key] = port
        return port

    async def _start_server(self, upstream):
        handler = functools.partial(_handle_client, front=self.front, upstream=upstream)
        server = await asyncio.start_server(handler, "127.0.0.1", 0, backlog=32)
        self._servers.append(server)
        return server.sockets[0].getsockname()[1]

    def stop(self):
        if not self.loop:
            return
        # 先在事件循环线程里关掉所有 server，再停 loop，避免 "Task was destroyed but it is pending"
        async def _shutdown():
            for srv in self._servers:
                try:
                    srv.close()
                    await srv.wait_closed()
                except Exception:
                    pass
        try:
            fut = asyncio.run_coroutine_threadsafe(_shutdown(), self.loop)
            fut.result(timeout=5)
        except Exception:
            pass
        try:
            self.loop.call_soon_threadsafe(self.loop.stop)
        except Exception:
            pass


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
