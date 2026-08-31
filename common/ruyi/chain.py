"""I 组:代理链中继(前置代理)。launch/register/unlock 三处共用的自包含实现。

解决"上游代理需外网 IP 才能连"(如 1024proxy 做 IP 白名单):Firefox 不能链两段
SOCKS5,故本地起一个 SOCKS5 中继:
  Firefox --SOCKS5(无认证/任意账号)--> 127.0.0.1:LOCAL
  中继 --(前置代理 front)--> 上游代理 upstream --(目标 host:port)
front / upstream 均支持 socks5(可带账号) 与 http(CONNECT+Basic) 两种协议。

三层用法:
- launch:直接 new ChainRelayHub(front) + start()/port_for()/stop()(自有实例,生命周期随进程)。
- register/unlock:用进程级单例 ensure_front_relay() + front_relay_url() 映射,
  daemon 线程 + os._exit 强杀路径天然清理,无需显式 stop。
- front_relay_url(upstream, front=..., per_tab=...) 是唯一入口:front 空 → 原样返回;
  非空 → socks5://127.0.0.1:{port}(per_tab=True 加 :u:p 后缀,中继对任意账号放行)。

自包含:只依赖标准库;解析走 _utils._parse_ruoyi_proxy(纯函数)。
"""

import asyncio
import base64
import functools
import os
import socket
import struct
import threading

from ._logging import log
from ._utils import _parse_ruoyi_proxy

__all__ = [
    "parse_front_proxy",
    "ChainRelayHub",
    "ensure_front_relay",
    "front_relay_url",
    "front_proxy_raw",
    "FRONT_PROXY_ENV",
]


# ---------------- 解析 ----------------

FRONT_PROXY_ENV = "LAUNCH_FRONT_PROXY"


def parse_front_proxy(url):
    """代理 URL -> {scheme, host, port, username, password}。非法抛 ValueError。

    scheme 缺省 socks5;port 强转 int。socks5/http(CONNECT) 两种协议。"""
    p = _parse_ruoyi_proxy(url)
    if not p or not p.get("host") or not p.get("port"):
        raise ValueError(f"非法代理 URL: {url}")
    p = dict(p)
    p["scheme"] = (p.get("scheme") or "socks5").lower()
    p["port"] = int(p["port"])
    return p


def front_proxy_raw(value=None):
    """CLI/opts 传入值优先,空则读 env FRONT_PROXY_ENV(LAUNCH_FRONT_PROXY),strip 后返回。"""
    raw = str(value if value is not None else "" or "").strip()
    if not raw:
        raw = str(os.environ.get(FRONT_PROXY_ENV, "") or "").strip()
    return raw


# ---------------- 隧道原语 ----------------


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
        log(f"上游 {up_scheme} 连接失败({type(exc).__name__}: {str(exc)[:80]})，回退 {alt} 重试", "WARN")
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
            log(f"[relay {tag}] 链路建立失败 {target_host}:{target_port} -> {type(exc).__name__}: {str(exc)[:100]}", "WARN")
            client_writer.write(b"\x05\x05\x00\x01\x00\x00\x00\x00\x00\x00")  # connection refused
            await client_writer.drain()
            return
        client_writer.write(b"\x05\x00\x00\x01\x00\x00\x00\x00\x00\x00")  # success
        await client_writer.drain()
        log(f"[relay {tag}] {target_host}:{target_port} 链路就绪", "DEBUG")
        # 4) 双向透传
        await asyncio.gather(
            _pipe(client_reader, up_writer),
            _pipe(up_reader, client_writer),
        )
    except asyncio.IncompleteReadError:
        pass
    except Exception as exc:
        log(f"[relay {tag}] 处理异常: {type(exc).__name__}: {str(exc)[:80]}", "WARN")
    finally:
        try:
            client_writer.close()
        except Exception:
            pass


# ---------------- 中继 hub ----------------


class ChainRelayHub:
    """按上游代理去复用本地 SOCKS5 中继端口；单事件循环跑在后台线程。

    port_for 整段持锁:多 worker 并发对同一上游只起一个 server(register/unlock
    多账号并发场景);launch 单线程用法行为不变。"""

    def __init__(self, front):
        self.front = front
        self.loop = None
        self.thread = None
        self._servers = []
        self._port_by_key = {}
        self._port_lock = threading.Lock()

    def start(self):
        self.loop = asyncio.new_event_loop()

        def _run():
            asyncio.set_event_loop(self.loop)
            self.loop.run_forever()

        self.thread = threading.Thread(target=_run, name="chain-relay", daemon=True)
        self.thread.start()

    def port_for(self, upstream):
        key = f"{upstream.get('host')}:{upstream.get('port')}"
        # 整段持锁:check-then-act 不能拆——若锁只护缓存读写,多 worker 同时 miss
        # 会对同一上游各起一个 server(register/unlock 多账号并发场景)。server 起
        # 在 hub 自有事件循环线程上,不取该锁,持锁阻塞等待不会死锁。
        with self._port_lock:
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


# ---------------- 进程级单例(register/unlock 复用) ----------------

_HUB = None
_HUB_LOCK = threading.Lock()


def ensure_front_relay(front_raw, tag=""):
    """解析并启动前置代理中继(进程级单例,幂等)。front 非法抛 ValueError。返回 hub。

    同一进程重复调用:front 相同 → 复用已起 hub;front 不同 → 以首个为准并 WARN
    (本地中继端口已分发给 Firefox,换 hub 会导致已映射端口失效)。"""
    global _HUB
    raw = str(front_raw or "").strip()
    if not raw:
        return None
    front = parse_front_proxy(raw)  # 非法抛 ValueError
    with _HUB_LOCK:
        if _HUB is not None:
            return _HUB
        hub = ChainRelayHub(front)
        hub.start()
        _HUB = hub
    log(f"{tag} 代理链中继已启动，前置代理: {front['scheme']}://{front['host']}:{front['port']}", "OK")
    return _HUB


def front_relay_url(upstream, front="", per_tab=False, tag=""):
    """upstream 代理串 -> Firefox 可直连的代理串。

    front 空 → 原样返回 upstream(无前置,行为不变);
    非空 → 起进程级中继并映射: socks5://127.0.0.1:{port}
           per_tab=True → socks5://127.0.0.1:{port}:u:p(中继对任意账号放行)。
    upstream 非法/解析失败时:front 开着还硬塞原串,Firefox 直连必然失败且无报错
    指引 —— 直接抛 ValueError 让调用方按不合规代理处理(跳过该条)。"""
    raw_front = front_proxy_raw(front)
    up = str(upstream or "").strip()
    if not up:
        return up
    if not raw_front:
        return up
    hub = ensure_front_relay(raw_front, tag=tag)
    parsed = parse_front_proxy(up)  # 非法抛 ValueError
    port = hub.port_for(parsed)
    if per_tab:
        url = f"socks5://127.0.0.1:{port}:u:p"
    else:
        url = f"socks5://127.0.0.1:{port}"
    log(f"{tag} 前置代理链: {parsed['host']}:{parsed['port']} -> 127.0.0.1:{port}", "OK")
    return url
