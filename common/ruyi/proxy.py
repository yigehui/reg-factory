"""D 组:代理池 ConsumableProxyPool + load/parse/fetch。

中性 env 名 + 通用默认名(proxies.txt)。_PROXY_LIST 单例经 _state 读写。
fetch_proxy_list_http 用 urllib 不用 requests。单向依赖 probe(_proxy_exit_key)。
"""

import os
import random
import threading
import urllib.parse
import urllib.request
from types import SimpleNamespace

from ._logging import log
from . import _state
from ._utils import _proxy_url_to_ruoyi, _parse_ruoyi_proxy, mask_ruoyi_proxy
from .probe import _proxy_exit_key

PROXY_FILE = os.environ.get("RUOYI_PROXY_FILE", "proxies.txt")
PROXY_URL = os.environ.get("RUOYI_PROXY_URL", "")
RUOYI_PROXY_SOURCE = os.environ.get("RUOYI_PROXY_SOURCE", "file")


def _parse_proxy_lines(lines, source_label):
    out = []
    for raw in lines or []:
        ln = str(raw or "").strip().lstrip("﻿")
        if not ln or ln.startswith("#"):
            continue
        normalized = _proxy_url_to_ruoyi(ln)
        if not normalized:
            log(f"跳过非法代理行({source_label}): {ln[:80]}", "WARN")
            continue
        out.append(normalized)
    return out


def parse_proxy_pool(path):
    """读本地代理文件,支持 URL、user:pass@host:port、host:port。"""
    if not path:
        return []
    if not os.path.isfile(path):
        log(f"代理池文件不存在: {path}", "WARN")
        return []
    with open(path, "r", encoding="utf-8") as f:
        return _parse_proxy_lines(f, path)


def fetch_proxy_list_http(proxy_url, timeout=12):
    proxy_url = str(proxy_url or "").strip()
    if not proxy_url:
        raise RuntimeError("RUOYI_PROXY_URL/--proxy-url 为空")
    parsed = urllib.parse.urlsplit(proxy_url)
    if str(parsed.scheme or "").lower() not in {"http", "https"} or not parsed.netloc:
        raise RuntimeError(f"代理列表地址仅支持 http/https: {proxy_url}")
    req = urllib.request.Request(proxy_url, headers={
        "Accept": "text/plain, */*",
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/138.0.0.0 Safari/537.36"
        ),
    })
    # 强制直连拉列表:urllib 默认会吃 HTTP_PROXY/HTTPS_PROXY/WinIE 系统代理,
    # 用空 ProxyHandler 显式禁用代理,确保直连到代理池 API。
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    with opener.open(req, timeout=timeout) as resp:
        data = resp.read().decode("utf-8", errors="replace")
    return _parse_proxy_lines(data.splitlines(), proxy_url)


def _proxy_source_label(args, default=None):
    source = str(getattr(args, "proxy_source", "") or default or RUOYI_PROXY_SOURCE or "file").strip().lower()
    if source in {"file", "http"}:
        return source
    return "file"


def load_proxy_list(args):
    """按当前来源加载代理 list。"""
    source = _proxy_source_label(args)
    proxy_url = str(getattr(args, "proxy_url", "") or PROXY_URL).strip()
    if source == "http":
        return fetch_proxy_list_http(proxy_url)
    batch = parse_proxy_pool(getattr(args, "proxy_file", "") or PROXY_FILE)
    if batch or not proxy_url:
        return batch
    log("proxy source=file 但本地代理文件为空,回退到 proxy-url HTTP 列表", "WARN")
    return fetch_proxy_list_http(proxy_url)


# 兼容旧名
load_proxy_batch = load_proxy_list


def build_proxy_source(args):
    """兼容旧调用:返回当前来源的一批代理 list。"""
    return load_proxy_list(args)


class ConsumableProxyPool:
    """任务级代理 list(就一个 list,不是消息队列)。

    启动 load → 注册 pop 一条 → list 空了再 load → 任务停 clear。
    加锁只是为了并发注册不抢同一条。
    """

    def __init__(self, source_args):
        self.source_args = source_args
        self.source = _proxy_source_label(source_args)
        self._lock = threading.Lock()
        self._list = []
        self._active_exit_keys = set()
        self._active_proxy_keys = {}

    @classmethod
    def from_args(cls, args):
        source_args = SimpleNamespace(
            proxy_file=getattr(args, "proxy_file", "") or PROXY_FILE,
            proxy_source=getattr(args, "proxy_source", "") or RUOYI_PROXY_SOURCE,
            proxy_url=getattr(args, "proxy_url", "") or PROXY_URL,
        )
        return cls(source_args)

    def __bool__(self):
        with self._lock:
            return bool(self._list)

    def remaining(self):
        with self._lock:
            return len(self._list)

    def stats(self):
        with self._lock:
            return {
                "remaining": len(self._list),
                "source": self.source,
                "active_exit_keys": len(self._active_exit_keys),
            }

    def _load_locked(self):
        """文件 load / 接口重调,结果塞进 self._list。"""
        try:
            batch = load_proxy_list(self.source_args) or []
        except Exception as exc:
            log(f"proxy load failed ({self.source}): {type(exc).__name__}: {exc}", "WARN")
            batch = []
        # 启动只做 host:port 去重;出口 IP 延后到 take() 按需探测,避免启动前全量预检。
        seen = set()
        fresh = []
        dup_host = 0
        for p in batch:
            p = str(p or "").strip()
            if not p:
                continue
            parsed = _parse_ruoyi_proxy(p)
            key = f"{parsed['host']}:{parsed['port']}" if parsed and parsed.get("host") and parsed.get("port") else p
            if key in seen:
                dup_host += 1
                continue
            seen.add(key)
            fresh.append(p)
        self._list.extend(fresh)
        log(
            f"proxy load source={self.source} got={len(fresh)} list={len(self._list)} dup_host={dup_host}",
            "INFO" if fresh else "WARN",
        )
        return len(fresh)

    def start(self):
        """任务启动:初始化 list。"""
        with self._lock:
            self._list = []
            self._load_locked()
            if not self._list:
                log(f"proxy list empty after start (source={self.source})", "WARN")
            return self

    def reload(self):
        """list 空时重载:文件再读 / 接口再调。"""
        with self._lock:
            return self._load_locked()

    def take(self):
        """随机取一条未与当前活动出口 IP 冲突的代理并删除。"""
        with self._lock:
            if not self._list:
                self._load_locked()
            if not self._list:
                return []
            deferred = []
            chosen = None
            while self._list:
                idx = random.randrange(len(self._list))
                proxy = self._list.pop(idx)
                exit_key = _proxy_exit_key(proxy)
                if exit_key in self._active_exit_keys:
                    deferred.append(proxy)
                    continue
                self._active_exit_keys.add(exit_key)
                self._active_proxy_keys[proxy] = exit_key
                chosen = proxy
                break
            if deferred:
                self._list.extend(deferred)
            if not chosen:
                log("proxy take blocked: remaining proxies share active exit IPs", "WARN")
                return []
            log(f"proxy take -> {mask_ruoyi_proxy(chosen)} remaining={len(self._list)} active={len(self._active_exit_keys)}", "DEBUG")
            return [chosen]

    def release(self, proxy):
        proxy = str(proxy or "").strip()
        if not proxy:
            return False
        with self._lock:
            exit_key = self._active_proxy_keys.pop(proxy, None) or _proxy_exit_key(proxy)
            released = exit_key in self._active_exit_keys
            self._active_exit_keys.discard(exit_key)
            return released

    def borrow_reuse(self):
        """代理不足(take 返回 [])时强制复用池里任意一条。

        不删除、不加入 _active_exit_keys、release 无效 -- 仅供并发 worker 数 > 代理数时共享出口 IP。
        返回 [proxy] 或 []。"""
        with self._lock:
            if not self._list:
                self._load_locked()
            if not self._list:
                return []
            return [self._list[random.randrange(len(self._list))]]

    def discard(self, proxy):
        """永久移除一条代理(失效/不可用),不再分配。take 不会再取到它。"""
        proxy = str(proxy or "").strip()
        if not proxy:
            return False
        with self._lock:
            exit_key = self._active_proxy_keys.pop(proxy, None) or _proxy_exit_key(proxy)
            self._active_exit_keys.discard(exit_key)
            try:
                self._list.remove(proxy)
            except ValueError:
                pass
            return True

    def stop(self):
        """任务结束:销毁 list。"""
        with self._lock:
            n = len(self._list)
            self._list = []
            self._active_exit_keys.clear()
            self._active_proxy_keys.clear()
            log(f"proxy list destroyed (cleared {n})", "INFO")


def get_consumable_proxy_pool():
    return _state._PROXY_LIST


def set_consumable_proxy_pool(pool):
    """任务启动绑定 list,结束传 None。"""
    _state._PROXY_LIST = pool
    return _state._PROXY_LIST


# 兼容旧名
def get_session_proxy_runtime():
    return _state._PROXY_LIST


def set_session_proxy_runtime(runtime):
    return set_consumable_proxy_pool(runtime)


SessionProxyRuntime = ConsumableProxyPool


def select_proxy_for_account(proxy_pool=None, runtime=None):
    """注册前从 list 取一条(取后删除)。"""
    pool = runtime if runtime is not None else None
    if pool is None and isinstance(proxy_pool, ConsumableProxyPool):
        pool = proxy_pool
    if pool is None:
        pool = _state._PROXY_LIST
    if isinstance(pool, ConsumableProxyPool):
        return pool.take()
    if not proxy_pool:
        return []
    items = list(proxy_pool)
    if not items:
        return []
    return [items.pop(random.randrange(len(items)))]


def release_proxy_for_account(proxy=None, proxy_pool=None, runtime=None):
    pool = runtime if runtime is not None else None
    if pool is None and isinstance(proxy_pool, ConsumableProxyPool):
        pool = proxy_pool
    if pool is None:
        pool = _state._PROXY_LIST
    if isinstance(pool, ConsumableProxyPool):
        return pool.release(proxy)
    return False
