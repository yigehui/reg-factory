"""B 组:IP/代理探测闭环。

整条探测闭环:_probe_proxy_identity → _log_current_ip → _probe_proxy_before_browser。
本模块是包内唯一需要 requests 的模块。PROXY_PRECHECK_TARGETS 参数化,默认空(不预检)。
"""

import time
from urllib.parse import quote

import requests

from ._logging import log
from . import _state
from ._utils import (
    _parse_ruoyi_proxy,
    _proxy_host_port_key,
    _geo_timezone_from_payload,
    _coerce_float_or_none,
)


def _proxy_identity_cache_get(proxy_str):
    key = str(proxy_str or "").strip()
    if not key:
        return None
    now = time.time()
    with _state._CURRENT_IP_INFO_LOCK:
        cached = _state._CURRENT_IP_INFO.get(key)
        if not isinstance(cached, dict):
            return None
        expires_at = float(cached.get("expires_at") or 0.0)
        if expires_at and expires_at < now:
            _state._CURRENT_IP_INFO.pop(key, None)
            return None
        return dict(cached)


def _proxy_identity_cache_put(proxy_str, identity):
    key = str(proxy_str or "").strip()
    if not key or not isinstance(identity, dict):
        return identity
    cached = dict(identity)
    cached["expires_at"] = time.time() + max(1.0, float(_state.PROXY_IDENTITY_CACHE_TTL or 1.0))
    with _state._CURRENT_IP_INFO_LOCK:
        _state._CURRENT_IP_INFO[key] = cached
    return dict(cached)


def _probe_proxy_identity(proxy_or_pool, timeout=None, use_cache=True):
    if timeout is None:
        timeout = _state.PROXY_IDENTITY_TIMEOUT
    if isinstance(proxy_or_pool, str):
        proxy_str = str(proxy_or_pool or "").strip()
        proxy_pool = [proxy_str] if proxy_str else []
    else:
        proxy_pool = list(proxy_or_pool or [])
        proxy_str = str(proxy_pool[0] or "").strip() if proxy_pool else ""
    if not proxy_str:
        return None

    if use_cache:
        cached = _proxy_identity_cache_get(proxy_str)
        if cached:
            return cached

    proxies = _proxy_for_ip_lookup(proxy_pool, "[proxy-id]")
    session = requests.Session()
    session.trust_env = False
    request_timeout = max(0.1, float(timeout or _state.PROXY_IDENTITY_TIMEOUT))

    for source_name, ip_endpoint in _state.IP_INFO_ENDPOINTS:
        try:
            resp = session.get(
                ip_endpoint,
                headers={"Accept": "application/json, text/plain;q=0.9, */*;q=0.8"},
                proxies=proxies,
                timeout=request_timeout,
            )
            resp.raise_for_status()
            data = resp.json()
            ip = str(data.get("ip") or data.get("query") or "").strip()
            country = str(
                data.get("country_name")
                or data.get("country")
                or data.get("countryCode")
                or data.get("country_code")
                or ""
            ).strip()
            if not ip:
                continue
            identity = {
                "proxy": proxy_str,
                "ip": ip,
                "country": country,
                "country_code": str(data.get("countryCode") or data.get("country_code") or data.get("country") or "").strip().upper(),
                "timezone": _geo_timezone_from_payload(data),
                "latitude": _coerce_float_or_none(data.get("latitude") or data.get("lat")),
                "longitude": _coerce_float_or_none(data.get("longitude") or data.get("lon")),
                "source": source_name,
                "endpoint": ip_endpoint,
                "exit_key": f"ip:{ip}",
            }
            return _proxy_identity_cache_put(proxy_str, identity) if use_cache else identity
        except Exception:
            continue

    fallback = {
        "proxy": proxy_str,
        "ip": "",
        "country": "",
        "country_code": "",
        "timezone": "",
        "latitude": None,
        "longitude": None,
        "source": "fallback",
        "endpoint": "",
        "exit_key": f"proxy:{_proxy_host_port_key(proxy_str)}",
    }
    return _proxy_identity_cache_put(proxy_str, fallback) if use_cache else fallback


def _proxy_exit_key(proxy_str, timeout=None, use_cache=True):
    if timeout is None:
        timeout = _state.PROXY_IDENTITY_TIMEOUT
    identity = _probe_proxy_identity(proxy_str, timeout=timeout, use_cache=use_cache)
    if identity and identity.get("exit_key"):
        return str(identity["exit_key"])
    return f"proxy:{_proxy_host_port_key(proxy_str)}"


def _proxy_for_ip_lookup(proxy_pool, tag):
    if not proxy_pool:
        return None
    proxy_str = str(proxy_pool[0] or "").strip()
    p = _parse_ruoyi_proxy(proxy_str)
    if not p:
        log(f"  {tag} IP probe proxy format invalid: {proxy_str[:80]!r}", "WARN")
        return None
    user = quote(p.get("username") or "", safe="")
    pwd = quote(p.get("password") or "", safe="")
    auth = f"{user}:{pwd}@" if (user or pwd) else ""
    scheme = str(p.get("scheme") or "socks5").strip().lower() or "socks5"
    if scheme in {"socks", "socks5"}:
        scheme = "socks5h"
    proxy_url = f"{scheme}://{auth}{p['host']}:{p['port']}"
    return {"http": proxy_url, "https": proxy_url}


def _log_current_ip(proxy_pool, tag):
    identity = _probe_proxy_identity(proxy_pool, timeout=15)
    if identity and identity.get("ip"):
        country_text = str(identity.get("country") or "UNKNOWN").strip() or "UNKNOWN"
        log(f"  {tag} current IP: {identity['ip']!r} country: {country_text!r} via {identity.get('source')}")
        return
    log(f"  {tag} current IP/country probe failed", "WARN")


def _probe_proxy_targets(proxy_pool, timeout=None, targets=None):
    if timeout is None:
        timeout = _state.PROXY_PRECHECK_TIMEOUT
    effective_targets = targets if targets is not None else _state.PROXY_PRECHECK_TARGETS
    if not effective_targets:
        return None
    proxies = _proxy_for_ip_lookup(proxy_pool, "[proxy-precheck]")
    if not proxies:
        return None
    session = requests.Session()
    session.trust_env = False
    request_timeout = max(0.1, float(timeout or _state.PROXY_PRECHECK_TIMEOUT))
    for target_url in effective_targets:
        try:
            resp = session.get(
                target_url,
                headers={"Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"},
                proxies=proxies,
                timeout=request_timeout,
                allow_redirects=False,
                stream=True,
            )
            status_code = int(getattr(resp, "status_code", 0) or 0)
            try:
                resp.close()
            except Exception:
                pass
            if 100 <= status_code < 500:
                return {
                    "url": target_url,
                    "status_code": status_code,
                }
        except Exception:
            continue
    return None


def _probe_proxy_before_browser(proxy_pool, tag, timeout=None, targets=None):
    if timeout is None:
        timeout = _state.PROXY_PRECHECK_TIMEOUT
    effective_targets = targets if targets is not None else _state.PROXY_PRECHECK_TARGETS
    if not effective_targets:
        # 无预检目标(新包默认):跳过预检,视为通过。
        log(f"  {tag} proxy precheck skipped (no targets)", "DEBUG")
        return True
    proxies = _proxy_for_ip_lookup(proxy_pool, tag)
    if not proxies:
        return True
    result = _probe_proxy_targets(proxy_pool, timeout=max(0.1, float(timeout or _state.PROXY_PRECHECK_TIMEOUT)), targets=effective_targets)
    if result:
        log(
            f"  {tag} proxy precheck ok: target={result['url']} status={result['status_code']}",
            "INFO",
        )
        return True
    log(f"  {tag} proxy precheck failed: target unreachable via proxy targets={','.join(effective_targets)}", "WARN")
    return False


# 兼容旧名:空 targets 时为 ""
PROXY_PRECHECK_URL = _state.PROXY_PRECHECK_TARGETS[0] if _state.PROXY_PRECHECK_TARGETS else ""
