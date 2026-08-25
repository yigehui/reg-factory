"""A 组纯函数:从 register 迁入,逐字保持行为。

只 import 标准库 + 本包 _logging。无业务依赖。
"""

import os
import re
import urllib.parse

from ._logging import _normalize_log_level  # noqa: F401  re-export 用


def _env_bool(name, default=False):
    value = os.environ.get(name)
    if value is None:
        return bool(default)
    return str(value).strip().lower() not in ("0", "false", "no", "off", "")


def _mask_ua(ua):
    raw = str(ua or "")
    m = re.search(r"Firefox/([\d.]+)", raw)
    ver = m.group(1) if m else "?"
    return f"Firefox/{ver}"


def _strip_proxy_scheme(value):
    s = str(value or "").strip()
    for pfx in ("socks5h://", "socks5://", "socks4://", "http://", "https://"):
        if s.lower().startswith(pfx):
            return s[len(pfx):]
    return s


def _proxy_url_to_ruoyi(value):
    """Normalize proxy URL / user:pass@host:port / host:port to ruyipage format."""
    raw = str(value or "").strip()
    if not raw:
        return ""
    if "://" in raw:
        try:
            parsed = urllib.parse.urlsplit(raw)
            scheme = str(parsed.scheme or "").strip().lower()
            host = parsed.hostname or ""
            port = parsed.port or 0
            user = urllib.parse.unquote(parsed.username or "")
            pwd = urllib.parse.unquote(parsed.password or "")
            if scheme and host and port:
                auth = (
                    f"{urllib.parse.quote(user, safe='')}:{urllib.parse.quote(pwd, safe='')}@"
                    if (user or pwd)
                    else ""
                )
                return f"{scheme}://{auth}{host}:{port}"
        except Exception:
            pass
    s = _strip_proxy_scheme(raw).replace(",", "@", 1) if "@" not in raw and "," in raw else _strip_proxy_scheme(raw)
    if "@" in s:
        auth, hostport = s.rsplit("@", 1)
        if ":" in auth and ":" in hostport:
            user, pwd = auth.split(":", 1)
            host, port = hostport.rsplit(":", 1)
            return f"socks5://{urllib.parse.quote(user, safe='')}:{urllib.parse.quote(pwd, safe='')}@{host}:{port}"
    parts = s.split(":")
    if len(parts) == 4:
        host, port = parts[0], parts[1]
        user = parts[2]
        pwd = ":".join(parts[3:])
        return f"socks5://{urllib.parse.quote(user, safe='')}:{urllib.parse.quote(pwd, safe='')}@{host}:{port}"
    if len(parts) == 2 and parts[1].isdigit():
        return f"socks5://{s}"
    return ""


def _parse_ruoyi_proxy(proxy_str):
    s = str(proxy_str or "").strip()
    if not s:
        return None
    if "://" in s:
        try:
            parsed = urllib.parse.urlsplit(s)
            scheme = str(parsed.scheme or "").strip().lower() or "socks5"
            host = parsed.hostname or ""
            port = parsed.port or 0
            user = urllib.parse.unquote(parsed.username or "")
            pwd = urllib.parse.unquote(parsed.password or "")
            if host and port:
                return {"scheme": scheme, "host": host, "port": str(port), "username": user, "password": pwd}
        except Exception:
            pass
    parts = s.split(":")
    if len(parts) >= 4:
        host, port = parts[0], parts[1]
        user = parts[2]
        pwd = ":".join(parts[3:])
        return {"scheme": "socks5", "host": host, "port": port, "username": user, "password": pwd}
    if len(parts) == 2 and parts[1].isdigit():
        return {"scheme": "socks5", "host": parts[0], "port": parts[1], "username": "", "password": ""}
    normalized = _proxy_url_to_ruoyi(s)
    if normalized and normalized != s:
        return _parse_ruoyi_proxy(normalized)
    return None


def mask_ruoyi_proxy(proxy_str):
    p = _parse_ruoyi_proxy(proxy_str)
    if not p:
        return "***" if proxy_str else "noproxy"
    auth = f"{p['username'][:8]}...@" if p.get("username") else ""
    return f"{p.get('scheme') or 'http'}://{auth}{p['host']}:{p['port']}"


def _proxy_host_port_key(proxy_str):
    parsed = _parse_ruoyi_proxy(proxy_str)
    if parsed and parsed.get("host") and parsed.get("port"):
        return f"{parsed['host']}:{parsed['port']}"
    return str(proxy_str or "").strip()


def _coerce_float_or_none(value):
    try:
        if value in (None, ""):
            return None
        return float(value)
    except Exception:
        return None


def _geo_timezone_from_payload(payload):
    timezone_value = payload.get("timezone")
    if isinstance(timezone_value, dict):
        timezone_value = (
            timezone_value.get("id")
            or timezone_value.get("name")
            or timezone_value.get("timezone")
            or ""
        )
    timezone_value = timezone_value or payload.get("time_zone") or ""
    return str(timezone_value or "").strip()


def _browser_model_name(browser_path):
    parts = [p for p in os.path.normpath(str(browser_path or "")).split(os.sep) if p]
    for part in reversed(parts):
        low = part.lower()
        if low.startswith("firefox-") or "ruyi" in low:
            return part
    return os.path.basename(str(browser_path or "")) or "unknown"
