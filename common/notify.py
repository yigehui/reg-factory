# -*- coding: utf-8 -*-
"""
common/notify.py - Telegram 通知发送，供任意注册/导出/循环流程复用。

配置走 .env（也可调用时直接传参覆盖）：
  TG_BOT_TOKEN  bot token
  TG_CHAT_ID    目标 chat id
  TG_PROXY      访问 api.telegram.org 的代理(国内网络必填)，支持
                http://  https://  socks5://  socks5h://(被 DNS 污染的域名用 socks5h，远程解析)
                带认证代理写 user:pass@host:port，如
                socks5h://yigehui:zxcv2684@163.192.35.163:41099

设计要点：
  - 用 requests 发送：支持 socks5 代理与代理认证(urllib 两者都不支持)。
    socks5 代理需装 PySocks(pip install pysocks)，否则会报 "Missing dependencies for SOCKS support"。
  - 自带重试(代理偶发 reset)；POST 仅对连接级错误重试，避免重复发消息。
  - token/chat_id 缺失或发送失败仅告警，绝不抛异常，不阻断调用方主流程。
  - 失败时把具体原因(HTTP 状态码/异常)打到日志，不再静默。
  - log_fn 可注入宿主日志函数 log_fn(msg, level)；不传则用 logging，兼容独立运行。
"""

import logging
import os

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from config import _load_dotenv

_load_dotenv()

_logger = logging.getLogger(__name__)

TG_API_BASE = "https://api.telegram.org"
TG_DEFAULT_TIMEOUT = 15
TG_MAX_RETRIES = 4  # 应对代理偶发 reset

TG_TEXT_LIMIT = 4096  # Telegram sendMessage 单条文本上限


def _split_tg_text(text, limit=TG_TEXT_LIMIT):
    """按 limit 切分长文本，优先在换行处断开；按字符切不会截断 UTF-8 多字节。"""
    chunks = []
    remaining = str(text or "")
    while len(remaining) > limit:
        cut = remaining.rfind("\n", 0, limit)
        if cut <= 0:  # 该段无合适换行，硬切
            cut = limit
        chunks.append(remaining[:cut])
        remaining = remaining[cut:].lstrip("\n")
    if remaining:
        chunks.append(remaining)
    return chunks


def _send_one_tg(sess, url, chat_id, text, timeout):
    """发单条 sendMessage；返回 (ok, error_desc)。"""
    resp = sess.post(url, data={"chat_id": chat_id, "text": text}, timeout=timeout)
    if resp.ok and resp.json().get("ok"):
        return True, None
    try:
        desc = resp.json().get("description", resp.text[:120])
    except Exception:
        desc = resp.text[:120]
    return False, f"HTTP {resp.status_code} {desc}"


def _default_log(msg, level="INFO"):
    """未注入 log_fn 时的兜底：映射到标准 logging。"""
    lvl = (level or "INFO").strip().upper() or "INFO"
    if lvl in ("WARN", "ERR", "ERROR"):
        _logger.warning(msg)
    elif lvl == "DEBUG":
        _logger.debug(msg)
    else:  # INFO / OK
        _logger.info(msg)


def _make_retry():
    """构造 Retry；POST 也纳入重试(连接级 reset 时请求并未发出，重试不会重复发消息)。
    urllib3 1.26+ 用 allowed_methods，更老版本用 method_whitelist，做兼容。"""
    common = dict(total=TG_MAX_RETRIES, backoff_factor=0.8,
                  status_forcelist=[429, 500, 502, 503, 504])
    try:
        return Retry(allowed_methods=frozenset(["GET", "POST"]), **common)
    except TypeError:  # pragma: no cover - 老版 urllib3
        return Retry(method_whitelist=frozenset(["GET", "POST"]), **common)


def _normalize_proxy(proxy):
    """补全代理 scheme(裸 host:port 按 http 处理)；返回 (proxy, warning)。
    socks5 必须显式写 socks5:// 或 socks5h://，无法自动推断。"""
    if not proxy:
        return "", ""
    p = proxy.strip()
    if "://" not in p:
        return "http://" + p, (f"TG_PROXY 未带 scheme，已按 http 处理；"
                               f"若实为 socks5 请显式写 socks5h://user:pass@host:port: {p}")
    return p, ""


def _build_session(proxy, timeout):
    """构造带重试的 requests.Session；socks5 代理需运行环境已装 PySocks。"""
    sess = requests.Session()
    adapter = HTTPAdapter(max_retries=_make_retry())
    sess.mount("https://", adapter)
    sess.mount("http://", adapter)
    if proxy:
        sess.proxies.update({"http": proxy, "https": proxy})
    return sess


def send_tg_message(text, bot_token=None, chat_id=None, proxy=None,
                    timeout=TG_DEFAULT_TIMEOUT, log_fn=None):
    """发送一条 Telegram 文本消息；成功返回 True，失败/未配置返回 False。

    bot_token/chat_id/proxy 缺省时分别读 TG_BOT_TOKEN/TG_CHAT_ID/TG_PROXY。
    proxy 支持 http:// https:// socks5:// socks5h://，带认证代理写 user:pass@host:port。
    socks5 代理需安装 PySocks(pip install pysocks)。
    log_fn(msg, level) 可注入宿主日志(如 ruoyi 的 log)；不传走 logging。
    任何异常都吞掉只记日志，保证不阻断调用方主流程。
    """
    log = log_fn or _default_log
    bot_token = bot_token or os.environ.get("TG_BOT_TOKEN", "")
    chat_id = chat_id or os.environ.get("TG_CHAT_ID", "")
    proxy = proxy or os.environ.get("TG_PROXY", "")
    if not bot_token or not chat_id:
        return False
    proxy, warn = _normalize_proxy(proxy)
    if warn:
        log(warn, "WARN")
    try:
        sess = _build_session(proxy, timeout)
        url = f"{TG_API_BASE}/bot{bot_token}/sendMessage"
        chunks = _split_tg_text(text, TG_TEXT_LIMIT)
        if len(chunks) > 1:
            log(f"TG 文本 {len(str(text))} 字符超 4096 上限，拆 {len(chunks)} 条发送", "INFO")
        ok_all = True
        first_err = None
        for _chunk in chunks:
            ok, err = _send_one_tg(sess, url, chat_id, _chunk, timeout)
            if not ok:
                ok_all = False
                if first_err is None:
                    first_err = err
        if ok_all:
            log("TG 通知已发送", "OK")
            return True
        log(f"send_tg_message 失败: {first_err}", "WARN")
        return False
    except Exception as exc:
        msg = f"send_tg_message failed: {type(exc).__name__}: {exc}"
        if "SOCKS" in str(exc) or "socks" in str(exc).lower():
            msg += "（socks5 代理需 pip install pysocks）"
        log(msg, "WARN")
        return False


def test_tg_connection(bot_token=None, proxy=None, timeout=TG_DEFAULT_TIMEOUT, log_fn=None):
    """getMe 验证 bot token + 代理可达性(不发消息，不校验 chat_id)。返回 (ok, message)。
    供 webui「测试 Telegram」按钮调用，与发送走同一套代理/重试逻辑。"""
    log = log_fn or _default_log
    bot_token = bot_token or os.environ.get("TG_BOT_TOKEN", "")
    proxy = proxy or os.environ.get("TG_PROXY", "")
    if not bot_token:
        return False, "未配置 TG_BOT_TOKEN"
    proxy, warn = _normalize_proxy(proxy)
    if warn:
        log(warn, "WARN")
    try:
        sess = _build_session(proxy, timeout)
        url = f"{TG_API_BASE}/bot{bot_token}/getMe"
        r = sess.get(url, timeout=timeout)
        data = r.json()
    except Exception as e:
        tip = "（国内需配 TG_PROXY；socks5 代理需 pip install pysocks）" if not proxy else ""
        return False, f"Telegram 请求失败：{type(e).__name__}: {str(e)[:80]}{tip}"
    if data.get("ok"):
        uname = (data.get("result") or {}).get("username", "")
        via = f"，经代理 {proxy}" if proxy else "（直连）"
        return True, f"Telegram 连通 ✓ bot=@{uname}{via}"
    return False, f"Telegram 返回错误：{data.get('description') or str(data)[:80]}"
