"""H 组:首屏 warmup(首页预热)+ 挑战 cookie 种子池。

解决"新开 ruyi 浏览器零 cookie 直达注册页被挑战系统拦":
1) warmup():先落地目标站首页停留数秒,让站点自然种下第一方挑战 cookie
   (_pxhd / __cf_bm / canary 等),再进注册页;
2) 种子池:注册成功后把"放行类"挑战 cookie 连同 (proxy, ua, 采集时间)
   存进 cookies/<platform>/seeds.json;后续同代理+同 UA 的新实例启动后
   回灌,等价"探路 cookie 共享"--同 IP 段并发里第一个过挑战,后面复用。

铁律:
- 只回灌/入池白名单挑战 cookie(匿名,换账号无影响);登录态 cookie
  绝不入池(白名单制,ESTSAUTH/session 之类天然进不来)。
- 挑战 cookie 绑出口 IP + UA + 半衰期约 30-60 分钟,种子必须
  (proxy, ua) 全等 + TTL 内才可用,否则灌进去立刻失效还留指纹矛盾。

自包含:只依赖标准库;page 走 duck-type(get/wait_loading/set_cookies/
get_cookies),与 ruyipage FirefoxPage 兼容。
"""

import json
import os
import random
import threading
import time
import urllib.parse

from ._logging import log

# 白名单:匿名挑战/放行 cookie。只认名单内,其余一律不入池(登录态防串)。
CHALLENGE_COOKIE_ALLOWLIST = frozenset({
    # Cloudflare
    "cf_clearance", "__cf_bm", "_cfuvid",
    # PerimeterX
    "_px", "_px2", "_px3", "_pxhd", "_pxde", "_pxvid",
    # Microsoft 挑战相关 + 匿名统计(非登录态)
    "canary", "brcap", "MUID", "MSCC",
    # MS 登录流程匿名 cookie(实测 login.live.com 首页自然种下,无凭据绑定)
    "MSPRequ", "MSPOK", "OParams", "uaid", "ChainedAuthProbe",
})

# 种子有效期(秒):挑战 cookie 半衰期 30-60min,默认 45min 过期即弃。
SEED_TTL_SEC = 2700.0
# 种子池条目上限(按 key 去重后),防文件无限膨胀。
SEED_MAX_ENTRIES = 200

_SEED_FILE_LOCK = threading.Lock()


def _seed_file(platform):
    """种子池文件路径:cwd/cookies/<platform>/seeds.json(与 common/cookies.py 同根)。"""
    return os.path.join(os.getcwd(), "cookies", str(platform or "default").strip() or "default", "seeds.json")


def _load_seed_entries(path):
    """读种子池(坏文件/缺文件一律返回 [],不影响主流程)。"""
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return []
    entries = data.get("entries") if isinstance(data, dict) else None
    if not isinstance(entries, list):
        return []
    return [e for e in entries if isinstance(e, dict)]


def _entry_fresh(entry, now, ttl):
    try:
        ts = float(entry.get("ts") or 0.0)
    except Exception:
        return False
    return ttl <= 0 or (now - ts) <= ttl


def _entry_key(entry):
    return (str(entry.get("proxy") or "").strip(), str(entry.get("ua") or "").strip())


def _normalize_cookie(obj):
    """CookieInfo / 原始 dict -> set_cookies 兼容的纯 dict;白名单外返回 None。"""
    raw = getattr(obj, "raw", None)
    if not isinstance(raw, dict):
        raw = obj if isinstance(obj, dict) else None
    if not raw:
        return None
    name = str(raw.get("name") or "").strip()
    if not name or name not in CHALLENGE_COOKIE_ALLOWLIST:
        return None
    value = raw.get("value", "")
    if isinstance(value, dict):
        value = value.get("value", "")
    cookie = {"name": name, "value": str(value)}
    domain = str(raw.get("domain") or "").strip()
    if domain:
        cookie["domain"] = domain
    for key in ("path", "secure", "sameSite", "expiry"):
        val = raw.get(key)
        if val is None:
            continue
        if key == "secure" and isinstance(val, str):
            continue
        try:
            cookie[key] = val
        except Exception:
            continue
    return cookie


def load_seed(platform, proxy, user_agent, max_age_sec=None, seed_file=None):
    """取匹配 (proxy, ua) 且未过期的种子 cookie 列表;无则 []。

    只认最新一条同 key 种子;过期即返回 [](过期种子灌入弊大于利)。"""
    path = seed_file or _seed_file(platform)
    ttl = SEED_TTL_SEC if max_age_sec is None else float(max_age_sec or 0.0)
    key = (str(proxy or "").strip(), str(user_agent or "").strip())
    if not key[0] or not key[1]:
        return []
    with _SEED_FILE_LOCK:
        entries = _load_seed_entries(path)
    now = time.time()
    matched = None
    for entry in reversed(entries):
        if _entry_key(entry) != key:
            continue
        matched = entry
        break
    if matched is None:
        return []
    if not _entry_fresh(matched, now, ttl):
        return []
    cookies = matched.get("cookies")
    if not isinstance(cookies, list):
        return []
    return [dict(c) for c in cookies if isinstance(c, dict)]


def save_seed(page, platform, proxy, user_agent, tag="", max_age_sec=None, seed_file=None):
    """注册成功后收集当前页 cookie,白名单过滤入池。返回入池条数(0=没入库)。"""
    if page is None:
        return 0
    key_proxy = str(proxy or "").strip()
    key_ua = str(user_agent or "").strip()
    if not key_proxy or not key_ua:
        log(f"  {tag} 种子入池跳过:缺 proxy/ua 键,无法绑定复用条件", "WARN")
        return 0
    get_cookies = getattr(page, "get_cookies", None)
    if not callable(get_cookies):
        log(f"  {tag} 种子入池跳过:page 无 get_cookies API", "WARN")
        return 0
    raw_cookies = []
    try:
        raw_cookies = get_cookies(True) or []
    except TypeError:
        try:
            raw_cookies = get_cookies() or []
        except Exception as exc:
            log(f"  {tag} 种子采集失败: {type(exc).__name__}: {exc}", "WARN")
            return 0
    except Exception as exc:
        log(f"  {tag} 种子采集失败: {type(exc).__name__}: {exc}", "WARN")
        return 0

    picked = []
    for obj in raw_cookies:
        cookie = _normalize_cookie(obj)
        if cookie:
            picked.append(cookie)
    if not picked:
        log(f"  {tag} 种子入池跳过:白名单挑战 cookie 一条都没有({len(raw_cookies)} cookies)", "DEBUG")
        return 0

    path = seed_file or _seed_file(platform)
    ttl = SEED_TTL_SEC if max_age_sec is None else float(max_age_sec or 0.0)
    now = time.time()
    entry = {"proxy": key_proxy, "ua": key_ua, "ts": now, "cookies": picked}

    with _SEED_FILE_LOCK:
        entries = _load_seed_entries(path)
        entries = [e for e in entries if _entry_fresh(e, now, ttl)]
        entries = [e for e in entries if _entry_key(e) != (key_proxy, key_ua)]
        entries.append(entry)
        if len(entries) > SEED_MAX_ENTRIES:
            entries = entries[-SEED_MAX_ENTRIES:]
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            tmp_path = path + ".tmp"
            with open(tmp_path, "w", encoding="utf-8") as f:
                json.dump({"version": 1, "entries": entries}, f, ensure_ascii=False, indent=2)
            os.replace(tmp_path, path)
        except Exception as exc:
            log(f"  {tag} 种子池写入失败: {type(exc).__name__}: {exc}", "WARN")
            return 0

    log(f"  {tag} 种子入池 {len(picked)} 条 -> {path}")
    return len(picked)


def warmup(page, tag="", *, home_url, seed_cookies=None, min_wait=2.0, max_wait=5.0):
    """启动后首屏预热:落地首页 -> 随机停留 -> 回灌种子 cookie。

    page 已在首页域名上(直接以首页 URL 启动)则跳过导航只做停留+回灌。
    返回 True 表示预热完成;导航失败返回 False(调用方继续原流程,不阻断)。"""
    if page is None or not str(home_url or "").strip():
        return False
    home_host = urllib.parse.urlsplit(str(home_url)).netloc.lower()
    current_host = ""
    try:
        current_host = urllib.parse.urlsplit(str(getattr(page, "url", "") or "")).netloc.lower()
    except Exception:
        pass
    already_home = bool(home_host) and bool(current_host) and (
        current_host == home_host or current_host.endswith("." + home_host)
    )
    if not already_home:
        try:
            page.get(home_url)
        except Exception as exc:
            log(f"  {tag} warmup 导航失败({home_url}): {type(exc).__name__}: {exc}", "WARN")
            return False
    try:
        wait_loading = getattr(page, "wait_loading", None)
        if callable(wait_loading):
            wait_loading(15)
    except Exception:
        pass
    wait_sec = random.uniform(max(0.0, min_wait), max(min_wait, max_wait))
    time.sleep(wait_sec)
    log(f"  {tag} warmup 首页就位 {current_host or home_host}, 停留 {wait_sec:.1f}s 种下第一方 cookie")

    if seed_cookies:
        try:
            page.set_cookies([dict(c) for c in seed_cookies])
            log(f"  {tag} warmup 回灌种子 cookie {len(seed_cookies)} 条(同代理同UA)")
        except Exception as exc:
            log(f"  {tag} warmup 回灌种子失败: {type(exc).__name__}: {exc}", "WARN")
    return True
