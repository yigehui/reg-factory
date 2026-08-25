"""C 组:UA 池。

10 条 UA 常量逐字复制自 common/ruyi_browser(单一来源在此)。env 名参数化,
新包默认读中性名 RUOYI_UA_POOL。HEADLESS_* 读 RUOYI_HEADLESS_* 中性名。
"""

import os
import re

from . import _state

_DEFAULT_UA_POOL = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:155.0) Gecko/20100101 Firefox/155.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:154.0) Gecko/20100101 Firefox/154.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:153.0) Gecko/20100101 Firefox/153.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:152.0) Gecko/20100101 Firefox/152.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:151.0) Gecko/20100101 Firefox/151.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:150.0) Gecko/20100101 Firefox/150.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:149.0) Gecko/20100101 Firefox/149.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:148.0) Gecko/20100101 Firefox/148.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:147.0) Gecko/20100101 Firefox/147.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:146.0) Gecko/20100101 Firefox/146.0",
)

HEADLESS_WINDOW_WIDTH = int(os.environ.get("RUOYI_HEADLESS_WIDTH", "1280") or "1280")
HEADLESS_WINDOW_HEIGHT = int(os.environ.get("RUOYI_HEADLESS_HEIGHT", "800") or "800")
HEADLESS_USER_AGENT = os.environ.get("RUOYI_HEADLESS_UA", _DEFAULT_UA_POOL[0])


def _load_ua_pool(env_name="RUOYI_UA_POOL"):
    """UA 池:env 优先,否则内置 10 条 Firefox。"""
    raw = str(os.environ.get(env_name, "") or "").strip()
    pool = []
    if raw:
        for part in re.split(r"[\n|]+", raw):
            ua = part.strip()
            if ua:
                pool.append(ua)
    if not pool:
        # 单条 HEADLESS_UA 放队首,再拼内置池去重
        seed = str(HEADLESS_USER_AGENT or "").strip()
        seen = set()
        for ua in (([seed] if seed else []) + list(_DEFAULT_UA_POOL)):
            if ua and ua not in seen:
                seen.add(ua)
                pool.append(ua)
    return pool or list(_DEFAULT_UA_POOL)


def _pick_user_agent(idx=None, env_name="RUOYI_UA_POOL"):
    """按任务序号轮询 UA;idx 为空时线程安全全局自增。"""
    pool = _load_ua_pool(env_name)
    if not pool:
        return HEADLESS_USER_AGENT

    if idx is not None:
        try:
            n = int(idx)
        except Exception:
            n = 1
        n = max(1, n)
        ua = pool[(n - 1) % len(pool)]
        if n <= len(pool) or str(os.environ.get(env_name, "") or "").strip():
            return ua
        cycle = (n - 1) // len(pool)
        m = re.search(r"rv:(\d+)\.0\).*Firefox/(\d+)\.0", ua)
        if not m:
            return ua
        ver = max(115, int(m.group(1)) - (cycle * len(pool)))
        ua = re.sub(r"rv:\d+\.0", f"rv:{ver}.0", ua, count=1)
        ua = re.sub(r"Firefox/\d+\.0", f"Firefox/{ver}.0", ua, count=1)
        return ua

    with _state._UA_RR_LOCK:
        ua = pool[_state._UA_RR_IDX % len(pool)]
        _state._UA_RR_IDX += 1
        return ua
