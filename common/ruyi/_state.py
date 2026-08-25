"""包内全局状态与可调常量集中。

所有子模块 `from ._state import ...` 读单例/锁/常量。中性 env 名,无业务依赖。
"""

import os
import threading

# ── 可调常量(读中性 env) ──────────────────────────────────────────

IP_INFO_ENDPOINTS = [
    ("ipwhois", "https://ipwho.is/"),
    ("ipify", "https://api.ipify.org?format=json"),
]

PROXY_PRECHECK_TIMEOUT = float(os.environ.get("RUOYI_PROXY_PRECHECK_TIMEOUT", "30") or "30")
PROXY_IDENTITY_TIMEOUT = float(os.environ.get("RUOYI_PROXY_IDENTITY_TIMEOUT", "12") or "12")
PROXY_IDENTITY_CACHE_TTL = float(os.environ.get("RUOYI_PROXY_IDENTITY_CACHE_TTL", "1800") or "1800")
BROWSER_QUIT_TIMEOUT = float(os.environ.get("RUOYI_BROWSER_QUIT_TIMEOUT", "5") or "5")
RUOYI_FIREFOX_EXIT_WAIT = float(os.environ.get("RUOYI_FIREFOX_EXIT_WAIT", "8") or "8")
RUOYI_FIREFOX_FORCEKILL_WAIT = float(os.environ.get("RUOYI_FIREFOX_FORCEKILL_WAIT", "6") or "6")
RUOYI_PROFILE_STALE_SEC = float(os.environ.get("RUOYI_PROFILE_STALE_SEC", "600") or "600")
RUOYI_PROFILE_CLEANUP_RETRIES = int(os.environ.get("RUOYI_PROFILE_CLEANUP_RETRIES", "5") or "5")
RUOYI_PROFILE_CLEANUP_RETRY_DELAY = float(os.environ.get("RUOYI_PROFILE_CLEANUP_RETRY_DELAY", "0.5") or "0.5")

# 代理预检业务域名列表:默认空(不预检)。调用方(register)按需注入业务值。
PROXY_PRECHECK_TARGETS = ()

# ── 包内全局状态(单例/锁) ────────────────────────────────────────

_CURRENT_IP_INFO = {}
_CURRENT_IP_INFO_LOCK = threading.Lock()

_UA_RR_IDX = 0
_UA_RR_LOCK = threading.Lock()

_PROXY_LIST = None  # 代理池单例

_ACTIVE_RUOYI_PROFILE_DIRS = set()
_ACTIVE_RUOYI_PROFILE_DIRS_LOCK = threading.Lock()

_ACTIVE_BROWSER_PAGES = {}
_ACTIVE_BROWSER_LOCK = threading.Lock()

_SHUTDOWN_HANDLERS_INSTALLED = False

# 可注入位
_FIREFOX_PATH_CACHE = None  # firefox 路径延迟求值缓存
_PROFILE_ROOT = None  # profile 根目录(None 时用默认 cwd/profiles_ruoyi_tmp)
