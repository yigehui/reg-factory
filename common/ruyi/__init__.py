"""ruyi 浏览器公共能力:自包含、零业务依赖的独立包。

只 import 标准库 + ruyipage + requests。可整目录拷到别的项目直接用。
提供:代理池 / profile 目录与销毁 / 进程树 kill / 启动栈 / UA 池 / IP 探测。

中性 env 名(RUOYI_*)。业务侧(register_outlook_ruoyi)做兼容层认旧名(OUTLOOK_*)。
"""

# ── _state:常量/全局状态 ──
from . import _state

# ── _logging:骨架 log + 过滤钩子 ──
from ._logging import (
    log,
    debug,
    set_log_level,
    register_log_filter,
    _normalize_log_level,
    _log_level_value,
    LOG_LEVELS,
)

# ── _utils:A 组纯函数 ──
from ._utils import (
    _env_bool,
    _mask_ua,
    _strip_proxy_scheme,
    _proxy_url_to_ruoyi,
    _parse_ruoyi_proxy,
    mask_ruoyi_proxy,
    _proxy_host_port_key,
    _coerce_float_or_none,
    _geo_timezone_from_payload,
    _browser_model_name,
)

# ── firefox:G 组 firefox 路径 + 进程树 kill ──
from .firefox import (
    DEFAULT_RUOYI_FIREFOX,
    _resolve_ruoyi_firefox_path,
    get_firefox_path,
    set_firefox_path,
    _ruoyi_firefox_running_count_by_path,
    _force_kill_ruoyi_firefox,
    _kill_ruoyi_firefox_by_profile,
)

# PEP 562 延迟求值:module.RUOYI_FIREFOX_PATH
def __getattr__(name):
    if name == "RUOYI_FIREFOX_PATH":
        return get_firefox_path()
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


# ── profile:E 组 profile 目录 + 清理 ──
from .profile import (
    get_profile_root,
    set_profile_root,
    _ruoyi_profile_dir,
    _cleanup_ruoyi_profile_root,
    _cleanup_stale_ruoyi_profile_root,
    _cleanup_ruoyi_slot_root,
    _cleanup_ruoyi_run_profile_dir,
)

# ── ua:C 组 UA 池 ──
from .ua import (
    _DEFAULT_UA_POOL,
    HEADLESS_WINDOW_WIDTH,
    HEADLESS_WINDOW_HEIGHT,
    HEADLESS_USER_AGENT,
    _load_ua_pool,
    _pick_user_agent,
)

# ── proxy:D 组代理池 ──
from .proxy import (
    PROXY_FILE,
    PROXY_URL,
    RUOYI_PROXY_SOURCE,
    _parse_proxy_lines,
    parse_proxy_pool,
    fetch_proxy_list_http,
    _proxy_source_label,
    load_proxy_list,
    load_proxy_batch,
    build_proxy_source,
    ConsumableProxyPool,
    get_consumable_proxy_pool,
    set_consumable_proxy_pool,
    get_session_proxy_runtime,
    set_session_proxy_runtime,
    SessionProxyRuntime,
    select_proxy_for_account,
    release_proxy_for_account,
)

# ── probe:B 组 IP/代理探测 ──
from .probe import (
    _proxy_identity_cache_get,
    _proxy_identity_cache_put,
    _probe_proxy_identity,
    _proxy_exit_key,
    _proxy_for_ip_lookup,
    _log_current_ip,
    _probe_proxy_targets,
    _probe_proxy_before_browser,
    PROXY_PRECHECK_URL,
)

# ── launch:F 组 + 启动栈 + 资源拦截 ──
from .launch import (
    _try_option_call,
    _all_contexts,
    _track_browser_page,
    _untrack_browser_page,
    _close_tracked_browser_pages,
    _quit_browser_page,
    _install_shutdown_handlers,
    set_shutdown_handlers_installed,
    _apply_ruoyi_browser_ua,
    _apply_ruoyi_quiet_prefs,
    _apply_ruoyi_headless_options,
    build_browser_options,
    _apply_ruoyi_headless_emulation,
    _build_headless_patch_js,
    RUOYI_HEADLESS_PATCH_JS,
    _ensure_ruoyi_headless_preload,
    _apply_ruoyi_headless_page_patches,
    _ruoyi_should_block_resource_request,
    _start_ruoyi_resource_blocking,
    after_launch,
    # 资源拦截常量(register/unlock re-export 用,单一来源)
    _RUOYI_RESOURCE_BLOCK_KINDS,
    _RUOYI_RESOURCE_ALLOW_HOST_HINTS,
    _RUOYI_TELEMETRY_HOSTS,
    _RUOYI_RESOURCE_BLOCK_EXTS,
)

# ── warmup:H 组首屏预热 + 挑战 cookie 种子池 ──
from .warmup import (
    CHALLENGE_COOKIE_ALLOWLIST,
    SEED_TTL_SEC,
    SEED_MAX_ENTRIES,
    load_seed,
    save_seed,
    warmup,
)

# ── 可调常量(re-export 自 _state,便于 from common.ruyi import RUOYI_FIREFOX_EXIT_WAIT) ──
from ._state import (
    IP_INFO_ENDPOINTS,
    PROXY_PRECHECK_TIMEOUT,
    PROXY_IDENTITY_TIMEOUT,
    PROXY_IDENTITY_CACHE_TTL,
    BROWSER_QUIT_TIMEOUT,
    RUOYI_FIREFOX_EXIT_WAIT,
    RUOYI_FIREFOX_FORCEKILL_WAIT,
    RUOYI_PROFILE_STALE_SEC,
    RUOYI_PROFILE_CLEANUP_RETRIES,
    RUOYI_PROFILE_CLEANUP_RETRY_DELAY,
    PROXY_PRECHECK_TARGETS,
)

__all__ = [
    # state
    "_state",
    "IP_INFO_ENDPOINTS", "PROXY_PRECHECK_TIMEOUT", "PROXY_IDENTITY_TIMEOUT",
    "PROXY_IDENTITY_CACHE_TTL", "BROWSER_QUIT_TIMEOUT", "RUOYI_FIREFOX_EXIT_WAIT",
    "RUOYI_FIREFOX_FORCEKILL_WAIT", "RUOYI_PROFILE_STALE_SEC",
    "RUOYI_PROFILE_CLEANUP_RETRIES", "RUOYI_PROFILE_CLEANUP_RETRY_DELAY",
    "PROXY_PRECHECK_TARGETS",
    # logging
    "log", "debug", "set_log_level", "register_log_filter",
    "_normalize_log_level", "_log_level_value", "LOG_LEVELS",
    # utils
    "_env_bool", "_mask_ua", "_strip_proxy_scheme", "_proxy_url_to_ruoyi",
    "_parse_ruoyi_proxy", "mask_ruoyi_proxy", "_proxy_host_port_key",
    "_coerce_float_or_none", "_geo_timezone_from_payload", "_browser_model_name",
    # firefox
    "DEFAULT_RUOYI_FIREFOX", "RUOYI_FIREFOX_PATH", "_resolve_ruoyi_firefox_path",
    "get_firefox_path", "set_firefox_path", "_ruoyi_firefox_running_count_by_path",
    "_force_kill_ruoyi_firefox", "_kill_ruoyi_firefox_by_profile",
    # profile
    "get_profile_root", "set_profile_root", "_ruoyi_profile_dir",
    "_cleanup_ruoyi_profile_root", "_cleanup_stale_ruoyi_profile_root",
    "_cleanup_ruoyi_slot_root", "_cleanup_ruoyi_run_profile_dir",
    # ua
    "_DEFAULT_UA_POOL", "HEADLESS_WINDOW_WIDTH", "HEADLESS_WINDOW_HEIGHT",
    "HEADLESS_USER_AGENT", "_load_ua_pool", "_pick_user_agent",
    # proxy
    "PROXY_FILE", "PROXY_URL", "RUOYI_PROXY_SOURCE", "_parse_proxy_lines",
    "parse_proxy_pool", "fetch_proxy_list_http", "_proxy_source_label",
    "load_proxy_list", "load_proxy_batch", "build_proxy_source",
    "ConsumableProxyPool", "get_consumable_proxy_pool", "set_consumable_proxy_pool",
    "get_session_proxy_runtime", "set_session_proxy_runtime", "SessionProxyRuntime",
    "select_proxy_for_account", "release_proxy_for_account",
    # probe
    "_proxy_identity_cache_get", "_proxy_identity_cache_put",
    "_probe_proxy_identity", "_proxy_exit_key", "_proxy_for_ip_lookup",
    "_log_current_ip", "_probe_proxy_targets", "_probe_proxy_before_browser",
    "PROXY_PRECHECK_URL",
    # launch
    "_try_option_call", "_all_contexts", "_track_browser_page",
    "_untrack_browser_page", "_close_tracked_browser_pages", "_quit_browser_page",
    "_install_shutdown_handlers", "set_shutdown_handlers_installed",
    "_apply_ruoyi_browser_ua", "_apply_ruoyi_quiet_prefs",
    "_apply_ruoyi_headless_options", "build_browser_options",
    "_apply_ruoyi_headless_emulation", "_build_headless_patch_js",
    "RUOYI_HEADLESS_PATCH_JS", "_ensure_ruoyi_headless_preload",
    "_apply_ruoyi_headless_page_patches", "_ruoyi_should_block_resource_request",
    "_start_ruoyi_resource_blocking", "after_launch",
    "_RUOYI_RESOURCE_BLOCK_KINDS", "_RUOYI_RESOURCE_ALLOW_HOST_HINTS",
    "_RUOYI_TELEMETRY_HOSTS", "_RUOYI_RESOURCE_BLOCK_EXTS",
    # warmup
    "CHALLENGE_COOKIE_ALLOWLIST", "SEED_TTL_SEC", "SEED_MAX_ENTRIES",
    "load_seed", "save_seed", "warmup",
]
