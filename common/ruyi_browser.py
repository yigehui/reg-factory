"""ruyi 浏览器启动栈 + 反检测 + 资源拦截 —— 兼容薄壳。

本体已迁入自包含独立包 ``common/ruyi/``(launch 模块,零业务依赖,可整目录拷到
别的项目直接用)。本文件保留为 re-export 薄壳,兼容期不删,旧消费者
``from common.ruyi_browser import build_browser_options, after_launch`` 继续可用。

新代码请直接 ``from common.ruyi import build_browser_options, after_launch``。
"""

from common.ruyi.launch import (  # noqa: F401  re-export
    _DEFAULT_UA_POOL,
    HEADLESS_WINDOW_WIDTH,
    HEADLESS_WINDOW_HEIGHT,
    HEADLESS_USER_AGENT,
    _mask_ua,
    _env_bool,
    _try_option_call,
    _all_contexts,
    log,
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
    _RUOYI_RESOURCE_BLOCK_KINDS,
    _RUOYI_RESOURCE_ALLOW_HOST_HINTS,
    _RUOYI_TELEMETRY_HOSTS,
    _RUOYI_RESOURCE_BLOCK_EXTS,
)

__all__ = [
    "_DEFAULT_UA_POOL",
    "HEADLESS_WINDOW_WIDTH",
    "HEADLESS_WINDOW_HEIGHT",
    "HEADLESS_USER_AGENT",
    "_mask_ua",
    "_env_bool",
    "_try_option_call",
    "_all_contexts",
    "log",
    "_apply_ruoyi_browser_ua",
    "_apply_ruoyi_quiet_prefs",
    "_apply_ruoyi_headless_options",
    "build_browser_options",
    "_apply_ruoyi_headless_emulation",
    "_build_headless_patch_js",
    "RUOYI_HEADLESS_PATCH_JS",
    "_ensure_ruoyi_headless_preload",
    "_apply_ruoyi_headless_page_patches",
    "_ruoyi_should_block_resource_request",
    "_start_ruoyi_resource_blocking",
    "after_launch",
    "_RUOYI_RESOURCE_BLOCK_KINDS",
    "_RUOYI_RESOURCE_ALLOW_HOST_HINTS",
    "_RUOYI_TELEMETRY_HOSTS",
    "_RUOYI_RESOURCE_BLOCK_EXTS",
]
