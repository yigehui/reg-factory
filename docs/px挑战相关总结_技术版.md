# px挑战相关总结（技术版）

> 适用范围：`D:/officeProject/yigehui/reg-factory/register_outlook_ruoyi.py` 当前工作树  
> 目标读者：内部研发、自动化维护、问题排查、方案复用人员  
> 文档定位：面向实现细节，强调“当前代码实际上做了什么、为什么这么做、关键参数在哪里、出问题先看哪一层”。

---

# 1. 文档范围与背景

当前 ruoyi 链路中的 PX（PerimeterX）处理，不是单独一个“长按验证码”的函数，而是四层能力叠加：

1. **页面/浏览器一致性层**：尽量让 Firefox headless 环境在常见探针上更像真实用户环境。
2. **挑战识别层**：在主文档和 iframe 中稳定找出真正可按的 hold target。
3. **动作执行层**：用每会话不同的短轨迹和按压时长联动，替代固定中心点机械长按。
4. **按后状态机层**：一次按压后进入等待/观察，不是立刻判失败。

当前工作树还额外叠加了：

- **资源拦截层（可选）**：屏蔽图片/字体/媒体，减少噪音与加载负担。
- **复盘层**：截图、HTML、frame 内容、HAR、PX metrics。
- **WebUI 开关层**：把关键参数暴露给面板，便于调试。

---

# 2. 相关文件

## 2.1 核心实现文件

- `D:/officeProject/yigehui/reg-factory/register_outlook_ruoyi.py`
- `D:/officeProject/yigehui/reg-factory/webui/scripts.py`
- `D:/officeProject/yigehui/reg-factory/webui/server.py`

## 2.2 配套测试文件

- `D:/officeProject/yigehui/reg-factory/test_ruoyi_px_hold_early_release.py`
- `D:/officeProject/yigehui/reg-factory/test_ruoyi_px_press_screenshots.py`
- `D:/officeProject/yigehui/reg-factory/test_ruoyi_captcha_waits.py`
- `D:/officeProject/yigehui/reg-factory/test_ruoyi_geo_emulation.py`
- `D:/officeProject/yigehui/reg-factory/test_webui_ruoyi_resource_block_flag.py`

## 2.3 相关历史提交（演进脉络）

- `7d8d0ec` `优化 ruoyi 验证码等待与 profile 清理策略`
- `52d5ead` `调整 ruoyi 并发授权与验证码等待日志`
- `c64ac05` `chore: checkpoint ruoyi headless and px handling`

> 说明：本文不是按某一个 commit 展开，而是以**当前工作树最终状态**为准。以上 commit 主要用于追溯演进来源。

---

# 3. 外部入口与参数

## 3.1 CLI 入口

核心入口文件：

- `register_outlook_ruoyi.py`

PX 相关参数：

- `--max-press`：PX 最大按压次数
- `--headless`：无头模式
- `--px-press-screenshots`：失败时补充 PX 相关截图
- `--block-resources`：启用图片/字体/媒体拦截
- `--ua-pool`：UA 池
- `--timeout`：单号总超时

关键代码：

```python
ap.add_argument("--headless", action="store_true", help="无头模式")

ap.add_argument("--block-resources", action="store_true",
                default=_env_bool("OUTLOOK_RUOYI_BLOCK_RESOURCES", False),
                help="屏蔽 image/font/media 资源请求")

ap.add_argument("--max-press", default=os.environ.get("OUTLOOK_REG_MAX_PRESS", "5"), help="按住次数上限")

ap.add_argument("--px-press-screenshots", action=argparse.BooleanOptionalAction,
                default=_env_bool("OUTLOOK_PX_PRESS_SCREENSHOTS", False),
                help="保存 ruoyi PX/失败相关截图")
```

## 3.2 WebUI 入口

WebUI 已暴露 PX 相关开关，便于现场调试：

- `max-press`
- `headless`
- `block-resources`
- `px-press-screenshots`

关键代码：

```python
{"flag": "--max-press", "type": "str", "default": "5", "help": "验证码按住次数上限"},
{"flag": "--headless", "type": "bool", "default": False, "help": "无头模式"},
{"flag": "--block-resources", "type": "bool", "default": False, "help": "屏蔽 image/font/media 请求"},
{"flag": "--px-press-screenshots", "type": "bool", "default": False,
 "help": "保存 ruoyi PX/失败相关截图"},
```

---

# 4. 总体执行链路

当前代码在进入浏览器主流程后，PX 相关初始化顺序大致是：

1. 选择 UA
2. 初始化 Firefox options / profile / per-tab 代理
3. 如为 headless，则应用 headless 参数与首选项
4. 启动 Firefox
5. 为 page 生成每会话 motion profile
6. 如有代理，则做 proxy exit 的 timezone/geolocation 对齐
7. 如启用资源拦截，则开启 request intercept
8. 如是 headless，则做页面层 patch（含 preload）
9. 进入 signup 页面
10. 在页面流程中识别 challenge、执行按压、进入按后状态机

关键代码：

```python
user_agent = _pick_user_agent(idx)
_apply_ruoyi_browser_ua(tb, tag, user_agent)

if is_headless:
    _apply_ruoyi_headless_options(tb, tag, user_agent=user_agent)
    tb.headless(True)

page = browser_page
setattr(page, "_ruoyi_px_motion_profile", _new_ruoyi_px_motion_profile())

if proxy_pool:
    _apply_ruoyi_proxy_geo_emulation(page, proxy_pool, tag)

if bool(getattr(opts, "block_resources", False)):
    _start_ruoyi_resource_blocking(page, tag)

if is_headless and not signup_opened:
    _apply_ruoyi_headless_page_patches(page, tag, log_once=True, user_agent=user_agent)
```

---

# 5. 参数与节奏控制

当前 PX 节奏是显式参数化的，不是写死一段“看到就按”。

## 5.1 关键时间参数

```python
INITIAL_PRESS_DELAY = 5
PX_HOLD_EARLY_RELEASE_AFTER = 5.0
PX_HOLD_EARLY_RELEASE_INTERVAL = 0.5
PX_HOLD_SECONDS_MIN = 10.5
PX_HOLD_SECONDS_MAX = 11.0
POST_MAX_PRESS_WAIT = 5
POST_PRESS_LOADING_CHECK = 8
CAPTCHA_STATE_TIMEOUT = float(os.environ.get("OUTLOOK_RUOYI_CAPTCHA_STATE_TIMEOUT", "20") or "20")
POST_PRESS_RETRY_GAP_MIN = 2.0
POST_PRESS_RETRY_GAP_MAX = 4.0
```

## 5.2 含义解释

### INITIAL_PRESS_DELAY

challenge 变成“可按”后，不会立刻触发按压，而是先等 5 秒。

目的：

- 降低“元素刚可见就秒按”的机械感
- 给页面一点稳定时间
- 避免刚渲染完目标还没完全稳定时去按

### PX_HOLD_SECONDS_MIN / MAX

单次基准 hold 时长在 10.5~11.0 秒之间，属于较长 hold。

注意：

- 这不是最终一定 hold 满的时长
- 如果 label 先隐藏，可以提前释放

### PX_HOLD_EARLY_RELEASE_AFTER / INTERVAL

hold 开始 5 秒后，每 0.5 秒轮询一次 label 隐藏状态。

目的是：

- 不必一直按到理论满时长
- 如果 challenge 已进入下一阶段，则尽早释放

### POST_MAX_PRESS_WAIT

达到 `max_press` 后，不是立即判输，而是再给 5 秒看是否跳转或进入下一步。

### POST_PRESS_LOADING_CHECK

一次按压后，如果在 8 秒内一直没有进入 loading / validating / reappear 等状态，才认为当前这一按很可能没起作用。

### CAPTCHA_STATE_TIMEOUT

任何单个 captcha 中间状态（如 validating、reappear wait）最长等待 20 秒，超时后退出，避免僵死。

---

# 6. 指纹一致性层

这一层的核心目标不是“做全功能指纹伪装框架”，而是：

- 保持 Firefox 身份
- 让 headless Firefox 在常见 PX / 通用自动化探针前，不要过于裸露
- 尽量让 `UA / locale / language / screen / timezone / geolocation` 互相协调

## 6.1 UA 池与每实例 UA

当前实例不会全部使用同一条默认 UA，而是通过 UA 池轮换，减少并发实例完全同模版的风险。

代码：

```python
_DEFAULT_UA_POOL = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:151.0) Gecko/20100101 Firefox/151.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:150.0) Gecko/20100101 Firefox/150.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:149.0) Gecko/20100101 Firefox/149.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:148.0) Gecko/20100101 Firefox/148.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:147.0) Gecko/20100101 Firefox/147.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:146.0) Gecko/20100101 Firefox/146.0",
)
```

```python
user_agent = _pick_user_agent(idx)
log(f"  {tag} ua pool pick -> {_mask_ua(user_agent)} ({user_agent[:72]}...)")
```

## 6.2 浏览器级 UA override

浏览器启动前，优先把 UA 写入 Firefox preference 或其他兼容 option API。

代码：

```python
def _apply_ruoyi_browser_ua(tb, tag, user_agent):
    ok, method = _try_option_call(
        tb,
        ("set_preference", "set_pref", "set_prefs", "set_option"),
        "general.useragent.override",
        ua,
    )
```

意义：

- 尽量让浏览器层和页面层的 UA 一致
- 减少仅 JS 覆写 `navigator.userAgent` 导致的上下游不一致

## 6.3 Headless 启动参数与 preference

headless 下除了 `tb.headless(True)`，还会主动补以下信息：

- width
- height
- window-size
- lang
- `general.useragent.override`
- `intl.accept_languages`
- `dom.webdriver.enabled = False`
- `privacy.resistFingerprinting = False`

代码：

```python
arg_pairs = (
    ("--width", str(HEADLESS_WINDOW_WIDTH)),
    ("--height", str(HEADLESS_WINDOW_HEIGHT)),
    ("--window-size", f"{HEADLESS_WINDOW_WIDTH},{HEADLESS_WINDOW_HEIGHT}"),
    ("--lang", "en-US"),
)

prefs = {
    "general.useragent.override": ua,
    "intl.accept_languages": "en-US,en",
    "dom.webdriver.enabled": False,
    "privacy.resistFingerprinting": False,
}
```

## 6.4 Emulation 统一

如果 `page.emulation` 接口可用，会继续补：

- locale
- screen size
- device scale factor
- user-agent
- `Accept-Language`

代码：

```python
calls = (
    ("set_locale", ("en-US",)),
    ("set_screen_size", (1920, 1080)),
    ("set_device_scale_factor", (1,)),
    ("set_user_agent", (ua,)),
    ("set_extra_headers", ({"Accept-Language": "en-US,en;q=0.9"},)),
)
```

## 6.5 Headless 页面 patch

这一层处理最常见的 headless 暴露点：

- `document.hidden`
- `document.visibilityState`
- `document.hasFocus()`
- `navigator.webdriver`
- `navigator.userAgent`
- `navigator.appVersion`
- `navigator.languages`
- `screen.*`
- 删除 `window.__playwright` / `window.__pwInitScripts` / `cdc_*`

代码：

```javascript
try {
  Object.defineProperty(document, 'hidden', {get: () => false, configurable: true});
  Object.defineProperty(document, 'visibilityState', {get: () => 'visible', configurable: true});
} catch(e) {}

try { document.hasFocus = function(){ return true; }; } catch(e) {}

try {
  Object.defineProperty(navigator, 'webdriver', {get: () => undefined, configurable: true});
} catch(e) {}

try {
  Object.defineProperty(navigator, 'userAgent', {get: () => UA, configurable: true});
  Object.defineProperty(navigator, 'appVersion', {get: () => APP_VERSION, configurable: true});
} catch(e) {}

try {
  Object.defineProperty(navigator, 'languages', {get: () => ['en-US', 'en'], configurable: true});
} catch(e) {}

try {
  delete window.__playwright;
  delete window.__pwInitScripts;
  for (const k of Object.getOwnPropertyNames(window)) {
    if (/^cdc_|webdriver|selenium|driver/i.test(k)) {
      try { delete window[k]; } catch(e) {}
    }
  }
} catch(e) {}
```

## 6.6 Preload + 运行时双层注入

当前不是只在页面打开后打一遍 JS，而是优先走 preload，再补一轮运行时 patch。

代码：

```python
def _ensure_ruoyi_headless_preload(page, tag=None, log_once=False, user_agent=None):
    add_script = getattr(page, "add_preload_script", None)
    preload_js = f"() => {{{_build_headless_patch_js(ua)}}}"
    add_script(preload_js)
```

```python
def _apply_ruoyi_headless_page_patches(page, tag=None, log_once=False, user_agent=None):
    _apply_ruoyi_headless_emulation(page, tag, log_once=False, user_agent=ua)
    _ensure_ruoyi_headless_preload(page, tag, log_once=log_once, user_agent=ua)
    patch_js = _build_headless_patch_js(ua)
```

## 6.7 timezone + geolocation 对齐代理出口

当前 proxy exit 的真实 IP 会被探测出来，再用其 `timezone / latitude / longitude` 去修页面环境。

代码：

```python
def _apply_ruoyi_proxy_geo_emulation(page, proxy_pool, tag):
    identity = _probe_proxy_identity(proxy_pool, timeout=15)
    timezone_id = str(identity.get("timezone") or "").strip()
    latitude = identity.get("latitude")
    longitude = identity.get("longitude")

    if timezone_id:
        emu.set_timezone(timezone_id)

    if latitude is not None and longitude is not None:
        emu.set_geolocation(latitude, longitude, accuracy=100)
```

目的：

- 避免“代理在美国，但页面时区/定位像亚洲”这种明显不一致
- 尽量让 challenge 判断链路看到的环境更统一

---

# 7. PX 挑战识别层

这一层解决的是：**challenge 到底在哪、该按哪个元素、如何识别它已经进入 validating/loading**。

## 7.1 多语言 hold / challenge 关键词

当前不是只靠英文关键字，而是带多语言匹配。

代码：

```python
hold_patterns = [
    r"press\s*(?:and|&)?\s*hold",
    r"long\s*press",
    "長押し",
    "按住|长按",
    r"appuyer\s*et\s*maintenir",
    r"halten",
]
challenge_patterns = [
    r"challenge|human|verify|verification",
    "チャレンジ|ヒューマン|検証|確認",
    "验证|驗證|人机|人機",
    "défi|vérifi",
    "prüfung|verifiz",
]
```

## 7.2 主目标查找 `_find_hold_target()`

候选来源包括：

1. `#px-captcha`
2. 带 `aria-label` 的 button / input / a[role=button]
3. 文本标签向上追到祖先 button
4. 纯按钮文本包含 hold 的按钮

返回的 target 会带上：

- `x / y`
- `left / top / width / height`
- `text`
- `id / tag / role / ariaLabel`
- `quality`
- `source`
- `holdLabelId`

代码：

```javascript
const addCandidate = (el, quality, source, labelEl = null) => {
  if (!el || seen.has(el) || !visible(el) || !buttonish(el)) return;
  seen.add(el);
  candidates.push(pack(el, quality, source, labelEl));
};

if (px && visible(px) && buttonish(px)) addCandidate(px, 0, '#px-captcha');
...
if (hasHold && hasChallenge) addCandidate(el, 1, 'aria-label:hold+challenge');
else if (hasHold) addCandidate(el, 2, 'aria-label:hold');
```

## 7.3 iframe 优先策略与 fallback

`_find_hold_context()` 会遍历所有 context，并优先：

- `hsprotect.net`
- `arkose`
- `funcaptcha`

如果没有明确命中的 frame，则退化为用可见 iframe box 估算一个目标区域。

代码：

```python
def _find_hold_context(page, min_quality=5):
    for ctx in _all_contexts(page):
        target = _find_hold_target(ctx)
        if any(k in url for k in ["hsprotect.net", "arkose", "funcaptcha"]):
            frame_bonus = -10
    ...
    box = _find_hsprotect_iframe_box(page)
    if box is not None:
        return page, box
```

## 7.4 challenge 可见性判断

当前 challenge 是否“存在”，不是只依赖一个 DOM 选择器，而是组合判断：

- 有没有高质量 hold target
- 页面里有没有 challenge iframe hint
- body 文案里有没有 `press and hold / captcha / perimeterx / 按住 / 长按` 等字样

代码：

```python
def _captcha_visible(page):
    for ctx in _all_contexts(page):
        t = _find_hold_target(ctx)
        if t and int(t.get("quality", 9)) <= 5:
            return True
    if _context_has_iframe_hint(page):
        return True
```

## 7.5 validating 判断

当前 validating 判断有两条主要路径：

1. `_px_captcha_completed_wait(page)`：看 `human challenge completed, please wait`
2. `_captcha_is_validating(page)`：
   - 页面有 challenge iframe，但暂时找不到 hold target
   - 或 body 文案中出现 `verifying / checking / loading / please wait / just a moment`

代码：

```python
def _px_captcha_completed_wait(page):
    phrase = "human challenge completed, please wait"
```

```python
if _px_captcha_completed_wait(page):
    return True
if iframe_hint and not _find_hold_context(page)[1]:
    return True
```

## 7.6 hold label 定位与状态读取

当前按压时不是盲按到底，而是会尝试找到和 target 最近的 hold label，并读取其 `display / visibility / opacity` 等状态。

用途：

- 判断当前 hold 是否已经进入下一阶段
- 如果 label 已 `display:none`，说明挑战界面发生变化，可以提前松开

代码：

```python
def _resolve_hold_label_for_target(ctx, target):
    hits.sort((a, b) => (a.distance - b.distance) || ((b.id ? 1 : 0) - (a.id ? 1 : 0)))
    return hits[0]
```

```python
def _px_hold_instruction_state(ctx, hold_p_id=None):
    return {
      id: el.id || '',
      text: raw.trim().slice(0, 80),
      display: s.display || '',
      visibility: s.visibility || '',
      opacity: s.opacity || '',
      hidden: s.display === 'none'
    }
```

---

# 8. PX 动作执行层

这一层对应 `_perform_hold()` 及其前置点位计算。

## 8.1 每会话 motion profile

每个 page 启动后都会生成一套独立运动参数，而不是所有实例共用同一条轨迹模版。

代码：

```python
setattr(page, "_ruoyi_px_motion_profile", _new_ruoyi_px_motion_profile())
```

```python
def _new_ruoyi_px_motion_profile():
    return {
        "press_ratio_x": random.uniform(0.42, 0.58),
        "press_ratio_y": random.uniform(0.48, 0.64),
        "settle_ratio_x": random.uniform(0.44, 0.56),
        "settle_ratio_y": random.uniform(0.50, 0.62),
        "micro_ratio_x": random.uniform(0.45, 0.57),
        "micro_ratio_y": random.uniform(0.49, 0.63),
        "jitter_x": random.uniform(1.0, 2.8),
        "jitter_y": random.uniform(0.8, 2.0),
        "settle_jitter_x": random.uniform(0.4, 1.0),
        "settle_jitter_y": random.uniform(0.3, 0.9),
        "micro_jitter_x": random.uniform(0.2, 0.6),
        "micro_jitter_y": random.uniform(0.2, 0.6),
        "lead_dx": random.choice((-1, 1)) * random.uniform(8.0, 16.0),
        "lead_dy": random.uniform(-6.0, 6.0),
        "move_ratio": random.uniform(0.024, 0.036),
        "lead_ratio": random.uniform(0.42, 0.56),
        "settle_ratio": random.uniform(0.20, 0.28),
        "micro_ratio": random.uniform(0.10, 0.18),
        "hover_ratio": random.uniform(0.008, 0.016),
    }
```

## 8.2 非中心点取点

`_ruoyi_point_inside_target()` 不会固定取中心点，而是按矩形比例 + jitter 算点。

代码：

```python
def _ruoyi_point_inside_target(target, ratio_x=0.5, ratio_y=0.55, jitter_x=0.0, jitter_y=0.0):
    left = float((target or {}).get("left", base_x) or base_x)
    top = float((target or {}).get("top", base_y) or base_y)
    width = max(1.0, width)
    height = max(1.0, height)
    ratio_x = min(0.9, max(0.1, float(ratio_x)))
    ratio_y = min(0.9, max(0.1, float(ratio_y)))
```

意义：

- 避免总落在绝对中心点
- 同类挑战里不同会话坐标更自然

## 8.3 轨迹结构：lead → settle → micro → final

按压前移动不是单段，而是四段短轨迹：

1. `lead`
2. `settle`
3. `micro`
4. `final`

代码：

```python
move_ms = int(min(420, max(200, hold_sec * float(motion_profile.get("move_ratio", 0.03)) * 1000.0)))
lead_ms = int(move_ms * float(motion_profile.get("lead_ratio", 0.48)))
settle_ms = int(move_ms * float(motion_profile.get("settle_ratio", 0.24)))
micro_ms = int(move_ms * float(motion_profile.get("micro_ratio", 0.14)))
final_ms = move_ms - lead_ms - settle_ms - micro_ms
```

```python
actions.move_to({"x": lead_x, "y": lead_y}, duration=lead_ms)
actions.move_to({"x": settle_x, "y": settle_y}, duration=settle_ms)
actions.move_to({"x": micro_x, "y": micro_y}, duration=micro_ms)
actions.move_to({"x": cx, "y": cy}, duration=final_ms)
if pre_hover_sec > 0:
    actions.wait(pre_hover_sec)
actions.hold().perform()
```

## 8.4 轨迹与 hold 时长联动

这里不是“轨迹时间固定 + hold 时间固定”，而是：

- `hold_sec` 先随机
- `move_ms` 再按 `hold_sec * move_ratio` 派生

这让动作整体更连贯，而不是割裂的“先瞬移再长按”。

## 8.5 提前释放

hold 开始后，5 秒后开始轮询 label，一旦 `display:none` 就跳出并 release。

代码：

```python
next_check = hold_started + PX_HOLD_EARLY_RELEASE_AFTER
...
release_state = _px_hold_instruction_state(state_ctx, hold_p_id)
if release_state and str(release_state.get("display") or "").strip().lower() == "none":
    break
```

## 8.6 label 刷新

如果原始 label 读不到，会重新找 context / target / label，以适应 DOM/iframe 变化。

代码：

```python
if release_state is None:
    refreshed_ctx, refreshed_target = _find_hold_context(page)
    refreshed_label = _resolve_hold_label_for_target(refreshed_ctx, refreshed_target) or {}
    refreshed_id = str(refreshed_label.get("id") or "").strip()
    if refreshed_id and refreshed_id != hold_p_id:
        state_ctx = refreshed_ctx
        hold_p_id = refreshed_id
```

## 8.7 动作异常兜底

如果 hold 执行中抛异常，会尝试 `release_all()`，防止动作状态残留。

代码：

```python
except Exception as exc:
    try:
        actions.release_all()
    except Exception:
        pass
    log(f"  {tag} hold failed: {type(exc).__name__}: {exc}", "WARN")
    return False
```

---

# 9. PX 按后状态机

这一层是当前实现里最容易出 bug、也是最关键的一层。

## 9.1 首次可按前等待

当前策略是：challenge visible + actionable 后，先等 `INITIAL_PRESS_DELAY`，再触发第一次按压。

代码：

```python
if visible and actionable:
    if initial_press_wait_started is None:
        initial_press_wait_started = time.time()
        log(f"  {tag} captcha visible, wait {INITIAL_PRESS_DELAY}s before press")
        time.sleep(INITIAL_PRESS_DELAY)
        continue
```

## 9.2 一次按压成功后进入“等待结果态”

如果 hold 成功，不立即允许下一次 press，而是先进入按后观察状态：

- 记 `px_hold_elapsed`
- 记 `validation_wait_started`
- `awaiting_reappear = True`
- `post_press_started_at = now`

代码：

```python
if hold_elapsed:
    px_hold_elapsed += float(hold_elapsed)
    validation_wait_started = time.time()
    awaiting_reappear = True
    post_press_saw_gap = False
    post_press_started_at = time.time()
```

## 9.3 `max_press` 的真实语义

当前实现已经修正过：

- `max_press` 只限制“还能不能发起下一次新 press”
- 不影响“当前这一 press 之后继续等待其结果”

代码：

```python
def _should_enter_post_press_reappear_wait(awaiting_reappear, press_count, max_press):
    if not awaiting_reappear:
        return False
    # max_press 只限制“不能再按下一次”，不影响“当前这一次按完后等待校验结果”。
    return True
```

## 9.4 按后优先看 loading / validation / reappear

如果按后 challenge 暂时不可按：

- 若 validating：记为进入 loading/validation
- 若只是暂不可按：记为等待 reappear
- 每 1.5 秒轮询一次

代码：

```python
if _should_enter_post_press_reappear_wait(awaiting_reappear, press_count, max_press):
    if not actionable:
        if validating:
            log(f"  {tag} press entered loading/validation, wait for captcha to reappear")
        else:
            log(f"  {tag} captcha not actionable after press, waiting for reappear")
        time.sleep(1.5)
        continue
```

## 9.5 captcha 重新出现后的处理

如果 challenge 真的经历了消失/重新出现，当前逻辑会：

- 结束当前 reappear wait
- 重置等待状态
- 走一次 retry gap
- 然后再进入下一轮 press 机会

代码：

```python
if post_press_saw_gap:
    awaiting_reappear = False
    post_press_saw_gap = False
    post_press_started_at = None
    initial_press_wait_started = time.time() - INITIAL_PRESS_DELAY
    gone_rounds = 0
    _wait_before_next_captcha_press(tag, "challenge failed and captcha reappeared")
    continue
```

## 9.6 按后没有 loading 的失败节奏

如果按后一直没有 loading / validating / reappear，到达 `POST_PRESS_LOADING_CHECK` 后，会按“当前这一 press 没起作用”的节奏重试，而不是一直悬空。

代码：

```python
if gap_waited < POST_PRESS_LOADING_CHECK:
    time.sleep(0.5)
    continue

awaiting_reappear = False
post_press_saw_gap = False
post_press_started_at = None
initial_press_wait_started = time.time() - INITIAL_PRESS_DELAY
_wait_before_next_captcha_press(
    tag,
    f"challenge failed with no loading after {POST_PRESS_LOADING_CHECK}s",
)
```

## 9.7 达到最大 press 后的最后观察窗口

达到 `max_press` 后，不会刚按完就直接 FAIL，而是再观察 `POST_MAX_PRESS_WAIT` 秒看是否跳转。

代码：

```python
if press_count >= max_press:
    if press_wait_started is None:
        press_wait_started = time.time()
        log(f"  {tag} max press {max_press} reached, wait up to {POST_MAX_PRESS_WAIT}s")
    waited = time.time() - press_wait_started
    if waited >= POST_MAX_PRESS_WAIT:
        log(f"  {tag} no redirect after {max_press} presses and {POST_MAX_PRESS_WAIT}s, give up", "WARN")
        _shot(page, "press_fail", idx)
        return _finish(reason="failure")
```

## 9.8 validating/loading 超时保护

所有中间态都挂统一超时，避免某个状态永远不退出。

代码：

```python
def _wait_state_timed_out(started_at, *, now=None, timeout=CAPTCHA_STATE_TIMEOUT):
    if started_at is None:
        return False
    return (now - started_at) >= timeout
```

```python
if _wait_state_timed_out(validation_wait_started, timeout=CAPTCHA_STATE_TIMEOUT):
    log(f"  {tag} captcha post-press state stuck for {int(waited)}s, give up", "WARN")
```

## 9.9 Microsoft Loading 页面单独识别

PX iframe 有时仍残留，但微软页面其实已经进入下一步 loading。当前实现会把这类页面单独识别出来，避免误认为 challenge 失败。

代码：

```python
def _microsoft_loading_page(page):
    if first not in ("loading", "loading...") and not first.startswith("loading"):
        return False
    if "press and hold" in low or "human challenge completed" in low:
        return False
    if "privacy and cookies" in low or "terms of use" in low or len(low) < 400:
        return True
```

---

# 10. 资源拦截（可选）

资源拦截不是 PX 主逻辑，而是当前工作树里的可选优化层。

## 10.1 目的

- 减少无关资源加载
- 降低页面噪音
- 尽量不碰 challenge 自身依赖资源

## 10.2 白名单 host

代码：

```python
_RUOYI_RESOURCE_ALLOW_HOST_HINTS = (
    "fpt.live.com",
    "hsprotect.net",
    "px-cloud.net",
    "px-cdn.net",
    "client.px-cloud.net",
)
```

## 10.3 拦截规则

当前按三类信号判断是否拦截：

1. `Sec-Fetch-Dest`
2. `Accept`
3. URL 后缀

代码：

```python
_RUOYI_RESOURCE_BLOCK_KINDS = ("image", "font", "media")
_RUOYI_RESOURCE_BLOCK_EXTS = (
    ".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".ico", ".bmp", ".avif",
    ".woff", ".woff2", ".ttf", ".otf", ".eot",
    ".mp4", ".webm", ".mp3", ".wav", ".ogg", ".m4a",
)
```

```python
def _ruoyi_should_block_resource_request(req):
    if any(host in low for host in _RUOYI_RESOURCE_ALLOW_HOST_HINTS):
        return False

    dest = str(headers.get("Sec-Fetch-Dest") or headers.get("sec-fetch-dest") or "").strip().lower()
    if dest in _RUOYI_RESOURCE_BLOCK_KINDS:
        return True

    accept = str(headers.get("Accept") or headers.get("accept") or "").strip().lower()
    if "image/" in accept or "font/" in accept or "audio/" in accept or "video/" in accept:
        return True
```

## 10.4 启用方式

当前默认通过 `--block-resources` 控制，不强制开启。

代码：

```python
if bool(getattr(opts, "block_resources", False)):
    _start_ruoyi_resource_blocking(page, tag)
```

## 10.5 退出时关闭拦截器

资源拦截会在 `finally` 中显式关闭，避免结束后拦截器残留。

代码：

```python
finally:
    _stop_ruoyi_resource_blocking(page)
```

---

# 11. 诊断与复盘层

## 11.1 失败截图

当前策略已经收敛为：**失败相关截图为主**，不是整个 PX 全流程狂截。

此外，`_perform_hold_with_px_screenshots()` 只会在：

- `enabled=True`
- 且 `_perform_hold()` 返回失败

时补一张 `press_fail_last_*` 截图。

代码：

```python
def _perform_hold_with_px_screenshots(page, ctx, target, idx, press_count, tag, enabled=False):
    ok = _perform_hold(page, ctx, target, idx, press_count, tag)
    if enabled and not ok:
        _save_screenshot(page, f"press_fail_last_{press_count}", idx, tag)
    return ok
```

## 11.2 失败前缀控制

当前失败截图会基于前缀判断是否值得留存。

代码：

```python
failure_prefixes = (
    "press_fail",
    "captcha_no_target",
    "email_fail",
    "email_input_fail",
    "email_empty_value",
    "pwd_fail",
    "bday_fail",
    "name_fail",
)
```

## 11.3 press_fail 去重

为了避免同一轮失败短时间内重复落图，对 `press_fail` 做了去重窗口。

代码：

```python
_PRESS_FAIL_SCREENSHOT_DEDUP_SEC = 8.0

if not str(name or "").startswith("press_fail"):
    return False
```

## 11.4 HTML / frame 内容保存

失败现场不仅有 png，还有：

- 当前页 outerHTML
- 当前页 innerText
- 所有 frame 的 text / html

这能帮助判断：

- challenge 是否在 iframe
- DOM 是否切换
- 页面到底卡在 Microsoft loading 还是 PX 自身

代码：

```python
html = page.run_js_loaded("return document.documentElement ? document.documentElement.outerHTML : '';") or ""
text = page.run_js_loaded("return document.documentElement ? document.documentElement.innerText : '';") or ""
frames = page.get_all_frames() or []
```

## 11.5 HAR 收集

如果启用 `--har`，会启动 `_RuoyiHarCollector`，保留 challenge 前后请求轨迹。

代码：

```python
if capture_har:
    har_collector = _RuoyiHarCollector(page, tag, idx, email_getter=lambda: email or "")
    har_collector.start()
```

## 11.6 PX metrics

当前批量结果会统计：

- `max_presses`
- `px_elapsed`
- `reg_elapsed`
- `PX_SUMMARY`
- `PX_DETAIL`

代码：

```python
def _normalize_px_metrics(idx, metrics=None):
    return {
        "idx": int(data.get("idx") or idx or 0),
        "max_presses": int(data.get("max_presses") or 0),
        "px_elapsed": float(data.get("px_elapsed") or 0.0),
        "reg_elapsed": float(data.get("reg_elapsed") or 0.0),
    }
```

```python
lines.append(
    f"PX_SUMMARY: max_presses {int(summary['max_px_presses'] or 0)} | avg_px_elapsed {float(summary['avg_px_elapsed'] or 0.0):.2f}s"
)
```

---

# 12. 测试覆盖

当前 PX 方案不是只有运行时逻辑，也有几组配套测试：

- `D:/officeProject/yigehui/reg-factory/test_ruoyi_px_hold_early_release.py`：验证 hold label 隐藏后的提前释放逻辑
- `D:/officeProject/yigehui/reg-factory/test_ruoyi_px_press_screenshots.py`：验证 PX 截图策略
- `D:/officeProject/yigehui/reg-factory/test_ruoyi_captcha_waits.py`：验证 captcha 等待状态机
- `D:/officeProject/yigehui/reg-factory/test_ruoyi_geo_emulation.py`：验证 timezone / geolocation 模拟
- `D:/officeProject/yigehui/reg-factory/test_webui_ruoyi_resource_block_flag.py`：验证 WebUI 的 `--block-resources` 开关拼参

这些测试的价值主要在于固定最容易回归的规则：

- 提前释放
- 按后等待语义
- 失败截图行为
- geo 模拟参数落地
- WebUI 开关是否真的传入 CLI

---

# 13. 已知实现边界

## 13.1 这不是重型指纹框架

当前实现没有上 `smart_fingerprint()` 全套，也没有做全浏览器统一伪装，而是：

- 保持 Firefox 身份
- 补最常见的 headless 暴露点
- 尽量让页面环境和代理出口协调

## 13.2 资源拦截是辅助层

资源拦截更多是：

- 减少噪音
- 控制无关加载
- 提升调试效率

它不是 PX 主通关逻辑本身。

## 13.3 真正效果仍受外部条件影响

即使本地策略合理，最终通过率仍会受这些外部条件影响：

- 代理出口质量
- 微软页当前风控强度
- challenge 当日策略变化
- headless / headed 模式差异
- 并发规模与同时出现的环境一致性

---

# 14. 技术口径总结

如果从技术实现角度概括当前 ruoyi 的 PX 方案，可以总结为：

1. **环境一致性先行**：UA、语言、地区、时区、定位、窗口、screen 尽量统一。
2. **识别逻辑多信号**：DOM、aria、iframe、body text、completed-wait 文案联合判断，不依赖单一选择器。
3. **动作逻辑更像人**：每会话不同 motion profile，非中心点、非瞬移、轨迹和 hold 时长联动。
4. **状态机优先于次数**：一次 press 后优先等结果，不因为达到 `max_press` 就立即断言失败。
5. **复盘能力完整**：失败截图、HTML、frame、HAR、PX metrics 都有出口。

这五点，就是当前工作树里 PX 处理的核心技术结构。
