# px挑战相关总结

> 目标：整理 **当前工作树** 中 `register_outlook_ruoyi.py` 针对 PerimeterX（PX）挑战的完整处理方案，覆盖识别、按压、等待状态机、指纹一致性、资源拦截、诊断与入口开关，方便学习、复盘和分享。

---

## 1. 相关代码与提交

### 1.1 当前核心文件

- `D:/officeProject/yigehui/reg-factory/register_outlook_ruoyi.py`
- `D:/officeProject/yigehui/reg-factory/webui/scripts.py`
- `D:/officeProject/yigehui/reg-factory/webui/server.py`
- `D:/officeProject/yigehui/reg-factory/test_ruoyi_px_hold_early_release.py`
- `D:/officeProject/yigehui/reg-factory/test_ruoyi_px_press_screenshots.py`
- `D:/officeProject/yigehui/reg-factory/test_ruoyi_captcha_waits.py`
- `D:/officeProject/yigehui/reg-factory/test_ruoyi_geo_emulation.py`
- `D:/officeProject/yigehui/reg-factory/test_webui_ruoyi_resource_block_flag.py`

### 1.2 相关提交（历史演进）

- `7d8d0ec` `优化 ruoyi 验证码等待与 profile 清理策略`
- `52d5ead` `调整 ruoyi 并发授权与验证码等待日志`
- `c64ac05` `chore: checkpoint ruoyi headless and px handling`

### 1.3 当前工作树补充

当前工作树在上述提交基础上，还额外包含两块 PX 相关补充：

- `image/font/media` 资源拦截
- WebUI 对 `--block-resources` 的勾选开关

本文以 **当前代码实际状态** 为准，而不是只按某一次提交回看。

---

## 2. 总体思路

当前 ruoyi 的 PX 处理，不是单点“暴力长按”，而是一整套链路：

1. **浏览器与页面一致性修补**：尽量让 headless Firefox 的 `UA / language / locale / screen / visibility / webdriver / timezone / geolocation` 协调一致。
2. **挑战识别**：在主页面和 iframe 中查找 PX challenge，优先找真正可按的 hold target。
3. **动作执行**：每会话生成一套短轨迹 motion profile，非中心点、非瞬移、按压时长和移动时长联动。
4. **提前释放**：hold 中轮询 label，如果 label 已隐藏，则提前松手，不死等到固定秒数。
5. **按后状态机**：一次按压后，不立刻判输，而是先等待 loading / validating / captcha reappear / redirect 的真实结果。
6. **诊断与复盘**：失败截图、HTML 状态保存、HAR、PX 统计汇总、WebUI 开关。

主线入口代码：

```python
if proxy_pool:
    _apply_ruoyi_proxy_geo_emulation(page, proxy_pool, tag)

if bool(getattr(opts, "block_resources", False)):
    _start_ruoyi_resource_blocking(page, tag)

if is_headless and not signup_opened:
    _apply_ruoyi_headless_page_patches(page, tag, log_once=True, user_agent=user_agent)
```

---

## 3. 配置项与入口开关

### 3.1 CLI 参数

当前与 PX 处理直接相关的参数：

- `--max-press`：最大按压次数
- `--headless`：无头模式
- `--px-press-screenshots`：PX 失败相关截图
- `--block-resources`：启用 image/font/media 资源拦截
- `--ua-pool`：UA 池
- `--timeout`：单号超时

代码：

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

### 3.2 WebUI 表单入口

WebUI 已暴露对应开关，便于调试：

- `max-press`
- `headless`
- `block-resources`
- `px-press-screenshots`

代码：

```python
{"flag": "--max-press", "type": "str", "default": "5", "help": "验证码按住次数上限"},
{"flag": "--headless", "type": "bool", "default": False, "help": "无头模式"},
{"flag": "--block-resources", "type": "bool", "default": False, "help": "屏蔽 image/font/media 请求"},
{"flag": "--px-press-screenshots", "type": "bool", "default": False,
 "help": "保存 ruoyi PX/失败相关截图"},
```

---

## 4. 时间参数与基本节奏

PX 当前并不是“看到就按”，而是有明确节奏参数：

- `INITIAL_PRESS_DELAY = 5`：challenge 变成可按后，先等 5 秒再按
- `PX_HOLD_EARLY_RELEASE_AFTER = 5.0`：hold 开始 5 秒后，开始轮询 label 是否隐藏
- `PX_HOLD_EARLY_RELEASE_INTERVAL = 0.5`：轮询间隔 0.5 秒
- `PX_HOLD_SECONDS_MIN / MAX = 10.5 ~ 11.0`：基准 hold 时长
- `POST_MAX_PRESS_WAIT = 5`：达到最大 press 后，再等 5 秒看是否跳转
- `POST_PRESS_LOADING_CHECK = 8`：按后 8 秒内仍看不到 loading，再当成失败节奏
- `CAPTCHA_STATE_TIMEOUT = 20`：单个 captcha 状态等待上限
- `POST_PRESS_RETRY_GAP_MIN / MAX = 2 ~ 4`：按后失败时下次尝试前退避

代码：

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

---

## 5. 指纹一致性层

### 5.1 UA 池与每实例 UA

ruoyi 当前会先从 UA 池里为每个实例选一条 UA，避免多个并发完全同模版。

代码：

```python
_DEFAULT_UA_POOL = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:151.0) Gecko/20100101 Firefox/151.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:150.0) Gecko/20100101 Firefox/150.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:149.0) Gecko/20100101 Firefox/149.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:148.0) Gecko/20100101 Firefox/148.0",
)

user_agent = _pick_user_agent(idx)
log(f"  {tag} ua pool pick -> {_mask_ua(user_agent)} ({user_agent[:72]}...)")
```

### 5.2 浏览器级 UA override

浏览器启动前，先尽量把 UA 写到 Firefox preference / option API。

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

### 5.3 Headless 启动参数与 preference

headless 下不是只开 `headless(True)`，还会补齐窗口和语言，避免过于默认。

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

### 5.4 Emulation 层统一

如果 `page.emulation` 可用，还会补：

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

### 5.5 页面补丁：最小反自动化痕迹处理

这层修的是“常见 headless 暴露点”，不是全量伪造显卡/插件/跨浏览器特征。

处理目标：

- `navigator.webdriver`
- `navigator.userAgent`
- `navigator.appVersion`
- `navigator.languages`
- `document.hidden`
- `document.visibilityState`
- `document.hasFocus()`
- `screen.*`
- 删除 `__playwright` / `__pwInitScripts` / `cdc_*`

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

### 5.6 Preload + 运行时双重注入

除了当前页 `run_js`，还会尽量通过 preload 脚本提前注入，减少首屏暴露窗口。

代码：

```python
def _ensure_ruoyi_headless_preload(page, tag=None, log_once=False, user_agent=None):
    add_script = getattr(page, "add_preload_script", None)
    preload_js = f"() => {{{_build_headless_patch_js(ua)}}}"
    add_script(preload_js)
```

### 5.7 代理出口对应 timezone + geolocation

这是当前 PX 一致性里非常重要的一层：页面地理信息尽量与代理出口一致。

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

---

## 6. PX 挑战识别层

### 6.1 多语言 hold/challenge 文案匹配

PX 目标识别不是只写英文，当前包含多语言 hold/challenge 关键字。

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

### 6.2 主目标查找 `_find_hold_target()`

主查找逻辑会在当前 context 里收集候选按钮：

- `#px-captcha`
- 带 `aria-label` 的 button / input / a[role=button]
- 文本 label 向上找祖先按钮
- 纯按钮文本包含 hold 的按钮

同时会记录：

- 中心点
- left/top/width/height
- 文案
- id / tag / role / ariaLabel
- quality
- source
- holdLabelId

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

### 6.3 iframe 优先与 fallback

`_find_hold_context()` 会遍历所有 context，优先：

- `hsprotect.net`
- `arkose`
- `funcaptcha`

如果一个都没命中，再退回到页面上可见 iframe 的 box 坐标作为 fallback。

代码：

```python
def _find_hold_context(page, min_quality=5):
    for ctx in _all_contexts(page):
        target = _find_hold_target(ctx)
        ...
        if any(k in url for k in ["hsprotect.net", "arkose", "funcaptcha"]):
            frame_bonus = -10
    ...
    box = _find_hsprotect_iframe_box(page)
    if box is not None:
        return page, box
```

### 6.4 captcha 可见性判断

challenge 是否存在，不只看一个选择器，还会看：

- 当前 frame 是否已找到高质量 hold target
- 页面里是否存在 hsprotect/arkose/funcaptcha iframe
- 页面 body 文案是否出现 `press and hold / captcha / perimeterx / 按住 / 长按` 等

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

### 6.5 validating / completed-wait 判断

当前 validating 检测分两层：

1. `_px_captcha_completed_wait(page)`：扫描 `#px-captcha` 子树里是否有 `human challenge completed, please wait`
2. `_captcha_is_validating(page)`：
   - 有 iframe hint 但暂时找不到 hold target
   - 或页面文案里出现 `verifying / checking / loading / please wait / just a moment` 等

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

### 6.6 hold label 定位与状态读取

当前按压不是盲按到底，而是会把目标附近的 hold label 找出来，并在 hold 中读取它的显示状态。

用途：

- 判断 label 是否已经 `display:none`
- 已隐藏则认为挑战进入下一阶段，可以提前释放

代码：

```python
def _resolve_hold_label_for_target(ctx, target):
    ...
    hits.sort((a, b) => (a.distance - b.distance) || ((b.id ? 1 : 0) - (a.id ? 1 : 0)));
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

## 7. PX 动作执行层

### 7.1 每会话 motion profile

每个 page 启动后都会生成一套独立运动参数，而不是所有实例同一路径。

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
        "lead_dx": random.choice((-1, 1)) * random.uniform(8.0, 16.0),
        "lead_dy": random.uniform(-6.0, 6.0),
        "move_ratio": random.uniform(0.024, 0.036),
        "lead_ratio": random.uniform(0.42, 0.56),
        "settle_ratio": random.uniform(0.20, 0.28),
        "micro_ratio": random.uniform(0.10, 0.18),
        "hover_ratio": random.uniform(0.008, 0.016),
    }
```

### 7.2 非中心点坐标选择

当前不会永远点元素中心。`_ruoyi_point_inside_target()` 会在目标矩形内按比例和 jitter 取点。

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

### 7.3 短轨迹：lead → settle → micro → final

当前按压前的移动不是单段，而是四段短轨迹：

1. `lead`
2. `settle`
3. `micro`
4. `final`

并且总移动时长与 `hold_sec` 联动。

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

### 7.4 提前释放

hold 开始后，不是死等 `10.5~11s` 完整结束；5 秒后开始轮询 hold label，如果已经隐藏则提前释放。

代码：

```python
next_check = hold_started + PX_HOLD_EARLY_RELEASE_AFTER
...
release_state = _px_hold_instruction_state(state_ctx, hold_p_id)
if release_state and str(release_state.get("display") or "").strip().lower() == "none":
    break
```

### 7.5 label id 刷新

如果原 hold label 读不到，会重新找 context / target / label，并刷新 `hold_p_id`，减少 iframe / DOM 变化导致的误判。

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

### 7.6 hold 失败兜底

如果 hold 行为抛异常，会尝试 `release_all()`，避免鼠标状态残留。

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

## 8. PX 按后状态机

### 8.1 首次可按前的等待

当前设计是：challenge 变成 visible + actionable 后，先等 `INITIAL_PRESS_DELAY` 再按，不抢第一时间。

代码：

```python
if visible and actionable:
    if initial_press_wait_started is None:
        initial_press_wait_started = time.time()
        log(f"  {tag} captcha visible, wait {INITIAL_PRESS_DELAY}s before press")
        time.sleep(INITIAL_PRESS_DELAY)
        continue
```

### 8.2 一次按压成功后，先进入“等待结果”态

如果 hold 成功：

- 记录 `px_hold_elapsed`
- 标记 `awaiting_reappear = True`
- 进入按后状态机
- 这时不会立刻允许下一次 press

代码：

```python
if hold_elapsed:
    px_hold_elapsed += float(hold_elapsed)
    validation_wait_started = time.time()
    awaiting_reappear = True
    post_press_saw_gap = False
    post_press_started_at = time.time()
```

### 8.3 `max_press` 语义修正

当前 `max_press` 的语义是：

- **限制下一次是否还能再按**
- **不打断当前这一按之后的判定等待**

代码：

```python
def _should_enter_post_press_reappear_wait(awaiting_reappear, press_count, max_press):
    if not awaiting_reappear:
        return False
    return True
```

### 8.4 按后先看 loading / validation / captcha reappear

按后如果 challenge 暂时不可按：

- 如果正在 validating：打印 `press entered loading/validation, wait for captcha to reappear`
- 否则：打印 `captcha not actionable after press, waiting for reappear`
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

### 8.5 captcha 消失后重新出现的处理

如果 challenge 真的经历了“消失/不可按 → 重新出现”，当前代码会：

- 结束当前 reappear wait
- 重置一些等待变量
- 再按退避间隔继续下一轮

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

### 8.6 按后没有 loading 的失败节奏

如果按后一直没看到 loading / validating / reappear，到达 `POST_PRESS_LOADING_CHECK`，就按“当前轮 challenge failed”处理，而不是一直卡住。

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

### 8.7 达到最大 press 后的最后等待

达到 `max_press` 之后，不是立刻 FAIL，而是再给 `POST_MAX_PRESS_WAIT` 秒看是否跳转。

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

### 8.8 validating/loading 超时保护

当前对 validating / loading / reappear 都有统一超时保护，避免一直僵死。

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

### 8.9 Microsoft Loading 页面单独识别

PX iframe 有时还在 DOM 里，但微软其实已经进入下一步 loading 页面。当前代码会单独识别这种页面，避免误当成 challenge 失败。

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

## 9. 资源拦截（可选）

### 9.1 目的

资源拦截不是挑战主逻辑本身，而是一个可选优化层：

- 少加载无关图片 / 字体 / 媒体
- 降低页面噪音
- 保留 PX 相关 host，不误伤挑战本身

### 9.2 白名单 host

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

### 9.3 拦截规则

当前会根据三类信息判断是否拦截：

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

### 9.4 启用方式

当前默认通过开关控制，不强制开启。

代码：

```python
if bool(getattr(opts, "block_resources", False)):
    _start_ruoyi_resource_blocking(page, tag)
```

### 9.5 退出时关闭拦截器

资源拦截不是一次性 fire-and-forget，当前在 `finally` 里会显式停掉，避免页面结束后拦截器残留。

代码：

```python
finally:
    _stop_ruoyi_resource_blocking(page)
```

---

## 10. 诊断与复盘层

### 10.1 失败截图

PX 当前失败相关截图已经收敛为“主要在失败时保留”，而不是整个过程全程狂截。

另外，`_perform_hold_with_px_screenshots()` 只会在 `enabled=True` 且本次 hold 失败时补一张 `press_fail_last_*` 截图。

代码：

```python
def _perform_hold_with_px_screenshots(page, ctx, target, idx, press_count, tag, enabled=False):
    ok = _perform_hold(page, ctx, target, idx, press_count, tag)
    if enabled and not ok:
        _save_screenshot(page, f"press_fail_last_{press_count}", idx, tag)
    return ok
```

失败前缀包括：

- `press_fail`
- `captcha_no_target`
- `timeout_captcha_reappear`
- `timeout_captcha_state`
- 其他注册阶段失败

代码：

```python
failure_prefixes = (
    "press_fail",
    "captcha_no_target",
    "email_fail",
    "email_input_fail",
    "pwd_fail",
    "bday_fail",
    "name_fail",
)
```

### 10.2 press_fail 去重

为避免同一失败窗口短时间连续截图，当前对 `press_fail` 做了时间去重。

代码：

```python
_PRESS_FAIL_SCREENSHOT_DEDUP_SEC = 8.0

if not str(name or "").startswith("press_fail"):
    return False
```

### 10.3 页面 HTML / frame 状态保存

失败截图不是只有 png，还会附带：

- 页面 outerHTML
- 页面 innerText
- 各 frame 的 text + html

这对事后分析 challenge 卡在哪一层很有帮助。

代码：

```python
html = page.run_js_loaded("return document.documentElement ? document.documentElement.outerHTML : '';") or ""
text = page.run_js_loaded("return document.documentElement ? document.documentElement.innerText : '';") or ""
frames = page.get_all_frames() or []
```

### 10.4 HAR

如果开了 `--har`，会把完整链路 HAR-like 数据收集下来，方便复盘挑战前后请求。

代码：

```python
if capture_har:
    har_collector = _RuoyiHarCollector(page, tag, idx, email_getter=lambda: email or "")
    har_collector.start()
```

### 10.5 PX 指标统计

当前批量结果会单独统计 PX：

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

## 11. 当前实现的要点归纳

### 11.1 最核心的四条

1. **最小 Firefox 风格指纹一致性修补**，不是重型 `smart_fingerprint()`。
2. **每会话不同的短轨迹按压**，非中心点、非瞬移、和 hold 时长联动。
3. **一次按压后优先等待真实结果**，不是刚按完就因为 `max_press` 直接判失败。
4. **失败可复盘**，截图、HTML、HAR、PX metrics 都能回看。

### 11.2 当前实现的边界

- 这套方案主要是在 **Firefox + ruyi/ruyipage** 约束下，尽量做“更像真人”的最小修补。
- 它不是通用反检测框架，也不是全浏览器统一伪装层。
- 资源拦截是可选优化，不是 PX 主逻辑本体。

---

## 12. 配套测试覆盖

当前 PX 方案不是只有运行时逻辑，也有几组配套测试用于兜底：

- `D:/officeProject/yigehui/reg-factory/test_ruoyi_px_hold_early_release.py`：验证 hold label 隐藏后的提前释放逻辑
- `D:/officeProject/yigehui/reg-factory/test_ruoyi_px_press_screenshots.py`：验证 PX 截图策略
- `D:/officeProject/yigehui/reg-factory/test_ruoyi_captcha_waits.py`：验证 captcha 等待状态机
- `D:/officeProject/yigehui/reg-factory/test_ruoyi_geo_emulation.py`：验证 timezone / geolocation 模拟
- `D:/officeProject/yigehui/reg-factory/test_webui_ruoyi_resource_block_flag.py`：验证 WebUI 的 `--block-resources` 开关拼参

这些测试的意义，不是覆盖全部真实 PX 场景，而是把最容易回归的几条规则固定住：

- 提前释放
- 最后一次按压后的等待语义
- 失败截图行为
- geo 模拟参数落地
- WebUI 开关是否真的传到 CLI

---

## 13. 对外分享时可以直接引用的一段总结

当前 ruoyi 的 PX 方案，本质上不是“简单长按验证码”，而是：

- 先把 Firefox headless 的常见暴露点做最小一致性修补；
- 再让页面 `UA / language / locale / timezone / geolocation` 尽量和代理出口协调；
- 然后在 challenge 出现时，通过多语言文案、iframe、按钮语义去找真正可按的 hold target；
- 实际执行时使用每会话独立的短轨迹 motion profile，非中心点、非瞬移、按住过程中还能根据 hold label 是否隐藏决定是否提前释放；
- 按完后进入等待状态机，先看 loading / validating / captcha 是否重新出现，再决定是否重试或失败；
- 最后把截图、HTML、HAR、PX 指标都沉淀下来，便于复盘。
