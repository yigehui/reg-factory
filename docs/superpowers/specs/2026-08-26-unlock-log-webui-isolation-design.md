# 解锁日志与 WebUI 日志隔离 设计

## 背景(三个问题)

1. **WebUI 运行日志互相干预**:左侧菜单多个脚本共用一个日志区。跑着解锁账号时点"唤醒 ruyi 浏览器",
   `runScript()` 清空日志区 + 切到新 SSE,解锁日志丢失。后端 `RUNS` 里旧 run 的 lines 仍在,
   前端却看不到了。用户诉求:**每个菜单(脚本)日志独立**,可同时跑多个脚本,切回某菜单看它的日志。

2. **解锁日志不如注册清晰**:解锁全是裸 `print`,无时间戳、无级别、无阶段分隔、无 step 计时、缩进层级混乱。
   用户诉求:对齐 ruyi 注册 —— **加时间戳、日志级别、worker/线程标记、step 简单计时**。

3. **错误凭证等到固定超时**:账号/密码错误时微软登录页显示错误文案,但 `classify` 没识别,
   一直匹配 `email_form`/`login_form` 反复重填,等 ~60s 超时才失败。应在 ~3s 内快速失败。

## 改动范围

- `unlock_outlook.py` —— 问题 2(日志格式) + 问题 3(错误凭证快失败)
- `webui/static/app.js` —— 问题 1(每脚本日志独立)
- `tests/` —— 对应测试

三个问题互不耦合,可独立实现与验证。

---

## 问题 1:WebUI 每脚本日志独立

### 现状根因
- `app.js`: 全局单 `curRun`/`evtSrc`/`#log`;`runScript()` 启动新任务时
  `log.textContent=''` 清空 + `evtSrc.close()` + `curRun` 覆盖。
- 切脚本(`selectScript`)只换表单,不动日志区 → 旧脚本日志被新脚本覆盖。
- 后端 `RUNS = {run_id: {lines, done, ...}}` 每 run 独立、`/api/logs/{run_id}` 按 run_id 推,
  **后端已支持多 run 并存**,问题纯在前端。

### 方案:每脚本独立日志缓冲 + 按脚本聚合
前端按 `script_id` 维护日志状态(不按 run_id —— 同一脚本多次运行归到同一菜单):

```
let scriptLogs = {};   // { script_id: { lines: [], run_id: null, done: true, evt: null } }
let curLogScript = null;  // 当前日志区展示的脚本 id
```

行为:
- **启动任务**(`runScript`):给当前 `curSrc.id` 建/取 `scriptLogs[id]`,
  - 若该脚本已有未 done 的旧 run,先 `evt.close()`(同脚本不并发跑两个,保持现状语义)。
  - 不清空 `lines`,而是**追加一条分隔线** `[webui] ===== 重新运行 =====`(保留历史,便于回看)。
  - 存 `run_id`,开 SSE 推进 `lines`,`done=false`。
  - 自动把日志区切到该脚本(`curLogScript = curSrc.id`),渲染 `lines`。
- **切菜单**(`selectScript`):渲染表单(现状不变) + **同时把日志区切到该脚本**:
  `curLogScript = id`,渲染 `scriptLogs[id]?.lines || []`。若该脚本有未 done 的 run,SSE 仍在后台推,
  切回时自动接续(下一帧 poll 渲染)。
- **SSE 消息**:onmessage 时把 line 追加进 `scriptLogs[id].lines`;`done` 事件置 `done=true` + `evt.close()`。

这样:解锁跑着(SSE 后台推 `scriptLogs['unlock']`),点"唤醒浏览器"启动 → 日志区切到 `scriptLogs['launch']`;
切回解锁菜单 → 日志区渲染 `scriptLogs['unlock']` 已收到的全部 lines,SSE 仍在推,新 line 继续追加。
两脚本日志互不覆盖。

### 渲染细节
- `#log` 渲染:切脚本时 `log.textContent = lines.join('\n')` + 滚到底(`LogAutoScroll`)。
- `#log-title`:`运行日志 — {script.title}`(现状已类似,保持)。
- `#btn-stop`:对 `curLogScript` 对应的 `run_id` 发 stop;无活动 run 则 disabled。
- `#cmd-preview`:展示 `curLogScript` 对应的 cmd。
- 不引入 tab DOM,纯 JS 状态切换(改动最小,UI 不变,行为变正确)。

### 边界
- 同一脚本连点两次运行:关旧 SSE、追加分隔线、开新 SSE(同现状"覆盖"语义,但不清空历史)。
- 后端 `RUNS` 在任务 done 后仍保留 lines(现状),前端切回能看;长时间运行多任务内存增长由后端
  现有清理逻辑兜底(本次不动后端)。

---

## 问题 2:解锁日志对齐注册

### 现状根因
解锁裸 `print(f"    [{name}] {state}")` 等:无时间戳、无级别、无 step 计时、缩进混乱。
注册用 `log(msg, level)` 输出 `[HH:MM:SS] [INFO] msg`,且有 `========== 注册 #1/1 ==========`、
`step xxx: N.NNs`、`SUMMARY:` 等结构。

### 方案:复用注册同款 log + step 计时
`unlock_outlook.py` 内:

1. **引入 log 函数**(对齐注册,简化版 —— 解锁不接业务关键词过滤):
   ```python
   def log(msg, level="INFO"):
       ts = time.strftime("%H:%M:%S")
       print(f"[{ts}] [{level}] {msg}", flush=True)
   ```
   - 不复用 `common.ruyi.log`(那个是包骨架,无业务过滤但解锁要自己的 worker 标记);
     也避免改 register 的业务版。解锁自建轻量版,单一来源在本文件。

2. **worker/线程标记**:保留现有 `tag = f"w{worker_id}` 概念,日志里带 `[#1][w0]` 形式
   (账号序号 + worker),对齐注册 `[#1][ruoyi]`。

3. **step 计时**:每个关键步骤记 `t0=time.perf_counter()`,完成 `log(f"  step {name}: {elapsed:.2f}s")`:
   - `step precheck`、`step login`、`step px`、`step finalize`。
   - 用 `perf_counter` 差值,非墙上时钟。

4. **阶段分隔**:
   - 每账号开头:`log(f"========== 解锁 {idx}/{total} ==========")`
   - 每账号结果:`log(f"[#{idx}] 结果: {outcome} {email} total={elapsed:.2f}s")`
   - 批次结束:`log(f"SUMMARY: unlocked {u} | needs_phone {n} | failed {f} | total {t}")` +
     `log(f"SUMMARY_TIME: total_elapsed {e:.2f}s")`(与注册口径一致,便于 webui 解析)。

5. **替换裸 print**:把现有 `print(f"    [{name}] {state}")`、`print(f"  {tag} ...")` 等逐个改 `log(...)`,
   统一缩进(账号级 0 缩进,步骤级 `  ` 2 空格)。保留 `file=sys.stderr` 的 WARN/ERR 行改 `log(..., "WARN")`。

### 不动的东西
- 浏览器内 JS 注入日志(`inject px error` 等)保留 print,那是浏览器侧异常,不影响主日志清晰度。
- `[px]` 子模块内部日志(API 轮询)保留原样,属子模块细节。

---

## 问题 3:错误凭证快速失败

### 方案
`unlock_outlook.py`:

1. **`classify` 加 `login_error` 态**(在 `email_form`/`login_form` 判定之前插入):
   ```python
   if any(x in t for x in [
       "we couldn't find an account",          # 账号不存在
       "your account or password is incorrect", # 密码错误
   ]): return "login_error"
   ```
   - 文案经联网确认是微软登录页标准错误文案(账号错误在邮箱提交后,密码错误在密码提交后)。
   - 放在 L235 `enter your password` / L236 `email or phone` 之前,避免被抢先匹配。

2. **主循环 Step 1 快速失败**(L539 循环内,与 `sms_verify`/`fido_setup` 同级):
   ```python
   if state == "login_error":
       log(f"[{tag}] 凭证错误,跳过(不等超时)", "WARN")
       return "bad_credentials"
   ```

3. **输出映射**:`outcome="bad_credentials"` 自动落 failed 文件
   (L864 `failed = [r for r in results if r[3] not in (...)]` 兜底),无需改输出逻辑。
   - failed 文件里能区分是凭证问题而非超时(日志有 WARN 记录)。

### 不动的东西
- `error_page`(Something went wrong)仍走 Try-again 逻辑,不与 `login_error` 混淆
  (文案不同,且 `error_page` 判定在前)。

---

## 测试

### 问题 1(webui 前端)
- 现有 webui 测试是后端逻辑(`test_webui_server_*`)。前端纯 JS 状态改造,加一个
  `tests/test_webui_log_isolation.py` 静态检查:`app.js` 存在 `scriptLogs` 且 `runScript` 不含
  `log.textContent=''` 全清空(防回退)。轻量,不做 DOM 渲染测试(无 jsdom)。

### 问题 2(解锁日志格式)
- `tests/test_unlock_log_format.py`:
  - `test_log_has_timestamp_and_level`:断言 `log("x","INFO")` 输出匹配 `^\[\d{2}:\d{2}:\d{2}\] \[INFO\] x`。
  - `test_summary_lines_format`:断言 SUMMARY/SUMMARY_TIME 行格式与注册一致。

### 问题 3(错误凭证)
- `tests/test_unlock_bad_credentials.py`:
  - `test_classify_login_error_account_not_found`:`classify` 对 "we couldn't find an account" 返回 `login_error`。
  - `test_classify_login_error_wrong_password`:"your account or password is incorrect" → `login_error`。
  - `test_login_error_fast_fail`:mock `classify` 返回 `login_error`,主循环返回 `"bad_credentials"`(不等 deadline)。

---

## 验证

- `pytest tests/` 全绿(含新测试)。
- 问题 1:webui 手测 —— 同时跑 unlock + launch 前端,切菜单看各自日志不丢。
- 问题 2:跑 `unlock_outlook.py --input <现成账号>` 看日志有 `[HH:MM:SS] [INFO]` + step 计时 + SUMMARY。
- 问题 3:用错误账号跑,~3s 内 `=> bad_credentials` 落 failed 文件,不等超时。

## 风险
- R1 问题 1:同脚本重复运行不清空历史,日志可能很长 → 分隔线区分 + 依赖后端清理,可接受。
- R2 问题 2:改 print 为 log 可能影响 webui 的 summary 解析 → 保持 SUMMARY/SUMMARY_TIME 前缀不变
  (webui scripts.py 按这些前缀解析)。
- R3 问题 3:错误文案漏判 → 文案列表保守,只放两条确定的标准文案;漏判时退回超时路径,不误判成功。
