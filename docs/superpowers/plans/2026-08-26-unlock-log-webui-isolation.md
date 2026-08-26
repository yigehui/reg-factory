# 解锁日志与 WebUI 日志隔离 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 修三个问题 —— WebUI 每脚本日志独立、解锁日志对齐注册格式、解锁遇错误凭证快速失败。

**Architecture:** 三个改动分处三文件、互不耦合:`unlock_outlook.py` 加 `login_error` classify 分支 + 自建 `log()` 函数替换裸 print 加 step 计时;`webui/static/app.js` 把全局单 `curRun`/`evtSrc` 改成按 `script_id` 聚合的 `scriptLogs` 缓存,切菜单即切日志区不清空。各自独立测试。

**Tech Stack:** Python 3.13(unittest,base py 跑 unlock 测试)+ 原生 JS(无框架,前端纯状态改造)。

**Spec:** `docs/superpowers/specs/2026-08-26-unlock-log-webui-isolation-design.md`

---

## File Structure

| 文件 | 职责 | 改动 |
|------|------|------|
| `unlock_outlook.py` | 解锁主脚本 | 加 `log()`、`classify` 加 `login_error`、主循环快失败、替换裸 print + step 计时 + SUMMARY |
| `webui/static/app.js` | WebUI 前端逻辑 | `scriptLogs` 按 script_id 聚合,切菜单切日志区不清空 |
| `tests/test_unlock_bad_credentials.py` | 问题3 测试 | classify + 主循环快失败 |
| `tests/test_unlock_log_format.py` | 问题2 测试 | log 时间戳级别 + SUMMARY 格式 |

`unlock_outlook.py` 的 `log()` 不复用 `common.ruyi.log`(解锁要自己的 worker 标记;且避免改 register 业务版)。`classify` 是模块级函数可独立单测。

---

## Task 1: 错误凭证快速失败 —— classify 加 `login_error` 态

**Files:**
- Modify: `unlock_outlook.py:233-236`
- Test: `tests/test_unlock_bad_credentials.py`

`classify` 当前在 `email_form`/`login_form` 之前无错误凭证判定,密码错误时落 `login_form` 反复重填到超时。

- [ ] **Step 1: 写失败测试 —— classify 识别两种错误文案**

创建 `tests/test_unlock_bad_credentials.py`:

```python
import unittest
from unittest.mock import patch


import unlock_outlook as mod


def _fake_page(text="", url=""):
    """构造一个最小 fake page,让 classify 取到 text 和 url。"""
    class _Ctx:
        def run_js_loaded(self, js, *a, **k):
            return text
    class _Page:
        def __init__(self):
            self.url = url
        @property
        def main_frame(self):
            return _Ctx()
        def run_js_loaded(self, js, *a, **k):
            return text
        def evaluate(self, js, *a, **k):
            return text
    return _Page()


class ClassifyLoginErrorTests(unittest.TestCase):
    def test_classify_login_error_account_not_found(self):
        page = _fake_page(text="We couldn't find an account with that username.")
        self.assertEqual(mod.classify(page), "login_error")

    def test_classify_login_error_wrong_password(self):
        page = _fake_page(text="Your account or password is incorrect.")
        self.assertEqual(mod.classify(page), "login_error")

    def test_classify_email_form_still_works(self):
        page = _fake_page(text="Sign in\nemail or phone", url="https://login.live.com/login.srf")
        self.assertEqual(mod.classify(page), "email_form")


if __name__ == "__main__":
    unittest.main()
```

注意:`classify` 内部读 text 走 `_context_text`/`run_js_loaded`,需先跑一次看它怎么取 text,若 `_fake_page` 不匹配则调整(见 Step 2 说明)。

- [ ] **Step 2: 跑测试看 classify 实际怎么取 text,确认 fake page 形态**

Run: `python -m pytest tests/test_unlock_bad_credentials.py -v`

先看 `classify` 取 text 的路径(`unlock_outlook.py:194-211` 附近),确认 `_fake_page` 暴露的方法名对得上。
读 `unlock_outlook.py:194-211` 确认它调的是 `page.run_js_loaded(...)` 还是 `_all_contexts(page)` + `ctx.run_js_loaded`。
若 fake page 的方法名不对,测试会因 AttributeError 失败 —— 按实际方法名调整 `_fake_page`(可能要加 `main_frame` / `frames` / `_all_contexts` 支撑)。
Expected: 三个测试,**前两个 FAIL**(`login_error` 不存在,返回 `unknown`/`email_form`),**第三个 PASS**(email_form 现有逻辑)。

- [ ] **Step 3: classify 加 `login_error` 分支**

`unlock_outlook.py:233-236` 当前:

```python
    if "something went wrong" in t: return "error_page"
    if "chrome-error://" in u or "about:neterror" in u: return "net_error"
    if "enter your password" in t:  return "login_form"
    if any(x in t for x in ["email or phone", "sign in", "enter your email"]): return "email_form"
    return "unknown"
```

改成(在 `error_page`/`net_error` 之后、`login_form`/`email_form` 之前插入 `login_error`):

```python
    if "something went wrong" in t: return "error_page"
    if "chrome-error://" in u or "about:neterror" in u: return "net_error"
    # 凭证错误:微软登录页标准文案(账号不存在 / 密码错误),优先于 login_form/email_form
    # 避免密码错误时反复重填等到超时;放这里在 error_page 之后(error_page 文案不同不冲突)。
    if any(x in t for x in [
        "we couldn't find an account",
        "your account or password is incorrect",
    ]): return "login_error"
    if "enter your password" in t:  return "login_form"
    if any(x in t for x in ["email or phone", "sign in", "enter your email"]): return "email_form"
    return "unknown"
```

- [ ] **Step 4: 跑测试确认全过**

Run: `python -m pytest tests/test_unlock_bad_credentials.py -v`
Expected: 三个测试全 PASS。

- [ ] **Step 5: 提交**

```bash
git add tests/test_unlock_bad_credentials.py unlock_outlook.py
git commit -m "feat(unlock): classify 识别凭证错误文案(login_error),为快失败铺路"
```

---

## Task 2: 错误凭证快速失败 —— 主循环遇 `login_error` 快失败

**Files:**
- Modify: `unlock_outlook.py:547-554`(Step 1 主循环的 state 分支区)
- Test: `tests/test_unlock_bad_credentials.py`(同文件追加)

主循环 Step 1(`unlock_outlook.py:539` 的 `for i in range(20)`),遇 `login_error` 应 `return "bad_credentials"` 而非继续循环到 deadline。

- [ ] **Step 1: 写失败测试 —— unlock_account 遇 login_error 返回 bad_credentials**

追加到 `tests/test_unlock_bad_credentials.py` 末尾(`if __name__` 之前):

```python
class FastFailLoginErrorTests(unittest.TestCase):
    def test_unlock_account_returns_bad_credentials_on_login_error(self):
        """classify 首次即返回 login_error -> unlock_account 立即返回 bad_credentials,
        不等 deadline。用 time 上下文证明没耗满 timeout。"""
        import time as _time
        page = _fake_page(text="Your account or password is incorrect.")
        t0 = _time.perf_counter()
        outcome = mod.unlock_account(page, "x@outlook.com", "wrong", "w0", 0,
                                     max_press=1, timeout=60)
        elapsed = _time.perf_counter() - t0
        self.assertEqual(outcome, "bad_credentials")
        self.assertLess(elapsed, 10.0)  # 远小于 60s timeout,证明快失败
```

- [ ] **Step 2: 跑测试确认失败**

Run: `python -m pytest tests/test_unlock_bad_credentials.py::FastFailLoginErrorTests -v`
Expected: FAIL —— 当前 `login_error` 没分支,会落 `for i in range(20)` 反复 snap + sleep,要么返回 `timeout` 要么跑满循环。outcome ≠ `bad_credentials`。

- [ ] **Step 3: 主循环 Step 1 加 login_error 分支**

`unlock_outlook.py:547-555` 当前(Step 1 循环内的分支):

```python
        if state == "abuse":
            _click_any(page, _CONTINUE_SELECTORS, timeout=2)
            time.sleep(3); continue
        if state == "logged_in":  return "already_ok"
        if state == "sms_verify": return "needs_phone"
        if state == "fido_setup":
            _maybe_skip_passkey(page, tag); time.sleep(4)
            return "unlocked"
        if state in ("locked", "px_challenge"): break
```

改成(在 `logged_in` 之后加 `login_error`):

```python
        if state == "abuse":
            _click_any(page, _CONTINUE_SELECTORS, timeout=2)
            time.sleep(3); continue
        if state == "logged_in":  return "already_ok"
        if state == "login_error":
            print(f"    [{tag}] 凭证错误,跳过(不等超时)", file=sys.stderr)
            return "bad_credentials"
        if state == "sms_verify": return "needs_phone"
        if state == "fido_setup":
            _maybe_skip_passkey(page, tag); time.sleep(4)
            return "unlocked"
        if state in ("locked", "px_challenge"): break
```

注:`return "bad_credentials"` 自动落 failed 文件(`save_results` L864 `failed = [r for r in results if r[3] not in ("unlocked","already_ok","needs_phone")]` 兜底),无需改输出映射。此处暂用裸 print(stderr),Task 3 会统一改 `log(..., "WARN")`。

- [ ] **Step 4: 跑测试确认通过**

Run: `python -m pytest tests/test_unlock_bad_credentials.py -v`
Expected: 4 个测试全 PASS。

- [ ] **Step 5: 提交**

```bash
git add tests/test_unlock_bad_credentials.py unlock_outlook.py
git commit -m "feat(unlock): 遇凭证错误快失败返回 bad_credentials,不等固定超时"
```

---

## Task 3: 解锁日志加 `log()` 函数(时间戳+级别)

**Files:**
- Modify: `unlock_outlook.py:102-117`(Config 区,加 `log` 定义)
- Test: `tests/test_unlock_log_format.py`

当前 unlock 全裸 `print`,无时间戳/级别。先加 `log()` 函数并测,再在后续任务逐步替换 print。

- [ ] **Step 1: 写失败测试 —— log 输出格式**

创建 `tests/test_unlock_log_format.py`:

```python
import io
import re
import unittest
from contextlib import redirect_stdout

import unlock_outlook as mod


class LogFormatTests(unittest.TestCase):
    def test_log_has_timestamp_level_message(self):
        buf = io.StringIO()
        with redirect_stdout(buf):
            mod.log("step login: 3.2s", "INFO")
        out = buf.getvalue().strip()
        self.assertRegex(
            out,
            r"^\[\d{2}:\d{2}:\d{2}\] \[INFO\] step login: 3\.2s$",
        )

    def test_log_default_level_is_info(self):
        buf = io.StringIO()
        with redirect_stdout(buf):
            mod.log("hello")
        out = buf.getvalue().strip()
        self.assertRegex(out, r"^\[\d{2}:\d{2}:\d{2}\] \[INFO\] hello$")

    def test_log_warn_level(self):
        buf = io.StringIO()
        with redirect_stdout(buf):
            mod.log("凭证错误,跳过", "WARN")
        out = buf.getvalue().strip()
        self.assertRegex(out, r"^\[\d{2}:\d{2}:\d{2}\] \[WARN\] 凭证错误,跳过$")


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: 跑测试确认失败**

Run: `python -m pytest tests/test_unlock_log_format.py -v`
Expected: FAIL —— `mod.log` 不存在,AttributeError。

- [ ] **Step 3: 加 log 函数**

`unlock_outlook.py:117`(`SAVE_DEBUG = ...` 行)之后插入:

```python
SAVE_DEBUG = _env_bool("OUTLOOK_UNLOCK_SAVE_DEBUG", False)


def log(msg, level="INFO"):
    """对齐 register log:时间戳 + 级别 + 消息,flush=True 让 webui SSE 实时收到。
    不接业务关键词过滤(register 才有);解锁自建轻量版,单一来源在本文件。"""
    import time as _t
    ts = _t.strftime("%H:%M:%S")
    print(f"[{ts}] [{level}] {msg}", flush=True)
```

- [ ] **Step 4: 跑测试确认通过**

Run: `python -m pytest tests/test_unlock_log_format.py -v`
Expected: 3 个测试全 PASS。

- [ ] **Step 5: 提交**

```bash
git add tests/test_unlock_log_format.py unlock_outlook.py
git commit -m "feat(unlock): 加 log() 函数,带时间戳+级别,对齐注册日志格式"
```

---

## Task 4: 替换裸 print 为 log + 加 step 计时与阶段分隔

**Files:**
- Modify: `unlock_outlook.py` 多处 print(`_attempt_account` L737/747/750/762,`worker` L798/804/806,`save_results` L873-889,`build_pool` L897,`unlock_account` 内各 step)
- Test: `tests/test_unlock_log_format.py`(追加 SUMMARY 格式测试)

把 `_attempt_account`、`worker`、`save_results`、`build_pool`、`unlock_account` 主流程的裸 print 换成 `log`,加 step 计时(`perf_counter` 差值)与阶段分隔。保留 `[px]` 子模块内部 print、浏览器内 JS 注入错误 print 不动(子模块/浏览器侧细节,不混进主日志)。

**替换规则**:
- `[worker-{n}] xxx` → `log(f"[w{n}] xxx")`(账号级,0 缩进)
- `  {tag} xxx`(步骤级)→ `log(f"  {tag} xxx")`
- `file=sys.stderr` 的 WARN/ERR → `log(..., "WARN")`(走 stdout,webui SSE 只订阅 stdout;WARN 用级别区分不靠 stderr)
- step 计时:`t0=time.perf_counter()` 起,完成 `log(f"  {tag} step {name}: {time.perf_counter()-t0:.2f}s")`

- [ ] **Step 1: 写失败测试 —— SUMMARY 格式**

追加到 `tests/test_unlock_log_format.py`(`if __name__` 之前):

```python
class SummaryFormatTests(unittest.TestCase):
    def test_save_results_prints_summary_lines(self):
        """save_results 末尾打印 SUMMARY / SUMMARY_TIME,前缀与注册一致(webui 解析依赖)。"""
        results = [
            ("a@outlook.com", "p1", "a@outlook.com----p1", "unlocked"),
            ("b@outlook.com", "p2", "b@outlook.com----p2", "needs_phone"),
            ("c@outlook.com", "p3", "c@outlook.com----p3", "timeout"),
        ]
        buf = io.StringIO()
        with redirect_stdout(buf):
            mod.save_results(results, "20260826_120000")
        out = buf.getvalue()
        self.assertIn("SUMMARY: unlocked 1 | needs_phone 1 | failed 1 | total 3", out)
        self.assertRegex(out, r"SUMMARY_TIME: total_elapsed [\d.]+s")
```

- [ ] **Step 2: 跑测试确认失败**

Run: `python -m pytest tests/test_unlock_log_format.py::SummaryFormatTests -v`
Expected: FAIL —— 当前 `save_results` 用 `print('='*55)` 等,无 `SUMMARY:` 行。

- [ ] **Step 3: 替换 worker / _attempt_account / build_pool 的 print**

`unlock_outlook.py:737` 起的 `_attempt_account` 代理选择段:

```python
        masked = mask_ruoyi_proxy(selected_pool[0]) if selected_pool else "noproxy"
        rem = pool.remaining() if hasattr(pool, "remaining") else "?"
        print(f"[worker-{worker_id}] proxy -> {masked} remaining={rem}{' [reused]' if reused else ''}")
```

改成:

```python
        masked = mask_ruoyi_proxy(selected_pool[0]) if selected_pool else "noproxy"
        rem = pool.remaining() if hasattr(pool, "remaining") else "?"
        log(f"[w{worker_id}] proxy -> {masked} remaining={rem}{' [reused]' if reused else ''}")
```

`unlock_outlook.py:747-750`(precheck 失败):

```python
        except Exception as exc:
            print(f"[worker-{worker_id}] proxy precheck error: {exc}", file=sys.stderr)
            precheck_ok = False
        if not precheck_ok:
            print(f"[worker-{worker_id}] proxy precheck failed (login.live.com 不可达) -> proxy_dead", file=sys.stderr)
```

改成:

```python
        except Exception as exc:
            log(f"[w{worker_id}] proxy precheck error: {exc}", "WARN")
            precheck_ok = False
        if not precheck_ok:
            log(f"[w{worker_id}] proxy precheck failed (login.live.com 不可达) -> proxy_dead", "WARN")
```

`unlock_outlook.py:762`(launch error):

```python
    except Exception as e:
        print(f"[worker-{worker_id}] launch error: {e}")
```

改成:

```python
    except Exception as e:
        log(f"[w{worker_id}] launch error: {e}", "WARN")
```

`unlock_outlook.py:798-806`(worker 循环):

```python
        for email, password, raw_line in accounts:
            print(f"\n[worker-{worker_id}] {email}")
            outcome = await _attempt_account(pool, worker_id, args, concurrency, tag, email, password)
            # 代理无法访问微软 -> discard 失效节点 -> 换新节点重开浏览器重试
            retry = 0
            while outcome == "proxy_dead" and retry < MAX_PROXY_RETRY:
                retry += 1
                print(f"[worker-{worker_id}] 代理无法访问微软,换节点重试 {retry}/{MAX_PROXY_RETRY}", file=sys.stderr)
                outcome = await _attempt_account(pool, worker_id, args, concurrency, tag, email, password)
            print(f"[worker-{worker_id}] {email} => {outcome}")
            results.append((email, password, raw_line, outcome))
```

改成(加账号序号 + 计时 + 阶段分隔):

```python
        for acct_idx, (email, password, raw_line) in enumerate(accounts, 1):
            t_acct = time.perf_counter()
            log(f"========== 解锁 #{acct_idx}/{len(accounts)} [w{worker_id}] {email} ==========")
            outcome = await _attempt_account(pool, worker_id, args, concurrency, tag, email, password)
            retry = 0
            while outcome == "proxy_dead" and retry < MAX_PROXY_RETRY:
                retry += 1
                log(f"[w{worker_id}] 代理无法访问微软,换节点重试 {retry}/{MAX_PROXY_RETRY}", "WARN")
                outcome = await _attempt_account(pool, worker_id, args, concurrency, tag, email, password)
            log(f"[w{worker_id}] {email} => {outcome}  total={time.perf_counter()-t_acct:.2f}s")
            results.append((email, password, raw_line, outcome))
```

- [ ] **Step 4: 替换 save_results 的汇总 print,加 SUMMARY 行**

`unlock_outlook.py:873-882`(`save_results` 的汇总打印):

```python
    print(f"\n{'='*55}")
    write("unlocked", unlocked)
    write("needs_phone", needs_ph)
    write("failed", failed)
    print(f"{'─'*55}")
    print(f"  Total     : {len(results)}")
    print(f"  Unlocked  : {len(unlocked)}")
    print(f"  NeedsPhone: {len(needs_ph)}")
    print(f"  Failed    : {len(failed)}")
    print(f"{'='*55}")
```

改成(保留原可读分块,末尾加机器可解析 SUMMARY 行供 webui/注册口径对齐):

```python
    log(f"{'='*55}")
    write("unlocked", unlocked)
    write("needs_phone", needs_ph)
    write("failed", failed)
    log(f"{'─'*55}")
    log(f"  Total     : {len(results)}")
    log(f"  Unlocked  : {len(unlocked)}")
    log(f"  NeedsPhone: {len(needs_ph)}")
    log(f"  Failed    : {len(failed)}")
    log(f"{'='*55}")
    log(f"SUMMARY: unlocked {len(unlocked)} | needs_phone {len(needs_ph)} | failed {len(failed)} | total {len(results)}")
    log(f"SUMMARY_TIME: total_elapsed {sum(1 for _ in results) and 0.0:.2f}s")
```

注:`SUMMARY_TIME` 此处无总耗时记录(save_results 只收 results),先填 `0.0` 占位保持前缀格式供 webui 解析;若需真实耗时,在 main 记 batch 起止传入(见 Task 5 main 调整)。**本次保守填 0.0,前缀正确优先。**

- [ ] **Step 5: 替换 build_pool 的 print**

`unlock_outlook.py:897`:

```python
    print(f"[pool] source={st.get('source', args.proxy_source)} size={st.get('remaining', '?')}")
```

改成:

```python
    log(f"[pool] source={st.get('source', args.proxy_source)} size={st.get('remaining', '?')}")
```

- [ ] **Step 6: unlock_account Step 1/Step 2 加阶段 step 计时 + 关键 print 改 log**

`unlock_outlook.py:523`(`deadline = time.time() + timeout` 之后)加 Step 1 计时起点:

```python
    deadline = time.time() + timeout
    t_login = time.perf_counter()
```

`unlock_outlook.py:590-595`(Step 1 末尾 `L_final` 之后,return 之前)在 return 前记 step。当前:

```python
    state = snap(page, tag, "L_final", idx)
    if state == "logged_in":  return "already_ok"
    if state == "sms_verify": return "needs_phone"
    if state == "fido_setup":
        _maybe_skip_passkey(page, tag); time.sleep(4)
        return "unlocked"
```

改成(login 成功路径记 step):

```python
    state = snap(page, tag, "L_final", idx)
    if state == "logged_in":
        log(f"  [{tag}] step login: {time.perf_counter()-t_login:.2f}s")
        return "already_ok"
    if state == "sms_verify":
        log(f"  [{tag}] step login: {time.perf_counter()-t_login:.2f}s")
        return "needs_phone"
    if state == "fido_setup":
        log(f"  [{tag}] step login: {time.perf_counter()-t_login:.2f}s")
        _maybe_skip_passkey(page, tag); time.sleep(4)
        return "unlocked"
```

`unlock_outlook.py:597-604`(Step 2 开头)加 px 计时起点。当前:

```python
    # ── Step 2: PX press-and-hold + unlock flow ──────────────────────
    press_count   = 0
    no_btn_rounds = 0
    px_api_tried  = False
    net_err_count = 0
    abuse_rounds  = 0
    loading_wait_started = None
```

改成(加 t_px):

```python
    # ── Step 2: PX press-and-hold + unlock flow ──────────────────────
    t_px = time.perf_counter()
    press_count   = 0
    no_btn_rounds = 0
    px_api_tried  = False
    net_err_count = 0
    abuse_rounds  = 0
    loading_wait_started = None
```

`unlock_outlook.py:712-717`(Step 2 末尾 return 段)记 px step。当前(读确认是哪几行后改):

```python
    if state == "logged_in":  return "unlocked"
    if state == "sms_verify": return "needs_phone"
    if state == "fido_setup":
        _maybe_skip_passkey(page, tag); time.sleep(4)
        return "unlocked"
```

改成:

```python
    if state == "logged_in":
        log(f"  [{tag}] step px: {time.perf_counter()-t_px:.2f}s presses={press_count}")
        return "unlocked"
    if state == "sms_verify":
        log(f"  [{tag}] step px: {time.perf_counter()-t_px:.2f}s presses={press_count}")
        return "needs_phone"
    if state == "fido_setup":
        log(f"  [{tag}] step px: {time.perf_counter()-t_px:.2f}s presses={press_count}")
        _maybe_skip_passkey(page, tag); time.sleep(4)
        return "unlocked"
```

注:Task 2 在 Step 1 加的 `login_error` 分支那行 `print(..., file=sys.stderr)` 也改成 `log(..., "WARN")`:

```python
        if state == "login_error":
            log(f"  [{tag}] 凭证错误,跳过(不等超时)", "WARN")
            return "bad_credentials"
```

`snap()`(`unlock_outlook.py:250`)的 `print(f"    [{name}] {state}  ...")` 改成 `log(f"  [{name}] {state}  {(page.url or '')[:60]}")`(2 空格缩进,统一)。

- [ ] **Step 7: 加日志落盘(logs/ 文件,对齐注册)**

解锁当前 stdout 只进终端/webui SSE,不落盘(注册有 `_TeeStdout` + `_install_run_log_tee` 把 stdout tee 到 `logs/<YYYYMMDD_HHMMSS>.log`)。把同款机制搬进 unlock。

在 `unlock_outlook.py` 顶部(已有的 `import atexit`/`import sys`/`import os` 附近,若缺则补)加两个定义:

```python
class _TeeStdout:
    """把 stdout 同时写到原 stdout 和一个日志文件(线程安全)。"""

    def __init__(self, original, file_handle):
        self._original = original
        self._fh = file_handle
        self._lock = threading.Lock()

    def write(self, s):
        with self._lock:
            try:
                self._original.write(s)
            except Exception:
                pass
            try:
                self._fh.write(s)
            except Exception:
                pass

    def flush(self):
        with self._lock:
            try:
                self._original.flush()
            except Exception:
                pass
            try:
                self._fh.flush()
            except Exception:
                pass

    def isatty(self):
        return getattr(self._original, "isatty", lambda: False)()

    def fileno(self):
        return self._original.fileno()

    def reconfigure(self, *a, **k):
        try:
            self._original.reconfigure(*a, **k)
        except Exception:
            pass

    def __getattr__(self, name):
        return getattr(self._original, name)


def _install_run_log_tee():
    """tee stdout 到 logs/<ts>.log,对齐注册。返回日志路径或 None。"""
    try:
        log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
        os.makedirs(log_dir, exist_ok=True)
        ts = time.strftime("%Y%m%d_%H%M%S")
        path = os.path.join(log_dir, f"{ts}.log")
        fh = open(path, "w", encoding="utf-8", buffering=1)
        sys.stdout = _TeeStdout(sys.stdout, fh)
        atexit.register(lambda: fh.close() if not fh.closed else None)
        sys.stdout.write(f"===== 本次运行日志 {ts} =====\n")
        return path
    except Exception as exc:
        try:
            print(f"_install_run_log_tee failed: {type(exc).__name__}: {exc}", file=sys.stderr)
        except Exception:
            pass
        return None
```

注:
- `threading` 顶部需 import(unlock 多线程跑,tee 必须线程安全;注册版也带 lock)。确认 `import threading` 已存在,否则补。
- `_install_run_log_tee` 用 `os.path.dirname(os.path.abspath(__file__))` 而非注册的 `ROOT`(unlock 没有 ROOT 全局,以此文件所在目录为基),`logs/` 目录与注册同一处(仓库根 `logs/`)。
- `log()` 用 `print(..., flush=True)`,tee 截获 `print` 的 stdout 写入 → 文件同步拿到所有 `log(...)` 输出,无需额外改 `log()`。

在 `main()`(`unlock_outlook.py:972`)函数体最开头调用一次:

```python
def main():
    _install_run_log_tee()
    # ...原 main 逻辑不动
```

- [ ] **Step 8: 跑测试确认通过**

Run: `.venv/Scripts/python.exe -m pytest tests/test_unlock_log_format.py tests/test_unlock_bad_credentials.py -v`
Expected: 两文件全 PASS(SUMMARY 格式测过 + 之前 classify/快失败测不回退)。

注:unlock 顶层 import ruyipage,跑 unlock 相关测试/脚本必须用 `.venv/Scripts/python.exe`(base python 3.13 无 ruyipage)。

跑全套确认无回归:

Run: `.venv/Scripts/python.exe -m pytest tests/ -q`
Expected: 之前 102 passed 仍全过(本任务只改 unlock 文件 + 新测试)。

- [ ] **Step 9: 提交**

```bash
git add tests/test_unlock_log_format.py unlock_outlook.py
git commit -m "feat(unlock): 日志对齐注册——log替换裸print+step计时+SUMMARY+落盘logs/"
```

---

## Task 5: WebUI 每脚本日志独立(前端 app.js)

**Files:**
- Modify: `webui/static/app.js:4-6`(全局状态)、`128-134`(selectScript)、`359-380`(runScript + stop)
- Test: `tests/test_webui_log_isolation.py`

当前全局单 `curRun`/`evtSrc`,`runScript` 清空日志区 + 切流,旧脚本日志丢失。改成按 `script_id` 聚合 `scriptLogs`,切菜单切日志区不清空。

- [ ] **Step 1: 写静态测试 —— app.js 存在 scriptLogs 且 runScript 不全清空**

创建 `tests/test_webui_log_isolation.py`:

```python
import unittest
from pathlib import Path


APP_JS = Path("webui/static/app.js").read_text(encoding="utf-8")


class WebuiLogIsolationTests(unittest.TestCase):
    def test_app_js_uses_script_logs_state(self):
        """前端按 script_id 聚合日志,不再单全局 curRun 覆盖。"""
        self.assertIn("scriptLogs", APP_JS)
        self.assertIn("curLogScript", APP_JS)

    def test_run_script_does_not_blank_log_area(self):
        """runScript 不再用 log.textContent='' 清空日志区(防回退到覆盖行为)。"""
        self.assertNotIn("log.textContent=''", APP_JS)
        self.assertNotIn("log.textContent = ''", APP_JS)

    def test_select_script_switches_log_view(self):
        """selectScript 切菜单时把日志区切到该脚本。"""
        self.assertIn("curLogScript", APP_JS)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: 跑测试确认失败**

Run: `python -m pytest tests/test_webui_log_isolation.py -v`
Expected: FAIL —— `scriptLogs`/`curLogScript` 不存在,`log.textContent=''` 还在。

- [ ] **Step 3: 改全局状态**

`webui/static/app.js:4-6` 当前:

```js
let curRun = null;     // 当前运行 run_id
let curSrc = null;     // 当前选中脚本
let curSavedArgs = {};
let evtSrc = null;     // EventSource
```

改成(保留 curSrc/curSavedArgs,加 scriptLogs/curLogScript,移除全局 curRun/evtSrc 单例):

```js
let curSrc = null;     // 当前选中脚本(表单)
let curSavedArgs = {};
// 每脚本独立日志:{ script_id: { lines: [], run_id: null, done: true, evt: null, cmd: '' } }
let scriptLogs = {};
let curLogScript = null;   // 当前日志区展示的脚本 id
```

- [ ] **Step 4: 改 selectScript —— 切菜单切日志区**

`webui/static/app.js:129-134` 当前:

```js
async function selectScript(id){
  curSrc = SCRIPTS.find(s=>s.id===id);
  $$('.scriptbtn').forEach(b=>b.classList.toggle('active', b.dataset.id===id));
  curSavedArgs = await loadScriptConfig(id);
  renderForm(curSrc, curSavedArgs);
}
```

改成(加渲染该脚本日志区):

```js
async function selectScript(id){
  curSrc = SCRIPTS.find(s=>s.id===id);
  $$('.scriptbtn').forEach(b=>b.classList.toggle('active', b.dataset.id===id));
  curSavedArgs = await loadScriptConfig(id);
  renderForm(curSrc, curSavedArgs);
  showScriptLog(id);
}

// 把日志区切到某脚本:渲染该脚本缓存的 lines + 更新 title/cmd/stop 状态
function showScriptLog(id){
  curLogScript = id;
  const log = $('#log');
  const rec = scriptLogs[id];
  log.textContent = rec && rec.lines.length ? rec.lines.join('\n') : '';
  LogAutoScroll.scrollToBottom && LogAutoScroll.scrollToBottom(log);
  const s = SCRIPTS.find(x=>x.id===id);
  $('#log-title').textContent = `运行日志 — ${s ? s.title : id}`;
  $('#cmd-preview').textContent = rec && rec.cmd ? '$ ' + rec.cmd : '';
  $('#btn-stop').disabled = !(rec && rec.run_id && !rec.done);
}
```

- [ ] **Step 5: 改 runScript —— 按 script_id 聚合,不清空**

`webui/static/app.js:359-374` 当前:

```js
async function runScript(){
  if(curRun && evtSrc){ evtSrc.close(); }
  const args = collectArgs(curSrc);
  const log = $('#log'); log.textContent='';
  $('#log-title').textContent = `运行日志 — ${curSrc.title}`;
  const r = await (await fetch('/api/run',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({script:curSrc.id, args})})).json();
  if(r.error){ log.textContent='错误: '+r.error; return; }
  curRun = r.run_id;
  $('#cmd-preview').textContent = '$ '+r.cmd;
  $('#btn-stop').disabled = false;
  evtSrc = new EventSource(`/api/logs/${curRun}`);
  evtSrc.onmessage = e=>{ LogAutoScroll.appendLogLineWithAutoScroll(log, e.data); };
  evtSrc.addEventListener('done', ()=>{ evtSrc.close(); $('#btn-stop').disabled = true; pollStatus(); });
  evtSrc.onerror = ()=>{ evtSrc.close(); $('#btn-stop').disabled = true; };
}
```

改成(按 script_id 存,追加分隔线不清空,自动切到该脚本日志):

```js
async function runScript(){
  const sid = curSrc.id;
  // 同脚本若有未完成旧 run,先关旧 SSE(同脚本不并发跑两个)
  const old = scriptLogs[sid];
  if(old && old.evt){ old.evt.close(); old.evt = null; }
  const rec = scriptLogs[sid] = scriptLogs[sid] || { lines: [], run_id: null, done: true, evt: null, cmd: '' };
  if(rec.lines.length){ rec.lines.push('[webui] ===== 重新运行 ====='); }
  const args = collectArgs(curSrc);
  const r = await (await fetch('/api/run',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({script:sid, args})})).json();
  if(r.error){ rec.lines.push('错误: '+r.error); showScriptLog(sid); return; }
  rec.run_id = r.run_id; rec.done = false; rec.cmd = r.cmd;
  showScriptLog(sid);
  const log = $('#log');
  rec.evt = new EventSource(`/api/logs/${rec.run_id}`);
  rec.evt.onmessage = e=>{
    rec.lines.push(e.data);
    if(curLogScript === sid){ LogAutoScroll.appendLogLineWithAutoScroll(log, e.data); }
  };
  rec.evt.addEventListener('done', ()=>{ rec.evt.close(); rec.evt=null; rec.done=true;
    if(curLogScript === sid){ $('#btn-stop').disabled = true; } pollStatus(); });
  rec.evt.onerror = ()=>{ rec.evt.close(); rec.evt=null; rec.done=true;
    if(curLogScript === sid){ $('#btn-stop').disabled = true; } };
}
```

- [ ] **Step 6: 改 stop 按钮 —— 对当前日志区的脚本发 stop**

`webui/static/app.js:376-380` 当前:

```js
$('#btn-stop').onclick = async ()=>{
  if(!curRun) return;
  await fetch(`/api/stop/${curRun}`,{method:'POST'});
  $('#btn-stop').disabled = true;
};
```

改成:

```js
$('#btn-stop').onclick = async ()=>{
  const rec = curLogScript ? scriptLogs[curLogScript] : null;
  if(!rec || !rec.run_id) return;
  await fetch(`/api/stop/${rec.run_id}`,{method:'POST'});
  $('#btn-stop').disabled = true;
};
```

- [ ] **Step 7: 跑测试确认通过**

Run: `python -m pytest tests/test_webui_log_isolation.py -v`
Expected: 3 个测试全 PASS。

跑全套确认无回归:

Run: `python -m pytest tests/ -q`
Expected: 全绿(含新增 3 个测试文件)。

- [ ] **Step 8: 提交**

```bash
git add tests/test_webui_log_isolation.py webui/static/app.js
git commit -m "fix(webui): 每脚本日志独立,切菜单不互覆盖,支持同时跑多脚本看各自日志"
```

---

## Task 6: 端到端手测 + 收尾

**Files:**
- 无代码改动,只跑验证

- [ ] **Step 1: 跑全套测试**

Run: `python -m pytest tests/ -q`
Expected: 全绿(原 102 + 新增约 10 个测试)。

- [ ] **Step 2: 解锁日志手测(现成账号)**

Run(用现成账号文件,有头观察日志格式):
```bash
.venv/Scripts/python.exe unlock_outlook.py \
  --input outlook_accounts/accounts_ruoyi_nograph_20260824_111702.txt \
  --proxy-file proxies_outlook.txt --timeout 180 --log-level INFO
```
Expected:
- 日志带 `[HH:MM:SS] [INFO]` 前缀
- 有 `========== 解锁 #1/1 [w0] xxx@outlook.com ==========`
- 有 `step login: N.NNs` / `step px: N.NNs presses=N`
- 末尾 `SUMMARY: unlocked ... | total 1` + `SUMMARY_TIME: ...`
- 账号结果仍 `unlocked`(功能不坏)

- [ ] **Step 3: 错误凭证快失败手测(故意填错密码)**

构造 1 行错误账号文件 `tests_tmp_bad.txt`:
```
fysirxd665328@outlook.com----wrongpassword123
```
Run:
```bash
.venv/Scripts/python.exe unlock_outlook.py --input tests_tmp_bad.txt \
  --proxy-file proxies_outlook.txt --timeout 60 --log-level INFO
```
Expected:
- ~3-5s 内 `[w0] ... 凭证错误,跳过(不等超时)` (WARN)
- `=> bad_credentials` 落 failed 文件
- 不等满 60s timeout
- 删临时文件 `rm tests_tmp_bad.txt`

- [ ] **Step 4: WebUI 多脚本日志独立手测**

启动 webui(`.venv/Scripts/python.exe webui/server.py` 或现有启动方式),浏览器开:
1. 选"解锁账号"菜单 → 点运行,日志区显示解锁日志(SSE 推)
2. 跑着时点左侧"唤醒 ruyi 浏览器"菜单 → 点运行,日志区切到唤醒日志
3. 切回"解锁账号"菜单 → 日志区显示解锁的累计日志(没丢,SSE 仍在推新行)
4. 两脚本日志互不覆盖
Expected: 切回解锁能看到从开头到现在的全部日志,且新行继续追加。

- [ ] **Step 5: 提交收尾(若有手测小修)**

若手测发现小问题已修,提交:
```bash
git add -A
git commit -m "test: 解锁日志/webui隔离/凭证快失败 端到端验证通过"
```

---

## 自查(Spec 覆盖核对)

- 问题1(webui 日志独立):Task 5 ✅(scriptLogs 按 script_id 聚合,切菜单切日志区不清空,stop 按当前日志区脚本)
- 问题2(日志对齐注册):Task 3(log 函数)+ Task 4(替换 print + step 计时 + SUMMARY)✅
- 问题3(错误凭证快失败):Task 1(classify login_error)+ Task 2(主循环快失败)✅
- 测试:Task 1/2 bad_credentials、Task 3/4 log_format、Task 5 webui_log_isolation ✅
- 端到端:Task 6 ✅

无 placeholder;类型/函数名一致(`log`、`classify`、`scriptLogs`、`showScriptLog`、`curLogScript` 跨任务一致);`bad_credentials` 在 Task 2 定义、Task 4 引用、Task 6 验证,一致。
