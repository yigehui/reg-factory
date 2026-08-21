# -*- coding: utf-8 -*-
"""
Outlook Account Batch Unlock — ruoyi 版
用 ruyipage 定制 Firefox + ruoyi _perform_hold 按住验证解锁被锁账号。
代理走 --proxy-url 池(同 ruoyi 注册)。结果输出 unlock_results/。

Usage:
  # 抽样20回归
  python unlock_outlook_ruoyi.py --input email_abuse.txt --limit 20 --concurrency 2 --headless \
    --proxy-url "http://163.192.58.188:8787/api/pool/proxies/text?token=...&return_type=socks5"
  # 全量
  python unlock_outlook_ruoyi.py --input email_abuse.txt --concurrency 2 --headless --proxy-url "..."
  # WebUI 默认 input=email_abuse.txt, proxy-url 走 .env OUTLOOK_PROXY_URL
"""
import argparse, asyncio, os, sys, time
from types import SimpleNamespace

if sys.platform == "win32":
    # hasattr 守卫:pytest 默认捕获 stdin 时 sys.stdin 是 DontReadFromInput,
    # 无 reconfigure 属性,不守卫会让所有 import 本模块的测试在收集期崩溃(需 -s 才能跑)。
    if hasattr(sys.stdout, "reconfigure"): sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stdin, "reconfigure"):  sys.stdin.reconfigure(encoding="utf-8")

# 吃项目根 .env(不覆盖已有环境变量),同 launch_ruoyi_browser
ROOT = os.path.dirname(os.path.abspath(__file__))
def _load_dotenv():
    try:
        p = os.path.join(ROOT, ".env")
        if not os.path.isfile(p): return
        for line in open(p, encoding="utf-8"):
            s = line.strip()
            if not s or s.startswith("#") or "=" not in s: continue
            k, _, v = s.partition("=")
            k = k.strip(); v = v.strip().strip('"').strip("'")
            if k and k not in os.environ: os.environ[k] = v
    except Exception: pass
_load_dotenv()

import register_outlook_ruoyi as rr  # 库式复用,模块级 import 不触发注册流程
from ruyipage import FirefoxOptions, FirefoxPage

# ── Config ───────────────────────────────────────────────────────────
OUTPUT_DIR     = "unlock_results"
SCREENSHOT_DIR = "screenshots_unlock"
UNLOCK_TIMEOUT = 300   # seconds per account
DEFAULT_INPUT  = "email_abuse.txt"


# ── Page state classifier(从 unlock_outlook.py 移植,锁号场景已验证有效)──
def classify(text, url):
    t, u = text.lower(), url.lower()
    if "account.microsoft.com" in u and "unlock" not in u: return "logged_in"
    if "account.live.com" in u and "proofs" in u:          return "logged_in"
    if "fido/create" in u or "fido/update" in u:           return "fido_setup"
    if "setting up your passkey" in t or "passkey" in t:   return "fido_setup"
    if any(x in t for x in ["your account has been locked", "we've locked",
                              "locked for your protection", "帐户已锁定"]): return "locked"
    if "let's prove you're human" in t or "press and hold" in t: return "px_challenge"
    if any(x in t for x in ["enter the code", "we texted", "we sent", "verification code",
                              "验证码", "短信"]): return "sms_verify"
    if any(x in t for x in ["verify your identity", "unusual activity"]): return "verify_needed"
    if "something went wrong" in t: return "error_page"
    if "chrome-error://" in u:      return "net_error"
    if "enter your password" in t:  return "login_form"
    if any(x in t for x in ["email or phone", "sign in", "enter your email"]): return "email_form"
    return "unknown"


# ── 快照:截图 + 取正文 + 分类 ──────────────────────────────────────────
def snap(page, tag, name):
    """同步:截图到 SCREENSHOT_DIR,取 body 文案,分类状态。返回 (state, text)。

    注意:不直接复用 rr._shot——它内部走库的 SCREENSHOT_DIR(screenshots_ruoyi/)
    且按 failure 前缀名+OUTLOOK_PX_PRESS_SCREENSHOTS 过滤,本模块的 tag_name
    (如 w0_L00)不匹配前缀,会被静默跳过→截图目录空。这里用 page.screenshot
    直写 SCREENSHOT_DIR,每帧都留现场图(spec §7 验收#3:px_challenge/locked 现场佐证)。
    """
    os.makedirs(SCREENSHOT_DIR, exist_ok=True)
    try:
        page.screenshot(path=os.path.join(SCREENSHOT_DIR, f"{tag}_{name}.png"), full_page=True)
    except Exception:
        pass
    url = page.url or ""
    text = rr._body_text(page) or ""
    state = classify(text, url)
    print(f"    [{name}] {state}  {url[:60]}")
    return state, text


# ── FIDO 跳过(复用 ruoyi _maybe_skip_passkey,基于 url 子串 + 点 Skip)──────
# 同步实现:被 unlock_account_sync 在 to_thread 里调用,不能 asyncio.run(会报已有事件循环)。
def skip_fido(page):
    if rr._maybe_skip_passkey(page, "[unlock]"):
        time.sleep(4)
        return True
    try:
        page.get("https://account.microsoft.com/", timeout=20000)
        return True
    except Exception:
        return False


# ── 填邮箱/密码的同步小工具(用 ruoyi _ele/_safe_input/_click_any)────────
def _fill_email_sync(page, email):
    el = rr._ele(page, 'input[type="email"],input[name="loginfmt"]', timeout=8)
    if el is None:
        return False
    rr._safe_input(el, email)
    try:
        page.actions.press("").perform()  # Enter(Selenium Keys.ENTER)
    except Exception:
        rr._click_any(page, ['#idSIButton9', 'text:Next', 'input[type="submit"]', 'button[type="submit"]'], timeout=3)
    return True

def _fill_password_sync(page, password):
    el = rr._ele(page, 'input[type="password"]', timeout=8)
    if el is None:
        return False
    cur = rr._read_input_value(el)  # ruyipage 元素无 .attrs.value,用 ruoyi 读取器
    if not cur:
        rr._safe_input(el, password)
    try:
        page.actions.press("").perform()  # Enter
    except Exception:
        rr._click_any(page, ['#idSIButton9', 'text:Next', 'input[type="submit"]', 'button[type="submit"]'], timeout=3)
    return True


# ── 核心解锁逻辑(同步,跑在 to_thread 里)──────────────────────────────
def unlock_account_sync(page, ctx, email, password, tag, *, max_press=5):
    os.makedirs(SCREENSHOT_DIR, exist_ok=True)

    # ── Step 1: 登录 ──────────────────────────────────────────────
    page.get("https://login.live.com/login.srf", timeout=60000)
    time.sleep(2)

    net_err_count = 0
    for i in range(20):
        state, _ = snap(page, tag, f"L{i:02d}")
        if state == "logged_in":  return "already_ok"
        if state == "sms_verify": return "needs_phone"
        if state == "fido_setup":
            skip_fido(page); return "unlocked"
        if state in ("locked", "px_challenge"): break
        if state == "net_error":
            net_err_count += 1
            if net_err_count >= 3: return "failed_net_error"
            page.get("https://login.live.com/login.srf", timeout=60000)
            time.sleep(3); continue
        if state == "email_form":
            _fill_email_sync(page, email); time.sleep(3); continue
        if state == "login_form":
            _fill_password_sync(page, password); time.sleep(3); continue
        # 通用 Next/submit
        if rr._click_any(page, ['#idSIButton9', 'text:Next', 'input[type="submit"]', 'button[type="submit"]'], timeout=3):
            time.sleep(2); continue
        time.sleep(2)

    state, _ = snap(page, tag, "L_final")
    if state == "logged_in":  return "already_ok"
    if state == "sms_verify": return "needs_phone"
    if state == "fido_setup":
        skip_fido(page); return "unlocked"

    # ── Step 2: PX 按住 + 解锁流(纯浏览器,无 EZCaptcha)─────────────
    press_count   = 0
    no_btn_rounds = 0
    net_err_count = 0
    err_page_count = 0

    for i in range(60):
        state, _ = snap(page, tag, f"U{i:02d}")

        if state == "logged_in":  return "unlocked"
        if state == "sms_verify": return "needs_phone"
        if state == "fido_setup":
            skip_fido(page); return "unlocked"
        if state == "error_page":
            err_page_count += 1
            if err_page_count >= 3:
                return "failed_error_page"
            if not rr._click_any(page, ['text:Try again', 'text:重试', 'text:再试一次'], timeout=3):
                # 不用 history.back()——它会把半登录态的 login.srf 当成 login_form,
                # 之后通用循环不停重填密码却再也回不到 Abuse(死循环→占满 20 轮)。
                # 改为重新导航到 Abuse 解锁入口,重建流程。
                try: page.get("https://account.live.com/Abuse", timeout=60000)
                except Exception: pass
                time.sleep(4)
            else:
                time.sleep(5)
            continue
        if state == "net_error":
            net_err_count += 1
            if net_err_count >= 5: return "failed_net_error"
            page.get("https://login.live.com/login.srf", timeout=60000)
            time.sleep(3); continue

        if state == "locked":
            rr._click_any(page, ['button[type="submit"]', 'text:Next', 'text:下一步', 'input[type="submit"]'], timeout=3)
            time.sleep(6); continue

        if state == "px_challenge":
            if press_count < max_press:
                hold_ctx, target = rr._find_hold_context(page)
                if target:
                    press_count += 1
                    try:
                        rr._perform_hold(page, hold_ctx, target, 0, press_count, "[unlock]")
                    except Exception as e:
                        print(f"    [unlock] press #{press_count} error: {e}")
                    print(f"    held (#{press_count}/{max_press})")
                    rr._wait_before_next_captcha_press("[unlock]", "challenge failed")
                    no_btn_rounds = 0
                else:
                    no_btn_rounds += 1
                    print(f"    no hold target (round {no_btn_rounds})")
                    if no_btn_rounds >= 5:
                        print("    no target 5 rounds — give up"); break
                    time.sleep(3)
            else:
                print(f"    all {max_press} PX presses exhausted"); break
            continue

        # 通用 next/submit
        if rr._click_any(page, ['button[type="submit"]', 'text:Next', 'text:下一步', '#idSIButton9', 'input[type="submit"]'], timeout=3):
            time.sleep(4)
        else:
            time.sleep(3)

    state, _ = snap(page, tag, "U_final")
    if state == "logged_in":  return "unlocked"
    if state == "sms_verify": return "needs_phone"
    if state == "fido_setup":
        skip_fido(page); return "unlocked"
    return f"failed_{state}"


# ── 浏览器启动:有头/无头都支持 ─────────────────────────────────────────
def _build_unlock_options(opts, idx, *, proxy_str, headless):
    tb = FirefoxOptions()
    tb.set_browser_path(rr.RUOYI_FIREFOX_PATH)
    profile_dir = rr._ruoyi_profile_dir(opts, idx)
    tb.set_profile(profile_dir)
    if proxy_str:
        tb.set_proxy(proxy_str)
        print(f"[unlock-ruoyi] proxy: {rr.mask_ruoyi_proxy(proxy_str)}")
    else:
        print("[unlock-ruoyi] 没挂代理——直接本机出口", file=sys.stderr)
    ua = rr._pick_user_agent(idx)
    rr._apply_ruoyi_browser_ua(tb, "[unlock]", ua)
    rr._apply_ruoyi_quiet_prefs(tb, "[unlock]")
    if headless:
        rr._apply_ruoyi_headless_options(tb, "[unlock]", user_agent=ua)
        tb.headless(True)
    return tb, profile_dir


# ── 代理摄入:复用 ruoyi fetch_proxy_list_http / parse_proxy_pool / _proxy_url_to_ruoyi ─
def _load_unlock_proxies(args):
    explicit = str(getattr(args, "proxy", "") or "").strip()
    if explicit:
        norm = rr._proxy_url_to_ruoyi(explicit)
        if not norm:
            raise SystemExit(f"[ERR] 非法代理 --proxy: {explicit}")
        return [norm]
    url = str(getattr(args, "proxy_url", "") or "").strip()
    if url:
        print(f"[unlock-ruoyi] 拉取代理池: {url[:80]}...")
        return rr.fetch_proxy_list_http(url)
    pf = str(getattr(args, "proxy_file", "") or "").strip()
    if pf:
        return rr.parse_proxy_pool(pf)
    return []  # 无代理直连


# ── File I/O(从 unlock_outlook.py 移植,新增 limit 切片)──────────────────
def load_accounts(path, limit=0):
    accounts = []
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"): continue
            parts = line.split("----")
            if len(parts) >= 2:
                accounts.append((parts[0].strip(), parts[1].strip(), line))
            else:
                print(f"[warn] skip: {line[:60]}")
    if limit and limit > 0:
        accounts = accounts[:limit]
    return accounts

def save_results(results, ts):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    unlocked = [r for r in results if r[3] in ("unlocked", "already_ok")]
    needs_ph = [r for r in results if r[3] == "needs_phone"]
    failed   = [r for r in results if r[3] not in ("unlocked", "already_ok", "needs_phone")]

    def write(name, rows):
        p = os.path.join(OUTPUT_DIR, f"{name}_{ts}.txt")
        with open(p, "w", encoding="utf-8") as f:
            for email, password, raw, outcome in rows:
                f.write(f"{raw}----{outcome}\n")
        print(f"  {name:<22s} {len(rows):4d}  -> {p}")

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

    ok_path = os.path.join(OUTPUT_DIR, f"unlocked_clean_{ts}.txt")
    with open(ok_path, "w", encoding="utf-8") as f:
        for email, password, _, _ in unlocked:
            f.write(f"{email}----{password}\n")
    if unlocked:
        print(f"\n  Clean unlocked list: {ok_path}")


# ── Worker:单账号全生命周期 ─────────────────────────────────────────────
async def worker(accounts, proxy_list, worker_id, results, sem, *, args, launch_lock, next_launch_at):
    async with sem:
        for email, password, raw_line in accounts:
            tag = f"w{worker_id}"
            print(f"\n[worker-{worker_id}] {email}")
            pid_profile_dir = None
            browser_page = None
            selected_proxy = None
            outcome = None
            try:
                # 取代理(按 email hash 轮询 proxy_list,同号尽量同出口)
                if proxy_list:
                    selected_proxy = proxy_list[hash(email) % len(proxy_list)]
                    print(f"[worker-{worker_id}] proxy -> {rr.mask_ruoyi_proxy(selected_proxy)}")
                    # 探活(失败换下一条,最多试 3 条)
                    probe_ok = False
                    tried = set()
                    for _ in range(3):
                        if not selected_proxy or selected_proxy in tried:
                            break
                        tried.add(selected_proxy)
                        if await asyncio.to_thread(rr._probe_proxy_before_browser, [selected_proxy], f"[w{worker_id}]"):
                            probe_ok = True; break
                        idx_next = (proxy_list.index(selected_proxy) + 1) % len(proxy_list)
                        selected_proxy = proxy_list[idx_next]
                    if not probe_ok and proxy_list:
                        print(f"[worker-{worker_id}] 代理探活全失败,回退直连")
                        selected_proxy = None

                # 启动错峰(规避 XPCOM 文件锁)
                async with launch_lock:
                    wait = max(0.0, next_launch_at[0] - time.time())
                    if wait > 0:
                        await asyncio.sleep(wait)
                    next_launch_at[0] = time.time() + args.launch_stagger

                run_args = SimpleNamespace(**vars(args))
                run_args.concurrency = args.concurrency
                tb, profile_dir = _build_unlock_options(run_args, worker_id, proxy_str=selected_proxy, headless=args.headless)
                pid_profile_dir = profile_dir

                # 同步启动 + 跑(放 to_thread,不阻塞 loop)
                def _run():
                    bp = FirefoxPage(tb)
                    rr._track_browser_page(bp)
                    try:
                        bp.close_other_tabs(bp)
                    except Exception:
                        pass
                    page = bp
                    # 装请求拦截(对齐 ruoyi 注册:用 rr._start_ruoyi_resource_blocking)
                    try:
                        rr._start_ruoyi_resource_blocking(page, "[unlock]")
                    except Exception:
                        pass
                    setattr(page, "_ruoyi_px_motion_profile", rr._new_ruoyi_px_motion_profile())
                    return bp, page
                browser_page, page = await asyncio.to_thread(_run)

                outcome = await asyncio.wait_for(
                    asyncio.to_thread(unlock_account_sync, page, browser_page, email, password, tag, max_press=args.max_press),
                    timeout=args.timeout)
                print(f"[worker-{worker_id}] {email} => {outcome}")

            except asyncio.TimeoutError:
                outcome = "timeout"
                print(f"[worker-{worker_id}] {email} => timeout")
            except Exception as e:
                outcome = f"error: {str(e)[:80]}"
                print(f"[worker-{worker_id}] {email} => error: {e}")
            finally:
                if browser_page is not None:
                    try:
                        await asyncio.to_thread(lambda: browser_page.quit(timeout=5))
                    except Exception:
                        pass
                if pid_profile_dir:
                    try:
                        await asyncio.to_thread(rr._cleanup_ruoyi_run_profile_dir, pid_profile_dir)
                    except Exception:
                        pass
                results.append((email, password, raw_line, outcome or "failed_unknown"))


# ── 并发编排:按 concurrency 切片,gather 所有 worker ────────────────────
async def run(accounts, proxy_list, args):
    if not accounts:
        print("[error] no accounts found"); return
    print(f"Input     : {args.input}")
    print(f"Accounts  : {len(accounts)}" + (f" (limit {args.limit})" if args.limit else ""))
    print(f"Concurrency: {args.concurrency}")
    print(f"Proxies   : {len(proxy_list)}" + (" (直连)" if not proxy_list else ""))

    results = []
    sem = asyncio.Semaphore(args.concurrency)
    chunks = [[] for _ in range(args.concurrency)]
    for i, acc in enumerate(accounts):
        chunks[i % args.concurrency].append(acc)
    launch_lock = asyncio.Lock()
    next_launch_at = [time.time()]

    await asyncio.gather(*[
        worker(chunks[i], proxy_list, i, results, sem,
               args=args, launch_lock=launch_lock, next_launch_at=next_launch_at)
        for i in range(args.concurrency) if chunks[i]
    ])

    from datetime import datetime
    save_results(results, datetime.now().strftime("%Y%m%d_%H%M%S"))


def build_parser():
    ap = argparse.ArgumentParser(
        description="批量解锁被锁 Outlook(ruyipage Firefox + ruoyi 按住)",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--input", "-i", default=DEFAULT_INPUT,
        help=f"账号文件(每行 email----password...;默认 {DEFAULT_INPUT})")
    ap.add_argument("--limit", type=int, default=0,
        help="只跑前 N 个(0=全部);抽样回归用")
    ap.add_argument("--concurrency", "-c", type=int, default=2, help="并发数(默认2)")
    ap.add_argument("--proxy-url", default=os.environ.get("OUTLOOK_PROXY_URL", ""),
        help="HTTP 拉取代理池地址(return_type=socks5);留空读 .env OUTLOOK_PROXY_URL")
    ap.add_argument("--proxy-file", default=os.environ.get("OUTLOOK_PROXY_FILE", ""),
        help="本地代理池文件(与 --proxy-url 二选一)")
    ap.add_argument("--proxy", default=os.environ.get("LAUNCH_UPSTREAM_PROXY", ""),
        help="单条代理(整浏览器走它);优先级最高")
    ap.add_argument("--headless", action="store_true", default=False,
        help="无头模式(默认有头;abuse 风控重,有头更稳)")
    ap.add_argument("--max-press", type=int,
        default=int(os.environ.get("OUTLOOK_REG_MAX_PRESS", "5") or 5),
        help="按住最大次数(默认5)")
    ap.add_argument("--timeout", type=float, default=UNLOCK_TIMEOUT,
        help="单账号超时秒数(默认300)")
    ap.add_argument("--launch-stagger", type=float, default=float(os.environ.get("OUTLOOK_LAUNCH_STAGGER", "3.0") or 3.0),
        help="浏览器启动错峰秒数(规避 XPCOM 文件锁)")
    ap.add_argument("--log-level", default=os.environ.get("OUTLOOK_LOG_LEVEL", "INFO"),
        choices=["DEBUG", "INFO", "WARN", "PROD", "ERR"], help="日志等级")
    return ap


def main():
    args = build_parser().parse_args()
    rr.set_log_level(args.log_level)
    rr._install_shutdown_handlers()
    if not os.path.isfile(rr.RUOYI_FIREFOX_PATH):
        print(f"[ERR] 定制 Firefox 内核不存在: {rr.RUOYI_FIREFOX_PATH}", file=sys.stderr)
        print("请先运行: .venv\\Scripts\\python.exe -m ruyipage install", file=sys.stderr)
        sys.exit(1)

    if not os.path.exists(args.input):
        print(f"[error] 账号文件不存在: {args.input}", file=sys.stderr)
        sys.exit(1)

    accounts = load_accounts(args.input, limit=args.limit)
    if not accounts:
        print("[info] no accounts to unlock."); sys.exit(0)

    # 启动前清理残留 ruoyi Firefox(对齐 ruoyi main 的 _force_kill_ruoyi_firefox)
    try:
        rr._force_kill_ruoyi_firefox()
    except Exception:
        pass

    proxy_list = _load_unlock_proxies(args)
    print(f"[unlock-ruoyi] input={args.input} limit={args.limit} conc={args.concurrency} "
          f"headless={args.headless} proxies={len(proxy_list)}")

    try:
        asyncio.run(run(accounts, proxy_list, args))
    except KeyboardInterrupt:
        print("\n[unlock-ruoyi] 收到 Ctrl-C,退出")

if __name__ == "__main__":
    main()
