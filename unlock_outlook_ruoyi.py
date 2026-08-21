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
    """同步:截图到 SCREENSHOT_DIR,取 body 文案,分类状态。返回 (state, text)。"""
    os.makedirs(SCREENSHOT_DIR, exist_ok=True)
    try:
        rr._shot(page, f"{tag}_{name}", 0)
    except Exception:
        pass
    url = page.url or ""
    text = rr._body_text(page) or ""
    state = classify(text, url)
    print(f"    [{name}] {state}  {url[:60]}")
    return state, text


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
    print(f"[unlock-ruoyi] input={args.input} limit={args.limit} conc={args.concurrency} "
          f"headless={args.headless} proxy_url={'yes' if args.proxy_url else 'no'}")

if __name__ == "__main__":
    main()
