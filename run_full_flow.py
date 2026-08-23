# -*- coding: utf-8 -*-
"""
run_full_flow.py — 端到端全流程编排（含邮箱注册）

把项目原有的两个阶段串成一条龙，方便"跑一遍看看"：

  Stage A  邮箱注册   outlook_reg_loop.py        -> 产出新 outlook 号写进 emails.txt
  Stage B  平台注册   register_three_platforms   -> 用该号去 claude/chatgpt/grok 注册

Stage A 本身是个常驻循环，这里把它当子进程拉起、盯着 emails.txt，**一旦冒出
一个新的可用号就立刻杀掉循环**进入 Stage B，所以是"注册到一个邮箱就往下走"。

前置：BitBrowser(54345) 在线、Clash Verge(控制器 9097 / 混合端口 7897) 在线。
默认自动注入 HTTP(S)_PROXY 与 CLASH_API/SECRET/GROUP，让邮箱注册能换节点绕 MS 风控。

用法：
  python run_full_flow.py                          # 注册1个邮箱 -> 在 claude 上注册
  python run_full_flow.py --platforms claude chatgpt
  python run_full_flow.py --platforms chatgpt --rounds 10   # 循环注册 10 个号
  python run_full_flow.py --platforms chatgpt --rounds 0    # 无限循环（Ctrl+C 停）
  python run_full_flow.py --skip-email --email a@outlook.com --password xxx   # 跳过邮箱注册
  python run_full_flow.py --email-attempts 20 --email-timeout 180
  python run_full_flow.py --dry-run                # 只打印将要执行的命令
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import time
from datetime import datetime

from process_utils import child_creationflags, stop_process_gracefully

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

ROOT = os.path.dirname(os.path.abspath(__file__))
EMAILS_FILE = os.path.join(ROOT, "emails.txt")

# 导入 config 以触发 .env 加载（CLASH_SECRET 等环境变量来自 .env / 真实环境）。
try:
    import config  # noqa: F401
except Exception:
    pass

# 默认基建端点（密钥走环境变量，端点可被环境变量覆盖）。
CLASH_API_DEFAULT = os.environ.get("CLASH_API", "http://127.0.0.1:9097")
CLASH_SECRET_DEFAULT = os.environ.get("CLASH_SECRET", "")
PROXY_DEFAULT = os.environ.get("CLASH_PROXY", "http://127.0.0.1:7897")


LOG_LEVELS = {
    "DEBUG": 10,
    "INFO": 20,
    "OK": 20,
    "A": 20,
    "B": 20,
    "WARN": 30,
    "PROD": 35,
    "ERR": 40,
}
LOG_LEVEL = str(os.environ.get("OUTLOOK_LOG_LEVEL", "INFO") or "INFO").strip().upper()


def _normalize_log_level(value, default="INFO"):
    raw = str(value or default).strip().upper()
    aliases = {
        "WARNING": "WARN",
        "ERROR": "ERR",
        "PRODUCTION": "PROD",
        "SUCCESS": "OK",
    }
    raw = aliases.get(raw, raw)
    return raw if raw in LOG_LEVELS else default


def _log_level_value(value):
    return LOG_LEVELS.get(_normalize_log_level(value), LOG_LEVELS["INFO"])


def set_log_level(value):
    global LOG_LEVEL
    LOG_LEVEL = _normalize_log_level(value)
    os.environ["OUTLOOK_LOG_LEVEL"] = LOG_LEVEL
    return LOG_LEVEL


def _should_keep_prod_log(msg, level):
    if _normalize_log_level(level) == "ERR":
        return True
    low = str(msg or "").strip().lower()
    if not low:
        return False
    return low.startswith("本轮结束") or low.startswith("全部结束")


def log(msg, level="INFO"):
    rendered = _normalize_log_level(level, default="INFO")
    if LOG_LEVEL == "PROD":
        if not _should_keep_prod_log(msg, rendered):
            return
    elif _log_level_value(rendered) < _log_level_value(LOG_LEVEL):
        return
    print(f"[{datetime.now().strftime('%H:%M:%S')}] [{rendered}] {msg}", flush=True)


def _env_bool(name, default=False):
    value = os.environ.get(name)
    if value is None:
        return bool(default)
    return str(value).strip().lower() not in ("0", "false", "no", "off", "")


# ---------------------------------------------------------------- emails.txt
def read_fresh_emails():
    """返回 emails.txt 里全部 (email, password, token, client_id) 条目（含已 reserve 的，纯快照用于 diff）。
    token/client_id 由 outlook_reg_loop 注册成功后抽 Graph 写入；缺列回退空串。"""
    out = []
    if not os.path.isfile(EMAILS_FILE):
        return out
    with open(EMAILS_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            parts = line.split("----")
            email = parts[0].strip()
            password = parts[1].strip() if len(parts) > 1 else ""
            # emails.txt 新格式: [2]=client_id [3]=refresh_token(原 [2]=token [3]=client_id,已对调)
            token = parts[3].strip() if len(parts) > 3 else ""
            client_id = parts[2].strip() if len(parts) > 2 else ""
            out.append((email, password, token, client_id))
    return out


# ---------------------------------------------------------------- env
def build_child_env(args):
    env = dict(os.environ)
    env.setdefault("PYTHONUNBUFFERED", "1")
    if args.proxy:
        env["HTTP_PROXY"] = env["HTTPS_PROXY"] = args.proxy
        env["http_proxy"] = env["https_proxy"] = args.proxy
        # 关键：localhost API(BitBrowser 54345 / Clash 控制器 9097) 必须直连，
        # 否则 urllib 把它们也塞进 7897 代理 -> 502 Bad Gateway。
        no_proxy = "127.0.0.1,localhost,::1"
        env["NO_PROXY"] = env["no_proxy"] = no_proxy
    # 让 outlook_reg_loop 的 _clash_verge 能连控制器换节点
    env.setdefault("CLASH_API", args.clash_api)
    env.setdefault("CLASH_SECRET", args.clash_secret)
    env.setdefault("CLASH_GROUP", args.clash_group)
    env["OUTLOOK_LOG_LEVEL"] = args.log_level
    return env


def build_stage_a_env(env):
    stage_env = dict(env)
    for key in (
        "HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy",
        "CLASH_API", "CLASH_SECRET", "CLASH_GROUP", "CLASH_PROXY",
    ):
        stage_env.pop(key, None)
    return stage_env


# ---------------------------------------------------------------- Stage A
def stage_email(args, env):
    """拉起 outlook_reg_loop.py，盯 emails.txt，拿到一个新号就停。返回 (email, password) 或 None。"""
    before = {e for e, _p, _t, _c in read_fresh_emails()}
    log(f"Stage A 邮箱注册启动；emails.txt 现有 {len(before)} 个号", "A")

    cmd = [
        sys.executable, "outlook_reg_loop.py",
        "--engine", args.outlook_engine,
        "--count", str(args.email_attempts),
        "--timeout", str(args.email_timeout),
        "--max-press", str(args.max_press),
        "--sleep", "3",
        "--log-level", args.log_level,
    ]
    if args.outlook_proxy_file:
        cmd += ["--proxy-file", args.outlook_proxy_file]
    if args.outlook_proxy_source:
        cmd += ["--proxy-source", args.outlook_proxy_source]
    if args.outlook_proxy_url:
        cmd += ["--proxy-url", args.outlook_proxy_url]
    if args.outlook_headless:
        cmd.append("--headless")
    if args.outlook_har:
        cmd.append("--har")
    if args.email_confirm_before_register:
        cmd.append("--confirm-before-register")
    log(f"Stage A cmd: {' '.join(cmd)}", "A")
    if args.dry_run:
        return ("dry-run@outlook.com", "DryRunPass1!", "", "")

    proc = subprocess.Popen(
        cmd, cwd=ROOT, env=build_stage_a_env(env),
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        text=True, encoding="utf-8", errors="replace", bufsize=1,
        creationflags=child_creationflags(),
    )
    new_email = None
    try:
        deadline = time.time() + args.email_total_timeout
        # 边读子进程日志边透传，同时每读几行 poll 一次 emails.txt
        assert proc.stdout is not None
        last_check = 0.0
        while True:
            if proc.poll() is not None:
                log("Stage A 子进程已退出", "A")
                break
            line = proc.stdout.readline()
            if line:
                print(f"  [outlook] {line}", end="", flush=True)
            now = time.time()
            if now - last_check >= 2:
                last_check = now
                cur = read_fresh_emails()
                fresh = [t for t in cur if t[0] not in before]
                if fresh:
                    new_email = fresh[-1]
                    log(f"检测到新邮箱：{new_email[0]} —— 停止 Stage A 循环", "A")
                    break
            if now > deadline:
                log(f"Stage A 总超时 {args.email_total_timeout}s 仍无新号", "A")
                break
            if not line:
                time.sleep(0.2)
    finally:
        if proc.poll() is None:
            try:
                stop_process_gracefully(proc, timeout=15)
            except subprocess.TimeoutExpired:
                proc.kill()
    return new_email


# ---------------------------------------------------------------- Stage B
def stage_platforms(args, env, email, password, token="", client_id=""):
    log(f"Stage B 平台注册：{email}  platforms={','.join(args.platforms)}", "B")
    # token 由 Stage A 注册时抽 Graph 写入 emails.txt；有真 token 走 Graph API 直收码(免浏览器)，
    # 没有(抽取失败回退 fresh/空)则下游退化到浏览器/broker 取码。
    cmd = [
        sys.executable, "register_three_platforms.py",
        "--email", email,
        "--password", password or "",
        "--token", (token or "fresh"),
        "--platforms", *args.platforms,
        "--node", args.node,
        "--timeout", str(args.platform_timeout),
        "--broker", args.broker,
    ]
    if client_id and client_id != "fresh":
        cmd += ["--client-id", client_id]
    if args.keep_on_fail:
        cmd.append("--keep-on-fail")
    if args.import_c2a:
        cmd.append("--import-c2a")  # 透传给 register_three_platforms -> register_chatgpt
    if args.codex:
        cmd.append("--codex")  # 透传给 register_three_platforms -> register_chatgpt
        if args.codex_group:
            cmd += ["--codex-group", args.codex_group]
        if args.codex_manual_phone:
            cmd.append("--codex-manual-phone")
    log(f"Stage B cmd: {' '.join(cmd)}", "B")
    if args.dry_run:
        return 0
    proc = subprocess.Popen(cmd, cwd=ROOT, env=env)
    return proc.wait()


# ---------------------------------------------------------------- 单轮
def run_once(args, env):
    """跑一轮 A+B。返回 Stage B 的 exit code（0=成功）；没拿到邮箱返回 1。"""
    t0 = time.time()
    # Stage A
    token = client_id = ""
    if args.skip_email:
        if not args.email:
            raise SystemExit("--skip-email 需要同时给 --email")
        email, password = args.email.strip(), args.password.strip()
        log(f"跳过邮箱注册，直接用 {email}", "A")
    else:
        got = stage_email(args, env)
        if not got:
            log("Stage A 没拿到可用邮箱，本轮终止", "ERR")
            return 1, ""
        email, password, token, client_id = got
        # emails.txt 里可能没记密码，用快照里的
        password = password or args.password

    # Stage B
    print("=" * 64)
    rc = stage_platforms(args, env, email, password, token, client_id)
    print("=" * 64)
    dt = time.time() - t0
    log(f"本轮结束  email={email}  Stage B exit={rc}  用时 {dt:.0f}s",
        "OK" if rc == 0 else "WARN")
    return rc, email


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description="端到端全流程（邮箱注册 + 平台注册）")
    # Stage A
    ap.add_argument("--skip-email", action="store_true", help="跳过邮箱注册，直接用 --email")
    ap.add_argument("--email", default="", help="跳过邮箱注册时指定现成邮箱")
    ap.add_argument("--password", default="", help="配合 --email")
    ap.add_argument("--email-attempts", type=int, default=30, help="邮箱注册最多尝试次数")
    ap.add_argument("--email-timeout", type=int, default=180, help="单次邮箱注册硬超时(s)")
    ap.add_argument("--email-total-timeout", type=int, default=1800, help="Stage A 总超时(s)")
    ap.add_argument("--outlook-engine", choices=["ruoyi", "standalone"],
                    default=os.environ.get("OUTLOOK_REG_ENGINE", "ruoyi"),
                    help="Outlook 自注册后端；默认 ruoyi")
    ap.add_argument("--outlook-proxy-file", default=os.environ.get("OUTLOOK_PROXY_FILE", "proxies_outlook.txt"),
                    help="ruoyi 后端代理池文件")
    ap.add_argument("--outlook-proxy-source", default=os.environ.get("OUTLOOK_RUOYI_PROXY_SOURCE", "file"),
                    choices=["file", "http"],
                    help="Outlook ruoyi 代理来源：本地文件 / HTTP txt 列表")
    ap.add_argument("--outlook-proxy-url", default=os.environ.get("OUTLOOK_PROXY_URL", ""),
                    help="Outlook ruoyi HTTP GET 代理列表地址，返回 txt")
    ap.add_argument("--outlook-headless", action="store_true",
                    help="仅 ruoyi 后端：Outlook 注册阶段使用无头模式")
    ap.add_argument("--outlook-har", action="store_true",
                    help="仅 ruoyi 后端：保存 Outlook 注册完整链路 HAR；开启后成功/失败都会保存")
    ap.add_argument("--max-press", default="3", help="人机验证按住次数上限")
    ap.add_argument("--email-confirm-before-register", action="store_true",
                    help="邮箱注册页打开后自动点确认，再开始填写")
    # 循环
    ap.add_argument("--rounds", type=int, default=1,
                    help="循环注册轮数；1=只跑一次(默认)，0=无限循环(Ctrl+C 停)")
    ap.add_argument("--round-sleep", type=int, default=5, help="每轮之间间隔(s)")
    # Stage B
    ap.add_argument("--platforms", nargs="+", choices=["claude", "chatgpt", "grok"],
                    default=["claude"], help="默认只跑 claude（最稳）；grok 已知死结")
    ap.add_argument("--node", default="auto", help="claude/grok 走的 Clash 节点")
    ap.add_argument("--platform-timeout", type=int, default=600)
    ap.add_argument("--broker", default="", help="共享取码服务URL；默认空=各脚本自行开 Outlook 取码")
    ap.add_argument("--keep-on-fail", action="store_true")
    ap.add_argument("--import-c2a", action="store_true",
                    help="chatgpt 注册成功后即时把 token 导入 chatgpt2api（透传到底层 register_chatgpt.py）")
    ap.add_argument("--codex", action="store_true",
                    help="chatgpt 注册成功后走 Codex OAuth 提取 rt 导入 SUB2API（透传到底层 register_chatgpt.py）")
    ap.add_argument("--codex-group", default=None,
                    help="SUB2API 目标分组名（透传，默认取 config.SUB2API_GROUP）")
    ap.add_argument("--codex-manual-phone", action="store_true",
                    help="Codex add-phone 手动模式：不接码，自己在浏览器填号收码（透传）")
    # 基建
    ap.add_argument("--proxy", default=PROXY_DEFAULT, help="HTTP(S)_PROXY；传空串禁用")
    ap.add_argument("--clash-api", default=CLASH_API_DEFAULT)
    ap.add_argument("--clash-secret", default=CLASH_SECRET_DEFAULT)
    ap.add_argument("--clash-group", default="GLOBAL",
                    help="Clash 组名（proxy_switch 探节点用）；global 模式下出口由 GLOBAL 决定，"
                         "传 'auto' 会 404。claude/grok 的节点选择模式见 --node")
    ap.add_argument("--log-level", default=os.environ.get("OUTLOOK_LOG_LEVEL", "INFO"),
                    choices=["DEBUG", "INFO", "WARN", "PROD", "ERR"],
                    help="Outlook/Graph 日志等级")
    ap.add_argument("--dry-run", action="store_true", help="只打印命令不执行")
    args = ap.parse_args()

    if args.skip_email and args.rounds != 1:
        # --skip-email 每轮都用同一个固定邮箱，循环没意义（甚至会重复注册同号）
        raise SystemExit("--skip-email 只能跑单轮，不能配合 --rounds")

    set_log_level(args.log_level)
    env = build_child_env(args)
    t_all = time.time()
    print("=" * 64)
    mode = "无限" if args.rounds == 0 else f"{args.rounds} 轮"
    log(
        f"全流程开始（循环 {mode}）  outlook={args.outlook_engine}  "
        f"proxy={args.proxy or 'OFF'}  clash={args.clash_api}"
    )
    print("=" * 64)

    ok = fail = 0
    rnd = 0
    last_rc = 0
    try:
        while args.rounds == 0 or rnd < args.rounds:
            rnd += 1
            print("#" * 64)
            log(f"===== 第 {rnd} 轮{'' if args.rounds == 0 else f'/{args.rounds}'} =====")
            print("#" * 64)
            rc, _email = run_once(args, env)
            last_rc = rc
            if rc == 0:
                ok += 1
            else:
                fail += 1
            # 最后一轮不必再睡
            more = args.rounds == 0 or rnd < args.rounds
            if more and args.round_sleep > 0 and not args.dry_run:
                log(f"本轮完成，{args.round_sleep}s 后进入下一轮（Ctrl+C 可停）")
                time.sleep(args.round_sleep)
    except KeyboardInterrupt:
        log("收到 Ctrl+C，停止循环", "WARN")

    dt = time.time() - t_all
    print("=" * 64)
    log(f"全部结束  共 {rnd} 轮  成功 {ok}  失败 {fail}  总用时 {dt:.0f}s",
        "OK" if fail == 0 and ok > 0 else "WARN")
    # 退出码：全成功 0；否则沿用最后一轮的非零码（单轮场景行为不变）
    return 0 if fail == 0 and ok > 0 else (last_rc or 1)


if __name__ == "__main__":
    sys.exit(main())
