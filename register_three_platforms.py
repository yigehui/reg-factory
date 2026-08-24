# -*- coding: utf-8 -*-
"""
Run one Outlook mailbox through Claude, ChatGPT, and Grok registration flows.

Default mode is sequential because all three flows may need to read the same
mailbox. Use --parallel only when debugging isolated browser/profile behavior.

Examples:
    python register_three_platforms.py --email a@outlook.com --password xxx --token REFRESH --client-id CID
    python register_three_platforms.py --from-pool --platforms claude chatgpt grok
    python register_three_platforms.py --from-pool --parallel --keep-on-fail
"""

import argparse
import asyncio
import os
import sys
from datetime import datetime

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

from common import emails as email_pool


ROOT = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(ROOT, "tri_register_logs")


def build_command(platform, args, account):
    email, password, token, client_id = account
    timeout = str(args.timeout)

    if platform == "claude":
        cmd = [
            sys.executable, "-u", "register.py",
            "--count", "1",
            "--concurrency", "1",
            "--timeout", timeout,
            "--email", email,
            "--password", password or "",
        ]
        if token:
            cmd += ["--token", token]
        return cmd

    if platform == "chatgpt":
        cmd = [
            sys.executable, "-u", "register_chatgpt.py",
            "--count", "1",
            "--concurrency", "1",
            "--timeout", timeout,
            "--email", email,
            "--password", password or "",
        ]
        if token:
            cmd += ["--refresh-token", token]
        if client_id:
            cmd += ["--client-id", client_id]
        if args.keep_on_fail:
            cmd.append("--keep-on-fail")
        if getattr(args, "import_c2a", False):
            cmd.append("--import-c2a")  # 注册成功后即时导入 chatgpt2api
        if getattr(args, "codex", False):
            cmd.append("--codex")  # 注册成功后走 Codex OAuth 提取 rt 导入 SUB2API
            if getattr(args, "codex_group", None):
                cmd += ["--codex-group", args.codex_group]
            if getattr(args, "codex_manual_phone", False):
                cmd.append("--codex-manual-phone")
        return cmd

    if platform == "grok":
        cmd = [
            sys.executable, "-u", "register_grok.py",
            "--count", "1",
            "--concurrency", "1",
            "--timeout", timeout,
            "--email", email,
            "--password", password or "",
        ]
        if args.keep_on_fail:
            cmd.append("--keep-on-fail")
        return cmd

    raise ValueError(f"unknown platform: {platform}")


async def run_platform(platform, cmd, run_id, child_env=None):
    os.makedirs(LOG_DIR, exist_ok=True)
    log_path = os.path.join(LOG_DIR, f"{run_id}_{platform}.log")
    print(f"\n[{platform}] start")
    print(f"[{platform}] log: {log_path}")

    proc = await asyncio.create_subprocess_exec(
        *cmd,
        cwd=ROOT,
        env=child_env,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.STDOUT,
    )

    saw_success = False
    with open(log_path, "w", encoding="utf-8", errors="replace") as log:
        assert proc.stdout is not None
        while True:
            line = await proc.stdout.readline()
            if not line:
                break
            text = line.decode("utf-8", errors="replace")
            if "success: 1/1" in text.lower():
                saw_success = True
            log.write(text)
            log.flush()
            print(f"[{platform}] {text}", end="")

    rc = await proc.wait()
    ok = rc == 0 and saw_success
    status = "OK" if ok else f"FAIL(exit={rc}, success_marker={saw_success})"
    print(f"[{platform}] done: {status}")
    return platform, ok, rc, log_path


def parse_account(args):
    if args.from_pool:
        em = email_pool.next_email("tri")
        if not em:
            raise SystemExit("no email available in emails.txt")
        return em

    if not args.email:
        raise SystemExit("provide --email/--password or use --from-pool")

    return (
        args.email.strip(),
        (args.password or "").strip(),
        (args.token or "").strip(),
        (args.client_id or "").strip(),
    )


def broker_release(broker_url, email):
    """三平台都跑完后，释放该邮箱在共享取码服务里的 Outlook 会话（关浏览器窗口）。"""
    if not broker_url:
        return
    try:
        import requests
        requests.post(broker_url.rstrip("/") + "/release", json={"email": email}, timeout=30)
        print(f"  [broker] released {email}")
    except Exception as e:
        print(f"  [broker] release failed: {e}")


def child_env_for(args):
    """子进程环境：注入 MAILBOX_BROKER 让三脚本走共享取码（不再各自开 Outlook）。
    代理统一走 ruyi 代理池：把 --proxy-file/--proxy-source/--proxy-url 注入 env，
    三个 register 子进程的 argparse 默认读这些 env 构建 ConsumableProxyPool。"""
    env = dict(os.environ)
    if args.broker:
        env["MAILBOX_BROKER"] = args.broker
        env["GROK_BROKER_TIMEOUT"] = str(args.grok_timeout)
    if getattr(args, "proxy_file", None):
        env["OUTLOOK_PROXY_FILE"] = args.proxy_file
    if getattr(args, "proxy_source", None):
        env["OUTLOOK_RUOYI_PROXY_SOURCE"] = args.proxy_source
    if getattr(args, "proxy_url", None):
        env["OUTLOOK_PROXY_URL"] = args.proxy_url
    env.setdefault("PYTHONUNBUFFERED", "1")
    return env


async def process_account(account, args, child_env):
    email = account[0]
    run_id = datetime.now().strftime("%Y%m%d_%H%M%S_") + email.split("@")[0][:8]
    print("=" * 60)
    print(f"  account: {email}  platforms={','.join(args.platforms)}  mode={'parallel' if args.parallel else 'sequential'}")
    print("=" * 60)

    jobs = [(p, build_command(p, args, account)) for p in args.platforms]
    if args.parallel:
        results = await asyncio.gather(*(run_platform(p, cmd, run_id, child_env) for p, cmd in jobs))
    else:
        results = []
        for platform, cmd in jobs:
            results.append(await run_platform(platform, cmd, run_id, child_env))

    broker_release(args.broker, email)   # 释放该号 Outlook 会话
    print(f"\n  Summary [{email}]")
    for platform, ok, rc, log_path in results:
        print(f"    {platform}: {'OK' if ok else f'FAIL(exit={rc})'}  log={log_path}")
    return results


async def main():
    parser = argparse.ArgumentParser(description="Register one mailbox on three platforms (broker + loop)")
    parser.add_argument("--email", default=None)
    parser.add_argument("--password", default="")
    parser.add_argument("--token", default="", help="Outlook refresh_token")
    parser.add_argument("--client-id", default="", help="Outlook OAuth client_id")
    parser.add_argument("--from-pool", action="store_true", help="reserve one mailbox from emails.txt")
    parser.add_argument("--platforms", nargs="+", choices=["claude", "chatgpt", "grok"], default=["claude", "chatgpt", "grok"])
    parser.add_argument("--parallel", action="store_true", help="run platforms in parallel; default is sequential")
    parser.add_argument("--timeout", type=int, default=600)
    parser.add_argument("--proxy-file", default=os.environ.get("OUTLOOK_PROXY_FILE", "proxies_outlook.txt"),
                        help="代理列表文件(每行 user:pass@host:port 或 socks5://...)")
    parser.add_argument("--proxy-source", default=os.environ.get("OUTLOOK_RUOYI_PROXY_SOURCE", "file"),
                        choices=["file", "http"],
                        help="file=本地代理文件;http=HTTP GET 拉 txt 列表(配合 --proxy-url)")
    parser.add_argument("--proxy-url", default=os.environ.get("OUTLOOK_PROXY_URL", ""),
                        help="HTTP GET 代理列表地址(配合 --proxy-source=http)")
    parser.add_argument("--keep-on-fail", action="store_true")
    parser.add_argument("--import-c2a", action="store_true",
                        help="chatgpt 注册成功后即时把 token 导入 chatgpt2api（透传给 register_chatgpt.py）")
    parser.add_argument("--codex", action="store_true",
                        help="chatgpt 注册成功后走 Codex OAuth 提取 rt 导入 SUB2API（透传给 register_chatgpt.py）")
    parser.add_argument("--codex-group", default=None,
                        help="SUB2API 目标分组名（透传给 register_chatgpt.py，默认取 config.SUB2API_GROUP）")
    parser.add_argument("--codex-manual-phone", action="store_true",
                        help="Codex add-phone 手动模式（透传给 register_chatgpt.py）")
    # broker + loop
    parser.add_argument("--broker", default="http://127.0.0.1:8765", help="共享取码服务 URL；传空串 '' 禁用")
    parser.add_argument("--grok-timeout", type=int, default=40, help="Grok 取码 broker 超时(秒，outlook 注定超时故调短)")
    parser.add_argument("--loop", action="store_true", help="持续从池取号循环注册（消费侧常驻）")
    parser.add_argument("--max-inflight", type=int, default=1, help="同时在处理的邮箱数（每号峰值≈3注册窗口+1broker窗口）")
    parser.add_argument("--poll-wait", type=int, default=20, help="池空时等待产号的轮询秒数")
    args = parser.parse_args()
    child_env = child_env_for(args)

    if args.loop:
        print(f"  [loop] consumer started  max_inflight={args.max_inflight}  broker={args.broker or 'OFF'}  platforms={','.join(args.platforms)}")
        tasks = set()

        async def guarded(acc):
            try:
                await process_account(acc, args, child_env)
            except Exception as e:
                print(f"  [loop] account {acc[0]} error: {e}")

        while True:
            # 节流：处理中的邮箱达到上限就等空位，避免把池里的号一次性 reserve 光
            while len(tasks) >= args.max_inflight:
                await asyncio.sleep(2)
            acc = email_pool.next_email("tri")
            if not acc:
                print(f"  [loop] pool empty, waiting for producer... ({args.poll_wait}s)")
                await asyncio.sleep(args.poll_wait)
                continue
            t = asyncio.create_task(guarded(acc))
            tasks.add(t)
            t.add_done_callback(tasks.discard)
    else:
        account = parse_account(args)
        await process_account(account, args, child_env)


if __name__ == "__main__":
    asyncio.run(main())
