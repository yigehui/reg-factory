#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
bind_secondary_email_http.py —— 阶段A:纯 HTTP(协议)给 MS 账号绑 cf 辅助邮箱。

背景:8/6 起微软对新号 Graph 授权收紧,无辅助邮箱的号一律 access_denied(no_graph)。
绑了辅助邮箱的号才能过 Graph OAuth(已用 probe 浏览器版在 joseph_lee239 验证)。
本脚本把绑定流程从浏览器改成纯协议,绑定成功的写 email_auth.txt。

流程(复用 extract_graph_tokens 的登录 + bind_proof_in_session 的绑定):
  GET authorize → parse sFT/sCtx/urlPost → POST ppsecure 登录 → 跟跳板 →
  proofs/Add(AddProof+EmailAddress) → 微软发码到 cf → cf 收码 → proofs/Verify(VerifyProof+iOttText)
  → Consent/Update(已是登录态,绑定即成功)→ 拿 auth code 作成功判据。

输出: email_auth.txt  每行 `邮箱----密码----client_id----辅助邮箱`

用法:
  python bind_secondary_email_http.py --from-nograph 3            # email_nograph.txt 第3个号
  python bind_secondary_email_http.py --email a@outlook.com --password 'Aa1!...'
  python bind_secondary_email_http.py --from-nograph 0 --proxy socks5h://user:pass@ip:port
"""
from __future__ import annotations

import argparse
import os
import sys
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))


def _load_dotenv():
    try:
        path = os.path.join(ROOT, ".env")
        if not os.path.isfile(path):
            return
        for line in open(path, encoding="utf-8"):
            s = line.strip()
            if not s or s.startswith("#") or "=" not in s:
                continue
            k, _, v = s.partition("=")
            k = k.strip(); v = v.strip().strip('"').strip("'")
            if k and k not in os.environ:
                os.environ[k] = v
    except Exception:
        pass


_load_dotenv()

import extract_graph_tokens as gt  # noqa: E402
from common import cloudflare_mail as cm  # noqa: E402

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stdin.reconfigure(encoding="utf-8")

CLIENT_ID = gt.CLIENT_ID
OUTPUT_FILE = os.path.join(ROOT, "email_auth.txt")
NOGRAPH_FILE = os.path.join(ROOT, "email_nograph.txt")


def _log(msg, level="INFO"):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] [{level}] {msg}", flush=True)


def _interprocess_lock(target_path):
    """多进程文件锁(复用 register_outlook_ruoyi 同款实现,避免循环 import 就地重写)。"""
    import threading
    lock = _interprocess_lock._locks.setdefault(target_path, threading.Lock())
    return lock


_interprocess_lock._locks = {}


def _proxy_to_requests(proxy_str):
    """socks5://... 或 socks5h://... 或 http://... → requests proxies dict。None=直连。"""
    if not proxy_str:
        return None
    p = proxy_str.strip()
    if p.startswith("socks5://"):
        p = "socks5h://" + p[len("socks5://"):]  # 远程 DNS
    return {"http": p, "https": p}


def append_to_email_auth(email, password, client_id, secondary_email, secondary_password=None):
    """绑定成功落盘,按 email 去重。
    格式: 邮箱----密码----client_id----辅助邮箱----辅助邮箱密码(密码功能未开则末段空)"""
    existing = set()
    with _interprocess_lock(OUTPUT_FILE):
        if os.path.isfile(OUTPUT_FILE):
            with open(OUTPUT_FILE, encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#"):
                        existing.add(line.split("----")[0].strip().lower())
        if email.lower() in existing:
            _log(f"email_auth.txt 已有 {email},跳过写入")
            return True
        with open(OUTPUT_FILE, "a", encoding="utf-8") as f:
            f.write(f"{email}----{password}----{client_id}----{secondary_email}----{secondary_password or ''}\n")
    _log(f"email_auth.txt += {email} (secondary={secondary_email})", "OK")
    return True


def bind_one(email, password, idx, proxy_str):
    """单账号协议绑定。返回 dict(含 ok/status/secondary)。"""
    tag = f"[#{idx}]"
    _log(f"{tag} 开始绑定 {email}")
    # 1) 建/取 cf 辅助邮箱(收码要靠它)。cf worker 走代理访问更稳(本地 DNS 常不稳)。
    try:
        cm.set_proxy(proxy_str) if proxy_str else cm.set_proxy(os.environ.get("CF_MAIL_PROXY") or None)
        d = cm.create_or_get_address(email)
        cf_address = d["address"]; cf_jwt = d["jwt"]
        cf_use_admin = d.get("use_admin", False)
        cf_password = d.get("password")  # 地址密码(ENABLE_ADDRESS_PASSWORD 开启才有;None=未开启)
        if cf_password:
            _log(f"{tag} cf 邮箱密码已生成: {cf_address} / {cf_password}")
        _log(f"{tag} cf 辅助邮箱就绪: {cf_address} (use_admin={cf_use_admin})")
    except Exception as e:
        _log(f"{tag} 建 cf 辅助邮箱失败: {type(e).__name__}: {e}", "ERR")
        return {"ok": False, "email": email, "status": "cf_fail", "error": str(e)}

    # 2) 调 get_graph_token,传 bind_secondary 让 proofs/Add 真绑而非 Skip。
    #    绑定成功后 get_graph_token 会继续跟 oauth 拿 code 并换 RT——RT 顺带就拿到了,
    #    但本阶段只关心绑定成功(email_auth.txt),RT 留给阶段B 重跑(或此处也保存)。
    proxies = _proxy_to_requests(proxy_str)
    bs = {
        "cf_address": cf_address,
        "cf_jwt": cf_jwt,
        "use_admin": cf_use_admin,
        "cm": cm,
    }
    try:
        result = gt.get_graph_token(email, password, idx=idx, proxies=proxies, bind_secondary=bs)
    except Exception as e:
        _log(f"{tag} 协议流程异常: {type(e).__name__}: {e}", "ERR")
        return {"ok": False, "email": email, "status": "exc", "error": str(e), "secondary": cf_address}

    if result and result.get("refresh_token"):
        rt = result["refresh_token"]
        _log(f"{tag} 绑定成功且已拿到 RT(顺带): rt={rt[:30]}...", "OK")
        append_to_email_auth(email, password, CLIENT_ID, cf_address, cf_password)
        return {"ok": True, "email": email, "status": "ok", "secondary": cf_address,
                "secondary_password": cf_password,
                "refresh_token": rt, "client_id": CLIENT_ID}
    else:
        _log(f"{tag} 流程未拿到 RT(可能绑定后微软要求别的校验/或登录失败)", "WARN")
        return {"ok": False, "email": email, "status": "no_rt", "secondary": cf_address}


def _load_accounts_from_nograph(limit=0, start=1):
    """读 email_nograph.txt,每行 `邮箱----密码`。start 从1开始。"""
    if not os.path.isfile(NOGRAPH_FILE):
        _log(f"无 {NOGRAPH_FILE}", "ERR"); return []
    lines = [l.strip() for l in open(NOGRAPH_FILE, encoding="utf-8") if l.strip()]
    if start > 1:
        lines = lines[start - 1:]
    if limit > 0:
        lines = lines[:limit]
    out = []
    for l in lines:
        parts = l.split("----")
        if len(parts) >= 2:
            out.append((parts[0].strip(), parts[1].strip()))
    return out


def main():
    ap = argparse.ArgumentParser(description="阶段A:协议绑定辅助邮箱")
    ap.add_argument("--email", default="")
    ap.add_argument("--password", default="")
    ap.add_argument("--from-nograph", type=int, default=0,
                    help="读 email_nograph.txt;0=不读;>0 从第N个开始;传负数=从第1个读全部")
    ap.add_argument("--limit", type=int, default=0, help="最多处理几个号(0=不限)")
    ap.add_argument("--proxy", default=os.environ.get("OUTLOOK_BIND_PROXY", ""),
                    help="MS 登录端点代理(socks5h://...),留空=直连(默认)")
    ap.add_argument("--concurrency", "-c", type=int, default=1)
    ap.add_argument("--log-level", default=os.environ.get("OUTLOOK_LOG_LEVEL", "INFO"),
                    choices=["DEBUG", "INFO", "WARN", "PROD", "ERR"])
    args = ap.parse_args()
    os.environ["OUTLOOK_LOG_LEVEL"] = gt._normalize_log_level(args.log_level)

    # 收集账号
    accounts = []
    if args.email and args.password:
        accounts.append((args.email, args.password))
    elif args.from_nograph != 0:
        start = abs(args.from_nograph)
        accounts = _load_accounts_from_nograph(limit=args.limit, start=start)
    if not accounts:
        _log("无账号可处理(用 --email/--password 或 --from-nograph N)", "ERR"); sys.exit(1)

    _log(f"账号数={len(accounts)} 并发={args.concurrency} proxy={gt._mask_proxy_url(args.proxy)}")
    _log(f"client_id={CLIENT_ID} 输出={OUTPUT_FILE}")

    results = []
    if args.concurrency <= 1 or len(accounts) == 1:
        for i, (e, p) in enumerate(accounts, 1):
            results.append(bind_one(e, p, i, args.proxy))
    else:
        with ThreadPoolExecutor(max_workers=args.concurrency) as pool:
            fut = {pool.submit(bind_one, e, p, i, args.proxy): (e, p) for i, (e, p) in enumerate(accounts, 1)}
            for f in as_completed(fut):
                results.append(f.result())

    ok = [r for r in results if r.get("ok")]
    _log("=" * 50)
    _log(f"绑定完成: {len(ok)}/{len(results)} 成功")
    for r in ok:
        _log(f"  [OK] {r['email']} -> {r.get('secondary')}", "OK")
    fails = [r for r in results if not r.get("ok")]
    if fails:
        _log(f"失败 {len(fails)}:", "WARN")
        for r in fails:
            _log(f"  [FAIL] {r['email']} status={r.get('status')} err={r.get('error','')[:80]}")


if __name__ == "__main__":
    main()
