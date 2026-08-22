#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
auth_bound_accounts.py —— 阶段B:读 email_auth.txt,对已绑辅助邮箱的号走纯协议 Graph OAuth 拿 RT。

前置:阶段A(bind_secondary_email_http.py)已把这些号绑了 cf 辅助邮箱并写入 email_auth.txt。
本阶段:复用 get_graph_token(bind_secondary=...) —— 登录后 proofs/Add 真绑(已是绑定态,会直接
过或复绑确认)→ Consent → code → authorization_code 换 refresh_token。

注意:已绑辅助邮箱的号,登录可能触发「Verify your email」挑战(微软要确认辅助邮箱地址)。
当前 get_graph_token 的纯协议路径暂不处理该登录挑战(那是浏览器 probe 的 _handle_login_verify_email)。
若登录时未触发挑战(账号状态正常),则正常走 proofs/Add→Consent→code 拿到 RT。

输出: email_all.txt  每行 `邮箱----密码----client_id----refresh_token----辅助邮箱`

用法:
  python auth_bound_accounts.py                         # 处理 email_auth.txt 全部
  python auth_bound_accounts.py --limit 5                # 只处理前5个
  python auth_bound_accounts.py --proxy socks5h://...    # 指定代理(默认直连)
"""
from __future__ import annotations

import argparse
import os
import sys
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
AUTH_FILE = os.path.join(ROOT, "email_auth.txt")
ALL_FILE = os.path.join(ROOT, "email_all.txt")


def _log(msg, level="INFO"):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] [{level}] {msg}", flush=True)


def _interprocess_lock(target_path):
    import threading
    lock = _interprocess_lock._locks.setdefault(target_path, threading.Lock())
    return lock


_interprocess_lock._locks = {}


def _proxy_to_requests(proxy_str):
    if not proxy_str:
        return None
    p = proxy_str.strip()
    if p.startswith("socks5://"):
        p = "socks5h://" + p[len("socks5://"):]
    return {"http": p, "https": p}


def append_to_email_all(email, password, client_id, refresh_token, secondary_email, secondary_password=None):
    """最终闭环格式: 邮箱----密码----client_id----refresh_token----辅助邮箱----辅助邮箱密码"""
    existing = set()
    with _interprocess_lock(ALL_FILE):
        if os.path.isfile(ALL_FILE):
            with open(ALL_FILE, encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#"):
                        existing.add(line.split("----")[0].strip().lower())
        if email.lower() in existing:
            _log(f"email_all.txt 已有 {email},跳过")
            return True
        with open(ALL_FILE, "a", encoding="utf-8") as f:
            f.write(f"{email}----{password}----{client_id}----{refresh_token}----{secondary_email}----{secondary_password or ''}\n")
    _log(f"email_all.txt += {email} (rt={refresh_token[:24]}...)", "OK")
    return True


def auth_one(email, password, client_id, secondary_email, secondary_password=None, idx=1, proxy_str=""):
    """对已绑辅助邮箱的号走协议 Graph OAuth 拿 RT。
    email_auth.txt 里 client_id 可能是默认 Thunderbird;若行里带了就用行里的。"""
    tag = f"[#{idx}]"
    _log(f"{tag} 授权 {email} (secondary={secondary_email})")
    # cf 模块设代理(收码场景;此处可能不需收码,但若登录触发挑战会用到)
    try:
        cm.set_proxy(proxy_str) if proxy_str else cm.set_proxy(os.environ.get("CF_MAIL_PROXY") or None)
    except Exception:
        pass

    proxies = _proxy_to_requests(proxy_str)
    # 传 bind_secondary:已绑号登录后若再到 proofs/Add 会复绑确认(幂等),拿到 cf 配置以防需要收码
    bs = {
        "cf_address": secondary_email,
        "cf_jwt": None,          # 已存在邮箱 use_admin 收码,无需 JWT
        "use_admin": True,       # 已存在邮箱走 admin API
        "cm": cm,
    }
    try:
        result = gt.get_graph_token(email, password, idx=idx, proxies=proxies, bind_secondary=bs)
    except Exception as e:
        _log(f"{tag} 协议异常: {type(e).__name__}: {e}", "ERR")
        return {"ok": False, "email": email, "status": "exc", "error": str(e)}

    if result and result.get("refresh_token"):
        rt = result["refresh_token"]
        _log(f"{tag} 拿到 RT: {rt[:30]}...", "OK")
        append_to_email_all(email, password, client_id or CLIENT_ID, rt, secondary_email, secondary_password)
        return {"ok": True, "email": email, "refresh_token": rt, "client_id": client_id or CLIENT_ID}
    else:
        _log(f"{tag} 未拿到 RT(可能登录触发挑战/或 access_denied)", "WARN")
        return {"ok": False, "email": email, "status": "no_rt"}


def _load_auth_accounts(limit=0):
    """读 email_auth.txt: 邮箱----密码----client_id----辅助邮箱"""
    if not os.path.isfile(AUTH_FILE):
        _log(f"无 {AUTH_FILE}(先跑阶段A绑定)", "ERR"); return []
    out = []
    for line in open(AUTH_FILE, encoding="utf-8"):
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        parts = line.split("----")
        if len(parts) >= 4:
            email, pwd, cid, sec = parts[0], parts[1], parts[2], parts[3]
            # 第5段:辅助邮箱密码(新格式);老行没有则空
            sec_pwd = parts[4].strip() if len(parts) >= 5 else ""
            out.append((email.strip(), pwd.strip(), cid.strip(), sec.strip(), sec_pwd))
    if limit > 0:
        out = out[:limit]
    return out


def main():
    ap = argparse.ArgumentParser(description="阶段B:已绑辅助邮箱号走协议 Graph 授权拿RT")
    ap.add_argument("--from-auth", action="store_true", help="读 email_auth.txt(默认即此)")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--proxy", default=os.environ.get("OUTLOOK_AUTH_PROXY", ""),
                    help="MS 登录端点代理(socks5h://...),留空=直连")
    ap.add_argument("--concurrency", "-c", type=int, default=1)
    ap.add_argument("--log-level", default=os.environ.get("OUTLOOK_LOG_LEVEL", "INFO"),
                    choices=["DEBUG", "INFO", "WARN", "PROD", "ERR"])
    args = ap.parse_args()
    os.environ["OUTLOOK_LOG_LEVEL"] = gt._normalize_log_level(args.log_level)

    accounts = _load_auth_accounts(limit=args.limit)
    if not accounts:
        _log("无账号(先跑 bind_secondary_email_http.py 生成 email_auth.txt)", "ERR"); sys.exit(1)

    _log(f"账号数={len(accounts)} 并发={args.concurrency} proxy={gt._mask_proxy_url(args.proxy)}")
    _log(f"client_id={CLIENT_ID} 输出={ALL_FILE}")

    results = []
    if args.concurrency <= 1 or len(accounts) == 1:
        for i, (e, p, c, s, sp) in enumerate(accounts, 1):
            results.append(auth_one(e, p, c, s, sp, i, args.proxy))
    else:
        with ThreadPoolExecutor(max_workers=args.concurrency) as pool:
            fut = {pool.submit(auth_one, e, p, c, s, sp, i, args.proxy): e
                   for i, (e, p, c, s, sp) in enumerate(accounts, 1)}
            for f in as_completed(fut):
                results.append(f.result())

    ok = [r for r in results if r.get("ok")]
    _log("=" * 50)
    _log(f"授权完成: {len(ok)}/{len(results)} 成功")
    for r in ok:
        _log(f"  [OK] {r['email']} rt={r.get('refresh_token','')[:30]}...", "OK")
    fails = [r for r in results if not r.get("ok")]
    if fails:
        _log(f"失败 {len(fails)}:", "WARN")
        for r in fails:
            _log(f"  [FAIL] {r['email']} status={r.get('status')}")


if __name__ == "__main__":
    main()
