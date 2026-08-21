#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
auth_nograph_to_all.py —— 直接吃 email_nograph.txt,自动判断「已绑直授 / 未绑先绑后授」,一次拿 RT 落 email_all.txt。

背景:8 月起微软对新号 Graph 授权收紧,无辅助邮箱的号 access_denied,失败号进 email_nograph.txt。
原流程是两步:阶段A bind_secondary_email_http.py(绑→email_auth.txt)+ 阶段B auth_bound_accounts.py(已绑号重授→email_all.txt)。
本脚本合并成一步:对每个 nograph 号,单次 get_graph_token(bind_secondary=bs) 自调度——
  - 已绑号:登录后重定向链不落 proofs/Add,直奔 Consent/code,bind_secondary 闲置,拿到 RT。
  - 未绑号:落到 proofs/Add,因传了 bind_secondary,走 bind_proof_in_session 真绑(填邮箱→收码→回填)后继续 oauth 拿 RT。
成功即按完整 6 段格式写 email_all.txt(邮箱----密码----client_id----refresh_token----辅助邮箱----辅助邮箱密码)。
失败号分两类:
  - 确认废号(Abuse 风控 / 账号不存在)→ email_abuse.txt(邮箱----密码----原因----时间,同 email 去重留最新);
  - 可重试失败(限流/登录失败/绑定未过)→ email_auth_fail.txt(邮箱----密码----时间,同 email 去重只留最新)。
--skip-failed 同时跳过 fail 和 abuse 里的号,避免反复撞死号。
授权与绑定辅助邮箱全程支持代理:优先单条 --proxy,其次代理池 --proxy-url(每号按 idx 轮询一条 socks5,绕开本机 IP 撞 429),留空=直连。
只写 email_all.txt / email_auth_fail.txt / email_abuse.txt,不写 email_auth.txt(本脚本即最终闭环,无需再走阶段B)。

用法:
  python auth_nograph_to_all.py --from-nograph 1 --limit 3 --proxy ""        # 根目录 nograph 前3个,直连
  python auth_nograph_to_all.py --from-nograph 5 --proxy socks5h://user:pass@ip:port
  python auth_nograph_to_all.py --from-nograph 1 --proxy-url "http://.../proxies/text?...&return_type=socks5" -c 5
  python auth_nograph_to_all.py --email a@outlook.com --password 'Aa1!...'
  python auth_nograph_to_all.py --from-nograph 1 --nograph-file outlook_accounts/email_nograph.txt
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
from auth_bound_accounts import append_to_email_all  # noqa: E402

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stdin.reconfigure(encoding="utf-8")

CLIENT_ID = gt.CLIENT_ID
DEFAULT_NOGRAPH_FILE = os.path.join(ROOT, "email_nograph.txt")
FAIL_FILE = os.path.join(ROOT, "email_auth_fail.txt")  # 失败号:邮箱----密码----时间(同 email 去重,只留最新)
ABUSE_FILE = os.path.join(ROOT, "email_abuse.txt")  # Abuse风控/账号不存在等废号:邮箱----密码----原因----时间
GRAPH_DEBUG_DIR = gt.OUTPUT_DIR  # = "outlook_accounts",get_graph_token 存失败 debug html 的目录


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


def _load_nograph(path, limit=0, start=1):
    """读 nograph 文件,每行 `邮箱----密码`。start 从1开始。跳空行与 # 注释行。
    路径参数化(默认根目录 email_nograph.txt,可由 --nograph-file 指向 outlook_accounts/ 那份)。"""
    if not os.path.isfile(path):
        _log(f"无 {path}", "ERR"); return []
    lines = [l.strip() for l in open(path, encoding="utf-8")
             if l.strip() and not l.strip().startswith("#")]
    if start > 1:
        lines = lines[start - 1:]  # 1-based 切片
    if limit > 0:
        lines = lines[:limit]
    out = []
    for l in lines:
        parts = l.split("----")
        if len(parts) >= 2:
            out.append((parts[0].strip(), parts[1].strip()))
    return out


def _load_first_fields(path):
    """读文件第一段(邮箱)成小写集合,用于跳过判断。跳空行/#注释。"""
    s = set()
    if not os.path.isfile(path):
        return s
    for raw in open(path, encoding="utf-8"):
        raw = raw.strip()
        if not raw or raw.startswith("#"):
            continue
        s.add(raw.split("----")[0].strip().lower())
    return s


def append_to_email_fail(email, password):
    """失败号落盘 email_auth_fail.txt,格式 邮箱----密码----时间。
    同 email 去重只留最新一条(重写该行),便于重跑时一眼看出稳定失败号、避免反复处理。
    注:时间用 datetime.now(),本脚本单进程内调用安全;多进程并发时进程锁兜底整文件。"""
    line = f"{email}----{password}----{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    with _interprocess_lock(FAIL_FILE):
        existing = []
        if os.path.isfile(FAIL_FILE):
            with open(FAIL_FILE, encoding="utf-8") as f:
                for raw in f:
                    raw = raw.rstrip("\n")
                    if not raw.strip() or raw.startswith("#"):
                        continue
                    if raw.split("----")[0].strip().lower() != email.lower():
                        existing.append(raw)
        existing.append(line)
        with open(FAIL_FILE, "w", encoding="utf-8") as f:
            for r in existing:
                f.write(r + "\n")
    _log(f"email_auth_fail.txt ~= {email} (status=fail)", "WARN")
    return True


def append_to_email_abuse(email, password, reason):
    """废号(Abuse风控/账号不存在)落 email_abuse.txt,格式 邮箱----密码----原因----时间。同 email 去重留最新。"""
    line = f"{email}----{password}----{reason}----{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    with _interprocess_lock(ABUSE_FILE):
        existing = []
        if os.path.isfile(ABUSE_FILE):
            with open(ABUSE_FILE, encoding="utf-8") as f:
                for raw in f:
                    raw = raw.rstrip("\n")
                    if not raw.strip() or raw.startswith("#"):
                        continue
                    if raw.split("----")[0].strip().lower() != email.lower():
                        existing.append(raw)
        existing.append(line)
        with open(ABUSE_FILE, "w", encoding="utf-8") as f:
            for r in existing:
                f.write(r + "\n")
    _log(f"email_abuse.txt ~= {email} (reason={reason})", "WARN")
    return True


def _classify_failure(email, idx):
    """get_graph_token 返回 None 后,判失败原因。
    返回: 'abuse' / 'noexist' / 'rate_limited' / 'pwd_incorrect' / 'unknown'。
    优先读线程内内存原因(gt.get_thread_fail_reason);内存没有才兜底扫最新一份 graph_debug_*.html。
    debug 文件名格式(见 extract_graph_tokens._save_graph_debug_html):
      graph_debug_{idx}_{reason}_{email_safe}_{ts}.html , email_safe = email.replace(@,_, /,_)

    判定依据来自微软 Abuse 中断页的 serverdata 标记(实测,2026-08):
      fchallengeabusiveaccount / account_serviceabuseinterruptpage / iabusereason
    (不是 URL 里的 /abuse 路径——那串在登录页 js 里也有,会误判。)
    """
    import glob
    # 1) 内存原因优先(get_graph_token 各失败分支记的线程内原因):不依赖磁盘,
    #    网页保存开关(--save-debug / OUTLOOK_SAVE_DEBUG_HTML)关着也能分类。
    reason = gt.get_thread_fail_reason()
    if reason:
        if reason in ("abuse", "noexist", "rate_limited", "pwd_incorrect"):
            return reason
        return "unknown"  # no_flow_token/consent_update/proofs_add/bind_fail/token_error/exc 等一律可重试
    # 2) 兜底:扫历史保存的 debug html(仅 --save-debug 开着且文件存在时有效),取最新一份
    safe = str(email).replace("@", "_").replace("/", "_")
    # 匹配该号任意 idx/reason 的 debug html,取最新一份
    fs = glob.glob(os.path.join(GRAPH_DEBUG_DIR, f"graph_debug_*_*_{safe}_*.html"))
    if not fs:
        return "unknown"
    fs.sort(key=os.path.getmtime, reverse=True)
    try:
        h = open(fs[0], encoding="utf-8", errors="ignore").read().lower()
    except Exception:
        return "unknown"
    # Abuse 中断页:微软要风控验证/挑战,号被标记为 abusive account
    if ("fchallengeabusiveaccount" in h or "account_serviceabuseinterruptpage" in h
            or "iabusereason" in h or "frequiresfluentserviceabuseflow" in h):
        return "abuse"
    if "doesn't exist" in h or "does not exist" in h or "that account doesn" in h:
        return "noexist"
    if "too many requests" in h or "status=429" in h:
        return "rate_limited"
    # 密码错误(80041012):号本身存在但凭据不对,常因出口 IP 声誉差被微软挡;可重试(换 IP)。
    if "password is incorrect" in h or "80041012" in h:
        return "pwd_incorrect"
    return "unknown"


def _fetch_proxy_pool(proxy_url):
    """从代理池 HTTP API 拉一条 socks5 列表。复用 register_outlook_ruoyi.fetch_proxy_list_http。"""
    from register_outlook_ruoyi import fetch_proxy_list_http
    pool = fetch_proxy_list_http(proxy_url)
    # 归一化成 requests 可用的 socks5h:// 形式(_proxy_to_requests 再兜一层)
    return [str(p).strip() for p in pool if str(p).strip()]


def _pick_proxy(pool, idx):
    """从代理池按 idx 轮询取一条;池空返回 ''(直连)。"""
    if not pool:
        return ""
    return pool[idx % len(pool)]


def process_one(email, password, idx, proxy_str, proxy_pool=None):
    """单 nograph 号:单次 get_graph_token(bind_secondary) 自调度「已绑直授 / 未绑先绑后授」。
    返回 dict(含 ok/status/secondary)。"""
    tag = f"[#{idx}]"
    # 代理优先级: 单条 --proxy > 代理池 --proxy-url(按idx轮询) > 直连
    use_proxy = proxy_str
    if not use_proxy and proxy_pool:
        use_proxy = _pick_proxy(proxy_pool, idx)
    proxy_tag = gt._mask_proxy_url(use_proxy) if use_proxy else "direct"
    _log(f"{tag} 处理 {email} proxy={proxy_tag}")
    # 1) cf 代理(收码要经代理访问 worker);建/取 cf 辅助邮箱
    try:
        cm.set_proxy(use_proxy) if use_proxy else cm.set_proxy(os.environ.get("CF_MAIL_PROXY") or None)
        d = cm.create_or_get_address(email)
        cf_address = d["address"]; cf_jwt = d["jwt"]
        cf_use_admin = d.get("use_admin", False)
        cf_password = d.get("password")
        if cf_password:
            _log(f"{tag} cf 邮箱密码已生成: {cf_address} / {cf_password}")
        _log(f"{tag} cf 辅助邮箱就绪: {cf_address} (use_admin={cf_use_admin})")
    except Exception as e:
        _log(f"{tag} 建 cf 辅助邮箱失败: {type(e).__name__}: {e}", "ERR")
        append_to_email_fail(email, password)
        return {"ok": False, "email": email, "status": "cf_fail", "error": str(e)}

    # 2) 构造 bind_secondary(用 create_or_get_address 的返回原样,不硬编码 use_admin)
    bs = {
        "cf_address": cf_address,
        "cf_jwt": cf_jwt,
        "use_admin": cf_use_admin,
        "cm": cm,
    }
    # 3) 单次 get_graph_token:已绑号不进 proofs/Add 直拿 code;未绑号进 proofs/Add 真绑后拿 code
    proxies = _proxy_to_requests(use_proxy)
    try:
        result = gt.get_graph_token(email, password, idx=idx, proxies=proxies, bind_secondary=bs)
    except Exception as e:
        _log(f"{tag} 协议流程异常: {type(e).__name__}: {e}", "ERR")
        append_to_email_fail(email, password)
        return {"ok": False, "email": email, "status": "exc", "error": str(e), "secondary": cf_address}

    # 4) 拿到 RT 即闭环成功,落 email_all.txt(6 段)
    if result and result.get("refresh_token"):
        rt = result["refresh_token"]
        _log(f"{tag} 授权成功,拿到 RT: {rt[:30]}...", "OK")
        append_to_email_all(email, password, CLIENT_ID, rt, cf_address, cf_password)
        return {"ok": True, "email": email, "status": "ok", "secondary": cf_address,
                "secondary_password": cf_password, "refresh_token": rt, "client_id": CLIENT_ID}
    else:
        cls = _classify_failure(email, idx)
        if cls == "abuse":
            _log(f"{tag} 未拿到 RT: 账号被微软风控(Abuse)", "WARN")
            append_to_email_abuse(email, password, "abuse")
        elif cls == "noexist":
            _log(f"{tag} 未拿到 RT: 账号不存在(已注销)", "WARN")
            append_to_email_abuse(email, password, "noexist")
        else:
            # rate_limited / unknown / other:限流或暂时失败,落 fail 待重试(不冤枉进abuse)
            _log(f"{tag} 未拿到 RT(限流/登录失败/绑定未过,cls={cls})", "WARN")
            append_to_email_fail(email, password)
        return {"ok": False, "email": email, "status": cls, "secondary": cf_address}


def main():
    ap = argparse.ArgumentParser(
        description="直接吃 email_nograph.txt:自动判断已绑直授/未绑先绑后授,落 email_all.txt")
    ap.add_argument("--email", default="")
    ap.add_argument("--password", default="")
    ap.add_argument("--from-nograph", type=int, default=0,
                    help="读 nograph 文件;0=不读;>0 从第N个开始;传负数=从第1个读全部")
    ap.add_argument("--nograph-file", default=DEFAULT_NOGRAPH_FILE,
                    help=f"nograph 文件路径(默认 {DEFAULT_NOGRAPH_FILE};可指 outlook_accounts/email_nograph.txt)")
    ap.add_argument("--limit", type=int, default=0, help="最多处理几个号(0=不限)")
    ap.add_argument("--proxy", default=os.environ.get("OUTLOOK_AUTH_PROXY", ""),
                    help="单条代理(socks5h://...),留空看 --proxy-url 或直连")
    ap.add_argument("--proxy-url", default=os.environ.get("OUTLOOK_PROXY_URL", ""),
                    help="代理池 HTTP API(GET 返回 socks5 列表,参考 ruyi 注册);留空=不用代理池")
    ap.add_argument("--save-debug", action="store_true",
                    help="授权失败时保存网页到 outlook_accounts/(默认不保存;也可 env OUTLOOK_SAVE_DEBUG_HTML=1)")
    ap.add_argument("--concurrency", "-c", type=int, default=1)
    ap.add_argument("--skip-failed", action="store_true",
                    help="跳过 email_auth_fail.txt 里已记录失败的号(避免反复撞风控号)")
    ap.add_argument("--skip-done", action="store_true",
                    help="跳过 email_all.txt 里已成功的号(重跑时不重复授权)")
    ap.add_argument("--log-level", default=os.environ.get("OUTLOOK_LOG_LEVEL", "INFO"),
                    choices=["DEBUG", "INFO", "WARN", "PROD", "ERR"])
    args = ap.parse_args()
    os.environ["OUTLOOK_LOG_LEVEL"] = gt._normalize_log_level(args.log_level)
    if args.save_debug:
        gt.set_save_debug_html(True)
        _log("已开启失败网页保存(--save-debug):失败将落 outlook_accounts/graph_debug_*.html")

    # 收集账号
    accounts = []
    if args.email and args.password:
        accounts.append((args.email, args.password))
    elif args.from_nograph != 0:
        start = abs(args.from_nograph)
        accounts = _load_nograph(args.nograph_file, limit=args.limit, start=start)
    if not accounts:
        _log("无账号可处理(用 --email/--password 或 --from-nograph N)", "ERR"); sys.exit(1)

    # 跳过集合:已失败 / 已成功(按需)
    # --skip-failed 同时把 email_abuse.txt(Abuse风控/账号不存在=确认废号)一并跳过,避免反复撞死号。
    skip_fail = set()
    if args.skip_failed:
        skip_fail = _load_first_fields(FAIL_FILE)
        skip_fail |= _load_first_fields(ABUSE_FILE)
        _log(f"--skip-failed: 跳过 {len(skip_fail)} 个已失败/废号(fail+abuse)")
    skip_done = set()
    if args.skip_done:
        skip_done = _load_first_fields(os.path.join(ROOT, "email_all.txt"))
        _log(f"--skip-done: 跳过 {len(skip_done)} 个已成功号")
    skipset = skip_fail | skip_done
    if skipset:
        before = len(accounts)
        accounts = [(e, p) for e, p in accounts if e.lower() not in skipset]
        _log(f"过滤后待处理: {len(accounts)}(原 {before},跳过 {before - len(accounts)})")
    if not accounts:
        _log("过滤后无账号可处理", "WARN"); sys.exit(0)

    # 加载代理池(--proxy-url):每号按 idx 轮询取一条,绕开本机 IP 撞 429。
    proxy_pool = []
    if args.proxy_url:
        _log(f"拉取代理池: {gt._mask_proxy_url(args.proxy_url)} ...")
        try:
            proxy_pool = _fetch_proxy_pool(args.proxy_url)
            _log(f"代理池就绪: {len(proxy_pool)} 条")
        except Exception as e:
            _log(f"拉代理池失败: {type(e).__name__}: {e}(将退回直连/单条 --proxy)", "WARN")
            proxy_pool = []

    src = gt._mask_proxy_url(args.proxy) if args.proxy else (f"pool({len(proxy_pool)})" if proxy_pool else "direct")
    _log(f"账号数={len(accounts)} 并发={args.concurrency} proxy={src}")
    _log(f"nograph_file={args.nograph_file} client_id={CLIENT_ID} 输出=email_all.txt")

    results = []
    if args.concurrency <= 1 or len(accounts) == 1:
        for i, (e, p) in enumerate(accounts, 1):
            results.append(process_one(e, p, i, args.proxy, proxy_pool))
    else:
        with ThreadPoolExecutor(max_workers=args.concurrency) as pool:
            fut = {pool.submit(process_one, e, p, i, args.proxy, proxy_pool): (e, p)
                   for i, (e, p) in enumerate(accounts, 1)}
            for f in as_completed(fut):
                results.append(f.result())

    ok = [r for r in results if r.get("ok")]
    _log("=" * 50)
    _log(f"完成: {len(ok)}/{len(results)} 成功")
    for r in ok:
        _log(f"  [OK] {r['email']} -> secondary={r.get('secondary')} rt={r.get('refresh_token','')[:30]}...", "OK")
    fails = [r for r in results if not r.get("ok")]
    if fails:
        _log(f"失败 {len(fails)}:", "WARN")
        for r in fails:
            _log(f"  [FAIL] {r['email']} status={r.get('status')} err={r.get('error','')[:80]}")


if __name__ == "__main__":
    main()
