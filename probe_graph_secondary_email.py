#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
probe_graph_secondary_email.py —— 探测微软 Graph 授权流程里「绑定辅助邮箱」这一步的真实形态。

目标(只读观测,不真绑):
  1. 用代理池一条 socks5 起 ruyipage Firefox
  2. 登录微软账号(email_nograph 失败账号)
  3. 打开 Graph 授权 URL(login.microsoftonline.com/consumers/oauth2/...)
  4. 一路跟到「辅助邮箱绑定 / proofs/Add / 安全信息」页
  5. 把每一步的 URL / title / HTML / 截图 / HAR 全 dump 到 outlook_accounts/probe_*/
     供后续分析「辅助邮箱提交」的请求序列(HTTP 重放用)

产物目录: outlook_accounts/probe_<email>_<ts>/
  - 01_landed_<step>.png        每步截图
  - 01_step_<n>.html            每步页面 HTML
  - steps.json                  每步 URL/title 摘要
  - network.har                 全程网络包(类 HAR)
  - console.log                 浏览器 console

用法:
  python probe_graph_secondary_email.py --email linda_johnson072@outlook.com --password 'Aa1!...'
  python probe_graph_secondary_email.py --from-nograph 2      # 用 email_nograph.txt 第2个账号
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import threading
import time
import urllib.parse
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

import register_outlook_ruoyi as rr  # noqa: E402
from ruyipage import FirefoxOptions, FirefoxPage  # noqa: E402

# Graph 授权常量(与 extract_graph_tokens.py 一致)
CLIENT_ID = "9e5f94bc-e8a4-4e73-b8be-63364c29d753"
REDIRECT_URI = "http://localhost"
SCOPE = "offline_access https://graph.microsoft.com/Mail.Read"


def _log(msg, level="INFO"):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] [{level}] {msg}", flush=True)


def _build_auth_url():
    return (
        f"https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize"
        f"?client_id={CLIENT_ID}"
        f"&response_type=code"
        f"&redirect_uri={urllib.parse.quote(REDIRECT_URI, safe='')}"
        f"&scope={urllib.parse.quote(SCOPE)}"
        f"&response_mode=query"
    )


class ProbeDumper:
    """每步把 URL/title/HTML/截图/网络包存盘。"""

    def __init__(self, out_dir, page):
        self.dir = out_dir
        os.makedirs(out_dir, exist_ok=True)
        self.page = page
        self.steps = []
        self.n = 0
        # 网络监听
        self.entries = []
        self._stop = threading.Event()
        self.listen_started = False
        # 绑定模式状态
        self.bind_mode = False
        self.cf_address = None
        self.cf_jwt = None
        self.cf_use_admin = False
        self.last_auth_code = None

    def start_listen(self):
        try:
            listen = getattr(self.page, "listen", None)
            if listen is None:
                _log("page.listen 不可用,跳过 HAR", "WARN")
                return
            listen.start(targets=True, collect_response=True)
            self.listen_started = True

            def _worker():
                while not self._stop.is_set():
                    try:
                        pkt = listen.wait(timeout=0.5)
                    except Exception:
                        continue
                    if pkt is not None:
                        self.entries.append(self._pkt(pkt))

            t = threading.Thread(target=_worker, name="probe-listen", daemon=True)
            t.start()
            _log("网络监听已启动")
        except Exception as e:
            _log(f"网络监听启动失败(忽略): {type(e).__name__}: {e}", "WARN")

    def _pkt(self, pkt):
        try:
            d = {
                "url": getattr(pkt, "url", "") or "",
                "method": getattr(pkt, "method", "") or "",
                "status": getattr(pkt, "status", 0),
                "event_type": getattr(pkt, "event_type", "") or "",
                "timestamp": getattr(pkt, "timestamp", "") or "",
            }
            # 尝试取请求头/响应头(可能不存在)
            for attr in ("headers", "request", "response"):
                try:
                    val = getattr(pkt, attr, None)
                    if val is not None:
                        d[attr] = str(val)[:2000]
                except Exception:
                    pass
            # 响应体(已 collect_response=True)
            try:
                body = getattr(pkt, "response_body", None)
                if body is None:
                    body = pkt.get_response_body() if hasattr(pkt, "get_response_body") else None
                if body:
                    d["response_body"] = str(body)[:4000]
            except Exception:
                pass
            return d
        except Exception as e:
            return {"_error": f"{type(e).__name__}: {e}", "url": str(pkt)[:200]}

    def snapshot(self, label, note=""):
        """记录当前页面状态:URL/title/html/截图。"""
        self.n += 1
        url = title = html = ""
        try:
            url = self.page.url or ""
        except Exception:
            pass
        try:
            title = self.page.title or ""
        except Exception:
            pass
        try:
            html = self.page.html or ""
        except Exception:
            html = ""

        safe_label = re.sub(r"[^A-Za-z0-9_-]+", "_", label)[:40]
        png = os.path.join(self.dir, f"{self.n:02d}_{safe_label}.png")
        htmlf = os.path.join(self.dir, f"{self.n:02d}_{safe_label}.html")
        try:
            self.page.screenshot(path=png, full_page=False)
        except Exception as e:
            _log(f"截图失败: {e}", "WARN")
        try:
            with open(htmlf, "w", encoding="utf-8") as f:
                f.write(html or "")
        except Exception:
            pass

        self.steps.append({
            "n": self.n, "label": label, "note": note,
            "url": url, "title": title, "png": os.path.basename(png),
            "html": os.path.basename(htmlf), "ts": datetime.now().strftime("%H:%M:%S"),
        })
        _log(f"[step {self.n}] {label} | url={url[:90]} | title={title[:40]}")
        # 增量写 steps.json,防止最后 save 卡住丢数据
        try:
            with open(os.path.join(self.dir, "steps.json"), "w", encoding="utf-8") as f:
                json.dump(self.steps, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
        return url, title, html

    def save(self):
        self._stop.set()
        time.sleep(0.3)
        if self.listen_started:
            try:
                self.page.listen.stop()
            except Exception:
                pass
        with open(os.path.join(self.dir, "steps.json"), "w", encoding="utf-8") as f:
            json.dump(self.steps, f, ensure_ascii=False, indent=2)
        with open(os.path.join(self.dir, "network.har"), "w", encoding="utf-8") as f:
            json.dump(self.entries, f, ensure_ascii=False, indent=2)
        _log(f"产物已存: {self.dir}  (steps={len(self.steps)} packets={len(self.entries)})")


# ---- 微软登录页元素选择器(参考 register_outlook_ruoyi / standalone)----
EMAIL_SEL = 'css:input[name="loginfmt"], input[type="email"], input[name="MemberName"]'
PWD_SEL = 'css:input[type="password"], input[name="passwd"]'
NEXT_SELS = [
    'css:input[type="submit"]#idSIButton9',
    'css:input[id="idSIButton9"]',
    'css:button[type="submit"]',
    'css:input[type="submit"]',
    'text:Next', 'text:下一步', 'text:Sign in', 'text:登录',
]
# 「是/保持登录」
STAY_SELS = [
    'css:input[id="idSIButton9"]', 'css:input[type="submit"][value="Yes"]',
    'text:Yes', 'text:是', 'text:Continue', 'text:继续',
]
# 辅助邮箱/安全信息页关键信号
# 注意:只留强信号。`secondary`/`email`/`alternate`/`recovery` 等通用词会被 Consent
# 授权同意页(正文含 "email")误命中,导致进 secondary 死循环——那些词不能放这。
SECONDARY_HINTS = [
    "proofs/Add", "proofs/add", "/AddSecurityInfo", "SecurityInfo",
    "Send code", "Get code", "verify your email",
    "Add sign-in method", "添加登录方式", "security info", "安全信息",
    "验证", "辅助", "恢复",
]


def _click_any(page, locators, timeout=2):
    for sel in locators:
        try:
            el = page.ele(sel, timeout=timeout)
            if el and getattr(el, "_is_NoneElement", False) is False and el:
                el.click()
                return sel
        except Exception:
            continue
    return None


def _fill(page, sel, value, timeout=4):
    try:
        el = page.ele(sel, timeout=timeout)
        if el and getattr(el, "_is_NoneElement", False) is False:
            try:
                el.input(value, clear=True)
            except Exception:
                # 回退 JS setter
                page.run_js(
                    f"""(v)=>{{const el=document.querySelector('{sel.split("css:")[1] if sel.startswith('css:') else sel}');
                    if(!el)return false; el.value=v; el.dispatchEvent(new Event('input',{{bubbles:true}}));
                    el.dispatchEvent(new Event('change',{{bubbles:true}})); return true;}}""",
                    value,
                )
            return True
    except Exception:
        pass
    return False


def _fill_any(page, locators, value, timeout=4):
    """逐个选择器尝试填值(DrissionPage 的 ele() 对逗号分隔的复合 CSS 会整体判失败,
    必须拆开一个个试)。命中即返回该 selector。"""
    for sel in locators:
        if _fill(page, sel, value, timeout=timeout):
            return sel
    return None


# Verify 页验证码框候选选择器(实测 16_bind_no_code_input.html:
#   <input type="tel" id="iOttText" name="iOttText" maxlength="16" aria-label="Enter the code we sent you">)
# iOttText 是微软 proofs/Verify 的标准验证码框;其余是不同 uiflavor/旧版兜底。
CODE_INPUT_SELS = [
    'css:#iOttText',
    'css:input[name="iOttText"]',
    'css:#iCode',
    'css:input[name="iCode"]',
    'css:input[id="otc"]',
    'css:input[name="otc"]',
    'css:input[name="otp"]',
    'css:input[name="Code"]',
    # 模糊兜底:name 含 code/otp/ott
    'css:input[name*="ode" i]',
    'css:input[name*="otp" i]',
    'css:input[name*="ott" i]',
]


def _fill_code_entry_digits(page, code, max_wait=15):
    """处理登录「Enter your code」分位数输入框:6 个独立 input
    (id="codeEntry-0..5", maxlength=1)。逐位填入。填成功返回 True。
    code 长度若不足位数则只填前几位。"""
    deadline = time.time() + max_wait
    digits = list(str(code))
    while time.time() < deadline:
        # 等第一个框出现
        first = None
        try:
            first = page.ele('css:#codeEntry-0', timeout=0.3)
            if first and getattr(first, "_is_NoneElement", False) is not False:
                first = None
        except Exception:
            first = None
        if not first:
            time.sleep(0.5)
            continue
        filled_ok = True
        for i, d in enumerate(digits):
            try:
                el = page.ele(f'css:#codeEntry-{i}', timeout=1)
                if not el or getattr(el, "_is_NoneElement", False) is not False:
                    filled_ok = False
                    break
                el.input(d, clear=True)
            except Exception:
                filled_ok = False
                break
        if filled_ok:
            _log(f"分位码框已填 {len(digits)} 位: {code}")
            return True
        time.sleep(0.5)
    _log("分位码框等待超时或填入失败", "WARN")
    return False


def _wait_for_code_input(page, locators=None, max_wait=25, poll=0.5):
    """轮询等验证码输入框渲染出来。微软 proofs/Verify 是 SPA
    (初始 HTML 是 <style>body{display:none}</style>+frame-bust JS,表单靠
    addproofpackage.js 动态渲染),固定等待拿不到 #iOttText,必须轮询。
    返回 (selector, element) 或 (None, None)。"""
    locators = locators or CODE_INPUT_SELS
    deadline = time.time() + max_wait
    while time.time() < deadline:
        for sel in locators:
            try:
                el = page.ele(sel, timeout=0.3)
                if el and getattr(el, "_is_NoneElement", False) is False:
                    # 确认可见(SPA 有 display:none 的占位 div)
                    vis = True
                    try:
                        vis = bool(el.is_displayed()) if hasattr(el, "is_displayed") else True
                    except Exception:
                        vis = True
                    if vis:
                        return sel, el
            except Exception:
                continue
        time.sleep(poll)
    return None, None


def _wait_stable(page, after=2):
    """等 URL/页面稳定。"""
    try:
        page.wait(after)
    except Exception:
        time.sleep(after)


def _handle_login_verify_email(page, dumper, cf_address, cf_jwt, use_admin):
    """处理登录时的「Verify your email」挑战页(已绑辅助邮箱的号复登会触发)。
    微软要你输入完整辅助邮箱地址确认身份:
      页面文案 "We'll send a code to ms*****@nuo.dpdns.org. enter it here"
      输入框 id="proof-confirmation-email-input"
    填入完整 cf 辅助邮箱地址 -> 提交。后续若要码,drive_to_auth 会跟。
    返回 True 表示处理过(无论是否成功),False 表示不在该页。"""
    from common import cloudflare_mail as cm
    try:
        title = (page.title or "").lower()
        html = page.html or ""
    except Exception:
        title = html = ""
    is_verify = ("verify your email" in title
                 or "proof-confirmation-email-input" in html
                 or "we'll send a code to" in html.lower())
    if not is_verify:
        return False
    _log("登录遇到「Verify your email」挑战(已绑辅助邮箱复登触发),填完整辅助邮箱地址")
    dumper.snapshot("login_verify_email")
    filled = _fill_any(page, [
        'css:#proof-confirmation-email-input',
        'css:input[id="proof-confirmation-email-input"]',
        'css:input[name="proof-confirmation-email-input"]',
    ], cf_address, timeout=6)
    if not filled:
        _log("未找到 Verify your email 输入框", "WARN")
        dumper.snapshot("login_verify_no_input")
        return True
    _log(f"已填辅助邮箱地址: {cf_address}")
    _wait_stable(page, 1)
    dumper.snapshot("login_verify_email_filled")
    # 提交前先记下收件箱最高 id 作为基线。微软发码很快,
    # 若提交后才取基线,码邮件可能已到,基线就把它排除了 -> 永远等不到。
    base_last_id = 0
    try:
        if use_admin:
            raws = cm.fetch_admin_mails(cf_address, limit=20)
            base_last_id = max([m.get("id", 0) for m in raws] or [0])
        else:
            mails = cm.fetch_parsed_mails(cf_jwt, limit=20)
            base_last_id = max([m.get("id", 0) for m in mails] or [0])
    except Exception as e:
        _log(f"取初始收件箱失败(继续): {e}", "WARN")
    _log(f"提交前收件箱最高 id={base_last_id}")
    # 提交
    _click_any(page, ['css:button[type="submit"]', 'css:input[type="submit"]',
                      'text:Next', 'text:下一步', 'text:Send code', 'text:Continue'], timeout=3)
    _wait_stable(page, 5)
    dumper.snapshot("login_verify_after_submit")
    # 提交后可能直接进 proofs/Verify 要验证码(微软发码到辅助邮箱)
    try:
        h2 = page.html or ""
    except Exception:
        h2 = ""
    if re.search(r'(enter.*code|code we sent|iOttText|proofs/Verify)', h2, re.I):
        _log("Verify your email 提交后进入验证码页,从 cf 收码并回填")
        code = cm.wait_for_code(cf_jwt, received_after_id=base_last_id, max_wait=180, poll=6,
                                use_admin=use_admin, address=cf_address)
        if not code:
            _log("Verify your email 取码超时", "WARN")
            dumper.snapshot("login_verify_no_code")
            return True
        _log(f"取到验证码: {code}")
        # 登录验证码有两种 UI:
        #   a) 单框 #iOttText(proofs/Verify 风格)
        #   b) 分位框 codeEntry-0..5(每框一位)
        sel_hit, _ = _wait_for_code_input(page, max_wait=8, poll=0.5)
        if sel_hit:
            _fill_any(page, CODE_INPUT_SELS, code, timeout=2)
        else:
            # 试分位框
            _log("未找到 #iOttText,尝试分位码框 codeEntry-N")
            if not _fill_code_entry_digits(page, code, max_wait=15):
                _log("登录验证码框未渲染(#iOttText 和 codeEntry 都没找到)", "WARN")
                dumper.snapshot("login_verify_no_code_input")
                return True
        _click_any(page, ['css:#iNext', 'css:button[type="submit"]', 'css:input[type="submit"]',
                          'text:Next', 'text:Verify'], timeout=3)
        _wait_stable(page, 5)
        dumper.snapshot("login_verify_code_submitted")
    return True


def login_microsoft(page, email, password, dumper):
    """登录微软账号。返回最终 URL。"""
    _log("打开 login.live.com ...")
    page.get("https://login.live.com/", timeout=60)
    _wait_stable(page, 3)
    dumper.snapshot("login_landed")

    # 邮箱
    if _fill(page, EMAIL_SEL, email, timeout=6):
        _log(f"已填邮箱 {email}")
        _click_any(page, NEXT_SELS)
        _wait_stable(page, 4)
        dumper.snapshot("after_email")
    else:
        _log("未找到邮箱框(可能已登录或被重定向)", "WARN")
        dumper.snapshot("no_email_box")

    # 登录挑战:Verify your email(已绑辅助邮箱的号复登触发)——在密码页之前先处理
    if getattr(dumper, "bind_mode", False) and getattr(dumper, "cf_address", None):
        if _handle_login_verify_email(page, dumper, dumper.cf_address, dumper.cf_jwt, dumper.cf_use_admin):
            _wait_stable(page, 3)
            dumper.snapshot("after_login_verify_email")

    # 密码
    if _fill(page, PWD_SEL, password, timeout=6):
        _log("已填密码")
        _click_any(page, NEXT_SELS)
        _wait_stable(page, 5)
        dumper.snapshot("after_password")
    else:
        _log("未找到密码框", "WARN")
        dumper.snapshot("no_password_box")

    # 过中间页:Stay signed in / 隐私 / passkey
    for i in range(6):
        try:
            url = (page.url or "").lower()
            title = (page.title or "").lower()
        except Exception:
            url = title = ""
        # 已经登录成功
        if "account.live.com" in url or "account.microsoft.com" in url or "outlook.live.com" in url:
            _log(f"已进入账号页: {url[:80]}")
            dumper.snapshot(f"post_login_{i}")
            break
        # Stay signed in
        clicked = _click_any(page, STAY_SELS, timeout=2)
        if clicked:
            _log(f"点掉中间页按钮: {clicked}")
            _wait_stable(page, 3)
            dumper.snapshot(f"interstitial_{i}")
            continue
        # 通用继续/下一步
        _click_any(page, ["text:Continue", "text:继续", "text:OK", "text:确定", "text:Next", "text:下一步"], timeout=1)
        _wait_stable(page, 2)
        dumper.snapshot(f"interstitial_{i}")

    return page.url or ""


def bind_secondary_email(page, dumper, cf_address, cf_jwt, use_admin=False):
    """在 proofs/Add 真表单页:填 cf 辅助邮箱 → 提交 → 收码 → 回填 → 完成绑定。
    绑定成功后页面会跳回 oauth,drive_to_auth 继续跟到 code。返回状态。"""
    from common import cloudflare_mail as cm

    _log(f"=== 绑定辅助邮箱: {cf_address} (use_admin={use_admin}) ===")
    # 记下发码前的收件箱最高 id,避免取到旧码
    base_last_id = 0
    try:
        if use_admin:
            raws = cm.fetch_admin_mails(cf_address, limit=20)
            base_last_id = max([m.get("id", 0) for m in raws] or [0])
        else:
            mails = cm.fetch_parsed_mails(cf_jwt, limit=20)
            base_last_id = max([m.get("id", 0) for m in mails] or [0])
    except Exception as e:
        _log(f"取初始收件箱失败(继续): {e}", "WARN")
    _log(f"发码前收件箱最高 id={base_last_id}")

    # 1) 填 EmailAddress
    filled = _fill(page, 'css:#EmailAddress, input[name="EmailAddress"]', cf_address, timeout=6)
    if not filled:
        _log("未找到 EmailAddress 输入框", "WARN")
        dumper.snapshot("bind_no_email_input")
        return "bind_fail_no_input"
    _log(f"已填辅助邮箱 {cf_address}")
    _wait_stable(page, 1)
    dumper.snapshot("bind_email_filled")

    # 2) 提交(action=AddProof)——点 iNext 按钮
    clicked = _click_any(page, ['css:#iNext', 'css:input[type="submit"]', 'text:Next'], timeout=3)
    _log(f"提交绑定(点 {clicked})")
    _wait_stable(page, 5)
    dumper.snapshot("bind_after_submit")

    # 3) 判断提交后是「验证码输入页」还是报错
    try:
        url = (page.url or "").lower()
        html = page.html or ""
    except Exception:
        url = html = ""
    # 验证码页信号
    code_page = bool(re.search(r'(enter.*code|verification code|enter the code|验证码|code we sent|otp|iOttText)'
                               r'|proofs/Verify|id="iOttText"|id="iCode"|name="iCode"|name="otp"|id="otc"',
                               html, re.I))
    # 报错信号
    err = ""
    m = re.search(r'(?:sErrTxt|data-error|error-message)[^>]*>([^<]{3,120})<', html, re.I)
    if m:
        err = m.group(1).strip()
    if err:
        _log(f"提交后页面报错: {err}", "WARN")
    if not code_page and not err:
        # 兜底:正文里找 "code"/"verify"
        code_page = "code" in (re.sub(r'<[^>]+>', ' ', html).lower())

    if not code_page:
        _log("未进入验证码输入页(可能已绑过/或流程变了)", "WARN")
        dumper.snapshot("bind_no_code_page")
        return "bind_fail_no_code_page"
    _log("已进入验证码输入页,等微软发码...")
    dumper.snapshot("bind_code_page")

    # 4) 从 cf 收码
    code = cm.wait_for_code(cf_jwt, received_after_id=base_last_id, max_wait=180, poll=6,
                            use_admin=use_admin, address=cf_address)
    if not code:
        _log("cf 取码超时", "WARN")
        dumper.snapshot("bind_no_code")
        return "bind_fail_no_code"
    _log(f"取到验证码: {code}")

    # 5) 回填验证码并提交(Verify 页:input id="iOttText" name="iOttText" type="tel")
    # 注意:Verify 是 SPA,提交后到 #iOttText 渲染出来有时间差,必须轮询等元素,
    # 不能用固定 timeout(实测提交后页面先停在 body{display:none} 的中间态)。
    sel_hit, el_hit = _wait_for_code_input(page, max_wait=25, poll=0.5)
    if not sel_hit:
        _log("等待验证码输入框超时(Verify SPA 未渲染出 #iOttText)", "WARN")
        dumper.snapshot("bind_no_code_input")
        return "bind_fail_no_code_input"
    _log(f"验证码框已渲染: {sel_hit}")
    code_filled = _fill_any(page, CODE_INPUT_SELS, code, timeout=2)
    if not code_filled:
        # 元素等到了但 input() 失败,直接 JS 兜底写值
        try:
            page.run_js(
                """(v)=>{const el=document.querySelector('#iOttText, input[name="iOttText"]');
                if(!el)return false; el.value=v; el.dispatchEvent(new Event('input',{bubbles:true}));
                el.dispatchEvent(new Event('change',{bubbles:true})); return true;}""",
                code,
            )
            code_filled = True
            _log("input() 失败,已用 JS 兜底写值")
        except Exception as e:
            _log(f"JS 兜底填值也失败: {e}", "WARN")
    _log("已填验证码")
    _wait_stable(page, 1)
    dumper.snapshot("bind_code_filled")

    # 提交验证码
    _click_any(page, ['css:#iNext', 'css:input[type="submit"]', 'text:Next', 'text:Verify'], timeout=3)
    _wait_stable(page, 6)
    dumper.snapshot("bind_after_code_submit")

    # 6) 判断绑定成功(可能回到 proofs/Add 显示已绑,或直接回 oauth)
    try:
        url2 = page.url or ""
    except Exception:
        url2 = ""
    _log(f"验证码提交后 url={url2[:90]}")
    if "proofs" in url2.lower() and "add" in url2.lower():
        # 还在 proofs/Add,可能要求再绑一种或显示成功
        # 看看有没有 "Skip"/"Looks good"/"Next" 推进
        _click_any(page, ['text:Looks good', 'text:Next', 'text:下一步', 'text:Skip', 'text:跳过',
                          'css:#iNext'], timeout=2)
        _wait_stable(page, 4)
        dumper.snapshot("bind_advance")
    _log("辅助邮箱绑定流程结束,回到 drive_to_auth 继续跟 oauth")
    return "bind_done"


def drive_to_auth(page, dumper):
    """打开 Graph 授权 URL,跟到辅助邮箱/安全信息页。"""
    auth_url = _build_auth_url()
    _log(f"打开授权 URL ...")
    page.get(auth_url, timeout=60)
    _wait_stable(page, 4)
    dumper.snapshot("auth_landed", note="授权页起点")

    # 一路跟重定向/中间表单,直到命中辅助邮箱信号 或 localhost(code) 或 超时
    last_url = ""
    stuck = 0
    seen_secondary = False
    for step in range(25):
        try:
            url = page.url or ""
        except Exception:
            url = ""
        if url == last_url:
            stuck += 1
        else:
            stuck = 0
            last_url = url

        # 拿到 code 了 -> 授权成功(不需要绑邮箱)
        if "localhost" in url and "code=" in url:
            parsed = urllib.parse.urlparse(url)
            params = urllib.parse.parse_qs(parsed.query)
            code_val = params.get("code", [None])[0]
            dumper.last_auth_code = code_val
            _log("授权成功,已拿到 code(无需辅助邮箱)")
            dumper.snapshot("auth_done_code", note=f"code={code_val[:20] if code_val else ''}...")
            return "code"
        if "localhost" in url and "error" in url.lower():
            _log(f"授权返回错误: {url[:120]}", "WARN")
            dumper.snapshot("auth_error", note=url[:200])
            return "error"

        # 命中辅助邮箱/安全信息信号?
        try:
            html = page.html or ""
        except Exception:
            html = ""
        low = (url + " " + html).lower()
        hit = [h for h in SECONDARY_HINTS if h.lower() in low]

        # —— Consent(授权同意)页:优先级最高,先于 SECONDARY_HINTS 处理 ——
        # 绑完辅助邮箱后微软会跳到 Consent/Update「Let this app access your info?」,
        # 其正文含 secondary/email 等词会被 SECONDARY_HINTS 误判进死循环。
        is_consent = (
            "/consent/update" in low or "consent/update" in low
            or "appconsent" in low or "access your info" in low
            or "let this app access" in low
        )
        if is_consent or "appConsentPrimaryButton" in html:
            _log("Consent(授权同意)页,点 PrimaryButton 接受")
            clicked = _click_any(page, [
                'css:button[data-testid="appConsentPrimaryButton"]',
                'css:button[data-testid="appConsentPrimary"]',
                'css:#idBtn_Accept',
                'css:input[id="idBtn_Accept"]',
                'css:#acceptButton',
            ], timeout=2)
            if not clicked:
                # 兜底:JS 找 primary 按钮点
                try:
                    page.run_js("""()=>{const b=document.querySelector('button[data-testid="appConsentPrimaryButton"],#idBtn_Accept,#acceptButton');if(b){b.click();return true;}return false;}""")
                    clicked = "js-fallback"
                except Exception as e:
                    _log(f"Consent JS 兜底点击失败: {e}", "WARN")
            _log(f"Consent 点击: {clicked}")
            _wait_stable(page, 5)
            dumper.snapshot(f"consent_{step}")
            continue

        if hit:
            _log(f"命中辅助邮箱信号: {hit}", "OK")
            dumper.snapshot("secondary_email_page", note=f"hits={hit}")
            # 关键:这是 DoSubmit 跳板页(fmHF 自动 POST 到 proofs/Add)还是真正的表单页?
            # 跳板页特征:<body onload="DoSubmit()"> 且 <form name="fmHF"> 且 无可见 input[邮箱]
            is_trampoline = ("DoSubmit()" in html and 'name="fmHF"' in html)
            has_real_email_input = bool(re.search(
                r'<input[^>]*(type="email"|name="(MemberName|email|EmailAddress)"|placeholder[^>]*email)',
                html, re.I))
            _log(f"  跳板页={is_trampoline} 真邮箱框={has_real_email_input}")
            if is_trampoline and not seen_secondary:
                # 主动提交 fmHF,跟到真正的 proofs/Add 表单页
                seen_secondary = True
                _log("提交 DoSubmit 跳板页 fmHF -> 真正 proofs/Add 页")
                try:
                    page.run_js("""()=>{const f=document.forms['fmHF']||document.getElementById('fmHF');if(f){f.submit();return true;}return false;}""")
                except Exception as e:
                    _log(f"提交 fmHF 失败: {e}", "WARN")
                _wait_stable(page, 5)
                dumper.snapshot("after_trampoline")
                continue
            elif has_real_email_input:
                _log("已到真·辅助邮箱输入页")
                dumper.snapshot("secondary_real_form")
                if dumper.bind_mode:
                    bres = bind_secondary_email(page, dumper, dumper.cf_address, dumper.cf_jwt, dumper.cf_use_admin)
                    if not bres or not bres.startswith("bind_done"):
                        _log(f"辅助邮箱绑定未成功: {bres},放弃后续授权", "WARN")
                        return bres
                    _log("辅助邮箱绑定成功,继续跟 oauth 拿 code")
                    # 绑定后微软通常会自动重定向回 oauth(authorize),
                    # 但也可能停在 proofs 页要手动推进。继续主循环跟重定向即可。
                    _wait_stable(page, 4)
                    dumper.snapshot("after_bind_done")
                    # 清掉 seen_secondary,让后续若再次命中 proofs 能正常推进(如要求绑第二种)
                    seen_secondary = False
                    continue
                return "secondary"
            else:
                # 命中信号但既非跳板也非真表单,等渲染
                _wait_stable(page, 3)
                dumper.snapshot("secondary_rendering")
                continue

        # proofs/Add 跳过按钮(项目原有逻辑是 Skip;这里我们要看到表单,先不 skip)
        # 但如果是 proofs/Add 的纯 Skip 页,点 Skip 后可能直接给 code
        try:
            skip_btn = page.ele('css:input[name="action"][value="Skip"]', timeout=1)
        except Exception:
            skip_btn = None

        # consent 页
        try:
            accept = page.ele('css:#idBtn_Accept, input[id="idBtn_Accept"]', timeout=1)
        except Exception:
            accept = None
        if accept and getattr(accept, "_is_NoneElement", False) is False:
            _log("consent 页,点 Accept")
            try:
                accept.click(); _wait_stable(page, 4)
            except Exception:
                pass
            dumper.snapshot(f"consent_{step}")
            continue

        # 卡住或没有明显信号:dump 当前态,尝试点「下一步/继续」类按钮推进
        if stuck >= 2:
            _log(f"疑似卡住(stuck={stuck}),dump 当前态 url={url[:90]}")
            dumper.snapshot(f"stuck_{step}")
            # 尝试通用推进
            clicked = _click_any(page, NEXT_SELS + ["text:Continue", "text:继续", "text:OK", "text:确定"], timeout=1)
            if clicked:
                _wait_stable(page, 3)
                dumper.snapshot(f"advance_{step}")
                stuck = 0
            else:
                # 真没东西可点了
                dumper.snapshot(f"deadend_{step}")
                return "stuck"
        else:
            dumper.snapshot(f"drive_{step}")
            _click_any(page, NEXT_SELS, timeout=1)
            _wait_stable(page, 3)

    _log("跟到 20 步上限", "WARN")
    dumper.snapshot("maxsteps")
    return "maxsteps"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--email", default="")
    ap.add_argument("--password", default="")
    ap.add_argument("--from-nograph", type=int, default=0,
                    help="用 email_nograph.txt 第 N 个账号(从1开始)")
    ap.add_argument("--proxy-url", default=os.environ.get("OUTLOOK_PROXY_URL", ""),
                    help="代理池 HTTP API(GET 返回 socks5 列表);留空读 .env OUTLOOK_PROXY_URL")
    ap.add_argument("--proxy", default="",
                    help="直接指定单条代理(socks5://user:pass@ip:port 或 socks5h://...),跳过代理池 API")
    ap.add_argument("--headless", action="store_true")
    ap.add_argument("--keep", action="store_true", help="跑完不关浏览器(人工接管)")
    ap.add_argument("--bind", action="store_true", help="执行绑定模式:建cf邮箱+真填+收码+回填,跑通绑定")
    args = ap.parse_args()

    rr.set_log_level("INFO")
    rr._install_shutdown_handlers()

    # 取账号
    email, password = args.email, args.password
    if (not email or not password) and args.from_nograph:
        f = os.path.join(ROOT, "email_nograph.txt")
        lines = [l.strip() for l in open(f, encoding="utf-8") if l.strip()]
        if args.from_nograph > len(lines):
            _log(f"email_nograph 只有 {len(lines)} 行", "ERR"); sys.exit(1)
        parts = lines[args.from_nograph - 1].split("----")
        email, password = parts[0], parts[1]
    if not email or not password:
        _log("需要 --email/--password 或 --from-nograph N", "ERR"); sys.exit(1)
    _log(f"账号: {email}")

    # 取代理:优先 --proxy 单条直传,否则 --proxy-url 代理池 API
    proxies_list = []
    if args.proxy:
        proxies_list = [args.proxy.strip()]
        _log(f"使用 --proxy 直传单条: {rr.mask_ruoyi_proxy(proxies_list[0])}")
    else:
        proxy_url = args.proxy_url
        if not proxy_url:
            _log("需要 --proxy <socks5> 或 --proxy-url(代理池 API)或 .env OUTLOOK_PROXY_URL", "ERR"); sys.exit(1)
        _log("拉取代理列表 ...")
        import requests as _rq
        try:
            txt = _rq.get(proxy_url, timeout=30).text
            proxies_list = [l.strip() for l in txt.splitlines() if l.strip()]
        except Exception as e:
            _log(f"拉代理失败: {e}", "ERR"); sys.exit(1)
        if not proxies_list:
            _log("代理列表为空", "ERR"); sys.exit(1)
        _log(f"代理池 {len(proxies_list)} 条,取第1条: {rr.mask_ruoyi_proxy(proxies_list[0])}")

    # 起浏览器
    tb = FirefoxOptions()
    tb.set_browser_path(rr.RUOYI_FIREFOX_PATH)
    profile_dir = rr._ruoyi_profile_dir(argparse.Namespace(ruoyi_slot=None), idx=1)
    tb.set_profile(profile_dir)
    _log(f"profile: {profile_dir}")
    ruoyi_proxy = rr._proxy_url_to_ruoyi(proxies_list[0])
    tb.set_proxy(ruoyi_proxy)
    _log(f"代理已挂: {rr.mask_ruoyi_proxy(ruoyi_proxy)}")
    if args.headless:
        try:
            rr._apply_ruoyi_headless_options(tb, "[probe]", user_agent=rr._pick_user_agent(1))
            tb.headless(True)
        except Exception as e:
            _log(f"headless 失败: {e}", "WARN")

    # 产物目录
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe = email.replace("@", "_at_")
    out_dir = os.path.join(ROOT, "outlook_accounts", f"probe_{safe}_{ts}")

    # 绑定模式:建 cf 辅助邮箱 + 给 cf 模块设代理(收码要经代理访问 cf worker)
    cf_address = cf_jwt = None
    cf_use_admin = False
    if args.bind:
        from common import cloudflare_mail as cm
        cm.set_proxy(proxies_list[0].replace("socks5://", "socks5h://", 1))
        try:
            d = cm.create_or_get_address(email)
            cf_address, cf_jwt = d["address"], d["jwt"]
            cf_use_admin = d.get("use_admin", False)
            _log(f"cf 辅助邮箱就绪: {cf_address} (use_admin={cf_use_admin})")
        except Exception as e:
            _log(f"建 cf 辅助邮箱失败: {e}", "ERR"); sys.exit(1)

    page = None
    dumper = None
    try:
        page = FirefoxPage(tb)
        rr._track_browser_page(page)
        try:
            page.close_other_tabs(page)
        except Exception:
            pass
        dumper = ProbeDumper(out_dir, page)
        if args.bind and cf_address:
            dumper.bind_mode = True
            dumper.cf_address = cf_address
            dumper.cf_jwt = cf_jwt
            dumper.cf_use_admin = cf_use_admin
        dumper.start_listen()

        # 先探出口 IP
        try:
            page.get("https://ipinfo.io/json", timeout=30)
            _wait_stable(page, 3)
            body = page.run_js_loaded("return document.body.innerText") or ""
            _log(f"ruyi 出口 IP: {body.strip()[:120]}")
            dumper.snapshot("exit_ip", note=body.strip()[:200])
        except Exception as e:
            _log(f"探出口失败: {e}", "WARN")

        # 登录
        login_microsoft(page, email, password, dumper)

        # 开授权
        result = drive_to_auth(page, dumper)
        _log(f"流程结束,结果: {result}")

        # 若绑完已拿到 code,用纯 HTTP 换 RT 验证协议授权可行
        if result == "code" and args.bind:
            _log("=== 绑定后授权成功,用纯 HTTP 换 refresh_token 验证协议路径 ===")
            try:
                auth_code = dumper.last_auth_code
                if auth_code:
                    import requests as _rq2
                    sess = _rq2.Session(); sess.trust_env = False
                    tr = sess.post("https://login.microsoftonline.com/consumers/oauth2/v2.0/token",
                        data={"client_id": CLIENT_ID, "grant_type": "authorization_code",
                              "code": auth_code, "redirect_uri": REDIRECT_URI, "scope": SCOPE}, timeout=30)
                    td = tr.json()
                    rt = td.get("refresh_token", "")
                    _log(f"协议换 RT: {'成功' if rt else '失败'} rt前30={rt[:30]}...")
                    with open(os.path.join(out_dir, "token_result.json"), "w", encoding="utf-8") as f:
                        json.dump({"email": email, "refresh_token": rt[:30] + "...", "ok": bool(rt)}, f, ensure_ascii=False, indent=2)
                else:
                    _log("未捕获到 auth code,无法换 RT", "WARN")
            except Exception as e:
                _log(f"换 RT 异常: {e}", "WARN")

    except KeyboardInterrupt:
        _log("收到 Ctrl-C", "WARN")
    except Exception as e:
        _log(f"运行异常: {type(e).__name__}: {e}", "ERR")
        import traceback
        traceback.print_exc()
    finally:
        if dumper:
            try:
                dumper.save()
            except Exception as e:
                _log(f"保存产物失败: {e}", "WARN")
        if page is not None and not args.keep:
            try:
                page.quit(timeout=5)
                _log("浏览器已关闭")
            except Exception as e:
                _log(f"关浏览器失败: {e}", "WARN")
        try:
            rr._cleanup_ruoyi_run_profile_dir(profile_dir)
        except Exception:
            pass

    _log(f"产物目录: {out_dir}")


if __name__ == "__main__":
    main()
