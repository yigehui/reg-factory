"""Edge 注册链路 spike：真 Edge(Chromium)+ DrissionPage 在未烧 IP 上能否过 PX。

一次性验证(verification spike, 见 Plan docs/superpowers/plans/2026-09-02-edge-register.md)。
不是 Plan A 本身。3 次 ≥2 次 PASSED_PX → 继续实现 common/edge + register_outlook_edge.py。
3 次全 BLOCKED → 停,IP 段烧穿,先换代理池。
"""

import argparse
import os
import random
import string
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "common"))

from DrissionPage import ChromiumOptions, ChromiumPage
# 4.1.1.4 默认超时不抛、返 falsy NoneElement(见移植表);这里显式开抛,_ele 包 try 统一返 None。
from DrissionPage._functions.settings import Settings as _DSettings

_DSettings.set_raise_when_ele_not_found(True)

_EPL = "https://signup.live.com/signup?lic=1"
_EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"


def _edge_ua():
    """从 msedge.exe FileVersion 拼 UA(默认 headless-UA 特征是 PX 杀点之一)。"""
    ver = "152"
    try:
        import subprocess
        out = subprocess.check_output(
            ["powershell", "-NoProfile", "-NonInteractive",
             "-Command", "(Get-Item '%s').VersionInfo.FileVersion" % _EDGE],
            text=True, timeout=10)
        lines = (out or "").strip().splitlines()
        if lines:
            ver = lines[-1].strip().split(".")[0] or "152"
    except Exception:
        pass
    return ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/%s.0.0.0 Safari/537.36 Edg/%s.0.0.0") % (ver, ver)


def _stealth_js(ua):
    """PX 过不过的三大件：webdriver 抹除、hardwareConcurrency=16、UA 对齐(port _build_headless_patch_js)。"""
    return """
(() => {
  Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
  Object.defineProperty(navigator, 'hardwareConcurrency', { get: () => 16 });
  try {
    Object.defineProperty(navigator, 'userAgent', { get: () => '%(ua)s' });
  } catch (e) {}
  Object.defineProperty(navigator, 'languages', { get: () => ['en-US', 'en'] });
  Object.defineProperty(navigator, 'plugins', { get: () => [1, 2, 3, 4, 5] });
  Object.defineProperty(screen, 'width',  { get: () => 1920 });
  Object.defineProperty(screen, 'height', { get: () => 1080 });
  Object.defineProperty(screen, 'availHeight', { get: () => 1040 });
  Object.defineProperty(screen, 'colorDepth', { get: () => 24 });
  window.chrome = window.chrome || { runtime: {} };
  const s = document.createElement('script');
  s.textContent = 'window.__spike=' + navigator.hardwareConcurrency;
  (document.head || document.documentElement).appendChild(s);
})()
""" % {"ua": ua}


def _ele(page, loc, timeout=5):
    try:
        return page.ele(loc, timeout=timeout)
    except Exception:
        return None


def _better_ele(frame, loc, timeout=4):
    try:
        return frame.ele(loc, timeout=timeout)
    except Exception:
        return None


# ---- selectors 与 ruyi 逐字一致(避免双头维护 drift) ----
_SEL_EMAIL = ('css:input[type="email"], input[name="Email"], '
             '#iSignupEmail, input[data-testid*="email" i], input[id*="email" i]')
_SEL_PASS = ('css:input[type="password"], input[name="Password"], '
             '#iPwd, input[data-testid*="password" i], input[id*="pwd" i]')
_SEL_BDAY_Y = ('css:select[name="BirthYear"], input[name="BirthYear"], '
               '#BirthYear, #BirthYearInput, input[id*="BirthYear" i]')
_SEL_NAME_F = ('css:#firstNameInput, input[name="firstNameInput"], input[name="FirstName"], '
               '#FirstName, input[id*="firstName" i]')
_HOLD_JS = """
(() => {
  const nodes = [...document.querySelectorAll('[data-task], [data-part], [role], p, span, div')];
  const hits = nodes.filter(e => {
    const t = (e.innerText || e.textContent || '').trim().toLowerCase();
    return /press and hold|按住|hold to verify|hold down/.test(t);
  });
  const el = hits.find(e => e.offsetParent !== null) || hits[0] || null;
  if (!el) return { found: false };
  const r = el.getBoundingClientRect();
  const cs = getComputedStyle(el);
  return {
    found: true,
    text: (el.innerText || el.textContent || '').slice(0, 40),
    display: cs.display,
    visibility: cs.visibility,
    rect: { x: r.x, y: r.y, w: r.width, h: r.height },
  };
})()
"""


def _rand_hold(lo, hi):
    return random.uniform(lo, hi)


def _hold_px(page, frame, px_el, tag, idx):
    """按住 PX 按钮。px_el 是 frame 内 '#px-captcha' 的 ChromiumElement。
    ruyi 的 hold 是 10.5-11s，前 5s 后每 0.5s 探测 hold 指令是否 display:none → 提前松手。"""
    r = px_el.rect  # 顶层视口
    sz = r.size     # tuple (width, height)
    w = float(sz[0]) if sz and sz[0] else 98.0
    h = float(sz[1]) if sz and sz[1] else 48.0
    mp = px_el.rect.viewport_midpoint
    base = {"x": float(mp[0]), "y": float(mp[1]), "width": w, "height": h}

    hold_secs = _rand_hold(10.5, 11.0)
    # 运动参数 port _ruoyi_point_inside_target / _perform_hold
    def pt(ratio_x, ratio_y, jx=0.0, jy=0.0):
        return (int(round(base["x"] + (base["width"] * ratio_x) + random.uniform(-jx, jx))),
                int(round(base["y"] + (base["height"] * ratio_y) + random.uniform(-jy, jy))))
    lead = pt(0.50, 0.55, 2, 1)              # 先落上一个点
    settle = pt(0.47, 0.58, 1, 1)            # 再挪近
    micro = pt(0.50, 0.59, 1, 1)             # 微移
    press = pt(0.48, 0.60, 1, 2)             # 落按钮中心偏下

    try:
        ac = frame.actions   # ChromiumFrame.actions 会先等 doc_loaded
        ac.move_to(px_el, duration=0.9)
        ac.move_to((lead[0], lead[1]), duration=0.42)
        ac.move_to((settle[0], settle[1]), duration=0.30)
        ac.move_to((micro[0], micro[1]), duration=0.18)
        ac.hold()
    except Exception as exc:
        return {"ok": False, "why": "move_pre:%s" % exc}

    t0 = time.time()
    released_early = False
    probe_every = 0.5
    probe_at = 5.0
    while True:
        held = time.time() - t0
        if held >= hold_secs:
            break
        if held >= probe_at:
            try:
                st = frame.run_js_loaded(_HOLD_JS)
                if isinstance(st, dict) and st.get("display") == "none":
                    released_early = True
                    break
            except Exception:
                pass
            probe_at += probe_every
        time.sleep(0.2)

    try:
        ac.release()
        ac.wait(0.25)
    except Exception:
        try:
            ac = frame.actions
            ac.release()
        except Exception:
            pass

    return {"ok": True, "held": round(time.time() - t0, 2), "early": released_early}


def _run_one(n, proxy):
    tag = "[#%d][edge-spike]" % n
    # 代理链：front 来自 LAUNCH_FRONT_PROXY
    from common.ruyi import chain as rchain
    try:
        front = rchain.front_proxy_raw(os.environ.get("LAUNCH_FRONT_PROXY", ""))
        relay = rchain.front_relay_url(proxy, front=front, tag=tag)
    except Exception:
        return "PROXY_FAIL"
    if not relay:
        return "PROXY_FAIL"
    uid = "spike%s" % "".join(random.choices(string.ascii_lowercase + string.digits, k=10))
    user = "%s_%s@outlook.com" % (uid, random.randint(0, 9999))

    ua = _edge_ua()   # 每次 run 只查一次 PowerShell(原实现每 run 两次,10s 超时 spawn 风险)
    opts = ChromiumOptions()
    opts.set_browser_path(_EDGE)
    opts.set_proxy(relay)
    opts.set_user_agent(ua)
    opts.auto_port()   # 每次独立 CDP,不留相连
    opts.set_argument("--disable-blink-features=AutomationControlled")
    opts.set_argument("--disable-gpu")
    opts.set_argument("--no-first-run")
    # 保留:真 Edge 首次运行会弹欢迎页,干扰 DOM
    opts.set_argument("--no-default-browser-check")
    opts.set_argument("--disable-sync")
    opts.set_argument("--blink-settings=imagesEnabled=false")
    opts.headless(True)

    page = None
    try:
        page = ChromiumPage(opts)
        page.add_init_js(_stealth_js(ua))
        page.get(_EPL)
        time.sleep(3)

        em = _ele(page, _SEL_EMAIL, 12)
        if em is None:
            page.get_screenshot(path="spike_t0_%d.png" % n)
            txt = (page.html or "")[:1800]
            return "BLOCKED" if ("blocked" in txt.lower() or "we can't" in txt.lower()
                                 or "unusual" in txt.lower()) else "FORM_FAIL"
        em.input(user)
        nxt = _ele(page, 'css:button[data-testid="primaryButton"], #iNext, css:input[type="submit"]', 4)
        if nxt is None:
            return "FORM_FAIL"
        nxt.click()
        time.sleep(2.5)

        pw = _ele(page, _SEL_PASS, 8)
        if pw is None:
            return "FORM_FAIL"
        pw.input("EdgeSp!ke-%d%s" % (n, random.randint(1000, 9999)))
        nxt2 = _ele(page, 'css:button[data-testid="primaryButton"], #iNext, css:input[type="submit"]', 4)
        if nxt2 is not None:
            nxt2.click()
        time.sleep(3)

        # 等 PX iframe(hsprotect/px)里出现 #px-captcha
        frame = None
        px = None
        for _ in range(30):
            try:
                for f in page.get_frames():
                    try:
                        fr = page.get_frame(f)   # element → ChromiumFrame(OOPIF 亦解析)
                    except Exception:
                        continue
                    furl = (getattr(fr, "url", None) or "").lower()
                    if "hsprotect" in furl or "#px" in furl or "px-captcha" in furl:
                        el_px = _better_ele(fr, "#px-captcha", 2)
                        if el_px is not None:
                            frame, px = fr, el_px
                            break
            except Exception:
                pass
            if frame is not None:
                break
            time.sleep(0.5)

        if frame is None:
            page.get_screenshot(path="spike_no_px_%d.png" % n)
            return "BLOCKED"

        res = _hold_px(page, frame, px, tag, n)
        time.sleep(2.0)

        # PX 后判过:两类信号轮询 6s —— ① 已跳登录成功/Account 页;② #px-captcha 消失(iframe 卸载)。
        # 单一 2s 查询会撞 teardown 竞态(通过但未跳转时 iframe 仍在)→ 误判 BLOCKED。
        page.get_screenshot(path="spike_px_%d.png" % n)
        if res.get("ok"):
            for _ in range(6):
                try:
                    url = str(page.url).lower()
                except Exception:
                    url = ""
                if any(h in url for h in ("login.live.com", "account.live.com")):
                    return "PASSED_PX"
                still = _better_ele(page, 'css:[id*="px-captcha"], iframe[src*="px"]', 1.5)
                if not still:
                    return "PASSED_PX"
                time.sleep(1.0)
        return "BLOCKED"
    finally:
        if page is not None:
            try:
                page.quit()
            except Exception:
                pass


def main():
    ap = argparse.ArgumentParser(description="Edge spike")
    ap.add_argument("--proxy-file", default=os.environ.get("OUTLOOK_PROXY_FILE", "proxy.txt"))
    ap.add_argument("--count", type=int, default=3)
    args = ap.parse_args()
    with open(args.proxy_file, "r", encoding="utf-8") as f:
        lines = [l.split()[-1].strip() for l in f if l.strip() and not l.strip().startswith("#")]
    if not lines:
        print("PROXY_FILE 为空", flush=True)
        return 1
    results = []
    for i in range(1, args.count + 1):
        proxy = lines[(i - 1) % len(lines)]
        print("run %d proxy=%s" % (i, proxy.split("@")[-1] if "@" in proxy else proxy), flush=True)
        r = _run_one(i, proxy)
        results.append(r)
        print("RESULT[%d]: %s" % (i, r), flush=True)
    px = [r for r in results if r == "PASSED_PX"]
    print("PX_SUMMARY: pass=%d/%d runs=%s" % (len(px), len(results), results), flush=True)
    return 0 if len(px) >= 2 else 1


if __name__ == "__main__":
    sys.exit(main())
