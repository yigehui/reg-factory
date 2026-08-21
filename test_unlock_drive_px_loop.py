# -*- coding: utf-8 -*-
"""TDD for _drive_px_loop:忠实移植 ruoyi 滑块状态机到解锁场景。

用 monkeypatch 打桩 rr 上的 helper + 可控时钟,测纯逻辑分支(不起浏览器)。
"""
import importlib
import types

import pytest

import unlock_outlook_ruoyi as uo
import register_outlook_ruoyi as rr


# ── 公共测试脚手架 ──────────────────────────────────────────────────────
class FakePage:
    """最小 page 替身:带 .url、记录 .get 调用。"""

    def __init__(self, url="https://login.live.com/Abuse"):
        self.url = url
        self.gets = []

    def get(self, url, timeout=None):
        self.gets.append(url)
        self.url = url


class FakeClock:
    """可控时钟:list 每读一次 time.time() 弹一个,空则停在最后值。"""

    def __init__(self, start=0.0, step=1.0):
        self._t = start
        self._step = step
        self._overrides = []
        self.sleeps = []

    def time(self):
        if self._overrides:
            return self._overrides.pop(0)
        self._t += self._step
        return self._t

    def sleep(self, secs):
        self.sleeps.append(secs)
        # 真推进时钟,模拟 sleep 期间时间流逝
        self._t += float(secs)


def _install_clock(monkeypatch, clock):
    monkeypatch.setattr(uo.time, "time", clock.time)
    monkeypatch.setattr(uo.time, "sleep", clock.sleep)


def _stub_rr(monkeypatch, *, visible=False, validating=False, hold_ctx=None,
             hold_target=None, target_quality=9, loading=False,
             perform_hold_ret=None, body_text=""):
    """打桩 rr 上的 helper,返回调用记录 dict。"""
    calls = {"perform_hold": [], "headless_patches": [], "find_hold": 0,
             "click_any": 0}

    def _visible(page):
        return visible

    def _validating(page):
        return validating

    def _find(page, min_quality=5):
        calls["find_hold"] += 1
        return hold_ctx, hold_target

    def _quality(t):
        return target_quality

    def _loading(page):
        return loading

    def _perform_hold(page, ctx, target, idx, press_count, tag):
        calls["perform_hold"].append((press_count, tag))
        return perform_hold_ret

    def _patches(page, tag=None, log_once=False, user_agent=None):
        calls["headless_patches"].append((tag, log_once, user_agent))

    def _wait(*a, **k):
        pass

    def _body_text(page):
        return body_text

    monkeypatch.setattr(rr, "_captcha_visible", _visible)
    monkeypatch.setattr(rr, "_captcha_is_validating", _validating)
    monkeypatch.setattr(rr, "_find_hold_context", _find)
    monkeypatch.setattr(rr, "_target_quality", _quality)
    monkeypatch.setattr(rr, "_microsoft_loading_page", _loading)
    monkeypatch.setattr(rr, "_perform_hold", _perform_hold)
    monkeypatch.setattr(rr, "_apply_ruoyi_headless_page_patches", _patches)
    monkeypatch.setattr(rr, "_wait_before_next_captcha_press", _wait)
    monkeypatch.setattr(rr, "_body_text", _body_text)
    return calls


# ── 1. 成功:logged_in ────────────────────────────────────────────────────
def test_returns_unlocked_when_logged_in(monkeypatch):
    clock = FakeClock(start=1000.0, step=1.0)
    _install_clock(monkeypatch, clock)
    _stub_rr(monkeypatch)
    page = FakePage()
    snaps = [("logged_in", "welcome")]

    def snap_cb(p, t, name):
        return snaps.pop(0) if snaps else ("logged_in", "x")

    rc = uo._drive_px_loop(page, "[t]", max_press=5,
                           deadline=clock.time() + 300,
                           headless=False, user_agent="UA", snap_cb=snap_cb)
    assert rc == "unlocked"


# ── 2. needs_phone ──────────────────────────────────────────────────────
def test_returns_needs_phone(monkeypatch):
    clock = FakeClock(start=1000.0, step=1.0)
    _install_clock(monkeypatch, clock)
    _stub_rr(monkeypatch)
    page = FakePage()

    def snap_cb(p, t, name):
        return ("sms_verify", "we texted")

    rc = uo._drive_px_loop(page, "[t]", max_press=5,
                           deadline=clock.time() + 300,
                           headless=False, user_agent="UA", snap_cb=snap_cb)
    assert rc == "needs_phone"


# ── 3. 首次按压跨过 INITIAL_PRESS_DELAY ───────────────────────────────
def test_first_press_after_initial_delay(monkeypatch):
    # 初始延迟设 0 让第1轮就能进按压路径
    monkeypatch.setattr(rr, "INITIAL_PRESS_DELAY", 0)
    monkeypatch.setattr(rr, "POST_MAX_PRESS_WAIT", 0)
    monkeypatch.setattr(rr, "CAPTCHA_STATE_TIMEOUT", 1000)
    clock = FakeClock(start=1000.0, step=1.0)
    _install_clock(monkeypatch, clock)
    calls = _stub_rr(monkeypatch, visible=True, validating=False,
                     hold_ctx="ctx", hold_target={"quality": 3, "x": 10, "y": 10},
                     target_quality=3, perform_hold_ret=0.5)
    page = FakePage()

    def snap_cb(p, t, name):
        return ("px_challenge", "press and hold")

    rc = uo._drive_px_loop(page, "[t]", max_press=5,
                           deadline=clock.time() + 300,
                           headless=False, user_agent="UA", snap_cb=snap_cb)
    # stub 让 captcha 一直 actionable(真实场景按压后会消失/校验),所以状态机
    # 会一路按压到 max_press=5,再等 POST_MAX_PRESS_WAIT 后返回 failed_px_challenge。
    # 这里验证:按压确实发生、press_count 递增、终态正确。
    assert len(calls["perform_hold"]) == 5, "应按压满 max_press=5 次"
    assert [c[0] for c in calls["perform_hold"]] == [1, 2, 3, 4, 5]
    assert rc == "failed_px_challenge"


# ── 4. headless 应用补丁 ─────────────────────────────────────────────────
def test_headless_applies_patches(monkeypatch):
    monkeypatch.setattr(rr, "INITIAL_PRESS_DELAY", 0)
    clock = FakeClock(start=1000.0, step=1.0)
    _install_clock(monkeypatch, clock)
    calls = _stub_rr(monkeypatch, visible=False, validating=False)

    def snap_cb(p, t, name):
        return ("logged_in", "ok")

    page = FakePage()
    uo._drive_px_loop(page, "[t]", max_press=5,
                      deadline=clock.time() + 300,
                      headless=True, user_agent="UA1", snap_cb=snap_cb)
    assert len(calls["headless_patches"]) >= 1
    assert calls["headless_patches"][0] == ("[t]", True, "UA1")

    calls["headless_patches"].clear()
    page2 = FakePage()
    uo._drive_px_loop(page2, "[t]", max_press=5,
                      deadline=clock.time() + 300,
                      headless=False, user_agent="UA2", snap_cb=snap_cb)
    assert calls["headless_patches"] == [], "headless=False 不应打补丁"


# ── 5. validating 阻止按压 ──────────────────────────────────────────────
def test_validating_blocks_press(monkeypatch):
    monkeypatch.setattr(rr, "INITIAL_PRESS_DELAY", 0)
    monkeypatch.setattr(rr, "CAPTCHA_STATE_TIMEOUT", 1000)  # 避免验证等待超时
    clock = FakeClock(start=1000.0, step=1.0)
    _install_clock(monkeypatch, clock)
    # visible=True,validating=True → had_captcha 初始 False,走 else 分支点 Next
    # 这里关键是即便 visible 也不调 _perform_hold
    calls = _stub_rr(monkeypatch, visible=True, validating=True,
                     hold_ctx="ctx", hold_target={"quality": 3},
                     target_quality=3, perform_hold_ret=0.5)

    def snap_cb(p, t, name):
        return ("px_challenge", "press and hold")

    page = FakePage()
    rc = uo._drive_px_loop(page, "[t]", max_press=5,
                           deadline=clock.time() + 300,
                           headless=False, user_agent="UA", snap_cb=snap_cb)
    # 不会按压(因为 had_captcha 起步 False,但 visible+validating 触发 had_captcha 分支不按)
    # 实际上首轮 visible+validating+not had_captcha → 走 else 分支点 Next;
    # 没按压,snap 一直 px_challenge,deadline 到 → failed_timeout
    assert len(calls["perform_hold"]) == 0, "validating 期间不应按压"
    assert rc == "failed_timeout"


# ── 6. max_press 耗尽 ────────────────────────────────────────────────────
def test_max_press_exhausted_returns_failed_px(monkeypatch):
    monkeypatch.setattr(rr, "INITIAL_PRESS_DELAY", 0)
    monkeypatch.setattr(rr, "POST_MAX_PRESS_WAIT", 0)
    # 让 POST_PRESS_LOADING_CHECK 也不挡,但 reappear 等待需要走完
    monkeypatch.setattr(rr, "CAPTCHA_STATE_TIMEOUT", 1000)
    clock = FakeClock(start=1000.0, step=1.0)
    _install_clock(monkeypatch, clock)
    calls = _stub_rr(monkeypatch, visible=True, validating=False,
                     hold_ctx="ctx", hold_target={"quality": 3},
                     target_quality=3, perform_hold_ret=0.5)

    def snap_cb(p, t, name):
        return ("px_challenge", "press and hold")

    page = FakePage()
    rc = uo._drive_px_loop(page, "[t]", max_press=2,
                           deadline=clock.time() + 300,
                           headless=False, user_agent="UA", snap_cb=snap_cb)
    assert rc == "failed_px_challenge"
    assert len(calls["perform_hold"]) == 2, "应按压满 max_press=2 次"


# ── 7. error_page 3 次 → failed_error_page ──────────────────────────────
def test_error_page_three_times_failed_error_page(monkeypatch):
    clock = FakeClock(start=1000.0, step=1.0)
    _install_clock(monkeypatch, clock)
    _stub_rr(monkeypatch)
    page = FakePage()

    def snap_cb(p, t, name):
        return ("error_page", "something went wrong")

    rc = uo._drive_px_loop(page, "[t]", max_press=5,
                           deadline=clock.time() + 300,
                           headless=False, user_agent="UA", snap_cb=snap_cb)
    assert rc == "failed_error_page"
    # 验证不用 history.back,用 page.get 重建
    assert any("Abuse" in u or "login.live" in u for u in page.gets), \
        "应通过 page.get 重建而非 history.back"
