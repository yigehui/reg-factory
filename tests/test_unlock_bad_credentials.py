import unittest
from unittest.mock import patch


import unlock_outlook as mod


def _fake_page(text="", url=""):
    """构造一个最小 fake page,让 classify 取到 text 和 url。

    classify 实际调用链(经核对 unlock_outlook.py:194-237 + common.ruyi._all_contexts
    + register_outlook_ruoyi._context_text/_body_text/_microsoft_loading_page):
      - page.url                       (属性)
      - page.get_all_frames()          (_all_contexts 里 extend,try/except 包裹)
      - ctx.run_js_loaded(script)       (_context_text 取 body.innerText)
      - ctx.run_js_loaded(heading js)   (取 h1/h2 innerText)
      - page.run_js_loaded(script)      (_microsoft_loading_page -> _body_text)
    全部 JS 调用统一返回 text,足够让 classify 的文案匹配走通。
    """
    class _Ctx:
        def run_js_loaded(self, js, *a, **k):
            return text
    ctx = _Ctx()

    class _Page:
        def __init__(self):
            self.url = url
        def get_all_frames(self):
            return []
        def run_js_loaded(self, js, *a, **k):
            return text
    page = _Page()
    # 让 _all_contexts 返回 [page] + frames;这里 frames 为空即可
    # (classify 对 frames 的 try/except 容错,空列表不影响主 context)
    return page


class ClassifyLoginErrorTests(unittest.TestCase):
    def test_classify_login_error_account_not_found(self):
        page = _fake_page(text="We couldn't find an account with that username.")
        self.assertEqual(mod.classify(page), "login_error")

    def test_classify_login_error_wrong_password(self):
        page = _fake_page(text="Your account or password is incorrect.")
        self.assertEqual(mod.classify(page), "login_error")

    def test_classify_email_form_still_works(self):
        page = _fake_page(text="Sign in\nemail or phone", url="https://login.live.com/login.srf")
        self.assertEqual(mod.classify(page), "email_form")


class FastFailLoginErrorTests(unittest.TestCase):
    def test_unlock_account_returns_bad_credentials_on_login_error(self):
        """classify 首次即返回 login_error -> unlock_account 立即返回 bad_credentials,
        不等 deadline。用 time 上下文证明没耗满 timeout。"""
        import time as _time
        page = _fake_page(text="Your account or password is incorrect.")
        # fake page 的 run_js_loaded 返回错误文案(非空字符串),会让 _try_again_button_scan
        # 误判"找到 Try again 按钮"返回 True(因为它用 run_js_loaded 当 JS 结果,bool(非空)=True),
        # 进而主循环反复 _click_try_again_if_present -> continue 直到 deadline,
        # 永远走不到 login_error 分支。故 patch 掉 snap(直接返回 login_error)与
        # _click_try_again_if_present(返回 False,模拟没点到 Try again),让 unlock_account
        # 首轮即进 login_error 分支验证快失败。
        with patch.object(mod, "snap", return_value="login_error"), \
             patch.object(mod, "_click_try_again_if_present", return_value=False):
            t0 = _time.perf_counter()
            outcome = mod.unlock_account(page, "x@outlook.com", "wrong", "w0", 0,
                                         max_press=1, timeout=60)
            elapsed = _time.perf_counter() - t0
        self.assertEqual(outcome, "bad_credentials")
        self.assertLess(elapsed, 10.0)  # 远小于 60s timeout,证明快失败


if __name__ == "__main__":
    unittest.main()
