import unittest

from playwright.sync_api import sync_playwright

import register_outlook_ruoyi as mod


class _PlaywrightCtx:
    def __init__(self, page):
        self.page = page

    def run_js_loaded(self, script):
        return self.page.evaluate(f"() => {{ {script} }}")

    # run_js 别名:热路径探测已切 run_js(不等 doc_loaded)
    run_js = run_js_loaded


class PxI18nSelectorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._pw = sync_playwright().start()
        try:
            cls._browser = cls._pw.chromium.launch(headless=True)
        except Exception as exc:
            cls._pw.stop()
            raise unittest.SkipTest(f"Playwright chromium unavailable: {exc}")

    @classmethod
    def tearDownClass(cls):
        cls._browser.close()
        cls._pw.stop()

    def setUp(self):
        self.page = self._browser.new_page(viewport={"width": 800, "height": 600})
        self.ctx = _PlaywrightCtx(self.page)

    def tearDown(self):
        self.page.close()

    def _set_japanese_px_markup(self, hidden=False):
        hidden_style = "display:none;" if hidden else "display:block;"
        self.page.set_content(
            f"""
<!doctype html>
<html>
  <body>
    <a id="outer-link" role="button" aria-label="アクセス可能なチャレンジ"
       style="display:block;width:320px;height:42px;">
      <div id="real-btn" role="button" aria-label="長押しヒューマンチャレンジ"
           style="display:block;width:275px;height:42px;text-align:center;">
        <div id="inner-wrap">
          <p id="hold-label" style="{hidden_style}margin:0;">長押し</p>
          <span>ヒューマンチャレンジには検証が必要です。</span>
        </div>
      </div>
    </a>
  </body>
</html>
"""
        )

    def test_find_hold_target_prefers_japanese_role_button_with_aria_label(self):
        self._set_japanese_px_markup()
        target = mod._find_hold_target(self.ctx)
        self.assertEqual("real-btn", (target or {}).get("id"))

    def test_resolve_hold_label_for_target_matches_japanese_hold_text(self):
        self._set_japanese_px_markup()
        label = mod._resolve_hold_label_for_target(
            self.ctx,
            {"id": "real-btn", "x": 150, "y": 21, "left": 0, "top": 0, "width": 275, "height": 42},
        )
        self.assertEqual("hold-label", (label or {}).get("id"))
        self.assertIn("長押し", (label or {}).get("text", ""))

    def test_px_hold_instruction_state_accepts_japanese_hold_label(self):
        self._set_japanese_px_markup(hidden=True)
        state = mod._px_hold_instruction_state(self.ctx, "hold-label")
        self.assertEqual("hold-label", (state or {}).get("id"))
        self.assertEqual("none", (state or {}).get("display"))


if __name__ == "__main__":
    unittest.main()
