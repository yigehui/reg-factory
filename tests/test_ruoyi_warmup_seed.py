"""warmup 首屏预热 + 挑战 cookie 种子池测试。

覆盖:
- save_seed: 白名单过滤(登录态 ESTSAUTH 绝不入池) / BiDi value dict 归一化 / 去重
- load_seed: (proxy, ua) 全等才命中 / TTL 过期返回 []
- warmup: 已在首页不重复导航 / 种子回灌 / about:blank 时导航
"""

import json
import os
import tempfile
import time
import unittest

from common.ruyi import (
    CHALLENGE_COOKIE_ALLOWLIST,
    load_seed,
    save_seed,
    warmup,
)

PROXY = "http://user:pass@1.2.3.4:8080"
UA = "Mozilla/5.0 UA-1"


class _FakeCookie(object):
    def __init__(self, raw):
        self.raw = raw


class _FakePage(object):
    url = "https://login.live.com/"

    def __init__(self, cookies=None):
        self._cookies = cookies if cookies is not None else _default_raw()
        self.set_called = None
        self.nav = None

    def get_cookies(self, all_info=False):
        return [_FakeCookie(c) for c in self._cookies]

    def set_cookies(self, cookies):
        self.set_called = list(cookies)
        return None

    def wait_loading(self, timeout):
        return True

    def get(self, url):
        self.nav = url


def _default_raw():
    return [
        {"name": "_pxhd", "value": "abc", "domain": ".live.com", "path": "/", "secure": True},
        {"name": "__cf_bm", "value": {"type": "string", "value": "cf123"}, "domain": ".live.com"},
        {"name": "ESTSAUTH", "value": "SECRET_LOGIN", "domain": ".login.live.com"},
        {"name": "canary", "value": "c9", "domain": "login.live.com"},
    ]


class WarmupSeedTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.mkdtemp()
        self.seed_file = os.path.join(self._tmp, "seeds.json")

    def test_save_seed_filters_login_cookies(self):
        n = save_seed(_FakePage(), "unittest_pl", PROXY, UA, "T", seed_file=self.seed_file)
        self.assertEqual(n, 3)
        with open(self.seed_file, encoding="utf-8") as f:
            data = json.load(f)
        entry = data["entries"][0]
        names = sorted(c["name"] for c in entry["cookies"])
        self.assertEqual(names, ["__cf_bm", "_pxhd", "canary"])
        self.assertNotIn("ESTSAUTH", names)
        # BiDi dict value 归一化成纯字符串
        by_name = {c["name"]: c for c in entry["cookies"]}
        self.assertEqual(by_name["__cf_bm"]["value"], "cf123")

    def test_load_seed_requires_proxy_and_ua_match(self):
        save_seed(_FakePage(), "unittest_pl", PROXY, UA, "T", seed_file=self.seed_file)
        self.assertEqual(len(load_seed("unittest_pl", PROXY, UA, seed_file=self.seed_file)), 3)
        self.assertEqual(load_seed("unittest_pl", PROXY, "UA-OTHER", seed_file=self.seed_file), [])
        self.assertEqual(load_seed("unittest_pl", "http://other:5.5.5.5:80", UA, seed_file=self.seed_file), [])

    def test_load_seed_expired_returns_empty(self):
        save_seed(_FakePage(), "unittest_pl", PROXY, UA, "T", seed_file=self.seed_file)
        with open(self.seed_file, encoding="utf-8") as f:
            data = json.load(f)
        data["entries"][0]["ts"] = time.time() - 99999
        with open(self.seed_file, "w", encoding="utf-8") as f:
            json.dump(data, f)
        self.assertEqual(load_seed("unittest_pl", PROXY, UA, seed_file=self.seed_file), [])

    def test_save_seed_replaces_same_key(self):
        save_seed(_FakePage(), "unittest_pl", PROXY, UA, "T", seed_file=self.seed_file)
        save_seed(_FakePage(), "unittest_pl", PROXY, UA, "T", seed_file=self.seed_file)
        with open(self.seed_file, encoding="utf-8") as f:
            data = json.load(f)
        self.assertEqual(len(data["entries"]), 1)

    def test_warmup_no_nav_when_already_home_and_injects_seed(self):
        page = _FakePage()
        ok = warmup(
            page, "T", home_url="https://login.live.com/",
            seed_cookies=[{"name": "_pxhd", "value": "x", "domain": ".live.com"}],
            min_wait=0.01, max_wait=0.02,
        )
        self.assertTrue(ok)
        self.assertIsNone(page.nav)
        self.assertEqual(len(page.set_called), 1)

    def test_warmup_navigates_when_not_home(self):
        page = _FakePage()
        page.url = "about:blank"
        ok = warmup(page, "T", home_url="https://login.live.com/", min_wait=0.01, max_wait=0.02)
        self.assertTrue(ok)
        self.assertEqual(page.nav, "https://login.live.com/")

    def test_allowlist_has_no_login_cookies(self):
        for forbidden in ("ESTSAUTH", "ESTSAUTHPERSISTENT", "SignInStateCookie",
                          "LP", "PPAuth", "RPSSecAuth"):
            self.assertNotIn(forbidden, CHALLENGE_COOKIE_ALLOWLIST)


if __name__ == "__main__":
    unittest.main()
