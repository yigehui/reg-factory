import unittest
from types import SimpleNamespace

import outlook_reg_loop as loop


class OutlookRegLoopRuoyiConcurrencyTests(unittest.TestCase):
    def test_one_attempt_ruoyi_passes_concurrency_into_register_opts(self):
        seen = {}

        def fake_register(opts, _proxy_pool, _idx):
            seen["concurrency"] = opts.concurrency
            return "a", "b", "", {}

        fake_mod = SimpleNamespace(
            register_outlook=fake_register,
            PROXY_FILE="proxies.txt",
        )

        loop._one_attempt_ruoyi(
            fake_mod,
            proxy_file="proxies.txt",
            proxy_source="file",
            aimili_url="",
            aimili_token="",
            idx=7,
            timeout=120,
            max_press=3,
            confirm_before_register=False,
            headless=True,
            concurrency=10,
            consumable_pool=None,
        )

        self.assertEqual(seen["concurrency"], 10)

    def test_one_attempt_ruoyi_releases_selected_proxy(self):
        released = []

        fake_pool = SimpleNamespace(
            remaining=lambda: 0,
            take=lambda: ["10.0.0.1:1080:user:pwd"],
        )

        fake_mod = SimpleNamespace(
            register_outlook=lambda opts, _proxy_pool, _idx: ("a", "b", "", {}),
            release_proxy_for_account=lambda proxy, runtime=None: released.append((proxy, runtime)),
            mask_ruoyi_proxy=lambda proxy: proxy,
            PROXY_FILE="proxies.txt",
        )

        loop._one_attempt_ruoyi(
            fake_mod,
            proxy_file="proxies.txt",
            proxy_source="file",
            aimili_url="",
            aimili_token="",
            idx=1,
            timeout=120,
            max_press=3,
            confirm_before_register=False,
            headless=True,
            concurrency=2,
            consumable_pool=fake_pool,
        )

        self.assertEqual(released, [("10.0.0.1:1080:user:pwd", fake_pool)])


if __name__ == "__main__":
    unittest.main()
