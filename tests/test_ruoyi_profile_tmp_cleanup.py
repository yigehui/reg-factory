import os
import tempfile
import time
import unittest
from types import SimpleNamespace
from unittest.mock import patch

import register_outlook_ruoyi as mod
from common import ruyi as _ruyi_pkg


class RuoyiProfileTmpCleanupTests(unittest.TestCase):
    def test_ruoyi_profile_dir_uses_configured_tmp_root(self):
        with tempfile.TemporaryDirectory() as tmp_root:
            with patch.object(mod, "RUOYI_PROFILE_ROOT", tmp_root):
                path = mod._ruoyi_profile_dir(SimpleNamespace(), 9)
                try:
                    self.assertTrue(path.startswith(tmp_root))
                    self.assertTrue(os.path.isdir(path))
                    self.assertTrue(os.path.basename(path).startswith("run_0009_"))
                finally:
                    mod._cleanup_ruoyi_run_profile_dir(path, tmp_root)

    def test_cleanup_stale_ruoyi_profile_root_only_removes_old_run_dirs(self):
        with tempfile.TemporaryDirectory() as tmp_root:
            old_run = os.path.join(tmp_root, "run_old")
            fresh_run = os.path.join(tmp_root, "run_fresh")
            keep_dir = os.path.join(tmp_root, "slot_01")
            os.makedirs(old_run)
            os.makedirs(fresh_run)
            os.makedirs(keep_dir)
            old_slot_run = os.path.join(keep_dir, "run_old")
            os.makedirs(old_slot_run)
            old = time.time() - 3600
            os.utime(old_run, (old, old))
            os.utime(old_slot_run, (old, old))

            cleaned = mod._cleanup_stale_ruoyi_profile_root(tmp_root, max_age_sec=10)

            self.assertEqual(2, cleaned)
            self.assertFalse(os.path.exists(old_run))
            self.assertFalse(os.path.exists(old_slot_run))
            self.assertTrue(os.path.isdir(fresh_run))
            self.assertTrue(os.path.isdir(keep_dir))

    def test_cleanup_ruoyi_run_profile_dir_removes_current_run_dir(self):
        with tempfile.TemporaryDirectory() as tmp_root:
            run_dir = os.path.join(tmp_root, "run_0001_test")
            os.makedirs(run_dir)

            self.assertTrue(mod._cleanup_ruoyi_run_profile_dir(run_dir, tmp_root))
            self.assertFalse(os.path.exists(run_dir))

    def test_cleanup_stale_ruoyi_profile_root_skips_active_run_dirs(self):
        with tempfile.TemporaryDirectory() as tmp_root:
            with patch.object(mod, "RUOYI_PROFILE_ROOT", tmp_root):
                run_dir = mod._ruoyi_profile_dir(SimpleNamespace(ruoyi_slot=1), 1)
                old = time.time() - 3600
                os.utime(run_dir, (old, old))

                self.assertEqual(0, mod._cleanup_stale_ruoyi_profile_root(tmp_root, max_age_sec=10))
                self.assertTrue(os.path.isdir(run_dir))
                self.assertTrue(mod._cleanup_ruoyi_run_profile_dir(run_dir, tmp_root))

    def test_cleanup_ruoyi_run_profile_dir_retries_then_succeeds(self):
        with tempfile.TemporaryDirectory() as tmp_root:
            run_dir = os.path.join(tmp_root, "run_0001_test")
            os.makedirs(run_dir)
            calls = []

            def fake_rmtree(path):
                calls.append(path)
                if len(calls) == 1:
                    raise PermissionError("locked")
                return None

            with (
                patch.object(_ruyi_pkg._state, "RUOYI_PROFILE_CLEANUP_RETRIES", 2),
                patch.object(_ruyi_pkg._state, "RUOYI_PROFILE_CLEANUP_RETRY_DELAY", 0.0),
                patch.object(_ruyi_pkg.profile.shutil, "rmtree", side_effect=fake_rmtree),
            ):
                self.assertTrue(mod._cleanup_ruoyi_run_profile_dir(run_dir, tmp_root))

            self.assertEqual([run_dir, run_dir], calls)

    def test_quit_browser_page_uses_force_quit(self):
        seen = {}

        class FakePage:
            def quit(self, timeout=None, force=False):
                seen["timeout"] = timeout
                seen["force"] = force

            def close(self):
                seen["closed"] = True

        self.assertTrue(mod._quit_browser_page(FakePage(), tag="[#1][ruoyi]", timeout=0.01))
        self.assertTrue(seen["timeout"] >= 2.0)
        self.assertTrue(seen["force"])
        self.assertNotIn("closed", seen)

    def test_unlock_reused_cleanup_helpers_exist(self):
        """unlock_outlook.close_firefox 复用 register 的清理能力销毁临时 profile,
        确保 unlock 引用的符号都存在(防重构改名单边失配导致解锁不再清理 -> 磁盘跑满)。"""
        for name in (
            "_kill_ruoyi_firefox_by_profile",
            "_cleanup_ruoyi_run_profile_dir",
            "RUOYI_FIREFOX_EXIT_WAIT",
        ):
            self.assertTrue(hasattr(mod, name), f"register 缺 unlock 复用的符号: {name}")
        self.assertTrue(callable(mod._kill_ruoyi_firefox_by_profile))
        self.assertTrue(callable(mod._cleanup_ruoyi_run_profile_dir))
        self.assertGreater(float(mod.RUOYI_FIREFOX_EXIT_WAIT), 0.0)

    def test_unlock_close_firefox_cleanup_sequence_removes_profile_dir(self):
        """模拟 unlock close_firefox 的清理序列(精准杀进程树 -> rmtree 临时 profile):
        在没真实 firefox 的测试环境直接走 _cleanup_ruoyi_run_profile_dir 验证目录被删。
        对齐 register 10323-10343 的销毁后清理路径。"""
        with tempfile.TemporaryDirectory() as tmp_root:
            with patch.object(mod, "RUOYI_PROFILE_ROOT", tmp_root):
                profile_dir = mod._ruoyi_profile_dir(SimpleNamespace(), 7)
                self.assertTrue(os.path.isdir(profile_dir))

                # unlock close_firefox 等价路径:进程树清理在无 firefox 时 no-op,随后 rmtree
                mod._kill_ruoyi_firefox_by_profile(profile_dir, timeout=0.01)
                removed = mod._cleanup_ruoyi_run_profile_dir(profile_dir, tmp_root)

                self.assertTrue(removed)
                self.assertFalse(os.path.exists(profile_dir))



if __name__ == "__main__":
    unittest.main()
