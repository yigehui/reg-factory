import importlib.util
import unittest
from pathlib import Path
from unittest.mock import patch

import register_outlook_ruoyi as ruoyi
import outlook_reg_loop as loop
import run_full_flow as full


def _load_webui_scripts():
    path = Path(__file__).resolve().parent / 'webui' / 'scripts.py'
    spec = importlib.util.spec_from_file_location('webui_scripts_for_test', path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


class ProdLogLevelTests(unittest.TestCase):
    def _capture(self, mod, level, entries):
        old_level = getattr(mod, 'LOG_LEVEL', None)
        try:
            if hasattr(mod, 'set_log_level'):
                mod.set_log_level(level)
            else:
                mod.LOG_LEVEL = level
            with patch('builtins.print') as mocked:
                for msg, msg_level in entries:
                    mod.log(msg, msg_level)
            return [call.args[0] for call in mocked.call_args_list]
        finally:
            if old_level is not None:
                if hasattr(mod, 'set_log_level'):
                    mod.set_log_level(old_level)
                else:
                    mod.LOG_LEVEL = old_level

    def test_ruoyi_prod_keeps_start_proxy_and_result_lines_with_account_index(self):
        lines = self._capture(
            ruoyi,
            'PROD',
            [
                ('开始: count=1 concurrency=1 launch_stagger=10s ua_pool=6 timeout=150s', 'INFO'),
                ('proxy list ready: source=file+aimili-list size=87', 'INFO'),
                ('========== 注册 #1/1 ==========', 'INFO'),
                ('#1 代理 -> socks5://yigehui...@163.192.58.188:52025', 'INFO'),
                ('  [#1][ruoyi] signup entry DOM step=email', 'INFO'),
                ('  [#1][ruoyi] captcha timeout', 'WARN'),
                ('  [#1][ruoyi] OK: foo@outlook.com / Passw0rd!', 'OK'),
                ('[#1][ruoyi] 授权结果: OK', 'OK'),
                ('[#1][ruoyi] 结果: OK foo@outlook.com total=12.34s', 'OK'),
                ('DONE: success=1/1 fail=0 no_graph=0', 'OK'),
                ('SUMMARY: success 1 | fail 0 | no_graph 0 | total 1', 'OK'),
                ('SUMMARY_TIME: total_elapsed 12.34s | avg_success_elapsed 12.34s', 'OK'),
                ('[#1][ruoyi] register task raised RuntimeError: boom', 'ERR'),
            ],
        )
        rendered = '\n'.join(lines)
        self.assertIn('开始: count=1 concurrency=1 launch_stagger=10s ua_pool=6 timeout=150s', rendered)
        self.assertIn('proxy list ready: source=file+aimili-list size=87', rendered)
        self.assertIn('========== 注册 #1/1 ==========', rendered)
        self.assertIn('#1 代理 -> socks5://yigehui...@163.192.58.188:52025', rendered)
        self.assertNotIn('signup entry DOM step=email', rendered)
        self.assertNotIn('captcha timeout', rendered)
        self.assertIn('OK: foo@outlook.com / Passw0rd!', rendered)
        self.assertIn('授权结果: OK', rendered)
        self.assertIn('结果: OK foo@outlook.com total=12.34s', rendered)
        self.assertIn('DONE: success=1/1 fail=0 no_graph=0', rendered)
        self.assertIn('SUMMARY: success 1 | fail 0 | no_graph 0 | total 1', rendered)
        self.assertIn('SUMMARY_TIME: total_elapsed 12.34s | avg_success_elapsed 12.34s', rendered)
        self.assertIn('register task raised RuntimeError: boom', rendered)

    def test_existing_warn_level_behavior_stays_unchanged(self):
        lines = self._capture(
            ruoyi,
            'WARN',
            [
                ('[#1][ruoyi] 授权结果: OK', 'OK'),
                ('[#1][ruoyi] 结果: OK foo@outlook.com total=12.34s', 'OK'),
                ('[#1][ruoyi] captcha timeout', 'WARN'),
            ],
        )
        rendered = '\n'.join(lines)
        self.assertNotIn('授权结果: OK', rendered)
        self.assertNotIn('结果: OK foo@outlook.com total=12.34s', rendered)
        self.assertIn('captcha timeout', rendered)

    def test_loop_prod_keeps_only_result_style_lines(self):
        lines = self._capture(
            loop,
            'PROD',
            [
                ('=== attempt #1  (pool=0, succ=0, fail=0) ===', 'INFO'),
                ('OK in 12.3s: foo@outlook.com -> rec.json (pool now 1)', 'OK'),
                ('FAIL in 9.0s (success rate 0/1 = 0%)', 'WARN'),
                ('attempt raised RuntimeError: boom', 'WARN'),
                ('write_record FAILED: RuntimeError: boom', 'ERR'),
            ],
        )
        rendered = '\n'.join(lines)
        self.assertNotIn('=== attempt #1', rendered)
        self.assertIn('OK in 12.3s: foo@outlook.com', rendered)
        self.assertIn('FAIL in 9.0s', rendered)
        self.assertNotIn('attempt raised RuntimeError: boom', rendered)
        self.assertIn('write_record FAILED: RuntimeError: boom', rendered)

    def test_run_full_flow_prod_keeps_only_final_summary_and_errors(self):
        lines = self._capture(
            full,
            'PROD',
            [
                ('Stage A 邮箱注册启动；emails.txt 现有 0 个号', 'A'),
                ('本轮结束  email=foo@outlook.com  Stage B exit=0  用时 99s', 'OK'),
                ('全部结束  共 1 轮  成功 1  失败 0  总用时 99s', 'OK'),
                ('Stage A 没拿到可用邮箱，本轮终止', 'ERR'),
            ],
        )
        rendered = '\n'.join(lines)
        self.assertNotIn('Stage A 邮箱注册启动', rendered)
        self.assertIn('本轮结束  email=foo@outlook.com  Stage B exit=0  用时 99s', rendered)
        self.assertIn('全部结束  共 1 轮  成功 1  失败 0  总用时 99s', rendered)
        self.assertIn('Stage A 没拿到可用邮箱，本轮终止', rendered)

    def test_webui_and_cli_choices_expose_prod(self):
        scripts = _load_webui_scripts()
        for sid in ('run_full_flow', 'outlook_reg_loop', 'register_outlook_ruoyi'):
            script = next(item for item in scripts.SCRIPTS if item['id'] == sid)
            field = next(arg for arg in script['args'] if arg.get('flag') == '--log-level')
            self.assertIn('PROD', field['choices'])
        env_item = next(item for group in scripts.ENV_SCHEMA for item in group['items'] if item.get('key') == 'OUTLOOK_LOG_LEVEL')
        self.assertIn('PROD', env_item['choices'])


if __name__ == '__main__':
    unittest.main()
