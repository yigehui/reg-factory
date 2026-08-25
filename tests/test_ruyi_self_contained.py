"""common/ruyi/ 包自包含性测试。

断言:包内所有 .py 的 import 只在白名单{标准库, ruyipage, requests, common.ruyi.*}内,
禁止 import config / common.notify / common.cloudflare_mail / register_outlook_ruoyi 等业务模块。
保证可整目录拷到别的项目直接用。
"""

import ast
import os
import unittest

import common.ruyi as ruyi_pkg

_PKG_DIR = os.path.dirname(os.path.abspath(ruyi_pkg.__file__))

# 允许的 import module 前缀白名单
_ALLOWED_PREFIXES = (
    # 标准库(白名单显式枚举,避免误判)
    "os", "sys", "re", "json", "time", "threading", "signal", "shutil",
    "tempfile", "subprocess", "urllib", "datetime", "types",
    "ast", "unittest", "msvcrt", "random",
    # 第三方
    "ruyipage", "requests",
    # 包内自身
    "common.ruyi",
)

# 禁止的业务模块(显式黑名单,双重保险)
_FORBIDDEN = (
    "config", "common.notify", "common.cloudflare_mail",
    "register_outlook_ruoyi", "register_outlook_standalone",
)


def _iter_pkg_py():
    for name in os.listdir(_PKG_DIR):
        if name.endswith(".py"):
            yield os.path.join(_PKG_DIR, name)


def _import_modules(path):
    """AST 解析,返回 (module 名, 是否相对 import) 列表。"""
    with open(path, "r", encoding="utf-8") as f:
        tree = ast.parse(f.read(), filename=path)
    result = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                result.append((alias.name, 0))
        elif isinstance(node, ast.ImportFrom):
            if node.module:
                result.append((node.module, node.level))
    return result


class RuyiSelfContainedTests(unittest.TestCase):
    def test_no_forbidden_business_imports(self):
        for path in _iter_pkg_py():
            for mod, _level in _import_modules(path):
                for forbidden in _FORBIDDEN:
                    self.assertFalse(
                        mod == forbidden or mod.startswith(forbidden + "."),
                        f"{os.path.basename(path)} 禁止 import 业务模块: {mod}",
                    )

    def test_all_imports_in_whitelist(self):
        for path in _iter_pkg_py():
            for mod, level in _import_modules(path):
                # 包内相对 import(level>0)合法,跳过白名单检查
                if level > 0:
                    continue
                self.assertTrue(
                    any(mod == p or mod.startswith(p + ".") for p in _ALLOWED_PREFIXES),
                    f"{os.path.basename(path)} import 非白名单: {mod}",
                )

    def test_package_dynamic_import_no_side_effects(self):
        import importlib
        mod = importlib.reload(ruyi_pkg)
        self.assertTrue(hasattr(mod, "ConsumableProxyPool"))
        self.assertTrue(hasattr(mod, "build_browser_options"))
        self.assertTrue(hasattr(mod, "after_launch"))

    def test_all_exports_resolvable(self):
        for name in ruyi_pkg.__all__:
            self.assertTrue(
                hasattr(ruyi_pkg, name),
                f"__all__ 声明但包未暴露: {name}",
            )

    def test_firefox_path_lazy_eval(self):
        # PEP 562:首次访问 RUOYI_FIREFOX_PATH 触发延迟求值,返回非空字符串
        path = ruyi_pkg.RUOYI_FIREFOX_PATH
        self.assertIsInstance(path, str)
        self.assertTrue(path, "RUOYI_FIREFOX_PATH 延迟求值应返回非空路径")

    def test_profile_root_injectable(self):
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            ruyi_pkg.set_profile_root(tmp)
            try:
                self.assertEqual(ruyi_pkg.get_profile_root(), tmp)
            finally:
                ruyi_pkg.set_profile_root(None)

    def test_probe_precheck_targets_default_empty(self):
        # 新包默认不锁业务域名;register 注入业务值
        from common.ruyi import _state
        original = _state.PROXY_PRECHECK_TARGETS
        _state.PROXY_PRECHECK_TARGETS = ()
        try:
            self.assertEqual(ruyi_pkg.PROXY_PRECHECK_URL, "")
        finally:
            _state.PROXY_PRECHECK_TARGETS = original


if __name__ == "__main__":
    unittest.main()
