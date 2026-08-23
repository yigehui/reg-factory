"""pytest 共享配置:把项目根加入 sys.path,让 tests/ 下的测试仍可
`import register_outlook_ruoyi` / `import webui.server` 等根目录模块。

测试从根目录用 `pytest tests/` 运行;根目录的源码不动,测试集中到 tests/。
"""
import os
import sys

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)
