# 升级 ruyipage 库 + Firefox 155 内核 — 设计文档

- 日期: 2026-08-22
- 范围: 把项目依赖的 `ruyipage` 库从 `1.2.54` 升到 `1.2.62`(PyPI 最新),并用 `python -m ruyipage install` 把配套 Firefox 内核从 `151-ruyi`(`firefox-151.0a1`)换到 `v1.2.58` 发布的 `155` 内核(`firefox-155.0a1`,子目录 `firefox-155.0a1-v1.2.58-win64`);顺带修 `register_outlook_ruoyi.py` 里指向旧内核的错误默认路径。
- 方案: 方案 A(库 + 内核一次性全量升级),不锁 `requirements.txt` 版本,验证只跑纯逻辑测试。
- 状态: 待用户终审

## 背景

现状:

- `requirements.txt` 写的是无版本锁的 `ruyipage`;venv 实装 `1.2.54`。
- 1.2.54 的 `_runtime/manifest.py` 钉内核为 `RELEASE_TAG="151-ruyi"` / `FIREFOX_VERSION="151.0a1"` / `install_subdir="firefox-151.0a1-151-ruyi-win64"`;磁盘 `%LOCALAPPDATA%\ruyipage\browsers\firefox-151.0a1-151-ruyi-win64` 即此。
- 1.2.62(即 GitHub `main` 头)的 `manifest.py` 钉内核为 `RELEASE_TAG="v1.2.58"` / `FIREFOX_VERSION="155.0a1"` / `install_subdir="firefox-155.0a1-v1.2.58-win64"` / `asset="firefox-155.0a1.en-US.win64-20260803.zip"`。
- `register_outlook_ruoyi.py:88` 的 `DEFAULT_RUOYI_FIREFOX` 硬编码为 `C:\Users\Administrator\AppData\Local\ruyipage\browsers\firefox-151.0a1-151-ruyi-win64\firefox\firefox.exe` —— 用户名是 `Administrator`(本机是 `yigehui`),子目录是旧 151。此默认值在本机**本就指向不存在的路径**,靠环境变量 `RUOYI_FIREFOX_PATH` 覆盖才工作;升级到 155 后子目录名也会变,错误雪上加霜。

1.2.54 → 1.2.62 八个版本的提交主题(GitHub `LoseNine/ruyipage`):

- 1.2.57: BiDi 预加载脚本与 Developer Tools 检测解耦;经 BiDi 元数据读受保护上下文 URL(为 Firefox 155 适配)。
- 1.2.58: 普通启动会话落到脚本可访问的 about:blank,让 Firefox 155 能立即读页面属性;`private=True` 保留原生 about:privatebrowsing 落地页。
- 1.2.59/1.2.60: 智能指纹运行时配置对齐;异步参数在同步边界解包(#27,影响 Actions/TouchActions/drag/run_js)。
- 1.2.61/1.2.62: 受管 Firefox 升到 release `v1.2.58`(即 155 内核);`ci: run browser tests with managed Firefox 155`;`fix: restore Firefox 155 capture compatibility`。

结论:八版本全是为 Firefox 155 做的兼容/BiDi 修复,无公开 breaking change。库与内核解耦(`set_browser_path` 显式传路径),但 1.2.57+ 的改动专为 155 设计,新库跑旧 151 内核属上游未测组合 → 方案 A 库+内核一起升是上游意图。

## 关键决策(已与用户确认)

1. 方案 A:库 + 155 内核一次性全量升级。
2. `requirements.txt` **不锁版本**(维持 `ruyipage`)。
3. 验证只跑**纯逻辑测试**(不依赖浏览器/网络/代理)。
4. `_tmp_c64_ruoyi.py` 是未被任何代码 import 的陈旧副本,**不在本次范围**(动它纯属无关 churn)。

## 项目实际用到的 ruyipage API 面(兼容性核查)

import 的符号(全项目):

- `from ruyipage import FirefoxOptions, FirefoxPage` —— `launch_ruoyi_browser.py`、`probe_graph_secondary_email.py`、`unlock_outlook_ruoyi.py`、`register_outlook_ruoyi.py`(函数内局部 import)
- `from ruyipage import NoneElement` —— `register_outlook_ruoyi.py:3294`

页面/元素方法(`register_outlook_ruoyi.py` + `unlock_outlook_ruoyi.py` 实测调用计数):

- `page.run_js_loaded` (16)、`page.get` (7)、`.run_js(` (5)、`.move_to(` (4)、`page.quit` (3)、`page.eles` (1)、`page.ele` (1)、`page.close_other_tabs` (1)
- 另:`new_container_tab`(`launch_ruoyi_browser.py`)、`FirefoxOptions.set_browser_path/set_profile/set_proxy/set_per_tab_proxies/headless`、`FirefoxPage(tb)` 构造、`close_other_tabs(page)`、`mask_ruoyi_proxy` 等项目自有封装。

这些核心符号/方法在 1.2.54→1.2.62 之间无重命名/删除(API 兼容风险低)。真正风险面是 151→155 的**指纹基线变化**(canvas/WebGL/UA),那只有用真号实测才能判断,纯逻辑测试覆盖不到 —— 用户已知悉并接受。

## §1 执行步骤

### 1.1 升级库

```
.venv/Scripts/python.exe -m pip install --upgrade ruyipage
```

升级后核对:

```
.venv/Scripts/python.exe -m pip show ruyipage | grep -i version   # 期望 1.2.62
.venv/Scripts/python.exe -c "import ruyipage; print(ruyipage.__version__)"
.venv/Scripts/python.exe -c "from ruyipage import FirefoxOptions, FirefoxPage, NoneElement; print('imports ok')"
```

### 1.2 安装 155 内核

```
.venv/Scripts/python.exe -m ruyipage install
```

预期:下载 `firefox-155.0a1.en-US.win64-20260803.zip`,解压到 `%LOCALAPPDATA%\ruyipage\browsers\firefox-155.0a1-v1.2.58-win64\firefox\firefox.exe`。旧 151 内核目录保留(库按 `install_subdir` 区分,不冲突;保留它便于回退)。

核对:

```
.venv/Scripts/python.exe -m ruyipage path        # 应输出 155 子目录的 firefox.exe
.venv/Scripts/python.exe -m ruyipage install --dry-run --json   # release=v1.2.58 version=155.0a1 installed=true
```

### 1.3 修 `DEFAULT_RUOYI_FIREFOX`

`register_outlook_ruoyi.py:88-94` 现状(硬编码,用户名/子目录双错):

```python
DEFAULT_RUOYI_FIREFOX = (
    r"C:\Users\Administrator\AppData\Local\ruyipage\browsers"
    r"\firefox-151.0a1-151-ruyi-win64\firefox\firefox.exe"
)
RUOYI_FIREFOX_PATH = os.environ.get("RUOYI_FIREFOX_PATH", DEFAULT_RUOYI_FIREFOX)
```

改为**动态解析库自带的 `get_executable_path`**,失败再回退硬编码 155 路径:

```python
def _resolve_default_ruoyi_firefox():
    """优先用 ruyipage 库自带的内核路径解析；失败回退到 155 内核硬编码路径。"""
    try:
        from ruyipage._runtime.installer import get_executable_path
        return get_executable_path(strict=False) or ""
    except Exception:
        return ""

_DEFAULT_RUOYI_FIREFOX_FALLBACK = os.path.join(
    os.environ.get("LOCALAPPDATA", r"C:\Users\yigehui\AppData\Local"),
    r"ruyipage\browsers\firefox-155.0a1-v1.2.58-win64\firefox\firefox.exe",
)

DEFAULT_RUOYI_FIREFOX = _resolve_default_ruoyi_firefox() or _DEFAULT_RUOYI_FIREFOX_FALLBACK
RUOYI_FIREFOX_PATH = os.environ.get("RUOYI_FIREFOX_PATH", DEFAULT_RUOYI_FIREFOX)
```

要点:

- `strict=False` —— 未装内核时返回空串而非抛错,触发回退。
- 用 `LOCALAPPDATA` 环境变量而非硬编码用户名,适配当前机器与未来迁移。
- 回退路径用 155 子目录 `firefox-155.0a1-v1.2.58-win64`(与新 manifest 一致)。
- 环境变量 `RUOYI_FIREFOX_PATH` 仍最高优先级,不破坏现有覆盖用法。

`_browser_model_name`(1502 行)按路径分段解析,版本无关,无需改 —— 升级后会正确打印 `firefox-155.0a1-v1.2.58-win64`。

### 1.4 验证:纯逻辑测试

只跑不依赖浏览器/网络/代理的纯逻辑测试。先看哪些是纯逻辑:

```
.venv/Scripts/python.exe -m pytest test_unlock_ruoyi_classify.py -q
```

(其余 `test_ruoyi_*.py` 多数要真浏览器/代理/账号,本次不跑 —— 用户已确认范围。)

若 `test_unlock_ruoyi_classify.py` 之外还有明显纯逻辑的(如 CLI 参数解析、代理归一化纯函数),可一并跑;需联网/启浏览器的跳过并在结果里标注。

## §2 回退方案

- 库回退:`.venv/Scripts/python.exe -m pip install ruyipage==1.2.54`
- 内核回退:旧 151 目录 `firefox-151.0a1-151-ruyi-win64` 仍在磁盘,设 `RUOYI_FIREFOX_PATH` 指回它即可(库的 `set_browser_path` 接受任意路径,1.2.62 也能跑 151 内核,只是上游未测组合)。
- 路径修复是纯增量,不回退也不影响环境变量覆盖。

## §3 不做的事(YAGNI)

- 不锁 `requirements.txt`(用户决定)。
- 不动 `_tmp_c64_ruoyi.py`(陈旧副本,无人 import)。
- 不跑浏览器冒烟/真号 PX 实测(用户限定纯逻辑测试)。
- 不改 `launch_ruoyi_browser.py` / `unlock_outlook_ruoyi.py` —— 它们从 `register_outlook_ruoyi` 复用 `RUOYI_FIREFOX_PATH`,路径修好后自动跟着对。
- 不重构 `_browser_model_name` 等版本无关代码。
