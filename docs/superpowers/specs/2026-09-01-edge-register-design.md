# Edge 注册链路（spike → 方案 A）— 设计文档

- 日期: 2026-09-01
- 范围: 参照 ruyi 注册链路（`register_outlook_ruoyi.py` + `common/ruyi/`），新增 Microsoft Edge（Chromium/CDP）驱动的同目标注册版本。两步走：先 spike 验证 Edge 指纹对 PX 的通过率，再建 `common/edge/` 包 + `register_outlook_edge.py` 完整版。
- 状态: 设计已获用户逐节批准（§1~§4），待实现

## 背景与动机

ruyi 注册链路用 ruyipage 定制 Firefox 155（BiDi）跑 signup.live.com 注册。8/28 定位的 blocked 根因（见记忆 `ruyi-outlook-blocked-rootcause`）：

1. **主因**：代理 IP 段信誉烧穿（批量注册源被微软风控标记）
2. **次因**：本机浏览器指纹矛盾（Firefox + zh-CN locale + WebGL 报错卡 + timezone/locale/IP 四方互斥）

换 Edge 跑同一条注册流程，是针对**次因**的指纹品类替换（Firefox → Chromium）。主因（IP 烧穿）不在本设计范围，但 spike 验收标准必须隔离两个变量——在烧穿的 IP 上验证新链路会得出错误结论。

## 调研结论（2026-09-01 实测）

1. **ruyipage 是纯 Firefox 库**（`_pages/` 全是 firefox_*，BiDi 协议），无 Edge/Chromium 支持 → 驱动栈必须换
2. **本机已有 Edge**：`C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe`（152.0.4191.53）
3. **DrissionPage 4.1.1.4** 为 PyPI 最新，未安装（待 spike 时装入 venv）；Chromium 系专用，API 风格与 ruyipage 相近（同步、`ele/eles`、中文文档），迁移成本最低
4. **DrissionPage `set_proxy()` 不支持带账号密码的代理**（Chromium 启动参数限制）；现有代理池全是 `user:pass@host:port` 格式 → 解法：统一走 `common.ruyi.chain.ChainRelayHub` 本地中继（对任意账号放行的无认证 SOCKS5），认证问题一次性解决，且与 ruoyi 代理链行为完全一致
5. **可原样复用**（纯 Python，与浏览器无关）：`ConsumableProxyPool` 代理池、`ChainRelayHub` 中继、IP/代理探测、UA 池骨架、profile 目录管理思路、helpers（`generate_email_password`/`generate_birthday`/`generate_name`/`verify_registered_outlook`/`extract_graph_token_http` Graph HTTP 授权链）
6. **必须重写**（绑定 ruyipage API/浏览器行为）：元素定位/输入封装、signup 各步、PX captcha 按压（`page.actions` → DrissionPage `tab.actions`）、资源拦截
7. **per-tab 代理轮换**是 ruyipage 定制内核的 container-tab 特性，Edge 没有 → 一号一实例一代理

## §1 Spike（第一步，一次性验证脚本）

**目标**：只回答一个问题——真 Edge + DrissionPage + 现有认证代理，PX 挑战能否按压放行、signup 能否走通。

**文件**：`spike_edge_register.py`（项目根，验证完可删或留作参考）

**内容**（~250 行）：
1. 安装 `DrissionPage==4.1.1.4` 进 venv
2. 代理挂载：不折腾插件，上游 `user:pass@host:port` → `ChainRelayHub` → `socks5://127.0.0.1:{port}`（无认证），Edge `--proxy-server` 直接吃。同时解掉"Chromium 不支持认证代理"与"代理链"两个问题
3. Edge 启动：`ChromiumOptions().set_browser_path(...)` + `set_user_data_path(tempfile.mkdtemp())` + `set_proxy(中继)` + 无头可选
4. 反检测最小集：`--disable-blink-features=AutomationControlled`、UA 从 msedge.exe 文件版本号动态生成（不硬编码）
5. 流程：开 signup.live.com → 填邮箱/密码/生日/姓名（选择器抄 ruoyi 的 `_email_input_selector` 等）→ PX 用 `tab.actions.move_to(...).hold()` 按压 → 输出放行与否
6. 判定输出：`RESULT: PASSED_PX / BLOCKED / FORM_FAIL / PROXY_FAIL` + 截图

**验收标准**：在一条**未烧穿**的代理 IP 上跑 3 次，≥2 次过 PX。3 次全 blocked → 停，先换 IP 池，方案 A 暂缓。`FORM_FAIL`/`PROXY_FAIL` 不算 PX 数据，只有 `PASSED_PX`/`BLOCKED` 计入。

**不做**：并发、批跑、WebUI、Graph 授权。

## §2 方案 A 完整架构

```
common/edge/                  # 自包含：标准库 + DrissionPage + 可选 import common.ruyi 纯函数
├── __init__.py               # 公开 API（对齐 common/ruyi/__init__.py 风格）
├── _state.py                 # EDGE_PATH 缓存、profile root、全局锁
├── browser.py                # Edge 路径解析(注册表+默认路径 glob) + 进程清理(taskkill 按路径/profile)
├── options.py                # ChromiumOptions 构建：UA/反检测参数/无头/窗口/代理(经中继)
├── proxy.py                  # 代理→中继映射：user:pass@upstream → socks5://127.0.0.1:{port}
│                             #   (复用 common.ruyi.chain 的 ChainRelayHub，零改动)
├── stealth.py                # CDP 反检测 patch JS（对齐 launch.py 的 _build_headless_patch_js，
│                             #   加 Chromium 特有：webdriver/chrome.runtime/plugins/hardwareConcurrency=16）
├── intercept.py              # 资源拦截：启动参数为主(--blink-settings=imageEnabled=false 等)，
│                             #   黑名单/扩展名常量直接 import common.ruyi.launch 的（单一来源）
├── actions.py                # 动作链映射：ruoyi 的 hold 语义 → tab.actions
│                             #   (move_to 绝对坐标 → move(offset)/hold()/release()，按压轨迹算法照搬)
└── profile.py                # tempfile.mkdtemp per-run profile + 清理

register_outlook_edge.py      # 业务流（预估 ~1800 行）
├── re-export 层              # helpers/代理池/探测/Graph 授权 全部 from register_outlook_ruoyi import
├── 表单流                    # _fill_email/_fill_password/_fill_birthday/_fill_name_and_terms
│                             #   移植：选择器原样保留(css: 前缀 DrissionPage 兼容)
├── PX captcha                # _find_hold_context/_perform_hold 移植，ctx.actions → tab.actions
├── register_outlook_edge()   # 单号主函数，签名/返回四元组与 register_outlook() 一致
└── main()/批跑               # _run_one_direct/_run_direct_batch 结构照搬（asyncio + to_thread）
```

**关键设计决策**：

1. **依赖方向**：`common/edge` → `common/ruyi`（只 import 纯函数：chain/_utils/launch 常量），反向禁止。`register_outlook_edge` → 两者 + `register_outlook_ruoyi`（helpers 复用）
2. **helpers 复用方式**：直接 `import register_outlook_ruoyi as rr` 后 `rr._load_helpers()`——不复制代码。ruoyi 文件是纯 import 安全的（`main()` 有 `if __name__` 保护）
3. **代理统一走中继**：所有上游代理（带不带账号密码）经 `ChainRelayHub` 映射成本地无认证 SOCKS5 喂给 Edge。与 ruoyi 代理链行为一致，`front_proxy` 天然支持
4. **一号一实例**：无 per-tab 轮换。`--concurrency N` = N 个独立 Edge 进程，各挂各的中继端口
5. **环境变量**：新业务 env 用 `EDGE_*` 前缀读但默认值与 ruoyi 相同；代理池/前置直接复用 `OUTLOOK_PROXY_FILE`/`LAUNCH_FRONT_PROXY`（同一池子同一前置，不引入第二份配置）
6. **WebUI**：`webui/scripts.py` 加 `register_outlook_edge` 条目（照抄 ruoyi 条目改 id/file/标题）

**错误处理**：`fail_reason` 语义与 ruoyi 完全一致（`blocked`/`failure`/`exception_*`/`success`），WebUI 日志解析行（`结果:`/`授权结果:`）格式不动，下游 `outlook_reg_loop` 无需感知差异。

## §3 数据流与生命周期（单号视角）

```
main() / _run_direct_batch (asyncio)
  │  ConsumableProxyPool.start() → select_proxy_for_account(idx) → selected_proxy
  │  _probe_proxy_before_browser(selected) ──失败──→ release_proxy, FAIL(proxy_precheck_failed)
  ▼
asyncio.to_thread(register_outlook_edge, run_args, selected_pool, idx)
  │
  │  1. ensure_front_relay(front_proxy)          # 进程级单例，幂等
  │  2. relay_url = ChainRelayHub.port_for(upstream) → socks5://127.0.0.1:{port}
  │  3. profile_dir = tempfile.mkdtemp(prefix="edge_tmp_")   # per-run 隔离，同 ruoyi
  │  4. co = ChromiumOptions:
  │       set_browser_path(resolve_edge_path())   # 注册表/glob 解析,缓存
  │       set_user_data_path(profile_dir)
  │       set_proxy(relay_url)
  │       set_argument('--disable-blink-features=AutomationControlled')
  │       headless? + 窗口尺寸 + UA(msedge 版本对齐真内核)
  │  5. browser = Chromium(co); tab = browser.latest_tab
  │  6. tab.run_js(stealth patch)  ← 首页 landing 后
  │  7. tab.get(SIGNUP_URL) → wait 装载
  │  8. 表单流: email → password → birthday → name   # 每步失败截图+fail_reason
  │  9. PX captcha 循环: _find_hold_target(主文档+iframe) → _perform_hold(tab.actions)
  │       放行 → post_signup 验证 → verify_registered_outlook()
  │       不放行 → max_press 次后 FAIL
  ▼
返回四元组 (email, password, fail_reason, px_metrics)
  ▼
_finish_direct_graph_auth()          # 原样复用 ruoyi 的：Graph HTTP 授权 + 落盘
  ▼
finally: browser.quit() + taskkill 残留 + rmtree(profile_dir) + release_proxy
```

**与 ruoyi 的 API 差异映射**：

| 维度 | ruoyi (Firefox/BiDi) | edge (Chromium/CDP) |
|---|---|---|
| 浏览器实例 | `FirefoxPage(tb)` | `Chromium(co)` → `latest_tab` |
| 元素查找 | `page.ele('css:...', timeout=N)` | `tab.ele('css:...', timeout=N)` 同款语法 |
| JS 注入 | `run_js_loaded(IIFE)` | `tab.run_js(...)`（包 IIFE 兼容） |
| iframe 遍历 | `page.get_all_frames()` | DrissionPage iframe 伪元素直接下钻 |
| 按压 | `actions.move_to({x,y}, duration)` `hold()` `release()` | `actions.move_to((x,y), duration)` `hold()` `release()`（坐标元组格式差异，轨迹算法零改动） |
| 资源拦截 | `page.intercept.start_requests(handler)` | 启动参数 `--blink-settings=imageEnabled=false` 等为主 |
| 资源释放 | `quit(timeout, force)` + powershell 按 profile 杀 | `browser.quit()` + `taskkill /F` 按启动路径过滤（对齐 firefox.py 做法） |

**资源清理铁律**：finally 里 quit → 按路径杀残留进程（等进程真消失，`--user-data-dir` 文件锁与 Firefox XPCOM 锁同类）→ 删 profile。

**超时体系**：`REGISTER_TIMEOUT=300` / `NO_PROGRESS_TIMEOUT` / `SUBMIT_STUCK_NOCHANGE_TIMEOUT` / `max_press` 全部沿用同名常量同语义，`EDGE_` 前缀 env 读但默认值相同。

## §4 验收边界与测试计划

**A 版验收**：
1. `python register_outlook_edge.py -n 1` 单号直跑：注册成功 → email_reg/live_file 落盘格式与 ruoyi 完全一致
2. `--concurrency 2 --launch-stagger 10`：两 Edge 实例并存，profile/端口/代理互不串
3. `--skip-graph-auth` 与默认授权两条路径都通
4. WebUI 里 edge 条目可启动、日志行可解析
5. ruoyi 链路回归：现有 `tests/test_ruoyi_*` 全绿（证明没碰坏共用层）

**测试范围**（纯逻辑，不依赖浏览器/网络，对齐现有 test_ruoyi_* 风格）：
- `tests/test_edge_paths.py`：Edge 路径解析优先级（env 显式 > 注册表 > 默认 glob > 报错）、缓存
- `tests/test_edge_proxy_relay.py`：user:pass@upstream → relay URL 映射、无 front 行为、非法代理抛错
- `tests/test_edge_stealth.py`：patch JS 含关键 defineProperty（webdriver/hardwareConcurrency=16/UA）
- `tests/test_edge_reexport.py`：edge 文件 import 后 helpers/代理池符号与 ruoyi 同源（`is` 断言）
- `tests/test_edge_form_selectors.py`：选择器常量与 ruoyi 原版逐字一致（防移植走样）

**明确不做**（YAGNI）：
- per-tab 代理轮换（Edge 无此特性）
- warmup/种子池（PX seed cookie 是 Firefox 会话实测优化，Edge 品类需重新实测——spike 数据出来后再议）
- unlock/launch 的 edge 版（注册跑通后自然延伸）
- HAR 采集（CDP 侧重做性价比低，需要时单开）

**风险与回退**：
- DrissionPage 4.1.1.4 与 Edge 152 兼容性：spike 第一步就验证；不行降级 Edge 或换 DrissionPage 4.0 稳定线
- 反检测 patch 与真实指纹冲突（伪装 UA 版本 ≠ 内核真实版本）：UA 从 msedge.exe 文件版本号动态生成
- helpers 复用引入 9600 行 import 开销：实测 <1s 可接受；超预期再把 helpers 独立成模块（不在本次做）

**提交切分**：
1. spike 脚本 + pip 依赖（一个 commit）
2. `common/edge/` 包（一个 commit）
3. `register_outlook_edge.py` + tests（一个 commit）
4. WebUI 条目（一个 commit）
