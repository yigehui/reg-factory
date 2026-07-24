# -*- coding: utf-8 -*-
"""
webui/scripts.py — GUI 的数据核心：把每个入口脚本的命令行参数、以及 .env 配置项
结构化成 schema，前端据此自动生成表单。新增脚本/配置项只改这里。

参数类型(type)对应前端控件：
  bool   -> 勾选框(store_true，勾上才追加 --flag)
  int    -> 数字输入
  str    -> 文本输入
  choice -> 下拉(配 choices)
  multi  -> 多选(nargs="+"，选中的值空格分隔追加在 --flag 后)
positional=True 表示位置参数(不带 --，直接拼值)。
"""

# ============================================================ 入口脚本 schema
SCRIPTS = [
    # ---------------------------------------------------------------- 主流程
    {
        "id": "run_full_flow",
        "file": "run_full_flow.py",
        "category": "主流程",
        "title": "端到端全流程",
        "desc": "注册 Outlook 邮箱 → 在所选平台注册账号。最常用入口。",
        "args": [
            {"flag": "--platforms", "type": "multi", "choices": ["claude", "chatgpt", "grok"],
             "default": ["claude"], "help": "要注册的平台(可多选)"},
            {"flag": "--rounds", "type": "int", "default": 1, "help": "循环注册轮数(0=无限循环)"},
            {"flag": "--codex", "type": "bool", "default": False, "help": "chatgpt 注册后走 Codex OAuth 导入 SUB2API(会接码过手机)"},
            {"flag": "--import-c2a", "type": "bool", "default": False, "help": "chatgpt 注册后即时导入 chatgpt2api(需配 CHATGPT2API_*)"},
            {"flag": "--codex-manual-phone", "type": "bool", "default": False, "help": "Codex add-phone 手动模式(自己在浏览器填号收码)"},
            {"flag": "--codex-group", "type": "str", "default": "", "help": "SUB2API 目标分组名(默认取配置)"},
            {"flag": "--skip-email", "type": "bool", "default": False, "help": "跳过邮箱注册，直接用下面指定的邮箱"},
            {"flag": "--email", "type": "str", "default": "", "help": "配合 --skip-email：现成邮箱"},
            {"flag": "--password", "type": "str", "default": "", "help": "配合 --email 的密码"},
            {"flag": "--node", "type": "str", "default": "auto", "help": "claude/grok 走的 Clash 节点"},
            {"flag": "--email-attempts", "type": "int", "default": 30, "help": "邮箱注册最多尝试次数"},
            {"flag": "--outlook-engine", "type": "choice", "choices": ["ruoyi", "standalone"],
             "default": "ruoyi", "help": "Outlook 自注册后端"},
            {"flag": "--outlook-proxy-file", "type": "str", "default": "proxies_outlook.txt",
             "help": "ruoyi 后端代理池文件"},
            {"flag": "--outlook-aimili-url", "type": "str", "default": "",
             "help": "AimiliVPN 列表接口；会和本地代理文件一起加载进同一个 list"},
            {"flag": "--outlook-aimili-token", "type": "str", "default": "",
             "help": "AimiliVPN 代理池 API Token"},
            {"flag": "--outlook-headless", "type": "bool", "default": False,
             "help": "ruoyi 后端：Outlook 注册阶段使用无头模式"},
            {"flag": "--outlook-har", "type": "bool", "default": False,
             "help": "ruoyi 后端：保存 Outlook 注册完整链路 HAR；成功/失败都保存"},
            {"flag": "--log-level", "type": "choice", "default": "INFO",
             "choices": ["DEBUG", "INFO", "WARN", "ERR"],
             "help": "Outlook/Graph 日志等级"},
            {"flag": "--platform-timeout", "type": "int", "default": 600, "help": "平台注册单号超时(秒)"},
            {"flag": "--email-confirm-before-register", "type": "bool", "default": False,
             "help": "邮箱注册页打开后自动点确认，再开始填写"},
            {"flag": "--dry-run", "type": "bool", "default": False, "help": "只打印将执行的命令，不真注册(安全预览)"},
        ],
    },
    {
        "id": "register_three_platforms",
        "file": "register_three_platforms.py",
        "category": "主流程",
        "title": "三平台注册(已有邮箱)",
        "desc": "用现成邮箱(或从 emails.txt 池)在 Claude/ChatGPT/Grok 注册。",
        "args": [
            {"flag": "--from-pool", "type": "bool", "default": False, "help": "从 emails.txt 池取一个邮箱"},
            {"flag": "--email", "type": "str", "default": "", "help": "指定邮箱(不从池取时)"},
            {"flag": "--password", "type": "str", "default": "", "help": "邮箱密码"},
            {"flag": "--token", "type": "str", "default": "", "help": "Outlook refresh_token(走 Graph 取码)"},
            {"flag": "--client-id", "type": "str", "default": "", "help": "Outlook OAuth client_id"},
            {"flag": "--platforms", "type": "multi", "choices": ["claude", "chatgpt", "grok"],
             "default": ["claude", "chatgpt", "grok"], "help": "要注册的平台"},
            {"flag": "--parallel", "type": "bool", "default": False, "help": "并行跑各平台(默认顺序)"},
            {"flag": "--loop", "type": "bool", "default": False, "help": "持续从池取号循环注册(常驻)"},
            {"flag": "--codex", "type": "bool", "default": False, "help": "chatgpt 后走 Codex OAuth"},
            {"flag": "--import-c2a", "type": "bool", "default": False, "help": "chatgpt 后导入 chatgpt2api"},
            {"flag": "--node", "type": "str", "default": "auto", "help": "Grok Clash 节点"},
            {"flag": "--timeout", "type": "int", "default": 600, "help": "单平台超时(秒)"},
        ],
    },
    {
        "id": "oauth_codex",
        "file": "oauth_codex.py",
        "category": "主流程",
        "title": "Codex OAuth 授权 → SUB2API",
        "desc": "用已存 cookie 重登 ChatGPT，走 OAuth 拿带 refresh_token 的凭据建到 SUB2API/CPA。默认直接接码过 add-phone。",
        "args": [
            {"flag": "--cookie", "type": "str", "default": "", "help": "cookie 文件路径(默认最新 full_*.json)"},
            {"flag": "--phone-skip", "type": "int", "default": 0, "help": "先赌免手机直连次数(0=直接接码不赌)"},
            {"flag": "--manual-phone", "type": "bool", "default": False, "help": "手动模式：自己在浏览器填号+输码"},
            {"flag": "--phone", "type": "str", "default": "", "help": "半自动：脚本填该号(E.164)+选WhatsApp，你只手输码"},
            {"flag": "--group", "type": "str", "default": "", "help": "SUB2API 目标分组(默认取配置)"},
            {"flag": "--skip-cpa", "type": "bool", "default": False, "help": "不推 CPA(默认配好就推)"},
            {"flag": "--keep", "type": "bool", "default": False, "help": "失败保留窗口便于排查"},
        ],
    },
    # ---------------------------------------------------------------- 单平台注册
    {
        "id": "register_chatgpt",
        "file": "register_chatgpt.py",
        "category": "单平台注册",
        "title": "ChatGPT 注册",
        "desc": "ChatGPT 单平台注册(可绕过邮箱池指定邮箱)。",
        "args": [
            {"flag": "--count", "type": "int", "default": 1, "help": "注册数量"},
            {"flag": "--concurrency", "type": "int", "default": 1, "help": "并发数"},
            {"flag": "--timeout", "type": "int", "default": 480, "help": "单号超时(秒)"},
            {"flag": "--email", "type": "str", "default": "", "help": "指定邮箱(绕过池)"},
            {"flag": "--password", "type": "str", "default": "", "help": "邮箱密码"},
            {"flag": "--refresh-token", "type": "str", "default": "", "help": "Outlook refresh_token"},
            {"flag": "--client-id", "type": "str", "default": "", "help": "Outlook OAuth client_id"},
            {"flag": "--import-c2a", "type": "bool", "default": False, "help": "注册后导入 chatgpt2api"},
            {"flag": "--codex", "type": "bool", "default": False, "help": "注册后走 Codex OAuth"},
            {"flag": "--codex-manual-phone", "type": "bool", "default": False, "help": "Codex 手动填号收码"},
            {"flag": "--keep-on-fail", "type": "bool", "default": False, "help": "失败保留窗口"},
        ],
    },
    {
        "id": "register_grok",
        "file": "register_grok.py",
        "category": "单平台注册",
        "title": "Grok 注册",
        "desc": "Grok 单平台注册(需能过 Cloudflare 的干净节点)。",
        "args": [
            {"flag": "--count", "type": "int", "default": 1, "help": "注册数量"},
            {"flag": "--concurrency", "type": "int", "default": 1, "help": "并发数"},
            {"flag": "--timeout", "type": "int", "default": 600, "help": "单号超时(秒)"},
            {"flag": "--node", "type": "str", "default": "auto", "help": "Clash 出口节点(过 grok CF)"},
            {"flag": "--email", "type": "str", "default": "", "help": "指定邮箱(绕过池)"},
            {"flag": "--password", "type": "str", "default": "", "help": "邮箱密码"},
            {"flag": "--keep-on-fail", "type": "bool", "default": False, "help": "失败保留窗口"},
        ],
    },
    {
        "id": "register_claude",
        "file": "register.py",
        "category": "单平台注册",
        "title": "Claude 注册",
        "desc": "Claude 单平台注册(claude.ai 区域封锁需走干净节点)。",
        "args": [
            {"flag": "--count", "type": "int", "default": 1, "help": "注册数量"},
            {"flag": "--concurrency", "type": "int", "default": 1, "help": "并发数"},
            {"flag": "--timeout", "type": "int", "default": 480, "help": "单号超时(秒)"},
            {"flag": "--email", "type": "str", "default": "", "help": "指定邮箱(调试)"},
            {"flag": "--password", "type": "str", "default": "", "help": "邮箱密码"},
            {"flag": "--token", "type": "str", "default": "", "help": "refresh token"},
            {"flag": "--node", "type": "str", "default": "none", "help": "Clash 节点(none=不切)"},
        ],
    },
    {
        "id": "register_github",
        "file": "register_github.py",
        "category": "单平台注册",
        "title": "GitHub 注册",
        "desc": "GitHub 注册(含 Arkose 验证视觉求解，需配 VISION_*/VOTE_*)。",
        "args": [
            {"flag": "--auto", "type": "bool", "default": False, "help": "走完整流程(含取 launch code)"},
            {"flag": "--email", "type": "str", "default": "", "help": "指定邮箱(默认从 _outlook_pool 取)"},
            {"flag": "--password", "type": "str", "default": "", "help": "邮箱密码"},
            {"flag": "--no-keep", "type": "bool", "default": False, "help": "结束后删窗口(默认保留)"},
            {"flag": "--timeout", "type": "int", "default": 600, "help": "超时(秒)"},
        ],
    },
    # ---------------------------------------------------------------- 养号 / 邮箱
    {
        "id": "outlook_reg_loop",
        "file": "outlook_reg_loop.py",
        "category": "养号/邮箱",
        "title": "Outlook 自注册养号",
        "desc": "使用 standalone 后端持续自注册 Outlook，产出到 _outlook_pool/ 与 emails.txt。count=0 为无限循环。",
        "fixed_args": ["--engine", "standalone"],
        "args": [
            {"flag": "--count", "type": "int", "default": 0, "help": "注册数量(0=无限循环)"},
            {"flag": "--target-pool", "type": "int", "default": 0, "help": "池达到此数量就停(0=不限)"},
            {"flag": "--proxy-file", "type": "str", "default": "proxies_outlook.txt",
             "help": "BitBrowser 账号代理池；每次注册随机取一条，留空则读 OUTLOOK_PROXIES"},
            {"flag": "--timeout", "type": "int", "default": 180, "help": "单号注册超时(秒)"},
            {"flag": "--log-level", "type": "choice", "default": "INFO",
             "choices": ["DEBUG", "INFO", "WARN", "ERR"],
             "help": "Outlook/Graph 日志等级"},
            {"flag": "--sleep", "type": "int", "default": 5, "help": "每次注册间隔(秒)"},
            {"flag": "--sleep-when-full", "type": "int", "default": 60, "help": "池达到 target-pool 后每轮等待秒数"},
        ],
    },
    {
        "id": "register_outlook_ruoyi",
        "file": "register_outlook_ruoyi.py",
        "category": "养号/邮箱",
        "title": "Outlook 自注册(ruoyi)",
        "desc": "ruyipage Firefox BiDi 版 Outlook 自注册；支持 SOCKS5 代理池、验证码按住、Graph token 落盘。",
        "args": [
            {"flag": "--count", "type": "int", "default": 1, "help": "注册数量"},
            {"flag": "--concurrency", "type": "int", "default": 1, "help": "并发注册数(建议 4 以内)"},
            {"flag": "--launch-stagger", "type": "float", "default": 10,
             "help": "并发启动错峰秒数；4 并发建议 10；0=关闭"},
            {"flag": "--ua-pool", "type": "str", "default": "",
             "help": "UA 池，用 | 分隔多条 Firefox UA；空=内置 6 条轮询"},
            {"flag": "--aimili-url", "type": "str", "default": "",
             "help": "AimiliVPN 列表接口；会和本地代理文件一起加载进同一个 list"},
            {"flag": "--aimili-token", "type": "str", "default": "",
             "help": "AimiliVPN 代理池 API Token"},
            {"flag": "--proxy-file", "type": "str", "default": "proxies_outlook.txt", "help": "代理池文件(每行一个 user:pass@host:port)"},
            {"flag": "--timeout", "type": "int", "default": 300, "help": "单号超时(秒)"},
            {"flag": "--max-press", "type": "str", "default": "5", "help": "验证码按住次数上限"},
            {"flag": "--headless", "type": "bool", "default": False, "help": "无头模式"},
            {"flag": "--log-level", "type": "choice", "default": "INFO",
             "choices": ["DEBUG", "INFO", "WARN", "ERR"],
             "help": "Outlook/Graph 日志等级"},
            {"flag": "--px-press-screenshots", "type": "bool", "default": False,
             "help": "保存 ruoyi PX/失败相关截图"},
            {"flag": "--har", "type": "bool", "default": False, "help": "保存完整链路 HAR；成功/失败都保存"},
            {"flag": "--email-suffixes", "type": "multi", "default": ["outlook.com"],
             "join": ",",
             "choices": [
                 {"label": "outlook.com", "value": "outlook.com"},
                 {"label": "hotmail.com", "value": "hotmail.com"},
             ],
             "help": "邮箱后缀（可多选，随机池）"},
            {"flag": "--account-format-mode", "type": "choice", "default": "name",
             "choices": [
                 {"label": "随机格式", "value": "random"},
                 {"label": "姓名格式", "value": "name"},
                 {"label": "姓名+数字格式", "value": "name_digits"},
                 {"label": "指定格式", "value": "custom"},
             ],
             "help": "选择账号格式；指定格式时使用下面的账号模板"},
            {"flag": "--account-format", "type": "str", "default": "",
             "help": "指定格式模板，如 {first}.{last}{digits:3}；账号格式选择“指定格式”时生效"},
            {"flag": "--password-format", "type": "str", "default": "",
             "help": "随机密码模板，如 Aa1!{rand:12}；留空用默认随机"},
            {"flag": "--no-verify", "type": "bool", "default": False, "help": "注册后不校验 Outlook 登录"},
            {"flag": "--confirm-before-register", "type": "bool", "default": False,
             "help": "注册页打开后自动点确认，再开始填写"},
        ],
    },
    {
        "id": "register_outlook_standalone",
        "file": "register_outlook_standalone.py",
        "category": "养号/邮箱",
        "title": "Outlook 自注册(独立模式)",
        "desc": "单次/批量跑 Outlook 自注册；支持 auto、protocol、headless、browser 模式，以及代理池文件。",
        "args": [
            {"flag": "--count", "type": "int", "default": 5, "help": "注册数量"},
            {"flag": "--concurrency", "type": "int", "default": 2, "help": "并发数"},
            {"flag": "--mode", "type": "choice", "choices": ["auto", "protocol", "headless", "browser"],
             "default": "auto", "help": "auto=protocol→headless→browser 回退链"},
            {"flag": "--proxy-file", "type": "str", "default": "proxies_outlook.txt", "help": "代理池文件(每行一个 user:pass@host:port)"},
            {"flag": "--account-file", "type": "str", "default": "",
             "help": "可选账号规格文件：账号 / 账号@后缀 / 账号----密码 / 账号@后缀----密码"},
            {"flag": "--email-suffixes", "type": "multi", "default": ["outlook.com"],
             "join": ",",
             "choices": [
                 {"label": "outlook.com", "value": "outlook.com"},
                 {"label": "hotmail.com", "value": "hotmail.com"},
             ],
             "help": "邮箱后缀（可多选，随机池）"},
            {"flag": "--account-format-mode", "type": "choice", "default": "name",
             "choices": [
                 {"label": "随机格式", "value": "random"},
                 {"label": "姓名格式", "value": "name"},
                 {"label": "姓名+数字格式", "value": "name_digits"},
                 {"label": "指定格式", "value": "custom"},
             ],
             "help": "选择账号格式；指定格式时使用下面的账号模板"},
            {"flag": "--account-format", "type": "str", "default": "",
             "help": "指定格式模板，如 {first}.{last}{digits:3}；账号格式选择“指定格式”时生效"},
            {"flag": "--password-format", "type": "str", "default": "",
             "help": "随机密码模板，如 Aa1!{rand:12}；留空用默认随机"},
            {"flag": "--no-proxy", "type": "bool", "default": False, "help": "禁用代理"},
            {"flag": "--timeout", "type": "int", "default": 300, "help": "单号超时(秒)"},
            {"flag": "--no-verify", "type": "bool", "default": False, "help": "注册后不校验 Outlook 登录"},
            {"flag": "--confirm-before-register", "type": "bool", "default": False,
             "help": "注册页打开后自动点确认，再开始填写"},
            {"flag": "--px-press-screenshots", "type": "bool", "default": False,
             "help": "保存 PX 按压前/按压后截图"},
            {"flag": "--protocol-captcharun-px", "type": "bool", "default": False,
             "help": "standalone 协议模式：调用 CaptchaRun PxCaptcha2 + Microsoft risk/verify 后再 CreateAccount"},
            {"flag": "--protocol-captcharun-token", "type": "str", "default": "",
             "help": "standalone 协议模式：CaptchaRun token；留空则读 CAPTCHARUN_API_KEY"},
            {"flag": "--protocol-captcharun-country", "type": "str", "default": "",
             "help": "standalone 协议模式：CaptchaRun 国家代码；留空自动按代理 IP 判断"},
            {"flag": "--protocol-captcharun-timezone", "type": "str", "default": "",
             "help": "standalone 协议模式：CaptchaRun 时区；留空自动按代理 IP 判断"},
            {"flag": "--protocol-captcharun-timeout", "type": "int", "default": 120,
             "help": "standalone 协议模式：CaptchaRun 单阶段轮询超时秒数"},
        ],
    },
    {
        "id": "unlock_outlook",
        "file": "unlock_outlook.py",
        "category": "养号/邮箱",
        "title": "解锁被锁 Outlook",
        "desc": "批量解锁被锁账号，结果分类输出到 unlock_results/。",
        "args": [
            {"flag": "--input", "type": "str", "default": "", "help": "账号文件(每行 email----password；默认找最新 locked)"},
            {"flag": "--proxy-file", "type": "str", "default": "", "help": "住宅代理池文件"},
            {"flag": "--concurrency", "type": "int", "default": 1, "help": "并发数"},
        ],
    },
    {
        "id": "extract_graph_tokens",
        "file": "extract_graph_tokens.py",
        "category": "养号/邮箱",
        "title": "提取 Graph refresh_token",
        "desc": "用账号密码换 Microsoft Graph refresh_token(免浏览器)，输出到 outlook_accounts/。",
        "args": [
            {"flag": "accounts_file", "type": "str", "default": "", "positional": True,
             "help": "账号文件(每行 email----password----...；留空自动扫 unlock_results/)"},
            {"flag": "--email", "type": "str", "default": "", "help": "单个邮箱"},
            {"flag": "--password", "type": "str", "default": "", "help": "单个邮箱密码"},
            {"flag": "--concurrency", "type": "int", "default": 5, "help": "并发数"},
        ],
    },
    {
        "id": "mailbox_broker",
        "file": "mailbox_broker.py",
        "category": "养号/邮箱",
        "title": "共享取码服务(常驻)",
        "desc": "并行流水线时起共享取码服务，避免三窗口并发登录同一邮箱。常驻运行。",
        "args": [
            {"flag": "--host", "type": "str", "default": "127.0.0.1", "help": "监听地址"},
            {"flag": "--port", "type": "int", "default": 8765, "help": "监听端口"},
            {"flag": "--idle", "type": "int", "default": 480, "help": "空闲会话回收秒数"},
        ],
    },
    # ---------------------------------------------------------------- 导出 / 上传
    {
        "id": "upload_tokens",
        "file": "upload_tokens.py",
        "category": "导出/上传",
        "title": "上传标准 token",
        "desc": "把 tokens/ 下标准 token 上传到 CPA/SUB2API/webchat2api。位置参数：all/chatgpt/grok。",
        "args": [
            {"flag": "target", "type": "choice", "choices": ["all", "chatgpt", "grok"],
             "default": "all", "positional": True, "help": "上传目标"},
        ],
    },
    {
        "id": "export_chatgpt2api",
        "file": "export_chatgpt2api.py",
        "category": "导出/上传",
        "title": "导出/上传 chatgpt2api",
        "desc": "聚合普通网页号导入 chatgpt2api(--post 直传 / 默认导出 txt)。",
        "args": [
            {"flag": "--post", "type": "str", "default": "", "help": "直接 POST 到的 host(留空只导出文件)"},
            {"flag": "--key", "type": "str", "default": "", "help": "chatgpt2api admin key(配合 --post)"},
            {"flag": "--json", "type": "bool", "default": False, "help": "导出 JSON(默认一行一个 access_token)"},
            {"flag": "--out", "type": "str", "default": "", "help": "输出文件路径"},
        ],
    },
    {
        "id": "export_accounts",
        "file": "export_accounts.py",
        "category": "导出/上传",
        "title": "导出账号 cookie",
        "desc": "导出已注册账号 cookie 供直登扩展使用(无参=全部平台)。",
        "args": [],
    },
]


def script_by_id(sid):
    for s in SCRIPTS:
        if s["id"] == sid:
            return s
    return None


# ============================================================ 外部工具链接
# 不在本机跑的 web 服务/工具，面板上以"打开链接"卡片呈现(新标签打开)。
EXTERNAL_LINKS = [
]


# ============================================================ 内嵌功能页
# 直接在面板里 iframe 内嵌的外部页面 + 可选 sms-man 接码助手。
EMBED_PAGES = [
    {"id": "gmail", "title": "Gmail 注册", "url": "https://gmails.zeabur.app/gq",
     "desc": "在线 Gmail 注册(内嵌)。右侧接码助手可租号 + 同号多次取验证码,人工复制粘贴到左侧表单。",
     "sms_helper": True, "sms_service_default": "google"},
]


# ============================================================ .env 配置 schema
# group: 分组标题；key: 变量名；required: 是否必填(运行对应功能时)；help: 说明；
# secret: True 时前端用密码框；default: 模板默认值(仅展示)。
ENV_SCHEMA = [
    {"group": "Clash 代理(节点切换/出口)", "tests": [{"target": "clash", "label": "测试 Clash 连通"}], "items": [
        {"key": "CLASH_SECRET", "required": True, "secret": True,
         "help": "Clash Verge → 设置 → 外部控制器(External Controller) 里设的 secret/密钥。"
                 "若该处留空,这里也留空。设了密钥不填会连不上控制器(节点切换失效)。"},
        {"key": "CLASH_API", "default": "http://127.0.0.1:9097",
         "help": "Clash 控制器地址。Clash Verge 默认端口 9097,mihomo 内核默认 9090。在 外部控制器 页可看到。"},
        {"key": "CLASH_PROXY", "default": "http://127.0.0.1:7897",
         "help": "Clash 混合代理端口(mixed-port),脚本走它出网。Verge 默认 7897。"},
        {"key": "CLASH_GROUP", "default": "GLOBAL",
         "help": "决定出口的代理组名。global 模式下填 GLOBAL;规则模式填你的节点选择组名。"},
    ]},
    {"group": "指纹浏览器", "tests": [{"target": "bitbrowser", "label": "测试 指纹浏览器连通"}], "items": [
        {"key": "FINGERPRINT_BROWSER", "type": "choice", "choices": ["bitbrowser", "adspower"],
         "default": "bitbrowser", "help": "选择当前指纹浏览器"},
        {"key": "BITBROWSER_API", "default": "http://127.0.0.1:54345", "help": "比特浏览器本地 API"},
        {"key": "ADSPOWER_API", "default": "http://127.0.0.1:50325", "help": "AdsPower 本地 API"},
        {"key": "ADSPOWER_API_KEY", "secret": True, "help": "AdsPower API key，未启用鉴权时留空"},
        {"key": "ADSPOWER_GROUP_ID", "default": "0", "help": "AdsPower 新建 profile 的分组 ID"},
    ]},
    {"group": "短信接码", "tests": [{"target": "smsman", "label": "测试 sms-man"}, {"target": "firefox", "label": "测试 firefox.fun"}], "items": [
        {"key": "SMS_TOKEN", "secret": True, "help": "firefox.fun 接码 token"},
        {"key": "HERO_SMS_API_KEY", "secret": True, "help": "hero-sms.com 备用接码 key"},
        {"key": "SMSMAN_TOKEN", "secret": True, "help": "sms-man.com 接码 key(Codex add-phone 主用)"},
        {"key": "SMSMAN_APP_ID_OPENAI", "default": "openai", "help": "sms-man OpenAI 服务 id(openai 自动解析为 2754)"},
        {"key": "SMSMAN_APP_ID_GMAIL", "default": "google", "help": "sms-man Gmail/Google 服务 id(google 自动解析;接码助手默认用它)"},
        {"key": "SMSMAN_COUNTRY_ID_OPENAI", "default": "0", "help": "国家 id(0=随机/按价格)"},
    ]},
    {"group": "打码平台(可选)", "items": [
        {"key": "CAPTCHARUN_API_KEY", "secret": True,
         "help": "CaptchaRun token，用于 standalone 协议模式；对应页面 token 留空时自动读取"},
        {"key": "CAPSOLVER_API_KEY", "secret": True, "help": "CapSolver 打码 key"},
        {"key": "EZCAPTCHA_API_KEY", "secret": True, "help": "EZ-Captcha 打码 key(解锁 Outlook 用)"},
        {"key": "YESCAPTCHA_API_KEY", "secret": True, "help": "YesCaptcha key(GitHub Arkose 备用)"},
    ]},
    {"group": "Outlook 邮箱来源", "items": [
        {"key": "OUTLOOK_CARD", "secret": True, "help": "闪客云邮箱卡密(接口取号用)"},
        {"key": "OUTLOOK_PROXIES", "help": "Outlook 自注册住宅代理池(换行/逗号分隔)"},
        {"key": "OUTLOOK_LOG_LEVEL", "type": "choice", "choices": ["DEBUG", "INFO", "WARN", "ERR"],
         "default": "INFO", "help": "Outlook/Graph 默认日志等级"},
        {"key": "OUTLOOK_AIMILI_POOL_URL", "default": "", "help": "AimiliVPN URL：管理端根地址或 /api/pool/proxies(/random) 完整地址"},
        {"key": "OUTLOOK_AIMILI_POOL_TOKEN", "secret": True, "help": "AimiliVPN 代理池 API Token"},

    ]},
    {"group": "SUB2API(Codex 导入)", "items": [
        {"key": "SUB2API_URL", "help": "SUB2API 管理接口地址(用 --codex 时必填)"},
        {"key": "SUB2API_EMAIL", "help": "SUB2API 登录邮箱"},
        {"key": "SUB2API_PASSWORD", "secret": True, "help": "SUB2API 登录密码"},
        {"key": "SUB2API_GROUP", "default": "codex", "help": "目标分组名(需后台先建好)"},
    ]},
    {"group": "CPA(codex 授权文件导入)", "items": [
        {"key": "CPA_URL", "help": "CPA 管理接口地址"},
        {"key": "CPA_MGMT_KEY", "secret": True, "help": "CPA 管理 key"},
    ]},
    {"group": "chatgpt2api(普通网页号)", "items": [
        {"key": "CHATGPT2API_URL", "help": "chatgpt2api host(用 --import-c2a 时必填)"},
        {"key": "CHATGPT2API_KEY", "secret": True, "help": "chatgpt2api admin key"},
    ]},
    {"group": "webchat2api(Grok sso)", "items": [
        {"key": "WEBCHAT2API_URL", "help": "webchat2api 地址(用 Grok 时)"},
        {"key": "WEBCHAT2API_KEY", "secret": True, "help": "webchat2api key"},
    ]},
    {"group": "Codex add-phone 接码调参", "items": [
        {"key": "CODEX_PHONE_SKIP_ATTEMPTS", "default": "0", "help": "先赌免手机次数(0=直接接码)"},
        {"key": "CODEX_ADDPHONE_ATTEMPTS", "default": "2", "help": "接码换号上限次数"},
        {"key": "CODEX_SMS_TIMEOUT", "default": "150", "help": "单号等码超时(秒)"},
        {"key": "SMS_COUNTRY_BLACKLIST_OPENAI", "default": "261,63", "help": "拉黑号段(dialing code)"},
    ]},
    {"group": "GitHub Arkose 视觉投票(可选)", "items": [
        {"key": "VISION_API_BASE", "help": "主视觉网关(OpenAI 兼容)"},
        {"key": "VISION_API_KEY", "secret": True, "help": "主视觉网关 key"},
        {"key": "VOTE_ZZ_BASE", "help": "投票中转网关(gemini/gpt)"},
        {"key": "VOTE_ZZ_KEY", "secret": True, "help": "投票网关 gemini key"},
        {"key": "VOTE_GPT_KEY", "secret": True, "help": "投票网关 gpt key"},
        {"key": "VOTE_OPUS_BASE", "help": "claude opus 专用网关"},
        {"key": "VOTE_OPUS_KEY", "secret": True, "help": "opus 网关 key"},
    ]},
]


# 所有 schema 里出现的 .env key（用于读 .env 时补齐未在模板里的项不丢）
def env_keys():
    keys = []
    for g in ENV_SCHEMA:
        for it in g["items"]:
            keys.append(it["key"])
    return keys
