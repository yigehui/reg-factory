# -*- coding: utf-8 -*-
"""
config.py — 全局配置。

所有密钥/凭据都从环境变量读取（默认空），不在仓库里留明文。
支持把变量写进同目录的 .env 文件（见 .env.example）；.env 只在对应环境
变量尚未设置时生效，不会覆盖真实的进程环境变量。
"""

import os


# ---------------------------------------------------------------- .env 加载
def _load_dotenv(path=None):
    """零依赖 .env 读取器：解析 KEY=VALUE，忽略空行与 # 注释。
    只在 os.environ 里尚未设置该 KEY 时填入（真实环境变量优先）。"""
    if path is None:
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if not os.path.isfile(path):
        return
    try:
        with open(path, encoding="utf-8") as f:
            for raw in f:
                line = raw.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, _, val = line.partition("=")
                key = key.strip()
                val = val.strip().strip('"').strip("'")
                if key and key not in os.environ:
                    os.environ[key] = val
    except Exception:
        pass


_load_dotenv()


def _env(name, default=""):
    return os.environ.get(name, default)


# ---------------------------------------------------------------- 本地基建
# Fingerprint browser provider: bitbrowser / adspower
FINGERPRINT_BROWSER = _env("FINGERPRINT_BROWSER", "bitbrowser").strip().lower()

# BitBrowser 本地 API 地址
BITBROWSER_API = _env("BITBROWSER_API", "http://127.0.0.1:54345")

# AdsPower 本地 API 地址
ADSPOWER_API = _env("ADSPOWER_API", "http://127.0.0.1:50325")
ADSPOWER_API_KEY = _env("ADSPOWER_API_KEY", "")
ADSPOWER_GROUP_ID = _env("ADSPOWER_GROUP_ID", "0")

# Claude.ai 注册相关 URL
CLAUDE_LOGIN_URL = "https://claude.ai/login"

# Cookie 输出目录
COOKIE_OUTPUT_DIR = "cookies"

# ---------------------------------------------------------------- 域名邮箱（备用）
MAIL_DOMAIN = _env("MAIL_DOMAIN", "")
MAIL_API_BASE = _env("MAIL_API_BASE", "")
MAIL_ADMIN_USER = _env("MAIL_ADMIN_USER", "admin")
MAIL_ADMIN_PASS = _env("MAIL_ADMIN_PASS", "")
# JWT token（从浏览器抓取，可能会过期需要更新）
MAIL_AUTH_TOKEN = _env("MAIL_AUTH_TOKEN", "")
# 新建邮箱统一密码
MAIL_NEW_PASS = _env("MAIL_NEW_PASS", "")

# ---------------------------------------------------------------- Outlook 邮箱 API (闪客云邮箱)
OUTLOOK_API_BASE = _env("OUTLOOK_API_BASE", "http://api.shankeyun.com")
OUTLOOK_CARD = _env("OUTLOOK_CARD", "")  # 闪客云卡密
OUTLOOK_TYPE = _env("OUTLOOK_TYPE", "outlook")  # outlook / hotmail / any

# ---------------------------------------------------------------- 短信接码平台 (firefox.fun)
SMS_API_BASE = _env("SMS_API_BASE", "http://www.firefox.fun/yhapi.ashx")
SMS_TOKEN = _env("SMS_TOKEN", "")  # 接码平台 token
SMS_PROJECT_ID = _env("SMS_PROJECT_ID", "2313")  # claude 项目
# 优先国家列表，按顺序尝试，""=任意(排除黑名单)
SMS_COUNTRY_PREFER = ["60", "56", "57", "44", ""]  # 60=马来西亚 56=智利 57=哥伦比亚 44=英国 ""=任意
SMS_COUNTRY_BLACKLIST = ["63"]  # 菲律宾

# ---------------------------------------------------------------- 备用短信平台 (hero-sms.com)
HERO_SMS_API_BASE = _env("HERO_SMS_API_BASE", "https://hero-sms.com/stubs/handler_api.php")
HERO_SMS_API_KEY = _env("HERO_SMS_API_KEY", "")  # 备用接码 api_key
HERO_SMS_SERVICE = _env("HERO_SMS_SERVICE", "acz")  # Claude 专用服务
# 优先国家: 7=马来西亚 52=泰国 16=英国 56=西班牙 39=阿根廷 86=意大利 34=爱沙尼亚 49=立陶宛 36=中国
HERO_SMS_COUNTRY_PREFER = [7, 52, 16, 56, 39, 86, 34, 49, 36]

# ---------------------------------------------------------------- 打码平台
# CapSolver 验证码打码平台
CAPSOLVER_API_KEY = _env("CAPSOLVER_API_KEY", "")

# EZ-Captcha 验证码打码平台
EZCAPTCHA_API_KEY = _env("EZCAPTCHA_API_KEY", "")
EZCAPTCHA_API_BASE = _env("EZCAPTCHA_API_BASE", "https://api.ez-captcha.com")

# YesCaptcha 打码平台（解 Arkose FunCaptcha，GitHub 注册用）。API 与 CapSolver 兼容。
YESCAPTCHA_API_KEY = _env("YESCAPTCHA_API_KEY", "")
YESCAPTCHA_API_BASE = _env("YESCAPTCHA_API_BASE", "https://api.yescaptcha.com")

# ---------------------------------------------------------------- agent-captcha 视觉投票求解器
# GitHub Arkose 拼图用多模态大模型「投票」求解（common/agent_captcha.py）。
# 各家网关 OpenAI 兼容(/v1/chat/completions)；claude/opus 走 Anthropic 原生(/v1/messages)。
# 主视觉网关（gpt-5.x，图像增强 gpt-image-2 也在此）
VISION_API_BASE = _env("VISION_API_BASE", "")
VISION_API_KEY = _env("VISION_API_KEY", "")
# 图像增强兜底网关（gpt-image-2 images/edits）
IMAGE_EDIT_BASE2 = _env("IMAGE_EDIT_BASE2", "")
IMAGE_EDIT_KEY2 = _env("IMAGE_EDIT_KEY2", "")
# 投票池：中转网关(gemini/gpt) + claude 专用网关。逗号分隔的 key 留空则该模型不参与。
VOTE_ZZ_BASE = _env("VOTE_ZZ_BASE", "")          # 中转网关(gemini-3.5-flash / gemini-3.1-pro / gpt-5.5)
VOTE_ZZ_KEY = _env("VOTE_ZZ_KEY", "")            # 上面网关里 gemini 用的 key
VOTE_GPT_KEY = _env("VOTE_GPT_KEY", "")          # 同网关里 gpt-5.5 用的 key（可与 ZZ_KEY 不同）
VOTE_OPUS_BASE = _env("VOTE_OPUS_BASE", "")      # claude opus 专用网关（Anthropic /v1/messages）
VOTE_OPUS_KEY = _env("VOTE_OPUS_KEY", "")
# gemma 免费兜底文本网关（可选）
GEMMA_API_BASE = _env("GEMMA_API_BASE", "")
GEMMA_API_KEY = _env("GEMMA_API_KEY", "")

# ---------------------------------------------------------------- Telegram 通知
# 配置后 common/notify.send_tg_message 会把消息发到该 chat；留空则不发。
# 任意注册/导出/循环流程均可复用（ruoyi 循环养号每批汇总即走此通道）。
# TG 走的 HTTP 代理(国内网络必填，否则 api.telegram.org 直连不通)，如 Clash 的 http://127.0.0.1:7897
TG_BOT_TOKEN = _env("TG_BOT_TOKEN", "")
TG_CHAT_ID = _env("TG_CHAT_ID", "")
TG_PROXY = _env("TG_PROXY", "")

# ---------------------------------------------------------------- 标准 token 导出/上传
# 注册成功后落地的标准格式 token 目录（CPA codex / SUB2API content / grok sso）
TOKEN_OUTPUT_DIR = _env("TOKEN_OUTPUT_DIR", "tokens")

# CPA 管理接口（ChatGPT codex 授权文件导入）
CPA_URL = _env("CPA_URL", "")
CPA_MGMT_KEY = _env("CPA_MGMT_KEY", "")

# SUB2API 管理接口（ChatGPT codex-session 导入）
SUB2API_URL = _env("SUB2API_URL", "")
SUB2API_EMAIL = _env("SUB2API_EMAIL", "")
SUB2API_PASSWORD = _env("SUB2API_PASSWORD", "")
SUB2API_GROUP = _env("SUB2API_GROUP", "codex")  # 目标分组名，需先在 SUB2API 后台建好

# webchat2api（Grok sso 注入）
WEBCHAT2API_URL = _env("WEBCHAT2API_URL", "")
WEBCHAT2API_KEY = _env("WEBCHAT2API_KEY", "")

# chatgpt2api（basketikun/chatgpt2api 普通网页号导入，POST <url>/api/accounts）
# register_chatgpt.py --import-c2a 注册成功后逐个上传时用
CHATGPT2API_URL = _env("CHATGPT2API_URL", "")  # 对端 host（见 .env）
CHATGPT2API_KEY = _env("CHATGPT2API_KEY", "")  # 对端 admin key（Authorization: Bearer）

# ---------------------------------------------------------------- 订阅授权入口
# Claude / SuperGrok 订阅入口（激活码 CDK 流程「敬请期待」，后续支持授权到 SUB2API / CPA）
CLAUDE_SUB_URL = _env("CLAUDE_SUB_URL", "https://6661231.xyz/#/claude")
GROK_SUB_URL = _env("GROK_SUB_URL", "https://6661231.xyz/#/grok")
# 激活码 CDK 池（预留，逗号/换行/空格分隔）
CLAUDE_SUB_CDK = [c.strip() for c in _env("CLAUDE_SUB_CDK", "").replace("\n", ",").replace(" ", ",").split(",") if c.strip()]
GROK_SUB_CDK = [c.strip() for c in _env("GROK_SUB_CDK", "").replace("\n", ",").replace(" ", ",").split(",") if c.strip()]

# ---------------------------------------------------------------- ChatGPT OAuth add-phone 接码
# OpenAI/ChatGPT 在接码平台的服务号（按平台分，跟 Claude 的不同）
SMS_PROJECT_ID_OPENAI = _env("SMS_PROJECT_ID_OPENAI", "")  # firefox.fun 的 ChatGPT 项目 iid（待填）
HERO_SMS_SERVICE_OPENAI = _env("HERO_SMS_SERVICE_OPENAI", "dr")  # hero-sms/sms-activate OpenAI 服务码默认 dr
# firefox.fun 价格上限：'0' 只取最便宜(垃圾号易被 OpenAI 拒)，给够才摸得到智利等好号
SMS_MAXPRICE_OPENAI = _env("SMS_MAXPRICE_OPENAI", "20")
# OpenAI add-phone 拉黑的号段(dialing code)：261 马达加斯加、63 菲律宾 等 OpenAI 常拒的
SMS_COUNTRY_BLACKLIST_OPENAI = [c.strip() for c in _env("SMS_COUNTRY_BLACKLIST_OPENAI", "261,63").split(",") if c.strip()]

# ---------------------------------------------------------------- 接码平台 (sms-man.com)
# sms-man.com API v2.0：base/control，token 鉴权，JSON 响应。过 Codex add-phone 主用。
# 返回的 number 已含国家码。app_id 支持数字 application_id 或 code/名(运行时查 /applications 解析)。
SMSMAN_API_BASE = _env("SMSMAN_API_BASE", "https://api.sms-man.com/control")
SMSMAN_TOKEN = _env("SMSMAN_TOKEN", "")  # sms-man.com API key（profile 页获取）
SMSMAN_APP_ID_OPENAI = _env("SMSMAN_APP_ID_OPENAI", "openai")  # 数字 application_id 或 code/名(自动解析)
SMSMAN_COUNTRY_ID_OPENAI = _env("SMSMAN_COUNTRY_ID_OPENAI", "0")  # 0=随机国家
SMSMAN_MAXPRICE_OPENAI = _env("SMSMAN_MAXPRICE_OPENAI", "")  # 价格上限（sms-man 币种），空=不限
