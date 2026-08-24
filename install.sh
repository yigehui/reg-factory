#!/usr/bin/env bash
# reg-factory 一键安装 (mac/linux)
set -e
cd "$(dirname "$0")"

echo "============================================================"
echo "  reg-factory 一键安装 (Python 环境 + 依赖 + 浏览器内核)"
echo "============================================================"

# ---- 0. 本机代理（pip / GitHub 下载）----
# 覆盖: REG_FACTORY_PROXY=http://127.0.0.1:7897 ./install.sh
# 关闭: REG_FACTORY_PROXY=off ./install.sh
REG_FACTORY_PROXY="${REG_FACTORY_PROXY:-http://127.0.0.1:10808}"
if [ "$REG_FACTORY_PROXY" != "off" ] && [ "$REG_FACTORY_PROXY" != "0" ] && [ -n "$REG_FACTORY_PROXY" ]; then
  export HTTP_PROXY="$REG_FACTORY_PROXY"
  export HTTPS_PROXY="$REG_FACTORY_PROXY"
  export ALL_PROXY="$REG_FACTORY_PROXY"
  export http_proxy="$REG_FACTORY_PROXY"
  export https_proxy="$REG_FACTORY_PROXY"
  export all_proxy="$REG_FACTORY_PROXY"
  echo "[0/6] 下载代理: $REG_FACTORY_PROXY"
else
  echo "[0/6] 下载代理: 关闭"
fi

PY=python3
command -v python3 >/dev/null 2>&1 || PY=python
command -v "$PY" >/dev/null 2>&1 || { echo "[错误] 没找到 Python。请先安装 Python 3.10+"; exit 1; }
echo "[1/6] 使用 Python: $($PY --version)"

if [ -x ".venv/bin/python" ]; then
  echo "[2/6] 虚拟环境已存在,跳过创建。"
else
  echo "[2/6] 创建虚拟环境 .venv ..."
  "$PY" -m venv .venv
fi
VENV_PY=".venv/bin/python"

echo "[3/6] 安装依赖 ..."
"$VENV_PY" -m pip install --upgrade pip >/dev/null
"$VENV_PY" -m pip install -r requirements.txt

echo "[4/6] 安装 Playwright Chromium 内核 ..."
"$VENV_PY" -m playwright install chromium || echo "[警告] 内核安装失败,可稍后手动跑: .venv/bin/playwright install chromium"

echo "[5/6] 安装 ruyipage Firefox 内核 (Outlook ruoyi 后端) ..."
RUOYI_PATH=""
if "$VENV_PY" -m ruyipage install; then
  RUOYI_PATH="$("$VENV_PY" -m ruyipage path 2>/dev/null || true)"
else
  echo "[警告] ruyipage 直连安装失败，尝试经代理镜像下载 ..."
  RUOYI_ZIP="${TMPDIR:-/tmp}/firefox-ruyi.zip"
  RUOYI_URL="https://github.com/LoseNine/ruyipage/releases/download/v1.2.58/firefox-155.0a1.en-US.win64-20260803.zip"
  case "$(uname -s)" in
    Darwin) RUOYI_URL="https://github.com/LoseNine/ruyipage/releases/download/v1.2.58/firefox-155.0a1.en-US.mac.zip" ;;
    Linux)  RUOYI_URL="https://github.com/LoseNine/ruyipage/releases/download/v1.2.58/firefox-155.0a1.en-US.linux-x86_64.tar.xz" ;;
  esac
  if command -v curl >/dev/null 2>&1; then
    if [ -n "$REG_FACTORY_PROXY" ] && [ "$REG_FACTORY_PROXY" != "off" ] && [ "$REG_FACTORY_PROXY" != "0" ]; then
      curl -L --retry 3 --connect-timeout 20 -x "$REG_FACTORY_PROXY" -o "$RUOYI_ZIP" "$RUOYI_URL" || true
    else
      curl -L --retry 3 --connect-timeout 20 -o "$RUOYI_ZIP" "$RUOYI_URL" || true
    fi
  fi
  if [ -f "$RUOYI_ZIP" ] && [ -s "$RUOYI_ZIP" ]; then
    "$VENV_PY" -m ruyipage install --from-file "$RUOYI_ZIP" || true
  fi
  RUOYI_PATH="$("$VENV_PY" -m ruyipage path 2>/dev/null || true)"
fi

if [ -n "$RUOYI_PATH" ]; then
  echo "      Firefox 路径: $RUOYI_PATH"
else
  echo "[警告] ruyipage Firefox 安装失败,可稍后手动跑: .venv/bin/python -m ruyipage install"
fi

if [ -f ".env" ]; then
  echo "[6/6] .env 已存在,保留你的配置。"
elif [ -f ".env.example" ]; then
  cp .env.example .env
  echo "[6/6] 已从模板生成 .env,稍后在面板"配置"页填写密钥。"
fi

# 若装好了 ruyipage 且 .env 还没有 RUOYI_FIREFOX_PATH，写入实际路径
if [ -n "$RUOYI_PATH" ] && [ -f ".env" ]; then
  if ! grep -q '^RUOYI_FIREFOX_PATH=' .env 2>/dev/null; then
    {
      echo ""
      echo "# ruyipage Firefox runtime path (from install.sh)"
      echo "RUOYI_FIREFOX_PATH=$RUOYI_PATH"
    } >> .env
    echo "      已将 RUOYI_FIREFOX_PATH 写入 .env"
  fi
fi

echo ""
echo "============================================================"
echo "  安装完成! 确保 BitBrowser/AdsPower 已打开,"
echo "  然后运行: ./start.sh  打开控制面板"
echo "  Outlook ruoyi 后端依赖 ruyipage Firefox（步骤 5）"
echo "  下载代理: $REG_FACTORY_PROXY  (REG_FACTORY_PROXY=off 可关闭)"
echo "============================================================"
