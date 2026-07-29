@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ============================================================
echo   reg-factory installer (Python venv + deps + browser core)
echo ============================================================
echo.

REM ---- 0. local proxy for pip / GitHub downloads (Clash/V2Ray etc.) ----
REM Override: set REG_FACTORY_PROXY=http://127.0.0.1:7897 before running
REM Disable:  set REG_FACTORY_PROXY=off
if not defined REG_FACTORY_PROXY set "REG_FACTORY_PROXY=http://127.0.0.1:10808"
if /I not "%REG_FACTORY_PROXY%"=="off" if /I not "%REG_FACTORY_PROXY%"=="0" if not "%REG_FACTORY_PROXY%"=="" (
  set "HTTP_PROXY=%REG_FACTORY_PROXY%"
  set "HTTPS_PROXY=%REG_FACTORY_PROXY%"
  set "ALL_PROXY=%REG_FACTORY_PROXY%"
  set "http_proxy=%REG_FACTORY_PROXY%"
  set "https_proxy=%REG_FACTORY_PROXY%"
  set "all_proxy=%REG_FACTORY_PROXY%"
  echo [0/6] download proxy: %REG_FACTORY_PROXY%
) else (
  echo [0/6] download proxy: off
)

REM ---- 1. find Python (>=3.10) ----
set PY=
where py >nul 2>nul && set PY=py -3
if "%PY%"=="" (
  where python >nul 2>nul && set PY=python
)
if "%PY%"=="" (
  echo [ERROR] Python not found. Install Python 3.10+ from https://www.python.org/downloads/
  echo         Check "Add Python to PATH" during install, then run this script again.
  pause
  exit /b 1
)
echo [1/6] Python: %PY%
%PY% --version

REM ---- 2. create venv ----
if exist ".venv\Scripts\python.exe" (
  echo [2/6] venv exists, skip create.
) else (
  echo [2/6] creating venv .venv ...
  %PY% -m venv .venv
  if errorlevel 1 ( echo [ERROR] venv create failed & pause & exit /b 1 )
)

set VENV_PY=.venv\Scripts\python.exe

REM ---- 3. install deps ----
echo [3/6] installing deps (pip install -r requirements.txt) ...
"%VENV_PY%" -m pip install --upgrade pip >nul 2>nul
"%VENV_PY%" -m pip install -r requirements.txt
if errorlevel 1 ( echo [ERROR] deps install failed, check network/pip mirror & pause & exit /b 1 )

REM ---- 4. install Playwright Chromium ----
echo [4/6] installing Playwright Chromium ...
"%VENV_PY%" -m playwright install chromium
if errorlevel 1 ( echo [WARN] playwright core install failed, run later: .venv\Scripts\playwright install chromium )

REM ---- 5. install ruyipage Firefox runtime (ruoyi backend) ----
echo [5/6] installing ruyipage Firefox runtime (for Outlook ruoyi backend) ...
"%VENV_PY%" -m ruyipage install
if errorlevel 1 (
  echo [WARN] ruyipage direct install failed, try mirror download via proxy ...
  set "RUOYI_ZIP=%TEMP%\firefox-ruyi-win64.zip"
  set "RUOYI_URL=https://github.com/LoseNine/ruyipage/releases/download/151-ruyi/firefox-151.0a1.en-US.win64.zip"
  curl -L --retry 3 --connect-timeout 20 -x "%REG_FACTORY_PROXY%" -o "!RUOYI_ZIP!" "!RUOYI_URL!"
  if exist "!RUOYI_ZIP!" (
    "%VENV_PY%" -m ruyipage install --from-file "!RUOYI_ZIP!"
  )
)
if errorlevel 1 (
  echo [WARN] ruyipage Firefox install failed, run later: .venv\Scripts\python.exe -m ruyipage install
) else (
  for /f "usebackq delims=" %%i in (`"%VENV_PY%" -m ruyipage path 2^>nul`) do (
    if not "%%i"=="" (
      echo       Firefox path: %%i
      if not defined RUOYI_FIREFOX_PATH set "RUOYI_FIREFOX_PATH=%%i"
    )
  )
)

REM ---- 6. prepare .env ----
if exist ".env" (
  echo [6/6] .env exists, keep your config.
) else (
  if exist ".env.example" (
    copy ".env.example" ".env" >nul
    echo [6/6] .env created from template, fill keys later in the web panel Config page.
  ) else (
    echo [6/6] .env.example not found, skip.
  )
)

REM If ruyipage path is known and .env has no RUOYI_FIREFOX_PATH, append it.
if defined RUOYI_FIREFOX_PATH (
  if exist ".env" (
    findstr /B /C:"RUOYI_FIREFOX_PATH=" ".env" >nul 2>nul
    if errorlevel 1 (
      echo.>>".env"
      echo # ruyipage Firefox runtime path (from install.bat)>>".env"
      echo RUOYI_FIREFOX_PATH=%RUOYI_FIREFOX_PATH%>>".env"
      echo       wrote RUOYI_FIREFOX_PATH into .env
    )
  )
)

echo.
echo ============================================================
echo   Install done!
echo   - Start BitBrowser/AdsPower and Clash Verge clients
echo   - Double-click start.bat to open the control panel
echo   - Outlook ruoyi backend needs ruyipage Firefox (step 5)
echo   - Proxy used: %REG_FACTORY_PROXY%  (set REG_FACTORY_PROXY=off to disable)
echo ============================================================
pause
