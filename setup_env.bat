@echo off
chcp 65001 >nul
setlocal EnableExtensions
title SD WebUI(Aki) 提示词图集 - 环境安装

rem ============================================================
rem  一次性环境安装脚本
rem  1. 检查 Python（没有则提示并可选自动下载安装）
rem  2. 在脚本目录创建 .venv 虚拟环境
rem  3. 仅在本次运行中设置 pip 临时加速源并安装依赖
rem  4. 依赖只作用于 .venv，不污染系统 Python
rem ============================================================

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

set "VENV_DIR=%SCRIPT_DIR%.venv"
set "VPY=%VENV_DIR%\Scripts\python.exe"
set "PKGS=pillow python-docx"

echo ============================================================
echo   SD WebUI(Aki) 提示词图集 - 环境安装
echo   脚本目录: %SCRIPT_DIR%
echo ============================================================
echo.

rem ---------------- 1. 选择临时加速源 ----------------
set "MIRROR_ARG=%~1"
if /i "%MIRROR_ARG%"=="tuna"     goto :m_tuna
if /i "%MIRROR_ARG%"=="aliyun"   goto :m_aliyun
if /i "%MIRROR_ARG%"=="tencent"  goto :m_tencent
if /i "%MIRROR_ARG%"=="official" goto :m_official

echo 请选择本次安装使用的 pip 临时加速源（只在本窗口生效，关闭即失效）：
echo   [1] 清华 TUNA       https://pypi.tuna.tsinghua.edu.cn/simple
echo   [2] 阿里云           https://mirrors.aliyun.com/pypi/simple
echo   [3] 腾讯云           https://mirrors.cloud.tencent.com/pypi/simple
echo   [4] 官方 PyPI（不加速）
echo.
choice /c 1234 /n /m "请输入 1-4（默认 1）: "
if errorlevel 4 goto :m_official
if errorlevel 3 goto :m_tencent
if errorlevel 2 goto :m_aliyun
goto :m_tuna

:m_tuna
set "PIP_INDEX_URL=https://pypi.tuna.tsinghua.edu.cn/simple"
set "PIP_TRUSTED_HOST=pypi.tuna.tsinghua.edu.cn"
set "MIRROR_NAME=清华 TUNA"
goto :mirror_ready
:m_aliyun
set "PIP_INDEX_URL=https://mirrors.aliyun.com/pypi/simple"
set "PIP_TRUSTED_HOST=mirrors.aliyun.com"
set "MIRROR_NAME=阿里云"
goto :mirror_ready
:m_tencent
set "PIP_INDEX_URL=https://mirrors.cloud.tencent.com/pypi/simple"
set "PIP_TRUSTED_HOST=mirrors.cloud.tencent.com"
set "MIRROR_NAME=腾讯云"
goto :mirror_ready
:m_official
set "PIP_INDEX_URL=https://pypi.org/simple"
set "PIP_TRUSTED_HOST=pypi.org"
set "MIRROR_NAME=官方 PyPI"
goto :mirror_ready

:mirror_ready
set "PIP_DISABLE_PIP_VERSION_CHECK=1"
set "PIP_DEFAULT_TIMEOUT=20"
echo.
echo [加速源] %MIRROR_NAME%  %PIP_INDEX_URL%
echo          ^(仅当前窗口临时生效^)
echo.

rem ---------------- 2. 检查 Python ----------------
rem 只认真正可执行的 Python 3.9+，避免把 Microsoft Store 的占位别名当成 Python
set "PY_CMD="
python -c "import sys;sys.exit(0 if sys.version_info>=(3,9) else 1)" >nul 2>nul && set "PY_CMD=python"
if not defined PY_CMD (
    py -3 -c "import sys;sys.exit(0 if sys.version_info>=(3,9) else 1)" >nul 2>nul && set "PY_CMD=py -3"
)

if defined PY_CMD goto :python_ready

echo [警告] 没有找到可用的 Python（未安装，或版本低于 3.9）。
echo.
echo   本脚本需要 Python 3.9 或更高版本。
echo   你可以手动安装：https://mirrors.huaweicloud.com/python/3.11.9/python-3.11.9-amd64.exe
echo   安装时请勾选 "Add python.exe to PATH"。
echo.
choice /c yn /n /m "是否现在自动下载并安装 Python 3.11.9？(y/n): "
if errorlevel 2 goto :no_python

set "PY_INSTALLER=%TEMP%\python-3.11.9-amd64.exe"
set "PY_URL=https://mirrors.huaweicloud.com/python/3.11.9/python-3.11.9-amd64.exe"
set "PY_URL_BACKUP=https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"
echo.
echo [下载] %PY_URL%
where curl >nul 2>nul
if errorlevel 1 (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "try { Invoke-WebRequest -Uri '%PY_URL%' -OutFile '%PY_INSTALLER%' } catch { exit 1 }"
) else (
    curl -L --fail --retry 3 -o "%PY_INSTALLER%" "%PY_URL%"
)
if not exist "%PY_INSTALLER%" (
    echo [提示] 镜像下载失败，改用 Python 官方地址...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "try { Invoke-WebRequest -Uri '%PY_URL_BACKUP%' -OutFile '%PY_INSTALLER%' } catch { exit 1 }"
)
if not exist "%PY_INSTALLER%" goto :no_python

echo [安装] 正在静默安装（仅当前用户，自动加入 PATH）...
start /wait "" "%PY_INSTALLER%" /passive InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_launcher=1
del /q "%PY_INSTALLER%" >nul 2>nul

rem 刷新当前窗口的 PATH，让新装的 Python 立即可用
set "PATH=%LOCALAPPDATA%\Programs\Python\Python311;%LOCALAPPDATA%\Programs\Python\Python311\Scripts;%PATH%"
python -c "import sys;sys.exit(0 if sys.version_info>=(3,9) else 1)" >nul 2>nul && set "PY_CMD=python"
if not defined PY_CMD (
    py -3 -c "import sys;sys.exit(0 if sys.version_info>=(3,9) else 1)" >nul 2>nul && set "PY_CMD=py -3"
)
if not defined PY_CMD goto :no_python

:python_ready
for /f "delims=" %%v in ('%PY_CMD% -c "import sys;print(sys.version.split()[0])" 2^>nul') do set "PY_VER=%%v"
echo [Python] 已找到 Python %PY_VER%  ^(%PY_CMD%^)

rem ---------------- 3. 创建虚拟环境 ----------------
if exist "%VPY%" goto :venv_ready

echo [虚拟环境] 正在创建 .venv ...
%PY_CMD% -m venv "%VENV_DIR%"
if not exist "%VPY%" goto :venv_failed

:venv_ready
echo [虚拟环境] %VPY%

rem ---------------- 4. 安装依赖（临时加速源 + 自动换源重试） ----------------
call :do_install
if not errorlevel 1 goto :install_ok

if /i not "%PIP_TRUSTED_HOST%"=="mirrors.aliyun.com" (
    echo.
    echo [提示] %MIRROR_NAME% 安装失败，自动改用阿里云源重试...
    set "PIP_INDEX_URL=https://mirrors.aliyun.com/pypi/simple"
    set "PIP_TRUSTED_HOST=mirrors.aliyun.com"
    set "MIRROR_NAME=阿里云"
    call :do_install
    if not errorlevel 1 goto :install_ok
)

if /i not "%PIP_TRUSTED_HOST%"=="pypi.org" (
    echo.
    echo [提示] 仍失败，自动改用官方 PyPI 重试...
    set "PIP_INDEX_URL=https://pypi.org/simple"
    set "PIP_TRUSTED_HOST=pypi.org"
    set "MIRROR_NAME=官方 PyPI"
    call :do_install
    if not errorlevel 1 goto :install_ok
)

goto :install_failed

:do_install
echo.
echo [1/3] 更新 pip（%MIRROR_NAME%）...
"%VPY%" -m pip install --upgrade pip --index-url "%PIP_INDEX_URL%" --trusted-host "%PIP_TRUSTED_HOST%" --timeout 20 --retries 2
if errorlevel 1 echo [提示] pip 升级失败，继续用现有版本安装依赖。

echo.
echo [2/3] 安装依赖：%PKGS%（%MIRROR_NAME%）...
"%VPY%" -m pip install %PKGS% --index-url "%PIP_INDEX_URL%" --trusted-host "%PIP_TRUSTED_HOST%" --timeout 20 --retries 2
if errorlevel 1 exit /b 1

echo.
echo [3/3] 验证依赖 ...
"%VPY%" -c "import PIL, docx; print('Pillow', PIL.__version__); print('python-docx', docx.__version__)"
if errorlevel 1 exit /b 1
exit /b 0

:install_ok
echo.
echo ============================================================
echo   安装完成！依赖已装入 .venv，未改动系统 Python。
echo   接下来：把图片文件夹直接拖到 run_extract.bat 上即可。
echo   （加速源只在本次安装窗口生效，不会写入全局 pip 配置）
echo ============================================================
echo.
pause
exit /b 0

:install_failed
echo.
echo [失败] 依赖安装没有成功。可以尝试：
echo   1. 换一个加速源重新运行本脚本（例如 setup_env.bat aliyun）
echo   2. 关闭代理/VPN 后重试
echo   3. 手动执行：.venv\Scripts\python.exe -m pip install %PKGS%
echo.
pause
exit /b 1

:venv_failed
echo.
echo [失败] 无法创建 .venv 虚拟环境。
echo   通常是 Python 安装不完整（缺少 venv / ensurepip 组件）。
echo   请重新安装官方 Python 3.9+，安装时勾选 "pip" 和 "Add python.exe to PATH"：
echo   https://mirrors.huaweicloud.com/python/3.11.9/python-3.11.9-amd64.exe
echo.
pause
exit /b 1

:no_python
echo.
echo [中止] 未安装 Python，环境准备失败。
echo   请安装 Python 3.9+ 后重新运行本脚本：
echo   https://mirrors.huaweicloud.com/python/3.11.9/python-3.11.9-amd64.exe
echo.
pause
exit /b 1
