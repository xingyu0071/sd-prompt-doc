@echo off
chcp 65001 >nul
setlocal EnableExtensions
title SD WebUI(Aki) 提示词图集 - 生成文档

rem ============================================================
rem  用法：
rem    1) 把图片文件夹直接拖到本 bat 图标上
rem    2) 或双击运行，按提示输入文件夹路径
rem  可选参数直接跟在后面，例如：
rem    run_extract.bat "D:\输出图" --page-break --only-aki
rem ============================================================

set "SCRIPT_DIR=%~dp0"
set "VPY=%SCRIPT_DIR%.venv\Scripts\python.exe"
set "PY_SCRIPT=%SCRIPT_DIR%sd_aki_prompt_doc.py"

set "TARGET=%~1"
if not defined TARGET goto :ask_path
rem %* 不受 shift 影响，直接原样保存全部参数（第一个就是文件夹）
set "EXTRA=%*"
goto :check_env

:ask_path
echo.
set /p "TARGET=请把图片文件夹拖到这里，或输入路径（不含引号）: "
if not defined TARGET goto :bad_path

:check_env
rem 去掉用户可能带上的首尾引号
set "TARGET=%TARGET:"=%"

if not exist "%TARGET%" goto :bad_path
if not exist "%PY_SCRIPT%" (
    echo [错误] 找不到 sd_aki_prompt_doc.py，请确认它和本 bat 在同一文件夹。
    pause
    exit /b 1
)

if exist "%VPY%" goto :run

echo [提示] 还没有安装运行环境，先调用 setup_env.bat ...
call "%SCRIPT_DIR%setup_env.bat"
if not exist "%VPY%" goto :no_env

:run
echo.
echo [运行] 扫描文件夹：%TARGET%
echo.
if defined EXTRA (
    "%VPY%" "%PY_SCRIPT%" %EXTRA%
) else (
    "%VPY%" "%PY_SCRIPT%" "%TARGET%"
)
set "RC=%ERRORLEVEL%"
echo.
if not "%RC%"=="0" (
    echo [结束] 脚本返回错误码 %RC%，请查看上面的提示。
    pause
    exit /b %RC%
)

echo [完成] 文档已生成。
if /i not "%NO_OPEN%"=="1" (
    echo        正在打开图片文件夹...
    start "" explorer "%TARGET%"
)
echo.
pause
exit /b 0

:bad_path
echo.
echo [错误] 文件夹不存在：%TARGET%
echo        请检查路径是否输入正确（可以右键文件夹 - 复制文件地址）。
echo.
pause
exit /b 1

:no_env
echo.
echo [错误] 运行环境仍未准备好，请先双击 setup_env.bat 完成安装。
pause
exit /b 1
