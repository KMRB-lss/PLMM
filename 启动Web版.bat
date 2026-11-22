@echo off
chcp 65001 >nul
title 批量文件重命名工具 · Web 版
cd /d "%~dp0"

echo ============================================
echo   批量文件重命名工具 · Web 控制台
echo ============================================
echo.

REM 依次尝试可用的 Python：PATH 里的 python -> py 启动器 -> 常见安装目录
set "PYTHON="
set "PF86=%ProgramFiles(x86)%"
where python >nul 2>nul && set "PYTHON=python"
if not defined PYTHON (
  where py >nul 2>nul && set "PYTHON=py -3"
)
if not defined PYTHON (
  for /d %%D in ("%LocalAppData%\Programs\Python\Python3*" "%LocalAppData%\Python\python*" "%ProgramFiles%\Python3*" "%PF86%\Python3*" "C:\Python3*") do (
    if not defined PYTHON if exist "%%~fD\python.exe" set "PYTHON=%%~fD\python.exe"
  )
)
if not defined PYTHON (
  echo [错误] 未找到 Python，请先安装 Python 3.9+（安装时勾选 Add python.exe to PATH）后重试
  echo        下载地址: https://www.python.org/downloads/
  pause
  exit /b 1
)

echo 使用 Python: %PYTHON%
echo 服务启动后自动打开浏览器: http://localhost:8890
echo 按 Ctrl+C 停止服务
echo.

%PYTHON% server.py 8890
if errorlevel 1 echo [提示] 程序非正常退出，请查看上方错误信息
pause
