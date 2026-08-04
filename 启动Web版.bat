@echo off
chcp 65001 >nul
title 批量文件重命名工具 · Web 版
cd /d "%~dp0"

echo ============================================
echo   批量文件重命名工具 · Web 控制台
echo ============================================
echo.

REM 依次尝试可用的 Python 解释器
set "PYTHON="
where python >nul 2>nul && set "PYTHON=python"
if not defined PYTHON (
  where py >nul 2>nul && set "PYTHON=py -3"
)
if not defined PYTHON (
  if exist "C:\Users\Lenovo\AppData\Local\Python\pythoncore-3.14-64\python.exe" set "PYTHON=C:\Users\Lenovo\AppData\Local\Python\pythoncore-3.14-64\python.exe"
)
if not defined PYTHON (
  if exist "C:\Python314\python.exe" set "PYTHON=C:\Python314\python.exe"
)
if not defined PYTHON (
  echo [错误] 未找到 Python，请先安装 Python 3.8+ 或修改本脚本中的路径
  pause
  exit /b 1
)

echo 使用 Python: %PYTHON%
echo 启动后浏览器访问: http://localhost:8890
echo 按 Ctrl+C 停止服务
echo.

%PYTHON% server.py 8890
pause
