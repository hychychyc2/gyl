@echo off
REM 采购订单自动化脚本启动器
REM 每天晚上10点运行

echo ========================================
echo 采购订单自动化脚本
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: Python未安装或未添加到PATH
    echo 请从 https://www.python.org/downloads/ 下载安装
    pause
    exit /b 1
)

REM 设置工作目录
cd /d "%~dp0"

REM 运行脚本
echo 正在执行采购订单自动化...
python purchase_order_automation.py

echo.
echo 执行完成，按任意键退出...
pause >nul