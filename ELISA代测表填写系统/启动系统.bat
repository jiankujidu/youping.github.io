@echo off
chcp 65001 >nul
title ELISA代测表填写系统

echo ========================================
echo    ELISA代测表填写系统
echo ========================================
echo.

REM 获取当前目录
set "CURRENT_DIR=%~dp0"
cd /d "%CURRENT_DIR%"

echo [1/2] 启动后端服务...
echo 当前目录: %CURRENT_DIR%

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误：未找到Python，请先安装Python 3.x
    pause
    exit /b 1
)

REM 检查模板文件是否存在
if not exist "优品Elisa代测表.xlsx" (
    echo 错误：找不到Excel模板文件"优品Elisa代测表.xlsx"
    echo 请确保该文件与此启动脚本在同一目录
    pause
    exit /b 1
)

REM 启动Python服务器
start "ELISA后端服务" python "%CURRENT_DIR%server.py"

REM 等待服务启动
echo 等待服务启动...
timeout /t 3 >nul

echo [2/2] 打开填写页面...
start "" "%CURRENT_DIR%index.html"

echo.
echo ========================================
echo    系统已启动！
echo    请在浏览器中填写表单后点击导出
echo    关闭此窗口将停止服务
echo ========================================
echo.

REM 保持窗口打开
pause
