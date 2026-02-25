@echo off
chcp 65001 >nul
echo ========================================
echo 监考老师排班系统 - 启动程序
echo ========================================
echo.

REM 获取exe所在目录
cd /d "%~dp0"

REM 检查exe是否存在
if not exist "监考老师排班系统.exe" (
    echo [错误] 未找到程序文件：监考老师排班系统.exe
    echo 请确保文件在正确的目录下
    pause
    exit /b 1
)

REM 启动程序
echo 正在启动监考老师排班系统...
echo 浏览器将自动打开 http://localhost:5000
echo.

REM 等待2秒让程序启动
timeout /t 2 /nobreak >nul

REM 尝试打开浏览器
start http://localhost:5000

REM 运行exe
"监考老师排班系统.exe"

pause
