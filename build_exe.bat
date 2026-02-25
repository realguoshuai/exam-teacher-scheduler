@echo off
title 监考老师排班系统 - 打包工具
echo.
echo ========================================
echo 监考老师排班系统 - 打包工具
echo ========================================
echo.
echo 正在执行打包脚本...
echo.

cd /d "%~dp0"

REM 使用Python执行打包脚本
python build_package.py

if %errorlevel% neq 0 (
    echo.
    echo [错误] 打包失败
    pause
) else (
    echo.
    echo ========================================
    echo 打包完成！
    echo ========================================
    echo.
    echo 输出目录: dist_release
    echo.
    pause
)
