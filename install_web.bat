@echo off
title Exam Teacher Scheduler - Install Web Dependencies
color 0A

echo ================================================================================
echo Exam Teacher Scheduler - Install Web Dependencies
echo ================================================================================
echo.

cd /d "%~dp0"

echo Checking Python environment...
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not detected. Please install Python 3.8 or higher.
    echo.
    echo Download: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [SUCCESS] Python environment OK
echo.

echo Installing web dependencies...
pip install -r requirements_web.txt

if errorlevel 1 (
    echo.
    echo [ERROR] Dependency installation failed
    pause
    exit /b 1
)

echo.
echo ================================================================================
echo [SUCCESS] Web dependencies installed!
echo ================================================================================
echo.
echo You can now:
echo   1. Run web version: run_web.bat
echo   2. Run CLI version: run.bat
echo.
pause
