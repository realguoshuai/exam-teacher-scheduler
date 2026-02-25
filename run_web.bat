@echo off
title Exam Teacher Scheduler - Web Version
color 0C

cls
echo ================================================================================
echo Exam Teacher Scheduler - Web Version
echo ================================================================================
echo.

cd /d "%~dp0"

echo Checking Python environment...
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not detected
    echo Please run install_web.bat first
    pause
    exit /b 1
)

echo [SUCCESS] Python environment OK
echo.

echo Installing web dependencies...
pip install -r requirements_web.txt

if errorlevel 1 (
    echo.
    echo [ERROR] Web dependencies installation failed
    pause
    exit /b 1
)

echo.
echo [SUCCESS] Dependencies installed
echo.
echo ================================================================================
echo Starting web server...
echo ================================================================================
echo.
echo The web interface will be available at: http://localhost:5000
echo Press Ctrl+C to stop the server
echo.

python app.py

if errorlevel 1 (
    echo.
    echo [ERROR] Server startup failed
    pause
)
