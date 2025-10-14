@echo off
REM Desktop Management Mock Backend - Quick Start Script
REM Starts the mock backend server on port 80

echo ========================================
echo Desktop Management Mock Backend Server
echo ========================================
echo.

REM Check if running as Administrator
net session >nul 2>&1
if %errorLevel% == 0 (
    echo [OK] Running as Administrator
) else (
    echo [ERROR] This script requires Administrator privileges!
    echo Please right-click and select "Run as Administrator"
    echo.
    pause
    exit /b 1
)

echo.
echo [INFO] Starting Flask server on port 80...
echo [INFO] Configure DNS: gdpmappercb.nomura.com -^> 127.0.0.1
echo [INFO] Or edit hosts file: C:\Windows\System32\drivers\etc\hosts
echo.
echo Press Ctrl+C to stop the server
echo.

python mock_backend.py

pause

