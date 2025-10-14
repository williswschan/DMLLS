@echo off
REM Add mock backend entry to Windows hosts file
REM Run as Administrator

echo ========================================
echo Configure Hosts File for Mock Backend
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

set HOSTS_FILE=C:\Windows\System32\drivers\etc\hosts

echo [INFO] Backing up current hosts file...
copy "%HOSTS_FILE%" "%HOSTS_FILE%.backup_%date:~-4,4%%date:~-10,2%%date:~-7,2%" >nul

echo [INFO] Adding entry for gdpmappercb.nomura.com...

REM Check if entry already exists
findstr /C:"gdpmappercb.nomura.com" "%HOSTS_FILE%" >nul
if %errorLevel% == 0 (
    echo [WARNING] Entry already exists in hosts file
    echo Current entry:
    findstr /C:"gdpmappercb.nomura.com" "%HOSTS_FILE%"
    echo.
    echo Do you want to replace it? (Y/N)
    choice /C YN /N
    if errorLevel 2 goto :skip
    
    REM Remove old entry
    findstr /V /C:"gdpmappercb.nomura.com" "%HOSTS_FILE%" > "%HOSTS_FILE%.tmp"
    move /Y "%HOSTS_FILE%.tmp" "%HOSTS_FILE%" >nul
)

REM Add new entry
echo 127.0.0.1    gdpmappercb.nomura.com >> "%HOSTS_FILE%"
echo.
echo [OK] Added: 127.0.0.1    gdpmappercb.nomura.com
echo.
goto :done

:skip
echo [INFO] Skipped modification
echo.

:done
echo [INFO] Flushing DNS cache...
ipconfig /flushdns >nul

echo.
echo ========================================
echo Configuration Complete!
echo ========================================
echo.
echo Test DNS resolution:
echo   nslookup gdpmappercb.nomura.com
echo.
echo Should return: 127.0.0.1
echo.
pause

