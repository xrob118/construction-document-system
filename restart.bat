@echo off
chcp 65001 >nul 2>&1
title Construction Management System - Restart

echo ================================================
echo   Construction Management System - Quick Restart
echo ================================================
echo.

echo [1/4] Stopping old backend process...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":8080 " ^| findstr "LISTENING"') do (
    echo   Found process PID=%%a, terminating...
    taskkill /F /PID %%a >nul 2>&1
)
taskkill /F /IM uvicorn.exe >nul 2>&1
echo   Backend process stopped.
echo.

echo [2/4] Building frontend...
cd /d "%~dp0frontend"
call npm run build
if %errorlevel% neq 0 (
    echo.
    echo   [ERROR] Frontend build failed!
    pause
    exit /b 1
)
echo   Frontend build completed.
echo.

echo [3/4] Updating frontend files...
if exist "%~dp0backend\static" rmdir /s /q "%~dp0backend\static"
xcopy /e /i /q /y "%~dp0frontend\dist" "%~dp0backend\static" >nul
echo   Frontend files updated.
echo.

echo [4/4] Starting backend service...
cd /d "%~dp0backend"
start "Construction Backend" cmd /k python -m uvicorn main:app --host 127.0.0.1 --port 8080

echo   Waiting for service to start...
timeout /t 5 /nobreak >nul

echo   Opening frontend...
start http://localhost:8080

echo.
echo ================================================
echo   All services started!
echo   Backend running in separate window
echo   Frontend opened in browser
echo ================================================
echo.
echo   This window will close in 5 seconds...
timeout /t 5 /nobreak >nul
