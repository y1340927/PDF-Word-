@echo off
title PDF to Word Converter
cd /d "%~dp0"

echo ========================================================
echo          PDF to Word Converter
echo ========================================================
echo.

echo [1/3] Cleaning old processes...
taskkill /f /im python.exe >nul 2>&1
echo        Done.
echo.

echo [2/3] Checking Python environment...
call venv\Scripts\activate.bat >nul 2>&1
if errorlevel 1 (
    echo.
    echo [ERROR] Virtual environment not found!
    echo   Please create it: python -m venv venv
    echo.
    pause
    exit /b 1
)
echo        Python environment OK.
echo.

echo [3/3] Starting Flask server...
echo.
echo ========================================================
echo   URL: http://127.0.0.1:5000
echo   Close this window to stop server.
echo ========================================================
echo.

start /b cmd /c "ping -n 4 127.0.0.1 >nul & start http://127.0.0.1:5000"

python app.py

echo.
echo ========================================================
echo   Server stopped. Press any key to exit...
echo ========================================================
pause >nul
