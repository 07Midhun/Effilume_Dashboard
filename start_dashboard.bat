@echo off
title EFFILUME Data Consolidation Dashboard
echo ========================================================
echo   EFFILUME Automated Excel Consolidation Dashboard
echo ========================================================
echo.
cd /d "%~dp0"

echo Checking Python environment...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not found in PATH. Please install Python 3.10+ or check PATH.
    pause
    exit /b 1
)

echo Starting server on http://localhost:8000...
echo (Press Ctrl+C to stop the dashboard)
echo.

python run_server.py
pause
