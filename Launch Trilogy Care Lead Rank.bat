@echo off
title Trilogy Care Lead Rank — Launcher

echo.
echo  ============================================
echo   Trilogy Care Lead Rank — Starting up...
echo  ============================================
echo.

:: Check Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Python is not installed or not found.
    echo.
    echo  Please download and install Python from:
    echo  https://www.python.org/downloads/
    echo.
    echo  Make sure to tick "Add Python to PATH" during install.
    echo.
    pause
    exit /b 1
)

echo  Python found. Checking dependencies...
echo.

:: Install pandas silently if not already installed
python -c "import pandas" >nul 2>&1
if errorlevel 1 (
    echo  Installing pandas ^(one-time setup^)...
    pip install pandas --quiet
)

echo  Launching app...
echo.

:: Run the app from this script's own folder
cd /d "%~dp0"
python "trilogy_care_lead_rank.py"

if errorlevel 1 (
    echo.
    echo  Something went wrong. Press any key to see the error details.
    pause
)
