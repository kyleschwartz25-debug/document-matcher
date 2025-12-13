@echo off
REM Smart launcher for Document Matcher
REM Checks if Python is installed, launches app, or shows setup instructions

setlocal enabledelayedexpansion

title Document Matcher Launcher

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    cls
    echo.
    echo ====================================================
    echo FIRST TIME SETUP REQUIRED
    echo ====================================================
    echo.
    echo Python is not installed or not in PATH.
    echo.
    echo To fix this:
    echo.
    echo   1. Download Python 3.11 from:
    echo      https://www.python.org/downloads/
    echo.
    echo   2. Run the installer with these settings:
    echo      - CHECK "Add Python to PATH"
    echo      - CHECK "pip"
    echo.
    echo   3. After installation, run this launcher again
    echo.
    echo OR, for automated setup:
    echo.
    echo   - Double-click: SETUP.bat
    echo.
    echo ====================================================
    echo.
    pause
    exit /b 1
)

REM Check if PyPDF2 is installed
python -c "import PyPDF2" >nul 2>&1
if %errorlevel% neq 0 (
    cls
    echo.
    echo Installing required packages...
    echo.
    python -m pip install PyPDF2 --quiet
    if !errorlevel! neq 0 (
        echo Failed to install dependencies.
        echo Please run: SETUP.bat
        pause
        exit /b 1
    )
    echo Dependencies installed successfully.
    echo.
)

REM Launch the application
python "%~dp0document_matcher.py"
