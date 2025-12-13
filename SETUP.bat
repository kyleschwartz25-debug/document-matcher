@echo off
REM Setup Script for Document Matcher
REM This script installs Python and required dependencies

echo ============================================
echo Document Matcher - Setup Wizard
echo ============================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo.
    echo SOLUTION: Please install Python 3.11+ from https://www.python.org/downloads/
    echo.
    echo During installation, MAKE SURE to:
    echo   1. Check "Add Python to PATH"
    echo   2. Check "pip" during custom installation
    echo.
    echo After installation, close this window and run this script again.
    pause
    exit /b 1
)

echo Found Python:
python --version
echo.

REM Install required packages
echo Installing required packages...
echo This may take a few minutes on first run.
echo.

python -m pip install --upgrade pip
python -m pip install PyPDF2

if %errorlevel% neq 0 (
    echo.
    echo ERROR: Failed to install dependencies
    pause
    exit /b 1
)

echo.
echo ============================================
echo Setup Complete!
echo ============================================
echo.
echo You can now run the Document Matcher:
echo   Double-click: document_matcher.py
echo   Or from command line: python document_matcher.py
echo.
pause
