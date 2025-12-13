@echo off
REM Document Matcher Launcher - Run from Command Prompt
REM This batch file launches the SO vs PO comparison tool

echo Starting Document Matcher...
cd /d "%~dp0"

powershell -ExecutionPolicy Bypass -File "DocumentMatcher_v2.ps1"

pause
