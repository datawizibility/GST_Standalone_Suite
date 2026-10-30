@echo off
title GST Standalone Suite
echo ========================================================
echo   Starting GST Standalone Intelligence Suite...
echo   Offline & Secure - Zero-Installation
echo   Your browser will open automatically at:
echo   http://127.0.0.1:5055
echo ========================================================
cd /d "%~dp0"

if exist "dist\GST_Audit_Suite.exe" (
    echo Launching standalone binary...
    start "" "dist\GST_Audit_Suite.exe"
) else (
    echo Launching via Python runtime...
    python app.py
)
