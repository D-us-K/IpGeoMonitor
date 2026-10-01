@echo off
chcp 65001 >nul
title IP Geo Monitor - Debug Console
echo ============================================
echo   IP Geo Monitor - debug launch
echo   (keep this window open to see output)
echo ============================================
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0IpGeoMonitor.ps1" %*
echo.
echo --------------------------------------------
echo Process exited with code: %ERRORLEVEL%
echo --------------------------------------------
pause
