@echo off
chcp 65001 >nul
title 23HG SCHEDULE ASSISTANT PRO - Go cai dat
setlocal
set "ADDIN_DIR=%APPDATA%\Microsoft\AddIns"

echo Dong tat ca cua so Excel truoc khi tiep tuc.
pause

del /f /q "%ADDIN_DIR%\23HG_Schedule_Assistant_Pro.xlam" >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\23HGAddins" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\Excel\Options" /v OPEN /f >nul 2>&1

echo.
echo Da go Add-in 23HG. Ban sao luu (neu co) nam tai %APPDATA%\23HG_Backup
pause
