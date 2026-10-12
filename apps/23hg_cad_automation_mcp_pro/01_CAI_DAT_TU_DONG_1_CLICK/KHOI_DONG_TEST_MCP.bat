@echo off
chcp 65001 >nul
cls
echo ===============================================================================
echo        23HG CAD AUTOMATION MCP PRO - KIEM TRA KET NOI VA CHAY THU
echo               Tac gia: Nguyen Bao Tu (@baotuhg) ^| 23HG SYSTEM
echo ===============================================================================
echo.

:: 1. Kiem tra Python
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [LOI] Khong tim thay Python! Vui long chay file 'CAI_DAT_NHANH_1_CLICK.bat' truoc.
    echo.
    pause
    exit /b 1
)

:: 2. Chay test ket noi CAD
echo [*] Dang kiem tra ket noi voi AutoCAD / GstarCAD / ZWCAD...
echo.
cd /d "%~dp0..\02_MA_NGUON_VIET_HOA_100_PHAN_TRAM"
python cad_cowork_mcp\test_connect.py

echo.
echo ===============================================================================
echo Huong dan:
echo - Neu hien 'KET NOI THANH CONG': Ban da san sang su dung MCP voi Claude/Cursor/Antigravity!
echo - Neu chua ket noi duoc: Hay mo AutoCAD len truoc roi chay lai file bat nay nhe.
echo ===============================================================================
echo.
pause
