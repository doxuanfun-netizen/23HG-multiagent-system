@echo off
chcp 65001 >nul
title 23HG SCHEDULE ASSISTANT PRO - Cai dat
color 1F
setlocal

set "HERE=%~dp0"
set "XLSM=%HERE%23HG_DU_AN_MAU_TIEN_DO_CHUAN_G1_PRO.xlsm"
set "XLAM=%HERE%23HG_Schedule_Assistant_Pro.xlam"
set "ADDIN_DIR=%APPDATA%\Microsoft\AddIns"
set "XLSTART_DIR=%APPDATA%\Microsoft\Excel\XLSTART"
set "BACKUP_DIR=%APPDATA%\23HG_Backup"
set "LOG=%BACKUP_DIR%\cai_dat.log"

if not exist "%BACKUP_DIR%" mkdir "%BACKUP_DIR%"
call :log "Bat dau cai dat tu %HERE%"

if not exist "%XLAM%" ( echo [LOI] Khong tim thay %XLAM% & goto :fail )
if not exist "%XLSM%" ( echo [LOI] Khong tim thay %XLSM% & goto :fail )

echo ===============================================================================
echo   23HG SCHEDULE ASSISTANT PRO - CAI DAT ADD-IN EXCEL
echo   Tac gia: NBT (Nguyen Bao Tu) ^| @baotuhg ^| 23HG SYSTEM
echo ===============================================================================
echo.

echo [1/4] Kiem tra Microsoft Excel...
tasklist /fi "imagename eq EXCEL.EXE" 2>nul | find /i "EXCEL.EXE" >nul
if not errorlevel 1 (
    echo.
    echo   Excel dang mo. Hay LUU cong viec va DONG tat ca cua so Excel.
    choice /c CT /n /m "  Nhan C de tiep tuc khi da dong Excel, T de thoat: "
    if errorlevel 2 ( echo Da huy cai dat. & goto :end )
    tasklist /fi "imagename eq EXCEL.EXE" 2>nul | find /i "EXCEL.EXE" >nul
    if not errorlevel 1 ( echo [LOI] Excel van dang mo. & goto :fail )
)

echo [2/4] Sao luu va cai dat Add-in...
if exist "%XLSTART_DIR%\23HG_Schedule_Assistant_Pro.xlam" (
    copy /y "%XLSTART_DIR%\23HG_Schedule_Assistant_Pro.xlam" "%BACKUP_DIR%\" >nul
    if errorlevel 1 ( echo [LOI] Khong sao luu duoc ban cu trong XLSTART & goto :fail )
    del /f /q "%XLSTART_DIR%\23HG_Schedule_Assistant_Pro.xlam" >nul 2>&1
    call :log "Da sao luu va go ban trung trong XLSTART"
)
if not exist "%ADDIN_DIR%" mkdir "%ADDIN_DIR%"
copy /y "%XLAM%" "%ADDIN_DIR%\" >nul
if errorlevel 1 ( echo [LOI] Khong copy duoc Add-in vao %ADDIN_DIR% & goto :fail )
call :log "Da copy Add-in vao %ADDIN_DIR%"

echo [3/4] Cau hinh Registry (Trusted Location va tu dong nap Add-in)...
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\23HGAddins" /v Path /t REG_SZ /d "%ADDIN_DIR%" /f >nul
if errorlevel 1 ( echo [LOI] Khong ghi duoc Trusted Location & goto :fail )
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Options" /v OPEN /t REG_SZ /d "\"23HG_Schedule_Assistant_Pro.xlam\"" /f >nul
if errorlevel 1 ( echo [LOI] Khong ghi duoc cau hinh tu dong nap & goto :fail )
call :log "Da cau hinh Registry"

echo [4/4] Mo file mau Master PRO...
start "" "%XLSM%"

call :log "Cai dat thanh cong"
echo.
echo ===============================================================================
echo   CAI DAT THANH CONG. Thanh Ribbon "23HG SCHEDULE ASSISTANT" da san sang.
echo   Nhat ky: %LOG%
echo ===============================================================================
goto :end

:fail
call :log "Cai dat that bai"
echo.
echo   Cai dat KHONG thanh cong. Xem nhat ky de biet chi tiet: %LOG%
exit /b 1

:end
echo.
pause
exit /b 0

:log
echo [%date% %time%] %~1>>"%LOG%"
exit /b 0
