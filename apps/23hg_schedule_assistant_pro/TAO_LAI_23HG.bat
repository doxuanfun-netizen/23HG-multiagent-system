@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
echo ============================================================
echo  23HG - TAO LAI SAN PHAM (Master + Add-in) bang Excel that
echo ============================================================
echo Can: Python 3, Excel da cai, va trong Excel bat:
echo   File ^> Options ^> Trust Center ^> Trust Center Settings ^>
echo   Macro Settings ^> "Trust access to the VBA project object model"
echo Dong het cua so Excel truoc khi chay.
echo.
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>nul | find /I "EXCEL.EXE" >nul
if not errorlevel 1 (
  echo [!] Excel dang mo. Hay luu va dong Excel roi chay lai.
  pause
  exit /b 1
)
set STAMP=%DATE:~-4%%DATE:~3,2%%DATE:~0,2%_%TIME:~0,2%%TIME:~3,2%
set STAMP=%STAMP: =0%
set BK=3_LuuTru_PhienBanCu\TRUOC_TAO_LAI_%STAMP%
mkdir "%BK%" >nul 2>&1
copy /y "23HG_DU_AN_MAU_TIEN_DO_CHUAN_G1_PRO.xlsm" "%BK%\" >nul
copy /y "23HG_Schedule_Assistant_Pro.xlam" "%BK%\" >nul
echo Da sao luu ban hien tai vao %BK%
echo.
python -m pip install -r requirements.txt
if errorlevel 1 goto :fail
python 1_Scripts_TuDongHoa\rebuild_perfect_master_v3_professional.py
if errorlevel 1 goto :fail
python 1_Scripts_TuDongHoa\audit_master_v3.py
if errorlevel 1 goto :fail
echo.
echo HOAN TAT. Mo Excel va lam theo KIEM_THU_TREN_WINDOWS.md
pause
exit /b 0
:fail
echo.
echo [LOI] Co buoc that bai. Ban cu van con trong %BK%
pause
exit /b 1
