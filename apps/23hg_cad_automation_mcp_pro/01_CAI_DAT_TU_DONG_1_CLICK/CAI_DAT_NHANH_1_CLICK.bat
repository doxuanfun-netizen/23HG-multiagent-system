@echo off
chcp 65001 >nul
title Cài Đặt Tự Động MCP CAD 23HG SYSTEM - Nguyễn Bảo Tú
color 0A

echo ================================================================================
echo   BỘ CÔNG CỤ TỰ ĐỘNG HÓA CAD VÀ TRỢ LÝ AI ĐIỀU KHIỂN BẢN VẼ - 23HG SYSTEM
echo   Tác giả ^& Bản quyền: Nguyễn Bảo Tú (@baotuhg)
echo   Phiên bản: 2.0.0 Pro (100%% Tiếng Việt Chuyên Ngành AEC)
echo ================================================================================
echo.

echo [1/4] Đang kiểm tra môi trường Python trên máy...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    color 0C
    echo [LỖI] Không tìm thấy Python!
    echo Vui lòng tải và cài đặt Python 3.10+ từ https://www.python.org/downloads/
    echo LƯU Ý QUAN TRỌNG: Hãy tích chọn "Add python.exe to PATH" khi cài đặt!
    echo.
    pause
    exit /b 1
)

python -c "import sys; print(f'Đã tìm thấy: Python {sys.version.split()[0]}')"
echo.

echo [2/4] Đang cài đặt các thư viện phụ thuộc (mcp, pydantic, pywin32, ezdxf, openpyxl)...
python -m pip install --upgrade pip >nul 2>&1
python -m pip install -r "%~dp0..\02_MA_NGUON_VIET_HOA_100_PHAN_TRAM\requirements.txt"
if %errorlevel% neq 0 (
    color 0C
    echo [LỖI] Cài đặt thư viện thất bại! Vui lòng kiểm tra kết nối mạng.
    pause
    exit /b 1
)
echo -> Cài đặt thư viện thành công!
echo.

echo [3/4] Đang cài đặt gói lõi 23HG CAD MCP vào hệ thống...
cd /d "%~dp0..\02_MA_NGUON_VIET_HOA_100_PHAN_TRAM"
python -m pip install -e .
if %errorlevel% neq 0 (
    color 0C
    echo [LỖI] Cài đặt gói lõi thất bại!
    pause
    exit /b 1
)
echo -> Cài đặt gói lõi 23HG CAD MCP thành công!
echo.

echo [4/4] Đang tự động cấu hình tích hợp AI...
set "CLAUDE_CONFIG=%APPDATA%\Claude\claude_desktop_config.json"
if exist "%APPDATA%\Claude" (
    echo Đã phát hiện Claude Desktop! Mẫu cấu hình sẵn có trong thư mục 03.
)

echo.
echo ================================================================================
echo   CHÚC MỪNG BẠN ĐÃ CÀI ĐẶT THÀNH CÔNG 100%%!
echo   Hệ thống MCP CAD 23HG SYSTEM đã sẵn sàng phục vụ bạn:
echo   - 139 lệnh CAD API Việt hóa hoàn toàn (Vẽ hình, Dim TCVN, Khối lượng, Layer)
echo   - Tương thích tốt với AutoCAD, GstarCAD, ZWCAD
echo   - Hỗ trợ kết nối với Claude Desktop, Cursor, Windsurf, Antigravity
echo ================================================================================
echo.
pause
