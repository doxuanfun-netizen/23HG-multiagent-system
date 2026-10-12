# HƯỚNG DẪN CẤU HÌNH MCP CHO CÁC PHẦN MỀM AI
**23HG CAD AUTOMATION MCP PRO (v2.0.0)**  
*Tác giả: Nguyễn Bảo Tú (@baotuhg) | 23HG SYSTEM*

---

## 1. CLAUDE DESKTOP
1. Nhấn tổ hợp phím `Win + R`, dán đường dẫn sau và nhấn Enter:
   ```text
   %APPDATA%\Claude
   ```
2. Mở file `claude_desktop_config.json` (nếu chưa có thì tạo mới).
3. Sao chép nội dung từ file [claude_desktop_config.json](file:///C:/Users/baotu/Downloads/23HG_CAD_AUTOMATION_MCP_PRO/03_FILE_CAU_HINH_CAC_PHAN_MEM_AI/claude_desktop_config.json) vào:
   ```json
   {
     "mcpServers": {
       "cad-mcp": {
         "command": "python",
         "args": [
           "-m",
           "cad_cowork_mcp"
         ],
         "env": {
           "PYTHONIOENCODING": "utf-8"
         }
       }
     }
   }
   ```
4. Khởi động lại Claude Desktop. Biểu tượng búa (Tools) sẽ xuất hiện 5 công cụ CAD tự động.

---

## 2. CURSOR IDE
### Cách 1: Cài đặt trực tiếp trên giao diện
1. Mở Cursor -> Vào **Settings** (`Ctrl + ,`) -> Chọn tab **Features** -> **MCP Servers**.
2. Nhấn **+ Add New MCP Server**:
   - **Name**: `cad-mcp`
   - **Type**: `command`
   - **Command**: `python -m cad_cowork_mcp`
3. Nhấn Save, trạng thái sẽ hiện chấm xanh **Connected**.

### Cách 2: File cấu hình dự án
Tạo thư mục `.cursor` ở thư mục gốc dự án và tạo file `mcp.json` với nội dung từ [cursor_mcp_config.json](file:///C:/Users/baotu/Downloads/23HG_CAD_AUTOMATION_MCP_PRO/03_FILE_CAU_HINH_CAC_PHAN_MEM_AI/cursor_mcp_config.json).

---

## 3. ANTIGRAVITY IDE / AGY CLI
- Khi đã chạy cài đặt `CAI_DAT_NHANH_1_CLICK.bat`, thư viện `cad_cowork_mcp` đã được liên kết trực tiếp với Python của hệ thống.
- Antigravity sẽ tự động nhận diện MCP server `cad-mcp` thông qua cấu hình trong `antigravity_mcp_config.json`.

---

## 4. WINDSURF (CODEIUM)
1. Mở đường dẫn:
   ```text
   %USERPROFILE%\.codeium\windsurf\mcp_config.json
   ```
2. Thêm cấu hình từ [windsurf_mcp_config.json](file:///C:/Users/baotu/Downloads/23HG_CAD_AUTOMATION_MCP_PRO/03_FILE_CAU_HINH_CAC_PHAN_MEM_AI/windsurf_mcp_config.json).
3. Khởi động lại Windsurf Cascade.

---

## 5. MẸO QUAN TRỌNG: BẬT AUTO-PASS (KHÔNG CẦN BẤM DUYỆT LỆNH)
Khi AI vẽ CAD bằng hàng chục lệnh liên tục, việc mỗi lệnh AI lại hiện hộp thoại hỏi "Allow / Approve" sẽ làm gián đoạn trải nghiệm.
- **Trong Cursor:** Vào Settings -> Features -> MCP -> Tích chọn **Auto-approve** cho `cad-mcp`.
- **Trong Claude Desktop:** Có thể cấu hình `"autoApprove": ["cad-mcp"]` trong file config.
- **Trong Antigravity:** Thêm câu nhắc: `"Từ nay các thao tác điều khiển AutoCAD tự động chấp thuận (Auto-pass) không cần hỏi lại"`.
