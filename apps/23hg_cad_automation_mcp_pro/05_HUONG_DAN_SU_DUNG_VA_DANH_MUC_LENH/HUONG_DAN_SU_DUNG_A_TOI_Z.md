# HƯỚNG DẪN SỬ DỤNG TỪ A ĐẾN Z
# 23HG CAD AUTOMATION MCP PRO (v2.0.0)
**Tác giả:** Nguyễn Bảo Tú (@baotuhg) | **Đơn vị phát triển:** 23HG SYSTEM  
**Phiên bản:** v2.0.0 Pro Localized  

---

## 1. GIỚI THIỆU SẢN PHẨM
**23HG CAD AUTOMATION MCP PRO** là bộ công cụ tích hợp giao thức MCP (Model Context Protocol) chuyên biệt dành cho kỹ sư xây dựng, kiến trúc sư và kỹ sư cơ điện Việt Nam. 

Sản phẩm cho phép các trợ lý AI hàng đầu hiện nay (**Claude Desktop, Cursor, Antigravity, Windsurf**) có thể **trực tiếp nhìn thấy, đọc hiểu và điều khiển AutoCAD** (hoặc GstarCAD/ZWCAD) theo thời gian thực bằng ngôn ngữ tự nhiên thuần Việt 100%.

### Các Tính Năng Đột Phá:
1. **100% Việt Hóa:** Toàn bộ 135 lệnh API, 652 tham số, phản hồi lỗi, prompt hướng dẫn kỹ thuật đều bằng tiếng Việt chuyên ngành AEC.
2. **Cài Đặt Tự Động 1-Click:** Không cần am hiểu kỹ thuật lập trình sâu, chỉ cần nhấp đúp file `.bat`.
3. **Đa Nền Tảng AI:** Hỗ trợ mọi AI Client có giao thức MCP (Claude Desktop, Cursor, Antigravity, Windsurf...).
4. **Hỗ Trợ NETLOAD Cực Nhanh:** Đi kèm plugin .NET `cad_cowork_mcp.dll` giúp thao tác hàng nghìn đối tượng CAD chỉ trong vài mili-giây.
5. **Bộ Lệnh Dim & Ký Hiệu Chuyên Nghiệp:** Tích hợp quy chuẩn Dimstyle chuẩn TCVN, textstyle, hatch, linetype không bị lỗi font.

---

## 2. YÊU CẦU HỆ THỐNG
- **Hệ điều hành:** Windows 10 hoặc Windows 11 (64-bit).
- **AutoCAD:** AutoCAD 2018 - 2026 (hoặc GstarCAD / ZWCAD).
- **Python:** Phiên bản Python 3.10 trở lên (khuyên dùng Python 3.11 hoặc 3.12).  
  *(Lưu ý: Khi cài Python, hãy nhớ tích chọn **"Add Python to PATH"**)*.

---

## 3. CÁC BƯỚC CÀI ĐẶT NHANH (1-CLICK)

### Bước 1: Chạy file cài đặt tự động
1. Vào thư mục `01_CAI_DAT_TU_DONG_1_CLICK`.
2. Nhấp đúp vào file [`CAI_DAT_NHANH_1_CLICK.bat`](file:///C:/Users/baotu/Downloads/23HG_CAD_AUTOMATION_MCP_PRO/01_CAI_DAT_TU_DONG_1_CLICK/CAI_DAT_NHANH_1_CLICK.bat).
3. Đợi khoảng 10 - 20 giây để hệ thống tự động:
   - Cài đặt các thư viện cần thiết (`mcp`, `pywin32`, `pydantic`, `ezdxf`, `openpyxl`).
   - Đăng ký gói phần mềm `cad-cowork-mcp-23hg` vào hệ thống Windows.

### Bước 2: Kiểm tra kết nối CAD
1. Mở phần mềm AutoCAD lên, mở 1 file bản vẽ hoặc tạo bản vẽ mới (`Ctrl + N`).
2. Chạy file [`KHOI_DONG_TEST_MCP.bat`](file:///C:/Users/baotu/Downloads/23HG_CAD_AUTOMATION_MCP_PRO/01_CAI_DAT_TU_DONG_1_CLICK/KHOI_DONG_TEST_MCP.bat).
3. Màn hình báo `✅ KẾT NỐI THÀNH CÔNG!` là xong.

---

## 4. NẠP PLUGIN .NET VÀO AUTOCAD (NETLOAD)
Để tăng tốc tối đa tốc độ vẽ và cho phép AI can thiệp sâu vào cơ sở dữ liệu CAD:
1. Trong cửa sổ dòng lệnh AutoCAD, gõ lệnh: `NETLOAD` và nhấn **Enter**.
2. Duyệt đến thư mục `04_PLUGIN_CAD_NETLOAD` và chọn file [`cad_cowork_mcp.dll`](file:///C:/Users/baotu/Downloads/23HG_CAD_AUTOMATION_MCP_PRO/04_PLUGIN_CAD_NETLOAD/cad_cowork_mcp.dll).
3. Nhấn **Open** (nếu AutoCAD hỏi quyền bảo mật, hãy chọn *Always Load*).

*(Mẹo: Bạn có thể thêm file dll này vào `Startup Suite` trong bảng lệnh `APPLOAD` của AutoCAD để mỗi lần bật CAD nó tự động nạp).*

---

## 5. CẤU HÌNH VÀO PHẦN MỀM AI CỦA BẠN

### Cho CLAUDE DESKTOP:
1. Nhấn `Win + R`, nhập `%APPDATA%\Claude` và Enter.
2. Mở file `claude_desktop_config.json`, chép nội dung từ thư mục `03_FILE_CAU_HINH_CAC_PHAN_MEM_AI`:
```json
{
  "mcpServers": {
    "cad-mcp": {
      "command": "python",
      "args": ["-m", "cad_cowork_mcp"],
      "env": {
        "PYTHONIOENCODING": "utf-8"
      }
    }
  }
}
```
3. Tắt và mở lại Claude Desktop.

### Cho CURSOR:
1. Mở Cursor -> Settings (`Ctrl + ,`) -> **Features** -> **MCP Servers**.
2. Nhấn **+ Add New MCP Server**:
   - **Name:** `cad-mcp`
   - **Type:** `command`
   - **Command:** `python -m cad_cowork_mcp`
3. Nhấn **Save**.

### Cấu Hình "Auto-Pass" (Không hỏi cấp quyền liên tục):
Trong Cursor hoặc Claude, bạn hãy bật chế độ **Auto-approve** cho `cad-mcp` để AI có thể tự động thực hiện chuỗi lệnh vẽ liên tục mà không bắt bạn phải click "Approve" từng lệnh.

---

## 6. CÁC CÂU LỆNH MẪU THỰC CHIẾN (PROMPT EXAMPLES)
Sau khi kết nối thành công, bạn chỉ cần ra lệnh cho AI bằng tiếng Việt thông thường:

### Ví dụ 1: Vẽ mặt bằng kết cấu dầm cột
> *"Hãy kiểm tra bản vẽ hiện hành trong AutoCAD. Tạo cho tôi các layer: `KC_COT` (màu đỏ, nét liền), `KC_DAM` (màu xanh lục), `KC_DIM` (màu 8). Sau đó vẽ hệ lưới trục A-B-C cách nhau 4.2m và 1-2-3 cách nhau 5.0m, đặt các cột 220x220 tại các nút giao."*

### Ví dụ 2: Tự động ghi kích thước (Dim) chuẩn kỹ thuật
> *"Hãy đo kích thước toàn bộ hệ trục trên bản vẽ. Dùng kiểu dim chuyên nghiệp có gạch chéo 45 độ (Oblique), chân dim cách đối tượng 100mm, đường gióng nhô 100mm, số dim nằm trên đường dóng."*

### Ví dụ 3: Bóc tách khối lượng và thống kê
> *"Hãy đếm số lượng block cửa `D1`, `D2` và block cột `C1` đang có trong bản vẽ, tổng hợp thành bảng thống kê số lượng và xuất ra màn hình cho tôi."*

### Ví dụ 4: Xoay bản vẽ và Zoom toàn bộ
> *"Hãy Zoom Extents toàn bộ bản vẽ và lưu lại file dưới tên `MB_TANG_1_FINAL.dwg`."*

---

## 7. XỬ LÝ LỖI THƯỜNG GẶP (TROUBLESHOOTING)

| Hiện tượng | Nguyên nhân | Cách khắc phục |
| :--- | :--- | :--- |
| **Không tìm thấy tiến trình CAD** | AutoCAD chưa bật hoặc chưa mở bản vẽ nào | Bật AutoCAD lên và mở/tạo 1 bản vẽ mới (.dwg). |
| **Lỗi UIPI / Quyền Admin** | AutoCAD bật bằng *Run as Administrator* nhưng AI/Python chạy quyền User | Tắt AutoCAD đi và mở lại bình thường (không chọn Run as Admin), hoặc cho cả hai cùng chạy quyền Admin. |
| **Chưa có thư viện pywin32** | Python thiếu thư viện giao tiếp COM Windows | Chạy lại file `CAI_DAT_NHANH_1_CLICK.bat`. |
| **Lỗi font tiếng Việt trong CAD** | Bản vẽ đang dùng font cũ TCVN3 / VNI | Chọn TextStyle chuẩn Arial hoặc Times New Roman Unicode trong CAD. |

---

## 8. HỖ TRỢ VÀ BẢN QUYỀN
- **Đóng gói phát triển:** Nguyễn Bảo Tú (@baotuhg) - 23HG SYSTEM.
- **Hỗ trợ kỹ thuật nội bộ:** Liên hệ trực tiếp qua Zalo nhóm 23HG.
- Mọi thắc mắc hoặc yêu cầu bổ sung lệnh API mới, anh em liên hệ trực tiếp để được update phiên bản mới nhất!
