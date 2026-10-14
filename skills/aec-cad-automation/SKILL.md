---
name: aec-cad-automation
description: Tự động hóa điều khiển AutoCAD/CAD trực tiếp qua giao thức MCP (autocad-live-mcp), bóc tách khối lượng hình học bê tông, ván khuôn, cốt thép BBS và đối soát kích thước thiết kế từ bản vẽ DWG/DXF không cần xuất file trung gian. Hỗ trợ giải mã font TCVN3/VNI sang Unicode và xuất bảng tính Excel 100% công thức sống.
---

# Tự động hóa CAD & Đo bóc Khối lượng Toàn diện (AEC-CAD Takeoff)
**Chủ quyền hệ thống & Tác giả:** Nguyễn Bảo Tú (@baotuhg) - 23HG SYSTEM  
**Đặc quyền:** Bộ công cụ và kỹ năng độc quyền được xây dựng thuần túy trên nền tảng Python COM Interop & ezdxf, tối ưu hóa riêng biệt cho môi trường làm việc cá nhân của Nguyễn Bảo Tú.

## 1. Mục tiêu & Kiến Trúc Hệ Thống MCP 3 Thành Phần

```text
[ Giao diện AI: Claude Desktop / Windsurf / Cursor / Antigravity ]
                               │
                               ▼ (Giao thức MCP: autocad-live-mcp)
               [ MCP Server điều khiển AutoCAD (Local) ]
                               │
                               ▼ (API / COM Interop: win32com / ezdxf)
      [ Bản vẽ DWG trong AutoCAD ] <───> [ Hồ sơ thiết kế (Excel/PDF) ]
```

Cung cấp khả năng cho AI Agent tương tác trực tiếp 2 chiều với phần mềm AutoCAD (AutoCAD 2026/GstarCAD) thông qua giao thức MCP:
- **Đọc thông tin bản vẽ đang mở:** Danh sách layer, blocks, text (tự động giải mã font TCVN3 `.VnTime` sang Unicode), dim, polyline.
- **Khám sức khỏe bản vẽ (Pre-flight Inspection):** Quét đơn vị `$INSUNITS`, phát hiện lỗi biến tỷ lệ `DIMLFAC != 1.0` (chống lỗi co giãn kích thước), rà soát cao độ 3D ($Z \ne 0$) và trùng nét.
- **Đo bóc Khối lượng Bê tông & Ván khuôn:** Tự động tính toán diện tích mặt cắt dầm, thể tích cọc khoan nhồi, bệ móng, thân mố trụ, xà mũ và diện tích ván khuôn tiếp xúc.
- **Đo bóc Đào đắp Đất đá Mặt cắt ngang:** Tự động tính diện tích khép kín (Shoelace) và khối lượng giữa các lý trình theo phương pháp diện tích trung bình 2 đầu (Average-End-Area).
- **Bóc tách Thống kê Cốt thép (BBS):** Trích xuất số hiệu thanh, đường kính $\Phi$, chiều dài, số lượng và quy đổi ra tổng chiều dài và trọng lượng (kg, tấn) theo TCVN 1651:2018.
- **Đối soát 2 chiều (`<───>`):** So khớp dữ liệu kích thước hình học từ CAD với số hiển thị sửa tay (TextOverride) để bắt cờ cảnh báo đỏ (Red Flag).
- **Xuất bảng tính Excel chuẩn mực:** 100% công thức sống (`=E*F*G*H*I`, `=SUM(...)`), tuyệt đối tuân thủ nguyên tắc Zero Dead Numbers.

---

## 2. Công cụ MCP Khả dụng (`autocad-live-mcp`)

Agent sử dụng 11 công cụ chuẩn hóa từ máy chủ `autocad-live-mcp`:

1. `cad_get_status`: Kiểm tra trạng thái kết nối AutoCAD, tên bản vẽ đang mở, đường dẫn, đơn vị ($INSUNITS).
2. `cad_list_drawings`: Liệt kê tất cả các bản vẽ DWG đang mở trong AutoCAD.
3. `cad_open_drawing`: Mở file DWG/DXF hoặc kích hoạt nếu đang mở.
4. `cad_check_health`: **GIAI ĐOẠN 1 (Khám sức khỏe)** - Quét `$INSUNITS`, kiểm tra biến tỷ lệ `DIMLFAC` (cảnh báo đỏ nếu != 1.0), quét lỗi bay cao độ 3D $Z \ne 0$, tổng hợp Layer.
5. `cad_takeoff_layers`: **GIAI ĐOẠN 2 (Đo bóc hình học)** - Trích xuất Polyline, Line, Hatch, Text theo layer. Tính chiều dài 1:1, diện tích khép kín Shoelace, giải mã tiếng Việt TCVN3/VNI sang Unicode.
6. `cad_check_discrepancies`: **GIAI ĐOẠN 3 (Đối soát chéo)** - Quét Dimension, so sánh kích thước đo thực tế 1:1 (`Measurement`) vs số hiển thị sửa tay (`TextOverride`). Bật cờ RED FLAG nếu lệch.
7. `cad_extract_rebar_bbs`: Trích xuất bảng thống kê cốt thép BBS, tính trọng lượng chuẩn theo TCVN 1651:2018 ($0.006165 \times d^2 \times L \times N$).
8. `cad_export_takeoff_excel`: **GIAI ĐOẠN 4 (Xuất file kết quả)** - Xuất file Excel 100% công thức sống, chuẩn cấu trúc WBS phân cấp, Zero Dead Numbers.
9. `cad_run_command`: Gửi lệnh AutoCAD hoặc AutoLISP an toàn (có Security Guard chặn lệnh độc hại, kiểm tra Quiescent state).
10. `cad_zoom_extents`: Phóng to toàn bộ bản vẽ trên màn hình CAD.
11. `cad_import_pdf`: Import trực tiếp bản vẽ vector từ file PDF vào AutoCAD (lệnh PDFIMPORT).

---

## 3. Quy trình Đo bóc Khối lượng Chi tiết

### Bước 1: Khảo sát hiện trạng bản vẽ & Thư mục hồ sơ
1. **Kiểm tra kết nối CAD:** Kiểm tra xem AutoCAD có đang mở bản vẽ hợp lệ trên máy tính không (qua COM Interop `win32com.client`).
2. **Quét thư mục hồ sơ (Batch Scan):** Nếu có thư mục chứa nhiều file bản vẽ `.dwg`/`.dxf`, tự động quét đệ quy và phân loại theo hạng mục công trình.
3. **Quét danh sách Layer:** Tìm các layer chứa nét hình học kết cấu (`KET_CAU`, `THIET_KE`, `COT_THEP`, `MAT_CAT`, `DIA_CHAT`).

### Bước 2: Đo bóc Bê tông & Ván khuôn Kết cấu
Áp dụng công thức giải tích hình học chuẩn mực:
- **Cọc khoan nhồi:** $V = N \times \pi \times \left(\frac{D}{2}\right)^2 \times L$
- **Bệ móng mố, trụ:** $V = N \times L \times W \times H$; Diện tích ván khuôn: $S_{vk} = N \times 2 \times (L + W) \times H$.
- **Dầm Super-T (L=38.2m):** $V = N \times A_{mc} \times L$ với diện tích mặt cắt dầm $A_{mc} = 0.945\text{ m}^2$.
- **Bản mặt cầu liên tục nhiệt:** $V = L \times B \times t$ (với $t = 0.20\text{ m}$).

### Bước 3: Đo bóc Khối lượng Đào đắp Mặt cắt ngang
1. **Truy vấn Polyline:** Đọc tọa độ danh sách đỉnh $[(x_1, y_1), (x_2, y_2), ..., (x_n, y_n)]$.
2. **Tính diện tích hình học (Shoelace Formula):**
   $$A = \frac{1}{2} \left| \sum_{i=1}^{n-1} (x_i y_{i+1} - x_{i+1} y_i) + (x_n y_1 - x_1 y_n) \right|$$
3. **Tính khối lượng đào đắp giữa các lý trình:**
   $$V = \frac{A_1 + A_2}{2} \times L$$

### Bước 4: Đo bóc Thống kê Cốt thép (BBS)
1. Giải mã text block thống kê thép TCVN3 sang Unicode.
2. Trích xuất: Số hiệu thanh, Đường kính $d$ (mm), Chiều dài $L$ (m), Số lượng $N$.
3. Tính trọng lượng cốt thép theo TCVN 1651:2018:
   $$W = L \times N \times \left(0.006165 \times d^2\right) \quad (\text{kg})$$

### Bước 5: Xuất Bảng tính Excel Đo bóc Chuẩn hóa
Gọi script động cơ đo bóc CAD chuyên dụng:
```bash
python skills/aec-cad-automation/scripts/cad_takeoff_engine.py --mode all --output templates/BANG_DO_BOC_KHOI_LUONG_CAD.xlsx
```

---

## 4. Xử lý Lỗi & Cơ chế Tự tiến hóa (Self-Evolution)
- **Lỗi UIPI (User Interface Privilege Isolation):** Đảm bảo tiến trình CAD và MCP Agent cùng cấp quyền (Elevation level).
- **Lỗi Handle không hợp lệ:** Khi một đối tượng bị xóa hoặc Undo, handle cũ sẽ mất hiệu lực; Agent phải truy vấn lại đối tượng mới nhất thay vì dùng handle cũ.
- **Xử lý font TCVN3:** Tự động dùng `TCVN3Decoder` chuyển đổi các ký tự lỗi hiển thị sang tiếng Việt Unicode chuẩn trước khi đưa vào bảng tính.
- **Tự động vá lỗi (FIX):** Khi một lệnh vẽ bị từ chối do trùng tên block hoặc sai kiểu dữ liệu tọa độ, Agent tự động hiệu chỉnh và thử lại.

---

## 5. Quy Trình Đo Bóc Minh Bạch 4 Giai Đoạn (Chuẩn 23HG / Nguyễn Bảo Tú)

Mọi tác vụ đo bóc CAD bắt buộc tuân thủ quy tắc toàn cầu tại `~/.gemini/config/rules/aec-cad-takeoff-workflow-rule.md`:
1. **Giai đoạn 1 - Khám sức khỏe bản vẽ:** Quét đơn vị `$INSUNITS`, kiểm tra biến hệ số tỷ lệ `DIMLFAC` và `Scale factor` trong DimStyle (phát hiện lỗi co giãn 0.8), rà soát cao độ 3D ($Z \ne 0$) và trùng nét.
2. **Giai đoạn 2 - Trích xuất từng bước:** Log tiến trình rõ ràng (Block $\rightarrow$ Polyline $\rightarrow$ Ghép nối không gian hạ lưu - thượng lưu).
3. **Giai đoạn 3 - Bảng đối soát chéo:** Bắt buộc so sánh song song:
   $$\text{Hình học thực tế (Line đo thật)} \quad \longleftrightarrow \quad \text{Text ghi chú / Dim thiết kế}$$
   Bật cờ cảnh báo đỏ nếu phát hiện vênh lệch vật tư.
4. **Giai đoạn 4 - Quản lý file sạch sẽ:** Tuyệt đối không sinh file ra thư mục cha (`Downloads`, `Desktop`). File kết quả nằm gọn trong folder dự án chỉ định, tự động dọn sạch file debug tạm ngay khi hoàn thành.

---

## 6. Tiêu Chuẩn Nét Dim & Tái Lập Bản Vẽ Chuyên Nghiệp (AEC Native Dimensioning)

Tuyệt đối **KHÔNG ĐƯỢC VẼ NÉT DIM THỦ CÔNG** bằng các đối tượng rời rạc (`LINE` + `TEXT` + gạch chéo 45 độ giả lập). Bắt buộc phải tạo **Đối tượng Kích thước Nguyên bản AutoCAD (`DIMENSION` / `AcDbRotatedDimension`)** với các thiết lập chuẩn TCVN:

### A. Cấu Hình DimStyle Chuẩn (`TCVN_KT_120`, `TCVN_KC_120`, `TCVN_KT_50`)
1. **Ký hiệu đầu đo (`DIMBLK`):** Sử dụng `_ArchTick` (gạch chéo kiến trúc nghiêng 45° thanh mảnh, sắc nét). Tuyệt đối không dùng mũi tên tam giác thô.
2. **Vị trí con số (`DIMTAD = 1`):** Con số kích thước luôn nằm ngay ngắn **phía TRÊN đường kích thước**, không cắt đứt đường đo ngang.
3. **Khoảng hở chân gióng (`DIMEXO = 2.0 mm`):** Tạo khoảng cách từ điểm đo đến chân đường gióng để bản vẽ thông thoáng, không dính sát nét tường/cột.
4. **Độ nhô đuôi gióng (`DIMEXE = 1.5 mm`):** Nhô đều đặn qua đường kích thước.
5. **Chiều cao chữ (`DIMTXT = 2.2 mm`):** In ra giấy đạt $2.2\text{ mm}$ tiêu chuẩn TCVN 4453/5574, trong Model Space tự động nhân theo `DIMSCALE`.
6. **Làm tròn số nguyên (`DIMDEC = 0`):** Hiển thị số nguyên mm chẵn ($1840, 2060, 3900, 5400, 7200, 28800$), triệt tiêu các đuôi thập phân `.00`.
7. **Font chữ kích thước (`DIMTXSTY`):** Sử dụng `VN_ARIAL` (`arial.ttf`), màu Text Color 7, màu đường Dim Color 3 (Green) hoặc Color 8/251.

### B. Quy Chuẩn 3 Cấp Kích Thước Bắt Buộc (4 Phía: Đáy, Đỉnh, Trái, Phải)
* **Cấp 1 (Phân khoang chi tiết / Sub-spans):** Kích thước từng khoang dầm phụ, dầm WC ($1840, 2060$), vách ngăn phòng.
* **Cấp 2 (Nhịp trục định vị / Bay spans):** Đo tim trục các bước cột chính ($3900, 5400, 2200, 7200, 2400$).
* **Cấp 3 (Kích thước phủ bì tổng thể / Overall):** Đo tổng chiều dài ($28,800\text{ mm}$) và tổng chiều rộng ($21,200\text{ mm}$).
* **Kích thước chi tiết nội bộ (Internal Dims):** Bắt buộc đo thông thủy hành lang ($2400\text{ mm}$), sảnh đón ($5400\text{ mm}$), lọt lòng buồng thang máy ($2000 \times 2050\text{ mm}$), chiều rộng từng vế thang bộ ($1750\text{ mm}$), và chi tiết từng buồng WC 1:50.

---

## 7. Quy Tắc Xử Lý Font Tiếng Việt & Đối Soát Xung Đột (Clash Detection)

### A. Khắc Phục Lỗi Font Tiếng Việt Trong AutoLISP
* **Nguyên nhân vỡ font:** Bộ thông dịch AutoLISP của AutoCAD chỉ đọc bảng mã ANSI (Windows-1252), ghi trực tiếp UTF-8 sẽ bị mất dấu hoặc vỡ ký tự (`MẶT BẰNG` $\rightarrow$ `MT BNG`).
* **Giải pháp chuẩn 100%:** Mã hóa toàn bộ chuỗi tiếng Việt sang định dạng **AutoCAD Unicode Escape Sequence** `\U+XXXX` (ví dụ: `\U+0102` cho `Ă`, `\U+0110` cho `Đ`, `\U+1ED6` cho `Ỗ`, `\U+1EBA` cho `Ả`) trên nền font `Arial.ttf` (`VN_ARIAL`). Khi xuất qua DXF (ezdxf), ghi chuỗi UTF-8 tiêu chuẩn với style `VN_ARIAL`.

### B. 5 Điểm Kiểm Tra Xung Đột Bắt Buộc Giữa Kết Cấu & Kiến Trúc
1. **Lưới trục & Kích thước bao:** Tường xây hoàn thiện bắt buộc phải nằm trên tim dầm chịu lực.
2. **Lỗ mở bản sàn vs Cầu thang bộ:** Không được vẽ cầu thang đè lên sàn bê tông đặc phẳng, và không được bố trí phòng/bàn ghế đè lên lỗ thông tầng kết cấu.
3. **Cụm Vệ sinh vs Sàn hạ cốt:** Cụm WC phải đặt trùng khớp với dải sàn hạ cốt kết cấu $-0.100$ và dầm phụ (dải E-F hoặc A-B) để tránh đục phá dầm khi đi ống thoát nước.
4. **Vách thang máy V1:** Đúng vị trí vách bê tông cốt thép, cửa mở hướng ra hành lang giao thông trung tâm.
5. **Cột kết cấu:** Loại bỏ các cột ảo ở góc biên không tồn tại trong thiết kế kết cấu (ví dụ góc A1, A8).

### C. Quy Chuẩn Đóng Gói Hồ Sơ Dự Án 5 Phân Khu
Toàn bộ dự án sau khi hoàn thành bắt buộc phải đóng gói vào một thư mục chuyên nghiệp duy nhất, tự động xóa sạch các file rải rác ngoài thư mục:
```
HO_SO_THIET_KE_<TEN_DU_AN>/
├── 01_BAN_VE_CAD_DWG/                  <-- Chứa toàn bộ file Master DWG và DXF
├── 02_TIEN_LUONG_VA_DU_TOAN_EXCEL/     <-- Bảng tính Excel 100% công thức sống + Sơ đồ PNG nhúng trực tiếp
├── 03_HINH_ANH_XEM_TRUOC_PREVIEW_4K/   <-- Ảnh xem trước 4K / sơ đồ hình học CAD siêu nét
├── 04_MA_NGUON_AUTOLISP_PYTHON/        <-- Mã nguồn LISP và Python tái tạo bản vẽ
├── 05_THUYET_MINH_VA_BAO_CAO_KY_THUAT/ <-- Báo cáo đối chiếu và thuyết minh
└── README.md                           <-- Sổ tay chỉ mục quản lý hồ sơ
```

---

## 8. Kết Xuất Sơ Đồ Hình Học & Trích Xuất Viewport Bản Vẽ CAD Nhúng Trực Tiếp Vào Excel

Khi lập bảng diễn giải khối lượng chi tiết (Takeoff) cho công trình Cầu Đường / Dân Dụng từ bản vẽ CAD (`.dxf` / `.dwg`), bắt buộc kết xuất sơ đồ hình học kỹ thuật kèm hình trích xuất trực tiếp từ bản vẽ CAD độ phân giải cao (`2400x1400` đến `3600x2400` px) và **nhúng trực tiếp vào bên phải bảng tính Excel** (`openpyxl.drawing.image.Image`):

1. **Giải mã Tiếng Việt TCVN3 & Ký hiệu AutoCAD (`decode_cad_text`):**
   - Ánh xạ đơn ký tự (single-pass lookup) cho cả bảng mã thường (`.VnTime`, `.VnArial`) và bảng mã chữ HOA (`.VnTimeH`, `.VnArialH`) sang Unicode UTF-8.
   - Đổi ký hiệu đặc biệt AutoCAD: `%%c` $\rightarrow$ `Ø`, `%%p` $\rightarrow$ `±`, `%%d` $\rightarrow$ `°`, `%%%` $\rightarrow$ `%`, và lọc sạch mã điều khiển `MTEXT` (`\P`, `\f...;`, `\S...\^...;`).
2. **Bung đệ quy thực thể ảo (`DIMENSION` & `INSERT` via `e.virtual_entities()`):**
   - Khi đọc `TEXT` trong `ezdxf`, lấy `e.dxf.get("height", 180.0)` (không gọi `e.dxf.get("char_height")` trên `TEXT` để tránh lỗi `DXFAttributeError`).
   - Bung đệ quy `e.virtual_entities()` cho mọi đối tượng `DIMENSION` và `INSERT` để hiển thị đầy đủ 100% đường gióng kích thước, mũi tên và con số kích thước nguyên bản của Tư vấn thiết kế.
3. **Chống đè chữ & Nhúng ảnh trực quan vào Excel:**
   - Sử dụng bảng màu kỹ thuật độ tương phản cao trên nền sáng (`ACI_COLORS`), hộp nền trắng viền mảnh cho nhãn kích thước và kiểm tra va chạm hộp bao (`bounding-box collision avoidance`).
   - Nhúng toàn bộ ảnh `.png` vào các ô neo bên phải bảng diễn giải (`L4`, `M4`, `N4`, ...) trên từng sheet Excel để kỹ sư QS / Thẩm tra đối chiếu trực tiếp giữa **Hình vẽ CAD** và **Công thức sống Excel** ngay khi mở file.
