# 23HG CAD AUTOMATION MCP PRO (v2.0.0)
### Bộ Tự Động Hóa AutoCAD Chuyên Nghiệp Dành Cho AI Hàng Đầu
**Tác giả:** Nguyễn Bảo Tú (@baotuhg) | **Phát triển bởi:** 23HG SYSTEM  
**Bản quyền phát hành:** Dành cho anh em cộng sự & kỹ sư xây dựng 23HG  

---

## 🌟 GIỚI THIỆU SẢN PHẨM
**23HG CAD AUTOMATION MCP PRO** là giải pháp cầu nối thông minh kết nối các hệ thống AI tiên tiến nhất hiện nay (**Claude Desktop, Cursor IDE, Google Antigravity, Windsurf**) với phần mềm thiết kế **AutoCAD** (hoặc GstarCAD / ZWCAD) thông qua giao thức chuẩn **MCP (Model Context Protocol)**.

Toàn bộ hệ thống đã được **bản địa hóa và Việt hóa 100%** theo thuật ngữ kỹ thuật Xây dựng - Kiến trúc - Cầu đường Việt Nam (AEC), loại bỏ hoàn toàn các rào cản ngôn ngữ và mã nguồn dư thừa.

---

## 📁 CẤU TRÚC BỘ SẢN PHẨM

```text
23HG_CAD_AUTOMATION_MCP_PRO/
│
├── 01_CAI_DAT_TU_DONG_1_CLICK/
│   ├── CAI_DAT_NHANH_1_CLICK.bat      <- Nhấp đúp để cài đặt toàn bộ môi trường
│   └── KHOI_DONG_TEST_MCP.bat         <- Chạy thử nghiệm và kiểm tra kết nối với AutoCAD
│
├── 02_MA_NGUON_VIET_HOA_100_PHAN_TRAM/
│   ├── cad_cowork_mcp/                <- Toàn bộ mã nguồn Python đã Việt hóa 100% (0 chữ tiếng Trung)
│   ├── requirements.txt               <- Danh sách thư viện phụ thuộc
│   └── setup.py                       <- File cài đặt package hệ thống
│
├── 03_FILE_CAU_HINH_CAC_PHAN_MEM_AI/
│   ├── claude_desktop_config.json     <- Cấu hình cho Claude Desktop
│   ├── cursor_mcp_config.json         <- Cấu hình cho Cursor IDE
│   ├── antigravity_mcp_config.json    <- Cấu hình cho Antigravity
│   ├── windsurf_mcp_config.json       <- Cấu hình cho Windsurf
│   └── HUONG_DAN_CAU_HINH_AI.md       <- Hướng dẫn chi tiết & mẹo bật Auto-pass
│
├── 04_TIEN_ICH_AUTOLISP_BRIDGE/
│   ├── 23hg_bridge.lsp                <- File tiện ích AutoLISP hỗ trợ phím tắt trong CAD
│   └── README.md                      <- Thuyết minh kiến trúc Native COM không cần NETLOAD
│
├── 05_HUONG_DAN_SU_DUNG_VA_DANH_MUC_LENH/
│   ├── HUONG_DAN_SU_DUNG_A_TOI_Z.md   <- Sách hướng dẫn vận hành chi tiết từ A tới Z
│   └── DANH_MUC_135_LENH_CAD_TIENG_VIET.md <- Tra cứu 135 lệnh CAD tự động bằng tiếng Việt
│
└── README.md                          <- File giới thiệu tổng quan bộ sản phẩm
```

---

## 🚀 HƯỚNG DẪN 2 BƯỚC KHỞI ĐỘNG NHANH (ZERO-PLUGIN)

1. **Cài Đặt Môi Trường (30 giây):**
   Vào thư mục `01_CAI_DAT_TU_DONG_1_CLICK`, nhấp đúp chạy `CAI_DAT_NHANH_1_CLICK.bat`.
2. **Cấu Hình AI & Khởi Động:**
   Chép nội dung file tương ứng trong `03_FILE_CAU_HINH_CAC_PHAN_MEM_AI` vào phần mềm AI bạn đang dùng (Claude / Cursor / Antigravity).
   *Hệ thống sử dụng kiến trúc **Native COM Interop** kết nối trực tiếp hai chiều với AutoCAD - **KHÔNG CẦN GÕ LỆNH NETLOAD**, loại bỏ 100% rủi ro bảo mật và xung đột phiên bản.*

---

## 🛠️ CÁC TÍNH NĂNG VƯỢT TRỘI
- **Động Cơ Bóc Tách Khối Lượng Tiên Lượng Kỹ Thuật (QS Takeoff Native COM Engine):**
  - `cad_get_units`: Đọc chính xác đơn vị bản vẽ (`INSUNITS` hỗ trợ toàn bộ 25 mã chuẩn AutoCAD 0–24, `MEASUREMENT`, `DIMLFAC`, `LUNITS`). Cảnh báo đỏ nguy cơ lệch 1000 lần khi `INSUNITS = 0` (chưa khai báo đơn vị) hoặc hệ số kích thước `DIMLFAC != 1.0`.
  - `cad_quantity_by_layer`: Quét sâu `ModelSpace`, đo bóc và cộng dồn chiều dài tim tuyến ($m$), diện tích khép kín ($m^2$), phân loại chi tiết Lines, Polylines kín/hở, Circles, Arcs, Hatches, Ellipses.
    - **Triệt tiêu lỗi cộng trùng diện tích (Zero Double-Counting):** Tách bạch hoàn toàn cột Diện tích Polyline kín (`closed_polyline_area_m2`) và Diện tích Hatch (`hatch_area_m2`).
    - **Loại bỏ số 0 âm thầm (Zero Silent Zeroes):** Tự động tính khoảng cách tọa độ đỉnh cho `AcDb2dPolyline` / `AcDb3dPolyline` khi thuộc tính `.Length` trả về 0. Trả về nhật ký chẩn đoán `diagnostics.skipped_entities` nếu có thực thể lỗi.
    - **Tự động phân nhóm 10 Hạng mục AEC (Trade Classification):** Bê tông Dầm/Giằng, Cột/Móng, Sàn/Bản thang, Cốt thép BBS, Xây tường, Cửa/Vách kính, Hoàn thiện lát nền, Cơ điện MEP, Kết cấu chung.
  - `cad_count_blocks`: Tự động nhận diện Dynamic Block (`EffectiveName`), trích xuất toàn bộ cặp thuộc tính `Attributes` (Tag & Value), phân nhóm theo Layer, phát hiện Xref và tỷ lệ co giãn phi đồng nhất (`is_scaled`).
  - `cad_export_boq`: Xuất bảng tiên lượng BoQ ra file Excel (.xlsx) đa sheet phân cấp WBS chuẩn Bộ Xây dựng hoặc CSV, tuân thủ nguyên tắc **Zero Dead Numbers** (100% công thức sống `=SUM(...)`, nhân tỷ lệ quy đổi `=D4*0.001`, `=F4*1e-06`). Tự động tái sử dụng dữ liệu đệm (Cache) để tối ưu hiệu năng.
- **Kiến Trúc Native COM Độc Quyền (Zero-Plugin / Zero-DLL):** Điều khiển AutoCAD trực tiếp qua Windows COM API (AutoCAD ActiveX Automation), tương thích hoàn hảo mọi phiên bản AutoCAD (2020 - 2026+), không cần chạy `NETLOAD` DLL, loại bỏ hoàn toàn các phụ thuộc rủi ro từ plugin cũ.
- **Lớp Bảo Mật Security Guard:** Tự động phát hiện và chặn đứng mọi lệnh nguy hiểm (`QUIT`, `_ERASE ALL`, `PURGE ALL`, `SHELL`, `SCRIPT`) chống Prompt Injection.
- **Khử Treo Modal Dialog:** Tự động kiểm tra trạng thái `GetAcadState().IsQuiescent`, phát hiện hộp thoại cảnh báo để không bị treo timeout.
- **Bộ Kiểm Thử Tự Động Toàn Diện (Comprehensive Test Suite - 73+ Tests):**
  Hệ thống đi kèm bộ kiểm thử tự động đạt chuẩn công nghiệp, phân tách thành 2 nhóm kiểm thử rõ ràng:

  1. **Nhóm Unit Test & Mock CI/CD (Chạy độc lập mọi môi trường CI/CD, không cần mở CAD):**
     ```bash
     cd 02_MA_NGUON_VIET_HOA_100_PHAN_TRAM/cad_cowork_mcp
     python -m unittest discover -s tests -v
     ```
     *Bao phủ 72 bài test: Kiểm soát Whitelist 4 tầng, chặn AutoLISP nhúng `_.EXPLODE (ssget "X")`, phục hồi đứt kết nối COM `RPC_E_DISCONNECTED`, chống treo Modal Dialog, giải mã DAG batch dependencies, bộ định tuyến Scorer, công cụ QS Takeoff bóc tách hình học và kiểm tra Zero CJK.*

  2. **Nhóm Live Integration Test (Kiểm thử thực tế hai chiều với AutoCAD đang chạy):**
     ```bash
     cd 02_MA_NGUON_VIET_HOA_100_PHAN_TRAM/cad_cowork_mcp
     python -m unittest tests/test_live_cad_integration.py -v
     python -m unittest tests/test_live_known_geometry.py -v
     ```
     *Kết nối trực tiếp tới tiến trình AutoCAD thật (`acad.exe`):*
     - `test_live_cad_integration.py`: Đọc thông tin bản vẽ hiện hành (`Drawing1.dwg`), quét đơn vị `$INSUNITS`, bóc tách khối lượng hình học theo Layer, đếm Block & thiết bị kèm Attributes và xuất file Excel BoQ thực tế qua Native COM Interop.
     - `test_live_known_geometry.py`: Thực nghiệm hình học mẫu 5000×4000 mm (5×4 m) có chu vi 18.0 m, diện tích Pline kín 20.0 m2 và Hatch 20.0 m2, kiểm chứng tách rời diện tích chống cộng trùng và công thức sống Excel.

---
**23HG SYSTEM - ĐỒNG HÀNH CÙNG KỸ SƯ VIỆT**  
*Mọi quyền sở hữu và phát triển thuộc về Nguyễn Bảo Tú (@baotuhg).*
