# 23HG GOLDEN MASTER BASELINE — NGUỒN SỰ THẬT DUY NHẤT (SSoT)

Tài liệu này khóa chặt **Đội hình Mẫu Ưu Tú Nhất (Dream Team Golden Master)** của hệ thống **23HG Multi-Agent System**, làm chuẩn gốc duy nhất cho mọi dự án Cầu, Đường, Cống hộp và Hạ tầng kỹ thuật. Tuyệt đối không sử dụng lại các script rút gọn hoặc kiến trúc thư mục nhân bản chéo cũ.

---

## 1. Kiến Trúc Đóng Gói Thư Mục Chuẩn Gốc (Dream Team 6 Phân Hệ — 0% Trùng Lặp)

Mọi dự án xuất xưởng đều được tổ chức thành **06 phân hệ nghiệp vụ độc lập**, mỗi tệp tin chỉ nằm ở **đúng 1 nơi duy nhất**:

| Mã thư mục | Tên phân hệ | Đối tượng sử dụng | Danh mục hồ sơ chuẩn bên trong |
| :--- | :--- | :--- | :--- |
| **`00_BAN_CHI_HUY_MASTER`** | Ban Chỉ Huy & Điều Hành Tổng Thể | Giám đốc dự án, Chỉ huy trưởng | • Sổ cái Master 15 Sheet (`Ho_So_KCS_QS_TienDo_*.xlsx`)<br>• Tiến độ Thi công Pro VBA (`.xlsm`) + MS Project (`.mpp`, `.xml`)<br>• Ma trận phân quyền RACI (`.xlsx`, `.md`)<br>• Thuyết minh BPTC (`.md`) & Báo cáo Thẩm tra AEC Audit (`.md`)<br>• Kê khai toàn vẹn MD5 (`DISPATCH_MANIFEST.json`) |
| **`01_HIEN_TRUONG_QLCL_KCS`** | Hiện Trường, QLCL, Thí Nghiệm & As-Built | Kỹ sư QA/QC, KCS, TVGS, Phòng LAS-XD | • **Bảng Theo Dõi Tần Suất Thí Nghiệm Bê Tông & Thép Chuyên Sâu** (`Bang_Theo_Doi_Tan_Suat_Thi_Nghiem_Be_Tong_Va_Thep_*.xlsx` — 4 Sheets)<br>• **Sơ Đồ Tiến Độ Trực Quan As-Built** (`So_Do_Tien_Do_Truc_Quan_AsBuilt_*.xlsx` & `.html`)<br>• Hồ sơ Word 22 Biên bản Nghiệm thu KCS (`.docx`)<br>• Bộ 05 (Cấp phối), Bộ 10 (Gantt), Bộ 11 (Danh mục KCS), Bộ 12–14 (Mẫu A4 nghiệm thu) |
| **`02_XUONG_TIEN_CHE_COT_THEP`** | Xưởng Tiền Chế Cốt Thép & Máy Cắt CNC | Tổ trưởng gia công thép, Vận hành CNC, Thủ kho | • **Bản vẽ phôi thép GDI + VBA 12 Sheet** (`RebarCut_Pro_Excel_V6_*.xlsm`)<br>• Dashboard tổng hợp cắt thép theo đường kính (`00_BANG_TONG_HOP_CAT_THEP_THEO_PHI.xlsx`)<br>• Bộ 01 (`01_To_Hop_Cat_Thep_11m7_RebarCut.xlsx`) & Bộ 04 (`04_Thong_Ke_Thep_Chi_Tiet_BBS.xlsx`)<br>• Thư mục `THEO_TUNG_DUONG_KINH_PHI/` (tách riêng từng file `.xlsx` theo từng $\Phi$)<br>• Thư mục `LENH_CAT_CNC_CSV/` & `CSV_XUAT_XUONG/` |
| **`03_KINH_TE_QS_DU_TOAN_THANH_TOAN`** | Kinh Tế, QS, Dự Toán & Thanh Toán 03a | Kỹ sư QS, Phòng Kế hoạch - Tài chính, Chủ đầu tư | • Bộ 02 (`02_Khoi_Luong_Dao_Dap_Trinh_Dien.xlsx`)<br>• Bộ 03 (`03_QS_Dien_Giai_Chi_Tiet_Takeoff.xlsx`)<br>• Bộ 06 (`06_Phan_Tich_Vat_Tu_Chi_Tiet_WBS.xlsx`)<br>• Bộ 07 (`07_Tong_Hop_Nhu_Cau_Vat_Tu_BOM_4_Giai_Doan.xlsx`)<br>• Bộ 08 (`08_Du_Toan_GXD_Thong_Tu_11_2021.xlsx`)<br>• Bộ 09 (`09_Thanh_Toan_Khoi_Luong_Phu_Luc_03a.xlsx`)<br>• Bộ 15 (`15_Nghi_Van_So_Lieu_Can_Xac_Nhan.xlsx`) |
| **`04_CO_GIOI_THIET_BI_VA_DAU_DIEZEL`** | Cơ Giới, Thiết Bị, Nhân Lực & Dầu Diezel | Đội xe máy, Ban vật tư nhiên liệu | • **Bảng Kế hoạch Ca Xe - Ca Máy - Nhân Lực - Dầu Diezel 3 Tầng Hợp Nhất** (`TDTC_CaXe_CaMay_DauDiezel_*.xlsx` — 5 Sheets chuẩn Vincons, Sheet 01 tích hợp 3 tầng công thức sống)<br>• Tiến độ ca máy MS Project (`Tien_Do_CaMay_*.xml`) |
| **`05_DU_LIEU_GOC_SCAN_MARKER`** | Kho Lưu Trữ Dữ Liệu Gốc & OCR Marker | Hệ thống AI / Kỹ sư tra cứu bản gốc | • Toàn bộ dữ liệu đầu vào: `bang_so_lieu.json`, `bang_so_lieu.xlsx`, `du_lieu.json`, `chia_doan.jsonl`, `noi_dung.md`, `thep_cho_to_hop_cat.json`, `tien_luong_du_toan_boq.json`, bản vẽ `.pdf` và `images/` |

---

## 2. Bản Đồ Các Engine Ưu Tú Nhất (Golden Master Generators)

| Phân hệ chuyên môn | Engine chuẩn gốc (BẮT BUỘC DÙNG) | Các bản cũ / rút gọn (CẤM DÙNG LẠI) |
| :--- | :--- | :--- |
| **1. Đóng gói 6 Phân hệ Dream Team** | `AECPackageDispatcher.reorganize_to_dream_team_6_folders` trong [`tools/package_dispatcher.py`](../../tools/package_dispatcher.py) | Không để nguyên 3 thư mục `BO_HO_SO_01`, `BO_HO_SO_02`, `03_HO_SO_THUC_CHIEN` gây trùng lặp file |
| **2. Gói Cơ giới & Dầu Diezel 3 Tầng (5 Sheets)** | `build_bep_an_3tier_fleet_workbook` trong [`tools/generate_bep_an_3tier_schedule.py`](../../tools/generate_bep_an_3tier_schedule.py) & [`examples/generate_km19_machine_schedule.py`](../../examples/generate_km19_machine_schedule.py) | Các hàm tự viết 11 cột tĩnh không có Gantt ngày và không có 3 tầng trên Sheet 01 |
| **3. Tần suất Thí nghiệm Bê tông & Thép (4 Sheets)** | [`tools/generate_tan_suat_thi_nghiem_excel.py`](../../tools/generate_tan_suat_thi_nghiem_excel.py) (Km19) & [`tools/generate_khai_hoang_2_elite_qlcl.py`](../../tools/generate_khai_hoang_2_elite_qlcl.py) (Khai Hoang 2) | Bảng tần suất chỉ có số tĩnh hoặc dùng `INT(x/20)` làm mất lần lấy mẫu đầu tiên |
| **4. Sơ đồ Tiến độ Trực quan As-Built** | [`tools/generate_visual_bridge_progress.py`](../../tools/generate_visual_bridge_progress.py) & [`tools/generate_khai_hoang_2_elite_qlcl.py`](../../tools/generate_khai_hoang_2_elite_qlcl.py) | Không có sơ đồ mặt cắt đứng trực quan |
| **5. Tổ hợp Cắt thép 1D CSP & Phôi xưởng** | [`tools/cutting_stock_solver.py`](../../tools/cutting_stock_solver.py) + [`tools/rebarcut_export.py`](../../tools/rebarcut_export.py) + [`examples/run_rebarcut_v6_native_com.py`](../../examples/run_rebarcut_v6_native_com.py) | Bảng cắt thép giả lập 15 dòng trong `pho_bang_master_builder.py` |
| **6. Sổ cái Master 15 Sheet & Word KCS** | [`examples/generate_khai_hoang_2_project.py`](../../examples/generate_khai_hoang_2_project.py) & [`examples/build_khai_hoang_2_full_pipeline.py`](../../examples/build_khai_hoang_2_full_pipeline.py) | [`tools/export_cau_khai_hoang_2.py`](../../tools/export_cau_khai_hoang_2.py) (bản 30 dòng thép mẫu rút gọn) |
| **7. Tiến độ Thi công Pro VBA (`.xlsm`)** | [`examples/build_khai_hoang_2_pro_schedule.py`](../../examples/build_khai_hoang_2_pro_schedule.py) | Bảng tiến độ tĩnh không có EVM / S-Curve |
