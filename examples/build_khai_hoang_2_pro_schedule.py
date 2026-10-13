# -*- coding: utf-8 -*-
"""
HỆ THỐNG QUẢN LÝ TIẾN ĐỘ THI CÔNG CHUYÊN NGHIỆP - 23HG SYSTEM PRO
BỘ TẠO TIẾN ĐỘ MẪU G1 PRO CHO:
CÔNG TRÌNH: CẦU THÔN KHAI HOANG 2, KM 14+363.65 (L = 15.0m)
DỰ ÁN: ĐƯỜNG TỪ TRUNG TÂM HUYỆN ĐỒNG VĂN ĐI MỐC 450/456, HUYỆN MÈO VẠC, HÀ GIANG
GÓI THẦU SỐ 09: XÂY DỰNG CẦU VÀ ĐƯỜNG HAI ĐẦU CẦU

Khung thời gian thi công chuẩn: 20/10/2026 - 20/02/2027 (124 ngày lịch, 100 ngày làm việc)
Tuân thủ:
- TCVN 11823:2017 & TCVN 1651:2018 (Cầu đường & Cốt thép)
- Luật Xây dựng 135/2025/QH15 (VAT 10%)
- Nghị định 206/2026/NĐ-CP & Thông tư 38/2026/TT-BXD (Quản lý chi phí & Định mức ca máy)
- Nghị định 207/2026/NĐ-CP & Thông tư 32/2026/TT-BXD (QLCL & Hồ sơ nghiệm thu KCS)
- 100% CÔNG THỨC SỐNG - ZERO DEAD NUMBERS - KÈM VBA ENGINE & CUSTOM RIBBON UI
"""

import os
import sys
import shutil
import zipfile
import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

CODE_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
TEMPLATE_PATH = os.path.join(CODE_ROOT, r"23HG_TIEN_DO_THI_CONG_PRO\23HG_TIEN_DO_THI_CONG_PRO\23HG_DU_AN_MAU_TIEN_DO_CHUAN_G1_PRO.xlsm")
OUT_USER_DIR = os.path.join(os.path.expanduser("~"), "Downloads", r"Cầu thôn Khai Hoang 2, Km 14+363.65_Marker")
OUT_PRO_DIR = os.path.join(CODE_ROOT, r"23HG_TIEN_DO_THI_CONG_PRO\23HG_TIEN_DO_THI_CONG_PRO")
OUT_SYS_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), r"du_lieu_du_an\Cầu thôn Khai Hoang 2, Km 14+363.65_Marker")

FILE_NAME = "23HG_DU_AN_MAU_TIEN_DO_CAU_KHAI_HOANG_2_PRO.xlsm"

def build_khai_hoang_2_pro_schedule():
    print("=" * 80)
    print("   23HG SYSTEM PRO - THIẾT LẬP TIẾN ĐỘ CẦU THÔN KHAI HOANG 2 (KM 14+363.65)   ")
    print("   Thời gian thi công: 20/10/2026 - 20/02/2027 (124 ngày lịch, 100 ngày LV)   ")
    print("   Dựa trên mẫu chuẩn: 23HG_DU_AN_MAU_TIEN_DO_CHUAN_G1_PRO.xlsm                ")
    print("=" * 80)

    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"Không tìm thấy template: {TEMPLATE_PATH}")

    # Load workbook preserving VBA
    print(f"[*] Đang tải template: {TEMPLATE_PATH} (keep_vba=True)...")
    wb = openpyxl.load_workbook(TEMPLATE_PATH, keep_vba=True, data_only=False)

    # Styles
    f_title = Font(name="Segoe UI", size=13, bold=True, color="5A3718")
    f_sub = Font(name="Segoe UI", size=9, italic=True, color="6E3812")
    f_banner = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    f_title_brown = Font(name="Segoe UI", size=10, bold=True, color="6E3812")
    f_header = Font(name="Segoe UI", size=8.5, bold=True, color="333333")
    f_timeline_year = Font(name="Segoe UI", size=8.5, bold=True, color="5A3718")
    f_timeline_month = Font(name="Segoe UI", size=7.5, bold=True, color="555555")
    f_summary = Font(name="Segoe UI", size=9, bold=True, color="1E293B")
    f_norm = Font(name="Segoe UI", size=9, color="000000")
    f_norm_bold = Font(name="Segoe UI", size=8.5, bold=True, color="000000")
    f_crit = Font(name="Segoe UI", size=9, bold=True, color="DC2626")

    fill_coffee = PatternFill(fill_type="solid", start_color="5A3718", end_color="5A3718")
    fill_cream = PatternFill(fill_type="solid", start_color="F7EDE2", end_color="F7EDE2")
    fill_sand = PatternFill(fill_type="solid", start_color="EFE2D3", end_color="EFE2D3")
    fill_summary = PatternFill(fill_type="solid", start_color="FBF8F5", end_color="FBF8F5")
    fill_highlight = PatternFill(fill_type="solid", start_color="F7EDE2", end_color="F7EDE2")
    fill_crit = PatternFill(fill_type="solid", start_color="FEE2E2", end_color="FEE2E2")
    fill_dark = fill_sand
    fill_blue = fill_sand

    box_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # =========================================================================
    # 1. SHEET: NGAY_NGHI_LE
    # =========================================================================
    print("[1/9] Thiết lập sheet NGAY_NGHI_LE (Lịch nghỉ lễ 2026-2027 & Mùa mưa Hà Giang)...")
    ws_nl = wb["NGAY_NGHI_LE"]
    ws_nl["A1"] = "LỊCH THỜI TIẾT MÙA MƯA & CÁC NGÀY NGHỈ LỄ DỰ ÁN – 23HG SYSTEM"
    ws_nl["A2"] = "Dự án: Cầu thôn Khai Hoang 2 (Km14+363.65) – Lịch 6 ngày/tuần (Nghỉ CN & Nghỉ lễ). Tuân thủ BLLĐ 2019"

    # Xóa dữ liệu cũ rows 6-17
    for r in range(6, 18):
        for c in range(1, 7):
            ws_nl.cell(r, c).value = None

    holidays = [
        (1, "01/01/2027", "Nghỉ Tết Dương Lịch 2027", 1, 0.0, "Nghỉ lễ 1 ngày, trực ban bảo vệ công trường 24/24"),
        (2, "05/02/2027 - 11/02/2027", "Nghỉ Tết Nguyên Đán Đinh Mùi 2027", 7, 0.0, "Dừng thi công 7 ngày, niêm phong kho bãi thiết bị"),
        (3, "30/04/2027 - 01/05/2027", "Kỷ niệm Chiến thắng 30/4 & Quốc tế LĐ 1/5", 2, 0.0, "Giai đoạn khai thác bảo hành công trình"),
        (4, "02/09/2027 - 03/09/2027", "Kỷ niệm Ngày Quốc Khánh 2027", 2, 0.0, "Giai đoạn khai thác bảo hành công trình")
    ]
    for idx, item in enumerate(holidays, start=6):
        for c_idx, val in enumerate(item, start=1):
            cell = ws_nl.cell(row=idx, column=c_idx, value=val)
            cell.font = f_norm
            cell.border = box_border
            if c_idx in [1, 2, 4, 5]: cell.alignment = align_center
            else: cell.alignment = align_left

    ws_nl["A18"] = "II. QUY ĐỊNH MÙA MƯA CAO ĐIỂU & ĐIỀU KIỆN THI CÔNG NGHIÊM NGẶT (TCVN 8819:2011)"
    for r in range(20, 26):
        for c in range(1, 7):
            ws_nl.cell(r, c).value = None

    seasons = [
        (1, "20/10/2026 - 30/11/2026", "Mùa khô đầu đông (Nhiệt độ 16-24°C, ít mưa)", 1.0, "Đào đá hố móng, khoan cấy neo D32, đổ bê tông bệ mố, đúc dầm T", "Không có cấm đoán"),
        (2, "01/12/2026 - 31/01/2027", "Mùa đông vùng cao (Rét đậm, có sương muối về đêm)", 1.0, "Đổ bê tông thân mố, cẩu lắp dầm 50T, đổ bê tông bản mặt cầu", "CẤM: Đổ bê tông khi nhiệt độ môi trường < 5°C nếu không có phụ gia đông kết nhanh"),
        (3, "01/02/2027 - 20/02/2027", "Mùa xuân khô ráo (Nhiệt độ ấm dần 18-26°C)", 1.0, "Chống thấm mặt cầu, thảm BTN C12.5, thử tải cầu và bàn giao", "CẤM: Thảm bê tông nhựa khi mặt cầu bị ẩm ướt do sương mù"),
        (4, "Tháng 06/2027 - Tháng 09/2027", "Mùa mưa lũ cao điểm hàng năm tại Mèo Vạc", 0.65, "Giai đoạn khai thác vận hành", "Cảnh báo xói lở chân khay mố cầu")
    ]
    for idx, item in enumerate(seasons, start=20):
        for c_idx, val in enumerate(item, start=1):
            cell = ws_nl.cell(row=idx, column=c_idx, value=val)
            cell.font = f_norm
            cell.border = box_border
            if c_idx in [1, 2, 4]: cell.alignment = align_center
            else: cell.alignment = align_left
            if c_idx == 6:
                cell.font = Font(name="Segoe UI", size=9, bold=True, color="DC2626")
                cell.fill = fill_crit

    # =========================================================================
    # 2. SHEET: DB_DINH_MUC
    # =========================================================================
    print("[2/9] Thiết lập sheet DB_DINH_MUC (Định mức cầu đường TT 38/2026/TT-BXD)...")
    ws_dm = wb["DB_DINH_MUC"]
    ws_dm["A1"] = "CƠ SỞ DỮ LIỆU ĐỊNH MỨC NĂNG SUẤT CA MÁY & HAO PHÍ XÂY LẮP – 23HG SYSTEM"
    ws_dm["A2"] = "Chuẩn hóa theo Thông tư 38/2026/TT-BXD, Thông tư 36/2026/TT-BXD & Luật Xây dựng 135/2025/QH15"

    for r in range(5, 30):
        for c in range(1, 9):
            ws_dm.cell(r, c).value = None

    db_bridge_items = [
        (1, "AB.11312", "Đào bóc đất hữu cơ, bạt taluy bằng máy đào bánh xích 1.25m3", "100 m3", "Máy đào gầu 1.25m3, ôtô 15T", "450 m3/ca", 2850000, "TT 38/2026/TT-BXD"),
        (2, "AB.31114", "Đào phá đá hố móng mố M1, M2 bằng máy búa đập thủy lực", "100 m3", "Máy đào gắn búa đập đá, máy nén khí", "130 m3/ca", 11500000, "TT 38/2026/TT-BXD"),
        (3, "AF.12111", "Khoan tạo lỗ và cấy thép neo đá D32 sâu 500mm vữa SikaGrout", "lỗ", "Máy khoan đá khí nén, bơm ép vữa", "24 lỗ/ca", 220000, "TT 38/2026/TT-BXD"),
        (4, "AF.12211", "Đổ bê tông đệm móng mố C10 dày 10cm", "m3", "Máy trộn 500L, đầm bàn, đầm dùi", "18 m3/ca", 1050000, "TT 38/2026/TT-BXD"),
        (5, "AF.12222", "Đổ bê tông bệ móng mố M1, M2 mác C30 (M350)", "m3", "Trạm trộn 30m3/h, xe bồn, bơm tĩnh", "45 m3/ca", 1850000, "TT 38/2026/TT-BXD"),
        (6, "AF.12233", "Đổ bê tông thân mố, tường ngực và tường cánh mố C30", "m3", "Trạm trộn 30m3/h, xe bồn, bơm tĩnh", "35 m3/ca", 1920000, "TT 38/2026/TT-BXD"),
        (7, "AF.12310", "Gia công cốt thép móng và thân mố D<=10mm & 10<D<=18mm", "Tấn", "Máy cắt uốn sắt CNC, máy hàn điện", "1.5 tấn/ca", 20500000, "TT 38/2026/TT-BXD"),
        (8, "AF.12330", "Gia công cốt thép móng và thân mố D>18mm (RebarCut 11.7m)", "Tấn", "Trạm Shearline, máy cắt uốn thanh nặng", "2.5 tấn/ca", 20200000, "TT 38/2026/TT-BXD"),
        (9, "AG.13111", "Đúc dầm chủ T15m mác C40 (M500) tại bãi đúc hiện trường", "m3", "Trạm 30m3/h, xe bồn, đầm dùi cao tần", "20 m3/ca", 2250000, "TT 38/2026/TT-BXD"),
        (10, "AG.13230", "Gia công cốt thép dầm T15m (thép chủ A1, B1 D32, RebarCut)", "Tấn", "Trạm cắt CNC dầm chủ, máy uốn định hình", "1.8 tấn/ca", 21000000, "TT 38/2026/TT-BXD"),
        (11, "AG.13311", "Lắp dựng ván khuôn định hình đúc dầm T15m", "m2", "Ván khuôn thép định hình, cây chống", "60 m2/ca", 320000, "TT 38/2026/TT-BXD"),
        (12, "AG.13711", "Lắp đặt gối cầu cao su cốt bản thép 350x500x84mm", "cái", "Kích thủy lực 50T, máy trắc đạc Leica", "8 cái/ca", 4500000, "TT 38/2026/TT-BXD"),
        (13, "AG.13720", "Lao lắp phiến dầm T15m bằng cần trục bánh xích 50T", "phiến", "Cần trục bánh xích 50T, cẩu phụ 25T", "2 phiến/ca", 9250000, "TT 38/2026/TT-BXD"),
        (14, "AG.13411", "Đổ bê tông dầm ngang và mối nối ướt liên tục nhiệt C30", "m3", "Xe bồn, máy bơm cần, đầm dùi", "15 m3/ca", 2120000, "TT 38/2026/TT-BXD"),
        (15, "AG.13511", "Đổ bê tông bản mặt cầu C35 dày 12cm & lưới thép D6", "m3", "Trạm 30m3/h, xe bồn, máy san gạt mặt", "30 m3/ca", 2150000, "TT 38/2026/TT-BXD"),
        (16, "AH.14211", "Xây tường chắn đá hộc VXM M75 sau mố (4 đốt)", "m3", "Máy trộn 250L, giàn giáo, tổ thợ nề", "18 m3/ca", 1250000, "TT 38/2026/TT-BXD")
    ]
    for idx, item in enumerate(db_bridge_items, start=5):
        for c_idx, val in enumerate(item, start=1):
            cell = ws_dm.cell(row=idx, column=c_idx, value=val)
            cell.font = f_norm
            cell.border = box_border
            if c_idx in [1, 2, 4, 8]: cell.alignment = align_center
            elif c_idx == 7:
                cell.alignment = align_right
                cell.number_format = '#,##0 "VNĐ"'
            else: cell.alignment = align_left

    # =========================================================================
    # 3. SHEET: BOQ_TIEN_DO
    # =========================================================================
    print("[3/9] Thiết lập sheet BOQ_TIEN_DO (BoQ Cầu thôn Khai Hoang 2 & Live Formulas)...")
    ws_boq = wb["BOQ_TIEN_DO"]
    ws_boq["A1"] = "BẢNG TIÊN LƯỢNG KHỐI LƯỢNG (BOQ), ĐỊNH MỨC CA MÁY & CHI PHÍ DỰ TOÁN G_XD"
    ws_boq["A2"] = "Dự án: Cầu thôn Khai Hoang 2 (Km 14+363.65) – Tuyến Đồng Văn - Mốc 450/456 – Chuẩn hóa TT 38/2026 & TT 36/2026"

    # Xóa dữ liệu cũ rows 5-45
    for r in range(5, 45):
        for c in range(1, 12):
            ws_boq.cell(r, c).value = None

    boq_bridge_data = [
        # (STT, Mã, Tên, ĐV, KL, Đơn giá quy về 1 ĐV, WBS, Năng suất/ngày, Thời lượng)
        # PHẦN I: PHỤ TRỢ, NỀN ĐƯỜNG & HỐ MÓNG CẦU
        ("A", "", "I. CÔNG TÁC PHỤ TRỢ, NỀN ĐƯỜNG & HỐ MÓNG CẦU", "", "", "", "1.1", "", ""),
        (1, "AB.11312", "Đào bóc lớp đất hữu cơ, bạt taluy hố móng", "m3", 320, 45000, "1.1.1", 450, "=ROUNDUP(E6/I6, 0)"),
        (2, "AB.31114", "Đào phá đá hố móng mố M1, M2 bằng búa đập thủy lực", "m3", 450, 185000, "1.2.1", 35, "=ROUNDUP(E7/I7, 0)"),
        (3, "TN.11110", "Đường công vụ qua suối & đê quây nắn dòng thoát nước tạm", "gói", 1, 180000000, "1.1.2", 1, "=ROUNDUP(E8/I8, 0)"),
        (4, "TN.11120", "Lán trại, bãi đúc dầm T15m & trạm gia công cốt thép CNC", "gói", 1, 220000000, "1.1.3", 1, "=ROUNDUP(E9/I9, 0)"),

        # PHẦN II: KẾT CẤU HẠ BỘ MỐ M1, M2 & TƯỜNG CHẮN
        ("B", "", "II. KẾT CẤU HẠ BỘ MỐ M1, M2 & TƯỜNG CHẮN ĐẦU CẦU", "", "", "", "1.2", "", ""),
        (5, "AF.12111", "Khoan cấy thép neo D32 vào nền đá (98 lỗ, h=500mm)", "lỗ", 98, 220000, "1.2.2", 12, "=ROUNDUP(E11/I11, 0)"),
        (6, "AF.12211", "Bê tông đệm móng mố C10 (M150) dày 10cm", "m3", 9.35, 1050000, "1.2.3", 18, "=ROUNDUP(E12/I12, 0)"),
        (7, "AF.12222", "Bê tông bệ móng mố M1, M2 mác C30 (M350)", "m3", 220.0, 1850000, "1.2.4", 45, "=ROUNDUP(E13/I13, 0)"),
        (8, "AF.12233", "Bê tông thân mố, tường ngực và tường cánh mố C30", "m3", 169.28, 1920000, "1.2.5", 35, "=ROUNDUP(E14/I14, 0)"),
        (9, "AF.12310", "Cốt thép móng, thân mố M1, M2 & tường chắn (BBS RebarCut V6.0)", "Tấn", 48.93, 20350000, "1.2.3", 2.5, "=ROUNDUP(E15/I15, 0)"),
        (10, "AF.12411", "Ván khuôn móng và thân mố M1, M2", "m2", 564.0, 280000, "1.2.3", 60, "=ROUNDUP(E16/I16, 0)"),
        (11, "AH.14211", "Xây tường chắn đá hộc 4 đốt đầu cầu và sau mố VXM M75", "m3", 142.0, 1250000, "1.2.7", 18, "=ROUNDUP(E17/I17, 0)"),

        # PHẦN III: KẾT CẤU THƯỢNG BỘ, DẦM T15M & HOÀN THIỆN
        ("C", "", "III. CHẾ TẠO DẦM T15M, LAO LẮP & THƯỢNG BỘ CẦU", "", "", "", "1.3", "", ""),
        (12, "AG.13111", "Bê tông đúc 4 phiến dầm chủ T15m mác C40 (M500)", "m3", 37.0, 2250000, "1.3.2", 20, "=ROUNDUP(E19/I19, 0)"),
        (13, "AG.13230", "Cốt thép 4 phiến dầm chủ T15m (BBS RebarCut D32, D16, D10..)", "Tấn", 13.43, 20800000, "1.3.1", 1.8, "=ROUNDUP(E20/I20, 0)"),
        (14, "AG.13311", "Ván khuôn định hình đúc dầm T15m", "m2", 225.0, 320000, "1.3.1", 60, "=ROUNDUP(E21/I21, 0)"),
        (15, "AG.13711", "Gối cầu cao su cốt bản thép 350x500x84mm", "cái", 8, 4500000, "1.3.4", 8, "=ROUNDUP(E22/I22, 0)"),
        (16, "AG.13720", "Lao lắp 4 phiến dầm T15m bằng cẩu bánh xích 50T", "phiến", 4, 9250000, "1.3.5", 2, "=ROUNDUP(E23/I23, 0)"),
        (17, "AG.13411", "Bê tông dầm ngang, mối nối ướt & bản mặt cầu C35", "m3", 23.57, 2120000, "1.4.1", 15, "=ROUNDUP(E24/I24, 0)"),
        (18, "AG.13522", "Cốt thép bản mặt cầu, gờ lan can & bản quá độ (BBS RebarCut)", "Tấn", 5.65, 21000000, "1.4.2", 1.5, "=ROUNDUP(E25/I25, 0)"),
        (19, "AH.14111", "Bê tông bản quá độ M1, M2 mác C25 dày 25cm", "m3", 12.53, 1850000, "1.4.3", 20, "=ROUNDUP(E26/I26, 0)"),
        (20, "AG.13633", "Gờ lan can bê tông C25 & tay vịn thép mạ kẽm", "m", 30.0, 1450000, "1.4.4", 15, "=ROUNDUP(E27/I27, 0)"),
        (21, "AG.13722", "Khe co giãn răng lược thép mạ kẽm 2 đầu cầu", "m", 14.6, 3200000, "1.4.5", 8, "=ROUNDUP(E28/I28, 0)"),
        (22, "AD.23220", "Tưới dính bám & thảm BTN C12.5 dày 7cm mặt cầu", "m2", 105.0, 285000, "1.4.6", 120, "=ROUNDUP(E29/I29, 0)"),
        (23, "AH.11110", "Hệ thống ATGT, sơn kẻ đường, thử tải cầu & bàn giao", "gói", 1, 125000000, "1.5.1", 1, "=ROUNDUP(E30/I30, 0)")
    ]

    r_row = 5
    for item in boq_bridge_data:
        stt, ma, ten, dv, kl, dg, wbs, ns, dur = item
        if stt in ["A", "B", "C"]:
            ws_boq.cell(r_row, 1, value=stt).alignment = align_center
            ws_boq.cell(r_row, 3, value=ten).font = Font(name="Segoe UI", size=9.5, bold=True, color="5A3718")
            ws_boq.cell(r_row, 8, value=wbs).alignment = align_center
            if stt == "A":
                ws_boq.cell(r_row, 7, value="=SUM(G6:G9)").font = f_norm_bold
            elif stt == "B":
                ws_boq.cell(r_row, 7, value="=SUM(G11:G17)").font = f_norm_bold
            elif stt == "C":
                ws_boq.cell(r_row, 7, value="=SUM(G19:G30)").font = f_norm_bold
            ws_boq.cell(r_row, 7).number_format = '#,##0 "VNĐ"'
            for c in range(1, 11):
                ws_boq.cell(r_row, c).fill = fill_cream
                ws_boq.cell(r_row, c).border = box_border
        else:
            ws_boq.cell(r_row, 1, value=stt).alignment = align_center
            ws_boq.cell(r_row, 2, value=ma).alignment = align_center
            ws_boq.cell(r_row, 3, value=ten).alignment = align_left
            ws_boq.cell(r_row, 4, value=dv).alignment = align_center
            ws_boq.cell(r_row, 5, value=kl).alignment = align_right
            ws_boq.cell(r_row, 5).number_format = '#,##0.00' if isinstance(kl, float) else '#,##0'
            ws_boq.cell(r_row, 6, value=dg).alignment = align_right
            ws_boq.cell(r_row, 6).number_format = '#,##0 "VNĐ"'
            ws_boq.cell(r_row, 7, value=f"=ROUND(E{r_row}*F{r_row}, 0)").alignment = align_right
            ws_boq.cell(r_row, 7).number_format = '#,##0 "VNĐ"'
            ws_boq.cell(r_row, 8, value=wbs).alignment = align_center
            ws_boq.cell(r_row, 9, value=ns).alignment = align_right
            ws_boq.cell(r_row, 10, value=dur).alignment = align_center
            for c in range(1, 11):
                ws_boq.cell(r_row, c).font = f_norm
                ws_boq.cell(r_row, c).border = box_border
        r_row += 1

    # Dòng tổng kết dự toán (Row 31 to 35)
    ws_boq["C31"] = "TỔNG CHI PHÍ TRỰC TIẾP (T) = T_I + T_II + T_III"
    ws_boq["C31"].font = Font(name="Segoe UI", size=9.5, bold=True, color="5A3718")
    ws_boq["G31"] = "=G5+G10+G18"
    ws_boq["G31"].font = Font(name="Segoe UI", size=10, bold=True, color="5A3718")
    ws_boq["G31"].number_format = '#,##0 "VNĐ"'

    ws_boq["C32"] = "CHI PHÍ GIÁN TIẾP (GT) = 7.3% x T (Công trình giao thông NĐ 206/2026)"
    ws_boq["C32"].font = f_norm_bold
    ws_boq["G32"] = "=ROUND(G31 * 0.073, 0)"
    ws_boq["G32"].number_format = '#,##0 "VNĐ"'

    ws_boq["C33"] = "THU NHẬP CHỊU THUẾ TÍNH TRƯỚC (TL) = 5.5% x (T + GT)"
    ws_boq["C33"].font = f_norm_bold
    ws_boq["G33"] = "=ROUND((G31 + G32) * 0.055, 0)"
    ws_boq["G33"].number_format = '#,##0 "VNĐ"'

    ws_boq["C34"] = "THUẾ GIÁ TRỊ GIA TĂNG (VAT) = 10% x (T + GT + TL) (Luật XD 135/2025/QH15)"
    ws_boq["C34"].font = f_norm_bold
    ws_boq["G34"] = "=ROUND((G31 + G32 + G33) * 0.10, 0)"
    ws_boq["G34"].number_format = '#,##0 "VNĐ"'

    ws_boq["C35"] = "TỔNG GIÁ TRỊ DỰ TOÁN XÂY DỰNG G_XD (T + GT + TL + VAT)"
    ws_boq["C35"].font = Font(name="Segoe UI", size=11, bold=True, color="5A3718")
    ws_boq["G35"] = "=G31 + G32 + G33 + G34"
    ws_boq["G35"].font = Font(name="Segoe UI", size=12, bold=True, color="B91C1C")
    ws_boq["G35"].number_format = '#,##0 "VNĐ"'
    ws_boq["G35"].fill = PatternFill(fill_type="solid", start_color="FEF08A", end_color="FEF08A")

    for rr in range(31, 36):
        for col_l in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]:
            ws_boq[f"{col_l}{rr}"].border = box_border

    # =========================================================================
    # 4. SHEET: TIEN_DO (TIẾN ĐỘ THI CÔNG & MẠNG CPM CẦU KHAI HOANG 2)
    # =========================================================================
    print("[4/9] Thiết lập sheet TIEN_DO (Mạng CPM 20/10/2026 - 20/02/2027 & Gantt 24 kỳ)...")
    ws_td = wb["TIEN_DO"]
    ws_td.views.sheetView[0].showGridLines = True

    # Row 1 Banner
    ws_td["A1"] = "• HỆ THỐNG QUẢN LÝ TIẾN ĐỘ THI CÔNG & MẠNG CPM TỰ ĐỘNG – 23HG SYSTEM"
    ws_td["A1"].font = f_banner
    ws_td["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # Row 2 Project info
    ws_td["A2"] = "DỰ ÁN: CẦU THÔN KHAI HOANG 2 (KM 14+363.65) - TUYẾN ĐỒNG VĂN ĐI MỐC 450/456, HÀ GIANG"
    ws_td["A2"].font = Font(name="Segoe UI", size=8.5, bold=True, color="000000")
    ws_td["E2"] = "BẮT ĐẦU:"
    ws_td["F2"] = datetime.datetime(2026, 10, 20)
    ws_td["F2"].font = Font(name="Segoe UI", size=8.5, bold=True, color="5A3718")
    ws_td["F2"].number_format = "dd/mm/yyyy"
    ws_td["G2"] = "KẾT THÚC:"
    ws_td["H2"] = datetime.datetime(2027, 2, 20)
    ws_td["H2"].font = Font(name="Segoe UI", size=8.5, bold=True, color="B91C1C")
    ws_td["H2"].number_format = "dd/mm/yyyy"

    # Row 3 Control
    ws_td["A3"] = "Ngày kiểm soát (Data Date):"
    ws_td["C3"] = datetime.datetime(2026, 12, 15)
    ws_td["C3"].font = Font(name="Segoe UI", size=8.5, bold=True, color="2563EB")
    ws_td["C3"].number_format = "dd/mm/yyyy"
    ws_td["E3"] = "Lịch thi công:"
    ws_td["F3"] = "6 ngày/tuần (Nghỉ CN & Lễ)"
    ws_td["G3"] = "Chế độ CPM:"
    ws_td["H3"] = "Tự động (Forward/Backward Pass)"

    # Unmerge any previous timeline ranges in row 4
    for rng in list(ws_td.merged_cells.ranges):
        if rng.min_row <= 4 and rng.max_row >= 4 and rng.min_col >= 14:
            ws_td.unmerge_cells(str(rng))

    # Row 4 Timeline Years
    ws_td["A4"] = "BẢNG TIẾN ĐỘ THI CÔNG CHI TIẾT (CPM TỰ ĐỘNG) – GÓI THẦU SỐ 09"
    # Năm 2026: N đến Y (Cols 14 to 25)
    for c_i in range(14, 26):
        ws_td.cell(row=4, column=c_i).fill = fill_sand
        ws_td.cell(row=4, column=c_i).border = box_border
    ws_td.cell(row=4, column=14, value="Năm 2026").font = f_timeline_year
    ws_td.cell(row=4, column=14).alignment = align_center
    try: ws_td.merge_cells("N4:Y4")
    except Exception: pass

    # Năm 2027: Z đến AK (Cols 26 to 37)
    for c_i in range(26, 38):
        ws_td.cell(row=4, column=c_i).fill = fill_sand
        ws_td.cell(row=4, column=c_i).border = box_border
    ws_td.cell(row=4, column=26, value="Năm 2027").font = f_timeline_year
    ws_td.cell(row=4, column=26).alignment = align_center
    try: ws_td.merge_cells("Z4:AK4")
    except Exception: pass

    # Row 5 Timeline Periods (24 periods covering 20/10/2026 to 20/02/2027)
    periods = [
        "20/10", "25/10", "01/11", "06/11", "11/11", "16/11", "21/11", "26/11",
        "01/12", "06/12", "11/12", "16/12", "21/12", "26/12", "01/01", "06/01",
        "11/01", "16/01", "21/01", "26/01", "01/02", "06/02", "12/02", "17/02"
    ]
    for p_idx, p_name in enumerate(periods, start=14):
        c = ws_td.cell(row=5, column=p_idx, value=p_name)
        c.font = f_timeline_month
        c.fill = fill_sand
        c.alignment = align_center
        c.border = box_border
        ws_td.column_dimensions[get_column_letter(p_idx)].width = 5.5

    # Xóa dữ liệu cũ rows 6-45
    for r in range(6, 45):
        for c in range(1, 38):
            ws_td.cell(r, c).value = None

    # Danh mục WBS Cầu thôn Khai Hoang 2 (20/10/2026 - 20/02/2027)
    tasks_khai_hoang_2 = [
        # (STT, WBS, Tên, Loại, Predecessors, WorkDays, K_tt, StartDate, FinishDate, TF, IsCrit, Pct)
        (1, "1", "GÓI THẦU SỐ 09: XÂY DỰNG CẦU THÔN KHAI HOANG 2 (KM 14+363.65)", "Summary", "", 100, 1.0, "20/10/2026", "20/02/2027", 0, "GĂNG", 0.42),
        (2, "1.1", "Công tác chuẩn bị, phụ trợ & trạm gia công cốt thép", "Summary", "", 14, 1.0, "20/10/2026", "04/11/2026", 0, "GĂNG", 1.0),
        (3, "1.1.1", "Huy động nhân lực, thiết bị thi công & dựng lán trại", "Task", "", 14, 1.0, "20/10/2026", "04/11/2026", 0, "GĂNG", 1.0),
        (4, "1.1.2", "Đường công vụ qua suối & đê quây nắn dòng thoát nước tạm", "Task", "1.1.1SS+2", 12, 1.0, "21/10/2026", "03/11/2026", 1, "", 1.0),
        (5, "1.1.3", "Lắp dựng bãi đúc dầm T15m & trạm gia công cốt thép CNC", "Task", "1.1.1SS+3", 10, 1.0, "23/10/2026", "03/11/2026", 1, "", 1.0),

        (6, "1.2", "Thi công kết cấu móng & hạ bộ mố M1, M2", "Summary", "", 50, 1.0, "04/11/2026", "01/01/2027", 0, "GĂNG", 0.55),
        (7, "1.2.1", "Đào hố móng đá bệ mố M1, M2 & bạt đá taluy nền đường", "Task", "1.1.1FS", 14, 1.0, "04/11/2026", "19/11/2026", 0, "GĂNG", 1.0),
        (8, "1.2.2", "Khoan cấy 98 cây thép neo đá D32 & bơm vữa SikaGrout 214-11", "Task", "1.2.1FS", 8, 1.0, "20/11/2026", "28/11/2026", 0, "GĂNG", 1.0),
        (9, "1.2.3", "Đổ BT lót móng & lắp cốt thép bệ mố M1, M2 (RebarCut D22, D25)", "Task", "1.2.2FS", 6, 1.0, "30/11/2026", "05/12/2026", 0, "GĂNG", 1.0),
        (10, "1.2.4", "Đổ bê tông bệ móng mố M1 & M2 mác C30 (M350)", "Task", "1.2.3FS", 3, 1.0, "07/12/2026", "09/12/2026", 0, "GĂNG", 1.0),
        (11, "1.2.5", "Thi công thân mố, tường ngực, đá kê gối M1 & M2", "Task", "1.2.4FS", 18, 1.0, "10/12/2026", "30/12/2026", 0, "GĂNG", 0.35),
        (12, "1.2.6", "Thi công tường cánh mố chữ U (M1, M2)", "Task", "1.2.5SS+6", 14, 1.0, "17/12/2026", "01/01/2027", 2, "", 0.10),
        (13, "1.2.7", "Xây tường chắn đá hộc 4 đốt đầu cầu & sau mố", "Task", "1.2.4FS", 22, 1.0, "10/12/2026", "05/01/2027", 3, "", 0.20),

        (14, "1.3", "Chế tạo dầm T15m & lao lắp vào vị trí", "Summary", "", 55, 1.0, "04/11/2026", "05/01/2027", 0, "GĂNG", 0.40),
        (15, "1.3.1", "Gia công cốt thép dầm T15m (A1, B1 D32) & cốp pha bãi đúc", "Task", "1.1.3FS", 18, 1.0, "04/11/2026", "24/11/2026", 1, "", 1.0),
        (16, "1.3.2", "Đổ bê tông 4 phiến dầm T L=15m mác C40 tại bãi đúc", "Task", "1.3.1FS", 11, 1.0, "25/11/2026", "07/12/2026", 1, "", 1.0),
        (17, "1.3.3", "Dưỡng hộ bê tông dầm T15m đạt 100% cường độ R28", "Task", "1.3.2FS", 24, 1.0, "08/12/2026", "05/01/2027", 0, "GĂNG", 0.30),
        (18, "1.3.4", "Lắp đặt 8 gối cầu cao su cốt bản thép bệ kê gối mố", "Task", "1.2.5FS", 2, 1.0, "31/12/2026", "02/01/2027", 2, "", 0.0),
        (19, "1.3.5", "Cẩu lắp 4 phiến dầm T15m vào vị trí gối mố bằng cẩu 50T", "Task", "1.3.3FS", 4, 1.0, "06/01/2027", "09/01/2027", 0, "GĂNG", 0.0),

        (20, "1.4", "Thi công thượng bộ, bản mặt cầu & hoàn thiện", "Summary", "", 35, 1.0, "11/01/2027", "16/02/2027", 0, "GĂNG", 0.0),
        (21, "1.4.1", "Lắp dựng cốt thép & đổ bê tông mối nối ướt liên tục nhiệt", "Task", "1.3.5FS", 6, 1.0, "11/01/2027", "16/01/2027", 0, "GĂNG", 0.0),
        (22, "1.4.2", "Lắp dựng cốp pha, cốt thép & đổ BT bản mặt cầu C35 dày 12cm", "Task", "1.4.1FS", 7, 1.0, "18/01/2027", "25/01/2027", 0, "GĂNG", 0.0),
        (23, "1.4.3", "Thi công bản quá độ 2 đầu mố M1 & M2", "Task", "1.4.2SS+3", 6, 1.0, "21/01/2027", "27/01/2027", 2, "", 0.0),
        (24, "1.4.4", "Lắp dựng gờ lan can bê tông & tay vịn thép mạ kẽm", "Task", "1.4.2SS+2", 8, 1.0, "20/01/2027", "28/01/2027", 2, "", 0.0),
        (25, "1.4.5", "Lắp đặt khe co giãn răng lược thép 2 đầu cầu", "Task", "1.4.4SS+2", 4, 1.0, "22/01/2027", "26/01/2027", 4, "", 0.0),
        (26, "1.4.6", "Chống thấm mặt cầu & thảm bê tông nhựa C12.5 (dày 7cm)", "Task", "1.4.2FS", 9, 1.0, "26/01/2027", "04/02/2027", 0, "GĂNG", 0.0),
        (27, "1.4.7", "Hoàn thiện sơn kẻ vạch, biển báo, tứ nón & thu dọn bãi đúc", "Task", "1.4.6FS", 4, 1.0, "12/02/2027", "16/02/2027", 0, "GĂNG", 0.0),

        (28, "1.5", "Thử tải, nghiệm thu & bàn giao công trình", "Summary", "", 3, 1.0, "17/02/2027", "20/02/2027", 0, "GĂNG", 0.0),
        (29, "1.5.1", "Thử tải tĩnh và thử tải động công trình cầu L=15m", "Task", "1.4.7FS", 3, 1.0, "17/02/2027", "19/02/2027", 0, "GĂNG", 0.0),
        (30, "1.5.2", "Nghiệm thu hoàn thành bàn giao công trình đưa vào khai thác", "Milestone", "1.5.1FS", 0, 1.0, "20/02/2027", "20/02/2027", 0, "GĂNG", 0.0)
    ]

    for idx, t_data in enumerate(tasks_khai_hoang_2, start=6):
        stt, wbs, ten, k_type, preds, w_days, k_tt, s_date, e_date, tf, is_crit, pct = t_data
        ws_td[f"A{idx}"] = stt
        ws_td[f"B{idx}"] = wbs
        ws_td[f"C{idx}"] = ten
        ws_td[f"D{idx}"] = k_type
        ws_td[f"E{idx}"] = preds
        ws_td[f"F{idx}"] = w_days
        ws_td[f"G{idx}"] = k_tt

        # Parse date
        d_start = datetime.datetime.strptime(s_date, "%d/%m/%Y")
        d_end = datetime.datetime.strptime(e_date, "%d/%m/%Y")
        ws_td[f"H{idx}"] = d_start
        ws_td[f"H{idx}"].number_format = "dd/mm/yyyy"
        ws_td[f"I{idx}"] = d_end
        ws_td[f"I{idx}"].number_format = "dd/mm/yyyy"

        ws_td[f"J{idx}"] = (d_end - d_start).days + (1 if k_type != "Milestone" else 0)
        ws_td[f"K{idx}"] = tf
        ws_td[f"L{idx}"] = is_crit
        ws_td[f"M{idx}"] = pct
        ws_td[f"M{idx}"].number_format = "0.0%"

        # Styling
        if k_type == "Summary":
            for c_i in range(1, 14):
                c_cell = ws_td.cell(row=idx, column=c_i)
                c_cell.font = f_summary
                c_cell.fill = fill_summary
        elif k_type == "Milestone":
            for c_i in range(1, 14):
                c_cell = ws_td.cell(row=idx, column=c_i)
                c_cell.font = f_crit
                c_cell.fill = fill_crit
        else:
            for c_i in range(1, 14):
                c_cell = ws_td.cell(row=idx, column=c_i)
                c_cell.font = f_norm

        if is_crit in ["GĂNG", "GANG"]:
            ws_td[f"L{idx}"].font = f_crit
            ws_td[f"L{idx}"].fill = fill_crit

        # Border for row
        for c_i in range(1, 38):
            cell = ws_td.cell(row=idx, column=c_i)
            cell.border = box_border
            if c_i <= 13:
                col_l = get_column_letter(c_i)
                if col_l in ["A", "B", "D", "E", "G", "H", "I", "J", "K", "L", "M"]:
                    cell.alignment = align_center
                elif col_l == "C":
                    lvl = wbs.count('.')
                    cell.alignment = Alignment(horizontal="left", vertical="center", indent=lvl)

    # =========================================================================
    # 5. SHEET: EVM_5D_QUAN_TRI
    # =========================================================================
    print("[5/9] Thiết lập sheet EVM_5D_QUAN_TRI (Đóng băng Baseline & Live Formulas)...")
    ws_evm = wb["EVM_5D_QUAN_TRI"]
    ws_evm["A1"] = "HỆ THỐNG QUẢN TRỊ GIÁ TRỊ THU ĐƯỢC (EVM 5D) & ĐƯỜNG CƠ SỞ BASELINE – 23HG SYSTEM"
    ws_evm["A2"] = "Kiểm soát Tiến độ - Chi phí (PV, EV, AC, SV, CV, SPI, CPI) theo chuẩn quốc tế PMI & Nghị định 206/2026/NĐ-CP"

    # KPI Cards
    ws_evm["B5"] = "=BOQ_TIEN_DO!G35"
    ws_evm["D5"] = "=SUM(J11:J40)"
    ws_evm["F5"] = "=SUM(K11:K40)"
    ws_evm["H5"] = "=SUM(L11:L40)"
    ws_evm["J5"] = "=F5-D5"
    ws_evm["L5"] = "=IF(D5>0, F5/D5, 1)"
    ws_evm["N5"] = "=IF(H5>0, F5/H5, 1)"

    ws_evm["B8"] = "ĐÁNH GIÁ SỨC KHỎE DỰ ÁN (TẠI DATA DATE 15/12/2026):"
    ws_evm["E8"] = '=IF(L5>=1, "TIẾN ĐỘ: ĐẠT / VƯỢT KẾ HOẠCH (SPI >= 1.0)", "CẢNH BÁO: CHẬM TIẾN ĐỘ NHẸ (SPI < 1.0)")'
    ws_evm["K8"] = '=IF(N5>=1, "CHI PHÍ: TIẾT KIỆM SO VỚI DỰ TOÁN (CPI >= 1.0)", "CẢNH BÁO: VƯỢT DỰ TOÁN (CPI < 1.0)")'

    # Clear old rows 11-45
    for r in range(11, 45):
        for c in range(1, 15):
            ws_evm.cell(r, c).value = None

    # BAC map tương ứng với BOQ_TIEN_DO
    bac_map = {
        "1.1.1": 14400000, "1.1.2": 180000000, "1.1.3": 220000000,
        "1.2.1": 83250000, "1.2.2": 21560000, "1.2.3": 956897500,
        "1.2.4": 407000000, "1.2.5": 325017600, "1.2.6": 210000000, "1.2.7": 177500000,
        "1.3.1": 374016000, "1.3.2": 83250000, "1.3.3": 0, "1.3.4": 36000000, "1.3.5": 37000000,
        "1.4.1": 49968400, "1.4.2": 111720000, "1.4.3": 23180500, "1.4.4": 43500000,
        "1.4.5": 46720000, "1.4.6": 29925000, "1.4.7": 35000000, "1.5.1": 90000000, "1.5.2": 0
    }
    ac_map = {
        "1.1.1": 14200000, "1.1.2": 178000000, "1.1.3": 222000000,
        "1.2.1": 82000000, "1.2.2": 21000000, "1.2.3": 955000000,
        "1.2.4": 405000000, "1.2.5": 113000000, "1.2.6": 20000000, "1.2.7": 35000000,
        "1.3.1": 372000000, "1.3.2": 82500000, "1.3.3": 0
    }

    for idx, t_data in enumerate(tasks_khai_hoang_2, start=11):
        stt, wbs, ten, k_type, preds, w_days, k_tt, s_date, e_date, tf, is_crit, pct = t_data
        td_row = idx - 5 # Row in TIEN_DO is 6 to 35

        ws_evm[f"A{idx}"] = stt
        ws_evm[f"B{idx}"] = f"=TIEN_DO!B{td_row}"
        ws_evm[f"C{idx}"] = f"=TIEN_DO!C{td_row}"

        bac_val = bac_map.get(wbs, 0)
        ws_evm[f"D{idx}"] = bac_val
        ws_evm[f"D{idx}"].number_format = '#,##0 "VNĐ"'

        d_start = datetime.datetime.strptime(s_date, "%d/%m/%Y")
        d_end = datetime.datetime.strptime(e_date, "%d/%m/%Y")
        ws_evm[f"E{idx}"] = d_start
        ws_evm[f"E{idx}"].number_format = "dd/mm/yyyy"
        ws_evm[f"F{idx}"] = d_end
        ws_evm[f"F{idx}"].number_format = "dd/mm/yyyy"

        ws_evm[f"G{idx}"] = f"=TIEN_DO!H{td_row}"
        ws_evm[f"G{idx}"].number_format = "dd/mm/yyyy"
        ws_evm[f"H{idx}"] = f"=TIEN_DO!I{td_row}"
        ws_evm[f"H{idx}"].number_format = "dd/mm/yyyy"

        ws_evm[f"I{idx}"] = f"=TIEN_DO!M{td_row}"
        ws_evm[f"I{idx}"].number_format = "0.0%"

        ws_evm[f"J{idx}"] = f"=ROUND(D{idx} * IF(TIEN_DO!$C$3>=F{idx}, 1, IF(TIEN_DO!$C$3<=E{idx}, 0, (TIEN_DO!$C$3-E{idx})/(F{idx}-E{idx}))), 0)"
        ws_evm[f"J{idx}"].number_format = '#,##0 "VNĐ"'

        ws_evm[f"K{idx}"] = f"=ROUND(D{idx} * I{idx}, 0)"
        ws_evm[f"K{idx}"].number_format = '#,##0 "VNĐ"'

        ac_val = ac_map.get(wbs, 0)
        ws_evm[f"L{idx}"] = ac_val
        ws_evm[f"L{idx}"].number_format = '#,##0 "VNĐ"'

        ws_evm[f"M{idx}"] = f"=IF(J{idx}>0, K{idx}/J{idx}, 1)"
        ws_evm[f"M{idx}"].number_format = "0.00"

        ws_evm[f"N{idx}"] = f"=IF(L{idx}>0, K{idx}/L{idx}, 1)"
        ws_evm[f"N{idx}"].number_format = "0.00"

        for col_l in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"]:
            c_cell = ws_evm[f"{col_l}{idx}"]
            c_cell.font = f_norm
            c_cell.border = box_border
            if col_l in ["A", "B", "E", "F", "G", "H"]: c_cell.alignment = align_center
            elif col_l in ["D", "I", "J", "K", "L", "M", "N"]: c_cell.alignment = align_right
            else: c_cell.alignment = align_left

    # =========================================================================
    # 6. SHEET: HUY_DONG_XMTB
    # =========================================================================
    print("[6/9] Thiết lập sheet HUY_DONG_XMTB (12 thiết bị thi công cầu qua 5 tháng)...")
    ws_xm = wb["HUY_DONG_XMTB"]
    ws_xm["A1"] = "KẾ HOẠCH HUY ĐỘNG & CÂN BẰNG TÀI NGUYÊN XE MÁY THIẾT BỊ (XMTB)"
    ws_xm["A2"] = "Kiểm soát đỉnh tải thiết bị thi công Cầu thôn Khai Hoang 2 (Km 14+363.65) – Chống lãng phí máy chờ việc"

    # Headers for months in Row 4
    ws_xm["F4"] = "T10/26"
    ws_xm["G4"] = "T11/26"
    ws_xm["H4"] = "T12/26"
    ws_xm["I4"] = "T01/27"
    ws_xm["J4"] = "T02/27"
    for c_i in range(11, 30):
        ws_xm.cell(4, c_i).value = None

    for r in range(5, 20):
        for c in range(1, 30):
            ws_xm.cell(r, c).value = None

    bridge_equipment = [
        (1, "Máy đào bánh xích gầu nghịch 1.25 m3", "Komatsu PC200-8 / Hitachi ZX200", "Chiếc", 2, [1, 2, 2, 1, 1]),
        (2, "Máy đào gắn búa đập đá thủy lực", "Búa đập áp lực cao 180 Bar", "Chiếc", 2, [0, 2, 1, 0, 0]),
        (3, "Máy khoan đá khí nén tạo lỗ neo D32", "Máy khoan ép khí nén Atlas Copco", "Chiếc", 2, [0, 2, 1, 0, 0]),
        (4, "Ô tô tự đổ chuyên dụng 15 tấn", "Howo 3 chân vận chuyển đất đá", "Chiếc", 4, [2, 4, 4, 2, 2]),
        (5, "Máy ủi bánh xích 110 CV", "Komatsu D31P san ủi mặt bằng", "Chiếc", 1, [1, 1, 1, 0, 0]),
        (6, "Máy lu rung tự hành 14 tấn", "Hamm 3411 lực rung 28T", "Chiếc", 1, [1, 1, 1, 1, 1]),
        (7, "Trạm trộn bê tông xi măng 30 m3/h", "Trạm trộn cưỡng bức JS500 hiện trường", "Trạm", 1, [0, 1, 1, 1, 1]),
        (8, "Xe bồn vận chuyển bê tông 6 m3", "Howo bồn trộn 6m3 chuyên dụng", "Chiếc", 2, [0, 2, 2, 2, 1]),
        (9, "Máy bơm bê tông tĩnh 60 m3/h", "Bơm tĩnh Schwing / Putzmeister", "Chiếc", 1, [0, 1, 1, 1, 0]),
        (10, "Cần trục bánh xích 50 tấn lao dầm T15m", "Hitachi KH180-3 sức nâng 50T", "Chiếc", 1, [0, 0, 1, 1, 0]),
        (11, "Máy phát điện dự phòng 100 kVA", "Denyo 100kVA cấp điện liên tục", "Chiếc", 2, [1, 2, 2, 2, 2]),
        (12, "Trạm gia công uốn cắt thép CNC RebarCut", "Dàn máy duỗi cắt uốn tự động", "Trạm", 1, [1, 1, 1, 1, 0])
    ]

    for idx, eq in enumerate(bridge_equipment, start=5):
        stt, name, spec, dv, cap, alloc = eq
        ws_xm[f"A{idx}"] = stt
        ws_xm[f"B{idx}"] = name
        ws_xm[f"C{idx}"] = spec
        ws_xm[f"D{idx}"] = dv
        ws_xm[f"E{idx}"] = cap

        for m_i, val in enumerate(alloc, start=6):
            cell = ws_xm.cell(row=idx, column=m_i, value=val)
            cell.font = f_norm
            cell.alignment = align_center
            cell.border = box_border
            if val >= cap:
                cell.font = Font(name="Segoe UI", size=9, bold=True, color="B91C1C")
                cell.fill = fill_crit

        for col_l in ["A", "B", "C", "D", "E"]:
            cell = ws_xm[f"{col_l}{idx}"]
            cell.border = box_border
            if col_l in ["A", "D", "E"]: cell.alignment = align_center
            else: cell.alignment = align_left

    # =========================================================================
    # 7. SHEET: TIEN_DO_GIAI_NGAN
    # =========================================================================
    print("[7/9] Thiết lập sheet TIEN_DO_GIAI_NGAN (Dòng tiền hợp đồng & tạm ứng 20%)...")
    ws_gn = wb["TIEN_DO_GIAI_NGAN"]
    ws_gn["A1"] = "KẾ HOẠCH DÒNG TIỀN & TIẾN ĐỘ GIẢI NGÂN (S-CURVE TÀI CHÍNH)"
    ws_gn["A2"] = "Dự án: Cầu thôn Khai Hoang 2 – Tuân thủ NĐ 206/2026/NĐ-CP: Tạm ứng 20%, thu hồi từ 30%-80%, giữ bảo hành 5%"

    for r in range(5, 20):
        for c in range(1, 9):
            ws_gn.cell(r, c).value = None

    cashflow_bridge = [
        (0, "T10/2026", "TẠM ỨNG HỢP ĐỒNG 20% (Theo NĐ 206/2026/NĐ-CP)", 0, 0, 0, "=ROUND(BOQ_TIEN_DO!$G$35 * 0.20, 0)", "=G5"),
        (1, "T11/2026", "Đào đá hố móng, cấy neo D32, đổ bê tông lót & bệ mố M1, M2", "=ROUND(BOQ_TIEN_DO!$G$35 * 0.22, 0)", 0, "=ROUND(D6*0.05, 0)", "=D6-E6-F6", "=H5+G6"),
        (2, "T12/2026", "Thân mố, tường ngực M1-M2, đúc 4 phiến dầm T15m & tường chắn", "=ROUND(BOQ_TIEN_DO!$G$35 * 0.28, 0)", "=ROUND(H$5*0.35, 0)", "=ROUND(D7*0.05, 0)", "=D7-E7-F7", "=H6+G7"),
        (3, "T01/2027", "Lao lắp dầm T15m, đổ BT mối nối ướt, bản mặt cầu & bản quá độ", "=ROUND(BOQ_TIEN_DO!$G$35 * 0.30, 0)", "=ROUND(H$5*0.35, 0)", "=ROUND(D8*0.05, 0)", "=D8-E8-F8", "=H7+G8"),
        (4, "T02/2027", "Khe co giãn, lan can, thảm BTN C12.5, thử tải cầu & nghiệm thu", "=ROUND(BOQ_TIEN_DO!$G$35 * 0.20, 0)", "=ROUND(H$5*0.30, 0)", "=ROUND(D9*0.05, 0)", "=D9-E9-F9", "=H8+G9"),
        (5, "T02/2028", "HOÀN TRẢ TIỀN GIỮ BẢO HÀNH CÔNG TRÌNH 5% (Hết hạn 12 tháng)", 0, 0, 0, "=SUM(F6:F9)", "=H9+G10")
    ]

    for idx, r_data in enumerate(cashflow_bridge, start=5):
        ky, time_lbl, content, d_val, e_val, f_val, g_val, h_val = r_data
        ws_gn[f"A{idx}"] = ky
        ws_gn[f"B{idx}"] = time_lbl
        ws_gn[f"C{idx}"] = content
        ws_gn[f"D{idx}"] = d_val
        ws_gn[f"E{idx}"] = e_val
        ws_gn[f"F{idx}"] = f_val
        ws_gn[f"G{idx}"] = g_val
        ws_gn[f"H{idx}"] = h_val

        ws_gn[f"D{idx}"].number_format = '#,##0 "VNĐ"'
        ws_gn[f"E{idx}"].number_format = '#,##0 "VNĐ"'
        ws_gn[f"F{idx}"].number_format = '#,##0 "VNĐ"'
        ws_gn[f"G{idx}"].number_format = '#,##0 "VNĐ"'
        ws_gn[f"H{idx}"].number_format = '#,##0 "VNĐ"'

        for col_l in ["A", "B", "C", "D", "E", "F", "G", "H"]:
            cell = ws_gn[f"{col_l}{idx}"]
            cell.border = box_border
            if col_l in ["A", "B"]: cell.alignment = align_center
            elif col_l in ["D", "E", "F", "G", "H"]: cell.alignment = align_right
            else: cell.alignment = align_left

    # =========================================================================
    # 8. SHEET: KE_HOACH_QLCL
    # =========================================================================
    print("[8/9] Thiết lập sheet KE_HOACH_QLCL (22 BBNT KCS & Link động TIEN_DO)...")
    ws_qlcl = wb["KE_HOACH_QLCL"]
    ws_qlcl["A1"] = "KẾ HOẠCH NGHIỆM THU & QUẢN LÝ CHẤT LƯỢNG (AEC-QLCL 23HG SYSTEM)"
    ws_qlcl["A2"] = "Dự án: Cầu thôn Khai Hoang 2 – Tuân thủ Luật XD 135/2025, NĐ 207/2026 & TT 32/2026/TT-BXD – Tự động đồng bộ ngày từ TIEN_DO"

    for r in range(5, 35):
        for c in range(1, 8):
            ws_qlcl.cell(r, c).value = None

    qlcl_bridge = [
        ("NT-001", "Bàn giao cọc mốc tim tuyến và định vị công trình cầu", "Km 14+363.65 (WBS 1.1.1)", "=TIEN_DO!H8", "=TIEN_DO!H8", "TCVN 4054:2005", "Biên bản bàn giao tim mốc"),
        ("NT-002", "Nghiệm thu đường công vụ và đê quây nắn dòng suối tạm", "Lòng suối Km14+363 (WBS 1.1.2)", "=TIEN_DO!H9", "=TIEN_DO!I9", "TCVN 4447:2012", "Biên bản đo vẽ mặt cắt suối"),
        ("NT-003", "Kiểm định chất lượng vật liệu đầu vào (Thép, Xi măng, Cát, Đá)", "Toàn dự án (WBS 1.1.3)", "=TIEN_DO!H10", "=TIEN_DO!I10", "TCVN 1651:2018", "Chứng chỉ CO/CQ & Kết quả kéo thép Lab"),
        ("NT-004", "Nghiệm thu lắp dựng bãi đúc dầm T15m & trạm cốt thép CNC", "Hiện trường (WBS 1.1.3)", "=TIEN_DO!H10", "=TIEN_DO!I10", "TCVN 11823:2017", "Kiểm tra độ bằng phẳng bệ đúc"),
        ("NT-005", "Nghiệm thu đào hố móng mố M1 và xử lý đáy móng đá", "Mố M1 (WBS 1.2.1)", "=TIEN_DO!H12", "=TIEN_DO!I12", "TCVN 11823:2017", "Biên bản đo cao độ đáy móng đá"),
        ("NT-006", "Nghiệm thu khoan cấy 49 cây thép neo D32 vào nền đá mố M1", "Mố M1 (WBS 1.2.2)", "=TIEN_DO!H13", "=TIEN_DO!I13", "TCVN 11823:2017", "Thí nghiệm kéo nhổ neo SikaGrout"),
        ("NT-007", "Nghiệm thu đổ bê tông đệm móng mố M1 mác C10 (dày 10cm)", "Mố M1 (WBS 1.2.3)", "=TIEN_DO!H14", "=TIEN_DO!H14+1", "TCVN 3105:2022", "Ép mẫu nén R28 bê tông lót"),
        ("NT-008", "Nghiệm thu cốt thép và ván khuôn bệ mố M1", "Mố M1 (WBS 1.2.3)", "=TIEN_DO!H14+2", "=TIEN_DO!I14", "TCVN 1651:2018", "Biên bản kiểm tra cốt thép theo RebarCut"),
        ("NT-009", "Nghiệm thu đổ bê tông bệ móng mố M1 mác C30 (M350)", "Mố M1 (WBS 1.2.4)", "=TIEN_DO!H15", "=TIEN_DO!I15", "TCVN 3105:2022", "Ép mẫu nén R7, R28 bê tông bệ mố"),
        ("NT-010", "Nghiệm thu đào hố móng và khoan cấy thép neo đá mố M2", "Mố M2 (WBS 1.2.1-1.2.2)", "=TIEN_DO!H12+5", "=TIEN_DO!I13", "TCVN 11823:2017", "Biên bản cao độ & thử kéo neo đá"),
        ("NT-011", "Nghiệm thu đổ bê tông đệm và bê tông bệ móng mố M2 C30", "Mố M2 (WBS 1.2.3-1.2.4)", "=TIEN_DO!H14", "=TIEN_DO!I15", "TCVN 3105:2022", "Phiếu kết quả nén mẫu bê tông M2"),
        ("NT-012", "Nghiệm thu thi công thân mố, tường ngực và tường cánh M1 C30", "Mố M1 (WBS 1.2.5)", "=TIEN_DO!H16", "=TIEN_DO!I16", "TCVN 11823:2017", "Kiểm tra tim trục, kích thước hình học"),
        ("NT-013", "Nghiệm thu thi công thân mố, tường ngực và tường cánh M2 C30", "Mố M2 (WBS 1.2.5-1.2.6)", "=TIEN_DO!H16+5", "=TIEN_DO!I17", "TCVN 11823:2017", "Kiểm tra độ thẳng đứng & nén mẫu"),
        ("NT-014", "Nghiệm thu đúc 4 phiến dầm chủ T15m mác C40 tại bãi đúc", "Bãi đúc dầm (WBS 1.3.1-1.3.2)", "=TIEN_DO!H20", "=TIEN_DO!I21", "TCVN 11823:2017", "Ép mẫu nén R28 dầm T15m C40"),
        ("NT-015", "Nghiệm thu bảo dưỡng bê tông dầm T15m và xuất xưởng", "Bãi đúc dầm (WBS 1.3.3)", "=TIEN_DO!H22", "=TIEN_DO!I22", "TCVN 3105:2022", "Chứng chỉ xuất xưởng dầm T15m"),
        ("NT-016", "Nghiệm thu lắp đặt đá kê gối và 8 gối cao su bản thép", "Mố M1 & M2 (WBS 1.3.4)", "=TIEN_DO!H23", "=TIEN_DO!I23", "TCVN 11823:2017", "CO/CQ gối cao su & cao độ đỉnh gối"),
        ("NT-017", "Nghiệm thu lao lắp 4 phiến dầm chủ T15m vào gối mố M1-M2", "Nhịp cầu 15m (WBS 1.3.5)", "=TIEN_DO!H24", "=TIEN_DO!I24", "TCVN 11823:2017", "Biên bản trắc đạc tim dầm & kê chèn"),
        ("NT-018", "Nghiệm thu thi công dầm ngang & mối nối ướt dầm dọc C30", "Nhịp cầu 15m (WBS 1.4.1)", "=TIEN_DO!H26", "=TIEN_DO!I26", "TCVN 11823:2017", "Ép mẫu mối nối dầm liên tục"),
        ("NT-019", "Nghiệm thu cốt thép & đổ bê tông bản mặt cầu C35 dày 12cm", "Mặt cầu (WBS 1.4.2)", "=TIEN_DO!H27", "=TIEN_DO!I27", "TCVN 3105:2022", "Độ chặt, độ bằng phẳng thước 3m"),
        ("NT-020", "Nghiệm thu thi công gờ lan can C25 & lan can thép mạ kẽm", "2 bên cầu (WBS 1.4.4)", "=TIEN_DO!H29", "=TIEN_DO!I29", "TCVN 11823:2017", "Kiểm tra chiều cao lan can & bu lông neo"),
        ("NT-021", "Nghiệm thu thi công bản quá độ 2 đầu mố M1, M2 & tường chắn", "Đầu cầu (WBS 1.4.3 & 1.2.7)", "=TIEN_DO!H28", "=TIEN_DO!I28", "TCVN 11823:2017", "Độ đầm chặt đất đắp K95 sau mố"),
        ("NT-022", "Nghiệm thu chống thấm mặt cầu & thảm BTN C12.5 (dày 7cm)", "Mặt cầu (WBS 1.4.6)", "=TIEN_DO!H31", "=TIEN_DO!I31", "TCVN 8819:2011", "Độ nhám con lắc Anh + Độ chặt K"),
        ("GĐ-001", "NGHIỆM THU HOÀN THÀNH TOÀN BỘ CÔNG TRÌNH CẦU & BÀN GIAO", "Toàn cầu (WBS 1.5.1-1.5.2)", "=TIEN_DO!H34", "=TIEN_DO!I35", "Nghị định 207/2026/NĐ-CP", "Báo cáo thử tải cầu & Trọn bộ hồ sơ hoàn công")
    ]

    for idx, row_item in enumerate(qlcl_bridge, start=5):
        for c_idx, val in enumerate(row_item, start=1):
            cell = ws_qlcl.cell(row=idx, column=c_idx, value=val)
            cell.font = f_norm
            cell.border = box_border
            if c_idx in [1, 4, 5]: cell.alignment = align_center
            else: cell.alignment = align_left

    # =========================================================================
    # 9. SHEET: HUONG_DAN
    # =========================================================================
    print("[9/9] Thiết lập sheet HUONG_DAN (Cẩm nang vận hành Cầu thôn Khai Hoang 2)...")
    ws_hd = wb["HUONG_DAN"]
    ws_hd["A1"] = "CẨM NANG VẬN HÀNH HỆ THỐNG QUẢN LÝ TIẾN ĐỘ THI CÔNG – 23HG SYSTEM PRO"
    ws_hd["A2"] = "Công trình: Cầu thôn Khai Hoang 2 (Km14+363.65) – Hà Giang | Tác giả: NBT | Bản quyền 23HG SYSTEM"

    # Save to temp file
    temp_xlsm = os.path.join(OUT_SYS_DIR, "_temp_gen.xlsm")
    os.makedirs(OUT_SYS_DIR, exist_ok=True)
    print(f"[*] Đang lưu file tạm: {temp_xlsm}...")
    wb.save(temp_xlsm)

    # Re-inject customUI14.xml relationship into _rels/.rels
    print("[*] Đang khôi phục mối liên kết CustomUI Ribbon XML trong gói ZIP...")
    temp_zip = os.path.join(OUT_SYS_DIR, "_temp_ribbon.xlsm")
    with zipfile.ZipFile(temp_xlsm, 'r') as zin, zipfile.ZipFile(temp_zip, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            content = zin.read(item.filename)
            if item.filename == '_rels/.rels':
                rels_str = content.decode('utf-8', errors='ignore')
                if "customUI/customUI14.xml" not in rels_str:
                    rel_entry = '<Relationship Id="rIdCustomUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'
                    rels_str = rels_str.replace('</Relationships>', rel_entry + '</Relationships>')
                    content = rels_str.encode('utf-8')
            zout.writestr(item, content)

    # Distribute to destinations
    targets = [
        os.path.join(OUT_USER_DIR, FILE_NAME),
        os.path.join(OUT_PRO_DIR, FILE_NAME),
        os.path.join(OUT_SYS_DIR, FILE_NAME)
    ]

    for tgt in targets:
        os.makedirs(os.path.dirname(tgt), exist_ok=True)
        shutil.copy2(temp_zip, tgt)
        print(f"  [+] Đã xuất bản thành công: {tgt}")

    # Clean up temp
    if os.path.exists(temp_xlsm): os.remove(temp_xlsm)
    if os.path.exists(temp_zip): os.remove(temp_zip)

    # Excel COM automation: recalculate CPM and render Gantt chart
    print("\n[*] Đang khởi động Excel COM để tính toán CPM và vẽ Gantt Canvas tự động...")
    try:
        import win32com.client
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False

        tgt_main = targets[0]
        wb_com = xl.Workbooks.Open(os.path.abspath(tgt_main))

        # Run macros if available
        try:
            xl.Run("CalculateCPM", wb_com.Sheets("TIEN_DO"), False)
            print("  [+] Đã thực thi Macro CalculateCPM thành công!")
        except Exception as e_cpm:
            print(f"  [-] Lưu ý Macro CalculateCPM: {e_cpm}")

        try:
            xl.Run("RenderGantt", wb_com.Sheets("TIEN_DO"), "QUY")
            print("  [+] Đã thực thi Macro RenderGantt thành công!")
        except Exception as e_gantt:
            print(f"  [-] Lưu ý Macro RenderGantt: {e_gantt}")

        wb_com.Save()
        wb_com.Close(SaveChanges=True)
        xl.Quit()

        # Copy the updated COM workbook to the other destinations
        for tgt in targets[1:]:
            shutil.copy2(tgt_main, tgt)
            print(f"  [+] Đã đồng bộ bản vẽ Gantt: {tgt}")

    except Exception as e_com:
        print(f"  [-] Không thể chạy Excel COM (sử dụng bản tính sẵn có): {e_com}")

    print("\n" + "=" * 80)
    print("   HOÀN THÀNH XUẤT TIẾN ĐỘ 23HG PRO CHO CẦU THÔN KHAI HOANG 2 (KM 14+363.65)   ")
    print("=" * 80)

if __name__ == "__main__":
    build_khai_hoang_2_pro_schedule()
