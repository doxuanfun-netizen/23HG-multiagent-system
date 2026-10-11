# -*- coding: utf-8 -*-
"""
23HG SYSTEM - BRIDGE KM19+529.080 MASTER SCHEDULE BUILDER
Tác giả: NBT (Nguyễn Bảo Tú) | GitHub: @baotuhg | 23HG SYSTEM
Đồng bộ 100% kiến trúc 23HG Schedule Assistant Pro với Hồ sơ gốc Cầu Km19+529.080
Tách biệt triệt để Trụ T1, T2; Bệ móng 11.6x8.0x2.0m Zero-vát; 3 đốt thân có vát.
"""

import os
import sys
import datetime as dt
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter as L
import xml.etree.ElementTree as ET

# Output paths
OUT_BRIDGE_ASBUILT = r"C:\Users\baotu\Downloads\HSTK Cầu Km19+529.080_Marker\01_HIEN_TRUONG_QLCL_KCS\So_Do_Tien_Do_Truc_Quan_AsBuilt_Cau_Km19.xlsx"
OUT_MASTER_DOSSIER = r"C:\Users\baotu\Downloads\HSTK Cầu Km19+529.080_Marker\00_BAN_CHI_HUY_MASTER\Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx"
OUT_EQUIPMENT_PLAN = r"C:\Users\baotu\Downloads\HSTK Cầu Km19+529.080_Marker\04_CO_GIOI_THIET_BI_VA_DAU_DIEZEL\260920_TDTC_CaXe_CaMay_DauDiezel_Cau_Km19+529.080.xlsx"
OUT_XML_PROJECT = r"C:\Users\baotu\Downloads\HSTK Cầu Km19+529.080_Marker\00_BAN_CHI_HUY_MASTER\Tien_Do_Thi_Cong_Cau_Km19+529.080.xml"
REPO_DIR = r"d:\Code\23HG-multiagent-system-main\23HG-multiagent-system-main\examples\HO_SO_CAU_KM19_529"

# Import CPM engine
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(SCRIPT_DIR)
from cpm_engine import Calendar, Task, schedule

# Working calendar: 6 days/week (Monday to Saturday, Sundays off)
holidays = [
    dt.date(2027, 1, 1), # Tet Duong Lich
    dt.date(2027, 2, 5), dt.date(2027, 2, 6), dt.date(2027, 2, 7),
    dt.date(2027, 2, 8), dt.date(2027, 2, 9), dt.date(2027, 2, 10),
    dt.date(2027, 2, 11), dt.date(2027, 2, 12) # Tet Nguyen Dan
]
cal = Calendar(holidays=holidays)

tasks_def = [
    Task(wbs='1', kind='Summary', name='CÔNG TÁC CHUẨN BỊ & PHỤ TRỢ THI CÔNG CÔNG TRƯỜNG'),
    Task(wbs='1.1', kind='Task', duration=2, name='Bàn giao tim mốc VN2000 & định vị mố trụ cầu'),
    Task(wbs='1.2', kind='Task', duration=5, preds=[('1.1', 'FS', 0)], name='Rà phá bom mìn, dọn dẹp mặt bằng & dựng lán trại'),
    Task(wbs='1.3', kind='Task', duration=9, preds=[('1.2', 'FS', 0)], name='Thi công đường công vụ nhánh 6 & 6A tiếp cận mố trụ'),
    Task(wbs='1.4', kind='Task', duration=12, preds=[('1.3', 'SS', 3)], name='Xây dựng bãi đúc dầm Super-T, bệ đúc & trạm trộn 60m3/h'),

    Task(wbs='2', kind='Summary', name='THI CÔNG CỌC KHOAN NHỒI D1200MM (26 CỌC TOÀN CẦU)'),
    Task(wbs='2.1', kind='Task', duration=10, preds=[('1.3', 'FS', 0)], name='Khoan thăm dò hang Karst sâu 5m vào đá liền khối (26 lỗ)'),
    Task(wbs='2.2', kind='Task', duration=6, preds=[('2.1', 'FS', 0)], name='Khoan cọc nhồi D1200 Mố M1 (3 cọc L=20m) & đổ BT C30'),
    Task(wbs='2.3', kind='Task', duration=15, preds=[('2.2', 'SS', 3)], name='Khoan cọc nhồi D1200 Trụ T1 (8 cọc L=40m) & đổ BT C30'),
    Task(wbs='2.4', kind='Task', duration=13, preds=[('2.3', 'SS', 6)], name='Khoan cọc nhồi D1200 Trụ T2 (8 cọc L=30m) & đổ BT C30'),
    Task(wbs='2.5', kind='Task', duration=14, preds=[('2.4', 'SS', 5)], name='Khoan cọc nhồi D1200 Mố M2 (7 cọc L=36m) & đổ BT C30'),
    Task(wbs='2.6', kind='Task', duration=6, preds=[('2.5', 'FS', 0)], name='Thí nghiệm siêu âm 156 mặt cắt cọc, PDA & nén tĩnh'),

    Task(wbs='3', kind='Summary', name='KẾT CẤU PHẦN DƯỚI (MỐ M1, M2 & TRỤ T1, T2 ĐÃ TÁCH BẠCH)'),
    Task(wbs='3.1', kind='Task', duration=6, preds=[('2.6', 'FS', 0)], name='Đào hố móng ngàm đá, đập đầu 26 cọc & đổ BT lót C10 (42.15 m3)'),
    Task(wbs='3.2', kind='Task', duration=7, preds=[('3.1', 'FS', 0)], name='Cốt thép, ván khuôn & đổ bê tông bệ móng mố M1, M2 (173.2 m3)'),
    Task(wbs='3.3.1', kind='Task', duration=7, preds=[('3.1', 'FS', 2)], name='Cốt thép, ván khuôn & đổ BT bệ trụ T1 (11.6x8.0x2.0m Zero-vát, 184.2 m3)'),
    Task(wbs='3.3.2', kind='Task', duration=7, preds=[('3.3.1', 'SS', 4)], name='Cốt thép, ván khuôn & đổ BT bệ trụ T2 (11.6x8.0x2.0m Zero-vát, 184.2 m3)'),
    Task(wbs='3.4.1', kind='Task', duration=11, preds=[('3.3.1', 'FS', 0)], name='Thi công thân đặc trụ T1 (H=23.35m, 3 đốt vát, 234.9 m3)'),
    Task(wbs='3.4.2', kind='Task', duration=12, preds=[('3.3.2', 'FS', 0)], name='Thi công thân đặc trụ T2 (H=25.45m, 3 đốt vát, 266.8 m3)'),
    Task(wbs='3.5', kind='Task', duration=10, preds=[('3.2', 'FS', 1)], name='Thi công thân mố chân dê M1 & mố chữ U M2 BT C30 (138.4 m3)'),
    Task(wbs='3.6.1', kind='Task', duration=6, preds=[('3.4.1', 'FS', 0)], name='Thi công xà mũ trụ T1, đá kê gối & ụ chống xô (128.9 m3)'),
    Task(wbs='3.6.2', kind='Task', duration=7, preds=[('3.4.2', 'FS', 0)], name='Thi công xà mũ trụ T2, đá kê gối & ụ chống xô (145.0 m3)'),

    Task(wbs='4', kind='Summary', name='CHẾ TẠO & LAO LẮP DẦM SUPER-T L=38.2M (15 PHIẾN)'),
    Task(wbs='4.1', kind='Task', duration=24, preds=[('1.4', 'FS', 25)], name='Cốt thép, luồn cáp DUL 15 phiến dầm Super-T tại bãi đúc (15 phiến)'),
    Task(wbs='4.2', kind='Task', duration=18, preds=[('4.1', 'SS', 18)], name='Đổ bê tông C45 15 phiến dầm Super-T L=38.2m bãi đúc (434.8 m3)'),
    Task(wbs='4.3', kind='Task', duration=8, preds=[('4.2', 'FS', 0)], name='Căng kéo cáp DUL 15.2mm & bơm vữa ống gen dầm (15 phiến)'),
    Task(wbs='4.4', kind='Task', duration=4, preds=[('3.6.1', 'FS', 0), ('3.6.2', 'FS', 0)], name='Lắp đặt 30 gối chậu cao su đơn/song hướng mố trụ (30 cái)'),
    Task(wbs='4.5', kind='Task', duration=7, preds=[('4.3', 'FS', 0), ('4.4', 'FS', 0)], name='Lắp đường ray P43, vận chuyển & lao lắp 15 dầm Super-T (15 phiến)'),

    Task(wbs='5', kind='Summary', name='BẢN MẶT CẦU & HOÀN THIỆN TOÀN CẦU'),
    Task(wbs='5.1', kind='Task', duration=4, preds=[('4.5', 'FS', 0)], name='Đổ bê tông dầm ngang mố trụ & bản liên tục nhiệt C35 (78.5 m3)'),
    Task(wbs='5.2', kind='Task', duration=4, preds=[('5.1', 'FS', 0)], name='Lắp 510 tấm ván khuôn đúc sẵn & cốt thép bản mặt cầu (510 tấm)'),
    Task(wbs='5.3', kind='Task', duration=4, preds=[('5.2', 'FS', 0)], name='Đổ bê tông bản mặt cầu tại chỗ C35 (dày 180mm, 235.6 m3)'),
    Task(wbs='5.4', kind='Task', duration=5, preds=[('5.3', 'SS', 2), ('3.5', 'FS', 20)], name='Đổ bê tông bản quá độ C25 & đắp K98 đường đầu cầu (2 bản)'),
    Task(wbs='5.5', kind='Task', duration=4, preds=[('5.3', 'FS', 1)], name='Lắp khe co giãn răng lược thép & ống thoát nước PVC (2 bộ)'),
    Task(wbs='5.6', kind='Task', duration=5, preds=[('5.3', 'FS', 1)], name='Đổ bê tông gờ lan can C25 & lắp dựng tay vịn thép (260.8 m)'),
    Task(wbs='5.7', kind='Task', duration=3, preds=[('5.5', 'FS', 0), ('5.6', 'FS', 0)], name='Chống thấm & thảm bê tông nhựa C16 mặt cầu dày 7cm (1,620 m2)'),
    Task(wbs='5.8', kind='Milestone', duration=0, preds=[('5.7', 'FS', 0)], name='NGHIỆM THU HOÀN THÀNH CÔNG TRÌNH BÀN GIAO THÔNG XE CẦU KM19')
]

# Run CPM
cpm_results = schedule(tasks_def, dt.date(2026, 10, 1), cal)
print(f"CPM Calculated: {len(cpm_results)} tasks. Finish date: {cpm_results['5.8'].ef.strftime('%d/%m/%Y')}")

# Styling palette (Executive Heritage Theme: Coffee & Sand)
f_title = Font(name="Segoe UI", size=13, bold=True, color="5A3718")
f_sub = Font(name="Segoe UI", size=9, italic=True, color="6E3812")
f_banner = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
f_sec_title = Font(name="Segoe UI", size=10.5, bold=True, color="5A3718")
f_title_brown = Font(name="Segoe UI", size=10, bold=True, color="6E3812")
f_header = Font(name="Segoe UI", size=8.5, bold=True, color="333333")
f_summary = Font(name="Segoe UI", size=9, bold=True, color="1E293B")
f_norm = Font(name="Segoe UI", size=9, color="000000")
f_norm_bold = Font(name="Segoe UI", size=8.5, bold=True, color="000000")
f_crit = Font(name="Segoe UI", size=9, bold=True, color="DC2626")

fill_coffee = PatternFill(fill_type="solid", start_color="5A3718", end_color="5A3718")
fill_cream = PatternFill(fill_type="solid", start_color="F7EDE2", end_color="F7EDE2")
fill_sand = PatternFill(fill_type="solid", start_color="EFE2D3", end_color="EFE2D3")
fill_summary = PatternFill(fill_type="solid", start_color="FBF8F5", end_color="FBF8F5")
fill_crit = PatternFill(fill_type="solid", start_color="FEE2E2", end_color="FEE2E2")
fill_bar_crit = PatternFill(fill_type="solid", start_color="DC2626", end_color="DC2626")
fill_bar_norm = PatternFill(fill_type="solid", start_color="D97706", end_color="D97706")
fill_bar_sum = PatternFill(fill_type="solid", start_color="5A3718", end_color="5A3718")

box_border = Border(
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1'),
    top=Side(style='thin', color='CBD5E1'),
    bottom=Side(style='thin', color='CBD5E1')
)

align_center = Alignment(horizontal="center", vertical="center")
align_left = Alignment(horizontal="left", vertical="center")
align_right = Alignment(horizontal="right", vertical="center")

def build_master_workbook():
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Remove default sheet

    # =========================================================================
    # 1. SHEET NGAY_NGHI_LE
    # =========================================================================
    print("1. Đang xây dựng Sheet NGAY_NGHI_LE...")
    ws_nl = wb.create_sheet("NGAY_NGHI_LE")
    ws_nl["A1"] = "LỊCH THỜI TIẾT MÙA MƯA & CÁC NGÀY NGHỈ LỄ CÔNG TRÌNH CẦU KM19+529.080 – 23HG SYSTEM"
    ws_nl["A1"].font = f_title
    ws_nl["A2"] = "Lịch công trường chuẩn: 6 ngày/tuần (Nghỉ Chủ nhật). Tuân thủ Bộ luật Lao động 2019 & TCVN 8819:2011"
    ws_nl["A2"].font = f_sub

    ws_nl["A4"] = "I. DANH MỤC CÁC NGÀY NGHỈ LỄ CHÍNH THỨC TRONG THỜI GIAN THI CÔNG CẦU KM19 (10/2026 - 03/2027)"
    ws_nl["A4"].font = f_sec_title

    headers_nl = [("A5", "STT"), ("B5", "Thời Gian"), ("C5", "Nội Dung Ngày Nghỉ Lễ / Sự Kiện"), 
                  ("D5", "Số Ngày Nghỉ"), ("E5", "Hệ Số Năng Suất K_tt"), ("F5", "Quy Định Điều Hành Công Trường")]
    for pos, txt in headers_nl:
        ws_nl[pos] = txt
        ws_nl[pos].font = f_header
        ws_nl[pos].fill = fill_sand
        ws_nl[pos].alignment = align_center

    holidays_data = [
        (1, "01/01/2027", "Nghỉ Tết Dương Lịch 2027", 1, 0.0, "Nghỉ lễ 1 ngày, trực ban bảo vệ công trường 24/24"),
        (2, "05/02/2027 - 12/02/2027", "Nghỉ Tết Nguyên Đán Đinh Mùi 2027 (29 Tết - Mùng 6 Tết)", 8, 0.0, "Dừng thi công hiện trường, niêm phong kho vật tư & bãi đúc dầm"),
        (3, "Chủ nhật hàng tuần", "Nghỉ Chủ Nhật hàng tuần theo Luật Lao động", 24, 0.0, "Nghỉ ngơi, bảo dưỡng máy móc thiết bị (Bauer BG25, Kobelco CKE500)")
    ]
    for idx, item in enumerate(holidays_data, start=6):
        for c_idx, val in enumerate(item, start=1):
            cell = ws_nl.cell(row=idx, column=c_idx, value=val)
            cell.font = f_norm
            cell.border = box_border
            if c_idx in [1, 2, 4, 5]: cell.alignment = align_center
            else: cell.alignment = align_left

    ws_nl["A11"] = "II. ĐIỀU KIỆN KHÍ HẬU VÙNG NÚI TÂY BẮC & QUY ĐỊNH MÙA KHÔ / MÙA MƯA (TCVN 8819:2011)"
    ws_nl["A11"].font = f_sec_title

    rain_headers = [("A12", "STT"), ("B12", "Khung Thời Gian"), ("C12", "Đặc Điểm Thời Tiết Khí Hậu"), 
                    ("D12", "Hệ Số Năng Suất K_tt"), ("E12", "Hạng Mục Cho Phép Thi Công"), ("F12", "Hạng Mục TUYỆT ĐỐI CẤM Thi Công")]
    for pos, txt in rain_headers:
        ws_nl[pos] = txt
        ws_nl[pos].font = f_header
        ws_nl[pos].fill = fill_sand
        ws_nl[pos].alignment = align_center

    rain_data = [
        (1, "Tháng 10/2026 - Tháng 03/2027", "Mùa khô cao điểm vùng Tây Bắc (Ít mưa, nền nhiệt 15-28°C)", 1.0, 
         "Tổng lực thi công toàn diện: Cọc khoan nhồi hang Karst, bệ móng, thân trụ T1, T2 cao 25.5m, đúc & lao 15 dầm Super-T, thảm BTN C16 mặt cầu", 
         "Nghiêm cấm thảm BTN khi trời mưa hoặc mặt cầu ẩm ướt (TCVN 8819:2011 Điều 8.2)")
    ]
    for idx, item in enumerate(rain_data, start=13):
        for c_idx, val in enumerate(item, start=1):
            cell = ws_nl.cell(row=idx, column=c_idx, value=val)
            cell.font = f_norm
            cell.border = box_border
            if c_idx in [1, 2, 4]: cell.alignment = align_center
            else: cell.alignment = align_left

    ws_nl.column_dimensions["A"].width = 8
    ws_nl.column_dimensions["B"].width = 25
    ws_nl.column_dimensions["C"].width = 45
    ws_nl.column_dimensions["D"].width = 16
    ws_nl.column_dimensions["E"].width = 50
    ws_nl.column_dimensions["F"].width = 45

    # =========================================================================
    # 2. SHEET DB_DINH_MUC
    # =========================================================================
    print("2. Đang xây dựng Sheet DB_DINH_MUC...")
    ws_dm = wb.create_sheet("DB_DINH_MUC")
    ws_dm["A1"] = "CƠ SỞ DỮ LIỆU ĐỊNH MỨC NĂNG SUẤT CA MÁY & HAO PHÍ CẦU KM19+529.080 – 23HG SYSTEM"
    ws_dm["A1"].font = f_title
    ws_dm["A2"] = "Chuẩn hóa theo Thông tư 38/2026/TT-BXD, Thông tư 12/2021/TT-BXD & QCKT Việt Nam"
    ws_dm["A2"].font = f_sub

    headers_dm = [("A4", "STT"), ("B4", "Mã Máy"), ("C4", "Chủng Loại Xe Máy Thiết Bị Thi Công"), 
                  ("D4", "Thông Số Kỹ Thuật / Model"), ("E4", "Định Mức Dầu (lít/ca)"), 
                  ("F4", "Năng Suất 1 Ca Máy"), ("G4", "Nhiệm Vụ Kỹ Thuật Trên Công Trường Cầu Km19")]
    for pos, txt in headers_dm:
        ws_dm[pos] = txt
        ws_dm[pos].font = f_header
        ws_dm[pos].fill = fill_sand
        ws_dm[pos].alignment = align_center

    machinery_db = [
        (1, "MK", "Máy khoan cọc nhồi D1200", "Bauer BG25 / Lực xoay 245 kNm", 145, "8.5 - 16.8 m/ca", "Khoan hang Karst 26 cọc D1200 ngàm đá M1, M2, T1, T2"),
        (2, "CX", "Cần cẩu bánh xích 50 tấn", "Kobelco CKE500 / Cần 38m", 62, "Phục vụ 24/24", "Hạ lồng thép cọc D1200, lắp ván khuôn & đà giáo leo thân trụ T1, T2"),
        (3, "CL", "Cần cẩu bánh lốp 25 tấn", "Kato KR250 / 4 đoạn cần", 48, "Phục vụ bãi đúc", "Phục vụ bãi đúc dầm Super-T, bốc dỡ cốt thép, ván khuôn đúc sẵn"),
        (4, "MD", "Máy đào bánh xích 1.25 m3", "Komatsu PC200-8 / Kèm búa đập đá", 68, "130 - 250 m3/ca", "Đào ngàm đá hố móng bệ trụ, phá đầu 26 cọc nhồi nhô 1.0m"),
        (5, "OT", "Ô tô tự đổ 15 tấn", "Howo 371HP 3 chân", 52, "12 - 16 chuyến/ca", "Vận chuyển đất đá ngàm móng ra bãi thải cự ly 5km"),
        (6, "XB", "Xe bồn vận chuyển bê tông 9 m3", "Hino / Hyundai Trộn xoay", 42, "6 - 8 chuyến/ca", "Vận chuyển bê tông C30 cọc, bệ móng, thân mố trụ và C45 dầm Super-T"),
        (7, "BM", "Xe bơm bê tông cần 42m", "Putzmeister BSF 42-5", 46, "35 - 50 m3/h", "Bơm cao áp bê tông thân trụ T1, T2 (H=23.35m - 25.45m), xà mũ, dầm ngang"),
        (8, "MN", "Máy nén khí 7.5 m3/phút", "Airman PDS265S áp lực cao", 38, "Phục vụ thổi rửa", "Thổi rửa đáy 26 hố khoan cọc nhồi & đục tỉa bê tông đầu cọc"),
        (9, "MH", "Máy phát hàn điện 500A", "Denyo DCW-480ESW / 2 kìm hàn", 28, "Hàn lồng thép cọc", "Gia công nối lồng thép cọc D1200, hàn khung sườn dầm Super-T"),
        (10, "TT", "Trạm trộn BTXM 60 m3/h", "Sicoma Italy 2 trục ngang", 58, "40 - 55 m3/h", "Sản xuất toàn bộ 2,348 m3 BTXM C10-C45 tại hiện trường"),
        (11, "GL", "Giá lao dầm Super-T 78.88T", "Hệ dầm kép ray P43 tời kéo", 35, "1.5 - 2 phiến/ngày", "Vận chuyển & lao lắp an toàn 15 phiến dầm Super-T 38.2m"),
        (12, "LU", "Máy lu rung tự hành 14-25T", "Hamm 3411 lực rung 25 tấn", 45, "400 m3/ca", "Đầm lèn K98 nền đường đầu cầu mố M1, M2"),
        (13, "XT", "Xe téc cấp dầu lưu động 9 m3", "Isuzu téc bồn gắn đồng hồ đo", 44, "Cấp phát 24/24", "Tiếp nhiên liệu diezel tại chỗ cho máy khoan Bauer, cẩu xích, máy đào"),
        (14, "MP", "Máy phát điện 3 pha 250kVA", "Cummins Diesel dự phòng", 55, "Phát điện 24/24", "Cung cấp nguồn điện 3 pha liên tục cho trạm trộn, kích căng kéo và bãi đúc")
    ]
    for idx, item in enumerate(machinery_db, start=5):
        for c_idx, val in enumerate(item, start=1):
            cell = ws_dm.cell(row=idx, column=c_idx, value=val)
            cell.font = f_norm
            cell.border = box_border
            if c_idx in [1, 2, 5]: cell.alignment = align_center
            elif c_idx == 6: cell.alignment = align_right
            else: cell.alignment = align_left

    ws_dm.column_dimensions["A"].width = 8
    ws_dm.column_dimensions["B"].width = 12
    ws_dm.column_dimensions["C"].width = 32
    ws_dm.column_dimensions["D"].width = 32
    ws_dm.column_dimensions["E"].width = 22
    ws_dm.column_dimensions["F"].width = 22
    ws_dm.column_dimensions["G"].width = 55

    # =========================================================================
    # 3. SHEET BOQ_TIEN_DO: 100% CÔNG THỨC SỐNG
    # =========================================================================
    print("3. Đang xây dựng Sheet BOQ_TIEN_DO (100% công thức sống)...")
    ws_boq = wb.create_sheet("BOQ_TIEN_DO")
    ws_boq["A1"] = "BẢNG TIÊN LƯỢNG KHỐI LƯỢNG (BOQ), ĐỊNH MỨC CA MÁY & CHI PHÍ DỰ TOÁN G_XD CẦU KM19+529.080"
    ws_boq["A1"].font = f_title
    ws_boq["A2"] = "Dự án: Cao tốc Tuyên Quang - Hà Giang (Gói 09-XL) | Tiêu chuẩn: Thông tư 38/2026/TT-BXD, TT 36/2026 & TCVN"
    ws_boq["A2"].font = f_sub

    headers_boq = [("A4", "STT"), ("B4", "Mã WBS"), ("C4", "Nội Dung Hạng Mục Công Tác Xây Lắp"), 
                   ("D4", "Đơn Vị"), ("E4", "Khối Lượng Thiết Kế"), ("F4", "Đơn Giá Trực Tiếp (VNĐ)"), 
                   ("G4", "Thành Tiền Trực Tiếp T (VNĐ)"), ("H4", "Năng Suất Định Mức / Ca"), 
                   ("I4", "Số Ca Máy Dự Kiến"), ("J4", "Thiết Bị Chủ Lực")]
    for pos, txt in headers_boq:
        ws_boq[pos] = txt
        ws_boq[pos].font = f_header
        ws_boq[pos].fill = fill_sand
        ws_boq[pos].alignment = align_center

    boq_items = [
        # WBS 1: Chuẩn bị & phụ trợ
        ("1.1", "Bàn giao tim mốc VN2000 & định vị tọa độ mố trụ cầu", "Điểm", 1.0, 15000000, 0.5, "Máy toàn đạc Leica TS06"),
        ("1.2", "Rà phá bom mìn, dọn dẹp mặt bằng & dựng lán trại", "m2", 12500.0, 18500, 2500, "Máy xúc PC200 + Ô tô 15T"),
        ("1.3", "Thi công đường công vụ nhánh 6 & 6A tiếp cận mố trụ", "m3", 1536.5, 95000, 170, "Máy xúc PC200 + Lu rung Hamm"),
        ("1.4", "Xây dựng bãi đúc dầm Super-T, bệ đúc & trạm trộn 60m3/h", "Hệ", 1.0, 450000000, 0.08, "Cần cẩu lốp 25T + Máy xúc"),

        # WBS 2: Cọc khoan nhồi D1200
        ("2.1", "Khoan thăm dò hang Karst sâu 5m vào đá liền khối (26 lỗ)", "Lỗ", 26.0, 8500000, 2.6, "Máy khoan đá Bauer BG25"),
        ("2.2", "Khoan cọc nhồi D1200 Mố M1 (3 cọc L=20m) & đổ BT C30", "m", 60.0, 3850000, 10.0, "Bauer BG25 + Cẩu 50T + Xe bồn"),
        ("2.3", "Khoan cọc nhồi D1200 Trụ T1 (8 cọc L=40m) & đổ BT C30", "m", 320.0, 4200000, 21.3, "2 Máy khoan Bauer + Cẩu 50T"),
        ("2.4", "Khoan cọc nhồi D1200 Trụ T2 (8 cọc L=30m) & đổ BT C30", "m", 240.0, 4150000, 18.5, "2 Máy khoan Bauer + Cẩu 50T"),
        ("2.5", "Khoan cọc nhồi D1200 Mố M2 (7 cọc L=36m) & đổ BT C30", "m", 252.0, 4100000, 18.0, "2 Máy khoan Bauer + Cẩu 50T"),
        ("2.6", "Thí nghiệm siêu âm 156 mặt cắt cọc, PDA & nén tĩnh", "Gói", 1.0, 185000000, 0.17, "Hệ thống siêu âm cọc đa kênh"),

        # WBS 3: Kết cấu phần dưới (ĐÃ TÁCH TRỤ T1, T2 & BỆ KHÔNG VÁT)
        ("3.1", "Đào hố móng ngàm đá, đập đầu 26 cọc & đổ BT lót C10 (42.15 m3)", "Cọc", 26.0, 4500000, 4.3, "Máy đào PC200 búa đập đá"),
        ("3.2", "Cốt thép, ván khuôn & đổ bê tông bệ móng mố M1, M2 (173.2 m3)", "m3", 173.2, 2850000, 24.7, "Cẩu 25T + Bơm 42m + Xe bồn"),
        ("3.3.1", "Cốt thép, VK & đổ BT bệ trụ T1 (11.6x8.0x2.0m Không vát, 184.243 m3)", "m3", 184.243, 2920000, 26.3, "Cẩu xích 50T + Bơm 42m"),
        ("3.3.2", "Cốt thép, VK & đổ BT bệ trụ T2 (11.6x8.0x2.0m Không vát, 184.243 m3)", "m3", 184.243, 2920000, 26.3, "Cẩu xích 50T + Bơm 42m"),
        ("3.4.1", "Thi công thân đặc trụ T1 (H=23.35m, 3 đốt vát, 234.922 m3) BT C30", "m3", 234.922, 3450000, 21.4, "Cẩu xích 50T + Đà giáo leo"),
        ("3.4.2", "Thi công thân đặc trụ T2 (H=25.45m, 3 đốt vát, 266.791 m3) BT C30", "m3", 266.791, 3450000, 22.2, "Cẩu xích 50T + Đà giáo leo"),
        ("3.5", "Thi công thân mố chân dê M1 & mố chữ U M2 BT C30 (138.4 m3)", "m3", 138.4, 3200000, 13.8, "Cẩu lốp 25T + Bơm 42m"),
        ("3.6.1", "Thi công xà mũ trụ T1, đá kê gối & khối chống xô (128.851 m3)", "m3", 128.851, 3650000, 21.5, "Cẩu 50T + Ván khuôn dầm đỡ"),
        ("3.6.2", "Thi công xà mũ trụ T2, đá kê gối & khối chống xô (145.026 m3)", "m3", 145.026, 3650000, 20.7, "Cẩu 50T + Ván khuôn dầm đỡ"),

        # WBS 4: Chế tạo & Lao lắp dầm Super-T
        ("4.1", "Cốt thép, luồn cáp DUL 15 phiến dầm Super-T tại bãi đúc (15 phiến)", "Phiến", 15.0, 85000000, 0.63, "Dàn cắt uốn CNC + Cẩu 25T"),
        ("4.2", "Đổ bê tông C45 15 phiến dầm Super-T L=38.2m bãi đúc (434.8 m3)", "m3", 434.8, 3850000, 24.2, "Trạm trộn 60m3/h + Cẩu 25T"),
        ("4.3", "Căng kéo cáp DUL 15.2mm & bơm vữa ống gen dầm (15 phiến)", "Phiến", 15.0, 42000000, 1.88, "Bơm kích thủy lực 250T"),
        ("4.4", "Lắp đặt 30 gối chậu cao su đơn/song hướng mố trụ (30 cái)", "Cái", 30.0, 12500000, 7.5, "Cần cẩu lốp 25T + Kích căn"),
        ("4.5", "Lắp đường ray P43, vận chuyển & lao lắp 15 dầm Super-T", "Phiến", 15.0, 58000000, 2.14, "Giá lao dầm ray P43 78.88T"),

        # WBS 5: Bản mặt cầu & Hoàn thiện
        ("5.1", "Đổ bê tông dầm ngang mố trụ & bản liên tục nhiệt C35 (78.5 m3)", "m3", 78.5, 3350000, 19.6, "Xe bơm cần 42m + Xe bồn"),
        ("5.2", "Lắp 510 tấm ván khuôn đúc sẵn & cốt thép bản mặt cầu (510 tấm)", "Tấm", 510.0, 450000, 127.5, "Cần cẩu lốp 25T bốc dỡ"),
        ("5.3", "Đổ bê tông bản mặt cầu tại chỗ C35 (dày 180mm, 235.6 m3)", "m3", 235.6, 3250000, 58.9, "Xe bơm cần 42m + 3 Xe bồn"),
        ("5.4", "Đổ bê tông bản quá độ C25 & đắp K98 đường đầu cầu (2 bản)", "Bản", 2.0, 65000000, 0.4, "Máy xúc PC200 + Lu rung"),
        ("5.5", "Lắp khe co giãn răng lược thép & ống thoát nước PVC (2 bộ)", "Bộ", 2.0, 95000000, 0.5, "Máy phát hàn + Cẩu nhẹ"),
        ("5.6", "Đổ bê tông gờ lan can C25 & lắp dựng tay vịn thép (260.8 m)", "m", 260.8, 1150000, 52.2, "Xe bồn 9m3 + Máy phát hàn"),
        ("5.7", "Chống thấm & thảm bê tông nhựa C16 mặt cầu dày 7cm (1,620 m2)", "m2", 1620.0, 310000, 540.0, "Xe tưới nhựa + Máy rải Vogele"),
        ("5.8", "Thử tải kiểm định động học & nghiệm thu thông xe hoàn thành", "Gói", 1.0, 120000000, 1.0, "Đoàn xe thử tải 3 trục 30T")
    ]

    for idx, item in enumerate(boq_items, start=5):
        wbs_val, name_val, unit_val, qty_val, price_val, cap_rate, eq_val = item
        ws_boq.cell(row=idx, column=1, value=idx-4).alignment = align_center
        ws_boq.cell(row=idx, column=2, value=wbs_val).alignment = align_center
        ws_boq.cell(row=idx, column=3, value=name_val).alignment = align_left
        ws_boq.cell(row=idx, column=4, value=unit_val).alignment = align_center
        
        c_qty = ws_boq.cell(row=idx, column=5, value=qty_val)
        c_qty.number_format = "#,##0.00" if isinstance(qty_val, float) else "#,##0"
        c_qty.alignment = align_right

        c_pri = ws_boq.cell(row=idx, column=6, value=price_val)
        c_pri.number_format = "#,##0"
        c_pri.alignment = align_right

        # Công thức sống Thành tiền: =E{idx}*F{idx}
        c_amt = ws_boq.cell(row=idx, column=7, value=f"=E{idx}*F{idx}")
        c_amt.number_format = "#,##0"
        c_amt.font = f_norm_bold
        c_amt.alignment = align_right

        c_cap = ws_boq.cell(row=idx, column=8, value=cap_rate)
        c_cap.number_format = "#,##0.00"
        c_cap.alignment = align_right

        # Công thức sống Số ca: =ROUNDUP(E{idx}/H{idx}, 1)
        c_sh = ws_boq.cell(row=idx, column=9, value=f"=ROUNDUP(E{idx}/H{idx}, 1)")
        c_sh.number_format = "#,##0.0"
        c_sh.alignment = align_right

        ws_boq.cell(row=idx, column=10, value=eq_val).alignment = align_left

        for c_i in range(1, 11):
            cell = ws_boq.cell(row=idx, column=c_i)
            cell.font = f_norm if c_i != 7 else f_norm_bold
            cell.border = box_border

    # Dòng Tổng cộng Chi phí trực tiếp T
    tot_row = len(boq_items) + 5
    ws_boq.cell(row=tot_row, column=3, value="TỔNG CHI PHÍ TRỰC TIẾP (T) CẦU KM19+529.080").font = f_title_brown
    ws_boq.cell(row=tot_row, column=7, value=f"=SUM(G5:G{tot_row-1})").number_format = "#,##0"
    ws_boq.cell(row=tot_row, column=7).font = Font(name="Segoe UI", size=10, bold=True, color="5A3718")
    ws_boq.cell(row=tot_row, column=7).fill = fill_sand

    # Bảng phân tích chi phí xây dựng G_XD theo Thông tư 36/2026/TT-BXD
    r_gt = tot_row + 1
    ws_boq.cell(row=r_gt, column=3, value="Chi phí gián tiếp (GT = 7.3% * T)").font = f_norm
    ws_boq.cell(row=r_gt, column=7, value=f"=G{tot_row}*0.073").number_format = "#,##0"

    r_tl = tot_row + 2
    ws_boq.cell(row=r_tl, column=3, value="Thu nhập chịu thuế tính trước (TL = 5.5% * (T + GT))").font = f_norm
    ws_boq.cell(row=r_tl, column=7, value=f"=(G{tot_row}+G{r_gt})*0.055").number_format = "#,##0"

    r_vat = tot_row + 3
    ws_boq.cell(row=r_vat, column=3, value="Thuế giá trị gia tăng (VAT = 10% * (T + GT + TL))").font = f_norm
    ws_boq.cell(row=r_vat, column=7, value=f"=(G{tot_row}+G{r_gt}+G{r_tl})*0.1").number_format = "#,##0"

    r_gxd = tot_row + 4
    ws_boq.cell(row=r_gxd, column=3, value="TỔNG DỰ TOÁN XÂY LẮP G_XD CẦU KM19+529.080 (SAU THUẾ)").font = f_title
    ws_boq.cell(row=r_gxd, column=7, value=f"=G{tot_row}+G{r_gt}+G{r_tl}+G{r_vat}").number_format = "#,##0"
    ws_boq.cell(row=r_gxd, column=7).font = Font(name="Segoe UI", size=11, bold=True, color="DC2626")
    ws_boq.cell(row=r_gxd, column=7).fill = fill_cream

    for r_k in range(tot_row, r_gxd + 1):
        for c_k in range(1, 11):
            ws_boq.cell(row=r_k, column=c_k).border = box_border

    ws_boq.column_dimensions["A"].width = 6
    ws_boq.column_dimensions["B"].width = 10
    ws_boq.column_dimensions["C"].width = 52
    ws_boq.column_dimensions["D"].width = 8
    ws_boq.column_dimensions["E"].width = 18
    ws_boq.column_dimensions["F"].width = 18
    ws_boq.column_dimensions["G"].width = 24
    ws_boq.column_dimensions["H"].width = 20
    ws_boq.column_dimensions["I"].width = 16
    ws_boq.column_dimensions["J"].width = 30

    # =========================================================================
    # 4. SHEET TIEN_DO: CPM TỰ ĐỘNG & GANTT CANVAS
    # =========================================================================
    print("4. Đang xây dựng Sheet TIEN_DO (CPM & Gantt Canvas)...")
    ws_td = wb.create_sheet("TIEN_DO")
    ws_td["A1"] = "• HỆ THỐNG QUẢN LÝ TIẾN ĐỘ THI CÔNG & MẠNG CPM TỰ ĐỘNG – 23HG SYSTEM"
    ws_td["A1"].font = f_banner
    ws_td["A1"].fill = fill_coffee

    ws_td["A2"] = "DỰ ÁN: CAO TỐC TUYÊN QUANG - HÀ GIANG (GÓI 09-XL) — CẦU KM19+529.080 (3 NHỊP DẦM SUPER-T L=38.2M)"
    ws_td["A2"].font = f_title
    ws_td["E2"] = "BẮT ĐẦU:"
    ws_td["E2"].font = f_norm_bold
    ws_td["F2"] = dt.datetime(2026, 10, 1)
    ws_td["F2"].number_format = "dd/mm/yyyy"
    ws_td["F2"].font = f_norm_bold
    ws_td["G2"] = "KẾT THÚC:"
    ws_td["G2"].font = f_norm_bold
    ws_td["H2"] = dt.datetime(2027, 3, 11)
    ws_td["H2"].number_format = "dd/mm/yyyy"
    ws_td["H2"].font = Font(name="Segoe UI", size=9, bold=True, color="DC2626")

    ws_td["A3"] = "Ngày kiểm soát (Data Date):"
    ws_td["A3"].font = f_sub
    ws_td["C3"] = dt.datetime(2026, 12, 15)
    ws_td["C3"].number_format = "dd/mm/yyyy"
    ws_td["C3"].font = f_norm_bold
    ws_td["E3"] = "Lịch thi công:"
    ws_td["E3"].font = f_sub
    ws_td["F3"] = "6 ngày/tuần (Nghỉ CN & Lễ)"
    ws_td["G3"] = "Chế độ CPM:"
    ws_td["G3"].font = f_sub
    ws_td["H3"] = "Tự động (Forward/Backward Pass)"

    # Timeline Headers (Columns 14 to 37)
    # 24 timeline columns representing 24 weeks from Oct 2026 to Mar 2027
    headers_td = [
        ("A5", "STT"), ("B5", "Mã WBS"), ("C5", "Nội Dung Hạng Mục Công Việc"), ("D5", "Loại"),
        ("E5", "Công Tác Tiền Nhiệm (Predecessors)"), ("F5", "Thời Lượng (Work Days)"),
        ("G5", "Hệ Số Mưa K_tt"), ("H5", "Khởi Công (ES)"), ("I5", "Hoàn Thành (EF)"),
        ("J5", "Ngày Lịch (Cal Days)"), ("K5", "Dự Trữ TF"), ("L5", "Đường Găng"),
        ("M5", "% Hoàn Thành (Tại Data Date)")
    ]
    for pos, txt in headers_td:
        ws_td[pos] = txt
        ws_td[pos].font = f_header
        ws_td[pos].fill = fill_sand
        ws_td[pos].alignment = align_center

    # Timeline Months row 4 & Weeks row 5
    weeks = []
    w_start = dt.date(2026, 10, 1)
    cur = w_start
    col_idx = 14
    while cur <= dt.date(2027, 3, 20):
        w_end = cur + dt.timedelta(days=6)
        lbl = f"T{cur.month:02d}/W{(cur.day-1)//7 + 1}"
        ws_td.cell(row=5, column=col_idx, value=lbl).font = Font(name="Segoe UI", size=7.5, bold=True, color="555555")
        ws_td.cell(row=5, column=col_idx).fill = fill_cream
        ws_td.cell(row=5, column=col_idx).alignment = align_center
        weeks.append((col_idx, cur, w_end))
        col_idx += 1
        cur += dt.timedelta(days=7)

    # Fill Tasks
    row_idx = 6
    for t in tasks_def:
        r_cpm = cpm_results[t.wbs]
        ws_td.cell(row=row_idx, column=1, value=row_idx-5).alignment = align_center
        ws_td.cell(row=row_idx, column=2, value=t.wbs).alignment = align_left
        
        # Indent WBS in column C
        depth = len(t.wbs.split('.')) - 1
        indent_spaces = "   " * depth
        c_name = ws_td.cell(row=row_idx, column=3, value=indent_spaces + t.name)
        c_name.alignment = align_left

        c_kind = ws_td.cell(row=row_idx, column=4, value=t.kind)
        c_kind.alignment = align_center

        # Predecessor string
        pred_str = ", ".join([f"{p[0]}{p[1]}{f'+{p[2]}' if p[2]>0 else ''}" for p in t.preds]) if t.preds else ""
        ws_td.cell(row=row_idx, column=5, value=pred_str).alignment = align_left

        c_dur = ws_td.cell(row=row_idx, column=6, value=t.duration)
        c_dur.alignment = align_right
        ws_td.cell(row=row_idx, column=7, value=1.0).alignment = align_center

        c_es = ws_td.cell(row=row_idx, column=8, value=dt.datetime(r_cpm.es.year, r_cpm.es.month, r_cpm.es.day))
        c_es.number_format = "dd/mm/yyyy"
        c_es.alignment = align_center

        c_ef = ws_td.cell(row=row_idx, column=9, value=dt.datetime(r_cpm.ef.year, r_cpm.ef.month, r_cpm.ef.day))
        c_ef.number_format = "dd/mm/yyyy"
        c_ef.alignment = align_center

        cal_days = (r_cpm.ef - r_cpm.es).days + 1 if t.kind != 'Milestone' else 0
        ws_td.cell(row=row_idx, column=10, value=cal_days).alignment = align_right

        c_tf = ws_td.cell(row=row_idx, column=11, value=r_cpm.tf)
        c_tf.alignment = align_right

        crit_str = "GANG" if r_cpm.critical else ""
        c_gang = ws_td.cell(row=row_idx, column=12, value=crit_str)
        c_gang.alignment = align_center
        if r_cpm.critical:
            c_gang.font = f_crit

        # % Completion at Data Date (15/12/2026)
        data_date = dt.date(2026, 12, 15)
        if r_cpm.ef <= data_date:
            pct = 1.0
        elif r_cpm.es > data_date:
            pct = 0.0
        else:
            tot = (r_cpm.ef - r_cpm.es).days + 1
            el = (data_date - r_cpm.es).days + 1
            pct = round(el / tot, 2)
        c_pct = ws_td.cell(row=row_idx, column=13, value=pct)
        c_pct.number_format = "0.0%"
        c_pct.alignment = align_right

        # Font & Fill per Type
        if t.kind == 'Summary':
            for c_i in range(1, 14):
                cell = ws_td.cell(row=row_idx, column=c_i)
                cell.font = f_summary
                cell.fill = fill_summary
                cell.border = box_border
        elif r_cpm.critical:
            for c_i in range(1, 14):
                cell = ws_td.cell(row=row_idx, column=c_i)
                cell.font = f_norm_bold if c_i in [2, 3, 12] else f_norm
                cell.border = box_border
        else:
            for c_i in range(1, 14):
                cell = ws_td.cell(row=row_idx, column=c_i)
                cell.font = f_norm
                cell.border = box_border

        # Gantt Canvas Bar
        for w_col, w_s, w_e in weeks:
            cell_gantt = ws_td.cell(row=row_idx, column=w_col)
            cell_gantt.border = box_border
            if not (r_cpm.ef < w_s or r_cpm.es > w_e):
                if t.kind == 'Summary':
                    cell_gantt.fill = fill_bar_sum
                elif r_cpm.critical:
                    cell_gantt.fill = fill_bar_crit
                else:
                    cell_gantt.fill = fill_bar_norm

        row_idx += 1

    # Freeze panes at D6
    ws_td.freeze_panes = "D6"

    # Set column widths
    ws_td.column_dimensions["A"].width = 6
    ws_td.column_dimensions["B"].width = 10
    ws_td.column_dimensions["C"].width = 52
    ws_td.column_dimensions["D"].width = 10
    ws_td.column_dimensions["E"].width = 24
    ws_td.column_dimensions["F"].width = 14
    ws_td.column_dimensions["G"].width = 12
    ws_td.column_dimensions["H"].width = 14
    ws_td.column_dimensions["I"].width = 14
    ws_td.column_dimensions["J"].width = 12
    ws_td.column_dimensions["K"].width = 12
    ws_td.column_dimensions["L"].width = 12
    ws_td.column_dimensions["M"].width = 16
    for w_col, _, _ in weeks:
        ws_td.column_dimensions[L(w_col)].width = 5.5

    # =========================================================================
    # 5. SHEET EVM_5D_QUAN_TRI
    # =========================================================================
    print("5. Đang xây dựng Sheet EVM_5D_QUAN_TRI...")
    ws_evm = wb.create_sheet("EVM_5D_QUAN_TRI")
    ws_evm["A1"] = "HỆ THỐNG QUẢN TRỊ GIÁ TRỊ THU ĐƯỢC (EVM 5D) & ĐƯỜNG CƠ SỞ BASELINE CẦU KM19+529.080"
    ws_evm["A1"].font = f_title
    ws_evm["A2"] = "Kiểm soát Tiến độ - Chi phí (PV, EV, AC, SV, CV, SPI, CPI) theo chuẩn PMI & Nghị định 206/2026/NĐ-CP"
    ws_evm["A2"].font = f_sub

    evm_kpis = [
        ("TỔNG DỰ TOÁN (BAC)", "=BOQ_TIEN_DO!G41", "#,##0", "Tổng dự toán xây lắp G_XD phê duyệt sau thuế"),
        ("KẾ HOẠCH TÍCH LŨY (PV)", "=C5*0.485", "#,##0", "Giá trị công việc theo kế hoạch tại Data Date 15/12/2026"),
        ("GIÁ TRỊ ĐẠT ĐƯỢC (EV)", "=C5*0.512", "#,##0", "Khối lượng thực tế hoàn thành đã nghiệm thu KCS"),
        ("CHI PHÍ THỰC TẾ (AC)", "=C5*0.498", "#,##0", "Tổng chi phí vật tư, nhân công, ca máy đã giải ngân"),
        ("CHÊNH LỆCH TIẾN ĐỘ (SV = EV - PV)", "=C7-C6", "#,##0", "Giá trị dương: Tiến độ thi công đang vượt trước kế hoạch"),
        ("CHÊNH LỆCH CHI PHÍ (CV = EV - AC)", "=C7-C8", "#,##0", "Giá trị dương: Tiết kiệm chi phí, không vượt ngân sách"),
        ("CHỈ SỐ HIỆU SUẤT TIẾN ĐỘ (SPI = EV/PV)", "=C7/C6", "0.000", "SPI >= 1.0: Đạt và vượt tiến độ yêu cầu"),
        ("CHỈ SỐ HIỆU SUẤT CHI PHÍ (CPI = EV/AC)", "=C7/C8", "0.000", "CPI >= 1.0: Quản trị chi phí đạt hiệu quả cao"),
        ("DỰ BÁO HOÀN THÀNH (EAC = BAC/CPI)", "=C5/C12", "#,##0", "Dự báo tổng chi phí khi hoàn thành công trình"),
        ("CHI PHÍ CÒN LẠI (ETC = EAC - AC)", "=C13-C8", "#,##0", "Dự toán chi phí cần thiết để thi công nốt phần còn lại"),
        ("CHÊNH LỆCH KHI XONG (VAC = BAC - EAC)", "=C5-C13", "#,##0", "Giá trị dương: Dự kiến hoàn thành dưới ngân sách")
    ]
    ws_evm["A4"] = "CHỈ SỐ QUẢN TRỊ"
    ws_evm["C4"] = "GIÁ TRỊ TÍNH TOÁN"
    ws_evm["E4"] = "GIẢI NGHĨA KỸ THUẬT & QUẢN TRỊ 5D"
    for p in ["A4", "C4", "E4"]:
        ws_evm[p].font = f_header
        ws_evm[p].fill = fill_sand

    for idx, (lbl, fmla, fmt, expl) in enumerate(evm_kpis, start=5):
        ws_evm.cell(row=idx, column=1, value=lbl).font = f_norm_bold
        c_val = ws_evm.cell(row=idx, column=3, value=fmla)
        c_val.font = Font(name="Segoe UI", size=10, bold=True, color="5A3718")
        c_val.number_format = fmt
        c_val.alignment = align_right
        ws_evm.cell(row=idx, column=5, value=expl).font = f_sub
        for c_k in range(1, 7):
            ws_evm.cell(row=idx, column=c_k).border = box_border

    ws_evm.column_dimensions["A"].width = 38
    ws_evm.column_dimensions["B"].width = 4
    ws_evm.column_dimensions["C"].width = 25
    ws_evm.column_dimensions["D"].width = 4
    ws_evm.column_dimensions["E"].width = 60

    # =========================================================================
    # 6. SHEET HUY_DONG_XMTB
    # =========================================================================
    print("6. Đang xây dựng Sheet HUY_DONG_XMTB...")
    ws_xmtb = wb.create_sheet("HUY_DONG_XMTB")
    ws_xmtb["A1"] = "KẾ HOẠCH HUY ĐỘNG & CÂN BẰNG TÀI NGUYÊN XE MÁY THIẾT BỊ CẦU KM19+529.080"
    ws_xmtb["A1"].font = f_title
    ws_xmtb["A2"] = "Kiểm soát đỉnh tải thiết bị - Khớp năng suất định mức ca máy theo Thông tư 38/2026/TT-BXD"
    ws_xmtb["A2"].font = f_sub

    xmtb_headers = [("A4", "STT"), ("B4", "Mã"), ("C4", "Chủng Loại Xe Máy Thiết Bị"), 
                    ("D4", "Định Mức Dầu"), ("E4", "T10/26"), ("F4", "T11/26"), 
                    ("G4", "T12/26"), ("H4", "T01/27"), ("I4", "T02/27"), ("J4", "T03/27"), 
                    ("K4", "Tổng Ca Máy"), ("L4", "Tổng Nhiên Liệu (Lít)")]
    for pos, txt in xmtb_headers:
        ws_xmtb[pos] = txt
        ws_xmtb[pos].font = f_header
        ws_xmtb[pos].fill = fill_sand
        ws_xmtb[pos].alignment = align_center

    # Machinery monthly distribution
    mach_dist = [
        (1, "MK", "Máy khoan cọc nhồi D1200 Bauer BG25", 145, 12, 60, 48, 0, 0, 0, 120),
        (2, "CX", "Cần cẩu bánh xích 50T Kobelco CKE500", 62, 10, 45, 45, 45, 25, 10, 180),
        (3, "CL", "Cần cẩu bánh lốp 25T Kato KR250", 48, 15, 25, 25, 25, 20, 10, 120),
        (4, "MD", "Máy đào bánh xích 1.25m3 Komatsu PC200", 68, 20, 15, 30, 15, 10, 5, 95),
        (5, "OT", "Ô tô tự đổ 15 tấn Howo 371HP", 52, 30, 25, 40, 25, 15, 5, 140),
        (6, "XB", "Xe bồn vận chuyển bê tông 9m3", 42, 10, 50, 55, 45, 35, 15, 210),
        (7, "BM", "Xe bơm bê tông cần 42m Putzmeister", 46, 5, 20, 25, 20, 10, 5, 85),
        (8, "MN", "Máy nén khí 7.5m3/phút Airman PDS265", 38, 10, 35, 35, 15, 10, 5, 110),
        (9, "MH", "Máy phát hàn điện 500A Denyo 500A", 28, 20, 45, 45, 40, 25, 15, 190),
        (10, "TT", "Trạm trộn BTXM 60m3/h Sicoma", 58, 15, 35, 40, 35, 25, 10, 160),
        (11, "GL", "Giá lao dầm Super-T 78.88T ray P43", 35, 0, 0, 0, 15, 40, 5, 60),
        (12, "LU", "Máy lu rung 14-25T Hamm 3411", 45, 10, 0, 0, 0, 5, 15, 30),
        (13, "XT", "Xe téc cấp dầu lưu động 9m3", 44, 20, 25, 25, 25, 20, 15, 130),
        (14, "MP", "Máy phát điện 3 pha công nghiệp 250kVA", 55, 25, 30, 30, 30, 25, 15, 155)
    ]
    for idx, item in enumerate(mach_dist, start=5):
        stt, code, name, dm_oil, m10, m11, m12, m01, m02, m03, tot_ca = item
        ws_xmtb.cell(row=idx, column=1, value=stt).alignment = align_center
        ws_xmtb.cell(row=idx, column=2, value=code).alignment = align_center
        ws_xmtb.cell(row=idx, column=3, value=name).alignment = align_left
        ws_xmtb.cell(row=idx, column=4, value=dm_oil).alignment = align_right

        ws_xmtb.cell(row=idx, column=5, value=m10).alignment = align_right
        ws_xmtb.cell(row=idx, column=6, value=m11).alignment = align_right
        ws_xmtb.cell(row=idx, column=7, value=m12).alignment = align_right
        ws_xmtb.cell(row=idx, column=8, value=m01).alignment = align_right
        ws_xmtb.cell(row=idx, column=9, value=m02).alignment = align_right
        ws_xmtb.cell(row=idx, column=10, value=m03).alignment = align_right

        # Công thức sống Tổng ca: =SUM(E{idx}:J{idx})
        c_tot = ws_xmtb.cell(row=idx, column=11, value=f"=SUM(E{idx}:J{idx})")
        c_tot.number_format = "#,##0"
        c_tot.font = f_norm_bold
        c_tot.alignment = align_right

        # Công thức sống Nhiên liệu: =K{idx}*D{idx}
        c_oil = ws_xmtb.cell(row=idx, column=12, value=f"=K{idx}*D{idx}")
        c_oil.number_format = "#,##0"
        c_oil.font = f_norm_bold
        c_oil.alignment = align_right

        for c_k in range(1, 13):
            cell = ws_xmtb.cell(row=idx, column=c_k)
            cell.font = f_norm if c_k not in [11, 12] else f_norm_bold
            cell.border = box_border

    tot_xmtb = len(mach_dist) + 5
    ws_xmtb.cell(row=tot_xmtb, column=3, value="TỔNG CỘNG TOÀN CÔNG TRƯỜNG CẦU KM19").font = f_title_brown
    ws_xmtb.cell(row=tot_xmtb, column=11, value=f"=SUM(K5:K{tot_xmtb-1})").number_format = "#,##0"
    ws_xmtb.cell(row=tot_xmtb, column=12, value=f"=SUM(L5:L{tot_xmtb-1})").number_format = "#,##0"
    ws_xmtb.cell(row=tot_xmtb, column=12).font = Font(name="Segoe UI", size=10, bold=True, color="DC2626")
    ws_xmtb.cell(row=tot_xmtb, column=12).fill = fill_sand

    for c_k in range(1, 13):
        ws_xmtb.cell(row=tot_xmtb, column=c_k).border = box_border

    ws_xmtb.column_dimensions["A"].width = 6
    ws_xmtb.column_dimensions["B"].width = 8
    ws_xmtb.column_dimensions["C"].width = 38
    ws_xmtb.column_dimensions["D"].width = 16
    for c_l in ["E", "F", "G", "H", "I", "J"]:
        ws_xmtb.column_dimensions[c_l].width = 11
    ws_xmtb.column_dimensions["K"].width = 16
    ws_xmtb.column_dimensions["L"].width = 22

    # =========================================================================
    # 7. SHEET TIEN_DO_GIAI_NGAN
    # =========================================================================
    print("7. Đang xây dựng Sheet TIEN_DO_GIAI_NGAN...")
    ws_gn = wb.create_sheet("TIEN_DO_GIAI_NGAN")
    ws_gn["A1"] = "KẾ HOẠCH DÒNG TIỀN & TIẾN ĐỘ GIẢI NGÂN (S-CURVE TÀI CHÍNH) CẦU KM19+529.080"
    ws_gn["A1"].font = f_title
    ws_gn["A2"] = "Tuân thủ Nghị định 206/2026/NĐ-CP: Tạm ứng 20%, thu hồi tạm ứng theo lũy tiến, bảo hành 5%"
    ws_gn["A2"].font = f_sub

    gn_headers = [("A4", "Kỳ TT"), ("B4", "Thời Gian"), ("C4", "Nội Dung Nghiệm Thu Trọng Tâm Cầu Km19"), 
                  ("D4", "Giá Trị Hoàn Thành (VNĐ)"), ("E4", "Khấu Trừ Tạm Ứng (VNĐ)"), 
                  ("F4", "Giữ Lại Bảo Hành 5% (VNĐ)"), ("G4", "Thực Nhận Kỳ Này (VNĐ)")]
    for pos, txt in gn_headers:
        ws_gn[pos] = txt
        ws_gn[pos].font = f_header
        ws_gn[pos].fill = fill_sand
        ws_gn[pos].alignment = align_center

    gn_data = [
        (0, "T10/2026", "TẠM ỨNG HỢP ĐỒNG 20% (Theo Nghị định 206/2026/NĐ-CP)", 0, 0, 0, "=BOQ_TIEN_DO!G41*0.2"),
        (1, "T11/2026", "Nghiệm thu chuẩn bị, đường công vụ & cọc M1, T1", 3850000000, 0, "=D6*0.05", "=D6-E6-F6"),
        (2, "T12/2026", "Nghiệm thu cọc T2, M2 & bệ mố trụ M1, M2, T1, T2", 5200000000, "=G5*0.3", "=D7*0.05", "=D7-E7-F7"),
        (3, "T01/2027", "Nghiệm thu thân trụ T1, T2, xà mũ & đúc 15 dầm Super-T", 5800000000, "=G5*0.4", "=D8*0.05", "=D8-E8-F8"),
        (4, "T02/2027", "Nghiệm thu lao dầm, dầm ngang, bản mặt cầu C35", 4100000000, "=G5*0.3", "=D9*0.05", "=D9-E9-F9"),
        (5, "T03/2027", "Nghiệm thu bản quá độ, gờ lan can, thảm BTN & bàn giao", 2250000000, 0, "=D10*0.05", "=D10-E10-F10")
    ]
    for idx, item in enumerate(gn_data, start=5):
        for c_idx, val in enumerate(item, start=1):
            cell = ws_gn.cell(row=idx, column=c_idx, value=val)
            cell.font = f_norm
            cell.border = box_border
            if c_idx in [1, 2]: cell.alignment = align_center
            elif c_idx == 3: cell.alignment = align_left
            else:
                cell.alignment = align_right
                cell.number_format = "#,##0"
                if c_idx == 7: cell.font = f_norm_bold

    tot_gn = len(gn_data) + 5
    ws_gn.cell(row=tot_gn, column=3, value="TỔNG CỘNG GIÁ TRỊ GIẢI NGÂN").font = f_title_brown
    ws_gn.cell(row=tot_gn, column=4, value=f"=SUM(D5:D{tot_gn-1})").number_format = "#,##0"
    ws_gn.cell(row=tot_gn, column=5, value=f"=SUM(E5:E{tot_gn-1})").number_format = "#,##0"
    ws_gn.cell(row=tot_gn, column=6, value=f"=SUM(F5:F{tot_gn-1})").number_format = "#,##0"
    ws_gn.cell(row=tot_gn, column=7, value=f"=SUM(G5:G{tot_gn-1})").number_format = "#,##0"
    ws_gn.cell(row=tot_gn, column=7).font = Font(name="Segoe UI", size=10, bold=True, color="5A3718")
    ws_gn.cell(row=tot_gn, column=7).fill = fill_sand

    for c_k in range(1, 8):
        ws_gn.cell(row=tot_gn, column=c_k).border = box_border

    ws_gn.column_dimensions["A"].width = 8
    ws_gn.column_dimensions["B"].width = 14
    ws_gn.column_dimensions["C"].width = 52
    ws_gn.column_dimensions["D"].width = 24
    ws_gn.column_dimensions["E"].width = 24
    ws_gn.column_dimensions["F"].width = 24
    ws_gn.column_dimensions["G"].width = 24

    # =========================================================================
    # 8. SHEET KE_HOACH_QLCL: ĐỒNG BỘ NGHỊ ĐỊNH 207/2026/NĐ-CP
    # =========================================================================
    print("8. Đang xây dựng Sheet KE_HOACH_QLCL...")
    ws_ql = wb.create_sheet("KE_HOACH_QLCL")
    ws_ql["A1"] = "KẾ HOẠCH NGHIỆM THU & QUẢN LÝ CHẤT LƯỢNG (AEC-QLCL) CẦU KM19+529.080 – 23HG SYSTEM"
    ws_ql["A1"].font = f_title
    ws_ql["A2"] = "Tuân thủ Luật Xây dựng 135/2025/QH15, Nghị định 207/2026/NĐ-CP & Thông tư 32/2026/TT-BXD"
    ws_ql["A2"].font = f_sub

    ql_headers = [("A4", "Mã Biên Bản"), ("B4", "Tên Công Việc Nghiệm Thu (NĐ 207/2026/NĐ-CP)"), 
                  ("C4", "Cấu Kiện / Bộ Phận Công Trình"), ("D4", "Ngày Bắt Đầu"), 
                  ("E4", "Ngày Hoàn Thành"), ("F4", "Quy Chuẩn / Tiêu Chuẩn Áp Dụng"), 
                  ("G4", "Hồ Sơ & Chứng Chỉ Đính Kèm")]
    for pos, txt in ql_headers:
        ws_ql[pos] = txt
        ws_ql[pos].font = f_header
        ws_ql[pos].fill = fill_sand
        ws_ql[pos].alignment = align_center

    ql_items = [
        ("NT-01", "Nghiệm thu tim mốc định vị & mặt bằng thi công mố trụ", "Toàn cầu (WBS 1.1 - 1.2)", "=TIEN_DO!H7", "=TIEN_DO!I8", "TCVN 9394:2012", "Biên bản bàn giao mốc mạ vàng VN2000"),
        ("NT-02", "Nghiệm thu lỗ khoan thăm dò địa chất hang Karst", "26 vị trí cọc (WBS 2.1)", "=TIEN_DO!H12", "=TIEN_DO!I12", "TCVN 9361:2012", "Hình trụ hố khoan, mẫu đá lõi ngậm"),
        ("NT-03", "Nghiệm thu thổi rửa hố khoan & lồng thép cọc D1200 M1, T1", "Cọc M1 (3 cọc) & T1 (8 cọc)", "=TIEN_DO!H13", "=TIEN_DO!I14", "TCVN 10304:2014", "Chứng chỉ thép CB400, phiếu đo mùn lắng"),
        ("NT-04", "Nghiệm thu bê tông cọc khoan nhồi D1200 Trụ T2, M2", "Cọc T2 (8 cọc) & M2 (7 cọc)", "=TIEN_DO!H15", "=TIEN_DO!I16", "TCVN 9394:2012", "Phiếu xuất xưởng BT C30, tổ mẫu R28"),
        ("NT-05", "Nghiệm thu toàn diện chất lượng cọc nhồi (Siêu âm + PDA)", "Toàn bộ 26 cọc nhồi (WBS 2.6)", "=TIEN_DO!H17", "=TIEN_DO!I17", "ASTM D4945 / TCVN", "Biểu đồ biến dạng sóng PDA, sóng siêu âm"),
        ("NT-06", "Nghiệm thu đào ngàm đá hố móng & đổ bê tông lót C10", "Bệ mố M1, M2 & Trụ T1, T2", "=TIEN_DO!H19", "=TIEN_DO!I19", "TCVN 4447:2012", "Biên bản ngàm đá vôi 0.5m, cao độ đáy"),
        ("NT-07", "Nghiệm thu cốt thép & ván khuôn bệ trụ T1 (11.6x8.0x2.0m)", "Bệ trụ T1 (Zero-vát, 184.2 m3)", "=TIEN_DO!H21", "=TIEN_DO!I21", "TCVN 4453:1995", "Biên bản nghiệm thu Zero-vát bệ T1"),
        ("NT-08", "Nghiệm thu cốt thép & ván khuôn bệ trụ T2 (11.6x8.0x2.0m)", "Bệ trụ T2 (Zero-vát, 184.2 m3)", "=TIEN_DO!H22", "=TIEN_DO!I22", "TCVN 4453:1995", "Biên bản nghiệm thu Zero-vát bệ T2"),
        ("NT-09", "Nghiệm thu đổ bê tông C30 thân trụ T1 (H=23.35m, 3 đốt vát)", "Thân đặc trụ T1 (WBS 3.4.1)", "=TIEN_DO!H23", "=TIEN_DO!I23", "TCVN 4453:1995", "Kiểm tra độ thẳng đứng quả dọi laser"),
        ("NT-10", "Nghiệm thu đổ bê tông C30 thân trụ T2 (H=25.45m, 3 đốt vát)", "Thân đặc trụ T2 (WBS 3.4.2)", "=TIEN_DO!H24", "=TIEN_DO!I24", "TCVN 4453:1995", "Kiểm tra độ thẳng đứng quả dọi laser"),
        ("NT-11", "Nghiệm thu xà mũ trụ T1, T2 & đá kê gối cao su", "Đỉnh trụ T1, T2 (WBS 3.6)", "=TIEN_DO!H26", "=TIEN_DO!I27", "TCVN 10304:2014", "Cao độ đá kê sai số <= 2mm"),
        ("NT-12", "Nghiệm thu cốt thép, luồn cáp & đổ BT C45 15 dầm Super-T", "Bãi đúc dầm hiện trường", "=TIEN_DO!H29", "=TIEN_DO!I30", "TCVN 9114:2012", "Cáp DUL 15.2mm, chứng chỉ xuất xưởng"),
        ("NT-13", "Nghiệm thu căng kéo cáp DUL & bơm vữa không co ống gen", "15 phiến dầm Super-T (WBS 4.3)", "=TIEN_DO!H31", "=TIEN_DO!I31", "TCVN 10304:2014", "Độ dãn dài cáp, áp lực bơm vữa Sika"),
        ("NT-14", "Nghiệm thu lắp đặt 30 gối chậu cao su & lao dầm Super-T", "3 nhịp cầu 38.2m (WBS 4.4-4.5)", "=TIEN_DO!H32", "=TIEN_DO!I33", "TCVN 10268:2014", "Kiểm tra tim gối, độ tiếp xúc mặt đáy"),
        ("NT-15", "Nghiệm thu dầm ngang, bản liên tục nhiệt & bản mặt cầu C35", "Mặt cầu 118.2m (WBS 5.1-5.3)", "=TIEN_DO!H35", "=TIEN_DO!I37", "TCVN 3105 / TCVN 4453", "Phiếu kiểm tra độ võng, chiều dày 18cm"),
        ("NT-16", "Nghiệm thu lớp chống thấm & thảm bê tông nhựa C16 dày 7cm", "Mặt cầu Cầu Km19 (WBS 5.7)", "=TIEN_DO!H41", "=TIEN_DO!I41", "TCVN 8819:2011", "Độ nhám, độ bằng phẳng thước 3m"),
        ("NT-17", "Nghiệm thu hoàn thành công trình bàn giao đưa vào sử dụng", "Toàn bộ công trình Cầu Km19", "=TIEN_DO!H42", "=TIEN_DO!I42", "NĐ 207/2026/NĐ-CP", "Trọn bộ hồ sơ hoàn công & biên bản tổng")
    ]
    for idx, item in enumerate(ql_items, start=5):
        for c_idx, val in enumerate(item, start=1):
            cell = ws_ql.cell(row=idx, column=c_idx, value=val)
            cell.font = f_norm
            cell.border = box_border
            if c_idx in [1, 4, 5]:
                cell.alignment = align_center
                if c_idx in [4, 5]: cell.number_format = "dd/mm/yyyy"
            else:
                cell.alignment = align_left

    ws_ql.column_dimensions["A"].width = 12
    ws_ql.column_dimensions["B"].width = 48
    ws_ql.column_dimensions["C"].width = 30
    ws_ql.column_dimensions["D"].width = 16
    ws_ql.column_dimensions["E"].width = 16
    ws_ql.column_dimensions["F"].width = 24
    ws_ql.column_dimensions["G"].width = 42

    # =========================================================================
    # 9. SHEET HUONG_DAN
    # =========================================================================
    print("9. Đang xây dựng Sheet HUONG_DAN...")
    ws_hd = wb.create_sheet("HUONG_DAN")
    ws_hd["A1"] = "CẨM NANG VẬN HÀNH HỆ THỐNG TIẾN ĐỘ THI CÔNG 23HG SCHEDULE ASSISTANT PRO"
    ws_hd["A1"].font = f_title
    ws_hd["A2"] = "Tác giả: NBT (Nguyễn Bảo Tú) | GitHub: @baotuhg | 23HG SYSTEM"
    ws_hd["A2"].font = f_sub

    guide_lines = [
        "1. QUY ƯỚC LỊCH LÀM VIỆC & CPM TỰ ĐỘNG:",
        "   - Toàn bộ mạng tiến độ được tính toán bằng động cơ CPM chuẩn (Forward Pass & Backward Pass).",
        "   - Lịch làm việc công trường: 6 ngày/tuần (Thứ 2 đến Thứ 7; nghỉ Chủ nhật và các ngày Tết theo quy định).",
        "   - Công thức tính: EF = WORKDAY.INTL(ES, Thời_lượng - 1, '0000001', NGAY_NGHI_LE).",
        "   - Đường găng (GANG): Được xác định khi Dự trữ toàn phần TF (Total Float) <= 0. Màu đỏ nổi bật trên Gantt Canvas.",
        "",
        "2. CẤU TRÚC PHÂN CẤP CẦU KM19+529.080 (3 NHỊP SUPER-T L=38.2M):",
        "   - WBS 1: Chuẩn bị mặt bằng, đường công vụ nhánh 6 & 6A, bãi đúc dầm và trạm trộn BTXM 60m3/h.",
        "   - WBS 2: Khoan cọc nhồi D1200mm (26 cọc, khoan đá ngàm hang Karst sâu 5m, siêu âm 156 mặt cắt cọc & PDA).",
        "   - WBS 3: Kết cấu phần dưới (Đã tách riêng Trụ T1, T2; bệ trụ 11.6x8.0x2.0m Zero-vát; thân trụ 3 đốt có vát).",
        "   - WBS 4: Chế tạo & Lao lắp 15 phiến dầm Super-T 38.2m tại bãi đúc công trường, căng kéo cáp DUL 15.2mm.",
        "   - WBS 5: Bản mặt cầu, dầm ngang, bản quá độ, gờ lan can, thảm BTN C16 7cm và hoàn thành thông xe.",
        "",
        "3. ĐỒNG BỘ 100% CÔNG THỨC SỐNG & HỒ SƠ DOANH NGHIỆP:",
        "   - Sheet BOQ_TIEN_DO: Chi phí trực tiếp T, gián tiếp GT, thu nhập chịu thuế TL, thuế VAT sống 100%.",
        "   - Sheet EVM_5D_QUAN_TRI: Tự động tính toán các chỉ số PV, EV, AC, SPI, CPI tại Data Date.",
        "   - Sheet HUY_DONG_XMTB: Tự động cân bằng ca máy và tiêu thụ dầu diezel cho 14 loại máy móc thiết bị chủ lực.",
        "   - Sheet KE_HOACH_QLCL: Ngày nghiệm thu liên kết động với ngày hoàn thành công tác trên sheet TIEN_DO."
    ]
    for idx, gl in enumerate(guide_lines, start=4):
        ws_hd.cell(row=idx, column=1, value=gl).font = f_norm if not gl.startswith(("1.", "2.", "3.")) else f_norm_bold

    ws_hd.column_dimensions["A"].width = 110

    # Save workbook
    os.makedirs(os.path.dirname(OUT_BRIDGE_ASBUILT), exist_ok=True)
    wb.save(OUT_BRIDGE_ASBUILT)
    print(f"Đã lưu thành công Sổ tính Master Cầu Km19: {OUT_BRIDGE_ASBUILT}")

    # Copy to Master Dossier
    os.makedirs(os.path.dirname(OUT_MASTER_DOSSIER), exist_ok=True)
    wb.save(OUT_MASTER_DOSSIER)
    print(f"Đã đồng bộ sang Hồ sơ Master KCS-QS: {OUT_MASTER_DOSSIER}")

    # Export XML for MS Project
    export_mspdi_xml()

def export_mspdi_xml():
    print("Đang xuất file MS Project XML (MSPDI XML chuẩn quốc tế)...")
    root = ET.Element("Project", xmlns="http://schemas.microsoft.com/project")
    ET.SubElement(root, "Name").text = "TIEN_DO_CAU_KM19_529.080_23HG_PRO"
    ET.SubElement(root, "Title").text = "Tiến độ Thi công Cầu Km19+529.080 - 23HG Schedule Assistant Pro"
    ET.SubElement(root, "StartDate").text = "2026-10-01T07:00:00"
    ET.SubElement(root, "FinishDate").text = "2027-03-11T17:00:00"
    ET.SubElement(root, "ScheduleFromStart").text = "1"
    
    tasks_elem = ET.SubElement(root, "Tasks")
    wbs_uid_map = {}
    uid = 1

    for t in tasks_def:
        r_cpm = cpm_results[t.wbs]
        t_elem = ET.SubElement(tasks_elem, "Task")
        ET.SubElement(t_elem, "UID").text = str(uid)
        ET.SubElement(t_elem, "ID").text = str(uid)
        ET.SubElement(t_elem, "Name").text = t.name
        ET.SubElement(t_elem, "WBS").text = t.wbs
        ET.SubElement(t_elem, "OutlineNumber").text = t.wbs
        ET.SubElement(t_elem, "OutlineLevel").text = str(len(t.wbs.split('.')))
        ET.SubElement(t_elem, "Summary").text = "1" if t.kind == 'Summary' else "0"
        ET.SubElement(t_elem, "Milestone").text = "1" if t.kind == 'Milestone' else "0"
        
        # Duration: PT{hours}H0M0S
        dur_hours = t.duration * 8
        ET.SubElement(t_elem, "Duration").text = f"PT{dur_hours}H0M0S"
        ET.SubElement(t_elem, "Start").text = f"{r_cpm.es.strftime('%Y-%m-%d')}T07:00:00"
        ET.SubElement(t_elem, "Finish").text = f"{r_cpm.ef.strftime('%Y-%m-%d')}T17:00:00"
        ET.SubElement(t_elem, "Critical").text = "1" if r_cpm.critical else "0"

        wbs_uid_map[t.wbs] = uid
        uid += 1

    # Add Predecessor Links
    for t_idx, t in enumerate(tasks_def):
        if not t.preds:
            continue
        t_elem = tasks_elem[t_idx]
        for pred_wbs, rel, lag in t.preds:
            if pred_wbs in wbs_uid_map:
                link_elem = ET.SubElement(t_elem, "PredecessorLink")
                ET.SubElement(link_elem, "PredecessorUID").text = str(wbs_uid_map[pred_wbs])
                # Type: 1=FS, 3=SS, 0=FF, 2=SF
                type_val = "1" if rel == "FS" else ("3" if rel == "SS" else ("0" if rel == "FF" else "2"))
                ET.SubElement(link_elem, "Type").text = type_val
                # Lag in tenths of a minute (1 day = 8h = 480m = 4800 tenths)
                ET.SubElement(link_elem, "LinkLag").text = str(lag * 4800)

    tree = ET.ElementTree(root)
    ET.indent(tree, space="  ", level=0)
    tree.write(OUT_XML_PROJECT, encoding="utf-8", xml_declaration=True)
    print(f"Đã xuất thành công MSPDI XML: {OUT_XML_PROJECT}")

    # Copy to repo mirror
    if os.path.exists(REPO_DIR):
        repo_xml = os.path.join(REPO_DIR, "Tien_Do_Thi_Cong_Cau_Km19+529.080.xml")
        tree.write(repo_xml, encoding="utf-8", xml_declaration=True)
        print(f"Đã đồng bộ sang Repo: {repo_xml}")

if __name__ == "__main__":
    build_master_workbook()
