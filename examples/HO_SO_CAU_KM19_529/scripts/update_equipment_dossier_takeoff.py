# -*- coding: utf-8 -*-
"""
Cập nhật Sheet 05_DoiChieu_BocTach trong 260920_TDTC_CaXe_CaMay_DauDiezel_Cau_Km19+529.080.xlsx
Tách biệt triệt để Trụ T1 và T2, cập nhật chuẩn hình học Bệ không vát 184.243 m3/bệ, 3 đốt thân có vát.
"""

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

p = r"C:\Users\baotu\Downloads\HSTK Cầu Km19+529.080_Marker\04_CO_GIOI_THIET_BI_VA_DAU_DIEZEL\260920_TDTC_CaXe_CaMay_DauDiezel_Cau_Km19+529.080.xlsx"
wb = openpyxl.load_workbook(p)
ws = wb['05_DoiChieu_BocTach']

# Formatting
f_title = Font(name="Segoe UI", size=11, bold=True, color="5A3718")
f_norm = Font(name="Segoe UI", size=9, color="000000")
f_norm_bold = Font(name="Segoe UI", size=9, bold=True, color="000000")
f_header = Font(name="Segoe UI", size=8.5, bold=True, color="333333")
fill_sand = PatternFill(fill_type="solid", start_color="EFE2D3", end_color="EFE2D3")

box_border = Border(
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1'),
    top=Side(style='thin', color='CBD5E1'),
    bottom=Side(style='thin', color='CBD5E1')
)

align_center = Alignment(horizontal="center", vertical="center")
align_left = Alignment(horizontal="left", vertical="center")
align_right = Alignment(horizontal="right", vertical="center")

# Clear existing rows below row 3
for r in range(4, ws.max_row + 10):
    for c in range(1, 9):
        ws.cell(r, c).value = None
        ws.cell(r, c).fill = PatternFill(fill_type=None)
        ws.cell(r, c).border = Border()

# Header
ws["A1"] = "BẢNG ĐỐI CHIẾU KHỐI LƯỢNG THIẾT KẾ CẦU KM19+529.080 (CHUẨN HÓA TÁCH BẠCH TRỤ T1 & T2)"
ws["A1"].font = f_title

headers = ["STT", "Hạng mục kết cấu Cầu Km19+529.080", "Quy cách / Kích thước", 
           "Khối lượng Bê tông (m3)", "Mác Bê tông", "Ván khuôn tiếp xúc (m2)", 
           "Cốt thép thường (tấn)", "Cáp DƯL / Ghi chú kỹ thuật"]
for c_idx, h in enumerate(headers, start=1):
    cell = ws.cell(row=3, column=c_idx, value=h)
    cell.font = f_header
    cell.fill = fill_sand
    cell.alignment = align_center
    cell.border = box_border

updated_takeoff = [
    (1, "26 Cọc khoan nhồi D1200mm (M1, M2, T1, T2)", "26 cọc, L_tb=33.5m, ngàm đá Karst", 1006.54, "C30", 3286.0, 116.35, "Khoan ngàm đá hang Karst sâu 5m"),
    (2, "Bê tông lót đáy bệ móng mố M1, M2 & trụ T1, T2", "Dày 10cm đá 1x2 đáy bệ móng", 42.15, "C10", 84.3, 0.0, "Đổ lót đáy móng sau đào đá"),
    (3, "Bệ móng mố M1 (78.4 m3) & mố M2 (94.8 m3)", "Bệ móng chữ nhật ngàm cọc", 173.20, "C30", 245.6, 28.50, "Cốt thép chịu lực Phi 28 - Phi 32"),
    (4, "Bệ móng trụ T1 (11.6m x 8.0m x 2.0m)", "Khối hộp 11.6x8.0x2.0m KHÔNG VÁT", 184.243, "C30", 78.4, 28.46, "Chuẩn hóa Zero-vát bệ T1 (184.243 m3)"),
    (5, "Bệ móng trụ T2 (11.6m x 8.0m x 2.0m)", "Khối hộp 11.6x8.0x2.0m KHÔNG VÁT", 184.243, "C30", 78.4, 28.46, "Chuẩn hóa Zero-vát bệ T2 (184.243 m3)"),
    (6, "Thân mố chân dê M1 (68.2 m3) & mố chữ U M2 (70.2 m3)", "Mố tường đặc kèm tường cánh", 138.40, "C30", 356.8, 24.15, "Bao gồm tường thân, tường cánh, đá kê"),
    (7, "Thân trụ đặc T1 (H=23.35m, 3 đốt thân có vát)", "Đốt 3 chân loe, Đốt 2 thân, Đốt 1 đỉnh", 234.922, "C30", 432.5, 41.20, "Vát góc 4 cạnh, đà giáo leo thi công 3 đợt"),
    (8, "Thân trụ đặc T2 (H=25.45m, 3 đốt thân có vát)", "Đốt 3 chân loe, Đốt 2 thân, Đốt 1 đỉnh", 266.791, "C30", 471.2, 46.80, "Vát góc 4 cạnh, đà giáo leo thi công 3 đợt"),
    (9, "Xà mũ trụ T1 & khối ụ chống xô", "Xà mũ vát đầu trụ T1, L=11.6m", 128.851, "C30", 242.5, 23.40, "Đá kê gối cầu sai số <= 2mm"),
    (10, "Xà mũ trụ T2 & khối ụ chống xô", "Xà mũ vát đầu trụ T2, L=11.6m", 145.026, "C30", 258.6, 26.20, "Đá kê gối cầu sai số <= 2mm"),
    (11, "15 Phiến dầm Super-T L=38.2m tại bãi đúc", "15 phiến dầm mặt cắt I loe", 434.80, "C45/55", 2610.0, 96.72, "27.78 tấn cáp DƯL 15.2mm (660 tao)"),
    (12, "Dầm ngang mố trụ & bản liên tục nhiệt", "Bê tông nối nhịp đỉnh trụ", 78.50, "C35", 185.0, 14.20, "Liên kết dầm Super-T trên đỉnh mố trụ"),
    (13, "Bản mặt cầu tại chỗ C35 (dày 180mm)", "Dày 180mm, L=118.2m, B=12.0m", 235.60, "C35", 1250.0, 42.50, "Phối hợp 510 tấm ván khuôn đúc sẵn"),
    (14, "Bản quá độ sau 2 mố & đắp K98 đầu cầu", "2 bản L=5.0m sau mố M1, M2", 35.80, "C25", 86.0, 5.80, "Đắp chọn lọc K98 đầm chặt"),
    (15, "Gờ lan can bê tông & khe co giãn, thảm BTN", "Gờ LC 2 bên L=260.8m, BTN C16 7cm", 52.40, "C25", 380.0, 8.50, "Thảm BTN C16 diện tích 1,620 m2")
]

for idx, row_vals in enumerate(updated_takeoff, start=4):
    for c_i, val in enumerate(row_vals, start=1):
        cell = ws.cell(row=idx, column=c_i, value=val)
        cell.font = f_norm
        cell.border = box_border
        if c_i in [1, 5]: cell.alignment = align_center
        elif c_i in [4, 6, 7]:
            cell.alignment = align_right
            cell.number_format = "#,##0.00" if isinstance(val, float) else "#,##0"
        else: cell.alignment = align_left

tot_r = len(updated_takeoff) + 4
ws.cell(row=tot_r, column=2, value="TỔNG CỘNG TOÀN CÔNG TRÌNH CẦU KM19+529.080 (100%)").font = f_norm_bold
ws.cell(row=tot_r, column=4, value=f"=SUM(D4:D{tot_r-1})").number_format = "#,##0.00"
ws.cell(row=tot_r, column=6, value=f"=SUM(F4:F{tot_r-1})").number_format = "#,##0.0"
ws.cell(row=tot_r, column=7, value=f"=SUM(G4:G{tot_r-1})").number_format = "#,##0.00"

for c_k in range(1, 9):
    c_tot = ws.cell(row=tot_r, column=c_k)
    c_tot.font = f_norm_bold
    c_tot.fill = fill_sand
    c_tot.border = box_border

ws.column_dimensions["A"].width = 6
ws.column_dimensions["B"].width = 46
ws.column_dimensions["C"].width = 38
ws.column_dimensions["D"].width = 18
ws.column_dimensions["E"].width = 12
ws.column_dimensions["F"].width = 18
ws.column_dimensions["G"].width = 18
ws.column_dimensions["H"].width = 42

wb.save(p)
print(f"Đã cập nhật thành công Sheet 05_DoiChieu_BocTach: {p}")
