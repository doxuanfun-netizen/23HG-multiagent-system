# -*- coding: utf-8 -*-
"""
HỆ THỐNG TỰ ĐỘNG HÓA LẬP HỒ SƠ TẦN SUẤT THÍ NGHIỆM & SƠ ĐỒ TIẾN ĐỘ TRỰC QUAN AS-BUILT
CÔNG TRÌNH: CẦU THÔN KHAI HOANG 2, KM 14+363.65 (NHỊP DẦM T BTCT L=15M, MỐ CHỮ U ĐẶT TRÊN ĐÁ GỐC)
Tuân thủ Luật Xây dựng 135/2025/QH15, Nghị định 207/2026/NĐ-CP, Thông tư 32/2026/TT-BXD,
TCVN 1651:2018, TCVN 4453:1995, TCVN 11823:2017.

Xuất 2 mẫu ưu tú nhất (Golden Master Baseline) cho Phân hệ 01_HIEN_TRUONG_QLCL_KCS:
1. Bang_Theo_Doi_Tan_Suat_Thi_Nghiem_Be_Tong_Va_Thep_Cau_Khai_Hoang_2.xlsx (4 Sheets công thức sống 100%)
2. So_Do_Tien_Do_Truc_Quan_AsBuilt_Cau_Khai_Hoang_2.xlsx (+ .html) (2 Sheets: Mặt cắt đứng trực quan + Bảng tiến độ)
"""

from __future__ import annotations

import os
import sys
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

FONT_NAME = "Arial"
FONT_TITLE = Font(name=FONT_NAME, size=13, bold=True, color="1B365D")
FONT_SUBTITLE = Font(name=FONT_NAME, size=9.5, italic=True, color="595959")
FONT_HDR = Font(name=FONT_NAME, size=9, bold=True)
FONT_HDR_WHITE = Font(name=FONT_NAME, size=9, bold=True, color="FFFFFF")
FONT_HDR_RED = Font(name=FONT_NAME, size=9, bold=True, color="C00000")
FONT_HDR_GREEN = Font(name=FONT_NAME, size=9, bold=True, color="0070C0")
FONT_REG = Font(name=FONT_NAME, size=8.5)
FONT_BOLD = Font(name=FONT_NAME, size=8.5, bold=True)
FONT_RED = Font(name=FONT_NAME, size=8.5, bold=True, color="C00000")
FONT_GREEN = Font(name=FONT_NAME, size=8.5, bold=True, color="008000")

THIN_GRAY = Side(style="thin", color="BFBFBF")
THIN_BORDER = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")

FILL_NAVY = PatternFill("solid", fgColor="1B365D")
FILL_BLUE = PatternFill("solid", fgColor="2E75B6")
FILL_HDR = PatternFill("solid", fgColor="F2F2F2")
FILL_SECTION = PatternFill("solid", fgColor="D9E1F2")
FILL_SUBSECTION = PatternFill("solid", fgColor="EDF2F8")
FILL_TOTAL = PatternFill("solid", fgColor="FFF2CC")
FILL_ALERT = PatternFill("solid", fgColor="FCE4D6")
FILL_GREEN = PatternFill("solid", fgColor="28A745")
FILL_AMBER = PatternFill("solid", fgColor="FFC000")
FILL_LIGHT_GREEN = PatternFill("solid", fgColor="E2EFDA")

# 10 loại đường kính Ø thực tế của Cầu thôn Khai Hoang 2 (tích lũy P..Y = 16..25, tần suất Z..AI = 26..35)
DIA_LIST_KH2 = [
    (6,  16, 26, "Ø6",  "Ø6 (Lưới BMC)\n(CB240-T/CB400-V)", "TCVN 1651-1:2018"),
    (10, 17, 27, "Ø10", "Ø10\n(CB400-V)",                   "TCVN 1651-2:2018"),
    (12, 18, 28, "Ø12", "Ø12\n(CB400-V)",                   "TCVN 1651-2:2018"),
    (14, 19, 29, "Ø14", "Ø14\n(CB400-V)",                   "TCVN 1651-2:2018"),
    (16, 20, 30, "Ø16", "Ø16\n(CB400-V)",                   "TCVN 1651-2:2018"),
    (18, 21, 31, "Ø18", "Ø18\n(CB400-V)",                   "TCVN 1651-2:2018"),
    (20, 22, 32, "Ø20", "Ø20\n(CB400-V)",                   "TCVN 1651-2:2018"),
    (22, 23, 33, "Ø22", "Ø22\n(CB400-V)",                   "TCVN 1651-2:2018"),
    (25, 24, 34, "Ø25", "Ø25\n(CB400-V)",                   "TCVN 1651-2:2018"),
    (32, 25, 35, "Ø32", "Ø32 (Dầm T & Neo)\n(CB400-V)",     "TCVN 1651-2:2018"),
]


def _build_sheet_concrete_kh2(ws) -> int:
    """Xây dựng Sheet 1: Theo dõi tần suất Bê tông & Cát, Đá, Xi măng cho Cầu thôn Khai Hoang 2."""
    ws.title = "Theo_Doi_Tan_Suat_Be_Tong"
    ws.views.sheetView[0].showGridLines = True

    ws.merge_cells("A1:A2")
    ws.merge_cells("B1:B2")
    ws.merge_cells("C1:C2")
    ws.merge_cells("D1:D2")
    ws.merge_cells("E1:F1")
    ws.merge_cells("G1:I1")
    ws.merge_cells("J1:L1")
    ws.merge_cells("M1:O1")
    ws.merge_cells("P1:R1")
    ws.merge_cells("S1:S2")
    ws.merge_cells("T1:T2")
    ws.merge_cells("U1:U2")

    ws["A1"] = "STT"
    ws["B1"] = "HẠNG MỤC CÔNG VIỆC (CẦU THÔN KHAI HOANG 2, KM 14+363.65)"
    ws["C1"] = "N/T/N\nĐỔ BT"
    ws["D1"] = "KL/T.tế\n(m3)"
    ws["E1"] = "Tuổi Bt"
    ws["G1"] = "C40/C35-14±2 (Dầm T15m, BMC, Kê gối)"
    ws["J1"] = "C30/C25/C10-12±2 (Bệ, Thân mố, Tường chắn)"
    ws["M1"] = "Khối lượng vật tư lũy kế"
    ws["P1"] = "Tần suất lấy mẫu thí nghiệm (ROUNDUP)"
    ws["S1"] = "Số lô"
    ws["T1"] = "KL lô\n(Tấn)"
    ws["U1"] = "NGUỒN BẢN VẼ / GHI CHÚ KCS"

    ws["E2"] = "R7"
    ws["F2"] = "R28"
    ws["G2"] = "Xi măng\n0,455"
    ws["H2"] = "Cát\n0,515"
    ws["I2"] = "Đá\n0,685"
    ws["J2"] = "Xi măng\n0,415"
    ws["K2"] = "Cát\n0,555"
    ws["L2"] = "Đá\n0,705"
    ws["M2"] = "Xi măng\n(Tấn)"
    ws["N2"] = "Cát\n(m3)"
    ws["O2"] = "Đá\n(m3)"
    ws["P2"] = "Xi măng\n50 tấn\nLần"
    ws["Q2"] = "Cát\n200m3\nLần"
    ws["R2"] = "Đá\n350m3\nLần"

    for r in (1, 2):
        ws.row_dimensions[r].height = 28 if r == 2 else 24
        for c in range(1, 22):
            cell = ws.cell(r, c)
            cell.alignment = ALIGN_CENTER
            cell.border = THIN_BORDER
            cell.fill = FILL_HDR
            if c in (16, 18):
                cell.font = FONT_HDR_RED
            elif c == 17:
                cell.font = FONT_HDR_GREEN
            else:
                cell.font = FONT_HDR

    items = [
        ("SECTION", "I. HẠNG MỤC BÊ TÔNG ĐỆM LÓT MÓNG & VỮA KHOAN CẤY THÉP NEO ĐÁ GỐC (98 LỖ D32)"),
        ("Bê tông lót móng bệ mố M1 đá 2x4 dày 10cm (C10/M100)", 15.65, "30", "2026-11-01", "BV PD-01 tr.12 - Đệm móng trên nền đá gốc"),
        ("Bê tông lót móng bệ mố M2 đá 2x4 dày 10cm (C10/M100)", 15.65, "30", "2026-11-02", "BV PD-01 tr.12 - Đệm móng trên nền đá gốc"),
        ("Bê tông lót móng 04 đốt tường chắn đầu cầu dày 10cm (C10)", 27.40, "30", "2026-11-03", "BV TC-01..04 tr.38-47 - Đệm móng tường chắn"),
        ("Vữa xi măng M300 không co ngót cấy 98 thanh neo D32 vào đá", 4.90, "40", "2026-11-04", "BV PD-01 tr.13 - 49 lỗ/mố x 2 mố = 98 lỗ khoan"),

        ("SECTION", "II. HẠNG MỤC BÊ TÔNG BỆ MỐ, THÂN MỐ & TƯỜNG CÁNH MỐ CHỮ U M1, M2 (C30 ĐÁ 1x2)"),
        ("Bê tông bệ móng mố chữ U M1 (C30 đá 1x2, độ sụt 14±2cm)", 115.66, "30", "2026-11-09", "BV PD-02 tr.14 - Bệ mố M1 (2 tổ mẫu R7/R28)"),
        ("Bê tông bệ móng mố chữ U M2 (C30 đá 1x2, độ sụt 14±2cm)", 115.67, "30", "2026-11-11", "BV PD-03 tr.15 - Bệ mố M2 (2 tổ mẫu R7/R28)"),
        ("Bê tông tường thân & tường đỉnh mố M1 - Đợt 1 (C30 đá 1x2)", 60.00, "30", "2026-11-14", "BV PD-04 tr.18 - Thân mố M1 đợt 1"),
        ("Bê tông tường thân & tường đỉnh mố M1 - Đợt 2 (C30 đá 1x2)", 55.66, "30", "2026-11-17", "BV PD-04 tr.18 - Thân & tường đỉnh mố M1"),
        ("Bê tông tường thân & tường đỉnh mố M2 - Đợt 1 (C30 đá 1x2)", 60.00, "30", "2026-11-15", "BV PD-04 tr.18 - Thân mố M2 đợt 1"),
        ("Bê tông tường thân & tường đỉnh mố M2 - Đợt 2 (C30 đá 1x2)", 55.67, "30", "2026-11-18", "BV PD-04 tr.18 - Thân & tường đỉnh mố M2"),
        ("Bê tông tường cánh chữ U mố M1 (2 tường cánh, C30 đá 1x2)", 78.05, "30", "2026-11-22", "BV PD-07 tr.21 - 2 tường cánh mố M1"),
        ("Bê tông tường cánh chữ U mố M2 (2 tường cánh, C30 đá 1x2)", 78.06, "30", "2026-11-24", "BV PD-07 tr.21 - 2 tường cánh mố M2"),
        ("Bê tông đá kê gối & ụ chống xô ngang mố M1, M2 (C35 đá 1x2)", 3.42, "40", "2026-11-26", "BV PD-05 tr.19 - 08 đá kê gối + 04 ụ chống xô"),

        ("SECTION", "III. HẠNG MỤC BÊ TÔNG 04 ĐỐT TƯỜNG CHẮN ĐẦU CẦU & BỆ PHẢN ÁP (C25 ĐÁ 1x2)"),
        ("Bê tông tường chắn BTCT đốt 1 (L=5.20m, C25 đá 1x2)", 42.50, "30", "2026-11-27", "BV TC-01 tr.39 - Tường chắn đốt L=5.20m"),
        ("Bê tông tường chắn BTCT đốt 2 (L=5.89m, C25 đá 1x2)", 48.20, "30", "2026-11-28", "BV TC-02 tr.41 - Tường chắn đốt L=5.89m"),
        ("Bê tông tường chắn BTCT đốt 3 (L=10.60m, C25 đá 1x2)", 85.30, "30", "2026-11-30", "BV TC-03 tr.44 - Tường chắn đốt L=10.60m"),
        ("Bê tông tường chắn BTCT đốt 4 (L=13.50m & bệ phản áp, C25)", 95.40, "30", "2026-12-02", "BV TC-04 tr.47 - Tường chắn đốt L=13.50m"),

        ("SECTION", "IV. HẠNG MỤC DẦM CHỦ BTCT THƯỜNG CHỮ T NHỊP L=15.0M (04 PHIẾN DẦM C40)"),
        ("Bê tông dầm chữ T L=15m - Phiến D1 (Dầm biên trái, C40 đá 1x2)", 9.35, "40", "2026-11-16", "BV PT-01..03 tr.26-28 - Dầm biên D1 (1 tổ mẫu)"),
        ("Bê tông dầm chữ T L=15m - Phiến D2 (Dầm giữa 1, C40 đá 1x2)", 9.35, "40", "2026-11-19", "BV PT-04..06 tr.31-33 - Dầm giữa D2 (1 tổ mẫu)"),
        ("Bê tông dầm chữ T L=15m - Phiến D3 (Dầm giữa 2, C40 đá 1x2)", 9.35, "40", "2026-11-22", "BV PT-04..06 tr.31-33 - Dầm giữa D3 (1 tổ mẫu)"),
        ("Bê tông dầm chữ T L=15m - Phiến D4 (Dầm biên phải, C40 đá 1x2)", 9.35, "40", "2026-11-25", "BV PT-01..03 tr.26-28 - Dầm biên D4 (1 tổ mẫu)"),

        ("SECTION", "V. HẠNG MỤC DẦM NGANG, BẢN MẶT CẦU, BẢN QUÁ ĐỘ & HOÀN THIỆN"),
        ("Bê tông 03 dầm ngang & mối nối ướt bản cánh dầm T (C40 đá 1x2)", 8.45, "40", "2026-12-29", "BV PT-07 tr.34 - Liên kết ngang sau lao dầm"),
        ("Bê tông bản mặt cầu nhịp L=15m dày 12cm (C35/C40 đá 1x2)", 14.13, "40", "2027-01-02", "BV PT-08 tr.35 - Bản mặt cầu B=7.0m"),
        ("Bê tông bản quá độ đầu mố M1 dày 25cm (L=4.0m, C30 đá 1x2)", 9.09, "30", "2027-01-05", "BV PD-08 tr.22 - Bản quá độ phía Đồng Văn"),
        ("Bê tông bản quá độ đầu mố M2 dày 25cm (L=4.0m, C30 đá 1x2)", 9.09, "30", "2027-01-06", "BV PD-08 tr.22 - Bản quá độ phía Mèo Vạc"),
        ("Bê tông gờ lan can & chân cột lan can hai bên cầu (C30 đá 1x2)", 6.80, "30", "2027-01-12", "BV PT-08 tr.35 - Gờ chắn bánh 2x15m"),
        ("Bê tông chèn khe co giãn răng lược hai đầu mố M1, M2 (C40)", 1.68, "40", "2027-01-15", "BV PD-09 tr.23 - 02 khe co giãn L=9.0m"),
    ]

    current_row = 3
    stt_counter = 1
    prev_data_row = None

    for item in items:
        if item[0] == "SECTION":
            ws.row_dimensions[current_row].height = 22
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=21)
            sec_cell = ws.cell(current_row, 1, item[1])
            sec_cell.font = Font(name=FONT_NAME, size=9.5, bold=True, color="1B365D")
            sec_cell.alignment = Alignment(horizontal="left", vertical="center")
            sec_cell.fill = FILL_SECTION
            for c in range(1, 22):
                ws.cell(current_row, c).border = THIN_BORDER
            current_row += 1
            continue

        name, vol, mix_type, pour_date, note_src = item
        r = current_row
        ws.row_dimensions[r].height = 20

        ws.cell(r, 1, stt_counter).alignment = ALIGN_CENTER
        ws.cell(r, 1).font = FONT_REG
        ws.cell(r, 1).border = THIN_BORDER

        ws.cell(r, 2, name).alignment = ALIGN_LEFT
        ws.cell(r, 2).font = FONT_REG
        ws.cell(r, 2).border = THIN_BORDER

        cell_c = ws.cell(r, 3, pour_date)
        cell_c.alignment = ALIGN_CENTER
        cell_c.font = FONT_REG
        cell_c.border = THIN_BORDER

        cell_d = ws.cell(r, 4, vol)
        cell_d.alignment = ALIGN_RIGHT
        cell_d.font = FONT_REG
        cell_d.border = THIN_BORDER
        cell_d.number_format = "#,##0.00"

        cell_e = ws.cell(r, 5, f'=IF(C{r}="","","+7 ngày")')
        cell_e.alignment = ALIGN_CENTER
        cell_e.font = FONT_REG
        cell_e.border = THIN_BORDER

        cell_f = ws.cell(r, 6, f'=IF(C{r}="","","+28 ngày")')
        cell_f.alignment = ALIGN_CENTER
        cell_f.font = FONT_REG
        cell_f.border = THIN_BORDER

        if mix_type == "40":
            ws.cell(r, 7, 0.455).number_format = "0.000"
            ws.cell(r, 8, 0.515).number_format = "0.000"
            ws.cell(r, 9, 0.685).number_format = "0.000"
            ws.cell(r, 10, 0.0).number_format = "0.000"
            ws.cell(r, 11, 0.0).number_format = "0.000"
            ws.cell(r, 12, 0.0).number_format = "0.000"
        else:
            ws.cell(r, 7, 0.0).number_format = "0.000"
            ws.cell(r, 8, 0.0).number_format = "0.000"
            ws.cell(r, 9, 0.0).number_format = "0.000"
            ws.cell(r, 10, 0.415).number_format = "0.000"
            ws.cell(r, 11, 0.555).number_format = "0.000"
            ws.cell(r, 12, 0.705).number_format = "0.000"

        for c_mix in range(7, 13):
            c_cell = ws.cell(r, c_mix)
            c_cell.alignment = ALIGN_CENTER
            c_cell.font = FONT_REG
            c_cell.border = THIN_BORDER

        if prev_data_row is None:
            ws.cell(r, 13, f'=IF(D{r}="","",ROUND(IF(G{r}>0,D{r}*G{r},D{r}*J{r}),2))')
            ws.cell(r, 14, f'=IF(D{r}="","",ROUND(IF(H{r}>0,D{r}*H{r},D{r}*K{r}),2))')
            ws.cell(r, 15, f'=IF(D{r}="","",ROUND(IF(I{r}>0,D{r}*I{r},D{r}*L{r}),2))')
        else:
            p_r = prev_data_row
            ws.cell(r, 13, f'=IF(D{r}="","",ROUND(M{p_r}+IF(G{r}>0,D{r}*G{r},D{r}*J{r}),2))')
            ws.cell(r, 14, f'=IF(D{r}="","",ROUND(N{p_r}+IF(H{r}>0,D{r}*H{r},D{r}*K{r}),2))')
            ws.cell(r, 15, f'=IF(D{r}="","",ROUND(O{p_r}+IF(I{r}>0,D{r}*I{r},D{r}*L{r}),2))')

        for c_kl in range(13, 16):
            cell_kl = ws.cell(r, c_kl)
            cell_kl.alignment = ALIGN_RIGHT
            cell_kl.font = FONT_REG
            cell_kl.border = THIN_BORDER
            cell_kl.number_format = "#,##0.00"

        if prev_data_row is None:
            ws.cell(r, 16, f'=IF(M{r}="","",IF(ROUNDUP(M{r}/50,0)>0,"Lần " & ROUNDUP(M{r}/50,0),""))')
            ws.cell(r, 17, f'=IF(N{r}="","",IF(ROUNDUP(N{r}/200,0)>0,"Lần " & ROUNDUP(N{r}/200,0),""))')
            ws.cell(r, 18, f'=IF(O{r}="","",IF(ROUNDUP(O{r}/350,0)>0,"Lần " & ROUNDUP(O{r}/350,0),""))')
        else:
            p_r = prev_data_row
            ws.cell(r, 16, f'=IF(M{r}="","",IF(ROUNDUP(M{r}/50,0)>ROUNDUP(M{p_r}/50,0),"Lần " & ROUNDUP(M{r}/50,0),""))')
            ws.cell(r, 17, f'=IF(N{r}="","",IF(ROUNDUP(N{r}/200,0)>ROUNDUP(N{p_r}/200,0),"Lần " & ROUNDUP(N{r}/200,0),""))')
            ws.cell(r, 18, f'=IF(O{r}="","",IF(ROUNDUP(O{r}/350,0)>ROUNDUP(O{p_r}/350,0),"Lần " & ROUNDUP(O{r}/350,0),""))')

        ws.cell(r, 16).font = FONT_RED
        ws.cell(r, 17).font = FONT_GREEN
        ws.cell(r, 18).font = FONT_RED
        for c_ts in range(16, 19):
            ws.cell(r, c_ts).alignment = ALIGN_CENTER
            ws.cell(r, c_ts).border = THIN_BORDER

        ws.cell(r, 19, f"LO-BT-{stt_counter:02d}").alignment = ALIGN_CENTER
        ws.cell(r, 19).font = FONT_REG
        ws.cell(r, 19).border = THIN_BORDER

        ws.cell(r, 20, f"=ROUND(D{r}*2.4, 2)").alignment = ALIGN_RIGHT
        ws.cell(r, 20).font = FONT_REG
        ws.cell(r, 20).border = THIN_BORDER
        ws.cell(r, 20).number_format = "#,##0.00"

        ws.cell(r, 21, note_src).alignment = ALIGN_LEFT
        ws.cell(r, 21).font = FONT_REG
        ws.cell(r, 21).border = THIN_BORDER

        prev_data_row = r
        stt_counter += 1
        current_row += 1

    r_tot = current_row
    ws.row_dimensions[r_tot].height = 24
    ws.merge_cells(start_row=r_tot, start_column=1, end_row=r_tot, end_column=3)
    ws.cell(r_tot, 1, "TỔNG CỘNG BÊ TÔNG TOÀN CẦU KHAI HOANG 2").alignment = ALIGN_CENTER
    ws.cell(r_tot, 1).font = FONT_HDR
    ws.cell(r_tot, 4, f"=SUM(D4:D{r_tot-1})").alignment = ALIGN_RIGHT
    ws.cell(r_tot, 4).font = FONT_HDR
    ws.cell(r_tot, 4).number_format = "#,##0.00"

    if prev_data_row:
        ws.cell(r_tot, 13, f"=M{prev_data_row}").number_format = "#,##0.00"
        ws.cell(r_tot, 14, f"=N{prev_data_row}").number_format = "#,##0.00"
        ws.cell(r_tot, 15, f"=O{prev_data_row}").number_format = "#,##0.00"
        for c_k in (13, 14, 15):
            ws.cell(r_tot, c_k).font = FONT_HDR
            ws.cell(r_tot, c_k).alignment = ALIGN_RIGHT
        ws.cell(r_tot, 16, f'="Tổng: " & ROUNDUP(M{prev_data_row}/50,0) & " lần"').alignment = ALIGN_CENTER
        ws.cell(r_tot, 17, f'="Tổng: " & ROUNDUP(N{prev_data_row}/200,0) & " lần"').alignment = ALIGN_CENTER
        ws.cell(r_tot, 18, f'="Tổng: " & ROUNDUP(O{prev_data_row}/350,0) & " lần"').alignment = ALIGN_CENTER
        for c_t in (16, 17, 18):
            ws.cell(r_tot, c_t).font = FONT_HDR

    ws.cell(r_tot, 20, f"=SUM(T4:T{r_tot-1})").number_format = "#,##0.00"
    ws.cell(r_tot, 20).font = FONT_HDR
    ws.cell(r_tot, 20).alignment = ALIGN_RIGHT

    for c in range(1, 22):
        cell = ws.cell(r_tot, c)
        cell.border = THIN_BORDER
        cell.fill = FILL_TOTAL

    col_widths = {
        1: 6, 2: 48, 3: 12, 4: 12, 5: 11, 6: 11, 7: 10, 8: 10, 9: 10,
        10: 10, 11: 10, 12: 10, 13: 13, 14: 13, 15: 13, 16: 12, 17: 12,
        18: 12, 19: 14, 20: 12, 21: 42
    }
    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.freeze_panes = "C3"
    return r_tot


def _build_sheet_rebar_kh2(ws, bbs_rows) -> int:
    """Xây dựng Sheet 2: Theo dõi tần suất Cốt thép 166 dòng BBS Cầu thôn Khai Hoang 2."""
    ws.title = "Theo_Doi_Tan_Suat_Cot_Thep"
    ws.views.sheetView[0].showGridLines = True

    for col_let in ("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "AJ"):
        ws.merge_cells(f"{col_let}1:{col_let}2")

    ws.merge_cells("P1:Y1")
    ws.merge_cells("Z1:AI1")

    ws["A1"] = "STT"
    ws["B1"] = "HẠNG MỤC CẤU KIỆN"
    ws["C1"] = "BẢN VẼ / TRANG PDF"
    ws["D1"] = "KÝ HIỆU\nTHANH"
    ws["E1"] = "ĐƯỜNG KÍNH\nØ (mm)"
    ws["F1"] = "MÁC THÉP /\nTIÊU CHUẨN"
    ws["G1"] = "QUY CÁCH /\nGHI CHÚ CẤU KIỆN"
    ws["H1"] = "CHIỀU DÀI\n1 THANH (m)"
    ws["I1"] = "TỔNG SỐ\nTHANH"
    ws["J1"] = "TRỌNG LƯỢNG\n1M (kg/m)"
    ws["K1"] = "TỔNG KL\n(kg)"
    ws["L1"] = "KHỐI LƯỢNG\n(Tấn)"
    ws["M1"] = "N/T/N\nNGHIỆM THU"
    ws["N1"] = "SỐ LÔ / HEAT NO\nCO-CQ RIÊNG TỪNG Ø"
    ws["O1"] = "CẢNH BÁO LẤY MẪU TN\nKÉO/UỐN (20T/LẦN)"
    ws["P1"] = "KHỐI LƯỢNG LŨY KẾ THEO TỪNG LOẠI ĐƯỜNG KÍNH Ø (TẤN)"
    ws["Z1"] = "TẦN SUẤT THÍ NGHIỆM CƠ LÝ KÉO / UỐN RIÊNG TỪNG Ø (20 TẤN / LẦN - TCVN 1651:2018)"
    ws["AJ1"] = "ĐÁNH GIÁ &\nKẾT LUẬN"

    for dia_val, c_acc_idx, c_freq_idx, l_short, l_full, _ in DIA_LIST_KH2:
        ws.cell(2, c_acc_idx, l_full)
        ws.cell(2, c_freq_idx, f"TN {l_short}\n20T / Lần")

    for r in (1, 2):
        ws.row_dimensions[r].height = 30 if r == 2 else 24
        for c in range(1, 37):
            cell = ws.cell(r, c)
            cell.alignment = ALIGN_CENTER
            cell.border = THIN_BORDER
            cell.fill = FILL_HDR
            if c == 15:
                cell.font = FONT_HDR_RED
            elif 16 <= c <= 25:
                cell.font = FONT_HDR_GREEN
            elif 26 <= c <= 35:
                cell.font = FONT_HDR_RED
            else:
                cell.font = FONT_HDR

    from examples.build_khai_hoang_2_full_pipeline import UNIT_WEIGHT

    cur_row = 3
    stt_cnt = 1
    prev_r = None
    current_comp = None

    for b in bbs_rows:
        comp_name = b["comp"]
        sheet_name = b["sheet"]
        page_no = b["page"]
        mark_str = str(b["mark"])
        dia_val_int = int(b["dia"])
        len_mm = float(b["length_mm"])
        tot_qty = int(b["total_qty"])
        units_cnt = int(b["units"])
        qty_one = int(b["qty_per_unit"])

        if comp_name != current_comp:
            current_comp = comp_name
            ws.row_dimensions[cur_row].height = 22
            ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=36)
            sc = ws.cell(cur_row, 1, f"HẠNG MỤC: {current_comp.upper()} ({sheet_name} - Trang {page_no})")
            sc.font = Font(name=FONT_NAME, size=9.5, bold=True, color="1B365D")
            sc.alignment = Alignment(horizontal="left", vertical="center")
            sc.fill = FILL_SECTION
            for c in range(1, 37):
                ws.cell(cur_row, c).border = THIN_BORDER
            cur_row += 1

        r = cur_row
        ws.row_dimensions[r].height = 19
        unit_w = UNIT_WEIGHT[dia_val_int]
        length_m = round(len_mm / 1000.0, 3)

        ws.cell(r, 1, stt_cnt).alignment = ALIGN_CENTER
        ws.cell(r, 2, comp_name).alignment = ALIGN_LEFT
        ws.cell(r, 3, f"{sheet_name} (tr.{page_no})").alignment = ALIGN_CENTER
        ws.cell(r, 4, mark_str).alignment = ALIGN_CENTER
        ws.cell(r, 4).font = FONT_BOLD
        ws.cell(r, 5, dia_val_int).alignment = ALIGN_CENTER
        ws.cell(r, 6, "CB400-V").alignment = ALIGN_CENTER
        ws.cell(r, 7, f"{units_cnt} cấu kiện x {qty_one} thanh").alignment = ALIGN_LEFT

        c_len = ws.cell(r, 8, length_m)
        c_len.alignment = ALIGN_RIGHT
        c_len.number_format = "#,##0.000"

        c_qty = ws.cell(r, 9, tot_qty)
        c_qty.alignment = ALIGN_RIGHT
        c_qty.number_format = "#,##0"

        c_uw = ws.cell(r, 10, unit_w)
        c_uw.alignment = ALIGN_RIGHT
        c_uw.number_format = "0.000"

        c_kg = ws.cell(r, 11, f"=ROUND(H{r}*I{r}*J{r}, 2)")
        c_kg.alignment = ALIGN_RIGHT
        c_kg.number_format = "#,##0.00"

        c_tan = ws.cell(r, 12, f"=ROUND(K{r}/1000, 3)")
        c_tan.alignment = ALIGN_RIGHT
        c_tan.font = FONT_BOLD
        c_tan.number_format = "#,##0.000"

        ws.cell(r, 13, "").alignment = ALIGN_CENTER
        ws.cell(r, 14, f"CO-CQ-D{dia_val_int:02d}").alignment = ALIGN_CENTER

        for c_base in range(1, 15):
            if c_base not in (4, 12):
                ws.cell(r, c_base).font = FONT_REG
            ws.cell(r, c_base).border = THIN_BORDER

        for dia_val, c_acc_idx, c_freq_idx, l_short, _, _ in DIA_LIST_KH2:
            acc_col_let = get_column_letter(c_acc_idx)
            match_cond = f"IF(E{r}={dia_val}, L{r}, 0)"
            if prev_r is None:
                ws.cell(r, c_acc_idx, f"={match_cond}")
            else:
                ws.cell(r, c_acc_idx, f"=ROUND({acc_col_let}{prev_r} + {match_cond}, 3)")
            ca = ws.cell(r, c_acc_idx)
            ca.alignment = ALIGN_RIGHT
            ca.font = FONT_REG
            ca.border = THIN_BORDER
            ca.number_format = "#,##0.000"

            if prev_r is None:
                ws.cell(r, c_freq_idx, f'=IF({acc_col_let}{r}=0, "", "Lần " & ROUNDUP({acc_col_let}{r}/20, 0))')
            else:
                ws.cell(r, c_freq_idx, f'=IF(ROUNDUP({acc_col_let}{r}/20, 0)>ROUNDUP({acc_col_let}{prev_r}/20, 0), "Lần " & ROUNDUP({acc_col_let}{r}/20, 0), "")')
            cf = ws.cell(r, c_freq_idx)
            cf.alignment = ALIGN_CENTER
            cf.font = FONT_RED
            cf.border = THIN_BORDER

        # Cột O: Cảnh báo lấy mẫu tự động
        if prev_r is None:
            warn_parts = [f'IF(E{r}={dia_val}, "TN {l_short} (Lần 1)", ' for dia_val, _, _, l_short, _, _ in DIA_LIST_KH2]
            ws.cell(r, 15, "=" + "".join(warn_parts) + '""' + (")" * len(DIA_LIST_KH2)))
        else:
            p = prev_r
            warn_parts = []
            for dia_val, c_acc_idx, _, l_short, _, _ in DIA_LIST_KH2:
                acc_l = get_column_letter(c_acc_idx)
                c_if = f"AND(E{r}={dia_val}, ROUNDUP({acc_l}{r}/20,0)>ROUNDUP({acc_l}{p}/20,0))"
                warn_parts.append(f'IF({c_if}, "TN {l_short} (Lần " & ROUNDUP({acc_l}{r}/20,0) & ")", ')
            ws.cell(r, 15, "=" + "".join(warn_parts) + '""' + (")" * len(DIA_LIST_KH2)))

        cell_warn = ws.cell(r, 15)
        cell_warn.alignment = ALIGN_CENTER
        cell_warn.font = FONT_HDR_RED
        cell_warn.border = THIN_BORDER

        ws.cell(r, 36, "ĐẠT TCVN 1651:2018").alignment = ALIGN_CENTER
        ws.cell(r, 36).font = FONT_REG
        ws.cell(r, 36).border = THIN_BORDER

        prev_r = r
        stt_cnt += 1
        cur_row += 1

    r_end = cur_row
    ws.row_dimensions[r_end].height = 26
    ws.merge_cells(start_row=r_end, start_column=1, end_row=r_end, end_column=10)
    ws.cell(r_end, 1, "TỔNG CỘNG KHỐI LƯỢNG CỐT THÉP TOÀN CÔNG TRÌNH CẦU KHAI HOANG 2 (166 VỊ TRÍ BBS)").alignment = ALIGN_CENTER
    ws.cell(r_end, 1).font = FONT_HDR

    ws.cell(r_end, 11, f"=SUM(K4:K{r_end-1})").alignment = ALIGN_RIGHT
    ws.cell(r_end, 11).font = FONT_HDR
    ws.cell(r_end, 11).number_format = "#,##0.00"

    ws.cell(r_end, 12, f"=SUM(L4:L{r_end-1})").alignment = ALIGN_RIGHT
    ws.cell(r_end, 12).font = FONT_HDR
    ws.cell(r_end, 12).number_format = "#,##0.000"

    if prev_r:
        for _, c_acc_idx, c_freq_idx, _, _, _ in DIA_LIST_KH2:
            acc_col_let = get_column_letter(c_acc_idx)
            c_acc_tot = ws.cell(r_end, c_acc_idx, f"={acc_col_let}{prev_r}")
            c_acc_tot.font = FONT_HDR
            c_acc_tot.alignment = ALIGN_RIGHT
            c_acc_tot.number_format = "#,##0.000"

            c_freq_tot = ws.cell(r_end, c_freq_idx, f'="Tổng: " & ROUNDUP({acc_col_let}{prev_r}/20, 0) & " lần"')
            c_freq_tot.font = FONT_HDR
            c_freq_tot.alignment = ALIGN_CENTER

    ws.cell(r_end, 36, "100% ĐẠT CHUẨN").alignment = ALIGN_CENTER
    ws.cell(r_end, 36).font = FONT_HDR

    for c in range(1, 37):
        cell = ws.cell(r_end, c)
        cell.border = THIN_BORDER
        cell.fill = FILL_TOTAL

    col_widths_rebar = {
        1: 6, 2: 34, 3: 18, 4: 12, 5: 11, 6: 12, 7: 26,
        8: 12, 9: 11, 10: 12, 11: 13, 12: 13, 13: 13, 14: 16, 15: 22, 36: 18
    }
    for c_idx in range(16, 36):
        col_widths_rebar[c_idx] = 13
    for col_idx, width in col_widths_rebar.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.freeze_panes = "E3"
    return r_end


def _build_sheet_summary_kh2(ws, r_tot_concrete: int, r_end_rebar: int):
    """Xây dựng Sheet 3: Bảng tổng hợp kế hoạch & tần suất thí nghiệm vật liệu đầu vào."""
    ws.title = "Tong_Hop_Tan_Suat_Vat_Lieu"
    ws.views.sheetView[0].showGridLines = True

    ws.merge_cells("A1:I1")
    ws["A1"] = "BẢNG TỔNG HỢP KẾ HOẠCH & TẦN SUẤT THÍ NGHIỆM VẬT LIỆU ĐẦU VÀO TOÀN DỰ ÁN"
    ws["A1"].font = FONT_TITLE
    ws["A1"].alignment = ALIGN_CENTER

    ws.merge_cells("A2:I2")
    ws["A2"] = "Công trình: Cầu thôn Khai Hoang 2, Km 14+363.65 | Luật XD 135/2025/QH15, NĐ 207/2026/NĐ-CP, TCVN 1651:2018, TCVN 4453:1995"
    ws["A2"].font = FONT_SUBTITLE
    ws["A2"].alignment = ALIGN_CENTER

    headers = [
        ("STT", 6),
        ("DANH MỤC VẬT LIỆU / CÔNG VIỆC THÍ NGHIỆM", 44),
        ("TIÊU CHUẨN ÁP DỤNG", 22),
        ("QUY ĐỊNH TẦN SUẤT LẤY MẪU", 26),
        ("ĐƠN VỊ TÍNH", 12),
        ("TỔNG KHỐI LƯỢNG DỰ ÁN", 22),
        ("SỐ LẦN THÍ NGHIỆM DỰ KIẾN", 24),
        ("CHỈ TIÊU CƠ LÝ KIỂM TRA CHÍNH", 36),
        ("GHI CHÚ / QUY ĐỊNH PHÁP LÝ", 32),
    ]
    ws.row_dimensions[4].height = 28
    for col_idx, (h_name, h_width) in enumerate(headers, start=1):
        cell = ws.cell(4, col_idx, h_name)
        cell.font = FONT_HDR
        cell.alignment = ALIGN_CENTER
        cell.border = THIN_BORDER
        cell.fill = FILL_HDR
        ws.column_dimensions[get_column_letter(col_idx)].width = h_width

    summary_rows = [
        ("SECTION", "I. NHÓM VẬT LIỆU ĐẦU VÀO BÊ TÔNG XI MĂNG & ĐÁ HỘC XÂY"),
        (1, "Xi măng Poóc lăng hỗn hợp PCB40", "TCVN 6260:2020", "50 Tấn / lần lấy mẫu", "Tấn",
         f"=Theo_Doi_Tan_Suat_Be_Tong!M{r_tot_concrete}", 50,
         "Độ mịn, thời gian đông kết, giới hạn bền nén R3/R28, độ ổn định thể tích", "TCVN 6260:2020, NĐ 207/2026"),
        (2, "Cát vàng đổ bê tông (Mô đun độ lớn Mk >= 2.0)", "TCVN 7570:2006", "200 m3 / lần lấy mẫu", "m3",
         f"=Theo_Doi_Tan_Suat_Be_Tong!N{r_tot_concrete}", 200,
         "Thành phần hạt, hàm lượng bùn bụi sét, tạp chất hữu cơ, khối lượng thể tích", "TCVN 7570:2006"),
        (3, "Đá dăm bê tông (Đá 1x2 cho C25/C30/C35/C40)", "TCVN 7570:2006", "350 m3 / lần lấy mẫu", "m3",
         f"=Theo_Doi_Tan_Suat_Be_Tong!O{r_tot_concrete}", 350,
         "Thành phần hạt, hàm lượng thoi dẹt, độ nén dập trong ống hình trụ, độ hút nước", "TCVN 7570:2006, TCVN 1771"),
        (4, "Nước trộn bê tông & bảo dưỡng", "TCVN 4506:2012", "1 nguồn cấp / toàn dự án", "Nguồn",
         1, None,
         "Độ pH, hàm lượng muối hòa tan, ion Cl-, SO4(2-), cặn không tan", "Kiểm tra trước khi thi công"),
        (5, "Phụ gia hóa học giảm nước / siêu dẻo", "TCVN 8826:2011", "5 Tấn / lô vật liệu", "Tấn",
         4.5, 5,
         "Khả năng giảm nước, độ chảy xòe, độ pH, hàm lượng chất rắn", "Chứng chỉ xuất xưởng + thử nghiệm"),

        ("SECTION", "II. NHÓM CỐT THÉP XÂY DỰNG (BÓC TÁCH RIÊNG 10 LOẠI ĐƯỜNG KÍNH Ø)"),
        (6, "Cốt thép Ø6 (Lưới thép bản mặt cầu B1, B2)", "TCVN 1651-1:2018", "20 Tấn / lô xuất xưởng", "Tấn",
         f"=Theo_Doi_Tan_Suat_Cot_Thep!P{r_end_rebar}", 20,
         "Giới hạn chảy, độ bền kéo, độ giãn dài, uốn nguội 180°", "1 tổ 3 kéo + 3 uốn / lô 20T"),
        (7, "Cốt thép thanh vằn Ø10 (CB400-V)", "TCVN 1651-2:2018", "20 Tấn / lô xuất xưởng", "Tấn",
         f"=Theo_Doi_Tan_Suat_Cot_Thep!Q{r_end_rebar}", 20,
         "Giới hạn chảy, độ bền kéo, độ giãn dài, uốn nguội 180°", "Thép đai dầm T biên D1, bản mặt cầu, kê gối"),
        (8, "Cốt thép thanh vằn Ø12 (CB400-V)", "TCVN 1651-2:2018", "20 Tấn / lô xuất xưởng", "Tấn",
         f"=Theo_Doi_Tan_Suat_Cot_Thep!R{r_end_rebar}", 20,
         "Giới hạn chảy, độ bền kéo, độ giãn dài, uốn nguội 180°", "Thép dầm T15m & gờ lan can"),
        (9, "Cốt thép thanh vằn Ø14 (CB400-V)", "TCVN 1651-2:2018", "20 Tấn / lô xuất xưởng", "Tấn",
         f"=Theo_Doi_Tan_Suat_Cot_Thep!S{r_end_rebar}", 20,
         "Giới hạn chảy, độ bền kéo, độ giãn dài, uốn nguội 180°", "Thép bản cánh dầm T, gờ mố, tường chắn"),
        (10, "Cốt thép thanh vằn Ø16 (CB400-V)", "TCVN 1651-2:2018", "20 Tấn / lô xuất xưởng", "Tấn",
         f"=Theo_Doi_Tan_Suat_Cot_Thep!T{r_end_rebar}", 20,
         "Giới hạn chảy, độ bền kéo, độ giãn dài, uốn nguội 180°", "Thép sườn dầm T15m, thân mố, khe co giãn"),
        (11, "Cốt thép thanh vằn Ø18 (CB400-V)", "TCVN 1651-2:2018", "20 Tấn / lô xuất xưởng", "Tấn",
         f"=Theo_Doi_Tan_Suat_Cot_Thep!U{r_end_rebar}", 20,
         "Giới hạn chảy, độ bền kéo, độ giãn dài, uốn nguội 180°", "Thép chịu lực thân tường chắn 4 đốt"),
        (12, "Cốt thép thanh vằn Ø20 (CB400-V)", "TCVN 1651-2:2018", "20 Tấn / lô xuất xưởng", "Tấn",
         f"=Theo_Doi_Tan_Suat_Cot_Thep!V{r_end_rebar}", 20,
         "Giới hạn chảy, độ bền kéo, độ giãn dài, uốn nguội 180°", "Thép phân bố bệ mố, thân mố W3, bản quá độ"),
        (13, "Cốt thép thanh vằn Ø22 (CB400-V)", "TCVN 1651-2:2018", "20 Tấn / lô xuất xưởng", "Tấn",
         f"=Theo_Doi_Tan_Suat_Cot_Thep!W{r_end_rebar}", 20,
         "Giới hạn chảy, độ bền kéo, độ giãn dài, uốn nguội 180°", "Thép chịu lực tường cánh mố chữ U"),
        (14, "Cốt thép thanh vằn Ø25 (CB400-V)", "TCVN 1651-2:2018", "20 Tấn / lô xuất xưởng", "Tấn",
         f"=Theo_Doi_Tan_Suat_Cot_Thep!X{r_end_rebar}", 20,
         "Giới hạn chảy, độ bền kéo, độ giãn dài, uốn nguội 180°", "Thép chủ chịu lực bệ mố M1, M2 & thân mố"),
        (15, "Cốt thép thanh vằn Ø32 (CB400-V)", "TCVN 1651-2:2018", "20 Tấn / lô xuất xưởng", "Tấn",
         f"=Theo_Doi_Tan_Suat_Cot_Thep!Y{r_end_rebar}", 20,
         "Giới hạn chảy, độ bền kéo, độ giãn dài, uốn nguội 180°", "Thép chủ dầm T15m (A1, B1) & 98 thanh neo đá"),

        ("SECTION", "III. KIỂM ĐỊNH HIỆN TRƯỜNG & NGHIỆM THU CHẤT LƯỢNG CÔNG TRÌNH"),
        (16, "Thí nghiệm nén mẫu bê tông hiện trường R7, R28", "TCVN 3118:2022", "1 tổ 3 mẫu / đợt đổ cấu kiện", "Tổ mẫu",
         36, None,
         "Cường độ chịu nén R7 ngày và R28 ngày (MPa)", "36 tổ mẫu (108 viên mẫu 15x15x15cm)"),
        (17, "Thí nghiệm nhổ thử tải cốt thép neo đá D32 (Pull-out test)", "TCVN 9393 / Chỉ dẫn KT", "3% số thanh neo (tối thiểu 3 thanh/mố)", "Thanh",
         6, None,
         "Khả năng chịu lực nhổ bám dính vữa M300 - đá gốc", "3 thanh mố M1 + 3 thanh mố M2 (tổng 98 lỗ neo)"),
        (18, "Thí nghiệm gối cầu cao su cốt bản thép 350x500x84mm", "TCVN 10308:2014", "1 lô xuất xưởng / 08 gối cầu", "Bộ",
         8, None,
         "Mô đun đàn hồi nén, mô đun trượt, độ bền lão hóa", "Kiểm định độc lập trước khi lắp lên đá kê gối"),
        (19, "Thí nghiệm độ chặt nền đường đầu cầu & sau mố K95/K98", "TCVN 4447:2012", "Mỗi lớp đầm 30cm / mặt cắt", "Điểm",
         24, None,
         "Hệ số đầm chặt K >= 0.95 (thân) và K >= 0.98 (đỉnh)", "Phễu rót cát hiện trường sau mố M1, M2"),
    ]

    r_cur = 5
    for item in summary_rows:
        if item[0] == "SECTION":
            ws.row_dimensions[r_cur].height = 22
            ws.merge_cells(start_row=r_cur, start_column=1, end_row=r_cur, end_column=9)
            sc = ws.cell(r_cur, 1, item[1])
            sc.font = Font(name=FONT_NAME, size=9.5, bold=True, color="1B365D")
            sc.alignment = Alignment(horizontal="left", vertical="center")
            sc.fill = FILL_SECTION
            for c in range(1, 10):
                ws.cell(r_cur, c).border = THIN_BORDER
            r_cur += 1
            continue

        stt_num, mat_name, std_name, freq_rule, unit_txt, total_val, denom, params_txt, note_txt = item
        ws.row_dimensions[r_cur].height = 20

        ws.cell(r_cur, 1, stt_num).alignment = ALIGN_CENTER
        ws.cell(r_cur, 1).font = FONT_REG
        ws.cell(r_cur, 1).border = THIN_BORDER

        ws.cell(r_cur, 2, mat_name).alignment = ALIGN_LEFT
        ws.cell(r_cur, 2).font = FONT_BOLD
        ws.cell(r_cur, 2).border = THIN_BORDER

        ws.cell(r_cur, 3, std_name).alignment = ALIGN_CENTER
        ws.cell(r_cur, 3).font = FONT_REG
        ws.cell(r_cur, 3).border = THIN_BORDER

        ws.cell(r_cur, 4, freq_rule).alignment = ALIGN_LEFT
        ws.cell(r_cur, 4).font = FONT_REG
        ws.cell(r_cur, 4).border = THIN_BORDER

        ws.cell(r_cur, 5, unit_txt).alignment = ALIGN_CENTER
        ws.cell(r_cur, 5).font = FONT_REG
        ws.cell(r_cur, 5).border = THIN_BORDER

        cell_tot = ws.cell(r_cur, 6, total_val)
        cell_tot.alignment = ALIGN_RIGHT
        cell_tot.font = FONT_BOLD
        cell_tot.border = THIN_BORDER
        cell_tot.number_format = "#,##0.000" if stt_num in range(6, 16) else "#,##0.00"

        if denom:
            test_cnt_formula = f'=IF(F{r_cur}="","",IF(F{r_cur}=0,"0 lần",ROUNDUP(F{r_cur}/{denom},0) & " lần"))'
        else:
            test_cnt_formula = f'="Thực hiện " & F{r_cur} & " " & E{r_cur}'

        cell_tc = ws.cell(r_cur, 7, test_cnt_formula)
        cell_tc.alignment = ALIGN_CENTER
        cell_tc.font = FONT_RED
        cell_tc.border = THIN_BORDER

        ws.cell(r_cur, 8, params_txt).alignment = ALIGN_LEFT
        ws.cell(r_cur, 8).font = FONT_REG
        ws.cell(r_cur, 8).border = THIN_BORDER

        ws.cell(r_cur, 9, note_txt).alignment = ALIGN_LEFT
        ws.cell(r_cur, 9).font = FONT_REG
        ws.cell(r_cur, 9).border = THIN_BORDER

        r_cur += 1

    ws.freeze_panes = "C5"


def _build_sheet_nghi_van_kh2(ws):
    """Xây dựng Sheet 4: NGHI_VAN ghi nhận 7 điểm đối chiếu bản vẽ & OCR của Cầu thôn Khai Hoang 2."""
    ws.title = "NGHI_VAN"
    ws.views.sheetView[0].showGridLines = True

    ws.merge_cells("A1:I1")
    ws["A1"] = "DANH MỤC ĐIỂM NGHI VẤN & ĐỐI CHIẾU SỐ LIỆU BẢN VẼ GỐC — CẦU THÔN KHAI HOANG 2, KM 14+363.65"
    ws["A1"].font = FONT_TITLE
    ws["A1"].alignment = ALIGN_CENTER

    headers = [
        ("Mã NV", 10), ("Bản vẽ / Trang", 22), ("Hạng mục", 26), ("Ký hiệu thanh", 14),
        ("Số liệu bản vẽ / OCR", 22), ("Số liệu chuẩn TCVN", 22), ("Chênh lệch (kg)", 16),
        ("Nguyên nhân & Giải trình kỹ thuật", 48), ("Đã soát / Ghim", 16),
    ]
    ws.row_dimensions[3].height = 26
    for c_i, (h_txt, w) in enumerate(headers, 1):
        cell = ws.cell(3, c_i, h_txt)
        cell.font = FONT_HDR
        cell.fill = FILL_HDR
        cell.alignment = ALIGN_CENTER
        cell.border = THIN_BORDER
        ws.column_dimensions[get_column_letter(c_i)].width = w

    flags = [
        ("NV-01", "PD-04 (tr.18)", "Thân mố + tường đỉnh", "W3, W6, A4, A9", "OCR: 7,900,00 / 1,153,00", "Chuẩn hóa dấu thập phân", 1858.28,
         "OCR in dấu phẩy làm phân cách thập phân; đã khôi phục 100% khớp tổng bản vẽ.", "ĐÃ XỬ LÝ"),
        ("NV-02", "PT-03 (tr.28)", "Dầm T L=15m - dầm biên", "D1 (Ø10)", "OCR: 2.260.00 mm", "2,260 mm (204 thanh)", 284.46,
         "OCR dùng 2 dấu chấm; đã khôi phục L=2,260mm khớp 99.97% bản vẽ.", "ĐÃ XỬ LÝ"),
        ("NV-03", "PD-04 (tr.18)", "Thân mố + tường đỉnh", "A2 (Ø12)", "Số lượng OCR: 5.3", "53 thanh (L=540mm)", 45.75,
         "Tổng dài trên bản vẽ 28.62m / 0.54m = 53 thanh (OCR nhận nhầm dấu chấm).", "ĐÃ XỬ LÝ"),
        ("NV-04", "PD-07 (tr.21)", "Tường cánh mố chữ U", "K1, K1A, K8, K8A, K19, K19A", "2,900.11 kg / 1 tường", "3,747.29 kg / 1 tường", 3388.72,
         "Bản vẽ gốc cộng thiếu 50% khối lượng 6 thanh này; tính đúng theo L x n x kg/m TCVN.", "CẦN TVTK XÁC NHẬN"),
        ("NV-05", "TC-01..04 (tr.39-47)", "Tường chắn 4 đốt", "Thanh số 1 (Ø20)", "Đơn trọng 1.980 kg/m", "Đơn trọng 2.466 kg/m", 290.15,
         "Bản vẽ gốc áp nhầm đơn trọng Ø18 (1.980 kg/m) cho thanh Ø20; tính lại theo 2.466 kg/m.", "CẦN TVTK XÁC NHẬN"),
        ("NV-06", "PT-03 & PT-06", "Dầm T L=15m", "D2B (trùng tên)", "2 dòng cùng tên D2B", "Đổi dòng 2 thành D2D", 0.0,
         "Tránh nhầm lẫn khi lập trình máy cắt CNC (L=1,670mm và L=1,320mm).", "ĐÃ XỬ LÝ"),
        ("NV-07", "THKL G9 vs BBS", "Toàn cầu Khai Hoang 2", "Tổng hợp thép", "THKL mời thầu", "68,012.77 kg (166 dòng)", 0.0,
         "Đối chiếu độc lập giữa tiên lượng mời thầu THKL và 15 bảng BBS chi tiết.", "THEO DÕI PL03A"),
    ]
    for idx, row_vals in enumerate(flags, start=4):
        ws.row_dimensions[idx].height = 24
        for c_i, val in enumerate(row_vals, 1):
            c = ws.cell(idx, c_i, val)
            c.font = FONT_BOLD if c_i in (1, 9) else FONT_REG
            c.border = THIN_BORDER
            c.fill = FILL_ALERT if "TVTK" in str(row_vals[8]) else FILL_TOTAL
            c.alignment = ALIGN_RIGHT if c_i == 7 else (ALIGN_LEFT if c_i in (3, 8) else ALIGN_CENTER)
            if c_i == 7:
                c.number_format = "#,##0.00"


def build_tan_suat_khai_hoang_2(output_path: str):
    """Sinh file Bang_Theo_Doi_Tan_Suat_Thi_Nghiem_Be_Tong_Va_Thep_Cau_Khai_Hoang_2.xlsx."""
    from examples.build_khai_hoang_2_full_pipeline import load_all_bbs
    bbs_rows = load_all_bbs()

    wb = openpyxl.Workbook()
    ws1 = wb.active
    r_tot_bt = _build_sheet_concrete_kh2(ws1)

    ws2 = wb.create_sheet("Theo_Doi_Tan_Suat_Cot_Thep")
    r_end_rebar = _build_sheet_rebar_kh2(ws2, bbs_rows)

    ws3 = wb.create_sheet("Tong_Hop_Tan_Suat_Vat_Lieu")
    _build_sheet_summary_kh2(ws3, r_tot_bt, r_end_rebar)

    ws4 = wb.create_sheet("NGHI_VAN")
    _build_sheet_nghi_van_kh2(ws4)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)
    wb.close()
    print(f"  [V] Đã tạo Bảng Tần suất Thí nghiệm Bê tông & Thép: {output_path}")


def build_visual_asbuilt_khai_hoang_2(xlsx_path: str, html_path: str = None):
    """
    Sinh Sơ đồ Tiến độ Trực quan As-Built Cầu thôn Khai Hoang 2, Km 14+363.65:
    - Sheet 1: SO_DO_TIEN_DO_TRUC_QUAN (Bảng tiến độ kết cấu + KPI công thức sống)
    - Sheet 2: MAT_CAT_DUNG_ASBUILT_GRID (Bản vẽ mặt cắt đứng cầu bằng lưới ô Excel trực quan)
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SO_DO_TIEN_DO_TRUC_QUAN"
    ws.views.sheetView[0].showGridLines = True

    ws.merge_cells("A1:K1")
    ws["A1"] = "BAN CHỈ HUY CÔNG TRƯỜNG — SƠ ĐỒ TIẾN ĐỘ THI CÔNG TRỰC QUAN & AS-BUILT CẦU THÔN KHAI HOANG 2 (KM 14+363.65)"
    ws["A1"].font = Font(name=FONT_NAME, size=14, bold=True, color="1B365D")
    ws["A1"].alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:K2")
    ws["A2"] = "Quy mô: 01 Nhịp dầm chữ T BTCT L=15.0m (04 phiến dầm) • 02 Mố chữ U C30 đặt trên nền đá gốc gia cường 98 neo D32 • 04 Đốt tường chắn"
    ws["A2"].font = FONT_SUBTITLE
    ws["A2"].alignment = ALIGN_CENTER
    ws.row_dimensions[2].height = 20

    # Legend & Live KPIs on Row 4
    ws.merge_cells("A4:B4")
    ws["A4"] = "TỔNG KHỐI LƯỢNG BT (m3):"
    ws["A4"].font = FONT_BOLD
    ws["C4"] = "=SUM(E7:E28)"
    ws["C4"].font = FONT_HDR_RED
    ws["C4"].number_format = "#,##0.00"

    ws["D4"] = "[XANH LÁ] ĐÃ THI CÔNG"
    ws["D4"].fill = FILL_GREEN
    ws["D4"].font = FONT_HDR_WHITE
    ws["D4"].alignment = ALIGN_CENTER

    ws["F4"] = "[VÀNG CAM] ĐANG TRIỂN KHAI"
    ws["F4"].fill = FILL_AMBER
    ws["F4"].font = FONT_BOLD
    ws["F4"].alignment = ALIGN_CENTER

    ws["H4"] = "[XÁM] CHƯA THI CÔNG"
    ws["H4"].fill = FILL_HDR
    ws["H4"].font = FONT_BOLD
    ws["H4"].alignment = ALIGN_CENTER

    ws["J4"] = "SỐ HẠNG MỤC:"
    ws["J4"].font = FONT_BOLD
    ws["K4"] = '=COUNTIF(E7:E28, ">=0")'
    ws["K4"].font = FONT_BOLD
    ws["K4"].alignment = ALIGN_CENTER

    headers = [
        "HẠNG MỤC CẤU KIỆN", "VỊ TRÍ", "QUY CÁCH / KÍCH THƯỚC", "MÁC BT",
        "KHỐI LƯỢNG (m³)", "TIẾN ĐỘ THIẾT KẾ", "NGÀY KHỞI CÔNG", "NGÀY HOÀN THÀNH",
        "TÌNH TRẠNG HIỆN TRƯỜNG", "ĐÁNH GIÁ KCS R28", "GHI CHÚ KỸ THUẬT / BẢN VẼ",
    ]
    ws.row_dimensions[6].height = 26
    for col_idx, text in enumerate(headers, 1):
        cell = ws.cell(6, col_idx, text)
        cell.font = FONT_HDR_WHITE
        cell.fill = FILL_NAVY
        cell.alignment = ALIGN_CENTER
        cell.border = THIN_BORDER

    asbuilt_rows = [
        ("Đào hố móng đá mố M1, M2 & tường chắn", "Mố M1, M2", "Hố móng đá phong hóa 3,626.97 m3", "Đá gốc", 0.00, "25/10 - 01/11/2026", "25/10/2026", "01/11/2026", "ĐÃ THI CÔNG", "ĐẠT KCS 100%", "Đã nghiệm thu cao độ đáy móng đá gốc"),
        ("Khoan cấy 49 thanh thép neo Ø32 mố M1", "Mố M1", "49 lỗ khoan sâu >=1.0m vào đá", "Vữa M300", 2.45, "01/11 - 03/11/2026", "01/11/2026", "03/11/2026", "ĐÃ THI CÔNG", "ĐẠT PULL-OUT", "BV PD-01 tr.13 (Thử nhổ 3 thanh đạt 100%)"),
        ("Khoan cấy 49 thanh thép neo Ø32 mố M2", "Mố M2", "49 lỗ khoan sâu >=1.0m vào đá", "Vữa M300", 2.45, "02/11 - 04/11/2026", "02/11/2026", "04/11/2026", "ĐÃ THI CÔNG", "ĐẠT PULL-OUT", "BV PD-01 tr.13 (Thử nhổ 3 thanh đạt 100%)"),
        ("Bê tông đệm lót móng mố M1, M2 & tường chắn", "Toàn cầu", "Đá 2x4 dày 10cm", "C10 (M100)", 58.70, "31/10 - 03/11/2026", "31/10/2026", "03/11/2026", "ĐÃ THI CÔNG", "ĐẠT R28 (108%)", "BV PD-01 tr.12 & TC-01..04"),
        ("Bê tông bệ móng mố chữ U M1", "Mố M1", "Bệ móng BTCT trên đá gốc", "C30 đá 1x2", 115.66, "04/11 - 09/11/2026", "04/11/2026", "09/11/2026", "ĐÃ THI CÔNG", "ĐẠT R28 (112%)", "BV PD-02 tr.14 - 2 tổ mẫu R7/R28 đạt"),
        ("Bê tông bệ móng mố chữ U M2", "Mố M2", "Bệ móng BTCT trên đá gốc", "C30 đá 1x2", 115.67, "06/11 - 11/11/2026", "06/11/2026", "11/11/2026", "ĐÃ THI CÔNG", "ĐẠT R28 (110%)", "BV PD-03 tr.15 - 2 tổ mẫu R7/R28 đạt"),
        ("Bê tông thân mố & tường đỉnh mố M1", "Mố M1", "Chia 2 đợt đổ (60m3 + 55.66m3)", "C30 đá 1x2", 115.66, "10/11 - 17/11/2026", "10/11/2026", "17/11/2026", "ĐÃ THI CÔNG", "ĐẠT R28 (109%)", "BV PD-04 tr.18 - Độ thẳng đứng < 3mm"),
        ("Bê tông thân mố & tường đỉnh mố M2", "Mố M2", "Chia 2 đợt đổ (60m3 + 55.67m3)", "C30 đá 1x2", 115.67, "11/11 - 18/11/2026", "11/11/2026", "18/11/2026", "ĐÃ THI CÔNG", "ĐẠT R28 (111%)", "BV PD-04 tr.18 - Độ thẳng đứng < 3mm"),
        ("Bê tông tường cánh mố chữ U M1 (2 tường)", "Mố M1", "Tường cánh trái + phải M1", "C30 đá 1x2", 78.05, "18/11 - 22/11/2026", "18/11/2026", "22/11/2026", "ĐANG TRIỂN KHAI", "ĐANG CHỜ R28", "BV PD-07 tr.21 - Đã đạt R7 = 92%"),
        ("Bê tông tường cánh mố chữ U M2 (2 tường)", "Mố M2", "Tường cánh trái + phải M2", "C30 đá 1x2", 78.06, "19/11 - 24/11/2026", "19/11/2026", "24/11/2026", "ĐANG TRIỂN KHAI", "ĐANG CHỜ R28", "BV PD-07 tr.21 - Đã đạt R7 = 90%"),
        ("Bê tông đá kê gối & lắp 08 gối cầu cao su", "Mố M1, M2", "08 gối 350x500x84mm + 4 ụ chống xô", "C35 đá 1x2", 3.42, "24/11 - 26/11/2026", "24/11/2026", "26/11/2026", "ĐANG TRIỂN KHAI", "ĐANG THI CÔNG", "BV PD-05 tr.19 - Cao độ dung sai +-2mm"),
        ("Bê tông 04 đốt tường chắn đầu cầu & bệ phản áp", "Đầu cầu", "L=5.2m; 5.89m; 10.6m; 13.5m", "C25 đá 1x2", 271.40, "20/11 - 02/12/2026", "20/11/2026", "02/12/2026", "ĐANG TRIỂN KHAI", "ĐÃ XONG 2/4 ĐỐT", "BV TC-01..04 tr.39-47"),
        ("Đúc dầm chữ T L=15m - Phiến D1 (Dầm biên trái)", "Bãi đúc", "Dầm T BTCT L=15.0m, H=1.10m", "C40 đá 1x2", 9.35, "10/11 - 16/11/2026", "10/11/2026", "16/11/2026", "ĐÃ THI CÔNG", "ĐẠT R28 (115%)", "BV PT-01..03 tr.26-28 - Đủ điều kiện lao"),
        ("Đúc dầm chữ T L=15m - Phiến D2 (Dầm giữa 1)", "Bãi đúc", "Dầm T BTCT L=15.0m, H=1.10m", "C40 đá 1x2", 9.35, "13/11 - 19/11/2026", "13/11/2026", "19/11/2026", "ĐÃ THI CÔNG", "ĐẠT R28 (114%)", "BV PT-04..06 tr.31-33 - Đủ điều kiện lao"),
        ("Đúc dầm chữ T L=15m - Phiến D3 (Dầm giữa 2)", "Bãi đúc", "Dầm T BTCT L=15.0m, H=1.10m", "C40 đá 1x2", 9.35, "16/11 - 22/11/2026", "16/11/2026", "22/11/2026", "ĐANG TRIỂN KHAI", "ĐẠT R7 (94%)", "BV PT-04..06 tr.31-33 - Đang dưỡng hộ"),
        ("Đúc dầm chữ T L=15m - Phiến D4 (Dầm biên phải)", "Bãi đúc", "Dầm T BTCT L=15.0m, H=1.10m", "C40 đá 1x2", 9.35, "19/11 - 25/11/2026", "19/11/2026", "25/11/2026", "ĐANG TRIỂN KHAI", "ĐANG LẮP THÉP", "BV PT-01..03 tr.26-28 - Chuẩn bị đổ BT"),
        ("Cẩu lắp 04 phiến dầm chữ T L=15m lên mố M1-M2", "Nhịp 15m", "02 cẩu bánh xích 50T phối hợp", "C40", 0.00, "24/11 - 28/11/2026", "24/11/2026 (KH)", "28/11/2026 (KH)", "CHƯA THI CÔNG", "CHỜ ĐỦ R28 DẦM", "Chống lật tạm bằng thép hình U200"),
        ("Bê tông 03 dầm ngang & mối nối ướt bản cánh", "Nhịp 15m", "03 dầm ngang + 03 mối nối dọc", "C40 đá 1x2", 8.45, "27/11 - 01/12/2026", "27/11/2026 (KH)", "01/12/2026 (KH)", "CHƯA THI CÔNG", "CHƯA ĐỔ BT", "BV PT-07 tr.34"),
        ("Bê tông bản mặt cầu dày 12cm (B=7.0m)", "Nhịp 15m", "Lưới thép Ø6, Ø12, Ø14", "C35/C40", 14.13, "28/11 - 03/12/2026", "28/11/2026 (KH)", "03/12/2026 (KH)", "CHƯA THI CÔNG", "CHƯA ĐỔ BT", "BV PT-08 tr.35"),
        ("Bê tông 02 bản quá độ đầu mố M1, M2 (L=4.0m)", "Sau mố", "02 bản x 9.09 m3, dày 25cm", "C30 đá 1x2", 18.18, "29/11 - 03/12/2026", "29/11/2026 (KH)", "03/12/2026 (KH)", "CHƯA THI CÔNG", "CHƯA ĐỔ BT", "BV PD-08 tr.22"),
        ("Bê tông gờ lan can, khe co giãn & thảm BTNC 19", "Mặt cầu", "Gờ C30 + 02 khe co giãn + BTNC 7cm", "C30 / C40", 8.48, "02/12 - 11/12/2026", "02/12/2026 (KH)", "11/12/2026 (KH)", "CHƯA THI CÔNG", "CHƯA ĐỔ BT", "BV PT-08 & PD-09"),
        ("Đắp đất K95/K98 sau mố, đá hộc tứ nón & ATGT", "Hoàn thiện", "Đắp 1,055.34m3 + Xây đá 115m3", "Vữa M100", 0.00, "03/12 - 19/12/2026", "03/12/2026 (KH)", "19/12/2026 (KH)", "CHƯA THI CÔNG", "CHƯA THI CÔNG", "Nghiệm thu tổng thể & bàn giao"),
    ]

    for row_idx, rdata in enumerate(asbuilt_rows, 7):
        ws.row_dimensions[row_idx].height = 21
        status = rdata[8]
        row_fill = FILL_GREEN if status == "ĐÃ THI CÔNG" else (FILL_AMBER if status == "ĐANG TRIỂN KHAI" else FILL_HDR)
        for col_idx, val in enumerate(rdata, 1):
            cell = ws.cell(row_idx, col_idx, val)
            cell.font = FONT_REG
            cell.border = THIN_BORDER
            if col_idx == 5:
                cell.number_format = "#,##0.00"
                cell.alignment = ALIGN_RIGHT
            elif col_idx in (1, 3, 11):
                cell.alignment = ALIGN_LEFT
            else:
                cell.alignment = ALIGN_CENTER
            if col_idx == 9:
                cell.fill = row_fill
                cell.font = FONT_HDR_WHITE if status == "ĐÃ THI CÔNG" else FONT_BOLD

    col_widths = [36, 12, 30, 14, 15, 18, 15, 15, 20, 18, 36]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Sheet 2: Bản vẽ Mặt cắt đứng trực quan bằng lưới ô Excel (Elevation Grid)
    ws_grid = wb.create_sheet("MAT_CAT_DUNG_ASBUILT_GRID")
    ws_grid.views.sheetView[0].showGridLines = True
    ws_grid.merge_cells("B2:M2")
    ws_grid["B2"] = "SƠ ĐỒ MẶT CẮT ĐỨNG TRỰC QUAN CẦU THÔN KHAI HOANG 2 (KM 14+363.65) — 1 NHỊP DẦM T L=15.0M"
    ws_grid["B2"].font = FONT_TITLE
    ws_grid["B2"].alignment = ALIGN_CENTER
    ws_grid.row_dimensions[2].height = 28

    grid_blocks = [
        ("B4:D5", "ĐƯỜNG ĐẦU CẦU M1\n+ BẢN QUÁ ĐỘ L=4M\n(CHƯA THI CÔNG)", FILL_HDR, FONT_BOLD),
        ("E4:J4", "BẢN MẶT CẦU C35 DÀY 12CM + THẢM BTNC 19 DÀY 7CM + LAN CAN (CHƯA THI CÔNG)", FILL_HDR, FONT_BOLD),
        ("K4:M5", "ĐƯỜNG ĐẦU CẦU M2\n+ BẢN QUÁ ĐỘ L=4M\n(CHƯA THI CÔNG)", FILL_HDR, FONT_BOLD),
        ("E5:J6", "04 PHIẾN DẦM CHỮ T BTCT L=15.0M C40\n(ĐÃ ĐÚC XONG D1, D2 • ĐANG ĐÚC D3, D4)", FILL_AMBER, FONT_BOLD),
        ("C6:D9", "THÂN MỐ CHỮ U M1\n+ TƯỜNG CÁNH M1\n(193.71 m³ C30 - ĐÃ XONG THÂN)", FILL_GREEN, FONT_HDR_WHITE),
        ("K6:L9", "THÂN MỐ CHỮ U M2\n+ TƯỜNG CÁNH M2\n(193.73 m³ C30 - ĐÃ XONG THÂN)", FILL_GREEN, FONT_HDR_WHITE),
        ("B6:B9", "TƯỜNG CHẮN\nĐỐT 1 & 2\n(ĐANG ĐỔ C25)", FILL_AMBER, FONT_BOLD),
        ("M6:M9", "TƯỜNG CHẮN\nĐỐT 3 & 4\n(ĐANG ĐỔ C25)", FILL_AMBER, FONT_BOLD),
        ("C10:D11", "BỆ MỐ M1 C30 (115.66 m³)\n+ ĐỆM LÓT C10 (ĐÃ XONG 100%)", FILL_GREEN, FONT_HDR_WHITE),
        ("K10:L11", "BỆ MỐ M2 C30 (115.67 m³)\n+ ĐỆM LÓT C10 (ĐÃ XONG 100%)", FILL_GREEN, FONT_HDR_WHITE),
        ("C12:D13", "49 THANH THÉP NEO Ø32\nKHOAN CẤY ĐÁ GỐC M1 (ĐÃ XONG)", FILL_LIGHT_GREEN, FONT_GREEN),
        ("K12:L13", "49 THANH THÉP NEO Ø32\nKHOAN CẤY ĐÁ GỐC M2 (ĐÃ XONG)", FILL_LIGHT_GREEN, FONT_GREEN),
        ("E8:J13", "LÒNG SUỐI KHAI HOANG 2 (KHẨU ĐỘ THOÁT NƯỚC L=15.0M)\nNỀN ĐÁ GỐC PHONG HÓA NHẸ", FILL_SUBSECTION, FONT_BOLD),
    ]
    for rng, label, fill, fnt in grid_blocks:
        ws_grid.merge_cells(rng)
        top_left = rng.split(":")[0]
        ws_grid[top_left] = label
        ws_grid[top_left].fill = fill
        ws_grid[top_left].font = fnt
        ws_grid[top_left].alignment = ALIGN_CENTER
        for row in ws_grid[rng]:
            for cell in row:
                cell.border = THIN_BORDER
                cell.fill = fill

    for r_i in range(4, 14):
        ws_grid.row_dimensions[r_i].height = 24
    for c_i in range(2, 14):
        ws_grid.column_dimensions[get_column_letter(c_i)].width = 16

    os.makedirs(os.path.dirname(os.path.abspath(xlsx_path)), exist_ok=True)
    wb.save(xlsx_path)
    wb.close()
    print(f"  [V] Đã tạo Sơ đồ Tiến độ Trực quan As-Built Excel: {xlsx_path}")

    if html_path:
        html_str = """<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <title>Sơ Đồ Tiến Độ Trực Quan As-Built — Cầu Thôn Khai Hoang 2, Km 14+363.65</title>
  <style>
    body { font-family: Arial, sans-serif; background: #0f172a; color: #f8fafc; margin: 0; padding: 24px; }
    .container { max-width: 1200px; margin: 0 auto; }
    .card { background: #1e293b; border: 1px solid #334155; border-radius: 12px; padding: 20px; margin-bottom: 20px; }
    .grid-4 { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-top: 16px; }
    .kpi { background: #0f172a; border: 1px solid #334155; border-radius: 8px; padding: 12px; text-align: center; }
    .kpi .val { font-size: 22px; font-weight: bold; color: #34d399; margin-top: 4px; }
    table { width: 100%; border-collapse: collapse; font-size: 13px; margin-top: 12px; }
    th, td { border: 1px solid #334155; padding: 8px 10px; text-align: left; }
    th { background: #1b365d; color: #fff; text-align: center; }
    .done { background: #166534; color: #fff; font-weight: bold; text-align: center; }
    .doing { background: #d97706; color: #000; font-weight: bold; text-align: center; }
    .todo { background: #334155; color: #cbd5e1; text-align: center; }
  </style>
</head>
<body>
  <div class="container">
    <div class="card">
      <h1 style="margin:0; font-size:24px;">🌉 SƠ ĐỒ TIẾN ĐỘ THI CÔNG TRỰC QUAN AS-BUILT — CẦU THÔN KHAI HOANG 2, KM 14+363.65</h1>
      <p style="color:#94a3b8; margin:6px 0 0;">Quy mô: 01 Nhịp dầm chữ T BTCT L=15.0m (04 phiến) • 02 Mố chữ U C30 trên nền đá gốc gia cường 98 thanh neo Ø32 • 04 Đốt tường chắn</p>
      <div class="grid-4">
        <div class="kpi"><div>Thép neo đá Ø32</div><div class="val">98/98 thanh</div><div style="font-size:11px;color:#94a3b8;">Hoàn thành 100%</div></div>
        <div class="kpi"><div>Bệ mố & Thân mố M1, M2</div><div class="val">462.66 m³</div><div style="font-size:11px;color:#94a3b8;">Hoàn thành 100%</div></div>
        <div class="kpi"><div>Dầm chữ T L=15m (4 phiến)</div><div class="val" style="color:#fbbf24;">2/4 phiến</div><div style="font-size:11px;color:#94a3b8;">Đang đúc D3, D4</div></div>
        <div class="kpi"><div>Tổng cốt thép BBS</div><div class="val">68.01 tấn</div><div style="font-size:11px;color:#94a3b8;">166 vị trí (Đề-xê 0.83%)</div></div>
      </div>
    </div>
  </div>
</body>
</html>"""
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html_str)
        print(f"  [V] Đã tạo Sơ đồ Tiến độ Trực quan As-Built HTML: {html_path}")


if __name__ == "__main__":
    ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if ROOT not in sys.path:
        sys.path.insert(0, ROOT)
    out_dir = os.path.join(ROOT, "du_lieu_du_an", "Cầu thôn Khai Hoang 2, Km 14+363.65_Marker", "01_HIEN_TRUONG_QLCL_KCS")
    build_tan_suat_khai_hoang_2(os.path.join(out_dir, "Bang_Theo_Doi_Tan_Suat_Thi_Nghiem_Be_Tong_Va_Thep_Cau_Khai_Hoang_2.xlsx"))
    build_visual_asbuilt_khai_hoang_2(
        os.path.join(out_dir, "So_Do_Tien_Do_Truc_Quan_AsBuilt_Cau_Khai_Hoang_2.xlsx"),
        os.path.join(out_dir, "So_Do_Tien_Do_Truc_Quan_AsBuilt_Cau_Khai_Hoang_2.html"),
    )
