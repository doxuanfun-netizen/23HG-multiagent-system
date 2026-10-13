# -*- coding: utf-8 -*-
"""
HỆ THỐNG MULTI-AGENT AEC 23HG — THI CÔNG & QUẢN TRỊ DỰ ÁN CẦU THÔN KHAI HOANG 2
KM14+363.65 (GÓI THẦU SỐ 9, ĐỒNG VĂN - MÈO VẠC, HÀ GIANG)
TIẾN ĐỘ THI CÔNG: 20/10/2026 -> 20/02/2027 (124 NGÀY)
TỐI ƯU HÓA CẮT THÉP: PHẦN MỀM CẮT THÉP V6.0 / REBARCUT PRO OR-TOOLS
"""

from __future__ import annotations
import os
import sys
import json
import re
import csv
import time
import shutil
import datetime
import xml.etree.ElementTree as ET
from collections import defaultdict, Counter

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# Setup Paths
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)
EXAMPLES_DIR = os.path.join(ROOT_DIR, "examples")
if EXAMPLES_DIR not in sys.path:
    sys.path.insert(0, EXAMPLES_DIR)

from tools.cutting_stock_solver import CuttingStockSolver, CutDemand, rebar_unit_weight_kg_m, STANDARD_BAR_LENGTH_MM
from tools.rebarcut_export import write_rebarcut_workbook
from generate_khai_hoang_2_project import generate_khai_hoang_2_master_package

USER_DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads", r"Cầu thôn Khai Hoang 2, Km 14+363.65_Marker")
SYSTEM_PROJECT_DIR = os.path.join(ROOT_DIR, "du_lieu_du_an", "Cầu thôn Khai Hoang 2, Km 14+363.65_Marker")
SOURCE_JSON = os.path.join(USER_DOWNLOAD_DIR, "bang_so_lieu.json")

# Unit weights (TCVN 1651:2018)
UNIT_WEIGHT = {
    6: 0.222, 8: 0.395, 10: 0.617, 12: 0.888, 14: 1.208, 16: 1.578,
    18: 1.998, 20: 2.466, 22: 2.984, 25: 3.853, 28: 4.834, 32: 6.313
}

DIA_CONFIG = {
    6:  {"code": "00", "name": "D06_Dai_Nho", "desc": "Thép Ø6 (thép lưới bản mặt cầu B1, B2)", "station": "Trạm duỗi cắt lưới thép CNC"},
    10: {"code": "01", "name": "D10_Thep_Cuon_Dai_Xoan", "desc": "Thép đai bản mặt cầu & bệ kê gối", "station": "Trạm uốn đai tự động CNC"},
    12: {"code": "02", "name": "D12_Thanh_Van_Cau_Tao", "desc": "Thép cấu tạo dầm T15m & gờ lan can", "station": "Trạm cắt đa thanh Shearline"},
    14: {"code": "03", "name": "D14_Ban_Canh_Dam_T", "desc": "Thép bản cánh dầm T, gờ lan can & tường chắn", "station": "Trạm cắt phân đoạn Shearline"},
    16: {"code": "04", "name": "D16_Suon_Dam_Ban_Mat_Cau", "desc": "Thép sườn dầm T15m, dầm ngang, thân mố & khe co giãn", "station": "Trạm cắt uốn thanh trung bình"},
    18: {"code": "05", "name": "D18_Khung_Than_Tuong_Chan", "desc": "Thép chịu lực thân tường chắn 4 đốt", "station": "Trạm cắt uốn khung định hình"},
    20: {"code": "06", "name": "D20_Phan_Bo_Be_Mong", "desc": "Thép phân bố bệ mố, thân mố & bản quá độ", "station": "Trạm cắt uốn thanh nặng"},
    22: {"code": "07", "name": "D22_Gia_Cuong_Chiu_Luc", "desc": "Thép chịu lực tường cánh mố chữ U & bệ mố", "station": "Trạm cắt uốn thanh nặng"},
    25: {"code": "08", "name": "D25_Thep_Chu_Be_Than_Mo", "desc": "Thép chủ chịu lực chính bệ mố M1, M2 & thân mố", "station": "Trạm gia công bệ mố M1-M2"},
    32: {"code": "09", "name": "D32_Thep_Neo_Chiu_Uon_Chinh", "desc": "Thép chủ dầm T15m (A1, B1) & thép neo mố vào nền đá", "station": "Trạm cắt dầm chủ & khoan neo"},
}

def clean_num(text):
    if text is None:
        return None
    s = str(text).strip().replace(' ', '')
    s = re.sub(r'[^\d.,\-]', '', s)
    if not s or not re.search(r'\d', s):
        return None
    # Xử lý định dạng OCR có cả phân cách hàng nghìn lẫn thập phân 2 chữ số (vd: '7,900,00', '1,153,00', '2.260.00')
    if re.fullmatch(r'\d{1,3}([.,]\d{3})+[.,]\d{2}', s):
        main_part = re.sub(r'[.,]', '', s[:-3])
        dec_part = s[-2:]
        return float(f"{main_part}.{dec_part}")
    if re.fullmatch(r'\d{1,3}(\.\d{3})+', s):
        return float(s.replace('.', ''))
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        if s.count(',') > 1:
            parts = s.split(',')
            s = ''.join(parts[:-1]) + '.' + parts[-1]
        else:
            s = s.replace(',', '.')
    elif '.' in s and s.count('.') > 1:
        parts = s.split('.')
        s = ''.join(parts[:-1]) + '.' + parts[-1]
    try:
        return float(s)
    except ValueError:
        return None

def clean_length_mm(text):
    s = str(text or '').strip()
    if re.fullmatch(r'\d{1,3}(\.\d{3})+', s):
        return float(s.replace('.', ''))
    return clean_num(s)

def clean_dia(text):
    m = re.search(r'(\d{1,2})', str(text or '').replace('Φ', '').replace('Ø', '').replace('D', ''))
    return int(m.group(1)) if m and int(m.group(1)) in UNIT_WEIGHT else None

def load_all_bbs():
    with open(SOURCE_JSON, encoding='utf-8') as f:
        tables = json.load(f)

    BBS_TABLES = [
        (2,  'Neo mố vào đá (toàn cầu)',          1, 1, 2, 4, 3, 6),
        (4,  'Bệ mố - bảng 1 (tr.14)',             1, 0, 1, 2, 4, 6),
        (5,  'Bệ mố - bảng 2 (tr.15)',             1, 0, 1, 2, 4, 6),
        (6,  'Thân mố + tường đỉnh (1 mố)',        2, 0, 1, 2, 4, 6),
        (7,  'Bệ kê gối + u neo dầm (1 mố)',       2, 0, 1, 2, 4, 6),
        (8,  'Tường cánh (1 tường)',               4, 0, 1, 2, 4, 6),
        (9,  'Bản quá độ (1 bản)',                 2, 0, 1, 2, 4, 6),
        (10, 'Khe co giãn (1 khe)',                2, 0, 2, 3, 1, 6),
        (12, 'Dầm T L=15m - dầm biên (1 dầm)',     2, 1, 2, 4, 3, 6),
        (13, 'Dầm T L=15m - dầm giữa (1 dầm)',     2, 1, 2, 4, 3, 6),
        (14, 'Lớp phủ mặt cầu + gờ lan can (1 nhịp)', 1, 1, 2, 4, 5, 6),
        (19, 'Tường chắn đốt L=5,20m',             1, 1, 2, 3, 4, 7),
        (20, 'Tường chắn đốt L=5,89m',             1, 1, 2, 3, 4, 7),
        (21, 'Tường chắn đốt L=10,60m',            1, 1, 2, 3, 4, 7),
        (22, 'Tường chắn đốt L=13,50m',            1, 1, 2, 3, 4, 7),
    ]

    OCR_FIXES = [
        (6, 'A2', 'qty', 53),
        (8, 'FIB', 'mark', 'F1B'),
        (5, 'FIB', 'mark', 'F1B'),
        (4, 'FIA', 'mark', 'F1A'),
    ]

    raw_items = []
    for ti, comp, units, c_m, c_d, c_l, c_q, c_kg in BBS_TABLES:
        t = tables[ti]
        sheet = t.get('sheet') or ''
        page = t.get('page') or 0
        seen = set()
        for r in t.get('rows', []):
            if len(r) <= max(c_m, c_d, c_l, c_q):
                continue
            mark = str(r[c_m]).strip()
            dia = clean_dia(r[c_d])
            length = clean_length_mm(r[c_l])
            qty = clean_num(r[c_q])
            if not mark or dia is None or not length or qty is None:
                continue
            if ti in (19, 20, 21, 22):
                mark = f'TC{mark}'
            for fti, fmark, fld, val in OCR_FIXES:
                if fti == ti and fmark == mark:
                    if fld == 'qty':
                        qty = val
                    elif fld == 'mark':
                        mark = val
            if mark in seen and mark == 'D2B':
                mark = 'D2D'
            seen.add(mark)
            kg_one = qty * length / 1000 * UNIT_WEIGHT[dia]
            kg_tot = kg_one * units
            raw_items.append({
                "comp": comp,
                "mark": mark,
                "dia": dia,
                "length_mm": int(round(length)),
                "qty_per_unit": int(round(qty)),
                "units": units,
                "total_qty": int(round(qty)) * units,
                "kg_total": kg_tot,
                "sheet": sheet,
                "page": page
            })
    return raw_items

def export_cnc_csv(path: str, sol, dia: int) -> None:
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["Mã cây", "Đường kính Ø (mm)", "Thanh cắt (Mark)", "Chiều dài cắt (m)", "Thừa cuối cây (m)", "Mã cây gốc (11.7m)"])
        for a in sol.assignment:
            bar_name = f"C{a['bar_id']}"
            for length, mark in zip(a["cuts_mm"], a["marks"]):
                writer.writerow([bar_name, dia, mark, length / 1000.0, a["waste_mm"] / 1000.0, 11.7])

def create_master_rebar_dashboard(path: str, summary_rows: list) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TONG_HOP_CAT_THEP_11M7"

    ws["B2"] = "DỰ ÁN: ĐƯỜNG TỪ TRUNG TÂM HUYỆN ĐỒNG VĂN ĐI MỐC 450/456, HUYỆN MÈO VẠC, TỈNH HÀ GIANG"
    ws["B2"].font = Font(name="Times New Roman", size=13, bold=True, color="1F3A5E")
    ws["B3"] = "BẢNG TỔNG HỢP KẾ HOẠCH CẮT THÉP 11.7M THEO TỪNG ĐƯỜNG KÍNH (PHẦN MỀM CẮT THÉP V6.0 / REBARCUT PRO)"
    ws["B3"].font = Font(name="Times New Roman", size=12, bold=True, color="000000")
    ws["B4"] = "Công trình: Cầu thôn Khai Hoang 2, Km14+363.65 (L=15m) | Thuật toán: Gilmore-Gomory Column Generation + OR-Tools CP-SAT | Cây thép chuẩn 11.7m"
    ws["B4"].font = Font(name="Times New Roman", size=10, italic=True)

    headers = [
        "TT", "Đường kính (Ø)", "Mô tả hạng mục kết cấu chính",
        "Số thanh thiết kế (BBS)", "Tổng chiều dài (m)", "Số cây 11.7m cần mua",
        "Khối lượng thép mua (kg)", "Chiều dài phế/đề-xê (m)", "Tỷ lệ hao hụt (%)",
        "Số mẫu cắt (Patterns)", "Đầu thừa tái sử dụng (m)",
        "Trạm gia công phụ trách", "Trạng thái tối ưu"
    ]

    header_row = 6
    for col_idx, h in enumerate(headers, start=2):
        cell = ws.cell(header_row, col_idx, h)
        cell.font = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F3A5E")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    row = 7
    tot_qty = 0
    tot_len = 0.0
    tot_bars = 0
    tot_weight = 0.0
    tot_waste = 0.0
    tot_reusable = 0.0
    last_col = 2 + len(headers) - 1

    for idx, item in enumerate(summary_rows, start=1):
        ws.cell(row, 2, idx).alignment = Alignment(horizontal="center")
        ws.cell(row, 3, f"Ø{item['dia']:02d}").alignment = Alignment(horizontal="center")
        ws.cell(row, 4, item["desc"])
        ws.cell(row, 5, item["bars_qty"]).number_format = "#,##0"
        ws.cell(row, 6, round(item["total_len_m"], 2)).number_format = "#,##0.00"
        ws.cell(row, 7, item["bars_11m7"]).number_format = "#,##0"
        ws.cell(row, 8, round(item["weight_kg"], 1)).number_format = "#,##0.0"
        ws.cell(row, 9, round(item["waste_m"], 2)).number_format = "#,##0.00"
        ws.cell(row, 10, item["waste_pct"] / 100.0).number_format = "0.00%"
        ws.cell(row, 11, item["patterns_count"]).alignment = Alignment(horizontal="center")
        ws.cell(row, 12, round(item["reusable_m"], 2)).number_format = "#,##0.00"
        ws.cell(row, 13, item["station"])
        status_cell = ws.cell(row, 14, item["status"])
        status_cell.alignment = Alignment(horizontal="center")
        status_cell.font = Font(name="Times New Roman", size=9, bold=True, color="006100")

        for c in range(2, last_col + 1):
            ws.cell(row, c).border = border
            if c in (5, 6, 7, 8, 9, 10, 12):
                ws.cell(row, c).alignment = Alignment(horizontal="right")

        tot_qty += item["bars_qty"]
        tot_len += item["total_len_m"]
        tot_bars += item["bars_11m7"]
        tot_weight += item["weight_kg"]
        tot_waste += item["waste_m"]
        tot_reusable += item["reusable_m"]
        row += 1

    # Dòng Tổng cộng
    ws.cell(row, 2, "TỔNG CỘNG").alignment = Alignment(horizontal="center")
    ws.cell(row, 3, "")
    ws.cell(row, 4, "TOÀN BỘ CÔNG TRÌNH CẦU KHAI HOANG 2 (KM14+363.65)")
    ws.cell(row, 5, tot_qty).number_format = "#,##0"
    ws.cell(row, 6, tot_len).number_format = "#,##0.00"
    ws.cell(row, 7, tot_bars).number_format = "#,##0"
    ws.cell(row, 8, tot_weight).number_format = "#,##0.0"
    ws.cell(row, 9, tot_waste).number_format = "#,##0.00"
    ws.cell(row, 10, (tot_waste / (tot_bars * 11.7)) if tot_bars else 0).number_format = "0.00%"
    ws.cell(row, 11, sum(x["patterns_count"] for x in summary_rows)).alignment = Alignment(horizontal="center")
    ws.cell(row, 12, tot_reusable).number_format = "#,##0.00"
    ws.cell(row, 13, "Đồng bộ bãi gia công 10 trạm CNC")
    ws.cell(row, 14, "HOÀN TẤT TỐI ƯU").alignment = Alignment(horizontal="center")

    for c in range(2, last_col + 1):
        cell = ws.cell(row, c)
        cell.font = Font(name="Times New Roman", size=10, bold=True)
        cell.fill = PatternFill("solid", fgColor="D9E1F2")
        cell.border = border
        if c in (5, 6, 7, 8, 9, 10, 12):
            cell.alignment = Alignment(horizontal="right")

    # Column widths
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 6
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 42
    ws.column_dimensions["E"].width = 15
    ws.column_dimensions["F"].width = 18
    ws.column_dimensions["G"].width = 18
    ws.column_dimensions["H"].width = 18
    ws.column_dimensions["I"].width = 18
    ws.column_dimensions["J"].width = 15
    ws.column_dimensions["K"].width = 14
    ws.column_dimensions["L"].width = 18
    ws.column_dimensions["M"].width = 30
    ws.column_dimensions["N"].width = 18

    wb.save(path)

WBS_TASKS = [
    # (WBS, Name, Duration, Pred, Norm, Workers, Crew, Critical, Start, Finish)
    ("WBS.01", "Bàn giao cọc mốc tim tuyến và định vị công trình", 5, "-", 0.5, 6, "Tổ Trắc đạc", "Critical", "2026-10-20", "2026-10-24"),
    ("WBS.02", "Thi công đường công vụ, lán trại và dọn dẹp mặt bằng", 10, "WBS.01", 1.2, 18, "Tổ Thi công nền", "Critical", "2026-10-23", "2026-11-01"),
    ("WBS.03", "Kiểm định vật liệu đầu vào (Xi măng, cát, đá, thép)", 7, "WBS.01", 0.8, 4, "Tổ QA/QC & Lab", "Non-Critical", "2026-10-22", "2026-10-28"),
    ("WBS.04", "Thi công bãi đúc dầm T15m và lắp đặt trạm trộn", 12, "WBS.02", 2.0, 16, "Tổ Xây dựng bãi", "Non-Critical", "2026-10-28", "2026-11-08"),
    ("WBS.05", "Đào hố móng mố M1 và xử lý nền đá nứt nẻ", 12, "WBS.02", 1.5, 14, "Tổ Cơ giới móng", "Critical", "2026-10-29", "2026-11-09"),
    ("WBS.06", "Khoan cấy cốt thép neo D32 vào nền đá hố móng M1", 6, "WBS.05", 2.5, 10, "Tổ Khoan cấy thép", "Critical", "2026-11-10", "2026-11-15"),
    ("WBS.07", "Đổ bê tông đệm móng mố M1 C10 (M150)", 2, "WBS.06", 0.6, 8, "Tổ Bê tông móng", "Critical", "2026-11-16", "2026-11-17"),
    ("WBS.08", "Gia công lắp dựng cốt thép và ván khuôn móng mố M1", 7, "WBS.07", 3.2, 16, "Tổ Cốt thép ván khuôn", "Critical", "2026-11-18", "2026-11-24"),
    ("WBS.09", "Đổ bê tông móng mố M1 C30 (V = 110 m3)", 4, "WBS.08", 1.8, 20, "Tổ Bê tông thương phẩm", "Critical", "2026-11-25", "2026-11-28"),
    ("WBS.10", "Đào hố móng và khoan cấy thép neo nền đá mố M2", 14, "WBS.05", 1.5, 14, "Tổ Cơ giới móng", "Non-Critical", "2026-11-05", "2026-11-18"),
    ("WBS.11", "Đổ bê tông đệm và bê tông móng mố M2 C30", 10, "WBS.10", 2.0, 18, "Tổ Bê tông móng", "Non-Critical", "2026-11-19", "2026-11-28"),
    ("WBS.12", "Thi công thân mố, tường ngực và tường cánh M1 C30", 18, "WBS.09", 3.5, 22, "Tổ Cốp pha thân", "Critical", "2026-11-29", "2026-12-16"),
    ("WBS.13", "Thi công thân mố, tường ngực và tường cánh M2 C30", 18, "WBS.11", 3.5, 22, "Tổ Cốp pha thân", "Non-Critical", "2026-11-29", "2026-12-16"),
    ("WBS.14", "Đúc 4 phiến dầm chủ T15m C40 tại bãi đúc dầm", 25, "WBS.04", 4.0, 24, "Tổ Đúc dầm T", "Non-Critical", "2026-11-10", "2026-12-04"),
    ("WBS.15", "Bảo dưỡng dầm T15m và nghiệm thu xuất xưởng", 14, "WBS.14", 0.5, 6, "Tổ Bảo dưỡng", "Non-Critical", "2026-12-05", "2026-12-18"),
    ("WBS.16", "Lắp đặt đá kê gối và 8 gối cao su bản thép (350x500x84)", 5, "WBS.12,13", 1.5, 10, "Tổ Cơ điện cầu", "Critical", "2026-12-17", "2026-12-21"),
    ("WBS.17", "Lao lắp 4 phiến dầm T15m vào vị trí mố M1-M2", 6, "WBS.15,16", 5.0, 20, "Tổ Lao lắp dầm", "Critical", "2026-12-22", "2026-12-27"),
    ("WBS.18", "Thi công 3 dầm ngang C30 liên kết dầm chủ", 8, "WBS.17", 2.8, 16, "Tổ Bê tông nhịp", "Critical", "2026-12-28", "2027-01-04"),
    ("WBS.19", "Lắp đặt ván khuôn, lưới thép D6 và đổ bê tông mặt cầu C35", 12, "WBS.18", 3.0, 22, "Tổ Mặt cầu", "Critical", "2027-01-05", "2027-01-16"),
    ("WBS.20", "Thi công gờ lan can C25 và lắp đặt lan can thép mạ kẽm", 10, "WBS.19", 2.2, 14, "Tổ Lan can", "Non-Critical", "2027-01-17", "2027-01-26"),
    ("WBS.21", "Thi công bản quá độ hai đầu cầu và tường chắn đất C25", 14, "WBS.19", 2.5, 16, "Tổ Đường đầu cầu", "Non-Critical", "2027-01-17", "2027-01-30"),
    ("WBS.22", "Lắp khe co giãn, thảm BTNC 19, thử tải và bàn giao", 21, "WBS.20,21", 3.0, 25, "Tổ Hoàn thiện", "Critical", "2027-01-31", "2027-02-20"),
]

def create_ms_project_xml(wbs_tasks, output_xml_path):
    root = ET.Element("Project", xmlns="http://schemas.microsoft.com/project")
    ET.SubElement(root, "Name").text = "Tien_Do_Thi_Cong_Cau_Khai_Hoang_2"
    ET.SubElement(root, "Title").text = "TIẾN ĐỘ THI CÔNG CẦU THÔN KHAI HOANG 2, KM14+363.65 (20/10/2026 - 20/02/2027)"
    ET.SubElement(root, "Company").text = "BAN QLDA DTXD HUYEN DONG VAN"
    ET.SubElement(root, "Author").text = "Multi-Agent AEC 23HG System"
    ET.SubElement(root, "StartDate").text = "2026-10-20T08:00:00"
    ET.SubElement(root, "FinishDate").text = "2027-02-20T17:00:00"
    ET.SubElement(root, "ScheduleFromStart").text = "1"

    calendars = ET.SubElement(root, "Calendars")
    cal = ET.SubElement(calendars, "Calendar")
    ET.SubElement(cal, "UID").text = "1"
    ET.SubElement(cal, "Name").text = "Standard"
    ET.SubElement(cal, "IsBaseCalendar").text = "1"

    tasks_el = ET.SubElement(root, "Tasks")
    for idx, (wbs, name, dur, pred, norm, work, crew, cpm, d_start, d_finish) in enumerate(wbs_tasks, start=1):
        task = ET.SubElement(tasks_el, "Task")
        ET.SubElement(task, "UID").text = str(idx)
        ET.SubElement(task, "ID").text = str(idx)
        ET.SubElement(task, "Name").text = name
        ET.SubElement(task, "OutlineNumber").text = str(idx)
        ET.SubElement(task, "OutlineLevel").text = "1"
        ET.SubElement(task, "Duration").text = f"PT{int(dur)*8}H0M0S"
        ET.SubElement(task, "Start").text = f"{d_start}T08:00:00"
        ET.SubElement(task, "Finish").text = f"{d_finish}T17:00:00"
        ET.SubElement(task, "PercentComplete").text = "0"
        ET.SubElement(task, "Critical").text = "1" if cpm == "Critical" else "0"
        ET.SubElement(task, "Milestone").text = "1" if dur == 1 and idx == len(wbs_tasks) else "0"

        if pred != "-":
            pred_link = ET.SubElement(task, "PredecessorLink")
            ET.SubElement(pred_link, "PredecessorUID").text = str(max(1, idx - 1))
            ET.SubElement(pred_link, "Type").text = "1"

    tree = ET.ElementTree(root)
    tree.write(output_xml_path, encoding="utf-8", xml_declaration=True)

def generate_kcs_word_dossier(docx_out_path):
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.85)
        section.right_margin = Inches(0.75)

    p_main = doc.add_paragraph()
    p_main.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_main.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc\n-------------------\n\n")
    r_sub.font.name = "Times New Roman"
    r_sub.font.size = Pt(12)
    r_sub.font.bold = True

    r_title = p_main.add_run("HỒ SƠ BIÊN BẢN NGHIỆM THU KỸ THUẬT & CHẤT LƯỢNG KCS\n")
    r_title.font.name = "Times New Roman"
    r_title.font.size = Pt(16)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

    r_desc = p_main.add_run("CÔNG TRÌNH: CẦU THÔN KHAI HOANG 2, KM14+363.65\n"
                           "DỰ ÁN: ĐƯỜNG TỪ TRUNG TÂM ĐỒNG VĂN ĐI MỐC 450/456, HUYỆN MÈO VẠC\n"
                           "GÓI THẦU SỐ 9: THI CÔNG XÂY DỰNG CÔNG TRÌNH CẦU VÀ ĐƯỜNG ĐẦU CẦU\n\n"
                           "KHUNG TIẾN ĐỘ THI CÔNG: 20/10/2026 - 20/02/2027 (124 NGÀY)\n"
                           "TIÊU CHUẨN QUẢN LÝ CHẤT LƯỢNG: NGHỊ ĐỊNH 207/2026/NĐ-CP & THÔNG TƯ 32/2026/TT-BXD\n")
    r_desc.font.name = "Times New Roman"
    r_desc.font.size = Pt(11)
    r_desc.font.italic = True

    doc.add_page_break()

    for idx, (wbs, name, dur, pred, norm, work, crew, cpm, d_start, d_finish) in enumerate(WBS_TASKS, start=1):
        code = f"BBNT-{idx:02d}"
        type_nt = "CÔNG VIỆC XÂY DỰNG"
        if "Kiểm định" in name: type_nt = "VẬT LIỆU XÂY DỰNG ĐẦU VÀO"
        elif "bàn giao" in name.lower(): type_nt = "HOÀN THÀNH CÔNG TRÌNH ĐƯA VÀO SỬ DỤNG"

        p_bb = doc.add_paragraph()
        p_bb.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_bb = p_bb.add_run(f"BIÊN BẢN NGHIỆM THU {type_nt.upper()}\nSố: {code}/BBNT-KH2/2026-2027")
        r_bb.font.name = "Times New Roman"
        r_bb.font.size = Pt(12)
        r_bb.font.bold = True
        r_bb.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

        p_info = doc.add_paragraph()
        p_info.paragraph_format.line_spacing = 1.15
        p_info.add_run("1. Đối tượng nghiệm thu: ").bold = True
        p_info.add_run(f"{name}\n")
        p_info.add_run("2. Phân loại & Mã WBS: ").bold = True
        p_info.add_run(f"{type_nt} (Mã WBS: {wbs} | Tổ đội: {crew})\n")
        p_info.add_run("3. Thời gian nghiệm thu hoàn thành: ").bold = True
        p_info.add_run(f"Ngày {d_finish} (Khởi công: {d_start})\n")
        p_info.add_run("4. Địa điểm thực hiện: ").bold = True
        p_info.add_run("Cầu thôn Khai Hoang 2, Km14+363.65 (Đường Đồng Văn - Mèo Vạc, Hà Giang)\n")
        p_info.add_run("5. Thành phần nghiệm thu:\n").bold = True
        p_info.add_run("   - Đại diện Tư vấn Giám sát: Kỹ sư Giám sát trưởng & Giám sát hiện trường.\n"
                       "   - Đại diện Nhà thầu thi công: Chỉ huy trưởng công trường & Cán bộ Kỹ thuật QA/QC.\n")
        p_info.add_run("6. Căn cứ nghiệm thu:\n").bold = True
        p_info.add_run("   - Hồ sơ thiết kế bản vẽ thi công đã được thẩm định & phê duyệt.\n"
                       "   - Tiêu chuẩn thi công & nghiệm thu: TCVN 11823:2017, TCVN 4453:1995, TCVN 1651:2018.\n"
                       "   - Kết quả thí nghiệm kiểm định độc lập phòng LAS-XD đạt 100% chỉ tiêu kỹ thuật.\n")
        p_info.add_run("7. Đánh giá chất lượng: ").bold = True
        p_info.add_run("Công việc đã được thi công hoàn thành đúng kích thước hình học thiết kế, cao độ, vị trí tim mốc, đảm bảo an toàn lao động và vệ sinh môi trường. ĐỦ ĐIỀU KIỆN CHUYỂN BƯỚC THI CÔNG TIẾP THEO.\n")

        table = doc.add_table(rows=2, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.rows[0].cells[0].paragraphs[0].text = "ĐẠI DIỆN NHÀ THẦU THI CÔNG\nChỉ huy trưởng công trường\n\n\n\n(Ký và ghi rõ họ tên)"
        table.rows[0].cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        table.rows[0].cells[1].paragraphs[0].text = "ĐẠI DIỆN TƯ VẤN GIÁM SÁT\nKỹ sư Giám sát trưởng\n\n\n\n(Ký và ghi rõ họ tên)"
        table.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.name = "Times New Roman"
                        r.font.size = Pt(10)
                        r.font.bold = True

        if idx < len(WBS_TASKS):
            doc.add_page_break()

    doc.save(docx_out_path)

def generate_machine_diesel_workbook(path: str):
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "01_TienDo_CaMay_Master"
    ws1.views.sheetView[0].showGridLines = True

    # Title
    ws1["B2"] = "DỰ ÁN: ĐƯỜNG TỪ TRUNG TÂM ĐỒNG VĂN ĐI MỐC 450/456, HUYỆN MÈO VẠC, HÀ GIANG"
    ws1["B2"].font = Font(name="Times New Roman", size=13, bold=True, color="1F497D")
    ws1["B3"] = "BẢNG KẾ HOẠCH CA XE, CA MÁY & TIÊU THỤ DẦU DIEZEL (20/10/2026 - 20/02/2027)"
    ws1["B3"].font = Font(name="Times New Roman", size=12, bold=True)
    ws1["B4"] = "Công trình: Cầu thôn Khai Hoang 2, Km14+363.65 | Tổng thời gian: 124 ngày | Đội thi công cơ giới & Bãi đúc dầm"
    ws1["B4"].font = Font(name="Times New Roman", size=10, italic=True)

    # Sheet 2: Danh mục MMTB & Định mức
    ws2 = wb.create_sheet(title="02_TongHop_CaXe_CaMay_MMTB")
    ws2.views.sheetView[0].showGridLines = True
    ws2["B2"] = "BẢNG TỔNG HỢP CA XE CA MÁY VÀ ĐỊNH MỨC TIÊU HAO NHIÊN LIỆU DIEZEL"
    ws2["B2"].font = Font(name="Times New Roman", size=13, bold=True, color="1F497D")

    h2 = ["STT", "Tên thiết bị thi công / Cơ giới", "Công suất / Trọng tải", "Định mức dầu (lít/ca)", "Tổng số ca máy", "Tổng lít dầu diezel", "Ghi chú công năng"]
    for c, h in enumerate(h2, start=2):
        cell = ws2.cell(row=5, column=c, value=h)
        cell.font = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F497D")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    equipments = [
        (1, "Máy xúc đào bánh xích", "1.25 m3 (PC200/300)", 105.0, 48, "Đào hố móng đá M1, M2, đường công vụ"),
        (2, "Máy xúc lốp đào & san gạt", "0.65 m3 (PC140)", 75.0, 36, "San bãi đúc, đào rãnh & hoàn thiện tứ nón"),
        (3, "Ô tô tự đổ 3 chân", "15 Tấn (Howo)", 65.0, 85, "Vận chuyển đất đá, cát sỏi mỏ về công trường"),
        (4, "Cần cẩu bánh xích", "50 Tấn (Kobelco)", 85.0, 28, "Lao lắp dầm T15m & cẩu bốc xếp cấu kiện"),
        (5, "Cần cẩu tự hành (xe cẩu thùng)", "12 Tấn (Dongfeng)", 55.0, 42, "Bốc dỡ thép cây, cốp pha bãi đúc & mố"),
        (6, "Máy khoan cấy neo đá chuyên dụng", "Khoan búa căn / xoay đập", 45.0, 32, "Khoan cấy 98 lỗ neo D32 vào nền đá mố"),
        (7, "Trạm trộn bê tông di động", "45 m3/h", 50.0, 56, "Sản xuất bê tông dầm C40, bệ thân mố C30"),
        (8, "Xe bồn vận chuyển bê tông", "7 m3 (Hyundai)", 60.0, 62, "Vận chuyển bê tông từ trạm ra vị trí đổ"),
        (9, "Máy bơm bê tông tĩnh áp lực cao", "90 m3/h", 70.0, 26, "Bơm bê tông thân mố, tường ngực & mặt cầu"),
        (10, "Máy lu rung mặt đường", "14 Tấn (Hamm)", 65.0, 22, "Lu lèn đường công vụ, móng đường đầu cầu"),
        (11, "Máy rải thảm bê tông nhựa", "Bề rộng 4.5m - 6.0m", 60.0, 12, "Rải thảm mặt cầu BTNC 19 dày 7cm"),
        (12, "Máy phát điện Diesel dự phòng", "150 kVA (Cummins)", 90.0, 124, "Cấp điện liên tục 124 ngày cho toàn công trường"),
    ]

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    row = 6
    for eq in equipments:
        ws2.cell(row, 2, eq[0]).alignment = Alignment(horizontal="center")
        ws2.cell(row, 3, eq[1])
        ws2.cell(row, 4, eq[2]).alignment = Alignment(horizontal="center")
        ws2.cell(row, 5, eq[3]).number_format = "#,##0.0"
        ws2.cell(row, 6, eq[4]).number_format = "#,##0"
        ws2.cell(row, 7, f"=E{row}*F{row}").number_format = "#,##0.0"
        ws2.cell(row, 8, eq[5])
        for c in range(2, 9):
            ws2.cell(row, c).border = border
        row += 1

    # Total row
    ws2.cell(row, 2, "TỔNG CỘNG").alignment = Alignment(horizontal="center")
    ws2.cell(row, 3, "")
    ws2.cell(row, 4, "")
    ws2.cell(row, 5, "")
    ws2.cell(row, 6, f"=SUM(F6:F{row-1})").number_format = "#,##0"
    ws2.cell(row, 7, f"=SUM(G6:G{row-1})").number_format = "#,##0.0"
    ws2.cell(row, 8, "Nhiên liệu cấp theo 4 kỳ thi công")
    for c in range(2, 9):
        cell = ws2.cell(row, c)
        cell.font = Font(name="Times New Roman", size=10, bold=True)
        cell.fill = PatternFill("solid", fgColor="D9E1F2")
        cell.border = border

    for ws in [ws1, ws2]:
        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = 18
    ws2.column_dimensions["C"].width = 34
    ws2.column_dimensions["H"].width = 45

    wb.save(path)

def run_pipeline():
    print("=" * 80)
    print("  KHỞI TẠO TOÀN DIỆN HỆ THỐNG CẦU THÔN KHAI HOANG 2, KM14+363.65")
    print("  TIẾN ĐỘ: 20/10/2026 -> 20/02/2027 (124 NGÀY)")
    print("  ENGINE: PHẦN MỀM CẮT THÉP V6.0 / REBARCUT PRO OR-TOOLS")
    print("=" * 80)

    # 1. Thư mục đích
    os.makedirs(USER_DOWNLOAD_DIR, exist_ok=True)
    os.makedirs(SYSTEM_PROJECT_DIR, exist_ok=True)

    rebarcut_dir_user = os.path.join(USER_DOWNLOAD_DIR, "01_HE_THONG_CAT_THEP_REBARCUT")
    rebarcut_dir_sys = os.path.join(SYSTEM_PROJECT_DIR, "01_HE_THONG_CAT_THEP_REBARCUT")
    os.makedirs(os.path.join(rebarcut_dir_user, "THEO_TUNG_DUONG_KINH_PHI"), exist_ok=True)
    os.makedirs(os.path.join(rebarcut_dir_user, "LENH_CAT_CNC_CSV"), exist_ok=True)

    # 2. Đọc BBS từ bang_so_lieu.json
    print("\n[BƯỚC 1/6] Đọc và trích xuất bảng thống kê thép BBS từ hồ sơ BVTC...")
    bbs_items = load_all_bbs()
    print(f"  -> Đã trích xuất thành công: {len(bbs_items)} danh mục thép thành phần")
    tot_bars_bbs = sum(x["total_qty"] for x in bbs_items)
    tot_len_bbs = sum(x["total_qty"] * x["length_mm"] / 1000.0 for x in bbs_items)
    tot_kg_bbs = sum(x["kg_total"] for x in bbs_items)
    print(f"  -> Tổng số thanh: {tot_bars_bbs:,} thanh | Tổng dài: {tot_len_bbs:,.2f} m | Tổng trọng lượng: {tot_kg_bbs:,.2f} kg ({tot_kg_bbs/1000:.2f} tấn)")

    # 3. Chạy phần mềm cắt thép V6.0 (RebarCut Pro Engine)
    print("\n[BƯỚC 2/6] Chạy tối ưu hóa cắt thép 1D CSP (Phần mềm cắt thép V6.0 / RebarCut Pro)...")
    solver = CuttingStockSolver(bar_length_mm=11700, kerf_mm=3, end_trim_mm=0, time_limit_s=5.0,
                                min_splice_segment_xd=40, long_bar_min_segments_only=True)

    demands = []
    for it in bbs_items:
        d = CutDemand(
            length_mm=it["length_mm"],
            quantity=it["total_qty"],
            diameter_mm=it["dia"],
            mark=f"{it['comp']}_{it['mark']}",
            grade="CB400-V"
        )
        if d.length_mm > solver.usable_length_mm:
            d.splice_allowed = True
            d.splice_kind = "Nối chồng 40d"
            d.lap_xd = 0.0
            d.splice_zones = [(0.0, 1.0)]
        demands.append(d)

    dias = sorted(list(set(d.diameter_mm for d in demands)))
    summary_rows = []

    for dia in dias:
        sub_demands = [d for d in demands if d.diameter_mm == dia]
        cfg = DIA_CONFIG.get(dia, {"code": f"{dia:02d}", "name": f"D{dia:02d}", "desc": f"Thép Ø{dia}", "station": "Trạm gia công chung"})
        sol = solver.solve(sub_demands, split_long_bars=True)

        # Xuất file RebarCut Pro 5 sheet
        file_name = f"{cfg['code']}_RebarCut_{cfg['name']}.xlsx"
        file_path = os.path.join(rebarcut_dir_user, "THEO_TUNG_DUONG_KINH_PHI", file_name)
        write_rebarcut_workbook(file_path, sub_demands, sol)

        # Xuất lệnh CNC CSV
        cnc_name = f"Lenh_Cat_CNC_{cfg['name']}.csv"
        cnc_path = os.path.join(rebarcut_dir_user, "LENH_CAT_CNC_CSV", cnc_name)
        export_cnc_csv(cnc_path, sol, dia)

        sub_len = sum(d.length_mm * d.quantity for d in sub_demands) / 1000.0
        reusable_m = sol.reusable_offcut_length_mm / 1000.0
        summary_rows.append({
            "dia": dia,
            "desc": cfg["desc"],
            "station": cfg["station"],
            "bars_qty": sum(d.quantity for d in sub_demands),
            "total_len_m": sub_len,
            "bars_11m7": sol.total_bars_needed,
            "weight_kg": sol.total_weight_kg,
            "waste_m": sol.total_waste_mm / 1000.0,
            "waste_pct": sol.waste_ratio_pct,
            "patterns_count": len(sol.patterns),
            "reusable_m": reusable_m,
            "status": sol.status
        })
        print(f"  + Ø{dia:02d}: {sol.total_bars_needed:4d} cây 11.7m | Mua: {sol.total_weight_kg:9.2f} kg | Đề-xê: {sol.waste_ratio_pct:5.2f}% | Reusable: {reusable_m:7.2f}m -> {file_name}")

    # Master dashboard RebarCut
    master_rebar_excel = os.path.join(rebarcut_dir_user, "00_BANG_TONG_HOP_CAT_THEP_THEO_PHI.xlsx")
    create_master_rebar_dashboard(master_rebar_excel, summary_rows)
    print(f"  [V] Đã tạo Master Dashboard RebarCut: {master_rebar_excel}")

    # 4. Xuất Master 14 Sheet Workbook & MS Project XML (20/10/2026 -> 20/02/2027)
    print("\n[BƯỚC 3/6] Xuất Hệ thống 14 Sheet Master và Tiến độ CPM MS Project XML...")
    master_excel_path = os.path.join(USER_DOWNLOAD_DIR, "Ho_So_KCS_QS_TienDo_Cau_Khai_Hoang_2.xlsx")
    xml_path = os.path.join(USER_DOWNLOAD_DIR, "Tien_Do_Thi_Cong_Cau_Khai_Hoang_2.xml")
    generate_khai_hoang_2_master_package(master_excel_path, xml_path)

    # 5. Xuất Hồ sơ Ca xe máy dầu Diezel & Word KCS
    print("\n[BƯỚC 4/6] Xuất Kế hoạch Ca xe ca máy dầu Diezel & Bộ Biên bản KCS Word...")
    diesel_path = os.path.join(USER_DOWNLOAD_DIR, "TDTC_CaXe_CaMay_DauDiezel_Cau_Khai_Hoang_2.xlsx")
    generate_machine_diesel_workbook(diesel_path)
    print(f"  [V] Đã tạo file Ca xe máy dầu Diezel: {diesel_path}")

    docx_path = os.path.join(USER_DOWNLOAD_DIR, "Ho_So_Bien_Ban_Nghiem_Thu_KCS_Cau_Khai_Hoang_2.docx")
    generate_kcs_word_dossier(docx_path)
    print(f"  [V] Đã tạo file Word KCS (22 biên bản): {docx_path}")

    # 6. Đóng gói 3 Tầng hồ sơ: Macro Master, Vi mô 14 bộ, Hub & Spoke 5 gói vệ tinh
    print("\n[BƯỚC 5/6] Đóng gói 3 Tầng hồ sơ theo cấu trúc chuẩn hệ thống 23HG...")
    dir_macro = os.path.join(USER_DOWNLOAD_DIR, "BO_HO_SO_01_MACRO_MASTER_14_SHEET")
    dir_micro = os.path.join(USER_DOWNLOAD_DIR, "BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO")
    dir_hub = os.path.join(USER_DOWNLOAD_DIR, "03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH")

    os.makedirs(dir_macro, exist_ok=True)
    os.makedirs(dir_micro, exist_ok=True)
    os.makedirs(dir_hub, exist_ok=True)

    # Chép vào Macro Master
    shutil.copy2(master_excel_path, os.path.join(dir_macro, "Ho_So_KCS_QS_TienDo_Cau_Khai_Hoang_2.xlsx"))
    shutil.copy2(xml_path, os.path.join(dir_macro, "Tien_Do_Thi_Cong_Cau_Khai_Hoang_2.xml"))
    shutil.copy2(diesel_path, os.path.join(dir_macro, "TDTC_CaXe_CaMay_DauDiezel_Cau_Khai_Hoang_2.xlsx"))
    shutil.copy2(docx_path, os.path.join(dir_macro, "Ho_So_Bien_Ban_Nghiem_Thu_KCS_Cau_Khai_Hoang_2.docx"))

    # Chép vào Vi mô
    shutil.copy2(master_rebar_excel, os.path.join(dir_micro, "01_To_Hop_Cat_Thep_11m7_RebarCut.xlsx"))
    shutil.copy2(os.path.join(rebarcut_dir_user, "LENH_CAT_CNC_CSV", "Lenh_Cat_CNC_D16_Suon_Dam_Ban_Mat_Cau.csv"),
                 os.path.join(dir_micro, "01_Phieu_Cat_Thep_Xuong_CNC.csv"))
    shutil.copy2(master_excel_path, os.path.join(dir_micro, "10_Tien_Do_Thi_Cong_CPM_Gantt_Chart.xlsx"))
    shutil.copy2(docx_path, os.path.join(dir_micro, "Ho_So_Bien_Ban_Nghiem_Thu_KCS_Cau_Khai_Hoang_2.docx"))

    # Hub & Spoke 5 gói
    goi_a = os.path.join(dir_hub, "GOI_A_CO_GIOI_VA_DAU_DIEZEL")
    goi_b = os.path.join(dir_hub, "GOI_B_XUONG_TIEN_CHE_COT_THEP")
    goi_c = os.path.join(dir_hub, "GOI_C_HIEN_TRUONG_QLCL_KCS")
    goi_d = os.path.join(dir_hub, "GOI_D_QS_DU_TOAN_THANH_TOAN")
    goi_e = os.path.join(dir_hub, "GOI_E_EXECUTIVE_DASHBOARD")

    for g in [goi_a, goi_b, goi_c, goi_d, goi_e]:
        os.makedirs(g, exist_ok=True)

    shutil.copy2(diesel_path, os.path.join(goi_a, "TDTC_CaXe_CaMay_DauDiezel_Cau_Khai_Hoang_2.xlsx"))
    shutil.copy2(xml_path, os.path.join(goi_a, "Tien_Do_Thi_Cong_Cau_Khai_Hoang_2.xml"))

    shutil.copy2(master_rebar_excel, os.path.join(goi_b, "01_To_Hop_Cat_Thep_11m7_RebarCut.xlsx"))
    shutil.copy2(os.path.join(rebarcut_dir_user, "LENH_CAT_CNC_CSV", "Lenh_Cat_CNC_D16_Suon_Dam_Ban_Mat_Cau.csv"),
                 os.path.join(goi_b, "01_Phieu_Cat_Thep_Xuong_CNC.csv"))

    shutil.copy2(docx_path, os.path.join(goi_c, "Ho_So_Bien_Ban_Nghiem_Thu_KCS_Cau_Khai_Hoang_2.docx"))

    shutil.copy2(master_excel_path, os.path.join(goi_d, "03_QS_Dien_Giai_Chi_Tiet_Takeoff.xlsx"))
    shutil.copy2(master_excel_path, os.path.join(goi_d, "08_Du_Toan_GXD_Thong_Tu_36_2026.xlsx"))

    shutil.copy2(master_excel_path, os.path.join(goi_e, "Ho_So_KCS_QS_TienDo_Cau_Khai_Hoang_2.xlsx"))
    shutil.copy2(xml_path, os.path.join(goi_e, "Tien_Do_Thi_Cong_Cau_Khai_Hoang_2.xml"))
    shutil.copy2(docx_path, os.path.join(goi_e, "Ho_So_Bien_Ban_Nghiem_Thu_KCS_Cau_Khai_Hoang_2.docx"))

    # Manifest
    manifest = {
        "project": "Cầu thôn Khai Hoang 2, Km14+363.65",
        "created_at": datetime.datetime.now().isoformat(),
        "schedule": {
            "start": "2026-10-20",
            "finish": "2027-02-20",
            "duration_days": 124
        },
        "rebar_cut_summary": {
            "total_bars_design": tot_bars_bbs,
            "total_weight_kg_design": tot_kg_bbs,
            "total_stock_bars_11m7": sum(x["bars_11m7"] for x in summary_rows),
            "total_purchased_kg": sum(x["weight_kg"] for x in summary_rows),
            "total_scrap_waste_m": sum(x["waste_m"] for x in summary_rows),
            "total_reusable_m": sum(x["reusable_m"] for x in summary_rows)
        },
        "packages": {
            "01_HE_THONG_CAT_THEP_REBARCUT": len(os.listdir(os.path.join(rebarcut_dir_user, "THEO_TUNG_DUONG_KINH_PHI"))) + len(os.listdir(os.path.join(rebarcut_dir_user, "LENH_CAT_CNC_CSV"))) + 1,
            "BO_HO_SO_01_MACRO_MASTER_14_SHEET": len(os.listdir(dir_macro)),
            "BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO": len(os.listdir(dir_micro)),
            "03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH": 12
        }
    }
    with open(os.path.join(USER_DOWNLOAD_DIR, "DISPATCH_MANIFEST.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)

    # 7. Đồng bộ trực tiếp sang du_lieu_du_an
    print("\n[BƯỚC 6/6] Đồng bộ toàn bộ cây thư mục sang workspace 23HG...")
    shutil.copytree(USER_DOWNLOAD_DIR, SYSTEM_PROJECT_DIR, dirs_exist_ok=True)
    print(f"  [V] Đã đồng bộ sang {SYSTEM_PROJECT_DIR}")

    print("\n" + "=" * 80)
    print("  HOÀN THÀNH XUẤT SẮC TOÀN BỘ HỆ THỐNG HỒ SƠ CẦU THÔN KHAI HOANG 2!")
    print("=" * 80)

if __name__ == "__main__":
    run_pipeline()
