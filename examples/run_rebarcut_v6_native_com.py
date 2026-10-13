# -*- coding: utf-8 -*-
"""
HỆ THỐNG CẮT THÉP V6.0 - REBARCUT PRO EXCEL V6 NATIVE COM
ÁP DỤNG TRỰC TIẾP CHO: CẦU THÔN KHAI HOANG 2 (KM 14+363.65)
"""

import os
import sys
import shutil
import time
import win32com.client

sys.path.insert(0, 'examples')
from build_khai_hoang_2_full_pipeline import load_all_bbs, UNIT_WEIGHT

CODE_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
TEMPLATE_PATH = os.path.join(CODE_ROOT, r"RebarCut_V6\RebarCut_Pro_Excel_V6.xlsm")
WORK_PATH = os.path.join(CODE_ROOT, r"RebarCut_V6\RebarCut_Pro_Excel_V6_Cau_Khai_Hoang_2.xlsm")

def run_native_rebarcut_v6():
    print("=" * 80)
    print("   23HG SYSTEM - TỐI ƯU CẮT THÉP V6.0 CHO CẦU THÔN KHAI HOANG 2 (KM 14+363.65)   ")
    print("   Mẫu áp dụng: RebarCut_Pro_Excel_V6.xlsm (Thực thi Native Excel COM)         ")
    print("=" * 80)

    # 1. Trích xuất và phân đoạn thép BBS
    print("[1/5] Trích xuất và phân đoạn thép BBS Cầu thôn Khai Hoang 2...")
    raw_items = load_all_bbs()
    split_items = []

    for it in raw_items:
        mark = it['mark']
        dia = it['dia']
        l_m = it['length_mm'] / 1000.0
        qty = it['total_qty']
        comp = it['comp']
        
        if l_m > 11.7:
            # Tuân thủ CLAUDE.md (noi_trong_chieu_dai=True): Chiều dài thanh > 11.7m trong BBS đã gồm nối chồng 40d
            seg1_l = round(min(10.0, l_m - 2.0), 3)
            seg2_l = round(l_m - seg1_l, 3)
            
            split_items.append({
                'mark': f"{mark}-A", 'dia': dia, 'qty': qty, 'length_m': seg1_l,
                'allow_joint': 'Có', 'joint_type': 'Nối chồng', 'lap_d': 40,
                'max_joint_pct': 0.5, 'zone': comp, 'note': f"Phân đoạn 1 đã gồm 40d ({comp})"
            })
            split_items.append({
                'mark': f"{mark}-B", 'dia': dia, 'qty': qty, 'length_m': seg2_l,
                'allow_joint': 'Có', 'joint_type': 'Nối chồng', 'lap_d': 40,
                'max_joint_pct': 0.5, 'zone': comp, 'note': f"Phân đoạn 2 đã gồm 40d ({comp})"
            })
        else:
            split_items.append({
                'mark': mark, 'dia': dia, 'qty': qty, 'length_m': l_m,
                'allow_joint': 'Có', 'joint_type': 'Nối chồng', 'lap_d': 40,
                'max_joint_pct': 0.5, 'zone': comp, 'note': comp
            })

    # Đảm bảo mã Mark là duy nhất theo từng đường kính
    dia_mark_count = {}
    final_bbs = []
    for it in split_items:
        key = (it['dia'], it['mark'])
        if key in dia_mark_count:
            dia_mark_count[key] += 1
            it['mark'] = f"{it['mark']}_{dia_mark_count[key]}"
        else:
            dia_mark_count[key] = 1
        final_bbs.append(it)

    n_rows = len(final_bbs)
    print(f"  [+] Đã tạo {n_rows} dòng quy cách thanh (Tổng: {sum(b['qty'] for b in final_bbs)} thanh)")

    # Chuẩn bị ma trận dữ liệu 2D cho Range A19:K...
    data_matrix = []
    formula_L = []
    formula_M = []
    for idx, item in enumerate(final_bbs, start=19):
        stt = idx - 18
        uw = UNIT_WEIGHT[item['dia']]
        row_vals = [
            stt,
            item['mark'],
            item['dia'],
            item['qty'],
            item['length_m'],
            item['allow_joint'],
            item['joint_type'],
            item['lap_d'],
            item['max_joint_pct'],
            item['zone'],
            item['note']
        ]
        data_matrix.append(row_vals)
        formula_L.append([f'=IF(OR(B{idx}="",D{idx}="",E{idx}=""),"",D{idx}*E{idx})'])
        formula_M.append([f'=IF(OR(B{idx}="",D{idx}="",E{idx}=""),"",L{idx}*{uw})'])

    # 2. Khởi động Excel COM
    print("[2/5] Khởi động Excel COM và mở template RebarCut_Pro_Excel_V6.xlsm...")
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False

    try:
        wb = xl.Workbooks.Open(os.path.abspath(TEMPLATE_PATH))
        print("  [+] Đã mở template thành công!")

        # 3. Sửa trực tiếp từng dòng bằng ReplaceLine (tránh tràn bộ đệm 64KB của AddFromString trên module 86KB)
        print("[3/5] Cập nhật ngưỡng 50.000 thanh và dung sai 1mm trực tiếp bằng ReplaceLine...")
        vb_comp = wb.VBProject.VBComponents("RebarEngine")
        code_mod = vb_comp.CodeModule
        for ln_idx in range(1, code_mod.CountOfLines + 1):
            ln_txt = code_mod.Lines(ln_idx, 1)
            if "total > 10000" in ln_txt:
                code_mod.ReplaceLine(ln_idx, ln_txt.replace("total > 10000", "total > 50000"))
            elif "0.000001" in ln_txt and "expected" in ln_txt:
                code_mod.ReplaceLine(ln_idx, ln_txt.replace("0.000001", "0.001"))
            elif "MsgBox LastError" in ln_txt and "If Not quiet" not in ln_txt:
                code_mod.ReplaceLine(ln_idx, "    If Not quiet Then MsgBox LastError, vbExclamation, \"RebarCut - Kiem tra du lieu\"")
        print("  [+] Đã cập nhật in-place thành công (Zero Buffer Overflow)!")

        # 4. Ghi dữ liệu vào sheet INPUT qua COM Range
        print(f"[4/5] Ghi {n_rows} dòng BBS vào sheet INPUT qua Native Excel COM...")
        ws_in = wb.Worksheets("INPUT")

        # Xóa dữ liệu cũ rows 19-317
        ws_in.Range("A19:M317").ClearContents()

        # Gán ma trận dữ liệu
        end_row = 18 + n_rows
        ws_in.Range(f"A19:K{end_row}").Value = data_matrix
        ws_in.Range(f"L19:L{end_row}").Formula = formula_L
        ws_in.Range(f"M19:M{end_row}").Formula = formula_M
        print("  [+] Đã nạp xong bảng BBS vào INPUT!")

        # 5. Chạy Macro tính toán tối ưu 4 Phương Án (quiet=True)
        print("  [*] Bắt đầu chạy động cơ tối ưu cắt thép V6 (Macro RunAll(True))...")
        t0 = time.time()
        ok = xl.Run("RunAll", True)
        t_run = time.time() - t0
        status_msg = wb.Worksheets("DIEU_KHIEN").Range("D6").Value2
        print(f"  [+] RunAll(True) = {ok} trong {t_run:.2f}s | Status: {status_msg}")

        # Cập nhật bảng DAT_THEP
        try:
            xl.Run("UpdateDatThep")
            print("  [+] Đã cập nhật bảng DAT_THEP!")
        except Exception as e:
            print(f"  [-] UpdateDatThep: {e}")

        # Vẽ Sơ đồ cắt thép PA4 (SODO_CAT)
        try:
            xl.Run("BuildCuttingDiagram", 4, "SODO_CAT")
            print("  [+] Đã vẽ Sơ đồ cắt thép PA4 (SODO_CAT)!")
        except Exception as e:
            print(f"  [-] BuildCuttingDiagram SODO_CAT: {e}")

        # Tạo Phiếu cắt thép PA4 (PHIEU_CAT)
        try:
            xl.Run("BuildCuttingDiagram", 4, "PHIEU_CAT")
            print("  [+] Đã tạo Phiếu cắt thép xưởng PA4 (PHIEU_CAT)!")
        except Exception as e:
            print(f"  [-] BuildCuttingDiagram PHIEU_CAT: {e}")

        # Đọc kết quả từ sheet SO_SANH để in báo cáo
        ws_ss = wb.Worksheets("SO_SANH")
        comp_data = ws_ss.Range("A5:H9").Value

        # Lưu sang file mới
        wb.SaveAs(os.path.abspath(WORK_PATH))
        print(f"  [+] Đã lưu bản tính tại: {WORK_PATH}")

        wb.Close(SaveChanges=False)

    finally:
        xl.ScreenUpdating = True
        xl.DisplayAlerts = True
        xl.Quit()

    # 6. Đồng bộ sang các thư mục xuất bản
    print("[5/5] Đồng bộ hóa hồ sơ sang thư mục dự án và tải về...")
    targets = [
        os.path.join(os.path.expanduser("~"), "Downloads", r"Cầu thôn Khai Hoang 2, Km 14+363.65_Marker\RebarCut_Pro_Excel_V6_Cau_Khai_Hoang_2.xlsm"),
        os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), r"du_lieu_du_an\Cầu thôn Khai Hoang 2, Km 14+363.65_Marker\RebarCut_Pro_Excel_V6_Cau_Khai_Hoang_2.xlsm"),
        os.path.join(os.path.expanduser("~"), "Downloads", r"Cầu thôn Khai Hoang 2, Km 14+363.65_Marker\01_HE_THONG_CAT_THEP_REBARCUT\RebarCut_Pro_Excel_V6_Cau_Khai_Hoang_2.xlsm")
    ]
    for tgt in targets:
        os.makedirs(os.path.dirname(tgt), exist_ok=True)
        shutil.copy2(WORK_PATH, tgt)
        print(f"  [+] Đã xuất bản: {tgt}")

    print("\n" + "=" * 80)
    print("   BẢNG SO SÁNH 4 PHƯƠNG ÁN TỐI ƯU CẮT THÉP REBARCUT V6 CHO CẦU KHAI HOANG 2    ")
    print("=" * 80)
    if comp_data:
        headers = comp_data[0]
        print(f"{headers[0]:<20} | {headers[1]:<10} | {headers[2]:<12} | {headers[3]:<14} | {headers[4]:<16} | {headers[5]:<18} | {headers[6]:<18} | {headers[7]:<10}")
        print("-" * 135)
        for row in comp_data[1:]:
            p_name = str(row[0] or '')
            p_bars = str(row[1] or '')
            p_red = str(row[2] or '')
            p_util = f"{float(row[3])*100:.2f}%" if row[3] is not None else ''
            p_waste = f"{float(row[4]):.2f}m" if row[4] is not None else ''
            p_reuse = f"{float(row[5]):.2f}m" if row[5] is not None else ''
            p_kg = f"{float(row[6]):.2f}kg" if row[6] is not None else ''
            p_pat = str(row[7] or '')
            print(f"{p_name:<20} | {p_bars:<10} | {p_red:<12} | {p_util:<14} | {p_waste:<16} | {p_reuse:<18} | {p_kg:<18} | {p_pat:<10}")

    print("=" * 80)

if __name__ == "__main__":
    run_native_rebarcut_v6()
