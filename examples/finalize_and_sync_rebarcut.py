# -*- coding: utf-8 -*-
"""
ĐỒNG BỘ VÀ XUẤT TOÀN BỘ FILE HỆ THỐNG REBARCUT V6.0 CHO CẦU THÔN KHAI HOANG 2
"""
import os
import shutil
import win32com.client

CODE_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
SRC_XLSM = os.path.join(CODE_ROOT, r"RebarCut_V6\RebarCut_Pro_Excel_V6_Cau_Khai_Hoang_2.xlsm")

TARGETS = [
    os.path.join(os.path.expanduser("~"), "Downloads", r"Cầu thôn Khai Hoang 2, Km 14+363.65_Marker\RebarCut_Pro_Excel_V6_Cau_Khai_Hoang_2.xlsm"),
    os.path.join(os.path.expanduser("~"), "Downloads", r"Cầu thôn Khai Hoang 2, Km 14+363.65_Marker\01_HE_THONG_CAT_THEP_REBARCUT\RebarCut_Pro_Excel_V6_Cau_Khai_Hoang_2.xlsm"),
    os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), r"du_lieu_du_an\Cầu thôn Khai Hoang 2, Km 14+363.65_Marker\RebarCut_Pro_Excel_V6_Cau_Khai_Hoang_2.xlsm")
]

CSV_FOLDER = os.path.join(os.path.expanduser("~"), "Downloads", r"Cầu thôn Khai Hoang 2, Km 14+363.65_Marker\01_HE_THONG_CAT_THEP_REBARCUT\CSV_XUAT_XUONG")
CSV_FOLDER_PROJECT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), r"du_lieu_du_an\Cầu thôn Khai Hoang 2, Km 14+363.65_Marker\01_HE_THONG_CAT_THEP_REBARCUT\CSV_XUAT_XUONG")

def sync_all():
    print("[1] Sao chép file RebarCut_Pro_Excel_V6_Cau_Khai_Hoang_2.xlsm đến các thư mục đích...")
    for tgt in TARGETS:
        os.makedirs(os.path.dirname(tgt), exist_ok=True)
        shutil.copy2(SRC_XLSM, tgt)
        print(f"  [+] Đã lưu: {tgt}")

    os.makedirs(CSV_FOLDER, exist_ok=True)
    os.makedirs(CSV_FOLDER_PROJECT, exist_ok=True)

    print("\n[2] Xuất 8 file CSV UTF-8 từ workbook...")
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(os.path.abspath(TARGETS[0]))
        num_csv = xl.Run("ExportTo", os.path.abspath(CSV_FOLDER))
        print(f"  [+] Đã xuất {num_csv} file CSV tại: {CSV_FOLDER}")
        
        # Copy CSV sang cả project folder
        for f in os.listdir(CSV_FOLDER):
            if f.endswith(".csv"):
                shutil.copy2(os.path.join(CSV_FOLDER, f), os.path.join(CSV_FOLDER_PROJECT, f))
        print(f"  [+] Đã đồng bộ CSV sang: {CSV_FOLDER_PROJECT}")
        wb.Close(SaveChanges=False)
    except Exception as e:
        print(f"  [-] Lỗi xuất CSV: {e}")
    finally:
        xl.Quit()

    print("\n[3] Tổng hợp danh mục file hệ thống:")
    for tgt in TARGETS:
        size_kb = os.path.getsize(tgt) / 1024.0
        print(f"  - {os.path.basename(tgt)}: {size_kb:.1f} KB -> {tgt}")

if __name__ == "__main__":
    sync_all()
