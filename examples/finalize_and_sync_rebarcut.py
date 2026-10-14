# -*- coding: utf-8 -*-
"""
ĐỒNG BỘ VÀ XUẤT TOÀN BỘ FILE HỆ THỐNG REBARCUT V6.0 CHO CẦU THÔN KHAI HOANG 2
"""
import os
import shutil
import sys

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from examples._paths import project_path

REBARCUT_DIR = os.environ.get("AEC_REBARCUT_DIR", project_path("RebarCut_V6"))
SRC_XLSM = os.path.join(REBARCUT_DIR, "RebarCut_Pro_Excel_V6_Cau_Khai_Hoang_2.xlsm")
PROJ_DIR = project_path("Cầu thôn Khai Hoang 2, Km 14+363.65_Marker")

TARGETS = [
    os.path.join(PROJ_DIR, "RebarCut_Pro_Excel_V6_Cau_Khai_Hoang_2.xlsm"),
    os.path.join(PROJ_DIR, "02_XUONG_TIEN_CHE_COT_THEP", "RebarCut_Pro_Excel_V6_Cau_Khai_Hoang_2.xlsm"),
]

CSV_FOLDER = os.path.join(PROJ_DIR, "02_XUONG_TIEN_CHE_COT_THEP", "CSV_XUAT_XUONG")
CSV_FOLDER_PROJECT = CSV_FOLDER

def sync_all():
    import win32com.client

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
