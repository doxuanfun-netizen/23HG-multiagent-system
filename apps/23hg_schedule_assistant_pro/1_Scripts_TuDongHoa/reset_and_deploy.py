# -*- coding: utf-8 -*-
import win32com.client
import os
import shutil

wb_path = os.path.abspath(r"23HG_DU_AN_MAU_TIEN_DO_CHUAN_G1_PRO.xlsm")
addin_path = os.path.abspath(r"23HG_Schedule_Assistant_Pro.xlam")

xl = win32com.client.Dispatch("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False

try:
    wb = xl.Workbooks.Open(wb_path)
    ws = wb.Sheets("TIEN_DO")

    # Set exact baseline serial date (45261 = 01/12/2023)
    ws.Range("F2").Value2 = 45261
    ws.Range("E11").Value = 25
    xl.Run(f"'{wb.Name}'!CalculateCPM", ws, False)
    xl.Run(f"'{wb.Name}'!RenderGantt", ws, "QUY")

    print(f"F2 Khởi công: {ws.Range('F2').Text}")
    print(f"H2 Kết thúc: {ws.Range('H2').Text}")
    print(f"Row 9 Gói thầu: {ws.Range('F9').Text} -> {ws.Range('G9').Text} | Ngày: {ws.Range('E9').Value}")
    print(f"Row 11 Task 1.1.1: {ws.Range('F11').Text} -> {ws.Range('G11').Text}")

    # Đồng bộ sang QLCL
    ws_qlcl = wb.Sheets("KE_HOACH_QLCL")
    print(f"QLCL NT-001 (D5, E5): {ws_qlcl.Range('D5').Text} -> {ws_qlcl.Range('E5').Text}")
    print(f"QLCL GĐ-001 (D15, E15): {ws_qlcl.Range('D15').Text} -> {ws_qlcl.Range('E15').Text}")

    # Check EVM BAC and dates
    ws_evm = wb.Sheets("EVM_5D_QUAN_TRI")
    print(f"EVM BAC: {ws_evm.Range('B5').Text} | PV: {ws_evm.Range('D5').Text} | EV: {ws_evm.Range('F5').Text}")

    wb.Save()
    wb.SaveAs(addin_path, 55)

    # Deploy to system folders
    appdata = os.environ.get("APPDATA", "")
    addins_dir = os.path.join(appdata, "Microsoft", "AddIns")
    xlstart_dir = os.path.join(appdata, "Microsoft", "Excel", "XLSTART")
    os.makedirs(addins_dir, exist_ok=True)
    os.makedirs(xlstart_dir, exist_ok=True)
    shutil.copy2(addin_path, os.path.join(addins_dir, "23HG_Schedule_Assistant_Pro.xlam"))
    shutil.copy2(addin_path, os.path.join(xlstart_dir, "23HG_Schedule_Assistant_Pro.xlam"))
    print("\n-> TRIỂN KHAI ADD-IN THÀNH CÔNG VÀO HỆ THỐNG!")

finally:
    wb.Close(False)
    xl.Quit()

print("HOÀN TẤT 100%!")
