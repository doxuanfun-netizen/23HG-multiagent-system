# -*- coding: utf-8 -*-
import os, sys, shutil, zipfile
import win32com.client

base_dir = os.path.abspath('.')
test_xlsm = os.path.join(base_dir, 'test_build.xlsm')
vba_file = os.path.join(base_dir, 'core_engine_clean.vba')
xml_file = os.path.join(base_dir, 'customUI14_exact.xml')
master_xlsm = os.path.join(base_dir, '23HG_DU_AN_MAU_TIEN_DO_CHUAN_G1_PRO.xlsm')
addin_xlam = os.path.join(base_dir, '23HG_Schedule_Assistant_Pro.xlam')

print("1. Cập nhật mã nguồn VBA vào test_build.xlsm...")
xl = win32com.client.Dispatch('Excel.Application')
xl.Visible = False
xl.DisplayAlerts = False

wb = xl.Workbooks.Open(test_xlsm)

# Replace M_23HG_CoreEngine
for c in list(wb.VBProject.VBComponents):
    if c.Name == 'M_23HG_CoreEngine':
        wb.VBProject.VBComponents.Remove(c)

c_new = wb.VBProject.VBComponents.Add(1)
c_new.Name = 'M_23HG_CoreEngine'
with open(vba_file, 'r', encoding='utf-8') as f:
    c_new.CodeModule.AddFromString(f.read())

# Keep Sheet1 clean
s1 = wb.VBProject.VBComponents('Sheet1')
if s1.CodeModule.CountOfLines > 0:
    s1.CodeModule.DeleteLines(1, s1.CodeModule.CountOfLines)

wb.Save()
wb.Close(False)
xl.Quit()
del wb
del xl
import gc, time
gc.collect()
time.sleep(1.5)
print("-> Đã nạp VBA thành công!")

# 2. Inject customUI14_exact.xml into test_build.xlsm
print("2. Đóng gói Ribbon XML chuẩn SCHEDULE ASSISTANT...")
with open(xml_file, 'rb') as f:
    new_xml_bytes = f.read()

temp_zip = test_xlsm + ".temp.zip"
if os.path.exists(temp_zip):
    os.remove(temp_zip)

with zipfile.ZipFile(test_xlsm, 'r') as zin:
    with zipfile.ZipFile(temp_zip, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == 'customUI/customUI14.xml':
                zout.writestr(item, new_xml_bytes)
            else:
                zout.writestr(item, zin.read(item.filename))

time.sleep(0.5)
if os.path.exists(master_xlsm):
    os.remove(master_xlsm)
shutil.move(temp_zip, master_xlsm)
if os.path.exists(test_xlsm):
    os.remove(test_xlsm)
print(f"-> Đã tạo Master Workbook hoàn hảo tại: {master_xlsm}")

# 3. Create Add-in xlam from master_xlsm
print("3. Xuất tệp Add-in 23HG_Schedule_Assistant_Pro.xlam...")
xl = win32com.client.Dispatch('Excel.Application')
xl.Visible = False
xl.DisplayAlerts = False

wb = xl.Workbooks.Open(master_xlsm)
if os.path.exists(addin_xlam):
    try:
        os.remove(addin_xlam)
    except Exception as e:
        print("Note removing old xlam:", e)

wb.SaveAs(Filename=addin_xlam, FileFormat=55) # 55 = xlOpenXMLAddIn
wb.Close(False)
xl.Quit()

# 4. Copy to %APPDATA%\Microsoft\AddIns\
addin_dest = os.path.expandvars(r'%APPDATA%\Microsoft\AddIns\23HG_Schedule_Assistant_Pro.xlam')
shutil.copy2(addin_xlam, addin_dest)
print(f"-> Đã sao chép Add-in vào thư mục chuẩn: {addin_dest}")

# 5. Activate Add-in in Excel
print("5. Kích hoạt Add-in trong Microsoft Excel...")
xl = win32com.client.Dispatch('Excel.Application')
xl.Visible = False
xl.DisplayAlerts = False
for ai in xl.AddIns:
    if '23HG' in ai.Name:
        ai.Installed = True
        print(f"-> Add-in {ai.Name} đã được kích hoạt: Installed = {ai.Installed}")
xl.Quit()

print("HOÀN TẤT 100%! Giao diện Ribbon SCHEDULE ASSISTANT và Master Workbook đã sẵn sàng.")
