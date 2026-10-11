import win32com.client
import os

xl = win32com.client.Dispatch('Excel.Application')
xl.Visible = False
xl.DisplayAlerts = False

wb_path = os.path.abspath(r'23HG_DU_AN_MAU_TIEN_DO_CHUAN_G1_PRO.xlsm')
wb = xl.Workbooks.Open(wb_path)
ws = wb.Sheets('TIEN_DO')

out_xml = os.path.abspath('2_BaoCao_XuatBan/test_vba_export.xml')
out_xer = os.path.abspath('2_BaoCao_XuatBan/test_vba_export.xer')
if os.path.exists(out_xml): os.remove(out_xml)
if os.path.exists(out_xer): os.remove(out_xer)

macro_xml = f"'{wb.Name}'!ExportTasksToMSPDI"
res_xml = xl.Run(macro_xml, ws, out_xml)
print('VBA ExportTasksToMSPDI result:', res_xml)
print('XML exists and size:', os.path.exists(out_xml), os.path.getsize(out_xml) if os.path.exists(out_xml) else 0)

macro_import_xml = f"'{wb.Name}'!ImportTasksFromMSPDI"
res_import = xl.Run(macro_import_xml, ws, out_xml)
print('VBA ImportTasksFromMSPDI result:', res_import)

macro_xer = f"'{wb.Name}'!ExportTasksToXER"
res_xer = xl.Run(macro_xer, ws, out_xer)
print('VBA ExportTasksToXER result:', res_xer)
print('XER exists and size:', os.path.exists(out_xer), os.path.getsize(out_xer) if os.path.exists(out_xer) else 0)

wb.Close(False)
xl.Quit()

assert res_xml == True, "XML export failed"
assert res_import == True, "XML import failed"
assert res_xer == True, "XER export failed"

if os.path.exists(out_xml): os.remove(out_xml)
if os.path.exists(out_xer): os.remove(out_xer)

print("ALL VBA COM BRIDGE RUNTIME TESTS (EXPORT & IMPORT) PASSED 100%!")
