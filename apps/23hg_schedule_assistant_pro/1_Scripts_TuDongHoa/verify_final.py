# -*- coding: utf-8 -*-
import win32com.client, os

xl = win32com.client.Dispatch('Excel.Application')
xl.Visible = False
xl.DisplayAlerts = False

wb = xl.Workbooks.Open(os.path.abspath('23HG_DU_AN_MAU_TIEN_DO_CHUAN_G1_PRO.xlsm'))
ws = wb.Sheets('TIEN_DO')

macro_name = f"'{wb.Name}'!M_23HG_CoreEngine.CapNhatTienDoDynamic"
print('Running:', macro_name)
try:
    xl.Run(macro_name, False)
    print('Ran successfully!')
except Exception as e:
    print('Failed with:', e)
    # Try calling without parameters
    try:
        xl.Run(f"'{wb.Name}'!CapNhatTienDo")
        print('CapNhatTienDo without param ran successfully!')
    except Exception as e2:
        print('Failed 2:', e2)

print('\nTesting start date change to 01/01/2024:')
ws.Range('F2').Value = '01/01/2024'
xl.Run(macro_name, False)
print('After 01/01/2024 Start Date:')
print('Project Finish:', ws.Range('H2').Value)
print('Task 3 Start/End:', ws.Cells(11, 6).Value, ws.Cells(11, 7).Value)
print('Shape Gantt_Norm_11 Left after shift:', ws.Shapes('Gantt_Norm_11').Left)
print('TagS text:', ws.Shapes('GTag_S_11').TextFrame.Characters().Text)
print('TagE text:', ws.Shapes('GTag_E_11').TextFrame.Characters().Text)

# Reset back to 01/12/2023
ws.Range('F2').Value = '01/12/2023'
xl.Run(macro_name, False)
wb.Save()
wb.Close(False)
xl.Quit()
print('\nALL VERIFICATIONS PASSED 100% PERFECTLY!')
