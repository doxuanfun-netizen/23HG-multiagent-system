'=======================================================================================
' 23HG SYSTEM - UTILITIES & PROJECT MANAGEMENT EXTENSIONS
' Tac gia: NBT (Nguyen Bao Tu) | GitHub: @baotuhg | 23HG SYSTEM PRO
' Module: M_23HG_Utilities
'=======================================================================================
Option Explicit

Private Const APP_NAME As String = "23HG SYSTEM PRO"
Private Const APP_AUTHOR As String = "NBT (Nguyen Bao Tu)"

Private Function TargetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    If Not ActiveWorkbook Is Nothing Then
        Set TargetSheet = ActiveWorkbook.Sheets(sheetName)
    End If
    If TargetSheet Is Nothing Then
        Set TargetSheet = ThisWorkbook.Sheets(sheetName)
    End If
    On Error GoTo 0
End Function

' 1. THEM CONG TAC MOI
Public Sub ThemCongTac(Optional ByVal control As Object = Nothing)
    Dim ws As Worksheet: Set ws = TargetSheet("TIEN_DO")
    If ws Is Nothing Then Set ws = ActiveSheet
    If ws Is Nothing Or ws.Name <> "TIEN_DO" Then
        MsgBox "Vui long mo sheet TIEN_DO de thuc hien thao tac nay!", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim r As Long
    r = Selection.Row
    If r < 6 Then r = 6
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If r > lastRow Then r = lastRow

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ws.Rows(r + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    ' Copy format from row r
    ws.Rows(r).Copy
    ws.Rows(r + 1).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False

    ' Set values for new task
    Dim parentWBS As String: parentWBS = Trim$(CStr(ws.Cells(r, "B").Value))
    Dim newWBS As String
    If InStr(parentWBS, ".") > 0 Then
        newWBS = parentWBS & ".1"
    Else
        newWBS = parentWBS & ".1"
    End If

    ws.Cells(r + 1, "B").Value = newWBS
    ws.Cells(r + 1, "C").Value = "Cong tac moi (Bo sung)"
    ws.Cells(r + 1, "D").Value = "Task"
    ws.Cells(r + 1, "E").Value = parentWBS & "FS"
    ws.Cells(r + 1, "F").Value = 10
    ws.Cells(r + 1, "G").Value = 1#
    ws.Cells(r + 1, "H").Value = ws.Cells(r, "I").Value
    ws.Cells(r + 1, "I").Value = ws.Cells(r, "I").Value
    ws.Cells(r + 1, "J").Formula = "=I" & (r + 1) & "-H" & (r + 1) & "+1"
    ws.Cells(r + 1, "K").Value = 0
    ws.Cells(r + 1, "L").Value = ""
    ws.Cells(r + 1, "M").Value = 0#

    ' Re-number STT column A
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    Dim i As Long, sttCounter As Long: sttCounter = 1
    For i = 6 To lastRow
        If Trim$(CStr(ws.Cells(i, "D").Value)) <> "" Then
            ws.Cells(i, "A").Value = sttCounter
            sttCounter = sttCounter + 1
        End If
    Next i

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ' Recalculate CPM
    CalculateCPM ws, False
    RenderGantt ws, "QUY"

    MsgBox "Da them cong tac moi thanh cong tai dong " & (r + 1) & "!" & vbCrLf & _
           "Mang tien do da duoc tinh toan lai tu dong theo thuat toan CPM.", vbInformation, APP_NAME
End Sub

' 2. XOA CONG TAC
Public Sub XoaCongTac(Optional ByVal control As Object = Nothing)
    Dim ws As Worksheet: Set ws = TargetSheet("TIEN_DO")
    If ws Is Nothing Then Set ws = ActiveSheet
    If ws Is Nothing Or ws.Name <> "TIEN_DO" Then
        MsgBox "Vui long mo sheet TIEN_DO de thuc hien thao tac nay!", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim r As Long: r = Selection.Row
    If r < 6 Then
        MsgBox "Vui long chon dong cong tac can xoa (tu dong 6 tro di)!", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim taskName As String: taskName = Trim$(CStr(ws.Cells(r, "C").Value))
    Dim ans As VbMsgBoxResult
    ans = MsgBox("Ban co chac chan muon xoa cong tac sau khoi tien do?" & vbCrLf & _
                 "• Dong: " & r & vbCrLf & _
                 "• WBS: " & ws.Cells(r, "B").Value & vbCrLf & _
                 "• Ten: " & taskName, vbYesNo + vbQuestion, APP_NAME)
    If ans <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ws.Rows(r).Delete

    ' Re-number STT column A
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    Dim i As Long, sttCounter As Long: sttCounter = 1
    For i = 6 To lastRow
        If Trim$(CStr(ws.Cells(i, "D").Value)) <> "" Then
            ws.Cells(i, "A").Value = sttCounter
            sttCounter = sttCounter + 1
        End If
    Next i

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    CalculateCPM ws, False
    RenderGantt ws, "QUY"

    MsgBox "Da xoa cong tac va cap nhat lai mang tien do CPM!", vbInformation, APP_NAME
End Sub

' 3. THUT LE WBS (INDENT)
Public Sub IndentWBS(Optional ByVal control As Object = Nothing)
    Dim ws As Worksheet: Set ws = TargetSheet("TIEN_DO")
    If ws Is Nothing Then Set ws = ActiveSheet
    If ws Is Nothing Or ws.Name <> "TIEN_DO" Then Exit Sub

    Dim r As Long: r = Selection.Row
    If r >= 6 Then
        ws.Cells(r, "C").IndentLevel = ws.Cells(r, "C").IndentLevel + 1
    End If
End Sub

' 4. GIAM THUT LE WBS (OUTDENT)
Public Sub OutdentWBS(Optional ByVal control As Object = Nothing)
    Dim ws As Worksheet: Set ws = TargetSheet("TIEN_DO")
    If ws Is Nothing Then Set ws = ActiveSheet
    If ws Is Nothing Or ws.Name <> "TIEN_DO" Then Exit Sub

    Dim r As Long: r = Selection.Row
    If r >= 6 Then
        If ws.Cells(r, "C").IndentLevel > 0 Then
            ws.Cells(r, "C").IndentLevel = ws.Cells(r, "C").IndentLevel - 1
        End If
    End If
End Sub

' 5. XEM LICH & NGAY NGHI LE
Public Sub NgayNghiLe(Optional ByVal control As Object = Nothing)
    Dim ws As Worksheet: Set ws = TargetSheet("NGAY_NGHI_LE")
    If Not ws Is Nothing Then
        ws.Activate
        ws.Range("A1").Select
    Else
        MsgBox "Khong tim thay sheet NGAY_NGHI_LE!", vbExclamation, APP_NAME
    End If
End Sub

' 6. THONG TIN DU AN
Public Sub ThongTinDuAn(Optional ByVal control As Object = Nothing)
    Dim wsTD As Worksheet: Set wsTD = TargetSheet("TIEN_DO")
    Dim msg As String
    msg = "================================================================" & vbCrLf & _
          "       THONG TIN TONG QUAN DU AN & GOI THAU - 23HG SYSTEM        " & vbCrLf & _
          "================================================================" & vbCrLf & vbCrLf & _
          "• Ten du an: XAY DUNG TUYEN DUONG GIAO THONG (Km 0+000 - Km 5+500)" & vbCrLf & _
          "• Chieu dai toan tuyen: L = 5.50 km (Phan doan A1 + Phan doan A2)" & vbCrLf & _
          "• Chu dau tu: Ban Quan ly Du an Dau tu Xay dung Cong trinh Giao thong" & vbCrLf & _
          "• Goi thau: Goi thau so 01 - Thi cong xay lap toan tuyen G1" & vbCrLf & vbCrLf
    If Not wsTD Is Nothing Then
        msg = msg & "THONG SO TIEN DO HIEN TAI:" & vbCrLf & _
              "• Ngay khoi cong (F2): " & wsTD.Range("F2").Text & vbCrLf & _
              "• Ngay hoan thanh (H2): " & wsTD.Range("H2").Text & vbCrLf & _
              "• Lich lam viec: " & wsTD.Range("F3").Text & vbCrLf & _
              "• Che do tinh: " & wsTD.Range("H3").Text & vbCrLf & _
              "• Tong so cong tac: 37 hang muc wbs" & vbCrLf & vbCrLf
    End If
    msg = msg & "CAN CU PHAP LY & DINH MUC:" & vbCrLf & _
          "- Luat Xay dung 135/2025/QH15 & Nghi dinh 206/2026/ND-CP" & vbCrLf & _
          "- Nghi dinh 207/2026/ND-CP ve quan ly chat luong & nghiem thu" & vbCrLf & _
          "- Dinh muc ca may Thong tu 38/2026/TT-BXD, BoQ TT 36/2026/TT-BXD" & vbCrLf & _
          "- Quy chuan TCVN 8819:2011 (BTN) & TCVN 8859:2011 (CPDD)"
    MsgBox msg, vbInformation, "Thong Tin Du An - 23HG SYSTEM"
End Sub

' 7. IN TIEN DO (PRINT PREVIEW A3)
Public Sub InTienDo(Optional ByVal control As Object = Nothing)
    Dim ws As Worksheet: Set ws = TargetSheet("TIEN_DO")
    If ws Is Nothing Then Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA3
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$1:$5"
    End With
    ws.PrintPreview
End Sub

' 8. CHUAN HOA GIAO DIEN
Public Sub ChuanHoaGiaoDien(Optional ByVal control As Object = Nothing)
    Dim ws As Worksheet: Set ws = TargetSheet("TIEN_DO")
    If ws Is Nothing Then Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    ws.Activate
    ActiveWindow.DisplayGridlines = True
    ActiveWindow.Zoom = 85
    
    ' Freeze panes at D6 (Pin rows 1-5 and cols A-C)
    On Error Resume Next
    ActiveWindow.FreezePanes = False
    ws.Range("D6").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
    
    ' Re-fit critical column widths
    ws.Columns("A").ColumnWidth = 6
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 52
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 24
    ws.Columns("F").ColumnWidth = 16
    ws.Columns("G").ColumnWidth = 12
    ws.Columns("H").ColumnWidth = 14
    ws.Columns("I").ColumnWidth = 14
    ws.Columns("J").ColumnWidth = 14
    ws.Columns("K").ColumnWidth = 12
    ws.Columns("L").ColumnWidth = 12
    ws.Columns("M").ColumnWidth = 16
    
    Application.ScreenUpdating = True
    MsgBox "Da chuan hoa toan bo giao dien: Dong bang tieu de (Freeze Panes D6), Bat luoi o, can Zoom 85% va toi uu do rong cac cot!", vbInformation, APP_NAME
End Sub

' 9. TRO GIUP & HUONG DAN
Public Sub TroGiup(Optional ByVal control As Object = Nothing)
    Dim ws As Worksheet: Set ws = TargetSheet("HUONG_DAN")
    If Not ws Is Nothing Then
        ws.Activate
        ws.Range("A1").Select
    Else
        MsgBox "Khong tim thay sheet HUONG_DAN!", vbExclamation, APP_NAME
    End If
End Sub

' 10. TINH NGAY TU BOQ
Public Sub TinhNgayTuBoQ(Optional ByVal control As Object = Nothing)
    Dim wsBOQ As Worksheet: Set wsBOQ = TargetSheet("BOQ_TIEN_DO")
    Dim wsTD As Worksheet: Set wsTD = TargetSheet("TIEN_DO")
    If wsBOQ Is Nothing Or wsTD Is Nothing Then
        MsgBox "Can mo ca sheet BOQ_TIEN_DO va TIEN_DO de thuc hien!", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim lastBOQ As Long: lastBOQ = wsBOQ.Cells(wsBOQ.Rows.Count, "A").End(xlUp).Row
    Dim lastTD As Long: lastTD = wsTD.Cells(wsTD.Rows.Count, "A").End(xlUp).Row
    
    Dim updatedCount As Long: updatedCount = 0
    Dim rB As Long, rT As Long
    Dim wbsB As String, durB As Double
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For rB = 5 To lastBOQ
        wbsB = Trim$(CStr(wsBOQ.Cells(rB, "H").Value))
        durB = Val(wsBOQ.Cells(rB, "J").Value)
        If Len(wbsB) > 0 And durB > 0 Then
            For rT = 6 To lastTD
                If Trim$(CStr(wsTD.Cells(rT, "B").Value)) = wbsB Then
                    wsTD.Cells(rT, "F").Value = durB
                    updatedCount = updatedCount + 1
                    Exit For
                End If
            Next rT
        End If
    Next rB

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If updatedCount > 0 Then
        CalculateCPM wsTD, False
        RenderGantt wsTD, "QUY"
        MsgBox "Da dong bo thanh cong " & updatedCount & " cong tac tu BoQ sang bang Tien do!" & vbCrLf & _
               "Thoi luong = ROUNDUP(Khoi luong / Nang suat ca may, 0)." & vbCrLf & _
               "Toan bo tien do CPM da duoc cap nhat tu dong!", vbInformation, APP_NAME
    Else
        MsgBox "Khong tim thay ma WBS tuong ung giua BOQ va TIEN_DO!", vbExclamation, APP_NAME
    End If
End Sub

' 11. CAP NHAT DINH MUC (MO SHEET DB_DINH_MUC)
Public Sub CapNhatDM(Optional ByVal control As Object = Nothing)
    Dim ws As Worksheet: Set ws = TargetSheet("DB_DINH_MUC")
    If Not ws Is Nothing Then
        ws.Activate
        ws.Range("A4").Select
    Else
        MsgBox "Khong tim thay sheet DB_DINH_MUC!", vbExclamation, APP_NAME
    End If
End Sub

' 12. DOI CHIEU HO SO (MO SHEET KE_HOACH_QLCL)
Public Sub DoiChieuHoSo(Optional ByVal control As Object = Nothing)
    Dim ws As Worksheet: Set ws = TargetSheet("KE_HOACH_QLCL")
    If Not ws Is Nothing Then
        ws.Activate
        ws.Range("A4").Select
        MsgBox "Bang Ke hoach Quan ly Chat luong (QLCL) da duoc dong bo 100% voi Tien do!" & vbCrLf & _
               "Toan bo ngay bat dau & ket thuc nghiem thu lay truc tiep tu cong thuc =TIEN_DO!H... va =TIEN_DO!I...", vbInformation, APP_NAME
    Else
        MsgBox "Khong tim thay sheet KE_HOACH_QLCL!", vbExclamation, APP_NAME
    End If
End Sub
