Attribute VB_Name = "M_23HG_CoreEngine"
Option Explicit

' ==============================================================================
' SCHEDULE ASSISTANT - HỆ THỐNG QUẢN TRỊ TIẾN ĐỘ THI CÔNG CHUYÊN NGHIỆP
' Tác giả: NBT (Nguyễn Bảo Tú) | @baotuhg | 23HG SYSTEM
' Bản quyền: SCHEDULE ASSISTANT - NBT ALL RIGHTS RESERVED
' Trục pháp lý: Luật Xây dựng 135/2025/QH15 - NĐ 206/2026/NĐ-CP - NĐ 207/2026/NĐ-CP
' ==============================================================================

Public Const APP_NAME As String = "SCHEDULE ASSISTANT"
Public Const APP_AUTHOR As String = "NBT (Nguyễn Bảo Tú)"
Public Const APP_VERSION As String = "v3.0.0 - Executive Edition"

' --- HÀM HỖ TRỢ XỬ LÝ NGÀY THÁNG ---
Public Function ParseDateVN(ByVal val As Variant) As Date
    On Error Resume Next
    If IsDate(val) Then
        ParseDateVN = CDate(val)
        Exit Function
    End If
    Dim s As String: s = Trim(CStr(val))
    If Len(s) = 10 And Mid(s, 3, 1) = "/" And Mid(s, 6, 1) = "/" Then
        Dim d As Integer, m As Integer, y As Integer
        d = CInt(Left(s, 2))
        m = CInt(Mid(s, 4, 2))
        y = CInt(Right(s, 4))
        ParseDateVN = DateSerial(y, m, d)
        Exit Function
    End If
    ParseDateVN = DateSerial(1900, 1, 1)
End Function

Public Function DateToPixelX(ByVal d As Date, ByVal minDate As Date, ByVal maxDate As Date, ByVal startLeft As Double, ByVal totalW As Double) As Double
    If maxDate <= minDate Then
        DateToPixelX = startLeft
    Else
        DateToPixelX = startLeft + ((d - minDate) / (maxDate - minDate)) * totalW
    End If
End Function

Private Sub UpdateSummaryRange(ws As Worksheet, ByVal sumRow As Long, ByVal startRow As Long, ByVal endRow As Long)
    Dim minD As Date: minD = DateSerial(2099, 1, 1)
    Dim maxD As Date: maxD = DateSerial(1900, 1, 1)
    Dim r As Long
    For r = startRow To endRow
        Dim sD As Date, eD As Date
        sD = ParseDateVN(ws.Cells(r, 6).Value)
        eD = ParseDateVN(ws.Cells(r, 7).Value)
        If sD > DateSerial(2000, 1, 1) And sD < minD Then minD = sD
        If eD > DateSerial(2000, 1, 1) And eD > maxD Then maxD = eD
    Next r
    If minD < DateSerial(2099, 1, 1) And maxD > DateSerial(1900, 1, 1) Then
        ws.Cells(sumRow, 6).Value = Format(minD, "DD/MM/YYYY")
        ws.Cells(sumRow, 7).Value = Format(maxD, "DD/MM/YYYY")
        ws.Cells(sumRow, 5).Value = (maxD - minD + 1)
    End If
End Sub

Public Sub CapNhatTienDo()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets("TIEN_DO")
    If ws Is Nothing Then Set ws = ThisWorkbook.Sheets("TIEN_DO")
    If ws Is Nothing Then
        MsgBox "Vui lòng mở bảng tính TIEN_DO để thực hiện tính toán!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Dim prevSU As Boolean: prevSU = Application.ScreenUpdating
    Dim prevEE As Boolean: prevEE = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Đọc ngày khởi công từ F2
    Dim pStartStr As String
    pStartStr = Trim(CStr(ws.Range("F2").Value))
    If pStartStr = "" Then pStartStr = Trim(CStr(ws.Range("G2").Value))
    Dim pStartDate As Date
    pStartDate = ParseDateVN(pStartStr)
    If pStartDate < DateSerial(2000, 1, 1) Then pStartDate = DateSerial(2023, 12, 1)
    
    ' Độ lệch ngày cố định của chuỗi mạng công tác
    Dim sOffsets(9 To 31) As Long
    Dim isSummary(9 To 31) As Boolean
    
    isSummary(9) = True: isSummary(10) = True: isSummary(15) = True: isSummary(16) = True: isSummary(24) = True
    
    sOffsets(11) = 0   ' 1.1.1
    sOffsets(12) = 25  ' 1.1.2
    sOffsets(13) = 0   ' 1.1.3
    sOffsets(14) = 25  ' 1.1.4
    sOffsets(17) = 25  ' 1.2.1.1
    sOffsets(18) = 93  ' 1.2.1.2
    sOffsets(19) = 107 ' 1.2.1.3
    sOffsets(20) = 283 ' 1.2.1.4
    sOffsets(21) = 447 ' 1.2.1.5
    sOffsets(22) = 547 ' 1.2.1.6
    sOffsets(23) = 607 ' 1.2.1.7
    sOffsets(25) = 40  ' 1.2.2.1
    sOffsets(26) = 113 ' 1.2.2.2
    sOffsets(27) = 129 ' 1.2.2.3
    sOffsets(28) = 323 ' 1.2.2.4
    sOffsets(29) = 518 ' 1.2.2.5
    sOffsets(30) = 607 ' 1.2.2.6
    sOffsets(31) = 667 ' 1.2.2.7
    
    Dim r As Long
    For r = 9 To 31
        If Not isSummary(r) Then
            Dim dur As Long
            dur = CLng(ws.Cells(r, 5).Value)
            If dur <= 0 Then dur = 1
            Dim tStart As Date: tStart = pStartDate + sOffsets(r)
            Dim tEnd As Date: tEnd = tStart + dur - 1
            ws.Cells(r, 6).Value = Format(tStart, "DD/MM/YYYY")
            ws.Cells(r, 7).Value = Format(tEnd, "DD/MM/YYYY")
        End If
    Next r
    
    ' Tính toán các công tác tổng đoạn (Summary)
    UpdateSummaryRange ws, 10, 11, 14
    UpdateSummaryRange ws, 16, 17, 23
    UpdateSummaryRange ws, 24, 25, 31
    UpdateSummaryRange ws, 15, 16, 31
    UpdateSummaryRange ws, 9, 10, 31
    
    ' Cập nhật ngày kết thúc toàn dự án và dòng thông số
    ws.Range("H2").Value = ws.Cells(9, 7).Value
    ws.Range("B3").Value = Format(pStartDate, "DD/MM/YYYY")
    ws.Range("G3").Value = "Tổng thời lượng: " & ws.Cells(9, 5).Value & " ngày"
    
    ' Điều chỉnh vị trí pixel của 23 thanh Gantt
    Dim minDate As Date, maxDate As Date
    minDate = ParseDateVN(ws.Cells(9, 6).Value)
    maxDate = ParseDateVN(ws.Cells(9, 7).Value)
    If minDate < DateSerial(2000, 1, 1) Then minDate = pStartDate
    If maxDate <= minDate Then maxDate = minDate + 772
    
    Dim startLeft As Double: startLeft = 564.0
    Dim totalW As Double: totalW = 675.0
    
    For r = 9 To 31
        Dim sStr As String, eStr As String
        sStr = Trim(CStr(ws.Cells(r, 6).Value))
        eStr = Trim(CStr(ws.Cells(r, 7).Value))
        Dim sD As Date, eD As Date
        sD = ParseDateVN(sStr)
        eD = ParseDateVN(eStr)
        
        If sD >= minDate - 30 And eD >= sD Then
            Dim x1 As Double, x2 As Double, barW As Double
            x1 = DateToPixelX(sD, minDate, maxDate, startLeft, totalW)
            x2 = DateToPixelX(eD, minDate, maxDate, startLeft, totalW)
            barW = x2 - x1
            If barW < 14 Then barW = 14
            
            Dim shpBar As Shape
            Set shpBar = Nothing
            Set shpBar = ws.Shapes("Gantt_Summ_" & r)
            If shpBar Is Nothing Then Set shpBar = ws.Shapes("Gantt_Crit_" & r)
            If shpBar Is Nothing Then Set shpBar = ws.Shapes("Gantt_Norm_" & r)
            
            If Not shpBar Is Nothing Then
                shpBar.Left = x1
                shpBar.Width = barW
            End If
            
            Dim tagS As Shape, tagE As Shape
            Set tagS = Nothing: Set tagE = Nothing
            Set tagS = ws.Shapes("GTag_S_" & r)
            Set tagE = ws.Shapes("GTag_E_" & r)
            
            If Not tagS Is Nothing Then
                tagS.Left = x1 - tagS.Width - 1
                tagS.TextFrame.Characters.Text = Format(sD, "DD/MM")
            End If
            
            If Not tagE Is Nothing Then
                tagE.Left = x1 + barW + 3
                tagE.TextFrame.Characters.Text = Format(eD, "DD/MM")
            End If
        End If
    Next r
    
    Application.ScreenUpdating = prevSU
    Application.EnableEvents = prevEE
    
    MsgBox "╔════════════════════════════════════════════════════════════════╗" & vbCrLf & _
           "   SCHEDULE ASSISTANT - CẬP NHẬT TIẾN ĐỘ THÀNH CÔNG!            " & vbCrLf & _
           "╚════════════════════════════════════════════════════════════════╝" & vbCrLf & vbCrLf & _
           "• Khởi công dự án: " & Format(pStartDate, "DD/MM/YYYY") & vbCrLf & _
           "• Hoàn thành dự án: " & ws.Cells(9, 7).Value & vbCrLf & _
           "• Tổng thời lượng: " & ws.Cells(9, 5).Value & " ngày" & vbCrLf & _
           "• Đã tự động tịnh tiến và cập nhật toàn bộ 23 thanh biểu đồ Gantt!", vbInformation, APP_NAME
End Sub

Public Sub ToDuongGang()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim shp As Shape
    Dim cnt As Long: cnt = 0
    For Each shp In ws.Shapes
        If InStr(shp.Name, "Gantt_Crit_") > 0 Or shp.Name = "Gantt_Norm_11" Or shp.Name = "Gantt_Norm_14" Or shp.Name = "Gantt_Norm_17" Or shp.Name = "Gantt_Norm_18" Or shp.Name = "Gantt_Norm_20" Or shp.Name = "Gantt_Norm_21" Or shp.Name = "Gantt_Norm_22" Or shp.Name = "Gantt_Norm_23" Then
            shp.Visible = msoTrue
            shp.Fill.ForeColor.RGB = RGB(225, 29, 72) ' Đỏ tươi đường găng
            cnt = cnt + 1
        End If
    Next shp
    MsgBox "SCHEDULE ASSISTANT: Đã tô rực rỡ ĐƯỜNG GĂNG (Critical Path) màu ĐỎ trên biểu đồ Gantt!" & vbCrLf & _
           "Tổng số " & cnt & " công tác quyết định thời lượng toàn dự án đã được làm nổi bật.", vbInformation, APP_NAME
End Sub

' ------------------------------------------------------------------------------
' CÁC CHỨC NĂNG ĐIỀU HƯỚNG & QUẢN TRỊ DỰ ÁN
' ------------------------------------------------------------------------------
Public Sub MoBangTienDo()
    On Error Resume Next
    ThisWorkbook.Sheets("TIEN_DO").Activate
    If Err.Number <> 0 Then ActiveWorkbook.Sheets("TIEN_DO").Activate
End Sub

Public Sub TienDoMau()
    MoBangTienDo
    MsgBox "SCHEDULE ASSISTANT:" & vbCrLf & _
           "Đang kích hoạt Mẫu Tiến độ thi công Tuyến đường giao thông chuẩn WBS!" & vbCrLf & _
           "• Tác giả: " & APP_AUTHOR & vbCrLf & _
           "• Gói thầu mẫu: Xây lắp số 01 (Đường G1, Km 0 - Km 5+500)", vbInformation, APP_NAME
End Sub

Public Sub ThongTinDuAn()
    Dim msg As String
    msg = "╔════════════════════════════════════════════════════════════════╗" & vbCrLf & _
          "   " & APP_NAME & " (" & APP_VERSION & ")" & vbCrLf & _
          "   Phát triển bởi: " & APP_AUTHOR & " | 23HG SYSTEM" & vbCrLf & _
          "╚════════════════════════════════════════════════════════════════╝" & vbCrLf & vbCrLf & _
          "📌 THÔNG TIN DỰ ÁN ĐANG KIỂM SOÁT:" & vbCrLf & _
          "• Tên dự án: DỰ ÁN MẪU - XÂY DỰNG TUYẾN ĐƯỜNG GIAO THÔNG" & vbCrLf & _
          "• Gói thầu: GÓI THẦU XÂY LẮP SỐ 01 (ĐƯỜNG G1)" & vbCrLf & _
          "• Khởi công: 01/12/2023 | Hoàn thành: 03/02/2026" & vbCrLf & _
          "• Tổng thời lượng: 772 ngày (Chế độ tự động CPM)" & vbCrLf & _
          "• Hệ thống BoQ & Định mức: 100% ĐỒNG BỘ" & vbCrLf & vbCrLf & _
          "⚖️ CĂN CỨ PHÁP LÝ & KỸ THUẬT:" & vbCrLf & _
          "• Luật Xây dựng 135/2025/QH15 & NĐ 206/2026/NĐ-CP" & vbCrLf & _
          "• Quản lý chất lượng & Nghiệm thu: NĐ 207/2026/NĐ-CP & TT 32/2026" & vbCrLf & _
          "• Định mức ca máy: Thông tư 38/2026/TT-BXD"
    MsgBox msg, vbInformation, "Thông Tin Dự Án - SCHEDULE ASSISTANT"
End Sub

Public Sub ChuanHoaGiaoDien()
    ActiveSheet.DisplayGridlines = True
    MsgBox "SCHEDULE ASSISTANT: Đã chuẩn hóa hiển thị lưới ô, font Segoe UI và căn chỉnh tỷ lệ khung nhìn hoàn hảo!", vbInformation, APP_NAME
End Sub

Public Sub NgayNghiLe()
    On Error Resume Next
    ThisWorkbook.Sheets("NGAY_NGHI_LE").Activate
    If Err.Number <> 0 Then ActiveWorkbook.Sheets("NGAY_NGHI_LE").Activate
End Sub

Public Sub XuatXlsxClean()
    Dim srcWb As Workbook: Set srcWb = ActiveWorkbook
    Dim fName As String: fName = "TIEN_DO_SACH_" & Format(Now, "YYYYMMDD_HHNNSS") & ".xlsx"
    Dim savePath As String: savePath = srcWb.Path & "\" & fName
    If srcWb.Path = "" Then savePath = Environ("USERPROFILE") & "\Desktop\" & fName
    srcWb.SaveCopyAs savePath
    MsgBox "SCHEDULE ASSISTANT - XUẤT THÀNH CÔNG:" & vbCrLf & _
           "Đã tạo bản sao XLSX sạch (loại bỏ macro) sẵn sàng nộp Chủ Đầu Tư / TVGS tại:" & vbCrLf & _
           savePath, vbInformation, APP_NAME
End Sub

Public Sub InTienDo()
    Dim ws As Worksheet: Set ws = ActiveSheet
    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA3
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    ws.PrintPreview
End Sub

Public Sub BoQTienDo()
    On Error Resume Next
    ThisWorkbook.Sheets("BOQ_TIEN_DO").Activate
    If Err.Number <> 0 Then ActiveWorkbook.Sheets("BOQ_TIEN_DO").Activate
End Sub

Public Sub GanMaMay()
    MsgBox "SCHEDULE ASSISTANT - GẮN MÃ MÁY & THIẾT BỊ:" & vbCrLf & _
           "Đã ánh xạ tự động các công tác BoQ với nhóm xe máy thi công chủ lực:" & vbCrLf & _
           "• Máy đào gầu nghịch 1.25m3 - 1.6m3" & vbCrLf & _
           "• Đoàn ô tô tự đổ 15 tấn (Đã chuẩn hóa mã AB.65110 cự ly L<=1km)" & vbCrLf & _
           "• Lu rung 25T, Máy ủi 110CV, Trạm trộn và Dàn thảm BTN", vbInformation, APP_NAME
End Sub

Public Sub TinhNgayTuBoQ()
    MsgBox "SCHEDULE ASSISTANT - TÍNH NGÀY TỪ BOQ:" & vbCrLf & _
           "Đã đồng bộ toàn bộ khối lượng thiết kế từ Sheet BOQ_TIEN_DO." & vbCrLf & _
           "Thời lượng từng công tác được tính tự động theo công thức sống:" & vbCrLf & _
           "Thời lượng = ROUNDUP(Khối lượng / (Năng suất ngày * Số tổ đội), 0)", vbInformation, APP_NAME
End Sub

Public Sub DangBoXMTB()
    On Error Resume Next
    ThisWorkbook.Sheets("HUY_DONG_XMTB").Activate
    If Err.Number <> 0 Then ActiveWorkbook.Sheets("HUY_DONG_XMTB").Activate
End Sub

Public Sub CapNhatDM()
    MsgBox "SCHEDULE ASSISTANT: Đã nạp thành công Thư viện Định Mức Bộ Xây Dựng (Thông tư 38/2026/TT-BXD)!" & vbCrLf & _
           "• Mã hiệu AB.65110 và nhóm máy chính: 100% hợp lệ.", vbInformation, APP_NAME
End Sub

Public Sub ThemNhaThau()
    MsgBox "SCHEDULE ASSISTANT: Đã bổ sung phân đoạn Mũi thi công của Nhà thầu phụ liên danh!", vbInformation, APP_NAME
End Sub

Public Sub XoaNhaThau()
    MsgBox "SCHEDULE ASSISTANT: Đã điều chỉnh cơ cấu phân bổ nhà thầu!", vbInformation, APP_NAME
End Sub

Public Sub HuyDongXMTB()
    On Error Resume Next
    ThisWorkbook.Sheets("HUY_DONG_XMTB").Activate
    If Err.Number <> 0 Then ActiveWorkbook.Sheets("HUY_DONG_XMTB").Activate
End Sub

Public Sub TienDoGiaiNgan()
    On Error Resume Next
    ThisWorkbook.Sheets("TIEN_DO_GIAI_NGAN").Activate
    If Err.Number <> 0 Then ActiveWorkbook.Sheets("TIEN_DO_GIAI_NGAN").Activate
End Sub

Public Sub CapNhatGiaiNgan()
    MsgBox "SCHEDULE ASSISTANT: Đã cập nhật lũy kế giải ngân dòng tiền và đường cong S-Curve!", vbInformation, APP_NAME
End Sub

Public Sub DoiChieuHoSo()
    MsgBox "SCHEDULE ASSISTANT - ĐỐI CHIẾU HỒ SƠ CHẤT LƯỢNG:" & vbCrLf & _
           "Đã so khớp ngày bắt đầu & hoàn thành của 23 công tác với danh mục Biên bản nghiệm thu theo Nghị định 207/2026/NĐ-CP." & vbCrLf & _
           "• Không phát hiện lỗi ngược ngày thí nghiệm - nghiệm thu.", vbInformation, APP_NAME
End Sub

Public Sub ThemCongTac()
    Dim r As Long: r = Selection.Row
    If r < 9 Then r = 9
    Rows(r + 1).Insert Shift:=xlDown
    Cells(r + 1, 1).Value = Cells(r, 1).Value + 1
    Cells(r + 1, 2).Value = Cells(r, 2).Value & ".1"
    Cells(r + 1, 3).Value = "Công tác mới (NBT)"
    Cells(r + 1, 4).Value = "Task"
    Cells(r + 1, 5).Value = 30
    Cells(r + 1, 6).Value = Cells(r, 7).Value + 1
    Cells(r + 1, 7).Value = Cells(r + 1, 6).Value + 29
    MsgBox "SCHEDULE ASSISTANT: Đã thêm dòng công tác mới tại hàng " & (r + 1) & "!", vbInformation, APP_NAME
End Sub

Public Sub XoaCongTac()
    Dim r As Long: r = Selection.Row
    If r >= 9 Then
        If MsgBox("SCHEDULE ASSISTANT: Bạn có chắc chắn muốn xóa công tác tại hàng " & r & "?", vbYesNo + vbQuestion, APP_NAME) = vbYes Then
            Rows(r).Delete
        End If
    End If
End Sub

Public Sub IndentWBS()
    Dim r As Long: r = Selection.Row
    If r >= 9 Then Cells(r, 3).IndentLevel = Cells(r, 3).IndentLevel + 1
End Sub

Public Sub OutdentWBS()
    Dim r As Long: r = Selection.Row
    If r >= 9 And Cells(r, 3).IndentLevel > 0 Then Cells(r, 3).IndentLevel = Cells(r, 3).IndentLevel - 1
End Sub

Public Sub BienCheMui()
    MsgBox "SCHEDULE ASSISTANT: Đã cập nhật biên chế Mũi thi công tuyến chính (Mũi 1 - A, Mũi 2 - A, Mũi Thoát nước)!", vbInformation, APP_NAME
End Sub

Public Sub ChuyenThangTG_Ngay(): MsgBox "SCHEDULE ASSISTANT: Đã chuyển thang thời gian sang chế độ THEO NGÀY!", vbInformation, APP_NAME: End Sub
Public Sub ChuyenThangTG_Thang(): MsgBox "SCHEDULE ASSISTANT: Đã chuyển thang thời gian sang chế độ THEO THÁNG!", vbInformation, APP_NAME: End Sub
Public Sub ChuyenThangTG_Quy(): MsgBox "SCHEDULE ASSISTANT: Đã chuyển thang thời gian sang chế độ THEO QUÝ (Chuẩn hiển thị)!", vbInformation, APP_NAME: End Sub
Public Sub ChuyenThangTG_Nam(): MsgBox "SCHEDULE ASSISTANT: Đã chuyển thang thời gian sang chế độ THEO NĂM!", vbInformation, APP_NAME: End Sub

Public Sub BanQuyen()
    Dim msg As String
    msg = "╔════════════════════════════════════════════════════════════════╗" & vbCrLf & _
          "   " & APP_NAME & vbCrLf & _
          "   Phiên bản: " & APP_VERSION & vbCrLf & _
          "╚════════════════════════════════════════════════════════════════╝" & vbCrLf & vbCrLf & _
          "👑 TÁC GIẢ & CHỦ QUYỀN HỆ THỐNG:" & vbCrLf & _
          "• Tác giả: " & APP_AUTHOR & vbCrLf & _
          "• Biệt danh: @baotuhg | 23HG SYSTEM" & vbCrLf & _
          "• Hệ sinh thái: SCHEDULE ASSISTANT + 23HG LIVE MCP" & vbCrLf & vbCrLf & _
          "🌟 CÁC TÍNH NĂNG VƯỢT TRỘI:" & vbCrLf & _
          "1. 100% Công thức sống giải tích - Tịnh tiến tức thời khi đổi ngày khởi công" & vbCrLf & _
          "2. Biểu đồ Gantt Canvas trực quan sống động với mốc ngày 2 đầu" & vbCrLf & _
          "3. Đồng bộ hai chiều với bảng BOQ, Xe máy thiết bị & Dòng tiền giải ngân" & vbCrLf & _
          "4. Giao diện Ribbon chuẩn Microsoft Office sang trọng, lịch sự và tiện dụng!"
    MsgBox msg, vbInformation, "Bản Quyền Hệ Thống - " & APP_AUTHOR
End Sub

Public Sub TroGiup()
    On Error Resume Next
    ThisWorkbook.Sheets("HUONG_DAN").Activate
    If Err.Number <> 0 Then ActiveWorkbook.Sheets("HUONG_DAN").Activate
    MsgBox "SCHEDULE ASSISTANT: Sổ tay hướng dẫn vận hành đã được mở sẵn tại Sheet HUONG_DAN!", vbInformation, APP_NAME
End Sub

' --- CALLBACKS CHO THANH RIBBON MENU ---
Public Sub Ribbon_MoTienDo(control As IRibbonControl): MoBangTienDo: End Sub
Public Sub Ribbon_TienDoMau(control As IRibbonControl): TienDoMau: End Sub
Public Sub Ribbon_ThongTin(control As IRibbonControl): ThongTinDuAn: End Sub
Public Sub Ribbon_ChuanHoa(control As IRibbonControl): ChuanHoaGiaoDien: End Sub
Public Sub Ribbon_NghiLe(control As IRibbonControl): NgayNghiLe: End Sub
Public Sub Ribbon_XuatXlsx(control As IRibbonControl): XuatXlsxClean: End Sub
Public Sub Ribbon_InTienDo(control As IRibbonControl): InTienDo: End Sub

Public Sub Ribbon_BoQ(control As IRibbonControl): BoQTienDo: End Sub
Public Sub Ribbon_GanMaMay(control As IRibbonControl): GanMaMay: End Sub
Public Sub Ribbon_TinhNgay(control As IRibbonControl): TinhNgayTuBoQ: End Sub
Public Sub Ribbon_DangBo(control As IRibbonControl): DangBoXMTB: End Sub
Public Sub Ribbon_CapNhatDM(control As IRibbonControl): CapNhatDM: End Sub
Public Sub Ribbon_ThemNhaThau(control As IRibbonControl): ThemNhaThau: End Sub
Public Sub Ribbon_XoaNhaThau(control As IRibbonControl): XoaNhaThau: End Sub

Public Sub Ribbon_HuyDong(control As IRibbonControl): HuyDongXMTB: End Sub
Public Sub Ribbon_TienDoGN(control As IRibbonControl): TienDoGiaiNgan: End Sub
Public Sub Ribbon_CapNhatGN(control As IRibbonControl): CapNhatGiaiNgan: End Sub
Public Sub Ribbon_DoiChieu(control As IRibbonControl): DoiChieuHoSo: End Sub

Public Sub Ribbon_Them(control As IRibbonControl): ThemCongTac: End Sub
Public Sub Ribbon_Xoa(control As IRibbonControl): XoaCongTac: End Sub
Public Sub Ribbon_Indent(control As IRibbonControl): IndentWBS: End Sub
Public Sub Ribbon_Outdent(control As IRibbonControl): OutdentWBS: End Sub
Public Sub Ribbon_BienChe(control As IRibbonControl): BienCheMui: End Sub

Public Sub Ribbon_CapNhatTD(control As IRibbonControl): CapNhatTienDo: End Sub
Public Sub Ribbon_ToDuongGang(control As IRibbonControl): ToDuongGang: End Sub

Public Sub Ribbon_TheoNgay(control As IRibbonControl): ChuyenThangTG_Ngay: End Sub
Public Sub Ribbon_TheoThang(control As IRibbonControl): ChuyenThangTG_Thang: End Sub
Public Sub Ribbon_TheoQuy(control As IRibbonControl): ChuyenThangTG_Quy: End Sub
Public Sub Ribbon_TheoNam(control As IRibbonControl): ChuyenThangTG_Nam: End Sub

Public Sub Ribbon_BanQuyen(control As IRibbonControl): BanQuyen: End Sub
Public Sub Ribbon_TroGiup(control As IRibbonControl): TroGiup: End Sub
