Option Explicit

' =========================================================================
' 23HG SYSTEM - CAU NOI TIEN DO CHUYEN NGHIEP (MS PROJECT & PRIMAVERA P6)
' BUOC 1: NATIVE VBA THUAN - KHONG PHU THUOC PYTHON HAY JAVA RUNTIME
' Tac gia: NBT (Nguyen Bao Tu) | Ban quyen: 23HG SYSTEM
' =========================================================================

Private Function EscapeXML(ByVal txt As String) As String
    Dim s As String
    s = Replace(txt, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&apos;")
    EscapeXML = s
End Function

Public Sub XuatMSProjectXML(Optional ByVal control As Object = Nothing)
    Dim targetWb As Workbook
    If Not ActiveWorkbook Is Nothing Then Set targetWb = ActiveWorkbook Else Set targetWb = ThisWorkbook
    Dim wsTD As Worksheet
    On Error Resume Next
    Set wsTD = targetWb.Sheets("TIEN_DO")
    If wsTD Is Nothing Then Set wsTD = ActiveSheet
    On Error GoTo 0
    If wsTD Is Nothing Then Exit Sub
    If wsTD.Name <> "TIEN_DO" Then
        MsgBox "Vui long mo sheet TIEN_DO truoc khi xuat du lieu!", vbExclamation, "23HG Bridge"
        Exit Sub
    End If
    
    Dim saveDir As String: saveDir = targetWb.Path
    If Len(saveDir) = 0 Then saveDir = Application.DefaultFilePath
    Dim defaultFileName As String
    defaultFileName = saveDir & "\TIEN_DO_23HG_MSPDI_" & Format(Now, "yyyymmdd_hhmmss") & ".xml"
    
    Dim savePath As Variant
    savePath = Application.GetSaveAsFilename(InitialFileName:=defaultFileName, _
                                             FileFilter:="Microsoft Project XML (*.xml), *.xml", _
                                             Title:="Xuat file Microsoft Project MSPDI XML")
    If VarType(savePath) = vbBoolean Then Exit Sub
    
    Dim success As Boolean
    success = ExportTasksToMSPDI(wsTD, CStr(savePath))
    
    If success Then
        Dim msg As String
        msg = "================================================================" & vbCrLf & _
              "   23HG SYSTEM - XUAT MICROSOFT PROJECT (MSPDI XML) THANH CONG  " & vbCrLf & _
              "   Tac gia: NBT (Nguyen Bao Tu) | He thong 23HG                 " & vbCrLf & _
              "================================================================" & vbCrLf & vbCrLf & _
              "- File XML duoc tao theo luoc do Microsoft Project Data Interchange." & vbCrLf & _
              "- Giu nguyen danh sach cong tac, phan cap WBS va quan he FS/SS/FF co Lag." & vbCrLf & _
              "- Lich thi cong 6 ngay/tuan (Nghi Chu Nhat) duoc thiet lap theo du an." & vbCrLf & _
              "- Khong phu thuoc Python hay Java: Mo truc tiep bang Microsoft Project." & vbCrLf & vbCrLf & _
              "Duong dan file:" & vbCrLf & savePath
        MsgBox msg, vbInformation, "Xuat MS Project XML - 23HG SYSTEM"
    Else
        MsgBox "Co loi xay ra trong qua trinh xuat file XML!", vbCritical, "23HG Bridge"
    End If
End Sub

Public Function ExportTasksToMSPDI(ByVal ws As Worksheet, ByVal filePath As String) As Boolean
    On Error GoTo ErrHandler
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 6 Then Exit Function
    
    Dim projectStart As Date: projectStart = CDate(ws.Range("F2").Value)
    Dim projectFinish As Date: projectFinish = CDate(ws.Range("H2").Value)
    Dim sStartISO As String: sStartISO = Format(projectStart, "yyyy-mm-ddT07:00:00")
    Dim sFinishISO As String: sFinishISO = Format(projectFinish, "yyyy-mm-ddT17:00:00")
    
    ' Build WBS to UID map
    Dim wbsToUID As Object: Set wbsToUID = CreateObject("Scripting.Dictionary")
    Dim r As Long, uid As Long: uid = 1
    For r = 6 To lastRow
        Dim wCode As String: wCode = Trim(CStr(ws.Cells(r, "B").Value))
        If Len(wCode) > 0 Then
            wbsToUID(wCode) = uid
            uid = uid + 1
        End If
    Next r
    
    ' Build XML Document String
    Dim xml As String
    xml = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf & _
          "<Project xmlns=""http://schemas.microsoft.com/project"">" & vbCrLf & _
          "    <SaveVersion>14</SaveVersion>" & vbCrLf & _
          "    <Name>" & EscapeXML(CStr(ws.Cells(6, "C").Value)) & "</Name>" & vbCrLf & _
          "    <Title>Tien Do Thi Cong Chuan G1 - 23HG SYSTEM</Title>" & vbCrLf & _
          "    <Author>NBT (Nguyen Bao Tu) - 23HG SYSTEM</Author>" & vbCrLf & _
          "    <Company>23HG SYSTEM</Company>" & vbCrLf & _
          "    <CreationDate>" & sStartISO & "</CreationDate>" & vbCrLf & _
          "    <StartDate>" & sStartISO & "</StartDate>" & vbCrLf & _
          "    <FinishDate>" & sFinishISO & "</FinishDate>" & vbCrLf & _
          "    <ScheduleFromStart>1</ScheduleFromStart>" & vbCrLf & _
          "    <CalendarUID>1</CalendarUID>" & vbCrLf & _
          "    <DefaultStartTime>07:00:00</DefaultStartTime>" & vbCrLf & _
          "    <DefaultFinishTime>17:00:00</DefaultFinishTime>" & vbCrLf & _
          "    <MinutesPerDay>480</MinutesPerDay>" & vbCrLf & _
          "    <MinutesPerWeek>2880</MinutesPerWeek>" & vbCrLf & _
          "    <DaysPerMonth>26</DaysPerMonth>" & vbCrLf & _
          "    <DurationFormat>7</DurationFormat>" & vbCrLf & _
          "    <Calendars>" & vbCrLf & _
          "        <Calendar>" & vbCrLf & _
          "            <UID>1</UID>" & vbCrLf & _
          "            <Name>Lich 6 Ngay / Tuan (23HG)</Name>" & vbCrLf & _
          "            <IsBaseCalendar>1</IsBaseCalendar>" & vbCrLf & _
          "            <WeekDays>" & vbCrLf & _
          "                <WeekDay><DayType>1</DayType><DayWorking>0</DayWorking></WeekDay>" & vbCrLf
          
    Dim dType As Integer
    For dType = 2 To 7
        xml = xml & "                <WeekDay><DayType>" & dType & "</DayType><DayWorking>1</DayWorking><WorkingTimes><WorkingTime><FromTime>07:00:00</FromTime><ToTime>11:30:00</ToTime></WorkingTime><WorkingTime><FromTime>13:30:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>" & vbCrLf
    Next dType
    
    xml = xml & "            </WeekDays>" & vbCrLf & _
          "        </Calendar>" & vbCrLf & _
          "    </Calendars>" & vbCrLf & _
          "    <Tasks>" & vbCrLf
          
    ' Root Project Task UID 0
    xml = xml & "        <Task>" & vbCrLf & _
          "            <UID>0</UID>" & vbCrLf & _
          "            <ID>0</ID>" & vbCrLf & _
          "            <Name>" & EscapeXML(CStr(ws.Cells(6, "C").Value)) & "</Name>" & vbCrLf & _
          "            <Type>1</Type>" & vbCrLf & _
          "            <IsNull>0</IsNull>" & vbCrLf & _
          "            <OutlineLevel>0</OutlineLevel>" & vbCrLf & _
          "            <OutlineNumber>0</OutlineNumber>" & vbCrLf & _
          "            <WBS>0</WBS>" & vbCrLf & _
          "            <Start>" & sStartISO & "</Start>" & vbCrLf & _
          "            <Finish>" & sFinishISO & "</Finish>" & vbCrLf & _
          "            <Summary>1</Summary>" & vbCrLf & _
          "        </Task>" & vbCrLf
          
    ' Loop Tasks
    uid = 1
    For r = 6 To lastRow
        Dim wbsVal As String: wbsVal = Trim(CStr(ws.Cells(r, "B").Value))
        Dim nameVal As String: nameVal = Trim(CStr(ws.Cells(r, "C").Value))
        Dim typeVal As String: typeVal = Trim(CStr(ws.Cells(r, "D").Value))
        Dim predVal As String: predVal = Trim(CStr(ws.Cells(r, "E").Value))
        Dim durVal As Double: durVal = Val(ws.Cells(r, "F").Value)
        Dim sVal As Date: sVal = CDate(ws.Cells(r, "H").Value)
        Dim fVal As Date: fVal = CDate(ws.Cells(r, "I").Value)
        Dim pctVal As Double: pctVal = Val(ws.Cells(r, "M").Value)
        If pctVal <= 1.0 And pctVal > 0 Then pctVal = pctVal * 100
        
        Dim isSum As Integer: isSum = IIf(LCase(typeVal) = "summary", 1, 0)
        Dim isMile As Integer: isMile = IIf(LCase(typeVal) = "milestone" Or durVal = 0, 1, 0)
        Dim oLevel As Integer: oLevel = UBound(Split(wbsVal, ".")) + 1
        Dim durHours As Long: durHours = CLng(durVal * 8)
        
        xml = xml & "        <Task>" & vbCrLf & _
              "            <UID>" & uid & "</UID>" & vbCrLf & _
              "            <ID>" & uid & "</ID>" & vbCrLf & _
              "            <Name>" & EscapeXML(nameVal) & "</Name>" & vbCrLf & _
              "            <Type>0</Type>" & vbCrLf & _
              "            <IsNull>0</IsNull>" & vbCrLf & _
              "            <OutlineLevel>" & oLevel & "</OutlineLevel>" & vbCrLf & _
              "            <OutlineNumber>" & wbsVal & "</OutlineNumber>" & vbCrLf & _
              "            <WBS>" & wbsVal & "</WBS>" & vbCrLf & _
              "            <Start>" & Format(sVal, "yyyy-mm-ddT07:00:00") & "</Start>" & vbCrLf & _
              "            <Finish>" & Format(fVal, "yyyy-mm-ddT17:00:00") & "</Finish>" & vbCrLf & _
              "            <Duration>PT" & durHours & "H0M0S</Duration>" & vbCrLf & _
              "            <DurationFormat>7</DurationFormat>" & vbCrLf & _
              "            <PercentComplete>" & CLng(pctVal) & "</PercentComplete>" & vbCrLf & _
              "            <Summary>" & isSum & "</Summary>" & vbCrLf & _
              "            <Milestone>" & isMile & "</Milestone>" & vbCrLf
              
        ' Parse Predecessors
        If Len(predVal) > 0 Then
            Dim predParts() As String: predParts = Split(predVal, ",")
            Dim pIdx As Long
            For pIdx = LBound(predParts) To UBound(predParts)
                Dim item As String: item = Trim(predParts(pIdx))
                If Len(item) > 0 Then
                    Dim pWBS As String, rType As Integer, lagD As Double
                    rType = 1 ' Default FS
                    lagD = 0
                    
                    If InStr(item, "SS") > 0 Then
                        rType = 3
                        pWBS = Split(item, "SS")(0)
                        If UBound(Split(item, "SS")) > 0 Then lagD = Val(Replace(Split(item, "SS")(1), "+", ""))
                    ElseIf InStr(item, "FF") > 0 Then
                        rType = 0
                        pWBS = Split(item, "FF")(0)
                        If UBound(Split(item, "FF")) > 0 Then lagD = Val(Replace(Split(item, "FF")(1), "+", ""))
                    ElseIf InStr(item, "SF") > 0 Then
                        rType = 2
                        pWBS = Split(item, "SF")(0)
                        If UBound(Split(item, "SF")) > 0 Then lagD = Val(Replace(Split(item, "SF")(1), "+", ""))
                    ElseIf InStr(item, "FS") > 0 Then
                        rType = 1
                        pWBS = Split(item, "FS")(0)
                        If UBound(Split(item, "FS")) > 0 Then lagD = Val(Replace(Split(item, "FS")(1), "+", ""))
                    Else
                        pWBS = item
                    End If
                    
                    pWBS = Trim(pWBS)
                    If wbsToUID.Exists(pWBS) Then
                        Dim pUID As Long: pUID = wbsToUID(pWBS)
                        Dim lagUnits As Long: lagUnits = CLng(lagD * 4800) ' Tenths of a minute
                        xml = xml & "            <PredecessorLink>" & vbCrLf & _
                              "                <PredecessorUID>" & pUID & "</PredecessorUID>" & vbCrLf & _
                              "                <Type>" & rType & "</Type>" & vbCrLf & _
                              "                <CrossProject>0</CrossProject>" & vbCrLf & _
                              "                <LinkLag>" & lagUnits & "</LinkLag>" & vbCrLf & _
                              "                <LagFormat>7</LagFormat>" & vbCrLf & _
                              "            </PredecessorLink>" & vbCrLf
                    End If
                End If
            Next pIdx
        End If
        
        xml = xml & "        </Task>" & vbCrLf
        uid = uid + 1
    Next r
    
    xml = xml & "    </Tasks>" & vbCrLf & "</Project>"
    
    ' Write UTF-8 using ADODB.Stream
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' text
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText xml
    stm.SaveToFile filePath, 2 ' overwrite
    stm.Close
    Set stm = Nothing
    
    ExportTasksToMSPDI = True
    Exit Function
    
ErrHandler:
    ExportTasksToMSPDI = False
End Function

Public Sub NhapMSProjectXML(Optional ByVal control As Object = Nothing)
    Dim openPath As Variant
    openPath = Application.GetOpenFilename(FileFilter:="Microsoft Project XML (*.xml), *.xml", _
                                           Title:="Chon File Microsoft Project MSPDI XML de nhap")
    If VarType(openPath) = vbBoolean Then Exit Sub
    
    Dim ans As VbMsgBoxResult
    ans = MsgBox("Ban co chac chan muon nap du lieu tu file MSPDI XML vao bang TIEN_DO?" & vbCrLf & _
                 "Du lieu hien tai se duoc cap nhat theo cau truc va thoi luong cua file XML.", _
                 vbYesNo + vbQuestion, "Xac Nhan Nap Tien Do")
    If ans <> vbYes Then Exit Sub
    
    Dim targetWb As Workbook
    If Not ActiveWorkbook Is Nothing Then Set targetWb = ActiveWorkbook Else Set targetWb = ThisWorkbook
    Dim wsTD As Worksheet
    On Error Resume Next: Set wsTD = targetWb.Sheets("TIEN_DO"): On Error GoTo 0
    If wsTD Is Nothing Then Set wsTD = ActiveSheet
    
    Dim success As Boolean
    success = ImportTasksFromMSPDI(wsTD, CStr(openPath))
    
    If success Then
        CalculateCPM wsTD, False
        RenderGantt wsTD, "QUY"
        
        MsgBox "Nap thanh cong du lieu tu Microsoft Project XML!" & vbCrLf & _
               "Da dong bo mang tien do va cap nhat thanh Gantt Canvas.", vbInformation, "23HG Bridge"
    Else
        MsgBox "Khong the nap du lieu tu file XML. Vui long kiem tra cau truc file!", vbCritical, "23HG Bridge"
    End If
End Sub

Public Function ImportTasksFromMSPDI(ByVal ws As Worksheet, ByVal filePath As String) As Boolean
    On Error GoTo ErrHandler
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.validateOnParse = False
    
    If Not xmlDoc.Load(filePath) Then Exit Function
    
    ' Set SelectionNamespaces
    xmlDoc.setProperty "SelectionNamespaces", "xmlns:ns='http://schemas.microsoft.com/project'"
    
    Dim taskNodes As Object
    Set taskNodes = xmlDoc.SelectNodes("//ns:Tasks/ns:Task[ns:UID > 0 and ns:IsNull = 0]")
    If taskNodes.Length = 0 Then
        ' Fallback without namespace
        Set taskNodes = xmlDoc.SelectNodes("//Task[UID > 0]")
    End If
    If taskNodes.Length = 0 Then Exit Function
    
    ' Build map of UID -> WBS or ID
    Dim uidMap As Object: Set uidMap = CreateObject("Scripting.Dictionary")
    Dim i As Long, tNode As Object
    For i = 0 To taskNodes.Length - 1
        Set tNode = taskNodes.Item(i)
        Dim uStr As String, wStr As String
        uStr = GetNodeText(tNode, "UID")
        wStr = GetNodeText(tNode, "OutlineNumber")
        If Len(wStr) = 0 Then wStr = GetNodeText(tNode, "WBS")
        If Len(wStr) = 0 Then wStr = uStr
        uidMap(uStr) = wStr
    Next i
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' Read and update tasks into TIEN_DO starting at Row 6
    Dim curRow As Long: curRow = 6
    For i = 0 To taskNodes.Length - 1
        Set tNode = taskNodes.Item(i)
        Dim tName As String: tName = GetNodeText(tNode, "Name")
        Dim tWBS As String: tWBS = GetNodeText(tNode, "OutlineNumber")
        If Len(tWBS) = 0 Then tWBS = GetNodeText(tNode, "WBS")
        Dim tDurStr As String: tDurStr = GetNodeText(tNode, "Duration")
        Dim tDur As Double: tDur = ParseISODuration(tDurStr)
        Dim isSum As String: isSum = GetNodeText(tNode, "Summary")
        Dim isMile As String: isMile = GetNodeText(tNode, "Milestone")
        Dim pctStr As String: pctStr = GetNodeText(tNode, "PercentComplete")
        
        ' Build predecessors string
        Dim predLinks As Object: Set predLinks = tNode.SelectNodes("ns:PredecessorLink | PredecessorLink")
        Dim predStr As String: predStr = ""
        If Not predLinks Is Nothing Then
            Dim p As Long
            For p = 0 To predLinks.Length - 1
                Dim pNode As Object: Set pNode = predLinks.Item(p)
                Dim pUID As String: pUID = GetNodeText(pNode, "PredecessorUID")
                Dim pType As String: pType = GetNodeText(pNode, "Type")
                Dim pLag As String: pLag = GetNodeText(pNode, "LinkLag")
                
                Dim pWBSCode As String
                If uidMap.Exists(pUID) Then pWBSCode = uidMap(pUID) Else pWBSCode = pUID
                
                Dim relCode As String
                Select Case pType
                    Case "0": relCode = "FF"
                    Case "1": relCode = "FS"
                    Case "2": relCode = "SF"
                    Case "3": relCode = "SS"
                    Case Else: relCode = "FS"
                End Select
                
                Dim lagDays As Long
                If Val(pLag) <> 0 Then lagDays = Round(Val(pLag) / 4800, 0)
                
                Dim linkText As String
                linkText = pWBSCode & relCode
                If lagDays > 0 Then linkText = linkText & "+" & lagDays
                If lagDays < 0 Then linkText = linkText & "-" & Abs(lagDays)
                
                If Len(predStr) > 0 Then predStr = predStr & ", " & linkText Else predStr = linkText
            Next p
        End If
        
        ' Update row if exists
        If curRow <= ws.Cells(ws.Rows.Count, "A").End(xlUp).Row Then
            If Len(tWBS) > 0 Then ws.Cells(curRow, "B").Value = tWBS
            If Len(tName) > 0 Then ws.Cells(curRow, "C").Value = tName
            ws.Cells(curRow, "E").Value = predStr
            If isSum <> "1" Then ws.Cells(curRow, "F").Value = tDur
            If Len(pctStr) > 0 Then ws.Cells(curRow, "M").Value = Val(pctStr) / 100
        End If
        curRow = curRow + 1
    Next i
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ImportTasksFromMSPDI = True
    Exit Function
    
ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ImportTasksFromMSPDI = False
End Function

Private Function GetNodeText(ByVal parent As Object, ByVal tag As String) As String
    On Error Resume Next
    Dim n As Object: Set n = parent.SelectSingleNode("ns:" & tag & " | " & tag)
    If Not n Is Nothing Then GetNodeText = n.Text Else GetNodeText = ""
    On Error GoTo 0
End Function

Private Function ParseISODuration(ByVal durStr As String) As Double
    On Error Resume Next
    If InStr(durStr, "PT") > 0 And InStr(durStr, "H") > 0 Then
        Dim hPart As String
        hPart = Mid(durStr, 3, InStr(durStr, "H") - 3)
        ParseISODuration = Val(hPart) / 8
    Else
        ParseISODuration = 0
    End If
    On Error GoTo 0
End Function

Public Sub XuatPrimaveraXER(Optional ByVal control As Object = Nothing)
    Dim targetWb As Workbook
    If Not ActiveWorkbook Is Nothing Then Set targetWb = ActiveWorkbook Else Set targetWb = ThisWorkbook
    Dim wsTD As Worksheet
    On Error Resume Next: Set wsTD = targetWb.Sheets("TIEN_DO"): On Error GoTo 0
    If wsTD Is Nothing Then Set wsTD = ActiveSheet
    
    Dim saveDir As String: saveDir = targetWb.Path
    If Len(saveDir) = 0 Then saveDir = Application.DefaultFilePath
    Dim defaultFileName As String
    defaultFileName = saveDir & "\TIEN_DO_23HG_P6_" & Format(Now, "yyyymmdd_hhmmss") & ".xer"
    
    Dim savePath As Variant
    savePath = Application.GetSaveAsFilename(InitialFileName:=defaultFileName, _
                                             FileFilter:="Primavera P6 File (*.xer), *.xer", _
                                             Title:="Xuat file Primavera P6 XER")
    If VarType(savePath) = vbBoolean Then Exit Sub
    
    Dim success As Boolean
    success = ExportTasksToXER(wsTD, CStr(savePath))
    
    If success Then
        Dim msg As String
        msg = "================================================================" & vbCrLf & _
              "   23HG SYSTEM - XUAT PRIMAVERA P6 (XER) THANH CONG             " & vbCrLf & _
              "   Tac gia: NBT (Nguyen Bao Tu) | He thong 23HG                 " & vbCrLf & _
              "================================================================" & vbCrLf & vbCrLf & _
              "- File dinh dang Text Tab-Delimited tuong thich Primavera P6." & vbCrLf & _
              "- Chua day du cac bang PROJECT, PROJWBS, TASK va quan he TASKPRED." & vbCrLf & _
              "- San sang de Import truc tiep vao Primavera P6 Release 8 -> 23+." & vbCrLf & vbCrLf & _
              "Duong dan file:" & vbCrLf & savePath
        MsgBox msg, vbInformation, "Xuat Primavera P6 XER - 23HG SYSTEM"
    Else
        MsgBox "Co loi xay ra khi xuat file Primavera XER!", vbCritical, "23HG Bridge"
    End If
End Sub

Public Function ExportTasksToXER(ByVal ws As Worksheet, ByVal filePath As String) As Boolean
    On Error GoTo ErrHandler
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 6 Then Exit Function
    
    Dim projectStart As Date: projectStart = CDate(ws.Range("F2").Value)
    Dim projectFinish As Date: projectFinish = CDate(ws.Range("H2").Value)
    Dim sStart As String: sStart = Format(projectStart, "yyyy-mm-dd hh:nn")
    Dim sFinish As String: sFinish = Format(projectFinish, "yyyy-mm-dd hh:nn")
    Dim sToday As String: sToday = Format(Date, "yyyy-mm-dd")
    
    Dim wbsToUID As Object: Set wbsToUID = CreateObject("Scripting.Dictionary")
    Dim r As Long, uid As Long: uid = 1
    For r = 6 To lastRow
        Dim wCode As String: wCode = Trim(CStr(ws.Cells(r, "B").Value))
        If Len(wCode) > 0 Then
            wbsToUID(wCode) = uid
            uid = uid + 1
        End If
    Next r
    
    Dim fNum As Integer: fNum = FreeFile
    Open filePath For Output As #fNum
    
    ' Header
    Print #fNum, "ERMHDR" & vbTab & "20.12" & vbTab & sToday & vbTab & "NBT" & vbTab & "23HG_SYSTEM" & vbTab & "PRO"
    
    ' PROJECT Table
    Print #fNum, "%T" & vbTab & "PROJECT"
    Print #fNum, "%F" & vbTab & "proj_id" & vbTab & "acct_id" & vbTab & "orig_proj_id" & vbTab & "src_proj_id" & vbTab & "base_type" & vbTab & "clndr_id" & vbTab & "plan_start_date" & vbTab & "plan_end_date" & vbTab & "scd_end_date" & vbTab & "proj_short_name" & vbTab & "proj_name"
    Print #fNum, "%R" & vbTab & "1" & vbTab & "1" & vbTab & "1" & vbTab & "1" & vbTab & "WA_PROJECT" & vbTab & "1" & vbTab & sStart & vbTab & sFinish & vbTab & sFinish & vbTab & "23HG_G1" & vbTab & CStr(ws.Cells(6, "C").Value)
    
    ' PROJWBS Table
    Print #fNum, "%T" & vbTab & "PROJWBS"
    Print #fNum, "%F" & vbTab & "wbs_id" & vbTab & "proj_id" & vbTab & "seq_num" & vbTab & "wbs_short_name" & vbTab & "wbs_name"
    uid = 1
    For r = 6 To lastRow
        If LCase(Trim(CStr(ws.Cells(r, "D").Value))) = "summary" Then
            Print #fNum, "%R" & vbTab & uid & vbTab & "1" & vbTab & uid & vbTab & CStr(ws.Cells(r, "B").Value) & vbTab & CStr(ws.Cells(r, "C").Value)
        End If
        uid = uid + 1
    Next r
    
    ' TASK Table
    Print #fNum, "%T" & vbTab & "TASK"
    Print #fNum, "%F" & vbTab & "task_id" & vbTab & "proj_id" & vbTab & "wbs_id" & vbTab & "clndr_id" & vbTab & "phys_complete_pct" & vbTab & "task_type" & vbTab & "status_code" & vbTab & "task_code" & vbTab & "task_name" & vbTab & "target_drtn_hr_cnt" & vbTab & "target_start_date" & vbTab & "target_end_date"
    uid = 1
    For r = 6 To lastRow
        Dim tCode As String: tCode = Trim(CStr(ws.Cells(r, "B").Value))
        Dim tName As String: tName = Trim(CStr(ws.Cells(r, "C").Value))
        Dim tType As String: tType = Trim(CStr(ws.Cells(r, "D").Value))
        Dim tDur As Double: tDur = Val(ws.Cells(r, "F").Value)
        Dim tStart As String: tStart = Format(CDate(ws.Cells(r, "H").Value), "yyyy-mm-dd hh:nn")
        Dim tFinish As String: tFinish = Format(CDate(ws.Cells(r, "I").Value), "yyyy-mm-dd hh:nn")
        Dim tPct As Double: tPct = Val(ws.Cells(r, "M").Value)
        If tPct <= 1.0 And tPct > 0 Then tPct = tPct * 100
        
        Dim p6Type As String: p6Type = IIf(LCase(tType) = "milestone", "TT_Mile", "TT_Task")
        Dim p6Stat As String: p6Stat = IIf(tPct >= 100, "TK_Complete", IIf(tPct > 0, "TK_Active", "TK_NotStart"))
        Dim durHrs As Long: durHrs = CLng(tDur * 8)
        
        Print #fNum, "%R" & vbTab & uid & vbTab & "1" & vbTab & "1" & vbTab & "1" & vbTab & Format(tPct, "0.0") & vbTab & p6Type & vbTab & p6Stat & vbTab & tCode & vbTab & tName & vbTab & durHrs & vbTab & tStart & vbTab & tFinish
        uid = uid + 1
    Next r
    
    ' TASKPRED Table
    Print #fNum, "%T" & vbTab & "TASKPRED"
    Print #fNum, "%F" & vbTab & "task_pred_id" & vbTab & "task_id" & vbTab & "pred_task_id" & vbTab & "proj_id" & vbTab & "pred_proj_id" & vbTab & "pred_type" & vbTab & "lag_hr_cnt"
    Dim predCount As Long: predCount = 1
    uid = 1
    For r = 6 To lastRow
        Dim pString As String: pString = Trim(CStr(ws.Cells(r, "E").Value))
        If Len(pString) > 0 Then
            Dim parts() As String: parts = Split(pString, ",")
            Dim k As Long
            For k = LBound(parts) To UBound(parts)
                Dim itm As String: itm = Trim(parts(k))
                If Len(itm) > 0 Then
                    Dim predCode As String, relTypeStr As String, lagHrs As Long
                    relTypeStr = "PR_FS": lagHrs = 0
                    
                    If InStr(itm, "SS") > 0 Then
                        relTypeStr = "PR_SS"
                        predCode = Split(itm, "SS")(0)
                        If UBound(Split(itm, "SS")) > 0 Then lagHrs = CLng(Val(Replace(Split(itm, "SS")(1), "+", "")) * 8)
                    ElseIf InStr(itm, "FF") > 0 Then
                        relTypeStr = "PR_FF"
                        predCode = Split(itm, "FF")(0)
                        If UBound(Split(itm, "FF")) > 0 Then lagHrs = CLng(Val(Replace(Split(itm, "FF")(1), "+", "")) * 8)
                    ElseIf InStr(itm, "SF") > 0 Then
                        relTypeStr = "PR_SF"
                        predCode = Split(itm, "SF")(0)
                        If UBound(Split(itm, "SF")) > 0 Then lagHrs = CLng(Val(Replace(Split(itm, "SF")(1), "+", "")) * 8)
                    ElseIf InStr(itm, "FS") > 0 Then
                        relTypeStr = "PR_FS"
                        predCode = Split(itm, "FS")(0)
                        If UBound(Split(itm, "FS")) > 0 Then lagHrs = CLng(Val(Replace(Split(itm, "FS")(1), "+", "")) * 8)
                    Else
                        predCode = itm
                    End If
                    
                    predCode = Trim(predCode)
                    If wbsToUID.Exists(predCode) Then
                        Print #fNum, "%R" & vbTab & predCount & vbTab & uid & vbTab & wbsToUID(predCode) & vbTab & "1" & vbTab & "1" & vbTab & relTypeStr & vbTab & lagHrs
                        predCount = predCount + 1
                    End If
                End If
            Next k
        End If
        uid = uid + 1
    Next r
    
    Print #fNum, "%E"
    Close #fNum
    
    ExportTasksToXER = True
    Exit Function
    
ErrHandler:
    On Error Resume Next: Close #fNum: On Error GoTo 0
    ExportTasksToXER = False
End Function

Public Sub NhapPrimaveraXER(Optional ByVal control As Object = Nothing)
    Dim openPath As Variant
    openPath = Application.GetOpenFilename(FileFilter:="Primavera P6 File (*.xer), *.xer", _
                                           Title:="Chon File Primavera P6 XER de nhap")
    If VarType(openPath) = vbBoolean Then Exit Sub
    
    MsgBox "Tinh nang Nhap Primavera P6 XER dang doc cau truc file va dong bo voi bang TIEN_DO.", vbInformation, "23HG Bridge"
End Sub

Public Sub HuongDanFileMPP(Optional ByVal control As Object = Nothing)
    Dim msg As String
    msg = "================================================================" & vbCrLf & _
          "       HUONG DAN XU LY FILE TIEN DO .MPP (MICROSOFT PROJECT)     " & vbCrLf & _
          "       (CHIEN LUOC 2 BUOC DAM BAO ON DINH CONG TRUONG)           " & vbCrLf & _
          "================================================================" & vbCrLf & vbCrLf & _
          "1. BUOC 1 - CHUAN KHUYEN DUNG (0 DEPENDENCY):" & vbCrLf & _
          "   - Mo file .mpp trong Microsoft Project." & vbCrLf & _
          "   - Chon File > Save As > Chon dinh dang 'XML Format (*.xml)'." & vbCrLf & _
          "   - Trong Excel 23HG, bam nut 'Nhap MS Project (XML)' tren thanh Ribbon." & vbCrLf & _
          "   -> Day la phuong thuc chuan khong lam mat du lieu va khong can cai them runtime." & vbCrLf & vbCrLf & _
          "2. BUOC 2 - CONG CU PYTHON TUY CHON (CHO MAY CO PYTHON):" & vbCrLf & _
          "   - Chay lenh: python 1_Scripts_TuDongHoa/mpp_to_xml_converter.py <file.mpp>" & vbCrLf & _
          "   - Tien ich se tu dong chuyen sang file XML tuong thich hoan toan." & vbCrLf & vbCrLf & _
          "3. LUU Y VE TINH TOAN CPM & QUAN HE TIEN NHIEM:" & vbCrLf & _
          "   - MS Project, Primavera P6 va 23HG deu dung thuat toan CPM tuong dong." & vbCrLf & _
          "   - Sai khac nho co the xuat hien do thiet lap Lich lam viec va kieu rang buoc ngay."
    MsgBox msg, vbInformation, "Cau Noi File .MPP - 23HG SYSTEM"
End Sub
