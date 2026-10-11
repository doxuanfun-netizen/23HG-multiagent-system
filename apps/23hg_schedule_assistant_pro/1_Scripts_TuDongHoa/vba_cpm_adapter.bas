Option Explicit

' ==============================================================================
' 23HG SYSTEM - CPM theo lich lam viec (Thu Hai - Thu Bay, tru ngay le)
' Tac gia: NBT (Nguyen Bao Tu) | @baotuhg | 23HG SYSTEM
' Phan tinh toan nam o module M_23HG_CPMCore (CpmCompute). Module nay chi doc/ghi sheet.
' ==============================================================================

Public Function ParseDateVN(ByVal val As Variant) As Date
    On Error GoTo Fallback
    Dim s As String: s = Trim(CStr(val))
    If InStr(s, " ") > 0 Then s = Left(s, InStr(s, " ") - 1)
    If InStr(s, "/") > 0 Then
        Dim p() As String: p = Split(s, "/")
        If UBound(p) = 2 Then
            Dim d As Long, m As Long, y As Long
            d = CLng(p(0)): m = CLng(p(1)): y = CLng(p(2))
            If y < 100 Then y = 2000 + y
            If m >= 1 And m <= 12 And d >= 1 And d <= 31 Then
                ParseDateVN = DateSerial(y, m, d)
                Exit Function
            End If
        End If
    End If
    If VarType(val) = vbDate Or IsDate(val) Then
        Dim vDate As Date: vDate = CDate(val)
        ParseDateVN = DateSerial(Year(vDate), Month(vDate), Day(vDate))
        Exit Function
    End If
Fallback:
    ParseDateVN = DateSerial(2023, 12, 1)
End Function

Private Sub AddHoliday(ByRef hol() As Long, ByRef nHol As Long, ByVal d As Long)
    nHol = nHol + 1
    ReDim Preserve hol(1 To nHol)
    hol(nHol) = d
End Sub

' Doc ngay nghi le tu sheet NGAY_NGHI_LE (cot B, hang 6-15): "dd/mm/yyyy" hoac "dd/mm/yyyy - dd/mm/yyyy".
' O trong bi bo qua (khong con bi tinh thanh ngay le 01/12/2023).
Private Sub LoadHolidays(ByRef hol() As Long, ByRef nHol As Long)
    Dim wsNL As Worksheet, r As Long, s As String, parts() As String
    Dim a As Long, b As Long, d As Long
    nHol = 0
    On Error Resume Next
    Set wsNL = ThisWorkbook.Sheets("NGAY_NGHI_LE")
    On Error GoTo 0
    If wsNL Is Nothing Then Exit Sub
    For r = 6 To 15
        If VarType(wsNL.Cells(r, "B").Value) = vbDate Then
            AddHoliday hol, nHol, CLng(wsNL.Cells(r, "B").Value)
        Else
            s = Trim$(wsNL.Cells(r, "B").Text)
            If Len(s) > 0 Then
                If InStr(s, "-") > 0 Then
                    parts = Split(s, "-")
                    a = CLng(ParseDateVN(parts(0))): b = CLng(ParseDateVN(parts(1)))
                    If b - a >= 0 And b - a <= 366 Then
                        For d = a To b
                            AddHoliday hol, nHol, d
                        Next d
                    End If
                Else
                    AddHoliday hol, nHol, CLng(ParseDateVN(s))
                End If
            End If
        End If
    Next r
End Sub

Public Sub CalculateCPM(ByVal ws As Worksheet, Optional ByVal showMessage As Boolean = False)
    Dim prevSU As Boolean, prevEvents As Boolean, prevCalc As Long
    prevSU = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 6 Then GoTo CleanExit

    Dim n As Long: n = lastRow - 5
    Dim inData As Variant: inData = ws.Range("A6:M" & lastRow).Value

    Dim hol() As Long, nHol As Long
    ReDim hol(1 To 1)
    LoadHolidays hol, nHol

    ' Ngay khoi cong lay tu F2
    Dim userF2 As Date: userF2 = ParseDateVN(ws.Range("F2").Text)
    If userF2 < DateSerial(2000, 1, 1) Then userF2 = ParseDateVN(ws.Range("F2").Value)
    If userF2 < DateSerial(2000, 1, 1) Then userF2 = DateSerial(2023, 12, 1)

    Dim wbs() As String, kind() As String, dur() As Long, preds() As String
    ReDim wbs(1 To n): ReDim kind(1 To n): ReDim dur(1 To n): ReDim preds(1 To n)
    Dim i As Long, k As Double, dv As Long
    For i = 1 To n
        wbs(i) = Trim$(CStr(inData(i, 2)))
        preds(i) = Trim$(CStr(inData(i, 5)))
        dv = 0
        If IsNumeric(inData(i, 6)) Then
            If Val(inData(i, 6)) > 0 Then dv = CLng(Val(inData(i, 6)))
        End If
        ' He so nang suat mua K_tt (0<K<1 keo dai thoi luong)
        k = 1
        If IsNumeric(inData(i, 7)) Then
            If Val(inData(i, 7)) > 0 Then k = CDbl(Val(inData(i, 7)))
        End If
        If k > 0 And k < 1 And dv > 0 Then dv = CLng(Round(dv / k))
        dur(i) = dv
        If Trim$(CStr(inData(i, 4))) = "Summary" Then
            kind(i) = "Summary"
        ElseIf dv = 0 Then
            kind(i) = "Milestone"
        Else
            kind(i) = "Task"
        End If
    Next i

    Dim es() As Long, ef() As Long, ls() As Long, lf() As Long, tf() As Long, crit() As Boolean
    ReDim es(1 To n): ReDim ef(1 To n): ReDim ls(1 To n): ReDim lf(1 To n): ReDim tf(1 To n): ReDim crit(1 To n)
    Dim errMsg As String
    errMsg = CpmCompute(n, wbs, kind, dur, preds, CLng(userF2), hol, nHol, es, ef, ls, lf, tf, crit)
    If Len(errMsg) > 0 Then
        ' Du lieu sai: khong ghi de len bang tien do
        If showMessage Then
            MsgBox "Khong tinh duoc tien do: " & errMsg, vbExclamation, "23HG Schedule Assistant"
        Else
            Application.StatusBar = "23HG CPM: " & errMsg
        End If
        GoTo CleanExit
    End If
    Application.StatusBar = False

    Dim outDates() As Variant: ReDim outDates(1 To n, 1 To 5)
    For i = 1 To n
        outDates(i, 1) = CDate(es(i))
        outDates(i, 2) = CDate(ef(i))
        outDates(i, 3) = ef(i) - es(i) + 1
        outDates(i, 4) = tf(i)
        If crit(i) Then outDates(i, 5) = "GANG" Else outDates(i, 5) = ""
    Next i
    ws.Range("H6:L" & lastRow).Value = outDates
    ws.Range("H6:I" & lastRow).NumberFormat = "dd/mm/yyyy"

    For i = 1 To n
        If crit(i) And kind(i) <> "Summary" Then
            ws.Cells(5 + i, "C").Font.Color = RGB(220, 38, 38)
            ws.Cells(5 + i, "C").Font.Bold = True
            ws.Cells(5 + i, "L").Font.Color = RGB(220, 38, 38)
            ws.Cells(5 + i, "L").Font.Bold = True
        Else
            ws.Cells(5 + i, "C").Font.Color = RGB(0, 0, 0)
            ws.Cells(5 + i, "C").Font.Bold = False
            ws.Cells(5 + i, "L").Font.Color = RGB(71, 85, 105)
            ws.Cells(5 + i, "L").Font.Bold = False
        End If
    Next i

    ' Banner khoi cong / ket thuc
    ws.Range("F2").NumberFormat = "dd/mm/yyyy"
    ws.Range("F2").Value = CDate(es(1))
    ws.Range("H2").NumberFormat = "dd/mm/yyyy"
    ws.Range("H2").Value = CDate(ef(1))

CleanExit:
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevSU
    Exit Sub

ErrHandler:
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevSU
    Application.StatusBar = "23HG CPM loi: " & Err.Description
End Sub
