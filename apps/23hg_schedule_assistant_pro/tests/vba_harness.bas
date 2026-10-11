Option VBASupport 1
Option Explicit
' Test-only wrapper: text in -> text out. Not shipped.
' Input lines: "START=serial", "HOL=s1,s2,...", then "wbs|kind|dur|preds" per task.
Public Function CpmRunText(ByVal txt As String) As String
    Dim lines() As String, f() As String, i As Long, n As Long, nh As Long, k As Long
    lines = Split(txt, vbLf)
    Dim start As Long
    Dim hol() As Long: ReDim hol(1 To 1)
    Dim wbs() As String, kind() As String, dur() As Long, preds() As String
    ReDim wbs(1 To UBound(lines) + 1): ReDim kind(1 To UBound(lines) + 1)
    ReDim dur(1 To UBound(lines) + 1): ReDim preds(1 To UBound(lines) + 1)
    For i = 0 To UBound(lines)
        If Left$(lines(i), 6) = "START=" Then
            start = CLng(Mid$(lines(i), 7))
        ElseIf Left$(lines(i), 4) = "HOL=" Then
            If Len(lines(i)) > 4 Then
                f = Split(Mid$(lines(i), 5), ",")
                nh = UBound(f) + 1
                ReDim hol(1 To nh)
                For k = 0 To UBound(f): hol(k + 1) = CLng(f(k)): Next k
            End If
        ElseIf Len(lines(i)) > 0 Then
            f = Split(lines(i), "|")
            n = n + 1
            wbs(n) = f(0): kind(n) = f(1): dur(n) = CLng(f(2)): preds(n) = f(3)
        End If
    Next i
    ReDim Preserve wbs(1 To n): ReDim Preserve kind(1 To n): ReDim Preserve dur(1 To n): ReDim Preserve preds(1 To n)
    Dim es() As Long, ef() As Long, ls() As Long, lf() As Long, tf() As Long, cr() As Boolean
    ReDim es(1 To n): ReDim ef(1 To n): ReDim ls(1 To n): ReDim lf(1 To n): ReDim tf(1 To n): ReDim cr(1 To n)
    Dim er As String
    er = CpmCompute(n, wbs, kind, dur, preds, start, hol, nh, es, ef, ls, lf, tf, cr)
    If Len(er) > 0 Then CpmRunText = "ERR:" & er: Exit Function
    Dim out As String
    For i = 1 To n
        out = out & wbs(i) & "|" & es(i) & "|" & ef(i) & "|" & ls(i) & "|" & lf(i) & "|" & tf(i) & "|" & IIf(cr(i), 1, 0) & vbLf
    Next i
    CpmRunText = out
End Function
