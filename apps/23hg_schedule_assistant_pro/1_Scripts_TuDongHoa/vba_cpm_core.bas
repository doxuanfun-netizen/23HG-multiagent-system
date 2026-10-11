Option Explicit
' 23HG CPM core - pure logic (no Worksheet, no Scripting.Dictionary).
' Semantics mirror cpm_engine.py:
'   work days = Mon-Sat minus holidays; duration in work days; EF = ES + (dur-1) work days;
'   milestone (Kind="Milestone"): EF = ES; lag in work days (may be negative);
'   TF in work days between ES and LS; critical when TF <= 0.
' Dates are passed as Long serial numbers (CLng(Date)).

Private mHol As Collection
Private mMap As Collection
Private mN As Long
Private mWbs() As String
Private mKind() As String
Private mDur() As Long
Private mState() As Long
Private mOrder() As Long
Private mOrderN As Long
Private mLinkN As Long
Private mFrom() As Long
Private mTo() As Long
Private mRel() As Long
Private mLag() As Long
Private mErr As String

Public Function CpmIsWork(ByVal d As Long) As Boolean
    CpmIsWork = False
    If Weekday(CDate(d), vbSunday) = 1 Then Exit Function
    Dim found As Boolean: found = False
    Dim v As Variant
    On Error Resume Next
    v = mHol.Item("h" & d)
    If Err.Number = 0 Then found = True
    Err.Clear
    On Error GoTo 0
    CpmIsWork = Not found
End Function

Public Function CpmOnOrAfter(ByVal d As Long) As Long
    Do While Not CpmIsWork(d)
        d = d + 1
    Loop
    CpmOnOrAfter = d
End Function

Public Function CpmAdd(ByVal d As Long, ByVal n As Long) As Long
    Dim stp As Long, k As Long
    If n >= 0 Then stp = 1 Else stp = -1
    k = Abs(n)
    Do While k > 0
        d = d + stp
        If CpmIsWork(d) Then k = k - 1
    Loop
    CpmAdd = d
End Function

Public Function CpmCount(ByVal a As Long, ByVal b As Long) As Long
    Dim n As Long, d As Long
    For d = a To b
        If CpmIsWork(d) Then n = n + 1
    Next d
    CpmCount = n
End Function

Public Function CpmFloat(ByVal es As Long, ByVal ls As Long) As Long
    If es <= ls Then
        CpmFloat = CpmCount(es, ls) - 1
    Else
        CpmFloat = -(CpmCount(ls, es) - 1)
    End If
End Function

Private Function Dd(ByVal i As Long) As Long
    If mKind(i) = "Milestone" Then
        Dd = 0
    ElseIf mDur(i) > 1 Then
        Dd = mDur(i) - 1
    Else
        Dd = 0
    End If
End Function

Private Function IsDigits(ByVal s As String) As Boolean
    Dim i As Long
    IsDigits = False
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    IsDigits = True
End Function

' Parses one predecessor such as "1.2.1.2SS+20", "1.1FS", "3.4", "2FS-1".
Private Function ParseOne(ByVal s As String, ByRef w As String, ByRef rel As Long, ByRef lag As Long) As Boolean
    Dim pos As Long, ch As String, tail As String, sg As Long
    ParseOne = False
    s = UCase$(Replace(Trim$(s), " ", ""))
    pos = 1
    Do While pos <= Len(s)
        ch = Mid$(s, pos, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Then
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop
    w = Left$(s, pos - 1)
    If Len(w) = 0 Then Exit Function
    rel = 1: lag = 0
    tail = Mid$(s, pos)
    If Len(tail) >= 2 Then
        Select Case Left$(tail, 2)
            Case "FS": rel = 1: tail = Mid$(tail, 3)
            Case "SS": rel = 2: tail = Mid$(tail, 3)
            Case "FF": rel = 3: tail = Mid$(tail, 3)
            Case "SF": rel = 4: tail = Mid$(tail, 3)
        End Select
    End If
    If Right$(tail, 1) = "D" Then tail = Left$(tail, Len(tail) - 1)
    If Len(tail) > 0 Then
        sg = 1
        If Left$(tail, 1) = "+" Then
            tail = Mid$(tail, 2)
        ElseIf Left$(tail, 1) = "-" Then
            sg = -1
            tail = Mid$(tail, 2)
        End If
        If Not IsDigits(tail) Then Exit Function
        lag = sg * CLng(tail)
    End If
    ParseOne = True
End Function

Private Function Visit(ByVal i As Long) As Boolean
    Visit = False
    If mState(i) = 2 Then Visit = True: Exit Function
    If mState(i) = 1 Then mErr = "Vong lap tien nhiem tai " & mWbs(i): Exit Function
    mState(i) = 1
    Dim k As Long
    For k = 1 To mLinkN
        If mTo(k) = i Then
            If Not Visit(mFrom(k)) Then Exit Function
        End If
    Next k
    mState(i) = 2
    mOrderN = mOrderN + 1
    mOrder(mOrderN) = i
    Visit = True
End Function

' Returns "" on success, otherwise an error message (outputs are then undefined).
' All array arguments are 1-based, size n. hol() may be unallocated when nHol = 0.
Public Function CpmCompute(ByVal n As Long, wbs() As String, kind() As String, dur() As Long, preds() As String, _
        ByVal projStart As Long, hol() As Long, ByVal nHol As Long, _
        es() As Long, ef() As Long, ls() As Long, lf() As Long, tf() As Long, crit() As Boolean) As String
    Dim i As Long, k As Long, j As Long, parts() As String, pw As String, rel As Long, lag As Long
    mErr = ""
    mN = n
    Set mHol = New Collection
    For i = 1 To nHol
        On Error Resume Next
        mHol.Add True, "h" & hol(i)
        Err.Clear
        On Error GoTo 0
    Next i
    Set mMap = New Collection
    ReDim mWbs(1 To n): ReDim mKind(1 To n): ReDim mDur(1 To n): ReDim mState(1 To n): ReDim mOrder(1 To n)
    For i = 1 To n
        mWbs(i) = Trim$(wbs(i)): mKind(i) = kind(i): mDur(i) = dur(i)
        If Len(mWbs(i)) = 0 Then CpmCompute = "Dong " & i & ": thieu ma WBS": Exit Function
        On Error Resume Next
        mMap.Add i, "w" & mWbs(i)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            CpmCompute = "WBS bi trung: " & mWbs(i)
            Exit Function
        End If
        On Error GoTo 0
    Next i

    ' ---- parse links ----
    mLinkN = 0
    ReDim mFrom(1 To 1): ReDim mTo(1 To 1): ReDim mRel(1 To 1): ReDim mLag(1 To 1)
    For i = 1 To n
        If mKind(i) <> "Summary" And Len(Trim$(preds(i))) > 0 Then
            parts = Split(Replace(preds(i), ";", ","), ",")
            For k = LBound(parts) To UBound(parts)
                If Len(Trim$(parts(k))) > 0 Then
                    If Not ParseOne(parts(k), pw, rel, lag) Then
                        CpmCompute = "Cong tac " & mWbs(i) & ": khong doc duoc tien nhiem '" & Trim$(parts(k)) & "'"
                        Exit Function
                    End If
                    j = 0
                    On Error Resume Next
                    j = mMap.Item("w" & pw)
                    Err.Clear
                    On Error GoTo 0
                    If j = 0 Then CpmCompute = "Cong tac " & mWbs(i) & ": tien nhiem '" & pw & "' khong ton tai": Exit Function
                    If mKind(j) = "Summary" Then CpmCompute = "Cong tac " & mWbs(i) & ": khong duoc noi tien nhiem vao dong Summary '" & pw & "'": Exit Function
                    mLinkN = mLinkN + 1
                    ReDim Preserve mFrom(1 To mLinkN): ReDim Preserve mTo(1 To mLinkN)
                    ReDim Preserve mRel(1 To mLinkN): ReDim Preserve mLag(1 To mLinkN)
                    mFrom(mLinkN) = j: mTo(mLinkN) = i: mRel(mLinkN) = rel: mLag(mLinkN) = lag
                End If
            Next k
        End If
    Next i

    ' ---- topological order of leaves ----
    mOrderN = 0
    For i = 1 To n
        If mKind(i) <> "Summary" Then
            If Not Visit(i) Then CpmCompute = mErr: Exit Function
        End If
    Next i
    If mOrderN = 0 Then CpmCompute = "Khong co cong tac nao de tinh": Exit Function

    ' ---- forward pass ----
    Dim pos As Long, w As Long, d As Long, esW As Long, efMin As Long, hasEfMin As Boolean, c As Long
    Dim s0 As Long: s0 = CpmOnOrAfter(projStart)
    Dim endD As Long
    For pos = 1 To mOrderN
        w = mOrder(pos)
        d = Dd(w)
        esW = s0: hasEfMin = False: efMin = 0
        For k = 1 To mLinkN
            If mTo(k) = w Then
                j = mFrom(k)
                Select Case mRel(k)
                    Case 1
                        c = CpmAdd(ef(j), 1 + mLag(k))
                        If c > esW Then esW = c
                    Case 2
                        c = CpmAdd(es(j), mLag(k))
                        If c > esW Then esW = c
                    Case 3
                        c = CpmAdd(ef(j), mLag(k))
                        If (Not hasEfMin) Or c > efMin Then efMin = c
                        hasEfMin = True
                    Case Else
                        c = CpmAdd(es(j), mLag(k))
                        If (Not hasEfMin) Or c > efMin Then efMin = c
                        hasEfMin = True
                End Select
            End If
        Next k
        If hasEfMin Then
            c = CpmAdd(efMin, -d)
            If c > esW Then esW = c
        End If
        es(w) = esW
        If d = 0 Then ef(w) = esW Else ef(w) = CpmAdd(esW, d)
        If pos = 1 Then
            endD = ef(w)
        ElseIf ef(w) > endD Then
            endD = ef(w)
        End If
    Next pos

    ' ---- backward pass ----
    Dim lfW As Long, lsW As Long, lsMin As Long, hasLs As Boolean
    For pos = mOrderN To 1 Step -1
        w = mOrder(pos)
        d = Dd(w)
        lfW = endD: hasLs = False: lsMin = 0
        For k = 1 To mLinkN
            If mFrom(k) = w Then
                j = mTo(k)
                Select Case mRel(k)
                    Case 1
                        c = CpmAdd(ls(j), -(1 + mLag(k)))
                        If c < lfW Then lfW = c
                    Case 2
                        c = CpmAdd(ls(j), -mLag(k))
                        If (Not hasLs) Or c < lsMin Then lsMin = c
                        hasLs = True
                    Case 3
                        c = CpmAdd(lf(j), -mLag(k))
                        If c < lfW Then lfW = c
                    Case Else
                        c = CpmAdd(lf(j), -mLag(k))
                        If (Not hasLs) Or c < lsMin Then lsMin = c
                        hasLs = True
                End Select
            End If
        Next k
        If d = 0 Then lsW = lfW Else lsW = CpmAdd(lfW, -d)
        If hasLs Then
            If lsMin < lsW Then lsW = lsMin
        End If
        ls(w) = lsW
        If d = 0 Then lf(w) = lsW Else lf(w) = CpmAdd(lsW, d)
        tf(w) = CpmFloat(es(w), ls(w))
        crit(w) = (tf(w) <= 0)
    Next pos

    ' ---- summary roll-up ----
    Dim hasKid As Boolean, pre As String
    For i = 1 To n
        If mKind(i) = "Summary" Then
            pre = mWbs(i) & "."
            hasKid = False
            For k = 1 To n
                If mKind(k) <> "Summary" And Left$(mWbs(k), Len(pre)) = pre Then
                    If Not hasKid Then
                        es(i) = es(k): ef(i) = ef(k): ls(i) = ls(k): lf(i) = lf(k): tf(i) = tf(k): crit(i) = crit(k)
                        hasKid = True
                    Else
                        If es(k) < es(i) Then es(i) = es(k)
                        If ef(k) > ef(i) Then ef(i) = ef(k)
                        If ls(k) < ls(i) Then ls(i) = ls(k)
                        If lf(k) > lf(i) Then lf(i) = lf(k)
                        If tf(k) < tf(i) Then tf(i) = tf(k)
                        If crit(k) Then crit(i) = True
                    End If
                End If
            Next k
            If Not hasKid Then CpmCompute = "Summary '" & mWbs(i) & "' khong co cong tac con": Exit Function
        End If
    Next i
    CpmCompute = ""
End Function
