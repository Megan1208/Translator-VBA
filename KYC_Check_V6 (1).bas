Option Explicit
' ============================================================
' KYC_Check_V6
' Compares PREVIOUS vs CURRENT export. Output = 3 sheets only:
'   Previous       : source data + Check Status (removed / changed rows in orange)
'   Current_Check  : source data + Check Status + Check Detail
'                    yellow = changed / new information, red = mandatory field blank
'   Change_Detail  : one row per change, new information (Current Value) in yellow
' Matching order (all in memory, arrays):
'   1) Party Name + ECID + Party Type   (same record, other fields may change)
'   2) ECID + Party Type                (Party Name changed)
'   3) Party Name + Party Type          (ECID changed)
'   anything left = added / removed
' Mandatory-field rules live in sheet "Requirement1" of THIS macro workbook.
' ============================================================

Private Const SEP As String = "|~|"
Private Const MAX_PAIR_SCAN As Long = 10000

Private oA() As String, nA() As String
Private oRowN As Long, nRowN As Long, oC As Long, nC As Long
Private oH As Object, nH As Object
Private oNameC As Long, oEcidC As Long, oTypeC As Long
Private nNameC As Long, nEcidC As Long, nTypeC As Long
Private oOK() As Boolean, nOK() As Boolean
Private fKey() As String, fOldC() As Long, fNewC() As Long, fCount As Long
Private oMate() As Long, nMate() As Long, nKind() As Long
Private oStat() As String, nStat() As String
Private nChg() As String, nMiss() As String, nNote() As String
Private oNames As Object, oNameType As Object, nNames As Object, nNameType As Object
Private missSeen As Object
Private dA() As String, dY() As Boolean, dN As Long, dCap As Long
Private mR() As Long, mC() As Long, mK() As Long, mN As Long, mCap As Long
Private rType() As String, rField() As String, rKey() As String, rScope() As String, rCount As Long
Private cPairs As Long, cSame As Long, cChanged As Long, cAdded As Long, cRemoved As Long
Private cNameChg As Long, cEcidChg As Long, cMissing As Long, cReview As Long

' ------------------------------------------------------------
' Main entry
' ------------------------------------------------------------
Public Sub RunKYCCheckV6()
    Dim oldPath As Variant, newPath As Variant, wb As Workbook, rulesWs As Worksheet
    Dim savedEvents As Boolean, savedScreen As Boolean, savedSecurity As Long
    Dim errText As String, filt As String
    savedEvents = Application.EnableEvents
    savedScreen = Application.ScreenUpdating
    savedSecurity = Application.AutomationSecurity
    filt = "Data files (*.csv;*.xlsx;*.xlsm;*.xls),*.csv;*.xlsx;*.xlsm;*.xls"
    On Error GoTo Failed
    oldPath = Application.GetOpenFilename(filt, , "1 - Select PREVIOUS file")
    If VarType(oldPath) = vbBoolean Then Exit Sub
    newPath = Application.GetOpenFilename(filt, , "2 - Select CURRENT file")
    If VarType(newPath) = vbBoolean Then Exit Sub
    If StrComp(CStr(oldPath), CStr(newPath), vbTextCompare) = 0 Then Err.Raise 5, , "Select two different source files."
    Application.EnableEvents = False
    Application.AutomationSecurity = 3
    Set rulesWs = GetRequirementSheet()
    ValidateRules rulesWs
    LoadRules rulesWs
    Application.ScreenUpdating = False
    ReadFile CStr(oldPath), oA, oRowN, oC
    ReadFile CStr(newPath), nA, nRowN, nC
    PrepareColumns
    RunEngine
    Set wb = Workbooks.Add(xlWBATWorksheet)
    WriteOutputs wb
    Application.ScreenUpdating = True
    wb.Worksheets("Current_Check").Activate
    MsgBox "Complete." & vbCrLf & vbCrLf & _
        "Matched records: " & cPairs & "  (unchanged " & cSame & ", changed " & cChanged & ")" & vbCrLf & _
        "   of which Party Name changed: " & cNameChg & ", ECID changed: " & cEcidChg & vbCrLf & _
        "Added: " & cAdded & "   Removed: " & cRemoved & "   Need review: " & cReview & vbCrLf & _
        "Mandatory fields blank (Current): " & cMissing & vbCrLf & vbCrLf & _
        "Save the result as .xlsx. Edit Requirement1 in the macro workbook and save the .xlsm to keep rule changes.", vbInformation
Cleanup:
    Application.EnableEvents = savedEvents
    Application.ScreenUpdating = savedScreen
    Application.AutomationSecurity = savedSecurity
    Exit Sub
Failed:
    errText = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
    MsgBox "Stopped: " & errText & vbCrLf & "No result workbook was kept.", vbExclamation
    GoTo Cleanup
End Sub

' ------------------------------------------------------------
' Engine (works only on the in-memory arrays)
' ------------------------------------------------------------
Private Sub RunEngine()
    ResetState
    BuildNameSets
    PassMatch 1
    PassMatch 2
    PassMatch 3
    BuildResults
    CheckMissing
End Sub

Private Sub ResetState()
    ReDim oMate(1 To oRowN): ReDim nMate(1 To nRowN): ReDim nKind(1 To nRowN)
    ReDim oStat(1 To oRowN): ReDim nStat(1 To nRowN)
    ReDim nChg(1 To nRowN): ReDim nMiss(1 To nRowN): ReDim nNote(1 To nRowN)
    dCap = 256: dN = 0
    ReDim dA(1 To dCap, 1 To 11): ReDim dY(1 To dCap)
    mCap = 256: mN = 0
    ReDim mR(1 To mCap): ReDim mC(1 To mCap): ReDim mK(1 To mCap)
    cPairs = 0: cSame = 0: cChanged = 0: cAdded = 0: cRemoved = 0
    cNameChg = 0: cEcidChg = 0: cMissing = 0: cReview = 0
    Set missSeen = CreateObject("Scripting.Dictionary")
End Sub

Private Sub PrepareColumns()
    Dim r As Long, k As Variant
    Set oH = HeaderMap(oA, oC, "Previous")
    Set nH = HeaderMap(nA, nC, "Current")
    oNameC = PickHeader(oH, "Party Name", "Previous")
    nNameC = PickHeader(nH, "Party Name", "Current")
    oEcidC = PickHeader(oH, "ECID", "Previous")
    nEcidC = PickHeader(nH, "ECID", "Current")
    oTypeC = PickHeader(oH, "Party Type", "Previous")
    nTypeC = PickHeader(nH, "Party Type", "Current")
    fCount = 0
    ReDim fKey(1 To nH.Count): ReDim fOldC(1 To nH.Count): ReDim fNewC(1 To nH.Count)
    For Each k In nH.Keys
        If oH.Exists(k) Then
            fCount = fCount + 1
            fKey(fCount) = CStr(k)
            fOldC(fCount) = oH(k)
            fNewC(fCount) = nH(k)
        End If
    Next k
    If oH.Exists("dob") Then NormDates oA, oRowN, oH("dob")
    If nH.Exists("dob") Then NormDates nA, nRowN, nH("dob")
    ReDim oOK(1 To oRowN): ReDim nOK(1 To nRowN)
    For r = 2 To oRowN: oOK(r) = IsDataRow(oA, r, oC): Next r
    For r = 2 To nRowN: nOK(r) = IsDataRow(nA, r, nC): Next r
End Sub

Private Sub BuildNameSets()
    Dim r As Long, nm As String
    Set oNames = CreateObject("Scripting.Dictionary")
    Set oNameType = CreateObject("Scripting.Dictionary")
    Set nNames = CreateObject("Scripting.Dictionary")
    Set nNameType = CreateObject("Scripting.Dictionary")
    For r = 2 To oRowN
        If oOK(r) Then
            nm = LCase$(oA(r, oNameC))
            oNames(nm) = True
            oNameType(nm & SEP & LCase$(oA(r, oTypeC))) = True
        End If
    Next r
    For r = 2 To nRowN
        If nOK(r) Then
            nm = LCase$(nA(r, nNameC))
            nNames(nm) = True
            nNameType(nm & SEP & LCase$(nA(r, nTypeC))) = True
        End If
    Next r
End Sub

Private Function KeyOf(ByVal isOld As Boolean, ByVal r As Long, ByVal mode As Long) As String
    Dim nm As String, id As String, tp As String
    If isOld Then
        nm = LCase$(oA(r, oNameC)): id = LCase$(oA(r, oEcidC)): tp = LCase$(oA(r, oTypeC))
    Else
        nm = LCase$(nA(r, nNameC)): id = LCase$(nA(r, nEcidC)): tp = LCase$(nA(r, nTypeC))
    End If
    KeyOf = ""
    Select Case mode
        Case 1
            KeyOf = nm & SEP & id & SEP & tp
        Case 2
            If Not IsMissingText(id) And Not IsMissingText(tp) Then KeyOf = id & SEP & tp
        Case 3
            If Not IsMissingText(nm) And Not IsMissingText(tp) Then KeyOf = nm & SEP & tp
    End Select
End Function

Private Sub AddToGroup(ByVal d As Object, ByVal k As String, ByVal r As Long)
    Dim c As Collection
    If d.Exists(k) Then
        Set c = d(k)
    Else
        Set c = New Collection
        d.Add k, c
    End If
    c.Add r
End Sub

Private Sub PassMatch(ByVal mode As Long)
    Dim od As Object, nd As Object, r As Long, k As String, key As Variant
    Dim oGrp As Collection, nGrp As Collection
    Set od = CreateObject("Scripting.Dictionary")
    Set nd = CreateObject("Scripting.Dictionary")
    For r = 2 To oRowN
        If oOK(r) And oMate(r) = 0 Then
            k = KeyOf(True, r, mode)
            If Len(k) > 0 Then AddToGroup od, k, r
        End If
    Next r
    For r = 2 To nRowN
        If nOK(r) And nMate(r) = 0 Then
            k = KeyOf(False, r, mode)
            If Len(k) > 0 Then AddToGroup nd, k, r
        End If
    Next r
    For Each key In od.Keys
        If nd.Exists(key) Then
            Set oGrp = od(key)
            Set nGrp = nd(key)
            If mode = 1 Then
                PairGroup oGrp, nGrp
            ElseIf oGrp.Count = 1 And nGrp.Count = 1 Then
                Link CLng(oGrp(1)), CLng(nGrp(1)), mode
            End If
        End If
    Next key
End Sub

Private Sub Link(ByVal o As Long, ByVal n As Long, ByVal mode As Long)
    oMate(o) = n
    nMate(n) = o
    nKind(n) = mode
End Sub

' Pair rows that share Name + ECID + Party Type: identical rows first, then best similarity.
Private Sub PairGroup(ByVal oGrp As Collection, ByVal nGrp As Collection)
    Dim i As Long, j As Long, oCnt As Long, nCnt As Long
    Dim oIdx() As Long, nIdx() As Long
    Dim best As Long, bi As Long, bj As Long, sc As Long, remO As Long, remN As Long
    oCnt = oGrp.Count: nCnt = nGrp.Count
    ReDim oIdx(1 To oCnt): ReDim nIdx(1 To nCnt)
    For i = 1 To oCnt: oIdx(i) = CLng(oGrp(i)): Next i
    For j = 1 To nCnt: nIdx(j) = CLng(nGrp(j)): Next j
    For i = 1 To oCnt
        For j = 1 To nCnt
            If nIdx(j) > 0 Then
                If SameScore(oIdx(i), nIdx(j)) = fCount Then
                    Link oIdx(i), nIdx(j), 1
                    oIdx(i) = 0: nIdx(j) = 0
                    Exit For
                End If
            End If
        Next j
    Next i
    For i = 1 To oCnt
        If oIdx(i) > 0 Then remO = remO + 1
    Next i
    For j = 1 To nCnt
        If nIdx(j) > 0 Then remN = remN + 1
    Next j
    If remO = 0 Or remN = 0 Then Exit Sub
    If remO * remN > MAX_PAIR_SCAN Then
        j = 1
        For i = 1 To oCnt
            If oIdx(i) > 0 Then
                Do While j <= nCnt
                    If nIdx(j) > 0 Then Exit Do
                    j = j + 1
                Loop
                If j > nCnt Then Exit For
                Link oIdx(i), nIdx(j), 1
                oIdx(i) = 0: nIdx(j) = 0
            End If
        Next i
        Exit Sub
    End If
    Do
        best = -1
        For i = 1 To oCnt
            If oIdx(i) > 0 Then
                For j = 1 To nCnt
                    If nIdx(j) > 0 Then
                        sc = SameScore(oIdx(i), nIdx(j))
                        If sc > best Then best = sc: bi = i: bj = j
                    End If
                Next j
            End If
        Next i
        If best < 0 Then Exit Do
        Link oIdx(bi), nIdx(bj), 1
        oIdx(bi) = 0: nIdx(bj) = 0
    Loop
End Sub

Private Function SameScore(ByVal o As Long, ByVal n As Long) As Long
    Dim k As Long, s As Long
    For k = 1 To fCount
        If StrComp(oA(o, fOldC(k)), nA(n, fNewC(k)), vbTextCompare) = 0 Then s = s + 1
    Next k
    SameScore = s
End Function

' ------------------------------------------------------------
' Results: statuses, Change_Detail rows and highlight marks
' ------------------------------------------------------------
Private Sub BuildResults()
    Dim r As Long
    For r = 2 To nRowN
        If nOK(r) Then
            If nMate(r) > 0 Then
                ReportPair nMate(r), r
            Else
                ReportAdded r
            End If
        End If
    Next r
    For r = 2 To oRowN
        If oOK(r) And oMate(r) = 0 Then ReportRemoved r
    Next r
    ReportColumns
End Sub

Private Sub ReportPair(ByVal o As Long, ByVal n As Long)
    Dim k As Long, ov As String, nv As String, chgText As String, lbl As String
    cPairs = cPairs + 1
    For k = 1 To fCount
        ov = oA(o, fOldC(k))
        nv = nA(n, fNewC(k))
        If StrComp(ov, nv, vbTextCompare) <> 0 Then
            Select Case fKey(k)
                Case "party name"
                    lbl = "Name changed": cNameChg = cNameChg + 1
                Case "ecid"
                    lbl = "ECID changed": cEcidChg = cEcidChg + 1
                Case Else
                    lbl = "Field changed"
            End Select
            AddDetail o, n, nA(1, fNewC(k)), ov, nv, lbl, True
            AddMark 1, n, fNewC(k)
            AddMark 2, o, fOldC(k)
            If Len(chgText) > 0 Then chgText = chgText & ", "
            chgText = chgText & nA(1, fNewC(k))
        End If
    Next k
    If Len(chgText) = 0 Then
        oStat(o) = "Unchanged": nStat(n) = "Unchanged": cSame = cSame + 1
    Else
        oStat(o) = "Changed": nStat(n) = "Changed": cChanged = cChanged + 1
        nChg(n) = chgText
    End If
End Sub

Private Sub ReportAdded(ByVal n As Long)
    Dim nm As String, st As String, j As Long, v As String, k As Variant
    nm = LCase$(nA(n, nNameC))
    If IsMissingText(nm) Then
        st = "Missing Party Name - review": cReview = cReview + 1
    ElseIf Not oNames.Exists(nm) Then
        st = "Party added": cAdded = cAdded + 1
    ElseIf Not oNameType.Exists(nm & SEP & LCase$(nA(n, nTypeC))) Then
        st = "Party Type added": cAdded = cAdded + 1
    Else
        st = "Record added (same name/role) - review": cReview = cReview + 1
    End If
    nStat(n) = st
    AddMark 4, n, 0
    For Each k In nH.Keys
        j = nH(k)
        v = nA(n, j)
        If Len(v) > 0 Then AddDetail 0, n, nA(1, j), "[none]", v, st, True
    Next k
End Sub

Private Sub ReportRemoved(ByVal o As Long)
    Dim nm As String, st As String, j As Long, v As String, k As Variant
    nm = LCase$(oA(o, oNameC))
    If IsMissingText(nm) Then
        st = "Missing Party Name - review": cReview = cReview + 1
    ElseIf Not nNames.Exists(nm) Then
        st = "Party removed": cRemoved = cRemoved + 1
    ElseIf Not nNameType.Exists(nm & SEP & LCase$(oA(o, oTypeC))) Then
        st = "Party Type removed": cRemoved = cRemoved + 1
    Else
        st = "Record removed (same name/role) - review": cReview = cReview + 1
    End If
    oStat(o) = st
    AddMark 5, o, 0
    For Each k In oH.Keys
        j = oH(k)
        v = oA(o, j)
        If Len(v) > 0 Then AddDetail o, 0, oA(1, j), v, "[none]", st, False
    Next k
End Sub

Private Sub ReportColumns()
    Dim k As Variant
    For Each k In oH.Keys
        If Not nH.Exists(k) Then AddDetail 0, 0, oA(1, oH(k)), "[column exists]", "[column absent]", "Column removed", False
    Next k
    For Each k In nH.Keys
        If Not oH.Exists(k) Then AddDetail 0, 0, nA(1, nH(k)), "[column absent]", "[column exists]", "Column added", True
    Next k
End Sub

Private Sub AddDetail(ByVal o As Long, ByVal n As Long, ByVal fld As String, ByVal ov As String, _
    ByVal nv As String, ByVal chg As String, ByVal hi As Boolean)
    Dim i As Long
    dN = dN + 1
    If dN > dCap Then GrowDetail
    i = dN
    dA(i, 5) = "[File level]"
    If o > 0 Then
        dA(i, 1) = oA(o, oNameC): dA(i, 3) = oA(o, oEcidC): dA(i, 5) = oA(o, oTypeC): dA(i, 10) = CStr(o)
    End If
    If n > 0 Then
        dA(i, 2) = nA(n, nNameC): dA(i, 4) = nA(n, nEcidC): dA(i, 5) = nA(n, nTypeC): dA(i, 11) = CStr(n)
    End If
    dA(i, 6) = fld
    dA(i, 7) = Shown(ov)
    dA(i, 8) = Shown(nv)
    dA(i, 9) = chg
    dY(i) = hi
End Sub

Private Sub GrowDetail()
    Dim t() As String, ty() As Boolean, i As Long, j As Long, newCap As Long
    newCap = dCap * 2
    ReDim t(1 To newCap, 1 To 11): ReDim ty(1 To newCap)
    For i = 1 To dCap
        For j = 1 To 11: t(i, j) = dA(i, j): Next j
        ty(i) = dY(i)
    Next i
    dA = t: dY = ty: dCap = newCap
End Sub

' kind: 1 = current cell yellow, 2 = previous cell orange, 3 = current cell red (missing),
'       4 = current row yellow (new), 5 = previous row orange (removed)
Private Sub AddMark(ByVal kind As Long, ByVal r As Long, ByVal c As Long)
    If mN >= mCap Then
        mCap = mCap * 2
        ReDim Preserve mR(1 To mCap): ReDim Preserve mC(1 To mCap): ReDim Preserve mK(1 To mCap)
    End If
    mN = mN + 1
    mR(mN) = r: mC(mN) = c: mK(mN) = kind
End Sub

' ------------------------------------------------------------
' Mandatory-field check on CURRENT (rules from Requirement1)
' ------------------------------------------------------------
Private Sub CheckMissing()
    Dim r As Long, i As Long, tp As String, ent As String, fk As String
    Dim hasEnt As Boolean, applies As Boolean
    hasEnt = nH.Exists("sub entity type")
    For r = 2 To nRowN
        If nOK(r) Then
            tp = LCase$(nA(r, nTypeC))
            If IsMissingText(nA(r, nNameC)) Then AddMissing r, nNameC, "Party Name"
            If IsMissingText(nA(r, nEcidC)) Then AddMissing r, nEcidC, "ECID"
            If IsMissingText(nA(r, nTypeC)) Then AddMissing r, nTypeC, "Party Type"
            ent = ""
            If hasEnt Then ent = LCase$(nA(r, nH("sub entity type")))
            For i = 1 To rCount
                If rType(i) = tp Or rType(i) = "*" Then
                    applies = True
                    If Len(rScope(i)) > 0 And hasEnt Then
                        If IsMissingText(ent) Then
                            applies = False
                            AddNote r, "Cannot evaluate " & rField(i) & ": Sub Entity Type missing"
                        Else
                            applies = (ent = rScope(i))
                        End If
                    End If
                    If applies Then
                        fk = rKey(i)
                        If Not nH.Exists(fk) Then
                            AddNote r, "Required column absent: " & rField(i)
                        ElseIf IsMissingText(nA(r, nH(fk))) Then
                            AddMissing r, nH(fk), rField(i)
                        End If
                    End If
                End If
            Next i
        End If
    Next r
End Sub

Private Sub AddMissing(ByVal r As Long, ByVal c As Long, ByVal lbl As String)
    Dim k As String
    k = r & ":" & c
    If missSeen.Exists(k) Then Exit Sub
    missSeen.Add k, True
    AddMark 3, r, c
    If Len(nMiss(r)) > 0 Then nMiss(r) = nMiss(r) & ", "
    nMiss(r) = nMiss(r) & lbl
    cMissing = cMissing + 1
End Sub

Private Sub AddNote(ByVal r As Long, ByVal txtNote As String)
    If InStr(1, nNote(r), txtNote, vbTextCompare) > 0 Then Exit Sub
    If Len(nNote(r)) > 0 Then nNote(r) = nNote(r) & "; "
    nNote(r) = nNote(r) & txtNote
End Sub

Private Function DetailText(ByVal r As Long) As String
    Dim s As String
    If Len(nChg(r)) > 0 Then s = "Changed: " & nChg(r)
    If Len(nMiss(r)) > 0 Then
        If Len(s) > 0 Then s = s & " | "
        s = s & "Missing: " & nMiss(r)
    End If
    If Len(nNote(r)) > 0 Then
        If Len(s) > 0 Then s = s & " | "
        s = s & "Note: " & nNote(r)
    End If
    DetailText = s
End Function

' ------------------------------------------------------------
' Output: Previous / Current_Check / Change_Detail
' ------------------------------------------------------------
Private Sub WriteOutputs(ByVal wb As Workbook)
    Dim wsP As Worksheet, wsC As Worksheet, wsD As Worksheet
    Set wsP = wb.Worksheets(1): wsP.Name = "Previous"
    Set wsC = wb.Worksheets.Add(After:=wsP): wsC.Name = "Current_Check"
    Set wsD = wb.Worksheets.Add(After:=wsC): wsD.Name = "Change_Detail"
    DumpPrevious wsP
    DumpCurrent wsC
    DumpDetail wsD
    FormatSheet wsP, oRowN, oC + 1
    FormatSheet wsC, nRowN, nC + 2
    FormatSheet wsD, dN + 1, 11
    ApplyMarks wsP, wsC
End Sub

Private Sub DumpPrevious(ByVal ws As Worksheet)
    Dim out() As String, r As Long, j As Long
    ReDim out(1 To oRowN, 1 To oC + 1)
    For r = 1 To oRowN
        For j = 1 To oC: out(r, j) = oA(r, j): Next j
    Next r
    out(1, oC + 1) = "Check Status"
    For r = 2 To oRowN: out(r, oC + 1) = oStat(r): Next r
    With ws.Range(ws.Cells(1, 1), ws.Cells(oRowN, oC + 1))
        .NumberFormat = "@"
        .Value = out
    End With
End Sub

Private Sub DumpCurrent(ByVal ws As Worksheet)
    Dim out() As String, r As Long, j As Long
    ReDim out(1 To nRowN, 1 To nC + 2)
    For r = 1 To nRowN
        For j = 1 To nC: out(r, j) = nA(r, j): Next j
    Next r
    out(1, nC + 1) = "Check Status"
    out(1, nC + 2) = "Check Detail"
    For r = 2 To nRowN
        If nOK(r) Then
            out(r, nC + 1) = nStat(r)
            out(r, nC + 2) = DetailText(r)
        End If
    Next r
    With ws.Range(ws.Cells(1, 1), ws.Cells(nRowN, nC + 2))
        .NumberFormat = "@"
        .Value = out
    End With
End Sub

Private Sub DumpDetail(ByVal ws As Worksheet)
    Dim out() As String, i As Long, j As Long, hdr As Variant
    hdr = Array("Party Name (Previous)", "Party Name (Current)", "ECID (Previous)", "ECID (Current)", _
        "Party Type", "Field", "Previous Value", "Current Value", "Change", "Previous Row", "Current Row")
    ReDim out(1 To dN + 1, 1 To 11)
    For j = 1 To 11: out(1, j) = hdr(j - 1): Next j
    For i = 1 To dN
        For j = 1 To 11: out(i + 1, j) = dA(i, j): Next j
    Next i
    With ws.Range(ws.Cells(1, 1), ws.Cells(dN + 1, 11))
        .NumberFormat = "@"
        .Value = out
    End With
    For i = 1 To dN
        If dY(i) Then ws.Cells(i + 1, 8).Interior.Color = ClrYellow()
    Next i
End Sub

Private Sub ApplyMarks(ByVal wsP As Worksheet, ByVal wsC As Worksheet)
    Dim m As Long
    For m = 1 To mN
        If mK(m) = 4 Then wsC.Range(wsC.Cells(mR(m), 1), wsC.Cells(mR(m), nC + 2)).Interior.Color = ClrYellow()
        If mK(m) = 5 Then wsP.Range(wsP.Cells(mR(m), 1), wsP.Cells(mR(m), oC + 1)).Interior.Color = ClrOrange()
    Next m
    For m = 1 To mN
        If mK(m) = 1 Then wsC.Cells(mR(m), mC(m)).Interior.Color = ClrYellow()
        If mK(m) = 2 Then wsP.Cells(mR(m), mC(m)).Interior.Color = ClrOrange()
    Next m
    For m = 1 To mN
        If mK(m) = 3 Then
            wsC.Cells(mR(m), mC(m)).Interior.Color = ClrRed()
            wsC.Cells(mR(m), mC(m)).Font.Color = RGB(150, 0, 0)
        End If
    Next m
    For m = 2 To nRowN
        If Len(nMiss(m)) > 0 Then wsC.Cells(m, nC + 2).Font.Color = RGB(150, 0, 0)
    Next m
End Sub

Private Sub FormatSheet(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    Dim j As Long
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
        .Font.Name = "Calibri"
        .Font.Size = 11
        .VerticalAlignment = xlTop
        .Columns.AutoFit
    End With
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Font.Bold = True
        .Interior.Color = RGB(210, 225, 240)
    End With
    For j = 1 To lastCol
        If ws.Columns(j).ColumnWidth > 45 Then ws.Columns(j).ColumnWidth = 45
        If ws.Columns(j).ColumnWidth < 10 Then ws.Columns(j).ColumnWidth = 10
    Next j
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
End Sub

Private Function ClrYellow() As Long
    ClrYellow = RGB(255, 240, 150)
End Function

Private Function ClrOrange() As Long
    ClrOrange = RGB(255, 217, 179)
End Function

Private Function ClrRed() As Long
    ClrRed = RGB(255, 190, 190)
End Function

' ------------------------------------------------------------
' Reading files -> string arrays (1-based, row 1 = header)
' ------------------------------------------------------------
Private Sub ReadFile(ByVal path As String, ByRef arr() As String, ByRef rowCount As Long, ByRef colCount As Long)
    Dim src As Workbook, sh As Worksheet, wbk As Workbook, owned As Boolean
    Dim lr As Long, lc As Long, v As Variant, i As Long, j As Long
    Dim pick As String, errText As String, fields() As Variant, n As Long
    On Error GoTo BadLoad
    For Each wbk In Application.Workbooks
        If StrComp(wbk.FullName, path, vbTextCompare) = 0 Then Set src = wbk: Exit For
    Next wbk
    If src Is Nothing Then
        owned = True
        If LCase$(Right$(path, 4)) = ".csv" Then
            n = CsvFieldCount(path) + 5
            If n > 16384 Then n = 16384
            ReDim fields(0 To n - 1)
            For i = 0 To n - 1: fields(i) = Array(i + 1, 2): Next i
            Workbooks.OpenText Filename:=path, Origin:=65001, DataType:=xlDelimited, _
                TextQualifier:=xlTextQualifierDoubleQuote, Comma:=True, FieldInfo:=fields, Local:=False
            Set src = ActiveWorkbook
        Else
            Set src = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
        End If
    End If
    Set sh = src.Worksheets(1)
    If src.Worksheets.Count > 1 Then
        pick = InputBox("Enter the data worksheet name in " & src.Name, "Worksheet", sh.Name)
        If Len(pick) = 0 Then Err.Raise 5, , "Worksheet selection cancelled."
        Set sh = src.Worksheets(pick)
    End If
    lr = LastRowOf(sh): lc = LastColOf(sh)
    If lr < 2 Then Err.Raise 5, , "No data rows found in " & src.Name
    v = sh.Range(sh.Cells(1, 1), sh.Cells(lr, lc)).Value
    ReDim arr(1 To lr, 1 To lc)
    For i = 1 To lr
        For j = 1 To lc
            arr(i, j) = CellText(v(i, j))
        Next j
    Next i
    rowCount = lr: colCount = lc
    If owned Then src.Close False
    Exit Sub
BadLoad:
    errText = Err.Description
    On Error Resume Next
    If owned And Not src Is Nothing Then src.Close False
    On Error GoTo 0
    Err.Raise 5, , errText
End Sub

Private Function CsvFieldCount(ByVal path As String) As Long
    Dim ff As Integer, ln As String, i As Long, q As Boolean, n As Long, ch As String
    ff = FreeFile
    Open path For Input As #ff
    If Not EOF(ff) Then Line Input #ff, ln
    Close #ff
    n = 1
    For i = 1 To Len(ln)
        ch = Mid$(ln, i, 1)
        If ch = """" Then
            q = Not q
        ElseIf ch = "," And Not q Then
            n = n + 1
        End If
    Next i
    CsvFieldCount = n
End Function

' Cleans a cell into plain comparable text (dates -> yyyy-mm-dd, no scientific numbers, single spaces).
Private Function CellText(ByVal v As Variant) As String
    Dim t As String
    Select Case VarType(v)
        Case vbEmpty, vbNull
            Exit Function
        Case vbError
            CellText = "[EXCEL ERROR]"
            Exit Function
        Case vbDate
            If v = Int(v) Then
                CellText = Format$(v, "yyyy-mm-dd")
            Else
                CellText = Format$(v, "yyyy-mm-dd hh:nn:ss")
            End If
            Exit Function
        Case vbDouble, vbSingle, vbCurrency, vbLong, vbInteger
            If v = Fix(v) And Abs(v) < 1E+15 Then t = Format$(v, "0") Else t = CStr(v)
        Case Else
            t = CStr(v)
    End Select
    t = Replace(Replace(Replace(Replace(Replace(t, ChrW(65279), ""), ChrW(160), " "), vbLf, " "), vbCr, " "), vbTab, " ")
    Do While InStr(t, "  ") > 0: t = Replace(t, "  ", " "): Loop
    CellText = Trim$(t)
End Function

Private Sub NormDates(ByRef arr() As String, ByVal rowCount As Long, ByVal col As Long)
    Dim r As Long, s As String
    For r = 2 To rowCount
        s = arr(r, col)
        If Len(s) > 0 Then
            If IsDate(s) Then arr(r, col) = Format$(CDate(s), "yyyy-mm-dd")
        End If
    Next r
End Sub

Private Function IsDataRow(ByRef arr() As String, ByVal r As Long, ByVal c As Long) As Boolean
    Dim j As Long, hasData As Boolean
    For j = 1 To c
        If Len(arr(r, j)) > 0 Then hasData = True: Exit For
    Next j
    If hasData Then
        ' Ignore known export footer instead of treating it as a person.
        IsDataRow = (InStr(1, arr(r, 1), "internal use only", vbTextCompare) = 0)
    End If
End Function

Private Function HeaderMap(ByRef arr() As String, ByVal c As Long, ByVal label As String) As Object
    Dim d As Object, j As Long, s As String, baseKey As String, k As Long
    Set d = CreateObject("Scripting.Dictionary")
    For j = 1 To c
        baseKey = HeaderKey(arr(1, j))
        If Len(baseKey) > 0 Then
            s = baseKey: k = 1
            Do While d.Exists(s)
                k = k + 1
                s = baseKey & " #" & k
            Loop
            d.Add s, j
        End If
    Next j
    If d.Count = 0 Then Err.Raise 5, , "No header row found in " & label
    Set HeaderMap = d
End Function

Private Function PickHeader(ByVal d As Object, ByVal label As String, ByVal fileLabel As String) As Long
    Dim s As String
    If d.Exists(HeaderKey(label)) Then
        PickHeader = d(HeaderKey(label))
        Exit Function
    End If
    s = InputBox("Enter the exact header of " & label & " in " & fileLabel & ". (Do not choose Party Role in Reference Data for Party Type.)", "Identify column", label)
    If Len(s) = 0 Then Err.Raise 5, , "Header not found: " & label & " (" & fileLabel & ")"
    If Not d.Exists(HeaderKey(s)) Then Err.Raise 5, , "Header not found: " & s
    PickHeader = d(HeaderKey(s))
End Function

' ------------------------------------------------------------
' Requirement1 (mandatory-field rules, from the email mapping)
' ------------------------------------------------------------
Private Function GetRequirementSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Requirement1")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "Requirement1"
        PopulateEmailRules ws
    End If
    Set GetRequirementSheet = ws
End Function

Public Sub SetupRequirement1V6()
    Dim ws As Worksheet
    On Error GoTo Failed
    Set ws = GetRequirementSheet()
    ValidateRules ws
    ws.Activate
    MsgBox "Requirement1 is ready. Existing rules were preserved. Set Mode to OPTIONAL to switch a rule off. Save your .xlsm after editing.", vbInformation
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
End Sub

' Each item: Array(Party Type, Field, Sub Entity Type scope, Notes)
Private Sub BuildDefaultRules(ByVal lst As Collection)
    Dim roles As Variant, flds As Variant, a As Variant, b As Variant
    roles = Array("Immediate Shareholder", "Board Member/Director", "Authorized Signer", _
        "China LDD Beneficial Owner", "Legal Representative or Equivalent Party")
    For Each a In roles
        lst.Add Array(CStr(a), "Country Tagging", "", "Y means a value is required, not literally Y")
    Next a
    roles = Array("Authorized Signer", "China LDD Beneficial Owner", "Legal Representative or Equivalent Party")
    flds = Array("DOB", "Country of Nationality", "Government Issued ID Type", _
        "Government ID Issuing Authority", "ID Number")
    For Each a In roles
        For Each b In flds
            lst.Add Array(CStr(a), CStr(b), "", "")
        Next b
    Next a
    lst.Add Array("China LDD Beneficial Owner", "Gender", "", "")
    lst.Add Array("*", "Gender", "Individual", "Applies to all Individual records. Set OPTIONAL to disable.")
End Sub

Private Sub PopulateEmailRules(ByVal ws As Worksheet)
    Dim lst As Collection, i As Long, it As Variant
    Set lst = New Collection
    BuildDefaultRules lst
    PutRow ws, 1, Array("Party Type", "Field", "Mode", "Source", "Sub Entity Type", "Notes")
    For i = 1 To lst.Count
        it = lst(i)
        PutRow ws, i + 1, Array(it(0), it(1), "REQUIRED", "Email mapping", it(2), it(3))
    Next i
    ws.Rows(1).Font.Bold = True
    ws.Range("A:F").Columns.AutoFit
    ws.Columns(6).ColumnWidth = 55
    ws.Range("A1:F" & (lst.Count + 1)).AutoFilter
    With ws.Range("C2:C" & (lst.Count + 1)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Formula1:="REQUIRED" & Application.International(xlListSeparator) & "OPTIONAL"
        .ErrorTitle = "Rule mode"
        .ErrorMessage = "Use REQUIRED or OPTIONAL."
        .ShowError = True
    End With
End Sub

Private Sub PutRow(ByVal ws As Worksheet, ByVal r As Long, ByVal vals As Variant)
    Dim j As Long
    For j = LBound(vals) To UBound(vals)
        ws.Cells(r, j + 1).NumberFormat = "@"
        ws.Cells(r, j + 1).Value2 = CStr(vals(j))
    Next j
End Sub

Private Sub ValidateRules(ByVal ws As Worksheet)
    Dim r As Long, m As String
    If Norm(ws.Cells(1, 1).Value2) <> "party type" Or Norm(ws.Cells(1, 2).Value2) <> "field" _
        Or Norm(ws.Cells(1, 3).Value2) <> "mode" Then
        Err.Raise 5, , "Requirement1: columns A/B/C must be Party Type / Field / Mode (E = Sub Entity Type)."
    End If
    For r = 2 To LastRowOf(ws)
        If Len(Norm(ws.Cells(r, 1).Value2)) > 0 Or Len(Norm(ws.Cells(r, 2).Value2)) > 0 Then
            If Len(Norm(ws.Cells(r, 1).Value2)) = 0 Or Len(HeaderKey(ws.Cells(r, 2).Value2)) = 0 Then
                Err.Raise 5, , "Incomplete Requirement1 row " & r
            End If
            m = UCase$(Trim$(Txt(ws.Cells(r, 3).Value2)))
            If m <> "REQUIRED" And m <> "OPTIONAL" Then
                Err.Raise 5, , "Requirement1 row " & r & ": Mode must be REQUIRED or OPTIONAL."
            End If
        End If
    Next r
End Sub

Private Sub LoadRules(ByVal ws As Worksheet)
    Dim lst As Collection, r As Long
    Set lst = New Collection
    For r = 2 To LastRowOf(ws)
        If Len(Norm(ws.Cells(r, 1).Value2)) > 0 Then
            If UCase$(Trim$(Txt(ws.Cells(r, 3).Value2))) = "REQUIRED" Then
                lst.Add Array(Norm(ws.Cells(r, 1).Value2), Txt(ws.Cells(r, 2).Value2), Norm(ws.Cells(r, 5).Value2))
            End If
        End If
    Next r
    LoadRulesFromList lst
End Sub

Private Sub LoadRulesFromList(ByVal lst As Collection)
    Dim i As Long, it As Variant
    rCount = lst.Count
    If rCount = 0 Then Exit Sub
    ReDim rType(1 To rCount): ReDim rField(1 To rCount): ReDim rKey(1 To rCount): ReDim rScope(1 To rCount)
    For i = 1 To rCount
        it = lst(i)
        rType(i) = LCase$(Trim$(CStr(it(0))))
        rField(i) = CStr(it(1))
        rKey(i) = HeaderKey(it(1))
        rScope(i) = LCase$(Trim$(CStr(it(2))))
    Next i
End Sub

' ------------------------------------------------------------
' Small helpers
' ------------------------------------------------------------
Private Function HeaderKey(ByVal value As Variant) As String
    Dim resultText As String
    resultText = Norm(value)
    Select Case resultText
        Case "partyname", "party name": resultText = "party name"
        Case "partytype", "party type": resultText = "party type"
        Case "date of birth", "dob": resultText = "dob"
        Case "nationality", "country of nationality": resultText = "country of nationality"
        Case "government issued id type", "government id type": resultText = "government issued id type"
        Case "government id issuing authority", "government issued id issuing authority": resultText = "government id issuing authority"
        Case "id number", "government id number": resultText = "id number"
        Case "sub entity type", "sub-entity type": resultText = "sub entity type"
    End Select
    HeaderKey = resultText
End Function

Private Function IsMissingText(ByVal s As String) As Boolean
    Select Case LCase$(Trim$(s))
        Case "", "--", "-", "n/a", "null": IsMissingText = True
    End Select
End Function

Private Function Shown(ByVal s As String) As String
    Shown = s
    If Len(s) = 0 Then Shown = "[blank]"
End Function

Private Function Txt(ByVal v As Variant) As String
    If IsError(v) Then
        Txt = "[EXCEL ERROR]"
    Else
        Txt = CStr(v)
    End If
End Function

Private Function Norm(ByVal v As Variant) As String
    Dim t As String
    t = Replace(Replace(Replace(Replace(Txt(v), ChrW(160), " "), vbLf, " "), vbCr, " "), vbTab, " ")
    Do While InStr(t, "  ") > 0: t = Replace(t, "  ", " "): Loop
    Norm = LCase$(Trim$(t))
End Function

Private Function LastRowOf(ByVal ws As Worksheet) As Long
    Dim c As Range
    Set c = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False, _
        SearchFormat:=False, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If c Is Nothing Then LastRowOf = 1 Else LastRowOf = c.Row
End Function

Private Function LastColOf(ByVal ws As Worksheet) As Long
    Dim c As Range
    Set c = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False, _
        SearchFormat:=False, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If c Is Nothing Then LastColOf = 1 Else LastColOf = c.Column
End Function

' ------------------------------------------------------------
' Self-test: runs the engine on synthetic arrays (no files, no sheets touched)
' ------------------------------------------------------------
Public Sub TestKYCCheckV6()
    Dim lst As Collection, i As Long, genderCol As Long, foundMark As Boolean
    Dim cntField As Long, cntEcid As Long, cntName As Long
    On Error GoTo Failed
    oRowN = 9: oC = 12: nRowN = 9: nC = 12
    ReDim oA(1 To oRowN, 1 To oC): ReDim nA(1 To nRowN, 1 To nC)
    SetRow oA, 1, Array("Party Name", "ECID", "Party Type", "Address", "DOB", "Gender", "Country of Nationality", _
        "Government Issued ID Type", "Government ID Issuing Authority", "ID Number", "Country Tagging", "Sub Entity Type")
    SetRow nA, 1, Array("ECID", "Party Type", "Party Name", "Gender", "Address", "DOB", "Country Tagging", _
        "Country of Nationality", "ID Number", "Government ID Issuing Authority", "Government Issued ID Type", "Sub Entity Type")
    FixOld 2, "Alpha", "100", "Authorized Signer", "A", "Female"
    FixOld 3, "Alpha", "100", "Board Member/Director", "A", "Female"
    FixOld 4, "Alpha", "100", "Senior Manager or Equivalent", "A", "Female"
    FixOld 5, "Switch", "001", "Authorized Signer", "A", "Male"
    FixOld 6, "Gone", "300", "Parent Entity", "A", "Male"
    FixOld 7, "Rename Old", "700", "Authorized Signer", "A", "Male"
    FixOld 8, "Repeat", "400", "Other Related Parties", "X", "Male"
    FixOld 9, "Repeat", "400", "Other Related Parties", "Y", "Male"
    ' rows and columns deliberately reordered in the current file
    FixNew 2, "Repeat", "400", "Other Related Parties", "Z", "Male"
    FixNew 3, "Alpha", "100", "Authorized Signer", "B", ""
    FixNew 4, "Switch", "002", "Authorized Signer", "C", "Male"
    FixNew 5, "Alpha", "100", "Senior Manager or Equivalent", "A", "Female"
    FixNew 6, "New Person", "600", "Authorized Signer", "A", "Male"
    FixNew 7, "Repeat", "400", "Other Related Parties", "X", "Male"
    FixNew 8, "Alpha", "100", "China LDD Beneficial Owner", "A", "Female"
    FixNew 9, "Rename New", "700", "Authorized Signer", "A", "Male"
    PrepareColumns
    Set lst = New Collection
    BuildDefaultRules lst
    LoadRulesFromList lst
    RunEngine
    For i = 1 To dN
        If dA(i, 9) = "Field changed" Then cntField = cntField + 1
        If dA(i, 9) = "ECID changed" Then
            cntEcid = cntEcid + 1
            AssertV6 dA(i, 7) = "001" And dA(i, 8) = "002", "ECID leading zeros lost"
            AssertV6 dY(i), "ECID change row not yellow"
        End If
        If dA(i, 9) = "Name changed" Then
            cntName = cntName + 1
            AssertV6 dA(i, 7) = "Rename Old" And dA(i, 8) = "Rename New", "Name change values wrong"
            AssertV6 dY(i), "Name change row not yellow"
        End If
    Next i
    AssertV6 cPairs = 6, "Expected 6 matched records, got " & cPairs
    AssertV6 cntField = 4, "Expected 4 field changes, got " & cntField
    AssertV6 cntEcid = 1, "Expected 1 ECID change, got " & cntEcid
    AssertV6 cntName = 1, "Expected 1 name change, got " & cntName
    AssertV6 cAdded = 2, "Expected 2 added, got " & cAdded
    AssertV6 cRemoved = 2, "Expected 2 removed, got " & cRemoved
    AssertV6 cReview = 0, "Unexpected review items: " & cReview
    AssertV6 nStat(6) = "Party added", "New Person status wrong: " & nStat(6)
    AssertV6 nStat(8) = "Party Type added", "Alpha LDD status wrong: " & nStat(8)
    AssertV6 oStat(6) = "Party removed", "Gone status wrong: " & oStat(6)
    AssertV6 oStat(3) = "Party Type removed", "Alpha Board status wrong: " & oStat(3)
    AssertV6 cMissing = 1, "Expected 1 missing mandatory field, got " & cMissing
    genderCol = nH("gender")
    For i = 1 To mN
        If mK(i) = 3 And mR(i) = 3 And mC(i) = genderCol Then foundMark = True
    Next i
    AssertV6 foundMark, "Blank Gender on current row 3 not flagged"
    MsgBox "SELF-TEST PASSED: name change, ECID change, field changes, added/removed party and role, " & _
        "repeated rows, shuffled rows/columns, mandatory-field check.", vbInformation
    Exit Sub
Failed:
    MsgBox "SELF-TEST FAILED: " & Err.Description, vbExclamation
End Sub

Private Sub SetRow(ByRef arr() As String, ByVal r As Long, ByVal v As Variant)
    Dim j As Long
    For j = LBound(v) To UBound(v)
        arr(r, j + 1) = CStr(v(j))
    Next j
End Sub

Private Sub FixOld(ByVal r As Long, ByVal nm As String, ByVal eid As String, ByVal roleName As String, _
    ByVal addr As String, ByVal gender As String)
    SetRow oA, r, Array(nm, eid, roleName, addr, "1990-01-01", gender, "CN", "Passport", "CN", "P0001", "CN", "Individual")
End Sub

Private Sub FixNew(ByVal r As Long, ByVal nm As String, ByVal eid As String, ByVal roleName As String, _
    ByVal addr As String, ByVal gender As String)
    SetRow nA, r, Array(eid, roleName, nm, gender, addr, "1990-01-01", "CN", "CN", "P0001", "CN", "Passport", "Individual")
End Sub

Private Sub AssertV6(ByVal condition As Boolean, ByVal messageText As String)
    If Not condition Then Err.Raise vbObjectError + 606, "KYC V6 self-test", messageText
End Sub
