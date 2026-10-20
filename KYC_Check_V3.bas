Attribute VB_Name = "KYC_Check_V3"
Option Explicit
' Windows desktop Excel. No additional VBA references required.
' Source files are read only. Header row must be row 1.
Public Sub RunKYCCheckV3()
    Dim p As Variant, q As Variant, a As Worksheet, b As Worksheet
    Dim out As Workbook, ch As Worksheet, issues As Worksheet, rules As Worksheet
    Dim ha As Object, hb As Object, ia As Object, ib As Object
    Dim ea As Long, eb As Long, ta As Long, tb As Long
    Dim r As Long, j As Long, oldRow As Long, k As Variant, h As Variant, typ As String
    Dim v As String, w As String, ruleIndex As Object, rk As String, n As Long
    Dim standard As Worksheet, hs As Object, es As Long, ts As Long, z As Variant
    Dim oldSecurity As Long, oldEvents As Boolean, oldScreen As Boolean
    oldSecurity = Application.AutomationSecurity
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    p = Application.GetOpenFilename("Data files (*.csv;*.xlsx;*.xlsm;*.xls),*.csv;*.xlsx;*.xlsm;*.xls", , "Select PREVIOUS")
    If VarType(p) = vbBoolean Then Exit Sub
    q = Application.GetOpenFilename("Data files (*.csv;*.xlsx;*.xlsm;*.xls),*.csv;*.xlsx;*.xlsm;*.xls", , "Select CURRENT")
    If VarType(q) = vbBoolean Then Exit Sub
    If StrComp(CStr(p), CStr(q), vbTextCompare) = 0 Then Err.Raise 5, , "Select two different files."
    Application.AutomationSecurity = 3
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set out = Workbooks.Add(xlWBATWorksheet)
    Set a = out.Worksheets(1): a.Name = "Previous"
    Set b = out.Worksheets.Add(After:=a): b.Name = "Current_Check"
    LoadFile CStr(p), a: LoadFile CStr(q), b
    Set ha = Headers(a): Set hb = Headers(b)
    ea = SelectHeader(a, ha, "ECID", "ECID")
    eb = SelectHeader(b, hb, "ECID", "ECID")
    ta = SelectHeader(a, ha, "Party Type", "Party Type")
    tb = SelectHeader(b, hb, "Party Type", "Party Type")
    Set ch = out.Worksheets.Add(After:=b): ch.Name = "Change_Detail"
    Set issues = out.Worksheets.Add(After:=ch): issues.Name = "Check_Issues"
    WriteRow ch, Array("ECID", "Party Type", "Field", "Previous", "Current", "Change", "Previous Row", "Current Row")
    WriteRow issues, Array("ECID", "Party Type", "Field", "Value", "Issue", "Source", "Row")
    ResolveDuplicates a, b, ha, hb, ea, eb, ta, tb, issues
    Set ia = IndexRows(a, ea, ta, issues): Set ib = IndexRows(b, eb, tb, issues)
    Dim scope As String, rr As Long, ruleType As String, ruleScope As String
    On Error Resume Next
    Set rules = ThisWorkbook.Worksheets("Requirement1")
    On Error GoTo Failed
    If rules Is Nothing Then
    Set standard = a: Set hs = ha: es = ea: ts = ta
    Application.ScreenUpdating = True
    If MsgBox("Use PREVIOUS as the standard for required fields? Yes = Previous; No = choose a separate standard file.", vbYesNo) = vbNo Then
        z = Application.GetOpenFilename("Data files (*.csv;*.xlsx;*.xlsm;*.xls),*.csv;*.xlsx;*.xlsm;*.xls", , "Select STANDARD")
        If VarType(z) = vbBoolean Then Err.Raise 5, , "Standard selection cancelled."
        Set standard = out.Worksheets.Add: standard.Name = "Standard"
        LoadFile CStr(z), standard
        Set hs = Headers(standard)
        es = SelectHeader(standard, hs, "ECID", "ECID")
        ts = SelectHeader(standard, hs, "Party Type", "Party Type")
    End If
    Application.ScreenUpdating = False
    Set rules = ThisWorkbook.Worksheets.Add: rules.Name = "Requirement1"
    WriteRow rules, Array("Party Type", "Field", "Mode", "Standard Row", "Sub Entity Type")
    Set ruleIndex = CreateObject("Scripting.Dictionary")
    For r = 2 To LastRow(standard)
        If Not Missing(standard.Cells(r, es).Value2) And Not Missing(standard.Cells(r, ts).Value2) Then
            typ = Txt(standard.Cells(r, ts).Value2): scope = EntityScope(standard, hs, r)
            For Each h In hs.Keys
                j = hs(h)
                If j <> es And j <> ts And h <> "__match_slot" Then
                    If Not Missing(standard.Cells(r, j).Value2) Then
                        rk = PairKey(PairKey(Norm(typ), scope), CStr(h))
                        If Not ruleIndex.Exists(rk) Then
                            ruleIndex.Add rk, True
                            WriteRow rules, Array(typ, standard.Cells(1, j).Value2, "REQUIRED", r, scope)
                        End If
                    End If
                End If
            Next h
        End If
    Next r
    End If
    If Norm(rules.Cells(1, 1).Value2) <> "party type" Or Norm(rules.Cells(1, 2).Value2) <> "field" Or Norm(rules.Cells(1, 3).Value2) <> "mode" Then
        Err.Raise 5, , "Requirement1 must have Party Type / Field / Mode in columns A / B / C."
    End If
    For r = 2 To LastRow(b)
        If Not Missing(b.Cells(r, eb).Value2) Then
            typ = Txt(b.Cells(r, tb).Value2): scope = EntityScope(b, hb, r)
            Dim covered As Boolean: covered = False
            For rr = 2 To LastRow(rules)
                ruleType = Norm(rules.Cells(rr, 1).Value2): ruleScope = Txt(rules.Cells(rr, 5).Value2)
                If ruleType = Norm(typ) And (Norm(ruleScope) = scope Or ruleScope = "") Then
                    covered = True: h = Norm(rules.Cells(rr, 2).Value2)
                    Dim mode As String: mode = UCase$(Trim$(Txt(rules.Cells(rr, 3).Value2)))
                    If mode <> "REQUIRED" And mode <> "OPTIONAL" Then
                        WriteRow issues, Array(b.Cells(r, eb).Value2, typ, rules.Cells(rr, 2).Value2, mode, "Invalid Requirement1 Mode", "Requirement1", rr)
                    End If
                    If mode = "REQUIRED" Then
                    If Not hb.Exists(h) Then
                        WriteRow issues, Array(b.Cells(r, eb).Value2, typ, rules.Cells(rr, 2).Value2, "", "Required column missing", "Current", r)
                    Else
                        j = hb(h)
                        If Missing(b.Cells(r, j).Value2) Then
                            WriteRow issues, Array(b.Cells(r, eb).Value2, typ, b.Cells(1, j).Value2, b.Cells(r, j).Value2, "Required information missing (standard-derived)", "Current", r)
                            b.Cells(r, j).Font.Color = RGB(190, 0, 0)
                            b.Cells(r, j).Interior.Color = RGB(255, 190, 190)
                        End If
                    End If
                    End If
                End If
            Next rr
            If Not covered Then WriteRow issues, Array(b.Cells(r, eb).Value2, typ, "Requirements", "", "Type / entity scope absent from standard", "Current", r)
            BasicChecks b, hb, r, eb, tb, issues
        End If
    Next r
    For Each h In ha.Keys
        If Not hb.Exists(h) Then WriteRow ch, Array("", "", a.Cells(1, ha(h)).Value2, "Column exists", "Column absent", "Column removed", "", "")
    Next h
    For Each h In hb.Keys
        If Not ha.Exists(h) Then WriteRow ch, Array("", "", b.Cells(1, hb(h)).Value2, "Column absent", "Column exists", "Column added", "", "")
    Next h
    For Each k In ib.Keys
        r = ib(k)
        If r > 0 Then
            If ia.Exists(k) Then
                oldRow = ia(k)
                If oldRow > 0 Then
                    For Each h In hb.Keys
                        If ha.Exists(h) Then
                            j = hb(h)
                            v = Txt(a.Cells(oldRow, ha(h)).Value2): w = Txt(b.Cells(r, j).Value2)
                            If v <> w And h <> "__match_slot" Then
                                WriteRow ch, Array(b.Cells(r, eb).Value2, b.Cells(r, tb).Value2, b.Cells(1, j).Value2, v, w, "Changed", oldRow, r)
                                b.Cells(r, j).Interior.Color = RGB(255, 240, 150)
                                ch.Cells(LastRow(ch), 5).Interior.Color = RGB(255, 240, 150)
                            End If
                        End If
                    Next h
                End If
            Else
                For Each h In hb.Keys
                    If h <> "__match_slot" Then WriteRow ch, Array(b.Cells(r, eb).Value2, b.Cells(r, tb).Value2, b.Cells(1, hb(h)).Value2, "", b.Cells(r, hb(h)).Value2, "Not matched in Previous", "", r)
                Next h
            End If
        End If
    Next k
    For Each k In ia.Keys
        r = ia(k)
        If r > 0 And Not ib.Exists(k) Then
            For Each h In ha.Keys
                If h <> "__match_slot" Then WriteRow ch, Array(a.Cells(r, ea).Value2, a.Cells(r, ta).Value2, a.Cells(1, ha(h)).Value2, a.Cells(r, ha(h)).Value2, "", "Not matched in Current", r, "")
            Next h
        End If
    Next k
    rules.Copy After:=out.Worksheets(out.Worksheets.Count)
    Dim summary As Worksheet
    Set summary = BuildSummary(a, b, ha, hb, ia, ib, ea, eb, ta, tb, ch, issues)
    Dim ws As Worksheet
    For Each ws In out.Worksheets
        ws.Rows(1).Font.Bold = True
        ws.Rows(1).Interior.Color = RGB(210, 225, 240)
        ws.UsedRange.AutoFilter
        ws.UsedRange.Columns.AutoFit
        For j = 1 To LastCol(ws)
            If ws.Columns(j).ColumnWidth > 45 Then ws.Columns(j).ColumnWidth = 45
        Next j
    Next ws
    a.Columns(ha("__match_slot")).Hidden = True
    b.Columns(hb("__match_slot")).Hidden = True
    out.Activate: summary.Activate
    summary.Range("A2").Select
    ActiveWindow.FreezePanes = True
    MsgBox "Complete. See Party_Summary and Change_Detail. Edit and SAVE Requirement1 in the MACRO workbook; future runs reuse it. Result contains a snapshot. Yellow = change. Red = issue.", vbInformation
Done:
    Application.AutomationSecurity = oldSecurity
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    Exit Sub
Failed:
    MsgBox "Stopped: " & Err.Description & vbCrLf & "Partial output must not be treated as complete.", vbExclamation
    Resume Done
End Sub

Private Sub LoadFile(ByVal path As String, ByVal target As Worksheet)
    Dim src As Workbook, sh As Worksheet, fields() As Variant, i As Long
    Dim count As Long, pick As String, errText As String
    Dim wb As Workbook, owned As Boolean, tempPath As String, ext As String
    On Error GoTo BadLoad
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then Set src = wb: Exit For
    Next wb
    If src Is Nothing Then
        ext = Mid$(path, InStrRev(path, "."))
        tempPath = Environ$("TEMP") & "\KYC_" & Format$(Now, "yyyymmddhhnnss") & "_" & CStr(CLng(Timer * 100)) & ext
        FileCopy path, tempPath
        path = tempPath: owned = True
    If LCase$(Right$(path, 4)) = ".csv" Then
        ' Excel sheet limit; all CSV columns imported as text.
        ReDim fields(0 To 16383)
        For i = 0 To 16383: fields(i) = Array(i + 1, 2): Next i
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
    target.Cells.NumberFormat = "@"
    target.Range("A1").Resize(LastRow(sh), LastCol(sh)).Value2 = sh.Range("A1").Resize(LastRow(sh), LastCol(sh)).Value2
    If owned Then src.Close False
    If Len(tempPath) > 0 Then Kill tempPath
    Exit Sub
BadLoad:
    errText = Err.Description
    On Error Resume Next
    If owned And Not src Is Nothing Then src.Close False
    If Len(tempPath) > 0 Then Kill tempPath
    On Error GoTo 0
    Err.Raise 5, , errText
End Sub
Private Function Headers(ByVal ws As Worksheet) As Object
    Dim d As Object, j As Long, s As String
    Set d = CreateObject("Scripting.Dictionary")
    For j = 1 To LastCol(ws)
        s = Norm(ws.Cells(1, j).Value2)
        If Len(s) = 0 Then Err.Raise 5, , "Blank header in " & ws.Name & " column " & j
        If d.Exists(s) Then Err.Raise 5, , "Duplicate header: " & s
        d.Add s, j
    Next j
    Set Headers = d
End Function
Private Function SelectHeader(ByVal ws As Worksheet, ByVal d As Object, ByVal label As String, ByVal suggested As String) As Long
    Dim s As String
    If d.Exists(Norm(suggested)) Then
        SelectHeader = d(Norm(suggested)): Exit Function
    End If
    s = InputBox("Enter exact " & label & " header in " & ws.Name & ". Do not choose Party Role in Reference Data for Party Type.", "Identify column", suggested)
    If Not d.Exists(Norm(s)) Then Err.Raise 5, , "Header not found: " & s
    SelectHeader = d(Norm(s))
End Function
Private Function IndexRows(ByVal ws As Worksheet, ByVal ec As Long, ByVal tp As Long, ByVal issues As Worksheet) As Object
    Dim d As Object, r As Long, k As String, e As String, t As String
    Set d = CreateObject("Scripting.Dictionary")
    For r = 2 To LastRow(ws)
        e = Txt(ws.Cells(r, ec).Value2): t = Txt(ws.Cells(r, tp).Value2)
        If Txt(ws.Cells(r, LastCol(ws)).Value2) = "UNRESOLVED" Then
            WriteRow issues, Array(e, t, "Matching", "", "Unpaired duplicate-group record; inspect Match_Log", ws.Name, r)
        ElseIf Missing(e) Or Missing(t) Then
            WriteRow issues, Array(e, t, "Key", "", "Missing ECID or Party Type; row not matched (including possible footer)", ws.Name, r)
        Else
            k = PairKey(PairKey(Trim$(e), Norm(t)), Txt(ws.Cells(r, LastCol(ws)).Value2))
            If d.Exists(k) Then
                If d(k) > 0 Then WriteRow issues, Array(e, t, "Key", "", "Duplicate ECID + Party Type; comparison skipped", ws.Name, d(k))
                d(k) = 0
                WriteRow issues, Array(e, t, "Key", "", "Duplicate ECID + Party Type; comparison skipped", ws.Name, r)
            Else
                d.Add k, r
            End If
        End If
    Next r
    Set IndexRows = d
End Function
Private Sub BasicChecks(ByVal ws As Worksheet, ByVal hd As Object, ByVal r As Long, ByVal ec As Long, ByVal tp As Long, ByVal issues As Worksheet)
    Dim x As Long, y As Long, v As String, w As String
    If hd.Exists("gender") And IsIndividual(ws, hd, r) Then
        x = hd("gender")
        If Missing(ws.Cells(r, x).Value2) Then
            WriteRow issues, Array(ws.Cells(r, ec).Value2, ws.Cells(r, tp).Value2, "Gender", ws.Cells(r, x).Value2, "Gender missing for Individual", "Current", r)
            ws.Cells(r, x).Font.Color = RGB(190, 0, 0)
            ws.Cells(r, x).Interior.Color = RGB(255, 190, 190)
        End If
    End If
    If hd.Exists("country of nationality") And hd.Exists("government id issuing authority") Then
        x = hd("country of nationality"): y = hd("government id issuing authority")
        v = Txt(ws.Cells(r, x).Value2): w = Txt(ws.Cells(r, y).Value2)
        If Not Missing(v) And Not Missing(w) Then
            If Country(v) <> Country(w) Then
                WriteRow issues, Array(ws.Cells(r, ec).Value2, ws.Cells(r, tp).Value2, "Nationality / Government ID Issuing Authority", v & " / " & w, "Review mismatch; country aliases may require mapping", "Current", r)
                ws.Cells(r, x).Font.Color = RGB(190, 0, 0)
            ws.Cells(r, x).Interior.Color = RGB(255, 190, 190)
                ws.Cells(r, y).Interior.Color = RGB(255, 190, 190)
            End If
        End If
    End If
End Sub
Private Function IsIndividual(ByVal ws As Worksheet, ByVal hd As Object, ByVal r As Long) As Boolean
    If hd.Exists("sub entity type") Then IsIndividual = (Norm(ws.Cells(r, hd("sub entity type")).Value2) = "individual")
End Function
Private Function Country(ByVal s As String) As String
    s = UCase$(Trim$(s))
    Select Case s
        Case "CHINA", "CN", "CHN": s = "CN"
        Case "SINGAPORE", "SG", "SGP": s = "SG"
        Case "INDIA", "IN", "IND": s = "IN"
        Case "VIETNAM", "VIET NAM", "VN", "VNM": s = "VN"
        Case "UNITED STATES", "UNITED STATES OF AMERICA", "US", "USA": s = "US"
    End Select
    Country = s
End Function
Private Function FirstTypeRow(ByVal ws As Worksheet, ByVal tp As Long, ByVal ec As Long, ByVal typ As String) As Long
    Dim r As Long
    For r = 2 To LastRow(ws)
        If Norm(ws.Cells(r, tp).Value2) = Norm(typ) And Not Missing(ws.Cells(r, ec).Value2) Then FirstTypeRow = r: Exit Function
    Next r
End Function
Private Function Missing(ByVal v As Variant) As Boolean
    Select Case UCase$(Trim$(Txt(v)))
        Case "", "--", "-", "N/A", "NULL": Missing = True
    End Select
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
Private Function PairKey(ByVal a As String, ByVal b As String) As String
    PairKey = Len(a) & ":" & a & Len(b) & ":" & b
End Function
Private Function LastRow(ByVal ws As Worksheet) As Long
    Dim c As Range
    Set c = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If c Is Nothing Then LastRow = 1 Else LastRow = c.Row
End Function
Private Function LastCol(ByVal ws As Worksheet) As Long
    Dim c As Range
    Set c = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If c Is Nothing Then LastCol = 1 Else LastCol = c.Column
End Function
Private Sub WriteRow(ByVal ws As Worksheet, ByVal values As Variant)
    Dim r As Long, j As Long
    r = LastRow(ws) + 1
    If Len(Txt(ws.Cells(1, 1).Value2)) = 0 And ws.UsedRange.Count = 1 Then r = 1
    For j = LBound(values) To UBound(values)
        ws.Cells(r, j + 1).NumberFormat = "@"
        ws.Cells(r, j + 1).Value2 = Txt(values(j))
    Next j
End Sub

Private Function EntityScope(ByVal ws As Worksheet, ByVal hd As Object, ByVal r As Long) As String
    If hd.Exists("sub entity type") Then EntityScope = Norm(ws.Cells(r, hd("sub entity type")).Value2)
End Function
Private Function GroupRows(ByVal ws As Worksheet, ByVal ec As Long, ByVal tp As Long) As Object
    Dim d As Object, c As Collection, r As Long, k As String
    Set d = CreateObject("Scripting.Dictionary")
    For r = 2 To LastRow(ws)
        If Not Missing(ws.Cells(r, ec).Value2) And Not Missing(ws.Cells(r, tp).Value2) Then
            k = PairKey(Trim$(Txt(ws.Cells(r, ec).Value2)), Norm(ws.Cells(r, tp).Value2))
            If Not d.Exists(k) Then
                Set c = New Collection: d.Add k, c
            End If
            d(k).Add r
        End If
    Next r
    Set GroupRows = d
End Function
Private Function EqualRows(ByVal a As Worksheet, ByVal b As Worksheet, ByVal ha As Object, ByVal hb As Object, ByVal ra As Long, ByVal rb As Long) As Boolean
    Dim h As Variant
    EqualRows = True
    For Each h In ha.Keys
        If h <> "__match_slot" And hb.Exists(h) Then
            If Txt(a.Cells(ra, ha(h)).Value2) <> Txt(b.Cells(rb, hb(h)).Value2) Then EqualRows = False: Exit Function
        End If
    Next h
End Function
Private Sub ResolveDuplicates(ByVal a As Worksheet, ByVal b As Worksheet, ByVal ha As Object, ByVal hb As Object, ByVal ea As Long, ByVal eb As Long, ByVal ta As Long, ByVal tb As Long, ByVal issues As Worksheet)
    Dim ga As Object, gb As Object, usedA As Object, usedB As Object
    Dim k As Variant, ra As Variant, rb As Variant, ca As Collection, cb As Collection
    Dim colA As Long, colB As Long, slot As Long, na As Long, nb As Long, lastA As Long, lastB As Long
    Dim log As Worksheet
    Set log = a.Parent.Worksheets.Add: log.Name = "Match_Log"
    WriteRow log, Array("ECID", "Party Type", "Previous Row", "Current Row", "Status")
    Set ga = GroupRows(a, ea, ta): Set gb = GroupRows(b, eb, tb)
    colA = LastCol(a) + 1: colB = LastCol(b) + 1
    a.Cells(1, colA).Value2 = "__match_slot": b.Cells(1, colB).Value2 = "__match_slot"
    ha.Add "__match_slot", colA: hb.Add "__match_slot", colB
    For Each k In ga.Keys
        Set ca = ga(k)
        If gb.Exists(k) Then
            Set cb = gb(k)
            Set usedA = CreateObject("Scripting.Dictionary"): Set usedB = CreateObject("Scripting.Dictionary")
            slot = 0
            For Each ra In ca
                For Each rb In cb
                    If Not usedB.Exists(CStr(rb)) Then
                        If EqualRows(a, b, ha, hb, CLng(ra), CLng(rb)) Then
                            slot = slot + 1
                            a.Cells(ra, colA).Value2 = CStr(slot): b.Cells(rb, colB).Value2 = CStr(slot)
                            usedA.Add CStr(ra), True: usedB.Add CStr(rb), True
                            WriteRow log, Array(a.Cells(ra, ea).Value2, a.Cells(ra, ta).Value2, ra, rb, "Identical record paired")
                            Exit For
                        End If
                    End If
                Next rb
            Next ra
            na = 0: nb = 0
            For Each ra In ca
                If Not usedA.Exists(CStr(ra)) Then na = na + 1: lastA = ra
            Next ra
            For Each rb In cb
                If Not usedB.Exists(CStr(rb)) Then nb = nb + 1: lastB = rb
            Next rb
            If na = 1 And nb = 1 Then
                slot = slot + 1
                a.Cells(lastA, colA).Value2 = CStr(slot): b.Cells(lastB, colB).Value2 = CStr(slot)
                WriteRow log, Array(a.Cells(lastA, ea).Value2, a.Cells(lastA, ta).Value2, lastA, lastB, "Unique remaining pair: compare fields")
            ElseIf na > 0 Or nb > 0 Then
                For Each ra In ca
                    If Not usedA.Exists(CStr(ra)) Then a.Cells(ra, colA).Value2 = "UNRESOLVED"
                Next ra
                For Each rb In cb
                    If Not usedB.Exists(CStr(rb)) Then b.Cells(rb, colB).Value2 = "UNRESOLVED"
                Next rb
                If na > 0 And nb > 0 Then
                    WriteRow issues, Array(a.Cells(ca(1), ea).Value2, a.Cells(ca(1), ta).Value2, "Matching", "", "Multiple remaining records; no reliable field pairing", "Both", "")
                End If
                WriteRow log, Array(a.Cells(ca(1), ea).Value2, a.Cells(ca(1), ta).Value2, na, nb, "Unpaired remaining row counts; inspect source rows")
            End If
        Else
            WriteRow log, Array(a.Cells(ca(1), ea).Value2, a.Cells(ca(1), ta).Value2, ca.Count, 0, "ECID + Party Type absent in Current")
        End If
    Next k
    For Each k In gb.Keys
        If Not ga.Exists(k) Then
            Set cb = gb(k)
            WriteRow log, Array(b.Cells(cb(1), eb).Value2, b.Cells(cb(1), tb).Value2, 0, cb.Count, "ECID + Party Type absent in Previous")
        End If
    Next k
End Sub

Private Function NameAt(ByVal ws As Worksheet, ByVal hd As Object, ByVal r As Long) As String
    If r < 2 Then Exit Function
    If hd.Exists("party name") Then NameAt = Txt(ws.Cells(r, hd("party name")).Value2)
End Function
Private Function DisplayValue(ByVal v As Variant) As String
    DisplayValue = Txt(v)
    If Len(DisplayValue) = 0 Then DisplayValue = "[blank]"
End Function
Private Function RowIssues(ByVal issues As Worksheet, ByVal r As Long) As String
    Dim i As Long, result As String
    For i = 2 To LastRow(issues)
        If Txt(issues.Cells(i, 6).Value2) = "Current" Or Txt(issues.Cells(i, 6).Value2) = "Current_Check" Then
            If Txt(issues.Cells(i, 7).Value2) = CStr(r) Then
                If Len(result) > 0 Then result = result & "; "
                result = result & Txt(issues.Cells(i, 3).Value2) & ": " & Txt(issues.Cells(i, 5).Value2)
            End If
        End If
    Next i
    RowIssues = result
End Function
Private Function BuildSummary(ByVal a As Worksheet, ByVal b As Worksheet, ByVal ha As Object, ByVal hb As Object, ByVal ia As Object, ByVal ib As Object, ByVal ea As Long, ByVal eb As Long, ByVal ta As Long, ByVal tb As Long, ByVal detail As Worksheet, ByVal issues As Worksheet) As Worksheet
    Dim ws As Worksheet, r As Long, prev As Long, i As Long, n As Long
    Dim k As String, status As String, problem As String, field As String
    Dim oldValue As String, newValue As String, found As Boolean, changed As Boolean
    Set ws = a.Parent.Worksheets.Add(Before:=a): ws.Name = "Party_Summary"
    WriteRow ws, Array("Party Type", "Party Name (Previous)", "Party Name (Current)", "ECID", "Status", "Changed Field", "Previous Value", "Current Value", "Check Issues", "Previous Row", "Current Row")
    For r = 2 To LastRow(b)
        If Not Missing(b.Cells(r, eb).Value2) Then
            prev = 0: status = "Not matched - review": found = False
            problem = RowIssues(issues, r)
            k = PairKey(PairKey(Trim$(Txt(b.Cells(r, eb).Value2)), Norm(b.Cells(r, tb).Value2)), Txt(b.Cells(r, hb("__match_slot")).Value2))
            If ib.Exists(k) And ia.Exists(k) Then
                If ib(k) > 0 And ia(k) > 0 Then prev = ia(k): status = "Unchanged"
            End If
            For i = 2 To LastRow(detail)
                If Txt(detail.Cells(i, 8).Value2) = CStr(r) And Txt(detail.Cells(i, 6).Value2) = "Changed" Then
                    field = Txt(detail.Cells(i, 3).Value2)
                    oldValue = DisplayValue(detail.Cells(i, 4).Value2): newValue = DisplayValue(detail.Cells(i, 5).Value2)
                    WriteRow ws, Array(b.Cells(r, tb).Value2, NameAt(a, ha, prev), NameAt(b, hb, r), b.Cells(r, eb).Value2, "Changed", field, oldValue, newValue, problem, IIf(prev > 0, CStr(prev), ""), r)
                    n = LastRow(ws)
                    ws.Cells(n, 8).Interior.Color = RGB(255, 240, 150)
                    If Norm(field) = "party name" Then ws.Cells(n, 3).Interior.Color = RGB(255, 240, 150)
                    If Len(problem) > 0 Then ws.Cells(n, 9).Font.Color = RGB(190, 0, 0)
                    found = True
                End If
            Next i
            If Not found Then
                If Len(problem) > 0 Then status = status & " / Check issue"
                WriteRow ws, Array(b.Cells(r, tb).Value2, NameAt(a, ha, prev), NameAt(b, hb, r), b.Cells(r, eb).Value2, status, "", "", "", problem, IIf(prev > 0, CStr(prev), ""), r)
                If Len(problem) > 0 Then ws.Cells(LastRow(ws), 9).Font.Color = RGB(190, 0, 0)
            End If
        End If
    Next r
    ' Retain old-only source records without silently treating unmatched keys as deletions.
    EachKeyOld ws, a, b, ha, hb, ia, ib, ea, eb, ta, tb
    If LastRow(ws) > 1 Then ws.Range("A1:K" & LastRow(ws)).Sort Key1:=ws.Range("A2"), Order1:=xlAscending, Key2:=ws.Range("D2"), Order2:=xlAscending, Header:=xlYes
    ws.Range("A:K").WrapText = True
    Set BuildSummary = ws
End Function
Private Sub EachKeyOld(ByVal ws As Worksheet, ByVal a As Worksheet, ByVal b As Worksheet, ByVal ha As Object, ByVal hb As Object, ByVal ia As Object, ByVal ib As Object, ByVal ea As Long, ByVal eb As Long, ByVal ta As Long, ByVal tb As Long)
    Dim r As Long, k As String
    For r = 2 To LastRow(a)
        If Not Missing(a.Cells(r, ea).Value2) Then
            k = PairKey(PairKey(Trim$(Txt(a.Cells(r, ea).Value2)), Norm(a.Cells(r, ta).Value2)), Txt(a.Cells(r, ha("__match_slot")).Value2))
            If Not ib.Exists(k) Then
                WriteRow ws, Array(a.Cells(r, ta).Value2, NameAt(a, ha, r), "[not matched]", a.Cells(r, ea).Value2, "Not matched in Current - review", "", "", "", "Inspect Match_Log; not a confirmed deletion", r, "")
            End If
        End If
    Next r
End Sub
