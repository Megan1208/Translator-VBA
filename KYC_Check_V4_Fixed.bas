Attribute VB_Name = "KYC_Check_V4_Fixed"
Option Explicit

Public Sub RunKYCCheckV4()
    Dim p As Variant, q As Variant, z As Variant, result As Workbook
    Dim a As Worksheet, b As Worksheet, standard As Worksheet, rules As Worksheet
    Dim ha As Object, hb As Object, hs As Object, ea As Long, eb As Long, ta As Long, tb As Long
    Dim sec As Long, events As Boolean, screen As Boolean, msg As String
    sec = Application.AutomationSecurity: events = Application.EnableEvents: screen = Application.ScreenUpdating
    On Error GoTo Failed
    p = Application.GetOpenFilename("Data files (*.csv;*.xlsx;*.xlsm;*.xls),*.csv;*.xlsx;*.xlsm;*.xls", , "1 - Select PREVIOUS")
    If VarType(p) = vbBoolean Then Exit Sub
    q = Application.GetOpenFilename("Data files (*.csv;*.xlsx;*.xlsm;*.xls),*.csv;*.xlsx;*.xlsm;*.xls", , "2 - Select CURRENT")
    If VarType(q) = vbBoolean Then Exit Sub
    If StrComp(CStr(p), CStr(q), vbTextCompare) = 0 Then Err.Raise 5, , "Select different source files."
    Application.AutomationSecurity = 3: Application.EnableEvents = False: Application.ScreenUpdating = False
    Set result = Workbooks.Add(xlWBATWorksheet)
    Set a = result.Worksheets(1): a.Name = "Previous"
    Set b = result.Worksheets.Add(After:=a): b.Name = "Current_Check"
    LoadFile CStr(p), a: LoadFile CStr(q), b
    Set ha = Headers(a): Set hb = Headers(b)
    ea = SelectHeader(a, ha, "ECID", "ECID"): eb = SelectHeader(b, hb, "ECID", "ECID")
    ta = SelectHeader(a, ha, "Party Type", "Party Type"): tb = SelectHeader(b, hb, "Party Type", "Party Type")
    On Error Resume Next
    Set rules = ThisWorkbook.Worksheets("Requirement1")
    On Error GoTo Failed
    If rules Is Nothing Then
        Set standard = a
        If MsgBox("Generate Requirement1 from PREVIOUS? Yes = Previous. No = choose your standard file.", vbYesNo) = vbNo Then
            z = Application.GetOpenFilename("Data files (*.csv;*.xlsx;*.xlsm;*.xls),*.csv;*.xlsx;*.xlsm;*.xls", , "3 - Select STANDARD")
            If VarType(z) = vbBoolean Then Err.Raise 5, , "Standard selection cancelled."
            Set standard = result.Worksheets.Add: standard.Name = "Standard"
            LoadFile CStr(z), standard
        End If
        Set hs = Headers(standard)
        Set rules = ThisWorkbook.Worksheets.Add: rules.Name = "Requirement1"
        GenerateRules standard, hs, SelectHeader(standard, hs, "ECID", "ECID"), SelectHeader(standard, hs, "Party Type", "Party Type"), rules
    End If
    CompareData a, b, ha, hb, ea, eb, ta, tb, rules
    rules.Copy After:=result.Worksheets(result.Worksheets.Count)
    FormatResults result
    result.Worksheets("Party_Summary").Activate
    MsgBox "Finished. See Run_Info for match counts, Party_Summary for names/status, Change_Detail for actual values. Edit Requirement1 in the macro workbook and SAVE it. Save result separately as .xlsx.", vbInformation
Cleanup:
    Application.AutomationSecurity = sec: Application.EnableEvents = events: Application.ScreenUpdating = screen
    Exit Sub
Failed:
    msg = Err.Description
    MsgBox "Stopped: " & msg & vbCrLf & "Do not use partial output as a complete result.", vbExclamation
    Resume Cleanup
End Sub

Private Sub GenerateRules(ByVal source As Worksheet, ByVal hd As Object, ByVal ec As Long, ByVal tp As Long, ByVal rules As Worksheet)
    Dim seen As Object, r As Long, h As Variant, k As String, scope As String
    Set seen = CreateObject("Scripting.Dictionary")
    WriteRow rules, Array("Party Type", "Field", "Mode", "Standard Row", "Sub Entity Type")
    For r = 2 To LastRow(source)
        If Not Missing(source.Cells(r, ec).Value2) And Not Missing(source.Cells(r, tp).Value2) Then
            scope = EntityScope(source, hd, r)
            For Each h In hd.Keys
                If CLng(hd(h)) <> ec And CLng(hd(h)) <> tp Then
                    k = PairKey(PairKey(Norm(source.Cells(r, tp).Value2), scope), CStr(h))
                    If Not Missing(source.Cells(r, hd(h)).Value2) And Not seen.Exists(k) Then
                        seen.Add k, True
                        WriteRow rules, Array(source.Cells(r, tp).Value2, source.Cells(1, hd(h)).Value2, "REQUIRED", r, scope)
                    End If
                End If
            Next h
        End If
    Next r
End Sub

' Outer dictionary: ECID. Inner dictionary: Party Type. Value: ALL source row numbers.
Private Function GroupByECID(ByVal ws As Worksheet, ByVal ec As Long, ByVal tp As Long, ByVal issues As Worksheet) As Object
    Dim people As Object, types As Object, rows As Collection, r As Long, e As String, t As String
    Set people = CreateObject("Scripting.Dictionary")
    For r = 2 To LastRow(ws)
        e = Trim$(Replace(Txt(ws.Cells(r, ec).Value2), ChrW(160), " ")): t = Norm(ws.Cells(r, tp).Value2)
        If Missing(e) Then
            If Application.WorksheetFunction.CountA(ws.Rows(r)) > 0 Then WriteRow issues, Array("", ws.Cells(r, tp).Value2, "ECID", "", "Missing ECID / footer: not compared", ws.Name, r)
        Else
            If Not people.Exists(e) Then
                Set types = CreateObject("Scripting.Dictionary"): people.Add e, types
            End If
            Set types = people(e)
            If Not types.Exists(t) Then
                Set rows = New Collection: types.Add t, rows
            End If
            types(t).Add r
            If Missing(t) Then WriteRow issues, Array(e, "", "Party Type", "", "Missing Party Type", ws.Name, r)
        End If
    Next r
    Set GroupByECID = people
End Function
Private Function UnionKeys(ByVal x As Object, ByVal y As Object) As Object
    Dim d As Object, k As Variant
    Set d = CreateObject("Scripting.Dictionary")
    For Each k In x.Keys: d(k) = True: Next k
    For Each k In y.Keys: d(k) = True: Next k
    Set UnionKeys = d
End Function
Private Function NewSheet(ByVal wb As Workbook, ByVal title As String, ByVal heads As Variant) As Worksheet
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)): ws.Name = title
    WriteRow ws, heads: Set NewSheet = ws
End Function
Private Sub CompareData(ByVal a As Worksheet, ByVal b As Worksheet, ByVal ha As Object, ByVal hb As Object, ByVal ea As Long, ByVal eb As Long, ByVal ta As Long, ByVal tb As Long, ByVal rules As Worksheet)
    Dim summary As Worksheet, detail As Worksheet, issues As Worksheet, log As Worksheet, info As Worksheet
    Dim pa As Object, pb As Object, people As Object, types As Object, xa As Object, xb As Object
    Dim e As Variant, t As Variant, ra As Variant, rb As Variant, rowsA As Collection, rowsB As Collection
    Dim matchedECIDCount As Long, pairs As Long, unresolved As Long, h As Variant
    Set summary = NewSheet(a.Parent, "Party_Summary", Array("Party Type", "Party Name (Previous)", "Party Name (Current)", "ECID", "Status", "Change Summary", "Check Issues", "Previous Row", "Current Row"))
    Set detail = NewSheet(a.Parent, "Change_Detail", Array("ECID", "Party Type", "Field", "Previous", "Current", "Change", "Previous Row", "Current Row"))
    Set issues = NewSheet(a.Parent, "Check_Issues", Array("ECID", "Party Type", "Field", "Value", "Issue", "Source", "Row"))
    Set log = NewSheet(a.Parent, "Match_Log", Array("ECID", "Party Type", "Previous Row", "Current Row", "Match Status"))
    Set info = NewSheet(a.Parent, "Run_Info", Array("Item", "Value"))
    Set pa = GroupByECID(a, ea, ta, issues): Set pb = GroupByECID(b, eb, tb, issues)
    Set people = UnionKeys(pa, pb)
    For Each e In people.Keys
        Set xa = CreateObject("Scripting.Dictionary"): Set xb = CreateObject("Scripting.Dictionary")
        If pa.Exists(e) Then Set xa = pa(e)
        If pb.Exists(e) Then Set xb = pb(e)
        If pa.Exists(e) And pb.Exists(e) Then matchedECIDCount = matchedECIDCount + 1
        Set types = UnionKeys(xa, xb)
        For Each t In types.Keys
            If xa.Exists(t) And xb.Exists(t) Then
                Set rowsA = xa(t): Set rowsB = xb(t)
                PairRole a, b, ha, hb, ea, eb, ta, tb, rowsA, rowsB, summary, detail, log, issues, pairs, unresolved
            ElseIf xb.Exists(t) Then
                For Each rb In xb(t)
                    If pa.Exists(e) Then
                        EmitUnpaired a, b, ha, hb, ea, eb, ta, tb, 0, CLng(rb), "Party Type added", summary, detail, log
                    Else
                        EmitUnpaired a, b, ha, hb, ea, eb, ta, tb, 0, CLng(rb), "ECID only in Current", summary, detail, log
                    End If
                Next rb
            Else
                For Each ra In xa(t)
                    If pb.Exists(e) Then
                        EmitUnpaired a, b, ha, hb, ea, eb, ta, tb, CLng(ra), 0, "Party Type removed", summary, detail, log
                    Else
                        EmitUnpaired a, b, ha, hb, ea, eb, ta, tb, CLng(ra), 0, "ECID only in Previous - verify ID format", summary, detail, log
                    End If
                Next ra
            End If
        Next t
    Next e
    For Each h In ha.Keys
        If Not hb.Exists(h) Then WriteRow detail, Array("", "", a.Cells(1, ha(h)).Value2, "[column exists]", "[column absent]", "Column removed", "", "")
    Next h
    For Each h In hb.Keys
        If Not ha.Exists(h) Then WriteRow detail, Array("", "", b.Cells(1, hb(h)).Value2, "[column absent]", "[column exists]", "Column added", "", "")
    Next h
    CheckRequirements b, hb, eb, tb, rules, issues
    AttachIssues summary, issues
    WriteRow info, Array("Previous ECID count", pa.Count)
    WriteRow info, Array("Current ECID count", pb.Count)
    WriteRow info, Array("matchedECIDCount ECID count", matchedECIDCount)
    WriteRow info, Array("Paired records", pairs)
    WriteRow info, Array("Unresolved records", unresolved)
    WriteRow info, Array("Previous source", a.Name)
    WriteRow info, Array("ECID matching", "Text; surrounding spaces removed. Leading zeros preserved. No name matching.")
    If matchedECIDCount = 0 Then WriteRow info, Array("WARNING", "ZERO matchedECIDCount ECIDs: comparison cannot identify field updates. Check selected ECID columns and original IDs in both sheets.")
    If LastRow(summary) > 1 Then summary.Range("A1:I" & LastRow(summary)).Sort Key1:=summary.Range("A2"), Order1:=xlAscending, Key2:=summary.Range("D2"), Order2:=xlAscending, Header:=xlYes
End Sub
Private Sub PairRole(ByVal a As Worksheet, ByVal b As Worksheet, ByVal ha As Object, ByVal hb As Object, ByVal ea As Long, ByVal eb As Long, ByVal ta As Long, ByVal tb As Long, ByVal ca As Collection, ByVal cb As Collection, ByVal summary As Worksheet, ByVal detail As Worksheet, ByVal log As Worksheet, ByVal issues As Worksheet, ByRef pairs As Long, ByRef unresolved As Long)
    Dim ua As Object, ub As Object, ra As Variant, rb As Variant, na As Long, nb As Long, la As Long, lb As Long
    Set ua = CreateObject("Scripting.Dictionary"): Set ub = CreateObject("Scripting.Dictionary")
    If ca.Count = 1 And cb.Count = 1 Then
        ComparePair a, b, ha, hb, ea, eb, ta, tb, CLng(ca(1)), CLng(cb(1)), summary, detail, log
        pairs = pairs + 1: Exit Sub
    End If
    For Each ra In ca
        For Each rb In cb
            If Not ub.Exists(CStr(rb)) Then
                If RowsEqual(a, b, ha, hb, CLng(ra), CLng(rb)) Then
                    ua(CStr(ra)) = True: ub(CStr(rb)) = True
                    ComparePair a, b, ha, hb, ea, eb, ta, tb, CLng(ra), CLng(rb), summary, detail, log
                    pairs = pairs + 1: Exit For
                End If
            End If
        Next rb
    Next ra
    For Each ra In ca
        If Not ua.Exists(CStr(ra)) Then na = na + 1: la = ra
    Next ra
    For Each rb In cb
        If Not ub.Exists(CStr(rb)) Then nb = nb + 1: lb = rb
    Next rb
    If na = 1 And nb = 1 Then
        ComparePair a, b, ha, hb, ea, eb, ta, tb, la, lb, summary, detail, log
        pairs = pairs + 1
    Else
        For Each ra In ca
            If Not ua.Exists(CStr(ra)) Then
                EmitUnpaired a, b, ha, hb, ea, eb, ta, tb, CLng(ra), 0, "Unresolved duplicate role", summary, detail, log
                unresolved = unresolved + 1
            End If
        Next ra
        For Each rb In cb
            If Not ub.Exists(CStr(rb)) Then
                EmitUnpaired a, b, ha, hb, ea, eb, ta, tb, 0, CLng(rb), "Unresolved duplicate role", summary, detail, log
                WriteRow issues, Array(b.Cells(rb, eb).Value2, b.Cells(rb, tb).Value2, "Matching", "", "Multiple rows in same ECID / role; review candidates", "Current_Check", rb)
                unresolved = unresolved + 1
            End If
        Next rb
    End If
End Sub
Private Function RowsEqual(ByVal a As Worksheet, ByVal b As Worksheet, ByVal ha As Object, ByVal hb As Object, ByVal ra As Long, ByVal rb As Long) As Boolean
    Dim h As Variant
    RowsEqual = True
    For Each h In ha.Keys
        If hb.Exists(h) Then
            If Txt(a.Cells(ra, ha(h)).Value2) <> Txt(b.Cells(rb, hb(h)).Value2) Then RowsEqual = False: Exit Function
        End If
    Next h
End Function
Private Function PartyName(ByVal ws As Worksheet, ByVal hd As Object, ByVal r As Long) As String
    If r = 0 Then PartyName = "[no matched record]": Exit Function
    If hd.Exists("party name") Then PartyName = Txt(ws.Cells(r, hd("party name")).Value2)
End Function
Private Function ShowValue(ByVal v As Variant) As String
    ShowValue = Txt(v)
    If Len(ShowValue) = 0 Then ShowValue = "[blank]"
End Function
Private Sub ComparePair(ByVal a As Worksheet, ByVal b As Worksheet, ByVal ha As Object, ByVal hb As Object, ByVal ea As Long, ByVal eb As Long, ByVal ta As Long, ByVal tb As Long, ByVal ra As Long, ByVal rb As Long, ByVal summary As Worksheet, ByVal detail As Worksheet, ByVal log As Worksheet)
    Dim h As Variant, fields As String, n As Long, status As String, nameChanged As Boolean
    ' Name is compared first; ALL other common fields are still compared afterwards.
    If ha.Exists("party name") And hb.Exists("party name") Then
        nameChanged = Txt(a.Cells(ra, ha("party name")).Value2) <> Txt(b.Cells(rb, hb("party name")).Value2)
        If nameChanged Then
            fields = "Party Name"
            EmitChange a, b, ha, hb, ea, eb, ta, tb, ra, rb, "party name", detail
        End If
    End If
    For Each h In hb.Keys
        If h <> "party name" And CLng(hb(h)) <> eb And CLng(hb(h)) <> tb Then
            If ha.Exists(h) Then
                If Txt(a.Cells(ra, ha(h)).Value2) <> Txt(b.Cells(rb, hb(h)).Value2) Then
                    If Len(fields) > 0 Then fields = fields & "; "
                    fields = fields & Txt(b.Cells(1, hb(h)).Value2)
                    EmitChange a, b, ha, hb, ea, eb, ta, tb, ra, rb, CStr(h), detail
                End If
            End If
        End If
    Next h
    status = "Unchanged": If Len(fields) > 0 Then status = "Changed"
    WriteRow summary, Array(b.Cells(rb, tb).Value2, PartyName(a, ha, ra), PartyName(b, hb, rb), b.Cells(rb, eb).Value2, status, fields, "", ra, rb)
    n = LastRow(summary)
    If Len(fields) > 0 Then summary.Cells(n, 6).Interior.Color = RGB(255, 240, 150)
    If nameChanged Then summary.Cells(n, 3).Interior.Color = RGB(255, 240, 150)
    WriteRow log, Array(b.Cells(rb, eb).Value2, b.Cells(rb, tb).Value2, ra, rb, status)
End Sub
Private Sub EmitChange(ByVal a As Worksheet, ByVal b As Worksheet, ByVal ha As Object, ByVal hb As Object, ByVal ea As Long, ByVal eb As Long, ByVal ta As Long, ByVal tb As Long, ByVal ra As Long, ByVal rb As Long, ByVal h As String, ByVal detail As Worksheet)
    WriteRow detail, Array(b.Cells(rb, eb).Value2, b.Cells(rb, tb).Value2, b.Cells(1, hb(h)).Value2, ShowValue(a.Cells(ra, ha(h)).Value2), ShowValue(b.Cells(rb, hb(h)).Value2), "Changed", ra, rb)
    detail.Cells(LastRow(detail), 5).Interior.Color = RGB(255, 240, 150)
    b.Cells(rb, hb(h)).Interior.Color = RGB(255, 240, 150)
End Sub
Private Sub EmitUnpaired(ByVal a As Worksheet, ByVal b As Worksheet, ByVal ha As Object, ByVal hb As Object, ByVal ea As Long, ByVal eb As Long, ByVal ta As Long, ByVal tb As Long, ByVal ra As Long, ByVal rb As Long, ByVal status As String, ByVal summary As Worksheet, ByVal detail As Worksheet, ByVal log As Worksheet)
    Dim e As String, t As String, h As Variant
    If rb > 0 Then
        e = Txt(b.Cells(rb, eb).Value2): t = Txt(b.Cells(rb, tb).Value2)
    Else
        e = Txt(a.Cells(ra, ea).Value2): t = Txt(a.Cells(ra, ta).Value2)
    End If
    WriteRow summary, Array(t, PartyName(a, ha, ra), PartyName(b, hb, rb), e, status, "See Change_Detail / Match_Log", "", IIf(ra > 0, CStr(ra), ""), IIf(rb > 0, CStr(rb), ""))
    WriteRow log, Array(e, t, IIf(ra > 0, CStr(ra), ""), IIf(rb > 0, CStr(rb), ""), status)
    If ra > 0 Then
        For Each h In ha.Keys
            WriteRow detail, Array(e, t, a.Cells(1, ha(h)).Value2, ShowValue(a.Cells(ra, ha(h)).Value2), "[no matched record]", status, ra, "")
        Next h
    Else
        For Each h In hb.Keys
            WriteRow detail, Array(e, t, b.Cells(1, hb(h)).Value2, "[no matched record]", ShowValue(b.Cells(rb, hb(h)).Value2), status, "", rb)
        Next h
    End If
End Sub
Private Sub CheckRequirements(ByVal ws As Worksheet, ByVal hd As Object, ByVal ec As Long, ByVal tp As Long, ByVal rules As Worksheet, ByVal issues As Worksheet)
    Dim r As Long, i As Long, mode As String, h As String, scope As String, covered As Boolean
    If Norm(rules.Cells(1, 1).Value2) <> "party type" Or Norm(rules.Cells(1, 2).Value2) <> "field" Or Norm(rules.Cells(1, 3).Value2) <> "mode" Then Err.Raise 5, , "Requirement1 needs Party Type / Field / Mode in A / B / C."
    For r = 2 To LastRow(ws)
        If Not Missing(ws.Cells(r, ec).Value2) Then
            covered = False
            For i = 2 To LastRow(rules)
                scope = Norm(rules.Cells(i, 5).Value2)
                If Norm(rules.Cells(i, 1).Value2) = Norm(ws.Cells(r, tp).Value2) And (scope = "" Or scope = EntityScope(ws, hd, r)) Then
                    covered = True: mode = UCase$(Trim$(Txt(rules.Cells(i, 3).Value2))): h = Norm(rules.Cells(i, 2).Value2)
                    If mode = "REQUIRED" Then
                        If Not hd.Exists(h) Then
                            WriteRow issues, Array(ws.Cells(r, ec).Value2, ws.Cells(r, tp).Value2, rules.Cells(i, 2).Value2, "", "Required column missing", "Current_Check", r)
                        ElseIf Missing(ws.Cells(r, hd(h)).Value2) Then
                            WriteRow issues, Array(ws.Cells(r, ec).Value2, ws.Cells(r, tp).Value2, rules.Cells(i, 2).Value2, ShowValue(ws.Cells(r, hd(h)).Value2), "Required information missing", "Current_Check", r)
                            ws.Cells(r, hd(h)).Font.Color = RGB(190, 0, 0)
                            If ws.Cells(r, hd(h)).Interior.Color <> RGB(255, 240, 150) Then ws.Cells(r, hd(h)).Interior.Color = RGB(255, 190, 190)
                        End If
                    ElseIf mode <> "OPTIONAL" Then
                        WriteRow issues, Array(ws.Cells(r, ec).Value2, ws.Cells(r, tp).Value2, rules.Cells(i, 2).Value2, mode, "Invalid rule mode", "Current_Check", r)
                    End If
                End If
            Next i
            If Not covered Then WriteRow issues, Array(ws.Cells(r, ec).Value2, ws.Cells(r, tp).Value2, "Requirements", "", "Type / entity scope has no rules", "Current_Check", r)
            BasicChecks ws, hd, r, ec, tp, issues
        End If
    Next r
End Sub
Private Sub AttachIssues(ByVal summary As Worksheet, ByVal issues As Worksheet)
    Dim r As Long, i As Long, s As String, rowID As String
    For r = 2 To LastRow(summary)
        s = "": rowID = Txt(summary.Cells(r, 9).Value2)
        If Len(rowID) > 0 Then
            For i = 2 To LastRow(issues)
                If (Txt(issues.Cells(i, 6).Value2) = "Current_Check" Or Txt(issues.Cells(i, 6).Value2) = "Current") And Txt(issues.Cells(i, 7).Value2) = rowID Then
                    If Len(s) > 0 Then s = s & "; "
                    s = s & Txt(issues.Cells(i, 3).Value2) & ": " & Txt(issues.Cells(i, 5).Value2)
                End If
            Next i
        End If
        summary.Cells(r, 7).Value2 = s
        If Len(s) > 0 Then summary.Cells(r, 7).Font.Color = RGB(190, 0, 0)
    Next r
End Sub
Private Sub FormatResults(ByVal wb As Workbook)
    Dim ws As Worksheet, j As Long
    For Each ws In wb.Worksheets
        ws.Rows(1).Font.Bold = True: ws.Rows(1).Interior.Color = RGB(210, 225, 240)
        ws.UsedRange.AutoFilter: ws.UsedRange.Columns.AutoFit
        For j = 1 To LastCol(ws)
            If ws.Columns(j).ColumnWidth > 55 Then ws.Columns(j).ColumnWidth = 55
        Next j
        ws.UsedRange.WrapText = True
    Next ws
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
Private Sub BasicChecks(ByVal ws As Worksheet, ByVal hd As Object, ByVal r As Long, ByVal ec As Long, ByVal tp As Long, ByVal issues As Worksheet)
    Dim x As Long, y As Long, v As String, w As String
    If hd.Exists("gender") And IsIndividual(ws, hd, r) Then
        x = hd("gender")
        If Missing(ws.Cells(r, x).Value2) Then
            WriteRow issues, Array(ws.Cells(r, ec).Value2, ws.Cells(r, tp).Value2, "Gender", ws.Cells(r, x).Value2, "Gender missing for Individual", "Current", r)
            ws.Cells(r, x).Font.Color = RGB(190, 0, 0)
            If ws.Cells(r, x).Interior.Color <> RGB(255, 240, 150) Then ws.Cells(r, x).Interior.Color = RGB(255, 190, 190)
        End If
    End If
    If hd.Exists("country of nationality") And hd.Exists("government id issuing authority") Then
        x = hd("country of nationality"): y = hd("government id issuing authority")
        v = Txt(ws.Cells(r, x).Value2): w = Txt(ws.Cells(r, y).Value2)
        If Not Missing(v) And Not Missing(w) Then
            If Country(v) <> Country(w) Then
                WriteRow issues, Array(ws.Cells(r, ec).Value2, ws.Cells(r, tp).Value2, "Nationality / Government ID Issuing Authority", v & " / " & w, "Review mismatch; country aliases may require mapping", "Current", r)
                ws.Cells(r, x).Font.Color = RGB(190, 0, 0)
            If ws.Cells(r, x).Interior.Color <> RGB(255, 240, 150) Then ws.Cells(r, x).Interior.Color = RGB(255, 190, 190)
                ws.Cells(r, y).Font.Color = RGB(190, 0, 0)
                If ws.Cells(r, y).Interior.Color <> RGB(255, 240, 150) Then ws.Cells(r, y).Interior.Color = RGB(255, 190, 190)
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

' Runs the REAL comparison engine against synthetic, reordered source sheets.
Public Sub TestKYCCheckV4()
    Dim wb As Workbook, a As Worksheet, b As Worksheet, rules As Worksheet
    Dim ha As Object, hb As Object, detail As Worksheet, summary As Worksheet
    Dim r As Long, changes As Long, additions As Long, removals As Long, changedRows As Long
    On Error GoTo BadTest
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set a = wb.Worksheets(1): a.Name = "Previous"
    Set b = wb.Worksheets.Add: b.Name = "Current_Check"
    Set rules = wb.Worksheets.Add: rules.Name = "Requirement1"
    WriteRow a, Array("ECID", "Party Type", "Party Name", "Address", "Gender", "Sub Entity Type")
    WriteRow b, Array("Party Name", "Gender", "Party Type", "ECID", "Sub Entity Type", "Address")
    WriteRow a, Array("00123", "Signer", "Old Name", "Shanghai", "Female", "Individual")
    WriteRow a, Array("00123", "Director", "Old Name", "Shanghai", "Female", "Individual")
    WriteRow a, Array("00123", "Manager", "Old Name", "Shanghai", "Female", "Individual")
    WriteRow a, Array("00456", "Signer", "Another Name", "Address X", "Male", "Individual")
    WriteRow a, Array("00456", "Signer", "Another Name", "Address Y", "Male", "Individual")
    WriteRow b, Array("Another Name", "Male", "Signer", "00456", "Individual", "Address Z")
    WriteRow b, Array("New Name", "Female", "Signer", "00123", "Individual", "Beijing")
    WriteRow b, Array("Old Name", "Female", "Owner", "00123", "Individual", "Shanghai")
    WriteRow b, Array("Old Name", "", "Director", "00123", "Individual", "Shanghai")
    WriteRow b, Array("Another Name", "Male", "Signer", "00456", "Individual", "Address X")
    WriteRow rules, Array("Party Type", "Field", "Mode", "Standard Row", "Sub Entity Type")
    WriteRow rules, Array("Signer", "Gender", "REQUIRED", "", "individual")
    WriteRow rules, Array("Director", "Gender", "REQUIRED", "", "individual")
    WriteRow rules, Array("Owner", "Gender", "REQUIRED", "", "individual")
    Set ha = Headers(a): Set hb = Headers(b)
    CompareData a, b, ha, hb, 1, 4, 2, 3, rules
    Set detail = wb.Worksheets("Change_Detail"): Set summary = wb.Worksheets("Party_Summary")
    For r = 2 To LastRow(detail)
        If Txt(detail.Cells(r, 6).Value2) = "Changed" Then
            changes = changes + 1
            If Txt(detail.Cells(r, 4).Value2) = Txt(detail.Cells(r, 5).Value2) Then Err.Raise 5, , "Equal values reported as change"
        End If
    Next r
    For r = 2 To LastRow(summary)
        Select Case Txt(summary.Cells(r, 5).Value2)
            Case "Party Type added": additions = additions + 1
            Case "Party Type removed": removals = removals + 1
            Case "Changed": changedRows = changedRows + 1
        End Select
    Next r
    If changes <> 4 Then Err.Raise 5, , "Expected 4 field changes; got " & changes
    If additions <> 1 Or removals <> 1 Then Err.Raise 5, , "Role add/remove test failed"
    If changedRows <> 3 Then Err.Raise 5, , "Expected 3 changed records"
    If LastRow(summary) <> 7 Then Err.Raise 5, , "Expected 6 summary records"
    If b.Cells(3, 1).Interior.Color <> RGB(255, 240, 150) Then Err.Raise 5, , "Name highlight failed"
    If b.Cells(3, 6).Interior.Color <> RGB(255, 240, 150) Then Err.Raise 5, , "Address highlight failed"
    If b.Cells(5, 2).Interior.Color <> RGB(255, 240, 150) Then Err.Raise 5, , "Cleared field highlight failed"
    If b.Cells(5, 2).Font.Color <> RGB(190, 0, 0) Then Err.Raise 5, , "Missing required field red font failed"
    FormatResults wb
    summary.Activate
    MsgBox "SELF-TEST PASSED: reordered rows/columns, name and address updates, role addition/removal, cleared required field, duplicate-role pairing and highlights.", vbInformation
    Exit Sub
BadTest:
    MsgBox "SELF-TEST FAILED: " & Err.Description, vbExclamation
End Sub
