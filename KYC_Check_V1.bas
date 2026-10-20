Attribute VB_Name = "KYC_Check_V1"
Option Explicit
' Windows desktop Excel. No additional VBA references required.
' Source files are read only. Header row must be row 1.
Public Sub RunKYCCheck()
    Dim p As Variant, q As Variant, a As Worksheet, b As Worksheet
    Dim out As Workbook, ch As Worksheet, issues As Worksheet, rules As Worksheet
    Dim ha As Object, hb As Object, ia As Object, ib As Object
    Dim ea As Long, eb As Long, ta As Long, tb As Long
    Dim r As Long, j As Long, oldRow As Long, k As Variant, h As Variant, typ As String
    Dim v As String, w As String, ruleIndex As Object, rk As String, n As Long
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
    Set ia = IndexRows(a, ea, ta, issues): Set ib = IndexRows(b, eb, tb, issues)
    ' Rules live in the macro workbook and survive subsequent runs when saved.
    On Error Resume Next
    Set rules = ThisWorkbook.Worksheets("Type_Requirements")
    On Error GoTo Failed
    If rules Is Nothing Then
        Set rules = ThisWorkbook.Worksheets.Add
        rules.Name = "Type_Requirements"
        WriteRow rules, Array("Party Type", "Field", "Mode", "Note")
    End If
    Set ruleIndex = CreateObject("Scripting.Dictionary")
    For r = 2 To LastRow(rules)
        rk = PairKey(Norm(rules.Cells(r, 1).Value2), Norm(rules.Cells(r, 2).Value2))
        If Not ruleIndex.Exists(rk) Then ruleIndex.Add rk, r
    Next r
    ' Generate every type/field combination; never infer mandatory rules from filled cells.
    For r = 2 To LastRow(b)
        If Not Missing(b.Cells(r, eb).Value2) Then
            typ = Txt(b.Cells(r, tb).Value2)
            For Each h In hb.Keys
                j = hb(h)
                If j <> eb And j <> tb Then
                    rk = PairKey(Norm(typ), CStr(h))
                    If Not ruleIndex.Exists(rk) Then
                        WriteRow rules, Array(typ, b.Cells(1, j).Value2, "REVIEW", "Confirm REQUIRED / OPTIONAL / INDIVIDUAL; REVIEW is not a completed check")
                        ruleIndex.Add rk, LastRow(rules)
                    End If
                    n = ruleIndex(rk)
                    v = UCase$(Trim$(Txt(rules.Cells(n, 3).Value2)))
                    If v = "REQUIRED" Or (v = "INDIVIDUAL" And IsIndividual(b, hb, r)) Then
                        If Missing(b.Cells(r, j).Value2) Then
                            WriteRow issues, Array(b.Cells(r, eb).Value2, typ, b.Cells(1, j).Value2, b.Cells(r, j).Value2, "Required information missing", "Current", r)
                            b.Cells(r, j).Interior.Color = RGB(255, 190, 190)
                        End If
                    ElseIf v <> "OPTIONAL" And v <> "INDIVIDUAL" Then
                        ' One issue per rule, not per source record.
                        If r = FirstTypeRow(b, tb, eb, typ) Then WriteRow issues, Array("", typ, b.Cells(1, j).Value2, "", "Requirement not confirmed", "Rules", n)
                    End If
                End If
            Next h
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
                            If v <> w Then
                                WriteRow ch, Array(b.Cells(r, eb).Value2, b.Cells(r, tb).Value2, b.Cells(1, j).Value2, v, w, "Changed", oldRow, r)
                                If b.Cells(r, j).Interior.Color <> RGB(255, 190, 190) Then b.Cells(r, j).Interior.Color = RGB(255, 240, 150)
                            End If
                        End If
                    Next h
                End If
            Else
                WriteRow ch, Array(b.Cells(r, eb).Value2, b.Cells(r, tb).Value2, "[Record]", "", "Present", "Added", "", r)
            End If
        End If
    Next k
    For Each k In ia.Keys
        r = ia(k)
        If r > 0 And Not ib.Exists(k) Then WriteRow ch, Array(a.Cells(r, ea).Value2, a.Cells(r, ta).Value2, "[Record]", "Present", "", "Removed", r, "")
    Next k
    ' Flag configured fields which disappeared from the current export.
    For r = 2 To LastRow(rules)
        If Not hb.Exists(Norm(rules.Cells(r, 2).Value2)) Then
            If UCase$(Trim$(Txt(rules.Cells(r, 3).Value2))) <> "OPTIONAL" Then WriteRow issues, Array("", rules.Cells(r, 1).Value2, rules.Cells(r, 2).Value2, "", "Configured field not in current file", "Rules", r)
        End If
    Next r
    rules.Copy After:=out.Worksheets(out.Worksheets.Count)
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
    out.Activate: ch.Activate
    MsgBox "Complete. Save this result as .xlsx. REVIEW rules are unconfirmed. Edit rules in the macro workbook, save it, and rerun.", vbInformation
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
    On Error GoTo BadLoad
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
    Set sh = src.Worksheets(1)
    If src.Worksheets.Count > 1 Then
        pick = InputBox("Enter the data worksheet name in " & src.Name, "Worksheet", sh.Name)
        If Len(pick) = 0 Then Err.Raise 5, , "Worksheet selection cancelled."
        Set sh = src.Worksheets(pick)
    End If
    target.Cells.NumberFormat = "@"
    target.Range("A1").Resize(LastRow(sh), LastCol(sh)).Value2 = sh.Range("A1").Resize(LastRow(sh), LastCol(sh)).Value2
    src.Close False
    Exit Sub
BadLoad:
    errText = Err.Description
    On Error Resume Next
    If Not src Is Nothing Then src.Close False
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
        If Missing(e) Or Missing(t) Then
            WriteRow issues, Array(e, t, "Key", "", "Missing ECID or Party Type; row not matched (including possible footer)", ws.Name, r)
        Else
            k = PairKey(Trim$(e), Norm(t))
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
            ws.Cells(r, x).Interior.Color = RGB(255, 190, 190)
        End If
    End If
    If hd.Exists("nationality") And hd.Exists("government id issuing authority") Then
        x = hd("nationality"): y = hd("government id issuing authority")
        v = Txt(ws.Cells(r, x).Value2): w = Txt(ws.Cells(r, y).Value2)
        If Not Missing(v) And Not Missing(w) Then
            If Country(v) <> Country(w) Then
                WriteRow issues, Array(ws.Cells(r, ec).Value2, ws.Cells(r, tp).Value2, "Nationality / Government ID Issuing Authority", v & " / " & w, "Review mismatch; country aliases may require mapping", "Current", r)
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
    Norm = LCase$(Trim$(Replace(Replace(Txt(v), ChrW(160), " "), vbLf, " ")))
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
