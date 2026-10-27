Attribute VB_Name = "modCDDPDFImport"
Option Explicit
'Run SetupCDDImporter once, or run ImportCDDPDF directly.
'Windows desktop Excel with the From PDF connector required.
Private Function T(ByVal h As String) As String
 Dim i As Long, n As Long
 For i = 1 To Len(h) Step 4
  n = CLng("&H" & Mid$(h, i, 4))
  If n > 32767 Then n = n - 65536
  T = T & ChrW(n)
 Next i
End Function
Private Function MString(ByVal value As String) As String
 value = Replace(value, "#", "#(#)")
 MString = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function
Private Function FreeSheetName(ByVal stem As String) As String
 Dim i As Long, nm As String, w As Worksheet, exists As Boolean
 For i = 0 To 999
  nm = Left$(stem, 26)
  If i > 0 Then nm = nm & "_" & CStr(i)
  exists = False
  For Each w In ThisWorkbook.Worksheets
   If StrComp(w.Name, nm, vbTextCompare) = 0 Then exists = True: Exit For
  Next w
  If Not exists Then FreeSheetName = nm: Exit Function
 Next i
 Err.Raise vbObjectError + 710, , T("65E06CD5521B5EFA65B076845DE54F5C8868540D79F03002")
End Function
Public Sub SetupCDDImporter()
 On Error GoTo EH
 Dim w As Worksheet, b As Shape
 If ThisWorkbook.ProtectStructure Then Err.Raise vbObjectError + 711, , T("5DE54F5C7C3F7ED3678453D74FDD62A4FF0C65E06CD5521B5EFA5DE5517798753002")
 Set w = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
 w.Name = FreeSheetName(T("0043004400445DE55177"))
 w.Range("A1").Value = T("0043004400440020005000440046002081EA52A85BFC5165")
 w.Range("A1").Font.Size = 18: w.Range("A1").Font.Bold = True
 w.Range("A3").Value = T("70B951FB630994AE9009005000440046300281EA52A88BFB53D6516890E8005400610062006C0065FF0C4E0D752852FE900962167F168F9167E58BE23002")
 w.Range("A4").Value = T("6BCF6B21521B5EFA65B07684539F59CB6570636E9875FF0C4FDD75595DF2670951855BB93002")
 w.Range("A5").Value = T("672C72485BFC5165539F59CB8868683CFF1B5BA26237548C0042004F5B576BB563D053D65C0657284E0B4E0096366BB55B8C62103002")
 w.Range("A6").Value = T("598253E65B58540E630994AE59316548FF0C0041006C0074002B0046003876F463A58FD0884C0049006D0070006F007200740043004400440050004400463002")
 w.Columns("A").ColumnWidth = 90
 Set b = w.Shapes.AddShape(1, 20, 155, 230, 48)
 b.TextFrame.Characters.Text = T("900962E95E765BFC516500200043004400440020005000440046")
 b.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!ImportCDDPDF"
 w.Activate
 Exit Sub
EH:
 MsgBox Err.Description, vbExclamation, "CDD Import"
End Sub
Public Sub ImportCDDPDF()
 On Error GoTo EH
 Dim picked As Variant, pdf As String, stage As String, token As String
 Dim w As Worksheet, lo As ListObject, qt As QueryTable, q As Object
 Dim queries As Object, formula As String, queryName As String, i As Long
 Dim oldStatus As Variant, oldEvents As Boolean, oldScreen As Boolean
 Dim stateSaved As Boolean, queryMade As Boolean, completed As Boolean
 Dim refreshed As Boolean, errNum As Long, errText As String
 stage = T("900962E9005000440046")
 If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 712, , T("5DE551775DE54F5C7C3F662F53EA8BFB7684FF0C8BF7514853E65B584E3A53EF519976840078006C0073006D3002")
 If ThisWorkbook.ProtectStructure Then Err.Raise vbObjectError + 713, , T("5DE54F5C7C3F7ED3678453D74FDD62A43002")
 picked = Application.GetOpenFilename(FileFilter:="PDF (*.pdf),*.pdf", Title:=T("900962E900200043004400440020005000440046"), MultiSelect:=False)
 If VarType(picked) = vbBoolean Then Exit Sub
 pdf = CStr(picked)
 If LCase$(Right$(pdf, 4)) <> ".pdf" Then Err.Raise vbObjectError + 714, , T("8BF7900962E900500044004665874EF63002")
 stage = T("68C067E50050006F007700650072002000510075006500720079")
 Set queries = CallByName(ThisWorkbook, "Queries", VbGet)
 token = Format$(Now, "yyyymmdd_hhnnss")
 queryName = "CDD_PDF_" & token
 i = 0
 Do
  Set q = Nothing
  On Error Resume Next
  Set q = queries.Item(queryName)
  On Error GoTo EH
  If q Is Nothing Then Exit Do
  i = i + 1: queryName = "CDD_PDF_" & token & "_" & CStr(i)
 Loop
 formula = BuildM(pdf)
 oldStatus = Application.StatusBar: oldEvents = Application.EnableEvents
 oldScreen = Application.ScreenUpdating: stateSaved = True
 Application.EnableEvents = False
 Application.ScreenUpdating = False
 Application.StatusBar = T("6B6357288BFB53D60043004400440020005000440046FF0C8BF77B495F85FF1B5927578B00500044004653EF80FD9700898165705206949F3002")
 stage = T("521B5EFA67E58BE2")
 queries.Add queryName, formula
 queryMade = True
 Set w = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
 w.Name = FreeSheetName(T("004300440044539F59CB6570636E005F") & Format$(Now, "mmdd_hhnnss"))
 w.Range("A1").Value = T("67656E90005000440046FF08539F65874EF6672A4FEE6539FF09")
 w.Range("B1").NumberFormat = "@": w.Range("B1").Value2 = pdf
 w.Range("A2").Value = T("5BFC516565F695F4")
 w.Range("B2").Value = Now: w.Range("B2").NumberFormat = "yyyy-mm-dd hh:mm:ss"
 w.Range("A3").Value = T("4EC554085E768BC6522B51FA7684005400610062006C0065FF1B0050006100670065300156FE7247548C80A1674356FE672A5BFC51653002539F59CB52175C1A672A66205C044E3A4E1A52A15B576BB53002")
 stage = T("8FDE63A50050006F007700650072002000510075006500720079")
 Set lo = w.ListObjects.Add(SourceType:=0, Source:= _
  "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & queryName & ";Extended Properties=""""", _
  Destination:=w.Range("A5"))
 Set qt = lo.QueryTable
 With qt
  .CommandType = 2
  .CommandText = "SELECT * FROM [" & queryName & "]"
  .RefreshStyle = xlInsertDeleteCells
  .AdjustColumnWidth = False
  .PreserveFormatting = True
  .RefreshOnFileOpen = False
  .RefreshPeriod = 0
  .BackgroundQuery = False
 End With
 stage = T("89E367900050004400465E7652A08F7D516890E8005400610062006C0065")
 refreshed = qt.Refresh(False)
 If Not refreshed Then Err.Raise vbObjectError + 715, , T("67E58BE2523765B05DF253D66D886216672A5B8C62103002")
 If lo.ListRows.Count = 0 Then Err.Raise vbObjectError + 716, , T("6CA167098BC6522B523053EF5BFC516576848868683C884C3002")
 stage = T("8BBE7F6E663E793A683C5F0F")
 w.Columns.ColumnWidth = 23
 w.Columns("A").ColumnWidth = 24
 lo.Range.WrapText = True
 lo.HeaderRowRange.Font.Bold = True
 lo.TableStyle = "TableStyleMedium2"
 completed = True
 Application.StatusBar = oldStatus
 Application.EnableEvents = oldEvents
 Application.ScreenUpdating = oldScreen
 stateSaved = False
 w.Activate
 MsgBox T("5BFC51655B8C6210FF0C5171") & CStr(lo.ListRows.Count) & T("884C3002") & vbCrLf & _
  T("67005DE64FA74FDD755967656E9088680049004430018868540D548C88685185884C53F73002") & vbCrLf & _
  T("8BF74FDD5B580078006C0073006D5DE54F5C7C3F30024E0B4E006B65518D505A5BA2623753CA0042004F5B576BB563D053D63002"), vbInformation, "CDD Import"
 Exit Sub
EH:
 errNum = Err.Number: errText = Err.Description
 On Error Resume Next
 'Only remove resources created during this unsuccessful run.
 If Not completed Then
  If Not lo Is Nothing Then lo.Delete
  If queryMade Then
   ThisWorkbook.Connections("Query - " & queryName).Delete
   queries.Item(queryName).Delete
  End If
  If Not w Is Nothing Then
   w.Range("A4").Value = T("5BFC516559318D25FF1A") & stage & " / " & errText
  End If
 End If
 If stateSaved Then
  Application.StatusBar = oldStatus
  Application.EnableEvents = oldEvents
  Application.ScreenUpdating = oldScreen
 End If
 On Error GoTo 0
 MsgBox T("5BFC5165672A5B8C62103002") & vbCrLf & T("6B659AA4FF1A") & stage & vbCrLf & _
  T("95198BEF53F7FF1A") & CStr(errNum) & vbCrLf & errText & vbCrLf & vbCrLf & _
  T("8BF763D04F9B8FD96BB595198BEF4FE1606FFF0C4FBF4E8E5B9A4F4D300265E097004FEE65395B8F5B8951688BBE7F6E3002"), vbExclamation, "CDD Import"
End Sub
Private Function BuildM(ByVal pdf As String) As String
 Dim m As String
 m = "let" & vbCrLf
 m = m & "    Source = Pdf.Tables(File.Contents(" & MString(pdf) & "))," & vbCrLf
 m = m & "    TablesOnly = Table.SelectRows(Source, each [Kind] = ""Table"")," & vbCrLf
 m = m & "    Checked = if Table.RowCount(TablesOnly) = 0 then error ""No tables were detected in this PDF. A scanned PDF may need OCR."" else TablesOnly," & vbCrLf
 m = m & "    WithSource = Table.AddColumn(Checked, ""Extracted"", each" & vbCrLf
 m = m & "        let" & vbCrLf
 m = m & "            SourceId = [Id]," & vbCrLf
 m = m & "            SourceName = [Name]," & vbCrLf
 m = m & "            Indexed = Table.AddIndexColumn([Data], ""CDD_SourceRow"", 1, 1, Int64.Type)," & vbCrLf
 m = m & "            AddId = Table.AddColumn(Indexed, ""CDD_SourceTableId"", each SourceId, type text)," & vbCrLf
 m = m & "            AddName = Table.AddColumn(AddId, ""CDD_SourceTableName"", each SourceName, type text)" & vbCrLf
 m = m & "        in AddName)," & vbCrLf
 m = m & "    Combined = Table.Combine(WithSource[Extracted])," & vbCrLf
 m = m & "    Meta = {""CDD_SourceTableId"", ""CDD_SourceTableName"", ""CDD_SourceRow""}," & vbCrLf
 m = m & "    Ordered = Table.ReorderColumns(Combined, Meta & List.RemoveItems(Table.ColumnNames(Combined), Meta))," & vbCrLf
 m = m & "    Result = if Table.RowCount(Ordered) > 1048571 then error ""Too many rows for one Excel worksheet."" else Ordered" & vbCrLf
 m = m & "in Result"
 BuildM = m
End Function
