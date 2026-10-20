Attribute VB_Name = "modBOReportV2"
Option Explicit
'Windows Excel + desktop Word. No references, APIs or ActiveX required.
Private Function U(ByVal h As String) As String
 Dim i As Long, n As Long
 For i = 1 To Len(h) Step 4
  n = CLng("&H" & Mid$(h, i, 4))
  If n > 32767 Then n = n - 65536
  U = U & ChrW(n)
 Next i
End Function
Private Function WS(ByVal nm As String) As Worksheet
 On Error Resume Next
 Set WS = ThisWorkbook.Worksheets(nm)
 On Error GoTo 0
End Function
Private Sub Fail(ByVal t As String)
 Err.Raise vbObjectError + 701, "BOReport", t
End Sub
Private Function S(ByVal r As Long) As String
 S = Trim$(CStr(WS("Settings").Cells(r, 2).Value2))
End Function
Private Function Yes(ByVal r As Long) As Boolean
 Yes = (S(r) = U("662F"))
End Function
Private Sub Choice(ByVal c As Range, ByVal vals As String)
 With c.Validation
  .Delete
  .Add Type:=3, Formula1:=Replace(vals, "|", Application.International(5))
  .IgnoreBlank = False
  .InCellDropdown = True
  .ShowError = True
  .ErrorTitle = U("8BF7900962E9521788684E2D7684503C")
  .ErrorMessage = U("8BF74F7F75284E0B62C9900998793002")
 End With
End Sub
Public Sub SetupReport()
 On Error GoTo EH
 Dim a, i As Long, w As Worksheet, b As Worksheet
 If Not WS("Settings") Is Nothing Or Not WS("BO_Review") Is Nothing Then Fail U("00530065007400740069006E006700730020621600200042004F005F00520065007600690065007700205DF25B5857283002521D59CB53164E0D4F1A898676D65DF267098BBE7F6EFF0C8BF74F7F752873B06709987597623002")
 Set w = ThisWorkbook.Worksheets.Add: w.Name = "Settings"
 a = Array(U("987976EE"), U("004D004900535DE54F5C8868540D79F0"), U("5339914D5217FF0852175B576BCDFF09"), U("5BA262375339914D503CFF084FE175284EE3780162165BA26237540DFF09"), U("5BA26237540D79F0FF085BFC516581EA52A8586B5199FF09"), U("7EDF4E00793E4F1A4FE175284EE37801FF085BFC516581EA52A8586B5199FF09"), U("94F6884C540D79F0"), U("590768484E3B4F537C7B578B"), U("767B8BB0673A5173"), U("6CE8518C8D44672CFF08542B53554F4D5E0179CDFF09"), U("6CD55B9A4EE388684EBA"), U("62107ACB65E5671FFF084E2D658762160059005900590059004D004D00440044FF09"), U("54084F5C4E1A52A17C7B578B"), U("99966B2153D173B05DEE5F0265E5671F"), U("5DEE5F026A215F0F"), U("662F54267A7F900F523081EA71364EBA"), U("662F54265F7154CD98CE96697B497EA7"), U("94885BF960277BA163A763AA65BD"), U("51764ED6886551458BF4660E"), U("5BA2623753CD9988"), _
  U("0042004F004D004900537ED3679C63CF8FF0FF0881EA884C68385B9E540E586B5199FF09"), U("521D6B215EFA7ACB4E1A52A151737CFB"), U("63017EED5C3D804C8C0367E5"), U("6D8953CA53EF75914EA46613"), U("53D173B05951673AFF1A51764ED6"), U("4E0A95E88D708BBF"), U("75358BDD6C9F901A"), U("5FAE4FE1300190AE4EF67B497F517EDC901A8BAF5DE55177"), U("7F517EDC516C5F004FE1606F67E58BE2"), U("4E0E91D1878D673A678481EA8EAB522976CA7D275BC676F8517376844E1A52A1"), U("5BA262378865514563D04F9B8D446599"), U("68385B9E63AA65BDFF1A51764ED6"), U("6C9F901A5BF98C61FF1A8BE5516C53F87ECF529E4EBA"), U("6C9F901A5BF98C61FF1A6BCD516C53F87ECF529E4EBA"), U("96444EF653CA4F9D636EFF1A5BA2623790AE4EF68BB05F55"), U("96444EF653CA4F9D636EFF1A80A167437ED3678456FE"), U("96444EF653CA4F9D636EFF1A516C5F004FE1606F"), _
  U("989D591696444EF6FF084E00884C4E009879FF0C0041006C0074002B0045006E0074006500726362884CFF09"), U("7ED367845206679053CA6BD44F8B7A7F900F8BF4660EFF084EBA5DE5586B5199FF09"), U("673A67848BC6522B8FC77A0BFF084EBA5DE5786E8BA45B8C657463CF8FF0FF09"), U("004D004900535BA26237540D79F05217"), U("004D004900534FE175284EE378015217"), U("004D004900534F1851480042004F59D3540D5217"), U("004D00490053590775280042004F59D3540D5217"), U("004D0049005356FD7C4D5217"), U("004D004900536027522B5217"), U("004D0049005351FA751F65E5671F5217"), U("004D0049005351737CFB7C7B578B5217"), U("004D0049005380A167436BD44F8B5217"), U("004D00490053653676CA002F886851B367436BD44F8B5217"), U("004D004900535F62621065E5671F5217"))
 For i = 0 To UBound(a): w.Cells(i + 1, 1).Value = a(i): Next i
 w.Cells(1, 2).Value = U("586B519900425217FF1B662F002F54264E0B62C94EE3886852FE9009FF1B62406709900962E99ED88BA45F85900962E9")
 w.Cells(2, 2).Value = "Report 1": w.Cells(3, 2).Value = "C"
 w.Cells(7, 2).Value = U("646968395927901A94F6884CFF084E2D56FDFF0967099650516C53F8")
 Choice w.Cells(15, 2), U("5F85900962E9007C5C1A672A8FDB884C59076848007C5B5857284EE54E0B5DEE5F02")
 For i = 16 To 17: w.Cells(i, 2).Value = U("5F85900962E9"): Choice w.Cells(i, 2), U("5F85900962E9007C662F007C5426"): Next i
 For i = 22 To 37
  If i <> 25 And i <> 32 Then
   w.Cells(i, 2).Value = U("5F85900962E9"): Choice w.Cells(i, 2), U("5F85900962E9007C662F007C5426")
  End If
 Next i
 w.Cells(15, 2).Value = U("5F85900962E9")
 a = Array("B", "C", "E", "D", "G", "I", "K", "Q", "U", "V", "Y")
 For i = 0 To UBound(a): w.Cells(i + 41, 2).Value = a(i): Next i
 w.Columns("A").ColumnWidth = 46: w.Columns("B").ColumnWidth = 78
 w.Columns("B").NumberFormat = "@": w.Range("A1:B51").WrapText = True
 w.Range("A1:B1").Font.Bold = True: w.Range("B2:B51").Interior.Color = RGB(255, 250, 210)
 Set b = ThisWorkbook.Worksheets.Add: b.Name = "BO_Review"
 a = Array(U("7EB3516562A5544AFF08662F002F5426FF09"), U("59D3540D"), U("56FD7C4D"), U("6027522B"), U("51FA751F65E5671F"), U("005151737CFB7C7B578B"), U("80A167436BD44F8B0055"), U("653676CA886851B367430056"), U("5F62621065E5671F0059"), U("51774F5351737CFB63CF8FF0FF084EBA5DE5786E8BA4FF09"), U("8BA45B9A680751C6FF084EBA5DE5786E8BA4FF09"), U("5DEE5F0263CF8FF0FF086BCF4EBAFF0C53EF591A9879FF09"), U("67656E90884C"), U("88655145002F5F025E3863D0793A"))
 For i = 0 To UBound(a): b.Cells(1, i + 1).Value = a(i): Next i
 b.Rows(1).Font.Bold = True: b.Columns("A:N").ColumnWidth = 20
 b.Columns("J:L").ColumnWidth = 48: b.Columns("N").ColumnWidth = 40
 b.Cells.NumberFormat = "@": b.Cells.WrapText = True
 b.Range("A1:N1").AutoFilter
 BuildMenu w
 MsgBox U("521D59CB53165B8C6210300251485C06004D004900535B8C6574590D52365230672C5DE54F5C7C3FFF0C4FDD7559539F59CB52174F4D7F6EFF0C518D586B00530065007400740069006E006700735E768FD0884C004C006F006100640042004F3002"), vbInformation
 Exit Sub
EH: MsgBox Err.Description, vbExclamation
End Sub
Private Function V(ByVal w As Worksheet, ByVal r As Long, ByVal settingRow As Long) As String
 Dim x As Variant
 If settingRow = 47 Or settingRow = 51 Then
  If VarType(w.Cells(r, w.Range(S(settingRow) & "1").Column).Value) = vbDate Then
   V = Format$(w.Cells(r, w.Range(S(settingRow) & "1").Column).Value, U("00790079007900795E74006D006D67080064006465E5")): Exit Function
  End If
 End If
 x = w.Cells(r, w.Range(S(settingRow) & "1").Column).Value2
 If IsError(x) Then Fail U("004D00490053670995198BEF53555143683CFF0C884CFF1A") & CStr(r)
 V = Trim$(CStr(x))
End Function
Private Function Pct(ByVal x As String) As String
 If IsNumeric(x) Then
  If CDbl(x) >= 0 And CDbl(x) <= 1 Then
   Pct = Format$(CDbl(x), "0.####%")
  Else
   Pct = x & "%"
  End If
 Else
  Pct = x
 End If
End Function
Private Function DText(ByVal x As String) As String
 Dim z As String, yy As Long, mm As Long, dd As Long, dt As Date
 z = Trim$(x)
 If Len(z) = 8 And IsNumeric(z) Then
  yy = CLng(Left$(z, 4)): mm = CLng(Mid$(z, 5, 2)): dd = CLng(Right$(z, 2))
  On Error GoTo Raw
  dt = DateSerial(yy, mm, dd)
  If Year(dt) = yy And Month(dt) = mm And Day(dt) = dd Then
   DText = CStr(yy) & U("5E74") & Format$(mm, "00") & U("6708") & Format$(dd, "00") & U("65E5"): Exit Function
  End If
 End If
Raw:
 'Keep DD/MM/YY and MM/YY unchanged: do not guess century or missing day.
 DText = z
End Function
Public Sub LoadBO()
 On Error GoTo EH
 Dim w As Worksheet, b As Worksheet, r As Long, last As Long, k As Long, rr As Long
 Dim nm As String, typ As String, key As String, keys As New Collection, found As Boolean, item As Variant
 If WS("Settings") Is Nothing Or WS("BO_Review") Is Nothing Then Fail U("8BF751488FD0884C00530065007400750070005200650070006F007200743002")
 Set w = WS(S(2)): Set b = WS("BO_Review")
 If w Is Nothing Then Fail U("627E4E0D5230004D004900535DE54F5C8868FF0C8BF768C067E500530065007400740069006E006700737B2C0032884C3002")
 If w.Name = "Settings" Or w.Name = "BO_Review" Then Fail U("004D0049005367656E904E0D80FD662F8BBE7F6E6216590D683898753002")
 If S(4) = "" Then Fail U("8BF7586B519900530065007400740069006E006700737B2C0034884C5BA262375339914D503C3002")
 If b.Cells(2, 2).Value <> "" Then Fail U("0042004F005F0052006500760069006500775DF2670951855BB930024E3A4FDD75594EBA5DE54FEE6539FF0C8BF7514853E65B585DE54F5C7C3FFF0C786E8BA4540E624B5DE56E0596647B2C0032884C4EE5540E768465E76570636E518D5BFC51653002")
 last = w.Cells(w.Rows.Count, w.Range(S(3) & "1").Column).End(xlUp).Row
 k = 1
 For r = 2 To last
  If Trim$(CStr(w.Cells(r, w.Range(S(3) & "1").Column).Value2)) = S(4) Then
   If Not found Then
    WS("Settings").Cells(5, 2).Value = V(w, r, 41)
    WS("Settings").Cells(6, 2).Value = V(w, r, 42)
    found = True
   End If
   nm = V(w, r, 43): If nm = "" Then nm = V(w, r, 44)
   typ = V(w, r, 48)
   'Same person + type + ratios + formation date: fold repeated control-method rows.
   key = LCase$(nm & "|" & V(w, r, 47) & "|" & typ & "|" & V(w, r, 49) & "|" & V(w, r, 50) & "|" & V(w, r, 51))
   rr = 0
   For Each item In keys
    If item(0) = key Then rr = CLng(item(1)): Exit For
   Next item
   If rr > 0 Then
    b.Cells(rr, 13).Value = b.Cells(rr, 13).Value & "," & CStr(r)
   Else
    k = k + 1: keys.Add Array(key, k)
    b.Cells(k, 1).Value = U("5F85900962E9"): Choice b.Cells(k, 1), U("5F85900962E9007C662F007C5426")
    b.Cells(k, 2).Value = nm: b.Cells(k, 3).Value = V(w, r, 45)
    b.Cells(k, 4).Value = V(w, r, 46): b.Cells(k, 5).Value = DText(V(w, r, 47))
    b.Cells(k, 6).Value = typ: b.Cells(k, 7).Value = Pct(V(w, r, 49))
    b.Cells(k, 8).Value = Pct(V(w, r, 50)): b.Cells(k, 9).Value = DText(V(w, r, 51))
    b.Cells(k, 13).Value = CStr(r)
    Select Case Left$(typ, 1)
     Case "1"
      b.Cells(k, 10).Value = U("95F463A562E56709") & b.Cells(k, 7).Value & U("80A16743")
      b.Cells(k, 11).Value = U("FF084E00FF09901A8FC776F463A565B95F0F6216800595F463A565B95F0F67007EC862E56709516C53F800320035002562164EE54E0A80A16743300180A14EFD6216800554084F19674376CA")
     Case "2"
      b.Cells(k, 10).Value = U("95F463A562E56709") & b.Cells(k, 8).Value & U("653676CA6743002F886851B36743")
      b.Cells(k, 11).Value = U("FF084E8CFF09867D672A6EE18DB37B2C4E009879680751C6FF0C4F4667007EC84EAB6709516C53F800320035002562164EE54E0A653676CA67433001886851B36743")
     Case "3", "5"
      b.Cells(k, 11).Value = U("FF084E09FF09867D672A6EE18DB37B2C4E009879680751C6FF0C4F46535572EC62168005805454085BF9516C53F88FDB884C5B9E964563A75236")
      b.Cells(k, 14).Value = U("8BF7886551455B9E964563A7523662164FE1625851737CFB51774F5363CF8FF03002")
     Case Else
      b.Cells(k, 14).Value = U("7C7B578B4E3A7A7A6216004F0074006800650072FF1A8BF763095B9E964560C551B58865514551737CFB548C8BA45B9A680751C63002")
    End Select
    If Left$(typ, 1) = "1" Or Left$(typ, 1) = "2" Then b.Cells(k, 14).Value = U("81EA52A863AA8F9E5F85590D6838FF1A786E8BA476F463A5002F95F463A530016BD44F8B300151774F536743522953CA680751C63002")
   End If
  End If
 Next r
 If Not found Then Fail U("6CA167095339914D5BA26237FF0C8BF768C067E55339914D521753CA5B8C65745339914D503C3002")
 MergeRelations b, k
 b.Activate
 MsgBox U("5DF25BFC51655E7654085E76540C540D540C51FA751F65E5671F4EBA54587684591A987951737CFB30028BF7590D6838004A3001004B3001004E5217FF0C900962E97EB35165662F002F5426300254085E76540E768467656E90884C4FDD75595728004D52173002"), vbInformation
 Exit Sub
EH: MsgBox Err.Description, vbExclamation
End Sub
Private Sub AddLine(ByRef t As String, ByVal x As String)
 t = t & x & vbCr
End Sub
Private Function Mark(ByVal r As Long, ByVal caption As String) As String
 If Yes(r) Then Mark = U("2611") & caption Else Mark = U("2610") & caption
End Function
Public Sub GenerateReport()
 On Error GoTo EH
 Dim b As Worksheet, r As Long, j As Long, n As Long, last As Long, t As String, nm As String
 Dim names As New Collection, x As Variant, at As String, obj As String, fn As Variant
 Dim wd As Object, doc As Object, title As String, first As Boolean
 If WS("Settings") Is Nothing Or WS("BO_Review") Is Nothing Then Fail U("8BF75148521D59CB53163002")
 For r = 4 To 14
  If S(r) = "" Then Fail U("00530065007400740069006E006700737B2C") & CStr(r) & U("884C4E0D80FD4E3A7A7A3002")
 Next r
 For r = 15 To 17
  If S(r) = "" Or S(r) = U("5F85900962E9") Then Fail U("8BF75B8C621000530065007400740069006E006700737B2C0031003581F300310037884C900962E93002")
 Next r
 For r = 22 To 37
  If r <> 25 And r <> 32 Then
   If S(r) <> U("662F") And S(r) <> U("5426") Then Fail U("8BF75B8C621000530065007400740069006E006700737B2C") & CStr(r) & U("884C7684662F002F5426900962E93002")
  End If
 Next r
 For Each x In Array(18, 20, 21, 40)
  If S(CLng(x)) = "" Then Fail U("00530065007400740069006E006700737B2C") & CStr(x) & U("884C4E0D80FD4E3A7A7AFF1B4E0D900275288BF7660E786E586B51993002")
 Next x
 If Yes(16) And S(39) = "" Then Fail U("5DF2900962E97A7F900F81EA71364EBAFF0C8BF7586B51997B2C00330039884C7ED36784520667908BF4660E3002")
 If S(15) <> U("5C1A672A8FDB884C59076848") And S(15) <> U("5B5857284EE54E0B5DEE5F02") Then Fail U("5DEE5F026A215F0F65E065483002")
 Set b = WS("BO_Review"): last = b.Cells(b.Rows.Count, 2).End(xlUp).Row
 For r = 2 To last
  If b.Cells(r, 1).Value <> U("662F") And b.Cells(r, 1).Value <> U("5426") Then Fail U("0042004F005F0052006500760069006500777B2C") & CStr(r) & U("884C8BF7900962E97EB35165662F002F54263002")
  If b.Cells(r, 1).Value = U("662F") Then
   For Each x In Array(2, 3, 4, 5, 9, 10, 11)
    If Trim$(CStr(b.Cells(r, CLng(x)).Value)) = "" Then Fail U("0042004F005F0052006500760069006500777B2C") & CStr(r) & U("884C5FC589814FE1606F7F3A59313002")
   Next x
   nm = LCase$(Trim$(CStr(b.Cells(r, 2).Value)))
   For Each x In names
    If x = nm Then Fail U("540C4E0059D3540D88AB7EB35165591A6B21FF0C8BF754085E7651737CFB63CF8FF05E7653EA7EB351654E00884C3002")
   Next x
   names.Add nm: n = n + 1
   If S(15) = U("5B5857284EE54E0B5DEE5F02") And Trim$(CStr(b.Cells(r, 12).Value)) = "" Then Fail U("5B5857285DEE5F026A215F0F4E0BFF0C6BCF540D7EB351654EBA54589700586B5199004C52175DEE5F0263CF8FF03002")
  End If
 Next r
 If n = 0 Then Fail U("81F35C117EB351654E00540D0042004F3002")
 title = S(7) & U("51734E8E") & S(5) & U("53D776CA624067094EBA4FE1606F76845DEE5F0262A5544A")
 AddLine t, U("5DEE5F0262A5544A"): AddLine t, title
 AddLine t, U("4E0030015BA2623757FA672C4FE1606F")
 AddLine t, U("0031002E590768484E3B4F53540D79F0FF1A") & S(5)
 AddLine t, U("0032002E7EDF4E00793E4F1A4FE175284EE37801FF1A") & S(6)
 AddLine t, U("0033002E590768484E3B4F537C7B578BFF1A") & S(8)
 AddLine t, U("0034002E767B8BB0673A5173FF1A") & S(9)
 AddLine t, U("0035002E6CE8518C8D44672CFF0851FA8D44989DFF09FF1A") & S(10)
 AddLine t, U("0036002E6CD55B9A4EE388684EBAFF1A") & S(11)
 AddLine t, U("0037002E62107ACB65E5671FFF1A") & DText(S(12))
 AddLine t, U("0038002E4E0E672C673A678454084F5C4E1A52A17C7B578BFF1A") & S(13)
 AddLine t, U("4E8C30015DEE5F02698251B5")
 AddLine t, U("FF084E00FF0999966B2153D173B05DEE5F0265E5671FFF1A") & DText(S(14))
 AddLine t, U("FF084E8CFF095DEE5F0253D173B05951673A")
 AddLine t, Mark(22, U("4E0E5BA26237521D6B215EFA7ACB4E1A52A151737CFB")) & "  " & Mark(23, U("63017EED5C3D804C8C0367E5")) & "  " & Mark(24, U("5BA262376D8953CA53EF75914EA46613"))
 If S(25) <> "" Then AddLine t, U("51764ED6FF1A") & S(25)
 AddLine t, U("FF084E09FF0953D776CA624067094EBA5DEE5F0251774F5360C551B5")
 If S(15) = U("5C1A672A8FDB884C59076848") Then
  AddLine t, U("7ECF6211884C5BF95BA26237") & S(5) & U("53D776CA624067094EBA8BC6522B68385B9E7ED3679C4E0E0042004F004D0049005368219A8C6BD45BF9FF0C53D173B05BA262375C1A672A8FDB884C590768483002")
 Else
  AddLine t, U("7ECF6211884C5BF95BA26237") & S(5) & U("53D776CA624067094EBA8BC6522B68385B9E7ED3679C4E0E0042004F004D0049005368219A8C6BD45BF9FF0C53D173B05BA262375B5857284EE54E0B5DEE5F02FF1A")
  j = 0
  For r = 2 To last
   If b.Cells(r, 1).Value = U("662F") Then
    j = j + 1
    AddLine t, U("FF08") & CStr(j) & U("FF09") & b.Cells(r, 2).Value & U("FF1A") & Replace(CStr(b.Cells(r, 12).Value), U("53D776CA624067094EBA90576F0F3001591A51FA"), U("672A4F5C4E3A53D776CA624067094EBA59076848"))
   End If
  Next r
 End If
 AddLine t, U("4E0930018BC6522B68385B9E8FC77A0B"): AddLine t, U("FF084E00FF095DF291C753D6768468385B9E63AA65BD")
 For r = 26 To 31: AddLine t, Mark(r, CStr(WS("Settings").Cells(r, 1).Value)): Next r
 If S(32) <> "" Then AddLine t, U("51764ED6FF1A") & S(32)
 AddLine t, U("FF084E8CFF098BC6522B4F9D636E53CA68385FC38BC1660E51855BB9")
 j = 0
 For r = 35 To 37
  If Yes(r) Then
   j = j + 1: at = Replace(CStr(WS("Settings").Cells(r, 1).Value), U("96444EF653CA4F9D636EFF1A"), "")
   AddLine t, U("96444EF6") & CStr(j) & U("FF1A") & at
  End If
 Next r
 If S(38) <> "" Then
  For Each x In Split(Replace(S(38), vbCr, ""), vbLf)
   If Trim$(CStr(x)) <> "" Then
    j = j + 1: AddLine t, U("96444EF6") & CStr(j) & U("FF1A") & CStr(x)
   End If
  Next x
 End If
 AddLine t, U("FF084E09FF0953D776CA624067094EBA8BC6522B68385B9E8FC77A0B")
 If Yes(16) Then
  AddLine t, U("0031002E53D776CA6240670967437ED3678452066790"): AddLine t, S(39)
  AddLine t, U("0032002E8BC6522B68385B9E51774F538FC77A0B")
 End If
 AddLine t, S(40)
 For r = 2 To last
  If b.Cells(r, 1).Value = U("662F") Then
   AddLine t, b.Cells(r, 2).Value & U("FF0C") & b.Cells(r, 4).Value & U("FF0C56FD7C4DFF1A") & b.Cells(r, 3).Value & U("FF0C51FA751F65E5671FFF1A") & b.Cells(r, 5).Value & U("FF0C5BF9") & S(5) & b.Cells(r, 10).Value & U("FF0C7B265408") & b.Cells(r, 11).Value & U("3002517653D776CA6240670967435F62621065E5671F4E3A") & b.Cells(r, 9).Value & U("3002")
  End If
 Next r
 AddLine t, U("6839636E0042004F004D0049005368389A8C7ED3679C53CD9988FF0C") & S(21)
 If Yes(33) Then obj = U("8BE5516C53F87ECF529E4EBA")
 If Yes(34) Then
  If obj <> "" Then obj = obj & U("53CA")
  obj = obj & U("51766BCD516C53F87ECF529E4EBA")
 End If
 If obj <> "" Then AddLine t, U("6211884C4E0E") & obj & U("8FDB884C4E866C9F901A3002")
 AddLine t, U("5BA2623753CD9988FF1A") & S(20)
 AddLine t, U("56DB300198CE966963A7523663AA65BD")
 If Yes(17) Then
  AddLine t, U("5DEE5F02662F54265F7154CD5BA2623798CE96697B497EA7FF1A2611662F002026105426")
 Else
  AddLine t, U("5DEE5F02662F54265F7154CD5BA2623798CE96697B497EA7FF1A2610662F002026115426")
 End If
 AddLine t, U("5BF95BA2623791C753D694885BF960277BA163A763AA65BDFF0859826709FF09FF1A") & S(18)
 AddLine t, U("4E94300151764ED6886551458BF4660E"): AddLine t, S(19)
 fn = Application.GetSaveAsFilename(InitialFileName:="BO_Report_" & Format$(Now, "yyyymmdd_hhnnss") & ".docx", FileFilter:="Word (*.docx),*.docx")
 If VarType(fn) = vbBoolean Then Exit Sub
 If LCase$(Right$(CStr(fn), 5)) <> ".docx" Then fn = CStr(fn) & ".docx"
 If Dir$(CStr(fn)) <> "" Then Fail U("76EE680765874EF65DF25B585728FF0C8BF763624E004E2A65874EF6540D3002")
 Set wd = CreateObject("Word.Application"): wd.Visible = True
 Set doc = wd.Documents.Add
 doc.Content.Text = t
 With doc.Content.Font: .Name = "Calibri": .NameFarEast = U("5B8B4F53"): .Size = 12: End With
 doc.Content.ParagraphFormat.SpaceAfter = 6
 For r = 1 To 2
  doc.Paragraphs(r).Range.Font.Bold = True
  doc.Paragraphs(r).Range.ParagraphFormat.Alignment = 1
 Next r
 doc.SaveAs CStr(fn), 12
 MsgBox U("62A5544A5DF2751F62105E7662535F0030028BF768385BF9540E4F7F752830025F53524D7248672C4E0D81EA52A863D2516580A167437ED3678456FE621696444EF63002"), vbInformation
 Exit Sub
EH:
 MsgBox Err.Description, vbExclamation
 'Leave visible Word document available for recovery if saving failed.
End Sub

Private Sub BuildMenu(ByVal w As Worksheet)
 Dim a, i As Long, sh As Shape
 a = Array("ImportMIS", "LoadBO", "ImportSP", "ApplySPRow", "GenerateReport")
 Dim labels
 labels = Array(U("0031002E900962E9004D0049005365874EF6"), U("0032002E63D053D65E7654085E760042004F"), U("0033002E900962E9005300505BFC51FA65874EF6"), U("0034002E8BFB53D690094E2D768400530050884C"), U("0035002E751F62100057006F0072006462A5544A"))
 w.Columns("D").ColumnWidth = 32
 For i = 0 To UBound(a)
  Set sh = w.Shapes.AddShape(1, w.Cells(2 + i * 3, 4).Left, w.Cells(2 + i * 3, 4).Top, 190, 30)
  sh.TextFrame.Characters.Text = labels(i)
  sh.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & a(i)
 Next i
 w.Cells(19, 4).Value = U("514853E65B584E3A0078006C0073006DFF0C518D4F7F7528630994AE3002")
 w.Cells(20, 4).Value = U("00530050FF1A5BFC5165540E90094E2D76EE68075BA262377684884CFF0C8FD0884C004100700070006C0079005300500052006F00773002")
 w.Cells(21, 4).Value = U("00530050591A90095DEE5F0253EA653E57284E0B65B9988489C8FF0C4E0D81EA52A85206914D7ED96BCF4E2A0042004F3002")
 w.Cells(53, 1).Value = U("0053005067656E905DE54F5C8868")
 w.Cells(54, 1).Value = U("0053005090095B9A67656E90884CFF088BB05F557528FF09")
 w.Cells(55, 1).Value = U("005300505BA262374FE1606F988489C8")
 w.Cells(56, 1).Value = U("0053005068389A8C7ED3679CFF0853EA8BFB988489C8FF09")
 w.Cells(57, 1).Value = U("0053005068389A8C65E5671FFF08786E8BA4540E586B7B2C00310034884CFF09")
 w.Cells(58, 1).Value = U("005300505DEE5F027C7B578BFF08786E8BA45BF95E940042004F540E586B590D6838004C5217FF09")
 w.Cells(59, 1).Value = U("0053005053CD9988FF08786E8BA4540E586B7B2C00320030884CFF09")
 w.Range("A53:B59").WrapText = True
End Sub
Public Sub RefreshButtons()
 On Error GoTo EH
 Dim w As Worksheet, i As Long
 Set w = WS("Settings")
 If w Is Nothing Then Fail U("8BF7514800530065007400750070005200650070006F007200743002")
 For i = w.Shapes.Count To 1 Step -1
  If w.Shapes(i).Type = 1 Then w.Shapes(i).Delete
 Next i
 BuildMenu w
 Exit Sub
EH: MsgBox Err.Description, vbExclamation
End Sub
Public Sub ImportMIS()
 ImportSource False
End Sub
Public Sub ImportSP()
 ImportSource True
End Sub
Private Sub ImportSource(ByVal sp As Boolean)
 On Error GoTo EH
 Dim f As Variant, book As Workbook, src As Worksheet, dest As Worksheet
 Dim nm As String, opts As String, x As Worksheet, target As String, i As Long
 Dim sec As Long, secSet As Boolean, data As Range
 If WS("Settings") Is Nothing Then Fail U("8BF751488FD0884C00530065007400750070005200650070006F007200743002")
 f = Application.GetOpenFilename("Excel (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", , U("900962E967656E900045007800630065006C65874EF6"))
 If VarType(f) = vbBoolean Then Exit Sub
 If StrComp(CStr(f), ThisWorkbook.FullName, vbTextCompare) = 0 Then Fail U("8BF7900962E9591690E867656E9065874EF63002")
 For Each book In Application.Workbooks
  If StrComp(book.FullName, CStr(f), vbTextCompare) = 0 Then Fail U("67656E9065874EF65DF27ECF62535F00FF0C8BF75148517395ED67656E9065874EF6518D5BFC51653002")
 Next book
 Set book = Nothing
 sec = Application.AutomationSecurity
 Application.AutomationSecurity = 3: secSet = True
 Set book = Application.Workbooks.Open(Filename:=CStr(f), UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
 Application.AutomationSecurity = sec: secSet = False
 For Each x In book.Worksheets: opts = opts & x.Name & vbCrLf: Next x
 nm = InputBox(U("8BF78F93516567656E905DE54F5C8868540D79F0FF1A") & vbCrLf & opts, U("900962E95DE54F5C8868"), book.Worksheets(1).Name)
 If nm = "" Then GoTo Clean
 On Error Resume Next
 Set src = book.Worksheets(nm)
 On Error GoTo EH
 If src Is Nothing Then Fail U("627E4E0D52308F93516576845DE54F5C8868540D79F03002")
 If sp Then target = "SP_" Else target = "MIS_"
 target = target & Format$(Now, "mmdd_hhnnss")
 If Not WS(target) Is Nothing Then Fail U("8BF77A0D540E91CD65B05BFC5165FF085DE54F5C8868540D79F091CD590DFF093002")
 Set data = src.UsedRange
 If data.Row <> 1 Or data.Column <> 1 Then Fail U("67656E90886897004ECE004100315F0059CBFF0C7B2C4E00884C662F8868593430028BF7572867656E90526F672C8C036574540E518D5BFC51653002")
 If data.Rows.Count < 2 Then Fail U("67656E9088686CA167096570636E884C3002")
 Set dest = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
 dest.Name = target
 dest.Range("A1").Resize(data.Rows.Count, data.Columns.Count).Value2 = data.Value2
 For i = 1 To data.Columns.Count
  dest.Columns(i).NumberFormat = src.Cells(2, i).NumberFormat
 Next i
 dest.Rows(1).Font.Bold = True: dest.Rows(1).WrapText = True
 dest.Columns.ColumnWidth = 18
 If sp Then
  WS("Settings").Cells(53, 2).Value = target
 Else
  WS("Settings").Cells(2, 2).Value = target
 End If
Clean:
 If Not book Is Nothing Then book.Close SaveChanges:=False
 If Not dest Is Nothing Then
  dest.Activate
  MsgBox U("5BFC51655B8C6210FF0C67656E9065874EF6672A4FEE653930028BF7786E8BA47B2C0031884C4E3A886859343002004D004900535BFC5165540E586B519900530065007400740069006E006700737B2C0034884C518D8FD0884C004C006F006100640042004FFF1B005300508BF790094E2D76EE68075BA26237884C518D8FD0884C004100700070006C0079005300500052006F00773002"), vbInformation
 End If
 Exit Sub
EH:
 Dim msg As String: msg = Err.Description
 On Error Resume Next
 If secSet Then Application.AutomationSecurity = sec
 If Not book Is Nothing Then book.Close SaveChanges:=False
 On Error GoTo 0
 MsgBox msg, vbExclamation
End Sub
Private Function HeaderCol(ByVal w As Worksheet, ByVal header As String) As Long
 Dim c As Long, v As String
 For c = 1 To w.Cells(1, w.Columns.Count).End(xlToLeft).Column
  v = LCase$(Trim$(Replace(Replace(CStr(w.Cells(1, c).Value2), vbCr, " "), vbLf, " ")))
  If v = LCase$(header) Then HeaderCol = c: Exit Function
 Next c
End Function
Public Sub ApplySPRow()
 On Error GoTo EH
 Dim w As Worksheet, r As Long, c As Long, i As Long, a, info As String, picked As Variant
 If WS("Settings") Is Nothing Then Fail U("8BF75148521D59CB53163002")
 Set w = WS(S(53))
 If w Is Nothing Then Fail U("8BF751480049006D0070006F00720074005300503002")
 If Not ActiveSheet Is w Then
  w.Activate
  picked = Application.InputBox(U("8BF78F93516576EE68075BA2623757280053005088684E2D7684884C53F7FF089996884C4E3A88685934FF09300253D66D8853EF514867E5770B6570636E3002"), U("900962E9005300508BB05F55"), Type:=1)
  If VarType(picked) = vbBoolean Then Exit Sub
  If picked <> Fix(picked) Or picked < 2 Or picked > w.Rows.Count Then Fail U("884C53F765E065483002")
  w.Cells(CLng(picked), 1).Select
 End If
 r = ActiveCell.Row
 If r < 2 Or r > w.Cells(w.Rows.Count, 1).End(xlUp).Row Then Fail U("8BF7900962E9670965486570636E884C3002")
 If Selection.Rows.Count <> 1 Then Fail U("4E006B2153EA900962E94E006761005300508BB05F553002")
 For c = 1 To 3
  info = info & CStr(w.Cells(1, c).Value2) & ": " & CStr(w.Cells(r, c).Value2) & vbLf
 Next c
 If MsgBox(U("8BF768385BF98FD9662F5F53524D5BA2623776848BB05F55FF1A") & vbCrLf & info & vbCrLf & U("5F53524D004D004900535BA26237FF1A") & S(5) & vbCrLf & U("5C0653EA66F465B000530050988489C8FF0C4E0D898676D65DF2586B5B576BB53002"), vbYesNo + vbQuestion) <> vbYes Then Exit Sub
 a = Array("BOMIS verification result", "BOMIS verification date", "Type of discrepancy", "Response")
 'Validate all headers before writing anything.
 For i = 0 To UBound(a)
  If HeaderCol(w, CStr(a(i))) = 0 Then Fail U("627E4E0D523088685934FF1A") & CStr(a(i)) & U("30028BF768385BF95BFC51FA88687B2C0031884CFF1B4E0D4F1A731C521753F73002")
 Next i
 WS("Settings").Cells(54, 2).Value = CStr(r)
 WS("Settings").Cells(55, 2).Value = info
 For i = 0 To UBound(a)
  c = HeaderCol(w, CStr(a(i)))
  If VarType(w.Cells(r, c).Value) = vbDate Then
   WS("Settings").Cells(56 + i, 2).Value = Format$(w.Cells(r, c).Value, U("00790079007900795E74006D006D67080064006465E5"))
  Else
   WS("Settings").Cells(56 + i, 2).Value = CStr(w.Cells(r, c).Value2)
  End If
 Next i
 WS("Settings").Activate
 MsgBox U("5DF28BFB53D6523000530065007400740069006E006700737B2C0035003581F300350039884C300268389A8C65E5671F786E8BA4662F99966B2153D173B065E5671F540E586B7B2C00310034884CFF1B53CD9988586B7B2C00320030884CFF1B7ED3679C786E8BA4540E586B7B2C00320031884C548C7B2C00310035884C6A215F0F30025DEE5F027C7B578B63095B9E96455BF95E9451737CFB586B0042004F005F0052006500760069006500777684004C52173002"), vbInformation
 Exit Sub
EH: MsgBox Err.Description, vbExclamation
End Sub
Private Sub AppendUnique(ByVal c As Range, ByVal text As String)
 If Trim$(text) = "" Then Exit Sub
 If Trim$(CStr(c.Value)) = "" Then
  c.Value = text
 ElseIf InStr(1, U("FF1B") & CStr(c.Value) & U("FF1B"), U("FF1B") & text & U("FF1B"), vbTextCompare) = 0 Then
  c.Value = CStr(c.Value) & U("FF1B") & text
 End If
End Sub
Private Sub MergeRelations(ByVal b As Worksheet, ByVal last As Long)
 Dim r As Long, j As Long, c As Long
 For r = last To 3 Step -1
  For j = 2 To r - 1
   If Trim$(CStr(b.Cells(r, 2).Value)) <> "" And _
     StrComp(Trim$(CStr(b.Cells(r, 2).Value)), Trim$(CStr(b.Cells(j, 2).Value)), vbTextCompare) = 0 And _
     Trim$(CStr(b.Cells(r, 5).Value)) = Trim$(CStr(b.Cells(j, 5).Value)) Then
    If b.Cells(r, 3).Value <> b.Cells(j, 3).Value Or b.Cells(r, 4).Value <> b.Cells(j, 4).Value Then
     AppendUnique b.Cells(j, 14), U("67656E9056FD7C4D62166027522B51B27A81FF0C5DF24FDD75594E24884CFF0C8BF7683867E53002")
     Exit For
    End If
    For c = 6 To 14
     AppendUnique b.Cells(j, c), CStr(b.Cells(r, c).Value)
    Next c
    AppendUnique b.Cells(j, 14), U("591A884C51737CFB5DF254085E76FF0C52A15FC5786E8BA463CF8FF03001680751C653CA540451737CFB5F62621065E5671F76845BF95E943002")
    b.Rows(r).Delete
    Exit For
   End If
  Next j
 Next r
End Sub
