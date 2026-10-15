Attribute VB_Name = "modKYCChecklistV2"
Option Explicit

' Local pilot only. Move to an approved team shared location for rollout.
Private Const MASTER_FILE As String = "\\CNACSC.acs.jpmchase.net.cn\lib\APCHNLIBHOME003\F776942\jpmDesk\Desktop\KYC_Checklist_Master.xlsm"

Public Sub InsertKYCChecklistV2()
    Dim xl As Object, src As Object, tmp As Object, sh As Object, dest As Object
    Dim insp As Object, mail As Object, wordDoc As Object
    Dim starts As Collection, ids As Collection, labels As Collection
    Dim wanted As Object, parts As Variant, part As Variant
    Dim menu As String, choice As String, selection As String, itemNo As String
    Dim noCol As Long, docCol As Long, commentsCol As Long, lastCol As Long
    Dim lastRow As Long, r As Long, c As Long, i As Long, startRow As Long, endRow As Long
    Dim nextOutRow As Long, rowsInBlock As Long, n As Long, countSelected As Long
    Dim errText As String, heading As String, cellText As String, lastCell As Object
    Dim oldSec As Long, firstPaste As Boolean
    On Error GoTo Failed
    Set insp = Application.ActiveInspector
    If insp Is Nothing Then Err.Raise vbObjectError + 101, , "Open a draft or reply in a separate Outlook window."
    Set mail = insp.CurrentItem
    If mail.Class <> 43 Then Err.Raise vbObjectError + 102, , "The active window is not an email."
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 3 ' msoAutomationSecurityForceDisable: don't execute workbook macros
    Set src = xl.Workbooks.Open(MASTER_FILE, 0, True)

    menu = "Choose a checklist type (sheet number):" & vbCrLf & vbCrLf
    For i = 1 To src.Worksheets.Count
        menu = menu & i & ". " & src.Worksheets(i).Name & vbCrLf
    Next
    choice = Trim$(InputBox(menu, "KYC Checklist Type"))
    If choice = "" Then GoTo Done
    If Not IsWholePositiveNumber(choice) Then Err.Raise vbObjectError + 103, , "Enter a valid sheet number."
    n = CLng(choice)
    If n > src.Worksheets.Count Then Err.Raise vbObjectError + 104, , "Sheet number not found."
    Set sh = src.Worksheets(n)

    ' Find headings from their labels, NOT from fixed columns.
    noCol = 0: docCol = 0: commentsCol = 0
    For c = 1 To sh.UsedRange.Column + sh.UsedRange.Columns.Count - 1
        heading = LCase$(Trim$(CStr(sh.Cells(1, c).Value)))
        If heading = "no." Or heading = "no" Or heading = "number" Then noCol = c
        If InStr(heading, "document requirement") > 0 Then docCol = c
        If InStr(heading, "comment") > 0 Then commentsCol = c
    Next
    If noCol = 0 Or docCol = 0 Or commentsCol = 0 Then _
        Err.Raise vbObjectError + 105, , "Row 1 must contain No., Document Requirement and Comments headers."
    If Not (noCol < docCol And docCol < commentsCol) Then _
        Err.Raise vbObjectError + 106, , "Header order should be No. / Document Requirement / Comments."
    ' Include the Comments column; merged Requirement cells in between are included automatically.
    lastCol = commentsCol
    Set lastCell = sh.Cells.Find("*", sh.Cells(1, 1), -4163, 2, 1, 2, False)
    If lastCell Is Nothing Then Err.Raise vbObjectError + 107, , "The selected sheet is empty."
    lastRow = lastCell.Row

    Set starts = New Collection: Set ids = New Collection: Set labels = New Collection
    For r = 2 To lastRow
        itemNo = Trim$(CStr(sh.Cells(r, noCol).Value2))
        If IsWholePositiveNumber(itemNo) Then
            ' The number should appear once, at the top left of its merged region.
            starts.Add r
            ids.Add CStr(CLng(itemNo))
            cellText = CStr(sh.Cells(r, docCol).Value)
            cellText = Replace(Replace(cellText, vbCr, " "), vbLf, " ")
            If Len(cellText) > 55 Then cellText = Left$(cellText, 55) & "..."
            labels.Add cellText
        End If
    Next
    If starts.Count = 0 Then Err.Raise vbObjectError + 108, , "No numbered checklist items found in the No. column."

    menu = "Type: " & sh.Name & vbCrLf & "Enter ALL or item numbers (e.g. 1,3,7):" & vbCrLf & vbCrLf
    For i = 1 To starts.Count
        menu = menu & ids(i) & ". " & labels(i) & vbCrLf
    Next
    selection = Trim$(InputBox(menu, "Select Checklist Items", "ALL"))
    If selection = "" Then GoTo Done

    Set wanted = CreateObject("Scripting.Dictionary")
    If UCase$(selection) <> "ALL" Then
        parts = Split(selection, ",")
        For Each part In parts
            itemNo = Trim$(CStr(part))
            If Not IsWholePositiveNumber(itemNo) Then Err.Raise vbObjectError + 109, , "Use ALL or whole numbers separated by commas."
            wanted(CStr(CLng(itemNo))) = True
        Next
    End If

    Set tmp = xl.Workbooks.Add
    Set dest = tmp.Worksheets(1)
    ' Copy full header including horizontally merged cells.
    sh.Range(sh.Cells(1, noCol), sh.Cells(1, lastCol)).Copy dest.Cells(1, 1)
    dest.Rows(1).RowHeight = sh.Rows(1).RowHeight
    nextOutRow = 2
    For i = 1 To starts.Count
        If UCase$(selection) = "ALL" Or wanted.Exists(ids(i)) Then
            startRow = CLng(starts(i))
            If i < starts.Count Then
                endRow = CLng(starts(i + 1)) - 1
            Else
                endRow = lastRow
            End If
            rowsInBlock = endRow - startRow + 1
            ' Copy the ENTIRE numbered item block, preserving merges and formatting.
            sh.Range(sh.Cells(startRow, noCol), sh.Cells(endRow, lastCol)).Copy dest.Cells(nextOutRow, 1)
            For r = 0 To rowsInBlock - 1
                dest.Rows(nextOutRow + r).RowHeight = sh.Rows(startRow + r).RowHeight
            Next
            nextOutRow = nextOutRow + rowsInBlock
            countSelected = countSelected + 1
        End If
    Next
    If countSelected = 0 Then Err.Raise vbObjectError + 110, , "No matching item numbers."

    For c = noCol To lastCol
        dest.Columns(c - noCol + 1).ColumnWidth = sh.Columns(c).ColumnWidth
    Next
    ' Copy as one continuous range, to create a single table in Word.
    dest.Range(dest.Cells(1, 1), dest.Cells(nextOutRow - 1, lastCol - noCol + 1)).Copy
    insp.Activate
    Set wordDoc = insp.WordEditor
    wordDoc.Range(0, 0).PasteExcelTable False, False, False
    wordDoc.Range(0, 0).InsertParagraphBefore
    MsgBox countSelected & " checklist item(s) inserted. Review the layout before sending.", vbInformation
Done:
    On Error Resume Next
    If Not tmp Is Nothing Then tmp.Close False
    If Not src Is Nothing Then src.Close False
    If Not xl Is Nothing Then xl.Quit
    Set tmp = Nothing: Set src = Nothing: Set xl = Nothing
    Exit Sub
Failed:
    errText = Err.Description
    MsgBox "Checklist insertion failed: " & errText, vbExclamation
    Resume Done
End Sub

Private Function IsWholePositiveNumber(ByVal s As String) As Boolean
    Dim i As Long
    s = Trim$(s)
    If Len(s) = 0 Or Len(s) > 8 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next
    IsWholePositiveNumber = (CLng(s) > 0)
End Function
