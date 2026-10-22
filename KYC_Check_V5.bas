Attribute VB_Name = "KYC_Check_V5"
Option Explicit
' Matching order: Party Name -> ECID -> Party Type -> other fields.
Private mOld As Worksheet, mNew As Worksheet
Private mOH As Object, mNH As Object
Private mON As Long, mNN As Long, mOE As Long, mNE As Long, mOT As Long, mNT As Long
Private mSummary As Worksheet, mDetail As Worksheet, mIssues As Worksheet, mMatch As Worksheet
Private mPairs As Long, mAmbiguous As Long, mUpdates As Long

Public Sub RunKYCCheckV5()
    Dim oldPath As Variant, newPath As Variant, outputBook As Workbook
    Dim ruleSheet As Worksheet, infoSheet As Worksheet, savedEvents As Boolean
    Dim savedScreen As Boolean, savedSecurity As Long, errorText As String
    savedEvents = Application.EnableEvents
    savedScreen = Application.ScreenUpdating
    savedSecurity = Application.AutomationSecurity
    On Error GoTo Failed
    oldPath = Application.GetOpenFilename("Data files (*.csv;*.xlsx;*.xlsm;*.xls),*.csv;*.xlsx;*.xlsm;*.xls", , "1 - Select PREVIOUS")
    If VarType(oldPath) = vbBoolean Then Exit Sub
    newPath = Application.GetOpenFilename("Data files (*.csv;*.xlsx;*.xlsm;*.xls),*.csv;*.xlsx;*.xlsm;*.xls", , "2 - Select CURRENT")
    If VarType(newPath) = vbBoolean Then Exit Sub
    If StrComp(CStr(oldPath), CStr(newPath), vbTextCompare) = 0 Then Err.Raise 5, , "Select two different source files."
    Application.EnableEvents = False
    Application.AutomationSecurity = 3
    Set outputBook = Workbooks.Add(xlWBATWorksheet)
    Set mOld = outputBook.Worksheets(1): mOld.Name = "Previous"
    Set mNew = outputBook.Worksheets.Add(After:=mOld): mNew.Name = "Current_Check"
    LoadFile CStr(oldPath), mOld
    LoadFile CStr(newPath), mNew
    BindColumns
    Application.ScreenUpdating = False
    Set ruleSheet = GetRequirementSheet()
    ValidateRules ruleSheet
    CompareAll
    CheckCurrent ruleSheet
    ApplyIssuesToSummary
    ruleSheet.Copy After:=outputBook.Worksheets(outputBook.Worksheets.Count)
    outputBook.Worksheets(outputBook.Worksheets.Count).Name = "Requirement1_Used"
    Set infoSheet = CreateReport(outputBook, "Run_Info", Array("Item", "Value"))
    WriteRow infoSheet, Array("Previous file", CStr(oldPath))
    WriteRow infoSheet, Array("Current file", CStr(newPath))
    WriteRow infoSheet, Array("Matched role records", mPairs)
    WriteRow infoSheet, Array("Reported changes (including name/role/ECID changes)", mUpdates)
    WriteRow infoSheet, Array("Ambiguous source records", mAmbiguous)
    WriteRow infoSheet, Array("Matching order", "Party Name -> ECID -> Party Type -> all common fields")
    WriteRow infoSheet, Array("Requirement source", "Macro workbook: Requirement1. Result copy is Requirement1_Used.")
    WriteRow infoSheet, Array("Country inconsistency check", "Removed per email mapping")
    FormatOutput outputBook
    mSummary.Activate
    Application.ScreenUpdating = True
    MsgBox "Complete. Read Party_Summary and Change_Detail. Save this result as .xlsx. Save your macro workbook to keep Requirement1 edits.", vbInformation
Cleanup:
    Application.EnableEvents = savedEvents
    Application.ScreenUpdating = savedScreen
    Application.AutomationSecurity = savedSecurity
    Exit Sub
Failed:
    errorText = Err.Description
    If Not outputBook Is Nothing Then
        On Error Resume Next
        outputBook.Worksheets(1).Range("A1").AddComment "INCOMPLETE RUN: " & errorText
        On Error GoTo 0
    End If
    MsgBox "Stopped: " & errorText & vbCrLf & "Partial output is INCOMPLETE.", vbExclamation
    GoTo Cleanup
End Sub

Private Sub BindColumns()
    Set mOH = Headers(mOld): Set mNH = Headers(mNew)
    mON = SelectHeader(mOld, mOH, "Party Name", "Party Name")
    mNN = SelectHeader(mNew, mNH, "Party Name", "Party Name")
    mOE = SelectHeader(mOld, mOH, "ECID", "ECID")
    mNE = SelectHeader(mNew, mNH, "ECID", "ECID")
    mOT = SelectHeader(mOld, mOH, "Party Type", "Party Type")
    mNT = SelectHeader(mNew, mNH, "Party Type", "Party Type")
End Sub

Private Function CreateReport(ByVal wb As Workbook, ByVal sheetTitle As String, ByVal headings As Variant) As Worksheet
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheetTitle
    WriteRow ws, headings
    Set CreateReport = ws
End Function

Private Sub CompareAll()
    Dim oldNames As Object, newNames As Object, allNames As Object, nameKey As Variant
    Dim oldRows As Collection, newRows As Collection, item As Variant
    mPairs = 0: mAmbiguous = 0: mUpdates = 0
    Set mSummary = CreateReport(mOld.Parent, "Party_Summary", Array( _
        "Party Type", "Party Name (Previous)", "Party Name (Current)", _
        "ECID (Previous)", "ECID (Current)", "Status", "Changed Fields", _
        "Check Issues", "Previous Row", "Current Row"))
    Set mDetail = CreateReport(mOld.Parent, "Change_Detail", Array( _
        "Party Name (Previous)", "Party Name (Current)", "ECID (Previous)", _
        "ECID (Current)", "Party Type", "Field", "Previous Value", _
        "Current Value", "Change", "Previous Row", "Current Row"))
    Set mIssues = CreateReport(mOld.Parent, "Check_Issues", Array( _
        "Party Name", "ECID", "Party Type", "Field", "Value", "Issue", "Current Row"))
    Set mMatch = CreateReport(mOld.Parent, "Match_Log", Array( _
        "Party Name", "ECID (Previous)", "ECID (Current)", "Party Type", _
        "Previous Row", "Current Row", "Result"))
    Set oldNames = GroupRowsBy(mOld, AllDataRows(mOld), mON)
    Set newNames = GroupRowsBy(mNew, AllDataRows(mNew), mNN)
    Set allNames = KeyUnion(oldNames, newNames)
    For Each nameKey In allNames.Keys
        If Len(CStr(nameKey)) = 0 Or Missing(CStr(nameKey)) Then
            If newNames.Exists(nameKey) Then
                For Each item In newNames(nameKey)
                    AddIssue CLng(item), "Party Name", "Party Name missing; not matched"
                    EmitSingle 0, CLng(item), "Missing Party Name - review"
                Next item
            End If
            If oldNames.Exists(nameKey) Then
                For Each item In oldNames(nameKey)
                    EmitSingle CLng(item), 0, "Missing Party Name - review"
                Next item
            End If
        ElseIf Not oldNames.Exists(nameKey) Then
            For Each item In newNames(nameKey)
                EmitSingle 0, CLng(item), "Party added"
            Next item
        ElseIf Not newNames.Exists(nameKey) Then
            For Each item In oldNames(nameKey)
                EmitSingle CLng(item), 0, "Party removed"
            Next item
        Else
            Set oldRows = oldNames(nameKey): Set newRows = newNames(nameKey)
            CompareName oldRows, newRows
        End If
    Next nameKey
    CompareColumns
End Sub

Private Function AllDataRows(ByVal ws As Worksheet) As Collection
    Dim result As New Collection, r As Long
    For r = 2 To LastRow(ws)
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, 1), ws.Cells(r, LastCol(ws)))) > 0 Then
            ' Ignore known export footer instead of treating it as a person.
            If InStr(1, Norm(ws.Cells(r, 1).Value2), "internal use only", vbTextCompare) = 0 Then result.Add r
        End If
    Next r
    Set AllDataRows = result
End Function

Private Function GroupRowsBy(ByVal ws As Worksheet, ByVal inputRows As Collection, ByVal colIndex As Long) As Object
    Dim groups As Object, rowGroup As Collection, item As Variant, keyText As String
    Set groups = CreateObject("Scripting.Dictionary")
    For Each item In inputRows
        If colIndex = mOE And (ws Is mOld) Then
            keyText = IDKey(ws.Cells(item, colIndex).Value2)
        ElseIf colIndex = mNE And (ws Is mNew) Then
            keyText = IDKey(ws.Cells(item, colIndex).Value2)
        Else
            keyText = Norm(ws.Cells(item, colIndex).Value2)
        End If
        If Not groups.Exists(keyText) Then
            Set rowGroup = New Collection
            groups.Add keyText, rowGroup
        End If
        Set rowGroup = groups(keyText)
        rowGroup.Add CLng(item)
    Next item
    Set GroupRowsBy = groups
End Function

Private Function KeyUnion(ByVal first As Object, ByVal second As Object) As Object
    Dim result As Object, item As Variant
    Set result = CreateObject("Scripting.Dictionary")
    For Each item In first.Keys: result(item) = True: Next item
    For Each item In second.Keys: result(item) = True: Next item
    Set KeyUnion = result
End Function

Private Sub CompareName(ByVal oldRows As Collection, ByVal newRows As Collection)
    Dim oldIDs As Object, newIDs As Object, oldOnly As Object, newOnly As Object
    Dim idItem As Variant, item As Variant, oldKey As Variant, newKey As Variant
    Set oldIDs = GroupRowsBy(mOld, oldRows, mOE)
    Set newIDs = GroupRowsBy(mNew, newRows, mNE)
    Set oldOnly = CreateObject("Scripting.Dictionary")
    Set newOnly = CreateObject("Scripting.Dictionary")
    For Each idItem In oldIDs.Keys
        If newIDs.Exists(idItem) And Not Missing(CStr(idItem)) Then
            CompareRoles oldIDs(idItem), newIDs(idItem)
        Else
            oldOnly.Add idItem, oldIDs(idItem)
        End If
    Next idItem
    For Each idItem In newIDs.Keys
        If Not oldIDs.Exists(idItem) Or Missing(CStr(idItem)) Then newOnly.Add idItem, newIDs(idItem)
    Next idItem
    If oldOnly.Count = 1 And newOnly.Count = 1 Then
        oldKey = oldOnly.Keys: newKey = newOnly.Keys
        If Not Missing(CStr(oldKey(0))) And Not Missing(CStr(newKey(0))) Then
            EmitECIDChange oldOnly(oldKey(0)), newOnly(newKey(0))
            Exit Sub
        End If
    End If
    For Each idItem In oldOnly.Keys
        For Each item In oldOnly(idItem)
            If oldOnly.Count > 0 And newOnly.Count > 0 Then
                EmitSingle CLng(item), 0, "Ambiguous ECID mapping - review"
                mAmbiguous = mAmbiguous + 1
            Else
                EmitSingle CLng(item), 0, "ECID association removed"
            End If
        Next item
    Next idItem
    For Each idItem In newOnly.Keys
        For Each item In newOnly(idItem)
            If oldOnly.Count > 0 And newOnly.Count > 0 Then
                EmitSingle 0, CLng(item), "Ambiguous ECID mapping - review"
                AddIssue CLng(item), "ECID", "Same name has multiple or missing ECIDs; no guessed pairing"
                mAmbiguous = mAmbiguous + 1
            Else
                EmitSingle 0, CLng(item), "ECID association added"
            End If
        Next item
    Next idItem
End Sub

Private Sub EmitECIDChange(ByVal oldRows As Collection, ByVal newRows As Collection)
    Dim oldRow As Long, newRow As Long, item As Variant, summaryRow As Long
    oldRow = oldRows(1): newRow = newRows(1)
    AddDetail oldRow, newRow, "[Name level]", "ECID", _
        mOld.Cells(oldRow, mOE).Value2, mNew.Cells(newRow, mNE).Value2, "ECID changed"
    For Each item In newRows
        AddSummary oldRow, CLng(item), "ECID changed", "ECID; role/other fields not compared at this stage"
        summaryRow = LastRow(mSummary)
        mSummary.Cells(summaryRow, 5).Interior.Color = RGB(255, 240, 150)
        mNew.Cells(item, mNE).Interior.Color = RGB(255, 240, 150)
    Next item
    mUpdates = mUpdates + 1
End Sub

Private Sub CompareRoles(ByVal oldRows As Collection, ByVal newRows As Collection)
    Dim oldRoles As Object, newRoles As Object, allRoles As Object
    Dim roleKey As Variant, item As Variant
    Set oldRoles = GroupRowsBy(mOld, oldRows, mOT)
    Set newRoles = GroupRowsBy(mNew, newRows, mNT)
    Set allRoles = KeyUnion(oldRoles, newRoles)
    For Each roleKey In allRoles.Keys
        If Missing(CStr(roleKey)) Then
            If oldRoles.Exists(roleKey) Then
                For Each item In oldRoles(roleKey): EmitSingle CLng(item), 0, "Missing Party Type - review": Next item
            End If
            If newRoles.Exists(roleKey) Then
                For Each item In newRoles(roleKey)
                    EmitSingle 0, CLng(item), "Missing Party Type - review"
                    AddIssue CLng(item), "Party Type", "Party Type missing; role not compared"
                Next item
            End If
        ElseIf oldRoles.Exists(roleKey) And newRoles.Exists(roleKey) Then
            PairRoleRows oldRoles(roleKey), newRoles(roleKey)
        ElseIf newRoles.Exists(roleKey) Then
            For Each item In newRoles(roleKey): EmitSingle 0, CLng(item), "Party Type added": Next item
        Else
            For Each item In oldRoles(roleKey): EmitSingle CLng(item), 0, "Party Type removed": Next item
        End If
    Next roleKey
End Sub

Private Sub PairRoleRows(ByVal oldRows As Collection, ByVal newRows As Collection)
    Dim usedOld As Object, usedNew As Object, oldItem As Variant, newItem As Variant
    Dim oldCount As Long, newCount As Long, lastOld As Long, lastNew As Long
    If oldRows.Count = 1 And newRows.Count = 1 Then
        ComparePair CLng(oldRows(1)), CLng(newRows(1))
        Exit Sub
    End If
    Set usedOld = CreateObject("Scripting.Dictionary")
    Set usedNew = CreateObject("Scripting.Dictionary")
    For Each oldItem In oldRows
        For Each newItem In newRows
            If Not usedNew.Exists(CStr(newItem)) Then
                If SameFields(CLng(oldItem), CLng(newItem)) Then
                    usedOld(CStr(oldItem)) = True: usedNew(CStr(newItem)) = True
                    ComparePair CLng(oldItem), CLng(newItem)
                    Exit For
                End If
            End If
        Next newItem
    Next oldItem
    For Each oldItem In oldRows
        If Not usedOld.Exists(CStr(oldItem)) Then oldCount = oldCount + 1: lastOld = oldItem
    Next oldItem
    For Each newItem In newRows
        If Not usedNew.Exists(CStr(newItem)) Then newCount = newCount + 1: lastNew = newItem
    Next newItem
    If oldCount = 1 And newCount = 1 Then
        ComparePair lastOld, lastNew
        Exit Sub
    End If
    For Each oldItem In oldRows
        If Not usedOld.Exists(CStr(oldItem)) Then
            EmitSingle CLng(oldItem), 0, "Repeated role row - review"
            mAmbiguous = mAmbiguous + 1
        End If
    Next oldItem
    For Each newItem In newRows
        If Not usedNew.Exists(CStr(newItem)) Then
            EmitSingle 0, CLng(newItem), "Repeated role row - review"
            AddIssue CLng(newItem), "Matching", "More than one remaining row with this name / ECID / role"
            mAmbiguous = mAmbiguous + 1
        End If
    Next newItem
End Sub

Private Function SameFields(ByVal oldRow As Long, ByVal newRow As Long) As Boolean
    Dim fieldKey As Variant
    SameFields = True
    For Each fieldKey In mNH.Keys
        If mOH.Exists(fieldKey) Then
            If Txt(mOld.Cells(oldRow, mOH(fieldKey)).Value2) <> Txt(mNew.Cells(newRow, mNH(fieldKey)).Value2) Then
                SameFields = False
                Exit Function
            End If
        End If
    Next fieldKey
End Function

Private Sub ComparePair(ByVal oldRow As Long, ByVal newRow As Long)
    Dim fieldKey As Variant, changedList As String, oldValue As String, newValue As String
    Dim resultRow As Long
    For Each fieldKey In mNH.Keys
        If mOH.Exists(fieldKey) Then
            oldValue = Txt(mOld.Cells(oldRow, mOH(fieldKey)).Value2)
            newValue = Txt(mNew.Cells(newRow, mNH(fieldKey)).Value2)
            If oldValue <> newValue Then
                If Len(changedList) > 0 Then changedList = changedList & "; "
                changedList = changedList & Txt(mNew.Cells(1, mNH(fieldKey)).Value2)
                AddDetail oldRow, newRow, Txt(mNew.Cells(newRow, mNT).Value2), _
                    Txt(mNew.Cells(1, mNH(fieldKey)).Value2), oldValue, newValue, "Field changed"
                mNew.Cells(newRow, mNH(fieldKey)).Interior.Color = RGB(255, 240, 150)
                mUpdates = mUpdates + 1
            End If
        End If
    Next fieldKey
    If Len(changedList) = 0 Then
        AddSummary oldRow, newRow, "Unchanged", ""
    Else
        AddSummary oldRow, newRow, "Information changed", changedList
        resultRow = LastRow(mSummary)
        mSummary.Cells(resultRow, 7).Interior.Color = RGB(255, 240, 150)
        If Txt(mOld.Cells(oldRow, mON).Value2) <> Txt(mNew.Cells(newRow, mNN).Value2) Then
            mSummary.Cells(resultRow, 3).Interior.Color = RGB(255, 240, 150)
        End If
    End If
    mPairs = mPairs + 1
End Sub

Private Sub EmitSingle(ByVal oldRow As Long, ByVal newRow As Long, ByVal statusText As String)
    Dim fieldKey As Variant, summaryRow As Long
    AddSummary oldRow, newRow, statusText, "See Change_Detail for source information"
    summaryRow = LastRow(mSummary)
    If newRow > 0 Then
        For Each fieldKey In mNH.Keys
            AddDetail 0, newRow, Txt(mNew.Cells(newRow, mNT).Value2), _
                Txt(mNew.Cells(1, mNH(fieldKey)).Value2), "[record absent or unpaired]", _
                mNew.Cells(newRow, mNH(fieldKey)).Value2, statusText
        Next fieldKey
    Else
        For Each fieldKey In mOH.Keys
            AddDetail oldRow, 0, Txt(mOld.Cells(oldRow, mOT).Value2), _
                Txt(mOld.Cells(1, mOH(fieldKey)).Value2), mOld.Cells(oldRow, mOH(fieldKey)).Value2, _
                "[record absent or unpaired]", statusText
        Next fieldKey
    End If
    If statusText = "Party added" Or statusText = "Party Type added" Then
        mSummary.Cells(summaryRow, 1).Interior.Color = RGB(255, 240, 150)
        mSummary.Cells(summaryRow, 3).Interior.Color = RGB(255, 240, 150)
        mNew.Cells(newRow, mNT).Interior.Color = RGB(255, 240, 150)
    End If
    mUpdates = mUpdates + 1
End Sub

Private Sub AddSummary(ByVal oldRow As Long, ByVal newRow As Long, ByVal statusText As String, ByVal changedText As String)
    Dim roleText As String, oldName As String, newName As String, oldID As String, newID As String
    Dim oldRef As String, newRef As String
    oldName = "[record absent or unpaired]": newName = oldName
    oldID = oldName: newID = oldName
    If oldRow > 0 Then
        oldName = Txt(mOld.Cells(oldRow, mON).Value2)
        oldID = Txt(mOld.Cells(oldRow, mOE).Value2)
        roleText = Txt(mOld.Cells(oldRow, mOT).Value2)
        oldRef = CStr(oldRow)
    End If
    If newRow > 0 Then
        newName = Txt(mNew.Cells(newRow, mNN).Value2)
        newID = Txt(mNew.Cells(newRow, mNE).Value2)
        roleText = Txt(mNew.Cells(newRow, mNT).Value2)
        newRef = CStr(newRow)
    End If
    WriteRow mSummary, Array(roleText, oldName, newName, oldID, newID, _
        statusText, changedText, "", oldRef, newRef)
    WriteRow mMatch, Array(IIf(newRow > 0, newName, oldName), oldID, newID, _
        roleText, oldRef, newRef, statusText)
End Sub

Private Sub AddDetail(ByVal oldRow As Long, ByVal newRow As Long, ByVal roleText As String, _
    ByVal fieldText As String, ByVal oldValue As Variant, ByVal newValue As Variant, ByVal statusText As String)
    Dim oldName As String, newName As String, oldID As String, newID As String
    Dim oldRef As String, newRef As String
    If oldRow > 0 Then
        oldName = Txt(mOld.Cells(oldRow, mON).Value2): oldID = Txt(mOld.Cells(oldRow, mOE).Value2)
        oldRef = CStr(oldRow)
    End If
    If newRow > 0 Then
        newName = Txt(mNew.Cells(newRow, mNN).Value2): newID = Txt(mNew.Cells(newRow, mNE).Value2)
        newRef = CStr(newRow)
    End If
    WriteRow mDetail, Array(oldName, newName, oldID, newID, roleText, fieldText, _
        Shown(oldValue), Shown(newValue), statusText, oldRef, newRef)
    If statusText = "Field changed" Or statusText = "ECID changed" Then
        mDetail.Cells(LastRow(mDetail), 8).Interior.Color = RGB(255, 240, 150)
    End If
End Sub

Private Sub CompareColumns()
    Dim fieldKey As Variant
    For Each fieldKey In mOH.Keys
        If Not mNH.Exists(fieldKey) Then
            AddDetail 0, 0, "[File level]", Txt(mOld.Cells(1, mOH(fieldKey)).Value2), _
                "[column exists]", "[column absent]", "Column removed"
        End If
    Next fieldKey
    For Each fieldKey In mNH.Keys
        If Not mOH.Exists(fieldKey) Then
            AddDetail 0, 0, "[File level]", Txt(mNew.Cells(1, mNH(fieldKey)).Value2), _
                "[column absent]", "[column exists]", "Column added"
        End If
    Next fieldKey
End Sub


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

Public Sub SetupRequirement1V5()
    Dim ws As Worksheet
    On Error GoTo Failed
    Set ws = GetRequirementSheet()
    ValidateRules ws
    ws.Activate
    MsgBox "Requirement1 is ready. Existing rules were preserved. Save your .xlsm after editing. Default rules are from the email, not inferred from data.", vbInformation
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub PopulateEmailRules(ByVal ws As Worksheet)
    Dim roles As Variant, fields As Variant, roleItem As Variant, fieldItem As Variant
    WriteRow ws, Array("Party Type", "Field", "Mode", "Source", "Sub Entity Type", "Notes")
    roles = Array("Immediate Shareholder", "Board Member/Director", "Authorized Signer", _
        "China LDD Beneficial Owner", "Legal Representative or Equivalent Party")
    For Each roleItem In roles
        WriteRow ws, Array(CStr(roleItem), "Country Tagging", "REQUIRED", "Email mapping", "", "Y means value required, not literal Y")
    Next roleItem
    roles = Array("Authorized Signer", "China LDD Beneficial Owner", "Legal Representative or Equivalent Party")
    fields = Array("DOB", "Country of Nationality", "Government Issued ID Type", _
        "Government ID Issuing Authority", "ID Number")
    For Each roleItem In roles
        For Each fieldItem In fields
            WriteRow ws, Array(CStr(roleItem), CStr(fieldItem), "REQUIRED", "Email mapping", "", "")
        Next fieldItem
    Next roleItem
    WriteRow ws, Array("China LDD Beneficial Owner", "Gender", "REQUIRED", "Email mapping", "", "")
    WriteRow ws, Array("*", "Gender", "REQUIRED", "Email separate Gender check", "Individual", _
        "Applies to all Individual records. OPTIONAL disables this rule.")
    ws.Rows(1).Font.Bold = True
    ws.Range("A:F").Columns.AutoFit
    ws.Columns(6).ColumnWidth = 55
    ws.Range("A1:F" & LastRow(ws)).AutoFilter
    With ws.Range("C2:C" & LastRow(ws)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="REQUIRED" & Application.International(xlListSeparator) & "OPTIONAL"
        .ErrorTitle = "Rule mode"
        .ErrorMessage = "Use REQUIRED or OPTIONAL."
        .ShowError = True
    End With
End Sub

Private Sub ValidateRules(ByVal ws As Worksheet)
    Dim r As Long, ruleMode As String, scope As String
    If Norm(ws.Cells(1, 1).Value2) <> "party type" Or Norm(ws.Cells(1, 2).Value2) <> "field" Then
        Err.Raise 5, , "Requirement1 columns A/B must be Party Type / Field. Keep Mode in C and Sub Entity Type in E."
    End If
    If Norm(ws.Cells(1, 3).Value2) <> "mode" Then Err.Raise 5, , "Requirement1 column C must be Mode."
    For r = 2 To LastRow(ws)
        If Len(Norm(ws.Cells(r, 1).Value2)) > 0 Or Len(Norm(ws.Cells(r, 2).Value2)) > 0 Then
            If Len(Norm(ws.Cells(r, 1).Value2)) = 0 Or Len(HeaderKey(ws.Cells(r, 2).Value2)) = 0 Then
                Err.Raise 5, , "Incomplete Requirement1 row " & r
            End If
            ruleMode = UCase$(Trim$(Txt(ws.Cells(r, 3).Value2)))
            If ruleMode <> "REQUIRED" And ruleMode <> "OPTIONAL" Then
                Err.Raise 5, , "Requirement1 row " & r & ": Mode must be REQUIRED or OPTIONAL."
            End If
        End If
    Next r
End Sub

Private Sub CheckCurrent(ByVal rules As Worksheet)
    Dim rowItem As Variant, ruleRow As Long, ruleType As String, roleType As String
    Dim scope As String, entity As String, fieldKey As String, modeText As String
    Dim scopeApplies As Boolean, covered As Boolean, r As Long
    For Each rowItem In AllDataRows(mNew)
        r = CLng(rowItem)
        roleType = Norm(mNew.Cells(r, mNT).Value2)
        If Missing(mNew.Cells(r, mNE).Value2) Then AddIssue r, "ECID", "ECID missing"
        entity = ""
        If mNH.Exists("sub entity type") Then entity = Norm(mNew.Cells(r, mNH("sub entity type")).Value2)
        covered = False
        For ruleRow = 2 To LastRow(rules)
            ruleType = Norm(rules.Cells(ruleRow, 1).Value2)
            If Len(ruleType) > 0 Then
                If ruleType = roleType Or ruleType = "*" Then
                    If ruleType = roleType Then covered = True
                    modeText = UCase$(Trim$(Txt(rules.Cells(ruleRow, 3).Value2)))
                    If modeText = "REQUIRED" Then
                        fieldKey = HeaderKey(rules.Cells(ruleRow, 2).Value2)
                        scope = Norm(rules.Cells(ruleRow, 5).Value2)
                        scopeApplies = (scope = "" Or scope = entity)
                        If Len(scope) > 0 And Missing(entity) Then
                            AddIssue r, "Sub Entity Type", "Cannot evaluate scoped requirement: entity type missing"
                        End If
                        If scopeApplies Then
                            If Not mNH.Exists(fieldKey) Then
                                AddIssue r, Txt(rules.Cells(ruleRow, 2).Value2), "Required column absent in Current"
                            ElseIf Missing(mNew.Cells(r, mNH(fieldKey)).Value2) Then
                                AddIssue r, Txt(rules.Cells(ruleRow, 2).Value2), "Required information missing"
                            End If
                        End If
                    End If
                End If
            End If
        Next ruleRow
        If Not covered Then
            AddIssue r, "Requirement1", "No type-specific mapping; only applicable global rules checked"
        End If
    Next rowItem
End Sub

Private Sub AddIssue(ByVal currentRow As Long, ByVal fieldText As String, ByVal reason As String)
    Dim i As Long, fieldKey As String
    fieldKey = HeaderKey(fieldText)
    For i = 2 To LastRow(mIssues)
        If Txt(mIssues.Cells(i, 7).Value2) = CStr(currentRow) Then
            If HeaderKey(mIssues.Cells(i, 4).Value2) = fieldKey And Txt(mIssues.Cells(i, 6).Value2) = reason Then Exit Sub
        End If
    Next i
    Dim currentValue As String
    If mNH.Exists(fieldKey) Then currentValue = Shown(mNew.Cells(currentRow, mNH(fieldKey)).Value2)
    WriteRow mIssues, Array(mNew.Cells(currentRow, mNN).Value2, mNew.Cells(currentRow, mNE).Value2, _
        mNew.Cells(currentRow, mNT).Value2, fieldText, currentValue, reason, currentRow)
    If mNH.Exists(fieldKey) Then
        With mNew.Cells(currentRow, mNH(fieldKey))
            .Font.Color = RGB(190, 0, 0)
            If .Interior.Color <> RGB(255, 240, 150) Then .Interior.Color = RGB(255, 190, 190)
        End With
    End If
End Sub

Private Sub ApplyIssuesToSummary()
    Dim r As Long, i As Long, resultText As String, rowRef As String
    For r = 2 To LastRow(mSummary)
        rowRef = Txt(mSummary.Cells(r, 10).Value2): resultText = ""
        If Len(rowRef) > 0 Then
            For i = 2 To LastRow(mIssues)
                If Txt(mIssues.Cells(i, 7).Value2) = rowRef Then
                    If Len(resultText) > 0 Then resultText = resultText & "; "
                    resultText = resultText & Txt(mIssues.Cells(i, 4).Value2) & ": " & Txt(mIssues.Cells(i, 6).Value2)
                End If
            Next i
        End If
        mSummary.Cells(r, 8).NumberFormat = "@"
        mSummary.Cells(r, 8).Value2 = resultText
        If Len(resultText) > 0 Then mSummary.Cells(r, 8).Font.Color = RGB(190, 0, 0)
    Next r
End Sub

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

Private Function IDKey(ByVal value As Variant) As String
    IDKey = Trim$(Replace(Txt(value), ChrW(160), " "))
End Function

Private Function Shown(ByVal value As Variant) As String
    Shown = Txt(value)
    If Len(Shown) = 0 Then Shown = "[blank]"
End Function

Private Sub FormatOutput(ByVal wb As Workbook)
    Dim ws As Worksheet, colItem As Long
    For Each ws In wb.Worksheets
        With ws.UsedRange
            .Font.Name = "Calibri"
            .Font.Size = 11
            .VerticalAlignment = xlTop
            .Columns.AutoFit
            .WrapText = True
        End With
        For colItem = 1 To LastCol(ws)
            If ws.Columns(colItem).ColumnWidth > 48 Then ws.Columns(colItem).ColumnWidth = 48
        Next colItem
        ws.Rows(1).Font.Bold = True
        ws.Rows(1).Interior.Color = RGB(210, 225, 240)
        If Not ws.AutoFilterMode Then ws.UsedRange.AutoFilter
        ws.UsedRange.Rows.AutoFit
    Next ws
    If LastRow(mSummary) > 1 Then
        mSummary.Range("A1:J" & LastRow(mSummary)).Sort _
            Key1:=mSummary.Range("A2"), Order1:=xlAscending, _
            Key2:=mSummary.Range("C2"), Order2:=xlAscending, Header:=xlYes
    End If
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
    target.Range("A1").Resize(LastRow(sh), LastCol(sh)).NumberFormat = "@"
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
        s = HeaderKey(ws.Cells(1, j).Value2)
        If Len(s) = 0 Then Err.Raise 5, , "Blank header in " & ws.Name & " column " & j
        If d.Exists(s) Then Err.Raise 5, , "Duplicate header: " & s
        d.Add s, j
    Next j
    Set Headers = d
End Function
Private Function SelectHeader(ByVal ws As Worksheet, ByVal d As Object, ByVal label As String, ByVal suggested As String) As Long
    Dim s As String
    If d.Exists(HeaderKey(suggested)) Then
        SelectHeader = d(HeaderKey(suggested)): Exit Function
    End If
    s = InputBox("Enter exact " & label & " header in " & ws.Name & ". Do not choose Party Role in Reference Data for Party Type.", "Identify column", suggested)
    If Not d.Exists(HeaderKey(s)) Then Err.Raise 5, , "Header not found: " & s
    SelectHeader = d(HeaderKey(s))
End Function
Private Function Missing(ByVal v As Variant) As Boolean
    Select Case Norm(v)
        Case "", "--", "-", "n/a", "null": Missing = True
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
    Set c = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False, SearchFormat:=False, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If c Is Nothing Then LastRow = 1 Else LastRow = c.Row
End Function
Private Function LastCol(ByVal ws As Worksheet) As Long
    Dim c As Range
    Set c = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False, SearchFormat:=False, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
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



' Executes the actual engine on a synthetic workbook; does not alter your Requirement1.
Public Sub TestKYCCheckV5()
    Dim wb As Workbook, rules As Worksheet, r As Long, fieldChanges As Long
    Dim idChanges As Long, addedRoles As Long, removedRoles As Long
    Dim addedPeople As Long, removedPeople As Long, issueFound As Boolean
    On Error GoTo Failed
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set mOld = wb.Worksheets(1): mOld.Name = "Previous"
    Set mNew = wb.Worksheets.Add(After:=mOld): mNew.Name = "Current_Check"
    Set rules = wb.Worksheets.Add(After:=mNew): rules.Name = "Requirement1"
    WriteRow mOld, Array("Party Name", "ECID", "Party Type", "Address", "DOB", "Gender", _
        "Country of Nationality", "Government Issued ID Type", "Government ID Issuing Authority", _
        "ID Number", "Country Tagging", "Sub Entity Type")
    WriteRow mNew, Array("ECID", "Party Type", "Party Name", "Gender", "Address", "DOB", _
        "Country Tagging", "Country of Nationality", "ID Number", "Government ID Issuing Authority", _
        "Government Issued ID Type", "Sub Entity Type")
    FixtureOld "Alpha", "100", "Authorized Signer", "A", "Female"
    FixtureOld "Alpha", "100", "Board Member/Director", "A", "Female"
    FixtureOld "Alpha", "100", "Senior Manager or Equivalent", "A", "Female"
    FixtureOld "Switch", "001", "Authorized Signer", "A", "Male"
    FixtureOld "Gone", "300", "Parent Entity", "A", "Male"
    FixtureOld "Repeat", "400", "Other Related Parties", "X", "Male"
    FixtureOld "Repeat", "400", "Other Related Parties", "Y", "Male"
    FixtureOld "Ambig", "501", "Authorized Signer", "A", "Male"
    FixtureOld "Ambig", "502", "Authorized Signer", "A", "Male"
    ' Rows and columns deliberately reordered.
    FixtureNew "Repeat", "400", "Other Related Parties", "Z", "Male"
    FixtureNew "Alpha", "100", "Authorized Signer", "B", ""
    FixtureNew "Switch", "002", "Authorized Signer", "C", "Male"
    FixtureNew "Alpha", "100", "Senior Manager or Equivalent", "A", "Female"
    FixtureNew "New Person", "600", "Authorized Signer", "A", "Male"
    FixtureNew "Repeat", "400", "Other Related Parties", "X", "Male"
    FixtureNew "Alpha", "100", "China LDD Beneficial Owner", "A", "Female"
    FixtureNew "Ambig", "601", "Authorized Signer", "A", "Male"
    FixtureNew "Ambig", "602", "Authorized Signer", "A", "Male"
    PopulateEmailRules rules
    ValidateRules rules
    BindColumns
    CompareAll
    CheckCurrent rules
    ApplyIssuesToSummary
    For r = 2 To LastRow(mDetail)
        If Txt(mDetail.Cells(r, 9).Value2) = "Field changed" Then fieldChanges = fieldChanges + 1
        If Txt(mDetail.Cells(r, 9).Value2) = "ECID changed" Then
            idChanges = idChanges + 1
            AssertV5 Txt(mDetail.Cells(r, 7).Value2) = "001", "Original ECID leading zero lost"
            AssertV5 Txt(mDetail.Cells(r, 8).Value2) = "002", "New ECID value incorrect"
        End If
    Next r
    For r = 2 To LastRow(mSummary)
        Select Case Txt(mSummary.Cells(r, 6).Value2)
            Case "Party added": addedPeople = addedPeople + 1
            Case "Party removed": removedPeople = removedPeople + 1
            Case "Party Type added": addedRoles = addedRoles + 1
            Case "Party Type removed": removedRoles = removedRoles + 1
        End Select
    Next r
    AssertV5 fieldChanges = 3, "Expected 3 field changes; got " & fieldChanges
    AssertV5 idChanges = 1, "Same-name ECID replacement not reported once"
    AssertV5 addedPeople = 1 And removedPeople = 1, "Name additions/removals failed"
    AssertV5 addedRoles = 1 And removedRoles = 1, "Role additions/removals failed"
    AssertV5 mPairs = 4, "Expected 4 paired role rows"
    AssertV5 mAmbiguous = 4, "Ambiguous same-name ECIDs were not retained for review"
    AssertV5 LastRow(mSummary) = 14, "Expected 13 summary rows"
    AssertV5 mNew.Cells(3, 5).Interior.Color = RGB(255, 240, 150), "Updated address not yellow"
    AssertV5 mNew.Cells(3, 4).Interior.Color = RGB(255, 240, 150), "Cleared Gender not yellow"
    AssertV5 mNew.Cells(3, 4).Font.Color = RGB(190, 0, 0), "Missing Gender not red text"
    For r = 2 To LastRow(mIssues)
        If Txt(mIssues.Cells(r, 7).Value2) = "3" And HeaderKey(mIssues.Cells(r, 4).Value2) = "gender" Then issueFound = True
    Next r
    AssertV5 issueFound, "Gender issue not generated"
    ' A manual OPTIONAL rule must suppress a missing DOB warning.
    Dim ruleIndex As Long
    For ruleIndex = 2 To LastRow(rules)
        If Norm(rules.Cells(ruleIndex, 1).Value2) = "authorized signer" Then
            If HeaderKey(rules.Cells(ruleIndex, 2).Value2) = "dob" Then rules.Cells(ruleIndex, 3).Value2 = "OPTIONAL"
        End If
    Next ruleIndex
    mNew.Cells(3, mNH("dob")).Value2 = ""
    ' A nationality / issuing authority difference is deliberately present, and is allowed.
    mNew.Cells(6, mNH("government id issuing authority")).Value2 = "US"
    mIssues.Range("A2:G" & LastRow(mIssues)).ClearContents
    CheckCurrent rules
    For ruleIndex = 2 To LastRow(mIssues)
        If Txt(mIssues.Cells(ruleIndex, 7).Value2) = "3" Then
            AssertV5 HeaderKey(mIssues.Cells(ruleIndex, 4).Value2) <> "dob", "OPTIONAL DOB rule was ignored"
        End If
        If Txt(mIssues.Cells(ruleIndex, 7).Value2) = "6" Then
            AssertV5 InStr(1, Txt(mIssues.Cells(ruleIndex, 6).Value2), "mismatch", vbTextCompare) = 0, "Removed country comparison is still active"
        End If
    Next ruleIndex
    ' Restore the synthetic data so comparison output still describes the displayed source.
    mNew.Cells(3, mNH("dob")).Value2 = "1990-01-01"
    mNew.Cells(6, mNH("government id issuing authority")).Value2 = "CN"
    mIssues.Range("A2:G" & LastRow(mIssues)).ClearContents
    CheckCurrent rules
    ApplyIssuesToSummary
    FormatOutput wb
    mSummary.Activate
    MsgBox "SELF-TEST PASSED: name additions/removals, ECID replacement, role changes, all-field updates, duplicate role rows, ambiguous names, shuffled rows/columns and yellow/red marks.", vbInformation
    Exit Sub
Failed:
    MsgBox "SELF-TEST FAILED: " & Err.Description, vbExclamation
End Sub

Private Sub FixtureOld(ByVal nm As String, ByVal eid As String, ByVal roleName As String, ByVal addressText As String, ByVal genderText As String)
    WriteRow mOld, Array(nm, eid, roleName, addressText, "1990-01-01", genderText, _
        "CN", "Passport", "CN", "P0001", "CN", "Individual")
End Sub

Private Sub FixtureNew(ByVal nm As String, ByVal eid As String, ByVal roleName As String, ByVal addressText As String, ByVal genderText As String)
    WriteRow mNew, Array(eid, roleName, nm, genderText, addressText, "1990-01-01", _
        "CN", "CN", "P0001", "CN", "Passport", "Individual")
End Sub

Private Sub AssertV5(ByVal condition As Boolean, ByVal messageText As String)
    If Not condition Then Err.Raise vbObjectError + 505, "KYC V5 self-test", messageText
End Sub
