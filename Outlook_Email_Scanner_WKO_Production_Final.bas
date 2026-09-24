Attribute VB_Name = "EmailScannerWKOProduction"
Option Explicit

' ============================================================
' Fast + safe Outlook Inbox -> Excel tracker scanner (final)
'
' Design goals:
'   1) Classic Outlook + fixed WKO shared mailbox Inbox only.
'   2) One pass through the selected date window; no costly Outlook sort.
'   3) Action subjects: any ORIGINAL message beginning "Action Required -".
'      The action text after the prefix may vary.
'   4) Completion subjects are handled separately and conservatively.
'   5) Replies/forwards never trigger Action/Completion automation.
'   6) A Completion resolves exactly one unique GSM. If no Excel row exists,
'      a completion-only row is created safely (GSM + Completed + Completion Date).
'      Ambiguous/duplicate cases are skipped, never guessed.
'   7) Existing Completed records are never moved back to Open.
'   8) Completion Date is filled only when blank.
'   9) "No requirement" on an Action makes that case Completed.
'  10) No filtering by Read/Unread or Flag status.
'  11) The user always chooses the scan start date in an input window.
'  12) Client ECID supports both IB and CB formats, e.g.
'      Client ECIDs: % IB:12345% or Client ECIDs: % CB:12345%.
'  13) Overall status uses Open/Completed and feeds Completed to Sub request STATUS.
'  14) CAU Handler is read from Outlook Categories for relevant emails.
'  15) Tracker formulas, ascending date sort, Arial 9 and two header colors are applied in bulk.
'
' Requirements:
'   - Windows desktop Excel
'   - Classic Outlook installed/configured
'   - Run from the worksheet containing the tracker
' ============================================================

Private Const OL_FOLDER_INBOX As Long = 6
Private Const OL_MAIL_ITEM_CLASS As Long = 43

' Fixed mailbox that contains the Client List Screening workflow emails.
Private Const TARGET_MAILBOX_SMTP As String = "wko.cnlib.gsm@jpmchase.com"

Private Const ACTION_PREFIX As String = "Action Required -"
Private Const COMPLETE_TEXT_1 As String = "No Further Action Required"
Private Const COMPLETE_TEXT_2A As String = "Client List Screening"
Private Const COMPLETE_TEXT_2B As String = "has been resolved"

Private Const STATUS_OPEN As String = "Open"
Private Const STATUS_COMPLETED As String = "Completed"
Private Const STATUS_PENDING_LEGACY As String = "Pending"
Private Const STATUS_COMPLETE_LEGACY As String = "Complete"

Private Const HDR_RECEIVED_DATE As String = "Received Date"
Private Const HDR_GSM_ID As String = "GSM ID"
Private Const HDR_HIT_TYPE As String = "Hit Type"
Private Const HDR_INQUIRY As String = "Inquiry information"
Private Const HDR_CLIENT_ECID As String = "Client ECID"
Private Const HDR_CLIENT_NAME As String = "Client name"
Private Const HDR_RELATED_ECID As String = "Related Party ECID"
Private Const HDR_RELATED_NAME As String = "Related Party Name"
Private Const HDR_STATUS As String = "Overall Request status"
Private Const HDR_STATUS_LEGACY As String = "Request status"
Private Const HDR_COMPLETION_DATE As String = "Completion date"
Private Const HDR_SUB_STATUS As String = "Sub request STATUS"
Private Const HDR_COMMENT As String = "Comment"
Private Const HDR_CAU_HANDLER As String = "CAU Handler"
Private Const HDR_AGING As String = "Aging"
Private Const HDR_INCOMING_MONTH As String = "Incoming Month"
Private Const HDR_COMPLETION_MONTH As String = "Completion month"
Private Const HDR_INCOMING_WEEK As String = "Incoming week"
Private Const HDR_COMPLETION_WEEK As String = "Completion week"
Private Const HDR_YEAR As String = "Year"

Private Const AUDIT_SHEET_NAME As String = "Email Scanner Audit"
Private Const ENABLE_AUDIT As Boolean = True

' Keep the prior business rule: do not process independent CDD workflows.
' This does NOT prevent extraction of a CDD-number-shaped Related Party ECID
' from a valid Client List Screening Action email.
Private Const IGNORE_CDD_ACTION_WORKFLOWS As Boolean = True


' ============================================================
' MAIN MACRO
' Alt+F8 -> ScanOutlookEmails -> Run
' ============================================================
Public Sub ScanOutlookEmails()

    Dim ws As Worksheet
    Dim headerRow As Long
    Dim headers As Object
    Dim rowByGsm As Object
    Dim duplicateGsm As Object

    Dim outlookApp As Object
    Dim outlookNamespace As Object
    Dim inbox As Object
    Dim allInboxItems As Object
    Dim scanItems As Object
    Dim currentItem As Object
    Dim itemClass As Long

    Dim scanStartDate As Date
    Dim suggestedStartDate As Date
    Dim filterErrorText As String
    Dim scannedFolderPath As String

    Dim scannedMailCount As Long
    Dim actionMailCount As Long
    Dim completionMailCount As Long
    Dim newCount As Long
    Dim updatedCount As Long
    Dim completedCount As Long
    Dim alreadyCompleteCount As Long
    Dim completionOnlyCreatedCount As Long
    Dim ambiguousGsmCount As Long
    Dim duplicateGsmSkipCount As Long
    Dim missingGsmCount As Long
    Dim replyForwardSkipCount As Long
    Dim perMailErrorCount As Long

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldCalculation As XlCalculation

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldCalculation = Application.Calculation

    On Error GoTo FatalError

    Set ws = ActiveSheet
    headerRow = FindHeaderRow(ws)

    If headerRow = 0 Then
        MsgBox "The 'GSM ID' header could not be found on the active sheet.", _
               vbExclamation, "Email Scanner"
        Exit Sub
    End If

    ' Minimal structural update for the new tracker version:
    ' rename the legacy status header if needed, then validate the layout.
    PrepareTrackerHeader ws, headerRow

    Set headers = BuildHeaderMap(ws, headerRow)
    If Not ValidateRequiredHeaders(headers) Then Exit Sub

    Set rowByGsm = CreateObject("Scripting.Dictionary")
    rowByGsm.CompareMode = vbTextCompare

    Set duplicateGsm = CreateObject("Scripting.Dictionary")
    duplicateGsm.CompareMode = vbTextCompare

    BuildGsmRowIndex ws, headerRow, headers(HDR_GSM_ID), _
                     rowByGsm, duplicateGsm

    suggestedStartDate = GetSuggestedStartDate( _
        ws, headerRow, headers(HDR_RECEIVED_DATE))

    If Not PromptForStartDate(suggestedStartDate, scanStartDate) Then Exit Sub

    ' Convert legacy Pending/Complete values to Open/Completed only after the
    ' user confirms the scan. This also synchronizes completed overall requests
    ' to Sub request STATUS.
    NormalizeExistingStatuses ws, headerRow, headers

    Set outlookApp = GetClassicOutlook()
    If outlookApp Is Nothing Then
        MsgBox "Outlook scanning could not start." & vbCrLf & vbCrLf & _
               "This macro requires Classic Outlook for Windows to be " & _
               "installed and configured." & vbCrLf & _
               "New Outlook is not supported by this VBA version.", _
               vbExclamation, "Email Scanner"
        Exit Sub
    End If

    Set outlookNamespace = outlookApp.GetNamespace("MAPI")

    ' IMPORTANT: always scan the WKO mailbox, never the user's personal
    ' default Inbox. This avoids different results on PCs with several
    ' Outlook mailboxes/accounts attached.
    Set inbox = GetMailboxInbox(outlookNamespace, TARGET_MAILBOX_SMTP)

    If inbox Is Nothing Then
        MsgBox "The WKO mailbox Inbox could not be opened:" & vbCrLf & _
               TARGET_MAILBOX_SMTP & vbCrLf & vbCrLf & _
               "Please make sure this mailbox is visible/accessible in " & _
               "Classic Outlook and that you have permission to open it." & _
               vbCrLf & vbCrLf & _
               "The scan was stopped. It did NOT fall back to your personal Inbox.", _
               vbExclamation, "Email Scanner"
        GoTo SafeExitNoSummary
    End If

    scannedFolderPath = SafeText(inbox.FolderPath)
    Set allInboxItems = inbox.Items

    ' IMPORTANT: zero results are a valid result. We do NOT silently fall
    ' back to the entire Inbox. If Restrict truly fails, stop safely.
    Set scanItems = GetDateRestrictedItems(allInboxItems, scanStartDate, _
                                           filterErrorText)

    If scanItems Is Nothing Then
        MsgBox "Outlook's date filter could not be applied, so the scan was " & _
               "stopped rather than falling back to the entire Inbox." & _
               vbCrLf & vbCrLf & filterErrorText, _
               vbExclamation, "Email Scanner"
        GoTo SafeExitNoSummary
    End If

    ' Do NOT sort the Outlook collection. Sorting can be expensive on Exchange.
    ' Order is no longer required for correctness:
    '   - Action first -> Pending, then Completion -> Complete.
    '   - Completion first -> creates/updates Complete row, then Action only
    '     supplements blank fields and never moves Complete back to Pending.

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Scanning WKO Inbox..."

    If ENABLE_AUDIT Then PrepareAuditSheet ws.Parent
    If ENABLE_AUDIT Then AuditExistingDuplicates ws.Parent, duplicateGsm

    For Each currentItem In scanItems

        ' Some Outlook stores can expose a stale/corrupt item that errors even
        ' when reading Class. Skip that one item instead of stopping the run.
        itemClass = 0
        On Error Resume Next
        Err.Clear
        itemClass = CLng(currentItem.Class)
        If Err.Number <> 0 Then
            perMailErrorCount = perMailErrorCount + 1
            Err.Clear
        End If
        On Error GoTo FatalError

        If itemClass = OL_MAIL_ITEM_CLASS Then
            scannedMailCount = scannedMailCount + 1

            If (scannedMailCount Mod 25) = 0 Then
                Application.StatusBar = "Scanning WKO Inbox... " & _
                                        Format$(scannedMailCount, "#,##0") & _
                                        " mail items checked"
            End If

            ProcessOneMail currentItem, ws, headerRow, headers, _
                           rowByGsm, duplicateGsm, _
                           actionMailCount, completionMailCount, _
                           newCount, updatedCount, completedCount, _
                           alreadyCompleteCount, completionOnlyCreatedCount, _
                           ambiguousGsmCount, _
                           duplicateGsmSkipCount, _
                           missingGsmCount, replyForwardSkipCount, _
                           perMailErrorCount
        End If

    Next currentItem

    ' Excel-only finishing steps are deliberately done once, in bulk, after
    ' Outlook scanning so they do not slow the per-email loop.
    FinalizeTracker ws, headerRow, headers

SafeExit:
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.Calculation = oldCalculation

    Set currentItem = Nothing
    Set scanItems = Nothing
    Set allInboxItems = Nothing
    Set inbox = Nothing
    Set outlookNamespace = Nothing
    Set outlookApp = Nothing

    MsgBox "Email scanning completed." & vbCrLf & vbCrLf & _
           "Inbox scanned: " & scannedFolderPath & vbCrLf & _
           "Scan start date: " & Format$(scanStartDate, "yyyy-mm-dd") & vbCrLf & _
           "Mail items checked: " & scannedMailCount & vbCrLf & _
           "Action emails identified: " & actionMailCount & vbCrLf & _
           "Completion emails identified: " & completionMailCount & vbCrLf & _
           "New records: " & newCount & vbCrLf & _
           "Existing records supplemented: " & updatedCount & vbCrLf & _
           "Cases newly completed: " & completedCount & vbCrLf & _
           "Already-completed cases confirmed: " & alreadyCompleteCount & vbCrLf & _
           "Completion-only records created: " & completionOnlyCreatedCount & vbCrLf & _
           "Ambiguous GSM emails skipped: " & ambiguousGsmCount & vbCrLf & _
           "Duplicate-GSM rows skipped: " & duplicateGsmSkipCount & vbCrLf & _
           "Emails with no GSM ID: " & missingGsmCount & vbCrLf & _
           "Reply/forward messages ignored: " & replyForwardSkipCount & vbCrLf & _
           "Per-email errors skipped: " & perMailErrorCount, _
           vbInformation, "Email Scanner"

    Exit Sub

SafeExitNoSummary:
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.Calculation = oldCalculation

    Set scanItems = Nothing
    Set allInboxItems = Nothing
    Set inbox = Nothing
    Set outlookNamespace = Nothing
    Set outlookApp = Nothing
    Exit Sub

FatalError:
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.Calculation = oldCalculation

    MsgBox "The scan stopped because of an unexpected error:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Email Scanner"

End Sub


' ============================================================
' ONE-MAIL PROCESSOR
' Each mail is isolated: one broken Outlook item cannot stop the whole scan.
' ============================================================
Private Sub ProcessOneMail(ByVal mail As Object, _
                           ByVal ws As Worksheet, _
                           ByVal headerRow As Long, _
                           ByVal headers As Object, _
                           ByVal rowByGsm As Object, _
                           ByVal duplicateGsm As Object, _
                           ByRef actionMailCount As Long, _
                           ByRef completionMailCount As Long, _
                           ByRef newCount As Long, _
                           ByRef updatedCount As Long, _
                           ByRef completedCount As Long, _
                           ByRef alreadyCompleteCount As Long, _
                           ByRef completionOnlyCreatedCount As Long, _
                           ByRef ambiguousGsmCount As Long, _
                           ByRef duplicateGsmSkipCount As Long, _
                           ByRef missingGsmCount As Long, _
                           ByRef replyForwardSkipCount As Long, _
                           ByRef perMailErrorCount As Long)

    Dim subjectText As String
    Dim bodyText As String
    Dim categoriesText As String
    Dim receivedTime As Date
    Dim isReplyForward As Boolean

    On Error GoTo MailError

    subjectText = SafeText(mail.Subject)
    receivedTime = CDate(mail.ReceivedTime)

    ' Ignore replies/forwards BEFORE classifying the workflow.
    isReplyForward = HasReplyForwardPrefix(subjectText)

    If isReplyForward Then
        If LooksLikeRelevantWorkflow(subjectText) Then
            replyForwardSkipCount = replyForwardSkipCount + 1
            If ENABLE_AUDIT Then
                WriteAudit ws.Parent, receivedTime, subjectText, "", _
                           "SKIPPED", "Reply/forward ignored"
            End If
        End If
        Exit Sub
    End If

    If IsActionEmailSubject(subjectText) Then

        actionMailCount = actionMailCount + 1
        bodyText = SafeText(mail.Body)
        categoriesText = GetMailCategoriesSafe(mail)

        ProcessActionMail ws, headerRow, headers, rowByGsm, duplicateGsm, _
                          subjectText, bodyText, categoriesText, receivedTime, _
                          newCount, updatedCount, completedCount, _
                          ambiguousGsmCount, duplicateGsmSkipCount, _
                          missingGsmCount

    ElseIf IsCompletionEmailSubject(subjectText) Then

        completionMailCount = completionMailCount + 1
        bodyText = SafeText(mail.Body)
        categoriesText = GetMailCategoriesSafe(mail)

        ProcessCompletionMail ws, headerRow, headers, rowByGsm, duplicateGsm, _
                              subjectText, bodyText, categoriesText, receivedTime, _
                              newCount, completedCount, alreadyCompleteCount, _
                              completionOnlyCreatedCount, ambiguousGsmCount, _
                              duplicateGsmSkipCount, missingGsmCount
    End If

    Exit Sub

MailError:
    perMailErrorCount = perMailErrorCount + 1
    If ENABLE_AUDIT Then
        WriteAudit ws.Parent, receivedTime, subjectText, "", _
                   "ERROR", CStr(Err.Number) & " - " & Err.Description
    End If
    Err.Clear

End Sub


' ============================================================
' ACTION
' ============================================================
Private Sub ProcessActionMail(ByVal ws As Worksheet, _
                              ByVal headerRow As Long, _
                              ByVal headers As Object, _
                              ByVal rowByGsm As Object, _
                              ByVal duplicateGsm As Object, _
                              ByVal subjectText As String, _
                              ByVal bodyText As String, _
                              ByVal categoriesText As String, _
                              ByVal receivedTime As Date, _
                              ByRef newCount As Long, _
                              ByRef updatedCount As Long, _
                              ByRef completedCount As Long, _
                              ByRef ambiguousGsmCount As Long, _
                              ByRef duplicateGsmSkipCount As Long, _
                              ByRef missingGsmCount As Long)

    Dim gsmId As String
    Dim gsmReason As String
    Dim targetRow As Long
    Dim isNew As Boolean
    Dim changed As Boolean
    Dim inquiryText As String
    Dim existingStatus As String
    Dim noRequirement As Boolean

    gsmId = ResolveSingleGsm(subjectText, bodyText, gsmReason)

    If Len(gsmId) = 0 Then
        If StrComp(gsmReason, "NO_GSM", vbTextCompare) = 0 Then
            missingGsmCount = missingGsmCount + 1
        Else
            ambiguousGsmCount = ambiguousGsmCount + 1
        End If

        If ENABLE_AUDIT Then
            WriteAudit ws.Parent, receivedTime, subjectText, "", _
                       "ACTION SKIPPED", gsmReason
        End If
        Exit Sub
    End If

    If duplicateGsm.Exists(gsmId) Then
        duplicateGsmSkipCount = duplicateGsmSkipCount + 1
        If ENABLE_AUDIT Then
            WriteAudit ws.Parent, receivedTime, subjectText, gsmId, _
                       "ACTION SKIPPED", "Duplicate GSM exists in Excel"
        End If
        Exit Sub
    End If

    If rowByGsm.Exists(gsmId) Then
        targetRow = CLng(rowByGsm(gsmId))
        isNew = False
    Else
        targetRow = GetNewDataRow(ws, headerRow, headers(HDR_GSM_ID))
        rowByGsm.Add gsmId, targetRow
        isNew = True
    End If

    inquiryText = ExtractLineValue(bodyText, "Suspect Decision Comments:")
    noRequirement = (InStr(1, inquiryText, "No requirement", _
                               vbTextCompare) > 0)

    changed = False

    changed = SetCellIfBlank(ws, targetRow, _
        headers(HDR_RECEIVED_DATE), DateValue(receivedTime), _
        "m/d/yyyy") Or changed

    changed = SetCellIfBlank(ws, targetRow, _
        headers(HDR_GSM_ID), gsmId, "@") Or changed

    changed = SetCellIfBlank(ws, targetRow, _
        headers(HDR_HIT_TYPE), _
        ExtractHitType(subjectText, gsmId)) Or changed

    changed = SetCellIfBlank(ws, targetRow, _
        headers(HDR_INQUIRY), inquiryText) Or changed

    changed = SetCellIfBlank(ws, targetRow, _
        headers(HDR_CLIENT_ECID), _
        ExtractClientEcid(bodyText), "@") Or changed

    changed = SetCellIfBlank(ws, targetRow, _
        headers(HDR_RELATED_ECID), _
        RegexFirst(bodyText, _
                   "Number\s*:\s*CDD-([0-9]+)"), "@") Or changed

    changed = SetCellIfBlank(ws, targetRow, _
        headers(HDR_RELATED_NAME), _
        ExtractRelatedPartyName(bodyText)) Or changed

    ' CAU Handler is fed from the Outlook Categories property. Existing/manual
    ' handler values are preserved.
    changed = SetCellIfBlank(ws, targetRow, _
        headers(HDR_CAU_HANDLER), NormalizeCategoryText(categoriesText)) Or changed

    existingStatus = Trim$(CStr(ws.Cells(targetRow, _
                                          headers(HDR_STATUS)).Value))

    If noRequirement Then
        ' Business rule: No requirement -> Completed.
        If Len(existingStatus) = 0 Or _
           StrComp(existingStatus, STATUS_OPEN, vbTextCompare) = 0 Or _
           StrComp(existingStatus, STATUS_PENDING_LEGACY, vbTextCompare) = 0 Then

            ws.Cells(targetRow, headers(HDR_STATUS)).Value = STATUS_COMPLETED
            SetSubStatusCompleted ws, targetRow, headers
            changed = True

            If Len(Trim$(CStr(ws.Cells(targetRow, _
                  headers(HDR_COMPLETION_DATE)).Value))) = 0 Then
                ws.Cells(targetRow, headers(HDR_COMPLETION_DATE)).Value = _
                    DateValue(receivedTime)
                ws.Cells(targetRow, headers(HDR_COMPLETION_DATE)).NumberFormat = _
                    "m/d/yyyy"
            End If

            completedCount = completedCount + 1
        ElseIf IsCompletedStatus(existingStatus) Then
            SetSubStatusCompleted ws, targetRow, headers
        End If

    Else
        ' Never move Completed or another manual status back to Open.
        If Len(existingStatus) = 0 Or _
           StrComp(existingStatus, STATUS_PENDING_LEGACY, vbTextCompare) = 0 Then
            ws.Cells(targetRow, headers(HDR_STATUS)).Value = STATUS_OPEN
            changed = True
        End If
    End If

    If isNew Then
        newCount = newCount + 1
    ElseIf changed Then
        updatedCount = updatedCount + 1
    End If

    If ENABLE_AUDIT Then
        WriteAudit ws.Parent, receivedTime, subjectText, gsmId, _
                   IIf(noRequirement, "ACTION -> COMPLETED", "ACTION PROCESSED"), _
                   "Excel row " & CStr(targetRow)
    End If

End Sub


' ============================================================
' COMPLETION
' ============================================================
Private Sub ProcessCompletionMail(ByVal ws As Worksheet, _
                                  ByVal headerRow As Long, _
                                  ByVal headers As Object, _
                                  ByVal rowByGsm As Object, _
                                  ByVal duplicateGsm As Object, _
                                  ByVal subjectText As String, _
                                  ByVal bodyText As String, _
                                  ByVal categoriesText As String, _
                                  ByVal receivedTime As Date, _
                                  ByRef newCount As Long, _
                                  ByRef completedCount As Long, _
                                  ByRef alreadyCompleteCount As Long, _
                                  ByRef completionOnlyCreatedCount As Long, _
                                  ByRef ambiguousGsmCount As Long, _
                                  ByRef duplicateGsmSkipCount As Long, _
                                  ByRef missingGsmCount As Long)

    Dim gsmId As String
    Dim gsmReason As String
    Dim targetRow As Long
    Dim currentBody As String
    Dim existingStatus As String
    Dim existingDate As String
    Dim statusCanComplete As Boolean
    Dim statusAlreadyComplete As Boolean
    Dim changed As Boolean

    gsmId = ResolveSingleGsm(subjectText, bodyText, gsmReason)

    If Len(gsmId) = 0 Then
        If StrComp(gsmReason, "NO_GSM", vbTextCompare) = 0 Then
            missingGsmCount = missingGsmCount + 1
        Else
            ambiguousGsmCount = ambiguousGsmCount + 1
        End If

        If ENABLE_AUDIT Then
            WriteAudit ws.Parent, receivedTime, subjectText, "", _
                       "COMPLETION SKIPPED", gsmReason
        End If
        Exit Sub
    End If

    If duplicateGsm.Exists(gsmId) Then
        duplicateGsmSkipCount = duplicateGsmSkipCount + 1
        If ENABLE_AUDIT Then
            WriteAudit ws.Parent, receivedTime, subjectText, gsmId, _
                       "COMPLETION SKIPPED", _
                       "Duplicate GSM exists in Excel; no row was guessed"
        End If
        Exit Sub
    End If

    ' rowByGsm was built from the ENTIRE Excel tracker before Outlook scanning.
    ' Therefore a Completion inside the chosen scan window updates an older
    ' Action row even when that Action email itself is outside the scan window.
    If Not rowByGsm.Exists(gsmId) Then
        ' Completion-only case: create a row rather than losing the event.
        ' Do NOT invent an Action Received Date. The Received Date column stays
        ' blank unless an actual Action email is later scanned for this GSM.
        targetRow = GetNewDataRow(ws, headerRow, headers(HDR_GSM_ID))
        rowByGsm.Add gsmId, targetRow

        currentBody = GetCurrentMessageBody(bodyText)

        SetCellIfBlank ws, targetRow, headers(HDR_GSM_ID), gsmId, "@"
        SetCellIfBlank ws, targetRow, headers(HDR_HIT_TYPE), _
                       ExtractHitType(subjectText, gsmId)
        SetCellIfBlank ws, targetRow, headers(HDR_INQUIRY), _
                       ExtractLineValue(currentBody, "Suspect Decision Comments:")
        SetCellIfBlank ws, targetRow, headers(HDR_CLIENT_ECID), _
                       ExtractClientEcid(currentBody), "@"
        SetCellIfBlank ws, targetRow, headers(HDR_RELATED_ECID), _
                       RegexFirst(currentBody, _
                                  "Number\s*:\s*CDD-([0-9]+)"), "@"
        SetCellIfBlank ws, targetRow, headers(HDR_RELATED_NAME), _
                       ExtractRelatedPartyName(currentBody)
        SetCellIfBlank ws, targetRow, headers(HDR_CAU_HANDLER), _
                       NormalizeCategoryText(categoriesText)

        ws.Cells(targetRow, headers(HDR_STATUS)).Value = STATUS_COMPLETED
        SetSubStatusCompleted ws, targetRow, headers
        ws.Cells(targetRow, headers(HDR_COMPLETION_DATE)).Value = _
            DateValue(receivedTime)
        ws.Cells(targetRow, headers(HDR_COMPLETION_DATE)).NumberFormat = _
            "m/d/yyyy"

        newCount = newCount + 1
        completionOnlyCreatedCount = completionOnlyCreatedCount + 1
        completedCount = completedCount + 1

        If ENABLE_AUDIT Then
            WriteAudit ws.Parent, receivedTime, subjectText, gsmId, _
                       "COMPLETION-ONLY ROW CREATED", _
                       "Excel row " & CStr(targetRow) & _
                       "; no Action email was required to create the record"
        End If
        Exit Sub
    End If

    targetRow = CLng(rowByGsm(gsmId))
    existingStatus = Trim$(CStr(ws.Cells(targetRow, _
                                         headers(HDR_STATUS)).Value))
    existingDate = Trim$(CStr(ws.Cells(targetRow, _
                                       headers(HDR_COMPLETION_DATE)).Value))

    statusCanComplete = (Len(existingStatus) = 0 Or _
                         StrComp(existingStatus, STATUS_OPEN, vbTextCompare) = 0 Or _
                         StrComp(existingStatus, STATUS_PENDING_LEGACY, _
                                 vbTextCompare) = 0)

    statusAlreadyComplete = IsCompletedStatus(existingStatus)

    ' Protect every other manual status. Do not change status OR date.
    If Not statusCanComplete And Not statusAlreadyComplete Then
        If ENABLE_AUDIT Then
            WriteAudit ws.Parent, receivedTime, subjectText, gsmId, _
                       "COMPLETION SKIPPED", _
                       "Manual status protected: " & existingStatus
        End If
        Exit Sub
    End If

    changed = False

    changed = SetCellIfBlank(ws, targetRow, _
        headers(HDR_CAU_HANDLER), NormalizeCategoryText(categoriesText)) Or changed

    If statusCanComplete Then
        ws.Cells(targetRow, headers(HDR_STATUS)).Value = STATUS_COMPLETED
        SetSubStatusCompleted ws, targetRow, headers
        changed = True
    ElseIf statusAlreadyComplete Then
        SetSubStatusCompleted ws, targetRow, headers
    End If

    If Len(existingDate) = 0 Then
        ws.Cells(targetRow, headers(HDR_COMPLETION_DATE)).Value = _
            DateValue(receivedTime)
        ws.Cells(targetRow, headers(HDR_COMPLETION_DATE)).NumberFormat = _
            "m/d/yyyy"
        changed = True
    End If

    If statusCanComplete Then
        completedCount = completedCount + 1
    ElseIf statusAlreadyComplete Then
        alreadyCompleteCount = alreadyCompleteCount + 1
    End If

    If ENABLE_AUDIT Then
        WriteAudit ws.Parent, receivedTime, subjectText, gsmId, _
                   IIf(statusCanComplete, "COMPLETED", _
                       "ALREADY COMPLETED"), _
                   "Excel row " & CStr(targetRow) & _
                   IIf(changed, "; updated", "; no change")
    End If

End Sub


' ============================================================
' SUBJECT CLASSIFICATION
' ============================================================
Private Function IsActionEmailSubject(ByVal subjectText As String) As Boolean

    Dim cleaned As String
    cleaned = Trim$(subjectText)

    If HasReplyForwardPrefix(cleaned) Then Exit Function

    If Len(cleaned) < Len(ACTION_PREFIX) Then Exit Function

    If StrComp(Left$(cleaned, Len(ACTION_PREFIX)), _
               ACTION_PREFIX, vbTextCompare) <> 0 Then Exit Function

    ' Preserve the earlier requirement to ignore independent CDD workflows.
    If IGNORE_CDD_ACTION_WORKFLOWS Then
        If InStr(1, cleaned, "CDD", vbTextCompare) > 0 And _
           InStr(1, cleaned, "Client List Screening", vbTextCompare) = 0 Then
            Exit Function
        End If
    End If

    IsActionEmailSubject = True

End Function


Private Function IsCompletionEmailSubject(ByVal subjectText As String) _
                                          As Boolean

    If HasReplyForwardPrefix(subjectText) Then Exit Function

    If InStr(1, subjectText, COMPLETE_TEXT_1, vbTextCompare) > 0 Then
        IsCompletionEmailSubject = True
        Exit Function
    End If

    If InStr(1, subjectText, COMPLETE_TEXT_2A, vbTextCompare) > 0 And _
       InStr(1, subjectText, COMPLETE_TEXT_2B, vbTextCompare) > 0 Then
        IsCompletionEmailSubject = True
    End If

End Function


Private Function LooksLikeRelevantWorkflow(ByVal subjectText As String) _
                                           As Boolean

    If InStr(1, subjectText, "Action Required", vbTextCompare) > 0 Or _
       InStr(1, subjectText, COMPLETE_TEXT_1, vbTextCompare) > 0 Or _
       (InStr(1, subjectText, COMPLETE_TEXT_2A, vbTextCompare) > 0 And _
        InStr(1, subjectText, COMPLETE_TEXT_2B, vbTextCompare) > 0) Then
        LooksLikeRelevantWorkflow = True
    End If

End Function


Private Function HasReplyForwardPrefix(ByVal subjectText As String) As Boolean

    Dim cleaned As String
    cleaned = UCase$(Trim$(subjectText))

    If Left$(cleaned, 3) = "RE:" Or _
       Left$(cleaned, 3) = "FW:" Or _
       Left$(cleaned, 4) = "FWD:" Then
        HasReplyForwardPrefix = True
    End If

End Function


' ============================================================
' SAFE GSM RESOLUTION
'
' Priority:
'   1) Subject: exactly one unique GSM -> use it.
'   2) If subject has none, current (unquoted) body must contain exactly one.
'   3) More than one unique GSM -> ambiguous -> skip.
' This prevents a reply/forward history from completing the wrong case.
' ============================================================
Private Function ResolveSingleGsm(ByVal subjectText As String, _
                                  ByVal bodyText As String, _
                                  ByRef reason As String) As String

    Dim ids As Object
    Dim currentBody As String
    Dim keyArray As Variant

    Set ids = ExtractUniqueGsmIds(subjectText)

    If ids.Count = 1 Then
        keyArray = ids.Keys
        ResolveSingleGsm = CStr(keyArray(0))
        reason = "OK_SUBJECT"
        Exit Function
    ElseIf ids.Count > 1 Then
        reason = "AMBIGUOUS_GSM_IN_SUBJECT"
        Exit Function
    End If

    currentBody = GetCurrentMessageBody(bodyText)
    Set ids = ExtractUniqueGsmIds(currentBody)

    If ids.Count = 1 Then
        keyArray = ids.Keys
        ResolveSingleGsm = CStr(keyArray(0))
        reason = "OK_BODY"
    ElseIf ids.Count = 0 Then
        reason = "NO_GSM"
    Else
        reason = "AMBIGUOUS_GSM_IN_BODY"
    End If

End Function


Private Function ExtractUniqueGsmIds(ByVal sourceText As String) As Object

    Dim result As Object
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim gsmId As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "GSM[0-9]+-[0-9]+"
    regex.IgnoreCase = True
    regex.Global = True
    regex.MultiLine = True

    If regex.Test(sourceText) Then
        Set matches = regex.Execute(sourceText)

        For Each match In matches
            gsmId = NormalizeGsm(CStr(match.Value))
            If Len(gsmId) > 0 Then
                If Not result.Exists(gsmId) Then result.Add gsmId, True
            End If
        Next match
    End If

    Set ExtractUniqueGsmIds = result

End Function


Private Function NormalizeGsm(ByVal gsmId As String) As String
    NormalizeGsm = UCase$(Trim$(gsmId))
End Function


' Remove common quoted-message/history sections before GSM extraction.
Private Function GetCurrentMessageBody(ByVal bodyText As String) As String

    Dim normalized As String
    Dim lines As Variant
    Dim i As Long
    Dim lineText As String
    Dim result As String
    Dim seenContent As Boolean

    normalized = Replace(bodyText, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)
    lines = Split(normalized, vbLf)

    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(CStr(lines(i)))

        If IsQuotedHistoryStart(lineText, seenContent) Then Exit For

        result = result & CStr(lines(i)) & vbLf
        If Len(lineText) > 0 Then seenContent = True
    Next i

    GetCurrentMessageBody = result

End Function


Private Function IsQuotedHistoryStart(ByVal lineText As String, _
                                      ByVal seenContent As Boolean) As Boolean

    Dim upperLine As String
    upperLine = UCase$(Trim$(lineText))

    If InStr(1, upperLine, "-----ORIGINAL MESSAGE-----", _
             vbTextCompare) > 0 Then
        IsQuotedHistoryStart = True
        Exit Function
    End If

    If Len(upperLine) >= 5 Then
        If Left$(upperLine, 5) = "_____" Then
            IsQuotedHistoryStart = True
            Exit Function
        End If
    End If

    ' Only treat From:/Sent: as quoted-history markers after some actual
    ' message content has already appeared. This avoids cutting a body whose
    ' first legitimate field happens to begin with a similar word.
    If seenContent Then
        If Left$(upperLine, 5) = "FROM:" Or _
           Left$(upperLine, 5) = "SENT:" Then
            IsQuotedHistoryStart = True
        End If
    End If

End Function


' ============================================================
' OUTLOOK DATE FILTER
' ============================================================
Private Function GetDateRestrictedItems(ByVal allInboxItems As Object, _
                                        ByVal scanStartDate As Date, _
                                        ByRef errorText As String) As Object

    Dim filteredItems As Object
    Dim filterText As String
    Dim errorNumber As Long
    Dim errorDescription As String

    ' Outlook Jet filters expect the local Windows short-date/time format.
    ' Midnight includes the complete selected start date.
    filterText = "[ReceivedTime] >= '" & _
                 Format$(DateValue(scanStartDate), "ddddd h:nn AMPM") & "'"

    On Error Resume Next
    Err.Clear
    Set filteredItems = allInboxItems.Restrict(filterText)
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Clear
    On Error GoTo 0

    If errorNumber <> 0 Or filteredItems Is Nothing Then
        errorText = "Filter: " & filterText & vbCrLf & _
                    CStr(errorNumber) & " - " & errorDescription
        Set GetDateRestrictedItems = Nothing
        Exit Function
    End If

    ' Do NOT treat Count = 0 as an error.
    Set GetDateRestrictedItems = filteredItems

End Function


' ============================================================
' START-DATE SUGGESTION
'
' The USER always chooses the actual scan start date in the InputBox.
' This function only supplies a convenient editable default.
' It deliberately does NOT force a scan back to the earliest Open case,
' because an old Open row could make every run unnecessarily slow.
' ============================================================
Private Function GetSuggestedStartDate(ByVal ws As Worksheet, _
                                       ByVal headerRow As Long, _
                                       ByVal receivedDateCol As Long) As Date

    Dim lastRow As Long
    Dim row As Long
    Dim cellValue As Variant
    Dim candidateDate As Date
    Dim latestDate As Date

    lastRow = ws.Cells(ws.Rows.Count, receivedDateCol).End(xlUp).Row

    For row = headerRow + 1 To lastRow
        cellValue = ws.Cells(row, receivedDateCol).Value

        If IsDate(cellValue) Then
            candidateDate = DateValue(CDate(cellValue))
            If candidateDate > latestDate Then latestDate = candidateDate
        End If
    Next row

    If latestDate > 0 Then
        GetSuggestedStartDate = latestDate - 1
    Else
        GetSuggestedStartDate = Date - 7
    End If

    If GetSuggestedStartDate > Date Then GetSuggestedStartDate = Date

End Function


Private Function PromptForStartDate(ByVal suggestedDate As Date, _
                                    ByRef scanStartDate As Date) As Boolean

    Dim answer As Variant
    Dim enteredText As String

    Do
        answer = Application.InputBox( _
            Prompt:="Enter the earliest Received Date to scan." & _
                   vbCrLf & vbCrLf & _
                   "Use YYYY-MM-DD, for example 2026-09-08." & vbCrLf & _
                   "The whole start date is included." & vbCrLf & _
                   "You control the scan window; the default is only a suggestion.", _
            Title:="Client List Screening Email Scanner", _
            Default:=Format$(suggestedDate, "yyyy-mm-dd"), _
            Type:=2)

        If VarType(answer) = vbBoolean Then
            If answer = False Then Exit Function
        End If

        enteredText = Trim$(CStr(answer))

        If TryParseIsoDate(enteredText, scanStartDate) Then
            PromptForStartDate = True
            Exit Function
        End If

        MsgBox "Please enter a valid date in YYYY-MM-DD format.", _
               vbExclamation, "Email Scanner"
    Loop

End Function


Private Function TryParseIsoDate(ByVal textValue As String, _
                                 ByRef parsedDate As Date) As Boolean

    Dim parts As Variant
    Dim y As Long
    Dim m As Long
    Dim d As Long
    Dim candidateDate As Date

    On Error GoTo InvalidDate

    parts = Split(textValue, "-")
    If UBound(parts) <> 2 Then GoTo InvalidDate

    If Not IsNumeric(parts(0)) Or _
       Not IsNumeric(parts(1)) Or _
       Not IsNumeric(parts(2)) Then GoTo InvalidDate

    y = CLng(parts(0))
    m = CLng(parts(1))
    d = CLng(parts(2))

    candidateDate = DateSerial(y, m, d)

    If Year(candidateDate) <> y Or _
       Month(candidateDate) <> m Or _
       Day(candidateDate) <> d Then GoTo InvalidDate

    parsedDate = candidateDate
    TryParseIsoDate = True
    Exit Function

InvalidDate:
    TryParseIsoDate = False

End Function



' ============================================================
' TRACKER STRUCTURE / STATUS NORMALIZATION / FINALIZATION
' ============================================================
Private Sub PrepareTrackerHeader(ByVal ws As Worksheet, _
                                 ByVal headerRow As Long)

    Dim newStatusCell As Range
    Dim oldStatusCell As Range

    Set newStatusCell = ws.Rows(headerRow).Find( _
        What:=HDR_STATUS, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

    If newStatusCell Is Nothing Then
        Set oldStatusCell = ws.Rows(headerRow).Find( _
            What:=HDR_STATUS_LEGACY, LookIn:=xlValues, _
            LookAt:=xlWhole, MatchCase:=False)

        If Not oldStatusCell Is Nothing Then
            oldStatusCell.Value = HDR_STATUS
        End If
    End If

End Sub


Private Sub NormalizeExistingStatuses(ByVal ws As Worksheet, _
                                      ByVal headerRow As Long, _
                                      ByVal headers As Object)

    Dim lastRow As Long
    Dim row As Long
    Dim statusText As String

    If Not headers.Exists(HDR_STATUS) Then Exit Sub
    If Not headers.Exists(HDR_GSM_ID) Then Exit Sub

    lastRow = GetTrackerLastRow(ws, headerRow, headers(HDR_GSM_ID))

    For row = headerRow + 1 To lastRow
        statusText = Trim$(CStr(ws.Cells(row, headers(HDR_STATUS)).Value))

        If StrComp(statusText, STATUS_PENDING_LEGACY, vbTextCompare) = 0 Then
            ws.Cells(row, headers(HDR_STATUS)).Value = STATUS_OPEN

        ElseIf StrComp(statusText, STATUS_COMPLETE_LEGACY, vbTextCompare) = 0 Then
            ws.Cells(row, headers(HDR_STATUS)).Value = STATUS_COMPLETED
            SetSubStatusCompleted ws, row, headers

        ElseIf StrComp(statusText, STATUS_COMPLETED, vbTextCompare) = 0 Then
            SetSubStatusCompleted ws, row, headers
        End If
    Next row

End Sub


Private Function IsCompletedStatus(ByVal statusText As String) As Boolean

    IsCompletedStatus = _
        (StrComp(Trim$(statusText), STATUS_COMPLETED, vbTextCompare) = 0 Or _
         StrComp(Trim$(statusText), STATUS_COMPLETE_LEGACY, vbTextCompare) = 0)

End Function


Private Sub SetSubStatusCompleted(ByVal ws As Worksheet, _
                                  ByVal targetRow As Long, _
                                  ByVal headers As Object)

    If headers.Exists(HDR_SUB_STATUS) Then
        ws.Cells(targetRow, headers(HDR_SUB_STATUS)).Value = STATUS_COMPLETED
    End If

End Sub


Private Function NormalizeCategoryText(ByVal categoriesText As String) As String

    Dim cleaned As String
    Dim separatorPosition As Long

    cleaned = Trim$(categoriesText)
    If Len(cleaned) = 0 Then Exit Function

    ' CAU Handler is a single field. If Outlook ever exposes more than one
    ' category, take the first category deterministically rather than storing
    ' a comma/semicolon-delimited list.
    separatorPosition = InStr(1, cleaned, ",", vbTextCompare)
    If separatorPosition = 0 Then
        separatorPosition = InStr(1, cleaned, ";", vbTextCompare)
    End If

    If separatorPosition > 0 Then
        cleaned = Trim$(Left$(cleaned, separatorPosition - 1))
    End If

    NormalizeCategoryText = cleaned

End Function


Private Sub FinalizeTracker(ByVal ws As Worksheet, _
                            ByVal headerRow As Long, _
                            ByVal headers As Object)

    ' Re-run normalization after the scan so every newly Completed case feeds
    ' "Completed" into Sub request STATUS.
    NormalizeExistingStatuses ws, headerRow, headers

    ' User requirement: oldest Received Date first.
    SortTrackerByReceivedDate ws, headerRow, headers

    ' Formulas are written once, in bulk, after sorting. This is materially
    ' faster than writing formulas inside the Outlook loop.
    ApplyTrackerFormulas ws, headerRow, headers

    ' Keep both status columns usable without manual setup.
    ApplyStatusValidations ws, headerRow, headers

    ' Formatting is applied in one bulk operation and therefore has negligible
    ' impact on the Outlook scan time.
    ApplyTrackerFormatting ws, headerRow, headers

End Sub


Private Function GetTrackerLastRow(ByVal ws As Worksheet, _
                                   ByVal headerRow As Long, _
                                   ByVal gsmCol As Long) As Long

    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, gsmCol).End(xlUp).Row
    If lastRow < headerRow Then lastRow = headerRow

    GetTrackerLastRow = lastRow

End Function


Private Sub SortTrackerByReceivedDate(ByVal ws As Worksheet, _
                                      ByVal headerRow As Long, _
                                      ByVal headers As Object)

    Dim lastRow As Long
    Dim lastCol As Long
    Dim sortRange As Range

    lastRow = GetTrackerLastRow(ws, headerRow, headers(HDR_GSM_ID))
    If lastRow <= headerRow + 1 Then Exit Sub

    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    Set sortRange = ws.Range(ws.Cells(headerRow, 1), _
                             ws.Cells(lastRow, lastCol))

    sortRange.Sort Key1:=ws.Cells(headerRow + 1, _
                                  headers(HDR_RECEIVED_DATE)), _
                   Order1:=xlAscending, Header:=xlYes, _
                   MatchCase:=False, Orientation:=xlTopToBottom

End Sub


Private Sub ApplyTrackerFormulas(ByVal ws As Worksheet, _
                                 ByVal headerRow As Long, _
                                 ByVal headers As Object)

    Dim lastRow As Long
    Dim holidayLastRow As Long

    Dim agingCol As Long
    Dim incomingMonthCol As Long
    Dim completionMonthCol As Long
    Dim incomingWeekCol As Long
    Dim completionWeekCol As Long
    Dim yearCol As Long

    Dim rDate As String
    Dim rHit As String
    Dim rStatus As String
    Dim rCompletion As String
    Dim holidayDates As String
    Dim designatedDays As String
    Dim formulaText As String

    lastRow = GetTrackerLastRow(ws, headerRow, headers(HDR_GSM_ID))
    If lastRow <= headerRow Then Exit Sub

    holidayLastRow = GetHolidayLastRow(ws.Parent)
    If holidayLastRow < 2 Then holidayLastRow = 2

    agingCol = headers(HDR_AGING)
    incomingMonthCol = headers(HDR_INCOMING_MONTH)
    completionMonthCol = headers(HDR_COMPLETION_MONTH)
    incomingWeekCol = headers(HDR_INCOMING_WEEK)
    completionWeekCol = headers(HDR_COMPLETION_WEEK)
    yearCol = headers(HDR_YEAR)

    holidayDates = "Holiday!R2C2:R" & CStr(holidayLastRow) & "C2"
    designatedDays = "Holiday!R2C3:R" & CStr(holidayLastRow) & "C3"

    ' Aging: same business logic as the approved tracker formula.
    ' OFAC uses calendar days. Other hit types use NETWORKDAYS, excluding
    ' China holidays and adding designated working days back in.
    rDate = RelativeR1C1(headers(HDR_RECEIVED_DATE), agingCol)
    rHit = RelativeR1C1(headers(HDR_HIT_TYPE), agingCol)
    rStatus = RelativeR1C1(headers(HDR_STATUS), agingCol)
    rCompletion = RelativeR1C1(headers(HDR_COMPLETION_DATE), agingCol)

    formulaText = "=IF(" & rDate & "="""","""",IF(" & rHit & "=""OFAC"","
    formulaText = formulaText & "IF(AND(" & rStatus & "=""Completed""," & rCompletion & "<>""""),"
    formulaText = formulaText & rCompletion & "-" & rDate & ",TODAY()-" & rDate & "),"
    formulaText = formulaText & "IF(AND(" & rStatus & "=""Completed""," & rCompletion & "<>""""),"
    formulaText = formulaText & "NETWORKDAYS(" & rDate & "," & rCompletion & "," & holidayDates & ")"
    formulaText = formulaText & "+COUNTIFS(" & designatedDays & ","">=""&" & rDate & ","
    formulaText = formulaText & designatedDays & ",""<=""&" & rCompletion & ")-1,"
    formulaText = formulaText & "NETWORKDAYS(" & rDate & ",TODAY()," & holidayDates & ")"
    formulaText = formulaText & "+COUNTIFS(" & designatedDays & ","">=""&" & rDate & ","
    formulaText = formulaText & designatedDays & ",""<=""&TODAY())-1)))"

    ws.Range(ws.Cells(headerRow + 1, agingCol), _
             ws.Cells(lastRow, agingCol)).FormulaR1C1 = formulaText

    ' Incoming Month
    rDate = RelativeR1C1(headers(HDR_RECEIVED_DATE), incomingMonthCol)
    formulaText = "=IF(" & rDate & "="""","""",TEXT(" & rDate & ",""mmm""))"
    ws.Range(ws.Cells(headerRow + 1, incomingMonthCol), _
             ws.Cells(lastRow, incomingMonthCol)).FormulaR1C1 = formulaText

    ' Completion month
    rCompletion = RelativeR1C1(headers(HDR_COMPLETION_DATE), completionMonthCol)
    formulaText = "=IF(" & rCompletion & "<>"""",TEXT(" & rCompletion & ",""mmm""),""nil"")"
    ws.Range(ws.Cells(headerRow + 1, completionMonthCol), _
             ws.Cells(lastRow, completionMonthCol)).FormulaR1C1 = formulaText

    ' Incoming week: Monday through Sunday, e.g. 21SEP~27SEP.
    rDate = RelativeR1C1(headers(HDR_RECEIVED_DATE), incomingWeekCol)
    formulaText = "=IF(" & rDate & "="""","""",TEXT(" & rDate & "-WEEKDAY(" & rDate & ",2)+1,""DDMMM"")"
    formulaText = formulaText & "&""~""&TEXT(" & rDate & "-WEEKDAY(" & rDate & ",2)+7,""DDMMM""))"
    ws.Range(ws.Cells(headerRow + 1, incomingWeekCol), _
             ws.Cells(lastRow, incomingWeekCol)).FormulaR1C1 = formulaText

    ' Completion week
    rCompletion = RelativeR1C1(headers(HDR_COMPLETION_DATE), completionWeekCol)
    formulaText = "=IF(" & rCompletion & "="""","""",TEXT(" & rCompletion & "-WEEKDAY(" & rCompletion & ",2)+1,""DDMMM"")"
    formulaText = formulaText & "&""~""&TEXT(" & rCompletion & "-WEEKDAY(" & rCompletion & ",2)+7,""DDMMM""))"
    ws.Range(ws.Cells(headerRow + 1, completionWeekCol), _
             ws.Cells(lastRow, completionWeekCol)).FormulaR1C1 = formulaText

    ' Year
    rDate = RelativeR1C1(headers(HDR_RECEIVED_DATE), yearCol)
    formulaText = "=IF(" & rDate & "="""","""",YEAR(" & rDate & "))"
    ws.Range(ws.Cells(headerRow + 1, yearCol), _
             ws.Cells(lastRow, yearCol)).FormulaR1C1 = formulaText

    ws.Range(ws.Cells(headerRow + 1, agingCol), _
             ws.Cells(lastRow, yearCol)).Calculate

End Sub


Private Function RelativeR1C1(ByVal referenceCol As Long, _
                              ByVal formulaCol As Long) As String

    Dim offsetValue As Long

    offsetValue = referenceCol - formulaCol

    If offsetValue = 0 Then
        RelativeR1C1 = "RC"
    Else
        RelativeR1C1 = "RC[" & CStr(offsetValue) & "]"
    End If

End Function


Private Function GetHolidayLastRow(ByVal wb As Workbook) As Long

    Dim holidayWs As Worksheet
    Dim lastB As Long
    Dim lastC As Long

    On Error Resume Next
    Set holidayWs = wb.Worksheets("Holiday")
    On Error GoTo 0

    If holidayWs Is Nothing Then
        GetHolidayLastRow = 34
        Exit Function
    End If

    lastB = holidayWs.Cells(holidayWs.Rows.Count, 2).End(xlUp).Row
    lastC = holidayWs.Cells(holidayWs.Rows.Count, 3).End(xlUp).Row

    GetHolidayLastRow = Application.Max(lastB, lastC, 2)

End Function


Private Sub ApplyStatusValidations(ByVal ws As Worksheet, _
                                   ByVal headerRow As Long, _
                                   ByVal headers As Object)

    Dim lastRow As Long
    Dim separatorText As String
    Dim statusRange As Range

    lastRow = GetTrackerLastRow(ws, headerRow, headers(HDR_GSM_ID))
    If lastRow <= headerRow Then Exit Sub

    separatorText = Application.International(xlListSeparator)
    Set statusRange = ws.Range(ws.Cells(headerRow + 1, headers(HDR_STATUS)), _
                               ws.Cells(lastRow, headers(HDR_STATUS)))

    On Error Resume Next
    statusRange.Validation.Delete
    On Error GoTo 0

    statusRange.Validation.Add Type:=xlValidateList, _
                               AlertStyle:=xlValidAlertStop, _
                               Operator:=xlBetween, _
                               Formula1:=STATUS_OPEN & separatorText & STATUS_COMPLETED

    statusRange.Validation.IgnoreBlank = True
    statusRange.Validation.InCellDropdown = True
    statusRange.Validation.ErrorTitle = "Invalid status"
    statusRange.Validation.ErrorMessage = _
        "Please select Open or Completed."
    statusRange.Validation.ShowError = True

    EnsureSubRequestValidation ws, headerRow, headers, lastRow

End Sub


Private Sub EnsureSubRequestValidation(ByVal ws As Worksheet, _
                                       ByVal headerRow As Long, _
                                       ByVal headers As Object, _
                                       ByVal lastRow As Long)

    Dim subCol As Long
    Dim sourceRow As Long
    Dim checkRow As Long
    Dim validationType As Long
    Dim formula1Text As String
    Dim targetRange As Range

    Dim listWs As Worksheet
    Dim firstUsed As Range
    Dim firstListRow As Long
    Dim listCol As Long
    Dim lastListRow As Long
    Dim firstValue As String
    Dim listName As String
    Dim refersToText As String

    If Not headers.Exists(HDR_SUB_STATUS) Then Exit Sub

    subCol = headers(HDR_SUB_STATUS)
    Set targetRange = ws.Range(ws.Cells(headerRow + 1, subCol), _
                               ws.Cells(lastRow, subCol))

    ' First preserve the preset dropdown already present in the workbook.
    sourceRow = 0
    For checkRow = headerRow + 1 To Application.Min(lastRow, headerRow + 50)
        validationType = 0
        On Error Resume Next
        validationType = ws.Cells(checkRow, subCol).Validation.Type
        On Error GoTo 0

        If validationType = xlValidateList Then
            sourceRow = checkRow
            Exit For
        End If
    Next checkRow

    If sourceRow > 0 Then
        On Error Resume Next
        formula1Text = ws.Cells(sourceRow, subCol).Validation.Formula1
        On Error GoTo 0

        If Len(formula1Text) > 0 Then
            On Error Resume Next
            targetRange.Validation.Delete
            On Error GoTo 0

            targetRange.Validation.Add Type:=xlValidateList, _
                                       AlertStyle:=xlValidAlertStop, _
                                       Operator:=xlBetween, _
                                       Formula1:=formula1Text
            targetRange.Validation.IgnoreBlank = True
            targetRange.Validation.InCellDropdown = True
            Exit Sub
        End If
    End If

    ' Fallback: build the dropdown from the dedicated "Sub request STATUS"
    ' sheet so newly created rows work without any manual setup.
    On Error Resume Next
    Set listWs = ws.Parent.Worksheets(HDR_SUB_STATUS)
    On Error GoTo 0
    If listWs Is Nothing Then Exit Sub

    Set firstUsed = listWs.Cells.Find(What:="*", _
                                      After:=listWs.Cells(1, 1), _
                                      LookIn:=xlValues, _
                                      LookAt:=xlPart, _
                                      SearchOrder:=xlByRows, _
                                      SearchDirection:=xlNext, _
                                      MatchCase:=False)
    If firstUsed Is Nothing Then Exit Sub

    firstListRow = firstUsed.Row
    listCol = firstUsed.Column
    lastListRow = listWs.Cells(listWs.Rows.Count, listCol).End(xlUp).Row

    firstValue = Trim$(CStr(listWs.Cells(firstListRow, listCol).Value))
    If InStr(1, firstValue, "status", vbTextCompare) > 0 And _
       lastListRow > firstListRow Then
        firstListRow = firstListRow + 1
    End If

    If lastListRow < firstListRow Then Exit Sub

    listName = "_EmailScanner_SubStatusList"
    refersToText = "='" & Replace(listWs.Name, "'", "''") & "'!" & _
                   listWs.Range(listWs.Cells(firstListRow, listCol), _
                                listWs.Cells(lastListRow, listCol)).Address

    On Error Resume Next
    ws.Parent.Names(listName).Delete
    On Error GoTo 0

    ws.Parent.Names.Add Name:=listName, RefersTo:=refersToText

    On Error Resume Next
    targetRange.Validation.Delete
    On Error GoTo 0

    targetRange.Validation.Add Type:=xlValidateList, _
                               AlertStyle:=xlValidAlertStop, _
                               Operator:=xlBetween, _
                               Formula1:="=" & listName
    targetRange.Validation.IgnoreBlank = True
    targetRange.Validation.InCellDropdown = True

End Sub


Private Sub ApplyTrackerFormatting(ByVal ws As Worksheet, _
                                   ByVal headerRow As Long, _
                                   ByVal headers As Object)

    Dim lastRow As Long
    Dim lastCol As Long
    Dim trackerRange As Range
    Dim headerRange As Range
    Dim greenFill As Long
    Dim orangeFill As Long
    Dim orangeHeaders As Variant
    Dim item As Variant

    lastRow = GetTrackerLastRow(ws, headerRow, headers(HDR_GSM_ID))
    If lastRow < headerRow Then lastRow = headerRow

    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    Set trackerRange = ws.Range(ws.Cells(headerRow, 1), _
                                ws.Cells(lastRow, lastCol))

    trackerRange.Font.Name = "Arial"
    trackerRange.Font.Size = 9

    ' Exactly two header colors, matching the green/orange convention in the
    ' supplied tracker. The fill is changed only on the header row.
    greenFill = RGB(216, 224, 189)
    orangeFill = RGB(244, 176, 91)

    Set headerRange = ws.Range(ws.Cells(headerRow, 1), _
                               ws.Cells(headerRow, lastCol))
    headerRange.Interior.Color = greenFill
    headerRange.Font.Name = "Arial"
    headerRange.Font.Size = 9
    headerRange.Font.Color = vbBlack

    orangeHeaders = Array( _
        HDR_CLIENT_ECID, HDR_CLIENT_NAME, HDR_STATUS, _
        HDR_COMPLETION_DATE, HDR_SUB_STATUS, HDR_COMMENT, HDR_CAU_HANDLER)

    For Each item In orangeHeaders
        If headers.Exists(CStr(item)) Then
            ws.Cells(headerRow, headers(CStr(item))).Interior.Color = orangeFill
        End If
    Next item

End Sub


' ============================================================
' EXCEL INDEX / DUPLICATE PROTECTION
' ============================================================
Private Sub BuildGsmRowIndex(ByVal ws As Worksheet, _
                             ByVal headerRow As Long, _
                             ByVal gsmCol As Long, _
                             ByVal rowByGsm As Object, _
                             ByVal duplicateGsm As Object)

    Dim lastRow As Long
    Dim row As Long
    Dim gsmId As String
    Dim firstRow As Long

    lastRow = ws.Cells(ws.Rows.Count, gsmCol).End(xlUp).Row

    For row = headerRow + 1 To lastRow
        gsmId = NormalizeGsm(CStr(ws.Cells(row, gsmCol).Value))

        If Len(gsmId) > 0 Then
            If duplicateGsm.Exists(gsmId) Then
                duplicateGsm(gsmId) = duplicateGsm(gsmId) & ", " & CStr(row)

            ElseIf rowByGsm.Exists(gsmId) Then
                firstRow = CLng(rowByGsm(gsmId))
                rowByGsm.Remove gsmId
                duplicateGsm.Add gsmId, CStr(firstRow) & ", " & CStr(row)

            Else
                rowByGsm.Add gsmId, row
            End If
        End If
    Next row

End Sub


Private Function FindHeaderRow(ByVal ws As Worksheet) As Long

    Dim foundCell As Range

    Set foundCell = ws.Cells.Find(What:=HDR_GSM_ID, _
                                  After:=ws.Cells(1, 1), _
                                  LookIn:=xlValues, _
                                  LookAt:=xlWhole, _
                                  SearchOrder:=xlByRows, _
                                  SearchDirection:=xlNext, _
                                  MatchCase:=False)

    If Not foundCell Is Nothing Then FindHeaderRow = foundCell.Row

End Function


Private Function BuildHeaderMap(ByVal ws As Worksheet, _
                                ByVal headerRow As Long) As Object

    Dim result As Object
    Dim lastCol As Long
    Dim col As Long
    Dim headerText As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    For col = 1 To lastCol
        headerText = Trim$(CStr(ws.Cells(headerRow, col).Value))
        If Len(headerText) > 0 Then
            If Not result.Exists(headerText) Then result.Add headerText, col
        End If
    Next col

    Set BuildHeaderMap = result

End Function


Private Function ValidateRequiredHeaders(ByVal headers As Object) As Boolean

    Dim requiredHeaders As Variant
    Dim item As Variant
    Dim missing As String

    requiredHeaders = Array( _
        HDR_RECEIVED_DATE, HDR_GSM_ID, HDR_HIT_TYPE, HDR_INQUIRY, _
        HDR_CLIENT_ECID, HDR_RELATED_ECID, HDR_RELATED_NAME, _
        HDR_STATUS, HDR_COMPLETION_DATE, HDR_SUB_STATUS, _
        HDR_CAU_HANDLER, HDR_AGING, HDR_INCOMING_MONTH, _
        HDR_COMPLETION_MONTH, HDR_INCOMING_WEEK, _
        HDR_COMPLETION_WEEK, HDR_YEAR)

    For Each item In requiredHeaders
        If Not headers.Exists(CStr(item)) Then
            missing = missing & vbCrLf & "- " & CStr(item)
        End If
    Next item

    If Len(missing) > 0 Then
        MsgBox "These required headers were not found:" & vbCrLf & _
               missing, vbExclamation, "Email Scanner"
        ValidateRequiredHeaders = False
    Else
        ValidateRequiredHeaders = True
    End If

End Function


Private Function GetNewDataRow(ByVal ws As Worksheet, _
                               ByVal headerRow As Long, _
                               ByVal gsmCol As Long) As Long

    Dim tbl As ListObject
    Dim newListRow As ListRow

    For Each tbl In ws.ListObjects
        If tbl.HeaderRowRange.Row = headerRow Then
            If gsmCol >= tbl.Range.Column And _
               gsmCol <= tbl.Range.Column + tbl.Range.Columns.Count - 1 Then

                Set newListRow = tbl.ListRows.Add
                GetNewDataRow = newListRow.Range.Row
                Exit Function
            End If
        End If
    Next tbl

    GetNewDataRow = ws.Cells(ws.Rows.Count, gsmCol).End(xlUp).Row + 1
    If GetNewDataRow <= headerRow Then GetNewDataRow = headerRow + 1

End Function


Private Function SetCellIfBlank(ByVal ws As Worksheet, _
                                ByVal row As Long, _
                                ByVal col As Long, _
                                ByVal newValue As Variant, _
                                Optional ByVal numberFormat As String = "") _
                                As Boolean

    If Len(Trim$(CStr(newValue))) = 0 Then Exit Function

    If Len(Trim$(CStr(ws.Cells(row, col).Value))) = 0 Then
        If Len(numberFormat) > 0 Then
            ws.Cells(row, col).NumberFormat = numberFormat
        End If

        ws.Cells(row, col).Value = newValue
        SetCellIfBlank = True
    End If

End Function


' ============================================================
' FIELD EXTRACTION
' ============================================================
Private Function ExtractHitType(ByVal subjectText As String, _
                                ByVal gsmId As String) As String

    Dim cleanedSubject As String
    Dim parts As Variant
    Dim i As Long

    cleanedSubject = Replace(subjectText, vbCr, " ")
    cleanedSubject = Replace(cleanedSubject, vbLf, " ")
    parts = Split(cleanedSubject, ",")

    For i = LBound(parts) To UBound(parts)
        If InStr(1, CStr(parts(i)), gsmId, vbTextCompare) > 0 Then
            If i < UBound(parts) Then
                ExtractHitType = Trim$(CStr(parts(i + 1)))
            End If
            Exit Function
        End If
    Next i

End Function


Private Function ExtractClientEcid(ByVal bodyText As String) As String

    ' Supports the two observed formats:
    '   Client ECID(s): % IB:123456%
    '   Client ECID(s): % CB:123456%
    ' Also tolerates spaces and either normal/full-width colons.
    ExtractClientEcid = RegexFirst(bodyText, _
        "Client ECIDs?\s*[:：]\s*%?\s*[IC]B\s*[:：]\s*([0-9]+)\s*%?")

End Function


Private Function ExtractRelatedPartyName(ByVal bodyText As String) As String

    Dim startPosition As Long
    Dim clientBlock As String
    Dim extractedName As String

    startPosition = InStr(1, bodyText, "Client Information", vbTextCompare)
    If startPosition = 0 Then Exit Function

    clientBlock = Mid$(bodyText, startPosition, 3000)

    extractedName = RegexFirst(clientBlock, _
        "^[ \t]*Name[ \t]*:[ \t]*([^\r\n]+)")

    If Len(extractedName) = 0 Then
        extractedName = RegexFirst(clientBlock, _
            "^[ \t]*Name[ \t]+([^\r\n]+)")
    End If

    If Len(extractedName) = 0 Then
        extractedName = ExtractLineValue(clientBlock, "Name")
    End If

    If InStr(1, extractedName, vbTab) > 0 Then
        extractedName = Trim$(Split(extractedName, vbTab)(0))
    End If

    ExtractRelatedPartyName = extractedName

End Function


Private Function ExtractLineValue(ByVal sourceText As String, _
                                  ByVal labelText As String) As String

    Dim labelPosition As Long
    Dim remainingText As String
    Dim lines As Variant
    Dim i As Long
    Dim valueText As String

    labelPosition = InStr(1, sourceText, labelText, vbTextCompare)
    If labelPosition = 0 Then Exit Function

    remainingText = Mid$(sourceText, labelPosition + Len(labelText))
    remainingText = Replace(remainingText, vbCrLf, vbLf)
    remainingText = Replace(remainingText, vbCr, vbLf)
    lines = Split(remainingText, vbLf)

    For i = LBound(lines) To UBound(lines)
        valueText = Trim$(CStr(lines(i)))

        If Len(valueText) > 0 Then
            ExtractLineValue = valueText
            Exit Function
        End If

        If i >= 3 Then Exit For
    Next i

End Function


Private Function RegexFirst(ByVal sourceText As String, _
                            ByVal patternText As String) As String

    Dim regex As Object
    Dim matches As Object

    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = patternText
    regex.IgnoreCase = True
    regex.Global = False
    regex.MultiLine = True

    If regex.Test(sourceText) Then
        Set matches = regex.Execute(sourceText)

        If matches(0).SubMatches.Count > 0 Then
            RegexFirst = Trim$(CStr(matches(0).SubMatches(0)))
        Else
            RegexFirst = Trim$(CStr(matches(0).Value))
        End If
    End If

End Function


Private Function SafeText(ByVal value As Variant) As String
    If IsNull(value) Or IsEmpty(value) Then
        SafeText = ""
    Else
        SafeText = CStr(value)
    End If
End Function


Private Function GetMailCategoriesSafe(ByVal mail As Object) As String

    On Error Resume Next
    Err.Clear
    GetMailCategoriesSafe = SafeText(mail.Categories)
    Err.Clear
    On Error GoTo 0

End Function


' ============================================================
' OUTLOOK CONNECTION
' ============================================================
Private Function GetMailboxInbox(ByVal outlookNamespace As Object, _
                                 ByVal mailboxSmtp As String) As Object

    Dim recipient As Object
    Dim sharedInbox As Object
    Dim rootFolder As Object
    Dim rootName As String
    Dim mailboxKey As String

    mailboxKey = LCase$(Trim$(mailboxSmtp))

    ' Preferred path for an Exchange/shared mailbox: resolve the SMTP
    ' address through Outlook/Exchange and ask for that mailbox's Inbox.
    On Error Resume Next
    Err.Clear
    Set recipient = outlookNamespace.CreateRecipient(mailboxSmtp)

    If Not recipient Is Nothing Then
        recipient.Resolve
        If recipient.Resolved Then
            Set sharedInbox = outlookNamespace.GetSharedDefaultFolder( _
                                  recipient, OL_FOLDER_INBOX)
        End If
    End If
    Err.Clear
    On Error GoTo 0

    If Not sharedInbox Is Nothing Then
        If IsOutlookFolderUsable(sharedInbox) Then
            Set GetMailboxInbox = sharedInbox
            Exit Function
        Else
            Set sharedInbox = Nothing
        End If
    End If

    ' Fallback for a mailbox that is already mounted in Outlook but cannot
    ' be opened through GetSharedDefaultFolder. Match its root display name
    ' conservatively against the SMTP address/local-part.
    On Error Resume Next
    For Each rootFolder In outlookNamespace.Folders
        rootName = LCase$(Trim$(SafeText(rootFolder.Name)))

        If rootName = mailboxKey Or _
           InStr(1, rootName, "wko.cnlib.gsm", vbTextCompare) > 0 Then

            Set sharedInbox = Nothing
            Set sharedInbox = rootFolder.Folders("Inbox")

            If Not sharedInbox Is Nothing Then
                If IsOutlookFolderUsable(sharedInbox) Then
                    Set GetMailboxInbox = sharedInbox
                    Exit Function
                Else
                    Set sharedInbox = Nothing
                End If
            End If
        End If
    Next rootFolder
    Err.Clear
    On Error GoTo 0

    Set GetMailboxInbox = Nothing

End Function


Private Function IsOutlookFolderUsable(ByVal folderObject As Object) As Boolean

    Dim testItems As Object
    Dim testPath As String

    On Error Resume Next
    Err.Clear
    testPath = SafeText(folderObject.FolderPath)
    Set testItems = folderObject.Items

    If Err.Number = 0 And Len(testPath) > 0 Then
        If Not testItems Is Nothing Then
            IsOutlookFolderUsable = True
        End If
    End If

    Err.Clear
    Set testItems = Nothing
    On Error GoTo 0

End Function


Private Function GetClassicOutlook() As Object

    Dim outlookApp As Object

    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")

    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If

    On Error GoTo 0
    Set GetClassicOutlook = outlookApp

End Function


' ============================================================
' AUDIT LOG
' ============================================================
Private Sub PrepareAuditSheet(ByVal wb As Workbook)

    Dim auditWs As Worksheet

    On Error Resume Next
    Set auditWs = wb.Worksheets(AUDIT_SHEET_NAME)
    On Error GoTo 0

    If auditWs Is Nothing Then
        Set auditWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        auditWs.Name = AUDIT_SHEET_NAME

        auditWs.Cells(1, 1).Value = "Logged At"
        auditWs.Cells(1, 2).Value = "Email Received"
        auditWs.Cells(1, 3).Value = "Subject"
        auditWs.Cells(1, 4).Value = "GSM ID"
        auditWs.Cells(1, 5).Value = "Result"
        auditWs.Cells(1, 6).Value = "Detail"
        auditWs.Rows(1).Font.Bold = True
        auditWs.Columns(1).NumberFormat = "m/d/yyyy h:mm"
        auditWs.Columns(2).NumberFormat = "m/d/yyyy h:mm"
    End If

End Sub


Private Sub WriteAudit(ByVal wb As Workbook, _
                       ByVal receivedTime As Date, _
                       ByVal subjectText As String, _
                       ByVal gsmId As String, _
                       ByVal resultText As String, _
                       ByVal detailText As String)

    Dim auditWs As Worksheet
    Dim nextRow As Long

    On Error GoTo AuditExit

    Set auditWs = wb.Worksheets(AUDIT_SHEET_NAME)
    nextRow = auditWs.Cells(auditWs.Rows.Count, 1).End(xlUp).Row + 1

    auditWs.Cells(nextRow, 1).Value = Now
    auditWs.Cells(nextRow, 2).Value = receivedTime
    auditWs.Cells(nextRow, 3).Value = subjectText
    auditWs.Cells(nextRow, 4).Value = gsmId
    auditWs.Cells(nextRow, 5).Value = resultText
    auditWs.Cells(nextRow, 6).Value = detailText

AuditExit:
    Err.Clear

End Sub


Private Sub AuditExistingDuplicates(ByVal wb As Workbook, _
                                    ByVal duplicateGsm As Object)

    Dim key As Variant

    If duplicateGsm Is Nothing Then Exit Sub

    For Each key In duplicateGsm.Keys
        WriteAudit wb, Now, "", CStr(key), _
                   "DUPLICATE GSM IN EXCEL", _
                   "Rows: " & CStr(duplicateGsm(key))
    Next key

End Sub
