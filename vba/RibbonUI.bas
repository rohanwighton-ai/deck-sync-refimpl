Attribute VB_Name = "RibbonUI"
Option Explicit

' Action logic for specs/ribbon-ui.md's four buttons -- gather what the
' engine needs (via DeckRegistry lookups + WorkbookBridge), call an
' existing Sub, report the result. New Period's picker and Resolve
' Unmatched Fields' role picker are the only "new" pieces, and both are
' InputBox chains, not new sync/matching logic.
'
' Each action is a plain, parameterless Public Sub (SyncNow/NewPeriod/
' OnboardNewType/ResolveUnmatchedFields) rather than a ribbon-style
' `(control As IRibbonControl)` callback -- see CommandBarUI.bas's header
' comment for why: a real customUI14.xml ribbon turned out to be
' impossible to ship for a .ppam add-in (confirmed 2026-07-26 -- its loader
' rejects the package if it contains anything beyond its exact expected
' part set, ribbon or not), so CommandBarUI.bas's toolbar buttons drive
' these directly. Kept as zero-argument Subs (not folded into
' CommandBarUI.bas itself) so a future real ribbon-hosted add-in (the
' "Office.js kept open" option DECISIONS.md's 2026-07-25 entry already
' flags as a later move) can still wire a `(control As IRibbonControl)`
' wrapper straight back to these without touching this module's logic.
'
' Shared result reporting (ribbon-ui.md's "one shared result form... not a
' bespoke dialog per action") is deliberately NOT a UserForm here -- see
' OnboardFlow.bas's header comment for why (no proven .frm/.frx precedent
' yet). ShowSyncResult below is the shared *reporting* logic ribbon-ui.md
' asks for (one function, called from every action, not divergent MsgBox
' calls scattered per-button) -- just backed by MsgBox instead of a form.
' Upgrading the display mechanism later only touches this one function.

' ---------------------------------------------------------------------
' R13: the sync flow. Rules in specs/sync-flow-rules.md (F1-F10).
' ---------------------------------------------------------------------
'
' Three entry points, one write path:
'   Sync Now         -- confirms the uniform batches, applies them, hands the
'                       remainder to the review sheet
'   Review Changes   -- builds the sheet for everything, writes nothing
'   Apply Approved   -- writes what was ticked on the sheet
'
' What changed on 2026-07-31: `Sync Now` used to plan every type, show a
' count-based confirmation ("19 slide(s) corrected -- Proceed?") and write the
' lot. That answered "how many will change?" and never "is this change to this
' slide right?", which is exactly what R13 forbids -- and the first real run
' took that path and moved 19 slides with nobody seeing a before-and-after.
'
' It is not a count any more. It shows every transformation in full, and it
' cannot reach a slide with anything a human has not seen.
'
' Carried over unchanged, because R13 supersedes none of it: the unsaved-
' workbook refusal (ResolveSyncContext) and the R9 duplicate-key warning
' (WarnOnDuplicateKeys, called by BOTH SyncNowCore and ReviewChangesCore).
'
' An earlier version of this comment claimed R9 lived in ResolveSyncContext. It
' did not, and SyncNowCore's fast path wrote to the deck without ever reaching
' the only place it did live. A comment asserting a guard that is not there is
' worse than no comment: it stops the next reader checking. Found by review,
' 2026-07-31.

' Sync Now, restored 2026-07-31 and batch-aware.
'
' It was briefly removed earlier the same day, on my reasoning that its
' confirmation showed a count and R13 demands a before-and-after. Rohan
' corrected it: R13.2 makes a verified uniform batch ONE decision, so a whole
' change set that collapses into a few uniform transformations can be approved
' in a dialog that shows all of them -- that is a complete before-and-after, not
' a count. See ReviewQueue's fast-path header.
'
' A second correction from him followed: the run does NOT have to be
' all-or-nothing. Applying the uniform part and deferring the rest is fully
' compliant -- the batched changes are still shown before writing, and the
' individual ones still go to the worksheet. So (F5):
'
'   - uniform batches present, few enough to read
'       -> show them all, one confirmation, apply just those
'   - individual changes remain afterwards
'       -> rebuild, write the review sheet, open it (F6)
'   - nothing batchable, or too many batches
'       -> the worksheet handles the whole run
'
' Every path goes through the same queue, revalidation and backup as Apply
' Approved. This is a different REVIEW SURFACE for changes that fit one, not a
' different write path -- there is still exactly one place that writes a field
' to a slide (F7).
Public Sub SyncNow()
    On Error GoTo Failed
    SyncNowCore
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult "Sync Now", RibbonUI.UnexpectedErrorText("Sync Now", Err.Number, Err.Description, Err.Source)
End Sub

Private Sub SyncNowCore()
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim wb As Object
    Dim types() As String
    Dim lo As Long, hi As Long
    If Not ResolveSyncContext("Sync Now", pres, wb, types, lo, hi) Then Exit Sub

    ' R9 BEFORE the queue is built, because the fast path writes without ever
    ' reaching ReviewChangesCore. Two slides on one key are invisible to the
    ' planner and to the applier alike -- both index by key into a Dictionary,
    ' last write wins -- so one slide is silently corrected and the other left
    ' stale, and the batch dialog cannot show it either, because its entity list
    ' comes from that same collapsed index.
    If Not WarnOnDuplicateKeys("Sync Now", types, lo, hi) Then Exit Sub

    ' THE DECK DECLARES THE PERIOD AND THE SHEET IS READ AS THAT PERIOD.
    ' Rows accumulate per period now, so an unfiltered read takes one row per
    ' slide out of several and files the rest in a counter nobody looks at.
    Dim deckPeriod As String
    deckPeriod = DeckRegistry.GetDeckPeriod(pres)
    Dim problem As String
    Dim refusals As String

    ' Build every type's queue first, so the decision covers the whole deck
    ' rather than the first type only -- the same reason SyncNowCore used to
    ' plan all types before confirming.
    Dim combined As ReviewQueueSet
    combined.SlideType = "(all types)"
    combined.RunStamp = ReviewQueue.MakeRunStamp()

    Dim i As Long
    For i = lo To hi
        Dim templateSld As Object
        Dim wsName As String
        If DeckRegistry.LookupType(pres, types(i), templateSld, wsName) Then
            Dim sheet As Sheet
            sheet = ExcelOutput.ReadSheetForDeckPeriod( _
                WorkbookBridge.GetOrAddWorksheet(wb, wsName), deckPeriod, problem)

            If problem <> "" Then
                refusals = refusals & "  " & types(i) & " -- " & problem & vbCrLf & vbCrLf
            Else
                Dim q As ReviewQueueSet
                q = ReviewQueue.BuildQueue(sheet, types(i))
                Dim n As Long
                For n = 1 To q.Count
                    combined.Count = combined.Count + 1
                    ReDim Preserve combined.Items(1 To combined.Count)
                    combined.Items(combined.Count) = q.Items(n)
                    ' Labels are per-type; prefix so two types' batches cannot
                    ' collide into one apparent decision in the combined view.
                    If combined.Items(combined.Count).BatchLabel <> "" Then
                        combined.Items(combined.Count).BatchLabel = types(i) & ":" & q.Items(n).BatchLabel
                    End If
                Next n
            End If
        End If
    Next i

    ' THE WHOLE RUN STOPS, not just the offending type. Sync Now writes, and
    ' syncing the readable types while one was refused is a partial write
    ' reported as a success -- the failure mode this codebase keeps paying for.
    If refusals <> "" Then
        MsgBox "Sync Now stopped. The deck declares period '" & deckPeriod & _
            "' and at least one Data sheet cannot be read for it:" & vbCrLf & vbCrLf & refusals & _
            "Nothing has been written. Fix the sheet, or set the deck's period, then run again.", _
            vbExclamation, "Sync Now"
        Exit Sub
    End If

    If combined.Count = 0 Then
        MsgBox ReviewQueue.FastPathRefusalText(combined), vbInformation, "Sync Now"
        Exit Sub
    End If

    ' F3/F4/F5: the change set picks the surface, and a run may use both. Nothing
    ' batchable (or too many batches to read) means the worksheet handles the
    ' whole run; otherwise the uniform part is confirmed here and the remainder
    ' is handed on.
    Dim fullReport As String
    If Not ReviewQueue.HasBatchableWork(combined) Then
        MsgBox ReviewQueue.FastPathRefusalText(combined), vbExclamation, "Sync Now"
        ReviewChangesCore False
        Exit Sub
    End If

    ' A macro-enabled DECK cannot be saved on a managed machine, and the block
    ' is silent -- so the sync would appear to succeed and then quietly not
    ' persist. Said before the confirmation, where it can still change the
    ' answer. See WorkbookBridge.MacroEnabledWarning for the incident.
    Dim deckMacroWarn As String
    deckMacroWarn = WorkbookBridge.MacroEnabledWarning(pres.fullName)
    If deckMacroWarn <> "" Then deckMacroWarn = deckMacroWarn & vbCrLf & vbCrLf

    If MsgBox(deckMacroWarn & ReviewQueue.ConfirmBatchText(combined), vbYesNo + vbQuestion, "Sync Now") <> vbYes Then
        Exit Sub
    End If

    ' Approved -- run it through the ordinary machinery rather than a shortcut.
    ' The grid is still written and still consumed, so a fast-path run leaves
    ' exactly the same audit trail as a reviewed one (F7).
    Dim logWs As Object
    Set logWs = WorkbookBridge.GetOrAddWorksheet(wb, "Sync Log")

    For i = lo To hi
        Dim tmplSld As Object
        Dim dataWsName As String
        If DeckRegistry.LookupType(pres, types(i), tmplSld, dataWsName) Then
            Dim dataSheet As Sheet
            ' Read the same way the queue was built. The loop above already
            ' refused anything unreadable, so `problem` cannot be set here --
            ' but reading it unfiltered would apply a different set of rows
            ' than the human just approved, which is worse than refusing.
            dataSheet = ExcelOutput.ReadSheetForDeckPeriod( _
                WorkbookBridge.GetOrAddWorksheet(wb, dataWsName), deckPeriod, problem)

            Dim tq As ReviewQueueSet
            tq = ReviewQueue.BuildQueue(dataSheet, types(i))

            ' ONLY the batched rows. The individual ones stay unapproved here and
            ' are picked up by the rebuild below -- this is the whole of F5.
            ReviewQueue.ApproveBatchedOnly tq

            Dim reviewWs As Object
            Set reviewWs = WorkbookBridge.GetOrAddWorksheet(wb, ReviewQueue.ReviewSheetNameFor(types(i)))
            ReviewQueue.WriteQueueSheet reviewWs, tq

            ' APPLY AND REBUILD ARE COUPLED PER TYPE, inside the loop.
            '
            ' They used to be apply-per-type in the loop and rebuild-once after
            ' it. ApplyApproved marks the sheet CONSUMED, so a COM error on
            ' type 2 jumped to the outer handler and skipped the rebuild for
            ' type 1 as well -- leaving type 1's slides already written, its
            ' review sheet stamped Consumed, and its individual prose rows
            ' behind a refusal ("already been applied") with nothing telling the
            ' human that Review Changes needed re-running to resurface them.
            '
            ' F6 says the rebuild is not optional. That only holds if it cannot
            ' be skipped by something happening to a different type.
            fullReport = fullReport & ReviewQueue.ApplyApproved(dataSheet, types(i), reviewWs, logWs) & vbCrLf

            Dim afterQ As ReviewQueueSet
            afterQ = ReviewQueue.BuildQueue(dataSheet, types(i))
            ReviewQueue.WriteQueueSheet reviewWs, afterQ
            If afterQ.Count > 0 Then
                fullReport = fullReport & "  " & afterQ.Count & " change(s) still pending for " & _
                    types(i) & " -- see its review sheet." & vbCrLf
            End If
        End If
    Next i

    fullReport = fullReport & PersistBothFiles(pres, wb)
    ShowSyncResult "Sync Now", fullReport

    ' The per-type rebuild above has already refreshed every sheet, so the
    ' remaining job is only to put the human in front of one when there is
    ' something left to read.
    If ReviewQueue.IndividualCount(combined) > 0 Then
        MsgBox ReviewQueue.IndividualCount(combined) & " change(s) still need reading one at a time." & vbCrLf & _
               "Their review sheet(s) have been refreshed to match the deck as it is now.", _
               vbInformation, "Sync Now"
    End If
    Exit Sub

End Sub

' The resolution every sync-family action needs, in one place: paired workbook,
' registered types, and the unsaved-workbook refusal.
'
' Factored out while restoring Sync Now, because there were then four copies of
' it drifting apart. Returns False when the caller should simply stop -- it has
' already told the human why.
Private Function ResolveSyncContext(title As String, pres As Object, ByRef wb As Object, _
                                    ByRef types() As String, ByRef lo As Long, ByRef hi As Long) As Boolean
    Dim workbookPath As String
    workbookPath = DeckRegistry.GetWorkbookPath(pres)
    If workbookPath = "" Then
        MsgBox "This deck has no paired workbook yet -- use 'Onboard New Slide Type' first.", vbExclamation, title
        Exit Function
    End If

    types = DeckRegistry.ListRegisteredTypes(pres)

    Dim hasTypes As Boolean
    On Error Resume Next
    lo = LBound(types): hi = UBound(types)
    hasTypes = (Err.Number = 0)
    On Error GoTo 0

    If Not hasTypes Then
        MsgBox "This deck has no registered slide types yet -- use 'Onboard New Slide Type' first.", vbExclamation, title
        Exit Function
    End If

    Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath)
    If wb Is Nothing Then
        MsgBox "Could not open the paired workbook at: " & workbookPath, vbCritical, title
        Exit Function
    End If

    If WorkbookBridge.IsDirty(wb) Then
        If MsgBox(WorkbookBridge.UnsavedWorkbookText(workbookPath), _
                  vbYesNo + vbExclamation, title) <> vbYes Then
            Exit Function
        End If
        wb.Save
    End If

    ResolveSyncContext = True
End Function

' R9, in ONE place, called by every path that can reach a write.
'
' Returns False when the human declined -- callers stop. Extracted after review
' found Sync Now's fast path had no R9 check at all while a comment said it did:
' two copies of a guard drift, and one copy plus a comment is not a guard.
'
' Warns rather than refuses, unchanged: a duplicate key is a data-entry mistake
' in the deck, not corruption, and refusing outright would block a quarter's
' reporting over a fixable typo. The default is No.
Private Function WarnOnDuplicateKeys(title As String, types() As String, lo As Long, hi As Long) As Boolean
    Dim dupType As Long
    For dupType = lo To hi
        Dim dupReport As DuplicateKeyReport
        dupReport = IdentityCheck.FindDuplicateKeys(types(dupType))
        If dupReport.HasDuplicates Then
            If MsgBox(IdentityCheck.DuplicateKeyWarningText(types(dupType), dupReport) & _
                      vbCrLf & vbCrLf & "Continue anyway?", _
                      vbYesNo + vbExclamation + vbDefaultButton2, title) <> vbYes Then
                Exit Function
            End If
        End If
    Next dupType
    WarnOnDuplicateKeys = True
End Function

' Toolbar entry point. The real work is in ReviewChangesCore; this exists only
' to catch anything that escapes it.
'
' A WRAPPER rather than an inline "On Error GoTo" on purpose. In VBA,
' "On Error GoTo 0" disables the enabled handler for the whole procedure, and
' these bodies are full of "On Error Resume Next / On Error GoTo 0" pairs -- an
' inline handler would be switched off by the first of them and read as
' protection while providing none. Putting the handler in a separate frame
' means nothing inside the body can turn it off, now or after a later edit.
Public Sub ReviewChanges()
    On Error GoTo Failed
    ReviewChangesCore False
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult "Review Changes", RibbonUI.UnexpectedErrorText("Review Changes", Err.Number, Err.Description, Err.Source)
End Sub

' The loosened setting (Round 13 §0.1). Builds the identical queue, then ticks
' every row.
'
' A SEPARATE BUTTON rather than a checkbox on the one above, deliberately. The
' RM's ruling permits wholesale approval only while the work runs on a carved
' copy, and the risk R13.2 names is that bulk approval "teaches the operator to
' click through". A distinct button someone presses by name keeps that a
' decision taken each time and visible in the report -- and makes tightening the
' deletion of one procedure rather than the unpicking of a flag.
Public Sub ReviewChangesApproveAll()
    On Error GoTo Failed
    ReviewChangesCore True
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult "Review Changes (approve all)", RibbonUI.UnexpectedErrorText("Review Changes (approve all)", Err.Number, Err.Description, Err.Source)
End Sub

Private Sub ReviewChangesCore(approveAll As Boolean)
    Dim title As String
    title = IIf(approveAll, "Review Changes (approve all)", "Review Changes")

    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim workbookPath As String
    workbookPath = DeckRegistry.GetWorkbookPath(pres)
    If workbookPath = "" Then
        MsgBox "This deck has no paired workbook yet -- use 'Onboard New Slide Type' first.", vbExclamation, title
        Exit Sub
    End If

    Dim types() As String
    types = DeckRegistry.ListRegisteredTypes(pres)

    Dim lo As Long, hi As Long, hasTypes As Boolean
    On Error Resume Next
    lo = LBound(types): hi = UBound(types)
    hasTypes = (Err.Number = 0)
    On Error GoTo 0

    If Not hasTypes Then
        MsgBox "This deck has no registered slide types yet -- use 'Onboard New Slide Type' first.", vbExclamation, title
        Exit Sub
    End If

    Dim wb As Object
    Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath)
    If wb Is Nothing Then
        MsgBox "Could not open the paired workbook at: " & workbookPath, vbCritical, title
        Exit Sub
    End If

    ' Refuse to build a queue out of Excel's unsaved buffer -- see
    ' WorkbookBridge.IsDirty for the live incident. Checked BEFORE planning, not
    ' after: the plan reads the sheet, so a queue built on unsaved data would
    ' show a human before-and-afters whose "after" exists in no file, and they
    ' would be approving values that could still change before Apply runs.
    If WorkbookBridge.IsDirty(wb) Then
        If MsgBox(WorkbookBridge.UnsavedWorkbookText(workbookPath), _
                  vbYesNo + vbExclamation, title) <> vbYes Then
            Exit Sub
        End If
        wb.Save
    End If

    ' R9: duplicate identity tags, checked BEFORE planning. Kept from
    ' SyncNowCore unchanged -- the planner cannot report this usefully, because
    ' to PlanRoutineSync two slides sharing a key is indistinguishable from one
    ' matched slide and one unmatched one. It is only visible across instances.
    If Not WarnOnDuplicateKeys(title, types, lo, hi) Then Exit Sub

    Dim fullReport As String
    Dim totalQueued As Long
    Dim firstSheet As Object

    Dim i As Long
    For i = lo To hi
        Dim templateSld As Object
        Dim wsName As String
        If DeckRegistry.LookupType(pres, types(i), templateSld, wsName) Then
            Dim ws As Object
            Set ws = WorkbookBridge.GetOrAddWorksheet(wb, wsName)

            Dim sheet As Sheet
            Dim problem As String
            sheet = ExcelOutput.ReadSheetForDeckPeriod(ws, DeckRegistry.GetDeckPeriod(pres), problem)

            If problem <> "" Then
                ' Reported and skipped rather than stopping the run: this builds
                ' a review sheet and writes nothing to a slide, so the other
                ' types' queues are still worth having. The refusal is the point
                ' -- an empty queue for this type would otherwise read as
                ' "nothing to change".
                fullReport = fullReport & "=== " & types(i) & " ===" & vbCrLf & _
                    "REFUSED at period '" & DeckRegistry.GetDeckPeriod(pres) & "': " & problem & vbCrLf & vbCrLf
            Else
                Dim q As ReviewQueueSet
                q = ReviewQueue.BuildQueue(sheet, types(i))
                totalQueued = totalQueued + q.Count

                Dim reviewWs As Object
                Set reviewWs = WorkbookBridge.GetOrAddWorksheet(wb, ReviewQueue.ReviewSheetNameFor(types(i)))
                ReviewQueue.WriteQueueSheet reviewWs, q
                If approveAll Then ReviewQueue.ApproveAllInSheet reviewWs
                If firstSheet Is Nothing Then Set firstSheet = reviewWs

                fullReport = fullReport & "=== " & types(i) & " ===" & vbCrLf & _
                    ReviewQueue.QueueSummaryText(q) & vbCrLf
            End If
        Else
            fullReport = fullReport & "SKIPPED " & types(i) & ": registered type's template slide no longer resolves (was it deleted?)" & vbCrLf
        End If
    Next i

    If approveAll And totalQueued > 0 Then
        fullReport = "APPROVE-ALL: every queued change has been ticked without" & vbCrLf & _
            "individual review. Permitted on a scratch copy only." & vbCrLf & vbCrLf & fullReport
    End If

    ' Bring the review sheet to the front. Leaving the human to go and find it
    ' is how a review becomes optional in practice while remaining mandatory on
    ' paper -- the same distinction R13 is about.
    If Not firstSheet Is Nothing And totalQueued > 0 Then
        On Error Resume Next
        firstSheet.Activate
        wb.Activate
        On Error GoTo 0
    End If

    ShowSyncResult title & " (nothing written)", fullReport
End Sub

' Toolbar entry point. The real work is in ApplyApprovedCore; this exists only
' to catch anything that escapes it. Same separate-frame reasoning as above.
Public Sub ApplyApprovedChanges()
    On Error GoTo Failed
    ApplyApprovedCore
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult "Apply Approved", RibbonUI.UnexpectedErrorText("Apply Approved", Err.Number, Err.Description, Err.Source)
End Sub

' The only path in this add-in that writes a field value to a slide as part of a
' sync. Everything it writes was ticked by a human and revalidated against the
' live slide immediately before the write.
Private Sub ApplyApprovedCore()
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim workbookPath As String
    workbookPath = DeckRegistry.GetWorkbookPath(pres)
    If workbookPath = "" Then
        MsgBox "This deck has no paired workbook yet -- nothing to apply.", vbExclamation, "Apply Approved"
        Exit Sub
    End If

    Dim types() As String
    types = DeckRegistry.ListRegisteredTypes(pres)

    Dim lo As Long, hi As Long, hasTypes As Boolean
    On Error Resume Next
    lo = LBound(types): hi = UBound(types)
    hasTypes = (Err.Number = 0)
    On Error GoTo 0

    If Not hasTypes Then
        MsgBox "This deck has no registered slide types yet -- nothing to apply.", vbExclamation, "Apply Approved"
        Exit Sub
    End If

    Dim wb As Object
    Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath)
    If wb Is Nothing Then
        MsgBox "Could not open the paired workbook at: " & workbookPath, vbCritical, "Apply Approved"
        Exit Sub
    End If

    ' The ticks live in the workbook, so an unsaved workbook means the
    ' approvals being read are on screen and not in any file. Same refusal as
    ' the review step, for the same reason.
    If WorkbookBridge.IsDirty(wb) Then
        If MsgBox(WorkbookBridge.UnsavedWorkbookText(workbookPath), _
                  vbYesNo + vbExclamation, "Apply Approved") <> vbYes Then
            Exit Sub
        End If
        wb.Save
    End If

    Dim logWs As Object
    Set logWs = WorkbookBridge.GetOrAddWorksheet(wb, "Sync Log")

    Dim fullReport As String
    Dim i As Long
    For i = lo To hi
        Dim templateSld As Object
        Dim wsName As String
        If DeckRegistry.LookupType(pres, types(i), templateSld, wsName) Then
            Dim reviewName As String
            reviewName = ReviewQueue.ReviewSheetNameFor(types(i))

            If Not WorkbookBridge.WorksheetExists(wb, reviewName) Then
                fullReport = fullReport & "=== " & types(i) & " ===" & vbCrLf & _
                    "No review has been built for this type. Run 'Review Changes' first." & vbCrLf & vbCrLf
            Else
                Dim ws As Object
                Set ws = WorkbookBridge.GetOrAddWorksheet(wb, wsName)

                Dim sheet As Sheet
                Dim problem As String
                sheet = ExcelOutput.ReadSheetForDeckPeriod(ws, DeckRegistry.GetDeckPeriod(pres), problem)

                If problem <> "" Then
                    ' This one WRITES TO SLIDES. The review sheet a human ticked
                    ' was built from a period-filtered read; applying a sheet
                    ' that can no longer be read that way would write a
                    ' different set of rows than the one they approved.
                    fullReport = fullReport & "=== " & types(i) & " ===" & vbCrLf & _
                        "REFUSED at period '" & DeckRegistry.GetDeckPeriod(pres) & "': " & problem & vbCrLf & _
                        "Nothing was written for this type." & vbCrLf & vbCrLf
                Else
                    Dim reviewWs As Object
                    Set reviewWs = WorkbookBridge.GetOrAddWorksheet(wb, reviewName)

                    fullReport = fullReport & _
                        ReviewQueue.ApplyApproved(sheet, types(i), reviewWs, logWs) & vbCrLf
                End If
            End If
        Else
            fullReport = fullReport & "SKIPPED " & types(i) & ": registered type's template slide no longer resolves (was it deleted?)" & vbCrLf
        End If
    Next i

    fullReport = fullReport & PersistBothFiles(pres, wb)
    ShowSyncResult "Apply Approved", fullReport
End Sub

' Toolbar entry point. The real work is in SyncPreviewCore; this exists only to
' catch anything that escapes it.
'
' A WRAPPER rather than an inline "On Error GoTo" on purpose. In VBA,
' "On Error GoTo 0" disables the enabled handler for the whole procedure, and
' these bodies are full of "On Error Resume Next / On Error GoTo 0" pairs -- an
' inline handler would be switched off by the first of them and read as
' protection while providing none. Putting the handler in a separate frame
' means nothing inside the body can turn it off, now or after a later edit.
Public Sub SyncPreview()
    On Error GoTo Failed
    SyncPreviewCore
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult "Preview Sync", RibbonUI.UnexpectedErrorText("Preview Sync", Err.Number, Err.Description, Err.Source)
End Sub

' Read-only twin of SyncNow: identical resolution path (same registry lookups,
' same workbook, same worksheets), but every registered type is run through
' RunSync.PreviewRoutineSync instead of RunRoutineSync, so the deck is never
' touched. Deliberately shares SyncNow's structure line for line -- a preview
' that resolves its inputs differently from the real thing can disagree with it
' about what would happen, which defeats the point.
Private Sub SyncPreviewCore()
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim workbookPath As String
    workbookPath = DeckRegistry.GetWorkbookPath(pres)
    If workbookPath = "" Then
        MsgBox "This deck has no paired workbook yet -- nothing to preview.", vbExclamation, "Preview Sync"
        Exit Sub
    End If

    Dim types() As String
    types = DeckRegistry.ListRegisteredTypes(pres)

    Dim lo As Long, hi As Long, hasTypes As Boolean
    On Error Resume Next
    lo = LBound(types): hi = UBound(types)
    hasTypes = (Err.Number = 0)
    On Error GoTo 0

    If Not hasTypes Then
        MsgBox "This deck has no registered slide types yet -- nothing to preview.", vbExclamation, "Preview Sync"
        Exit Sub
    End If

    Dim wb As Object
    Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath)
    If wb Is Nothing Then
        MsgBox "Could not open the paired workbook at: " & workbookPath, vbCritical, "Preview Sync"
        Exit Sub
    End If

    Dim fullReport As String
    Dim i As Long
    For i = lo To hi
        Dim templateSld As Object
        Dim wsName As String
        If DeckRegistry.LookupType(pres, types(i), templateSld, wsName) Then
            Dim ws As Object
            Set ws = WorkbookBridge.GetOrAddWorksheet(wb, wsName)
            fullReport = fullReport & RunSync.PreviewRoutineSync(ws, types(i)) & vbCrLf
        Else
            fullReport = fullReport & "SKIPPED " & types(i) & ": registered type's template slide no longer resolves (was it deleted?)" & vbCrLf
        End If
    Next i

    ' Warns where Sync Now refuses. The preview writes nothing, so unsaved data
    ' cannot damage the deck here -- but a preview of values that exist in no
    ' file is still a preview of something that might never be synced, and the
    ' whole worth of this report is that it can be trusted. Stated at the TOP:
    ' a caveat below a long report is a caveat nobody reads.
    If WorkbookBridge.IsDirty(wb) Then
        fullReport = "NOTE: the Data workbook has unsaved changes, so this preview " & _
            "reflects what is on screen in Excel, not what is in the file." & vbCrLf & _
            "Save it before syncing." & vbCrLf & vbCrLf & fullReport
    End If

    ' THE DETAIL GOES ON A SHEET; THE DIALOG GETS THE HEADLINE.
    '
    ' 2026-08-08: this modal ended mid-word at "would c". MsgBox truncates its
    ' prompt near 1024 characters and says nothing about it, so a preview of 27
    ' changes showed an unknown fraction of them -- and a preview you approve
    ' from, that is silently incomplete, is worse than no preview at all.
    WorkbookBridge.WriteRunLog wb, "Preview Sync -- nothing was written", fullReport

    Dim shortReport As String
    shortReport = "PREVIEW ONLY -- nothing was written to any slide." & vbCrLf & vbCrLf & _
        CountLines(fullReport, "would correct:") & " slide(s) would change." & vbCrLf & vbCrLf & _
        "The full before-and-after is on the '" & WorkbookBridge.RUN_LOG_SHEET_NAME & _
        "' sheet in the workbook, untruncated." & vbCrLf & vbCrLf & _
        "Read it there, then run Sync Now."

    If WorkbookBridge.IsDirty(wb) Then
        shortReport = "NOTE: the Data workbook has unsaved changes, so this preview " & _
            "reflects what is on screen in Excel, not what is in the file." & vbCrLf & vbCrLf & _
            shortReport
    End If

    ShowSyncResult "Preview Sync (nothing written)", shortReport
End Sub

' How many times a marker appears in a report. Used for the preview headline
' rather than a counter threaded through the loop above: the report is the
' thing being summarised, so counting IT cannot drift away from what it says.
Private Function CountLines(text As String, marker As String) As Long
    If Len(marker) = 0 Then Exit Function
    Dim pos As Long
    pos = InStr(1, text, marker, vbTextCompare)
    Do While pos > 0
        CountLines = CountLines + 1
        pos = InStr(pos + Len(marker), text, marker, vbTextCompare)
    Loop
End Function

' ---------------------------------------------------------------------
' Audit Fields -- "what on this slide is the tool not tracking?"
' See TemplateAudit.bas for the reasoning. Read-only; the only write is
' the audit worksheet.
' ---------------------------------------------------------------------

Public Sub AuditFields()
    On Error GoTo Failed
    AuditFieldsCore
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult "Audit Fields", RibbonUI.UnexpectedErrorText("Audit Fields", Err.Number, Err.Description, Err.Source)
End Sub

' Picks the subject slide by preference, and never requires a template.
'
' The fallback chain is the whole point (Rohan, 2026-07-30): this operation and
' field marking must each work whatever the other has or hasn't done, because
' decks arrive at different maturities. A deck that has never run Create
' Template Slide still needs to know which fields it is missing -- arguably
' more than a mature one does, since knowing the fields is what makes a
' template worth building. So:
'   1. the type's master template, if step 1 has been run
'   2. otherwise the registered slide (pre-step-1 decks: a real project slide)
'   3. otherwise the first instance of the type
' Each is a legitimate subject; only the ranking differs.
Private Sub AuditFieldsCore()
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim types() As String
    types = DeckRegistry.ListRegisteredTypes(pres)
    Dim lo As Long, hi As Long, hasTypes As Boolean
    On Error Resume Next
    lo = LBound(types): hi = UBound(types)
    hasTypes = (Err.Number = 0)
    On Error GoTo 0

    If Not hasTypes Then
        MsgBox "This deck has no registered slide types yet -- use 'Onboard New Slide Type' first.", vbExclamation, "Audit Fields"
        Exit Sub
    End If

    Dim slideType As String
    slideType = InputBox(BuildTypePickerPrompt(types), "Audit Fields -- Choose Type")
    slideType = ResolveTypeAnswer(slideType, types)
    If slideType = "" Then Exit Sub

    Dim instances() As Object
    instances = RunSync.GatherInstances(slideType)

    Dim subjectSld As Object
    Dim subjectLabel As String

    Set subjectSld = TemplateSlide.FindTemplateFor(slideType)
    If Not subjectSld Is Nothing Then
        subjectLabel = "the master template (slide " & subjectSld.SlideIndex & ")"
    End If

    Dim wsName As String
    If subjectSld Is Nothing Then
        Dim registeredSld As Object
        If DeckRegistry.LookupType(pres, slideType, registeredSld, wsName) Then
            Set subjectSld = registeredSld
            subjectLabel = "slide " & subjectSld.SlideIndex & " (no master template yet -- this type's registered slide)"
        End If
    Else
        Dim ignoredSld As Object
        DeckRegistry.LookupType pres, slideType, ignoredSld, wsName
    End If

    If subjectSld Is Nothing Then
        Dim iLo As Long, iHi As Long, hasInstances As Boolean
        On Error Resume Next
        iLo = LBound(instances): iHi = UBound(instances)
        hasInstances = (Err.Number = 0)
        On Error GoTo 0
        If hasInstances Then
            Set subjectSld = instances(iLo)
            subjectLabel = "slide " & subjectSld.SlideIndex & " (first slide of this type)"
        End If
    End If

    If subjectSld Is Nothing Then
        MsgBox "Nothing to audit -- type '" & slideType & "' has no slides in this deck.", vbExclamation, "Audit Fields"
        Exit Sub
    End If

    ' The subject is excluded from its own comparison set. Without this, a
    ' pre-step-1 deck audits a real project slide against a list that includes
    ' that same slide, so every text scores at least 1 and nothing can ever
    ' read "on no other slide" -- the verdict that carries all the signal.
    Dim comparisons() As Object
    comparisons = ExcludeSlide(instances, subjectSld)

    Dim rowCount As Long
    Dim trackedFields As String
    Dim rows() As AuditRow
    rows = TemplateAudit.BuildAudit(subjectSld, comparisons, rowCount, trackedFields)

    Dim trackedCount As Long
    trackedCount = 0
    If trackedFields <> "" Then trackedCount = UBound(Split(trackedFields, "|")) + 1

    Dim likelyDataCount As Long
    Dim cLo As Long, cHi As Long, hasComparisons As Boolean
    On Error Resume Next
    cLo = LBound(comparisons): cHi = UBound(comparisons)
    hasComparisons = (Err.Number = 0)
    On Error GoTo 0
    Dim comparisonCount As Long
    If hasComparisons Then comparisonCount = cHi - cLo + 1

    ' Asks TemplateAudit rather than re-implementing the prefix match. The two
    ' were separate hand-written comparisons until 2026-07-30 and only this one
    ' was right -- the other had the literal's length wrong by one and silently
    ' mis-sorted the whole grid.
    Dim i As Long
    For i = 1 To rowCount
        If TemplateAudit.IsLikelyProjectData(rows(i).Verdict) Then likelyDataCount = likelyDataCount + 1
    Next i

    ' Write the grid, if there is a workbook to write it to. A deck with no
    ' paired workbook still gets the counts -- the audit reads the DECK, so
    ' refusing outright would withhold an answer it already has.
    Dim wroteGrid As Boolean
    Dim workbookPath As String
    workbookPath = DeckRegistry.GetWorkbookPath(pres)
    If workbookPath <> "" And rowCount > 0 Then
        Dim wb As Object
        Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath)
        If Not wb Is Nothing Then
            Dim ws As Object
            Set ws = WorkbookBridge.GetOrAddWorksheet(wb, TemplateAudit.AUDIT_SHEET_NAME)
            TemplateAudit.WriteAuditGrid ws, rows, rowCount
            wroteGrid = True
        End If
    End If

    Dim report As String
    report = TemplateAudit.SummaryText(slideType, subjectLabel, trackedCount, rowCount, likelyDataCount, comparisonCount)
    If rowCount > 0 And Not wroteGrid Then
        report = report & vbCrLf & vbCrLf & "COULD NOT WRITE THE LIST: no paired workbook was reachable, so only the counts above are available."
    End If

    ShowSyncResult "Audit Fields", report
End Sub

' `slides` minus `dropSld`, matched by SlideID rather than object identity --
' same reasoning as AdoptFlow.ExcludeTemplateSlide, whose shape this mirrors
' (two references to one slide are not guaranteed to compare equal).
Public Function ExcludeSlide(slides() As Object, dropSld As Object) As Object()
    Dim result() As Object
    Dim n As Long
    n = 0

    Dim lo As Long, hi As Long, hasAny As Boolean
    On Error Resume Next
    lo = LBound(slides): hi = UBound(slides)
    hasAny = (Err.Number = 0)
    On Error GoTo 0
    If Not hasAny Then
        ExcludeSlide = result
        Exit Function
    End If

    Dim i As Long
    For i = lo To hi
        If slides(i).SlideID <> dropSld.SlideID Then
            n = n + 1
            ReDim Preserve result(1 To n)
            Set result(n) = slides(i)
        End If
    Next i

    ExcludeSlide = result
End Function

' ---------------------------------------------------------------------
' Create Template Slide -- specs/deck-compiler-concept.md progression
' step 1. Gives a type a master template slide that is never a real
' project, and re-points the type's registration at it, so new records
' stop being cloned from whichever real slide happened to be onboarded
' first. See TemplateSlide.bas's header for the hazard.
' ---------------------------------------------------------------------

' Toolbar entry point. The real work is in CreateTemplateSlideCore; this
' exists only to catch anything that escapes it.
'
' A WRAPPER rather than an inline "On Error GoTo" on purpose. In VBA,
' "On Error GoTo 0" disables the enabled handler for the whole procedure, and
' these bodies are full of "On Error Resume Next / On Error GoTo 0" pairs -- an
' inline handler would be switched off by the first of them and read as
' protection while providing none. Putting the handler in a separate frame
' means nothing inside the body can turn it off, now or after a later edit.
Public Sub CreateTemplateSlide()
    On Error GoTo Failed
    CreateTemplateSlideCore
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult "Create Template Slide", RibbonUI.UnexpectedErrorText("Create Template Slide", Err.Number, Err.Description, Err.Source)
End Sub

Private Sub CreateTemplateSlideCore()
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim types() As String
    types = DeckRegistry.ListRegisteredTypes(pres)
    Dim lo As Long, hi As Long, hasTypes As Boolean
    On Error Resume Next
    lo = LBound(types): hi = UBound(types)
    hasTypes = (Err.Number = 0)
    On Error GoTo 0

    If Not hasTypes Then
        MsgBox "This deck has no registered slide types yet -- use 'Onboard New Slide Type' first.", vbExclamation, "Create Template Slide"
        Exit Sub
    End If

    Dim slideType As String
    slideType = InputBox(BuildTypePickerPrompt(types), "Create Template Slide -- Choose Type")
    slideType = ResolveTypeAnswer(slideType, types)
    If slideType = "" Then Exit Sub

    ' Already has one: stop here rather than at MakeTemplateFrom's own guard,
    ' so the message can name the existing template's slide number. Both
    ' checks stay -- this one is for the human, that one is the invariant.
    Dim existing As Object
    Set existing = TemplateSlide.FindTemplateFor(slideType)
    If Not existing Is Nothing Then
        MsgBox "Type '" & slideType & "' already has a master template: slide " & existing.SlideIndex & "." & vbCrLf & vbCrLf & _
               "A type must have exactly one. Nothing was changed.", vbInformation, "Create Template Slide"
        Exit Sub
    End If

    Dim sourceSld As Object
    Dim wsName As String
    If Not DeckRegistry.LookupType(pres, slideType, sourceSld, wsName) Then
        MsgBox "Type '" & slideType & "' is registered but its slide no longer resolves (was it deleted?)." & vbCrLf & _
               "Re-onboard the type before creating its template.", vbExclamation, "Create Template Slide"
        Exit Sub
    End If

    ' Label the source by its instance key where it has one, falling back to
    ' the slide number -- the key is what the human recognises from the Data
    ' sheet, and "slide 3" is meaningless once the deck is reordered.
    Dim sourceInstance As SlideInstance
    sourceInstance = Resolve.ResolveSlideInstance(sourceSld)
    Dim sourceLabel As String
    sourceLabel = "slide " & sourceSld.SlideIndex
    If sourceInstance.HasInstanceKey Then sourceLabel = sourceInstance.InstanceKey & " (slide " & sourceSld.SlideIndex & ")"

    Dim sourceRoles() As String
    Dim sourceShapes() As Candidate
    sourceShapes = Onboarding.BuildTemplateFieldShapes(sourceSld, sourceRoles)
    Dim fLo As Long, fHi As Long, hasFields As Boolean
    On Error Resume Next
    fLo = LBound(sourceRoles): fHi = UBound(sourceRoles)
    hasFields = (Err.Number = 0)
    On Error GoTo 0
    Dim fieldCount As Long
    If hasFields Then fieldCount = fHi - fLo + 1

    If MsgBox(TemplateSlide.ConfirmTemplateText(slideType, sourceLabel, fieldCount), _
              vbYesNo + vbQuestion, "Create Template Slide") <> vbYes Then
        Exit Sub
    End If

    Dim mr As MakeTemplateResult
    mr = TemplateSlide.MakeTemplateFrom(sourceSld, slideType)

    Dim report As String
    If Not mr.Ok Then
        report = "FAILED to create a template for '" & slideType & "': " & mr.Reason
        ShowSyncResult "Create Template Slide", report
        Exit Sub
    End If

    ' Registration is the step that actually changes behaviour -- without it
    ' the template exists but nothing clones it, which is the quietest
    ' possible half-finished state. Done here rather than inside
    ' MakeTemplateFrom so that function stays testable with no registry.
    DeckRegistry.RegisterType pres, slideType, mr.NewSlide, wsName

    report = "Master template created for '" & slideType & "'." & vbCrLf & vbCrLf & _
        "    slide " & mr.NewSlide.SlideIndex & ", hidden from the slideshow" & vbCrLf & _
        "    " & mr.FieldCount & " field(s) set to placeholders" & vbCrLf & _
        "    new records will now be cloned from it, not from " & sourceLabel & vbCrLf & vbCrLf & _
        "It will not appear in Preview Sync or Sync Now reports -- a template" & vbCrLf & _
        "is not a record, so it is neither counted nor corrected." & vbCrLf & vbCrLf & _
        "Worth doing now: open it and clear anything the sync does not manage" & vbCrLf & _
        "(figures, chart data, notes, untagged text) that belonged to " & sourceLabel & "."
    ShowSyncResult "Create Template Slide", report
End Sub

Public Function BuildTypePickerPrompt(types() As String) As String
    Dim s As String
    s = "Choose a slide type (enter the number or the name):" & vbCrLf
    Dim lo As Long, hi As Long, i As Long
    lo = LBound(types): hi = UBound(types)
    For i = lo To hi
        s = s & i & ") " & types(i) & vbCrLf
    Next i
    BuildTypePickerPrompt = s
End Function

' Same number-or-name convention ResolveFields.PickRoleFromList already
' established -- kept as a separate function here (rather than a shared
' generic picker) since the two operate on different array element types
' with no natural common signature in VBA without a Variant-typed generic,
' which this project avoids per its own established style.
Public Function ResolveTypeAnswer(answer As String, types() As String) As String
    If Trim(answer) = "" Then
        ResolveTypeAnswer = ""
        Exit Function
    End If

    Dim lo As Long, hi As Long
    lo = LBound(types): hi = UBound(types)

    If IsNumeric(answer) Then
        Dim idx As Long
        idx = CLng(answer)
        If idx >= lo And idx <= hi Then
            ResolveTypeAnswer = types(idx)
            Exit Function
        End If
    End If

    Dim i As Long
    For i = lo To hi
        If LCase(types(i)) = LCase(Trim(answer)) Then
            ResolveTypeAnswer = types(i)
            Exit Function
        End If
    Next i

    ResolveTypeAnswer = ""
End Function

' ---------------------------------------------------------------------
' Onboard New Slide Type / Resolve Unmatched Fields -- thin wrappers over
' OnboardFlow.bas / ResolveFields.bas, the only new pieces being the
' DeckRegistry lookup Resolve Unmatched Fields needs to find its template
' (ribbon-ui.md's spec text assumed a caller would supply templateSld;
' DeckRegistry is that caller now).
' ---------------------------------------------------------------------

' Toolbar entry point. The real work is in OnboardNewTypeCore; this exists only to
' catch anything that escapes it.
'
' A WRAPPER rather than an inline "On Error GoTo" on purpose. In VBA,
' "On Error GoTo 0" disables the enabled handler for the whole procedure, and
' these bodies are full of "On Error Resume Next / On Error GoTo 0" pairs -- an
' inline handler would be switched off by the first of them and read as
' protection while providing none. Putting the handler in a separate frame
' means nothing inside the body can turn it off, now or after a later edit.
Public Sub OnboardNewType()
    On Error GoTo Failed
    OnboardNewTypeCore
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult "Onboard New Slide Type", RibbonUI.UnexpectedErrorText("Onboard New Slide Type", Err.Number, Err.Description, Err.Source)
End Sub

Private Sub OnboardNewTypeCore()
    Dim report As String
    report = OnboardFlow.PromptOnboardNewSlideType()
    If report <> "" Then
        ShowSyncResult "Onboard New Slide Type", report
    End If
End Sub

' Toolbar entry point. The real work is in ResolveUnmatchedFieldsCore; this exists only to
' catch anything that escapes it.
'
' A WRAPPER rather than an inline "On Error GoTo" on purpose. In VBA,
' "On Error GoTo 0" disables the enabled handler for the whole procedure, and
' these bodies are full of "On Error Resume Next / On Error GoTo 0" pairs -- an
' inline handler would be switched off by the first of them and read as
' protection while providing none. Putting the handler in a separate frame
' means nothing inside the body can turn it off, now or after a later edit.
Public Sub ResolveUnmatchedFields()
    On Error GoTo Failed
    ResolveUnmatchedFieldsCore
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult "Resolve Unmatched Fields", RibbonUI.UnexpectedErrorText("Resolve Unmatched Fields", Err.Number, Err.Description, Err.Source)
End Sub

Private Sub ResolveUnmatchedFieldsCore()
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim sel As Object
    Set sel = Application.ActiveWindow.Selection
    If sel.Type <> ppSelectionShapes Or sel.ShapeRange.count <> 1 Then
        MsgBox "Select exactly one shape on the slide first.", vbExclamation, "Resolve Unmatched Fields"
        Exit Sub
    End If

    ' .Parent is the containing Slide only for a top-level shape -- a shape
    ' selected from inside a group would resolve to the GroupShape instead.
    ' Not handled here: field shapes are expected to be top-level per this
    ' project's existing discovery convention (Discovery.bas recurses into
    ' groups to find candidates, but a human directly clicking one they
    ' want to resolve is the common case this flow targets); flagged as a
    ' known gap rather than a silently wrong assumption.
    Dim sld As Object
    Set sld = sel.ShapeRange(1).Parent

    Dim instance As SlideInstance
    instance = Resolve.ResolveSlideInstance(sld)
    If Not instance.HasTypeTag Then
        MsgBox "This slide has no slide type tag -- Resolve Unmatched Fields only applies to a slide already matched to a type.", vbExclamation, "Resolve Unmatched Fields"
        Exit Sub
    End If

    Dim templateSld As Object
    Dim wsName As String
    If Not DeckRegistry.LookupType(pres, instance.TypeTag, templateSld, wsName) Then
        MsgBox "Could not find a registered template for type '" & instance.TypeTag & "'.", vbExclamation, "Resolve Unmatched Fields"
        Exit Sub
    End If

    Dim result As String
    result = ResolveFields.PromptResolveUnmatchedField(templateSld)
    MsgBox result, vbInformation, "Resolve Unmatched Fields"
End Sub

' ---------------------------------------------------------------------
' Shared result reporting -- ribbon-ui.md's "one shared result form...
' reused after Sync Now, New Period, and the onboarding verify step."
' ---------------------------------------------------------------------


' BOTH FILES, SAVED AND CONFIRMED, AT THE END OF EVERY PATH THAT WROTE.
'
' 2026-08-08 on the rig: Apply Approved reported "16 written, 0 failed", took a
' backup, re-checked each change against its slide -- and the deck file was three
' hours stale. Nothing in this module called pres.Save at all. The workbook's
' review sheet had the same fate: the dialog said it had been "refreshed to match
' the deck as it is now" and the saved file contained no such sheet.
'
' Appended to the report rather than raised: the writes really did happen, so
' this is news about persistence, and a failure here must be impossible to miss
' while never discarding the report of what was written.
Private Function PersistBothFiles(pres As Object, wb As Object) As String
    Dim trouble As String

    Dim deckProblem As String
    deckProblem = DeckRegistry.SaveDeckVerified(pres)
    If deckProblem <> "" Then trouble = trouble & vbCrLf & deckProblem & vbCrLf

    Dim wbProblem As String
    If Not wb Is Nothing Then wbProblem = WorkbookBridge.SaveWorkbookVerified(wb)
    If wbProblem <> "" Then trouble = trouble & vbCrLf & wbProblem & vbCrLf

    If trouble = "" Then
        PersistBothFiles = vbCrLf & "Deck and workbook both saved." & vbCrLf
    Else
        PersistBothFiles = vbCrLf & "---- NOT SAVED ----" & trouble
    End If
End Function

Public Sub ShowSyncResult(title As String, report As String)
    MsgBox report, vbInformation, title
End Sub

' What the human sees when an action dies of something nobody anticipated.
'
' Every toolbar action is wrapped in a handler that ends here (see any entry
' point's own header for why a WRAPPER rather than an inline handler). Before
' 2026-07-29 there were none at all, so an unguarded raise anywhere below a
' button produced VBA's own Debug/End dialog -- which is not just ugly: End
' discards whatever the run had collected, and on 2026-07-29 that meant 45
' instance keys confirmed one prompt at a time.
'
' Pure, so the wording is testable without provoking a real error.
'
' Deliberately does NOT claim nothing was written. An error partway through a
' commit can leave real changes in the deck and the Data sheet, and a
' reassuring "no changes were made" would be a lie exactly when the human most
' needs the truth. Saying "check before re-running" is less comforting and
' actually correct.
Public Function UnexpectedErrorText(actionName As String, errNumber As Long, errDescription As String, errSource As String) As String
    Dim where As String
    where = Trim(errSource)
    If where = "" Then where = "an unidentified step"

    UnexpectedErrorText = _
        actionName & " stopped early -- something went wrong that this add-in didn't anticipate." & vbCrLf & vbCrLf & _
        "Error " & errNumber & ": " & errDescription & vbCrLf & _
        "Reported by: " & where & vbCrLf & vbCrLf & _
        "IMPORTANT: this run may have already changed your deck or its Data sheet before it stopped. Check both before running it again." & vbCrLf & vbCrLf & _
        "Nothing here needs the VBA editor -- if this keeps happening, the text above is the useful part to pass on."
End Function
