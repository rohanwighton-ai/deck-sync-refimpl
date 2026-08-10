Attribute VB_Name = "RibbonUI"
Option Explicit

' Cap for every report shown through ShowSyncResult -- see that Sub. Declared
' HERE because a module-level Const between two procedures is a syntax error
' that stops the whole project (learned the hard way, twice, 2026-08-08).
Public Const REPORT_CAP As Long = 900


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
' WHERE AM I -- rebuilds the readiness sheet and shows it.
'
' REBUILDS, never merely activates. A readiness surface that can be revisited
' without recomputation is the same defect as a verifier that reads a cache: it
' answers about a moment that has passed, and it is the surface a person would
' consult INSTEAD of checking. See Readiness.bas's header for the four rules.
' NO LONGER A BUTTON TARGET. The chain is the entry point; this stays as the
' error-handling wrapper its Core still needs. Private so the reachability
' check reports genuine orphans rather than adapters.
Private Sub WhereAmI()
    On Error GoTo Failed
    WhereAmICore
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult "Where am I", RibbonUI.UnexpectedErrorText("Where am I", Err.Number, Err.Description, Err.Source)
End Sub

Private Sub WhereAmICore()
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim workbookPath As String
    workbookPath = DeckRegistry.GetWorkbookPath(pres)
    If workbookPath = "" Then
        MsgBox "This deck has no paired workbook yet, so there is nothing to report " & _
               "on." & vbCrLf & vbCrLf & "Press '" & CommandBarUI.CAP_SYNC_NOW & "' -- it walks setup on a deck that has none.", _
               vbExclamation, "Where am I"
        Exit Sub
    End If

    Dim wb As Object
    Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath)
    If wb Is Nothing Then
        MsgBox "Could not open the paired workbook at: " & workbookPath, vbCritical, "Where am I"
        Exit Sub
    End If

    Dim r As ReadyReport
    r = Readiness.Build(pres, wb)
    Readiness.WriteSheet wb, pres, r

    ' The sheet is the answer; the dialog only carries the headline and points at
    ' it. Everything else would be truncated -- CapReport exists because MsgBox
    ' silently cuts near 1024 characters, and this report is longer than that.
    Readiness.ShowSheet wb
    MsgBox Readiness.Headline(r) & vbCrLf & vbCrLf & _
           "The full picture is on the '" & Readiness.READY_SHEET_NAME & _
           "' sheet, first tab of the workbook." & vbCrLf & vbCrLf & _
           "Nothing is disabled by what it says -- it reports, it does not gate.", _
           vbInformation, "Where am I"
End Sub

' NO LONGER A BUTTON TARGET. The chain is the entry point; this stays as the
' error-handling wrapper its Core still needs. Private so the reachability
' check reports genuine orphans rather than adapters.
Private Sub SyncNow()
    On Error GoTo Failed
    SyncNowCore
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult "Sync Now", RibbonUI.UnexpectedErrorText("Sync Now", Err.Number, Err.Description, Err.Source)
End Sub

' afterCreate caps the re-plan at ONE. Creating slides changes the deck under the
' queue, so the run must re-plan -- but if creation silently produced nothing, the
' orphans are still there and an unguarded re-entry would recurse forever, taking
' PowerPoint down with it. A guard that can only be crossed once is the difference
' between a re-plan and a loop.
Private Sub SyncNowCore(Optional ByVal afterCreate As Boolean = False)
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

    ' Collected per type so creation can be offered inside the ONE confirmation
    ' this run already shows, rather than as a separate button.
    Dim createTypes() As String
    Dim createTemplates() As Object
    Dim createSheets() As Sheet
    Dim createTypeCount As Long
    createTypeCount = 0

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

                ' AGGREGATED, or they would vanish exactly when there is more than
                ' one slide type -- the case where a wrong "nothing to sync" is
                ' hardest to notice. Only Items were merged before, and the loop
                ' above runs zero times for a type whose rows all reach no slide.
                '
                ' THESE LINES MUST STAY INSIDE THIS Else, DIRECTLY AFTER THE
                ' ASSIGNMENT ABOVE. VBA's Dim does not scope to a loop -- q is
                ' procedure-scoped and keeps the PREVIOUS type's contents on any
                ' iteration where BuildQueue is not reached (no registered type, or
                ' a refused sheet). Moved out one level, this would silently add the
                ' last type's orphans a second time. Nothing would raise; the count
                ' would just be wrong, which is the failure this whole change is
                ' about. (Rohan asked "is that the old q?" -- it is not, because
                ' q = BuildQueue(...) copies every field of the UDT, but only here.)
                combined.OrphanCount = combined.OrphanCount + q.OrphanCount
                If q.OrphanKeys <> "" Then
                    If combined.OrphanKeys <> "" Then combined.OrphanKeys = combined.OrphanKeys & ", "
                    combined.OrphanKeys = combined.OrphanKeys & q.OrphanKeys
                End If
                combined.FlaggedCount = combined.FlaggedCount + q.FlaggedCount
                combined.FlaggedNotes = combined.FlaggedNotes & q.FlaggedNotes
                combined.RowCount = combined.RowCount + q.RowCount
                combined.SlideCount = combined.SlideCount + q.SlideCount
                combined.SlideNoRowCount = combined.SlideNoRowCount + q.SlideNoRowCount
                If q.SlideNoRowKeys <> "" Then
                    If combined.SlideNoRowKeys <> "" Then combined.SlideNoRowKeys = combined.SlideNoRowKeys & ", "
                    combined.SlideNoRowKeys = combined.SlideNoRowKeys & q.SlideNoRowKeys
                End If

                ' Remembered per type, because creation needs THIS type's template
                ' slide and worksheet, and the confirmation below is for the whole
                ' deck. Only the types that actually have orphans are collected.
                If q.OrphanCount > 0 Then
                    createTypeCount = createTypeCount + 1
                    ReDim Preserve createTypes(1 To createTypeCount)
                    ReDim Preserve createTemplates(1 To createTypeCount)
                    ReDim Preserve createSheets(1 To createTypeCount)
                    createTypes(createTypeCount) = types(i)
                    Set createTemplates(createTypeCount) = templateSld
                    createSheets(createTypeCount) = sheet
                End If
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

    ' CREATION IS PART OF SYNC AGAIN, BOUNDED RATHER THAN ABSENT.
    '
    ' It was removed from the sync path on 2026-07-31 because the real deck had 43
    ' orphaned rows against 46 slides and one click would have duplicated the
    ' template 43 times -- and a hand-assembled board pack is missing most entities
    ' by design, so the old behaviour would have started creating at exactly the
    ' moment the content was finished. "Removing the capability beats warning about
    ' it" was right while the capability was unbounded.
    '
    ' Rohan, 2026-08-09: "why createmissingslides when it could be a function of
    ' sync? we need to simplify" -- and "I want the ppt <> excel to be at parity
    ' after a sync". Both are satisfied by bounding it instead of hiding it behind
    ' a button that never existed: a few orphans are new projects and get created
    ' inside the confirmation you already read; a lot of them mean this is not the
    ' deck you think it is, and the run refuses.
    Dim createdReport As String
    If combined.OrphanCount > 0 And Not afterCreate Then
        If Not ReviewQueue.OrphansLookLikeNewProjects(combined) Then
            MsgBox "Sync Now stopped -- this does not look like the right deck." & vbCrLf & vbCrLf & _
                combined.OrphanCount & " of " & combined.RowCount & " register row(s) match no slide." & vbCrLf & vbCrLf & _
                "A few unmatched rows are new projects. This many means the deck, the " & _
                "period, or the instance keys disagree with the register -- and creating " & _
                "a slide for each would duplicate the template across the deck." & vbCrLf & vbCrLf & _
                "Nothing was created and nothing was written." & vbCrLf & vbCrLf & _
                ReviewQueue.ParityText(combined), vbExclamation, "Sync Now"
            Exit Sub
        End If

        If MsgBox(combined.OrphanCount & " register row(s) have no slide:" & vbCrLf & vbCrLf & _
                  "    " & combined.OrphanKeys & vbCrLf & vbCrLf & _
                  "Create a slide for each, copied from the template and tagged?" & vbCrLf & vbCrLf & _
                  "Say No to sync the existing slides only and leave these for later.", _
                  vbYesNo + vbQuestion, "Sync Now -- new slides") = vbYes Then
            Dim ci As Long
            For ci = 1 To createTypeCount
                createdReport = createdReport & RunSync.CreateMissingSlides( _
                    createSheets(ci), createTypes(ci), createTemplates(ci), False) & vbCrLf
            Next ci
            ' Re-plan: the deck has changed underneath the queue, and the rows that
            ' were orphans now have slides whose fields need filling. Trusting the
            ' queue built before creation would sync the old set and report success.
            SyncNowCore True
            Exit Sub
        End If
    End If

    If combined.Count = 0 Then
        MsgBox ReviewQueue.FastPathRefusalText(combined) & vbCrLf & vbCrLf & _
               ReviewQueue.ParityText(combined), vbInformation, "Sync Now"
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

    ' CAPPED, AND THE QUESTION IS THE PART THAT SURVIVES.
    '
    ' This was a raw MsgBox while CapReport's own header named it as one of the
    ' eighteen callers left untouched. BatchSummaryText emits a before/after line
    ' plus every member key per batch -- roughly 450 characters for one 43-member
    ' batch -- so past about two batches this exceeded MsgBox's silent ~1024 limit.
    ' It is the dialog that authorises writing to slides, which makes a silent cut
    ' here the worst instance of the four this project has already fixed elsewhere.
    If MsgBox(CapReport(deckMacroWarn & ReviewQueue.ConfirmBatchText(combined), _
                        ReviewQueue.ConfirmBatchQuestion(combined)), _
              vbYesNo + vbQuestion, "Sync Now") <> vbYes Then
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

    ' PARITY IS STATED, NOT LEFT TO BE INFERRED FROM AN ABSENCE OF COMPLAINTS.
    '
    ' Measured AFTER the writes, from a fresh comparison, so it describes the deck
    ' as it now is rather than as it was planned to be. Both directions: rows with
    ' no slide, and slides with no row -- the second was invisible to every count
    ' this tool had, because the planner walks the register and simply never
    ' visits a slide the register does not mention.
    Dim parity As ReviewQueueSet
    Dim pj As Long
    For pj = lo To hi
        Dim pTemplate As Object
        Dim pWsName As String
        If DeckRegistry.LookupType(pres, types(pj), pTemplate, pWsName) Then
            Dim pProblem As String
            Dim pSheet As Sheet
            pSheet = ExcelOutput.ReadSheetForDeckPeriod( _
                WorkbookBridge.GetOrAddWorksheet(wb, pWsName), deckPeriod, pProblem)
            If pProblem = "" Then
                Dim pq As ReviewQueueSet
                pq = ReviewQueue.BuildQueue(pSheet, types(pj))
                parity.RowCount = parity.RowCount + pq.RowCount
                parity.SlideCount = parity.SlideCount + pq.SlideCount
                parity.OrphanCount = parity.OrphanCount + pq.OrphanCount
                parity.SlideNoRowCount = parity.SlideNoRowCount + pq.SlideNoRowCount
                If pq.OrphanKeys <> "" Then
                    If parity.OrphanKeys <> "" Then parity.OrphanKeys = parity.OrphanKeys & ", "
                    parity.OrphanKeys = parity.OrphanKeys & pq.OrphanKeys
                End If
                If pq.SlideNoRowKeys <> "" Then
                    If parity.SlideNoRowKeys <> "" Then parity.SlideNoRowKeys = parity.SlideNoRowKeys & ", "
                    parity.SlideNoRowKeys = parity.SlideNoRowKeys & pq.SlideNoRowKeys
                End If
            End If
        End If
    Next pj
    fullReport = fullReport & vbCrLf & ReviewQueue.ParityText(parity) & vbCrLf

    If createdReport <> "" Then fullReport = createdReport & vbCrLf & fullReport

    fullReport = fullReport & PersistBothFiles(pres, wb)
    WorkbookBridge.WriteRunLog wb, "Sync Now -- full report", fullReport
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
        MsgBox "This deck has no paired workbook yet. Press '" & CommandBarUI.CAP_SYNC_NOW & "' -- it walks setup on a deck that has none.", vbExclamation, title
        Exit Function
    End If

    types = DeckRegistry.ListRegisteredTypes(pres)

    Dim hasTypes As Boolean
    On Error Resume Next
    lo = LBound(types): hi = UBound(types)
    hasTypes = (Err.Number = 0)
    On Error GoTo 0

    If Not hasTypes Then
        MsgBox "This deck has no registered slide types yet. Press '" & CommandBarUI.CAP_SYNC_NOW & "' -- it walks setup on a deck that has none.", vbExclamation, title
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

        ' VERIFIED, because the whole point of this prompt is that the file and
        ' the screen agree before anything reads the register. A Save that
        ' reported nothing and moved nothing would leave us proceeding on exactly
        ' the mismatch the prompt exists to prevent.
        Dim promptSaveProblem As String
        promptSaveProblem = WorkbookBridge.SaveWorkbookVerified(wb)
        If promptSaveProblem <> "" Then
            MsgBox promptSaveProblem & vbCrLf & vbCrLf & _
                   "Stopping here rather than reading values that are not in the file.", _
                   vbCritical, title
            Exit Function
        End If
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
            ' Capped for the same reason as Sync Now's confirmation, and with the
            ' same tail preserved: this lists every duplicated key, so a deck with
            ' many of them is exactly the case where the question would be cut.
            If MsgBox(CapReport(IdentityCheck.DuplicateKeyWarningText(types(dupType), dupReport) & _
                      vbCrLf & vbCrLf & "Continue anyway?", "Continue anyway?"), _
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
' NO LONGER A BUTTON TARGET. The chain is the entry point; this stays as the
' error-handling wrapper its Core still needs. Private so the reachability
' check reports genuine orphans rather than adapters.
Private Sub ReviewChanges()
    On Error GoTo Failed
    ReviewChangesCore False
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult CommandBarUI.CAP_REVIEW_CHANGES, RibbonUI.UnexpectedErrorText(CommandBarUI.CAP_REVIEW_CHANGES, Err.Number, Err.Description, Err.Source)
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
' NO LONGER A BUTTON TARGET. The chain is the entry point; this stays as the
' error-handling wrapper its Core still needs. Private so the reachability
' check reports genuine orphans rather than adapters.
Private Sub ReviewChangesApproveAll()
    On Error GoTo Failed
    ReviewChangesCore True
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult "Review Changes (approve all)", RibbonUI.UnexpectedErrorText("Review Changes (approve all)", Err.Number, Err.Description, Err.Source)
End Sub

Private Sub ReviewChangesCore(approveAll As Boolean)
    Dim title As String
    title = IIf(approveAll, "Review Changes (approve all)", CommandBarUI.CAP_REVIEW_CHANGES)

    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim workbookPath As String
    workbookPath = DeckRegistry.GetWorkbookPath(pres)
    If workbookPath = "" Then
        MsgBox "This deck has no paired workbook yet. Press '" & CommandBarUI.CAP_SYNC_NOW & "' -- it walks setup on a deck that has none.", vbExclamation, title
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
        MsgBox "This deck has no registered slide types yet. Press '" & CommandBarUI.CAP_SYNC_NOW & "' -- it walks setup on a deck that has none.", vbExclamation, title
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

        ' VERIFIED, because the whole point of this prompt is that the file and
        ' the screen agree before anything reads the register. A Save that
        ' reported nothing and moved nothing would leave us proceeding on exactly
        ' the mismatch the prompt exists to prevent.
        Dim promptSaveProblem As String
        promptSaveProblem = WorkbookBridge.SaveWorkbookVerified(wb)
        If promptSaveProblem <> "" Then
            MsgBox promptSaveProblem & vbCrLf & vbCrLf & _
                   "Stopping here rather than reading values that are not in the file.", _
                   vbCritical, title
            Exit Sub
        End If
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
' ---------------------------------------------------------------------
' THE CHAIN -- "1. Sync Now"
' ---------------------------------------------------------------------
'
' One button that does the next right thing, whatever state the deck is in.
' It reads that state, states its plan before the first write, and stops only
' where there is a genuine decision.
'
' EVERY PUBLIC CAPABILITY IS CALLED FROM HERE, deliberately and by name. When
' this was three buttons the chains called the private Cores, which orphaned
' nine public entry points -- built, and unreachable by a person. The static
' check caught it. Calling the public wrappers keeps each one's own error
' handler in place too, which is why they exist.
'
' V1 HONESTY: DraftingUI's six entry points have no work/prompt split, so each
' stage still owns its prompts. The SEQUENCE is right; the stop-count discipline
' arrives with the Core splits.
Public Sub SyncNowChain()
    On Error GoTo Failed
    SyncNowChainCore
    Exit Sub
Failed:
    ' END COLLECTING IN THE HANDLER TOO. Left on, the next standalone run would
    ' swallow its own messages into a buffer nobody reads -- a silent tool is a
    ' worse failure than a noisy one.
    Dim partial As String
    partial = DraftingUI.EndCollecting()
    If partial <> "" Then partial = "What had happened before the error:" & vbCrLf & vbCrLf & partial & vbCrLf & vbCrLf
    RibbonUI.ShowSyncResult "Sync Now", partial & RibbonUI.UnexpectedErrorText("Sync Now", Err.Number, Err.Description, Err.Source)
End Sub

' Returns True to CARRY ON with the chain, False to stop.
'
' Stops only when the person chose to go and tag something, or cancelled. A
' scan that cannot run does NOT stop the chain: refusing to sync because a
' check was unable to look would be the check gating rather than offering, and
' Readiness.bas:51 governs the whole design -- it offers, it does not gate.
Private Function OfferMarkingForUnwiredFields(pres As Object, TITLE As String) As Boolean
    OfferMarkingForUnwiredFields = True

    Dim wb As Object
    On Error Resume Next
    Set wb = WorkbookBridge.OpenOrGetWorkbook(DeckRegistry.GetWorkbookPath(pres))
    On Error GoTo 0
    If wb Is Nothing Then Exit Function

    Dim period As String
    period = DeckRegistry.GetDeckPeriod(pres)
    If period = "" Then Exit Function

    Dim types() As String
    types = DeckRegistry.ListRegisteredTypes(pres)
    Dim lo As Long, hi As Long, hasTypes As Boolean
    On Error Resume Next
    lo = LBound(types): hi = UBound(types)
    hasTypes = (Err.Number = 0)
    On Error GoTo 0
    If Not hasTypes Then Exit Function

    Dim i As Long
    For i = lo To hi
        Dim templateSld As Object, wsName As String
        Set templateSld = Nothing
        If DeckRegistry.LookupType(pres, types(i), templateSld, wsName) Then
            If WorkbookBridge.WorksheetExists(wb, wsName) Then
                Dim problem As String
                Dim sheet As Sheet
                sheet = ExcelOutput.ReadSheetForDeckPeriod( _
                    WorkbookBridge.GetOrAddWorksheet(wb, wsName), period, problem)
                If problem = "" Then
                    Dim wiring As FieldWiringResult
                    wiring = FieldWiring.ScanFieldWiring(types(i), sheet.Fields, templateSld)

                    ' PARTIAL COVERAGE ALONE DOES NOT STOP YOU. A field part-way
                    ' across a deck is a real state, and prompting about it on
                    ' every press is how a dialog gets clicked through -- which
                    ' would cost the times it matters. It is reported on the
                    ' START HERE sheet instead.
                    If wiring.Scanned And (wiring.UnmarkedCount > 0 _
                            Or wiring.TemplateUnmarkedCount > 0 Or wiring.OrphanCount > 0) Then
                        ' NAMES THE FIELDS, NOT JUST A COUNT. Fix-list 1a: a
                        ' true count with no subject sends people to check the
                        ' wrong thing, four times over now.
                        Dim answer As VbMsgBoxResult
                        answer = MsgBox( _
                            "Slide type '" & types(i) & "' has fields with nothing to write into." & vbCrLf & vbCrLf & _
                            FieldWiring.WiringText(wiring) & vbCrLf & vbCrLf & _
                            "Syncing now would carry those fields all the way to the slide and " & _
                            "refuse them there, once per slide." & vbCrLf & vbCrLf & _
                            "Yes    -- tag them now, by clicking each shape." & vbCrLf & _
                            "No     -- sync anyway and leave them untagged." & vbCrLf & _
                            "Cancel -- change nothing.", _
                            vbYesNoCancel + vbExclamation, TITLE)

                        If answer = vbYes Then
                            ' TAG ON THE TEMPLATE WHEN THAT IS WHAT IS MISSING.
                            ' A new slide is a Duplicate of the template and
                            ' inherits its shape tags, so tagging the template
                            ' fixes every slide not yet made; tagging an
                            ' instance fixes only that one.
                            ' A SHAPE HAS TO BE SELECTED BEFORE MARKING CAN DO
                            ' ANYTHING, and this branch used to call it
                            ' regardless -- so the FIRST press always produced
                            ' "Select exactly one shape first" and always would.
                            ' Navigating to the template guarantees it: arriving
                            ' at a slide selects nothing. Rohan hit it twice.
                            '
                            ' A dialog that cannot succeed on its first showing
                            ' is the same defect as a check that cannot fail --
                            ' it looks like a step and is furniture.
                            Dim selCount As Long
                            selCount = 0
                            On Error Resume Next
                            If Application.ActiveWindow.Selection.Type = 2 Then   ' ppSelectionShapes
                                selCount = Application.ActiveWindow.Selection.ShapeRange.Count
                            End If
                            On Error GoTo 0

                            If selCount = 1 Then
                                BatchOnboardFlow.MarkFieldForBatch
                            Else
                                Dim whereTo As String
                                If wiring.TemplateUnmarkedCount > 0 And Not templateSld Is Nothing Then
                                    On Error Resume Next
                                    Application.ActiveWindow.View.GotoSlide templateSld.SlideIndex
                                    On Error GoTo 0
                                    whereTo = "You are now on the TEMPLATE slide. Tag the missing " & _
                                        "fields HERE: a new slide is a copy of this one and inherits " & _
                                        "its tags, so tagging it here fixes every slide you make " & _
                                        "from now on." & vbCrLf & vbCrLf
                                End If

                                MsgBox whereTo & _
                                    "Now CLICK THE SHAPE you want to tag, then press '" & _
                                    CommandBarUI.CAP_SYNC_NOW & "' again." & vbCrLf & vbCrLf & _
                                    "Nothing was changed." & vbCrLf & vbCrLf & _
                                    "(If the shape sits inside a group, one click selects the whole " & _
                                    "group -- that is fine, you will be offered the shapes inside it.)", _
                                    vbInformation, TITLE
                            End If
                            OfferMarkingForUnwiredFields = False
                            Exit Function
                        ElseIf answer = vbCancel Then
                            MsgBox "Nothing was changed.", vbInformation, TITLE
                            OfferMarkingForUnwiredFields = False
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Function

Private Sub SyncNowChainCore()
    Const TITLE As String = "1. Sync Now"

    Dim pres As Object
    Set pres = Application.ActivePresentation

    ' SETUP IS A PRECONDITION, NOT AN ACTIVITY -- once ever per slide type, and
    ' only on a deck that has none. A configured deck never sees this.
    Dim types() As String
    types = DeckRegistry.ListRegisteredTypes(pres)
    Dim lo As Long, hi As Long, hasTypes As Boolean
    On Error Resume Next
    lo = LBound(types): hi = UBound(types)
    hasTypes = (Err.Number = 0)
    On Error GoTo 0

    If Not hasTypes Then
        Dim setupAnswer As VbMsgBoxResult
        setupAnswer = MsgBox( _
            "This deck has no slide type set up yet, so there is nothing to sync." & vbCrLf & vbCrLf & _
            "Setting one up is a ONE-OFF and you only do it once per slide layout:" & vbCrLf & _
            "  - tag each field on a template slide by clicking it" & vbCrLf & _
            "  - link the other slides of that layout to the register" & vbCrLf & vbCrLf & _
            "Yes  -- tag every field on a slide at once, in an Excel grid (quicker)." & vbCrLf & _
            "No   -- tag them one at a time by clicking each shape." & vbCrLf & _
            "Cancel -- change nothing.", vbYesNoCancel + vbQuestion, TITLE)

        If setupAnswer = vbYes Then
            DiscoverUI.DiscoverFields
            BatchOnboardFlow.BatchOnboardType
        ElseIf setupAnswer = vbNo Then
            BatchOnboardFlow.MarkFieldForBatch
        Else
            ' STARTING THE TAGGING AGAIN lives here because this is the only
            ' place a person is tagging. Offered on the way out rather than as
            ' its own choice: it is what you want when the LAST attempt went
            ' wrong, which is exactly when you are cancelling.
            If MsgBox("Nothing was changed." & vbCrLf & vbCrLf & _
                      "Discard any marking already done on this deck and start clean?" & vbCrLf & _
                      "(It cannot remove just one field.)", _
                      vbYesNo + vbQuestion, TITLE) = vbYes Then
                BatchOnboardFlow.ClearMarkedFieldsForBatch
            End If
        End If
        Exit Sub
    End If

    ' WHERE YOU ARE, WITHOUT A BUTTON FOR IT. Folded in per Rohan 2026-08-09.
    ' Rebuilt first so the sheet reflects the state this run is about to act on,
    ' and so a person who cancels at the plan below still gets the picture.
    WhereAmI

    ' THE REPAIR, OFFERED ONLY WHEN IT IS THE ANSWER. Repoint sets the workbook
    ' path; it does NOT rebuild the slide-type link, so it is offered when the
    ' pairing is the thing that is broken and says so when it has not fully
    ' worked.
    If DeckRegistry.GetWorkbookPath(pres) = "" Then
        If MsgBox("This deck is not paired with a workbook, so there is nothing to sync." & vbCrLf & vbCrLf & _
                  "Point it at one now?", vbYesNo + vbExclamation, TITLE) = vbYes Then
            DraftingUI.RepointWorkbookUI
        End If
        Exit Sub
    End If

    ' A FIELD WITH NOTHING TO WRITE INTO, OFFERED BEFORE THE QUARTER IS SPENT
    ' DRAFTING IT.
    '
    ' Until 2026-08-10 `MarkFieldForBatch` had exactly ONE call site -- the
    ' `Not hasTypes` branch above -- so a deck whose slide type was already
    ' registered had NO WAY to tag a new field. Every field added to the rig
    ' after its first setup was tagged by Claude over COM, which is the
    ' dependence this tool exists to remove. This branch is what makes
    ' `RM_MARK_MISSING_FIELDS` a true remedy rather than advice to press a
    ' button that does not exist.
    '
    ' It fires only when there is something to tag, so a steady-state quarter
    ' never sees it -- the same rule the rest of the chain follows.
    If Not OfferMarkingForUnwiredFields(pres, TITLE) Then Exit Sub

    ' THE PLAN, BEFORE THE FIRST WRITE. This is the hazard a chain creates and a
    ' row of buttons did not: it can do more than the person expected. Saying
    ' what it will do, in order, is the answer.
    If MsgBox("This will, in order:" & vbCrLf & vbCrLf & _
              "  1. Set the deck's quarter (it asks you which)" & vbCrLf & _
              "  2. Offer to copy last quarter's rows forward" & vbCrLf & _
              "  3. Rebuild the sheets you write in" & vbCrLf & _
              "  4. Publish the rows you ticked, for one field" & vbCrLf & _
              "  5. Show every slide change and ASK before writing any of it" & vbCrLf & vbCrLf & _
              "Nothing reaches a slide until step 5, and you can cancel there." & vbCrLf & _
              "Cancelling at any step leaves everything before it in place." & vbCrLf & vbCrLf & _
              "Go ahead?", vbYesNo + vbQuestion, TITLE) <> vbYes Then
        MsgBox "Nothing was changed.", vbInformation, TITLE
        Exit Sub
    End If

    ' ONE REPORT FOR THE WHOLE PROLOGUE. Each stage's decisions still stop and
    ' ask; only its informational messages are collected. A stage with nothing
    ' to do now says so in the report instead of interrupting to say it.
    DraftingUI.BeginCollecting
    DraftingUI.StartQuarter
    DraftingUI.RollForwardUI
    DraftingUI.RefreshDraftingSheets
    DraftingUI.CopyAiDraftsToSubmit
    DraftingUI.PublishDraftsForField

    Dim staged As String
    staged = DraftingUI.EndCollecting()
    If staged <> "" Then
        MsgBox CapReport(staged, "Next: the slide changes."), vbInformation, TITLE
    End If

    ' The deck-level sync, with its own detection of unapplied ticks in front.
    PutItOnTheSlidesCore
End Sub

' ---------------------------------------------------------------------
' CHAIN C -- "3. Put it on the slides"
' ---------------------------------------------------------------------
'
' One button for the deck-level sync, sequencing the review and the apply
' instead of asking a person to know which of three buttons they are up to.
'
' A CHAIN, NOT A MENU. It reads the current state and runs the sequence,
' stopping only where there is a decision in the gap -- Rohan's rule, stated at
' BatchOnboardFlow.bas:2886: "a boundary earns its place only where a person has
' to do work or make a decision in the gap." There are exactly two stops here,
' and each one is named below.
'
' STOP 1 is the safety property of the whole design. ReviewChangesCore calls
' WriteQueueSheet unconditionally, so a chain that opened by rebuilding the
' queue would take every tick with it. Detect first, branch, never rebuild
' without being told to.
'
' STOP 2 is whatever the route it takes already asks -- the review sheet's own
' approvals, or ApplyApprovedCore's consent dialog. This chain adds no
' confirmation of its own, because a second dialog restating the first is the
' thing that teaches a person to click through.
' NO LONGER A BUTTON TARGET. The chain is the entry point; this stays as the
' error-handling wrapper its Core still needs. Private so the reachability
' check reports genuine orphans rather than adapters.
Private Sub PutItOnTheSlides()
    On Error GoTo Failed
    PutItOnTheSlidesCore
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult "Put it on the slides", RibbonUI.UnexpectedErrorText("Put it on the slides", Err.Number, Err.Description, Err.Source)
End Sub

Private Sub PutItOnTheSlidesCore()
    ' Titled from the caption that actually exists. This said "3. Put it on the
    ' slides" -- a caption from the three-button design, which lasted three hours.
    Dim TITLE As String
    TITLE = CommandBarUI.CAP_SYNC_NOW & " -- slide changes"

    ' The guards for a missing workbook, missing types and an unopenable file
    ' live in ReviewChangesCore and say why in each case. Duplicating them here
    ' would mean two sets of wording to keep true, so when the detection cannot
    ' run this delegates and lets that path do the explaining.
    Dim pending As Long
    Dim sheetNames As String
    Dim stamp As String
    pending = ScanPendingApprovals(sheetNames, stamp)

    If pending = 0 Then
        ' REVIEW AND APPROVE ARE PART OF THE CHAIN. Rohan, 2026-08-09. Bulk
        ' approval keeps the property RibbonUI.bas:573 argued a separate button
        ' was protecting -- it is never the default, it has to be chosen by name
        ' each time, and ReviewChangesCore still prepends the APPROVE-ALL banner
        ' to the report and the Run Log. The banner was always the mechanism
        ' doing that work; the button was just where it lived.
        Dim readAll As VbMsgBoxResult
        readAll = MsgBox( _
            "Ready to build the list of slide changes." & vbCrLf & vbCrLf & _
            "Yes  -- read each change and tick the ones you want (normal)." & vbCrLf & _
            "No   -- tick EVERYTHING without reading it. Scratch copies only." & vbCrLf & _
            "Cancel -- change nothing.", vbYesNoCancel + vbQuestion, TITLE)

        If readAll = vbYes Then
            ReviewChangesCore False
        ElseIf readAll = vbNo Then
            ReviewChangesCore True
        Else
            MsgBox "Nothing was changed.", vbInformation, TITLE
        End If
        Exit Sub
    End If

    ' STOP 1. Names the sheet, the count and when it was built, because a count
    ' without its subject sends you to check the wrong thing.
    Dim msg As String
    msg = sheetNames & " holds " & pending & " ticked change(s) from " & stamp & _
          ", not yet applied." & vbCrLf & vbCrLf & _
          "Yes -- apply those " & pending & " change(s) to the slides now." & vbCrLf & _
          "No -- build a fresh review instead. THE " & pending & " TICK(S) ARE LOST." & vbCrLf & _
          "Cancel -- change nothing."

    Dim answer As VbMsgBoxResult
    answer = MsgBox(msg, vbYesNoCancel + vbQuestion, TITLE)

    If answer = vbYes Then
        ApplyApprovedCore
    ElseIf answer = vbNo Then
        ReviewChangesCore False
    Else
        MsgBox "Nothing was changed. " & sheetNames & " still holds its " & _
               pending & " ticked change(s).", vbInformation, TITLE
    End If
End Sub

' Totals the unapplied ticks across every registered slide type, and names the
' sheets they are on. Returns 0 when anything needed to look is missing -- the
' caller then delegates to the path that reports why.
Private Function ScanPendingApprovals(ByRef sheetNames As String, ByRef stamp As String) As Long
    sheetNames = ""
    stamp = ""
    ScanPendingApprovals = 0

    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim workbookPath As String
    workbookPath = DeckRegistry.GetWorkbookPath(pres)
    If workbookPath = "" Then Exit Function

    Dim types() As String
    types = DeckRegistry.ListRegisteredTypes(pres)

    Dim lo As Long, hi As Long, hasTypes As Boolean
    On Error Resume Next
    lo = LBound(types): hi = UBound(types)
    hasTypes = (Err.Number = 0)
    On Error GoTo 0
    If Not hasTypes Then Exit Function

    Dim wb As Object
    Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath)
    If wb Is Nothing Then Exit Function

    Dim total As Long
    Dim i As Long
    For i = lo To hi
        Dim n As Long
        Dim wsName As String
        Dim thisStamp As String
        n = ReviewQueue.PendingApprovals(wb, types(i), wsName, thisStamp)
        If n > 0 Then
            total = total + n
            If sheetNames = "" Then
                sheetNames = "'" & wsName & "'"
                stamp = thisStamp
            Else
                sheetNames = sheetNames & ", '" & wsName & "'"
            End If
        End If
    Next i

    ScanPendingApprovals = total
End Function

' NO LONGER A BUTTON TARGET. The chain is the entry point; this stays as the
' error-handling wrapper its Core still needs. Private so the reachability
' check reports genuine orphans rather than adapters.
Private Sub ApplyApprovedChanges()
    On Error GoTo Failed
    ApplyApprovedCore
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult CommandBarUI.CAP_APPLY_APPROVED, RibbonUI.UnexpectedErrorText(CommandBarUI.CAP_APPLY_APPROVED, Err.Number, Err.Description, Err.Source)
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
        MsgBox "This deck has no paired workbook yet -- nothing to apply.", vbExclamation, CommandBarUI.CAP_APPLY_APPROVED
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
        MsgBox "This deck has no registered slide types yet -- nothing to apply.", vbExclamation, CommandBarUI.CAP_APPLY_APPROVED
        Exit Sub
    End If

    Dim wb As Object
    Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath)
    If wb Is Nothing Then
        MsgBox "Could not open the paired workbook at: " & workbookPath, vbCritical, CommandBarUI.CAP_APPLY_APPROVED
        Exit Sub
    End If

    ' The ticks live in the workbook, so an unsaved workbook means the
    ' approvals being read are on screen and not in any file. Same refusal as
    ' the review step, for the same reason.
    If WorkbookBridge.IsDirty(wb) Then
        If MsgBox(WorkbookBridge.UnsavedWorkbookText(workbookPath), _
                  vbYesNo + vbExclamation, CommandBarUI.CAP_APPLY_APPROVED) <> vbYes Then
            Exit Sub
        End If

        ' VERIFIED, because the whole point of this prompt is that the file and
        ' the screen agree before anything reads the register. A Save that
        ' reported nothing and moved nothing would leave us proceeding on exactly
        ' the mismatch the prompt exists to prevent.
        Dim promptSaveProblem As String
        promptSaveProblem = WorkbookBridge.SaveWorkbookVerified(wb)
        If promptSaveProblem <> "" Then
            MsgBox promptSaveProblem & vbCrLf & vbCrLf & _
                   "Stopping here rather than reading values that are not in the file.", _
                   vbCritical, CommandBarUI.CAP_APPLY_APPROVED
            Exit Sub
        End If
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
                    "No review has been built for this type. Press '" & CommandBarUI.CAP_SYNC_NOW & "' first." & vbCrLf & vbCrLf
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
    WorkbookBridge.WriteRunLog wb, "Apply Approved -- full report", fullReport
    ShowSyncResult CommandBarUI.CAP_APPLY_APPROVED, fullReport
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
            fullReport = fullReport & RunSync.PreviewRoutineSync(ws, types(i), DeckRegistry.GetDeckPeriod(pres)) & vbCrLf
        Else
            fullReport = fullReport & "SKIPPED " & types(i) & ": registered type's template slide no longer resolves (was it deleted?)" & vbCrLf
        End If
    Next i

    ' Warns where Sync Now refuses. The preview writes nothing, so unsaved data
    ' cannot damage the deck here -- but a preview of values that exist in no
    ' file is still a preview of something that might never be synced, and the
    ' whole worth of this report is that it can be trusted. Stated at the TOP:
    ' a caveat below a long report is a caveat nobody reads.
    ' MEASURED ONCE, BEFORE THIS RUN DIRTIES ANYTHING ITSELF.
    '
    ' There used to be a second IsDirty test further down, after WriteRunLog. That
    ' one could never be false: WriteRunLog clears and rewrites cells, so the
    ' workbook is unsaved BECAUSE OF THIS RUN by the time it was asked. The single
    ' warning that means "this preview cannot be trusted" was therefore shown on
    ' every preview including clean ones -- which is how you train someone to click
    ' past the one that will matter. Same shape as an always-true guard: it reads
    ' as care taken and stops anyone re-checking.
    Dim wbWasDirty As Boolean
    wbWasDirty = WorkbookBridge.IsDirty(wb)

    If wbWasDirty Then
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
        "Read it there, then press '" & CommandBarUI.CAP_SYNC_NOW & "' again."

    ' wbWasDirty, NOT a fresh IsDirty -- WriteRunLog above has dirtied the workbook
    ' by now, so re-asking would always answer yes. See the note at the first test.
    If wbWasDirty Then
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
        MsgBox "This deck has no registered slide types yet. Press '" & CommandBarUI.CAP_SYNC_NOW & "' -- it walks setup on a deck that has none.", vbExclamation, "Audit Fields"
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
Private Function ExcludeSlide(slides() As Object, dropSld As Object) As Object()
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
        MsgBox "This deck has no registered slide types yet. Press '" & CommandBarUI.CAP_SYNC_NOW & "' -- it walks setup on a deck that has none.", vbExclamation, "Create Template Slide"
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

' EVERY REPORT THAT GOES THROUGH HERE IS CAPPED, AND THE CUT IS ANNOUNCED.
'
' MsgBox truncates near 1024 characters and says nothing about it. That was found
' and fixed on the drafting dialog, then found again on Preview Sync -- and both
' times the fix went in where the failure was seen, leaving the eighteen callers
' of this Sub untouched, including Sync Now and Apply Approved with their full
' accumulated reports. Apply Approved already lists one line per written field;
' at 43 slides it would cut silently, and the summary lives at the BOTTOM.
'
' Capped here so it cannot recur in a caller nobody thought about. The callers
' that have a workbook to hand also write their full detail to the Run Log sheet;
' this is the floor under all of them, not a replacement for that.
' ONE PLACE THAT KNOWS ABOUT THE LIMIT.
'
' MsgBox truncates near 1024 characters and says nothing. That was fixed on the
' drafting dialog, then again on Preview Sync, then again in ShowSyncResult --
' three times, each only where the failure had been seen, and on 2026-08-08 it
' bit a FOURTH time in the publish confirmation. That dialog is the one that asks
' permission to write, so a silent cut there hides part of what is being
' authorised: it ended mid-word at "would", with the summary below the cut.
'
' Any dialog anywhere can now call this instead of inventing its own answer.
' mustKeep is a tail ALREADY PRESENT at the end of text that must survive the cut.
' Pass the question on any dialog that asks permission. ConfirmBatchText appends
' "Apply the N uniform change(s) above?" LAST, so a plain truncation removed the
' question and left Yes/No buttons answering nothing -- the person is then agreeing
' to a list they can only partly see, with no visible thing being agreed to.
' Not duplicated when the text fits: under the cap the tail is already there and
' the function returns early.
Public Function CapReport(text As String, Optional mustKeep As String = "") As String
    CapReport = text
    If Len(text) <= REPORT_CAP Then Exit Function

    Dim notice As String
    notice = vbCrLf & vbCrLf & "[shortened -- the full list is on the '" & _
        WorkbookBridge.RUN_LOG_SHEET_NAME & "' sheet in the workbook]"

    Dim room As Long
    room = REPORT_CAP - Len(notice) - Len(mustKeep)
    If room < 200 Then room = 200      ' a floor, so a long tail cannot erase the body

    CapReport = Left$(text, room) & notice
    If mustKeep <> "" Then CapReport = CapReport & vbCrLf & vbCrLf & mustKeep
End Function

Public Sub ShowSyncResult(title As String, report As String)
    ' THE BUILD, ON EVERY REPORT. So "is this fix in?" is answerable from the
    ' screen instead of by asking. Appended AFTER CapReport so the stamp cannot
    ' be the thing that gets truncated away.
    MsgBox CapReport(report) & vbCrLf & vbCrLf & "build: " & CommandBarUI.BUILD_STAMP, _
           vbInformation, title
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
