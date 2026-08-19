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
' "WHERE AM I" DELETED ENTIRELY, 2026-08-17 evening (Rohan: "delete the whole
' thing, keep anything useful but otherwise get rid of it"). It rebuilt a
' readiness sheet by re-verifying the deck's period and workbook path from
' the SAVED FILE'S OWN BYTES (two full copies of the deck plus slow
' Shell.Application ZIP extraction) and running a full ReviewQueue.BuildQueue
' diff per registered slide type -- all of it redundant with checks the real
' operations already make, cheaply, at the point they actually matter (see
' e.g. RunSync.bas's own "REFUSED: this slide type has no template slide
' registered"). Nothing here was worth relocating: every check traced back to
' something already caught elsewhere with its own clear message. `Readiness.
' bas` is deleted along with this. FIX-LIST/NEXT-SESSION carry the full
' reasoning for anyone looking for it later.


' SyncNow/SyncNowCore DELETED ENTIRELY, 2026-08-17/18 night (FIX-LIST item AJ).
' Superseded by SyncNowChain/SyncNowChainCore (below), the button's real
' handler since the "1. Set up my quarter" rename -- confirmed via grep before
' deletion that nothing called SyncNow or SyncNowCore except each other and
' SyncNowCore's own internal re-plan recursion (afterCreate:=True), which is
' moot if the function is never entered from outside in the first place.
' Contained the worst call pattern found in either of tonight's audits: four
' full ReviewQueue.BuildQueue calls per slide type, on every press, of a
' button nothing could press. The R13/F1-F10 design-rules header above this
' note predates the rename and is left in place -- WarnOnDuplicateKeys (R9)
' and the uniform-batch confirmation rules it documents are still shared with
' the live ReviewChangesCore, not exclusive to the deleted code.

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
        MsgBox "This deck has no paired workbook yet. Press '" & CommandBarUI.CAP_SET_UP_QUARTER & "' -- it walks setup on a deck that has none.", vbExclamation, title
        Exit Function
    End If

    types = DeckRegistry.ListRegisteredTypes(pres)

    Dim hasTypes As Boolean
    On Error Resume Next
    lo = LBound(types): hi = UBound(types)
    hasTypes = (Err.Number = 0)
    On Error GoTo 0

    If Not hasTypes Then
        MsgBox "This deck has no registered slide types yet. Press '" & CommandBarUI.CAP_SET_UP_QUARTER & "' -- it walks setup on a deck that has none.", vbExclamation, title
        Exit Function
    End If

    Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath)
    If wb Is Nothing Then
        MsgBox "Could not open the paired workbook at: " & workbookPath, vbCritical, title
        Exit Function
    End If

    ' IS THIS EVEN OUR REGISTER? Same check as DraftingUI.Resolve's, 2026-08-19
    ' -- this function has no callers today (confirmed by grep; BuildAllQueuesCore
    ' and ApplyApprovedCore both grew their own inline copy of this exact
    ' resolve-and-check sequence instead of calling it), but it is a live piece
    ' of this project's history of exactly that drift and worth keeping correct
    ' in case something calls it again.
    Dim pairNote As String
    pairNote = DeckRegistry.PairingProblem(pres, wb)
    If pairNote <> "" Then
        MsgBox pairNote, vbCritical, title
        Exit Function
    End If

    ' Saves quietly instead of asking -- see WorkbookBridge.EnsureSavedQuietly's
    ' own header for why the prompt this replaced (2026-08-19) no longer needs
    ' to interrupt to do its job. A genuine save failure is still never hidden.
    Dim promptSaveProblem As String
    promptSaveProblem = WorkbookBridge.EnsureSavedQuietly(wb, workbookPath)
    If promptSaveProblem <> "" Then
        MsgBox promptSaveProblem & vbCrLf & vbCrLf & _
               "Stopping here rather than reading values that are not in the file.", _
               vbCritical, title
        Exit Function
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
' A BUTTON TARGET AGAIN, 2026-08-14. Rohan: review is its own action, "clear
' what it is and isn't for, part of a sequence, or not." It was made private when
' the single chain swallowed it; the artifact split gives it back its own door.
Public Sub ReviewChanges()
    On Error GoTo Failed
    ReviewChangesCore
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult CommandBarUI.STAGE_REVIEW_CHANGES, RibbonUI.UnexpectedErrorText(CommandBarUI.STAGE_REVIEW_CHANGES, Err.Number, Err.Description, Err.Source)
End Sub

' BULK APPROVE IS DELETED, 2026-08-14, and this note is its headstone.
'
' It was the Round 13 §0.1 loosened setting: build the queue, then tick every row
' without reading them. Permitted by the RM's ruling only on a carved copy, and
' R13.2 named the risk itself -- bulk approval "teaches the operator to click
' through".
'
' It went because NO PERSON COULD REACH IT. Its wrapper was private, its only
' remaining door was the three-way "Ready to build the list" prompt, and that
' prompt was deleted tonight as ceremony in front of the real gate. What was left
' was a tested capability with no way in: ReviewQueue.ApproveAllInSheet had a
' test and no caller, which is this project's signature defect -- the tested
' picture injector, the tested progress bars, the tested publish path.
'
' Rohan, 2026-08-14: "delete useless bits incorporate useful bits". Unreachable
' code provides no capability, so it is not a capability being removed. If bulk
' approval is wanted again it comes back deliberately, with a door; git holds it
' at 6c74912.
' THE SLIDES OF ONE TYPE THAT THE REGISTER HAS NO ROW FOR -- as SLIDE OBJECTS,
' not keys, because the caller has to delete them and an index is not a handle:
' deleting one slide renumbers every slide after it.
'
' ReviewQueue counts this same condition (SlideNoRowCount) and keeps only the
' count and the keys, which is right for a report and useless for an action. The
' RULE is duplicated here deliberately and is one line -- Not sheet.Rows.Exists
' -- while what differs is what comes back.
'
' TWO GUARDS THE COUNT DOES NOT IMPLY:
'   - a TEMPLATE slide is never retired. It carries a slide type and no instance
'     key by design, so anything keying off "has a type" would delete the one
'     slide every other slide is cloned from.
'   - a slide with no instance key is never retired. That is an UNCLASSIFIED
'     slide -- title pages, dividers, anything a person added -- and the register
'     having no row for it is not evidence of anything.
' PUBLIC for the test that pins its guards. Those guards are the only thing
' between "the register no longer lists this project" and deleting the slide
' every other slide is cloned from, so they are worth asserting directly rather
' than through the dialog that calls them.
Public Function SlidesWithNoRow(pres As Object, sheet As Sheet, slideType As String) As Collection
    Dim found As Collection
    Set found = New Collection

    Dim sld As Object
    For Each sld In pres.Slides
        Dim inst As SlideInstance
        inst = Resolve.ResolveSlideInstance(sld)
        If inst.HasTypeTag And inst.TypeTag = slideType Then
            If inst.HasInstanceKey And Not inst.IsTemplate Then
                If Not sheet.Rows.Exists(inst.InstanceKey) Then
                    found.Add sld
                End If
            End If
        End If
    Next sld

    Set SlidesWithNoRow = found
End Function

' DECK MEMBERSHIP -- one question with two directions.
'
' Rohan, 2026-08-15: "have slide creation its own function, same for slide
' retirement, its own thing." He is right, and the code already agreed without
' anyone noticing: ReviewQueueSet carries BOTH directions side by side --
' OrphanCount (a register row with no slide) and SlideNoRowCount (a slide the
' register has no row for) -- and ParityText already reports them together.
' Both detections were built. NEITHER action was reachable.
'
' Creation existed and was orphaned: RunSync.CreateMissingSlides was called only
' from SyncNowCore, which was called only from a Private SyncNow that nothing
' called. Its comment said "NO LONGER A BUTTON TARGET. The chain is the entry
' point" -- the chain calls StartQuarter, RollForwardUI, RefreshDraftingSheets,
' marking and discovery, and never this. It had been made Private *specifically
' so the reachability check would not report it*, which is how a genuine orphan
' became invisible to the checker built to find orphans.
'
' WHY NOT INSIDE BUTTON 2. Creating a slide is a write, and the detection lives
' in ReviewChangesCore, whose contract is printed on a button: "Review changes
' (writes nothing)". Folding a write into it would make that caption false.
'
' RETIREMENT DELETES. Rohan chose delete over hide, 2026-08-15: the register is
' the source of truth and last quarter's saved deck is the archive, so hiding
' would grow the deck forever to avoid a loss already covered. The warning names
' every slide by index and key -- this is the one act in the tool that destroys
' content the register cannot restore, because the register is exactly what no
' longer mentions it.
'
' Not retiring was never the safe option. ReviewQueue.bas:1456: "a slide with no
' row keeps last period's text through every sync while every report says the
' run was clean."
' The button target for CAP_REPOINT_WORKBOOK. A wrapper and nothing more --
' DraftingUI.RepointWorkbookUI does the work, shows the current pairing, and
' verifies the new one against the file's own bytes before reporting. It goes
' through RibbonUI because every other button does, and because that is where
' the reachability check looks for a button target.
Public Sub ChangePairedWorkbook()
    On Error GoTo Failed
    DraftingUI.RepointWorkbookUI
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult CommandBarUI.CAP_REPOINT_WORKBOOK, _
        RibbonUI.UnexpectedErrorText(CommandBarUI.CAP_REPOINT_WORKBOOK, Err.Number, Err.Description, Err.Source)
End Sub

' The two halves of deck membership, each reached by its own button so the
' destructive one is never arrived at by momentum. retireMode is the declared
' intent: False = create only, True = delete only. Neither does the other's
' work, and each reports the other's count so nothing goes unseen.
'
' A Boolean rather than named constants deliberately: VBA requires module-level
' Const declarations above every procedure, and a pair declared up at the top of
' a 1700-line module is further from these call sites than the comment is.
' The button target for CAP_DISCOVER_FIELDS. DiscoverUI.DiscoverFields does the
' work and is self-contained; this only gives it a way in that does not require
' the deck to be unconfigured.
Public Sub DiscoverFieldsOnSlide()
    On Error GoTo Failed
    DiscoverUI.DiscoverFields
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult CommandBarUI.CAP_DISCOVER_FIELDS, _
        RibbonUI.UnexpectedErrorText(CommandBarUI.CAP_DISCOVER_FIELDS, Err.Number, Err.Description, Err.Source)
End Sub

' The entry point for the standing formatting-consistency check (see
' FormattingAudit.bas's own header) -- built and tested 2026-08-19 but never
' wired to anything a person can press. Read-only: compares real slides
' against each other, writes nothing.
Public Sub CheckFormatting()
    On Error GoTo Failed
    Dim report As String
    report = FormattingAudit.ScanFormattingOutliers(Application.ActivePresentation)
    RibbonUI.ShowSyncResult CommandBarUI.CAP_CHECK_FORMATTING, report
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult CommandBarUI.CAP_CHECK_FORMATTING, _
        RibbonUI.UnexpectedErrorText(CommandBarUI.CAP_CHECK_FORMATTING, Err.Number, Err.Description, Err.Source)
End Sub

Public Sub AddMissingSlides()
    On Error GoTo Failed
    SlideMembershipCore False
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult CommandBarUI.CAP_ADD_SLIDES, _
        RibbonUI.UnexpectedErrorText(CommandBarUI.CAP_ADD_SLIDES, Err.Number, Err.Description, Err.Source)
End Sub

Public Sub RetireSlides()
    On Error GoTo Failed
    SlideMembershipCore True
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult CommandBarUI.CAP_RETIRE_SLIDES, _
        RibbonUI.UnexpectedErrorText(CommandBarUI.CAP_RETIRE_SLIDES, Err.Number, Err.Description, Err.Source)
End Sub

Private Sub SlideMembershipCore(ByVal retireMode As Boolean)
    Dim TITLE As String
    If retireMode Then
        TITLE = CommandBarUI.CAP_RETIRE_SLIDES
    Else
        TITLE = CommandBarUI.CAP_ADD_SLIDES
    End If

    Dim pres As Object, wb As Object
    Dim types() As String
    Dim lo As Long, hi As Long
    If Not ResolveDeckContext(TITLE, pres, wb, types, lo, hi) Then Exit Sub

    Dim period As String
    period = DeckRegistry.GetDeckPeriod(pres)

    Dim createTypes() As String
    Dim createTemplates() As Object
    Dim createSheets() As Sheet
    Dim createCount As Long
    createCount = 0

    Dim orphanTotal As Long, noRowTotal As Long
    Dim orphanKeys As String, noRowKeys As String
    Dim refusals As String

    Dim retireSlides As Collection
    Set retireSlides = New Collection

    ' Same Sync Log wiring as BuildAllQueuesCore, same reason: this is the
    ' other button-reachable BuildQueue call site, and a crash inside its
    ' build chain (FIX-LIST item V) must leave a record naming the item.
    Dim logWs As Object
    Set logWs = WorkbookBridge.GetOrAddWorksheet(wb, WorkbookBridge.SYNC_LOG_SHEET_NAME)

    Dim i As Long
    For i = lo To hi
        Dim templateSld As Object
        Dim wsName As String
        If DeckRegistry.LookupType(pres, types(i), templateSld, wsName) Then
            Dim ws As Object
            Set ws = WorkbookBridge.GetOrAddWorksheet(wb, wsName)

            Dim sheet As Sheet
            Dim problem As String
            sheet = ExcelOutput.ReadSheetForDeckPeriod(ws, period, problem)

            If problem <> "" Then
                ' NAMED AND COUNTED, never skipped in silence: a type whose sheet
                ' cannot be read has an unknown membership, not a matching one.
                refusals = refusals & "  " & types(i) & ": " & problem & vbCrLf
            Else
                Dim q As ReviewQueueSet
                q = ReviewQueue.BuildQueue(sheet, types(i), logWs)

                If q.OrphanCount > 0 Then
                    orphanTotal = orphanTotal + q.OrphanCount
                    If orphanKeys <> "" Then orphanKeys = orphanKeys & ", "
                    orphanKeys = orphanKeys & q.OrphanKeys

                    createCount = createCount + 1
                    ReDim Preserve createTypes(1 To createCount)
                    ReDim Preserve createTemplates(1 To createCount)
                    ReDim Preserve createSheets(1 To createCount)
                    createTypes(createCount) = types(i)
                    Set createTemplates(createCount) = templateSld
                    createSheets(createCount) = sheet
                End If

                ' COLLECTED AS OBJECTS, BEFORE ANYTHING CHANGES THE DECK.
                ' Creation below inserts slides; a slide INDEX captured now would
                ' point somewhere else by the time deletion runs. An object
                ' reference does not move.
                Dim retireHere As Collection
                Set retireHere = SlidesWithNoRow(pres, sheet, types(i))
                Dim rr As Long
                For rr = 1 To retireHere.count
                    retireSlides.Add retireHere(rr)
                Next rr

                If q.SlideNoRowCount > 0 Then
                    noRowTotal = noRowTotal + q.SlideNoRowCount
                    If noRowKeys <> "" Then noRowKeys = noRowKeys & ", "
                    noRowKeys = noRowKeys & q.SlideNoRowKeys
                End If
            End If
        End If
    Next i

    ' THE OTHER DIRECTION, NAMED BUT NEVER ACTED ON. Declaring intent must not
    ' hide the drift the other button exists for, or the deck grows forever. So
    ' the half that was not asked for is REPORTED, together with the button that
    ' handles it, and never prompted. The scan above already computed both
    ' counts, so this costs nothing.
    Dim otherNote As String
    If retireMode Then
        If orphanTotal > 0 Then
            otherNote = vbCrLf & vbCrLf & orphanTotal & " register row(s) have no slide. Press '" & _
                CommandBarUI.CAP_ADD_SLIDES & "' to create them."
        End If
    Else
        If retireSlides.count > 0 Then
            otherNote = vbCrLf & vbCrLf & retireSlides.count & _
                " slide(s) carry a key the register has no row for. Press '" & _
                CommandBarUI.CAP_RETIRE_SLIDES & "' to deal with them."
        End If
    End If

    If refusals <> "" Then
        refusals = vbCrLf & vbCrLf & "Could not read (membership unknown for these):" & vbCrLf & refusals
    End If

    If orphanTotal = 0 And retireSlides.count = 0 Then
        ShowSyncResult TITLE, "The deck and the register agree for " & period & _
            " -- every row has a slide and every slide has a row." & refusals
        Exit Sub
    End If

    ' NOTHING TO DO IN THE DIRECTION ASKED FOR, said plainly and then stopped --
    ' rather than falling through to a prompt about the half nobody pressed.
    Dim outcome As String

    If retireMode And retireSlides.count = 0 Then
        ShowSyncResult TITLE, "Every slide for " & period & _
            " has a register row -- nothing to retire." & otherNote & refusals
        Exit Sub
    End If

    If (Not retireMode) And orphanTotal = 0 Then
        ShowSyncResult TITLE, "Every register row for " & period & _
            " has a slide -- nothing to add." & otherNote & refusals
        Exit Sub
    End If

    ' ADDING IS ASKED SEPARATELY FROM REMOVING. They are opposite acts with
    ' opposite consequences -- one copies a template, the other destroys work --
    ' and a single "make the deck match" confirmation would buy consent for the
    ' destructive half using the safe half's reasoning.
    If (Not retireMode) And orphanTotal > 0 Then
        If MsgBox(orphanTotal & " register row(s) for " & period & " have no slide:" & vbCrLf & vbCrLf & _
                  "  " & orphanKeys & vbCrLf & vbCrLf & _
                  "Create a slide for each, copied from the template and tagged?" & vbCrLf & vbCrLf & _
                  "Nothing else is touched -- existing slides are not changed by this." & _
                  otherNote & refusals, _
                  vbYesNo + vbQuestion, TITLE) = vbYes Then
            Dim ci As Long
            For ci = 1 To createCount
                outcome = outcome & RunSync.CreateMissingSlides( _
                    createSheets(ci), createTypes(ci), createTemplates(ci), False) & vbCrLf
            Next ci
        Else
            outcome = outcome & "Nothing was created." & vbCrLf
        End If
    End If

    ' RETIRING DELETES SLIDES. Rohan, 2026-08-15: "delete with warning prior."
    '
    ' The warning NAMES EVERY SLIDE, index and key, because this is the one act
    ' in the tool that destroys content a person cannot get back from the
    ' register -- the register is precisely what no longer mentions them. What
    ' they can be got back from is last quarter's saved deck, which is why
    ' delete was the right call over hiding: the archive already exists.
    If retireMode And retireSlides.count > 0 Then
        Dim names As String
        Dim k As Long
        For k = 1 To retireSlides.count
            Dim rInst As SlideInstance
            rInst = Resolve.ResolveSlideInstance(retireSlides(k))
            names = names & "  slide " & retireSlides(k).SlideIndex & " -- " & rInst.InstanceKey & vbCrLf
        Next k

        If MsgBox("DELETE " & retireSlides.count & " slide(s)?" & vbCrLf & vbCrLf & names & vbCrLf & _
                  "The register has no row for these in " & period & ", so the tool treats them as " & _
                  "retired projects." & vbCrLf & vbCrLf & _
                  "THIS DELETES THE SLIDES. They cannot be recovered from the register -- the " & _
                  "register is what no longer mentions them. Last quarter's saved deck is where " & _
                  "they still exist." & vbCrLf & vbCrLf & _
                  "No -- leave them in place and change nothing.", _
                  vbYesNo + vbExclamation, TITLE) = vbYes Then
            Dim deleted As Long
            For k = 1 To retireSlides.count
                retireSlides(k).Delete
                deleted = deleted + 1
            Next k
            outcome = outcome & deleted & " slide(s) retired (deleted)." & vbCrLf
        Else
            outcome = outcome & retireSlides.count & " slide(s) left in place." & vbCrLf
        End If
    End If

    ' SAVED AND VERIFIED FROM THE FILE. Slides created or deleted only in
    ' PowerPoint's memory are the same failure as a value written and not saved.
    outcome = outcome & PersistBothFiles(pres, wb)

    ShowSyncResult TITLE, outcome & refusals
End Sub

' EVERYTHING THAT MUST BE TRUE BEFORE ANY CODE READS THE REGISTER.
'
' Extracted from ReviewChangesCore 2026-08-15, unchanged, because deck
' MEMBERSHIP -- creating a slide for a register row that has none, and later
' retiring a slide the register no longer mentions -- needs exactly the same
' four guards, and PutItOnTheSlidesCore's own header already says why they must
' not be written twice: "Duplicating them here would mean two sets of wording to
' keep true."
'
' The two that matter most are not obvious:
'   - the UNSAVED-BUFFER refusal. Planning from Excel's buffer shows a person an
'     "after" that exists in no file, and they approve values that can still
'     change before the write runs.
'   - R9 duplicate identity tags, checked BEFORE planning, because to the planner
'     two slides sharing a key is indistinguishable from one matched slide and
'     one unmatched one. It is only visible across instances.
'
' Returns False when the caller should stop; every refusal has already been
' explained to the person by the time it returns.
Private Function ResolveDeckContext(title As String, ByRef pres As Object, ByRef wb As Object, _
                                    ByRef types() As String, ByRef lo As Long, ByRef hi As Long) As Boolean
    ResolveDeckContext = False

    Set pres = Application.ActivePresentation

    Dim workbookPath As String
    workbookPath = DeckRegistry.GetWorkbookPath(pres)
    If workbookPath = "" Then
        MsgBox "This deck has no paired workbook yet. Press '" & CommandBarUI.CAP_SET_UP_QUARTER & "' -- it walks setup on a deck that has none.", vbExclamation, title
        Exit Function
    End If

    types = DeckRegistry.ListRegisteredTypes(pres)

    Dim hasTypes As Boolean
    On Error Resume Next
    lo = LBound(types): hi = UBound(types)
    hasTypes = (Err.Number = 0)
    On Error GoTo 0

    If Not hasTypes Then
        MsgBox "This deck has no registered slide types yet. Press '" & CommandBarUI.CAP_SET_UP_QUARTER & "' -- it walks setup on a deck that has none.", vbExclamation, title
        Exit Function
    End If

    Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath)
    If wb Is Nothing Then
        MsgBox "Could not open the paired workbook at: " & workbookPath, vbCritical, title
        Exit Function
    End If

    ' Saves quietly instead of asking -- see WorkbookBridge.EnsureSavedQuietly's
    ' own header for why the prompt this replaced (2026-08-19) no longer needs
    ' to interrupt to do its job. A genuine save failure is still never hidden.
    Dim promptSaveProblem As String
    promptSaveProblem = WorkbookBridge.EnsureSavedQuietly(wb, workbookPath)
    If promptSaveProblem <> "" Then
        MsgBox promptSaveProblem & vbCrLf & vbCrLf & _
               "Stopping here rather than reading values that are not in the file.", _
               vbCritical, title
        Exit Function
    End If

    If Not WarnOnDuplicateKeys(title, types, lo, hi) Then Exit Function

    ResolveDeckContext = True
End Function

Private Sub ReviewChangesCore()
    Dim title As String
    title = CommandBarUI.STAGE_REVIEW_CHANGES

    Dim fullReport As String
    Dim totalQueued As Long
    If Not BuildAllQueuesCore(title, fullReport, totalQueued) Then Exit Sub

    ShowSyncResult title & " (nothing written)", fullReport
End Sub

' THE BUILD HALF OF BOTH "Review changes" AND "2. Put it on the slides",
' extracted 2026-08-18 so the two buttons cannot disagree about what got
' queued -- PutItOnTheSlidesCore's own header already explains why guards
' must not be written twice ("two sets of wording to keep true"), and this is
' the same rule applied to the queue-building loop itself.
'
' Returns False when the caller should stop; every refusal has already been
' explained to the person by the time it returns (same contract as
' ResolveDeckContext above).
Private Function BuildAllQueuesCore(title As String, ByRef fullReport As String, ByRef totalQueued As Long) As Boolean
    BuildAllQueuesCore = False
    fullReport = ""
    totalQueued = 0

    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim workbookPath As String
    workbookPath = DeckRegistry.GetWorkbookPath(pres)
    If workbookPath = "" Then
        MsgBox "This deck has no paired workbook yet. Press '" & CommandBarUI.CAP_SET_UP_QUARTER & "' -- it walks setup on a deck that has none.", vbExclamation, title
        Exit Function
    End If

    Dim types() As String
    types = DeckRegistry.ListRegisteredTypes(pres)

    Dim lo As Long, hi As Long, hasTypes As Boolean
    On Error Resume Next
    lo = LBound(types): hi = UBound(types)
    hasTypes = (Err.Number = 0)
    On Error GoTo 0

    If Not hasTypes Then
        MsgBox "This deck has no registered slide types yet. Press '" & CommandBarUI.CAP_SET_UP_QUARTER & "' -- it walks setup on a deck that has none.", vbExclamation, title
        Exit Function
    End If

    Dim wb As Object
    Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath)
    If wb Is Nothing Then
        MsgBox "Could not open the paired workbook at: " & workbookPath, vbCritical, title
        Exit Function
    End If

    ' IS THIS EVEN OUR REGISTER? Same check DraftingUI.Resolve now makes for
    ' its own callers, added 2026-08-19 -- this is the register-diff queue
    ' build behind "Review changes" and half of "Put it on the slides", and
    ' until now nothing here asked whether the workbook just opened actually
    ' belongs to this deck. See DeckRegistry's "THE PAIRING, BOTH WAYS" for
    ' the full history: the check existed, but only the AI-drafting publish
    ' path ever called it.
    Dim pairNote As String
    pairNote = DeckRegistry.PairingProblem(pres, wb)
    If pairNote <> "" Then
        MsgBox pairNote, vbCritical, title
        Exit Function
    End If

    ' Refuse to build a queue out of Excel's unsaved buffer -- see
    ' WorkbookBridge.IsDirty for the live incident. Checked BEFORE planning, not
    ' after: the plan reads the sheet, so a queue built on unsaved data would
    ' show a human before-and-afters whose "after" exists in no file, and they
    ' would be approving values that could still change before Apply runs.
    ' Saves quietly instead of asking -- see WorkbookBridge.EnsureSavedQuietly's
    ' own header for why the prompt this replaced (2026-08-19) no longer needs
    ' to interrupt to do its job. A genuine save failure is still never hidden.
    Dim promptSaveProblem As String
    promptSaveProblem = WorkbookBridge.EnsureSavedQuietly(wb, workbookPath)
    If promptSaveProblem <> "" Then
        MsgBox promptSaveProblem & vbCrLf & vbCrLf & _
               "Stopping here rather than reading values that are not in the file.", _
               vbCritical, title
        Exit Function
    End If

    ' R9: duplicate identity tags, checked BEFORE planning. Kept from
    ' SyncNowCore unchanged -- the planner cannot report this usefully, because
    ' to PlanRoutineSync two slides sharing a key is indistinguishable from one
    ' matched slide and one unmatched one. It is only visible across instances.
    If Not WarnOnDuplicateKeys(title, types, lo, hi) Then Exit Function

    ' The Sync Log, resolved here for the same reason ApplyApprovedCore
    ' resolves it: Error 50290's FOURTH occurrence (2026-08-19, FIX-LIST item
    ' V) landed in THIS build phase -- before any write ever reached the Sync
    ' Log -- so a crash during the build must now leave a record naming what
    ' was mid-flight, and it must survive even if PowerPoint dies before any
    ' dialog appears.
    Dim logWs As Object
    Set logWs = WorkbookBridge.GetOrAddWorksheet(wb, WorkbookBridge.SYNC_LOG_SHEET_NAME)

    Dim firstSheet As Object
    Dim stageErrNum As Long, stageErrDesc As String, stageErrSrc As String

    Dim i As Long
    For i = lo To hi
        Dim templateSld As Object
        Dim wsName As String
        If DeckRegistry.LookupType(pres, types(i), templateSld, wsName) Then
            Dim ws As Object
            Set ws = WorkbookBridge.GetOrAddWorksheet(wb, wsName)

            ' PER-STAGE CRASH CAPTURE around the chain's two heavy Excel-side
            ' stages (this read, and WriteQueueSheet below) -- the per-slide/
            ' per-field granularity lives one layer down, in BuildQueue and
            ' PlanRoutineSync's own traps. See ReviewQueue.LogAndReraiseCrash.
            Dim sheet As Sheet
            Dim problem As String
            On Error Resume Next
            Err.Clear
            sheet = ExcelOutput.ReadSheetForDeckPeriod(ws, DeckRegistry.GetDeckPeriod(pres), problem)
            stageErrNum = Err.Number: stageErrDesc = Err.Description: stageErrSrc = Err.Source
            On Error GoTo 0
            If stageErrNum <> 0 Then
                ReviewQueue.LogAndReraiseCrash logWs, "", "RibbonUI.BuildAllQueuesCore", "", "", _
                    "reading register sheet '" & wsName & "' for type '" & types(i) & "'", _
                    stageErrNum, stageErrDesc, stageErrSrc
            End If

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
                q = ReviewQueue.BuildQueue(sheet, types(i), logWs)
                totalQueued = totalQueued + q.Count

                Dim reviewWs As Object
                Set reviewWs = WorkbookBridge.GetOrAddWorksheet(wb, ReviewQueue.ReviewSheetNameFor(types(i)))
                On Error Resume Next
                Err.Clear
                ReviewQueue.WriteQueueSheet reviewWs, q
                stageErrNum = Err.Number: stageErrDesc = Err.Description: stageErrSrc = Err.Source
                On Error GoTo 0
                If stageErrNum <> 0 Then
                    ReviewQueue.LogAndReraiseCrash logWs, q.RunStamp, "RibbonUI.BuildAllQueuesCore", "", "", _
                        "writing review sheet for type '" & types(i) & "'", _
                        stageErrNum, stageErrDesc, stageErrSrc
                End If
                If firstSheet Is Nothing Then Set firstSheet = reviewWs

                fullReport = fullReport & "=== " & types(i) & " ===" & vbCrLf & _
                    ReviewQueue.QueueSummaryText(q) & vbCrLf
            End If
        Else
            fullReport = fullReport & "SKIPPED " & types(i) & ": registered type's template slide no longer resolves (was it deleted?)" & vbCrLf
        End If
    Next i

    ' Bring the review sheet to the front. Leaving the human to go and find it
    ' is how a review becomes optional in practice while remaining mandatory on
    ' paper -- the same distinction R13 is about.
    If Not firstSheet Is Nothing And totalQueued > 0 Then
        On Error Resume Next
        firstSheet.Activate
        wb.Activate
        On Error GoTo 0
    End If

    BuildAllQueuesCore = True
End Function

' Toolbar entry point. The real work is in ApplyApprovedCore; this exists only
' to catch anything that escapes it. Same separate-frame reasoning as above.
' ---------------------------------------------------------------------
' THE CHAIN -- the workbook side, captioned CAP_SET_UP_QUARTER
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
' The shape-first door. Offers to tag whatever single UNTAGGED shape is
' selected, whether or not the register knows about it yet -- which is the case
' the register-first route could never cover.
Private Function OfferMarkingForSelectedShape(TITLE As String) As Boolean
    OfferMarkingForSelectedShape = True

    Dim shp As Object
    On Error Resume Next
    If Application.ActiveWindow.Selection.Type = 2 Then          ' ppSelectionShapes
        If Application.ActiveWindow.Selection.ShapeRange.Count = 1 Then
            Set shp = Application.ActiveWindow.Selection.ShapeRange(1)
        End If
    End If
    On Error GoTo 0
    If shp Is Nothing Then Exit Function

    ' A GROUP IS STILL WORTH OFFERING -- the picker opens it and lists what is
    ' inside, which is exactly how the timeline's nine shapes get reached. Only
    ' an already-tagged single shape is skipped.
    Dim existingRole As String
    existingRole = ""
    On Error Resume Next
    ' READ FOR A GROUP TOO, since 2026-08-10. A group could not carry a role
    ' tag before then -- WalkForRoleTag stepped past it -- so excluding groups
    ' here was free. Now a DEVICE is a tagged group, and skipping the read would
    ' offer to tag a timeline that is already tagged, every single press.
    existingRole = shp.Tags("role")
    On Error GoTo 0
    If existingRole <> "" Then Exit Function

    If MsgBox("'" & shp.Name & "' is selected and is not tracked as a field yet." & vbCrLf & vbCrLf & _
              "Tag it now?" & vbCrLf & vbCrLf & _
              "Yes -- tag this shape." & vbCrLf & _
              "No  -- leave it and carry on with the sync.", _
              vbYesNo + vbQuestion, TITLE) = vbYes Then
        BatchOnboardFlow.MarkFieldForBatch
        OfferMarkingForSelectedShape = False
    End If
End Function

' Returns True to CARRY ON with the chain, False to stop.
'
' READING VALUES OFF THE SLIDES INTO THE REGISTER. See Harvest.bas's header for
' why this is a separate capability from adoption rather than a mode of it:
' adoption skips every already-linked slide before it matches anything, so on a
' live deck -- where every slide is linked -- it harvests nothing at all.
'
' NOT A FOURTH BUTTON, and not an unconditional prompt either. In Normal view a
' slide is ALWAYS selected, so "you have slides selected, harvest them?" would
' fire on nearly every press, which is the invariant prompt this chain deleted
' six of. The gate is a DRY RUN over the selected slides: if nothing would be
' written, nothing is said. On a steady-state deck every field already holds a
' value, so this is silent; it speaks only in the state it exists for, which is
' the press after new fields have been tagged.
'
' The dry run costs one register read plus a shape walk per field per selected
' slide. In Normal view that is one slide. In the slide sorter it is whatever
' the person deliberately selected, which is the moment they have asked for it.
Private Function OfferHarvestForSelectedSlides(pres As Object, TITLE As String) As Boolean
    OfferHarvestForSelectedSlides = True

    Dim sel As Object
    On Error Resume Next
    Set sel = Application.ActiveWindow.Selection
    On Error GoTo 0
    If sel Is Nothing Then Exit Function

    Dim isSlides As Boolean
    isSlides = False
    On Error Resume Next
    isSlides = (sel.Type = 1)                                ' ppSelectionSlides
    On Error GoTo 0
    If Not isSlides Then Exit Function

    Dim period As String
    period = DeckRegistry.GetDeckPeriod(pres)
    If period = "" Then Exit Function

    Dim wb As Object
    On Error Resume Next
    Set wb = WorkbookBridge.OpenOrGetWorkbook(DeckRegistry.GetWorkbookPath(pres))
    On Error GoTo 0
    If wb Is Nothing Then Exit Function

    ' IS THIS EVEN OUR REGISTER? Same check DraftingUI.Resolve now makes for
    ' its own callers, added 2026-08-19. This is the OTHER direction --
    ' slide content written INTO the register -- so a mismatch here means
    ' this project's real numbers land in a stranger's workbook, not just
    ' the reverse.
    Dim pairNote As String
    pairNote = DeckRegistry.PairingProblem(pres, wb)
    If pairNote <> "" Then
        MsgBox pairNote, vbCritical, TITLE
        Exit Function
    End If

    ' PASS ONE -- dry run. Writes nothing, to the deck or the register. Its only
    ' job is to decide whether there is anything worth interrupting for, and to
    ' say exactly what.
    Dim toStamp As Long, toRead As Long, slideCount As Long
    Dim detail As String, devices As String, collisions As String

    Dim sld As Object
    For Each sld In sel.SlideRange
        Dim ws As Object, tpl As Object
        Set ws = SheetForSlide(pres, wb, sld)
        Set tpl = TemplateForSlide(pres, sld)

        Dim slideNote As String
        slideNote = ""

        If Not tpl Is Nothing Then
            Dim pDry As PropagateOutcome
            pDry = Harvest.PropagateTemplateTags(sld, tpl, True)
            If pDry.Ran And pDry.Stamped > 0 Then
                toStamp = toStamp + pDry.Stamped
                slideNote = slideNote & pDry.Detail
            End If
            ' FROM Collisions, NOT Detail. Detail is what WOULD be stamped; mixing
            ' the two is what printed 16 successful stamps under a "Refused" header.
            If pDry.Collided > 0 Then _
                collisions = collisions & "Slide " & sld.SlideIndex & ":" & vbCrLf & pDry.Collisions
        End If

        If Not ws Is Nothing Then
            ' pDry.Pending IS THE FIX for the count. The dry harvest finds
            ' fields by their role tag, and in a dry run propagation has not
            ' written those tags yet -- so without this it counts only fields
            ' that were ALREADY tagged and omits every one this same press is
            ' about to label and read. Measured 2026-08-15: offered 10, wrote 34.
            Dim pending As Object
            Set pending = Nothing
            If Not tpl Is Nothing Then Set pending = pDry.Pending

            Dim dry As HarvestOutcome
            dry = Harvest.HarvestSlide(sld, ws, period, True, pending)
            If dry.Ran And dry.Written > 0 Then
                toRead = toRead + dry.Written
                slideNote = slideNote & dry.Detail
            End If
            If dry.Ran And dry.SkippedNotText > 0 Then devices = dry.Detail
        End If

        If slideNote <> "" Then
            slideCount = slideCount + 1
            detail = detail & "Slide " & sld.SlideIndex & ":" & vbCrLf & slideNote
        End If
    Next sld

    If toStamp = 0 And toRead = 0 Then Exit Function

    ' NAMES BOTH WRITES. This is the one approval in front of an operation that
    ' changes the DECK as well as the register, and a prompt that mentioned only
    ' the register would be asking for consent to less than it does.
    Dim ask As String
    ask = toStamp & " shape(s) would be labelled on the slides, and " & toRead & _
          " value(s) read into the register, across " & slideCount & " slide(s) for " & period & "." & vbCrLf & vbCrLf & _
          "This writes to BOTH files:" & vbCrLf & _
          "  - the deck gets a role tag on each labelled shape (nothing visible changes)" & vbCrLf & _
          "  - the register gets a value ONLY where it currently holds nothing" & vbCrLf & vbCrLf & _
          "A value already in the register is never overwritten." & vbCrLf & vbCrLf
    If collisions <> "" Then ask = ask & "Refused -- two fields matched one shape:" & vbCrLf & collisions & vbCrLf
    ask = ask & detail

    ' THE FULL PLAN GOES TO THE RUN LOG BEFORE THE DIALOG IS CAPPED, because
    ' CapReport's notice SAYS the full list is there. Writing the notice without
    ' writing the log would be a dialog telling a person where to look for
    ' something that is not there -- worse than truncating silently, since they
    ' would go and check.
    If Not wb Is Nothing Then
        WorkbookBridge.WriteRunLog wb, "Harvest plan for " & period, _
            IIf(collisions = "", "", "REFUSED -- two fields matched one shape:" & vbCrLf & collisions & vbCrLf) & detail
    End If

    ' CAPPED AS A WHOLE, not just the detail. Capping one part while another grew
    ' unbounded is what let MsgBox cut a collision line mid-word with no notice at
    ' all -- VBA's MsgBox truncates SILENTLY, so the cap has to cover everything
    ' that reaches it.
    If MsgBox(CapReport(ask), vbYesNo + vbQuestion, TITLE) <> vbYes Then
        OfferHarvestForSelectedSlides = False
        Exit Function
    End If

    ' PASS TWO -- the same walks, writing. LABEL FIRST, THEN READ, per slide:
    ' the harvest finds a field BY its role tag, so a tag stamped after the read
    ' would be a field labelled this run and not harvested until the next one --
    ' a half-done state that reports as success.
    '
    ' Re-derived rather than replayed from pass one: if anything changed between
    ' the two, each pass's own guard still governs, so the worst case is doing
    ' less than was offered, never more.
    Dim stamped As Long, written As Long
    For Each sld In sel.SlideRange
        Set tpl = TemplateForSlide(pres, sld)
        If Not tpl Is Nothing Then
            Dim pWet As PropagateOutcome
            pWet = Harvest.PropagateTemplateTags(sld, tpl, False)
            stamped = stamped + pWet.Stamped
        End If

        Set ws = SheetForSlide(pres, wb, sld)
        If Not ws Is Nothing Then
            Dim wet As HarvestOutcome
            wet = Harvest.HarvestSlide(sld, ws, period, False)
            written = written + wet.Written
        End If
    Next sld

    Dim report As String
    report = stamped & " shape(s) labelled, " & written & " value(s) written into the register for " & period & "."
    report = report & vbCrLf & PersistBothFiles(pres, wb)
    If devices <> "" Then
        report = report & vbCrLf & "Not read (no way to read one back yet):" & vbCrLf & devices
    End If

    ShowSyncResult TITLE, report
    OfferHarvestForSelectedSlides = False
End Function

' The TEMPLATE slide for this slide's own type. Separate from SheetForSlide
' because propagation needs the template and harvesting needs the sheet, and a
' slide can have one resolvable without the other.
Private Function TemplateForSlide(pres As Object, sld As Object) As Object
    Dim inst As SlideInstance
    inst = Resolve.ResolveSlideInstance(sld)
    If Not inst.HasTypeTag Then Exit Function
    If inst.IsTemplate Then Exit Function          ' never propagate onto itself

    Dim templateSld As Object, wsName As String
    If Not DeckRegistry.LookupType(pres, inst.TypeTag, templateSld, wsName) Then Exit Function

    Set TemplateForSlide = templateSld
End Function

' The register sheet for THIS slide's own type. Per-slide rather than resolved
' once for the selection: a selection can span types, and picking one sheet for
' all of them would write a slide's values into another type's rows.
Private Function SheetForSlide(pres As Object, wb As Object, sld As Object) As Object
    Dim inst As SlideInstance
    inst = Resolve.ResolveSlideInstance(sld)
    If Not inst.HasTypeTag Then Exit Function

    Dim templateSld As Object, wsName As String
    If Not DeckRegistry.LookupType(pres, inst.TypeTag, templateSld, wsName) Then Exit Function
    If Not WorkbookBridge.WorksheetExists(wb, wsName) Then Exit Function

    Set SheetForSlide = WorkbookBridge.GetOrAddWorksheet(wb, wsName)
End Function

' Returns True to CARRY ON with the chain, False to stop.
'
' The SLIDE-first door, mirroring OfferMarkingForSelectedShape above.
'
' AdoptFlow.AdoptExistingSlides was written AS a toolbar entry point -- its own
' header still said so -- and the 2026-08-14 split to three buttons left it with
' no button and no caller anywhere: built, tested and unreachable by a person.
' check_vba_static.py could not see it because AdoptFlow.bas was missing from
' its UI_MODULES set, so the checker built to catch exactly this reported clean.
' Both are fixed together; keeping one without the other re-opens the hole.
'
' OFFERED, NOT BUTTONED. The bar is deliberately three, and adoption needs a
' SELECTION to act on, so the selection is the trigger rather than a fourth
' caption. Slide selection and shape selection are different Selection.Types,
' so this cannot collide with the door above.
'
' IT FIRES ONLY WHEN THERE IS SOMETHING TO ADOPT, and that condition is load
' bearing rather than tidy. In Normal view a slide is ALWAYS selected, so
' prompting on "slides are selected" would put a dialog in front of nearly
' every press -- the invariant prompt this chain deleted six of. An already
' linked slide is skipped by PlanAdoption before any matching, so a selection
' containing only linked slides has genuinely nothing to offer.
Private Function OfferAdoptionForSelectedSlides(TITLE As String) As Boolean
    OfferAdoptionForSelectedSlides = True

    Dim sel As Object
    On Error Resume Next
    Set sel = Application.ActiveWindow.Selection
    On Error GoTo 0
    If sel Is Nothing Then Exit Function

    Dim isSlides As Boolean
    isSlides = False
    On Error Resume Next
    isSlides = (sel.Type = 1)                                ' ppSelectionSlides
    On Error GoTo 0
    If Not isSlides Then Exit Function

    ' COUNT THE UNLINKED ONES, not the selected ones. Resolve.ResolveSlideInstance
    ' is the same read PlanAdoption uses to decide, so this offer cannot promise
    ' work that the adoption itself will then skip.
    Dim unlinked As Long
    unlinked = 0
    Dim sld As Object
    On Error Resume Next
    For Each sld In sel.SlideRange
        Dim inst As SlideInstance
        inst = Resolve.ResolveSlideInstance(sld)
        If Not (inst.HasInstanceKey And inst.HasTypeTag) Then unlinked = unlinked + 1
    Next sld
    On Error GoTo 0
    If unlinked = 0 Then Exit Function

    If MsgBox(unlinked & " of the selected slides are not linked to the register yet." & vbCrLf & vbCrLf & _
              "Adopt them now?" & vbCrLf & vbCrLf & _
              "Yes -- read their fields against the template and link them to rows." & vbCrLf & _
              "No  -- leave them and carry on.", _
              vbYesNo + vbQuestion, TITLE) = vbYes Then
        AdoptFlow.AdoptExistingSlides
        OfferAdoptionForSelectedSlides = False
    End If
End Function

' Returns True to CARRY ON with the chain, False to stop.
'
' Stops only when the person chose to go and tag something, or cancelled. A
' scan that cannot run does NOT stop the chain: refusing to sync because a
' check was unable to look would be the check gating rather than offering, and
' Readiness.bas:51 governs the whole design -- it offers, it does not gate.
Private Function OfferMarkingForUnwiredFields(pres As Object, TITLE As String) As Boolean
    OfferMarkingForUnwiredFields = True

    ' Collected and reported, never prompted. See the block below for why.
    Dim unwiredNote As String

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
                            Or wiring.TemplateUnmarkedCount > 0 Or wiring.OrphanCount > 0 _
                            Or wiring.CaseMismatchCount > 0) Then
                        ' NAMES THE FIELDS, NOT JUST A COUNT. Fix-list 1a: a
                        ' true count with no subject sends people to check the
                        ' wrong thing, four times over now.
                        ' THE PROMPT IS GONE (2026-08-14). Its answer never varied.
                        '
                        ' It fired on EVERY press with the same 21 names --
                        ' MS1_LABEL..MS7_DONE -- which are not orphans at all.
                        ' They are the internals of the MILESTONE_TIMELINE
                        ' device, which injects as ONE addressable thing. The
                        ' only correct answer was always "No, sync anyway", and
                        ' answering "Yes" would have walked a person through
                        ' tagging 21 device internals as individual fields --
                        ' destroying the device this warning was pointing at.
                        '
                        ' A dialog whose answer is fixed is not a decision, it
                        ' is a toll. Worse, it trains the click-through that
                        ' eventually gets paid on the ONE dialog that matters.
                        ' Reported to the run report and the START HERE sheet
                        ' instead, which is already this file's stated policy
                        ' for partial coverage twelve lines above.
                        '
                        ' The tagging entry point is NOT lost -- it lives on
                        ' Discover Fields, which is where a person goes when
                        ' they mean to tag something, rather than mid-sync when
                        ' they meant to publish.
                        ' COUNTS ONLY (FieldWiring.CoverageSummaryLine), NOT
                        ' BlockingText's full field/slide detail -- Rohan,
                        ' 2026-08-19: "folded [into this shared dialog], I
                        ' think", choosing this over un-folding back to a
                        ' separate dialog once the two turned out not to fit
                        ' together (see the header above
                        ' OfferMarkingForUnwiredFields's caller for the full
                        ' story: the OTHER three sections in this combined
                        ' summary already use most of the shared 900-char
                        ' budget on their own).
                        unwiredNote = unwiredNote & "Slide type '" & types(i) & "': " & _
                            FieldWiring.CoverageSummaryLine(wiring) & vbCrLf
                    End If
                End If
            End If
        End If
    Next i

    ' NOT ALSO LOGGED HERE. WorkbookBridge.WriteRunLog clears the WHOLE
    ' sheet on every call ("REPLACED each run, not appended" -- its own
    ' header), and this chain calls it several more times after this point
    ' (drafting sheet rebuild, Field Spec/Sources validation, Lobby
    ' rebuild) -- confirmed live 2026-08-19: a call here was gone from the
    ' saved file by the time the run finished, silently, because a LATER
    ' call in the same chain overwrote it first. Claiming "full detail is
    ' in the Run Log" here would have been a promise this code cannot
    ' keep -- the modal below is the only place this information is ever
    ' actually shown, so it is written to say exactly that.

    ' A NOTICE HERE AGAIN, DELIBERATELY REVERSING THE 2026-08-14 MODAL
    ' REMOVAL -- see that date's comment, still directly above this
    ' function's own definition, for the full reasoning that removed it.
    '
    ' That removal was correct for what it fixed: the modal fired on EVERY
    ' press with the SAME 21 names (MS1_LABEL..MS7_DONE), which were never
    ' real gaps -- they are the milestone device's internal parts, asked
    ' about individually when they are addressed as one device. A dialog
    ' whose answer never varies is a toll, not a decision, and it trains the
    ' click-through that eventually gets paid on the one press that matters.
    '
    ' THAT FALSE-POSITIVE SOURCE IS NOW FIXED, not just avoided. ScanFieldWiring
    ' buckets device-owned columns into DeviceOwnedCount and excludes them from
    ' Unmarked/TemplateUnmarked entirely (MilestoneDevice.IsColumnForThisDevice's
    ' gate, above) -- confirmed by reading the function, not assumed by
    ' analogy. So `unwiredNote` reaching here is a genuine structural gap
    ' (a real field the register expects that the template or some slides do
    ' not carry), not device noise -- the exact distinction the killed
    ' modal never made. Rohan, 2026-08-19: "if the excel register is not
    ' complete as per slide type we actually get a notification for which
    ' field and which slides."
    '
    ' FOLDED INTO THE CHAIN'S ONE COMBINED DIALOG, NOT A SEPARATE MsgBox.
    ' A standalone MsgBox here was the first version of this fix, and it
    ' was wrong on two counts, both found live the same night: (1) it
    ' pushed "Set up my quarter" from LOBBY-DESIGN.md section 6's documented
    ' ~2-modal target back up to 3; (2) AppEvents.cls's mApp_SheetChange
    ' handler runs INSIDE this PowerPoint VBA project (WithEvents on
    ' Excel.Application, but the handler code is PowerPoint's), so ANY
    ' modal left open here blocks PowerPoint's single VBA thread from
    ' servicing that handler -- editing an Approve-column tick in Excel
    ' while this dialog sat open would hit the exact "source application
    ' may be busy" stall Rohan saw. DraftingUI.AppendCollected folds this
    ' into the SAME summary StartQuarter/RollForwardUI/RefreshDraftingSheets
    ' already build (moved to run inside that collecting window -- see
    ' SyncNowChainCore), same technique those three already use to become
    ' one dialog instead of three.
    '
    ' ONE MODAL FOR EVERYTHING, not one per slide type -- `unwiredNote`
    ' already accumulates across the whole loop above before anything is
    ' shown, same reasoning BlockingText's own per-type labelling relies on.
    ' Still never blocks: OfferMarkingForUnwiredFields returns True either
    ' way, same as before this change -- this is a notification, not a gate.
    If unwiredNote <> "" Then
        ' REUSES THE ESTABLISHED CAP, NOT A NEW ONE. CapReport is this
        ' project's own "ONE PLACE THAT KNOWS ABOUT THE LIMIT" (its own
        ' header: fixed this exact MsgBox-truncates-near-1024-characters
        ' defect four times already before settling here). A first version
        ' of this fix invented a separate line-based cap instead of
        ' checking for this one first -- found and corrected the same
        ' night, live: 20 lines still truncated mid-word, because
        ' CapReport's own header already explains why (character budget,
        ' not line count). Reusing it instead of a second, slightly
        ' different answer to the same question.
        '
        ' NOTICE OVERRIDDEN, not the default. CapReport's built-in notice
        ' says "the full list is on the Run Log sheet" -- true for its
        ' other callers, false for this one (see the comment above this
        ' function's own WriteRunLog removal): nothing this function
        ' writes survives to the saved file, because later calls in the
        ' same chain overwrite the sheet first.
        '
        ' EVEN A 350-CHAR CAP ON THIS SECTION WASN'T ENOUGH, live-confirmed
        ' the same night: the other three sections in this combined dialog
        ' already use most of the shared 900-char budget on their own, so
        ' the OUTER CapReport call (SyncNowChainCore) chopped straight
        ' through this section regardless of how tightly it was capped on
        ' its own -- the cut isn't section-aware. Rohan's call, given that:
        ' names removed (FieldWiring.CoverageSummaryLine instead of
        ' BlockingText -- see the header above where unwiredNote is built),
        ' counts kept. A summary this short should never need the cap
        ' below in practice; it stays on as a safety net, not the primary
        ' defence, for a deck registering many slide types at once.
        DraftingUI.AppendCollected _
            "The register is not fully wired for this deck:" & vbCrLf & vbCrLf & _
            CapReport(unwiredNote, "", "[shortened -- not every gap is listed here]", 200), _
            "Field Coverage"
    End If
End Function

Private Sub SyncNowChainCore()
    ' FROM THE CAPTION CONSTANT, NOT TYPED. This was the literal "1. Sync Now"
    ' and it survived the rename, so the first press of addin87 put a dialog on
    ' screen titled after a button that no longer exists. Exactly the class this
    ' project keeps paying for: a machine-knowable fact written a second time.
    Dim TITLE As String
    TITLE = CommandBarUI.CAP_SET_UP_QUARTER

    ' FIX-LIST P1. Every dialog below this point is a bare MsgBox/InputBox
    ' owned by PowerPoint's own process, and nothing else in this chain raises
    ' PowerPoint's window -- see DraftingUI.BringPowerPointToFront for the
    ' full incident. Called again below, after RollForwardUI, because that
    ' call deliberately (and correctly, for its own picker) raises Excel and
    ' undoes this.
    DraftingUI.BringPowerPointToFront

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

        ' THE QUARTER MUST BE SET BEFORE ONBOARDING, NOT AFTER.
        '
        ' Onboarding stamps every row it writes with the deck's period, and
        ' ExcelOutput.UpsertRow REFUSES a blank one on a sheet that has a
        ' Quarter column -- correctly, since such a row is invisible to every
        ' filtered read and would report as a clean sync of nothing.
        '
        ' But StartQuarter lived further down the chain, past the `hasTypes`
        ' exit above, so a deck being set up for the FIRST TIME could never
        ' reach it. The result was walking the entire marking grid and then
        ' hitting a raw error at the commit -- with slides already tagged, so
        ' the deck is left half-onboarded and the failure looks like the
        ' marking's fault rather than a missing period.
        '
        ' WORKFLOW.md flagged this ordering problem on 2026-08-04 and it was
        ' never fixed because the toolbar still had a step 0 button then. When
        ' the toolbar went to two buttons, the only way to set a period on a
        ' virgin deck went with it.
        If setupAnswer = vbYes Or setupAnswer = vbNo Then
            DraftingUI.StartQuarter
            If DeckRegistry.GetDeckPeriod(pres) = "" Then
                MsgBox "No quarter was set, so setup stopped before anything was tagged." & vbCrLf & vbCrLf & _
                       "Onboarding writes a register row per slide and every row has to say " & _
                       "which quarter it belongs to.", vbExclamation, TITLE
                Exit Sub
            End If
        End If

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

    ' "WHERE YOU ARE" REMOVED ENTIRELY, 2026-08-17 evening -- was folded in
    ' here per Rohan 2026-08-09, deleted per Rohan the same night this
    ' comment was rewritten. Every check it ran (deck period set, workbook
    ' paired, template slide present, partial-period leftover rows) is
    ' independently caught, cheaply, by the real operations below the
    ' moment they actually run and refuse with their own clear message --
    ' see RunSync.bas's own "REFUSED: this slide type has no template slide
    ' registered" for one concrete example. What it was NOT redundant with
    ' (a period reported-as-set but never actually saved) is already
    ' verified at the point of WRITING, by SetDeckPeriodVerified inside
    ' StartQuarter, called moments later in this same chain -- re-deriving
    ' it again from disk here was checking something already proven true.
    ' It paid for that redundancy at real cost: TWO full copies of the
    ' entire deck file plus slow Shell.Application ZIP extraction, and a
    ' full ReviewQueue.BuildQueue diff per registered slide type, on EVERY
    ' single press of this chain's own button -- all to produce one line of
    ' throwaway status text. Rohan: "just because it can be done doesn't
    ' mean it should be." `Readiness.bas` and `RibbonUI.WhereAmI`/
    ' `WhereAmICore` are deleted along with this call.

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
    ' TAGGING FROM THE SLIDE, which is the direction a person actually works in.
    '
    ' Rohan, 2026-08-10: "that's an obscure use." He was right. The only route to
    ' marking ran through the register being AHEAD of the deck -- to tag a shape
    ' you first had to invent a column for it. The natural direction is the
    ' opposite: point at the thing on the slide and say "track this".
    '
    ' Fires only on a single shape that carries NO role tag, so selecting a
    ' tagged field to look at it does not prompt, and neither does having nothing
    ' selected. Answering No carries straight on to the sync.
    If Not OfferMarkingForSelectedShape(TITLE) Then Exit Sub

    If Not OfferAdoptionForSelectedSlides(TITLE) Then Exit Sub

    ' AFTER adoption, deliberately: adoption LINKS a slide and harvest FILLS a
    ' linked one, so on a selection containing both, linking first is the order
    ' that leaves nothing stranded. Adoption stops the chain when it runs, so
    ' the two never fire on the same press -- pressing again picks up the rest.
    If Not OfferHarvestForSelectedSlides(pres, TITLE) Then Exit Sub

    ' THE PLAN USED TO BE A MODAL HERE, AND IT IS GONE. 2026-08-14.
    '
    ' It listed the five stages and asked "Go ahead?" -- pressing the button IS
    ' going ahead, so the answer never varied, and a prompt whose answer never
    ' varies is how the one that matters gets clicked past. Nothing reaches a
    ' slide before the review tick regardless, which is the property this was
    ' claiming to protect and the tick actually protects.
    '
    ' The ORIENTATION was real and does not die with the dialog -- "which step am
    ' I up to" is a genuine need. It belongs on START HERE, where it can be read
    ' before starting rather than dismissed while starting.

    ' ONE REPORT FOR THE WHOLE PROLOGUE. Each stage's decisions still stop and
    ' ask; only its informational messages are collected. A stage with nothing
    ' to do now says so in the report instead of interrupting to say it.
    ' THE WORKBOOK SIDE, AND ONLY THE WORKBOOK SIDE.
    '
    ' Split by ARTIFACT, not by step, and NEITHER SIDE TRIGGERS THE OTHER --
    ' Rohan's call, 2026-08-14, reversing the 2026-08-09 single-chain decision.
    ' The chain was right about the problem (orientation: "which step am I up
    ' to?") and wrong about the remedy (removing the choice). The coupling it
    ' created is what wiped 43 approve ticks on 2026-08-14: a person pressed a
    ' button to PUBLISH and it REBUILT first.
    '
    ' The boundary is WHERE YOU STOP TYPING. Everything here happens before a
    ' word is written; everything on button 2 happens after. So this ends with
    ' the drafting sheets in front of you, and stops.
    '
    ' CopyAiDraftsToSubmit is NO LONGER HERE. It ran at this point, before
    ' Copilot had written a single draft, and asked which field to copy drafts
    ' from -- a question about something that does not exist yet. It belongs on
    ' the other side of the typing, and that is where it now runs.
    DraftingUI.BeginCollecting
    DraftingUI.StartQuarter
    DraftingUI.RollForwardUI
    ' RollForwardUI just raised Excel for its own range picker -- undo that
    ' before the next dialog in the chain, which is PowerPoint's. FIX-LIST P1.
    DraftingUI.BringPowerPointToFront
    DraftingUI.RefreshDraftingSheets

    ' MOVED HERE, INSIDE THE COLLECTING WINDOW, 2026-08-19 -- see this
    ' function's own header for why (folds into one dialog instead of a
    ' separate MsgBox, and closes the cross-app busy-block window). Runs
    ' after the sheets refresh rather than before OfferHarvest above for no
    ' functional reason (ScanFieldWiring reads slide tags and register
    ' column names, neither of which RefreshDraftingSheets changes) --
    ' placed last so the completeness check reads as the final word on
    ' "is this quarter's setup actually done."
    '
    ' EndCollecting CALLED ON THIS EXIT TOO, even though
    ' OfferMarkingForUnwiredFields never actually returns False today (it
    ' is a notification, not a gate -- see its own header). BeginCollecting
    ' resets the buffer on the NEXT run regardless, so this was never a
    ' data-loss risk -- but skipping it would leave mCollecting stuck True
    ' until then, which is exactly the "fail closed" guarantee this
    ' module's own header states for every other exit from this window.
    If Not OfferMarkingForUnwiredFields(pres, TITLE) Then
        DraftingUI.EndCollecting
        Exit Sub
    End If

    Dim staged As String
    staged = DraftingUI.EndCollecting()

    If staged <> "" Then
        MsgBox CapReport(staged, "Your sheets are ready. Write your wording, then press '" & _
                                 CommandBarUI.CAP_PUT_ON_SLIDES & "'."), vbInformation, TITLE
    End If

    ' AND IT STOPS HERE. It used to fall straight into PutItOnTheSlidesCore.
    ' That is the coupling Rohan's artifact split exists to remove: setting up a
    ' quarter must not be able to change a slide, and pressing a button to
    ' publish must not be able to rebuild the sheets you typed into.
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
Public Sub PutItOnTheSlides()
    On Error GoTo Failed

    ' THE DECK SIDE STARTS BY PUBLISHING WHAT YOU TYPED -- every field, asking
    ' nothing. This is the other half of the artifact split: you press this when
    ' you have STOPPED TYPING, so the first thing it must do is take your words
    ' out of the drafting sheets and into the register.
    '
    ' Collected, so thirteen fields report once at the end instead of thirteen
    ' times on the way through.
    DraftingUI.BeginCollecting
    DraftingUI.PublishAllDraftedFields CommandBarUI.CAP_PUT_ON_SLIDES

    Dim published As String
    published = DraftingUI.EndCollecting()
    If published <> "" Then
        MsgBox CapReport(published, "Next: the slide changes."), vbInformation, CommandBarUI.CAP_PUT_ON_SLIDES
    End If

    PutItOnTheSlidesCore
    Exit Sub
Failed:
    Dim partialPub As String
    partialPub = DraftingUI.EndCollecting()
    RibbonUI.ShowSyncResult "Put it on the slides", RibbonUI.UnexpectedErrorText("Put it on the slides", Err.Number, Err.Description, Err.Source)
End Sub

Private Sub PutItOnTheSlidesCore()
    ' The guards for a missing workbook, missing types and an unopenable file
    ' live in ReviewChangesCore and say why in each case. Duplicating them here
    ' would mean two sets of wording to keep true, so when the detection cannot
    ' run this delegates and lets that path do the explaining.
    Dim pending As Long
    Dim sheetNames As String
    Dim stamp As String
    pending = ScanPendingApprovals(sheetNames, stamp)

    If pending = 0 Then
        ' BUILD, THEN ASK ONCE, THEN APPLY -- ALL IN THIS ONE PRESS.
        '
        ' Until 2026-08-18 this branch built the queue and STOPPED, telling the
        ' person to go look at the review sheet and press this same button a
        ' second time. Working as designed, not a bug -- but Rohan's own live
        ' first-ever run hit it and named the actual cost: "should only click
        ' put it on the slides once, then approve it... that's all." Pressing
        ' one button twice, with no visible reason the first press didn't just
        ' finish, reads as broken even when it isn't.
        '
        ' The three-way Yes/No/Cancel gate that stood here before THAT (deleted
        ' the same night, LOBBY-DESIGN.md phase 3) is not what this restores.
        ' That gate asked before anything was queued; this asks with the real,
        ' final, already-pre-ticked list in hand -- one question, "apply what
        ' you just built?", not three questions about what to build.
        Dim fullReport As String
        Dim totalQueued As Long
        If Not BuildAllQueuesCore(CommandBarUI.CAP_PUT_ON_SLIDES, fullReport, totalQueued) Then Exit Sub

        If totalQueued = 0 Then
            ShowSyncResult CommandBarUI.CAP_PUT_ON_SLIDES, fullReport & vbCrLf & "Nothing queued -- nothing to apply."
            Exit Sub
        End If

        If MsgBox(fullReport & vbCrLf & vbCrLf & totalQueued & _
                  " change(s) queued, pre-approved. Apply them now?", _
                  vbYesNo + vbQuestion, CommandBarUI.CAP_PUT_ON_SLIDES) <> vbYes Then
            ' Queue stays written to the review sheet either way -- saying No
            ' here does not lose it, same as the old two-press path never did.
            Exit Sub
        End If
    End If

    ' THE YES/NO/CANCEL GATE IS GONE. LOBBY-DESIGN.md phase 3, 2026-08-17.
    '
    ' It asked three questions this stage cannot answer any better than
    ' ApplyApproved already does, one ticked item at a time:
    '   - "Apply now?" -- yes, that is the entire point of pressing this button
    '     with pending ticks sitting in a review sheet.
    '   - "Or build a fresh review instead, losing the ticks?" -- a fresh
    '     review is exactly what a STALE tick needs, and ApplyApproved already
    '     revalidates every single item against the LIVE slide by hash before
    '     writing it (ReviewQueue.bas's own header: "THE REVALIDATION IS THE
    '     POINT, not a defensive extra"). A tick that has genuinely gone stale
    '     is DROPPED there, individually, with its own line in the report --
    '     "changed since you approved it; re-review" -- which is more precise
    '     than a blanket rebuild-or-not choice made before any item is looked
    '     at. This modal could only ever get that choice right by accident.
    '   - "Cancel, change nothing?" -- pressing THIS button, with ticks
    '     already sitting in a review sheet from an earlier pass, already
    '     communicated intent to apply them; a person who does not want that
    '     does not press it.
    ' Every field a queue ever shows now arrives pre-ticked (BuildQueue,
    ' ReviewQueue.bas) -- working the queue is REMOVING ticks from what should
    ' not sync, not adding them to bless what should (LOBBY-DESIGN.md section
    ' 5's full reasoning, including the deliberately-accepted residual risk
    ' for fields with no drafting-sheet gate). Requiring a second confirmation
    ' on top of that default is the same double-approval redundancy Rohan
    ' caught with ABOUT_BODY earlier the same night: approved once already,
    ' asked to approve the identical diff again with zero new information.
    ApplyApprovedCore
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
    RibbonUI.ShowSyncResult CommandBarUI.STAGE_APPLY_APPROVED, RibbonUI.UnexpectedErrorText(CommandBarUI.STAGE_APPLY_APPROVED, Err.Number, Err.Description, Err.Source)
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
        MsgBox "This deck has no paired workbook yet -- nothing to apply.", vbExclamation, CommandBarUI.STAGE_APPLY_APPROVED
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
        MsgBox "This deck has no registered slide types yet -- nothing to apply.", vbExclamation, CommandBarUI.STAGE_APPLY_APPROVED
        Exit Sub
    End If

    Dim wb As Object
    Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath)
    If wb Is Nothing Then
        MsgBox "Could not open the paired workbook at: " & workbookPath, vbCritical, CommandBarUI.STAGE_APPLY_APPROVED
        Exit Sub
    End If

    ' IS THIS EVEN OUR REGISTER? Same check DraftingUI.Resolve now makes for
    ' its own callers, added 2026-08-19. THIS is the actual slide-write step
    ' -- the highest-stakes of the three places this was missing, since a
    ' mismatch here means real content lands in another deck's register (or
    ' another register's values land on these slides) and every stage after
    ' reports success. See DeckRegistry's "THE PAIRING, BOTH WAYS" comment.
    Dim pairNote As String
    pairNote = DeckRegistry.PairingProblem(pres, wb)
    If pairNote <> "" Then
        MsgBox pairNote, vbCritical, CommandBarUI.STAGE_APPLY_APPROVED
        Exit Sub
    End If

    ' The ticks live in the workbook, so an unsaved workbook means the
    ' approvals being read are on screen and not in any file. Same refusal as
    ' the review step, for the same reason.
    ' Saves quietly instead of asking -- see WorkbookBridge.EnsureSavedQuietly's
    ' own header for why the prompt this replaced (2026-08-19) no longer needs
    ' to interrupt to do its job. A genuine save failure is still never hidden.
    Dim promptSaveProblem As String
    promptSaveProblem = WorkbookBridge.EnsureSavedQuietly(wb, workbookPath)
    If promptSaveProblem <> "" Then
        MsgBox promptSaveProblem & vbCrLf & vbCrLf & _
               "Stopping here rather than reading values that are not in the file.", _
               vbCritical, CommandBarUI.STAGE_APPLY_APPROVED
        Exit Sub
    End If

    Dim logWs As Object
    Set logWs = WorkbookBridge.GetOrAddWorksheet(wb, WorkbookBridge.SYNC_LOG_SHEET_NAME)

    Dim fullReport As String
    Dim totalWritten As Long, totalSkipped As Long, totalStale As Long, totalFailed As Long
    Dim refusedTypes As Long, skippedTypes As Long
    Dim i As Long
    For i = lo To hi
        Dim templateSld As Object
        Dim wsName As String
        If DeckRegistry.LookupType(pres, types(i), templateSld, wsName) Then
            Dim reviewName As String
            reviewName = ReviewQueue.ReviewSheetNameFor(types(i))

            If Not WorkbookBridge.WorksheetExists(wb, reviewName) Then
                fullReport = fullReport & "=== " & types(i) & " ===" & vbCrLf & _
                    "No review has been built for this type. Press '" & CommandBarUI.CAP_SET_UP_QUARTER & "' first." & vbCrLf & vbCrLf
                skippedTypes = skippedTypes + 1
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
                    refusedTypes = refusedTypes + 1
                Else
                    Dim reviewWs As Object
                    Set reviewWs = WorkbookBridge.GetOrAddWorksheet(wb, reviewName)

                    Dim tWritten As Long, tSkipped As Long, tStale As Long, tFailed As Long
                    fullReport = fullReport & _
                        ReviewQueue.ApplyApproved(sheet, types(i), reviewWs, logWs, _
                            tWritten, tSkipped, tStale, tFailed) & vbCrLf
                    totalWritten = totalWritten + tWritten
                    totalSkipped = totalSkipped + tSkipped
                    totalStale = totalStale + tStale
                    totalFailed = totalFailed + tFailed
                End If
            End If
        Else
            fullReport = fullReport & "SKIPPED " & types(i) & ": registered type's template slide no longer resolves (was it deleted?)" & vbCrLf
            skippedTypes = skippedTypes + 1
        End If
    Next i

    ' CAPTURED, NOT JUST APPENDED. Kingsbury-hound integrity audit, 2026-08-19:
    ' this return used to go ONLY into fullReport (the Run Log copy) -- the
    ' modal's own summary below never read it, so "deck and register both
    ' saved" printed even when PersistBothFiles had just returned
    ' "---- NOT SAVED ----". A person closing PowerPoint on that message loses
    ' every write this run just made, with nothing telling them to worry.
    Dim persistResult As String
    persistResult = PersistBothFiles(pres, wb)
    fullReport = fullReport & persistResult
    WorkbookBridge.WriteRunLog wb, "Apply Approved -- full report", fullReport

    ' MODAL GETS A SUMMARY, NEVER THE ITEM LIST -- Rohan, 2026-08-18: the
    ' modal listing every published item by ID and character count was the
    ' actual complaint, not the count of items. Full per-item detail is
    ' still unconditionally in the Run Log, written above, unchanged.
    ' Extracted to BuildApplyApprovedSummary so the persistResult branch can
    ' be proven with a fail-first test -- ApplyApprovedCore itself ends in a
    ' real ShowSyncResult MsgBox and can't run headless.
    Dim summary As String
    summary = BuildApplyApprovedSummary(persistResult, totalWritten, totalSkipped, totalStale, _
        totalFailed, refusedTypes, skippedTypes, DeckRegistry.GetDeckPeriod(pres), hi - lo + 1)

    ShowSyncResult CommandBarUI.STAGE_APPLY_APPROVED, summary
End Sub

Public Function BuildApplyApprovedSummary(persistResult As String, totalWritten As Long, _
        totalSkipped As Long, totalStale As Long, totalFailed As Long, refusedTypes As Long, _
        skippedTypes As Long, period As String, typeCount As Long) As String
    Dim summary As String
    summary = totalWritten & " written, " & totalSkipped & " not approved, " & _
        totalStale & " dropped as stale, " & totalFailed & " failed."
    If refusedTypes > 0 Then summary = summary & vbCrLf & refusedTypes & " type(s) refused -- see the Run Log."
    If skippedTypes > 0 Then summary = summary & vbCrLf & skippedTypes & " type(s) skipped -- see the Run Log."

    summary = summary & vbCrLf & vbCrLf & "Position: " & period & ", " & typeCount & " slide type(s)."

    ' SAVE FAILURE OUTRANKS EVERYTHING BELOW. Everything counted above
    ' (written/skipped/stale/failed) describes what was decided in memory --
    ' none of it is real until PersistBothFiles actually lands it on disk. A
    ' stale-item nudge or a "read the Run Log" hint is the wrong headline if
    ' this run's writes never made it past the object model.
    If InStr(1, persistResult, "---- NOT SAVED ----", vbTextCompare) > 0 Then
        summary = summary & vbCrLf & "Next: SAVE FAILED -- nothing from this run is on disk yet." & _
            vbCrLf & Trim(persistResult) & _
            vbCrLf & "Do not close without resolving this -- re-run Apply Approved once fixed."
    ElseIf totalStale > 0 Then
        summary = summary & vbCrLf & "Next: press '" & CommandBarUI.CAP_SET_UP_QUARTER & _
            "' to refresh the " & totalStale & " stale item(s) with their current before-and-after."
    ElseIf refusedTypes > 0 Or skippedTypes > 0 Then
        summary = summary & vbCrLf & "Next: read the Run Log for what needs fixing before the next sync."
    Else
        summary = summary & vbCrLf & "Next: nothing pending -- deck and register both saved."
    End If

    BuildApplyApprovedSummary = summary
End Function

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
        "Read it there, then press '" & CommandBarUI.CAP_PUT_ON_SLIDES & "' again."

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
        MsgBox "This deck has no registered slide types yet. Press '" & CommandBarUI.CAP_SET_UP_QUARTER & "' -- it walks setup on a deck that has none.", vbExclamation, "Audit Fields"
        Exit Sub
    End If

    Dim slideType As String
    slideType = PickType(types, "Audit Fields -- Choose Type")
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
    Dim carriedCount As Long, orphanedCount As Long
    Dim workbookPath As String
    workbookPath = DeckRegistry.GetWorkbookPath(pres)
    If workbookPath <> "" And rowCount > 0 Then
        Dim wb As Object
        Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath)
        If Not wb Is Nothing Then
            Dim ws As Object
            Set ws = WorkbookBridge.GetOrAddWorksheet(wb, TemplateAudit.AUDIT_SHEET_NAME)
            ' WriteAuditGrid now CARRIES every recorded field/chrome/drop
            ' decision forward by shape identity instead of refusing to
            ' rebuild over them (FIX-LIST P5) -- report what happened rather
            ' than checking for a refusal that no longer exists.
            TemplateAudit.WriteAuditGrid ws, rows, rowCount, carriedCount, orphanedCount
            wroteGrid = True
        End If
    End If

    Dim report As String
    report = TemplateAudit.SummaryText(slideType, subjectLabel, trackedCount, rowCount, likelyDataCount, comparisonCount)
    If wroteGrid And carriedCount > 0 Then
        report = report & vbCrLf & vbCrLf & "Carried " & carriedCount & " prior decision(s) forward."
    End If
    If wroteGrid And orphanedCount > 0 Then
        report = report & vbCrLf & vbCrLf & orphanedCount & " prior decision(s) could NOT be carried forward " & _
            "(that shape's text has changed, or the shape itself is gone) -- check the sheet's history if you need them."
    End If
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
    RibbonUI.ShowSyncResult CommandBarUI.CAP_CREATE_TEMPLATE, RibbonUI.UnexpectedErrorText(CommandBarUI.CAP_CREATE_TEMPLATE, Err.Number, Err.Description, Err.Source)
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
        MsgBox "This deck has no registered slide types yet. Press '" & CommandBarUI.CAP_SET_UP_QUARTER & "' -- it walks setup on a deck that has none.", vbExclamation, CommandBarUI.CAP_CREATE_TEMPLATE
        Exit Sub
    End If

    Dim slideType As String
    slideType = PickType(types, CommandBarUI.CAP_CREATE_TEMPLATE & " -- Choose Type")
    If slideType = "" Then Exit Sub

    ' THE SOURCE IS A REAL INSTANCE THE USER PICKS, not whatever the type's
    ' plain registration happens to point at. That registration (DeckSyncType:)
    ' gets overwritten to point at the FIRST template the moment one exists --
    ' RegisterType always replaces the single property -- so once a K template
    ' exists it can no longer supply "a representative real slide" for making
    ' an S template. The picked instance's OWN key is also where the letter
    ' comes from, below: no separate "which letter" prompt, because the letter
    ' is a fact about the slide the user is looking at, not a fact to ask for
    ' twice (Scenario 3, CHECKLIST.md).
    Dim wsName As String
    Dim sourceSld As Object
    Set sourceSld = PickTemplateSource(pres, slideType, wsName)
    If sourceSld Is Nothing Then Exit Sub

    ' Label the source by its instance key where it has one, falling back to
    ' the slide number -- the key is what the human recognises from the Data
    ' sheet, and "slide 3" is meaningless once the deck is reordered.
    Dim sourceInstance As SlideInstance
    sourceInstance = Resolve.ResolveSlideInstance(sourceSld)
    Dim sourceLabel As String
    sourceLabel = "slide " & sourceSld.SlideIndex
    If sourceInstance.HasInstanceKey Then sourceLabel = sourceInstance.InstanceKey & " (slide " & sourceSld.SlideIndex & ")"

    Dim letter As String
    letter = TemplateSlide.CodeLetterOf(sourceInstance.InstanceKey)

    ' Already has one FOR THIS LETTER: stop here rather than at MakeTemplateFrom's
    ' own guard, so the message can name the existing template's slide number.
    ' Both checks stay -- this one is for the human, that one is the invariant.
    ' TemplateSlide.ExistingTemplateForLetter is what actually generalised the
    ' old one-per-TYPE check to one-per-type-PER-LETTER (Scenario 3 step 4) --
    ' pulled out to its own testable function rather than inlined here,
    ' because this Sub's MsgBox/InputBox calls make it unreachable from the
    ' headless suite, and that logic is exactly the kind this project does not
    ' ship untested.
    Dim existing As Object
    Set existing = TemplateSlide.ExistingTemplateForLetter(pres, slideType, letter)
    If Not existing Is Nothing Then
        Dim letterNote As String
        letterNote = ""
        If letter <> "" Then letterNote = " for letter '" & letter & "'"
        MsgBox "Type '" & slideType & "'" & letterNote & " already has a master template: slide " & existing.SlideIndex & "." & vbCrLf & vbCrLf & _
               "A type/letter pair must have exactly one. Nothing was changed.", vbInformation, CommandBarUI.CAP_CREATE_TEMPLATE
        Exit Sub
    End If

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

    ' Same predicate DeckRegistry.RegisterNewTemplateLetter uses to decide
    ' whether it claims the type-level fallback -- computed here too so the
    ' confirmation the human reads matches the write it is confirming,
    ' instead of describing a fixed assumption that stopped being true the
    ' first time a type got a second letter.
    Dim willClaimFallback As Boolean
    If letter <> "" Then
        Dim fallbackPreviewSld As Object
        Dim fallbackPreviewWs As String
        If Not DeckRegistry.LookupType(pres, slideType, fallbackPreviewSld, fallbackPreviewWs) Then
            willClaimFallback = True
        ElseIf Not Resolve.IsTemplateSlide(fallbackPreviewSld) Then
            willClaimFallback = True
        End If
    End If

    If MsgBox(TemplateSlide.ConfirmTemplateText(slideType, sourceLabel, fieldCount, letter, willClaimFallback), _
              vbYesNo + vbQuestion, CommandBarUI.CAP_CREATE_TEMPLATE) <> vbYes Then
        Exit Sub
    End If

    Dim mr As MakeTemplateResult
    mr = TemplateSlide.MakeTemplateFrom(sourceSld, slideType)

    Dim report As String
    If Not mr.Ok Then
        report = "FAILED to create a template for '" & slideType & "': " & mr.Reason
        ShowSyncResult CommandBarUI.CAP_CREATE_TEMPLATE, report
        Exit Sub
    End If

    ' Registration is the step that actually changes behaviour -- without it
    ' the template exists but nothing clones it, which is the quietest
    ' possible half-finished state. Done here rather than inside
    ' MakeTemplateFrom so that function stays testable with no registry.
    ' DeckRegistry.RegisterNewTemplateLetter carries the same "pulled out so
    ' it's testable" reasoning as ExistingTemplateForLetter above.
    DeckRegistry.RegisterNewTemplateLetter pres, slideType, letter, mr.NewSlide, wsName

    report = "Master template created for '" & slideType & "'" & _
        IIf(letter <> "", " (letter '" & letter & "')", "") & "." & vbCrLf & vbCrLf & _
        "    slide " & mr.NewSlide.SlideIndex & ", hidden from the slideshow" & vbCrLf & _
        "    " & mr.FieldCount & " field(s) set to placeholders" & vbCrLf & _
        "    new records will now be cloned from it, not from " & sourceLabel & vbCrLf & vbCrLf & _
        "It will not appear in Preview Sync or Sync Now reports -- a template" & vbCrLf & _
        "is not a record, so it is neither counted nor corrected." & vbCrLf & vbCrLf & _
        "Worth doing now: open it and clear anything the sync does not manage" & vbCrLf & _
        "(figures, chart data, notes, untagged text) that belonged to " & sourceLabel & "."
    ShowSyncResult CommandBarUI.CAP_CREATE_TEMPLATE, report
End Sub

' Real, non-template instances of `slideType` -- these are what a template
' can actually be made FROM. Returns Nothing (having already told the human
' why) if there are none: the type is registered but nothing of it has been
' onboarded onto a real slide, or every real slide of it has already become
' a template.
'
' ByRef wsName carries back the type's worksheet name via DeckRegistry.LookupType
' regardless of what that lookup's SLIDE half currently resolves to (a real
' slide, a template, or nothing at all if deleted) -- the worksheet name does
' not change across a type's templates, only which slide is registered does,
' so this is the one place still allowed to read it that way.
Private Function PickTemplateSource(pres As Object, slideType As String, ByRef wsName As String) As Object
    Dim ignoredSld As Object
    wsName = ""
    DeckRegistry.LookupType pres, slideType, ignoredSld, wsName

    Dim instances() As Object
    instances = RunSync.GatherInstances(slideType)
    Dim lo As Long, hi As Long, hasAny As Boolean
    On Error Resume Next
    lo = LBound(instances): hi = UBound(instances): hasAny = (Err.Number = 0)
    On Error GoTo 0

    Dim candidates() As Object
    Dim keys() As String
    Dim n As Long
    n = 0
    If hasAny Then
        Dim i As Long
        For i = lo To hi
            Dim inst As SlideInstance
            inst = Resolve.ResolveSlideInstance(instances(i))
            If inst.HasInstanceKey Then
                n = n + 1
                ReDim Preserve candidates(1 To n)
                ReDim Preserve keys(1 To n)
                Set candidates(n) = instances(i)
                keys(n) = inst.InstanceKey
            End If
        Next i
    End If

    If n = 0 Then
        MsgBox "No real, already-onboarded slide of type '" & slideType & "' exists to build a template from." & vbCrLf & vbCrLf & _
               "Onboard at least one real project of this type first.", vbExclamation, CommandBarUI.CAP_CREATE_TEMPLATE
        Exit Function
    End If

    If n = 1 Then
        Set PickTemplateSource = candidates(1)
        Exit Function
    End If

    Dim prompt As String
    prompt = "Choose the real slide to build the template from (enter the number or the instance key):" & vbCrLf
    Dim j As Long
    For j = 1 To n
        Dim ltr As String
        ltr = TemplateSlide.CodeLetterOf(keys(j))
        prompt = prompt & j & ") " & keys(j) & IIf(ltr <> "", " (letter " & ltr & ")", "") & vbCrLf
    Next j

    Dim answer As String
    answer = InputBox(prompt, CommandBarUI.CAP_CREATE_TEMPLATE & " -- Choose Source")
    If Trim(answer) = "" Then Exit Function

    If IsNumeric(answer) Then
        Dim asNum As Long
        asNum = CLng(answer)
        If asNum >= 1 And asNum <= n Then
            Set PickTemplateSource = candidates(asNum)
            Exit Function
        End If
    End If
    For j = 1 To n
        If StrComp(Trim(answer), keys(j), vbTextCompare) = 0 Then
            Set PickTemplateSource = candidates(j)
            Exit Function
        End If
    Next j

    MsgBox "'" & answer & "' did not match a number or an instance key. Nothing was changed.", vbExclamation, CommandBarUI.CAP_CREATE_TEMPLATE
End Function

' A QUESTION WITH ONE POSSIBLE ANSWER IS NOT A QUESTION.
'
' Three call sites asked a person to TYPE a slide type -- by number or name --
' from a list that on every real deck so far has held exactly one entry. Rohan
' hit it twice inside a single setup run on 2026-08-13, and a typo returns ""
' which cancels the step outright. At one type the prompt had no upside at all:
' it could only cost something.
'
' It still asks whenever there is a genuine choice. FIX-LIST item 7, "things
' typed that could be picked", closed for this prompt.
'
' All three sites go through here rather than the fix landing only where it was
' noticed -- the same reason FieldForRun covers both of its call sites.
Public Function PickType(types() As String, caption As String) As String
    Dim lo As Long, hi As Long
    lo = LBound(types): hi = UBound(types)

    If lo = hi Then
        PickType = types(lo)
        Exit Function
    End If

    PickType = ResolveTypeAnswer(InputBox(BuildTypePickerPrompt(types), caption), types)
End Function

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
' noticeText is OVERRIDABLE (default unchanged for every existing caller) --
' the default notice promises the Run Log has the rest, which is true only
' when THIS call is the last thing to write there before the workbook
' saves. A caller inside a longer chain (RibbonUI.
' OfferMarkingForUnwiredFields is the first one, 2026-08-19) cannot make
' that promise, so it says something true instead of reusing a claim that
' happens to be false for it.
'
' maxLen is ALSO OVERRIDABLE (0 = use the shared REPORT_CAP, unchanged for
' every existing caller). Added the same night, live: a caller whose text
' becomes ONE SECTION of a larger combined report (DraftingUI.
' AppendCollected) still gets capped against the FULL 900-char REPORT_CAP
' by default, which does nothing to stop the section BEFORE it plus the
' sections after it from pushing the whole combined dialog over that same
' ceiling again, one layer up -- confirmed live: the coverage section
' alone was well under REPORT_CAP, and the OUTER CapReport call in
' SyncNowChainCore (wrapping Start a Quarter + Roll Forward + Refresh
' Drafting Sheets + this section together) still truncated mid-word, with
' ITS OWN default Run Log notice, because nothing had reserved headroom
' for the sections sharing the same dialog.
Public Function CapReport(text As String, Optional mustKeep As String = "", _
                          Optional noticeText As String = "", _
                          Optional maxLen As Long = 0) As String
    Dim cap As Long
    cap = IIf(maxLen > 0, maxLen, REPORT_CAP)

    CapReport = text
    If Len(text) <= cap Then Exit Function

    Dim notice As String
    If noticeText <> "" Then
        notice = vbCrLf & vbCrLf & noticeText
    Else
        notice = vbCrLf & vbCrLf & "[shortened -- the full list is on the '" & _
            WorkbookBridge.RUN_LOG_SHEET_NAME & "' sheet in the workbook]"
    End If

    Dim room As Long
    room = cap - Len(notice) - Len(mustKeep)
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
