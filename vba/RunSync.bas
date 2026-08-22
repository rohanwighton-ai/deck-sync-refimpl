Attribute VB_Name = "RunSync"
Option Explicit

' The orchestration driver: the piece every other module in this port
' explicitly draws a boundary around as "someone else's job" -- gathering
' live instances (SyncOperations' own non-goal), executing a decided
' duplication (SlideDuplication's caller's job), and reconciling deck order
' against Data-sheet row order every sync (specs/slide-duplication-
' trigger.md's standing invariant, whole-type in scope, not a single-call
' concern). This module is that someone else, composing discovery/
' identity_tags/matching/resolve/sync_operations/onboarding/excel_output/
' verification/slide_duplication into one real, runnable sync pass.
'
' Routine sync covers cases 1/3/4/6. Cases 5/7 remain non-goals.
'
' CASE 2 (period rollover) IS GONE, removed 2026-08-05. It duplicated a
' slide within the deck for the next period, which is the deck-accumulates
' model rejected on 2026-08-03: each period now gets its own deck file and
' last period's stays as the record. Rolling forward is a ROW operation
' (ExcelOutput.RollForwardPeriod), and no slide is created for it.

' ---------------------------------------------------------------------
' Gathering instances -- explicitly not SyncOperations' job (its own
' stated non-goal); this is where that gathering actually happens.
' ---------------------------------------------------------------------

' Every slide in the active presentation whose slide_type tag matches
' `slideType`, in current deck order, EXCLUDING the type's master template
' slide. Possibly unallocated if none exist yet (a genuinely empty type,
' e.g. before its first onboarding).
'
' The template exclusion lives here, at the one choke point, rather than at
' each of the five call sites below -- and excluding it here is what makes
' the whole of step 1 hold, because every consequence a template slide would
' otherwise have flows through this one function:
'   - PlanRoutineSync never sees it, so it is never reported as case 6
'     (unclassified_slide -- which is exactly what an untagged keyless typed
'     slide IS today, and would be noise on every single sync)
'   - ...nor corrected, nor counted in any summary the human reads
'   - ResequenceByRowOrder never sees it, so row order never moves it
'     (it happened to survive resequencing already, because anchorIndex only
'     tracks keyed slides -- but that was accidental, not asserted)
' Adding a sixth call site therefore inherits the exclusion for free, which
' is the point.
Public Function GatherInstances(slideType As String) As Object()
    Dim results() As Object
    Dim n As Long
    n = 0

    Dim sld As Object
    For Each sld In Application.ActivePresentation.Slides
        Dim resolved As SlideInstance
        resolved = Resolve.ResolveSlideInstance(sld)
        If resolved.HasTypeTag And resolved.TypeTag = slideType And Not resolved.IsTemplate Then
            n = n + 1
            ReDim Preserve results(1 To n)
            Set results(n) = sld
        End If
    Next sld

    GatherInstances = results
End Function

' ---------------------------------------------------------------------
' The main entry point
' ---------------------------------------------------------------------

' Runs one routine sync pass for `slideType`: reads `ws` (the paired Data
' sheet), gathers this type's current instances, dispatches
' SyncOperations.PlanRoutineSync's decisions (no_change/in_place_correction
' are already executed as a side effect of planning itself -- see
' SyncOperations.bas's own header comment; new_record is executed here via
' SlideDuplication.DuplicateAndTag; flagged is reported, never forced), and
' finally reconciles deck order against Data-sheet row order (the standing
' invariant, applied after every create/correct so it covers both new and
' pre-existing slides uniformly, not just what changed this pass).
'
' `templateSld` is the type's stored template (an already-onboarded
' reference slide) -- deciding where a template lives is explicitly out of
' scope everywhere else in this port (onboarding.md's own non-goal), so
' it's supplied here rather than looked up.
' DO NOT ADD A NEW CALLER OF THIS FUNCTION. ExcelOutput.ReadSheet reads with
' no period filter, and ReadSheetForPeriod's own comment says "first one
' wins" when an instance ID appears more than once -- on a register with
' more than one period's rows (every real register, past its first quarter),
' that means whichever period sits FIRST in the sheet, not the current one.
' Confirmed live 2026-08-21: a real sync through this function wrote
' Q3F26 prose onto the deck while reporting "43 corrected... SAVED to disk
' (verified)" -- both known callers (SyncRealDeck.bas, HiddenFixCheck.bas)
' were fixed the same night to build a period-aware Sheet themselves via
' ExcelOutput.ReadSheetForDeckPeriod and call RunRoutineSyncWithSheet
' directly, matching what RibbonUI.SyncNow (the real button) already did
' correctly all along. Kept only because deleting it risks breaking a caller
' this search didn't find; DerivedFieldTags-style "search before you add"
' applies here too -- grep for RunRoutineSync( before ever calling this one.
Public Function RunRoutineSync(ws As Object, slideType As String, templateSld As Object) As String
    Dim sheet As Sheet
    sheet = ExcelOutput.ReadSheet(ws)
    RunRoutineSync = RunRoutineSyncWithSheet(sheet, slideType, templateSld)
End Function

' The same sync, driven from an already-read sheet.
'
' Split out so the LONG-FORMAT REGISTER can feed the identical engine. That is
' the whole reason the format change is tractable: Register.ReadRegister
' returns the same `Sheet` UDT ExcelOutput.ReadSheet does, so the planner, the
' injector, the resequencer and the report below are untouched by the switch.
' Wide and long differ in how the sheet is READ, not in what the sync does with
' it -- and this function is where that claim is actually cashed.
'
' Both readers are live on purpose. Decks that have not migrated still use the
' wide sheet; deleting that path now would strand them.
Public Function RunRoutineSyncWithSheet(sheet As Sheet, slideType As String, templateSld As Object) As String
    Dim report As String
    report = "=== RunRoutineSync: " & slideType & " ===" & vbCrLf

    Dim instances() As Object
    instances = GatherInstances(slideType)

    ' srcWs DELIBERATELY NOT THREADED THROUGH HERE. This function (and
    ' RunRoutineSync, its only caller) has zero live callers of its own --
    ' slide creation left the sync path 2026-07-31, see this function's own
    ' comment above. A picture field would hit the same "Sources sheet was
    ' not available here" gap PreviewRoutineSync and PlanCounts were fixed
    ' for (2026-08-18) if this is ever reconnected -- fix it THEN, alongside
    ' whatever revives the call site, rather than threading a parameter
    ' through dead code on spec now.
    Dim actions() As SyncAction
    actions = SyncOperations.PlanRoutineSync(instances, sheet.InstanceOrder, sheet.Rows)

    Dim lo As Long, hi As Long, hasActions As Boolean
    On Error Resume Next
    lo = LBound(actions)
    hi = UBound(actions)
    hasActions = (Err.Number = 0)
    On Error GoTo 0

    Dim noChangeCount As Long, correctedCount As Long, newRecordCount As Long, failedCount As Long, flaggedCount As Long

    If hasActions Then
        Dim i As Long
        For i = lo To hi
            Select Case actions(i).Kind
                Case "no_change"
                    noChangeCount = noChangeCount + 1

                Case "in_place_correction"
                    ' Already written by PlanRoutineSync itself (it calls
                    ' InjectPrimitive per field directly) -- nothing left
                    ' to execute here, just report.
                    correctedCount = correctedCount + 1
                    report = report & "  corrected: " & actions(i).InstanceKey & vbCrLf
                    ' Field detail, same as the preview gives. Until 2026-07-30 this
                    ' branch printed the instance key and nothing else, so the report
                    ' you get AFTER a deck changes said less than the one you get
                    ' before -- and the moment you most want a record of what moved is
                    ' immediately after it moved. "was/now" rather than the preview's
                    ' "now/new": by this point the write has happened, and a report
                    ' that describes a completed change in the future tense is a lie.
                    Dim doneField As Variant
                    For Each doneField In actions(i).ChangedFieldCurrent.Keys
                        report = report & "      " & doneField & ":" & vbCrLf & _
                            "        was:  '" & BatchOnboardFlow.FieldPreview(CStr(actions(i).ChangedFieldCurrent(doneField))) & "'" & vbCrLf & _
                            "        now:  '" & BatchOnboardFlow.FieldPreview(CStr(actions(i).ChangedFieldNew(doneField))) & "'" & vbCrLf
                    Next doneField

                Case "new_record"
                    ' REPORTED, NOT CREATED, as of 2026-07-31.
                    '
                    ' A sync used to create a slide here. It no longer does, per
                    ' the Excel Control Layer requirement: "a register row with
                    ' no matching slide must never cause a slide to be created
                    ' as a consequence of a sync run. Creation is an operation a
                    ' person chooses."
                    '
                    ' This replaces a guard rather than adding one. Since
                    ' 2026-07-30 the confirmation stated slide creation in
                    ' capitals with its consequence spelled out, because the
                    ' real deck at the time had 43 orphaned rows against 46
                    ' slides and one click would have mass-duplicated. That
                    ' guard warned about a capability; removing the capability
                    ' is strictly better, and it is what makes an accidentally
                    ' synced composite deck merely wrong rather than
                    ' destructive -- a hand-assembled board pack is missing most
                    ' entities, so most rows are unmatched, so the old behaviour
                    ' would have set about creating them at exactly the moment
                    ' the content was finished.
                    '
                    ' D12 is preserved: a row with no slide is still a supported
                    ' case, serviced by CreateMissingSlides below.
                    newRecordCount = newRecordCount + 1
                    report = report & "  no slide for: " & actions(i).RowInstanceKey & _
                        "  (use Create Missing Slides -- sync does not create)" & vbCrLf

                Case "flagged"
                    flaggedCount = flaggedCount + 1
                    report = report & "  flagged: " & actions(i).Subject & " (" & actions(i).FlagKind & ") -- " & actions(i).Reason & vbCrLf
            End Select
        Next i
    End If

    report = report & "Summary: " & noChangeCount & " unchanged, " & correctedCount & " corrected, " & _
        newRecordCount & " with no slide, " & failedCount & " failed, " & flaggedCount & " flagged" & vbCrLf

    Dim moveCount As Long
    moveCount = ResequenceByRowOrder(slideType, sheet.InstanceOrder)
    report = report & "Resequenced " & moveCount & " slide(s) to match Data-sheet row order." & vbCrLf

    RunRoutineSyncWithSheet = report
End Function

' D12's supported case, as an operation a person chooses rather than a side
' effect of syncing: create a slide for every register row that has none.
'
' Split out of the sync on 2026-07-31. The engine always separated DECIDING to
' create (SyncOperations.PlanRoutineSync returns new_record) from DOING it
' (SlideDuplication.DuplicateAndTag), which is the only reason this is a
' rewiring rather than a rewrite -- the decision half stays exactly where it
' was and only the execution moves.
'
' dryRun defaults TRUE for the same reason TagMigration's does: this is the one
' remaining operation that can add slides in bulk, and the caller must ask for
' the write explicitly rather than get it by omission.
Public Function CreateMissingSlides(sheet As Sheet, slideType As String, templateSld As Object, _
                                    Optional dryRun As Boolean = True, Optional srcWs As Object = Nothing) As String
    Dim report As String
    report = IIf(dryRun, "=== PREVIEW: Create Missing Slides ===", "=== Create Missing Slides ===") & vbCrLf

    Dim instances() As Object
    instances = GatherInstances(slideType)

    ' dryRun:=True on the plan regardless of our own dryRun. PlanRoutineSync
    ' writes corrected field text WHILE planning (see its header), and this
    ' operation must not correct anything -- it creates. Letting it plan wet
    ' would make Create Missing Slides silently do a sync's work too.
    '
    ' srcWs NOT NEEDED HERE, unlike PreviewRoutineSync/PlanCounts (fixed
    ' 2026-08-18 for the same gap): this function only ever reads Kind =
    ' "new_record"/"flagged" from the result below, never "in_place_correction",
    ' so a picture field's WouldChange/ErrorMessage is computed and discarded
    ' either way. No `ws` in this function's own signature to resolve one from,
    ' and nothing downstream would consume it if there were.
    Dim actions() As SyncAction
    actions = SyncOperations.PlanRoutineSync(instances, sheet.InstanceOrder, sheet.Rows, True)

    Dim lo As Long, hi As Long, hasActions As Boolean
    On Error Resume Next
    lo = LBound(actions): hi = UBound(actions)
    hasActions = (Err.Number = 0)
    On Error GoTo 0

    ' REFUSE WHILE ANY SLIDE OF THIS TYPE IS UNCLASSIFIED.
    '
    ' The two classifications are not mutually exclusive, and this is the
    ' operation where that bites. A slide that keeps its slide_type tag but
    ' loses its instance_key is reported `flagged` -- AND the register row for
    ' the entity it should represent is reported `new_record`, because nothing
    ' claims that key any more. One broken slide, two categories.
    '
    ' Create Missing Slides is the documented remedy for new_record. Run in that
    ' state it duplicates the template into a SECOND slide for the entity while
    ' the original sits there untagged -- two slides silently claiming one
    ' project, and the remedy is what created the duplicate. Found by a MECE
    ' audit 2026-07-31; present in the Python reference implementation too, so
    ' it is original design rather than a port slip.
    '
    ' Refusing rather than warning: unlike a duplicate key, which is a fixable
    ' typo that should not block a quarter, this operation is the ONLY thing
    ' that adds slides in bulk, and the wrong answer is unrecoverable by the
    ' tool itself. The flagged slide almost certainly IS the entity about to be
    ' duplicated -- re-tagging it is the correct fix and costs one action.
    ' THE SOURCE MUST ACTUALLY BE A TEMPLATE, and nothing checked.
    '
    ' DeckRegistry.LookupType returns FindBySlideID of whatever slide the type was
    ' registered against. It does not assert is_template, and on the rig deck NO
    ' slide carries that tag at all -- preflight reports "0 template". So this
    ' would have cloned a REAL PROJECT'S SLIDE and tagged the copy with the new
    ' project's key.
    '
    ' Sync then corrects the fields it owns and leaves every other panel exactly
    ' as it was, so the new project's slide would carry the source project's
    ' Strategic Alignment, Problem and Project Progress -- the three panels
    ' nothing tracks. No check would ever fire: sync only inspects its own fields,
    ' and parity would report a matched pair.
    '
    ' Refusing rather than warning, for the reason stated below about flagged
    ' slides: this is the only operation that adds slides in bulk, and the wrong
    ' answer is not recoverable by the tool.
    If templateSld Is Nothing Then
        CreateMissingSlides = report & vbCrLf & _
            "REFUSED: this slide type has no template slide registered." & vbCrLf
        Exit Function
    End If
    If Not Resolve.IsTemplateSlide(templateSld) Then
        CreateMissingSlides = report & vbCrLf & _
            "REFUSED: the slide registered for '" & slideType & "' is NOT marked as a" & vbCrLf & _
            "template -- it is a real project's slide." & vbCrLf & vbCrLf & _
            "Copying it would give every new project that project's wording in any" & vbCrLf & _
            "panel this tool does not track, and nothing here would notice." & vbCrLf & vbCrLf & _
            "A template slide is created from inside setup, which '" & CommandBarUI.CAP_SET_UP_QUARTER & "' walks on a deck that has none." & vbCrLf
        Exit Function
    End If

    Dim flaggedCount As Long
    If hasActions Then
        Dim fi As Long
        For fi = lo To hi
            If actions(fi).Kind = "flagged" Then flaggedCount = flaggedCount + 1
        Next fi
    End If

    If flaggedCount > 0 Then
        report = report & vbCrLf & _
            "REFUSED: " & flaggedCount & " slide(s) of this type are UNCLASSIFIED." & vbCrLf & vbCrLf & _
            "A slide that has lost its instance key is reported both as flagged AND" & vbCrLf & _
            "as a missing slide for the entity it used to be. Creating slides now" & vbCrLf & _
            "would duplicate the template for entities that ALREADY HAVE a slide," & vbCrLf & _
            "leaving two slides claiming one project." & vbCrLf & vbCrLf & _
            "Re-tag the flagged slide(s) first, then run this again." & vbCrLf
        CreateMissingSlides = report
        Exit Function
    End If

    Dim createdCount As Long, failedCount As Long
    If hasActions Then
        Dim i As Long
        For i = lo To hi
            If actions(i).Kind = "new_record" Then
                ' PER-ROW TEMPLATE CHOICE, not the single type-level templateSld
                ' every row used to get. CodeLetterOf reads the letter straight
                ' off the row's own instance key ("" when it carries none, which
                ' is every key on a deck that predates per-letter registration),
                ' and LookupTemplateForLetter falls back to the type's plain
                ' registration on that same "" -- so this changes nothing for a
                ' deck with one template per type, and only picks a different
                ' template on a deck that has actually registered a second letter.
                Dim rowLetter As String
                rowLetter = TemplateSlide.CodeLetterOf(actions(i).RowInstanceKey)

                Dim rowTemplateSld As Object
                Dim rowWsName As String
                If Not DeckRegistry.LookupTemplateForLetter(Application.ActivePresentation, _
                        slideType, rowLetter, rowTemplateSld, rowWsName) Then
                    Set rowTemplateSld = templateSld
                End If

                ' Same check the type-level guard above already makes on
                ' templateSld, repeated here because a per-letter registration
                ' can point at a slide that was never actually marked
                ' is_template -- the exact defect class that guard exists to
                ' prevent, now reachable per-letter instead of only per-type.
                If Not Resolve.IsTemplateSlide(rowTemplateSld) Then
                    failedCount = failedCount + 1
                    report = report & "  FAILED " & actions(i).RowInstanceKey & _
                        ": the template registered for letter '" & rowLetter & _
                        "' is not marked as a template slide." & vbCrLf
                ElseIf dryRun Then
                    createdCount = createdCount + 1
                    report = report & "  would create: " & actions(i).RowInstanceKey & vbCrLf
                Else
                    Dim dr As DuplicateResult
                    dr = SlideDuplication.DuplicateAndTag(rowTemplateSld, slideType, _
                            actions(i).RowInstanceKey, actions(i).Values, instances, srcWs)
                    If dr.Ok Then
                        createdCount = createdCount + 1
                        report = report & "  created: " & actions(i).RowInstanceKey
                        If dr.MissingFieldCount > 0 Then
                            report = report & " (missing " & dr.MissingFieldCount & " field(s):"
                            Dim m As Long
                            For m = 1 To dr.MissingFieldCount
                                report = report & " " & dr.MissingFields(m)
                            Next m
                            report = report & ")"
                        End If
                        If dr.FailedFieldCount > 0 Then
                            report = report & " (FAILED TO WRITE " & dr.FailedFieldCount & " field(s):"
                            Dim fm As Long
                            For fm = 1 To dr.FailedFieldCount
                                report = report & " " & dr.FailedFields(fm)
                            Next fm
                            report = report & ")"
                        End If
                        report = report & vbCrLf
                    Else
                        failedCount = failedCount + 1
                        report = report & "  FAILED " & actions(i).RowInstanceKey & ": " & dr.Reason & vbCrLf
                    End If
                End If
            End If
        Next i
    End If

    report = report & "Summary: " & createdCount & IIf(dryRun, " would be created", " created") & _
        ", " & failedCount & " failed" & vbCrLf

    If Not dryRun And createdCount > 0 Then
        Dim moveCount As Long
        moveCount = ResequenceByRowOrder(slideType, sheet.InstanceOrder)
        report = report & "Resequenced " & moveCount & " slide(s) to match row order." & vbCrLf
    End If

    CreateMissingSlides = report
End Function

' Read-only twin of RunRoutineSync: says exactly what a routine sync would do
' to this deck, and does none of it.
'
' A routine sync has THREE mutation sites, and a preview is only trustworthy if
' every one of them is suppressed:
'   1. PlanRoutineSync -> InjectPrimitive, which writes corrected field text
'      *while planning* (surprising, but real -- see that function's header)
'   2. this function's own new_record branch -> SlideDuplication.DuplicateAndTag
'   3. ResequenceByRowOrder -> Slide.MoveTo
' 1 and 3 are suppressed by their own dryRun flags; 2 simply isn't reached
' here, because this function reports the action instead of executing it.
'
' Worth previewing before every real sync until Sync Now has a track record:
' an orphaned Data row (one whose instance key matches no slide) is classified
' new_record, so a deck whose linkage has drifted can quietly turn a sync into
' a mass slide duplication. That is not hypothetical -- it was the live state
' of the real deck on 2026-07-27 (43 orphaned rows against 46 slides), and
' only the button being absent from the toolbar prevented it.
' deckPeriod is REQUIRED, not optional with a "" default.
'
' It read the sheet through ExcelOutput.ReadSheet -- the UNFILTERED reader --
' while every sync-side read used ReadSheetForDeckPeriod. ExcelOutput's own header
' says in capitals "THE SYNC PATH MUST USE THIS, NOT ReadSheet"; that fix landed on
' the four RibbonUI call sites and missed this one, because it lives inside another
' module. So the preview read every period's rows at once, and where a project has
' a row in two periods the unfiltered reader keeps whichever sits HIGHER and files
' the rest in DuplicateInstances, which nothing here looked at.
'
' On the rig this currently gets the right answer by accident: all five projects
' with two rows happen to have Q4F26 directly above Q1F27. Reorder them, or roll
' forward so the new period lands first, and the preview silently starts previewing
' a different period from the one Sync Now would write. A defect that agrees with
' reality today is the kind testing now cannot catch -- hence a required argument
' rather than a default that preserves the old behaviour for any caller that forgets.
Public Function PreviewRoutineSync(ws As Object, slideType As String, deckPeriod As String) As String
    Dim report As String
    report = "=== PREVIEW (nothing written): " & slideType & " ===" & vbCrLf

    ' R9: duplicate identity tags, at the TOP of the report rather than the
    ' bottom. The preview's job is to tell a human what a sync would do, and if
    ' two slides share a key the honest answer is "one of these, and it is not
    ' defined which" -- which invalidates every count below it. Burying that
    ' under the summary would let someone read the counts and act on them.
    Dim dupReport As DuplicateKeyReport
    dupReport = IdentityCheck.FindDuplicateKeys(slideType)
    If dupReport.HasDuplicates Then
        report = report & IdentityCheck.DuplicateKeyWarningText(slideType, dupReport) & vbCrLf & vbCrLf
    End If

    Dim sheet As Sheet
    Dim readProblem As String
    sheet = ExcelOutput.ReadSheetForDeckPeriod(ws, deckPeriod, readProblem)
    If readProblem <> "" Then
        ' Reported and returned rather than previewed against a sheet that cannot
        ' be trusted. A preview writes nothing, so this is not dangerous in itself
        ' -- but a preview of the wrong rows is worse than no preview, because it
        ' is the thing someone then approves a real sync from.
        PreviewRoutineSync = report & "REFUSED at period '" & deckPeriod & "': " & _
            readProblem & vbCrLf
        Exit Function
    End If

    Dim instances() As Object
    instances = GatherInstances(slideType)

    ' SAME GAP ApplyApproved AND BuildQueue ALREADY CLOSE, found 2026-08-18 by
    ' a purpose-hound sweep after the BuildQueue fix landed: PlanRoutineSync
    ' has six call sites and only one had srcWs threaded through. Without it,
    ' a picture field can never report itself correctly unchanged here either
    ' -- every Preview would show it as "would correct" with "the Sources
    ' sheet was not available here", live and reachable from DraftingUI's
    ' post-publish "Preview only" prompt.
    Dim srcWs As Object
    Set srcWs = Nothing
    If WorkbookBridge.WorksheetExists(ws.Parent, Sources.SOURCES_SHEET_NAME) Then
        Set srcWs = ws.Parent.Worksheets(Sources.SOURCES_SHEET_NAME)
    End If

    Dim actions() As SyncAction
    actions = SyncOperations.PlanRoutineSync(instances, sheet.InstanceOrder, sheet.Rows, True, Nothing, "", srcWs)

    Dim lo As Long, hi As Long, hasActions As Boolean
    On Error Resume Next
    lo = LBound(actions)
    hi = UBound(actions)
    hasActions = (Err.Number = 0)
    On Error GoTo 0

    Dim noChangeCount As Long, wouldCorrectCount As Long, wouldCreateCount As Long, flaggedCount As Long

    If hasActions Then
        Dim i As Long
        For i = lo To hi
            Select Case actions(i).Kind
                Case "no_change"
                    noChangeCount = noChangeCount + 1

                Case "in_place_correction"
                    wouldCorrectCount = wouldCorrectCount + 1
                    report = report & "  would correct: " & actions(i).InstanceKey & vbCrLf
                    Dim fieldName As Variant
                    For Each fieldName In actions(i).ChangedFieldCurrent.Keys
                        ' Both halves, always. A preview that shows only "now"
                        ' makes the human open Excel to find out what they are
                        ' approving, which is the errand the preview exists to
                        ' save them -- found live 2026-07-30, first real use.
                        report = report & "      " & fieldName & ":" & vbCrLf & _
                            "        now:  '" & BatchOnboardFlow.FieldPreview(CStr(actions(i).ChangedFieldCurrent(fieldName))) & "'" & vbCrLf & _
                            "        new:  '" & BatchOnboardFlow.FieldPreview(CStr(actions(i).ChangedFieldNew(fieldName))) & "'" & vbCrLf
                    Next fieldName

                Case "new_record"
                    wouldCreateCount = wouldCreateCount + 1
                    report = report & "  REACHES NO SLIDE: " & actions(i).RowInstanceKey & _
                        " -- no slide carries this row's instance key" & vbCrLf

                Case "flagged"
                    flaggedCount = flaggedCount + 1
                    report = report & "  flagged: " & actions(i).Subject & " (" & actions(i).FlagKind & ") -- " & actions(i).Reason & vbCrLf
            End Select
        Next i
    End If

    report = report & "Summary: " & noChangeCount & " unchanged, " & wouldCorrectCount & " would be corrected, " & _
        wouldCreateCount & " row(s) reach no slide, " & flaggedCount & " flagged" & vbCrLf

    Dim outOfPosition As Long
    outOfPosition = ResequenceByRowOrder(slideType, sheet.InstanceOrder, True)
    report = report & outOfPosition & " slide(s) are not in Data-sheet row order." & vbCrLf

    ' THIS USED TO THREATEN MASS DUPLICATION, AND THAT IS NO LONGER TRUE.
    '
    ' It said "A real Sync Now would DUPLICATE the template slide once for each".
    ' Slide creation left the sync path on 2026-07-31 (DECISIONS.md) -- nothing in
    ' production calls RunRoutineSync any more, and ReviewQueue.BuildQueue never
    ' queues a new_record. So the two buttons described each other wrongly in BOTH
    ' directions: this one warned of a danger that cannot happen, while Sync Now
    ' said "every linked slide already matches the workbook" over the very same
    ' rows. Still called out on its own line, because a row that reaches nothing is
    ' content silently missing from the deck -- just not for the old reason.
    If wouldCreateCount > 0 Then
        report = report & vbCrLf & "WARNING: " & wouldCreateCount & " Data row(s) match no slide in this deck." & vbCrLf & _
            "Their text will NOT appear anywhere -- Sync Now does not create slides." & vbCrLf & _
            "Either those slides are missing, or their instance keys disagree with the register." & vbCrLf & _
            "There is no button for adding a missing slide yet: add it from the template" & vbCrLf & _
            "by hand and tag it, or fix the key on the slide that should carry the row." & vbCrLf
    End If

    PreviewRoutineSync = report
End Function

' The same plan the preview reports on, reduced to counts, so Sync Now can
' confirm before it writes instead of after.
'
' Counts rather than the preview's report string on purpose: a caller that
' needs to know "will this create slides?" must not have to parse prose to
' find out. Sync Now's guard has to be exactly as reliable as the planner,
' and a regex over a human-readable summary is not.
'
' dryRun:=True is load-bearing, not defensive. PlanRoutineSync's own
' InjectPrimitive call writes corrected text *while planning* (see that
' function's header), so counting the work without this flag would perform
' half of it -- the confirmation would be describing changes it had already
' made. Deliberately does NOT call ResequenceByRowOrder: that is the third
' mutation site, and nothing here needs its number.
Public Sub PlanCounts(ws As Object, slideType As String, ByRef unchangedCount As Long, _
                      ByRef correctCount As Long, ByRef createCount As Long, ByRef flagCount As Long)
    unchangedCount = 0
    correctCount = 0
    createCount = 0
    flagCount = 0

    Dim sheet As Sheet
    sheet = ExcelOutput.ReadSheet(ws)

    Dim instances() As Object
    instances = GatherInstances(slideType)

    ' SAME GAP AS PreviewRoutineSync, closed the same way -- ws is already
    ' available here, so there is no reason for this call site to be the odd
    ' one out among PlanRoutineSync's callers.
    Dim srcWs As Object
    Set srcWs = Nothing
    If WorkbookBridge.WorksheetExists(ws.Parent, Sources.SOURCES_SHEET_NAME) Then
        Set srcWs = ws.Parent.Worksheets(Sources.SOURCES_SHEET_NAME)
    End If

    Dim actions() As SyncAction
    actions = SyncOperations.PlanRoutineSync(instances, sheet.InstanceOrder, sheet.Rows, True, Nothing, "", srcWs)

    Dim lo As Long, hi As Long
    On Error Resume Next
    lo = LBound(actions)
    hi = UBound(actions)
    If Err.Number <> 0 Then Exit Sub   ' no actions at all -- all four stay 0
    On Error GoTo 0

    Dim i As Long
    For i = lo To hi
        Select Case actions(i).Kind
            Case "no_change":           unchangedCount = unchangedCount + 1
            Case "in_place_correction": correctCount = correctCount + 1
            Case "new_record":          createCount = createCount + 1
            Case "flagged":             flagCount = flagCount + 1
        End Select
    Next i
End Sub

' The text of Sync Now's confirmation. Pure, so the wording is testable
' without a live deck or a dialog.
'
' Slide creation is stated in capitals and on its own line whenever it is
' non-zero. The 2026-07-27 near-miss (43 orphaned rows against 46 slides)
' would have read as an unremarkable "43 created" buried in a summary; the
' one number that is painful to undo should not look like the others.
Public Function ConfirmSyncText(correctCount As Long, createCount As Long, flagCount As Long) As String
    Dim s As String
    s = "This will change the deck." & vbCrLf & vbCrLf & _
        "    " & correctCount & " slide(s) corrected" & vbCrLf

    ' No longer a warning, because a sync can no longer create anything --
    ' creation moved to CreateMissingSlides on 2026-07-31. What was the
    ' loudest line in this dialog is now an informational count.
    '
    ' Deliberately still SHOWN rather than dropped: a large number here is the
    ' signal that this deck's linkage has drifted, or that it is a
    ' hand-assembled composite that should not be synced at all. The number was
    ' always the useful part; the alarm was only needed while the alarm was the
    ' only thing standing between it and mass duplication.
    If createCount > 0 Then
        s = s & "    " & createCount & " row(s) have no slide -- NOT created by this action" & vbCrLf & _
            "      (a large number here usually means drifted linkage," & vbCrLf & _
            "       or that this deck is an assembled pack, not a source deck)" & vbCrLf
    Else
        s = s & "    0 rows without a slide" & vbCrLf
    End If

    If flagCount > 0 Then s = s & "    " & flagCount & " flagged" & vbCrLf

    s = s & vbCrLf & "Proceed?"
    ConfirmSyncText = s
End Function

' ---------------------------------------------------------------------
' Row-order resequencing (specs/slide-duplication-trigger.md's standing
' invariant)
' ---------------------------------------------------------------------

' Repositions every slide of `slideType` so their relative deck order
' matches `instanceOrder` (the Data-sheet's current row order) exactly --
' re-gathers instances itself (rather than reusing the caller's, since
' new_record may have just created more), and moves any that are out of
' place via Slide.MoveTo. Returns the number of slides actually moved.
'
' Design choice not fully pinned down by the spec: the resequenced block
' is anchored at the CURRENT lowest SlideIndex among this type's slides
' (not forced to the front of the deck) -- preserves wherever a human
' originally chose to place this type's slides within the overall deck,
' fixing only their relative order among themselves. A different, equally
' defensible choice (always push to the front, or to the end) was not
' made; this one seemed least surprising for a deck that mixes multiple
' slide types.
' `dryRun` counts without moving anything. The count is then "slides not
' currently sitting at their target index", which is an honest answer to "is
' the deck out of order, and roughly how badly" but is NOT guaranteed to equal
' the number of MoveTo calls a real run makes: each real move shifts the
' indices of other slides, so the two can differ. Reported as "out of position"
' rather than "would move" for exactly that reason -- a preview that quietly
' overstates its own precision is worse than one that admits the limit.
Public Function ResequenceByRowOrder(slideType As String, instanceOrder As Collection, Optional dryRun As Boolean = False) As Long
    Dim freshInstances() As Object
    freshInstances = GatherInstances(slideType)

    Dim lo As Long, hi As Long, hasAny As Boolean
    On Error Resume Next
    lo = LBound(freshInstances)
    hi = UBound(freshInstances)
    hasAny = (Err.Number = 0)
    On Error GoTo 0

    If Not hasAny Then
        ResequenceByRowOrder = 0
        Exit Function
    End If

    Dim byKey As Object
    Set byKey = CreateObject("Scripting.Dictionary")
    Dim anchorIndex As Long
    anchorIndex = -1

    Dim i As Long
    For i = lo To hi
        Dim resolved As SlideInstance
        resolved = Resolve.ResolveSlideInstance(freshInstances(i))
        If resolved.HasInstanceKey Then
            Set byKey(resolved.InstanceKey) = freshInstances(i)
            If anchorIndex = -1 Then
                anchorIndex = freshInstances(i).SlideIndex
            ElseIf freshInstances(i).SlideIndex < anchorIndex Then
                anchorIndex = freshInstances(i).SlideIndex
            End If
        End If
    Next i

    If anchorIndex = -1 Then
        ResequenceByRowOrder = 0
        Exit Function
    End If

    Dim moveCount As Long
    moveCount = 0
    Dim prevIndex As Long
    prevIndex = anchorIndex - 1

    Dim key As Variant
    For Each key In instanceOrder
        If byKey.Exists(CStr(key)) Then
            Dim sld As Object
            Set sld = byKey(CStr(key))
            Dim targetIndex As Long
            targetIndex = prevIndex + 1
            If sld.SlideIndex <> targetIndex Then
                If Not dryRun Then sld.MoveTo targetIndex
                moveCount = moveCount + 1
            End If
            prevIndex = targetIndex
        End If
    Next key

    ResequenceByRowOrder = moveCount
End Function
