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
' Routine sync (cases 1/3/4/6) and period rollover (case 2, via
' RunPeriodRollover below) are two distinct entry points, never conflated
' -- case 2 stays explicitly-invoked-only per specs/sync-operations.md's
' own "never inferred from routine sync" rule; RunRoutineSync itself never
' calls into it. Cases 5/7 remain non-goals throughout this project.

' ---------------------------------------------------------------------
' Gathering instances -- explicitly not SyncOperations' job (its own
' stated non-goal); this is where that gathering actually happens.
' ---------------------------------------------------------------------

' Every slide in the active presentation whose slide_type tag matches
' `slideType`, in current deck order. Possibly unallocated if none exist
' yet (a genuinely empty type, e.g. before its first onboarding).
Public Function GatherInstances(slideType As String) As Object()
    Dim results() As Object
    Dim n As Long
    n = 0

    Dim sld As Object
    For Each sld In Application.ActivePresentation.Slides
        Dim resolved As SlideInstance
        resolved = Resolve.ResolveSlideInstance(sld)
        If resolved.HasTypeTag And resolved.TypeTag = slideType Then
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
Public Function RunRoutineSync(ws As Object, slideType As String, templateSld As Object) As String
    Dim report As String
    report = "=== RunRoutineSync: " & slideType & " ===" & vbCrLf

    Dim sheet As Sheet
    sheet = ExcelOutput.ReadSheet(ws)

    Dim instances() As Object
    instances = GatherInstances(slideType)

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

                Case "new_record"
                    Dim dr As DuplicateResult
                    dr = SlideDuplication.DuplicateAndTag(templateSld, slideType, actions(i).RowInstanceKey, actions(i).Values, instances)
                    If dr.Ok Then
                        newRecordCount = newRecordCount + 1
                        report = report & "  created: " & actions(i).RowInstanceKey
                        If dr.MissingFieldCount > 0 Then
                            report = report & " (missing " & dr.MissingFieldCount & " field(s):"
                            Dim m As Long
                            For m = 1 To dr.MissingFieldCount
                                report = report & " " & dr.MissingFields(m)
                            Next m
                            report = report & ")"
                        End If
                        report = report & vbCrLf
                    Else
                        failedCount = failedCount + 1
                        report = report & "  FAILED to create " & actions(i).RowInstanceKey & ": " & dr.Reason & vbCrLf
                    End If

                Case "flagged"
                    flaggedCount = flaggedCount + 1
                    report = report & "  flagged: " & actions(i).Subject & " (" & actions(i).FlagKind & ") -- " & actions(i).Reason & vbCrLf
            End Select
        Next i
    End If

    report = report & "Summary: " & noChangeCount & " unchanged, " & correctedCount & " corrected, " & _
        newRecordCount & " created, " & failedCount & " failed, " & flaggedCount & " flagged" & vbCrLf

    Dim moveCount As Long
    moveCount = ResequenceByRowOrder(slideType, sheet.InstanceOrder)
    report = report & "Resequenced " & moveCount & " slide(s) to match Data-sheet row order." & vbCrLf

    RunRoutineSync = report
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
Public Function PreviewRoutineSync(ws As Object, slideType As String) As String
    Dim report As String
    report = "=== PREVIEW (nothing written): " & slideType & " ===" & vbCrLf

    Dim sheet As Sheet
    sheet = ExcelOutput.ReadSheet(ws)

    Dim instances() As Object
    instances = GatherInstances(slideType)

    Dim actions() As SyncAction
    actions = SyncOperations.PlanRoutineSync(instances, sheet.InstanceOrder, sheet.Rows, True)

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
                    report = report & "  WOULD CREATE A NEW SLIDE: " & actions(i).RowInstanceKey & _
                        " -- no slide carries this row's instance key" & vbCrLf

                Case "flagged"
                    flaggedCount = flaggedCount + 1
                    report = report & "  flagged: " & actions(i).Subject & " (" & actions(i).FlagKind & ") -- " & actions(i).Reason & vbCrLf
            End Select
        Next i
    End If

    report = report & "Summary: " & noChangeCount & " unchanged, " & wouldCorrectCount & " would be corrected, " & _
        wouldCreateCount & " new slide(s) would be created, " & flaggedCount & " flagged" & vbCrLf

    Dim outOfPosition As Long
    outOfPosition = ResequenceByRowOrder(slideType, sheet.InstanceOrder, True)
    report = report & outOfPosition & " slide(s) are not in Data-sheet row order." & vbCrLf

    ' The loud one. Mass duplication is the only outcome here that is painful
    ' to undo, so it gets called out on its own rather than left as a number
    ' in a summary line someone skims.
    If wouldCreateCount > 0 Then
        report = report & vbCrLf & "WARNING: " & wouldCreateCount & " Data row(s) match no slide in this deck." & vbCrLf & _
            "A real Sync Now would DUPLICATE the template slide once for each." & vbCrLf & _
            "If that is not what you want, fix the linkage before syncing." & vbCrLf
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

    Dim actions() As SyncAction
    actions = SyncOperations.PlanRoutineSync(instances, sheet.InstanceOrder, sheet.Rows, True)

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

    If createCount > 0 Then
        s = s & "    " & createCount & " NEW SLIDE(S) WILL BE CREATED" & vbCrLf & vbCrLf & _
            "Slides get created when a Data row matches no slide in" & vbCrLf & _
            "the deck. If the linkage has drifted, that is a mass" & vbCrLf & _
            "duplication, not an update. Run Preview Sync first if" & vbCrLf & _
            "you are not expecting new slides." & vbCrLf
    Else
        s = s & "    0 new slides created" & vbCrLf
    End If

    If flagCount > 0 Then s = s & "    " & flagCount & " flagged" & vbCrLf

    s = s & vbCrLf & "Proceed?"
    ConfirmSyncText = s
End Function

' ---------------------------------------------------------------------
' Period rollover (case 2) -- a distinct, explicitly-invoked entry point,
' never reachable from RunRoutineSync above (SyncOperations.
' PlanPeriodRollover's own "never inferred from routine sync" rule).
' ---------------------------------------------------------------------

' Executes an explicit period rollover for one named instance: resolves
' `sourceSld` (the instance's own current slide) into a SlideInstance,
' calls SyncOperations.PlanPeriodRollover to decide the rollover, then
' duplicates `sourceSld` itself into a new slide tagged `newInstanceKey`
' with `newValues` injected -- mirroring RunRoutineSync's own case-3
' handling (SlideDuplication.DuplicateAndTag), but driven from an explicit
' command against one known instance rather than an unmatched Data-sheet
' row. `sourceSld` is left untouched as history, per specs/sync-
' operations.md's case-2 requirement -- DuplicateAndTag never mutates its
' source, only the new duplicate.
'
' `sourceSld` is accepted un-resolved (rather than requiring the caller to
' pass an already-built SlideInstance) so a future picker UI can hand over
' exactly the slide a user selected without resolving it first.
Public Function RunPeriodRollover(sourceSld As Object, slideType As String, newInstanceKey As String, newValues As Object) As DuplicateResult
    Dim instance As SlideInstance
    instance = Resolve.ResolveSlideInstance(sourceSld)

    Dim rollover As PeriodRollover
    rollover = SyncOperations.PlanPeriodRollover(instance, newValues)

    Dim existingInstances() As Object
    existingInstances = GatherInstances(slideType)

    RunPeriodRollover = SlideDuplication.DuplicateAndTag(sourceSld, slideType, newInstanceKey, rollover.NewValues, existingInstances)
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
