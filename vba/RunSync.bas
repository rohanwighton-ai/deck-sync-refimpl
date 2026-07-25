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
Public Function ResequenceByRowOrder(slideType As String, instanceOrder As Collection) As Long
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
                sld.MoveTo targetIndex
                moveCount = moveCount + 1
            End If
            prevIndex = targetIndex
        End If
    Next key

    ResequenceByRowOrder = moveCount
End Function
