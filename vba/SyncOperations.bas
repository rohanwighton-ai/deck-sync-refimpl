Attribute VB_Name = "SyncOperations"
Option Explicit

' VBA port of src/sync_operations.py's plan_routine_sync()/
' plan_period_rollover(), per specs/vba-port.md's port order (module 4,
' together with Resolve.bas). Dispatches per specs/sync-operations.md:
' cases 1 (no_change), 3 (new_record), and 4 (in_place_correction) from
' PlanRoutineSync; case 6 (unclassified_slide) is folded into the same
' function, matching the Python original. Case 2 (period_rollover) is
' PlanPeriodRollover, a distinct, separately-invoked function never
' reachable from PlanRoutineSync -- never inferred from a value merely
' looking different. Cases 5/7 are non-goals per specs/sync-operations.md
' and are not produced anywhere in this file.
'
' See SPIKE_NOTES_Resolve.md for the full divergence list -- most notably,
' this port skips resolve.py's separate field_shapes pre-resolution step
' and calls InjectPrimitive.InjectPrimitive() directly per field, relying
' on its own native tag lookup as the single source of truth for "does
' this instance have a shape for this field."

Public Type SyncAction
    Kind As String              ' "no_change" | "in_place_correction" | "new_record" | "flagged"
    InstanceKey As String       ' no_change / in_place_correction
    ChangedFields As Object     ' in_place_correction: Scripting.Dictionary fieldName -> InjectResult
    RowInstanceKey As String    ' new_record: the Data-sheet row's instance key
    Values As Object            ' new_record: Scripting.Dictionary fieldName -> String value
    Subject As String           ' flagged: a readable handle on the flagged slide (its SlideID)
    FlagKind As String          ' flagged: always "unclassified_slide" here (cases 5/7 are non-goals)
    Reason As String
End Type

Public Type PeriodRollover
    SourceInstanceKey As String
    NewValues As Object         ' Scripting.Dictionary fieldName -> String value
    Reason As String
End Type

' Dispatch cases 1/3/4/6 across `instances` (live Slide objects already
' believed to belong to one type -- gathering that set is the caller's job,
' matching sync_operations.py's own non-goal of not discovering instances
' itself) against a Data-sheet already read into memory:
'   dataRows     - Scripting.Dictionary: instance id (String) -> Scripting.
'                  Dictionary of fieldName (String) -> value (String)
'   instanceOrder - Collection of instance-id strings, in the Data-sheet's
'                  row order
' `instances` must be a 1-based array, using the same "(1 To 0)" convention
' Discovery.bas's own output uses for "no items".
'
' Module 6 (Excel-side reads) is not yet ported (per specs/vba-port.md's
' port order, it's step 6, after onboarding) -- this function accepts the
' shape that reader will eventually produce rather than reading a worksheet
' itself, mirroring sync_operations.py's own separation: it never touches a
' file directly either, excel_output.py reads the Sheet it operates on.
Public Function PlanRoutineSync(instances() As Object, instanceOrder As Collection, dataRows As Object) As SyncAction()
    Dim actions() As SyncAction
    ReDim actions(1 To 0)
    Dim n As Long
    n = 0

    Dim lo As Long, hi As Long
    lo = LBound(instances)
    hi = UBound(instances)

    Dim resolved() As SlideInstance
    ReDim resolved(lo To hi)

    Dim i As Long
    For i = lo To hi
        resolved(i) = Resolve.ResolveSlideInstance(instances(i))
        If Not resolved(i).HasTypeTag Or Not resolved(i).HasInstanceKey Then
            n = n + 1
            ReDim Preserve actions(1 To n)
            actions(n).Kind = "flagged"
            actions(n).Subject = "SlideID " & instances(i).SlideID
            actions(n).FlagKind = "unclassified_slide"
            actions(n).Reason = "no recognized type tag / persistent instance key -- flagged for reclassification, not guessed"
        End If
    Next i

    Dim knownByKey As Object
    Set knownByKey = CreateObject("Scripting.Dictionary")
    For i = lo To hi
        If resolved(i).HasInstanceKey Then
            knownByKey(resolved(i).InstanceKey) = i
        End If
    Next i

    Dim instanceId As Variant
    For Each instanceId In instanceOrder
        Dim key As String
        key = CStr(instanceId)

        Dim rowValues As Object
        If dataRows.Exists(key) Then
            Set rowValues = dataRows(key)
        Else
            Set rowValues = CreateObject("Scripting.Dictionary")
        End If

        If Not knownByKey.Exists(key) Then
            n = n + 1
            ReDim Preserve actions(1 To n)
            actions(n).Kind = "new_record"
            actions(n).RowInstanceKey = key
            Set actions(n).Values = CloneStringDict(rowValues)
            actions(n).Reason = "no known slide instance carries this row's instance key"
        Else
            Dim idx As Long
            idx = knownByKey(key)
            Dim instanceSlide As Object
            Set instanceSlide = instances(idx)

            Dim changed As Object
            Set changed = CreateObject("Scripting.Dictionary")

            Dim fieldName As Variant
            For Each fieldName In rowValues.Keys
                Dim sourceValue As String
                sourceValue = rowValues(fieldName)

                Dim r As InjectResult
                r = InjectPrimitive.InjectPrimitive(instanceSlide, CStr(fieldName), sourceValue)

                ' r.Found = False covers both "no shape carries this
                ' field's tag" (skip -- matches resolve.py's
                ' field_shapes.get() returning None -> skip) and "more
                ' than one shape carries it" (ambiguous -- also skipped
                ' here, not separately flagged). Deliberately conflating
                ' those two Python-distinguishable situations into one
                ' skip outcome: disambiguating structural drift like a
                ' duplicate role tag is case-7 (deck_side_conflict)
                ' adjacent territory, a non-goal per
                ' specs/sync-operations.md. See SPIKE_NOTES_Resolve.md.
                If r.Found And r.Written Then
                    changed(fieldName) = r
                End If
            Next fieldName

            n = n + 1
            ReDim Preserve actions(1 To n)
            If changed.count > 0 Then
                actions(n).Kind = "in_place_correction"
                actions(n).InstanceKey = key
                Set actions(n).ChangedFields = changed
            Else
                actions(n).Kind = "no_change"
                actions(n).InstanceKey = key
            End If
        End If
    Next instanceId

    PlanRoutineSync = actions
End Function

' Case 2, per specs/sync-operations.md: only ever called explicitly against
' one specific known instance, never dispatched from PlanRoutineSync.
Public Function PlanPeriodRollover(instance As SlideInstance, newValues As Object) As PeriodRollover
    If Not instance.HasInstanceKey Then
        Err.Raise vbObjectError + 1, "SyncOperations.PlanPeriodRollover", _
            "cannot roll over a period for an unclassified instance (no instance_key)"
    End If

    Dim result As PeriodRollover
    result.SourceInstanceKey = instance.InstanceKey
    Set result.NewValues = CloneStringDict(newValues)
    result.Reason = "explicit period-rollover command"
    PlanPeriodRollover = result
End Function

Private Function CloneStringDict(source As Object) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In source.Keys
        result(k) = source(k)
    Next k
    Set CloneStringDict = result
End Function

' ---------------------------------------------------------------------
' Manual smoke test -- not a real test harness, same as every other module
' here. See SPIKE_NOTES_Resolve.md for the full recipe and expected
' values, cross-checked against tests/test_resolve.py's already-proven
' end-to-end results.
' ---------------------------------------------------------------------

' Run with slide 1 of the active presentation already tagged (via the
' Immediate window, before running this):
'   Application.ActivePresentation.Slides(1).Tags.Add "slide_type", "quarterly-update"
'   Application.ActivePresentation.Slides(1).Tags.Add "instance_key", "rec-1"
'   Application.ActivePresentation.Slides(1).Shapes(1).Tags.Add "role", "Title"
' (pick the actual title shape index/name for your slide if it isn't
' Shapes(1)). Seeds the Data-sheet row with the shape's own current text,
' so this should report "no_change" first, then "in_place_correction"
' after you change TITLE_TEXT below to something else and re-run.
Public Sub ManualSmokeTest_NoChangeThenInPlaceCorrection()
    Const TITLE_TEXT As String = "" ' leave blank to seed from the shape's current text (expect no_change)

    Dim sld As Object
    Set sld = Application.ActivePresentation.Slides(1)

    Dim instances(1 To 1) As Object
    Set instances(1) = sld

    Dim order As New Collection
    order.Add "rec-1"

    Dim rowsDict As Object
    Set rowsDict = CreateObject("Scripting.Dictionary")
    Dim rec1Fields As Object
    Set rec1Fields = CreateObject("Scripting.Dictionary")
    If TITLE_TEXT = "" Then
        rec1Fields("Title") = sld.Shapes(1).TextFrame.TextRange.Text
    Else
        rec1Fields("Title") = TITLE_TEXT
    End If
    Set rowsDict("rec-1") = rec1Fields

    Dim actions() As SyncAction
    actions = PlanRoutineSync(instances, order, rowsDict)

    Dim msg As String
    msg = "actions=" & (UBound(actions) - LBound(actions) + 1) & " Kind=" & actions(1).Kind & _
        " InstanceKey=" & actions(1).InstanceKey
    Debug.Print msg
    MsgBox msg & " (expected: no_change when TITLE_TEXT is blank, in_place_correction otherwise)"
End Sub
