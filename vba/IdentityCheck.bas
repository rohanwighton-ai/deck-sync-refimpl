Attribute VB_Name = "IdentityCheck"
Option Explicit

' R9 / D5 of the Excel Control Layer exchange (31 July 2026): identity tags
' must be checked for uniqueness before any run writes.
'
' The requirement is not theoretical. Probed against real Office 2026-07-31:
'
'   source slide:            SlideID 257, instance_key 'PROBE-001'
'   after Slide.Duplicate:   SlideID 258, instance_key 'PROBE-001'
'   after paste to new deck: SlideID 257, instance_key 'PROBE-001'
'
' So identity tags clone exactly the way shape NAMES do -- which is the same
' failure that got ShapeName rejected as a join key. Two consequences:
'
'   - duplicating a project slide produces two slides both claiming one
'     EntityCode, and nothing in the deck says so
'   - copying last quarter's deck to start this quarter carries the declared
'     period with it, so a Q4 deck reports itself as Q3
'
' FieldID (the shape `role` tag) cloning is CORRECT and wanted -- every project
' slide should carry PROGRESS_BODY. It is only the identity tags that must not
' duplicate. This module therefore checks instance_key and deliberately does
' not look at role tags at all.
'
' Why the check has to carry the whole load for the copied-deck case: SlideID
' IS reassigned on a within-deck duplicate, so PowerPoint natively distinguishes
' those two slides. But a slide pasted into another presentation kept SlideID
' 257 -- so across files there is no native discriminator, and copying a deck to
' start a new quarter is the OBVIOUS workflow, not an edge case. There is
' nothing else standing between that and two slides silently sharing a key.

Public Type DuplicateKeyReport
    HasDuplicates As Boolean
    Count As Long              ' number of KEYS duplicated, not slides affected
    Detail As String           ' one line per duplicated key, naming every slide
End Type

' Every instance_key claimed by more than one slide of `slideType`.
'
' Uses RunSync.GatherInstances, so the master template is excluded by
' construction -- a template is deliberately keyless and is not a competing
' claim on anything.
Public Function FindDuplicateKeys(slideType As String) As DuplicateKeyReport
    Dim result As DuplicateKeyReport

    Dim instances() As Object
    instances = RunSync.GatherInstances(slideType)

    Dim lo As Long, hi As Long, hasAny As Boolean
    On Error Resume Next
    lo = LBound(instances): hi = UBound(instances)
    hasAny = (Err.Number = 0)
    On Error GoTo 0
    If Not hasAny Then
        FindDuplicateKeys = result
        Exit Function
    End If

    ' key -> "3, 7, 11" (the slide indexes claiming it). Built as a plain
    ' string per key rather than a nested structure, because VBA cannot hold a
    ' UDT or an array inside a Dictionary (AGENTS.md, hit three times).
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = lo To hi
        Dim inst As SlideInstance
        inst = Resolve.ResolveSlideInstance(instances(i))
        If inst.HasInstanceKey Then
            Dim k As String
            k = inst.InstanceKey
            If seen.Exists(k) Then
                seen(k) = seen(k) & ", " & instances(i).SlideIndex
            Else
                seen(k) = CStr(instances(i).SlideIndex)
            End If
        End If
    Next i

    Dim key As Variant
    For Each key In seen.Keys
        If InStr(seen(key), ",") > 0 Then
            result.Count = result.Count + 1
            result.Detail = result.Detail & _
                "  '" & key & "' is claimed by slides " & seen(key) & vbCrLf
        End If
    Next key

    result.HasDuplicates = (result.Count > 0)
    FindDuplicateKeys = result
End Function

' The warning a human sees. Pure, so the wording is testable without a deck --
' same reason RunSync.ConfirmSyncText is.
'
' States the CONSEQUENCE, not just the fact. "Two slides share a key" means
' nothing to someone who has not read the matching spec; "only one of them will
' ever be updated, and which one is not defined" is the thing they need to
' decide on.
Public Function DuplicateKeyWarningText(slideType As String, report As DuplicateKeyReport) As String
    If Not report.HasDuplicates Then
        DuplicateKeyWarningText = ""
        Exit Function
    End If

    DuplicateKeyWarningText = _
        "DUPLICATE IDENTITY TAGS in '" & slideType & "'." & vbCrLf & vbCrLf & _
        report.Detail & vbCrLf & _
        "Each key should belong to exactly one slide. Duplicates almost always" & vbCrLf & _
        "mean a slide was copied -- identity tags clone on copy, the same way" & vbCrLf & _
        "shape names do." & vbCrLf & vbCrLf & _
        "Consequence if you continue: only ONE of the slides sharing a key will" & vbCrLf & _
        "be updated from the Data sheet, and which one is not defined. The" & vbCrLf & _
        "others will silently keep whatever they show now." & vbCrLf & vbCrLf & _
        "Fix the keys before syncing."
End Function
