Attribute VB_Name = "PlaceholderCheck"
Option Explicit

' V7, closing Amendment A of the Excel Control Layer exchange.
'
' D6 originally required that a surviving `<<...>>` "blocks publication and
' marks the output draft". There is no draft state in this tool and none is
' being built -- the word came out of D6 deliberately. What replaces it:
'
'   the placeholder text IS the marker.
'
' That was the recommendation and it was accepted. `<<PROJECT_STATUS>>` sitting
' on a slide is unmissable in a way no filename suffix, log line or watermark
' is, and inventing a second marker before the first has been shown to fail is
' how you end up with two that nobody reads.
'
' This module adds the machine-readable half: a deck property recording what
' the last run found, so the condition is queryable rather than only visible.
'
' THE STALENESS PROBLEM, and it is the Excel side's point rather than mine.
' A bare count cannot be told apart from a count left by an earlier run --
' "3 placeholders" might be from this morning's sync or from one three weeks
' ago against a different quarter, and those demand opposite responses. So the
' marker carries the PERIOD and the TIMESTAMP alongside the count, and reading
' it answers "is this current?" rather than merely "what was the number?".
' A marker that cannot be dated is barely better than no marker.

' `count|period|timestamp`. Pipe as separator, consistent with
' DeckRegistry.BuildTypeRegistration, and safe for the same reason: neither a
' period literal nor an ISO timestamp can contain one.
Public Const PLACEHOLDER_PROPERTY_NAME As String = "DeckSyncPlaceholders"

Public Type PlaceholderReport
    Count As Long
    Detail As String          ' one line per placeholder found, naming slide and field
End Type

' Every managed field still showing its template placeholder.
'
' Detection is an EXACT match against TemplateSlide.PlaceholderFor(the shape's
' OWN role tag) -- not a search for "<<" anywhere in the text. A substring
' search would flag legitimate content that happens to contain angle brackets,
' and false positives on a publication gate are how a gate gets ignored.
'
' Uses RunSync.GatherInstances, so the master template is excluded: its fields
' are SUPPOSED to read `<<...>>` forever. Counting them would mean the marker
' could never reach zero, and a warning that always fires stops being read --
' the same failure as an always-true guard.
Public Function FindPlaceholders(slideType As String) As PlaceholderReport
    Dim result As PlaceholderReport

    Dim instances() As Object
    instances = RunSync.GatherInstances(slideType)

    Dim lo As Long, hi As Long, hasAny As Boolean
    On Error Resume Next
    lo = LBound(instances): hi = UBound(instances)
    hasAny = (Err.Number = 0)
    On Error GoTo 0
    If Not hasAny Then
        FindPlaceholders = result
        Exit Function
    End If

    Dim i As Long
    For i = lo To hi
        Dim inst As SlideInstance
        inst = Resolve.ResolveSlideInstance(instances(i))
        WalkForPlaceholders instances(i).Shapes, _
            IIf(inst.HasInstanceKey, inst.InstanceKey, "<unkeyed>"), _
            instances(i).SlideIndex, result
    Next i

    FindPlaceholders = result
End Function

' Recurses groups -- managed fields are nested on real decks, and a flat walk
' would under-report, which on a publication gate means passing a deck that
' should have been stopped.
Private Sub WalkForPlaceholders(shapesColl As Object, instanceKey As String, _
                                slideIndex As Long, ByRef result As PlaceholderReport)
    Dim shp As Object
    For Each shp In shapesColl
        If shp.Type = msoGroup Then
            WalkForPlaceholders shp.GroupItems, instanceKey, slideIndex, result
        Else
            Dim role As String
            role = shp.Tags("role")
            If role <> "" Then
                Dim txt As String
                On Error Resume Next
                txt = shp.TextFrame.TextRange.Text
                On Error GoTo 0

                If Trim(txt) = TemplateSlide.PlaceholderFor(role) Then
                    result.Count = result.Count + 1
                    result.Detail = result.Detail & _
                        "  slide " & slideIndex & " (" & instanceKey & "): " & role & vbCrLf
                End If
            End If
        End If
    Next shp
End Sub

' ---------------------------------------------------------------------
' The deck marker
' ---------------------------------------------------------------------

Public Function BuildMarker(count As Long, period As String, stamp As String) As String
    BuildMarker = CStr(count) & "|" & period & "|" & stamp
End Function

' False when the property is absent or malformed -- both are routine states
' (a deck that has never been synced, or one written by an older build), not
' errors to raise.
Public Function ParseMarker(raw As String, ByRef count As Long, ByRef period As String, ByRef stamp As String) As Boolean
    count = 0: period = "": stamp = ""
    If raw = "" Then Exit Function

    Dim parts() As String
    parts = Split(raw, "|")
    If UBound(parts) <> 2 Then Exit Function
    If Not IsNumeric(parts(0)) Then Exit Function

    count = CLng(parts(0))
    period = parts(1)
    stamp = parts(2)
    ParseMarker = True
End Function

Public Sub WriteMarker(pres As Object, count As Long, period As String, stamp As String)
    Dim value As String
    value = BuildMarker(count, period, stamp)

    Dim prop As Object
    On Error Resume Next
    Set prop = pres.CustomDocumentProperties(PLACEHOLDER_PROPERTY_NAME)
    On Error GoTo 0

    If prop Is Nothing Then
        pres.CustomDocumentProperties.Add Name:=PLACEHOLDER_PROPERTY_NAME, _
            LinkToContent:=False, Type:=msoPropertyTypeString, Value:=value
    Else
        prop.Value = value
    End If
End Sub

Public Function ReadMarker(pres As Object) As String
    Dim prop As Object
    On Error Resume Next
    Set prop = pres.CustomDocumentProperties(PLACEHOLDER_PROPERTY_NAME)
    On Error GoTo 0
    If prop Is Nothing Then Exit Function
    ReadMarker = CStr(prop.Value)
End Function

' ---------------------------------------------------------------------
' What a human reads
' ---------------------------------------------------------------------

' The run report's HEADLINE when placeholders survive -- Amendment A asks for
' the list itself rather than a summary line, on the grounds that a count in a
' summary is skimmable and a named list is not.
'
' Pure, so the wording is testable.
Public Function HeadlineText(r As PlaceholderReport, period As String) As String
    If r.Count = 0 Then
        HeadlineText = ""
        Exit Function
    End If

    HeadlineText = _
        "NOT READY TO PUBLISH -- " & r.Count & " field(s) still show their template" & vbCrLf & _
        "placeholder for period '" & period & "':" & vbCrLf & vbCrLf & _
        r.Detail & vbCrLf & _
        "Each of these has no approved register row for this period, so nothing" & vbCrLf & _
        "was written and the template's scaffolding is still showing. They will" & vbCrLf & _
        "appear on the slide as <<FIELD_NAME>>."
End Function

' Whether a stored marker still describes the deck's CURRENT period.
'
' This is the whole reason the period and timestamp are stored. "3 placeholders"
' is unactionable on its own: from this morning it means fix them, from last
' quarter it means ignore it. Returning the distinction rather than the number
' is what makes the property worth having.
Public Function MarkerStatusText(raw As String, currentPeriod As String) As String
    Dim count As Long, period As String, stamp As String

    If Not ParseMarker(raw, count, period, stamp) Then
        MarkerStatusText = "No placeholder record on this deck -- it has not been synced by a build that writes one."
        Exit Function
    End If

    If StrComp(period, currentPeriod, vbTextCompare) <> 0 Then
        MarkerStatusText = "STALE record: " & count & " placeholder(s) were found for period '" & period & _
            "' at " & stamp & ", but this deck now declares '" & currentPeriod & "'." & vbCrLf & _
            "Re-run the sync before trusting that number -- it describes a different quarter."
        Exit Function
    End If

    If count = 0 Then
        MarkerStatusText = "Last run (" & stamp & ") found no placeholders for '" & period & "'."
    Else
        MarkerStatusText = count & " placeholder(s) outstanding for '" & period & "', as at " & stamp & "."
    End If
End Function
