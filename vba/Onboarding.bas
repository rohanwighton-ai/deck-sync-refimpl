Attribute VB_Name = "Onboarding"
Option Explicit

' VBA port of src/onboarding.py's match_slide_against_template()/
' onboard_new_instance()/confirm_field_match(), per specs/vba-port.md's
' port order (module 5, after discovery/identity_tags/matching/resolve).
'
' First-time onboarding of a type needs no code here at all -- per
' specs/onboarding.md, "the working copy IS becoming the reference": a first
' example slide is discovered, confirmed with a human, and tagged directly
' via the object model (Slide.Tags.Add / Shape.Tags.Add, exactly as
' SyncOperations.bas's own ManualSmokeTest already demonstrates). This
' module is specifically the gap onboarding.md identifies as real: matching
' a *subsequent* slide's candidates against an already-established template.
'
' See SPIKE_NOTES_Onboarding.md for the full divergence list, a real
' UDT/Dictionary gotcha found while building this, and the manual
' verification recipe.

Public Type FieldMatch
    Role As String
    Result As MatchResult
End Type

' ---------------------------------------------------------------------
' Template construction
' ---------------------------------------------------------------------

' Given an already-onboarded slide, build its field-shape reference set:
' every discovered candidate whose live shape carries a role tag, paired
' with that role. Mirrors resolve.py's resolve_slide_instance() -- but only
' the field_shapes half (slide_type/instance_key aren't needed by matching,
' only by SyncOperations' own dispatch, which Resolve.bas already covers) --
' and deliberately does NOT reuse Resolve.SlideInstance's shape (it carries
' no field_shapes at all; see Resolve.bas's own header comment on why that
' was left out, and SPIKE_NOTES_Onboarding.md's divergence 1 for why this
' module doesn't reopen that type rather than building its own).
'
' Returns parallel arrays (roles(), and the Candidate() return value) rather
' than a role->Candidate map: Candidate is a UDT, and VBA cannot store a UDT
' inside a Scripting.Dictionary or Variant (a genuine compile-time
' restriction, not a style choice -- see SPIKE_NOTES_Onboarding.md).
Public Function BuildTemplateFieldShapes(templateSld As Object, ByRef roles() As String) As Candidate()
    Dim allCandidates() As Candidate
    Dim allShapes() As Object
    allCandidates = Discovery.DiscoverSlideWithShapes(templateSld, allShapes)

    ' results()/roles() are deliberately NOT pre-ReDim'd to (1 To 0) -- a
    ' real, confirmed VBA restriction (AGENTS.md's Known Patterns) makes
    ' that throw "Subscript out of range" at runtime. Left unallocated
    ' instead; the first ReDim Preserve below allocates correctly on its
    ' own. A template with zero tagged fields returns a genuinely
    ' unallocated array -- callers must error-guard, never assume allocated.
    Dim results() As Candidate
    Dim n As Long
    n = 0

    Dim i As Long, lo As Long, hi As Long
    On Error Resume Next
    lo = LBound(allCandidates)
    hi = UBound(allCandidates)
    If Err.Number <> 0 Then
        BuildTemplateFieldShapes = results ' both results and roles stay unallocated -- zero candidates on the template slide
        Exit Function
    End If
    On Error GoTo 0

    For i = lo To hi
        Dim role As String
        role = allShapes(i).Tags("role")
        If role <> "" Then
            n = n + 1
            ReDim Preserve results(1 To n)
            ReDim Preserve roles(1 To n)
            results(n) = allCandidates(i)
            roles(n) = role
        End If
    Next i

    BuildTemplateFieldShapes = results
End Function

' ---------------------------------------------------------------------
' Matching a subsequent slide against the template
' ---------------------------------------------------------------------

' For each field role the template defines (templateRoles()/
' templateFieldShapes(), from BuildTemplateFieldShapes, same indices),
' score every untagged candidate field on `sld` against the template's
' reference shape for that role, per specs/matching.md's tier-2 path.
' Already-tagged candidates are excluded (either this same field from a
' prior pass, or a different field entirely -- neither is a fresh match
' target). Pure decoration (no text, not a picture) is excluded too, per
' specs/discovery.md -- src/onboarding.py's own Priority-10 fix
' (is_candidate_field) is what this mirrors.
'
' `untaggedShapes` is an out-parameter: the live Shape objects behind the
' untagged candidate pool actually scored, at the same indices Matching.
' Match()'s MatchResult.CandidateIndex refers to. Simple callers that only
' want the match report can pass a throwaway local array and ignore it;
' OnboardNewInstance needs it to actually write a tag onto whichever shape a
' high-confidence match accepted.
Public Function MatchSlideAgainstTemplate(sld As Object, templateRoles() As String, templateFieldShapes() As Candidate, ByRef untaggedShapes() As Object) As FieldMatch()
    Dim allCandidates() As Candidate
    Dim allShapes() As Object
    allCandidates = Discovery.DiscoverSlideWithShapes(sld, allShapes)

    ' untagged()/untaggedShapes()/results() are deliberately NOT pre-ReDim'd
    ' to (1 To 0) -- see BuildTemplateFieldShapes' identical comment. Every
    ' LBound/UBound call on a possibly-unallocated array below is
    ' error-guarded rather than assumed safe.
    Dim untagged() As Candidate
    Dim n As Long
    n = 0

    Dim i As Long, allLo As Long, allHi As Long
    On Error Resume Next
    allLo = LBound(allCandidates)
    allHi = UBound(allCandidates)
    Dim hasAllCandidates As Boolean
    hasAllCandidates = (Err.Number = 0)
    On Error GoTo 0

    If hasAllCandidates Then
        For i = allLo To allHi
            If IsCandidateField(allCandidates(i)) And allShapes(i).Tags("role") = "" Then
                n = n + 1
                ReDim Preserve untagged(1 To n)
                ReDim Preserve untaggedShapes(1 To n)
                untagged(n) = allCandidates(i)
                Set untaggedShapes(n) = allShapes(i)
            End If
        Next i
    End If

    Dim results() As FieldMatch
    Dim roleLo As Long, roleHi As Long
    On Error Resume Next
    roleLo = LBound(templateRoles)
    roleHi = UBound(templateRoles)
    Dim hasRoles As Boolean
    hasRoles = (Err.Number = 0)
    On Error GoTo 0

    If hasRoles Then
        Dim outIdx As Long, idx As Long
        outIdx = 0
        For idx = roleLo To roleHi
            outIdx = outIdx + 1
            ReDim Preserve results(1 To outIdx)
            results(outIdx).Role = templateRoles(idx)
            results(outIdx).Result = Matching.Match(untagged, templateFieldShapes(idx))
        Next idx
    End If

    MatchSlideAgainstTemplate = results
End Function

' Shape.HasText/is-picture check discovery.py's is_candidate_field property
' expresses on Candidate itself -- kept here rather than added to
' Discovery.Candidate because the filter belongs at the point candidates are
' handed to the matcher (specs/discovery.md's own framing, which
' src/onboarding.py's docstring quotes directly), not inside discovery
' itself (verify_structure and other callers still need the full,
' unfiltered shape list).
Private Function IsCandidateField(c As Candidate) As Boolean
    IsCandidateField = (c.ShapeType = "picture") Or c.HasText
End Function

' ---------------------------------------------------------------------
' Confirmation and instance onboarding
' ---------------------------------------------------------------------

' Write `role` onto `shp` -- the confirmation primitive for a match a human
' has decided on, whether that's accepting a medium-confidence result
' MatchSlideAgainstTemplate already scored, or a direct selection an
' eventual onboarding UI would make (specs/onboarding.md's Non-goals: this
' is the primitive that UI would call, not the UI itself). Shape.Tags.Add
' upserts natively (if "role" already exists on this shape, its value is
' replaced, not duplicated) -- no read-merge-write needed, unlike the Python
' original's hand-rolled OOXML.
Public Sub ConfirmFieldMatch(shp As Object, role As String)
    shp.Tags.Add "role", role
End Sub

' Tag a new instance's slide-level identity (supplied by whatever created
' it -- e.g. a sync-operations duplication, per specs/slide-duplication-
' trigger.md -- not matched) and auto-accept any high-confidence field
' matches against the template, writing their tags immediately so they
' become a tier-1 fast match next time (specs/matching.md's confidence_
' thresholds: "the system gets more self-healing over time, not less").
' Medium/low-confidence matches are returned but never auto-tagged -- a
' human decides via ConfirmFieldMatch, same as any other unresolved match.
'
' Slide.Tags.Add upserts the same way Shape.Tags.Add does, so this is
' unconditional and idempotent -- safe to call even if the slide already
' carries these tags from a prior pass.
Public Function OnboardNewInstance(sld As Object, templateRoles() As String, templateFieldShapes() As Candidate, slideType As String, instanceKey As String) As FieldMatch()
    sld.Tags.Add "slide_type", slideType
    sld.Tags.Add "instance_key", instanceKey

    Dim untaggedShapes() As Object
    Dim matches() As FieldMatch
    matches = MatchSlideAgainstTemplate(sld, templateRoles, templateFieldShapes, untaggedShapes)

    Dim i As Long, lo As Long, hi As Long, hasMatches As Boolean
    On Error Resume Next
    lo = LBound(matches)
    hi = UBound(matches)
    hasMatches = (Err.Number = 0)
    On Error GoTo 0

    If hasMatches Then
        For i = lo To hi
            If matches(i).Result.Confidence = "high" And matches(i).Result.HasCandidate Then
                ConfirmFieldMatch untaggedShapes(matches(i).Result.CandidateIndex), matches(i).Role
            End If
        Next i
    End If

    OnboardNewInstance = matches
End Function

' ---------------------------------------------------------------------
' Manual smoke tests -- not a real test harness, same as every other module
' here. See SPIKE_NOTES_Onboarding.md for the full recipe and expected
' values, cross-checked against tests/test_onboarding.py's already-proven
' results.
' ---------------------------------------------------------------------

' Run with test-fixtures/mst-slide-layouts.pptx's two layouts applied to
' slides 1 (template, already tagged) and 2 (new, untagged) of the active
' presentation, per SPIKE_NOTES_Onboarding.md's setup steps. Expects Title
' to score high (auto-tagged) and Body to score medium (left unresolved),
' matching tests/test_onboarding.py::test_onboard_new_instance_auto_accepts_
' high_confidence_and_leaves_medium_unresolved.
Public Sub ManualSmokeTest_OnboardNewInstance()
    Dim templateSld As Object, newSld As Object
    Set templateSld = Application.ActivePresentation.Slides(1)
    Set newSld = Application.ActivePresentation.Slides(2)

    Dim templateRoles() As String
    Dim templateFieldShapes() As Candidate
    templateFieldShapes = BuildTemplateFieldShapes(templateSld, templateRoles)

    Dim matches() As FieldMatch
    matches = OnboardNewInstance(newSld, templateRoles, templateFieldShapes, "quarterly-update", "rec-2")

    Dim i As Long, msg As String, lo As Long, hi As Long
    On Error Resume Next
    lo = LBound(matches)
    hi = UBound(matches)
    If Err.Number = 0 Then
        On Error GoTo 0
        For i = lo To hi
            msg = msg & matches(i).Role & ": confidence=" & matches(i).Result.Confidence & _
                " hasCandidate=" & matches(i).Result.HasCandidate & vbCrLf
        Next i
    Else
        On Error GoTo 0
        msg = "(no field matches -- template had zero tagged roles)"
    End If
    Debug.Print msg
    MsgBox msg & "(expected: one role high/True (auto-tagged), one role medium/False)"
End Sub
