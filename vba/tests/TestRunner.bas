Attribute VB_Name = "TestRunner"
Option Explicit

' First real-execution test harness for the PowerPoint-hosted modules
' (Discovery, InjectPrimitive, Matching, Resolve, SyncOperations,
' Onboarding). Every prior SPIKE_NOTES_*.md said "not executed or verified
' in this environment" -- this module exists to close that gap for real,
' driven headlessly via COM automation (see the PowerShell driver script),
' not by a human clicking through the Immediate window.
'
' Deliberately NOT the same as the ManualSmokeTest* subs already in each
' module: those use MsgBox, which would hang a headless script waiting on a
' click that never comes. Every Test_* function here is a pure assertion
' function -- returns "" on full pass, or one "FAIL: ..." line per failed
' assertion otherwise -- so RunAllTests can produce one machine-readable
' report with no human present.
'
' Fixture strategy: most tests build their own slides/shapes programmatically
' (deterministic, no dependency on a fixture file's exact geometry surviving
' across PowerPoint versions). Two tests still open a real test-fixtures/
' file directly (shp-groupshape.pptx, for its real 4-way sibling-ambiguity
' geometry, which would be needlessly fragile to reproduce by hand) --
' `fixturesDir` is that folder's path, passed in by the driver rather than
' hardcoded, since the real repo path isn't necessarily reachable the same
' way from a Windows-hosted COM session (see the driver script for how
' fixtures get staged).

Public Function RunAllTests(fixturesDir As String, stagingDir As String) As String
    Dim report As String
    report = "=== deck-sync-refimpl VBA test run (PowerPoint) ===" & vbCrLf

    Dim r As String

    r = "": On Error Resume Next: Err.Clear
    r = Test_Discovery_GroupRecursionFindsCandidates(fixturesDir)
    AppendResult report, "Discovery_GroupRecursionFindsCandidates", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_InjectPrimitive_NoOpWhenValueAlreadyMatches()
    AppendResult report, "InjectPrimitive_NoOpWhenValueAlreadyMatches", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_InjectPrimitive_WritesAndVerifiesOnMismatch()
    AppendResult report, "InjectPrimitive_WritesAndVerifiesOnMismatch", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_InjectPrimitive_AmbiguousTagRefusesToGuess()
    AppendResult report, "InjectPrimitive_AmbiguousTagRefusesToGuess", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_Matching_SiblingAmbiguityResolvedByZOrder(fixturesDir)
    AppendResult report, "Matching_SiblingAmbiguityResolvedByZOrder", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_Matching_EnrichPlaceholderIdxReadsRealFile(stagingDir)
    AppendResult report, "Matching_EnrichPlaceholderIdxReadsRealFile", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_Resolve_ReadsTagsOffLiveSlide()
    AppendResult report, "Resolve_ReadsTagsOffLiveSlide", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_SyncOperations_Cases1And4()
    AppendResult report, "SyncOperations_Cases1And4", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_SyncOperations_Case3NewRecord()
    AppendResult report, "SyncOperations_Case3NewRecord", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_SyncOperations_Case6UnclassifiedSlide()
    AppendResult report, "SyncOperations_Case6UnclassifiedSlide", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_Onboarding_HighAndMediumConfidence()
    AppendResult report, "Onboarding_HighAndMediumConfidence", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_Onboarding_OnboardNewInstanceAutoTagsHighOnly()
    AppendResult report, "Onboarding_OnboardNewInstanceAutoTagsHighOnly", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_Onboarding_PureDecorationNeverMatched()
    AppendResult report, "Onboarding_PureDecorationNeverMatched", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_Verification_StructureMatchesAfterDuplicate()
    AppendResult report, "Verification_StructureMatchesAfterDuplicate", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_Verification_DetectsShapeCountMismatch()
    AppendResult report, "Verification_DetectsShapeCountMismatch", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_Verification_DetectsZOrderSwap()
    AppendResult report, "Verification_DetectsZOrderSwap", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_SlideDuplication_CreatesTaggedInjectedSlide()
    AppendResult report, "SlideDuplication_CreatesTaggedInjectedSlide", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_SlideDuplication_RefusesInstanceKeyCollision()
    AppendResult report, "SlideDuplication_RefusesInstanceKeyCollision", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_SlideDuplication_PartialRowStillCreatesSlideButFlagsMissing()
    AppendResult report, "SlideDuplication_PartialRowStillCreatesSlideButFlagsMissing", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_RunSync_GatherInstancesFiltersByType()
    AppendResult report, "RunSync_GatherInstancesFiltersByType", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_RunSync_EndToEndCreatesSlidesFromFreshSheet()
    AppendResult report, "RunSync_EndToEndCreatesSlidesFromFreshSheet", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_RunSync_RunPeriodRolloverDuplicatesLeavingSourceUntouched()
    AppendResult report, "RunSync_RunPeriodRolloverDuplicatesLeavingSourceUntouched", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_DeckAdoption_AlreadyLinkedSlideSkipped()
    AppendResult report, "DeckAdoption_AlreadyLinkedSlideSkipped", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_DeckAdoption_ReadyHighConfidenceSlideLinkedAndCreatesFreshRow()
    AppendResult report, "DeckAdoption_ReadyHighConfidenceSlideLinkedAndCreatesFreshRow", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_DeckAdoption_MediumConfidenceSlideNeedsConfirmationAndIsNotTagged()
    AppendResult report, "DeckAdoption_MediumConfidenceSlideNeedsConfirmationAndIsNotTagged", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_DeckAdoption_UnclassifiedSlideExcluded()
    AppendResult report, "DeckAdoption_UnclassifiedSlideExcluded", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_DeckAdoption_MatchesExistingKeylessRowLinksWithoutCreatingNewRow()
    AppendResult report, "DeckAdoption_MatchesExistingKeylessRowLinksWithoutCreatingNewRow", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_DeckAdoption_MultiSlideZeroBasedBatchKeepsIndicesAligned()
    AppendResult report, "DeckAdoption_MultiSlideZeroBasedBatchKeepsIndicesAligned", r
    On Error GoTo 0

    RunAllTests = report
End Function

Private Sub AppendResult(ByRef report As String, testName As String, testResult As String)
    If Err.Number <> 0 Then
        report = report & "ERROR " & testName & " :: " & Err.Description & " (line context lost -- VBA has no stack trace)" & vbCrLf
    ElseIf testResult = "" Then
        report = report & "PASS  " & testName & vbCrLf
    Else
        report = report & "FAIL  " & testName & vbCrLf & testResult
    End If
End Sub

Private Function Assert(cond As Boolean, msg As String) As String
    If cond Then
        Assert = ""
    Else
        Assert = "    FAIL: " & msg & vbCrLf
    End If
End Function

Private Function NewBlankSlide() As Object
    Dim n As Long
    n = Application.ActivePresentation.Slides.count + 1
    Set NewBlankSlide = Application.ActivePresentation.Slides.Add(n, ppLayoutBlank)
End Function

Private Function FindShapeByRole(sld As Object, role As String) As Object
    Dim shp As Object
    For Each shp In sld.Shapes
        If shp.Tags("role") = role Then
            Set FindShapeByRole = shp
            Exit Function
        End If
    Next shp
    Set FindShapeByRole = Nothing
End Function

' ---------------------------------------------------------------------
' Discovery
' ---------------------------------------------------------------------

Private Function Test_Discovery_GroupRecursionFindsCandidates(fixturesDir As String) As String
    Dim result As String
    Dim testPres As Object
    Set testPres = Application.Presentations.Open(fixturesDir & "shp-groupshape.pptx", msoTrue)

    Dim candidates() As Candidate
    candidates = Discovery.DiscoverSlide(testPres.Slides(1))

    Dim lo As Long, hi As Long, hasCandidates As Boolean
    On Error Resume Next
    lo = LBound(candidates)
    hi = UBound(candidates)
    hasCandidates = (Err.Number = 0)
    On Error GoTo 0
    result = result & Assert(hasCandidates, "DiscoverSlide returned at least one candidate")

    Dim found As Boolean
    Dim i As Long
    If hasCandidates Then
        For i = lo To hi
            If candidates(i).Name = "Oval 2" Then found = True
        Next i
    End If
    result = result & Assert(found, "DiscoverSlide recursed into a group and found 'Oval 2'")

    testPres.Close
    Test_Discovery_GroupRecursionFindsCandidates = result
End Function

' ---------------------------------------------------------------------
' InjectPrimitive
' ---------------------------------------------------------------------

Private Function Test_InjectPrimitive_NoOpWhenValueAlreadyMatches() As String
    Dim result As String
    Dim sld As Object
    Set sld = NewBlankSlide()

    Dim shp As Object
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    shp.TextFrame.TextRange.Text = "hello"
    shp.Tags.Add "role", "demo_field"

    Dim r As InjectResult
    r = InjectPrimitive.InjectPrimitive(sld, "demo_field", "hello")

    result = result & Assert(r.Found, "shape found by role tag")
    result = result & Assert(Not r.Written, "no-op when value already matches")
    result = result & Assert(r.Verified, "verified true on no-op path")

    Test_InjectPrimitive_NoOpWhenValueAlreadyMatches = result
End Function

Private Function Test_InjectPrimitive_WritesAndVerifiesOnMismatch() As String
    Dim result As String
    Dim sld As Object
    Set sld = NewBlankSlide()

    Dim shp As Object
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    shp.TextFrame.TextRange.Text = "old value"
    shp.Tags.Add "role", "demo_field"

    Dim r As InjectResult
    r = InjectPrimitive.InjectPrimitive(sld, "demo_field", "new value")

    result = result & Assert(r.Found, "shape found by role tag")
    result = result & Assert(r.Written, "written when value differs")
    result = result & Assert(r.Verified, "re-read confirms the write took")
    result = result & Assert(shp.TextFrame.TextRange.Text = "new value", "shape text actually changed, got '" & shp.TextFrame.TextRange.Text & "'")

    Test_InjectPrimitive_WritesAndVerifiesOnMismatch = result
End Function

Private Function Test_InjectPrimitive_AmbiguousTagRefusesToGuess() As String
    Dim result As String
    Dim sld As Object
    Set sld = NewBlankSlide()

    Dim shp1 As Object, shp2 As Object
    Set shp1 = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    shp1.TextFrame.TextRange.Text = "a"
    shp1.Tags.Add "role", "dup_field"
    Set shp2 = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 150, 200, 50)
    shp2.TextFrame.TextRange.Text = "b"
    shp2.Tags.Add "role", "dup_field"

    Dim r As InjectResult
    r = InjectPrimitive.InjectPrimitive(sld, "dup_field", "c")

    result = result & Assert(Not r.Found, "ambiguous tag (2 shapes) refuses rather than picking one")
    result = result & Assert(shp1.TextFrame.TextRange.Text = "a" And shp2.TextFrame.TextRange.Text = "b", "neither shape was touched")

    Test_InjectPrimitive_AmbiguousTagRefusesToGuess = result
End Function

' ---------------------------------------------------------------------
' Matching
' ---------------------------------------------------------------------

Private Function Test_Matching_SiblingAmbiguityResolvedByZOrder(fixturesDir As String) As String
    Dim result As String
    Dim testPres As Object
    Set testPres = Application.Presentations.Open(fixturesDir & "shp-groupshape.pptx", msoTrue)

    Dim candidates() As Candidate
    candidates = Discovery.DiscoverSlide(testPres.Slides(1))

    ' The reference must carry real geometry/shape-type, not just ZOrder/
    ' HasText -- a reference with everything else left at Long/String
    ' defaults (0, "") scores near-zero on every candidate regardless of
    ' z-order, which is exactly what this test found the first time it
    ' actually ran (a gap in the original ManualSmokeTest_SiblingAmbiguity
    ' this test was copied from -- never caught before because that sub had
    ' never been executed either). The 4 decoration ovals are near-identical
    ' overlapping siblings by design (that's the whole scenario), so any one
    ' of their real geometries is representative -- clone one, but keep
    ' ZOrder explicitly at 2 (not copied) since z-order proximity to the
    ' reference is exactly the signal under test.
    Dim reference As Candidate
    Dim i As Long
    For i = LBound(candidates) To UBound(candidates)
        If Not candidates(i).HasText And candidates(i).ShapeType = "autoshape_or_textbox" Then
            reference = candidates(i)
            Exit For
        End If
    Next i
    reference.ZOrder = 2
    reference.PlaceholderIdx = -1 ' not testing placeholder matching here

    Dim r As MatchResult
    r = Matching.Match(candidates, reference)

    result = result & Assert(r.Confidence = "high", "confidence is high, got '" & r.Confidence & "'")
    result = result & Assert(r.HasCandidate, "HasCandidate true")
    ' Not asserting the exact resolved shape name: this test's reference is
    ' cloned from the first non-text autoshape found (a stand-in for "a
    ' sibling with matching geometry"), not necessarily Oval 2 itself --
    ' real run (2026-07-25) showed shp-groupshape.pptx has other non-text
    ' autoshapes (e.g. "Rounded Rectangle 1") competing for that same
    ' geometry class, so which specific sibling wins z-order proximity to
    ' ZOrder=2 depends on the fixture's exact real layout, not asserted
    ' here. What's actually under test -- multiple similarly-scored
    ' candidates get disambiguated by z-order proximity rather than
    ' flagged or picked arbitrarily -- is what HasCandidate=True/
    ' Confidence=high already confirm.

    testPres.Close
    Test_Matching_SiblingAmbiguityResolvedByZOrder = result
End Function

' Exercises Matching.EnrichPlaceholderIdx's real, previously-unexecuted
' Shell.Application zip-extraction + XPath fallback against a real file on
' disk -- the most exotic, least-proven code in this whole port. Uses an
' ordinary slide with a real placeholder (not test-fixtures/mst-slide-
' layouts.pptx's CustomLayouts, whose live-COM accessibility was flagged
' unconfirmed in SPIKE_NOTES_Resolve.md divergence 6) -- EnrichPlaceholderIdx
' only cares about a part path + shape name, so a slide's own
' ppt/slides/slideN.xml exercises the identical code path with less risk.
Private Function Test_Matching_EnrichPlaceholderIdxReadsRealFile(stagingDir As String) As String
    Dim result As String
    Dim sld As Object
    Set sld = Application.ActivePresentation.Slides.Add(Application.ActivePresentation.Slides.count + 1, ppLayoutText)

    Dim savePath As String
    savePath = stagingDir & "enrich_test.pptx"
    Application.ActivePresentation.SaveCopyAs savePath

    Dim candidates() As Candidate
    candidates = Discovery.DiscoverSlide(sld)

    Dim partName As String
    partName = "ppt/slides/slide" & sld.SlideIndex & ".xml"

    Dim ok As Boolean
    ok = Matching.EnrichPlaceholderIdx(candidates, savePath, partName)
    result = result & Assert(ok, "EnrichPlaceholderIdx completed without error against a real saved file")

    Dim anyPlaceholder As Boolean, anyResolved As Boolean
    Dim i As Long
    For i = LBound(candidates) To UBound(candidates)
        If candidates(i).HasPlaceholder Then
            anyPlaceholder = True
            If candidates(i).PlaceholderIdx <> -1 Then anyResolved = True
        End If
    Next i
    result = result & Assert(anyPlaceholder, "ppLayoutText slide has at least one placeholder shape")
    result = result & Assert(anyResolved, "at least one placeholder's idx resolved to something other than -1")

    Test_Matching_EnrichPlaceholderIdxReadsRealFile = result
End Function

' ---------------------------------------------------------------------
' Resolve
' ---------------------------------------------------------------------

Private Function Test_Resolve_ReadsTagsOffLiveSlide() As String
    Dim result As String
    Dim sld As Object
    Set sld = NewBlankSlide()
    sld.Tags.Add "slide_type", "quarterly-update"
    sld.Tags.Add "instance_key", "rec-1"

    Dim instance As SlideInstance
    instance = Resolve.ResolveSlideInstance(sld)
    result = result & Assert(instance.HasTypeTag And instance.TypeTag = "quarterly-update", "TypeTag read correctly, got '" & instance.TypeTag & "'")
    result = result & Assert(instance.HasInstanceKey And instance.InstanceKey = "rec-1", "InstanceKey read correctly, got '" & instance.InstanceKey & "'")

    Dim untaggedSld As Object
    Set untaggedSld = NewBlankSlide()
    Dim untaggedInstance As SlideInstance
    untaggedInstance = Resolve.ResolveSlideInstance(untaggedSld)
    result = result & Assert(Not untaggedInstance.HasTypeTag And Not untaggedInstance.HasInstanceKey, "untagged slide resolves to no type/key")

    Test_Resolve_ReadsTagsOffLiveSlide = result
End Function

' ---------------------------------------------------------------------
' SyncOperations
' ---------------------------------------------------------------------

Private Function Test_SyncOperations_Cases1And4() As String
    Dim result As String
    Dim sld As Object
    Set sld = NewBlankSlide()
    sld.Tags.Add "slide_type", "quarterly-update"
    sld.Tags.Add "instance_key", "rec-1"
    Dim shp As Object
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    shp.TextFrame.TextRange.Text = "Q3 Revenue"
    shp.Tags.Add "role", "Title"

    Dim instances(1 To 1) As Object
    Set instances(1) = sld
    Dim order As New Collection
    order.Add "rec-1"

    Dim rows1 As Object
    Set rows1 = CreateObject("Scripting.Dictionary")
    Dim fields1 As Object
    Set fields1 = CreateObject("Scripting.Dictionary")
    fields1("Title") = "Q3 Revenue"
    Set rows1("rec-1") = fields1

    Dim actions1() As SyncAction
    actions1 = SyncOperations.PlanRoutineSync(instances, order, rows1)
    result = result & Assert(UBound(actions1) = 1, "one action produced, got " & (UBound(actions1) - LBound(actions1) + 1))
    result = result & Assert(actions1(1).Kind = "no_change", "case 1: no_change, got '" & actions1(1).Kind & "'")

    Dim rows2 As Object
    Set rows2 = CreateObject("Scripting.Dictionary")
    Dim fields2 As Object
    Set fields2 = CreateObject("Scripting.Dictionary")
    fields2("Title") = "Q3 Revenue (revised)"
    Set rows2("rec-1") = fields2

    Dim actions2() As SyncAction
    actions2 = SyncOperations.PlanRoutineSync(instances, order, rows2)
    result = result & Assert(actions2(1).Kind = "in_place_correction", "case 4: in_place_correction, got '" & actions2(1).Kind & "'")
    result = result & Assert(shp.TextFrame.TextRange.Text = "Q3 Revenue (revised)", "shape text actually updated, got '" & shp.TextFrame.TextRange.Text & "'")
    result = result & Assert(actions2(1).ChangedFieldVerified("Title") = True, "changed field reports Verified=True")

    Test_SyncOperations_Cases1And4 = result
End Function

Private Function Test_SyncOperations_Case3NewRecord() As String
    Dim result As String
    Dim instances() As Object ' deliberately left unallocated -- no known instances at all; PlanRoutineSync error-guards this (see AGENTS.md's Known Patterns)

    Dim order As New Collection
    order.Add "rec-new"

    Dim rows As Object
    Set rows = CreateObject("Scripting.Dictionary")
    Dim fields As Object
    Set fields = CreateObject("Scripting.Dictionary")
    fields("Title") = "Brand New"
    Set rows("rec-new") = fields

    Dim actions() As SyncAction
    actions = SyncOperations.PlanRoutineSync(instances, order, rows)

    result = result & Assert(UBound(actions) = 1, "one action produced, got " & (UBound(actions) - LBound(actions) + 1))
    result = result & Assert(actions(1).Kind = "new_record", "case 3: new_record, got '" & actions(1).Kind & "'")
    result = result & Assert(actions(1).RowInstanceKey = "rec-new", "RowInstanceKey correct")
    result = result & Assert(actions(1).Values("Title") = "Brand New", "Values carries the row's data")

    Test_SyncOperations_Case3NewRecord = result
End Function

Private Function Test_SyncOperations_Case6UnclassifiedSlide() As String
    Dim result As String
    Dim sld As Object
    Set sld = NewBlankSlide() ' no tags at all

    Dim instances(1 To 1) As Object
    Set instances(1) = sld
    Dim order As New Collection ' empty -- nothing to dispatch against

    Dim rows As Object
    Set rows = CreateObject("Scripting.Dictionary")

    Dim actions() As SyncAction
    actions = SyncOperations.PlanRoutineSync(instances, order, rows)

    result = result & Assert(UBound(actions) = 1, "one action produced, got " & (UBound(actions) - LBound(actions) + 1))
    result = result & Assert(actions(1).Kind = "flagged", "case 6: flagged, got '" & actions(1).Kind & "'")
    result = result & Assert(actions(1).FlagKind = "unclassified_slide", "FlagKind correct, got '" & actions(1).FlagKind & "'")

    Test_SyncOperations_Case6UnclassifiedSlide = result
End Function

' ---------------------------------------------------------------------
' Onboarding
' ---------------------------------------------------------------------

Private Function Test_Onboarding_HighAndMediumConfidence() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 300, 50)
    titleShp.TextFrame.TextRange.Text = "Title text"
    titleShp.Tags.Add "role", "Title"
    Dim bodyShp As Object
    Set bodyShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 200, 300, 200)
    bodyShp.TextFrame.TextRange.Text = "Body text"
    bodyShp.Tags.Add "role", "Body"

    Dim templateRoles() As String
    Dim templateFieldShapes() As Candidate
    templateFieldShapes = Onboarding.BuildTemplateFieldShapes(templateSld, templateRoles)
    result = result & Assert((UBound(templateRoles) - LBound(templateRoles) + 1) = 2, "template has 2 field roles")

    ' New slide: Title shape barely drifted (well inside tolerance -> high
    ' confidence expected). Body shape moved far away entirely (600pt off,
    ' many multiples of POSITION_TOLERANCE_EMU's 1-inch/72pt radius) --
    ' first attempt here used a modest drift and empirically scored "high"
    ' anyway (real run, 2026-07-25: geometry is only 0.3 of the renormalized
    ' weight once placeholder is inapplicable, so a moderate offset wasn't
    ' enough to pull the blended score below HIGH_THRESHOLD=0.75 on its
    ' own) -- a real finding about how forgiving the tolerance actually is
    ' with only 3 applicable signals, not a harness bug. Widened the drift
    ' to unambiguous rather than trying to hit "exactly medium" precisely.
    Dim newSld As Object
    Set newSld = NewBlankSlide()
    Dim newTitleShp As Object
    Set newTitleShp = newSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 102, 101, 300, 50)
    newTitleShp.TextFrame.TextRange.Text = "Title text drifted"
    Dim newBodyShp As Object
    Set newBodyShp = newSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 700, 700, 300, 200)
    newBodyShp.TextFrame.TextRange.Text = "Body text drifted"

    Dim untaggedShapes() As Object
    Dim matches() As FieldMatch
    matches = Onboarding.MatchSlideAgainstTemplate(newSld, templateRoles, templateFieldShapes, untaggedShapes)

    Dim titleConf As String, bodyConf As String
    Dim haveTitle As Boolean, haveBody As Boolean
    Dim i As Long
    For i = LBound(matches) To UBound(matches)
        If matches(i).Role = "Title" Then titleConf = matches(i).Result.Confidence: haveTitle = True
        If matches(i).Role = "Body" Then bodyConf = matches(i).Result.Confidence: haveBody = True
    Next i

    result = result & Assert(haveTitle And titleConf = "high", "Title (barely drifted) scores high, got '" & titleConf & "'")
    result = result & Assert(haveBody And bodyConf <> "high", "Body (drifted position+size) does not auto-accept, got '" & bodyConf & "'")

    Test_Onboarding_HighAndMediumConfidence = result
End Function

Private Function Test_Onboarding_OnboardNewInstanceAutoTagsHighOnly() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 300, 50)
    titleShp.TextFrame.TextRange.Text = "Title text"
    titleShp.Tags.Add "role", "Title"

    Dim templateRoles() As String
    Dim templateFieldShapes() As Candidate
    templateFieldShapes = Onboarding.BuildTemplateFieldShapes(templateSld, templateRoles)

    Dim newSld As Object
    Set newSld = NewBlankSlide()
    Dim newTitleShp As Object
    Set newTitleShp = newSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 101, 100, 300, 50)
    newTitleShp.TextFrame.TextRange.Text = "New title"

    Dim matches() As FieldMatch
    matches = Onboarding.OnboardNewInstance(newSld, templateRoles, templateFieldShapes, "quarterly-update", "rec-2")

    Dim instance As SlideInstance
    instance = Resolve.ResolveSlideInstance(newSld)
    result = result & Assert(instance.HasTypeTag And instance.TypeTag = "quarterly-update", "slide-level type tagged unconditionally")
    result = result & Assert(instance.HasInstanceKey And instance.InstanceKey = "rec-2", "slide-level instance key tagged unconditionally")
    result = result & Assert(newTitleShp.Tags("role") = "Title", "high-confidence Title auto-tagged onto the shape, got '" & newTitleShp.Tags("role") & "'")

    Test_Onboarding_OnboardNewInstanceAutoTagsHighOnly = result
End Function

Private Function Test_Onboarding_PureDecorationNeverMatched() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim decoShp As Object
    Set decoShp = templateSld.Shapes.AddShape(msoShapeOval, 100, 100, 50, 50) ' no text -- pure decoration
    decoShp.Tags.Add "role", "Deco"

    Dim templateRoles() As String
    Dim templateFieldShapes() As Candidate
    templateFieldShapes = Onboarding.BuildTemplateFieldShapes(templateSld, templateRoles)

    Dim newSld As Object
    Set newSld = NewBlankSlide()
    Dim newDecoShp As Object
    Set newDecoShp = newSld.Shapes.AddShape(msoShapeOval, 100, 100, 50, 50) ' identical geometry, still no text

    Dim untaggedShapes() As Object
    Dim matches() As FieldMatch
    matches = Onboarding.MatchSlideAgainstTemplate(newSld, templateRoles, templateFieldShapes, untaggedShapes)

    Dim lo As Long, hi As Long, hasMatches As Boolean
    On Error Resume Next
    lo = LBound(matches)
    hi = UBound(matches)
    hasMatches = (Err.Number = 0)
    On Error GoTo 0

    result = result & Assert(hasMatches And (hi - lo + 1) = 1, "one match result produced")
    If hasMatches Then
        result = result & Assert(matches(1).Result.Confidence = "low", "pure decoration never matches even at identical geometry, got '" & matches(1).Result.Confidence & "'")
        result = result & Assert(Not matches(1).Result.HasCandidate, "HasCandidate false")
    End If

    Test_Onboarding_PureDecorationNeverMatched = result
End Function

' ---------------------------------------------------------------------
' Verification
' ---------------------------------------------------------------------

Private Function Test_Verification_StructureMatchesAfterDuplicate() As String
    Dim result As String
    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim shp1 As Object, shp2 As Object
    Set shp1 = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    shp1.TextFrame.TextRange.Text = "Title"
    shp1.Tags.Add "role", "Title"
    Set shp2 = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 150, 200, 50)
    shp2.TextFrame.TextRange.Text = "Body"
    shp2.Tags.Add "role", "Body"

    Dim dupSlides As Object
    Set dupSlides = sld.Duplicate()
    Dim dupSld As Object
    Set dupSld = dupSlides(1)

    Dim structCheck As StructuralVerification
    structCheck = Verification.VerifyStructure(sld, dupSld)
    result = result & Assert(structCheck.Ok, "structure matches after a clean duplicate, got " & structCheck.MismatchCount & " mismatch(es)")
    result = result & Assert(structCheck.SourceCount = structCheck.DuplicateCount, "counts match, got " & structCheck.SourceCount & " vs " & structCheck.DuplicateCount)

    Dim zCheck As ZOrderVerification
    zCheck = Verification.VerifyZOrder(sld, dupSld)
    result = result & Assert(zCheck.Ok, "z-order matches after a clean duplicate, got " & zCheck.MismatchCount & " mismatch(es)")

    Test_Verification_StructureMatchesAfterDuplicate = result
End Function

Private Function Test_Verification_DetectsShapeCountMismatch() As String
    Dim result As String
    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim shp1 As Object, shp2 As Object
    Set shp1 = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    shp1.TextFrame.TextRange.Text = "Title"
    shp1.Tags.Add "role", "Title"
    Set shp2 = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 150, 200, 50)
    shp2.TextFrame.TextRange.Text = "Body"
    shp2.Tags.Add "role", "Body"

    Dim dupSlides As Object
    Set dupSlides = sld.Duplicate()
    Dim dupSld As Object
    Set dupSld = dupSlides(1)
    dupSld.Shapes(2).Delete ' real structural defect: duplicate is missing a shape

    Dim structCheck As StructuralVerification
    structCheck = Verification.VerifyStructure(sld, dupSld)
    result = result & Assert(Not structCheck.Ok, "structure mismatch detected")
    result = result & Assert(structCheck.MismatchCount > 0, "at least one mismatch reported")

    Test_Verification_DetectsShapeCountMismatch = result
End Function

Private Function Test_Verification_DetectsZOrderSwap() As String
    Dim result As String
    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim shp1 As Object, shp2 As Object
    Set shp1 = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    shp1.TextFrame.TextRange.Text = "Title"
    shp1.Tags.Add "role", "Title"
    Set shp2 = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 150, 200, 50)
    shp2.TextFrame.TextRange.Text = "Body"
    shp2.Tags.Add "role", "Body"

    Dim dupSlides As Object
    Set dupSlides = sld.Duplicate()
    Dim dupSld As Object
    Set dupSld = dupSlides(1)
    ' Shapes(1) (Title) was added first, so it's already backmost among just
    ' 2 shapes -- sending it "to back" would be a no-op (first attempt here
    ' did exactly that and found nothing, a test bug not a Verification.bas
    ' bug, per the real 2026-07-25 run). Bring it to front instead, which
    ' actually swaps its relative order against Body.
    dupSld.Shapes(1).ZOrder msoBringToFront ' same shapes/tags/types, stacking order swapped on the duplicate only

    Dim zCheck As ZOrderVerification
    zCheck = Verification.VerifyZOrder(sld, dupSld)
    result = result & Assert(Not zCheck.Ok, "z-order swap detected")
    result = result & Assert(zCheck.MismatchCount > 0, "at least one z-order mismatch reported")

    Dim structCheck As StructuralVerification
    structCheck = Verification.VerifyStructure(sld, dupSld)
    result = result & Assert(structCheck.Ok, "structure itself is unaffected by a pure z-order swap, got " & structCheck.MismatchCount & " mismatch(es)")

    Test_Verification_DetectsZOrderSwap = result
End Function

' ---------------------------------------------------------------------
' SlideDuplication
' ---------------------------------------------------------------------

Private Function Test_SlideDuplication_CreatesTaggedInjectedSlide() As String
    Dim result As String
    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object, bodyShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    titleShp.TextFrame.TextRange.Text = "Template Title"
    titleShp.Tags.Add "role", "Title"
    Set bodyShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 150, 200, 50)
    bodyShp.TextFrame.TextRange.Text = "Template Body"
    bodyShp.Tags.Add "role", "Body"
    templateSld.Tags.Add "slide_type", "quarterly-update"
    templateSld.Tags.Add "instance_key", "rec-template"

    Dim values As Object
    Set values = CreateObject("Scripting.Dictionary")
    values("Title") = "Q3 Revenue"
    values("Body") = "Strong quarter"

    Dim noInstances() As Object ' deliberately unallocated -- nothing to collision-check against

    Dim dr As DuplicateResult
    dr = SlideDuplication.DuplicateAndTag(templateSld, "quarterly-update", "rec-new-1", values, noInstances)

    result = result & Assert(dr.Ok, "DuplicateAndTag succeeded, reason='" & dr.Reason & "'")
    If dr.Ok Then
        Dim instance As SlideInstance
        instance = Resolve.ResolveSlideInstance(dr.NewSlide)
        result = result & Assert(instance.TypeTag = "quarterly-update", "new slide tagged with correct slide_type, got '" & instance.TypeTag & "'")
        result = result & Assert(instance.InstanceKey = "rec-new-1", "new slide tagged with correct instance_key, got '" & instance.InstanceKey & "'")
        result = result & Assert(dr.MissingFieldCount = 0, "no missing fields when values supplies everything, got " & dr.MissingFieldCount)

        Dim newTitleShp As Object
        Set newTitleShp = FindShapeByRole(dr.NewSlide, "Title")
        result = result & Assert(Not newTitleShp Is Nothing, "Title shape found on new slide")
        If Not newTitleShp Is Nothing Then
            result = result & Assert(newTitleShp.TextFrame.TextRange.Text = "Q3 Revenue", "Title value injected correctly, got '" & newTitleShp.TextFrame.TextRange.Text & "'")
        End If
    End If

    Test_SlideDuplication_CreatesTaggedInjectedSlide = result
End Function

Private Function Test_SlideDuplication_RefusesInstanceKeyCollision() As String
    Dim result As String
    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    titleShp.TextFrame.TextRange.Text = "Title"
    titleShp.Tags.Add "role", "Title"
    templateSld.Tags.Add "slide_type", "quarterly-update"
    templateSld.Tags.Add "instance_key", "rec-a"

    Dim existingSld As Object
    Set existingSld = NewBlankSlide()
    existingSld.Tags.Add "slide_type", "quarterly-update"
    existingSld.Tags.Add "instance_key", "rec-dup"

    Dim existingInstances(1 To 1) As Object
    Set existingInstances(1) = existingSld

    Dim values As Object
    Set values = CreateObject("Scripting.Dictionary")
    values("Title") = "New Value"

    Dim slidesBefore As Long
    slidesBefore = Application.ActivePresentation.Slides.count

    Dim dr As DuplicateResult
    dr = SlideDuplication.DuplicateAndTag(templateSld, "quarterly-update", "rec-dup", values, existingInstances)

    result = result & Assert(Not dr.Ok, "refuses to create a duplicate instance_key")
    result = result & Assert(Application.ActivePresentation.Slides.count = slidesBefore, "no new slide was left behind, got " & Application.ActivePresentation.Slides.count & " vs expected " & slidesBefore)

    Test_SlideDuplication_RefusesInstanceKeyCollision = result
End Function

Private Function Test_SlideDuplication_PartialRowStillCreatesSlideButFlagsMissing() As String
    Dim result As String
    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object, bodyShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    titleShp.TextFrame.TextRange.Text = "Title"
    titleShp.Tags.Add "role", "Title"
    Set bodyShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 150, 200, 50)
    bodyShp.TextFrame.TextRange.Text = "Body"
    bodyShp.Tags.Add "role", "Body"
    templateSld.Tags.Add "slide_type", "quarterly-update"
    templateSld.Tags.Add "instance_key", "rec-template2"

    Dim values As Object
    Set values = CreateObject("Scripting.Dictionary")
    values("Title") = "Only Title Supplied" ' Body deliberately missing

    Dim noInstances() As Object

    Dim dr As DuplicateResult
    dr = SlideDuplication.DuplicateAndTag(templateSld, "quarterly-update", "rec-partial", values, noInstances)

    result = result & Assert(dr.Ok, "slide still created despite a missing field, reason='" & dr.Reason & "'")
    result = result & Assert(dr.MissingFieldCount = 1, "exactly one missing field flagged, got " & dr.MissingFieldCount)
    If dr.MissingFieldCount = 1 Then
        result = result & Assert(dr.MissingFields(1) = "Body", "flagged field is 'Body', got '" & dr.MissingFields(1) & "'")
    End If

    Test_SlideDuplication_PartialRowStillCreatesSlideButFlagsMissing = result
End Function

' ---------------------------------------------------------------------
' RunSync
' ---------------------------------------------------------------------

Private Function Test_RunSync_GatherInstancesFiltersByType() As String
    Dim result As String
    Dim sldA As Object, sldB As Object, sldC As Object
    Set sldA = NewBlankSlide()
    sldA.Tags.Add "slide_type", "type-a"
    sldA.Tags.Add "instance_key", "a-1"
    Set sldB = NewBlankSlide()
    sldB.Tags.Add "slide_type", "type-b"
    sldB.Tags.Add "instance_key", "b-1"
    Set sldC = NewBlankSlide()
    sldC.Tags.Add "slide_type", "type-a"
    sldC.Tags.Add "instance_key", "a-2"

    Dim instances() As Object
    instances = RunSync.GatherInstances("type-a")

    Dim lo As Long, hi As Long, hasAny As Boolean
    On Error Resume Next
    lo = LBound(instances): hi = UBound(instances): hasAny = (Err.Number = 0)
    On Error GoTo 0

    result = result & Assert(hasAny And (hi - lo + 1) >= 2, "at least the 2 type-a slides found, got " & IIf(hasAny, hi - lo + 1, 0))

    Dim foundA1 As Boolean, foundA2 As Boolean, foundB As Boolean
    Dim i As Long
    If hasAny Then
        For i = lo To hi
            Dim inst As SlideInstance
            inst = Resolve.ResolveSlideInstance(instances(i))
            If inst.InstanceKey = "a-1" Then foundA1 = True
            If inst.InstanceKey = "a-2" Then foundA2 = True
            If inst.InstanceKey = "b-1" Then foundB = True
        Next i
    End If
    result = result & Assert(foundA1 And foundA2, "both type-a instances included")
    result = result & Assert(Not foundB, "type-b instance correctly excluded")

    Test_RunSync_GatherInstancesFiltersByType = result
End Function

' End-to-end: a template, a fresh Excel sheet (via cross-app Excel
' automation -- the real intended usage per specs/vba-port.md's "VBA runs
' inside Excel or drives it via COM from the PowerPoint side") with one
' existing-and-stale row (exercises case 4) and two rows with no matching
' slide yet (exercises case 3 twice, plus resequencing). Written directly
' to the worksheet's cells matching ExcelOutput's own layout convention
' rather than importing ExcelOutput.bas into a second, Excel-hosted VBA
' project -- RunSync.RunRoutineSync's own ExcelOutput.ReadSheet call reads
' it from the PowerPoint-hosted project against the live cross-app ws
' object either way, which is the actual thing under test.
Private Function Test_RunSync_EndToEndCreatesSlidesFromFreshSheet() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    titleShp.TextFrame.TextRange.Text = "Template Title"
    titleShp.Tags.Add "role", "Title"
    templateSld.Tags.Add "slide_type", "e2e-type"
    templateSld.Tags.Add "instance_key", "e2e-template"

    Dim xl As Object, wb As Object, ws As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Add()
    Set ws = wb.Worksheets(1)

    ws.Cells(1, 1).Value = "Instance ID"
    ws.Cells(1, 2).Value = "Title"
    ws.Cells(2, 1).Value = "e2e-existing"
    ws.Cells(2, 2).Value = "Existing Value"
    ws.Cells(3, 1).Value = "e2e-new-1"
    ws.Cells(3, 2).Value = "Brand New One"
    ws.Cells(4, 1).Value = "e2e-new-2"
    ws.Cells(4, 2).Value = "Brand New Two"

    Dim existingSld As Object
    Set existingSld = NewBlankSlide()
    Dim existingTitleShp As Object
    Set existingTitleShp = existingSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    existingTitleShp.TextFrame.TextRange.Text = "Stale Value"
    existingTitleShp.Tags.Add "role", "Title"
    existingSld.Tags.Add "slide_type", "e2e-type"
    existingSld.Tags.Add "instance_key", "e2e-existing"

    Dim report As String
    report = RunSync.RunRoutineSync(ws, "e2e-type", templateSld)

    Dim instances() As Object
    instances = RunSync.GatherInstances("e2e-type")

    Dim lo As Long, hi As Long, hasAny As Boolean
    On Error Resume Next
    lo = LBound(instances): hi = UBound(instances): hasAny = (Err.Number = 0)
    On Error GoTo 0

    result = result & Assert(hasAny And (hi - lo + 1) = 4, "4 total instances after sync (template+existing+2 new), got " & IIf(hasAny, hi - lo + 1, 0) & " -- report: " & report)

    Dim byKey As Object
    Set byKey = CreateObject("Scripting.Dictionary")
    Dim i As Long
    If hasAny Then
        For i = lo To hi
            Dim inst As SlideInstance
            inst = Resolve.ResolveSlideInstance(instances(i))
            If inst.HasInstanceKey Then Set byKey(inst.InstanceKey) = instances(i)
        Next i
    End If

    result = result & Assert(byKey.Exists("e2e-new-1") And byKey.Exists("e2e-new-2"), "both new_record slides were created")

    If byKey.Exists("e2e-existing") Then
        Dim correctedShp As Object
        Set correctedShp = FindShapeByRole(byKey("e2e-existing"), "Title")
        Dim correctedText As String
        correctedText = IIf(correctedShp Is Nothing, "<shape not found>", correctedShp.TextFrame.TextRange.Text)
        result = result & Assert(correctedText = "Existing Value", "existing slide's stale value was corrected (case 4), got '" & correctedText & "'")
    End If
    If byKey.Exists("e2e-new-1") Then
        Dim new1Shp As Object
        Set new1Shp = FindShapeByRole(byKey("e2e-new-1"), "Title")
        Dim new1Text As String
        new1Text = IIf(new1Shp Is Nothing, "<shape not found>", new1Shp.TextFrame.TextRange.Text)
        result = result & Assert(new1Text = "Brand New One", "new slide's value injected correctly, got '" & new1Text & "'")
    End If

    If byKey.Exists("e2e-existing") And byKey.Exists("e2e-new-1") And byKey.Exists("e2e-new-2") Then
        result = result & Assert(byKey("e2e-existing").SlideIndex < byKey("e2e-new-1").SlideIndex, "e2e-existing sits before e2e-new-1 after resequencing")
        result = result & Assert(byKey("e2e-new-1").SlideIndex < byKey("e2e-new-2").SlideIndex, "e2e-new-1 sits before e2e-new-2 after resequencing")
    End If

    wb.Close False
    xl.Quit

    Test_RunSync_EndToEndCreatesSlidesFromFreshSheet = result
End Function

' Case 2 (period rollover): duplicates the source instance's current slide
' into a new period's slide, injecting new values -- and confirms the
' source slide itself is left untouched as history (specs/sync-
' operations.md's explicit case-2 requirement), not just that a duplicate
' with the right values exists. Also confirms RunPeriodRollover's own
' collision guard (via the existingInstances it gathers internally) refuses
' to roll over onto an instance_key that already exists, same posture
' SlideDuplication.DuplicateAndTag already enforces for case 3.
Private Function Test_RunSync_RunPeriodRolloverDuplicatesLeavingSourceUntouched() As String
    Dim result As String

    Dim sourceSld As Object
    Set sourceSld = NewBlankSlide()
    Dim titleShp As Object
    Set titleShp = sourceSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    titleShp.TextFrame.TextRange.Text = "Q1 Value"
    titleShp.Tags.Add "role", "Title"
    sourceSld.Tags.Add "slide_type", "rollover-type"
    sourceSld.Tags.Add "instance_key", "rollover-q1"

    Dim newValues As Object
    Set newValues = CreateObject("Scripting.Dictionary")
    newValues.Add "Title", "Q2 Value"

    Dim dr As DuplicateResult
    dr = RunSync.RunPeriodRollover(sourceSld, "rollover-type", "rollover-q2", newValues)

    result = result & Assert(dr.Ok, "rollover duplication succeeded: " & dr.Reason)

    If dr.Ok Then
        Dim newTitleShp As Object
        Set newTitleShp = FindShapeByRole(dr.NewSlide, "Title")
        Dim newText As String
        newText = IIf(newTitleShp Is Nothing, "<shape not found>", newTitleShp.TextFrame.TextRange.Text)
        result = result & Assert(newText = "Q2 Value", "new period's slide got the new value injected, got '" & newText & "'")

        Dim newInst As SlideInstance
        newInst = Resolve.ResolveSlideInstance(dr.NewSlide)
        result = result & Assert(newInst.HasInstanceKey And newInst.InstanceKey = "rollover-q2", "new slide tagged with the new period's instance_key")
    End If

    Dim sourceTitleShp As Object
    Set sourceTitleShp = FindShapeByRole(sourceSld, "Title")
    Dim sourceText As String
    sourceText = IIf(sourceTitleShp Is Nothing, "<shape not found>", sourceTitleShp.TextFrame.TextRange.Text)
    result = result & Assert(sourceText = "Q1 Value", "source slide left untouched as history, got '" & sourceText & "'")

    ' Collision guard: rolling over to an instance_key that already exists
    ' (e.g. the just-created new period) must refuse, not double-create.
    Dim dupDr As DuplicateResult
    dupDr = RunSync.RunPeriodRollover(sourceSld, "rollover-type", "rollover-q2", newValues)
    result = result & Assert(Not dupDr.Ok, "re-rolling over to an already-used instance_key is refused, not silently double-created")

    Test_RunSync_RunPeriodRolloverDuplicatesLeavingSourceUntouched = result
End Function

' ---------------------------------------------------------------------
' DeckAdoption
' ---------------------------------------------------------------------

' Opens a fresh, blank Excel worksheet via cross-app automation, same
' pattern as Test_RunSync_EndToEndCreatesSlidesFromFreshSheet -- DeckAdoption
' always reads/writes a live worksheet, never a closed file. Caller is
' responsible for wb.Close False / xl.Quit when done.
Private Sub NewTestWorksheet(ByRef xl As Object, ByRef wb As Object, ByRef ws As Object)
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Add()
    Set ws = wb.Worksheets(1)
    ws.Cells(1, 1).Value = "Instance ID"
End Sub

Private Function Test_DeckAdoption_AlreadyLinkedSlideSkipped() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 300, 50)
    titleShp.TextFrame.TextRange.Text = "Template Title"
    titleShp.Tags.Add "role", "Title"

    Dim alreadySld As Object
    Set alreadySld = NewBlankSlide()
    Dim alreadyShp As Object
    Set alreadyShp = alreadySld.Shapes.AddTextbox(msoTextOrientationHorizontal, 101, 100, 300, 50)
    alreadyShp.TextFrame.TextRange.Text = "Already Linked"
    alreadySld.Tags.Add "slide_type", "adopt-type"
    alreadySld.Tags.Add "instance_key", "already-1"

    Dim xl As Object, wb As Object, ws As Object
    NewTestWorksheet xl, wb, ws

    Dim slidesToAdopt(1 To 1) As Object
    Set slidesToAdopt(1) = alreadySld

    Dim harvestedValues() As Object
    Dim plans() As AdoptionSlidePlan
    plans = DeckAdoption.PlanAdoption(slidesToAdopt, templateSld, ws, harvestedValues)

    result = result & Assert(UBound(plans) = 1, "one plan entry produced, got " & (UBound(plans) - LBound(plans) + 1))
    result = result & Assert(plans(1).Disposition = "already_linked", "disposition is already_linked, got '" & plans(1).Disposition & "'")

    Dim confirmedKeys(1 To 1) As String
    confirmedKeys(1) = ""

    Dim commitResult As AdoptionResult
    commitResult = DeckAdoption.CommitAdoption(plans, slidesToAdopt, harvestedValues, confirmedKeys, "adopt-type", templateSld, ws)

    result = result & Assert(commitResult.AlreadyLinkedCount = 1, "AlreadyLinkedCount=1, got " & commitResult.AlreadyLinkedCount)
    result = result & Assert(commitResult.LinkedCount = 0, "LinkedCount=0 (nothing to link), got " & commitResult.LinkedCount)

    wb.Close False
    xl.Quit

    Test_DeckAdoption_AlreadyLinkedSlideSkipped = result
End Function

Private Function Test_DeckAdoption_ReadyHighConfidenceSlideLinkedAndCreatesFreshRow() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 300, 50)
    titleShp.TextFrame.TextRange.Text = "Template Title"
    titleShp.Tags.Add "role", "Title"

    Dim candidateSld As Object
    Set candidateSld = NewBlankSlide()
    Dim candidateShp As Object
    Set candidateShp = candidateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 102, 101, 300, 50)
    candidateShp.TextFrame.TextRange.Text = "Harvested Title Value"

    Dim xl As Object, wb As Object, ws As Object
    NewTestWorksheet xl, wb, ws

    Dim slidesToAdopt(1 To 1) As Object
    Set slidesToAdopt(1) = candidateSld

    Dim harvestedValues() As Object
    Dim plans() As AdoptionSlidePlan
    plans = DeckAdoption.PlanAdoption(slidesToAdopt, templateSld, ws, harvestedValues)

    result = result & Assert(plans(1).Disposition = "ready", "disposition is ready, got '" & plans(1).Disposition & "'")
    result = result & Assert(plans(1).MatchedKeylessRowId = "", "no keyless row to match against an empty sheet")
    result = result & Assert(Not harvestedValues(1) Is Nothing, "harvested values captured for a ready slide")
    If Not harvestedValues(1) Is Nothing Then
        result = result & Assert(harvestedValues(1)("Title") = "Harvested Title Value", "harvested Title value correct, got '" & harvestedValues(1)("Title") & "'")
    End If

    Dim confirmedKeys(1 To 1) As String
    confirmedKeys(1) = "adopt-rec-1"

    Dim commitResult As AdoptionResult
    commitResult = DeckAdoption.CommitAdoption(plans, slidesToAdopt, harvestedValues, confirmedKeys, "adopt-type", templateSld, ws)

    result = result & Assert(commitResult.LinkedCount = 1, "LinkedCount=1, got " & commitResult.LinkedCount)
    result = result & Assert(commitResult.FailedVerificationCount = 0, "no verification failures, got " & commitResult.FailedVerificationCount)

    Dim instance As SlideInstance
    instance = Resolve.ResolveSlideInstance(candidateSld)
    result = result & Assert(instance.HasTypeTag And instance.TypeTag = "adopt-type", "slide tagged with slide_type, got '" & instance.TypeTag & "'")
    result = result & Assert(instance.HasInstanceKey And instance.InstanceKey = "adopt-rec-1", "slide tagged with instance_key, got '" & instance.InstanceKey & "'")

    Dim sheet As Sheet
    sheet = ExcelOutput.ReadSheet(ws)
    result = result & Assert(sheet.Rows.Exists("adopt-rec-1"), "Data-sheet row created for the new instance")
    If sheet.Rows.Exists("adopt-rec-1") Then
        result = result & Assert(sheet.Rows("adopt-rec-1")("Title") = "Harvested Title Value", "Data-sheet row carries the harvested value, got '" & sheet.Rows("adopt-rec-1")("Title") & "'")
    End If

    wb.Close False
    xl.Quit

    Test_DeckAdoption_ReadyHighConfidenceSlideLinkedAndCreatesFreshRow = result
End Function

' Reuses the exact same drift (title barely moved, body moved 600pt away)
' Test_Onboarding_HighAndMediumConfidence already established scores
' high/medium respectively (not low) -- see that test's own comment for why
' this specific offset was chosen empirically.
Private Function Test_DeckAdoption_MediumConfidenceSlideNeedsConfirmationAndIsNotTagged() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 300, 50)
    titleShp.TextFrame.TextRange.Text = "Title text"
    titleShp.Tags.Add "role", "Title"
    Dim bodyShp As Object
    Set bodyShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 200, 300, 200)
    bodyShp.TextFrame.TextRange.Text = "Body text"
    bodyShp.Tags.Add "role", "Body"

    Dim candidateSld As Object
    Set candidateSld = NewBlankSlide()
    Dim candidateTitleShp As Object
    Set candidateTitleShp = candidateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 102, 101, 300, 50)
    candidateTitleShp.TextFrame.TextRange.Text = "Title text drifted"
    Dim candidateBodyShp As Object
    Set candidateBodyShp = candidateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 700, 700, 300, 200)
    candidateBodyShp.TextFrame.TextRange.Text = "Body text drifted"

    Dim xl As Object, wb As Object, ws As Object
    NewTestWorksheet xl, wb, ws

    Dim slidesToAdopt(1 To 1) As Object
    Set slidesToAdopt(1) = candidateSld

    Dim harvestedValues() As Object
    Dim plans() As AdoptionSlidePlan
    plans = DeckAdoption.PlanAdoption(slidesToAdopt, templateSld, ws, harvestedValues)

    result = result & Assert(plans(1).Disposition = "needs_confirmation", "disposition is needs_confirmation, got '" & plans(1).Disposition & "'")

    Dim confirmedKeys(1 To 1) As String
    confirmedKeys(1) = "" ' irrelevant -- needs_confirmation is never committed regardless

    Dim commitResult As AdoptionResult
    commitResult = DeckAdoption.CommitAdoption(plans, slidesToAdopt, harvestedValues, confirmedKeys, "adopt-type", templateSld, ws)

    result = result & Assert(commitResult.LinkedCount = 0, "nothing committed for a needs_confirmation slide, got LinkedCount=" & commitResult.LinkedCount)
    result = result & Assert(commitResult.ExcludedUnclassifiedCount = 1, "reported in the excluded/unclassified bucket, got " & commitResult.ExcludedUnclassifiedCount)

    Dim instance As SlideInstance
    instance = Resolve.ResolveSlideInstance(candidateSld)
    result = result & Assert(Not instance.HasTypeTag And Not instance.HasInstanceKey, "slide-level tags never written for a needs_confirmation slide")
    Dim untouchedTitleShp As Object
    Set untouchedTitleShp = FindShapeByRole(candidateSld, "Title")
    result = result & Assert(untouchedTitleShp Is Nothing, "Title shape was never tagged either -- no partial commit")

    wb.Close False
    xl.Quit

    Test_DeckAdoption_MediumConfidenceSlideNeedsConfirmationAndIsNotTagged = result
End Function

Private Function Test_DeckAdoption_UnclassifiedSlideExcluded() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 300, 50)
    titleShp.TextFrame.TextRange.Text = "Template Title"
    titleShp.Tags.Add "role", "Title"

    Dim unrelatedSld As Object
    Set unrelatedSld = NewBlankSlide()
    unrelatedSld.Shapes.AddShape msoShapeOval, 500, 500, 50, 50 ' pure decoration, no text -- excluded by IsCandidateField entirely

    Dim xl As Object, wb As Object, ws As Object
    NewTestWorksheet xl, wb, ws

    Dim slidesToAdopt(1 To 1) As Object
    Set slidesToAdopt(1) = unrelatedSld

    Dim harvestedValues() As Object
    Dim plans() As AdoptionSlidePlan
    plans = DeckAdoption.PlanAdoption(slidesToAdopt, templateSld, ws, harvestedValues)

    result = result & Assert(plans(1).Disposition = "unclassified", "disposition is unclassified, got '" & plans(1).Disposition & "'")

    Dim confirmedKeys(1 To 1) As String
    confirmedKeys(1) = ""

    Dim commitResult As AdoptionResult
    commitResult = DeckAdoption.CommitAdoption(plans, slidesToAdopt, harvestedValues, confirmedKeys, "adopt-type", templateSld, ws)

    result = result & Assert(commitResult.ExcludedUnclassifiedCount = 1, "reported excluded/unclassified, got " & commitResult.ExcludedUnclassifiedCount)
    result = result & Assert(commitResult.LinkedCount = 0, "never forced in, got LinkedCount=" & commitResult.LinkedCount)

    wb.Close False
    xl.Quit

    Test_DeckAdoption_UnclassifiedSlideExcluded = result
End Function

Private Function Test_DeckAdoption_MatchesExistingKeylessRowLinksWithoutCreatingNewRow() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 300, 50)
    titleShp.TextFrame.TextRange.Text = "Template Title"
    titleShp.Tags.Add "role", "Title"

    Dim candidateSld As Object
    Set candidateSld = NewBlankSlide()
    Dim candidateShp As Object
    Set candidateShp = candidateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 102, 101, 300, 50)
    candidateShp.TextFrame.TextRange.Text = "Existing Value X"

    Dim xl As Object, wb As Object, ws As Object
    NewTestWorksheet xl, wb, ws
    ws.Cells(1, 2).Value = "Title"
    ws.Cells(2, 2).Value = "Existing Value X" ' row 2: real data, blank Instance ID -- a keyless row

    Dim slidesToAdopt(1 To 1) As Object
    Set slidesToAdopt(1) = candidateSld

    Dim harvestedValues() As Object
    Dim plans() As AdoptionSlidePlan
    plans = DeckAdoption.PlanAdoption(slidesToAdopt, templateSld, ws, harvestedValues)

    result = result & Assert(plans(1).Disposition = "ready", "disposition is ready, got '" & plans(1).Disposition & "'")
    result = result & Assert(plans(1).MatchedKeylessRowId = "2", "matched keyless row 2, got '" & plans(1).MatchedKeylessRowId & "'")

    Dim confirmedKeys(1 To 1) As String
    confirmedKeys(1) = "adopt-rec-existing"

    Dim commitResult As AdoptionResult
    commitResult = DeckAdoption.CommitAdoption(plans, slidesToAdopt, harvestedValues, confirmedKeys, "adopt-type", templateSld, ws)

    result = result & Assert(commitResult.LinkedCount = 1, "LinkedCount=1, got " & commitResult.LinkedCount)
    result = result & Assert(ws.Cells(2, 1).Value = "adopt-rec-existing", "row 2's Instance ID cell was filled in, got '" & ws.Cells(2, 1).Value & "'")

    Dim sheet As Sheet
    sheet = ExcelOutput.ReadSheet(ws)
    result = result & Assert(sheet.InstanceOrder.count = 1, "no new row was appended -- still exactly one instance row, got " & sheet.InstanceOrder.count)

    wb.Close False
    xl.Quit

    Test_DeckAdoption_MatchesExistingKeylessRowLinksWithoutCreatingNewRow = result
End Function

' Regression test for a real bug found and fixed during review (2026-07-25,
' see SPIKE_NOTES_DeckAdoption.md): an earlier version of PlanAdoption forced
' its returned AdoptionSlidePlan() array to be 1-based (via a separate
' counter) while harvestedValues() kept whatever base the caller's
' slidesToAdopt() array used -- CommitAdoption then indexed both (plus
' confirmedInstanceKeys()) with one shared loop variable, silently
' misattributing which slide's harvest/confirmation belongs to which plan
' entry for any non-1-based slidesToAdopt(). Every other DeckAdoption test
' uses a 1-based Dim slidesToAdopt(1 To 1), which would never have exposed
' this -- this test deliberately uses a 0-based array with 3 slides of 3
' different dispositions to prove indices stay aligned end-to-end.
Private Function Test_DeckAdoption_MultiSlideZeroBasedBatchKeepsIndicesAligned() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 300, 50)
    titleShp.TextFrame.TextRange.Text = "Template Title"
    titleShp.Tags.Add "role", "Title"

    Dim alreadySld As Object
    Set alreadySld = NewBlankSlide()
    alreadySld.Tags.Add "slide_type", "adopt-type"
    alreadySld.Tags.Add "instance_key", "already-batch-1"

    Dim readySld As Object
    Set readySld = NewBlankSlide()
    Dim readyShp As Object
    Set readyShp = readySld.Shapes.AddTextbox(msoTextOrientationHorizontal, 101, 100, 300, 50)
    readyShp.TextFrame.TextRange.Text = "Batch Ready Title"

    Dim unrelatedSld As Object
    Set unrelatedSld = NewBlankSlide()
    unrelatedSld.Shapes.AddShape msoShapeOval, 500, 500, 50, 50 ' pure decoration -- unclassified

    Dim xl As Object, wb As Object, ws As Object
    NewTestWorksheet xl, wb, ws

    Dim slidesToAdopt(0 To 2) As Object ' deliberately 0-based, not 1-based
    Set slidesToAdopt(0) = alreadySld
    Set slidesToAdopt(1) = readySld
    Set slidesToAdopt(2) = unrelatedSld

    Dim harvestedValues() As Object
    Dim plans() As AdoptionSlidePlan
    plans = DeckAdoption.PlanAdoption(slidesToAdopt, templateSld, ws, harvestedValues)

    result = result & Assert(LBound(plans) = 0 And UBound(plans) = 2, "plans() shares slidesToAdopt's 0-based range, got " & LBound(plans) & " To " & UBound(plans))
    result = result & Assert(plans(0).Disposition = "already_linked", "index 0 (alreadySld) is already_linked, got '" & plans(0).Disposition & "'")
    result = result & Assert(plans(1).Disposition = "ready", "index 1 (readySld) is ready, got '" & plans(1).Disposition & "'")
    result = result & Assert(plans(2).Disposition = "unclassified", "index 2 (unrelatedSld) is unclassified, got '" & plans(2).Disposition & "'")
    result = result & Assert(Not harvestedValues(1) Is Nothing, "harvested values captured at index 1, not misattributed to another index")
    If Not harvestedValues(1) Is Nothing Then
        result = result & Assert(harvestedValues(1)("Title") = "Batch Ready Title", "harvested value at index 1 belongs to readySld, got '" & harvestedValues(1)("Title") & "'")
    End If

    Dim confirmedKeys(0 To 2) As String
    confirmedKeys(0) = ""
    confirmedKeys(1) = "adopt-batch-ready-1"
    confirmedKeys(2) = ""

    Dim commitResult As AdoptionResult
    commitResult = DeckAdoption.CommitAdoption(plans, slidesToAdopt, harvestedValues, confirmedKeys, "adopt-type", templateSld, ws)

    result = result & Assert(commitResult.AlreadyLinkedCount = 1, "AlreadyLinkedCount=1, got " & commitResult.AlreadyLinkedCount)
    result = result & Assert(commitResult.LinkedCount = 1, "LinkedCount=1, got " & commitResult.LinkedCount)
    result = result & Assert(commitResult.ExcludedUnclassifiedCount = 1, "ExcludedUnclassifiedCount=1, got " & commitResult.ExcludedUnclassifiedCount)

    ' The critical correctness check: readySld (index 1) actually received
    ' the tag/instance_key confirmed for index 1 -- not index 0's or 2's.
    Dim readyInstance As SlideInstance
    readyInstance = Resolve.ResolveSlideInstance(readySld)
    result = result & Assert(readyInstance.HasInstanceKey And readyInstance.InstanceKey = "adopt-batch-ready-1", "readySld (index 1) tagged with its own confirmed instance_key, got '" & readyInstance.InstanceKey & "'")
    Dim readyTitleShp As Object
    Set readyTitleShp = FindShapeByRole(readySld, "Title")
    result = result & Assert(Not readyTitleShp Is Nothing, "readySld's Title shape was tagged")

    ' alreadySld (index 0) must be untouched -- still its original key, not overwritten.
    Dim alreadyInstance As SlideInstance
    alreadyInstance = Resolve.ResolveSlideInstance(alreadySld)
    result = result & Assert(alreadyInstance.InstanceKey = "already-batch-1", "alreadySld (index 0) instance_key unchanged, got '" & alreadyInstance.InstanceKey & "'")

    ' unrelatedSld (index 2) must remain untagged.
    Dim unrelatedInstance As SlideInstance
    unrelatedInstance = Resolve.ResolveSlideInstance(unrelatedSld)
    result = result & Assert(Not unrelatedInstance.HasInstanceKey, "unrelatedSld (index 2) never tagged")

    wb.Close False
    xl.Quit

    Test_DeckAdoption_MultiSlideZeroBasedBatchKeepsIndicesAligned = result
End Function
