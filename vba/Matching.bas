Attribute VB_Name = "Matching"
Option Explicit

' VBA port of src/matching.py's score_candidate()/match(), per specs/vba-
' port.md's port order (module 3, after discovery and identity_tags).
' Operates purely on Discovery.Candidate data -- like its Python
' counterpart, this module doesn't care how a Candidate was produced or how
' IdentityTag got populated (that's discovery/identity_tags' job upstream).
' See SPIKE_NOTES_Matching.md for the full divergence list and manual
' verification recipe -- in particular, this module is what finally
' resolves the PlaceholderIdx gap Discovery.bas flagged and left at -1,
' via a raw-OOXML fallback (ResolvePlaceholderIdx/EnrichPlaceholderIdx)
' since PowerPoint's object model has no native placeholder-idx property.

' Signal weights, in specs/matching.md's reliability order -- identical
' values to matching.py's module-level constants.
Private Const PLACEHOLDER_WEIGHT As Double = 0.5
Private Const GEOMETRY_WEIGHT As Double = 0.3
Private Const SHAPE_TYPE_WEIGHT As Double = 0.15
Private Const CONTENT_WEIGHT As Double = 0.05

Private Const POSITION_TOLERANCE_EMU As Double = 914400
Private Const SIZE_TOLERANCE_EMU As Double = 914400

Private Const HIGH_THRESHOLD As Double = 0.75
Private Const MEDIUM_THRESHOLD As Double = 0.4
Private Const SIBLING_GAP_THRESHOLD As Double = 0.1

Public Type MatchResult
    HasCandidate As Boolean  ' False unless a match was accepted (mirrors Python's candidate=None)
    CandidateIndex As Long   ' index into the candidates() array passed to Match; valid only if HasCandidate
    Confidence As String     ' "high" | "medium" | "low"
    HasScore As Boolean      ' False only for a tier-1 tag-trust match (mirrors Python's score=None)
    Score As Double
    Reason As String
End Type

' ---------------------------------------------------------------------
' Scoring (specs/matching.md's tier-2 fallback)
' ---------------------------------------------------------------------

' Combine every applicable scored signal into one score in [0, 1], in
' specs/matching.md's reliability order. A signal that isn't applicable is
' excluded and the remaining weights are renormalized -- never padded with
' a fabricated value. In this VBA port, only the placeholder-index signal
' can be inapplicable (see PlaceholderScore); geometry/shape-type/content
' are always applicable here because the object model always reports
' Left/Top/Width/Height/Type/text for any shape, unlike raw OOXML where an
' <a:xfrm> or <p:ph> can simply be absent. See SPIKE_NOTES_Matching.md.
Public Function ScoreCandidate(candidate As Candidate, reference As Candidate) As Double
    Dim weightedSum As Double, totalWeight As Double
    Dim applicable As Boolean, s As Double

    s = PlaceholderScore(candidate, reference, applicable)
    If applicable Then
        weightedSum = weightedSum + PLACEHOLDER_WEIGHT * s
        totalWeight = totalWeight + PLACEHOLDER_WEIGHT
    End If

    weightedSum = weightedSum + GEOMETRY_WEIGHT * GeometryScore(candidate, reference)
    totalWeight = totalWeight + GEOMETRY_WEIGHT

    weightedSum = weightedSum + SHAPE_TYPE_WEIGHT * ShapeTypeScore(candidate, reference)
    totalWeight = totalWeight + SHAPE_TYPE_WEIGHT

    weightedSum = weightedSum + CONTENT_WEIGHT * ContentPatternScore(candidate, reference)
    totalWeight = totalWeight + CONTENT_WEIGHT

    If totalWeight = 0 Then
        ScoreCandidate = 0
    Else
        ScoreCandidate = weightedSum / totalWeight
    End If
End Function

' applicable=False when `reference` isn't a placeholder, OR when it is one
' but its PlaceholderIdx was never resolved via EnrichPlaceholderIdx (still
' the Discovery.bas sentinel -1). The latter case matters: without this
' guard, two shapes that are both placeholders but both still carry the
' unresolved -1 sentinel would compare -1=-1 and silently report a false
' full-confidence placeholder match. Treating "unresolved" the same as
' "not applicable" is the safe, conservative choice -- consistent with
' never fabricating a signal from missing data.
Private Function PlaceholderScore(candidate As Candidate, reference As Candidate, ByRef applicable As Boolean) As Double
    If Not reference.HasPlaceholder Or reference.PlaceholderIdx = -1 Then
        applicable = False
        PlaceholderScore = 0
        Exit Function
    End If
    applicable = True

    If Not candidate.HasPlaceholder Or candidate.PlaceholderIdx = -1 Then
        PlaceholderScore = 0
        Exit Function
    End If

    If candidate.PlaceholderType = reference.PlaceholderType And candidate.PlaceholderIdx = reference.PlaceholderIdx Then
        PlaceholderScore = 1
    Else
        PlaceholderScore = 0
    End If
End Function

Private Function GeometryScore(candidate As Candidate, reference As Candidate) As Double
    Dim dx As Double, dy As Double, dw As Double, dh As Double
    Dim posScore As Double, sizeScore As Double

    dx = candidate.PositionX - reference.PositionX
    dy = candidate.PositionY - reference.PositionY
    posScore = 1 - (Sqr(dx * dx + dy * dy) / POSITION_TOLERANCE_EMU)
    If posScore < 0 Then posScore = 0

    dw = candidate.SizeCx - reference.SizeCx
    dh = candidate.SizeCy - reference.SizeCy
    sizeScore = 1 - (Sqr(dw * dw + dh * dh) / SIZE_TOLERANCE_EMU)
    If sizeScore < 0 Then sizeScore = 0

    GeometryScore = (posScore + sizeScore) / 2
End Function

Private Function ShapeTypeScore(candidate As Candidate, reference As Candidate) As Double
    ShapeTypeScore = IIf(candidate.ShapeType = reference.ShapeType, 1, 0)
End Function

' Weakest signal, last resort per spec -- Discovery.bas only captures
' HasText, not actual text content, so like the Python original this
' degrades to a has-text match rather than a real pattern comparison.
Private Function ContentPatternScore(candidate As Candidate, reference As Candidate) As Double
    ContentPatternScore = IIf(candidate.HasText = reference.HasText, 1, 0)
End Function

Private Function ConfidenceFor(score As Double) As String
    If score >= HIGH_THRESHOLD Then
        ConfidenceFor = "high"
    ElseIf score >= MEDIUM_THRESHOLD Then
        ConfidenceFor = "medium"
    Else
        ConfidenceFor = "low"
    End If
End Function

' ---------------------------------------------------------------------
' Match: two-tier dispatch (specs/matching.md)
' ---------------------------------------------------------------------

' Match `reference`'s field against the best candidate in `candidates`.
' `validTags`, if supplied, is anything For-Each-able (Collection or array)
' of valid tag strings; omit it to trust any non-empty IdentityTag.
'
' Tier 1 (trust, no scoring): any candidate whose IdentityTag is non-empty
' (and, if validTags was given, contained in it) is trusted directly.
' Exactly one such candidate is an immediate high-confidence match; more
' than one is a same-tag collision, which can't be silently resolved.
'
' Tier 2 (scored fallback): every candidate is scored via ScoreCandidate().
' The top scorer is accepted only if it clears the high-confidence
' threshold AND isn't ambiguously close to other similarly-scored siblings
' (sibling_ambiguity: z-order is tried as a supplementary disambiguator
' before giving up and flagging). Medium confidence is always flagged;
' low confidence is always unmatched.
Public Function Match(candidates() As Candidate, reference As Candidate, Optional validTags As Variant) As MatchResult
    Dim result As MatchResult
    Dim lo As Long, hi As Long, i As Long
    ' `candidates` can be genuinely unallocated (a caller with zero
    ' candidates -- e.g. onboarding's untagged pool when every candidate is
    ' either already tagged or pure decoration -- never gets to ReDim it at
    ' all; see AGENTS.md's Known Patterns on why ReDim-to-(1 To 0) can't be
    ' used to represent that instead). Treat that the same as the existing
    ' hi < lo "no candidates" path below, not as a fresh error.
    On Error Resume Next
    lo = LBound(candidates)
    hi = UBound(candidates)
    If Err.Number <> 0 Then
        lo = 1
        hi = 0
    End If
    On Error GoTo 0

    Dim taggedIdx As Collection
    Set taggedIdx = New Collection
    If hi >= lo Then
        For i = lo To hi
            If candidates(i).IdentityTag <> "" Then
                If IsValidTag(candidates(i).IdentityTag, validTags) Then
                    taggedIdx.Add i
                End If
            End If
        Next i
    End If

    If taggedIdx.count = 1 Then
        result.HasCandidate = True
        result.CandidateIndex = taggedIdx(1)
        result.Confidence = "high"
        result.HasScore = False
        result.Reason = "existing identity tag"
        Match = result
        Exit Function
    ElseIf taggedIdx.count > 1 Then
        result.HasCandidate = False
        result.Confidence = "medium"
        result.HasScore = False
        result.Reason = taggedIdx.count & " candidates already carry this identity tag -- collision"
        Match = result
        Exit Function
    End If

    If hi < lo Then
        result.HasCandidate = False
        result.Confidence = "low"
        result.HasScore = False
        result.Reason = "no candidates to match against"
        Match = result
        Exit Function
    End If

    Dim scores() As Double
    ReDim scores(lo To hi)
    For i = lo To hi
        scores(i) = ScoreCandidate(candidates(i), reference)
    Next i

    Dim bestIdx As Long, bestScore As Double
    bestIdx = lo
    bestScore = scores(lo)
    For i = lo + 1 To hi
        If scores(i) > bestScore Then
            bestScore = scores(i)
            bestIdx = i
        End If
    Next i

    Dim conf As String
    conf = ConfidenceFor(bestScore)

    If conf = "high" Then
        Dim tiedIdx As Collection
        Set tiedIdx = New Collection
        For i = lo To hi
            If (bestScore - scores(i)) < SIBLING_GAP_THRESHOLD Then tiedIdx.Add i
        Next i

        If tiedIdx.count > 1 Then
            Dim j As Variant, dist As Long, minDist As Long
            minDist = -1
            For Each j In tiedIdx
                dist = Abs(candidates(j).ZOrder - reference.ZOrder)
                If minDist = -1 Or dist < minDist Then minDist = dist
            Next j

            Dim winners As Collection
            Set winners = New Collection
            For Each j In tiedIdx
                dist = Abs(candidates(j).ZOrder - reference.ZOrder)
                If dist = minDist Then winners.Add j
            Next j

            If winners.count = 1 Then
                bestIdx = winners(1)
                bestScore = scores(bestIdx)
                conf = ConfidenceFor(bestScore)
            Else
                result.HasCandidate = False
                result.Confidence = "medium"
                result.HasScore = True
                result.Score = bestScore
                result.Reason = "sibling ambiguity: " & tiedIdx.count & " candidates score within " & _
                    SIBLING_GAP_THRESHOLD & " of each other and z-order doesn't disambiguate"
                Match = result
                Exit Function
            End If
        End If
    End If

    Select Case conf
        Case "high"
            result.HasCandidate = True
            result.CandidateIndex = bestIdx
            result.Confidence = "high"
            result.HasScore = True
            result.Score = bestScore
            result.Reason = "scored match"
        Case "medium"
            result.HasCandidate = False
            result.Confidence = "medium"
            result.HasScore = True
            result.Score = bestScore
            result.Reason = "medium confidence -- flagged for human confirmation"
        Case Else
            result.HasCandidate = False
            result.Confidence = "low"
            result.HasScore = True
            result.Score = bestScore
            result.Reason = "low confidence -- unmatched"
    End Select

    Match = result
End Function

Private Function IsValidTag(tag As String, validTags As Variant) As Boolean
    If IsMissing(validTags) Then
        IsValidTag = True
        Exit Function
    End If
    Dim v As Variant
    For Each v In validTags
        If CStr(v) = tag Then
            IsValidTag = True
            Exit Function
        End If
    Next v
    IsValidTag = False
End Function

' ---------------------------------------------------------------------
' PlaceholderIdx raw-OOXML fallback -- closes the gap Discovery.bas
' flagged (PowerPoint's object model has no PlaceholderFormat.Idx
' equivalent). Per specs/vba-port.md: "only fall back to raw OOXML where
' VBA's object model genuinely has no path, and flag that fallback
' explicitly." See SPIKE_NOTES_Matching.md for the full explanation and
' its own real limitations (duplicate shape names, unsaved-edits skew).
' ---------------------------------------------------------------------

' Fill in PlaceholderIdx for every HasPlaceholder candidate in `candidates`
' by reading `partName`'s raw XML out of `pptxPath` on disk (NOT the live
' in-memory Shapes collection -- see SPIKE_NOTES_Matching.md's "reads
' last-saved state" divergence). Returns False if the fallback itself
' failed (extraction/parse error); callers should then treat every
' PlaceholderIdx as still unresolved (-1), never assume 0.
Public Function EnrichPlaceholderIdx(ByRef candidates() As Candidate, pptxPath As String, partName As String) As Boolean
    Dim dom As Object
    Set dom = LoadPartXml(pptxPath, partName)
    If dom Is Nothing Then
        EnrichPlaceholderIdx = False
        Exit Function
    End If

    Dim i As Long, lo As Long, hi As Long
    ' `candidates` may be genuinely unallocated (a slide with zero
    ' candidates) -- see Match()'s identical guard above.
    On Error Resume Next
    lo = LBound(candidates)
    hi = UBound(candidates)
    If Err.Number <> 0 Then
        EnrichPlaceholderIdx = True ' nothing to enrich is not a failure
        Exit Function
    End If
    On Error GoTo 0
    For i = lo To hi
        If candidates(i).HasPlaceholder Then
            candidates(i).PlaceholderIdx = PlaceholderIdxFromDom(dom, candidates(i).Name)
        End If
    Next i
    EnrichPlaceholderIdx = True
End Function

' Extracts `pptxPath` (a zip under the hood) to a temp folder and loads
' `partName`'s XML as an MSXML2.DOMDocument60, or Nothing on any failure.
'
' Originally used Shell.Application's Namespace()/CopyHere zip-folder
' support -- the standard VBA escape hatch for "no zip library in stock
' VBA" -- but that technique was confirmed BROKEN under COM automation on
' 2026-07-25: Namespace(zipPath) reliably returned Nothing even against an
' independently-verified-valid .pptx (opened cleanly by .NET's
' ZipFile.OpenRead, 39 real entries), reproduced 3x including with an
' explicit delay to rule out a shell-notification race. Root cause not
' fully isolated; see SPIKE_NOTES_Matching.md's original finding.
'
' Replaced with shelling out to tar.exe (bsdtar, bundled with Windows 10
' 1803+ and Windows 11 by default -- confirmed present and working against
' a real .pptx on this exact machine, including extracting a real slide
' part and loading its actual XML content, before this replacement was
' written). tar.exe auto-detects zip format despite the name; this avoids
' Shell.Application/the Windows shell namespace entirely, which is exactly
' the layer that was failing under automation.
Private Function LoadPartXml(pptxPath As String, partName As String) As Object
    On Error GoTo Fail

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim tempRoot As String
    tempRoot = Environ$("TEMP") & "\vba_ooxml_" & CStr(CLng(Timer * 1000))
    fso.CreateFolder tempRoot

    Dim zipPath As String, extractPath As String
    zipPath = tempRoot & "\copy.zip"
    extractPath = tempRoot & "\extracted"
    fso.CopyFile pptxPath, zipPath, True
    fso.CreateFolder extractPath

    ' -xf: extract from file. -C: change to this directory first. The
    ' final argument is the specific archive entry to extract (tar accepts
    ' forward slashes on Windows) -- only that one part is pulled out, not
    ' the whole package. WScript.Shell.Run's windowStyle=0 keeps this
    ' invisible; waitOnReturn=True (last arg) blocks until tar exits and
    ' returns its real exit code, so there is no CopyHere-style "did it
    ' actually finish yet" polling needed this time -- Run's own return is
    ' the completion signal.
    Dim cmd As String
    cmd = "cmd.exe /c tar -xf """ & zipPath & """ -C """ & extractPath & """ """ & partName & """"

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Dim exitCode As Long
    exitCode = wsh.Run(cmd, 0, True)
    If exitCode <> 0 Then GoTo Fail

    Dim targetPath As String
    targetPath = extractPath & "\" & Replace(partName, "/", "\")
    If Not fso.FileExists(targetPath) Then GoTo Fail

    ' ProgID is "MSXML2.DOMDocument.6.0" (dotted), not "MSXML2.DOMDocument60"
    ' (no dots) -- confirmed by direct probe (2026-07-25) that the no-dot
    ' form raises Err 429 "ActiveX component can't create object" on this
    ' machine while the dotted form succeeds and returns the identical real
    ' DOMDocument60 object (TypeName confirmed). A second, previously-masked
    ' bug: the original Shell.Application failure always short-circuited
    ' before execution ever reached this line, so this was never actually
    ' exercised until the zip-extraction technique was replaced.
    Dim dom As Object
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False
    dom.setProperty "SelectionLanguage", "XPath"
    If Not dom.Load(targetPath) Then GoTo Fail

    Set LoadPartXml = dom
    Exit Function

Fail:
    Set LoadPartXml = Nothing
End Function

' Finds `shapeName`'s <p:ph> element (if any) via its <p:cNvPr name="...">
' sibling-of-parent relationship and reads its idx attribute, applying
' OOXML's own default (0) when <p:ph> is present but idx is omitted --
' exactly matching discovery.py's _placeholder_info. Returns -1 (not
' resolvable) if the shape name doesn't appear exactly once in this part,
' or if it has no <p:ph> at all.
Private Function PlaceholderIdxFromDom(dom As Object, shapeName As String) As Long
    Dim cNvPrNodes As Object
    Set cNvPrNodes = dom.SelectNodes("//*[local-name()='cNvPr'][@name=" & XPathLiteral(shapeName) & "]")

    ' Zero matches (name not found in this part) or more than one (duplicate
    ' shape names on this slide/layout -- InjectPrimitive.bas's SPIKE_NOTES.md
    ' already documents this as real, observed behavior on a live deck) are
    ' both ambiguous or absent. Refuse to guess, same posture as
    ' FindShapeByRoleTag's multiplicity handling.
    If cNvPrNodes.Length <> 1 Then
        PlaceholderIdxFromDom = -1
        Exit Function
    End If

    Dim phNode As Object
    Set phNode = cNvPrNodes.Item(0).parentNode.SelectSingleNode("*[local-name()='nvPr']/*[local-name()='ph']")
    If phNode Is Nothing Then
        PlaceholderIdxFromDom = -1
        Exit Function
    End If

    Dim idxAttr As Object
    Set idxAttr = phNode.Attributes.getNamedItem("idx")
    If idxAttr Is Nothing Then
        PlaceholderIdxFromDom = 0
    Else
        PlaceholderIdxFromDom = CLng(idxAttr.Text)
    End If
End Function

' XPath 1.0 has no escape character for quotes inside a string literal --
' the standard workaround is concat() when a literal contains both quote
' kinds. Shape names containing a single quote (rare, but not impossible)
' are the reason this exists rather than a plain "'" & s & "'".
Private Function XPathLiteral(s As String) As String
    If InStr(s, "'") = 0 Then
        XPathLiteral = "'" & s & "'"
    ElseIf InStr(s, """") = 0 Then
        XPathLiteral = """" & s & """"
    Else
        Dim parts() As String, i As Long, result As String
        parts = Split(s, "'")
        result = "concat("
        For i = LBound(parts) To UBound(parts)
            If i > LBound(parts) Then result = result & ", ""'"", "
            result = result & "'" & parts(i) & "'"
        Next i
        result = result & ")"
        XPathLiteral = result
    End If
End Function

' ---------------------------------------------------------------------
' Manual smoke tests -- not a real test harness (no VBA unit-test
' framework wired up here, same as every other module in this project).
' See SPIKE_NOTES_Matching.md for the full recipe and expected values,
' cross-checked against tests/test_matching.py's already-proven results.
' ---------------------------------------------------------------------

' Run with test-fixtures/shp-groupshape.pptx open as the active
' presentation's first slide. Expects HasCandidate=True, name "Oval 2",
' Confidence="high" -- the same 4-way tie broken by z-order that
' tests/test_matching.py::test_shp_groupshape_sibling_ambiguity_resolved_by_zorder
' proves in Python.
Public Sub ManualSmokeTest_SiblingAmbiguity()
    Dim sld As Object
    Set sld = Application.ActivePresentation.Slides(1)

    Dim candidates() As Candidate
    candidates = Discovery.DiscoverSlide(sld)

    Dim reference As Candidate
    reference.ZOrder = 2
    reference.HasText = False
    reference.PlaceholderIdx = -1 ' not testing placeholder matching here

    Dim result As MatchResult
    result = Match(candidates, reference)

    Dim msg As String
    msg = "Confidence=" & result.Confidence & " HasCandidate=" & result.HasCandidate
    If result.HasCandidate Then msg = msg & " Name=" & candidates(result.CandidateIndex).Name
    Debug.Print msg
    MsgBox msg & " (expected: high / Oval 2)"
End Sub

' Run with test-fixtures/mst-slide-layouts.pptx's two layouts applied to
' slides 1 and 2 of the active presentation (or reached via
' Designs(1).SlideMaster.CustomLayouts(1)/(2) directly). `pptxPath` is the
' file's path on disk, needed for the placeholder-idx XML fallback.
' Expects Confidence="medium" -- placeholder-index match alone must not
' force an auto-accept when geometry has drifted this far, matching
' tests/test_matching.py::test_mst_slide_layouts_placeholder_index_alone_does_not_force_high_confidence.
Public Sub ManualSmokeTest_PlaceholderIndex(pptxPath As String, layout1 As Object, layout2 As Object)
    Dim candidates1() As Candidate, candidates2() As Candidate
    candidates1 = Discovery.DiscoverCustomLayout(layout1)
    candidates2 = Discovery.DiscoverCustomLayout(layout2)

    If Not EnrichPlaceholderIdx(candidates1, pptxPath, "ppt/slideLayouts/slideLayout1.xml") Then
        MsgBox "EnrichPlaceholderIdx failed for layout1 -- see SPIKE_NOTES_Matching.md"
        Exit Sub
    End If
    If Not EnrichPlaceholderIdx(candidates2, pptxPath, "ppt/slideLayouts/slideLayout2.xml") Then
        MsgBox "EnrichPlaceholderIdx failed for layout2 -- see SPIKE_NOTES_Matching.md"
        Exit Sub
    End If

    Dim i As Long, referenceIdx As Long
    referenceIdx = -1
    For i = LBound(candidates1) To UBound(candidates1)
        If candidates1(i).Name = "Text Placeholder 3" Then referenceIdx = i
    Next i
    If referenceIdx = -1 Then
        MsgBox "Could not find 'Text Placeholder 3' on layout 1"
        Exit Sub
    End If

    Dim result As MatchResult
    result = Match(candidates2, candidates1(referenceIdx))

    Dim msg As String
    msg = "Confidence=" & result.Confidence & " HasCandidate=" & result.HasCandidate
    Debug.Print msg
    MsgBox msg & " (expected: medium / False)"
End Sub
