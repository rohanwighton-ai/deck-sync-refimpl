Attribute VB_Name = "Verification"
Option Explicit

' VBA port of src/verification.py's verify_structure()/verify_z_order() --
' never ported before now. InjectPrimitive.bas only ever covered
' inject_primitive itself; specs/vba-port.md's 6-module port order never
' listed a separate "verification" module because nothing needed the
' structural/z-order checks until specs/slide-duplication-trigger.md made
' them mandatory before tagging any duplicate.
'
' Both functions take live Slide objects, not pre-discovered Candidate
' arrays: Discovery.Candidate.IdentityTag is always "" straight out of
' DiscoverSlide/DiscoverSlideWithShapes (discovery does not read tags, per
' discovery.md's non-goals -- true of the Python original too), but
' tag-based pairing is the entire point of both checks here (a pure
' z-order swap must never also look like a structural defect, which
' positional pairing alone cannot tell apart from a real mismatch). So
' each function does its own DiscoverSlideWithShapes + role-tag read
' internally, mirroring what Onboarding.BuildTemplateFieldShapes already
' does for the same underlying reason.
'
' Dictionary(tag As String -> index As Long) is used for tag lookups --
' Long is a primitive, Variant-safe value, unlike a Candidate (a UDT can't
' go into a Dictionary at all; see AGENTS.md's Known Patterns).

Public Type StructuralMismatch
    Index As Long   ' z_order for a tagged shape's mismatch, positional index for untagged; -1 for the shape_count case
    Kind As String  ' "shape_count" | "type" | "missing_in_duplicate" | "extra_in_duplicate"
    Detail As String
End Type

Public Type StructuralVerification
    SourceCount As Long
    DuplicateCount As Long
    MismatchCount As Long
    Mismatches() As StructuralMismatch ' possibly unallocated when MismatchCount = 0 -- see AGENTS.md's ReDim(1 To 0) Known Pattern; check MismatchCount, never LBound/UBound directly
    Ok As Boolean
End Type

Public Type ZOrderMismatch
    TagA As String
    TagB As String
    Detail As String
End Type

Public Type ZOrderVerification
    PairsChecked As Long
    MismatchCount As Long
    Mismatches() As ZOrderMismatch ' possibly unallocated when MismatchCount = 0 -- same reason as StructuralVerification.Mismatches
    Ok As Boolean
End Type

' ---------------------------------------------------------------------
' Shared: discover a slide's candidates + their real role tags
' ---------------------------------------------------------------------

' Returns candidates() (possibly unallocated if the slide has none) and,
' via the two ByRef out-params, parallel tags()/shapes() arrays at the same
' indices. tags(i) is "" for an untagged candidate, never Nothing/omitted --
' every index has an entry, unlike a Dictionary that would only hold tagged
' ones.
Private Function DiscoverWithTags(sld As Object, ByRef tags() As String, ByRef shapes() As Object) As Candidate()
    Dim candidates() As Candidate
    candidates = Discovery.DiscoverSlideWithShapes(sld, shapes)

    Dim lo As Long, hi As Long, hasAny As Boolean
    On Error Resume Next
    lo = LBound(candidates)
    hi = UBound(candidates)
    hasAny = (Err.Number = 0)
    On Error GoTo 0

    If Not hasAny Then
        DiscoverWithTags = candidates ' unallocated -- ReDim tags() to match
        Exit Function
    End If

    ReDim tags(lo To hi)
    Dim i As Long
    For i = lo To hi
        tags(i) = shapes(i).Tags("role")
    Next i

    DiscoverWithTags = candidates
End Function

' Builds a tag -> index Dictionary from a tags() array (only non-blank
' entries), safely handling an unallocated tags() array (zero candidates).
Private Function IndexByTag(tags() As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim lo As Long, hi As Long, hasAny As Boolean
    On Error Resume Next
    lo = LBound(tags)
    hi = UBound(tags)
    hasAny = (Err.Number = 0)
    On Error GoTo 0

    If Not hasAny Then
        Set IndexByTag = result
        Exit Function
    End If

    Dim i As Long
    For i = lo To hi
        If tags(i) <> "" Then result(tags(i)) = i
    Next i

    Set IndexByTag = result
End Function

Private Sub AppendStructuralMismatch(ByRef mismatches() As StructuralMismatch, ByRef n As Long, idx As Long, kind As String, detail As String)
    n = n + 1
    ReDim Preserve mismatches(1 To n)
    mismatches(n).Index = idx
    mismatches(n).Kind = kind
    mismatches(n).Detail = detail
End Sub

' ---------------------------------------------------------------------
' verify_structure
' ---------------------------------------------------------------------

' Shape count, type, and identity-tag correspondence between sourceSld and
' duplicateSld, checked explicitly rather than assumed from Slide.Duplicate
' succeeding. Tagged shapes are paired by role tag; untagged shapes fall
' back to positional pairing within just the untagged subsequence (no
' other signal exists for them) -- field-for-field against
' src/verification.py's verify_structure().
Public Function VerifyStructure(sourceSld As Object, duplicateSld As Object) As StructuralVerification
    Dim result As StructuralVerification
    Dim n As Long
    n = 0

    Dim sourceShapes() As Object, duplicateShapes() As Object
    Dim sourceTags() As String, duplicateTags() As String
    Dim sourceCandidates() As Candidate, duplicateCandidates() As Candidate
    sourceCandidates = DiscoverWithTags(sourceSld, sourceTags, sourceShapes)
    duplicateCandidates = DiscoverWithTags(duplicateSld, duplicateTags, duplicateShapes)

    Dim sLo As Long, sHi As Long, sHas As Boolean
    On Error Resume Next
    sLo = LBound(sourceCandidates): sHi = UBound(sourceCandidates): sHas = (Err.Number = 0)
    On Error GoTo 0
    Dim dLo As Long, dHi As Long, dHas As Boolean
    On Error Resume Next
    dLo = LBound(duplicateCandidates): dHi = UBound(duplicateCandidates): dHas = (Err.Number = 0)
    On Error GoTo 0

    result.SourceCount = IIf(sHas, sHi - sLo + 1, 0)
    result.DuplicateCount = IIf(dHas, dHi - dLo + 1, 0)

    If result.SourceCount <> result.DuplicateCount Then
        AppendStructuralMismatch result.Mismatches, n, -1, "shape_count", _
            "source has " & result.SourceCount & " shape(s), duplicate has " & result.DuplicateCount
    End If

    Dim sourceByTag As Object, duplicateByTag As Object
    Set sourceByTag = IndexByTag(sourceTags)
    Set duplicateByTag = IndexByTag(duplicateTags)

    Dim tag As Variant
    For Each tag In sourceByTag.Keys
        Dim si As Long
        si = sourceByTag(tag)
        If duplicateByTag.Exists(tag) Then
            Dim di As Long
            di = duplicateByTag(tag)
            If sourceCandidates(si).ShapeType <> duplicateCandidates(di).ShapeType Then
                AppendStructuralMismatch result.Mismatches, n, sourceCandidates(si).ZOrder, "type", _
                    "tagged shape '" & tag & "': source is '" & sourceCandidates(si).ShapeType & "', duplicate is '" & duplicateCandidates(di).ShapeType & "'"
            End If
        Else
            AppendStructuralMismatch result.Mismatches, n, sourceCandidates(si).ZOrder, "missing_in_duplicate", _
                "tagged shape '" & tag & "' has no counterpart in duplicate"
        End If
    Next tag
    For Each tag In duplicateByTag.Keys
        If Not sourceByTag.Exists(tag) Then
            Dim di2 As Long
            di2 = duplicateByTag(tag)
            AppendStructuralMismatch result.Mismatches, n, duplicateCandidates(di2).ZOrder, "extra_in_duplicate", _
                "duplicate has tagged shape '" & tag & "' with no source counterpart"
        End If
    Next tag

    ' Untagged positional pairing, within just the untagged subsequence.
    Dim sourceUntagged() As Candidate, duplicateUntagged() As Candidate
    Dim sun As Long, dun As Long
    sun = 0: dun = 0
    Dim i As Long
    If sHas Then
        For i = sLo To sHi
            If sourceTags(i) = "" Then
                sun = sun + 1
                ReDim Preserve sourceUntagged(1 To sun)
                sourceUntagged(sun) = sourceCandidates(i)
            End If
        Next i
    End If
    If dHas Then
        For i = dLo To dHi
            If duplicateTags(i) = "" Then
                dun = dun + 1
                ReDim Preserve duplicateUntagged(1 To dun)
                duplicateUntagged(dun) = duplicateCandidates(i)
            End If
        Next i
    End If

    Dim commonUntagged As Long
    commonUntagged = IIf(sun < dun, sun, dun)
    For i = 1 To commonUntagged
        If sourceUntagged(i).ShapeType <> duplicateUntagged(i).ShapeType Then
            AppendStructuralMismatch result.Mismatches, n, i - 1, "type", _
                "source shape " & (i - 1) & " is '" & sourceUntagged(i).ShapeType & "', duplicate is '" & duplicateUntagged(i).ShapeType & "'"
        End If
    Next i
    For i = commonUntagged + 1 To sun
        AppendStructuralMismatch result.Mismatches, n, i - 1, "missing_in_duplicate", _
            "source shape " & (i - 1) & " ('" & sourceUntagged(i).Name & "') has no counterpart"
    Next i
    For i = commonUntagged + 1 To dun
        AppendStructuralMismatch result.Mismatches, n, i - 1, "extra_in_duplicate", _
            "duplicate shape " & (i - 1) & " ('" & duplicateUntagged(i).Name & "') has no source counterpart"
    Next i

    result.MismatchCount = n
    result.Ok = (n = 0)
    VerifyStructure = result
End Function

' ---------------------------------------------------------------------
' verify_z_order
' ---------------------------------------------------------------------

' Kept distinct from VerifyStructure's count/type/tag correspondence: a
' duplicate can have exactly the right shapes, tags, and values while a
' stacking-order regression still makes an overlaid field invisible.
' Checks every pair of commonly-tagged shapes (not just adjacent ones), so
' a single swap deep in the stack is caught regardless of how many other
' shapes sit between the two that moved -- field-for-field against
' src/verification.py's verify_z_order().
Public Function VerifyZOrder(sourceSld As Object, duplicateSld As Object) As ZOrderVerification
    Dim result As ZOrderVerification
    Dim n As Long
    n = 0

    Dim sourceShapes() As Object, duplicateShapes() As Object
    Dim sourceTags() As String, duplicateTags() As String
    Dim sourceCandidates() As Candidate, duplicateCandidates() As Candidate
    sourceCandidates = DiscoverWithTags(sourceSld, sourceTags, sourceShapes)
    duplicateCandidates = DiscoverWithTags(duplicateSld, duplicateTags, duplicateShapes)

    Dim sourceByTag As Object, duplicateByTag As Object
    Set sourceByTag = IndexByTag(sourceTags)
    Set duplicateByTag = IndexByTag(duplicateTags)

    ' Common tags, as a plain array (needed for the i<j pairwise walk below;
    ' a Dictionary's .Keys already gives a 0-based array of Variants in VBA,
    ' which is fine to index directly).
    Dim commonTags() As String
    Dim cn As Long
    cn = 0
    Dim tag As Variant
    For Each tag In sourceByTag.Keys
        If duplicateByTag.Exists(tag) Then
            cn = cn + 1
            ReDim Preserve commonTags(1 To cn)
            commonTags(cn) = CStr(tag)
        End If
    Next tag

    Dim a As Long, b As Long
    For a = 1 To cn
        For b = a + 1 To cn
            Dim tagA As String, tagB As String
            tagA = commonTags(a)
            tagB = commonTags(b)

            Dim sourceBelow As Boolean, duplicateBelow As Boolean
            sourceBelow = (sourceCandidates(sourceByTag(tagA)).ZOrder < sourceCandidates(sourceByTag(tagB)).ZOrder)
            duplicateBelow = (duplicateCandidates(duplicateByTag(tagA)).ZOrder < duplicateCandidates(duplicateByTag(tagB)).ZOrder)

            If sourceBelow <> duplicateBelow Then
                n = n + 1
                ReDim Preserve result.Mismatches(1 To n)
                result.Mismatches(n).TagA = tagA
                result.Mismatches(n).TagB = tagB
                result.Mismatches(n).Detail = "'" & tagA & "' is " & IIf(sourceBelow, "below", "above") & " '" & tagB & "' in the source, but " & _
                    IIf(duplicateBelow, "below", "above") & " it in the duplicate"
            End If
        Next b
    Next a

    result.PairsChecked = cn * (cn - 1) \ 2
    result.MismatchCount = n
    result.Ok = (n = 0)
    VerifyZOrder = result
End Function
