Attribute VB_Name = "Discovery"
Option Explicit

' VBA port of src/discovery.py's discover(), per specs/vba-port.md's port
' order (module 1 -- the foundational one everything else reads). Ports the
' mechanism, not the Python workaround: the Python side hand-parses raw
' <p:spTree> OOXML because it has no host application; this port walks the
' live Shapes/GroupShapes collections PowerPoint already gives every open
' presentation, so there is no XML in this file at all. See
' SPIKE_NOTES_Discovery.md for the full list of deliberate divergences (a
' couple of real, unresolved gaps -- placeholder idx chief among them -- not
' just stylistic ones) and the manual verification recipe, since there is no
' Windows/Office install in this environment to run it against.
'
' 2026-07-25: added DiscoverSlideWithShapes/DiscoverCustomLayoutWithShapes
' (module 5, onboarding, needed a way back from a Candidate to its live
' Shape -- see the gap note above DiscoverSlideWithShapes and
' SPIKE_NOTES_Onboarding.md). DiscoverSlide/DiscoverCustomLayout's own
' behavior is unchanged.

Public Type Candidate
    Name As String              ' shape name, for humans -- never used as an identity key (see identity_tags.py)
    GroupPath As String         ' "/"-joined chain of enclosing group names, "" if top-level
    ZOrder As Long              ' 1-based document-order position among discovered leaf shapes
    ShapeType As String         ' "autoshape_or_textbox" | "picture" | SHAPE_TYPE_DEVICE
    HasPlaceholder As Boolean
    PlaceholderType As String   ' best-effort OOXML-style label ("title","body",...); "" if not a placeholder
    PlaceholderIdx As Long      ' always -1 here -- the object model exposes no such property, see notes
    HasText As Boolean
    PositionX As Long           ' EMU, converted from Shape.Left (see PointsToEmu)
    PositionY As Long           ' EMU, converted from Shape.Top
    SizeCx As Long              ' EMU, converted from Shape.Width
    SizeCy As Long              ' EMU, converted from Shape.Height
    IdentityTag As String       ' always "" out of DiscoverSlide -- discovery does not read tags, per discovery.md's non-goals
End Type

' THE SHAPE TYPE FOR A DEVICE -- a group addressed as ONE field.
'
' Named rather than typed as a literal in two modules: Onboarding.IsCandidateField
' has to admit it, and Matching.ShapeTypeScore compares it. A string spelled twice
' is the write-it-twice class, and this project has paid for that repeatedly.
Public Const SHAPE_TYPE_DEVICE As String = "device"


' Walk `sld`'s shape tree and return every candidate leaf shape (autoshape,
' textbox, placeholder, or picture), recursing into groups. Mirrors
' discovery.py's discover(): type-agnostic, never treats a group as one
' opaque candidate, tags leaves not containers.
Public Function DiscoverSlide(sld As Object) As Candidate()
    ' Deliberately NOT pre-ReDim'd to (1 To 0) here: a real, confirmed VBA
    ' restriction (see AGENTS.md's Known Patterns) makes ReDim-to-an-empty-
    ' range throw "Subscript out of range" at runtime. Left genuinely
    ' unallocated instead -- Walk's own ReDim Preserve on first append
    ' handles allocation correctly (confirmed), and a slide with zero
    ' candidates returns a genuinely unallocated array, which callers must
    ' check for with an error-guarded UBound/LBound, never assume allocated.
    Dim results() As Candidate
    Dim shapes() As Object ' unused by this entry point -- see DiscoverSlideWithShapes
    Dim count As Long
    count = 0
    Dim z As Long
    z = 0

    Walk sld.Shapes, "", results, shapes, count, z

    DiscoverSlide = results
End Function

' Same walk, but rooted at a slide layout's shapes instead of a slide's --
' the direct equivalent of discovery.py's discover_from_pptx_layout(), for
' matching against mst-slide-layouts.pptx-style fixtures (layouts/masters,
' no slides). PowerPoint exposes a layout's shapes via
' Slide.CustomLayout.Shapes (borrowed from an existing slide) or
' SlideMaster.CustomLayouts(i).Shapes directly -- callers pass whichever
' CustomLayout object they already have.
Public Function DiscoverCustomLayout(layout As Object) As Candidate()
    ' See DiscoverSlide's comment on why this is deliberately not
    ' pre-ReDim'd to (1 To 0).
    Dim results() As Candidate
    Dim shapes() As Object ' unused by this entry point -- see DiscoverCustomLayoutWithShapes
    Dim count As Long
    count = 0
    Dim z As Long
    z = 0

    Walk layout.Shapes, "", results, shapes, count, z

    DiscoverCustomLayout = results
End Function

' Gap closed for onboarding (port-order step 5): neither DiscoverSlide nor
' DiscoverCustomLayout expose the live Shape object behind each Candidate,
' because no caller needed one before now -- InjectPrimitive.bas re-locates
' a shape by its role TAG (FindShapeByRoleTag), never by Candidate identity.
' Onboarding is the first caller that must go the other direction: given a
' Candidate a human/matcher accepted, write a tag onto its actual shape (or,
' for a template, read the tag a shape already carries). Candidate.ZOrder is
' exactly "1-based position among discovered leaf shapes" (see the Public
' Type comment below) and Walk assigns it in one deterministic pass over the
' same tree DiscoverSlide/DiscoverCustomLayout already traverse -- so a
' second, identically-filtered walk that also stashes each leaf's live Shape
' reference at the same index is a safe, exact pairing, not a heuristic
' re-match. Purely additive: DiscoverSlide/DiscoverCustomLayout's own
' behavior/output is unchanged (see SPIKE_NOTES_Onboarding.md).
Public Function DiscoverSlideWithShapes(sld As Object, ByRef shapes() As Object) As Candidate()
    ' See DiscoverSlide's comment on why this is deliberately not
    ' pre-ReDim'd to (1 To 0).
    Dim results() As Candidate
    Dim count As Long
    count = 0
    Dim z As Long
    z = 0

    Walk sld.Shapes, "", results, shapes, count, z

    DiscoverSlideWithShapes = results
End Function

Public Function DiscoverCustomLayoutWithShapes(layout As Object, ByRef shapes() As Object) As Candidate()
    ' See DiscoverSlide's comment on why this is deliberately not
    ' pre-ReDim'd to (1 To 0).
    Dim results() As Candidate
    Dim count As Long
    count = 0
    Dim z As Long
    z = 0

    Walk layout.Shapes, "", results, shapes, count, z

    DiscoverCustomLayoutWithShapes = results
End Function

Private Sub Walk(shapesColl As Object, groupPath As String, ByRef results() As Candidate, ByRef shapes() As Object, ByRef count As Long, ByRef z As Long)
    Dim shp As Object
    For Each shp In shapesColl
        If shp.Type = msoGroup Then
            Dim childPath As String
            If groupPath = "" Then
                childPath = shp.Name
            Else
                childPath = groupPath & "/" & shp.Name
            End If

            ' A DEVICE IS ONE FIELD, NOT ITS PARTS.
            '
            ' The milestone timeline is a group of ~30 named shapes -- seven
            ' slots x three circles, plus a label and a date each, plus a track
            ' and a bar. Walking into it made every one of them a candidate, so
            ' Rohan's timeline appeared as 21 fields to tag by hand. He had
            ' already solved this in the deck: the shapes carry a naming
            ' convention (MS1_ON, MS2_DATE, MILESTONE_TIMELINE...), which is
            ' what "load shape modules as prenamed per slide entities" meant.
            '
            ' Injection has always treated the device as one addressable thing
            ' (InjectPrimitive:263 asks SlotCount before routing), and marking
            ' has too (BatchOnboardFlow:1503). DISCOVERY was the one that still
            ' saw the parts -- which is exactly the "writing is solved,
            ' recognising is not" gap recorded on 2026-08-13.
            '
            ' Recognised by COUNTING SLOTS rather than by matching the group's
            ' name, so a renamed timeline still resolves and a group that merely
            ' happens to be called MILESTONE_TIMELINE but carries no slots does
            ' not. The device's own integrity check reports a missing part by
            ' name, so nothing is silently swallowed by this shortcut.
            If MilestoneDevice.SlotCount(shp) > 0 Then
                z = z + 1

                Dim dev As Candidate
                dev.Name = shp.Name
                dev.GroupPath = groupPath
                dev.ZOrder = z
                dev.ShapeType = SHAPE_TYPE_DEVICE
                dev.HasPlaceholder = False
                dev.PlaceholderType = ""
                dev.PlaceholderIdx = -1
                ' HasText is False: the group carries no text of its own, its
                ' parts do. Candidacy comes from being a device, not from text
                ' -- see Onboarding.IsCandidateField.
                dev.HasText = False
                dev.PositionX = PointsToEmu(shp.Left)
                dev.PositionY = PointsToEmu(shp.Top)
                dev.SizeCx = PointsToEmu(shp.Width)
                dev.SizeCy = PointsToEmu(shp.Height)
                dev.IdentityTag = ""

                count = count + 1
                ReDim Preserve results(1 To count)
                ReDim Preserve shapes(1 To count)
                results(count) = dev
                Set shapes(count) = shp
            Else
                Walk shp.GroupItems, childPath, results, shapes, count, z
            End If
        ElseIf IsCandidateLeafType(shp) Then
            z = z + 1

            Dim c As Candidate
            c.Name = shp.Name
            c.GroupPath = groupPath
            c.ZOrder = z
            c.ShapeType = IIf(InjectPrimitive.IsPictureShape(shp), "picture", "autoshape_or_textbox")

            c.HasPlaceholder = (shp.Type = msoPlaceholder)
            If c.HasPlaceholder Then
                c.PlaceholderType = PlaceholderTypeLabel(shp.PlaceholderFormat.Type)
            Else
                c.PlaceholderType = ""
            End If
            c.PlaceholderIdx = -1 ' not exposed by the object model -- see SPIKE_NOTES_Discovery.md

            If InjectPrimitive.IsPictureShape(shp) Then
                c.HasText = False
            Else
                c.HasText = ShapeHasNonEmptyText(shp)
            End If

            c.PositionX = PointsToEmu(shp.Left)
            c.PositionY = PointsToEmu(shp.Top)
            c.SizeCx = PointsToEmu(shp.Width)
            c.SizeCy = PointsToEmu(shp.Height)

            c.IdentityTag = ""

            count = count + 1
            ReDim Preserve results(1 To count)
            ReDim Preserve shapes(1 To count)
            results(count) = c
            Set shapes(count) = shp
        End If
        ' Anything that is neither msoGroup nor a candidate-eligible leaf type
        ' (table, chart, connector, OLE object, ...) is skipped entirely here,
        ' matching discovery.py's walk(): it only ever recognizes the OOXML
        ' tags <p:grpSp>, <p:sp>, and <p:pic> -- a <p:graphicFrame> (table/
        ' chart) or <p:cxnSp> (connector) never matches any of those three tag
        ' names, so the Python side is *already* silently blind to them, not
        ' filtering them out by a type check. This port reproduces that same
        ' real scope rather than the more idealized "type-agnostic" language
        ' in discovery.md's requirements list, which describes the has-text/
        ' is-picture check *within* the sp/pic tags Python actually looks at.
    Next shp
End Sub

Private Function IsCandidateLeafType(shp As Object) As Boolean
    Select Case shp.Type
        Case msoAutoShape, msoTextBox, msoPlaceholder, msoFreeform
            IsCandidateLeafType = True
        Case msoPicture, msoLinkedPicture, msoGraphic
            ' msoGraphic (28) is an SVG/icon-inserted shape, not a raster
            ' picture -- same gap, same fix, as InjectPrimitive.
            ' IsPictureShape's header (probed live 2026-08-19).
            IsCandidateLeafType = True
        Case Else
            IsCandidateLeafType = False
    End Select
End Function

Private Function ShapeHasNonEmptyText(shp As Object) As Boolean
    If Not shp.HasTextFrame Then
        ShapeHasNonEmptyText = False
        Exit Function
    End If
    If Not shp.TextFrame.HasText Then
        ShapeHasNonEmptyText = False
        Exit Function
    End If
    ShapeHasNonEmptyText = (Trim(shp.TextFrame.TextRange.Text) <> "")
End Function

' Best-effort mapping from PowerPoint's ppPlaceholderType enum to the
' OOXML <p:ph type="..."> string discovery.py actually captures. This is a
' many-to-one approximation, not an authoritative reverse mapping -- see
' SPIKE_NOTES_Discovery.md for why a handful of enum values (vertical
' variants, OLE-ish placeholder kinds with no clean OOXML string equivalent
' exposed here) fall through to "unknown" instead of guessing.
Private Function PlaceholderTypeLabel(phType As Long) As String
    Select Case phType
        Case 13 ' ppPlaceholderTitle
            PlaceholderTypeLabel = "title"
        Case 15 ' ppPlaceholderCenterTitle
            PlaceholderTypeLabel = "ctrTitle"
        Case 16 ' ppPlaceholderSubtitle
            PlaceholderTypeLabel = "subTitle"
        Case 2  ' ppPlaceholderBody
            PlaceholderTypeLabel = "body"
        Case 7  ' ppPlaceholderObject
            PlaceholderTypeLabel = "obj"
        Case 8  ' ppPlaceholderChart
            PlaceholderTypeLabel = "chart"
        Case 12 ' ppPlaceholderTable
            PlaceholderTypeLabel = "tbl"
        Case 10 ' ppPlaceholderMediaClip
            PlaceholderTypeLabel = "media"
        Case 9  ' ppPlaceholderBitmap
            PlaceholderTypeLabel = "clipArt"
        Case 3  ' ppPlaceholderDate
            PlaceholderTypeLabel = "dt"
        Case 4  ' ppPlaceholderSlideNumber
            PlaceholderTypeLabel = "sldNum"
        Case 5  ' ppPlaceholderHeader
            PlaceholderTypeLabel = "hdr"
        Case 6  ' ppPlaceholderFooter
            PlaceholderTypeLabel = "ftr"
        Case Else
            PlaceholderTypeLabel = "unknown"
    End Select
End Function

' 1 point = 12700 EMU (914400 EMU/inch / 72 points/inch) -- matches
' discovery.py's EMU convention so a future matching.py port (step 3) can
' compare geometry across languages without a unit-mismatch bug. Shape.Left
' etc. are Single (single-precision float) in the object model, so this
' round-trip is not always bit-exact versus Python's raw integer EMU read
' straight off <a:off>/<a:ext> -- see SPIKE_NOTES_Discovery.md.
Private Function PointsToEmu(pts As Single) As Long
    PointsToEmu = CLng(pts * 12700)
End Function

' Manual smoke-test entry point -- not a real test harness (no VBA unit-test
' framework wired up here, matching InjectPrimitive.bas's ManualSmokeTest
' pattern). Run from the VBA IDE (F5) against the active presentation's
' first slide; prints one line per discovered candidate to the Immediate
' window. See SPIKE_NOTES_Discovery.md for the recipe to run this against
' the project's actual test-fixtures/*.pptx files and what to expect.
Public Sub ManualSmokeTest()
    Dim sld As Object
    Set sld = Application.ActivePresentation.Slides(1)

    Dim candidates() As Candidate
    candidates = DiscoverSlide(sld)

    ' candidates may be genuinely unallocated (zero candidates found) --
    ' UBound/LBound on it would throw, not return 0/1, so this must be
    ' error-guarded rather than assumed allocated. See AGENTS.md's Known
    ' Patterns on the (1 To 0) restriction this project ran into.
    Dim lo As Long, hi As Long
    On Error Resume Next
    lo = LBound(candidates)
    hi = UBound(candidates)
    If Err.Number <> 0 Then
        Debug.Print "DiscoverSlide found 0 candidate(s)."
        MsgBox "DiscoverSlide found 0 candidate(s)."
        Exit Sub
    End If
    On Error GoTo 0

    Dim i As Long
    Debug.Print "DiscoverSlide found " & (hi - lo + 1) & " candidate(s):"
    For i = lo To hi
        With candidates(i)
            Debug.Print "  z=" & .ZOrder & " name=" & .Name & " group=" & .GroupPath & _
                " type=" & .ShapeType & " hasPlaceholder=" & .HasPlaceholder & _
                " placeholderType=" & .PlaceholderType & " hasText=" & .HasText & _
                " pos=(" & .PositionX & "," & .PositionY & ") size=(" & .SizeCx & "," & .SizeCy & ")"
        End With
    Next i

    MsgBox "See Immediate window (Ctrl+G) for " & (UBound(candidates) - LBound(candidates) + 1) & " candidate(s)."
End Sub
