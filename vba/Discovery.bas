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

Public Type Candidate
    Name As String              ' shape name, for humans -- never used as an identity key (see identity_tags.py)
    GroupPath As String         ' "/"-joined chain of enclosing group names, "" if top-level
    ZOrder As Long              ' 1-based document-order position among discovered leaf shapes
    ShapeType As String         ' "autoshape_or_textbox" | "picture"
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

' Walk `sld`'s shape tree and return every candidate leaf shape (autoshape,
' textbox, placeholder, or picture), recursing into groups. Mirrors
' discovery.py's discover(): type-agnostic, never treats a group as one
' opaque candidate, tags leaves not containers.
Public Function DiscoverSlide(sld As Object) As Candidate()
    Dim results() As Candidate
    ReDim results(1 To 0)
    Dim count As Long
    count = 0
    Dim z As Long
    z = 0

    Walk sld.Shapes, "", results, count, z

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
    Dim results() As Candidate
    ReDim results(1 To 0)
    Dim count As Long
    count = 0
    Dim z As Long
    z = 0

    Walk layout.Shapes, "", results, count, z

    DiscoverCustomLayout = results
End Function

Private Sub Walk(shapesColl As Object, groupPath As String, ByRef results() As Candidate, ByRef count As Long, ByRef z As Long)
    Dim shp As Object
    For Each shp In shapesColl
        If shp.Type = msoGroup Then
            Dim childPath As String
            If groupPath = "" Then
                childPath = shp.Name
            Else
                childPath = groupPath & "/" & shp.Name
            End If
            Walk shp.GroupItems, childPath, results, count, z
        ElseIf IsCandidateLeafType(shp) Then
            z = z + 1

            Dim c As Candidate
            c.Name = shp.Name
            c.GroupPath = groupPath
            c.ZOrder = z
            c.ShapeType = IIf(IsPicture(shp), "picture", "autoshape_or_textbox")

            c.HasPlaceholder = (shp.Type = msoPlaceholder)
            If c.HasPlaceholder Then
                c.PlaceholderType = PlaceholderTypeLabel(shp.PlaceholderFormat.Type)
            Else
                c.PlaceholderType = ""
            End If
            c.PlaceholderIdx = -1 ' not exposed by the object model -- see SPIKE_NOTES_Discovery.md

            If IsPicture(shp) Then
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
            results(count) = c
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

Private Function IsPicture(shp As Object) As Boolean
    IsPicture = (shp.Type = msoPicture) Or (shp.Type = msoLinkedPicture)
End Function

Private Function IsCandidateLeafType(shp As Object) As Boolean
    Select Case shp.Type
        Case msoAutoShape, msoTextBox, msoPlaceholder, msoFreeform
            IsCandidateLeafType = True
        Case msoPicture, msoLinkedPicture
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

    Dim i As Long
    Debug.Print "DiscoverSlide found " & (UBound(candidates) - LBound(candidates) + 1) & " candidate(s):"
    For i = LBound(candidates) To UBound(candidates)
        With candidates(i)
            Debug.Print "  z=" & .ZOrder & " name=" & .Name & " group=" & .GroupPath & _
                " type=" & .ShapeType & " hasPlaceholder=" & .HasPlaceholder & _
                " placeholderType=" & .PlaceholderType & " hasText=" & .HasText & _
                " pos=(" & .PositionX & "," & .PositionY & ") size=(" & .SizeCx & "," & .SizeCy & ")"
        End With
    Next i

    MsgBox "See Immediate window (Ctrl+G) for " & (UBound(candidates) - LBound(candidates) + 1) & " candidate(s)."
End Sub
