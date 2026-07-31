Attribute VB_Name = "TagCloningProbe"
Option Explicit

' One-off empirical probe -- NOT part of the shipped add-in.
'
' Answers Q6 of the Excel Control Layer confirmation (31 July 2026), which
' correctly flagged that D5's uniqueness requirement was REASONED from the
' DeckSyncId finding rather than measured. This measures it.
'
' Four questions, none answered from memory:
'   1. Do SLIDE-level tags survive Slide.Duplicate?
'   2. Do SHAPE-level role tags survive it?
'   3. Does the duplicate get a NEW SlideID -- i.e. is there a PowerPoint-native
'      discriminator between a copy and its original within one deck?
'   4. Do tags survive Copy/Paste into a DIFFERENT presentation?
'
' Q3 is the one that matters most for D5's scope: if SlideID is genuinely
' reassigned on duplicate, then within-deck duplication is natively
' detectable and the uniqueness check has help. If it is preserved, the
' check is the whole defence.
Public Function TagCloningProbe() As String
    Dim r As String

    Dim pres As Object
    Set pres = Application.ActivePresentation

    ' --- Build a tagged source slide -------------------------------------
    Dim src As Object
    Set src = pres.Slides.Add(pres.Slides.count + 1, 12)   ' ppLayoutBlank
    src.Tags.Add "slide_type", "probe-type"
    src.Tags.Add "instance_key", "PROBE-001"

    Dim shp As Object
    Set shp = src.Shapes.AddTextbox(1, 50, 50, 300, 50)    ' msoTextOrientationHorizontal
    shp.TextFrame.TextRange.Text = "probe field"
    shp.Tags.Add "role", "PROGRESS_BODY"

    Dim srcId As Long
    srcId = src.SlideID

    r = "=== SOURCE ===" & vbCrLf & _
        "  SlideID:      " & srcId & vbCrLf & _
        "  slide_type:   '" & src.Tags("slide_type") & "'" & vbCrLf & _
        "  instance_key: '" & src.Tags("instance_key") & "'" & vbCrLf & _
        "  shape role:   '" & shp.Tags("role") & "'" & vbCrLf & vbCrLf

    ' --- Q1/Q2/Q3: Slide.Duplicate within the same presentation -----------
    Dim dupColl As Object
    Set dupColl = src.Duplicate()
    Dim dup As Object
    Set dup = dupColl(1)

    Dim dupShapeRole As String
    dupShapeRole = "<no shape>"
    If dup.Shapes.count >= 1 Then dupShapeRole = dup.Shapes(1).Tags("role")

    r = r & "=== AFTER Slide.Duplicate (same presentation) ===" & vbCrLf & _
        "  SlideID:      " & dup.SlideID & "   (source was " & srcId & ")" & vbCrLf & _
        "  SlideID differs from source: " & (dup.SlideID <> srcId) & vbCrLf & _
        "  slide_type:   '" & dup.Tags("slide_type") & "'" & vbCrLf & _
        "  instance_key: '" & dup.Tags("instance_key") & "'" & vbCrLf & _
        "  shape role:   '" & dupShapeRole & "'" & vbCrLf & vbCrLf

    ' --- Q4: Copy/Paste into a DIFFERENT presentation ---------------------
    ' The single-slide analogue of starting a new deck from an old one.
    Dim other As Object
    Set other = Application.Presentations.Add(msoFalse)
    src.Copy
    Dim pasted As Object
    Set pasted = other.Slides.Paste(1)(1)

    Dim pastedShapeRole As String
    pastedShapeRole = "<no shape>"
    If pasted.Shapes.count >= 1 Then pastedShapeRole = pasted.Shapes(1).Tags("role")

    r = r & "=== AFTER Copy/Paste into a NEW presentation ===" & vbCrLf & _
        "  SlideID:      " & pasted.SlideID & "   (source was " & srcId & ")" & vbCrLf & _
        "  slide_type:   '" & pasted.Tags("slide_type") & "'" & vbCrLf & _
        "  instance_key: '" & pasted.Tags("instance_key") & "'" & vbCrLf & _
        "  shape role:   '" & pastedShapeRole & "'" & vbCrLf & vbCrLf

    ' --- Verdicts, stated rather than left to be inferred ------------------
    r = r & "=== VERDICTS ===" & vbCrLf & _
        "  slide tags clone on Duplicate:        " & (dup.Tags("instance_key") = "PROBE-001") & vbCrLf & _
        "  shape role tags clone on Duplicate:   " & (dupShapeRole = "PROGRESS_BODY") & vbCrLf & _
        "  slide tags clone across presentations:" & (pasted.Tags("instance_key") = "PROBE-001") & vbCrLf & _
        "  SlideID reassigned on Duplicate:      " & (dup.SlideID <> srcId) & vbCrLf

    other.Saved = msoTrue
    other.Close
    dup.Delete
    src.Delete

    TagCloningProbe = r
End Function
