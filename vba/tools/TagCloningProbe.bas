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

    ' A PROGRESS field is not a lone textbox. It is a PAIR -- FIELD and
    ' FIELD.track -- and a template that carries the pair is likely to carry it
    ' GROUPED, which is the case InjectPrimitive's WalkForRoleTag recurses for.
    ' So the probe has to ask about a tag on a shape INSIDE a group, on a
    ' SUFFIXED role value, alongside a SECOND tag name on the same shape (the
    ' picture stamp). A lone top-level shape would answer an easier question
    ' than the one the design turns on.
    Dim trackShp As Object, doneShp As Object
    Set trackShp = src.Shapes.AddShape(1, 50, 200, 400, 20)  ' msoShapeRectangle
    trackShp.Tags.Add "role", "BAR_BODY.track"
    Set doneShp = src.Shapes.AddShape(1, 50, 200, 160, 20)
    doneShp.Tags.Add "role", "BAR_BODY"
    doneShp.Tags.Add InjectPrimitive.PICTURE_SOURCE_TAG, "S07"
    src.Shapes.Range(Array(trackShp.Name, doneShp.Name)).Group

    Dim srcId As Long
    srcId = src.SlideID

    r = "=== SOURCE ===" & vbCrLf & _
        "  SlideID:      " & srcId & vbCrLf & _
        "  slide_type:   '" & src.Tags("slide_type") & "'" & vbCrLf & _
        "  instance_key: '" & src.Tags("instance_key") & "'" & vbCrLf & _
        "  shape role:   '" & shp.Tags("role") & "'" & vbCrLf & _
        "  shapes:" & vbCrLf & DumpTags(src.Shapes) & vbCrLf

    ' --- Q1/Q2/Q3: Slide.Duplicate within the same presentation -----------
    Dim dupColl As Object
    Set dupColl = src.Duplicate()
    Dim dup As Object
    Set dup = dupColl(1)

    ' Found BY TAG, not by position. Reading dup.Shapes(1) would answer for
    ' whichever shape happens to come first and could report a pass off the
    ' wrong shape entirely.
    Dim dupShapeRole As String
    dupShapeRole = "<not found>"
    If Not FindByRole(dup.Shapes, "PROGRESS_BODY") Is Nothing Then dupShapeRole = "PROGRESS_BODY"

    Dim dupTrack As Object, dupDone As Object
    Set dupTrack = FindByRole(dup.Shapes, "BAR_BODY.track")
    Set dupDone = FindByRole(dup.Shapes, "BAR_BODY")

    r = r & "=== AFTER Slide.Duplicate (same presentation) ===" & vbCrLf & _
        "  SlideID:      " & dup.SlideID & "   (source was " & srcId & ")" & vbCrLf & _
        "  SlideID differs from source: " & (dup.SlideID <> srcId) & vbCrLf & _
        "  slide_type:   '" & dup.Tags("slide_type") & "'" & vbCrLf & _
        "  instance_key: '" & dup.Tags("instance_key") & "'" & vbCrLf & _
        "  shape role:   '" & dupShapeRole & "'" & vbCrLf & _
        "  shapes:" & vbCrLf & DumpTags(dup.Shapes) & vbCrLf

    ' --- Q4: Copy/Paste into a DIFFERENT presentation ---------------------
    ' The single-slide analogue of starting a new deck from an old one.
    Dim other As Object
    Set other = Application.Presentations.Add(msoFalse)
    src.Copy
    Dim pasted As Object
    Set pasted = other.Slides.Paste(1)(1)

    Dim pastedShapeRole As String
    pastedShapeRole = "<not found>"
    If Not FindByRole(pasted.Shapes, "PROGRESS_BODY") Is Nothing Then pastedShapeRole = "PROGRESS_BODY"

    r = r & "=== AFTER Copy/Paste into a NEW presentation ===" & vbCrLf & _
        "  SlideID:      " & pasted.SlideID & "   (source was " & srcId & ")" & vbCrLf & _
        "  slide_type:   '" & pasted.Tags("slide_type") & "'" & vbCrLf & _
        "  instance_key: '" & pasted.Tags("instance_key") & "'" & vbCrLf & _
        "  shape role:   '" & pastedShapeRole & "'" & vbCrLf & vbCrLf

    ' --- Verdicts, stated rather than left to be inferred ------------------
    r = r & "=== VERDICTS ===" & vbCrLf & _
        "  slide tags clone on Duplicate:          " & (dup.Tags("instance_key") = "PROBE-001") & vbCrLf & _
        "  shape role tag clones (top level):      " & (dupShapeRole = "PROGRESS_BODY") & vbCrLf & _
        "  SUFFIXED .track tag clones, IN A GROUP: " & (Not dupTrack Is Nothing) & vbCrLf & _
        "  its pair (BAR_BODY) clones too:         " & (Not dupDone Is Nothing) & vbCrLf & _
        "  picture stamp '" & InjectPrimitive.PICTURE_SOURCE_TAG & "' clones:          " & _
            (IIf(dupDone Is Nothing, "n/a", CStr(dupDone.Tags(InjectPrimitive.PICTURE_SOURCE_TAG) = "S07"))) & vbCrLf & _
        "  slide tags clone across presentations:  " & (pasted.Tags("instance_key") = "PROBE-001") & vbCrLf & _
        "  SlideID reassigned on Duplicate:        " & (dup.SlideID <> srcId) & vbCrLf & _
        vbCrLf & _
        "  CONTROL -- these MUST read False, or nothing above means anything:" & vbCrLf & _
        "  a role never tagged is found:           " & _
            (Not FindByRole(dup.Shapes, "NEVER_TAGGED_ANYWHERE") Is Nothing) & vbCrLf & _
        "  the track shape carries a picture stamp:" & _
            IIf(dupTrack Is Nothing, " n/a", " " & CStr(dupTrack.Tags(InjectPrimitive.PICTURE_SOURCE_TAG) <> "")) & vbCrLf

    other.Saved = msoTrue
    other.Close
    dup.Delete
    src.Delete

    TagCloningProbe = r
End Function

' Every shape on the slide, groups walked into, with the two tag names this
' project reads. Printed in full alongside the verdicts so a False verdict
' shows WHAT was there instead of only that the wanted thing was not.
' `inside` names the enclosing group, or is empty at the top level. Without it
' a grouped shape and a shape that was never grouped print the same line, so
' the dump could not tell a working group test from a grouping that silently
' did not happen -- which is the whole thing this probe is asking about.
Private Function DumpTags(shapesColl As Object, Optional inside As String = "") As String
    Dim shp As Object, s As String
    For Each shp In shapesColl
        If shp.Type = msoGroup Then
            s = s & "    [" & shp.Name & "] GROUP of " & shp.GroupItems.count & vbCrLf
            s = s & DumpTags(shp.GroupItems, shp.Name)
        Else
            s = s & "    " & IIf(inside = "", "", "  in " & inside & " -> ") & _
                "[" & shp.Name & "] role='" & shp.Tags("role") & _
                "' " & InjectPrimitive.PICTURE_SOURCE_TAG & "='" & _
                shp.Tags(InjectPrimitive.PICTURE_SOURCE_TAG) & "'" & vbCrLf
        End If
    Next
    DumpTags = s
End Function

' Same group recursion as InjectPrimitive.WalkForRoleTag, which is the code
' that will actually have to find these shapes at sync time.
Private Function FindByRole(shapesColl As Object, want As String) As Object
    Dim shp As Object
    For Each shp In shapesColl
        If shp.Type = msoGroup Then
            Dim inner As Object
            Set inner = FindByRole(shp.GroupItems, want)
            If Not inner Is Nothing Then
                Set FindByRole = inner
                Exit Function
            End If
        ElseIf shp.Tags("role") = want And want <> "" Then
            Set FindByRole = shp
            Exit Function
        End If
    Next
End Function
