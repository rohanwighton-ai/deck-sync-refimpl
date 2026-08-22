' Adds a real MS<n>_CALDATE textbox to every milestone slot on the given
' slide(s), positioned directly below that slot's own DATE label, measured
' from the DATE label's ACTUAL geometry on THIS slide (never a hardcoded
' absolute position) -- same "measured, not computed" convention
' MilestoneDevice.bas's bar/track already use, and the same technique
' add_pct_shapes.vbs proved (2026-08-22, same session). Positioned under
' the DATE label specifically, not the circle -- Rohan: "date could be
' added instead under each circle like we were going to with % deliverable
' complete", and that design was "as a smaller font separate label under
' the date label" (MILESTONE-PERCENTAGE-DESIGN.md). Hidden by default;
' DrawMilestones shows/writes it when a register MS<n>_CALDATE value
' exists -- NOT gated by achieved state (see COL_CALDATE's own header).
'
' Regrouping the device to add the new shape destroys the group's own name
' and role tag (confirmed empirically 2026-08-22, reproduced again here) --
' both are captured before and explicitly restored after.
'
' Usage: cscript //Nologo add_caldate_shapes.vbs <slideIndex>
' Run once against all real milestone-carrying slides + templates,
' 2026-08-22. Kept for the next time a slide needs the same retrofit.

Dim ppt
Set ppt = CreateObject("PowerPoint.Application")
ppt.Visible = True

Dim deckPath
deckPath = "C:\Users\rohan\OneDrive\Claude\3. Project Progress.pptx"
Dim pres
Set pres = ppt.Presentations.Open(deckPath, False, False, True)

' Checked by NAME, not the full path. Presentations.Open was called with
' the exact local path above, so PowerPoint cannot have opened a
' different file -- but a fully OneDrive-synced file's own FullName
' sometimes reports as the cloud https://d.docs.live.net/... alias
' instead of the local path afterward (confirmed live 2026-08-22: every
' one of 36 slides hit this and aborted safely, having written nothing).
' That is not a wrong-file risk, it is OneDrive's own path virtualization
' -- Open() itself is what guarantees the right file, this check exists
' only to catch a genuinely different file being opened by mistake.
If pres.Name <> "3. Project Progress.pptx" Then
    WScript.Echo "WARNING -- opened file does not match, aborting: " & pres.Name
    pres.Close()
    ppt.Quit()
    WScript.Quit 1
End If

Dim GAP_PT, HEIGHT_PT, FONT_SIZE
GAP_PT = 0.015 * 72
HEIGHT_PT = 0.10 * 72
FONT_SIZE = 5.5

' ONE SLIDE PER INVOCATION, DELIBERATELY -- see add_pct_shapes.vbs's own
' header for the "chained regroup -> Type mismatch" reason this matters.
Dim slideList
If WScript.Arguments.Count > 0 Then
    slideList = Array(WScript.Arguments(0))
Else
    slideList = Array("44")
End If

Dim totalAdded, totalSlides
totalAdded = 0
totalSlides = 0

Dim s
For Each s In slideList
    Dim idx
    idx = CInt(s)
    Dim sld
    Set sld = pres.Slides.Item(idx)

    Dim deviceGroup, shp, inner
    Set deviceGroup = Nothing
    For Each shp In sld.Shapes
        If shp.Type = 6 Then ' msoGroup
            For Each inner In shp.GroupItems
                If inner.Name = "MS_BAR" Then
                    Set deviceGroup = shp
                    Exit For
                End If
            Next
        End If
        If Not deviceGroup Is Nothing Then Exit For
    Next

    If deviceGroup Is Nothing Then
        WScript.Echo "Slide " & idx & " : no milestone device group found -- skipped"
    Else
        Dim origName, tagNames(), tagValues(), tagCount, i
        origName = deviceGroup.Name
        tagCount = deviceGroup.Tags.Count
        ReDim tagNames(tagCount - 1)
        ReDim tagValues(tagCount - 1)
        For i = 1 To tagCount
            tagNames(i - 1) = deviceGroup.Tags.Name(i)
            tagValues(i - 1) = deviceGroup.Tags.Value(i)
        Next

        Dim addedThisSlide, slot
        addedThisSlide = 0

        For slot = 1 To 7
            Dim dateName, dateShp, found
            dateName = "MS" & slot & "_DATE"
            Set dateShp = Nothing
            found = False
            For Each inner In deviceGroup.GroupItems
                If inner.Name = dateName Then
                    Set dateShp = inner
                    found = True
                    Exit For
                End If
            Next
            If found Then
                Dim caldName, already
                caldName = "MS" & slot & "_CALDATE"
                already = False
                For Each inner In deviceGroup.GroupItems
                    If inner.Name = caldName Then already = True
                Next

                If already Then
                    WScript.Echo "Slide " & idx & " slot " & slot & " : " & caldName & " already exists -- skipped"
                Else
                    Dim leftPt, widthPt, topPt
                    leftPt = dateShp.Left
                    widthPt = dateShp.Width
                    topPt = dateShp.Top + dateShp.Height + GAP_PT

                    Dim newShape
                    Set newShape = sld.Shapes.AddTextbox(1, leftPt, topPt, widthPt, HEIGHT_PT)
                    newShape.Name = caldName
                    ' AutoSize MUST be turned off before anything else -- see
                    ' add_pct_shapes.vbs's own header for the width=0 bug
                    ' this avoids.
                    newShape.TextFrame.AutoSize = 0 ' ppAutoSizeNone
                    newShape.TextFrame.MarginLeft = 0
                    newShape.TextFrame.MarginRight = 0
                    newShape.TextFrame.MarginTop = 0
                    newShape.TextFrame.MarginBottom = 0
                    newShape.TextFrame.TextRange.Text = ""
                    newShape.TextFrame.TextRange.Font.Size = FONT_SIZE
                    newShape.TextFrame.TextRange.Font.Bold = False
                    newShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2 ' ppAlignCenter
                    ' Re-assert geometry explicitly, after every text/font
                    ' property is set -- the one proven place autofit can
                    ' still silently override it.
                    newShape.Left = leftPt
                    newShape.Top = topPt
                    newShape.Width = widthPt
                    newShape.Height = HEIGHT_PT
                    newShape.Visible = False

                    Set deviceGroup = sld.Shapes.Range(Array(deviceGroup.Name, newShape.Name)).Group

                    deviceGroup.Name = origName
                    For i = 0 To tagCount - 1
                        deviceGroup.Tags.Add tagNames(i), tagValues(i)
                    Next

                    addedThisSlide = addedThisSlide + 1
                    totalAdded = totalAdded + 1
                End If
            End If
        Next

        WScript.Echo "Slide " & idx & " : " & addedThisSlide & " MS<n>_CALDATE shape(s) added"
        totalSlides = totalSlides + 1
    End If
Next

WScript.Echo "Total: " & totalAdded & " shapes across " & totalSlides & " slide(s)"

pres.Save()
pres.Close()
ppt.Quit()
WScript.Echo "Saved."
