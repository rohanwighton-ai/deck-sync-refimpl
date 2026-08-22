' Adds a real MS<n>_PCT textbox to every milestone slot on the given
' slide(s), positioned directly below that slot's own ON circle, measured
' from the circle's ACTUAL geometry on THIS slide (never a hardcoded
' absolute position) -- matches MilestoneDevice.bas's own "measured, not
' computed" convention for the bar/track. Hidden by default; DrawMilestones
' shows/writes it when a register MS<n>_PCT value exists.
'
' Regrouping the device to add the new shape destroys the group's own name
' and role tag (confirmed empirically, scratch copy, 2026-08-22) -- both
' are captured before and explicitly restored after.
'
' Usage: cscript //Nologo add_pct_shapes.vbs <slideIndex>
' Already run against all 46 milestone-carrying real slides + templates
' (1-44, 46, 47 -- 45 has no milestone device) on 2026-08-22. Kept for the
' next time a slide is added that needs the same shape retrofitting.

Dim ppt
Set ppt = CreateObject("PowerPoint.Application")
ppt.Visible = True

Dim deckPath
deckPath = "C:\Users\rohan\AppData\Local\deck-sync-quarter-20260820-0900\3. Project Progress.pptx"
Dim pres
Set pres = ppt.Presentations.Open(deckPath, False, False, True)

If pres.FullName <> deckPath Then
    WScript.Echo "WARNING -- opened path does not match, aborting: " & pres.FullName
    pres.Close()
    ppt.Quit()
    WScript.Quit 1
End If

Dim GAP_PT, HEIGHT_PT, FONT_SIZE
GAP_PT = 0.015 * 72
HEIGHT_PT = 0.10 * 72
FONT_SIZE = 5.5

' ONE SLIDE PER INVOCATION, DELIBERATELY. Chaining multiple slides' worth
' of regroup operations in one PowerPoint session produced a "Type
' mismatch" on the second slide's device-group lookup, reproducibly, even
' though the identical logic against that same slide in total isolation
' worked cleanly -- a cumulative COM state issue from the regrouping, not
' a per-slide content problem. Never root-caused further; one launch per
' slide (driven by a bash loop) sidesteps it entirely and matches the
' fresh-instance pattern already proven reliable elsewhere tonight.
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

    ' Find the device group (the one carrying MS_BAR).
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
        ' Capture name + every tag before any regroup touches it.
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
            Dim circleName, circle, found
            circleName = "MS" & slot & "_ON"
            Set circle = Nothing
            found = False
            For Each inner In deviceGroup.GroupItems
                If inner.Name = circleName Then
                    Set circle = inner
                    found = True
                    Exit For
                End If
            Next
            If found Then
                Dim pctName, already
                pctName = "MS" & slot & "_PCT"
                already = False
                For Each inner In deviceGroup.GroupItems
                    If inner.Name = pctName Then already = True
                Next

                If already Then
                    WScript.Echo "Slide " & idx & " slot " & slot & " : " & pctName & " already exists -- skipped"
                Else
                    Dim leftPt, widthPt, topPt
                    leftPt = circle.Left
                    widthPt = circle.Width
                    topPt = circle.Top + circle.Height + GAP_PT

                    Dim newShape
                    Set newShape = sld.Shapes.AddTextbox(1, leftPt, topPt, widthPt, HEIGHT_PT)
                    newShape.Name = pctName
                    ' AutoSize MUST be turned off before anything else -- an
                    ' empty textbox with WordWrap=False otherwise auto-fits
                    ' to its (empty) content and collapses to ~0 width.
                    ' Found live: first attempt shipped every PCT shape at
                    ' cx="65" EMU (effectively invisible) for exactly this
                    ' reason, caught by reading the saved file's own XML
                    ' before trusting the "7 shapes added" success message.
                    newShape.TextFrame.AutoSize = 0 ' ppAutoSizeNone
                    newShape.TextFrame.MarginLeft = 0
                    newShape.TextFrame.MarginRight = 0
                    newShape.TextFrame.MarginTop = 0
                    newShape.TextFrame.MarginBottom = 0
                    newShape.TextFrame.TextRange.Text = ""
                    newShape.TextFrame.TextRange.Font.Size = FONT_SIZE
                    newShape.TextFrame.TextRange.Font.Bold = True
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

                    ' Restore identity immediately -- regrouping just wiped it.
                    deviceGroup.Name = origName
                    For i = 0 To tagCount - 1
                        deviceGroup.Tags.Add tagNames(i), tagValues(i)
                    Next

                    addedThisSlide = addedThisSlide + 1
                    totalAdded = totalAdded + 1
                End If
            End If
        Next

        WScript.Echo "Slide " & idx & " : " & addedThisSlide & " MS<n>_PCT shape(s) added"
        totalSlides = totalSlides + 1
    End If
Next

WScript.Echo "Total: " & totalAdded & " shapes across " & totalSlides & " slide(s)"

pres.Save()
pres.Close()
ppt.Quit()
WScript.Echo "Saved."
