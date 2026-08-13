Attribute VB_Name = "MilestoneDevice"
Option Explicit

' A TIMELINE IS ONE DEVICE, NOT FIFTEEN FIELDS.
'
' Rohan, 2026-08-10, after an hour spent tagging a bar made of TWO shapes:
' "this is all too much, we need to simplify the approach for slides with these
' more complex devices" -- and then, decisively: "I'd like it all data drawn
' given I take a robotic approach when updating it anyway."
'
' His timeline is a bar, a track, six or seven circles and their labels. Tagging
' that shape-by-shape through the marking dialogs would be an evening per
' template and a dozen chances to answer wrong. The per-shape model is right for
' "this text box holds ABOUT_BODY" and nonsense for a device that is one thing
' in a person's head and fifteen objects in PowerPoint.
'
' ---------------------------------------------------------------------
' PARTS ARE FOUND BY SHAPE NAME, NOT BY TAG
' ---------------------------------------------------------------------
'
' Deliberate reversal of the convention everywhere else in this project, for a
' reason that only applies at this size: A TAG IS INVISIBLE. Nobody can see it,
' check it, or repair it without the tool -- fine for one shape, disqualifying
' for fifteen. PowerPoint's Selection Pane lists NAMES, so a person can see the
' whole structure at a glance, spot the one called wrong, and fix it themselves.
'
' The cost is that names are not unique and not enforced. That is why the whole
' device sits inside ONE tagged group: the tag says "this group is the milestone
' timeline", and names only have to be unique WITHIN it.
'
' ---------------------------------------------------------------------
' THE DATA IS THE CONTROL POINT
' ---------------------------------------------------------------------
'
' Rohan's word: these variables are the CCPs. Everything drawn is derived from
' the label / date / achieved triple, so wrong data produces a slide that is
' confidently wrong -- there is no second signal to notice it. Hence the
' refusals below are strict: three lists of different lengths cannot be aligned
' by guessing, so nothing is drawn at all.
'
' ---------------------------------------------------------------------
' WHAT IT NEVER DOES
' ---------------------------------------------------------------------
'
' Never creates, deletes, moves or reorders a shape. Rohan: "you can see the
' extremely accurate positioning? we need to maintain that and z order etc."
' So: visibility toggles, text writes, and TWO height changes -- the bar's and
' the track's. Both are heights, never positions: a height cannot move a shape
' sideways or change its z-order, which is why a computed one is acceptable
' here when a computed position would not be. (This said ONE until 2026-08-13,
' when a shorter milestone chain was found to leave track hanging below the
' last circle, pointing at nothing.) The
' template pre-places every slot; unused ones are hidden, not removed. Circle
' positions are READ to work out how far the bar should reach -- measured off
' the slide the template already laid out, never computed.

' Slot part names, as they appear in the Selection Pane. `n` is 1-based.
'   MS1_NOW    the circle shown when this is the CURRENT milestone (optional)
'   MS1_ON     the circle shown when it is achieved but not current
'   MS1_OFF    the circle shown when it is NOT achieved          (optional)
'   MS1_LABEL  its text
'   MS1_DATE   its date
' and for the bar itself:
'   MS_BAR     the part that grows, to the last ACHIEVED circle
'   MS_TRACK   the extent, shortened to the last USED circle
'
' Exactly one circle per used slot is ever visible. Absent optional circles
' degrade to the next best one the template carries, and the degradation is
' reported rather than faked by recolouring.
Public Const SLOT_PREFIX As String = "MS"
Public Const PART_ON As String = "_ON"
Public Const PART_OFF As String = "_OFF"
' The CURRENT position -- the last achieved milestone, drawn larger. Optional,
' like _OFF: a template without it simply shows _ON there and the difference
' goes unmarked, which is reported rather than faked.
'
' Rohan, 2026-08-13: four states, not two. Achieved-now is big; achieved-earlier
' and not-achieved are small and differ by colour; an unused slot shows nothing.
' His real slides already carry the big circle -- slot 3 of 3_P001 is 0.43"
' against everyone else's 0.35" -- so this formalises what he had drawn rather
' than introducing it.
Public Const PART_NOW As String = "_NOW"
Public Const PART_LABEL As String = "_LABEL"
Public Const PART_DATE As String = "_DATE"
Public Const NAME_BAR As String = "MS_BAR"
Public Const NAME_TRACK As String = "MS_TRACK"

Public Type MilestoneDrawResult
    Drawn As Long           ' slots given a milestone
    Hidden As Long          ' slots hidden for want of one
    SlotsFound As Long      ' slots the TEMPLATE carries
    BarSet As Boolean
    ErrorMessage As String
    Detail As String
End Type

' A COLUMN PER PART, NOT A PACKED CELL.
'
' The first version put every milestone in one cell as
' `Kickoff~Oct 2023~Y||Design~Mar 2024~Y`. Rohan: "with one row per project
' wouldn't the timeline be one column per interval not one row?" He was right,
' and it killed the packed format the same hour it was written.
'
' A packed cell is the spreadsheet equivalent of an invisible tag -- the exact
' thing this module rejects tags FOR. You cannot sort it, filter it, validate
' it, or change one date without string surgery in the formula bar.
'
' Columns also destroy the alignment hazard properly rather than guarding it:
' MS1_LABEL, MS1_DATE and MS1_DONE are aligned by COLUMN NAME within one row, so
' there is no position to get out of step. The parallel-list problem was
' self-inflicted by splitting one thing into three lists.
'
' AND THESE ARE FIELDS, NOT A NEW KIND OF THING. Rohan again: "I'm not sure I
' understand the difference between that and a field." There isn't one. A field
' is a register column whose value lands in a shape; MS1_LABEL is a register
' column whose value lands in a shape. The only difference is how the shape is
' ADDRESSED -- by role tag, or by name inside a tagged device group. Inventing a
' second category would have switched the coverage, case and empty-value checks
' off for two dozen columns and called it tidiness.
Public Const COL_LABEL As String = "_LABEL"
Public Const COL_DATE As String = "_DATE"
Public Const COL_DONE As String = "_DONE"

' The register column names this device reads, for slot `i`. Built here so the
' column names and the SHAPE names cannot drift apart -- they are deliberately
' the same strings, because a person looking at either should recognise the
' other.
Public Function ColumnFor(i As Long, part As String) As String
    ColumnFor = SLOT_PREFIX & i & part
End Function

' Tolerant of how a person writes "yes" in a spreadsheet. Blank means NOT done:
' a milestone nobody has marked achieved has not been achieved.
Public Function IsDoneWord(v As String) As Boolean
    Dim u As String
    u = UCase(Trim(v))
    IsDoneWord = (u = "Y" Or u = "YES" Or u = "TRUE" Or u = "1" Or u = "DONE" Or u = "ACHIEVED")
End Function

' Every shape inside the device group, by name. Groups walked into, because a
' slot's parts may themselves be grouped.
Private Function PartsOf(grp As Object) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    CollectNamed grp.GroupItems, d
    Set PartsOf = d
End Function

Private Sub CollectNamed(shapesColl As Object, ByRef d As Object)
    Dim shp As Object
    For Each shp In shapesColl
        If shp.Type = msoGroup Then
            CollectNamed shp.GroupItems, d
        End If
        ' The GROUP ITSELF is recorded too -- a slot may be a group whose name
        ' is the slot part, with its own members inside.
        Dim nm As String
        nm = UCase(Trim(shp.Name))
        If nm <> "" Then
            If Not d.Exists(nm) Then d.Add nm, shp
        End If
    Next shp
End Sub

' HOW MANY SLOTS THE TEMPLATE CARRIES, counted rather than configured.
'
' Asking a person for "the maximum number of milestones" invites an answer that
' is right today and wrong next quarter, held in a constant nobody revisits.
' The template already knows: it has as many slots as someone drew.
Public Function SlotCount(grp As Object) As Long
    If grp Is Nothing Then Exit Function
    Dim parts As Object
    Set parts = PartsOf(grp)

    Dim n As Long
    n = 0
    ' A SLOT EXISTS IF IT CARRIES ANY OF ITS THREE CIRCLES.
    '
    ' Counting only _ON would stop at the first slot drawn with _NOW alone --
    ' which is exactly how Rohan's real slides are drawn today, where slot 3
    ' carries the big current circle and nothing else. The count would have
    ' come back 2 on a seven-slot timeline and the rest would have been
    ' reported as "more milestones than slots" and refused.
    Do While parts.Exists(SLOT_PREFIX & (n + 1) & PART_ON) _
          Or parts.Exists(SLOT_PREFIX & (n + 1) & PART_NOW) _
          Or parts.Exists(SLOT_PREFIX & (n + 1) & PART_OFF)
        n = n + 1
    Loop
    SlotCount = n
End Function

' IS THIS DEVICE INTACT? -- the check that makes renaming safe.
'
' Parts are addressed by NAME, which is the whole reason a person can set this
' up in the Selection Pane in fifteen minutes instead of thirty marking
' sessions. The cost is that nothing protects a name: rename MS2_DATE and that
' milestone silently stops updating.
'
' The answer is NOT a second copy of the fact as a tag -- two copies of one fact
' drift, and then the Selection Pane lies. The answer is to check the set is
' complete and say what is missing, which catches a rename, a typo and a deleted
' shape with one mechanism.
'
' Rohan, 2026-08-10: "doesn't each shape in the group have its own tag if it
' contains variable info?" They are all fields. They are simply addressed by
' name, and this is what stands in for the robustness a tag would have given.
'
' Reports SLOT BY SLOT. "3 parts missing" would send someone hunting; naming
' MS2_DATE sends them to the shape.
Public Function DeviceIntegrity(grp As Object) As String
    If grp Is Nothing Then
        DeviceIntegrity = "no group given"
        Exit Function
    End If

    Dim parts As Object
    Set parts = PartsOf(grp)

    Dim slots As Long
    slots = SlotCount(grp)
    If slots = 0 Then
        DeviceIntegrity = "no milestone slots -- nothing named " & SLOT_PREFIX & "1" & PART_ON & _
            ". Name the shapes in PowerPoint's Selection Pane (Home > Select > Selection Pane)."
        Exit Function
    End If

    Dim missing As String
    Dim i As Long
    For i = 1 To slots
        ' _OFF is genuinely optional -- a template that shows achievement by
        ' something other than a second circle is a real design, and
        ' DrawMilestones already reports where it cannot show the difference.
        ' Reporting it as MISSING here would cry wolf on a valid template.
        missing = missing & MissingPart(parts, i, PART_LABEL)
        missing = missing & MissingPart(parts, i, PART_DATE)

        ' THE CIRCLES WERE NEVER CHECKED HERE AT ALL, and that was a real hole.
        ' It stayed hidden while SlotCount keyed on _ON alone: a slot that lost
        ' its _ON simply stopped being counted, and the stray-probe below caught
        ' the tail. Once SlotCount accepted any circle -- which it must, because
        ' Rohan's real slides carry slots drawn with only the big _NOW one -- a
        ' slot could count while being unable to render its main state, and
        ' nothing said so.
        '
        ' Two different faults, reported differently:
        ' NO "has no circle at all" BRANCH, deliberately. SlotCount only counts
        ' a slot that carries one, so inside this loop a circle-less slot cannot
        ' occur -- the loop would have stopped before reaching it, and the
        ' stray-part probe below catches its orphaned LABEL/DATE with a better
        ' message than a check here could give. It was written, found to be
        ' unreachable by construction, and removed: a branch that cannot execute
        ' reads as care taken and stops anyone looking again.
        If Not parts.Exists(UCase(SLOT_PREFIX & i & PART_ON)) Then
            ' Can render something, but not "achieved, earlier" -- which is the
            ' state most slots spend most of their life in. Named rather than
            ' counted, because the fix is to rename one shape.
            If missing <> "" Then missing = missing & ", "
            missing = missing & SLOT_PREFIX & i & PART_ON
        End If
    Next i

    ' THE SLOT AFTER THE LAST ONE, checked deliberately. A device with MS1..MS3
    ' complete and a stray MS5_LABEL means someone renamed MS4_ON and the whole
    ' tail went invisible -- SlotCount stops at the first gap, so without this
    ' the device would look perfectly healthy at three slots.
    Dim strays As String
    Dim probe As Long
    For probe = slots + 1 To slots + 6
        If parts.Exists(UCase(ColumnFor(probe, COL_LABEL))) _
           Or parts.Exists(UCase(ColumnFor(probe, COL_DATE))) _
           Or parts.Exists(UCase(SLOT_PREFIX & probe & PART_OFF)) Then
            If strays <> "" Then strays = strays & ", "
            strays = strays & SLOT_PREFIX & probe
        End If
    Next probe

    If parts.Exists(UCase(NAME_BAR)) = False Or parts.Exists(UCase(NAME_TRACK)) = False Then
        If missing <> "" Then missing = missing & ", "
        missing = missing & "the bar pair (" & NAME_BAR & " and/or " & NAME_TRACK & ")"
    End If

    Dim out As String
    If missing <> "" Then
        If Right(missing, 2) = ", " Then missing = Left(missing, Len(missing) - 2)
        out = slots & " slot(s) found, but these are missing: " & missing
    Else
        out = slots & " slot(s), all parts present"
    End If

    If strays <> "" Then
        out = out & ". AND parts exist for " & strays & " beyond the last complete slot -- " & _
            "something named " & SLOT_PREFIX & (slots + 1) & PART_ON & " is missing or misspelled, " & _
            "so every milestone after " & SLOT_PREFIX & slots & " is invisible to the tool."
    End If

    DeviceIntegrity = out
End Function

Private Function MissingPart(parts As Object, i As Long, part As String) As String
    If Not parts.Exists(UCase(SLOT_PREFIX & i & part)) Then
        MissingPart = SLOT_PREFIX & i & part & ", "
    End If
End Function

' THE ENTRY POINT SYNC USES: the row's values, straight from the register.
'
' `rowValues` is the slide's row as ExcelOutput hands it over -- field name to
' value. This reads MS1_LABEL / MS1_DATE / MS1_DONE for as many slots as the
' TEMPLATE carries, so adding a milestone means filling in columns, never
' changing code.
'
' A slot whose LABEL is blank is treated as absent, and that is the only rule
' about how many milestones there are. No count is stored anywhere: the register
' says what exists by having something in it.
Public Function DrawFromRow(grp As Object, rowValues As Object) As MilestoneDrawResult
    Dim result As MilestoneDrawResult

    If grp Is Nothing Then
        result.ErrorMessage = "no timeline group given"
        DrawFromRow = result
        Exit Function
    End If

    Dim slots As Long
    slots = SlotCount(grp)
    If slots = 0 Then
        result.ErrorMessage = "this group carries no milestone slots -- nothing named " & _
            SLOT_PREFIX & "1" & PART_ON & ". Name the shapes in PowerPoint's Selection Pane first."
        DrawFromRow = result
        Exit Function
    End If

    Dim labels() As String, dates() As String, done() As String
    ReDim labels(1 To slots): ReDim dates(1 To slots): ReDim done(1 To slots)

    Dim used As Long
    Dim gap As String
    Dim i As Long
    For i = 1 To slots
        labels(i) = ValueOr(rowValues, ColumnFor(i, COL_LABEL))
        dates(i) = ValueOr(rowValues, ColumnFor(i, COL_DATE))
        done(i) = ValueOr(rowValues, ColumnFor(i, COL_DONE))

        If Trim(labels(i)) <> "" Then
            ' A GAP IS REPORTED. Slots are drawn in order, so a filled MS3 with
            ' an empty MS2 means either a typo or a milestone someone meant to
            ' write. Drawing around it would quietly renumber the timeline.
            If used < i - 1 Then
                If gap <> "" Then gap = gap & ", "
                gap = gap & ColumnFor(i, COL_LABEL) & " is filled but " & _
                    ColumnFor(used + 1, COL_LABEL) & " is empty"
            End If
            used = i
        End If
    Next i

    If gap <> "" Then
        result.ErrorMessage = "the milestone columns have a gap: " & gap & _
            ". Nothing was drawn -- closing the gap here would renumber the timeline " & _
            "against what the register says."
        DrawFromRow = result
        Exit Function
    End If

    DrawFromRow = DrawMilestones(grp, labels, dates, done, used)
End Function

Private Function ValueOr(rowValues As Object, key As String) As String
    On Error Resume Next
    If Not rowValues Is Nothing Then
        If rowValues.Exists(key) Then ValueOr = CStr(rowValues(key))
    End If
    On Error GoTo 0
End Function

' Draws the device. `useCount` says how many of the parallel entries are real;
' DrawFromRow builds all three from one pass over the same columns, so they
' cannot disagree. The length check below guards the tests that call this
' directly.
Public Function DrawMilestones(grp As Object, labels() As String, dates() As String, _
                               done() As String, Optional useCount As Long = -1) As MilestoneDrawResult
    Dim result As MilestoneDrawResult

    If grp Is Nothing Then
        result.ErrorMessage = "no timeline group given"
        DrawMilestones = result
        Exit Function
    End If

    Dim nL As Long, nD As Long, nX As Long
    nL = UBound(labels) - LBound(labels) + 1
    nD = UBound(dates) - LBound(dates) + 1
    nX = UBound(done) - LBound(done) + 1

    ' `useCount` says how many entries are REAL. DrawFromRow fills all three
    ' arrays in one pass over the same slots, so they are the same length by
    ' construction and only the count differs -- overriding just nL made the
    ' guard compare 3 against 4 and refuse its own correctly-built data.
    ' The mismatch check below is for callers that build the arrays themselves.
    If useCount >= 0 Then
        nL = useCount: nD = useCount: nX = useCount
    End If

    If nL <> nD Or nL <> nX Then
        result.ErrorMessage = "the milestone lists are different lengths -- " & _
            nL & " label(s), " & nD & " date(s), " & nX & " done-flag(s). " & _
            "They are paired by position, so nothing was drawn: pairing them " & _
            "wrongly would put the right dates on the wrong milestones and the " & _
            "slide would look finished."
        DrawMilestones = result
        Exit Function
    End If

    Dim parts As Object
    Set parts = PartsOf(grp)

    result.SlotsFound = SlotCount(grp)
    If result.SlotsFound = 0 Then
        result.ErrorMessage = "this group carries no milestone slots -- nothing named " & _
            SLOT_PREFIX & "1" & PART_ON & ". Name the shapes in PowerPoint's Selection Pane first."
        DrawMilestones = result
        Exit Function
    End If

    If nL > result.SlotsFound Then
        result.ErrorMessage = nL & " milestone(s) in the register but only " & _
            result.SlotsFound & " slot(s) drawn on the template. Nothing was drawn -- " & _
            "showing the first " & result.SlotsFound & " would silently drop the rest."
        DrawMilestones = result
        Exit Function
    End If

    ' --- which slot is CURRENT, decided before anything is drawn ---------
    '
    ' lastAchieved used to be accumulated INSIDE the visibility loop, which was
    ' fine while there were two states: each slot's appearance depended only on
    ' its own done-flag. It is not fine with three. "Is this the current one?"
    ' is a question about the whole list, and a slot cannot answer it while the
    ' loop is still walking towards the answer -- slot 2 would have had to know
    ' whether slot 5 is achieved.
    '
    ' So it is computed first, from the data, before a single shape is touched.
    Dim lastAchieved As Long
    lastAchieved = 0

    Dim scan As Long
    For scan = 1 To nL
        If IsDoneWord(done(LBound(done) + scan - 1)) Then lastAchieved = scan
    Next scan

    Dim i As Long
    For i = 1 To result.SlotsFound
        Dim onShp As Object, offShp As Object, nowShp As Object
        Dim labShp As Object, datShp As Object
        Set onShp = PartOrNothing(parts, SLOT_PREFIX & i & PART_ON)
        Set offShp = PartOrNothing(parts, SLOT_PREFIX & i & PART_OFF)
        Set nowShp = PartOrNothing(parts, SLOT_PREFIX & i & PART_NOW)
        Set labShp = PartOrNothing(parts, SLOT_PREFIX & i & PART_LABEL)
        Set datShp = PartOrNothing(parts, SLOT_PREFIX & i & PART_DATE)

        If i <= nL Then
            Dim isDone As Boolean
            isDone = IsDoneWord(done(LBound(done) + i - 1))

            ' ACHIEVED IS SHOWN BY WHICHEVER CIRCLES THE TEMPLATE CARRIES.
            ' Two circles means toggle between them. One circle means the
            ' template has no way to show the difference, and that is REPORTED
            ' rather than faked by recolouring -- the tool does not invent
            ' formatting a person did not author.
            ' EXACTLY ONE CIRCLE IS SHOWN PER USED SLOT, and which one is the
            ' whole four-state model:
            '
            '   current (the last achieved)  -> _NOW   big
            '   achieved, earlier            -> _ON    small
            '   not achieved                 -> _OFF   small, other colour
            '   slot unused                  -> none   (handled below)
            '
            ' Every absent part DEGRADES to the next best circle the template
            ' actually carries and SAYS SO. It never recolours or resizes to
            ' fake a state the author did not draw -- that rule predates this
            ' and is the reason the device is trustworthy on a slide nobody
            ' has checked.
            Dim isCurrent As Boolean
            isCurrent = (i = lastAchieved)

            Dim shown As Object
            Set shown = Nothing

            If isCurrent Then
                Set shown = nowShp
                If shown Is Nothing Then
                    Set shown = onShp
                    NoteOnce result, SLOT_PREFIX & i & " is the current milestone but has no " & _
                        PART_NOW & " circle, so it looks the same as the earlier ones"
                End If
            ElseIf isDone Then
                Set shown = onShp
            Else
                Set shown = offShp
                If shown Is Nothing Then
                    Set shown = onShp
                    NoteOnce result, SLOT_PREFIX & i & " has no " & PART_OFF & _
                        " circle, so achieved and not-achieved look the same"
                End If
            End If

            ' Hide all three first, then show the one chosen. Setting them
            ' individually in each branch is how a stale circle survives a
            ' state change -- last quarter's _NOW would still be visible
            ' underneath this quarter's _ON.
            SetVisible onShp, False
            SetVisible offShp, False
            SetVisible nowShp, False
            If Not shown Is Nothing Then SetVisible shown, True

            SetVisible labShp, True
            SetVisible datShp, True
            WriteText labShp, labels(LBound(labels) + i - 1)
            WriteText datShp, dates(LBound(dates) + i - 1)

            result.Drawn = result.Drawn + 1
        Else
            ' A SLOT WITH NO MILESTONE IS HIDDEN, NOT EMPTIED. Blanking its text
            ' would leave a circle floating with nothing beside it.
            SetVisible onShp, False
            SetVisible offShp, False
            SetVisible nowShp, False
            SetVisible labShp, False
            SetVisible datShp, False
            result.Hidden = result.Hidden + 1
        End If
    Next i

    ' --- the bar ---------------------------------------------------------
    ' MEASURED, NOT COMPUTED. The bar reaches the last achieved circle, and
    ' where that circle sits is the template's decision -- read off the slide,
    ' never worked out from a fraction.
    Dim bar As Object, track As Object
    Set bar = PartOrNothing(parts, NAME_BAR)
    Set track = PartOrNothing(parts, NAME_TRACK)

    If bar Is Nothing Or track Is Nothing Then
        If result.Detail <> "" Then result.Detail = result.Detail & "; "
        result.Detail = result.Detail & "no " & NAME_BAR & "/" & NAME_TRACK & _
            " pair, so the progress bar was not touched"
        DrawMilestones = result
        Exit Function
    End If

    ' --- the track shortens to the last USED slot ------------------------
    '
    ' A seven-slot template running a five-milestone project used to leave two
    ' slots' worth of track hanging below the last circle, pointing at nothing.
    ' Rohan found it by asking what happens to a shorter chain.
    '
    ' The same measurement the bar already uses, aimed at a different target:
    ' the bar reaches the last ACHIEVED circle, the track reaches the last USED
    ' one. A height is one-dimensional, so neither can drift sideways or change
    ' z-order -- which is what made a computed height acceptable here when a
    ' computed position would not be.
    '
    ' ALWAYS SET, NEVER CONDITIONALLY. Set only when shortening is needed and a
    ' template that once ran a five-milestone project would keep the short track
    ' forever, silently truncating a seven-milestone one later. Recomputing from
    ' the data every run is what makes it self-correcting.
    Dim lastUsedCircle As Object
    If nL > 0 Then Set lastUsedCircle = SlotCircle(parts, nL)

    If lastUsedCircle Is Nothing Then
        NoteOnce result, "no circle found for slot " & nL & ", so the track was left at full length"
    Else
        Dim trackTarget As Single
        trackTarget = (lastUsedCircle.Top + lastUsedCircle.Height / 2) - track.Top
        If trackTarget < 0 Then trackTarget = 0
        track.Height = trackTarget
    End If

    Dim reachTo As Object
    If lastAchieved > 0 Then
        Set reachTo = SlotCircle(parts, lastAchieved)
    End If

    If reachTo Is Nothing Then
        ' Nothing achieved: the bar is empty, which is a real state and not an
        ' error. Height 0 rather than hidden -- hiding it would lose the shape
        ' from the stack the next run has to find.
        bar.Height = 0
        result.BarSet = True
    Else
        ' VERTICAL: the track runs top to bottom, so the bar's height is the
        ' distance from the track's top to the middle of the achieved circle.
        Dim target As Single
        target = (reachTo.Top + reachTo.Height / 2) - track.Top
        If target < 0 Then target = 0
        If target > track.Height Then target = track.Height
        bar.Height = target
        result.BarSet = True
    End If

    DrawMilestones = result
End Function

' Append a note to the report, skipping it if that exact sentence is already
' there. Named for what it does: seven slots missing the same part would
' otherwise say the same thing seven times and bury everything else.
Private Sub NoteOnce(ByRef result As MilestoneDrawResult, msg As String)
    If InStr(result.Detail, msg) > 0 Then Exit Sub
    If result.Detail <> "" Then result.Detail = result.Detail & "; "
    result.Detail = result.Detail & msg
End Sub

' The circle a slot actually carries, preferring the biggest statement of it.
' Used for MEASURING, where any of the three will do because they share a
' centre -- so this must never be used to decide what to SHOW.
Private Function SlotCircle(parts As Object, i As Long) As Object
    Dim s As Object
    Set s = PartOrNothing(parts, SLOT_PREFIX & i & PART_NOW)
    If s Is Nothing Then Set s = PartOrNothing(parts, SLOT_PREFIX & i & PART_ON)
    If s Is Nothing Then Set s = PartOrNothing(parts, SLOT_PREFIX & i & PART_OFF)
    Set SlotCircle = s
End Function

Private Function PartOrNothing(parts As Object, nm As String) As Object
    If parts.Exists(UCase(nm)) Then Set PartOrNothing = parts(UCase(nm))
End Function

Private Sub SetVisible(shp As Object, show As Boolean)
    If shp Is Nothing Then Exit Sub
    On Error Resume Next
    shp.Visible = IIf(show, msoTrue, msoFalse)
    On Error GoTo 0
End Sub

Private Sub WriteText(shp As Object, value As String)
    If shp Is Nothing Then Exit Sub
    On Error Resume Next
    If shp.HasTextFrame Then shp.TextFrame.TextRange.text = value
    On Error GoTo 0
End Sub
