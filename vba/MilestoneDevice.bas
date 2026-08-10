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
' So: visibility toggles, text writes, and ONE height change on the fill. The
' template pre-places every slot; unused ones are hidden, not removed. Circle
' positions are READ to work out how far the bar should reach -- measured off
' the slide the template already laid out, never computed.

' Slot part names, as they appear in the Selection Pane. `n` is 1-based.
'   MS1_ON     the circle shown when the milestone IS achieved
'   MS1_OFF    the circle shown when it is NOT           (optional -- see below)
'   MS1_LABEL  its text
'   MS1_DATE   its date
' and for the bar itself:
'   MS_BAR     the part that grows
'   MS_TRACK   the full extent, read and never written
Public Const SLOT_PREFIX As String = "MS"
Public Const PART_ON As String = "_ON"
Public Const PART_OFF As String = "_OFF"
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
    Do While parts.Exists(SLOT_PREFIX & (n + 1) & PART_ON)
        n = n + 1
    Loop
    SlotCount = n
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

    ' --- the slots -------------------------------------------------------
    Dim lastAchieved As Long
    lastAchieved = 0

    Dim i As Long
    For i = 1 To result.SlotsFound
        Dim onShp As Object, offShp As Object, labShp As Object, datShp As Object
        Set onShp = PartOrNothing(parts, SLOT_PREFIX & i & PART_ON)
        Set offShp = PartOrNothing(parts, SLOT_PREFIX & i & PART_OFF)
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
            If offShp Is Nothing Then
                If result.Detail <> "" Then result.Detail = result.Detail & "; "
                result.Detail = result.Detail & SLOT_PREFIX & i & " has no " & PART_OFF & _
                    " circle, so achieved and not-achieved look the same"
                SetVisible onShp, True
            Else
                SetVisible onShp, isDone
                SetVisible offShp, Not isDone
            End If

            SetVisible labShp, True
            SetVisible datShp, True
            WriteText labShp, labels(LBound(labels) + i - 1)
            WriteText datShp, dates(LBound(dates) + i - 1)

            If isDone Then lastAchieved = i
            result.Drawn = result.Drawn + 1
        Else
            ' A SLOT WITH NO MILESTONE IS HIDDEN, NOT EMPTIED. Blanking its text
            ' would leave a circle floating with nothing beside it.
            SetVisible onShp, False
            SetVisible offShp, False
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

    Dim reachTo As Object
    If lastAchieved > 0 Then
        Set reachTo = PartOrNothing(parts, SLOT_PREFIX & lastAchieved & PART_ON)
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
