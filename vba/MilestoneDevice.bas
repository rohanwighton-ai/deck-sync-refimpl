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
'   MS1_PCT    a Research Manager's % opinion, small text under the date
'              (optional -- see COL_PCT's header)
'   MS1_CALDATE a real calendar date ("Nov 2023"), small text under the
'              date -- same shape and position PART_PCT uses, different
'              content (optional -- see COL_CALDATE's header)
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
' Optional, like _OFF: a template without it simply shows no percentage,
' reported once rather than faked. See COL_PCT's header for the shape.
Public Const PART_PCT As String = "_PCT"
' A real calendar date, same optional-shape treatment as PART_PCT. See
' COL_CALDATE's header for why this is a separate field from PART_DATE
' (the existing month-count/glyph text) rather than a replacement for it.
Public Const PART_CALDATE As String = "_CALDATE"
Public Const NAME_BAR As String = "MS_BAR"
Public Const NAME_TRACK As String = "MS_TRACK"

Public Type MilestoneDrawResult
    Drawn As Long           ' slots given a milestone
    Hidden As Long          ' slots hidden for want of one
    SlotsFound As Long      ' slots the TEMPLATE carries
    BarSet As Boolean
    ErrorMessage As String
    Detail As String        ' mixed notes: genuine write failures AND template
                             ' limitations (a missing PART_PCT shape, etc.) --
                             ' NOT a reliable "did anything fail" signal on its
                             ' own. WriteFailureCount below is that signal.
    WriteFailureCount As Long ' incremented ONLY where a real write genuinely
                             ' did not take (a per-slot re-read mismatch),
                             ' never for a structural/template note
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

' A RESEARCH MANAGER'S JUDGMENT OF HOW FAR ALONG A MILESTONE IS.
' Rohan, 2026-08-22, after the on/off binary made a genuinely-progressing but
' not-yet-achieved milestone look identical to one nobody has touched:
' "perhaps we put a % done number in or next to each circle... this is where
' Research Managers give an informed opinion."
'
' TWO DECISIONS, NOT ONE, LANDED THE SAME EVENING. First tried as Option A
' (MILESTONE-PERCENTAGE-DESIGN.md): fold the number straight into MS1_LABEL's
' text, no new shape. Superseded within the hour by Rohan looking at a real
' rendered slide and asking a sharper question than the one this module had
' been answering: "isn't that separator needed?" led to "as a smaller font
' separate label under the date label" -- Option B, a real shape
' (MS1_PCT, PART_PCT above), and a second, bigger change alongside it:
'
' COLOUR BECOMES POSITIONAL, NOT PER-FLAG. "we spoke about making the gaps
' the same colour as done if they are before where the big circle currently
' is... colour are always contiguous." A slot's own DONE flag no longer
' decides ITS colour -- only whether it's <= the current slot, which is a
' question about the whole list (see lastAchieved in DrawMilestones). The
' flag still decides WHERE current sits; the percentage carries the honest
' "how far along" nuance that "done vs not done" colour used to lose.
'
' Blank means no opinion entered, same as every other part here -- the
' shape is simply left hidden.
Public Const COL_PCT As String = "_PCT"

' A REAL CALENDAR DATE, KEPT SEPARATE FROM PART_DATE ON PURPOSE.
' Rohan, 2026-08-22, looking at the promoted deck: "i also like the play
' button asterix and month number from start in the circles" (PART_DATE's
' existing content -- the play glyph/star/month-offset-from-start, all
' kept, none of this replaces it) "i think date could be added instead
' under each circle like we were going to with % deliverable complete" --
' the same optional-shape, small-text-under-the-date treatment PART_PCT
' uses, but showing the real month/year a milestone is due, not a
' percentage.
'
' NOT GATED BY `i <= lastAchieved` THE WAY PCT IS. A % complete is
' meaningless before a milestone has started; a due date is exactly the
' opposite -- most useful for a milestone that HASN'T happened yet, so
' this shape shows whenever a value exists, achieved or not.
'
' Sourced from `SRC_MILESTONES`' own "Resolved Due Date" column, which
' already existed for 308 of 370 real milestone rows across 38 projects
' (extracted from CRC's own tracker) and was simply never carried through
' to the register or a slide -- unlike PCT, which never had real data for
' more than the one project it was hand-tested against. That is the
' reason this feature can show up across the whole deck at once instead
' of on one slide with nothing behind the rest.
Public Const COL_CALDATE As String = "_CALDATE"

' The register column names this device reads, for slot `i`. Built here so the
' column names and the SHAPE names cannot drift apart -- they are deliberately
' the same strings, because a person looking at either should recognise the
' other.
Public Function ColumnFor(i As Long, part As String) As String
    ColumnFor = SLOT_PREFIX & i & part
End Function

' IS THIS REGISTER COLUMN OWNED BY THIS DEVICE, NOT AN ORDINARY FIELD?
'
' The device's columns (MS1_LABEL, MS1_DATE, MS1_DONE, MS2_LABEL, ...) are
' never individually role-tagged on a slide -- they are addressed by SHAPE
' NAME inside a group tagged with the device's own identity (InjectField's
' device route), exactly like Discovery already recognises the GROUP as one
' candidate instead of walking into its parts (Discovery.bas:165,
' `MilestoneDevice.SlotCount(shp) > 0`). FieldWiring.ScanFieldWiring asks
' "does any slide's role-tag set carry this exact field name" -- a question
' these columns can never answer yes to, because they were never meant to be
' answered that way. Reported as "21 fields on the register that no slide
' carries" on every run before this existed; not a false alarm about a real
' gap, a wrong question asked of columns it doesn't apply to.
'
' Recognised by SHAPE, matching `IsMilestoneInternalShape` below, not by a
' separate parallel pattern -- see that function's own header for why.
Public Function IsColumnForThisDevice(colName As String) As Boolean
    Dim c As String
    c = UCase(Trim(colName))
    If Left$(c, Len(SLOT_PREFIX)) <> SLOT_PREFIX Then Exit Function

    Dim rest As String
    rest = Mid$(c, Len(SLOT_PREFIX) + 1)

    Dim part As String
    If EndsWith(rest, COL_LABEL) Then
        part = COL_LABEL
    ElseIf EndsWith(rest, COL_DATE) Then
        part = COL_DATE
    ElseIf EndsWith(rest, COL_DONE) Then
        part = COL_DONE
    ElseIf EndsWith(rest, COL_PCT) Then
        part = COL_PCT
    ElseIf EndsWith(rest, COL_CALDATE) Then
        part = COL_CALDATE
    Else
        Exit Function
    End If

    Dim slotDigits As String
    slotDigits = Left$(rest, Len(rest) - Len(part))
    IsColumnForThisDevice = (slotDigits <> "" And IsAllDigits(slotDigits))
End Function

Private Function EndsWith(s As String, suffix As String) As Boolean
    EndsWith = (Len(s) >= Len(suffix)) And _
        (StrComp(Right$(s, Len(suffix)), suffix, vbTextCompare) = 0)
End Function

Private Function IsAllDigits(s As String) As Boolean
    Dim i As Long
    If s = "" Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    IsAllDigits = True
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
' STRING-RETURNING WRAPPER, EXISTS SOLELY FOR THE AUTOMATION BOUNDARY.
'
' Application.Run cannot marshal a UDT back to an external caller -- confirmed
' 2026-08-15 trying to call DrawFromRow directly from outside VBA: two other
' functions in this module (returning String, Boolean) worked from the same
' caller in the same session, and DrawFromRow alone failed with "Sub or
' function not defined", the generic error VBA gives for a signature Run
' cannot resolve. Nothing inside VBA is affected -- InjectPrimitive still
' calls DrawFromRow directly and always will. This exists only so a live
' slide can be exercised and inspected from PowerShell/COM, the same way
' every other verified-write function in this project already returns a
' plain string for exactly that reason.
Public Function DrawFromRowReport(grp As Object, rowValues As Object) As String
    Dim r As MilestoneDrawResult
    r = DrawFromRow(grp, rowValues)
    DrawFromRowReport = "Drawn=" & r.Drawn & " Hidden=" & r.Hidden & " SlotsFound=" & r.SlotsFound & _
        " BarSet=" & r.BarSet & " Error=[" & r.ErrorMessage & "] Detail=[" & r.Detail & "]"
End Function

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

    Dim labels() As String, dates() As String, done() As String, pcts() As String, caldates() As String
    ReDim labels(1 To slots): ReDim dates(1 To slots): ReDim done(1 To slots): ReDim pcts(1 To slots): ReDim caldates(1 To slots)

    Dim used As Long
    Dim gap As String
    Dim i As Long
    For i = 1 To slots
        labels(i) = ValueOr(rowValues, ColumnFor(i, COL_LABEL))
        dates(i) = ValueOr(rowValues, ColumnFor(i, COL_DATE))
        done(i) = ValueOr(rowValues, ColumnFor(i, COL_DONE))

        ' TRAILING "%" STRIPPED HERE, ONCE, so a Research Manager typing
        ' either "75" or "75%" renders identically -- WriteText below adds
        ' the "%" back on the way to the shape, never doubled.
        Dim pct As String
        pct = Trim(ValueOr(rowValues, ColumnFor(i, COL_PCT)))
        If Right$(pct, 1) = "%" Then pct = Trim(Left$(pct, Len(pct) - 1))
        pcts(i) = pct

        caldates(i) = ValueOr(rowValues, ColumnFor(i, COL_CALDATE))

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

    DrawFromRow = DrawMilestones(grp, labels, dates, done, pcts, caldates, used)
End Function

Private Function ValueOr(rowValues As Object, key As String) As String
    On Error Resume Next
    If Not rowValues Is Nothing Then
        If rowValues.Exists(key) Then ValueOr = CStr(rowValues(key))
    End If
    On Error GoTo 0
End Function

' Draws the device. `pcts` is required, same as labels/dates/done -- it is a
' field like the others, not a bolt-on (see COL_PCT's header). `useCount`
' says how many of the parallel entries are real; DrawFromRow builds all
' four from one pass over the same columns, so they cannot disagree. The
' length check below guards the tests that call this directly.
Public Function DrawMilestones(grp As Object, labels() As String, dates() As String, _
                               done() As String, pcts() As String, caldates() As String, _
                               Optional useCount As Long = -1) As MilestoneDrawResult
    Dim result As MilestoneDrawResult

    If grp Is Nothing Then
        result.ErrorMessage = "no timeline group given"
        DrawMilestones = result
        Exit Function
    End If

    Dim nL As Long, nD As Long, nX As Long, nP As Long, nC As Long
    nL = UBound(labels) - LBound(labels) + 1
    nD = UBound(dates) - LBound(dates) + 1
    nX = UBound(done) - LBound(done) + 1
    nP = UBound(pcts) - LBound(pcts) + 1
    nC = UBound(caldates) - LBound(caldates) + 1

    ' `useCount` says how many entries are REAL. DrawFromRow fills all five
    ' arrays in one pass over the same slots, so they are the same length by
    ' construction and only the count differs -- overriding just nL made the
    ' guard compare 3 against 4 and refuse its own correctly-built data.
    ' The mismatch check below is for callers that build the arrays themselves.
    If useCount >= 0 Then
        nL = useCount: nD = useCount: nX = useCount: nP = useCount: nC = useCount
    End If

    If nL <> nD Or nL <> nX Or nL <> nP Or nL <> nC Then
        result.ErrorMessage = "the milestone lists are different lengths -- " & _
            nL & " label(s), " & nD & " date(s), " & nX & " done-flag(s), " & nP & " pct(s), " & _
            nC & " caldate(s). " & _
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
        Dim labShp As Object, datShp As Object, pctShp As Object, caldShp As Object
        Set onShp = PartOrNothing(parts, SLOT_PREFIX & i & PART_ON)
        Set offShp = PartOrNothing(parts, SLOT_PREFIX & i & PART_OFF)
        Set nowShp = PartOrNothing(parts, SLOT_PREFIX & i & PART_NOW)
        Set labShp = PartOrNothing(parts, SLOT_PREFIX & i & PART_LABEL)
        Set datShp = PartOrNothing(parts, SLOT_PREFIX & i & PART_DATE)
        Set pctShp = PartOrNothing(parts, SLOT_PREFIX & i & PART_PCT)
        Set caldShp = PartOrNothing(parts, SLOT_PREFIX & i & PART_CALDATE)

        If i <= nL Then
            ' COLOUR IS POSITIONAL, NOT PER-FLAG -- Rohan, 2026-08-22: "colour
            ' are always contiguous." A slot's own DONE flag no longer decides
            ' ITS colour; only whether it sits at or before the current slot
            ' does. lastAchieved is already the HIGHEST done-flagged slot, so
            ' nothing after it could ever have been "done" anyway -- this only
            ' changes the render for a slot BEFORE current whose own flag is
            ' blank, which used to show as a visible gap and now doesn't. The
            ' flag still decides WHERE current sits (the scan loop above); the
            ' percentage (below) is what now carries the honest "not really
            ' 100%" nuance the colour used to.
            Dim isDone As Boolean
            isDone = (i <= lastAchieved)

            ' ACHIEVED IS SHOWN BY WHICHEVER CIRCLES THE TEMPLATE CARRIES.
            ' Two circles means toggle between them. One circle means the
            ' template has no way to show the difference, and that is REPORTED
            ' rather than faked by recolouring -- the tool does not invent
            ' formatting a person did not author.
            ' EXACTLY ONE CIRCLE IS SHOWN PER USED SLOT, and which one is the
            ' whole four-state model:
            '
            '   current (the last achieved)  -> _NOW   big
            '   at or before current         -> _ON    small
            '   after current                -> _OFF   small, other colour
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
            ' EVERY WRITE IS CHECKED, AND A FAILURE IS REPORTED, NOT SWALLOWED.
            ' The slot still counts as Drawn -- a declined write is a defect to
            ' flag, not a reason to pretend the slot was skipped.
            Dim slotOk As Boolean
            slotOk = True
            slotOk = SetVisible(onShp, False) And slotOk
            slotOk = SetVisible(offShp, False) And slotOk
            slotOk = SetVisible(nowShp, False) And slotOk
            If Not shown Is Nothing Then slotOk = SetVisible(shown, True) And slotOk

            slotOk = SetVisible(labShp, True) And slotOk
            slotOk = SetVisible(datShp, True) And slotOk
            slotOk = WriteText(labShp, labels(LBound(labels) + i - 1)) And slotOk
            slotOk = WriteText(datShp, dates(LBound(dates) + i - 1)) And slotOk
            slotOk = SetDateColourToMatch(datShp, shown) And slotOk

            ' PCT IS OPTIONAL AT THE VALUE LEVEL, unlike LABEL/DATE -- a used
            ' slot with no Research Manager opinion yet simply shows nothing,
            ' same "blank means absent" rule as everywhere else in this device.
            '
            ' ALSO GATED BY POSITION, same rule as the colour above -- Rohan,
            ' looking at the first real one on screen: "not worth having on
            ' future ones, should just turn on when big lead circle reaches
            ' that point." A slot AFTER current hasn't been reached yet, so
            ' any percentage sitting in the register for it is premature and
            ' stays suppressed until the current marker (lastAchieved) is
            ' at or past that slot -- even if a value happens to be present.
            Dim pctValue As String
            pctValue = pcts(LBound(pcts) + i - 1)
            If Trim(pctValue) <> "" And i <= lastAchieved Then
                If pctShp Is Nothing Then
                    NoteOnce result, SLOT_PREFIX & i & " has a percentage to show but no " & _
                        PART_PCT & " shape on this template, so it isn't shown"
                Else
                    slotOk = SetVisible(pctShp, True) And slotOk
                    slotOk = WriteText(pctShp, pctValue & "%") And slotOk
                End If
            Else
                slotOk = SetVisible(pctShp, False) And slotOk
            End If

            ' CALDATE IS OPTIONAL AT THE VALUE LEVEL, same "blank means
            ' absent" rule -- but NOT gated by lastAchieved the way PCT is.
            ' A % complete is meaningless before a milestone starts; a due
            ' date is exactly the opposite (see COL_CALDATE's own header),
            ' so it shows whenever the register holds one, achieved or not.
            Dim caldValue As String
            caldValue = caldates(LBound(caldates) + i - 1)
            If Trim(caldValue) <> "" Then
                If caldShp Is Nothing Then
                    NoteOnce result, SLOT_PREFIX & i & " has a calendar date to show but no " & _
                        PART_CALDATE & " shape on this template, so it isn't shown"
                Else
                    slotOk = SetVisible(caldShp, True) And slotOk
                    slotOk = WriteText(caldShp, caldValue) And slotOk
                    slotOk = SetDateColourToMatch(caldShp, shown) And slotOk
                End If
            Else
                slotOk = SetVisible(caldShp, False) And slotOk
            End If

            If Not slotOk Then
                NoteOnce result, SLOT_PREFIX & i & _
                    ": a write did not take -- the slide may not match this report"
                result.WriteFailureCount = result.WriteFailureCount + 1
            End If

            result.Drawn = result.Drawn + 1
        Else
            ' A SLOT WITH NO MILESTONE IS HIDDEN, NOT EMPTIED. Blanking its text
            ' would leave a circle floating with nothing beside it.
            Dim hideOk As Boolean
            hideOk = True
            hideOk = SetVisible(onShp, False) And hideOk
            hideOk = SetVisible(offShp, False) And hideOk
            hideOk = SetVisible(nowShp, False) And hideOk
            hideOk = SetVisible(labShp, False) And hideOk
            hideOk = SetVisible(datShp, False) And hideOk
            hideOk = SetVisible(pctShp, False) And hideOk
            hideOk = SetVisible(caldShp, False) And hideOk

            If Not hideOk Then
                NoteOnce result, SLOT_PREFIX & i & _
                    ": could not fully hide -- a visibility write did not take"
                result.WriteFailureCount = result.WriteFailureCount + 1
            End If

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

' FIX-LIST item Q, 2026-08-15. Both writers used to be Subs: suppress the
' error, write, restore handling, tell the caller nothing. DrawMilestones
' could report a milestone "drawn" against a shape that never actually
' changed -- this project's own signature defect ("reports success without
' confirming the effect"), sitting in the writer for 21 of the tool's 29
' `Given` fields. Found by an Invisible-Failure audit the same evening;
' `SlideDuplication.bas:115` and `TemplateSlide.bas:122` already guard the
' same class of write with an explicit postcondition -- these two are that
' shape, now applied here.
'
' Both now return whether the write is CONFIRMED, by reading the property
' back rather than trusting the assignment did not raise. `shp.Visible` is
' exactly the kind of property this project has already measured PowerPoint
' silently declining elsewhere (LockAspectRatio, 2026-08-10) -- a write that
' does not raise is not evidence it landed.
Private Function SetVisible(shp As Object, show As Boolean) As Boolean
    If shp Is Nothing Then
        SetVisible = True     ' nothing to fail -- there was no shape to write to
        Exit Function
    End If
    On Error Resume Next
    shp.Visible = IIf(show, msoTrue, msoFalse)
    Dim landed As Boolean
    landed = (shp.Visible = IIf(show, msoTrue, msoFalse))
    On Error GoTo 0
    SetVisible = landed
End Function

' THE DATE TEXT'S COLOUR MATCHES WHICHEVER CIRCLE IS SHOWN FOR ITS SLOT --
' measured, not computed, same philosophy the bar's height already uses
' ("read off the slide, never worked out from a fraction"). No type/palette
' table lives here: `circleShp` is the ON/OFF/NOW shape DrawMilestones just
' decided to show for this slot, already carrying whatever colour the
' template (or an earlier colour retrofit) gave it, so the date can only
' ever agree with what a person can actually see on screen.
'
' THE OUTLINE, NOT THE FILL. Rohan, 2026-08-22: "make sure text is
' harmonious and visible" -- an OFF circle's FILL is a deliberately pale
' tint (right for a small shape, wrong for text: pale-on-light-background
' text is close to illegible). Its OUTLINE is the type's saturated colour
' instead, and CU (same session) already made every OFF circle's outline
' match that slide's own ON outline -- so the outline is both legible and
' already uniform across achieved/not-achieved, unlike the fill.
'
' FOUND LIVE 2026-08-22 (Rohan, looking at the promoted deck): every
' MS<n>_DATE across all 43 real slides was hardcoded scheme:bg2 (the old
' uniform teal) for slots 1-3 and a type-hex for slots 4-7, on EVERY
' slide regardless of that project's own achieved point -- not a rendering
' bug, a static value baked into the template and propagated by
' duplication, never touched by DrawMilestones (no Font.Color/ForeColor
' call existed anywhere in this module before this fix).
Private Function SetDateColourToMatch(datShp As Object, circleShp As Object) As Boolean
    If datShp Is Nothing Or circleShp Is Nothing Then
        SetDateColourToMatch = True ' nothing to fail -- there was no shape to write to
        Exit Function
    End If
    If Not datShp.HasTextFrame Then
        SetDateColourToMatch = False
        Exit Function
    End If
    On Error Resume Next
    Dim wantRGB As Long
    wantRGB = circleShp.Line.ForeColor.RGB
    datShp.TextFrame.TextRange.Font.Color.RGB = wantRGB
    Dim landed As Boolean
    landed = (datShp.TextFrame.TextRange.Font.Color.RGB = wantRGB)
    On Error GoTo 0
    SetDateColourToMatch = landed
End Function

Private Function WriteText(shp As Object, value As String) As Boolean
    If shp Is Nothing Then
        WriteText = True      ' nothing to fail -- there was no shape to write to
        Exit Function
    End If
    If Not shp.HasTextFrame Then
        WriteText = False     ' the old code silently no-op'd here. This is that case, named.
        Exit Function
    End If

    ' SAME LINE-BREAK CONVENTION InjectPrimitive's own text writer already uses
    ' (InjectPrimitive.bas:38, LINE_BREAK_DELIMITER = "||") -- this module's own
    ' header says a milestone label is a register column like any other
    ' ("these are FIELDS, not a new kind of thing"), so a two-line label needs
    ' the same conversion a two-line plain-text field already gets. Missing
    ' here, "||" showed up as two literal pipe characters on a real slide
    ' instead of a line break -- found live 2026-08-19 comparing a genuine
    ' two-line milestone against Rohan's real original timeline.
    Dim converted As String
    converted = Replace(value, InjectPrimitive.LINE_BREAK_DELIMITER, vbCr)

    On Error Resume Next
    shp.TextFrame.TextRange.text = converted
    Dim landed As Boolean
    landed = (shp.TextFrame.TextRange.text = converted)
    On Error GoTo 0
    WriteText = landed
End Function
