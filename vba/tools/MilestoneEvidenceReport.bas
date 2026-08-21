Attribute VB_Name = "MilestoneEvidenceReport"
Option Explicit

' One-off, read-only diagnostic -- NOT part of the shipped add-in (not in
' build_ppam.ps1's module list, not wired to CommandBarUI), matching
' vba/tools/VerifyRealDeck.bas's own pattern exactly.
'
' Answers one question: for each project's current-period milestones, does
' the register's MS<n>_DONE flag agree with what the real CRC tracker
' (SRC_MILESTONES, a permanent pasted extract -- see that sheet's own A7:
' "Feeds MS1..MS7 (label/date/done)") actually shows?
'
' WHY THIS EXISTS -- 2026-08-22 gap analysis. Rohan spotted a slide (2_P012)
' whose entire milestone timeline showed as "not achieved" despite the
' project being well underway; the register's MS*_DONE was blank across the
' board while SRC_MILESTONES showed real completions. A follow-up scan found
' 8 more projects with a apparent gap between source-confirmed completions
' and recorded DONE flags -- but a root-cause pass (Fable, same night) found
' the naive gap-count overstates badly: SRC_MILESTONES holds one row per
' TRACKER milestone (sometimes 20+ for one project) while the register only
' has 7 MS slots, so many "missing" completions are really just several
' fine-grained tracker items that all belong under one already-Y register
' slot. `2_P004` was proven to need NO fix once grouped properly.
'
' SRC_MILESTONES's own claim to "feed" MS*_DONE was never actually wired to
' any code (confirmed: zero other VBA references to that sheet). It is, and
' has been since FIX-LIST.md item BV, a DELIBERATE manual field -- mirroring
' it mechanically from Completion% is wrong in both directions (a project
' can close early with genuinely-incomplete later milestones, correctly
' recorded as not-done; a stale 0% can sit under a comment that plainly says
' the work is finished). So THIS TOOL NEVER DECIDES OR WRITES ANYTHING. It
' reports grouped evidence -- comment text included, since a comment beats a
' stale percentage -- for Rohan to read and apply by hand, the same review-
' before-write norm ReviewQueue.bas already holds for every other field.
'
' GROUPING: a tracker milestone with a numeric "Due Month # (offset)" (e.g.
' 9) belongs to the first register MS slot whose own MS<n>_DATE is a numeric
' month offset >= that due month (a deliverable due at month 9, with
' checkpoints at 6 and 12, is reported at the 12-month checkpoint -- the
' next one after it's due). A slot whose MS<n>_DATE is not a plain number
' (the kickoff "▶"/end "★" markers, or a literal date string some projects
' use instead) cannot be matched against a month-offset and is reported as
' "no numeric date to group against" rather than silently skipped.
Public Function MilestoneEvidenceReport(deckPath As String, workbookPath As String) As String
    Dim report As String
    Dim detail As String

    Dim pres As Object
    Set pres = Application.Presentations.Open(deckPath, msoTrue, msoFalse, msoFalse)

    Dim xl As Object, wb As Object, ws As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(workbookPath, 0, True)
    Set ws = WorkbookBridge.RegisterOrFirstDataSheet(wb)

    Dim deckPeriod As String
    deckPeriod = DeckRegistry.GetDeckPeriod(pres)

    Dim sheet As ExcelOutput.Sheet
    Dim readProblem As String
    sheet = ExcelOutput.ReadSheetForDeckPeriod(ws, deckPeriod, readProblem)
    If readProblem <> "" Then
        MilestoneEvidenceReport = "REFUSED: could not read the register for period '" & deckPeriod & "': " & readProblem
        wb.Close False
        xl.Quit
        pres.Close
        Exit Function
    End If

    ' SRC_MILESTONES: real header row is 10 (title/notes banner above it),
    ' data from row 11. Columns confirmed against the live workbook:
    '   A=Project Number  D=Milestone Name  G=Due Month # (offset)
    '   J=Knack Reporting Comments  L=Deliverable Completion RM %
    '
    ' READ ONCE, AS A BLOCK. The first version of this tool called
    ' srcWs.Cells(r, col).Value in a per-project loop -- one COM round trip
    ' per cell, ~560 rows re-scanned for every one of ~40 projects, tens of
    ' thousands of cross-process calls. Measured live: still running after
    ' several minutes with CPU time flat between checks. Same class of
    ' mistake this project has paid for before (DRAFTING-SPEED-STRATEGY.md's
    ' whole "bulk read/write" phase exists because of it). A single
    ' Range.Value2 read is ONE COM call regardless of row count; everything
    ' after this point works on the in-memory array.
    Dim srcWs As Object
    Set srcWs = wb.Sheets("SRC_MILESTONES")
    Dim srcLastRow As Long
    srcLastRow = srcWs.Cells(srcWs.Rows.count, "A").End(-4162).Row ' xlUp

    Dim srcBlock As Variant
    If srcLastRow >= 11 Then
        srcBlock = srcWs.Range("A11:L" & srcLastRow).Value2
    End If
    ' srcBlock(row, col): row is 1-based from A11, col 1=A .. 12=L

    Dim projectsChecked As Long, projectsWithGaps As Long
    Dim errorDetail As String

    Dim instanceId As Variant
    Dim stage As String
    For Each instanceId In sheet.InstanceOrder
        ' PER-PROJECT ERROR TRAP. First live run crashed the whole batch on
        ' a "Subscript out of range" with no indication which project or
        ' line -- a modal dialog left sitting for minutes, mistaken for a
        ' slow COM call before the screen was actually checked. One bad
        ' project's data must not stop every other project's report.
        ' `stage` records which block was entered last, since the error
        ' dialog itself does not say -- included in the error report.
        On Error GoTo InstanceError
        stage = "rowValues"

        Dim rowValues As Object
        Set rowValues = sheet.Rows(CStr(instanceId))

        stage = "slot loop"
        Dim slotNum As Long
        Dim slotOffset(1 To 7) As Long
        Dim slotOffsetKnown(1 To 7) As Boolean
        Dim slotDone(1 To 7) As Boolean
        Dim slotHasLabel(1 To 7) As Boolean
        Dim anySlot As Boolean
        anySlot = False

        For slotNum = 1 To 7
            stage = "slot loop, slotNum=" & slotNum
            Dim lbl As String, dt As String, dn As String
            lbl = ValueOr(rowValues, MilestoneDevice.ColumnFor(slotNum, MilestoneDevice.COL_LABEL))
            dt = ValueOr(rowValues, MilestoneDevice.ColumnFor(slotNum, MilestoneDevice.COL_DATE))
            dn = ValueOr(rowValues, MilestoneDevice.ColumnFor(slotNum, MilestoneDevice.COL_DONE))
            slotHasLabel(slotNum) = (Trim(lbl) <> "")
            If slotHasLabel(slotNum) Then anySlot = True
            slotDone(slotNum) = (StrComp(Trim(dn), "Y", vbTextCompare) = 0)
            If IsNumeric(Trim(dt)) Then
                slotOffset(slotNum) = CLng(Trim(dt))
                slotOffsetKnown(slotNum) = True
            Else
                slotOffsetKnown(slotNum) = False
            End If
        Next slotNum

        If Not anySlot Then GoTo NextInstance ' nothing to check -- no milestone plan on this row

        projectsChecked = projectsChecked + 1

        ' --- collect this project's SRC_MILESTONES rows, from the in-memory
        ' block read once above -- no COM calls in this loop at all.
        stage = "src row scan setup"
        Dim srcCount As Long
        Dim srcOffset() As Long, srcComplete() As Boolean, srcComment() As String
        Dim srcRowCount As Long
        srcRowCount = IIf(srcLastRow >= 11, srcLastRow - 11 + 1, 0)
        ReDim srcOffset(1 To srcRowCount + 1)
        ReDim srcComplete(1 To srcRowCount + 1)
        ReDim srcComment(1 To srcRowCount + 1)
        srcCount = 0

        Dim r As Long
        For r = 1 To srcRowCount
            stage = "src row scan, r=" & r & " of " & srcRowCount
            If StrComp(Trim(CStr(srcBlock(r, 1))), CStr(instanceId), vbTextCompare) = 0 Then
                Dim dueRaw As Variant
                dueRaw = srcBlock(r, 7) ' column G
                If IsNumeric(dueRaw) Then
                    srcCount = srcCount + 1
                    srcOffset(srcCount) = CLng(dueRaw)
                    Dim compRaw As Variant
                    compRaw = srcBlock(r, 12) ' column L
                    srcComplete(srcCount) = (IsNumeric(compRaw) And CDbl(compRaw) >= 1)
                    srcComment(srcCount) = CStr(srcBlock(r, 10)) ' column J
                End If
            End If
        Next r

        If srcCount = 0 Then GoTo NextInstance ' no source rows for this project -- nothing to compare

        stage = "GroupSourceIntoSlots call"
        Dim slotGroupN(1 To 7) As Long, slotGroupComplete(1 To 7) As Long
        Dim slotLatestComment(1 To 7) As String
        GroupSourceIntoSlots slotOffset, slotOffsetKnown, slotHasLabel, _
            srcOffset, srcComplete, srcComment, srcCount, _
            slotGroupN, slotGroupComplete, slotLatestComment

        ' --- report disagreements only ------------------------------------
        stage = "report loop"
        Dim projectHasGap As Boolean
        projectHasGap = False
        Dim projectDetail As String
        For slotNum = 1 To 7
            If slotHasLabel(slotNum) And slotGroupN(slotNum) > 0 Then
                Dim allComplete As Boolean
                allComplete = (slotGroupComplete(slotNum) = slotGroupN(slotNum))
                If allComplete <> slotDone(slotNum) Then
                    projectHasGap = True
                    projectDetail = projectDetail & "  MS" & slotNum & ": register says " & _
                        IIf(slotDone(slotNum), "DONE", "not done") & _
                        "; " & slotGroupComplete(slotNum) & " of " & slotGroupN(slotNum) & _
                        " grouped tracker item(s) complete."
                    If slotLatestComment(slotNum) <> "" Then
                        projectDetail = projectDetail & " Comment: " & Left(slotLatestComment(slotNum), 200)
                    End If
                    projectDetail = projectDetail & vbCrLf
                End If
            ElseIf slotHasLabel(slotNum) And Not slotOffsetKnown(slotNum) Then
                projectDetail = projectDetail & "  MS" & slotNum & ": no numeric date to group tracker items against (date='" & _
                    ValueOr(rowValues, MilestoneDevice.ColumnFor(slotNum, MilestoneDevice.COL_DATE)) & "')." & vbCrLf
            End If
        Next slotNum

        If projectHasGap Then
            projectsWithGaps = projectsWithGaps + 1
            detail = detail & instanceId & ":" & vbCrLf & projectDetail & vbCrLf
        End If

        On Error GoTo 0
NextInstance:
    Next instanceId
    GoTo AfterLoop

InstanceError:
    errorDetail = errorDetail & instanceId & " [" & stage & "]: Err " & Err.Number & " -- " & Err.Description & vbCrLf
    On Error GoTo 0
    Resume NextInstance

AfterLoop:

    wb.Close False
    xl.Quit
    pres.Close

    report = "=== Milestone Evidence Report ===" & vbCrLf & _
        "Deck: " & deckPath & vbCrLf & _
        "Workbook: " & workbookPath & vbCrLf & _
        "Period: " & deckPeriod & vbCrLf & _
        "Run at: " & Now & vbCrLf & vbCrLf & _
        "Projects with a milestone plan checked: " & projectsChecked & vbCrLf & _
        "Projects where grouped tracker evidence disagrees with a DONE flag: " & projectsWithGaps & vbCrLf & _
        IIf(errorDetail <> "", "Projects that errored during check (see below, skipped, did not stop the rest): " & _
            (Len(errorDetail) - Len(Replace(errorDetail, vbCrLf, ""))) / Len(vbCrLf) & vbCrLf, "") & vbCrLf & _
        "This report NEVER decides or writes anything -- read the comment text before " & _
        "changing any flag; a comment can override a stale percentage in either direction." & vbCrLf & vbCrLf & _
        IIf(errorDetail <> "", "--- Errors (per project, did not stop the run) ---" & vbCrLf & errorDetail & vbCrLf, "") & _
        "--- Per-project detail (only disagreements listed) ---" & vbCrLf & detail

    MilestoneEvidenceReport = report
End Function

Private Function ValueOr(rowValues As Object, key As String) As String
    On Error Resume Next
    If Not rowValues Is Nothing Then
        If rowValues.Exists(key) Then ValueOr = CStr(rowValues(key))
    End If
    On Error GoTo 0
End Function

' PURE, NO FILE I/O -- separated from the COM-driving function above so it
' can be unit tested directly, same reasoning as MilestoneDevice.DrawMilestones
' taking plain arrays rather than a live shape group.
'
' Groups each source tracker row into the first register slot (1 To 7)
' whose own MS<n>_DATE offset is >= that row's due-month offset -- a
' deliverable due at month 9, with register checkpoints at 6 and 12, is
' reported at the 12-month checkpoint (the next one after it falls due),
' never the earlier one it hadn't reached yet. A slot with no numeric date,
' or a project with no matching slot at or after a row's due month, gets
' nothing grouped into it -- callers report that as "no numeric date to
' group against", not as silent agreement.
'
' slotLatestComment takes the LAST non-blank comment seen for a slot, not a
' concatenation of every comment for it -- multiple source rows reporting
' the same milestone across quarters is expected (see any real
' SRC_MILESTONES row), and the most recent one is what should override a
' stale percentage, not an older one further down the same tracker.
Public Sub GroupSourceIntoSlots(slotOffset() As Long, slotOffsetKnown() As Boolean, slotHasLabel() As Boolean, _
                                srcOffset() As Long, srcComplete() As Boolean, srcComment() As String, srcCount As Long, _
                                ByRef slotGroupN() As Long, ByRef slotGroupComplete() As Long, ByRef slotLatestComment() As String)
    Dim i As Long, slotNum As Long
    For i = 1 To srcCount
        Dim bestSlot As Long
        bestSlot = 0
        For slotNum = 1 To 7
            If slotHasLabel(slotNum) And slotOffsetKnown(slotNum) Then
                If slotOffset(slotNum) >= srcOffset(i) Then
                    ' NOT "bestSlot = 0 Or slotOffset(slotNum) < slotOffset(bestSlot)" --
                    ' VBA's Or does not short-circuit, so that form evaluates
                    ' slotOffset(bestSlot) = slotOffset(0) even when bestSlot is
                    ' still 0, and slotOffset is dimensioned 1 To 7. Confirmed
                    ' live: this was "Subscript out of range" on every single
                    ' project, every run, because bestSlot starts at 0 and this
                    ' line is the first place that ever reads it as an index.
                    If bestSlot = 0 Then
                        bestSlot = slotNum
                    ElseIf slotOffset(slotNum) < slotOffset(bestSlot) Then
                        bestSlot = slotNum
                    End If
                End If
            End If
        Next slotNum
        If bestSlot > 0 Then
            slotGroupN(bestSlot) = slotGroupN(bestSlot) + 1
            If srcComplete(i) Then slotGroupComplete(bestSlot) = slotGroupComplete(bestSlot) + 1
            If Trim(srcComment(i)) <> "" Then slotLatestComment(bestSlot) = srcComment(i)
        End If
    Next i
End Sub
