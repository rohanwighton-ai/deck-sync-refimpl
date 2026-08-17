Attribute VB_Name = "Timing"
Option Explicit

' -----------------------------------------------------------------------
' TIMING -- in-process instrumentation, so a slow run can be diagnosed from
' its own log instead of guessed at from outside via CPU polling.
' -----------------------------------------------------------------------
'
' WHY THIS EXISTS. Rohan, 2026-08-17, watching a live run: "can you please
' set timing machinery in the code at particular points to understand if
' the speeds we are getting are within an order of magnitude regarding what
' you expect for that part of the code?" External measurement (polling
' Get-Process CPU from PowerShell every 20s) can tell you THAT something is
' slow, but not WHICH part -- "1. Set up my quarter" and "2. Put it on the
' slides" are each several distinct stages (drafting-sheet crawl, Lobby
' rebuild, publish, review-queue build, apply), and only one of tonight's
' two real fixes (W: the log-write scan, Y: the shape-search cache) touches
' any given run. This writes a real timestamped duration for each stage to
' a sheet in the register, so a run's own numbers say which stage the time
' actually went to.
'
' TARGETS, in the SAME per-unit terms as the Sec/Unit column this module
' writes -- Rohan, 2026-08-17: "targets should be framed in those too" --
' so a logged rate can be read directly against a number, not against
' prose. All measured live 2026-08-17, before items W and Y were fixed --
' BEFORE numbers, kept as the floor a run under the FIXED code should beat:
'   - AppendLogLine, sec/call (old): grew with log length -- ~0.15 sec/call
'     at 100 calls in, uncompleted at 300 calls (would have been well over
'     0.4 sec/call). FIXED (item W): TARGET well under 0.01 sec/call,
'     flat regardless of log length -- if the logged rate grows as a run
'     goes on, the fix did not take.
'   - FindShapeByRoleTag, sec/item (old): ~4-5 sec/item on a real 221-item
'     apply run. FIXED (item Y): TARGET under 0.05 sec/item on a cache HIT;
'     a cache MISS (first use of a field, or genuine drift) still pays the
'     full walk, so an occasional item near the OLD number is expected --
'     the AVERAGE across a run is what should have moved, not every item.
' Nothing else in this codebase has a measured baseline yet -- the
' register-vs-slide diff scan inside BuildQueue/PlanRoutineSync in
' particular has NEVER been measured on its own, only inferred (FIX-LIST
' item X's update). This module exists to get that measurement, not to
' assert it in advance.
'
' WHY A SEPARATE SHEET, NOT FOLDED INTO Run Log OR Sync Log. Run Log is
' REPLACED each run (WorkbookBridge.WriteRunLog's own header) -- a timing
' history needs to survive across runs to be comparable. Sync Log is
' per-item outcomes, a different grain entirely. A timing entry is written
' the same safe way AppendLogLine now is (item W's own fix) -- Cells(Rows.
' Count,1).End(XL_UP).Row+1, never a rescan from row 2 -- so this
' instrumentation cannot reintroduce the exact cost it exists to diagnose.
Public Const TIMING_SHEET_NAME As String = "Timing"

Private Const COL_M_WHEN As Long = 1
Private Const COL_M_STAGE As Long = 2
Private Const COL_M_SECONDS As Long = 3
Private Const COL_M_UNITCOUNT As Long = 4
Private Const COL_M_UNITLABEL As Long = 5
Private Const COL_M_SECPERUNIT As Long = 6
Private Const COL_M_DETAIL As Long = 7

Private Const TIMING_HEADER_ROW As Long = 1
Private Const TIMING_FIRST_ROW As Long = 2

' Numeric, not the named constant xlUp -- same reason as ExcelOutput.bas's
' own XL_UP and ReviewQueue.bas's copy of it: this module can be driven
' cross-app from PowerPoint, where the named form does not resolve.
Private Const XL_UP As Long = -4162

' CANCEL BUDGET. Rohan, 2026-08-17: "include some cancel lines if target
' exceeded at a reasonable point" -- said right after tonight's live demo hit
' a run that went silent for ~2 minutes with no way to interrupt it short of
' force-closing both apps (FIX-LIST item X). This does NOT rescue that exact
' case: that stall was inside a single synchronous COM call (Ctrl+Break did
' not touch it, confirmed live), and a periodic between-items check can only
' ever run BETWEEN items, never inside one that has actually locked up --
' root-causing that is item X's own open problem, not this one. What this
' DOES catch is the more common shape of "running long": a run that is
' genuinely progressing, item by item, but doing so at several times the
' documented per-item target (e.g. a run with an unusually high cache-miss
' rate against ShapeAddressBook) -- rather than grinding through hundreds of
' remaining items at a bad rate with no visibility, a checkpoint every
' CANCEL_CHECK_INTERVAL_ITEMS offers a Yes/No only once elapsed time has
' genuinely blown past budget, never on a normal-speed run (elapsed stays
' under budget, so the check is silent) -- same "don't reintroduce a modal
' on the fast path" discipline as Phase 3's own pre-ticked queue.
Public Const CANCEL_CHECK_INTERVAL_ITEMS As Long = 10
Public Const CANCEL_CHECK_SEC_PER_ITEM_BUDGET As Double = 2#
Public Const CANCEL_CHECK_MIN_ELAPSED_SEC As Double = 15#

' Start a stopwatch. VBA's Timer (seconds since local midnight, sub-second
' resolution) is enough for this -- every stage measured here runs well
' inside one calendar day, and the one real failure mode (a run spanning
' midnight) would show as a negative duration, which LogTiming reports
' rather than hides, so a misleading number is at least a visibly odd one.
Public Function StartClock() As Double
    StartClock = Timer
End Function

' Writes one row: when, which stage, total seconds since StartClock was
' called for it, and -- the point of this redesign, Rohan 2026-08-17: "make
' sure the variables are quantifiable for the job it does so we understand
' per unit of whatever" -- how many UNITS of work that covered and the
' resulting SECONDS PER UNIT, as its own column, not buried in text. A raw
' total is only comparable to another run of the exact same size; a rate is
' comparable to this module's own documented expected ranges and to any
' other run regardless of how big it was.
'
' unitCount is Optional and defaults to 0 (no rate computed, e.g. a stage
' that genuinely has no natural unit) rather than required, so a call site
' with nothing sensible to count is not forced to invent one.
'
' Safe to call with wb=Nothing (does nothing) so instrumentation can be
' left in call sites that sometimes run before a workbook is resolved,
' without an extra guard at every site.
' excludeSeconds -- time already accounted for elsewhere (a human-wait
' bracketed with LogWait, below, inside this same stage) that must not
' also be counted as this stage's own processing time. Rohan, 2026-08-17:
' "include pauses for if the system is waiting for a human for whatever
' reason" -- the point is not to hide the wait, it is to stop it silently
' inflating the PROCESSING number next to it, which is the number these
' targets are actually meant to judge.
Public Sub LogTiming(wb As Object, stage As String, startedAt As Double, _
                     Optional unitCount As Long = 0, Optional unitLabel As String = "", _
                     Optional detail As String = "", Optional excludeSeconds As Double = 0)
    If wb Is Nothing Then Exit Sub

    Dim elapsed As Double
    elapsed = (Timer - startedAt) - excludeSeconds
    If elapsed < 0 Then elapsed = 0

    WriteTimingRow wb, stage, elapsed, unitCount, unitLabel, detail
End Sub

' For a duration already computed elsewhere -- an accumulator summed across
' a loop (ApplyApproved's probeSeconds/writeSeconds), not a fresh
' StartClock/Timer pair. LogTiming's startedAt/excludeSeconds arithmetic
' has no way to express "here is the number, directly" -- forcing one
' through it (excludeSeconds = -seconds, startedAt = 0) computes Timer +
' seconds instead, a real mistake made writing this module's own first
' draft and caught before being trusted, not a hypothetical one.
Public Sub LogDuration(wb As Object, stage As String, seconds As Double, _
                       Optional unitCount As Long = 0, Optional unitLabel As String = "", _
                       Optional detail As String = "")
    If wb Is Nothing Then Exit Sub
    If seconds < 0 Then seconds = 0
    WriteTimingRow wb, stage, seconds, unitCount, unitLabel, detail
End Sub

Private Sub WriteTimingRow(wb As Object, stage As String, elapsed As Double, _
                           unitCount As Long, unitLabel As String, detail As String)
    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(wb, TIMING_SHEET_NAME)

    If Trim(CStr(ws.Cells(TIMING_HEADER_ROW, COL_M_WHEN).Value)) = "" Then
        ws.Cells(TIMING_HEADER_ROW, COL_M_WHEN).Value = "When"
        ws.Cells(TIMING_HEADER_ROW, COL_M_STAGE).Value = "Stage"
        ws.Cells(TIMING_HEADER_ROW, COL_M_SECONDS).Value = "Seconds"
        ws.Cells(TIMING_HEADER_ROW, COL_M_UNITCOUNT).Value = "Units"
        ws.Cells(TIMING_HEADER_ROW, COL_M_UNITLABEL).Value = "Unit"
        ws.Cells(TIMING_HEADER_ROW, COL_M_SECPERUNIT).Value = "Sec/Unit"
        ws.Cells(TIMING_HEADER_ROW, COL_M_DETAIL).Value = "Detail"
        ws.Rows(TIMING_HEADER_ROW).Font.Bold = True
    End If

    Dim r As Long
    r = ws.Cells(ws.Rows.Count, COL_M_WHEN).End(XL_UP).Row + 1
    If r < TIMING_FIRST_ROW Then r = TIMING_FIRST_ROW

    ws.Cells(r, COL_M_WHEN).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(r, COL_M_STAGE).Value = stage
    ws.Cells(r, COL_M_SECONDS).Value = Format(elapsed, "0.000")
    If unitCount > 0 Then
        ws.Cells(r, COL_M_UNITCOUNT).Value = unitCount
        ws.Cells(r, COL_M_UNITLABEL).Value = unitLabel
        ws.Cells(r, COL_M_SECPERUNIT).Value = Format(elapsed / unitCount, "0.0000")
    End If
    ws.Cells(r, COL_M_DETAIL).Value = detail
End Sub

' A REGISTER OF WHAT WAS CLICKED, WHEN. Rohan, 2026-08-17: "also have a
' register of what I clicked when so that you can diagnose it." Written as
' its own row, prefixed "CLICKED:", at the moment each chain's own Resolve
' succeeds and a workbook first becomes available -- a few hundred ms after
' the real button press at most, close enough to correlate against
' everything else in this sheet by timestamp. Distinct from the "(total)"
' stage-duration row the same chain also logs: this one marks WHEN
' something started; that one reports how long it took once it finished.
' If a run hangs before ever reaching its "(total)" row, THIS row is still
' there, naming exactly which button and when -- the gap between it and
' whatever logs next (or the absence of anything logging next at all) is
' itself the diagnosis.
Public Sub LogClick(wb As Object, buttonCaption As String)
    If wb Is Nothing Then Exit Sub
    LogTiming wb, "CLICKED: " & buttonCaption, StartClock()
End Sub

' Logs a human-wait as ITS OWN row, prefixed so it reads unmistakably as
' "someone was looking at a dialog", not "the code was slow" -- a
' MsgBox/InputBox blocks VBA execution for as long as a person takes to
' answer it, which has nothing to do with how fast this add-in runs.
' Returns the waited duration so the caller can pass it to LogTiming's
' excludeSeconds for the stage this wait happened inside.
Public Function LogWait(wb As Object, description As String, startedAt As Double) As Double
    LogWait = Timer - startedAt
    LogTiming wb, "WAITING (human): " & description, startedAt
End Function

' Periodic checkpoint for a long item-by-item loop -- see the CANCEL BUDGET
' comment above for what this does and does not catch. Silent (returns False
' immediately, no dialog) whenever elapsedSoFar is still inside budget, so a
' normal-speed run never sees this. Only once elapsed genuinely exceeds
' itemsDone * secPerItemBudget, AND elapsed is past minElapsedFloor (so the
' ratio math is not just noise on a handful of items), does it ask. Adds its
' own wait time to waitedSeconds (ByRef) so the caller can exclude it from
' the surrounding stage's own processing rate, same discipline as LogWait.
Public Function CheckBudgetAndMaybeCancel(wb As Object, stage As String, _
        itemsDone As Long, itemsTotal As Long, elapsedSoFar As Double, _
        secPerItemBudget As Double, minElapsedFloor As Double, _
        ByRef waitedSeconds As Double) As Boolean
    CheckBudgetAndMaybeCancel = False
    If wb Is Nothing Then Exit Function

    Dim budget As Double
    budget = itemsDone * secPerItemBudget
    If elapsedSoFar <= budget Then Exit Function
    If elapsedSoFar <= minElapsedFloor Then Exit Function

    Dim tWait As Double
    tWait = StartClock()
    Dim msg As String
    msg = stage & " is running slower than expected: " & itemsDone & " of " & itemsTotal & _
        " item(s) done in " & Format(elapsedSoFar, "0") & "s (budget by this point: ~" & _
        Format(budget, "0") & "s)." & vbCrLf & vbCrLf & _
        "Anything already written stays written." & vbCrLf & _
        "Keep going, or stop here?"
    Dim answer As VbMsgBoxResult
    answer = MsgBox(msg, vbYesNo + vbQuestion, "Running long -- " & stage)
    waitedSeconds = waitedSeconds + LogWait(wb, stage & ": running-long check at item " & itemsDone, tWait)
    CheckBudgetAndMaybeCancel = (answer = vbNo)
End Function
