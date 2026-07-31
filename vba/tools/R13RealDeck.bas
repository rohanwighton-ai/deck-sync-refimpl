Attribute VB_Name = "R13RealDeck"
Option Explicit

' R13 driven against a COPY of the real 46-slide deck.
'
' Two phases, deliberately separate entry points so the read-only one cannot
' become the writing one by accident:
'
'   ReviewOnly  -- opens the deck READ-ONLY, builds the queue, writes the grid
'                  to its own workbook, reports. Cannot change a slide.
'   ApplyPhase  -- opens the deck read-write, reads the ticked grid back, and
'                  applies exactly what was approved.
'
' WHY THIS RIG MATTERS MORE THAN A FIXTURE. The register in this rig carries 46
' ABOUT_BODY rows, all Status=Approved, seeded from a .tsv. If any of that text
' was derived by parsing the .pptx XML rather than read through the object
' model, it disagrees with what PowerPoint reports for nearly every
' multi-paragraph field (AGENTS.md) -- so a sync would want to rewrite 46 prose
' fields with subtly wrong text, and the old count-based confirmation would have
' presented that as "46 slide(s) corrected -- Proceed?".
'
' That is precisely the failure R13 was issued to prevent, and it is sitting in
' this rig live rather than hypothetically. A correct run classifies every one of
' those 46 as PROSE -> individual -> the worksheet, and writes none of them
' without someone reading them.

' The Application.Run UDT warm-up trap (see WORKPLAN / E2EFirstField's probes):
' in a freshly Imported project a Public Function is only reachable via
' Application.Run once the cross-module Public UDTs it declares have been
' touched by an earlier Application.Run in the same session. It fails as "Sub or
' function not defined" and looks exactly like a compile error.
'
' This probe declares the two UDTs this module's real functions use, so calling
' it first makes them resolvable.
Public Function PingR13() As String
    Dim q As ReviewQueueSet
    Dim it As ReviewItem
    PingR13 = "pongR13"
End Function

Private Function OpenRegisterSheet(registerPath As String, period As String, _
                                   ByRef xl As Object, ByRef wb As Object, _
                                   ByRef diag As String) As Sheet
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(registerPath, 0, True)   ' read-only

    Dim reg As RegisterRead
    reg = Register.ReadRegister(wb.Worksheets(1), period, "q")

    diag = "--- register ---" & vbCrLf & _
        "  rows seen:       " & reg.RowsSeen & vbCrLf & _
        "  accepted:        " & reg.Accepted & _
            "  (period " & reg.AcceptedPeriod & ", static " & reg.AcceptedStatic & ")" & vbCrLf & _
        "  rejected status: " & reg.RejectedStatus & vbCrLf & _
        "  rejected period: " & reg.RejectedPeriod & vbCrLf & _
        "  missing columns: '" & reg.MissingColumns & "'" & vbCrLf & vbCrLf

    OpenRegisterSheet = reg.Data
End Function

' Phase 1. Opens the deck READ-ONLY -- the third argument to Presentations.Open
' is ReadOnly:=msoTrue, so PowerPoint itself refuses a write regardless of what
' this code asks for. The queue is built dry either way; this is belt and braces
' on a rig that holds a copy of real business content.
Public Function ReviewOnly(deckPath As String, registerPath As String, _
                           period As String, gridPath As String) As String
    Dim r As String

    Dim pres As Object
    Set pres = Application.Presentations.Open(deckPath, msoTrue, msoFalse, msoTrue)
    pres.Windows(1).Activate

    r = "Deck (READ-ONLY): " & deckPath & vbCrLf & _
        "Register:         " & registerPath & vbCrLf & _
        "Period:           " & period & vbCrLf & _
        "Slides:           " & pres.Slides.count & vbCrLf & vbCrLf

    Dim xl As Object, wb As Object, diag As String
    Dim sheet As Sheet
    sheet = OpenRegisterSheet(registerPath, period, xl, wb, diag)
    r = r & diag

    ' A QUEUE OF ZERO IS AMBIGUOUS and must never be read as "all good".
    ' It means "nothing differs" only if slides were actually gathered and
    ' matched. Zero gathered instances, or every row landing as new_record,
    ' produces the identical zero -- so the discriminating numbers are printed
    ' BEFORE the queue, not after.
    Dim instances() As Object
    instances = RunSync.GatherInstances("q")
    Dim ilo As Long, ihi As Long, iHas As Boolean, iCount As Long
    On Error Resume Next
    ilo = LBound(instances): ihi = UBound(instances)
    iHas = (Err.Number = 0)
    On Error GoTo 0
    If iHas Then iCount = ihi - ilo + 1

    Dim acts() As SyncAction
    acts = SyncOperations.PlanRoutineSync(instances, sheet.InstanceOrder, sheet.Rows, True)
    Dim alo As Long, ahi As Long, aHas As Boolean
    On Error Resume Next
    alo = LBound(acts): ahi = UBound(acts)
    aHas = (Err.Number = 0)
    On Error GoTo 0

    Dim kNo As Long, kCorr As Long, kNew As Long, kFlag As Long
    If aHas Then
        Dim ai As Long
        For ai = alo To ahi
            Select Case acts(ai).Kind
                Case "no_change":           kNo = kNo + 1
                Case "in_place_correction": kCorr = kCorr + 1
                Case "new_record":          kNew = kNew + 1
                Case "flagged":             kFlag = kFlag + 1
            End Select
        Next ai
    End If

    r = r & "--- what the planner saw ---" & vbCrLf & _
        "  slides gathered as type 'q': " & iCount & vbCrLf & _
        "  register instance keys:      " & sheet.InstanceOrder.count & vbCrLf & _
        "  no_change:          " & kNo & vbCrLf & _
        "  in_place_correction:" & kCorr & vbCrLf & _
        "  new_record (NO MATCHING SLIDE): " & kNew & vbCrLf & _
        "  flagged:            " & kFlag & vbCrLf & vbCrLf

    Dim q As ReviewQueueSet
    q = ReviewQueue.BuildQueue(sheet, "q")

    r = r & "--- R13 queue ---" & vbCrLf & _
        "  changes needing review: " & q.Count & vbCrLf & _
        "  individual:             " & ReviewQueue.IndividualCount(q) & vbCrLf & _
        "  distinct batches:       " & ReviewQueue.DistinctBatchCount(q) & vbCrLf & _
        "  fast path available:    " & ReviewQueue.HasBatchableWork(q) & vbCrLf & vbCrLf

    ' Per-FieldID split, which is the number that says whether R13 classified
    ' this rig correctly. Prose must be 100% individual.
    Dim byField As Object, byFieldInd As Object
    Set byField = CreateObject("Scripting.Dictionary")
    Set byFieldInd = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To q.Count
        Dim f As String
        f = q.Items(i).FieldID
        If byField.Exists(f) Then byField(f) = CLng(byField(f)) + 1 Else byField(f) = 1
        If q.Items(i).BatchLabel = "" Then
            If byFieldInd.Exists(f) Then byFieldInd(f) = CLng(byFieldInd(f)) + 1 Else byFieldInd(f) = 1
        End If
    Next i

    r = r & "--- by field ---" & vbCrLf
    Dim k As Variant
    For Each k In byField.Keys
        Dim ind As Long
        If byFieldInd.Exists(k) Then ind = CLng(byFieldInd(k)) Else ind = 0
        r = r & "  " & k & ": " & byField(k) & " change(s), " & ind & " individual, kind=" & _
            ReviewQueue.ContentKindOf(CStr(k)) & vbCrLf
    Next k
    r = r & vbCrLf

    If ReviewQueue.DistinctBatchCount(q) > 0 Then
        r = r & "--- batches ---" & vbCrLf & ReviewQueue.BatchSummaryText(q) & vbCrLf
    End If

    ' The grid, into its own workbook. Saved so the apply phase can read the
    ' same ticks back -- which is the whole resumability claim, exercised
    ' against a real file rather than asserted.
    Dim gwb As Object, gws As Object
    Set gwb = xl.Workbooks.Add()
    Set gws = gwb.Worksheets(1)
    gws.Name = "Sync Review q"
    ReviewQueue.WriteQueueSheet gws, q
    On Error Resume Next
    Kill gridPath
    On Error GoTo 0
    gwb.SaveAs gridPath
    gwb.Close False
    r = r & "Grid written: " & gridPath & vbCrLf

    wb.Close False
    xl.Quit
    pres.Close

    r = r & vbCrLf & "NOTHING WAS WRITTEN TO THE DECK." & vbCrLf
    ReviewOnly = r
End Function

' Phase 2. Reads the ticked grid and applies exactly what it approves.
'
' `approveAll` exists for the scratch-copy setting (Round 13 §0.1) and is a
' separate argument someone has to pass, never a default.
Public Function ApplyPhase(deckPath As String, registerPath As String, _
                           period As String, gridPath As String, approveAll As String) As String
    Dim r As String

    Dim pres As Object
    Set pres = Application.Presentations.Open(deckPath, msoFalse, msoFalse, msoTrue)
    pres.Windows(1).Activate

    r = "Deck (read-write): " & deckPath & vbCrLf & _
        "Grid:              " & gridPath & vbCrLf & _
        "Approve-all:       " & approveAll & vbCrLf & vbCrLf

    Dim xl As Object, wb As Object, diag As String
    Dim sheet As Sheet
    sheet = OpenRegisterSheet(registerPath, period, xl, wb, diag)
    r = r & diag

    Dim gwb As Object, gws As Object
    Set gwb = xl.Workbooks.Open(gridPath)
    Set gws = gwb.Worksheets(1)

    If UCase(approveAll) = "YES" Then
        ReviewQueue.ApproveAllInSheet gws
        r = r & "APPROVE-ALL applied to the grid -- scratch copy only." & vbCrLf & vbCrLf
    End If

    Dim logWs As Object
    Set logWs = gwb.Worksheets.Add(After:=gwb.Worksheets(gwb.Worksheets.count))
    logWs.Name = "Sync Log"

    r = r & ReviewQueue.ApplyApproved(sheet, "q", gws, logWs)

    gwb.Save
    gwb.Close False
    wb.Close False
    xl.Quit

    pres.Save
    pres.Close

    ApplyPhase = r
End Function

' Phase 3: the SHIPPING Sync Now path, driven headlessly.
'
' Mirrors RibbonUI.SyncNowCore step for step -- build, approve ONLY the batched
' rows, write the grid, apply, then report what remains. Deliberately not a
' simplified stand-in: the point is to exercise the code a human actually
' reaches through the button, including the grid round-trip through a real file.
'
' The confirmation dialog is the one thing not exercised here (a headless run
' cannot answer a MsgBox). Its TEXT is asserted by the unit tests, so what is
' unproven is only that the dialog appears -- not what it says.
Public Function SyncNowPhase(deckPath As String, registerPath As String, _
                             period As String, gridPath As String) As String
    Dim r As String

    Dim pres As Object
    Set pres = Application.Presentations.Open(deckPath, msoFalse, msoFalse, msoTrue)
    pres.Windows(1).Activate

    Dim xl As Object, wb As Object, diag As String
    Dim sheet As Sheet
    sheet = OpenRegisterSheet(registerPath, period, xl, wb, diag)
    r = "Deck: " & deckPath & vbCrLf & vbCrLf & diag

    Dim q As ReviewQueueSet
    q = ReviewQueue.BuildQueue(sheet, "q")

    r = r & "--- queue as built ---" & vbCrLf & _
        "  total:            " & q.Count & vbCrLf & _
        "  individual:       " & ReviewQueue.IndividualCount(q) & vbCrLf & _
        "  distinct batches: " & ReviewQueue.DistinctBatchCount(q) & vbCrLf & _
        "  fast path:        " & ReviewQueue.HasBatchableWork(q) & vbCrLf & vbCrLf & _
        ReviewQueue.BatchSummaryText(q) & vbCrLf & _
        "--- what the dialog would say ---" & vbCrLf & _
        ReviewQueue.ConfirmBatchText(q) & vbCrLf & vbCrLf

    ' F5: only the batched rows are approved by the dialog.
    ReviewQueue.ApproveBatchedOnly q

    Dim gwb As Object, gws As Object
    Set gwb = xl.Workbooks.Add()
    Set gws = gwb.Worksheets(1)
    gws.Name = "Sync Review q"
    ReviewQueue.WriteQueueSheet gws, q

    Dim logWs As Object
    Set logWs = gwb.Worksheets.Add(After:=gwb.Worksheets(gwb.Worksheets.count))
    logWs.Name = "Sync Log"

    r = r & ReviewQueue.ApplyApproved(sheet, "q", gws, logWs) & vbCrLf

    ' F6: rebuild, so what remains describes the deck AFTER those writes.
    Dim after As ReviewQueueSet
    after = ReviewQueue.BuildQueue(sheet, "q")
    r = r & "--- after applying, rebuilt ---" & vbCrLf & _
        "  still pending: " & after.Count & vbCrLf & _
        "  individual:    " & ReviewQueue.IndividualCount(after) & vbCrLf & _
        "  batches:       " & ReviewQueue.DistinctBatchCount(after) & vbCrLf

    On Error Resume Next
    Kill gridPath
    On Error GoTo 0
    gwb.SaveAs gridPath
    gwb.Close False
    wb.Close False
    xl.Quit

    pres.Save
    pres.Close

    SyncNowPhase = r
End Function
