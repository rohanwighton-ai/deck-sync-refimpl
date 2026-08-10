Attribute VB_Name = "Readiness"
Option Explicit

' WHERE YOU ARE, computed fresh, on the first tab of the workbook.
'
' Rohan, 2026-08-09: "cant like requirements stack until 'ready to sync'? should
' allow a cleaner interface? form guides the user?"
'
' The toolbar numbers its buttons 0, 0b, 1, 2, 3, 4 and that ordering is a LIE
' about the shape of the work. The true shape is:
'
'     a deck-level prologue  ->  N per-field lanes that do NOT wait for each
'                                other  ->  a deck-level sync
'
' RefreshDraftingSheets builds every prose field's sheet at once and says it is
' "safe to run at any time"; PublishDraftsForField publishes ONE field. So
' ABOUT_BODY can be published and synced while KEY_EVENTS_BODY is half drafted.
' A 0-to-4 ladder cannot express that, and a readiness LADDER would just be a
' prettier version of the same wrong claim. Hence a form with a per-field block,
' not a checklist.
'
' ---------------------------------------------------------------------
' THE DANGER, AND THE FOUR RULES THAT ANSWER IT
' ---------------------------------------------------------------------
'
' A green readiness list is EXACTLY this project's defining defect -- reports
' success without confirming the effect -- and it is worse than every previous
' instance, because it is designed to be the thing consulted INSTEAD of
' checking. Get it wrong and it does not hide one bug; it teaches a person to
' stop looking. So:
'
'   1. THREE STATES, NEVER TWO. OK / BLOCKED / CANNOT TELL. A check that could
'      not run says so and is never quietly counted as a pass. FastPathRefusalText
'      earned this rule: "the deck's worst state and its healthiest state
'      produced the same sentence."
'
'   2. EVERY LINE NAMES WHAT IT READ AND WHEN. Not "Period Q4F26 ok" but
'      "Q4F26 -- read from the saved .pptx". On 2026-08-08 the tool said "Deck
'      period is now Q4F26" against a file untouched for three days; a line that
'      omits its source recreates that at the top of the workbook.
'
'   3. THE WORD READY IS FORBIDDEN UNLESS EVERY CHECK RAN. No "READY, 3 checks
'      skipped". Anything uncomputed makes the headline CANNOT CONFIRM.
'      ReviewQueue's backup rule states the principle: a reported safety that is
'      not there "is the reason you feel safe running the destructive write".
'
'   4. RECOMPUTED, NEVER REVISITED. The button REBUILDS this sheet; it must
'      never merely Activate an existing tab, and the banner carries its own
'      computation time so a tab left open overnight indicts itself.
'
' AND IT OFFERS, IT DOES NOT GATE. No button is disabled anywhere on the basis
' of what this sheet says. Someone syncing one field while another is mid-draft
' is doing something legitimate, and a tool that argues with that is wrong.

Public Const READY_SHEET_NAME As String = "START HERE"

Public Const ST_OK As String = "ok"
Public Const ST_BLOCKED As String = "BLOCKED"
Public Const ST_UNKNOWN As String = "CANNOT TELL"

' ---------------------------------------------------------------------
' REMEDIES ARE A CLOSED SET, NOT FREE TEXT.
'
' Each line used to carry its remedy as a hand-typed string, and one of them
' said "Create Template Slide" -- a button this toolbar has never had. A person
' following that advice goes looking for something that does not exist, which is
' the precise failure the toolbar redesign exists to stop, shipping already.
'
' As an enum the code switches on, a remedy naming a dead button stops being a
' typo and becomes a compile-time impossibility: the text is resolved in ONE
' place, from CommandBarUI's caption constants, so renaming a button moves every
' remedy with it.
' ---------------------------------------------------------------------
Public Enum RemedyCode
    RM_NONE = 0
    RM_START_QUARTER
    RM_ROLL_FORWARD
    RM_REPOINT_WORKBOOK
    RM_ONBOARD_SLIDES
    RM_TEMPLATE_FROM_ONBOARDING
    RM_SAVE_DECK_THEN_REBUILD
    RM_SAVE_WORKBOOK_THEN_REBUILD
    RM_DECK_UNREADABLE
    RM_MARK_MISSING_FIELDS
End Enum


' One row of the sheet. Source is the evidence half of rule 2 -- what was read,
' not merely what was concluded.
Public Type ReadyLine
    Label As String
    Value As String
    State As String
    Source As String
    Remedy As String
End Type

Public Type ReadyReport
    Lines() As ReadyLine
    Count As Long
    Blocked As Long
    Unknown As Long
End Type

' The single place a remedy becomes words. Button remedies are built from the
' captions rather than repeating them.
Public Function RemedyText(code As RemedyCode) As String
    Select Case code
        Case RM_NONE:                    RemedyText = ""
        Case RM_START_QUARTER:           RemedyText = "Press '" & CommandBarUI.CAP_START_QUARTER & "'"
        Case RM_ROLL_FORWARD:            RemedyText = "Press '" & CommandBarUI.CAP_ROLL_FORWARD & "'"
        Case RM_REPOINT_WORKBOOK:        RemedyText = "Press '" & CommandBarUI.CAP_REPOINT_WORKBOOK & "'"
        Case RM_ONBOARD_SLIDES:          RemedyText = "Press '" & CommandBarUI.CAP_ONBOARD_SLIDES & "'"

        ' NOT A BUTTON, and saying so is the whole point of this entry. The
        ' template slide is created from inside onboarding; there has never been
        ' a way to press for it directly.
        Case RM_TEMPLATE_FROM_ONBOARDING: RemedyText = _
            "Create it from within '" & CommandBarUI.CAP_ONBOARD_SLIDES & "' -- there is no separate button"

        Case RM_SAVE_DECK_THEN_REBUILD:  RemedyText = "Save the deck, then rebuild this sheet"
        Case RM_SAVE_WORKBOOK_THEN_REBUILD: RemedyText = "Save the workbook, then rebuild this sheet"
        Case RM_DECK_UNREADABLE:         RemedyText = "Open the deck from a local folder, or see the trace"

        ' TRUE ONLY BECAUSE THE ROUTE WAS BUILT WITH IT. Until 2026-08-10 the
        ' honest remedy here would have been "there is no way to do this":
        ' MarkFieldForBatch's single call site sat behind `If Not hasTypes`, so
        ' a deck with a registered slide type could not tag a new field at all.
        ' This line and SyncNowChainCore's marking branch are one change; do not
        ' keep one without the other.
        Case RM_MARK_MISSING_FIELDS:     RemedyText = _
            "Press '" & CommandBarUI.CAP_SYNC_NOW & "' -- it offers to tag them before syncing"

        ' An unrecognised code must never render as a confident blank -- the
        ' sheet would show a line with no way forward and look complete.
        Case Else:                       RemedyText = "(no remedy recorded -- code " & code & ")"
    End Select
End Function

' The headline. Deliberately a function of the counts alone, so it cannot drift
' away from the lines beneath it.
Public Function Headline(r As ReadyReport) As String
    If r.Unknown > 0 Then
        Headline = "CANNOT CONFIRM -- " & r.Unknown & " of " & r.Count & _
            " check(s) could not run"
    ElseIf r.Blocked > 0 Then
        Headline = "NOT READY -- " & r.Blocked & " of " & r.Count & " check(s) blocked"
    ElseIf r.Count = 0 Then
        ' Zero checks is not a pass. Nothing ran.
        Headline = "CANNOT CONFIRM -- no checks ran"
    Else
        Headline = "READY -- " & r.Count & " check(s) passed"
    End If
End Function

Public Sub AddLine(ByRef r As ReadyReport, label As String, value As String, _
                   state As String, source As String, Optional remedy As String = "")
    r.Count = r.Count + 1
    ReDim Preserve r.Lines(1 To r.Count)
    r.Lines(r.Count).Label = label
    r.Lines(r.Count).Value = value
    r.Lines(r.Count).State = state
    r.Lines(r.Count).Source = source
    r.Lines(r.Count).Remedy = remedy
    If state = ST_BLOCKED Then r.Blocked = r.Blocked + 1
    If state = ST_UNKNOWN Then r.Unknown = r.Unknown + 1
End Sub

' Builds the report for a deck and its workbook.
'
' The deck lines are read OUT OF PROCESS, from the saved .pptx, because that is
' the only reading this project trusts about a property -- DeckRegistry.
' PropertyOnDisk opens a copy as a zip through Shell COM and parses
' docProps/custom.xml. The workbook lines are read LIVE, which is correct
' because Excel has it open, and made honest by refusing to report numbers from
' a workbook with unsaved edits.
Public Function Build(pres As Object, wb As Object) As ReadyReport
    Dim r As ReadyReport

    Dim deckPath As String
    On Error Resume Next
    deckPath = pres.FullName
    On Error GoTo 0

    ' --- the deck -----------------------------------------------------
    '
    ' READ FAILURE IS NOT ABSENCE, AND SAYING SO IS RULE 1.
    '
    ' Observed 2026-08-09 on a OneDrive-hosted deck: this reported "Period:
    ' BLOCKED -- not set in the saved file" while the file's bytes held Q3F26.
    ' PropertyOnDisk could not reach the file at all, and the only thing it could
    ' say about that was "", which read as a confident statement about the deck.
    ' BLOCKED sends a person to re-set a period that is already correct; CANNOT
    ' TELL sends them to find out why the file could not be read. The trace is
    ' carried into the Source column because rule 2 requires the line to name
    ' what it read, and "could not read it, here is how far it got" is that.
    Dim periodDisk As String, periodTrace As String, periodUnreadable As Boolean
    periodDisk = DeckRegistry.PropertyOnDisk(deckPath, DeckRegistry.PROP_DECK_PERIOD, _
                                             periodTrace, periodUnreadable)
    If periodUnreadable Then
        AddLine r, "Period", "COULD NOT READ THE SAVED FILE", ST_UNKNOWN, _
            "attempted: " & periodTrace, RemedyText(RM_DECK_UNREADABLE)
    ElseIf periodDisk = "" Then
        AddLine r, "Period", "(not set in the saved file)", ST_BLOCKED, _
            "saved .pptx", RemedyText(RM_START_QUARTER)
    Else
        AddLine r, "Period", periodDisk, ST_OK, "saved .pptx"
    End If

    ' A period held only in memory is the 2026-08-08 defect exactly: reported
    ' set, never written. Stated as its own line rather than folded into the one
    ' above, because "you have unsaved work" and "it never saved" need different
    ' actions from a person.
    Dim periodLive As String
    periodLive = DeckRegistry.GetDeckPeriod(pres)
    If periodLive <> "" And periodDisk <> "" Then
        If StrComp(periodLive, periodDisk, vbTextCompare) <> 0 Then
            AddLine r, "Period not saved", "PowerPoint says " & periodLive & _
                ", the file says " & periodDisk, ST_BLOCKED, _
                "saved .pptx vs PowerPoint", RemedyText(RM_SAVE_DECK_THEN_REBUILD)
        End If
    End If

    Dim wbPathDisk As String, wbTrace As String, wbUnreadable As Boolean
    wbPathDisk = DeckRegistry.WorkbookPathOnDisk(deckPath, wbTrace, wbUnreadable)
    If wbUnreadable Then
        AddLine r, "Paired workbook", "COULD NOT READ THE SAVED FILE", ST_UNKNOWN, _
            "attempted: " & wbTrace, RemedyText(RM_DECK_UNREADABLE)
    ElseIf wbPathDisk = "" Then
        AddLine r, "Paired workbook", "(none recorded in the saved file)", ST_BLOCKED, _
            "saved .pptx", RemedyText(RM_REPOINT_WORKBOOK)
    Else
        AddLine r, "Paired workbook", wbPathDisk, ST_OK, "saved .pptx"
    End If

    ' --- the workbook -------------------------------------------------
    '
    ' RULE 1 IN ITS MOST IMPORTANT PLACE. Everything below counts rows, and a
    ' workbook with unsaved edits would have this sheet report numbers that are
    ' not in any file. That is the failure this whole surface risks becoming, so
    ' it refuses rather than guessing, and says which.
    If wb Is Nothing Then
        AddLine r, "Register", "workbook not open", ST_UNKNOWN, "-- nothing read"
        Build = r
        Exit Function
    End If

    If WorkbookBridge.IsDirty(wb) Then
        AddLine r, "Register", "not read -- Excel is holding unsaved edits", ST_UNKNOWN, _
            "Excel (unsaved)", RemedyText(RM_SAVE_WORKBOOK_THEN_REBUILD)
        Build = r
        Exit Function
    End If

    Dim types() As String
    types = DeckRegistry.ListRegisteredTypes(pres)
    Dim lo As Long, hi As Long, hasTypes As Boolean
    On Error Resume Next
    lo = LBound(types): hi = UBound(types)
    hasTypes = (Err.Number = 0)
    On Error GoTo 0

    If Not hasTypes Then
        AddLine r, "Slide type", "(none registered)", ST_BLOCKED, "saved .pptx", _
            RemedyText(RM_ONBOARD_SLIDES)
        Build = r
        Exit Function
    End If

    Dim i As Long
    For i = lo To hi
        Dim templateSld As Object
        Dim wsName As String
        If Not DeckRegistry.LookupType(pres, types(i), templateSld, wsName) Then
            AddLine r, "Slide type " & types(i), "registered, but its slide is gone", _
                ST_BLOCKED, "saved .pptx", RemedyText(RM_ONBOARD_SLIDES)
        Else
            AddLine r, "Slide type", types(i) & " -> sheet '" & wsName & "'", ST_OK, "saved .pptx"

            ' A REGISTERED SLIDE IS NOT A TEMPLATE, and on the rig none is.
            ' Stated here because it is invisible everywhere else until the day
            ' a new project arrives and gets another project's slide.
            If templateSld Is Nothing Then
                AddLine r, "Template slide", "none for " & types(i), ST_BLOCKED, _
                    "deck", RemedyText(RM_TEMPLATE_FROM_ONBOARDING)
            ElseIf Not Resolve.IsTemplateSlide(templateSld) Then
                AddLine r, "Template slide", "NOT marked as a template -- a new project " & _
                    "would be copied from a real project's slide", ST_BLOCKED, "deck", _
                    RemedyText(RM_TEMPLATE_FROM_ONBOARDING)
            Else
                AddLine r, "Template slide", "present", ST_OK, "deck"
            End If

            If Not WorkbookBridge.WorksheetExists(wb, wsName) Then
                AddLine r, "Register sheet", "'" & wsName & "' is missing from this workbook", _
                    ST_BLOCKED, "open workbook", RemedyText(RM_REPOINT_WORKBOOK)
            Else
                Dim problem As String
                Dim sheet As Sheet
                sheet = ExcelOutput.ReadSheetForDeckPeriod( _
                    WorkbookBridge.GetOrAddWorksheet(wb, wsName), periodDisk, problem)
                If problem <> "" Then
                    AddLine r, "Rows for " & periodDisk, problem, ST_BLOCKED, _
                        "open workbook", RemedyText(RM_ROLL_FORWARD)
                Else
                    Dim q As ReviewQueueSet
                    q = ReviewQueue.BuildQueue(sheet, types(i))
                    AddLine r, "Rows for " & periodDisk, q.RowCount & " row(s), " & _
                        q.SlideCount & " tagged slide(s)", ST_OK, "open workbook"

                    ' Parity is the question the tool exists to answer, so it is
                    ' a line here and not a footnote.
                    If q.OrphanCount = 0 And q.SlideNoRowCount = 0 Then
                        AddLine r, "Parity", "deck and register agree", ST_OK, _
                            "deck + open workbook"
                    Else
                        AddLine r, "Parity", ReviewQueue.ParityText(q), ST_BLOCKED, _
                            "deck + open workbook", "Sync Now offers to create missing slides"
                    End If

                    ' IS EVERY FIELD ACTUALLY ATTACHED TO A SHAPE? Parity above
                    ' answers "do the SLIDES and ROWS line up"; this answers "does
                    ' each COLUMN have something on the slide to write into". A
                    ' deck can pass parity perfectly and still carry a register
                    ' field that nothing on any slide holds -- the field is then
                    ' carried all the way to the injector and refused there, once
                    ' per slide, after the drafting work is already done.
                    Dim wiring As FieldWiringResult
                    wiring = FieldWiring.ScanFieldWiring(types(i), sheet.Fields, templateSld)
                    If Not wiring.Scanned Then
                        AddLine r, "Fields tagged", FieldWiring.WiringText(wiring), _
                            ST_UNKNOWN, "-- the slides could not be read"
                    ElseIf wiring.UnmarkedCount > 0 Or wiring.OrphanCount > 0 _
                           Or wiring.TemplateUnmarkedCount > 0 Then
                        AddLine r, "Fields tagged", FieldWiring.WiringText(wiring), _
                            ST_BLOCKED, "deck + open workbook", _
                            RemedyText(RM_MARK_MISSING_FIELDS)
                    ElseIf Not wiring.TemplateScanned Then
                        ' Every field is on a slide, and nothing looked at the
                        ' template. That is not a pass: rule 1, a check that
                        ' could not run says so.
                        AddLine r, "Fields tagged", FieldWiring.WiringText(wiring), _
                            ST_UNKNOWN, "deck + open workbook"
                    Else
                        AddLine r, "Fields tagged", FieldWiring.WiringText(wiring), _
                            ST_OK, "deck + open workbook"
                    End If
                End If
            End If
        End If
    Next i

    Build = r
End Function

' Writes the report onto the first tab, replacing whatever was there.
'
' REBUILDS, never revisits (rule 4). The banner carries the time it was computed
' and both files' saved times, so a tab left open overnight says so itself.
Public Sub WriteSheet(wb As Object, pres As Object, r As ReadyReport)
    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(wb, READY_SHEET_NAME)
    ws.Cells.Clear

    ws.Cells(1, 1).Value = "DECK SYNC -- WHERE YOU ARE"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 11

    ws.Cells(2, 1).Value = "rebuilt " & Format(Now, "dd mmm yyyy hh:nn") & _
        "  --  this sheet is recomputed each time you press the button; it is not live."

    Dim deckName As String
    On Error Resume Next
    deckName = pres.Name
    On Error GoTo 0
    ws.Cells(3, 1).Value = "deck:     " & deckName
    ws.Cells(4, 1).Value = "workbook: " & wb.Name & _
        IIf(WorkbookBridge.IsDirty(wb), "   (UNSAVED EDITS -- numbers below are withheld)", "")

    ws.Cells(6, 1).Value = Headline(r)
    ws.Cells(6, 1).Font.Bold = True
    ws.Cells(6, 1).Font.Size = 10

    ws.Cells(8, 1).Value = "What"
    ws.Cells(8, 2).Value = "State"
    ws.Cells(8, 3).Value = "Value"
    ws.Cells(8, 4).Value = "Read from"
    ws.Cells(8, 5).Value = "Fix with"
    ws.Rows(8).Font.Bold = True

    Dim rowNo As Long
    rowNo = 9
    Dim i As Long
    For i = 1 To r.Count
        ws.Cells(rowNo, 1).Value = r.Lines(i).Label
        ws.Cells(rowNo, 2).Value = r.Lines(i).State
        ws.Cells(rowNo, 3).Value = r.Lines(i).Value
        ws.Cells(rowNo, 4).Value = r.Lines(i).Source
        ws.Cells(rowNo, 5).Value = r.Lines(i).Remedy
        If r.Lines(i).State <> ST_OK Then ws.Rows(rowNo).Font.Bold = True
        rowNo = rowNo + 1
    Next i

    ws.Cells(rowNo + 1, 1).Value = "Nothing here disables a button. Fields do not wait " & _
        "for each other -- you can sync one while another is still being drafted."
    ws.Cells(rowNo + 1, 1).Font.Italic = True

    ' The sheet index the old START HERE carried, kept below the readiness block
    ' rather than deleted: it is the only map of the tabs, and the tabs are where
    ' the evening actually goes.
    rowNo = rowNo + 3
    ws.Cells(rowNo, 1).Value = "THE TABS IN THIS WORKBOOK"
    ws.Cells(rowNo, 1).Font.Bold = True
    rowNo = rowNo + 1
    ws.Cells(rowNo, 1).Value = "Sheet"
    ws.Cells(rowNo, 3).Value = "What it is"
    ws.Cells(rowNo, 4).Value = "How long it lives"
    ws.Rows(rowNo).Font.Bold = True
    rowNo = rowNo + 1

    Dim sh As Object
    For Each sh In wb.Worksheets
        If sh.Name <> READY_SHEET_NAME Then
            ws.Cells(rowNo, 1).Value = sh.Name
            ws.Cells(rowNo, 3).Value = WorkbookBridge.DescribeSheet(sh.Name)
            ws.Cells(rowNo, 4).Value = WorkbookBridge.LifespanOf(sh.Name)
            rowNo = rowNo + 1
        End If
    Next sh

    ws.Cells.Font.Size = 8
    ws.Cells(1, 1).Font.Size = 11
    ws.Cells(6, 1).Font.Size = 10
    ws.Cells.VerticalAlignment = -4160        ' xlTop
    ws.Columns(1).ColumnWidth = 22
    ws.Columns(2).ColumnWidth = 12
    ws.Columns(3).ColumnWidth = 58
    ws.Columns(4).ColumnWidth = 22
    ws.Columns(5).ColumnWidth = 26
    ws.Columns(3).WrapText = True

    On Error Resume Next
    ws.Move Before:=wb.Worksheets(1)   ' first tab, so it is what you land on
    On Error GoTo 0
End Sub

' Brings the sheet to the front. Separate from WriteSheet so the writer stays
' testable without a visible Excel.
Public Sub ShowSheet(wb As Object)
    On Error Resume Next
    WorkbookBridge.GetOrAddWorksheet(wb, READY_SHEET_NAME).Activate
    wb.Activate
    On Error GoTo 0
End Sub
