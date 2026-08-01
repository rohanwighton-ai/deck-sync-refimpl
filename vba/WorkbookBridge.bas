Attribute VB_Name = "WorkbookBridge"
Option Explicit

' The sheet that explains the workbook. First tab, so it is what you land on.
Public Const INDEX_SHEET_NAME As String = "START HERE"

Public Const REGISTER_SHEET_NAME As String = "Register"

' THE REGISTER IS FOUND BY NAME, NEVER BY TAB POSITION.
'
' WriteWorkbookIndex ends with `ws.Move Before:=wb.Worksheets(1)` -- the index
' sheet deliberately puts itself at the front. The moment that shipped, every
' `wb.Worksheets(1)` in the codebase silently started returning the START HERE
' instructions sheet instead of the register.
'
' It failed silently because an empty register is a LEGAL state: no matching
' columns, no rows, no error. Callers reported "0 row(s) written" as a clean
' run. The drafting sheet went from 43 rows to 0 and nothing anywhere said why.
'
' E2EField.bas already carried the comment "Columns by header name, never by
' position" -- directly beneath a line picking the SHEET by position. The rule
' was known one level down and never applied one level up.
'
' Found 2026-08-01, after the FieldSpec compile error had hidden it for a day.
' Raises rather than returning Nothing: a workbook with no register is broken,
' and that must not be reportable as zero rows.
Public Function RegisterSheet(wb As Object) As Object
    Dim sh As Object
    For Each sh In wb.Worksheets
        If StrComp(sh.Name, REGISTER_SHEET_NAME, vbTextCompare) = 0 Then
            Set RegisterSheet = sh
            Exit Function
        End If
    Next sh
    Err.Raise vbObjectError + 513, "WorkbookBridge.RegisterSheet", _
        "No sheet named '" & REGISTER_SHEET_NAME & "' in this workbook. " & _
        "The register is located by name, not by tab position."
End Function

' Small shared primitive both RibbonUI.bas (Sync Now) and OnboardFlow.bas
' (Onboard New Slide Type, which establishes the pairing in the first
' place) need: given a workbook path, get a live Workbook object -- reusing
' an already-open instance if one matches, otherwise driving Excel via COM
' the same way this project's engine already does everywhere else (per
' vba-port.md: "VBA runs against a live Worksheet... via COM automation
' driven from the PowerPoint side"). Not a new sync/matching primitive --
' pure plumbing, split out once two ribbon-layer callers needed the exact
' same few lines rather than duplicating them.

' Reuses a running Excel instance if one exists (GetObject with no path
' argument attaches to it), otherwise starts a new one. Excel is left
' Visible so a user can see what's happening, same posture RunSync.bas's
' own cross-app calls already assume (ExcelOutput.bas operates on a live,
' visible Worksheet, not a hidden background instance).
Public Function GetExcelApp() As Object
    Dim xl As Object
    On Error Resume Next
    Set xl = GetObject(, "Excel.Application")
    On Error GoTo 0

    If xl Is Nothing Then
        Set xl = CreateObject("Excel.Application")
        xl.Visible = True
    End If

    Set GetExcelApp = xl
End Function

' Matches by full path against every open workbook first (avoids opening a
' second read-write handle onto a file someone already has open -- Excel
' itself would refuse or open read-only, neither of which this caller
' should silently paper over). Opens it fresh only if no match is found.
' Returns Nothing if `path` doesn't exist and can't be opened -- callers
' must handle that explicitly, this never raises.
Public Function OpenOrGetWorkbook(path As String) As Object
    Dim xl As Object
    Set xl = GetExcelApp()

    Dim wb As Object
    For Each wb In xl.Workbooks
        If LCase(wb.FullName) = LCase(path) Then
            Set OpenOrGetWorkbook = wb
            Exit Function
        End If
    Next wb

    On Error Resume Next
    Set wb = xl.Workbooks.Open(path)
    On Error GoTo 0

    Set OpenOrGetWorkbook = wb
End Function

' Creates `path` as a fresh, empty workbook if nothing exists there yet --
' the first-onboarding-on-this-deck case, where there is no paired workbook
' to open. Returns Nothing (never raises) if the path's containing folder
' doesn't exist or SaveAs otherwise fails.
Public Function CreateWorkbook(path As String) As Object
    Dim xl As Object
    Set xl = GetExcelApp()

    Dim wb As Object
    Set wb = xl.Workbooks.Add()

    On Error Resume Next
    wb.SaveAs path
    On Error GoTo 0

    ' Dir() must be guarded too, and wasn't -- this is the line that actually
    ' raised on 2026-07-29, not the SaveAs above it. Dir() throws runtime error
    ' 52 ("Bad file name or number") on a path it considers malformed, and an
    ' https:// URL is exactly that. So the one call written to CONFIRM the save
    ' worked was the one that blew the run up, while the risky-looking call it
    ' was checking sat safely inside a handler. This function's header promised
    ' "never raises" and was wrong for as long as it has existed.
    Dim landed As Boolean
    landed = False
    On Error Resume Next
    landed = (Dir(path) <> "")
    On Error GoTo 0

    If landed Then
        Set CreateWorkbook = wb
    Else
        Set CreateWorkbook = Nothing
    End If
End Function

' Sheet1 (Excel's own default first-sheet name) is reused for the first
' type registered against a fresh workbook rather than left as inert dead
' weight with a second, real sheet added alongside it -- confirmed safe
' since CreateSheet only refuses to reinitialize a sheet that already has a
' header in A1, and a brand-new Workbooks.Add() sheet is genuinely empty.
Public Function GetOrAddWorksheet(wb As Object, sheetName As String) As Object
    Dim ws As Object
    For Each ws In wb.Worksheets
        If ws.Name = sheetName Then
            Set GetOrAddWorksheet = ws
            Exit Function
        End If
    Next ws

    If wb.Worksheets.count = 1 And IsEmpty(wb.Worksheets(1).Cells(1, 1).Value) Then
        Set ws = wb.Worksheets(1)
        ws.Name = sheetName
    Else
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        ws.Name = sheetName
    End If

    Set GetOrAddWorksheet = ws
End Function

' Does `wb` already have a sheet by this name -- asked without creating one.
'
' GetOrAddWorksheet is the wrong tool for "is there a review to apply?": its Add
' half would answer the question by making the answer yes, leaving a blank sheet
' behind and reporting an empty queue as though a real review had come back with
' nothing ticked. Those two outcomes need to stay distinguishable, because one
' means "you approved nothing" and the other means "you never reviewed".
Public Function WorksheetExists(wb As Object, sheetName As String) As Boolean
    Dim ws As Object
    For Each ws In wb.Worksheets
        If ws.Name = sheetName Then
            WorksheetExists = True
            Exit Function
        End If
    Next ws
    WorksheetExists = False
End Function

' ---------------------------------------------------------------------
' The index sheet -- the workbook explaining itself
' ---------------------------------------------------------------------

' Writes a "START HERE" sheet listing every sheet, what it is for, and how long
' it lives.
'
' Rohan, 2026-08-01, on opening the register: "not clear on the sheets in it".
' The same failure as the drafting sheet earlier the same night -- a surface
' that assumes the reader already knows why it exists. A workbook that
' accumulates a register, a drafting sheet per field, a review grid per slide
' type and a log cannot be understood by looking at the tabs.
'
' The LIFESPAN column is the part that matters and the part nobody could infer.
' "Permanent" and "rebuilt every round" look identical as tabs, and the
' difference decides whether it is safe to type in one.
Public Sub WriteWorkbookIndex(wb As Object)
    Dim ws As Object
    Set ws = GetOrAddWorksheet(wb, INDEX_SHEET_NAME)
    ws.Cells.Clear

    ws.Cells(1, 1).Value = "WHAT IS IN THIS WORKBOOK"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 9

    ws.Cells(3, 1).Value = "Sheet"
    ws.Cells(3, 2).Value = "What it is"
    ws.Cells(3, 3).Value = "How long it lives"
    ws.Rows(3).Font.Bold = True

    Dim r As Long
    r = 4

    Dim sh As Object
    For Each sh In wb.Worksheets
        If sh.Name <> INDEX_SHEET_NAME Then
            ws.Cells(r, 1).Value = sh.Name
            ws.Cells(r, 2).Value = DescribeSheet(sh.Name)
            ws.Cells(r, 3).Value = LifespanOf(sh.Name)
            r = r + 1
        End If
    Next sh

    ws.Cells(r + 1, 1).Value = "The register is the record. The deck is a view of it -- " & _
        "if a slide and the register disagree, the register is what gets reviewed and applied."
    ws.Cells(r + 1, 1).Font.Italic = True

    ' 8pt, matching every other sheet the tools write. The title above keeps
    ' its own larger size -- set after this, so order matters.
    ws.Cells.Font.Size = 8
    ws.Cells(1, 1).Font.Size = 9
    ws.Cells.VerticalAlignment = -4160        ' xlTop
    ws.Columns(1).ColumnWidth = 26
    ws.Columns(2).ColumnWidth = 62
    ws.Columns(3).ColumnWidth = 30
    ws.Columns(2).WrapText = True

    On Error Resume Next
    ws.Move Before:=wb.Worksheets(1)   ' first tab, so it is what you land on
    On Error GoTo 0
End Sub

' Classified by name, because that is all a sheet carries. Pure, so the wording
' is testable without a workbook.
Public Function DescribeSheet(sheetName As String) As String
    If Left(sheetName, 4) = "TPL_" Then
        DescribeSheet = "Drafting sheet for " & Mid(sheetName, 5) & _
            ". Read column C, type new wording in F, put Y in G. Instructions are on the sheet."
    ElseIf Left(sheetName, 11) = "Sync Review" Then
        DescribeSheet = "Every change waiting to be approved before it reaches a slide. " & _
            "Tick what you agree with, then run Apply Approved."
    ElseIf sheetName = "Sync Log" Then
        DescribeSheet = "What was written to slides, and when. Written as it happens, so a run " & _
            "that dies halfway still leaves a record."
    ElseIf sheetName = "Field Spec" Then
        DescribeSheet = "How each field should be WRITTEN -- purpose, voice, length, and what " & _
            "not to do. Edit this to change the instructions the AI is given. Yours, not the tool's."
    ElseIf sheetName = "Sources" Then
        DescribeSheet = "WHERE THE WORDS CAME FROM. One row per source, referenced by ID from " & _
            "column E of a drafting sheet. Point at documents; do not paste them in here."
    ElseIf sheetName = "Register" Then
        DescribeSheet = "THE RECORD. One row per project, field and quarter, with its text and " & _
            "whether a human approved it. Everything else in this workbook feeds it or reads it."
    Else
        DescribeSheet = "(not created by this tool)"
    End If
End Function

Public Function LifespanOf(sheetName As String) As String
    If sheetName = "Register" Then
        LifespanOf = "PERMANENT -- grows each quarter"
    ElseIf Left(sheetName, 4) = "TPL_" Then
        LifespanOf = "Rebuilt each drafting round"
    ElseIf Left(sheetName, 11) = "Sync Review" Then
        LifespanOf = "One per run, then consumed"
    ElseIf sheetName = "Sync Log" Then
        LifespanOf = "Append-only history"
    ElseIf sheetName = "Field Spec" Then
        LifespanOf = "PERMANENT -- edit it freely"
    ElseIf sheetName = "Sources" Then
        LifespanOf = "PERMANENT -- accumulates, never rebuilt"
    Else
        LifespanOf = "unknown"
    End If
End Function

' Excel sheet names: max 31 chars, cannot contain \ / ? * [ ] : -- and
' cannot be blank. `slideType` is a free-form string with no such
' guarantee, so this is the one genuinely new piece of logic in this
' module (everything else above is COM plumbing) -- pure, no Excel object
' needed, testable directly.
Public Function SanitizeSheetName(rawName As String) As String
    Dim result As String
    result = rawName

    Dim badChars As String
    badChars = "\/?*[]:"
    Dim i As Long
    For i = 1 To Len(badChars)
        result = Replace(result, Mid(badChars, i, 1), "-")
    Next i

    If Len(result) > 31 Then
        result = Left(result, 31)
    End If

    If Trim(result) = "" Then
        result = "Data"
    End If

    SanitizeSheetName = result
End Function

' Does this workbook have edits that exist only in Excel's memory?
'
' Found live 2026-07-30, and it is not a nicety. GetExcelApp attaches to the
' RUNNING Excel instance, so the engine reads the workbook as it appears on
' screen, saved or not. A slide was created that evening from a row that
' existed in no file: the saved workbook held rows 1-4, the sync built a slide
' from row 5. Close Excel without saving at that point and the deck keeps a
' slide whose backing row is gone -- an orphan no future sync will ever update,
' with nothing to indicate it.
'
' The wider damage is to trust in Preview Sync. If data can change between the
' preview and the sync without any file changing, the preview is not a promise
' about the next write.
'
' Errors are treated as dirty, not clean. A workbook that cannot be asked
' whether it is saved is exactly the case not to assume the safe answer for.
Public Function IsDirty(wb As Object) As Boolean
    On Error GoTo Assume
    IsDirty = Not wb.Saved
    Exit Function
Assume:
    IsDirty = True
End Function

' What the human is asked when the paired workbook has unsaved edits. Pure, so
' the wording is pinned by a test rather than by a live click-through.
'
' Offers to save rather than refusing outright: the user is mid-task with Excel
' open in front of them, and "go and press Ctrl+S yourself" is friction with no
' safety benefit over doing it for them. Refusing is still the outcome if they
' decline -- syncing from a buffer is the thing being prevented, and there is
' no third option where it happens anyway.
Public Function UnsavedWorkbookText(workbookPath As String) As String
    UnsavedWorkbookText = _
        "The Data workbook has unsaved changes:" & vbCrLf & vbCrLf & _
        "    " & workbookPath & vbCrLf & vbCrLf & _
        "Syncing now would read values that exist only in Excel's memory, " & _
        "not in the file. If Excel is later closed without saving, any slide " & _
        "built from those values is left with no matching row." & vbCrLf & vbCrLf & _
        "Save the workbook and continue?"
End Function

' Format the register itself. It is the biggest sheet in the workbook and the
' one nothing had ever formatted -- it is written by the seeding and publishing
' paths, which are concerned with values, not with what it looks like to read.
'
' Cosmetic only: touches font, widths, alignment and the frozen header. It
' NEVER writes, moves or clears a cell value, because this is the record and a
' formatter has no business near its contents.
'
' Widths are chosen BY HEADER NAME, not by column position -- the register's
' column order is not guaranteed and assuming it is, is the exact mistake that
' cost 2026-08-01. An unrecognised header is left at whatever width it has.
Public Sub FormatRegisterSheet(ws As Object)
    On Error Resume Next          ' cosmetic: must never break a caller

    ws.Cells.Font.Size = 8
    ws.Cells.VerticalAlignment = -4160        ' xlTop
    ws.Rows(1).Font.Bold = True

    Dim c As Long
    For c = 1 To 20
        Dim h As String
        h = Trim(CStr(ws.Cells(1, c).Value))
        If h = "" Then
            ' keep going -- a gap does not mean the end of the header row
        ElseIf h = "Value" Then
            ws.Columns(c).ColumnWidth = 70
            ws.Columns(c).WrapText = True
        ElseIf h = "EntityCode" Or h = "FieldID" Or h = "SlideType" Then
            ws.Columns(c).ColumnWidth = 18
        ElseIf h = "Quarter" Or h = "Status" Or h = "FieldType" Then
            ws.Columns(c).ColumnWidth = 12
        ElseIf h = "CharCount" Then
            ws.Columns(c).ColumnWidth = 8
        ElseIf h = "UpdatedDate" Then
            ws.Columns(c).ColumnWidth = 13
        End If
    Next c

    ' Same reason as the drafting sheet: the Value column holds 350-500
    ' character paragraphs and a wrapped autofit turns every row into a page.
    ws.Rows(1).RowHeight = 26
    Dim lastRow As Long
    lastRow = 1
    Do While Trim(CStr(ws.Cells(lastRow + 1, 1).Value)) <> ""
        lastRow = lastRow + 1
    Loop
    If lastRow > 1 Then ws.Range(ws.Rows(2), ws.Rows(lastRow)).RowHeight = 40

    Dim xlApp As Object
    Set xlApp = ws.Application
    ws.Activate
    xlApp.ActiveWindow.FreezePanes = False
    ws.Cells(2, 1).Select
    xlApp.ActiveWindow.FreezePanes = True

    On Error GoTo 0
End Sub

' The register sheet, for callers that do not know which workbook shape they
' have been handed.
'
' TWO SHAPES ARE BOTH LEGITIMATE. The e2e rig uses a single sheet named
' "Register". A live pairing registers a sheet name per slide type. RegisterSheet
' above is exact and raises when there is no "Register" -- correct for the
' drafting path, wrong for tools pointed at either kind.
'
' So: the named register when it exists, otherwise the first sheet that is not
' one of the tool's OWN sheets. Those are excluded by name because they are the
' ones this tool creates, and the failure being fixed is precisely a tool
' reading its own instructions sheet and reporting an empty register as a clean
' run. Anything else is assumed to be the caller's data.
Public Function RegisterOrFirstDataSheet(wb As Object) As Object
    Dim sh As Object
    For Each sh In wb.Worksheets
        If StrComp(sh.Name, REGISTER_SHEET_NAME, vbTextCompare) = 0 Then
            Set RegisterOrFirstDataSheet = sh
            Exit Function
        End If
    Next sh

    For Each sh In wb.Worksheets
        If Not IsToolOwnedSheet(sh.Name) Then
            Set RegisterOrFirstDataSheet = sh
            Exit Function
        End If
    Next sh

    Err.Raise vbObjectError + 514, "WorkbookBridge.RegisterOrFirstDataSheet", _
        "This workbook contains only sheets created by the tool -- there is no " & _
        "register in it. Located by name, never by tab position."
End Function

' Sheets this tool creates and would never be a register.
Public Function IsToolOwnedSheet(sheetName As String) As Boolean
    If sheetName = INDEX_SHEET_NAME Then IsToolOwnedSheet = True
    If sheetName = "Field Spec" Then IsToolOwnedSheet = True
    If sheetName = "Sources" Then IsToolOwnedSheet = True
    If sheetName = "Sync Log" Then IsToolOwnedSheet = True
    If Left(sheetName, 4) = "TPL_" Then IsToolOwnedSheet = True
    If Left(sheetName, 11) = "Sync Review" Then IsToolOwnedSheet = True
    If Left(sheetName, 13) = "Template Audit" Then IsToolOwnedSheet = True
End Function
