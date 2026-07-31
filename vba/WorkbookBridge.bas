Attribute VB_Name = "WorkbookBridge"
Option Explicit

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
