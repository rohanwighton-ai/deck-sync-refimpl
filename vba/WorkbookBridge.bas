Attribute VB_Name = "WorkbookBridge"
Option Explicit

Public Const RUN_LOG_SHEET_NAME As String = "Run Log"

' THE SYNC LOG HAD NO CONSTANT, ALONE AMONG THE TOOL-OWNED SHEETS. It was a bare
' "Sync Log" literal in SEVEN places across two modules -- and two of them are
' GetOrAddWorksheet calls, which CREATE the sheet when the name does not match.
' So a single divergent literal would not fail: it would quietly start a second
' log sheet while IsToolOwnedSheet and ArrangeTabs went on guarding the first,
' splitting the audit trail in a way that looks entirely healthy. Same shape as
' every other stale-string defect in this codebase -- a wrong string with a
' plausible orphan beside it is worse than a wrong string on its own.
Public Const SYNC_LOG_SHEET_NAME As String = "Sync Log"

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

' The worksheet a slide type is REGISTERED against, or Nothing with a reason.
'
' THE ONE ANSWER TO "WHICH SHEET IS THIS TYPE'S REGISTER", for drafting,
' publish and sync alike. They used to answer it differently and that is a
' silent-wrong-answer bug, not an inconsistency:
'
'   sync     asked DeckRegistry for the worksheet name registered per slide type
'   drafting asked for a sheet literally NAMED "Register", and only fell back to
'            the registered name when no such sheet existed
'
' A workbook can easily have both -- a sheet called "Register" left by an early
' onboarding, plus a type registered against "Research Project Status". Publish
' then wrote one sheet while Sync Now read the other, and BOTH reported success.
' The text was really published and the slides really synced; they just were not
' the same text. The rig could never show it, because the rig's registered name
' IS "Register", so the two paths happened to coincide.
'
' IT REFUSES A MISSING SHEET RATHER THAN CREATING ONE. GetOrAddWorksheet creates,
' which is right when onboarding is establishing the pairing and catastrophic
' when something is merely looking the register up: a freshly invented blank
' sheet reads as a register with nothing in it, which is a clean sync of
' nothing. Creating a register is an act of onboarding, never of resolution.
Public Function WorksheetForSlideType(pres As Object, wb As Object, slideType As String, _
                                      ByRef problem As String) As Object
    problem = ""

    Dim templateSld As Object
    Dim wsName As String
    If Not DeckRegistry.LookupType(pres, slideType, templateSld, wsName) Then
        problem = "this deck has no slide type registered as '" & slideType & "'."
        Exit Function
    End If

    If Trim$(wsName) = "" Then
        problem = "slide type '" & slideType & "' is registered, but with no worksheet name."
        Exit Function
    End If

    If Not WorksheetExists(wb, wsName) Then
        problem = "slide type '" & slideType & "' is registered against a worksheet named '" & _
            wsName & "', and this workbook has no such sheet. Refusing to create it -- an " & _
            "invented blank sheet reads as a register with nothing in it, which reports as a " & _
            "clean run of nothing. Check the deck and workbook are the pair you meant."
        Exit Function
    End If

    Set WorksheetForSlideType = GetOrAddWorksheet(wb, wsName)
End Function

' Small shared primitive both RibbonUI.bas (Sync Now) and OnboardFlow.bas
' (Setup B: Onboard Slides, which establishes the pairing in the first
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

' Why this workbook cannot be written to, or "" if it can.
'
' READ-ONLY WAS NEVER CHECKED ANYWHERE. Confirmed 2026-08-01 across the whole
' vba/ tree: the only `.ReadOnly` references were two diagnostic prints in a
' tools module, and those were on the presentation, not the workbook.
'
' So a register that is locked -- open in someone else's Excel, on a read-only
' share, or opened read-only after a recovery prompt -- opens happily, every
' write appears to work, and the failure lands at wb.Save or nowhere at all.
' The returned object cannot distinguish "I opened it" from "I opened a copy you
' cannot keep", which is this project's signature failure one layer up from the
' values it has already been caught by.
'
' A separate function rather than a refusal inside OpenOrGetWorkbook, because
' read-only is legitimate for the read-only paths (Preview Sync opens the
' register to compare and never writes). The caller that is about to WRITE is
' the one that has to ask.
Public Function WriteBlockedReason(wb As Object) As String
    If wb Is Nothing Then
        WriteBlockedReason = "The workbook could not be opened at all."
        Exit Function
    End If

    Dim ro As Boolean
    ro = False
    On Error Resume Next
    ro = wb.ReadOnly
    On Error GoTo 0

    If ro Then
        WriteBlockedReason = _
            "THE REGISTER IS OPEN READ-ONLY, so nothing can be saved into it." & vbCrLf & vbCrLf & _
            wb.FullName & vbCrLf & vbCrLf & _
            "Usually this means it is already open somewhere else -- another " & _
            "Excel window, another machine, or a colleague. Close it there, then " & _
            "try again."
        Exit Function
    End If

    WriteBlockedReason = MacroEnabledWarning(wb.FullName)
End Function

' Creates `path` as a fresh, empty workbook if nothing exists there yet --
' the first-onboarding-on-this-deck case, where there is no paired workbook
' to open. Returns Nothing (never raises) if the path's containing folder
' doesn't exist or SaveAs otherwise fails.
Public Function CreateWorkbook(path As String) As Object
    ' REFUSES AN EXISTING FILE. Defence in depth, added 2026-08-13.
    '
    ' This function does `Workbooks.Add` then `SaveAs path`, which writes a
    ' BLANK workbook over whatever is already there. Its only caller reached it
    ' by asking a person for a path, and a person naming a path they already
    ' know is usually naming a file that already exists -- so the dangerous case
    ' was also the likeliest one. That caller now opens instead of creating, but
    ' the primitive should not be capable of it either: the next caller will not
    ' know, and the cost of being wrong here is a quarter's work.
    '
    ' Nothing in this codebase legitimately overwrites a workbook. A caller that
    ' one day needs to must say so explicitly rather than get it by omission.
    Dim cwfso As Object
    Set cwfso = CreateObject("Scripting.FileSystemObject")
    If cwfso.FileExists(path) Then
        Err.Raise vbObjectError + 515, "WorkbookBridge.CreateWorkbook", _
            "There is already a file at " & path & ". Refusing to create over it -- " & _
            "creating writes an EMPTY workbook and would destroy whatever is there. " & _
            "Open it instead, or choose a different name."
    End If

    Dim xl As Object
    Set xl = GetExcelApp()

    Dim wb As Object
    Set wb = xl.Workbooks.Add()

    ' THE ERROR WAS SWALLOWED ENTIRELY -- no Err check at all, so a failed
    ' SaveAs let onboarding carry on as though the paired workbook existed.
    On Error Resume Next
    wb.SaveAs path
    Dim createErr As String
    If Err.Number <> 0 Then createErr = "Error " & Err.Number & ": " & Err.Description
    Err.Clear
    On Error GoTo 0

    Dim cfso As Object
    Set cfso = CreateObject("Scripting.FileSystemObject")
    If Not cfso.FileExists(path) Then
        Err.Raise vbObjectError + 514, "WorkbookBridge.CreateWorkbookAt", _
            "Could not create the workbook at " & path & _
            IIf(createErr = "", "", " (" & createErr & ")") & _
            ". Nothing downstream can rely on a workbook that is not there."
    End If

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
' A MODAL IS THE WRONG CONTAINER FOR A RUN REPORT.
'
' Rohan, twice on 2026-08-08: "illegible, too long, and the user has no idea what
' is going on", then "still pretty hard to understand". Shortening the wording did
' not fix it, because the problem is not the wording -- it is that a dialog you
' must dismiss to continue is being used to deliver a page of detail you cannot
' scroll, copy, or come back to. MsgBox also truncates near 1024 characters, so
' the longer the report grows the more of it silently disappears.
'
' So the detail goes on a sheet, where it can be read at leisure, kept, and
' compared with the last run; the dialog keeps only what a person needs in the
' three seconds before they click OK.
' Saves the workbook and CONFIRMS the file changed, or says why not.
'
' Same defect, same day, same evidence as DeckRegistry.SaveDeckVerified: the
' sync path wrote a review sheet, told the user it had been "refreshed to match
' the deck as it is now", and the saved workbook contained no such sheet. The
' review the user is being asked to work from existed only on screen.
' Freeze the quarter being rolled out of, as its own file beside the register.
'
' THE FIRST HALF OF FILE-PER-QUARTER, AND DELIBERATELY THE HALF THAT CANNOT
' DESTROY ANYTHING. This only ever CREATES a file: it never edits, prunes or
' deletes, and it refuses rather than overwriting an archive that already exists.
' The pruning half -- dropping the old period's rows from the live register and
' retiring ParkSheetCopy -- is not built, and must not be until it has tests and a
' run at the keyboard. "Well-tested unit, unproven by a person" is the shape that
' wiped 43 approve ticks.
'
' WHY THIS IS THE FOUNDATION. Rohan's ruling is that the archive IS last quarter's
' file (DOCUMENT-MAP decision 6). Today that is aspirational: every quarter is
' stacked in one workbook, so the rollover clear fires in the only copy, and
' ParkSheetCopy exists solely to survive that. 26 park sheets in a 59-sheet
' workbook, growing 13 a quarter, is the cost. Once a real frozen file exists per
' quarter, the parking, the clear-in-place, and the whole class of defect where a
' rebuild lands on work worth keeping all become unnecessary.
'
' NOT A GATE YET, AND THAT IS A DECISION. If this fails, the caller reports it and
' CONTINUES, because today's roll-forward only APPENDS rows -- nothing is lost when
' the archive is missing. THE MOMENT THE PRUNE LANDS THIS MUST BECOME A HARD GATE:
' pruning without a verified archive is the destructive step this exists to make
' safe. Whoever builds the prune changes that here, and says so in the caller.
'
' Returns "" when the archive is confirmed on disk, otherwise what went wrong.
Public Function ArchiveWorkbookForPeriod(wb As Object, period As String) As String
    On Error GoTo Failed

    If Trim$(period) = "" Then
        ArchiveWorkbookForPeriod = "No period was named, so the archive could not be named either."
        Exit Function
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim livePath As String
    livePath = wb.FullName
    If LCase$(Left$(livePath, 4)) = "http" Then
        ArchiveWorkbookForPeriod = "This register is cloud-hosted, so a local archive path " & _
            "cannot be worked out from " & livePath & ". Nothing was written."
        Exit Function
    End If

    Dim folder As String, baseName As String, ext As String
    folder = fso.GetParentFolderName(livePath)
    baseName = fso.GetBaseName(livePath)
    ext = fso.GetExtensionName(livePath)
    If ext = "" Then ext = "xlsx"

    Dim target As String
    target = fso.BuildPath(folder, baseName & "-" & period & "." & ext)

    ' REFUSE, DO NOT OVERWRITE. An existing archive is a previous quarter's frozen
    ' record; silently replacing it would destroy exactly the thing this function
    ' exists to protect, and it would look like success.
    If fso.FileExists(target) Then
        ArchiveWorkbookForPeriod = "An archive for " & period & " already exists:" & vbCrLf & _
            target & vbCrLf & vbCrLf & "Nothing was written. Move or rename it if you " & _
            "genuinely want a fresh one."
        Exit Function
    End If

    ' SaveCopyAs, NOT SaveAs: SaveAs would re-point the OPEN workbook at the archive,
    ' so every subsequent write this session -- including the roll-forward about to
    ' run -- would land in the frozen file and not in the live register. The deck's
    ' pairing would still name the live one. That is a silent split, and it is the
    ' single worst thing this function could do.
    wb.SaveCopyAs target

    ' THE FILE HAS THE ONLY WORD. SaveCopyAs returning quietly is not evidence, and
    ' this project has been caught believing an API's self-report more than once.
    If Not fso.FileExists(target) Then
        ArchiveWorkbookForPeriod = "The archive was requested and did not appear on disk:" & _
            vbCrLf & target
        Exit Function
    End If
    If fso.GetFile(target).Size = 0 Then
        ArchiveWorkbookForPeriod = "The archive was written but is empty:" & vbCrLf & target
        Exit Function
    End If

    Exit Function

Failed:
    ArchiveWorkbookForPeriod = "Could not archive " & period & "." & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description
End Function

Public Function SaveWorkbookVerified(wb As Object) As String
    On Error GoTo Failed
    If wb Is Nothing Then Exit Function

    Dim path As String
    path = wb.FullName

    ' SAME DEFECT AS DeckRegistry.SaveDeckVerified, AND FOUND THE SAME WAY -- by
    ' reading the Run Log this function wrote, which said "This workbook has never
    ' been saved to a file: https://d.docs.live.net/...", printing the URL inside
    ' the sentence denying a file existed. FileSystemObject answers False for a URL
    ' rather than raising, so the guard fell through and wb.Save was NEVER CALLED.
    ' Fixed here as a class rather than only where it was noticed: this one has
    ' NINE call sites against the deck's two, so it is the larger half.
    Dim checkPath As String
    checkPath = DeckRegistry.LocalPathForUrl(path)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Attempt the save even when it cannot be verified, and say so plainly -- the
    ' third state. Skipping the write was the whole defect; reporting a write we
    ' did not witness would be the next one.
    If checkPath = "" Or Not fso.FileExists(checkPath) Then
        On Error Resume Next
        wb.Save
        Dim blindErr As String
        If Err.Number <> 0 Then blindErr = "Error " & Err.Number & ": " & Err.Description
        Err.Clear
        On Error GoTo 0

        SaveWorkbookVerified = "THE SAVE COULD NOT BE VERIFIED." & vbCrLf & vbCrLf & _
            path & vbCrLf & vbCrLf & _
            "The workbook was told to save, but its file could not be located on " & _
            "this PC to confirm the bytes moved. Check it yourself before closing " & _
            "Excel." & IIf(blindErr = "", "", vbCrLf & vbCrLf & blindErr)
        Exit Function
    End If

    Dim before As Date
    before = fso.GetFile(checkPath).DateLastModified

    On Error Resume Next
    wb.Save
    Dim saveErr As String
    If Err.Number <> 0 Then saveErr = "Error " & Err.Number & ": " & Err.Description
    Err.Clear
    On Error GoTo 0

    If fso.GetFile(checkPath).DateLastModified > before Then Exit Function      ' "" = saved

    SaveWorkbookVerified = "THE WORKBOOK WAS NOT SAVED." & vbCrLf & vbCrLf & _
        path & vbCrLf & vbCrLf & _
        "Anything written to it -- including the review sheet -- is on screen " & _
        "only. Save it yourself before closing Excel." & _
        IIf(saveErr = "", "", vbCrLf & vbCrLf & saveErr)
    Exit Function

Failed:
    SaveWorkbookVerified = "Could not save the workbook." & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description
End Function

Public Sub WriteRunLog(wb As Object, header As String, body As String)
    On Error GoTo Failed

    Dim ws As Object
    Set ws = GetOrAddWorksheet(wb, RUN_LOG_SHEET_NAME)

    ' REPLACED each run, not appended. A log that grows forever becomes a file
    ' nobody opens, and the question this answers is always "what just happened",
    ' never "what happened in March".
    ws.Cells.Clear

    ws.Cells(1, 1).Value = header
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(2, 1).Value = "Run at " & Format(Now, "yyyy-mm-dd hh:nn:ss")

    Dim lines As Variant
    lines = Split(Replace(body, vbCrLf, vbLf), vbLf)

    ' EVERY LINE IS FORCED TO TEXT, and that is not cosmetic.
    '
    ' A cell value beginning with "=" is a FORMULA to Excel. Every report this
    ' log receives opens with a banner like:
    '
    '     === PREVIEW (nothing written): project-status ===
    '
    ' so Excel tried to parse "== PREVIEW ..." as an expression, raised, and the
    ' handler below abandoned the rest of the log -- silently, by design, because
    ' a log must never stop the run it describes. Net effect: the body was lost
    ' on EVERY Preview Sync and Sync Now, while the dialog said "the full
    ' before-and-after is on the 'Run Log' sheet, untruncated".
    '
    ' Found 2026-08-09: the sheet held five lines (title, timestamp, and the
    ' unsaved-workbook note) against a report of 38 changes -- and the count in
    ' the dialog is derived from that same report, so the detail provably
    ' existed.
    '
    ' The leading apostrophe is Excel's own "treat as text" marker and is not
    ' stored in the value. The column is also formatted as text first, so a line
    ' like "-1.5" or "1/2" cannot be coerced into a number or a date either.
    ws.Columns(1).NumberFormat = "@"

    Dim i As Long, r As Long
    r = 4
    For i = LBound(lines) To UBound(lines)
        ws.Cells(r, 1).Value = "'" & CStr(lines(i))
        r = r + 1
    Next i

    ws.Columns(1).ColumnWidth = 110
    ws.Cells.Font.Size = 9
    Exit Sub

Failed:
    ' A run log that cannot be written must never stop the run it is describing.
End Sub

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
        ' COLUMN LETTERS ARE DERIVED, NEVER TYPED. This sentence said "type new
        ' wording in D, put Y in E" -- layout 3 letters, two layouts out of date,
        ' pointing a person at the SOURCES and AI DRAFT columns. A sentence cannot
        ' fail a test; the constants can.
        DescribeSheet = "Drafting sheet for " & Mid(sheetName, 5) & _
            ". Read column " & Chr$(64 + Drafting.COL_D_CURRENT) & _
            ", type new wording in " & Chr$(64 + Drafting.COL_D_SUBMIT) & _
            ", put Y in " & Chr$(64 + Drafting.COL_D_APPROVED) & _
            ". Instructions are on the sheet."
    ElseIf Left(sheetName, Len("Review ")) = "Review " _
        Or Left(sheetName, Len("Sync Review")) = "Sync Review" Then
        DescribeSheet = "Every change waiting to be approved before it reaches a slide. " & _
            "Tick what you agree with, then press '" & CommandBarUI.CAP_PUT_ON_SLIDES & "' again."
    ElseIf sheetName = SYNC_LOG_SHEET_NAME Then
        DescribeSheet = "What was written to slides, and when. Written as it happens, so a run " & _
            "that dies halfway still leaves a record."
    ElseIf sheetName = FieldSpec.SPEC_SHEET_NAME Then
        DescribeSheet = "How each field should be WRITTEN -- purpose, voice, length, and what " & _
            "not to do. Edit this to change the instructions the AI is given. Yours, not the tool's."
    ElseIf sheetName = Sources.SOURCES_SHEET_NAME Then
        DescribeSheet = "WHERE THE WORDS CAME FROM. One row per source, referenced by ID from " & _
            "column " & Chr$(64 + Drafting.COL_D_SOURCES) & " of a drafting sheet. " & _
            "Point at documents; do not paste them in here."
    ElseIf sheetName = REGISTER_SHEET_NAME Then
        ' Described the LONG register -- one row per project/field/quarter with an
        ' approval column -- a model retired 2026-08-03. The wide sheet is one row
        ' per SLIDE per quarter, one COLUMN per field, and carries no approval state.
        DescribeSheet = "THE RECORD. One row per slide per quarter, one column per field. " & _
            "Approval lives on the review sheet, not here. Everything else in this " & _
            "workbook feeds it or reads it."
    Else
        DescribeSheet = "(not created by this tool)"
    End If
End Function

Public Function LifespanOf(sheetName As String) As String
    If sheetName = REGISTER_SHEET_NAME Then
        LifespanOf = "PERMANENT -- grows each quarter"
    ElseIf Left(sheetName, 4) = "TPL_" Then
        LifespanOf = "Rebuilt each drafting round"
    ElseIf Left(sheetName, 11) = "Sync Review" Then
        LifespanOf = "One per run, then consumed"
    ElseIf sheetName = SYNC_LOG_SHEET_NAME Then
        LifespanOf = "Append-only history"
    ElseIf sheetName = FieldSpec.SPEC_SHEET_NAME Then
        LifespanOf = "PERMANENT -- edit it freely"
    ElseIf sheetName = Sources.SOURCES_SHEET_NAME Then
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
'
' TWO LIVE BUGS FOUND HERE BY REVIEW, 2026-08-01, both the house failure mode:
'
'   Left(sheetName, 13) = "Template Audit"   -- that literal is FOURTEEN
'   characters, so the comparison could never be True. An always-false guard
'   shipped in production, reading as care taken. Exactly what the project's own
'   zettel warns about, in the list written to prevent this class of mistake.
'
'   "Field Discovery" was absent. DiscoverUI writes that sheet INTO the paired
'   register workbook, so RegisterOrFirstDataSheet could hand a caller the
'   discovery grid and call it the register -- the very defect this denylist
'   exists to prevent, reintroduced by a module written seven hours after the
'   list.
'
' Both were possible because this is a hand-maintained list duplicating names
' that already exist as public constants. It now uses the constants, so a
' renamed sheet cannot drift out of step with the code that names it. The two
' prefix rules stay literal because there is no constant for a prefix, and their
' lengths are now derived with Len() rather than counted by hand -- which is how
' the 13-versus-14 error happened in the first place.
' THE REVIEW-SHEET PREFIX WAS STALE FOR THREE WEEKS, AND IT MATTERED.
' This matched only "Sync Review" while ReviewQueue.ReviewSheetNameFor has
' produced "Review <type>-<hash>" since the 3de4be8 rename -- so the ONE sheet
' a person actually works in was reported as "(not created by this tool)" and
' given lifespan "unknown" on the START HERE sheet. Rohan ticked 43 approvals
' into an unrecognised sheet on 2026-08-14.
'
' REGISTER AND RUN LOG WERE MISSING ENTIRELY, while LifespanOf two functions
' below has always known Register perfectly well. Two functions in one module
' disagreeing about the most important sheet in the workbook.
'
' Both prefixes are matched: "Review " is current, "Sync Review" is the legacy
' name and any workbook onboarded before the rename still carries one.
Public Function IsToolOwnedSheet(sheetName As String) As Boolean
    If sheetName = INDEX_SHEET_NAME Then IsToolOwnedSheet = True
    If sheetName = REGISTER_SHEET_NAME Then IsToolOwnedSheet = True
    If sheetName = FieldSpec.SPEC_SHEET_NAME Then IsToolOwnedSheet = True
    If sheetName = Sources.SOURCES_SHEET_NAME Then IsToolOwnedSheet = True
    If sheetName = DiscoverUI.DISCOVERY_SHEET_NAME Then IsToolOwnedSheet = True
    If sheetName = TemplateAudit.AUDIT_SHEET_NAME Then IsToolOwnedSheet = True
    If sheetName = RUN_LOG_SHEET_NAME Then IsToolOwnedSheet = True
    If sheetName = SYNC_LOG_SHEET_NAME Then IsToolOwnedSheet = True
    If Left(sheetName, Len("TPL_")) = "TPL_" Then IsToolOwnedSheet = True
    If Left(sheetName, Len("SRC_")) = "SRC_" Then IsToolOwnedSheet = True
    If Left(sheetName, Len("SAVED ")) = "SAVED " Then IsToolOwnedSheet = True
    If Left(sheetName, Len("Review ")) = "Review " Then IsToolOwnedSheet = True
    If Left(sheetName, Len("Sync Review")) = "Sync Review" Then IsToolOwnedSheet = True
End Function

' A warning if this path is a macro-enabled Office file, or "" if it is fine.
'
' MACRO-ENABLED FILES CANNOT BE SAVED ON A MANAGED MACHINE, AND FAIL SILENTLY.
'
' Rohan, 2026-08-01: his deck stopped saving with no message and no dialog --
' Ctrl+S simply did nothing. Cause: a rescue macro had been imported into the
' presentation's own VBA project, which made a .pptx macro-bearing, and company
' policy blocks macro-enabled DOCUMENTS (while permitting trusted add-ins). It
' cost 41 minutes of unsaved work and a diagnosis that started in the wrong
' place. He raised it again the same evening -- "remember the pptx vs pptm thing
' too (pptm won't save on work machine)" -- because it is a standing constraint,
' not a past incident.
'
' Nothing in this tool ever CREATES a macro-enabled file: there is no .pptm or
' .xlsm format constant anywhere in the source, and every path it constructs
' ends .xlsx. The exposure is writing to a file that already is one.
'
' Warns rather than refuses. On the author's personal machine these files save
' perfectly well, and a tool that refused to touch them there would be wrong.
' The person knows which machine they are on; the tool does not.
Public Function MacroEnabledWarning(path As String) As String
    If path = "" Then Exit Function

    Dim p As String
    p = LCase(Trim(path))

    Dim ext As String
    Dim dotAt As Long
    dotAt = InStrRev(p, ".")
    If dotAt = 0 Then Exit Function
    ext = Mid(p, dotAt)

    ' Length derived, never counted by hand -- see IsToolOwnedSheet, where a
    ' hand-counted 13-versus-14 shipped an always-false guard.
    Select Case ext
        Case ".pptm", ".ppsm", ".potm", ".xlsm", ".xlsb", ".xltm", ".docm"
            MacroEnabledWarning = _
                "WARNING: this is a MACRO-ENABLED file (" & ext & ")." & vbCrLf & _
                "On a managed work machine, saving it may be blocked by policy -- " & _
                "and the block is SILENT: no message, and Ctrl+S simply does nothing." & vbCrLf & vbCrLf & _
                "If that happens, use File > Save As and explicitly pick the " & _
                "non-macro format (.pptx / .xlsx), to a local folder."
    End Select
End Function

' TABS IN THE ORDER YOU USE THEM, not the order they were created.
'
' Rohan, 2026-08-10. The workbook had grown by accretion: two drafting sheets
' stranded among the reference sheets because they were made first, the other
' three at the end because they were made last, and the logs in the middle.
' Nothing grouped, and with five colour-coded drafting tabs the strip is now the
' index -- an index in creation order is not one.
'
'   START HERE   where am I
'   Field Spec   the recipes -- how each field should be written
'   Sources      the evidence available
'   TPL_...      where you work, in Field Spec order so the colours run in
'                sequence rather than scattering
'   Review ...   the gate
'   <the rest>   the register and anything else
'   Run Log      what happened, last, because it is read after the fact
'   Sync Log
'
' THE REGISTER SITS AFTER THE DRAFTING SHEETS ON PURPOSE. It is the output --
' derived, not edited. Putting it second would invite treating it as the working
' surface, which is the mistake column C already guards against one level down.
'
' draftOrder is a vbLf-delimited list of drafting sheet names in the order the
' caller built them; only the caller knows the Field Spec order.
'
' Never creates or deletes. A name that is not present is skipped, so this is
' safe on a workbook that has only some of them.
Public Sub ArrangeTabs(wb As Object, draftOrder As String)
    ' TAB ORDER FOLLOWS THE LIFECYCLE OF A QUARTER, not the order sheets happened
    ' to be created in. Rohan asked for logical numbering across the workbooks;
    ' this is that ordering expressed as POSITION rather than as renamed tabs.
    '
    ' WHY NOT NUMBER THE NAMES. Sheet names are this tool's addressing mechanism
    ' -- nine constants, dozens of literals, plus the TPL_, "Review " and "SAVED "
    ' prefix matches. Renaming to "01_FIELD_SPEC" and the like means a code change
    ' AND a migration of a live workbook holding drafted work. That is a
    ' deliberate migration to plan, not a tidy-up to slip into a session, and it
    ' is written up in NEXT-SESSION.md. Ordering costs nothing and cannot break a
    ' lookup, so it goes in now.
    '
    ' The sequence, and the reason for it:
    '   1  START HERE     where a person begins, and the readiness checklist
    '   2  Field Spec     what fields exist at all -- configuration before data
    '   3  Sources        the evidence values are allowed to cite
    '   4  SRC_*          harvested source data, feeding the above
    '   5  Register       THE DATA. The reason the workbook exists.
    '   6  TPL_*          where a person actually works, in Field Spec order
    '   7  Review *       the approval gate, between work and the deck
    '   8  diagnostics    Field Discovery, Template Audit -- read when something
    '                     looks wrong, not part of the normal path
    '   9  logs           the audit trail, appended to and rarely read forward
    '  10  SAVED *        parked archives, last, so they cannot be mistaken for
    '                     the live sheet they were copied from
    '
    ' REGISTER WAS PREVIOUSLY UNPLACED. It fell into "everything else in its
    ' current order" along with the diagnostics -- so the single most important
    ' sheet in the workbook sat wherever it happened to land. That is the defect
    ' this ordering fixes, more than any cosmetic gain.
    Dim wanted As String
    wanted = Readiness.READY_SHEET_NAME & vbLf & _
             FieldSpec.SPEC_SHEET_NAME & vbLf & _
             Sources.SOURCES_SHEET_NAME

    ' Source sheets, by prefix -- there is no constant per SRC_ sheet.
    Dim ws As Object
    For Each ws In wb.Worksheets
        If Left(ws.Name, 4) = "SRC_" Then wanted = wanted & vbLf & ws.Name
    Next ws

    wanted = wanted & vbLf & REGISTER_SHEET_NAME

    If Trim(draftOrder) <> "" Then wanted = wanted & vbLf & draftOrder

    ' Review sheets next, whatever they are called -- the name carries a slide
    ' type and a hash, so it is matched by prefix rather than named.
    For Each ws In wb.Worksheets
        If Left(ws.Name, 7) = "Review " Then wanted = wanted & vbLf & ws.Name
    Next ws

    ' Diagnostics: read when something looks wrong, not on the normal path.
    wanted = wanted & vbLf & DiscoverUI.DISCOVERY_SHEET_NAME & _
                      vbLf & TemplateAudit.AUDIT_SHEET_NAME

    ' Then anything unrecognised, in its current order -- a person's own sheet
    ' keeps its place rather than being shuffled by a tool that does not own it.
    ' Logs and parked archives are held back deliberately.
    For Each ws In wb.Worksheets
        If InStr(vbLf & wanted & vbLf, vbLf & ws.Name & vbLf) = 0 Then
            If ws.Name <> RUN_LOG_SHEET_NAME And ws.Name <> SYNC_LOG_SHEET_NAME _
               And Left(ws.Name, 6) <> "SAVED " Then
                wanted = wanted & vbLf & ws.Name
            End If
        End If
    Next ws

    ' Logs, then parked archives absolutely last. A "SAVED ..." copy sitting
    ' beside the live sheet it was taken from is a second place to type, and
    ' typing there is silent -- publish reads the live sheet only.
    wanted = wanted & vbLf & RUN_LOG_SHEET_NAME & vbLf & SYNC_LOG_SHEET_NAME
    For Each ws In wb.Worksheets
        If Left(ws.Name, 6) = "SAVED " Then wanted = wanted & vbLf & ws.Name
    Next ws

    Dim names() As String
    names = Split(wanted, vbLf)

    Dim placed As Long
    Dim i As Long
    For i = LBound(names) To UBound(names)
        Dim nm As String
        nm = Trim(names(i))
        If nm <> "" Then
            If WorksheetExists(wb, nm) Then
                On Error Resume Next
                If placed = 0 Then
                    wb.Worksheets(nm).Move Before:=wb.Worksheets(1)
                Else
                    wb.Worksheets(nm).Move After:=wb.Worksheets(placed)
                End If
                If Err.Number = 0 Then placed = placed + 1
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next i
End Sub
