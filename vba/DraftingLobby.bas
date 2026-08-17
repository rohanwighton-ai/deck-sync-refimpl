Attribute VB_Name = "DraftingLobby"
Option Explicit

' -----------------------------------------------------------------------
' THE DRAFTING LOBBY -- see LOBBY-DESIGN.md for the full design and why.
' -----------------------------------------------------------------------
'
' Rohan, 2026-08-16: "it's more like the work being pinned on a board by the
' author, and the crawler just looking at the board, not every author's
' desk... authorship pins on board next to other authors' notes for the
' period... it's pinned as part of authoring."
'
' ONE SHARED SHEET, NOT PER-FIELD. Every drafting sheet's APPROVE tick pins
' one row here. Publish reads ONLY this sheet -- never the 13 (or however
' many) drafting sheets directly -- which is the whole fix for the crawl
' (FIX-LIST item U) and, combined with pre-ticking the register-to-slide
' queue (LOBBY-DESIGN.md section 5), the two-press pattern.
'
' A FRESH RECOMPUTE PATH EXISTS AND IS NOT OPTIONAL (BuildLobbyFromScratch,
' below). The pin-on-tick mechanism (AppEvents.bas) is the fast path, not the
' only path -- a person hand-editing the workbook at work, with no Claude, no
' Python and no macro running, pins nothing, and this project already designs
' for that machine. The cold-start crawl is the same full read
' PublishAllDraftedFields already did before this module existed; it is not
' new cost, just no longer the ONLY path.

Public Const LOBBY_SHEET_NAME As String = "Drafting Lobby"

Private Const COL_L_SHEET As Long = 1
Private Const COL_L_ROW As Long = 2
Private Const COL_L_FIELDID As Long = 3
Private Const COL_L_ENTITY As Long = 4
Private Const COL_L_TIMESTAMP As Long = 5

Private Const LOBBY_HEADER_ROW As Long = 1
Private Const LOBBY_FIRST_ROW As Long = 2

' Numeric, not the named constant xlUp -- same reason as ExcelOutput.bas's own
' XL_UP: this module is PowerPoint-hosted, where the named form does not
' resolve.
Private Const XL_UP As Long = -4162

Public Type LobbyEntry
    SheetName As String
    Row As Long
    FieldId As String
    EntityKey As String
End Type

' HELD AT MODULE LEVEL, NOT LOCAL, OR IT WOULD STOP FIRING. A WithEvents
' object only sinks events for as long as something keeps a live reference
' to it; a local variable in EnsureWatching would be destroyed the instant
' that Sub returned, and the pin mechanism would silently do nothing from
' the very next keystroke. This is genuinely the only place in the codebase
' this pattern is needed (the only class module) -- named here so the next
' person does not "clean up" what looks like an unused module-level Object.
Private mAppEvents As AppEvents

' Wires the pin mechanism to whichever Excel instance actually owns `wb`.
' Idempotent and cheap to call on every resolve -- called from
' WorkbookBridge.OpenOrGetWorkbook, the one place every path through this
' add-in already goes through to reach the register workbook (LOBBY-DESIGN.
' md section 8). Re-pointing `.App` to the SAME instance it already watches
' is a harmless no-op; the check exists only to cover a second Excel
' instance (or the same workbook reopened in a fresh session) actually
' being different from whatever was last watched.
Public Sub EnsureWatching(wb As Object)
    If mAppEvents Is Nothing Then Set mAppEvents = New AppEvents

    ' Watch()/IsWatching(), not a direct property -- AppEvents.mApp is kept
    ' Private by design (see AppEvents.cls's own header for why, and for the
    ' real bug that first made this look like a WithEvents visibility rule
    ' rather than what it actually was: a mis-imported .cls file).
    If Not mAppEvents.IsWatching(wb.Application) Then mAppEvents.Watch wb.Application
End Sub

' Reverse of Drafting.DraftSheetNameFor -- "" if `sheetName` is not a real
' drafting sheet for any currently-declared Prose field. This is the whole
' safety net that lets AppEvents watch the entire Application rather than
' one workbook: anything that fails this check (every unrelated sheet on
' the machine, every non-drafting sheet in this workbook) exits in one
' comparison per field, cheap enough to run on every keystroke.
' FIX-LIST item AL. This function's own header comment above claims "exits
' in one comparison" for anything that is not a real drafting sheet name --
' it didn't. Every call used to run `DraftingUI.ProseFields(wb)` FIRST
' (`WorksheetExists` -- a For Each over every worksheet in the workbook --
' then a full scan of the Field Spec sheet, 2 cell reads per row) before the
' name comparison the comment describes ever ran. Found reading this
' function's actual first line, not trusting what the comment above it
' claims. `mApp_SheetChange` (`AppEvents.cls`) calls this on EVERY sheet
' change in EVERY open workbook -- at ~100-170 COM calls per miss, that
' ~0.3-1.2s tax landed on every cell written anywhere the tool wasn't
' already disabling events (item AL's other half).
'
' Every drafting sheet name is `Drafting.DraftSheetNameFor`'s own output:
' `"TPL_" & fieldId`, sanitized. A sheet whose name does not start with
' "TPL_" cannot possibly be one, and that is checkable with zero COM calls,
' before the workbook is ever touched -- making this comment's claim
' actually true instead of aspirational.
Public Function FieldIdForSheet(wb As Object, sheetName As String) As String
    ' Case-insensitive, matching the StrComp(..., vbTextCompare) the loop
    ' below already uses -- a behaviour-preserving fast path, not a stricter
    ' one.
    If StrComp(Left$(sheetName, 4), "TPL_", vbTextCompare) <> 0 Then Exit Function

    Dim fields As String
    fields = DraftingUI.ProseFields(wb)
    If Trim(fields) = "" Then Exit Function

    Dim parts() As String
    parts = Split(fields, ",")

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim fieldId As String
        fieldId = Trim(parts(i))
        If fieldId <> "" Then
            If StrComp(Drafting.DraftSheetNameFor(fieldId), sheetName, vbTextCompare) = 0 Then
                FieldIdForSheet = fieldId
                Exit Function
            End If
        End If
    Next i
End Function

Private Function EnsureLobbySheet(wb As Object) As Object
    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(wb, LOBBY_SHEET_NAME)

    If Trim(CStr(ws.Cells(LOBBY_HEADER_ROW, COL_L_SHEET).Value)) = "" Then
        ws.Cells(LOBBY_HEADER_ROW, COL_L_SHEET).Value = "SheetName"
        ws.Cells(LOBBY_HEADER_ROW, COL_L_ROW).Value = "Row"
        ws.Cells(LOBBY_HEADER_ROW, COL_L_FIELDID).Value = "FieldID"
        ws.Cells(LOBBY_HEADER_ROW, COL_L_ENTITY).Value = "EntityCode"
        ws.Cells(LOBBY_HEADER_ROW, COL_L_TIMESTAMP).Value = "Pinned"
        ws.Rows(LOBBY_HEADER_ROW).Font.Bold = True
    End If

    Set EnsureLobbySheet = ws
End Function

' FIX-LIST item AB: this used to be a VBA loop reading one cell per row from
' LOBBY_FIRST_ROW, the exact same shape item W fixed in AppendLogLine -- and
' worse here, because THREE separate call sites lean on this one function
' (FindLobbyRow's own bound, FindLobbyRow's comparison loop via that bound,
' and PinToLobby's own direct call), so every pin during a from-scratch
' rebuild paid for it three times over. End(XL_UP) is one native COM call
' instead of up to n. The header row always carries "FieldID" in this exact
' column (EnsureLobbySheet, above), so End(XL_UP) lands on row 1 -- same
' "nothing pinned yet" answer the old loop gave via its own r-1 arithmetic.
Private Function LastLobbyRow(ws As Object) As Long
    LastLobbyRow = ws.Cells(ws.Rows.Count, COL_L_FIELDID).End(XL_UP).Row
End Function

' Finds an existing pin for this exact (sheet, field, entity), or 0.
'
' KEYED ON SHEET+FIELD+ENTITY, NOT ON ROW NUMBER. A row number is where the
' field happens to sit today; the identity of "this field, for this project"
' is what a second tick on the same row should update in place rather than
' duplicate.
Private Function FindLobbyRow(ws As Object, sheetName As String, fieldId As String, _
                              entityKey As String) As Long
    Dim last As Long
    last = LastLobbyRow(ws)

    Dim r As Long
    For r = LOBBY_FIRST_ROW To last
        If StrComp(CStr(ws.Cells(r, COL_L_SHEET).Value), sheetName, vbTextCompare) = 0 And _
           StrComp(CStr(ws.Cells(r, COL_L_FIELDID).Value), fieldId, vbTextCompare) = 0 And _
           StrComp(CStr(ws.Cells(r, COL_L_ENTITY).Value), entityKey, vbTextCompare) = 0 Then
            FindLobbyRow = r
            Exit Function
        End If
    Next r
End Function

' THE PIN. Called from two places: AppEvents' tick-watcher (the fast path,
' one row at a time, as it happens) and BuildLobbyFromScratch (the cold-start
' path, all at once). Idempotent -- pinning the same (sheet, field, entity)
' twice updates the timestamp in place rather than growing the sheet.
Public Sub PinToLobby(wb As Object, sheetName As String, r As Long, fieldId As String, _
                      entityKey As String)
    Dim ws As Object
    Set ws = EnsureLobbySheet(wb)

    Dim existing As Long
    existing = FindLobbyRow(ws, sheetName, fieldId, entityKey)

    Dim targetRow As Long
    If existing > 0 Then
        targetRow = existing
    Else
        targetRow = LastLobbyRow(ws) + 1
        If targetRow < LOBBY_FIRST_ROW Then targetRow = LOBBY_FIRST_ROW
    End If

    ws.Cells(targetRow, COL_L_SHEET).Value = sheetName
    ws.Cells(targetRow, COL_L_ROW).Value = r
    ws.Cells(targetRow, COL_L_FIELDID).Value = fieldId
    ws.Cells(targetRow, COL_L_ENTITY).Value = entityKey
    ws.Cells(targetRow, COL_L_TIMESTAMP).Value = Now
End Sub

' Removes one entry after PublishDraftsForField has actually published it.
' Not found is not an error -- a person may have already cleared it, or the
' Lobby may have been rebuilt from scratch since the pin landed.
Public Sub ClearLobbyEntry(wb As Object, sheetName As String, fieldId As String, _
                           entityKey As String)
    If Not WorkbookBridge.WorksheetExists(wb, LOBBY_SHEET_NAME) Then Exit Sub
    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(wb, LOBBY_SHEET_NAME)

    Dim r As Long
    r = FindLobbyRow(ws, sheetName, fieldId, entityKey)
    If r > 0 Then ws.Rows(r).Delete
End Sub

' Everything currently pinned. Publish reads this and ONLY this -- never the
' drafting sheets directly -- which is the fix for the crawl.
'
' AN ARRAY OF THE UDT, NOT A Collection OR Dictionary OF IT. Both of those
' are late-bound/Variant-based internally and cannot hold a VBA Type --
' compile error, not a runtime one (AGENTS.md, hit four times now). Same
' shape as SyncAction()/ReviewItem() elsewhere in this codebase.
'
' EmptyLobbyEntries() is the (1 To 0) convention every other array-returning
' function here uses for "genuinely nothing" -- ReDim-ing straight to
' (1 To 0) throws at runtime (AGENTS.md), so an empty result is built by
' never entering the loop rather than by asking for a zero-length array.
Public Function ReadLobby(wb As Object) As LobbyEntry()
    Dim out() As LobbyEntry
    Dim n As Long
    n = 0

    If Not WorkbookBridge.WorksheetExists(wb, LOBBY_SHEET_NAME) Then
        ReadLobby = out
        Exit Function
    End If
    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(wb, LOBBY_SHEET_NAME)

    Dim last As Long
    last = LastLobbyRow(ws)

    Dim r As Long
    For r = LOBBY_FIRST_ROW To last
        n = n + 1
        ReDim Preserve out(1 To n)
        out(n).SheetName = CStr(ws.Cells(r, COL_L_SHEET).Value)
        out(n).Row = CLng(ws.Cells(r, COL_L_ROW).Value)
        out(n).FieldId = CStr(ws.Cells(r, COL_L_FIELDID).Value)
        out(n).EntityKey = CStr(ws.Cells(r, COL_L_ENTITY).Value)
    Next r

    ReadLobby = out
End Function

' Count of pending entries without building the array -- for callers (the
' pending-approvals check in RibbonUI) that only need "is there anything
' pending", not the entries themselves.
Public Function LobbyCount(wb As Object) As Long
    If Not WorkbookBridge.WorksheetExists(wb, LOBBY_SHEET_NAME) Then Exit Function
    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(wb, LOBBY_SHEET_NAME)
    LobbyCount = LastLobbyRow(ws) - LOBBY_FIRST_ROW + 1
    If LobbyCount < 0 Then LobbyCount = 0
End Function

' THE COLD-START / REPAIR PATH. The one full crawl this design still needs --
' proven safe to run because it is exactly the read PublishAllDraftedFields
' already did every single press, before this module existed. Not slower than
' today's tool; the fix is that it no longer has to run on every press, only
' here, on demand.
'
' Rebuilds the Lobby from the ACTUAL state of every drafting sheet's APPROVE
' column -- the source of truth is always the sheets, never the Lobby itself.
' Safe to run at any time: a person can call this to repair the Lobby after a
' hand-edit at work, or any time there is doubt about whether a pin was missed.
Public Function BuildLobbyFromScratch(wb As Object) As String
    Dim fields As String
    fields = DraftingUI.ProseFields(wb)
    If Trim(fields) = "" Then
        BuildLobbyFromScratch = "No Prose fields on the Field Spec sheet -- nothing to scan."
        Exit Function
    End If

    ' Fresh sheet each time -- this is a REBUILD from reality, not an
    ' incremental merge. Any pin the sheets no longer support (approve tick
    ' since removed) must not survive a rebuild, the same "never trust a
    ' maintained record over reality" rule that governs the review queue.
    If WorkbookBridge.WorksheetExists(wb, LOBBY_SHEET_NAME) Then
        ' wb.Application, NOT the bare `Application`. This VBA project is
        ' hosted in PowerPoint (COM-add-in-first, per DECISIONS.md); a bare
        ' `Application` reference resolves to PowerPoint.Application, not the
        ' Excel instance that actually owns `wb`. Every other place in this
        ' codebase that needs the workbook's own Application already knows
        ' this (DraftingUI.bas's wb.Application.Visible/.InputBox/.Caption
        ' calls) -- found live 2026-08-16 when this exact line silently
        ' suppressed the WRONG app's alerts, the sheet-delete confirmation
        ' fired on the real (correct) Excel instance regardless, and a
        ' rebuild left a stale pin behind. Caught by the test that proves a
        ' rebuild does not trust its own past content -- see
        ' Test_DraftingLobby_BuildFromScratchFindsOnlyApprovedRows.
        wb.Application.DisplayAlerts = False
        wb.Worksheets(LOBBY_SHEET_NAME).Delete
        wb.Application.DisplayAlerts = True
    End If
    Dim ws As Object
    Set ws = EnsureLobbySheet(wb)

    ' NO FindLobbyRow HERE, DELIBERATELY -- FIX-LIST item AB. The Lobby was
    ' just deleted and recreated above, in THIS SAME CALL, so every pin this
    ' function writes is provably new; going through PinToLobby's shared
    ' existing-pin lookup would scan the (growing) Lobby sheet, 3 cells per
    ' row, for a match that cannot exist. Measured live 2026-08-17: 503s for
    ' 559 rows scanned, 115 pinned -- almost entirely this exact redundant
    ' scan, repeated for every pin as the sheet grew. nextRow is tracked
    ' locally instead, since this function is the only writer of this sheet
    ' for the duration of this call. PinToLobby itself is untouched --
    ' AppEvents' real-time, one-pin-at-a-time path still needs the genuine
    ' existing-pin lookup, and still gets it.
    Dim nextRow As Long
    nextRow = LOBBY_FIRST_ROW

    Dim parts() As String
    parts = Split(fields, ",")

    Dim pinned As Long, scanned As Long
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim fieldId As String
        fieldId = Trim(parts(i))
        If fieldId <> "" Then
            Dim sheetName As String
            sheetName = Drafting.DraftSheetNameFor(fieldId)
            If WorkbookBridge.WorksheetExists(wb, sheetName) Then
                Dim dws As Object
                Set dws = WorkbookBridge.GetOrAddWorksheet(wb, sheetName)

                Dim r As Long
                r = Drafting.DRAFT_FIRST_ROW
                Do While Trim(CStr(dws.Cells(r, Drafting.COL_D_ENTITY).Value)) <> ""
                    scanned = scanned + 1
                    Dim mark As String
                    mark = CStr(dws.Cells(r, Drafting.COL_D_APPROVED).Value)
                    If ReviewQueue.IsApprovalMark(mark) Then
                        Dim entityKey As String
                        entityKey = Trim(CStr(dws.Cells(r, Drafting.COL_D_ENTITY).Value))
                        ws.Cells(nextRow, COL_L_SHEET).Value = sheetName
                        ws.Cells(nextRow, COL_L_ROW).Value = r
                        ws.Cells(nextRow, COL_L_FIELDID).Value = fieldId
                        ws.Cells(nextRow, COL_L_ENTITY).Value = entityKey
                        ws.Cells(nextRow, COL_L_TIMESTAMP).Value = Now
                        nextRow = nextRow + 1
                        pinned = pinned + 1
                    End If
                    r = r + 1
                Loop
            End If
        End If
    Next i

    BuildLobbyFromScratch = "Lobby rebuilt: " & pinned & " approved row(s) pinned, " & _
        scanned & " row(s) scanned across " & (UBound(parts) - LBound(parts) + 1) & " field(s)."
End Function
