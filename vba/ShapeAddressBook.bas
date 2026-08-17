Attribute VB_Name = "ShapeAddressBook"
Option Explicit

' -----------------------------------------------------------------------
' THE SHAPE ADDRESS BOOK -- a persistent, self-healing cache of "which
' shape, by name, answers to this role tag on this slide type".
' -----------------------------------------------------------------------
'
' WHY THIS EXISTS. InjectPrimitive.FindShapeByRoleTag walks every shape on a
' slide (recursing into every group) to find the one tagged with a given
' FieldID, every single time it is called -- twice per queued item (once for
' the dry-run probe, once for the real write). Measured live 2026-08-17 on a
' real 221-item apply run: roughly 4-5 seconds per item, confirmed NOT to be
' screen-redraw (PowerPoint has no ScreenUpdating property -- checked, does
' not exist) and NOT window focus (measured, no stall). A simple one-shape
' test slide did the same operation in under 1ms -- the real cost is walking
' real, content-heavy slides' real shape trees, repeatedly, for shapes whose
' location never actually changes between runs.
'
' Slide.Duplicate (SlideDuplication.bas's own mechanism for creating every
' instance) copies shapes in the same order with the same names as the
' template. Nothing in this codebase renames a shape after that. So "which
' shape answers to FieldID X on a slide of type Y" is the same answer on
' every instance of that type, every time -- worth remembering, not worth
' re-deriving on every single call.
'
' PowerPoint's Shapes.Item(Name) is a genuine fast path, not a marginal one
' -- confirmed against Microsoft's own guidance, not assumed: referencing a
' shape by name avoids the enumeration a lookup by index or a manual walk
' both require.
'
' NO SEPARATE "DISCOVERY" PASS. Rohan, 2026-08-17: "to avoid discovery" --
' the book is not bulk-populated ahead of time. It starts empty and fills
' itself lazily, one field at a time, the first time InjectPrimitive.
' FindShapeByRoleTag is ever asked to find it: cache miss -> full walk (the
' existing, correct, ambiguity-checked mechanism, unchanged) -> the shape's
' real .Name is recorded here before returning. Every subsequent call for
' that (slide type, field) tries the cached name FIRST, verifies its role
' tag still matches (cheap -- one shape, not a walk), and only falls back to
' the full walk again if that check fails. A wrong or stale entry is never
' trusted blindly; it is always verified, and self-heals by being
' overwritten with whatever the full walk finds instead.
'
' WHY A SEPARATE MODULE WITH A MODULE-LEVEL WORKBOOK REFERENCE, NOT A NEW
' PARAMETER ON EVERY INJECTOR FUNCTION. A persistent, cross-session cache
' has to live somewhere in the register workbook -- but InjectField,
' InjectPrimitive, InjectProgressVia, InjectPictureVia, InjectDeviceVia and
' every one of their callers (ApplyApproved, MilestoneDevice, SyncOperations
' .PlanRoutineSync's dry runs, RunSync) take a slide/shape, never a
' workbook. Threading `wb` through that entire family to reach one cache
' would be real, invasive surgery across the whole injector for a lookup
' that only FindShapeByRoleTag itself needs. Same shape as DraftingLobby.bas
' tonight: SetActiveWorkbook is called from the one place every path through
' this add-in already goes through to reach the register
' (WorkbookBridge.OpenOrGetWorkbook), and everything downstream reaches the
' cache through this module instead of carrying a workbook reference itself.
Public Const ADDRESS_BOOK_SHEET_NAME As String = "Shape Address Book"

' FIX-LIST item AR, 2026-08-17/18 night. This book only ever remembered a
' POSITIVE match -- a genuine miss (no shape on this slide type for this
' field) recorded nothing, so InjectorFor's up-to-four FindShapeByRoleTag
' calls per identity tag (base, ".1", ".track", ".rest") each re-walked the
' whole slide from scratch, every time, for every field with no matching
' shape -- the majority case once a register carries more populated columns
' than any one slide type has fields for (measured: ~4-5s/item on a real
' apply run, diagnosed by a cold audit as the actual dominant unmeasured
' cost in the tool, nothing to do with tonight's earlier AF/AL work).
' RecordAbsent below closes that gap using the SAME invariant the positive
' cache already relies on (Slide.Duplicate copies the template's shapes
' unchanged; nothing in this codebase adds or renames one afterwards) -- if
' a slide type's template has no shape for a field, no instance of that
' type ever will either, so "confirmed absent" is exactly as stable as
' "confirmed present". This sentinel is a normal string, not an empty one,
' specifically so it is not rejected by Record's own empty-shapeName guard
' and is distinguishable from Lookup's "" ("nothing cached yet") result.
Public Const NO_SHAPE_MARKER As String = "(no shape)"

Private Const COL_B_TYPE As Long = 1
Private Const COL_B_FIELDID As Long = 2
Private Const COL_B_SHAPENAME As Long = 3

Private Const ADDRESS_BOOK_HEADER_ROW As Long = 1
Private Const ADDRESS_BOOK_FIRST_ROW As Long = 2

' HELD AT MODULE LEVEL FOR THE SAME REASON DraftingLobby.mAppEvents IS: a
' Private workbook reference set once and reused across many calls, not
' rebuilt or looked up fresh each time. Unlike mAppEvents this is not a
' WithEvents sink -- it is simply "which workbook does the cache live in" --
' but the lifetime reasoning is identical: set once per paired workbook,
' reused for as long as this add-in session runs.
Private mWb As Object

' Wired from WorkbookBridge.OpenOrGetWorkbook, the one place every path
' through this add-in already goes through to reach the register workbook --
' same hook point DraftingLobby.EnsureWatching uses, for the same reason.
' Idempotent and cheap to call every resolve.
Public Sub SetActiveWorkbook(wb As Object)
    Set mWb = wb
End Sub

Private Function EnsureSheet() As Object
    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(mWb, ADDRESS_BOOK_SHEET_NAME)

    If Trim(CStr(ws.Cells(ADDRESS_BOOK_HEADER_ROW, COL_B_TYPE).Value)) = "" Then
        ws.Cells(ADDRESS_BOOK_HEADER_ROW, COL_B_TYPE).Value = "SlideType"
        ws.Cells(ADDRESS_BOOK_HEADER_ROW, COL_B_FIELDID).Value = "FieldID"
        ws.Cells(ADDRESS_BOOK_HEADER_ROW, COL_B_SHAPENAME).Value = "ShapeName"
        ws.Rows(ADDRESS_BOOK_HEADER_ROW).Font.Bold = True
    End If

    Set EnsureSheet = ws
End Function

Private Function LastBookRow(ws As Object) As Long
    Dim r As Long
    r = ADDRESS_BOOK_FIRST_ROW
    Do While Trim(CStr(ws.Cells(r, COL_B_FIELDID).Value)) <> ""
        r = r + 1
    Loop
    LastBookRow = r - 1
End Function

Private Function FindBookRow(ws As Object, slideType As String, fieldId As String) As Long
    Dim last As Long
    last = LastBookRow(ws)

    Dim r As Long
    For r = ADDRESS_BOOK_FIRST_ROW To last
        If StrComp(CStr(ws.Cells(r, COL_B_TYPE).Value), slideType, vbTextCompare) = 0 And _
           StrComp(CStr(ws.Cells(r, COL_B_FIELDID).Value), fieldId, vbTextCompare) = 0 Then
            FindBookRow = r
            Exit Function
        End If
    Next r
End Function

' "" if nothing is cached yet (a genuine miss -- the caller's own full walk
' is what populates this, not a separate discovery step) or if no workbook
' has been wired yet (SetActiveWorkbook never called -- callable safely
' before any workbook is open, same defensive shape DraftingLobby's
' functions already use).
Public Function Lookup(slideType As String, fieldId As String) As String
    If mWb Is Nothing Then Exit Function
    If Not WorkbookBridge.WorksheetExists(mWb, ADDRESS_BOOK_SHEET_NAME) Then Exit Function

    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(mWb, ADDRESS_BOOK_SHEET_NAME)

    Dim r As Long
    r = FindBookRow(ws, slideType, fieldId)
    If r > 0 Then Lookup = CStr(ws.Cells(r, COL_B_SHAPENAME).Value)
End Function

' THE SELF-HEALING WRITE. Called by InjectPrimitive.FindShapeByRoleTag every
' time its full walk finds a real answer -- on first use (cache was empty)
' and on drift (cache was wrong) alike, with no distinction between the two:
' both are just "here is the current truth, remember it." Idempotent --
' recording the same (slideType, fieldId) twice updates the row in place
' rather than growing the sheet, same rule DraftingLobby.PinToLobby follows.
Public Sub Record(slideType As String, fieldId As String, shapeName As String)
    If mWb Is Nothing Then Exit Sub
    If slideType = "" Or fieldId = "" Or shapeName = "" Then Exit Sub

    Dim ws As Object
    Set ws = EnsureSheet()

    Dim existing As Long
    existing = FindBookRow(ws, slideType, fieldId)

    Dim targetRow As Long
    If existing > 0 Then
        targetRow = existing
    Else
        targetRow = LastBookRow(ws) + 1
        If targetRow < ADDRESS_BOOK_FIRST_ROW Then targetRow = ADDRESS_BOOK_FIRST_ROW
    End If

    ws.Cells(targetRow, COL_B_TYPE).Value = slideType
    ws.Cells(targetRow, COL_B_FIELDID).Value = fieldId
    ws.Cells(targetRow, COL_B_SHAPENAME).Value = shapeName
End Sub

' THE NEGATIVE TWIN OF Record. Called by FindShapeByRoleTag only when its
' full walk finds ZERO matches -- deliberately NOT when it finds two or
' more (an ambiguous, exceptional state; caching that as "absent" would be
' actively wrong, and re-walking a rare ambiguous case costs nothing). Same
' idempotent update-in-place behaviour as Record. A genuinely later-added
' template shape would need this book cleared by hand to be seen -- no
' automatic invalidation exists for either cache direction today, and none
' is added here on the strength of a case that has not happened.
Public Sub RecordAbsent(slideType As String, fieldId As String)
    If mWb Is Nothing Then Exit Sub
    If slideType = "" Or fieldId = "" Then Exit Sub

    Dim ws As Object
    Set ws = EnsureSheet()

    Dim existing As Long
    existing = FindBookRow(ws, slideType, fieldId)

    Dim targetRow As Long
    If existing > 0 Then
        targetRow = existing
    Else
        targetRow = LastBookRow(ws) + 1
        If targetRow < ADDRESS_BOOK_FIRST_ROW Then targetRow = ADDRESS_BOOK_FIRST_ROW
    End If

    ws.Cells(targetRow, COL_B_TYPE).Value = slideType
    ws.Cells(targetRow, COL_B_FIELDID).Value = fieldId
    ws.Cells(targetRow, COL_B_SHAPENAME).Value = NO_SHAPE_MARKER
End Sub
