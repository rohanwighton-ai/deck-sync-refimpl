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

' FIX-LIST item AT, 2026-08-17 late. mSheet/mNextRow are the SAME "resolve
' once, reuse" fix as items W and AB, applied to the cache's own storage
' layer -- found because the cache built to fix item AR turned out to
' contain items W and AB's exact defect shape itself. Before this fix,
' every single Lookup/Record/RecordAbsent call re-resolved the sheet via
' WorkbookBridge.WorksheetExists and/or GetOrAddWorksheet, each a `For Each
' ws In wb.Worksheets` scan of the WHOLE WORKBOOK (~54 sheets on the real
' register) -- independent of how big the address book itself is. Measured
' live against the real register (17-row book): 508.5ms/call. At this
' project's own measured ~5,000+ Lookup calls per "Put it on the slides"
' press (FIX-LIST item AF/AR), that is ~42 minutes -- worse than the
' 14.5-minute run that got killed, and it explains why the process reads as
' frozen rather than slow: steady CPU work, zero visible output.
' mNextRow closes the matching gap in Record/RecordAbsent -- both called
' LastBookRow (its own O(book length) scan) a second time on every append,
' the identical pattern item AB fixed in BuildLobbyFromScratch/PinToLobby.
Private mSheet As Object
Private mNextRow As Long ' 0 = not yet resolved for the current mSheet

' THE PER-SLIDE TAG INDEX -- the second generation of the miss cache, and a
' correctness fix, not a tuning pass (2026-08-18, the TIMELINE_ELAPSED
' "dropped: changed since approval" incident). The first generation
' (mNegativeCache, FIX-LIST item AT) keyed confirmed absence by
' (slideType, fieldId) -- but ABSENCE IS A PER-SLIDE FACT, and keying it
' per TYPE let one slide's genuine lack of a shape veto another slide's
' confirmed presence of it. Live shape of the failure: 44 slides carry
' slide_type=project-progress and exactly ONE of them (the prototype,
' 3_P001) has the elapsed-time bar. Every "Put it on the slides" press
' built the queue (3_P001 probed first, bar found), then walked the 43
' barless siblings -- the first of which recorded a TYPE-WIDE "(no shape)"
' -- so seconds later ApplyApproved's dry probe of 3_P001 itself was told
' the bar did not exist, read CurrentValue="", and dropped the approved
' change as "changed since approval" on every single run. The positive
' book row for the bar sat on this very sheet the whole time; the
' in-memory miss simply outranked it, and nothing ever reconciled the two.
'
' Gist: the old notebook of "that shape doesn't exist" answers was filed
' under the slide's TYPE, so one slide that truly lacked a shape hid the
' same shape on the one slide that had it; this files every answer under
' the individual slide instead, so slides can no longer lie about each
' other.
'
' The replacement stores, per individual slide (keyed by SlideID), the
' full set of identity keys the walk saw on that slide -- built as a free
' byproduct of the first full walk of that slide each session, since the
' walk visits every shape anyway. Absence for (slide, tag) is then a
' memory lookup with zero COM calls, same cost as the old design, and the
' walk count per session is BOUNDED BY THE SLIDE COUNT (one indexing walk
' per slide) instead of by (type x absent-field x suffix). Same
' in-memory-only, reset-on-workbook-change lifetime as before, for the
' same reason: presence only needs to survive the current session, and an
' in-memory cache invalidates itself for free on every reopen.
Private mSlideTagIndex As Object ' Scripting.Dictionary: slideKey -> Dictionary of identity keys present

' Wired from WorkbookBridge.OpenOrGetWorkbook, the one place every path
' through this add-in already goes through to reach the register workbook --
' same hook point DraftingLobby.EnsureWatching uses, for the same reason.
' Idempotent and cheap to call every resolve -- cheaper still now: re-wiring
' the SAME workbook object (the common case) no longer invalidates the
' sheet/row cache at all.
Public Sub SetActiveWorkbook(wb As Object)
    If Not (mWb Is wb) Then
        Set mSheet = Nothing
        mNextRow = 0
        Set mSlideTagIndex = Nothing
    End If
    Set mWb = wb
End Sub

Private Function SlideIndex() As Object
    If mSlideTagIndex Is Nothing Then
        Set mSlideTagIndex = CreateObject("Scripting.Dictionary")
    End If
    Set SlideIndex = mSlideTagIndex
End Function

' The one place a slide's cache key is derived, so the walk's recorder, the
' absence check, and the stamp-notification below can never disagree about
' what identifies a slide. SlideID is PowerPoint's own stable per-
' presentation id (survives reorders, unlike SlideIndex); "" on anything
' that cannot produce one, and every consumer treats "" as "do not cache".
' Two presentations open in one session could in principle collide on
' SlideID -- accepted: real sessions work one deck at a time, and a deck
' switch goes through a workbook re-wire that resets this index anyway.
Public Function SlideKeyFor(sld As Object) As String
    On Error Resume Next
    SlideKeyFor = CStr(sld.SlideID)
    On Error GoTo 0
End Function

' The identity keys the last full walk saw on this slide, or Nothing if the
' slide has not been walked this session (or no workbook is wired). A
' caller holding Nothing MUST walk; a caller holding a dictionary may trust
' `Not .Exists(tag)` as confirmed absence ON THIS SLIDE -- the per-slide
' scope is the entire point, see mSlideTagIndex's header.
Public Function SlideTagsFor(slideKey As String) As Object
    If mWb Is Nothing Then Exit Function
    If slideKey = "" Then Exit Function
    If mSlideTagIndex Is Nothing Then Exit Function
    If mSlideTagIndex.Exists(slideKey) Then Set SlideTagsFor = mSlideTagIndex(slideKey)
End Function

' Called by InjectPrimitive.FindShapeByRoleTag every time its full walk
' completes, with the complete set of identity keys the walk saw. Always an
' overwrite, never a merge: the walk just visited every shape on the slide,
' so its answer supersedes whatever was recorded before -- which is also
' what lets a re-walk self-heal a stale entry, same reasoning as Record.
Public Sub RecordSlideTags(slideKey As String, tagsPresent As Object)
    If mWb Is Nothing Then Exit Sub
    If slideKey = "" Then Exit Sub
    If tagsPresent Is Nothing Then Exit Sub
    Set SlideIndex()(slideKey) = tagsPresent
End Sub

' CACHE COHERENCE FOR THE ONE EVENT THAT CAN CREATE A FALSE ABSENCE: a role
' tag stamped onto a shape mid-session (Onboarding.ConfirmFieldMatch,
' Harvest.PropagateTemplateTags, TagMigration, picture re-add). A stale
' "present" entry is benign -- it just buys a walk that answers correctly
' and re-records. A stale "absent" entry is the dangerous direction: it
' short-circuits BEFORE any walk, so without this call a slide indexed
' before the stamp would hide the newly-tagged shape for the rest of the
' session. Every code path that adds a role tag calls this; a tag added by
' hand in the PowerPoint UI mid-session is not covered (nothing to hook),
' matching the restart-to-clear limitation the old design documented.
' Gist: when the tool itself gives a shape a new name-tag, it also updates
' its own notebook so it doesn't keep believing the tag isn't there.
Public Sub NoteRoleTagAdded(shp As Object, identityTag As String)
    If mWb Is Nothing Then Exit Sub
    If identityTag = "" Then Exit Sub
    If mSlideTagIndex Is Nothing Then Exit Sub

    Dim k As String
    On Error Resume Next
    k = CStr(shp.Parent.SlideID)
    On Error GoTo 0
    If k = "" Then Exit Sub

    If mSlideTagIndex.Exists(k) Then mSlideTagIndex(k)(identityTag) = True
End Sub

' READ PATH. Deliberately does NOT create the sheet -- Lookup's "nothing
' cached yet" contract for a never-yet-written book depends on that. Caches
' mSheet only once the sheet is confirmed to exist, same invalidation as
' SetActiveWorkbook.
Private Function ResolveSheetForRead() As Object
    If Not (mSheet Is Nothing) Then
        Set ResolveSheetForRead = mSheet
        Exit Function
    End If
    If mWb Is Nothing Then Exit Function
    If Not WorkbookBridge.WorksheetExists(mWb, ADDRESS_BOOK_SHEET_NAME) Then Exit Function

    Set mSheet = WorkbookBridge.GetOrAddWorksheet(mWb, ADDRESS_BOOK_SHEET_NAME)
    mNextRow = LastBookRow(mSheet) + 1
    Set ResolveSheetForRead = mSheet
End Function

' WRITE PATH. Creates the sheet (and seeds its header) on first use, same
' as the old EnsureSheet -- just cached afterward instead of re-resolved.
Private Function ResolveSheetForWrite() As Object
    If Not (mSheet Is Nothing) Then
        Set ResolveSheetForWrite = mSheet
        Exit Function
    End If

    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(mWb, ADDRESS_BOOK_SHEET_NAME)

    If Trim(CStr(ws.Cells(ADDRESS_BOOK_HEADER_ROW, COL_B_TYPE).Value)) = "" Then
        ws.Cells(ADDRESS_BOOK_HEADER_ROW, COL_B_TYPE).Value = "SlideType"
        ws.Cells(ADDRESS_BOOK_HEADER_ROW, COL_B_FIELDID).Value = "FieldID"
        ws.Cells(ADDRESS_BOOK_HEADER_ROW, COL_B_SHAPENAME).Value = "ShapeName"
        ws.Rows(ADDRESS_BOOK_HEADER_ROW).Font.Bold = True
    End If

    Set mSheet = ws
    mNextRow = LastBookRow(ws) + 1
    Set ResolveSheetForWrite = mSheet
End Function

Private Function LastBookRow(ws As Object) As Long
    Dim r As Long
    r = ADDRESS_BOOK_FIRST_ROW
    Do While Trim(CStr(ws.Cells(r, COL_B_FIELDID).Value)) <> ""
        r = r + 1
    Loop
    LastBookRow = r - 1
End Function

' Takes the book's own known extent (mNextRow - 1) instead of recomputing it
' via LastBookRow's own scan -- FindBookRow used to call LastBookRow on
' EVERY invocation, which meant the mSheet/mNextRow caching above still left
' this exact same class of redundant rescan running underneath it, on every
' single Lookup call. Found by comparing a live measured number against
' what the fix should have produced and not accepting the gap unexplained
' (FIX-LIST item AT).
Private Function FindBookRow(ws As Object, slideType As String, fieldId As String) As Long
    Dim last As Long
    last = mNextRow - 1

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
    ' Misses are no longer answered here at all -- confirmed absence lives
    ' in the per-slide tag index (SlideTagsFor), because absence is a
    ' per-slide fact and answering it per type is the exact defect
    ' mSlideTagIndex's header records. This function now only ever reports
    ' the positive book: a name once found, or "" for never-recorded.
    ' (A legacy register whose sheet still carries "(no shape)" rows from
    ' the pre-AT persistent-miss era can still surface NO_SHAPE_MARKER
    ' through the sheet read below; FindShapeByRoleTag treats that as
    ' "no hint", never as a verdict.)
    Dim ws As Object
    Set ws = ResolveSheetForRead()
    If ws Is Nothing Then Exit Function

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
    Set ws = ResolveSheetForWrite()

    Dim existing As Long
    existing = FindBookRow(ws, slideType, fieldId)

    Dim targetRow As Long
    If existing > 0 Then
        targetRow = existing
    Else
        targetRow = mNextRow
        If targetRow < ADDRESS_BOOK_FIRST_ROW Then targetRow = ADDRESS_BOOK_FIRST_ROW
        mNextRow = targetRow + 1
    End If

    ws.Cells(targetRow, COL_B_TYPE).Value = slideType
    ws.Cells(targetRow, COL_B_FIELDID).Value = fieldId
    ws.Cells(targetRow, COL_B_SHAPENAME).Value = shapeName
End Sub

' RecordAbsent (the per-TYPE negative twin of Record) was retired
' 2026-08-18: recording one slide's zero-match walk as a type-wide verdict
' is the defect mSlideTagIndex's header documents. Absence is now recorded
' as a byproduct of RecordSlideTags above -- a tag not in a slide's
' recorded set is confirmed absent for THAT slide only.
