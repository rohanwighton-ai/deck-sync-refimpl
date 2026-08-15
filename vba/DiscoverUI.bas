Attribute VB_Name = "DiscoverUI"
Option Explicit

' MARK EVERY FIELD ON ONE SHEET, INSTEAD OF THREE DIALOGS AT A TIME.
'
' Rohan marked 52 fields on 2026-08-01. Each cost three modal prompts -- name,
' type, volatility -- so roughly 150 dialogs, about an hour. It then produced
' three of that day's thirteen findings, all of them consequences of the prompt
' chain rather than of marking itself: an unmarkable shape accepted and only
' refused at the end (3), no way to remove one bad mark (4), and an identifier
' that could not tell two shapes apart (11).
'
' THIS PROJECT ALREADY MADE THIS EXACT JOURNEY ONCE. Instance keys were a modal
' prompt per slide -- 45 of them -- until 2026-07-29, when Rohan said: "I'd
' rather fix the excel and skip all this having to confirm the instance key...
' too time consuming." They became one Excel grid. The recorded reason is worth
' repeating, because it is the whole argument for this module too:
'
'     modal prompts are irreplaceable input. A sheet is editable, re-openable,
'     and survives a validation failure with the human's edits intact. A chain
'     of dialogs is none of those.
'
' Field marking never got that treatment. This is it.
'
' ADDITIVE, DELIBERATELY. The per-shape marking flow is untouched and still works;
' it is reached from inside the Sync Now chain rather than from its own button.
' A new flow built in one evening and never run against Rohan's real deck should
' not be the only road; if this misbehaves he marks the old way and loses
' nothing but the time he would have spent anyway.

' PUBLIC, so WorkbookBridge.IsToolOwnedSheet can name it by constant rather
' than by a duplicated string literal. It was private and duplicated, and the
' duplicate was simply missing -- which let the register resolver treat this
' sheet as the user's data.
Public Const DISCOVERY_SHEET_NAME As String = "Field Discovery"
Private Const DISCOVERY_SHEET As String = "Field Discovery"

Private Const COL_ID As Long = 1
Private Const COL_SHAPE As Long = 2
Private Const COL_WHERE As Long = 3
Private Const COL_TEXT As Long = 4
Private Const COL_CHARS As Long = 5
Private Const COL_INCLUDE As Long = 6
Private Const COL_FIELD As Long = 7
Private Const COL_TYPE As Long = 8
Private Const COL_VOL As Long = 9

Private Const HEADER_ROW As Long = 6
Private Const FIRST_ROW As Long = 7

' ---------------------------------------------------------------------------
' BUTTON: Setup A2: Discover Fields
' ---------------------------------------------------------------------------
Public Sub DiscoverFields()
    ' Named for the button that reaches it, not for the routine. This said
    ' "Discover Fields", a caption matching no button on the toolbar -- the same
    ' defect fixed in RepointWorkbookUI the same day.
    Const CAP As String = CommandBarUI.CAP_DISCOVER_FIELDS
    On Error GoTo Failed

    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim sld As Object
    Set sld = Nothing
    On Error Resume Next
    Set sld = Application.ActiveWindow.View.Slide
    On Error GoTo 0
    If sld Is Nothing Then
        MsgBox "Open the slide you want to use as the template, in Normal view, then run this again.", _
               vbExclamation, CAP
        Exit Sub
    End If

    ' NO PAIRING YET? ESTABLISH ONE, DO NOT REFUSE.
    '
    ' This used to say "press Sync Now -- that is what establishes the pairing",
    ' which on a virgin deck is the button the person had just pressed: Sync
    ' Now's setup branch calls THIS, and the onboarding step that would have
    ' created the pairing runs afterwards and refuses for want of the fields
    ' this produces. Two refusals pointing at each other and no way in. Found on
    ' the real deck, 2026-08-13, on its first ever run.
    '
    ' Pairing is a question about WHICH WORKBOOK and has nothing to do with
    ' fields, so asking it here costs nothing and unblocks the only order a
    ' first-time user can actually follow.
    Dim workbookPath As String
    workbookPath = DeckRegistry.GetWorkbookPath(pres)
    If workbookPath = "" Then
        Dim pairCancel As String
        Dim wbPair As Object
        Set wbPair = BatchOnboardFlow.ResolveDataWorkbook(pres, workbookPath, pairCancel)
        If wbPair Is Nothing Or workbookPath = "" Then
            MsgBox "Discover Fields writes its grid into the paired workbook, so there has to be one." & vbCrLf & vbCrLf & _
                   IIf(pairCancel = "", "No workbook was chosen.", pairCancel), vbExclamation, CAP
            Exit Sub
        End If
    End If

    Dim wb As Object
    Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath)
    If wb Is Nothing Then
        MsgBox "Could not open the paired workbook at:" & vbCrLf & workbookPath, vbCritical, CAP
        Exit Sub
    End If

    Dim built As String
    built = BuildDiscoverySheet(sld, wb)
    If Left(built, 1) = "!" Then
        MsgBox Mid(built, 2), vbExclamation, CAP
        Exit Sub
    End If

    ' PRESENTING IS THE UI'S JOB, NOT THE BUILDER'S.
    ' BuildDiscoverySheet used to make Excel visible itself, which meant a
    ' headless test left a VISIBLE Excel behind -- and the harness's cleanup
    ' sweep only reaps windowless processes, so it survived every run and had to
    ' be killed by hand. A function that both computes and presents cannot be
    ' called without its side effects.
    On Error Resume Next
    wb.Application.Visible = True
    wb.Activate
    wb.Worksheets(DISCOVERY_SHEET).Activate
    wb.Worksheets(DISCOVERY_SHEET).Cells(FIRST_ROW, COL_INCLUDE).Select
    On Error GoTo 0

    ' --- wait for the human ---------------------------------------------
    ' Same shape as Bulk Onboard's own review pause. VBA cannot watch a
    ' worksheet and continue, so the dialog IS the wait -- and it is honest
    ' about that rather than pretending to be asynchronous.
    If MsgBox(built & vbCrLf & vbCrLf & _
              "Put Y in column F and a name in column G for each field you want tracked." & vbCrLf & vbCrLf & _
              "Click YES when you have finished editing, NO to cancel and mark nothing.", _
              vbYesNo + vbQuestion, CAP) <> vbYes Then
        MsgBox "Nothing was marked. The grid is still there if you want to come back to it.", vbInformation, CAP
        Exit Sub
    End If

    MsgBox ApplyDiscoverySheet(sld, wb), vbInformation, CAP
    Exit Sub

Failed:
    MsgBox "Discover Fields stopped early." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Nothing on the deck was changed -- this only writes a sheet and marks fields in memory.", _
           vbCritical, CAP
End Sub

' Writes the grid. Split out of DiscoverFields so it can be driven from a test
' without a human clicking a dialog -- a flow that has only ever been run by
' hand is a flow nobody can regression-test, which is how the prompt chain it
' replaces stayed broken for so long.
'
' Returns a human-readable summary, or "!message" if it could not proceed.
Public Function BuildDiscoverySheet(sld As Object, wb As Object) As String
    Dim cands() As Candidate
    Dim shapes() As Object
    cands = Discovery.DiscoverSlideWithShapes(sld, shapes)

    Dim lo As Long, hi As Long, hasAny As Boolean
    On Error Resume Next
    lo = LBound(cands): hi = UBound(cands)
    hasAny = (Err.Number = 0)
    On Error GoTo 0
    If Not hasAny Then
        BuildDiscoverySheet = "!Found no text shapes on this slide."
        Exit Function
    End If

    ' Reading order, not z-order. A grid ordered the way the slide is read is a
    ' grid a person can work down without hunting; z-order is an implementation
    ' detail of how the deck was built and correlates with nothing they can see.
    Dim order() As Long
    order = ReadingOrder(cands, lo, hi)

    ' --- the sheet ------------------------------------------------------
    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(wb, DISCOVERY_SHEET)

    ' THE CLEAR IS GONE, AND A CARRY DID NOT REPLACE IT.
    '
    ' This used to be `ws.Cells.Clear`. The only thing that bought was avoiding
    ' stale rows BELOW the new grid when a slide lost a shape -- a tail problem,
    ' solved by wiping the whole sheet including columns F and G, which are the
    ' two the sheet's own row 3 instructs a person to type into, and which
    ' ApplyDiscoverySheet reads back. Marking a slide, declining to apply, and
    ' pressing Discover Fields again silently destroyed the marks -- directly
    ' contradicting the message the tool shows on decline: "The grid is still
    ' there if you want to come back to it."
    '
    ' Rohan, 2026-08-15: "why is clear needed there?" It is not. The first
    ' answer drafted was to read F and G into a dictionary and write them back
    ' afterwards, which is a carry, which is a backup around a destructive call
    ' -- the tell DOCUMENT-MAP decision 1 names outright. The call goes instead.
    '
    ' So an existing row is UPDATED WHERE IT SITS, matched on the shape id the
    ' grid already carries in column A. The row never moves, so a person's mark
    ' cannot end up beside a different shape, and nothing has to be held in
    ' memory during the write.
    Dim rowById As Object
    Set rowById = ExistingRowsById(ws)

    ' A sheet with no rows to preserve is still built from scratch -- and this
    ' clear also drops stale FORMATTING from an earlier layout, which a
    ' ClearContents would leave behind.
    If rowById.count = 0 Then ws.Cells.Clear

    ws.Cells(1, 1).Value = "FIELD DISCOVERY  --  every text shape on your template slide"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(2, 1).Value = "1.  Work down the list. It is in reading order: top of the slide first."
    ws.Cells(3, 1).Value = "2.  For anything you want tracked, type a NAME in column G and put Y in column F."
    ws.Cells(4, 1).Value = "3.  Leave the rest alone. Save and come back to PowerPoint -- nothing is marked until you say so."

    ws.Cells(HEADER_ROW, COL_ID).Value = "Shape ID (do not edit)"
    ws.Cells(HEADER_ROW, COL_SHAPE).Value = "Shape"
    ws.Cells(HEADER_ROW, COL_WHERE).Value = "Inside group"
    ws.Cells(HEADER_ROW, COL_TEXT).Value = "What it says now"
    ws.Cells(HEADER_ROW, COL_CHARS).Value = "Chars"
    ws.Cells(HEADER_ROW, COL_INCLUDE).Value = "F -- Y TO TRACK IT"
    ws.Cells(HEADER_ROW, COL_FIELD).Value = "G -- FIELD NAME"
    ws.Cells(HEADER_ROW, COL_TYPE).Value = "Type"
    ws.Cells(HEADER_ROW, COL_VOL).Value = "Static/Variable"
    ws.Rows(HEADER_ROW).Font.Bold = True
    ws.Rows(HEADER_ROW).WrapText = True

    ' The append cursor sits below every row already on the sheet, so a shape
    ' that is genuinely new lands at the bottom rather than on top of somebody
    ' else's row. New shapes therefore break reading order, which is the honest
    ' trade: reading order is a convenience, a mark landing beside the wrong
    ' shape is a data defect.
    Dim r As Long, written As Long
    r = FIRST_ROW
    Dim rk As Variant
    For Each rk In rowById.Keys
        If rowById(rk) >= r Then r = rowById(rk) + 1
    Next rk

    ' Every shape id this run has seen, so rows for shapes that have left the
    ' slide can be identified after the loop rather than guessed at.
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim maxRow As Long
    maxRow = FIRST_ROW - 1

    Dim k As Long
    For k = LBound(order) To UBound(order)
        Dim i As Long
        i = order(k)
        If cands(i).HasText Then
            Dim shp As Object
            Set shp = shapes(i)

            Dim txt As String
            txt = ""
            On Error Resume Next
            txt = shp.TextFrame.TextRange.Text
            On Error GoTo 0
            txt = Replace(Replace(Replace(txt, vbCrLf, " / "), vbCr, " / "), vbLf, " / ")

            ' Where this shape's row is: the one it already had, or a new one at
            ' the bottom. COLUMNS F AND G ARE NOT TOUCHED IN EITHER CASE -- on an
            ' existing row they are the person's marks and must survive; on a
            ' fresh row they are already empty.
            Dim idKey As String
            idKey = CStr(shp.Id)

            Dim tr As Long
            If rowById.Exists(idKey) Then
                tr = rowById(idKey)
            Else
                tr = r
                r = r + 1
            End If
            seen(idKey) = True
            If tr > maxRow Then maxRow = tr

            ws.Cells(tr, COL_ID).Value = shp.Id
            ws.Cells(tr, COL_SHAPE).Value = "'" & cands(i).Name
            ws.Cells(tr, COL_WHERE).Value = "'" & cands(i).GroupPath
            ws.Cells(tr, COL_TEXT).Value = "'" & txt
            ws.Cells(tr, COL_CHARS).Value = Len(txt)
            ws.Cells(tr, COL_TYPE).Value = "text"
            ws.Cells(tr, COL_VOL).Value = "variable"
            written = written + 1
        End If
    Next k

    ' ONLY the rows whose shape has left the slide. This is the tail problem the
    ' old whole-sheet clear was really solving, solved where it actually lives --
    ' and a row is cleared entirely, marks included, because the shape those
    ' marks referred to no longer exists.
    Dim gone As Long
    For Each rk In rowById.Keys
        If Not seen.Exists(CStr(rk)) Then
            ws.Range(ws.Cells(rowById(rk), COL_ID), ws.Cells(rowById(rk), COL_VOL)).Clear
            gone = gone + 1
        End If
    Next rk

    ' The formatting below addresses rows FIRST_ROW..r-1, so `r` must end one
    ' past the last row in use -- which is the append cursor only when nothing
    ' was updated in place further down the sheet.
    If maxRow + 1 > r Then r = maxRow + 1

    ws.Cells.Font.Size = 8
    ws.Cells(1, 1).Font.Size = 9
    ws.Cells.VerticalAlignment = -4160          ' xlTop
    ws.Columns(COL_ID).ColumnWidth = 9
    ws.Columns(COL_SHAPE).ColumnWidth = 18
    ws.Columns(COL_WHERE).ColumnWidth = 22
    ws.Columns(COL_TEXT).ColumnWidth = 64
    ws.Columns(COL_TEXT).WrapText = True
    ws.Columns(COL_CHARS).ColumnWidth = 6
    ws.Columns(COL_INCLUDE).ColumnWidth = 11
    ws.Columns(COL_FIELD).ColumnWidth = 26
    ws.Columns(COL_TYPE).ColumnWidth = 10
    ws.Columns(COL_VOL).ColumnWidth = 13
    ws.Rows(HEADER_ROW).RowHeight = 28
    If r > FIRST_ROW Then ws.Range(ws.Rows(FIRST_ROW), ws.Rows(r - 1)).RowHeight = 30

    ' The two columns a person types in are the only ones that look like inputs.
    ws.Range(ws.Cells(FIRST_ROW, COL_INCLUDE), ws.Cells(r - 1, COL_FIELD)).Interior.Color = RGB(255, 249, 219)
    ws.Columns(COL_TEXT).Interior.Color = RGB(242, 242, 242)

    ' Says what SURVIVED, not just what was written. A person who marked rows
    ' and came back needs to know their marks are still there without having to
    ' go and look -- and if the number is ever 0 when they expected otherwise,
    ' that is the bug report.
    Dim kept As Long
    kept = CountMarks(ws, r - 1)

    BuildDiscoverySheet = written & " text shape(s) listed in '" & DISCOVERY_SHEET & "'."
    If kept > 0 Then
        BuildDiscoverySheet = BuildDiscoverySheet & "  " & kept & _
            " existing mark(s) kept."
    End If
    If gone > 0 Then
        BuildDiscoverySheet = BuildDiscoverySheet & "  " & gone & _
            " row(s) removed for shapes no longer on the slide."
    End If
End Function

' The shape id -> row map the in-place update needs, read from the grid itself
' rather than held across a rebuild.
'
' Walks from FIRST_ROW while column A holds an id, which is how every other
' reader of a tool-written grid in this project finds the end (ReadQueueSheet
' does the same on its hash column). Returns an empty dictionary for a sheet
' that has never been built, which is what makes the fresh-build branch a
' count test rather than a separate existence check.
Private Function ExistingRowsById(ws As Object) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim r As Long
    r = FIRST_ROW
    Do While Trim(CStr(ws.Cells(r, COL_ID).Value)) <> ""
        d(Trim(CStr(ws.Cells(r, COL_ID).Value))) = r
        r = r + 1
    Loop

    Set ExistingRowsById = d
End Function

' How many rows a person has actually marked -- a Y in F, or a name typed in G.
' Counts either, because a name with no Y is still work somebody did and still
' something a rebuild must not lose.
Private Function CountMarks(ws As Object, lastRow As Long) As Long
    Dim n As Long, r As Long
    For r = FIRST_ROW To lastRow
        If ReviewQueue.IsApprovalMark(CStr(ws.Cells(r, COL_INCLUDE).Value)) Then
            n = n + 1
        ElseIf Trim(CStr(ws.Cells(r, COL_FIELD).Value)) <> "" Then
            n = n + 1
        End If
    Next r
    CountMarks = n
End Function

' Reads the grid back and marks what was ticked. Re-discovers the slide rather
' than trusting an array captured before the human edited anything -- shapes can
' move, and a stale array would bind marks to whatever was there earlier.
Public Function ApplyDiscoverySheet(sld As Object, wb As Object) As String
    Dim cands() As Candidate
    Dim shapes() As Object
    cands = Discovery.DiscoverSlideWithShapes(sld, shapes)

    Dim lo As Long, hi As Long, hasAny As Boolean
    On Error Resume Next
    lo = LBound(cands): hi = UBound(cands)
    hasAny = (Err.Number = 0)
    On Error GoTo 0
    If Not hasAny Then
        ApplyDiscoverySheet = "No text shapes on this slide."
        Exit Function
    End If

    If Not WorkbookBridge.WorksheetExists(wb, DISCOVERY_SHEET) Then
        ApplyDiscoverySheet = "No '" & DISCOVERY_SHEET & "' sheet -- run Discover Fields first."
        Exit Function
    End If
    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(wb, DISCOVERY_SHEET)

    ' --- read it back ----------------------------------------------------
    Dim marked As Long, skippedNoName As Long, skippedNotFound As Long
    Dim unmarked As Long
    Dim problems As String, removedList As String

    Dim rr As Long
    rr = FIRST_ROW
    Do While Trim(CStr(ws.Cells(rr, COL_ID).Value)) <> ""
        Dim wantId As String
        wantId = Trim(CStr(ws.Cells(rr, COL_ID).Value))

        ' UNTICKING REMOVES THE MARK. The grid presents as declarative state --
        ' "what is ticked is what is tracked" -- and until now it behaved
        ' imperatively: ticked rows were appended and unticked rows did nothing.
        ' A sheet that looks like a checklist and acts like an in-tray is the
        ' same lie as every other finding here, and it meant this grid did NOT
        ' fix finding 4 even though it looked like it should.
        If Not ReviewQueue.IsApprovalMark(CStr(ws.Cells(rr, COL_INCLUDE).Value)) Then
            Dim maybeMarked As Object
            Set maybeMarked = ShapeById(shapes, lo, hi, wantId)
            If Not maybeMarked Is Nothing Then
                Dim goneName As String
                goneName = BatchOnboardFlow.UnmarkShapeForBatch(maybeMarked)
                If goneName <> "" Then
                    unmarked = unmarked + 1
                    removedList = removedList & "  " & goneName & vbCrLf
                End If
            End If
        End If

        If ReviewQueue.IsApprovalMark(CStr(ws.Cells(rr, COL_INCLUDE).Value)) Then
            Dim fname As String
            fname = Trim(CStr(ws.Cells(rr, COL_FIELD).Value))

            If fname = "" Then
                ' Ticked with no name is a person part-way through a thought,
                ' not an instruction. Reported by name of shape so they can find
                ' it, never guessed at.
                skippedNoName = skippedNoName + 1
                problems = problems & "  row " & rr & " (" & CStr(ws.Cells(rr, COL_SHAPE).Value) & _
                           ") is ticked but has no field name" & vbCrLf
            Else
                Dim target As Object
                Set target = ShapeById(shapes, lo, hi, wantId)
                If target Is Nothing Then
                    skippedNotFound = skippedNotFound + 1
                    problems = problems & "  row " & rr & " -- shape id " & wantId & " is no longer on the slide" & vbCrLf
                Else
                    Dim st As String
                    st = BatchOnboardFlow.MarkShapeForBatch(target, fname, _
                            BatchOnboardFlow.NormalizeFieldType(CStr(ws.Cells(rr, COL_TYPE).Value)), _
                            BatchOnboardFlow.NormalizeFieldVolatility(CStr(ws.Cells(rr, COL_VOL).Value)))
                    ' MarkShapeForBatch returns a human-readable status. Anything
                    ' that is not an ordinary "Marked field N" is surfaced rather
                    ' than counted as success -- a grid that reports 14 marked
                    ' when 3 refused is worse than one that reports nothing.
                    If InStr(1, st, "Marked field", vbTextCompare) = 1 Then
                        marked = marked + 1
                    Else
                        problems = problems & "  row " & rr & " (" & fname & "): " & st & vbCrLf
                    End If
                End If
            End If
        End If
        rr = rr + 1
    Loop

    Dim msg As String
    msg = marked & " field(s) marked from the grid."
    If unmarked > 0 Then msg = msg & vbCrLf & unmarked & " UNMARKED (unticked here, so no longer tracked):" & vbCrLf & removedList
    If skippedNoName > 0 Then msg = msg & vbCrLf & skippedNoName & " ticked but unnamed -- not marked."
    If skippedNotFound > 0 Then msg = msg & vbCrLf & skippedNotFound & " shape(s) no longer on the slide."
    If problems <> "" Then msg = msg & vbCrLf & vbCrLf & "NOT MARKED:" & vbCrLf & problems
    If marked > 0 Then msg = msg & vbCrLf & vbCrLf & "Now press '" & CommandBarUI.CAP_SET_UP_QUARTER & "' to link the rest of the slides."

    ApplyDiscoverySheet = msg
End Function

' Indices sorted top-to-bottom, then left-to-right. Insertion sort: the lists are
' tens of items, and a simple sort that is obviously correct beats a clever one
' nobody will re-read.
Private Function ReadingOrder(cands() As Candidate, lo As Long, hi As Long) As Long()
    Dim idx() As Long
    ReDim idx(lo To hi)
    Dim i As Long
    For i = lo To hi
        idx(i) = i
    Next i

    Dim j As Long, k As Long, tmp As Long
    For j = lo + 1 To hi
        tmp = idx(j)
        k = j - 1
        Do While k >= lo
            If IsAfter(cands(idx(k)), cands(tmp)) Then
                idx(k + 1) = idx(k)
                k = k - 1
            Else
                Exit Do
            End If
        Loop
        idx(k + 1) = tmp
    Next j

    ReadingOrder = idx
End Function

' Is a after b in reading order? Rows first, with a tolerance, because two
' shapes that look side by side are rarely at the identical Y.
Private Function IsAfter(a As Candidate, b As Candidate) As Boolean
    Const ROW_TOLERANCE_EMU As Long = 180000     ' ~0.2 inch
    If Abs(a.PositionY - b.PositionY) > ROW_TOLERANCE_EMU Then
        IsAfter = (a.PositionY > b.PositionY)
    Else
        IsAfter = (a.PositionX > b.PositionX)
    End If
End Function

' The shape carrying this Id. By Id, never by name -- see BatchOnboardFlow's
' serializer comment: 47 of 158 shapes on the real deck share a name.
Private Function ShapeById(shapes() As Object, lo As Long, hi As Long, wantId As String) As Object
    Dim i As Long
    For i = lo To hi
        On Error Resume Next
        Dim thisId As String
        thisId = CStr(shapes(i).Id)
        On Error GoTo 0
        If thisId = wantId Then
            Set ShapeById = shapes(i)
            Exit Function
        End If
    Next i
End Function
