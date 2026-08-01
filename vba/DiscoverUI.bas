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
' ADDITIVE, DELIBERATELY. "Setup A: Mark Fields" is untouched and still works.
' A new flow built in one evening and never run against Rohan's real deck should
' not be the only road; if this misbehaves he marks the old way and loses
' nothing but the time he would have spent anyway.

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
    Const CAP As String = "Discover Fields"
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

    Dim workbookPath As String
    workbookPath = DeckRegistry.GetWorkbookPath(pres)
    If workbookPath = "" Then
        MsgBox "This deck has no paired workbook yet." & vbCrLf & vbCrLf & _
               "Discover Fields writes its grid into the paired workbook, so there has to be one." & vbCrLf & _
               "Use 'Setup A: Mark Fields' and 'Setup B: Onboard Slides' for the very first slide type -- " & _
               "that is what establishes the pairing.", vbExclamation, CAP
        Exit Sub
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
    ws.Cells.Clear

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

    Dim r As Long, written As Long
    r = FIRST_ROW
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

            ws.Cells(r, COL_ID).Value = shp.Id
            ws.Cells(r, COL_SHAPE).Value = "'" & cands(i).Name
            ws.Cells(r, COL_WHERE).Value = "'" & cands(i).GroupPath
            ws.Cells(r, COL_TEXT).Value = "'" & txt
            ws.Cells(r, COL_CHARS).Value = Len(txt)
            ws.Cells(r, COL_TYPE).Value = "text"
            ws.Cells(r, COL_VOL).Value = "variable"
            written = written + 1
            r = r + 1
        End If
    Next k

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

    On Error Resume Next
    wb.Application.Visible = True
    wb.Activate
    ws.Activate
    ws.Cells(FIRST_ROW, COL_INCLUDE).Select
    On Error GoTo 0

    BuildDiscoverySheet = written & " text shape(s) listed in '" & DISCOVERY_SHEET & "'."
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
    Dim problems As String

    Dim rr As Long
    rr = FIRST_ROW
    Do While Trim(CStr(ws.Cells(rr, COL_ID).Value)) <> ""
        Dim wantId As String
        wantId = Trim(CStr(ws.Cells(rr, COL_ID).Value))

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
    If skippedNoName > 0 Then msg = msg & vbCrLf & skippedNoName & " ticked but unnamed -- not marked."
    If skippedNotFound > 0 Then msg = msg & vbCrLf & skippedNotFound & " shape(s) no longer on the slide."
    If problems <> "" Then msg = msg & vbCrLf & vbCrLf & "NOT MARKED:" & vbCrLf & problems
    If marked > 0 Then msg = msg & vbCrLf & vbCrLf & "Now run 'Setup B: Onboard Slides'."

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
