Attribute VB_Name = "Sources"
Option Explicit

' WHERE THE WORDS CAME FROM.
'
' A drafting sheet records what a field was changed TO. It has never recorded
' what that change was based on, so a quarter later "why does it say 90%?" has
' no answer except somebody's memory. This sheet is that answer.
'
' ONE ROW PER SOURCE, EVER. Drafting rows refer to sources by ID, not by
' description. A milestone report used by all 43 projects is written down once
' and referenced 43 times -- which is the difference between a record and a
' mess. Rohan's constraint, 2026-08-01: "Sources should be managed so it
' doesn't get silly."
'
' The three things that make it stay manageable:
'   1. IDs are short and stable (S01, S02 ...). Retyping a label 43 times gives
'      43 spellings; retyping an ID gives an ID.
'   2. The sheet is never rebuilt from scratch -- rows accumulate and survive
'      every drafting refresh, because a source outlives the field it was used
'      for.
'   3. A reference to an ID that does not exist is REPORTED, not silently
'      ignored. An unnoticed typo turns the record into decoration.
'
' Deliberately NOT here: source content. This holds pointers -- what it is,
' where it lives. Pasting document text into a spreadsheet is how the sheet
' becomes unreadable, which is the failure mode being designed against.

Public Const SOURCES_SHEET_NAME As String = "Sources"

Public Const COL_S_ID As Long = 1
Public Const COL_S_LABEL As Long = 2
Public Const COL_S_TYPE As Long = 3
Public Const COL_S_LOCATOR As Long = 4
Public Const COL_S_NOTES As Long = 5
Public Const COL_S_ADDED As Long = 6
Public Const COL_S_APPLIES As Long = 7

' SPELLED OUT, and deliberately NOT the token "ALL" -- DECISIONS.md 2026-08-04.
' "ALL" was the sentinel the wide-sheet model replaced, and reusing the word for
' a second meaning is the "one word doing two jobs" defect this project has
' already paid for once, when "static" meant both "not authored by a person" and
' "does not change between periods".
Public Const APPLIES_ALL As String = "All periods"

Public Const SRC_INTRO_ROW As Long = 1
Public Const SRC_HEADER_ROW As Long = 5
Public Const SRC_FIRST_ROW As Long = 6

' Creates the sheet if it is absent and (re)writes only the furniture --
' headings, instructions, widths. EXISTING ROWS ARE NEVER TOUCHED: this runs on
' every drafting build, and a function that ran 40 times a session and could
' delete the provenance record would be a liability rather than a convenience.
Public Function WriteSourcesSheet(ws As Object) As String
    ws.Cells(SRC_INTRO_ROW, 1).Value = "SOURCES  --  what the drafting sheets are based on"
    ws.Cells(SRC_INTRO_ROW, 1).Font.Bold = True
    ws.Cells(SRC_INTRO_ROW, 1).Font.Size = 9

    ws.Cells(2, 1).Value = "Add a row for anything you drafted from. Give it the next free ID, then put that ID in column G of the drafting sheet."
    ws.Cells(3, 1).Value = "One row per source, ever -- if 20 projects use the same report, they all reference the same ID. Do not paste document text in here; point at it."

    ws.Cells(4, 1).Value = "A drafting row referring to an ID that is not listed here is reported when you publish. It will not stop the publish, but it means the record is wrong."
    ws.Cells(4, 1).Font.Italic = True

    ws.Cells(3, COL_S_APPLIES).Value = "Applies to: pick " & APPLIES_ALL & " for something cited in every period (a contract, a project description), or pick the one period it belongs to (a period's own report). The list offers the periods your register actually holds. Blank counts as " & APPLIES_ALL & "."

    ws.Cells(SRC_HEADER_ROW, COL_S_ID).Value = "ID"
    ws.Cells(SRC_HEADER_ROW, COL_S_LABEL).Value = "What it is"
    ws.Cells(SRC_HEADER_ROW, COL_S_TYPE).Value = "Type"
    ws.Cells(SRC_HEADER_ROW, COL_S_LOCATOR).Value = "Where it lives"
    ws.Cells(SRC_HEADER_ROW, COL_S_NOTES).Value = "Notes"
    ws.Cells(SRC_HEADER_ROW, COL_S_ADDED).Value = "Added"
    ws.Cells(SRC_HEADER_ROW, COL_S_APPLIES).Value = "Applies to"
    ws.Rows(SRC_HEADER_ROW).Font.Bold = True

    ws.Cells.Font.Size = 8
    ws.Cells(SRC_INTRO_ROW, 1).Font.Size = 9
    ws.Columns(COL_S_ID).ColumnWidth = 7
    ws.Columns(COL_S_LABEL).ColumnWidth = 44
    ws.Columns(COL_S_TYPE).ColumnWidth = 10
    ws.Columns(COL_S_LOCATOR).ColumnWidth = 46
    ws.Columns(COL_S_NOTES).ColumnWidth = 34
    ws.Columns(COL_S_ADDED).ColumnWidth = 11
    ws.Columns(COL_S_APPLIES).ColumnWidth = 14
    ws.Cells.VerticalAlignment = -4160        ' xlTop

    Dim n As Long
    n = CountSources(ws)
    WriteSourcesSheet = "Sources: " & n & " source(s) on record. Next free ID is " & NextSourceId(ws) & "."
End Function

Public Function CountSources(ws As Object) As Long
    Dim r As Long
    r = SRC_FIRST_ROW
    Do While Trim(CStr(ws.Cells(r, COL_S_ID).Value)) <> ""
        CountSources = CountSources + 1
        r = r + 1
    Loop
End Function

' The lowest unused S-number, so two people adding a source on the same day do
' not both reach for S04. Numeric, not "count + 1": a deleted row in the middle
' would otherwise hand out an ID that is already in use further down.
Public Function NextSourceId(ws As Object) As String
    Dim highest As Long
    Dim r As Long
    r = SRC_FIRST_ROW
    Do While Trim(CStr(ws.Cells(r, COL_S_ID).Value)) <> ""
        Dim raw As String
        raw = UCase(Trim(CStr(ws.Cells(r, COL_S_ID).Value)))
        If Left(raw, 1) = "S" Then
            Dim numPart As String
            numPart = Mid(raw, 2)
            If IsNumeric(numPart) Then
                If CLng(numPart) > highest Then highest = CLng(numPart)
            End If
        End If
        r = r + 1
    Loop
    NextSourceId = "S" & Format(highest + 1, "00")
End Function

' Every ID currently on the sheet, upper-cased, as a Dictionary for lookup.
Public Function KnownSourceIds(ws As Object) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim r As Long
    r = SRC_FIRST_ROW
    Do While Trim(CStr(ws.Cells(r, COL_S_ID).Value)) <> ""
        d(UCase(Trim(CStr(ws.Cells(r, COL_S_ID).Value)))) = CStr(ws.Cells(r, COL_S_LABEL).Value)
        r = r + 1
    Loop
    Set KnownSourceIds = d
End Function

' THE PERIOD IS PICKED, NEVER TYPED.
'
' Periods are free text matched EXACTLY, so "Q4F26", "Q4 F26" and "q4f26" are
' three different periods, and a period that matches nothing reads as a clean run
' of zero rows. Every other guard in this project sits downstream of the typing;
' this one removes the typing. The list is not invented here either -- it is the
' periods the register actually holds, so a source can only be stamped with a
' period that exists.
'
' This is also what makes the time scale a non-question: months, halves or
' anything else work unchanged, because nothing here parses a period. It only
' offers back what is already on the sheet.
'
' FAILS LOUD. Excel's list argument is capped near 255 characters; past that the
' dropdown silently does not apply, and a validation that quietly is not there is
' worse than none -- it reads as care taken and stops anyone re-checking.
Public Function ApplyPeriodValidation(srcWs As Object, regWs As Object) As String
    If srcWs Is Nothing Or regWs Is Nothing Then
        ApplyPeriodValidation = "Sources validation: skipped (no Sources sheet or no register)."
        Exit Function
    End If

    Dim cInstance As Long, cQuarter As Long
    ExcelOutput.LocateStructuralColumns regWs, cInstance, cQuarter
    If cQuarter = 0 Then
        ApplyPeriodValidation = "Sources validation: the register has no '" & _
            ExcelOutput.QUARTER_HEADER & "' column, so there are no periods to offer."
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ExcelOutput.LastUsedRow(regWs)

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim list As String
    list = APPLIES_ALL

    Dim r As Long
    For r = 2 To lastRow
        Dim p As String
        p = Trim(CStr(regWs.Cells(r, cQuarter).Value))
        If p <> "" Then
            If Not seen.Exists(UCase(p)) Then
                seen(UCase(p)) = True
                list = list & "," & p
            End If
        End If
    Next r

    If seen.count = 0 Then
        ApplyPeriodValidation = "Sources validation: the register holds no periods yet."
        Exit Function
    End If

    If Len(list) > 255 Then
        ApplyPeriodValidation = "Sources validation: NOT APPLIED -- " & seen.count & _
            " periods is more than Excel's dropdown can hold (" & Len(list) & _
            " characters, limit 255). Type the period by hand and check the spelling."
        Exit Function
    End If

    Dim lastSrcRow As Long
    lastSrcRow = SRC_FIRST_ROW + CountSources(srcWs) - 1
    ' Room to grow: the rows below the last source get the dropdown too, so a
    ' source added tomorrow is constrained the same way as one added today.
    If lastSrcRow < SRC_FIRST_ROW then lastSrcRow = SRC_FIRST_ROW
    lastSrcRow = lastSrcRow + 50

    Dim rng As Object
    Set rng = srcWs.Range(srcWs.Cells(SRC_FIRST_ROW, COL_S_APPLIES), _
                          srcWs.Cells(lastSrcRow, COL_S_APPLIES))

    Dim failed As Boolean
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add 3, 1, 1, list          ' xlValidateList, xlValidAlertStop, xlBetween
    rng.Validation.IgnoreBlank = True
    rng.Validation.InCellDropdown = True
    If Err.Number <> 0 Then failed = True
    Err.Clear
    On Error GoTo 0

    If failed Then
        ApplyPeriodValidation = "Sources validation: NOT APPLIED -- Excel refused " & _
            "(the workbook may be read-only or the sheet protected)."
        Exit Function
    End If

    ' Anything already on the sheet that is not in the list. Existing rows are
    ' never rewritten -- they hold the provenance record -- so this is reported.
    Dim offending As String
    Dim n As Long
    For r = SRC_FIRST_ROW To SRC_FIRST_ROW + CountSources(srcWs) - 1
        Dim cur As String
        cur = Trim(CStr(srcWs.Cells(r, COL_S_APPLIES).Value))
        If cur <> "" And Not seen.Exists(UCase(cur)) And _
           StrComp(cur, APPLIES_ALL, vbTextCompare) <> 0 Then
            offending = offending & "  " & Trim(CStr(srcWs.Cells(r, COL_S_ID).Value)) & _
                " = """ & cur & """" & vbCrLf
            n = n + 1
        End If
    Next r

    ApplyPeriodValidation = "Sources validation: dropdown on 'Applies to' with " & _
        (seen.count + 1) & " choice(s)." & _
        IIf(n = 0, "", vbCrLf & n & " existing source(s) name a period the register " & _
            "does not have:" & vbCrLf & offending)
End Function

' What each source ID says it applies to, upper-cased, as a Dictionary.
' A blank cell -- which is every row on a sheet written before this column
' existed -- reads as APPLIES_ALL. That is a real default and it is chosen
' deliberately: the alternative makes every existing source start reporting
' against every period, which would train you to ignore the report on the day
' it was introduced.
Public Function SourceApplicability(ws As Object) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim r As Long
    r = SRC_FIRST_ROW
    Do While Trim(CStr(ws.Cells(r, COL_S_ID).Value)) <> ""
        Dim applies As String
        applies = Trim(CStr(ws.Cells(r, COL_S_APPLIES).Value))
        If applies = "" Then applies = APPLIES_ALL
        d(UCase(Trim(CStr(ws.Cells(r, COL_S_ID).Value)))) = UCase(applies)
        r = r + 1
    Loop
    Set SourceApplicability = d
End Function

' Which of the IDs cited by a drafting row belong to a DIFFERENT period than the
' one being published. Returns "" when every citation is either period-neutral
' or names this period.
'
' REPORTED, NEVER ENFORCED -- the same posture as UnknownRefs directly below, and
' for the same reason: refusing to publish a quarter's text over a citation is
' the tool getting in the way of the work it exists to do. A wrong citation makes
' the provenance record wrong, which is worth saying out loud and is not worth
' blocking the evening for.
'
' An ID that is not on the sheet at all is NOT reported here. UnknownRefs already
' owns that message, and saying it twice for one typo reads as two faults.
Public Function RefsForOtherPeriod(refs As String, applic As Object, period As String) As String
    If Trim(refs) = "" Then Exit Function
    If applic Is Nothing Then Exit Function
    If Trim(period) = "" Then Exit Function

    Dim parts As Variant
    parts = Split(Replace(Replace(refs, ";", ","), " ", ","), ",")

    Dim wrong As String
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim one As String
        one = UCase(Trim(CStr(parts(i))))
        If one <> "" Then
            If applic.Exists(one) Then
                Dim applies As String
                applies = applic(one)
                If applies <> UCase(APPLIES_ALL) And applies <> UCase(Trim(period)) Then
                    If InStr(wrong & ",", one & ",") = 0 Then
                        wrong = wrong & IIf(wrong = "", "", ",") & one & " (" & applies & ")"
                    End If
                End If
            End If
        End If
    Next i

    RefsForOtherPeriod = wrong
End Function

' Which of the IDs in a drafting row's Sources cell are not on the Sources
' sheet. Returns "" when they all check out.
'
' Pure and separated deliberately: it is the only part of this module with a
' decision in it, and a decision that cannot be driven both ways from a test is
' one nobody ever watches fail.
Public Function UnknownRefs(refs As String, known As Object) As String
    If Trim(refs) = "" Then Exit Function
    If known Is Nothing Then Exit Function

    Dim parts As Variant
    parts = Split(Replace(Replace(refs, ";", ","), " ", ","), ",")

    Dim bad As String
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim one As String
        one = UCase(Trim(CStr(parts(i))))
        If one <> "" Then
            If Not known.Exists(one) Then
                If InStr(bad & ",", one & ",") = 0 Then
                    bad = bad & IIf(bad = "", "", ",") & one
                End If
            End If
        End If
    Next i
    UnknownRefs = bad
End Function

' The sources cited on a drafting sheet, written out for the prompt.
'
' THE FLOW USED TO STOP HERE, AND THAT WAS THE WHOLE PROBLEM. Sources were
' consulted at PUBLISH only -- KnownSourceIds, UnknownRefs, RefsForOtherPeriod
' checked that a citation was valid and bound to the right period. Nothing ever
' put the cited document in front of the thing doing the writing.
'
' So a recipe could forbid presenting an inferred fact as a declared one, the
' fact could be declared in a cited document, and the model still had no way to
' confirm it -- because the prompt said "the workbook is the sole source of
' truth" and never mentioned the citation. WORKED-EXAMPLE-STRATEGIC-ALIGNMENT.md
' records exactly that outcome: "Declared linkages: [TBC]", and it would have
' stayed [TBC] however many sources were cited.
'
' Returns "" when nothing on the sheet cites anything, so a field that needs no
' evidence gets no extra words.
Public Function CitedBlockFor(srcWs As Object, draftWs As Object, _
                              sourcesCol As Long, firstRow As Long) As String
    If srcWs Is Nothing Or draftWs Is Nothing Then Exit Function

    ' Which IDs does this sheet actually cite, in the order first seen.
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim order As Object
    Set order = CreateObject("Scripting.Dictionary")

    Dim r As Long
    r = firstRow
    Do While Trim(CStr(draftWs.Cells(r, 1).Value)) <> ""
        Dim raw As String
        raw = Trim(CStr(draftWs.Cells(r, sourcesCol).Value))
        If raw <> "" Then
            Dim parts() As String
            parts = Split(Replace(raw, ";", ","), ",")
            Dim i As Long
            For i = LBound(parts) To UBound(parts)
                Dim id As String
                id = UCase(Trim(parts(i)))
                If id <> "" Then
                    If Not seen.Exists(id) Then
                        seen(id) = True
                        order(order.Count) = id
                    End If
                End If
            Next i
        End If
        r = r + 1
    Loop

    If order.Count = 0 Then Exit Function

    Dim s As String
    s = vbCrLf & vbCrLf & "SOURCES CITED ON THIS SHEET" & vbCrLf & _
        "These are named in column G against the rows that rely on them. Each" & vbCrLf & _
        "row's own column G says which apply TO THAT ROW -- do not carry a" & vbCrLf & _
        "source across to a row that does not cite it." & vbCrLf & vbCrLf

    Dim k As Long
    For k = 0 To order.Count - 1
        Dim wantId As String
        wantId = CStr(order(k))
        Dim rr As Long
        rr = SRC_FIRST_ROW
        Dim found As Boolean
        found = False
        Do While Trim(CStr(srcWs.Cells(rr, COL_S_ID).Value)) <> ""
            If UCase(Trim(CStr(srcWs.Cells(rr, COL_S_ID).Value))) = wantId Then
                s = s & wantId & "  " & Trim(CStr(srcWs.Cells(rr, COL_S_LABEL).Value)) & vbCrLf & _
                    "    kind:     " & Trim(CStr(srcWs.Cells(rr, COL_S_TYPE).Value)) & vbCrLf & _
                    "    where:    " & Trim(CStr(srcWs.Cells(rr, COL_S_LOCATOR).Value)) & vbCrLf & _
                    "    applies:  " & Trim(CStr(srcWs.Cells(rr, COL_S_APPLIES).Value)) & vbCrLf & vbCrLf
                found = True
                Exit Do
            End If
            rr = rr + 1
        Loop
        ' A cited ID with no Sources row is reported at publish. Saying so here
        ' too stops the model treating an unknown citation as evidence.
        If Not found Then
            s = s & wantId & "  -- NOT ON THE SOURCES SHEET. Do not treat this as evidence." & vbCrLf & vbCrLf
        End If
    Next k

    s = s & "The workbook AND these cited documents are your evidence. You may" & vbCrLf & _
        "confirm a fact from a document a row cites. You may NOT introduce" & vbCrLf & _
        "anything that is in neither." & vbCrLf & vbCrLf & _
        "IF YOU CANNOT OPEN ONE, SAY SO in column J and treat its facts as" & vbCrLf & _
        "unconfirmed. Do not infer the contents of a document you could not" & vbCrLf & _
        "read -- an unopened source is not a source."

    CitedBlockFor = s
End Function
