Attribute VB_Name = "FieldSpec"
Option Explicit

' Per-field drafting guidance, held in the workbook rather than in code.
'
' The drafting prompt used to be generated generically: "write an updated
' version of <FIELD>, stay close to column C, do not invent facts." True for
' every field and specifically right for none. ABOUT_BODY and KEY_EVENTS_BODY
' want almost opposite things -- one is a standing description of what a project
' IS, the other a list of what happened this quarter -- and both were being told
' the same thing.
'
' The Excel side had already written the real guidance for ABOUT_BODY in
' specs/ABOUT_BODY_Field_Package.md: purpose, voice, length, an own-job test and
' an explicit "do not". None of it reached the sheet.
'
' IN THE WORKBOOK, NOT IN CODE, and that is the whole point. This is content
' policy owned by whoever owns the report -- the RM, not the tool. Changing how
' a field should be written must be editing a spreadsheet row, not a code change
' and a rebuild. Putting it in VBA would have made every wording decision a
' developer task, which is exactly the dependency this project exists to remove.
'
' A MISSING SPEC IS NOT AN ERROR. A field with no row still drafts, using the
' generic guidance -- because refusing to draft a field until somebody writes
' its style guide would block the work on paperwork. The prompt says plainly
' when it is running unguided, so the gap is visible rather than silent.

Public Const SPEC_SHEET_NAME As String = "Field Spec"

Public Const COL_S_FIELDID As Long = 1
Public Const COL_S_KIND As Long = 2
Public Const COL_S_PURPOSE As Long = 3
Public Const COL_S_VOICE As Long = 4
Public Const COL_S_LENGTH As Long = 5
Public Const COL_S_OWNJOB As Long = 6
Public Const COL_S_DONOT As Long = 7

Public Const SPEC_HEADER_ROW As Long = 1
Public Const SPEC_FIRST_ROW As Long = 2

Public Type FieldGuidance
    Found As Boolean
    FieldId As String
    Kind As String          ' Controlled / Prose / Static -- see ReviewQueue
    Purpose As String
    Voice As String
    Length As String
    OwnJob As String
    DoNot As String
End Type

' Creates the sheet with its headers and, on a fresh workbook, the rows already
' settled by the Excel side. Existing rows are never overwritten -- this is
' their content, and a rebuild that silently reverted an edit would make the
' sheet untrustworthy in the same way a lost draft would.
Public Function WriteSpecSheet(ws As Object) As String
    Dim existing As Object
    Set existing = CreateObject("Scripting.Dictionary")

    Dim r As Long
    r = SPEC_FIRST_ROW
    On Error Resume Next
    Do While Trim(CStr(ws.Cells(r, COL_S_FIELDID).Value)) <> ""
        existing(UCase(Trim(CStr(ws.Cells(r, COL_S_FIELDID).Value)))) = True
        r = r + 1
    Loop
    On Error GoTo 0

    ws.Cells(SPEC_HEADER_ROW, COL_S_FIELDID).Value = "FieldID"
    ws.Cells(SPEC_HEADER_ROW, COL_S_KIND).Value = "Kind (Controlled/Prose/Static)"
    ws.Cells(SPEC_HEADER_ROW, COL_S_PURPOSE).Value = "Purpose -- the question this field answers"
    ws.Cells(SPEC_HEADER_ROW, COL_S_VOICE).Value = "Voice"
    ws.Cells(SPEC_HEADER_ROW, COL_S_LENGTH).Value = "Length"
    ws.Cells(SPEC_HEADER_ROW, COL_S_OWNJOB).Value = "Own-job test"
    ws.Cells(SPEC_HEADER_ROW, COL_S_DONOT).Value = "Do NOT"
    ws.Rows(SPEC_HEADER_ROW).Font.Bold = True
    ws.Rows(SPEC_HEADER_ROW).WrapText = True

    Dim added As Long
    If Not existing.Exists("ABOUT_BODY") Then
        SeedRow ws, r, "ABOUT_BODY", "Prose", _
            "The ""what"". A neutral, factual description of what the project is and does: the approach or technology, the target systems, and the aim.", _
            "Descriptive, present tense. Taciturn. No promotional language, no hedging, no padding. Assumes AMR literacy.", _
            "About one paragraph. Target 200-350 characters, advisory. The real anchor is this project's own prior value, not the portfolio median.", _
            "Does it describe WHAT THE PROJECT IS, without straying into why it is needed or what it is worth?", _
            "Justify the project. Pitch its strategic value. Describe this quarter's activity. Invent facts, figures, organisations or outcomes not in the workbook."
        r = r + 1: added = added + 1
    End If
    If Not existing.Exists("KEY_EVENTS_BODY") Then
        SeedRow ws, r, "KEY_EVENTS_BODY", "Prose", _
            "What actually happened this quarter. One line per event, most significant first.", _
            "Factual and dated. Each line stands alone. No narrative linking between events.", _
            "Multi-line, one event per line, separated by || in the register. Typically 3-6 lines.", _
            "Is every line something that HAPPENED in this period, rather than a standing fact about the project?", _
            "Restate what the project is (that is ABOUT_BODY). List intentions or plans as though they had occurred. Pad to fill space."
        r = r + 1: added = added + 1
    End If
    If Not existing.Exists("PROJECT_STATUS") Then
        SeedRow ws, r, "PROJECT_STATUS", "Controlled", _
            "The project's current state, from a fixed vocabulary.", _
            "One of the agreed values, exactly as spelled. Not a sentence.", _
            "A single controlled value.", _
            "Is it one of the agreed values, character for character?", _
            "Invent a new status. Vary the capitalisation. Explain or qualify it."
        r = r + 1: added = added + 1
    End If

    ws.Columns(COL_S_FIELDID).ColumnWidth = 20
    ws.Columns(COL_S_KIND).ColumnWidth = 16
    Dim c As Long
    For c = COL_S_PURPOSE To COL_S_DONOT
        ws.Columns(c).ColumnWidth = 46
        ws.Columns(c).WrapText = True
    Next c

    WriteSpecSheet = "Field Spec: " & existing.count & " row(s) kept, " & added & " seeded."
End Function

Private Sub SeedRow(ws As Object, r As Long, fieldId As String, kind As String, _
                    purpose As String, voice As String, length As String, _
                    ownJob As String, doNot As String)
    ws.Cells(r, COL_S_FIELDID).Value = fieldId
    ws.Cells(r, COL_S_KIND).Value = kind
    ws.Cells(r, COL_S_PURPOSE).Value = "'" & purpose
    ws.Cells(r, COL_S_VOICE).Value = "'" & voice
    ws.Cells(r, COL_S_LENGTH).Value = "'" & length
    ws.Cells(r, COL_S_OWNJOB).Value = "'" & ownJob
    ws.Cells(r, COL_S_DONOT).Value = "'" & doNot
End Sub

' Looks a field up. Found=False when there is no row -- the caller falls back to
' generic guidance and says so, rather than refusing to draft.
Public Function LookupGuidance(ws As Object, fieldId As String) As FieldGuidance
    Dim g As FieldGuidance
    g.FieldId = fieldId

    If ws Is Nothing Then
        LookupGuidance = g
        Exit Function
    End If

    Dim r As Long
    r = SPEC_FIRST_ROW
    On Error Resume Next
    Do While Trim(CStr(ws.Cells(r, COL_S_FIELDID).Value)) <> ""
        If StrComp(Trim(CStr(ws.Cells(r, COL_S_FIELDID).Value)), fieldId, vbTextCompare) = 0 Then
            g.Found = True
            g.Kind = Trim(CStr(ws.Cells(r, COL_S_KIND).Value))
            g.Purpose = Trim(CStr(ws.Cells(r, COL_S_PURPOSE).Value))
            g.Voice = Trim(CStr(ws.Cells(r, COL_S_VOICE).Value))
            g.Length = Trim(CStr(ws.Cells(r, COL_S_LENGTH).Value))
            g.OwnJob = Trim(CStr(ws.Cells(r, COL_S_OWNJOB).Value))
            g.DoNot = Trim(CStr(ws.Cells(r, COL_S_DONOT).Value))
            Exit Do
        End If
        r = r + 1
    Loop
    On Error GoTo 0

    LookupGuidance = g
End Function

' The prompt, built from the field's own guidance.
'
' Pure -- takes the guidance, not a worksheet -- so every wording decision here
' is testable without Excel. The generic fallback is deliberately explicit that
' it IS the fallback: an AI told "write this field" with no further guidance
' produces something plausible and unusable, and the person reading the output
' should know which of those they are looking at.
Public Function PromptFrom(g As FieldGuidance) As String
    Dim s As String
    s = "Read the workbook and the existing text in column C." & vbCrLf & vbCrLf

    If Not g.Found Then
        s = s & "NOTE: no Field Spec row exists for " & g.FieldId & ", so this is generic" & vbCrLf & _
            "guidance. Add a row to the 'Field Spec' sheet to say how this field" & vbCrLf & _
            "should be written." & vbCrLf & vbCrLf & _
            "Write an updated version of " & g.FieldId & " for each project into column F ONLY." & vbCrLf & vbCrLf
    Else
        s = s & "WHAT THIS FIELD IS FOR" & vbCrLf & g.Purpose & vbCrLf & vbCrLf & _
            "VOICE" & vbCrLf & g.Voice & vbCrLf & vbCrLf & _
            "LENGTH" & vbCrLf & g.Length & vbCrLf & vbCrLf & _
            "BEFORE YOU WRITE EACH ROW, ASK" & vbCrLf & g.OwnJob & vbCrLf & vbCrLf & _
            "DO NOT" & vbCrLf & g.DoNot & vbCrLf & vbCrLf & _
            "Write into column F ONLY. Leave every other column untouched." & vbCrLf & vbCrLf
    End If

    s = s & "Column C is the standard, not a draft to improve on. Stay close to it in" & vbCrLf & _
        "length and in voice. If the text in column C already does its job, say so" & vbCrLf & _
        "and leave the row blank." & vbCrLf & vbCrLf & _
        "The workbook is the sole source of truth. Do not introduce facts, figures," & vbCrLf & _
        "organisations or outcomes that are not in it. Where something needed is" & vbCrLf & _
        "missing or ambiguous, say so in column E and ask -- do not infer or fill" & vbCrLf & _
        "the gap."

    PromptFrom = s
End Function
