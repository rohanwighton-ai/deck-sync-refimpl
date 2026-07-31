Attribute VB_Name = "Drafting"
Option Explicit

' The drafting surface: where NEW text comes from.
'
' Everything built before this moves text that already exists -- the register was
' seeded from the slides, so a sync writes the deck back to itself. Real
' machinery, no benefit. This is the piece that closes the loop: a sheet where
' text is drafted against the previous value, approved a row at a time, and
' published into the register as Approved.
'
' Specified by the Excel side in specs/ABOUT_BODY_Field_Package.md §4 and left
' unbuilt. Built generically here rather than for ABOUT_BODY alone, because by
' the time it was written five fields had been through the same path -- the
' shape is known from cases, not guessed.
'
' TWO COLUMNS ARE THE WHOLE DESIGN
'
' The previous value sits immediately left of the draft, read-only. Their §4:
' "Copilot rewrites against a visible exemplar far better than it generates
' cold." That is not a formatting preference -- an AI asked to write a project
' description from nothing produces something generic and plausible, while the
' same AI shown the existing description produces something in the house voice
' at the right level of detail. The exemplar column is the difference between
' "write me a description" and "bring ours up to date".
'
' It also makes the human's job a comparison rather than a judgement in the
' abstract, which is faster and catches more.
'
' ONE DEFINITION OF "YES"
'
' The approval column is named and parsed identically to the sync review grid,
' via ReviewQueue.IsApprovalMark. Three grids had grown three spellings of the
' same idea -- Include (Y/N), Approve (Y/N), Approved -- which is how a tool
' starts feeling homemade, and how a blank cell in one surface comes to mean
' something different from a blank cell in another. Anything that is not an
' affirmative Y/Yes is a no, everywhere.
'
' PUBLISHING IS THE ONLY PLACE Approved IS EVER WRITTEN
'
' Seeding writes Seed. Drafting writes Draft. Only a human ticking a row
' produces Approved, and only through PublishDrafts. That is the whole of the
' synthesis/population boundary: the register's Status column means "a person
' read this and meant it", and this is the one function entitled to say so.

Public Const COL_D_ENTITY As Long = 1
Public Const COL_D_NAME As Long = 2
Public Const COL_D_CURRENT As Long = 3
Public Const COL_D_CHARS As Long = 4
Public Const COL_D_NOTES As Long = 5
Public Const COL_D_DRAFT As Long = 6
Public Const COL_D_APPROVED As Long = 7

Public Const DRAFT_HEADER_ROW As Long = 1
Public Const DRAFT_FIRST_ROW As Long = 2

' Name of a field's drafting sheet. Same hash-disambiguation reasoning as
' ReviewQueue.ReviewSheetNameFor: Excel truncates at 31 characters, and two
' FieldIDs sharing a prefix must not collapse onto one sheet.
Public Function DraftSheetNameFor(fieldId As String) As String
    DraftSheetNameFor = WorkbookBridge.SanitizeSheetName("TPL_" & fieldId)
End Function

' The instruction block that goes above the sheet, for whoever or whatever
' drafts. Pure, so its wording is testable without a workbook.
'
' Transcribed from the field package's own prompt rather than reworded: it
' encodes decisions made by the people who own the content, and paraphrasing
' those would quietly change them.
Public Function DraftingPromptFor(fieldId As String) As String
    Dim s As String
    s = "Read the workbook and the existing text in column C." & vbCrLf & vbCrLf & _
        "Write an updated version of " & fieldId & " for each project into column F ONLY." & vbCrLf & _
        "Leave every other column untouched." & vbCrLf & vbCrLf & _
        "Column C is the standard, not a draft to improve on. Stay close to it in" & vbCrLf & _
        "length and in voice. If the text in column C already does its job, say so" & vbCrLf & _
        "and leave the row blank." & vbCrLf & vbCrLf & _
        "The workbook is the sole source of truth. Do not introduce facts, figures," & vbCrLf & _
        "organisations or outcomes that are not in it. Where something needed is" & vbCrLf & _
        "missing or ambiguous, say so in column E and ask -- do not infer or fill" & vbCrLf & _
        "the gap." & vbCrLf & vbCrLf & _
        "Taciturn voice. No promotional language, no hedging, no padding."
    DraftingPromptFor = s
End Function

' Builds (or rebuilds) a field's drafting sheet from the register.
'
' Reads the REGISTER, not the deck. The register already holds every field's
' current value, and going to the deck would make drafting depend on a deck
' being open -- drafting is workbook work and should be possible on a laptop
' with no deck in front of you.
'
' REBUILDING PRESERVES DRAFT AND NOTES, and clears nothing a human typed. A
' rebuild happens when entities are added or the exemplar moves on, and losing
' half-written drafts to a refresh would make the sheet untrustworthy in exactly
' the way the modal prompts it replaces were. Approvals ARE cleared, because an
' approval is against a specific pairing of exemplar and draft, and a rebuild may
' have changed the exemplar.
Public Function WriteDraftingSheet(ws As Object, reg As Sheet, fieldId As String) As String
    Dim keptDraft As Object, keptNotes As Object
    Set keptDraft = CreateObject("Scripting.Dictionary")
    Set keptNotes = CreateObject("Scripting.Dictionary")

    Dim r As Long
    r = DRAFT_FIRST_ROW
    On Error Resume Next
    Do While Trim(CStr(ws.Cells(r, COL_D_ENTITY).Value)) <> ""
        Dim oldKey As String
        oldKey = Trim(CStr(ws.Cells(r, COL_D_ENTITY).Value))
        If Trim(CStr(ws.Cells(r, COL_D_DRAFT).Value)) <> "" Then keptDraft(oldKey) = CStr(ws.Cells(r, COL_D_DRAFT).Value)
        If Trim(CStr(ws.Cells(r, COL_D_NOTES).Value)) <> "" Then keptNotes(oldKey) = CStr(ws.Cells(r, COL_D_NOTES).Value)
        r = r + 1
    Loop
    On Error GoTo 0

    ws.Cells.Clear

    ws.Cells(DRAFT_HEADER_ROW, COL_D_ENTITY).Value = "EntityCode"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_NAME).Value = "Project name"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_CURRENT).Value = "Current " & fieldId & " (read-only)"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_CHARS).Value = "Chars"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_NOTES).Value = "Your source notes"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_DRAFT).Value = "Draft"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_APPROVED).Value = "Approved (Y/N)"
    ws.Rows(DRAFT_HEADER_ROW).Font.Bold = True

    Dim written As Long, restored As Long
    r = DRAFT_FIRST_ROW

    Dim k As Variant
    For Each k In reg.InstanceOrder
        Dim key As String
        key = CStr(k)
        If reg.Rows.Exists(key) Then
            Dim vals As Object
            Set vals = reg.Rows(key)
            ' Guarded on Exists: reading a missing key from a Scripting.Dictionary
            ' ADDS it and returns Empty, so an unguarded read would silently
            ' invent blank rows for entities this field does not cover.
            Dim current As String
            If vals.Exists(fieldId) Then current = CStr(vals(fieldId))
            Dim projName As String
            If vals.Exists("PROJECT_NAME") Then projName = CStr(vals("PROJECT_NAME"))

            ws.Cells(r, COL_D_ENTITY).Value = key
            ws.Cells(r, COL_D_NAME).Value = "'" & projName
            ws.Cells(r, COL_D_CURRENT).Value = "'" & current
            ws.Cells(r, COL_D_CHARS).Value = Len(current)
            If keptNotes.Exists(key) Then
                ws.Cells(r, COL_D_NOTES).Value = "'" & keptNotes(key)
                restored = restored + 1
            End If
            If keptDraft.Exists(key) Then ws.Cells(r, COL_D_DRAFT).Value = "'" & keptDraft(key)
            ws.Cells(r, COL_D_APPROVED).Value = ""
            written = written + 1
            r = r + 1
        End If
    Next k

    ws.Columns(COL_D_ENTITY).ColumnWidth = 12
    ws.Columns(COL_D_NAME).ColumnWidth = 34
    ws.Columns(COL_D_CURRENT).ColumnWidth = 60
    ws.Columns(COL_D_CHARS).ColumnWidth = 7
    ws.Columns(COL_D_NOTES).ColumnWidth = 26
    ws.Columns(COL_D_DRAFT).ColumnWidth = 60
    ws.Columns(COL_D_APPROVED).ColumnWidth = 14
    ws.Columns(COL_D_CURRENT).WrapText = True
    ws.Columns(COL_D_DRAFT).WrapText = True

    WriteDraftingSheet = written & " row(s) written for " & fieldId & _
        IIf(restored > 0, ", " & restored & " existing note(s)/draft(s) preserved", "") & "."
End Function

' Reads the sheet back and publishes approved drafts into the register.
'
' A row publishes only when BOTH the draft is non-empty AND the approval is an
' affirmative Y. Either alone is not consent: a draft with no tick has not been
' read, and a tick against an empty draft is a mis-click.
'
' Real line breaks become "||" on the way in, which is the register's own
' convention and the exact inverse of what InjectPrimitive does on the way out --
' proven exact by VerifyHarvest on 46 of 46 multi-line values. Any cell still
' holding a carriage return after conversion FAILS rather than being written,
' per the field package's publish step.
Public Function PublishDrafts(ws As Object, regWs As Object, fieldId As String, _
                              Optional dryRun As Boolean = True) As String
    Dim report As String
    report = IIf(dryRun, "=== PREVIEW: publish " & fieldId & " ===", "=== Publish " & fieldId & " ===") & vbCrLf

    Dim hdr As Object
    Set hdr = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To 20
        Dim h As String
        h = Trim(CStr(regWs.Cells(1, c).Value))
        If h <> "" Then hdr(h) = c
    Next c

    Dim needed As Variant
    For Each needed In Array("EntityCode", "FieldID", "Value", "Status")
        If Not hdr.Exists(needed) Then
            PublishDrafts = report & "STOPPED: register has no '" & needed & "' column."
            Exit Function
        End If
    Next needed

    Dim published As Long, skippedNoTick As Long, skippedEmpty As Long, failed As Long, noRow As Long

    Dim r As Long
    r = DRAFT_FIRST_ROW
    Do While Trim(CStr(ws.Cells(r, COL_D_ENTITY).Value)) <> ""
        Dim ent As String
        ent = Trim(CStr(ws.Cells(r, COL_D_ENTITY).Value))
        Dim draft As String
        draft = CStr(ws.Cells(r, COL_D_DRAFT).Value)
        Dim tick As Boolean
        tick = ReviewQueue.IsApprovalMark(CStr(ws.Cells(r, COL_D_APPROVED).Value))

        If Trim(draft) = "" Then
            If tick Then
                skippedEmpty = skippedEmpty + 1
                report = report & "  SKIPPED " & ent & " -- ticked but the draft is empty" & vbCrLf
            End If
        ElseIf Not tick Then
            skippedNoTick = skippedNoTick + 1
        Else
            Dim encoded As String
            encoded = Replace(draft, vbCrLf, "||")
            encoded = Replace(encoded, vbCr, "||")
            encoded = Replace(encoded, vbLf, "||")
            encoded = Replace(encoded, Chr(11), "||")

            If InStr(encoded, vbCr) > 0 Or InStr(encoded, vbLf) > 0 Then
                failed = failed + 1
                report = report & "  FAILED " & ent & " -- a line break survived conversion" & vbCrLf
            Else
                Dim target As Long
                target = FindRegisterRow(regWs, hdr, ent, fieldId)
                If target = 0 Then
                    noRow = noRow + 1
                    report = report & "  NO REGISTER ROW for " & ent & "/" & fieldId & vbCrLf
                Else
                    If Not dryRun Then
                        regWs.Cells(target, hdr("Value")).Value = encoded
                        regWs.Cells(target, hdr("Status")).Value = Register.STATUS_APPROVED
                        If hdr.Exists("CharCount") Then regWs.Cells(target, hdr("CharCount")).Value = CStr(Len(encoded))
                        If hdr.Exists("UpdatedDate") Then regWs.Cells(target, hdr("UpdatedDate")).Value = Format(Now, "yyyy-mm-dd")
                    End If
                    published = published + 1
                    report = report & "  " & IIf(dryRun, "would publish: ", "published: ") & ent & _
                        " (" & Len(encoded) & " chars)" & vbCrLf
                End If
            End If
        End If
        r = r + 1
    Loop

    report = report & vbCrLf & "Summary: " & published & IIf(dryRun, " would be published", " published") & _
        ", " & skippedNoTick & " drafted but not ticked, " & skippedEmpty & " ticked but empty, " & _
        noRow & " with no register row, " & failed & " failed" & vbCrLf

    If skippedNoTick > 0 Then
        report = report & vbCrLf & skippedNoTick & " draft(s) are waiting on a tick. Nothing reaches a" & vbCrLf & _
            "slide until they have one." & vbCrLf
    End If

    PublishDrafts = report
End Function

Private Function FindRegisterRow(regWs As Object, hdr As Object, entity As String, fieldId As String) As Long
    Dim r As Long
    r = 2
    Do While Trim(CStr(regWs.Cells(r, hdr("EntityCode")).Value)) <> ""
        If Trim(CStr(regWs.Cells(r, hdr("EntityCode")).Value)) = entity Then
            If Trim(CStr(regWs.Cells(r, hdr("FieldID")).Value)) = fieldId Then
                FindRegisterRow = r
                Exit Function
            End If
        End If
        r = r + 1
    Loop
    FindRegisterRow = 0
End Function
