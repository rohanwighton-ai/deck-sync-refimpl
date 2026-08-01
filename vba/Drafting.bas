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

' THREE TEXTS PER ROW, AND ONLY ONE OF THEM CAN REACH A SLIDE.
'
' ORIGINAL   what the slide says today. Read-only.
' AI DRAFT   what Copilot wrote. Never published -- not once, not ever.
' SUBMIT     what YOU are sending. Starts empty. Publish reads this and
'            nothing else.
'
' The split exists because "the AI wrote it and I didn't stop it" and "I sent
' it" are different acts, and the old single Draft column could not tell them
' apart. Copy AI -> Submit is one action away and fills only EMPTY Submit
' cells, so re-running it can never overwrite an edit you made.
Public Const COL_D_ENTITY As Long = 1
Public Const COL_D_NAME As Long = 2
Public Const COL_D_CURRENT As Long = 3
Public Const COL_D_CHARS As Long = 4
Public Const COL_D_SOURCES As Long = 5
Public Const COL_D_DRAFT As Long = 6
Public Const COL_D_SUBMIT As Long = 7
Public Const COL_D_SUBCHARS As Long = 8
Public Const COL_D_APPROVED As Long = 9
Public Const COL_D_NOTES As Long = 10
Public Const COL_D_LAYOUT As Long = 11
Public Const COL_D_PROMPT As Long = 12

' THE SHEET DECLARES WHICH LAYOUT IT WAS WRITTEN IN.
'
' Bump this whenever a COL_D_* constant changes meaning. Carrying a person's
' work across a rebuild means reading the OLD sheet with column numbers -- and
' if those numbers have been renumbered underneath, the read is silently
' wrong. It is not a crash; it is column 7 meaning "tick" on Monday and
' "SUBMIT" on Tuesday, so three ticks arrive as three pieces of publishable
' text.
'
' Caught live 2026-08-01 the first time this layout changed, by a count that
' did not match: "3 left alone (you had already written something there)" when
' only one row had been written. The number was the only evidence; nothing
' raised, and the sheet looked fine.
Public Const DRAFT_LAYOUT_VERSION As Long = 2

' The instruction block occupies rows 1-7, so the grid starts lower.
'
' The instructions used to be PRINTED TO THE CONSOLE when the sheet was built,
' and the sheet itself said nothing. Rohan opened it cold on 2026-07-31 and
' could not tell what he was being asked to do -- "no instruction as to what to
' type and where or how the information goes when I do". Correct, and the whole
' reason a human runs step 1 of the manual test.
'
' A form with its covering letter thrown away is not a form. They live ON the
' sheet now, where the person is.
Public Const DRAFT_INTRO_ROW As Long = 1
Public Const DRAFT_HEADER_ROW As Long = 9
Public Const DRAFT_FIRST_ROW As Long = 10

' Name of a field's drafting sheet. Same hash-disambiguation reasoning as
' ReviewQueue.ReviewSheetNameFor: Excel truncates at 31 characters, and two
' FieldIDs sharing a prefix must not collapse onto one sheet.
Public Function DraftSheetNameFor(fieldId As String) As String
    DraftSheetNameFor = WorkbookBridge.SanitizeSheetName("TPL_" & fieldId)
End Function

' Superseded by FieldSpec.PromptFrom, which builds the prompt from the field's
' own row in the workbook rather than from a generic template baked into code.
' Kept as the fallback wording only; callers pass a FieldGuidance.
Public Function DraftingPromptFor(fieldId As String) As String
    Dim g As FieldGuidance
    g.FieldId = fieldId
    DraftingPromptFor = FieldSpec.PromptFrom(g)
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
Public Function WriteDraftingSheet(ws As Object, reg As Sheet, fieldId As String, _
                                   Optional guidance As Variant) As String
    ' A REBUILD MUST NOT COST A PERSON THEIR WORK. Everything a human or an AI
    ' put on this sheet is carried across: the AI draft, the SUBMIT text they
    ' edited, the source IDs they assigned, and their notes. Only ORIGINAL and
    ' the character counts are re-derived, because only those come from the
    ' register. Losing a column here would be silent and would cost an evening.
    Dim keptDraft As Object, keptNotes As Object, keptSubmit As Object, keptSources As Object
    Set keptDraft = CreateObject("Scripting.Dictionary")
    Set keptNotes = CreateObject("Scripting.Dictionary")
    Set keptSubmit = CreateObject("Scripting.Dictionary")
    Set keptSources = CreateObject("Scripting.Dictionary")

    ' Only carry work across if the sheet was written by THIS layout. A sheet
    ' from an older layout is read with column numbers that have since been
    ' renumbered, which silently relabels a person's data rather than losing
    ' it -- the worse of the two failures, because it looks like content.
    Dim sheetLayout As Long
    Dim layoutMatches As Boolean
    On Error Resume Next
    sheetLayout = CLng(Val(CStr(ws.Cells(DRAFT_INTRO_ROW, COL_D_LAYOUT).Value)))
    On Error GoTo 0
    layoutMatches = (sheetLayout = DRAFT_LAYOUT_VERSION)

    Dim isNewSheet As Boolean
    isNewSheet = (Trim(CStr(ws.Cells(DRAFT_FIRST_ROW, COL_D_ENTITY).Value)) = "")

    Dim r As Long
    r = DRAFT_FIRST_ROW
    If layoutMatches Then
        On Error Resume Next
        Do While Trim(CStr(ws.Cells(r, COL_D_ENTITY).Value)) <> ""
            Dim oldKey As String
            oldKey = Trim(CStr(ws.Cells(r, COL_D_ENTITY).Value))
            If Trim(CStr(ws.Cells(r, COL_D_DRAFT).Value)) <> "" Then keptDraft(oldKey) = CStr(ws.Cells(r, COL_D_DRAFT).Value)
            If Trim(CStr(ws.Cells(r, COL_D_NOTES).Value)) <> "" Then keptNotes(oldKey) = CStr(ws.Cells(r, COL_D_NOTES).Value)
            If Trim(CStr(ws.Cells(r, COL_D_SUBMIT).Value)) <> "" Then keptSubmit(oldKey) = CStr(ws.Cells(r, COL_D_SUBMIT).Value)
            If Trim(CStr(ws.Cells(r, COL_D_SOURCES).Value)) <> "" Then keptSources(oldKey) = CStr(ws.Cells(r, COL_D_SOURCES).Value)
            r = r + 1
        Loop
        On Error GoTo 0
    End If

    ws.Cells.Clear

    ' --- instructions, ON the sheet, for the person ----------------------
    ws.Cells(DRAFT_INTRO_ROW, 1).Value = "WHAT TO DO ON THIS SHEET  --  " & fieldId
    ws.Cells(DRAFT_INTRO_ROW, 1).Font.Bold = True
    ws.Cells(DRAFT_INTRO_ROW, 1).Font.Size = 9

    ws.Cells(2, 1).Value = "1.  Read column C (ORIGINAL) -- what the slide says today."
    ws.Cells(3, 1).Value = "2.  List the source IDs you are working from in column E. Add new ones on the Sources sheet first."
    ws.Cells(4, 1).Value = "3.  Ask Copilot for a draft (prompt is in L2). It writes into column F (AI DRAFT). F is never published."
    ws.Cells(5, 1).Value = "4.  Run Copy AI to Submit, then EDIT column G (SUBMIT) until you are happy. G is what gets sent."
    ws.Cells(6, 1).Value = "5.  Type  Y  in column I, save and CLOSE the file, then run Publish and Apply."

    ws.Cells(7, 1).Value = "Only column G is published -- nothing the AI writes reaches a slide unless you have put it in SUBMIT and ticked it. " & _
                           "Column C is read-only: edit the register, not this sheet, to change what a slide says today."
    ws.Cells(7, 1).Font.Italic = True

    Dim introRow As Long
    For introRow = 2 To 7
        ws.Cells(introRow, 1).Font.Size = 8
    Next introRow

    ' --- the grid --------------------------------------------------------
    ' Headers say what to DO, not what the column is called internally.
    ' "Draft" and "Approved (Y/N)" were accurate and told a first-time reader
    ' nothing about where to type.
    ws.Cells(DRAFT_HEADER_ROW, COL_D_ENTITY).Value = "Project code"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_NAME).Value = "Project name"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_CURRENT).Value = "C  --  ORIGINAL, what the slide says now (read-only)"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_CHARS).Value = "Chars"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_SOURCES).Value = "E  --  SOURCES (IDs from the Sources sheet, e.g. S01,S03)"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_DRAFT).Value = "F  --  AI DRAFT (Copilot writes here -- NEVER published)"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_SUBMIT).Value = "G  --  SUBMIT, your text (THIS is what gets sent)"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_SUBCHARS).Value = "Chars"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_APPROVED).Value = "I  --  TYPE Y TO USE IT"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_NOTES).Value = "J  --  your notes (optional)"
    ws.Rows(DRAFT_HEADER_ROW).Font.Bold = True
    ws.Rows(DRAFT_HEADER_ROW).WrapText = True

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
            ' THE DELIMITER IS STORAGE, NOT SOMETHING A PERSON SHOULD READ.
            ' The register stores line breaks as "||". Column C was writing that
            ' straight out, so a human comparing text saw "a||b" and -- worse --
            ' so did Copilot, which then learns to emit "||" in its drafts.
            ' Rendered as real breaks here. Safe to copy from: publish re-encodes
            ' vbLf back to "||", so a value copied out of C and into SUBMIT makes
            ' the round trip intact.
            ws.Cells(r, COL_D_CURRENT).Value = "'" & Replace(current, "||", vbLf)
            ws.Cells(r, COL_D_CHARS).Value = Len(current)
            If keptNotes.Exists(key) Then
                ws.Cells(r, COL_D_NOTES).Value = "'" & keptNotes(key)
                restored = restored + 1
            End If
            If keptDraft.Exists(key) Then ws.Cells(r, COL_D_DRAFT).Value = "'" & keptDraft(key)
            If keptSources.Exists(key) Then ws.Cells(r, COL_D_SOURCES).Value = "'" & keptSources(key)
            If keptSubmit.Exists(key) Then
                ws.Cells(r, COL_D_SUBMIT).Value = "'" & keptSubmit(key)
                ws.Cells(r, COL_D_SUBCHARS).Value = Len(CStr(keptSubmit(key)))
            End If
            ws.Cells(r, COL_D_APPROVED).Value = ""
            written = written + 1
            r = r + 1
        End If
    Next k

    ' 8pt throughout. At 11pt three text columns of 350+ characters do not fit
    ' on a screen together, and the whole point of ORIGINAL / AI / SUBMIT is
    ' reading them side by side.
    ws.Cells.Font.Size = 8
    ws.Cells(DRAFT_INTRO_ROW, 1).Font.Size = 9

    ws.Columns(COL_D_ENTITY).ColumnWidth = 11
    ws.Columns(COL_D_NAME).ColumnWidth = 30
    ws.Columns(COL_D_CURRENT).ColumnWidth = 52
    ws.Columns(COL_D_CHARS).ColumnWidth = 6
    ws.Columns(COL_D_SOURCES).ColumnWidth = 14
    ws.Columns(COL_D_DRAFT).ColumnWidth = 52
    ws.Columns(COL_D_SUBMIT).ColumnWidth = 52
    ws.Columns(COL_D_SUBCHARS).ColumnWidth = 6
    ws.Columns(COL_D_APPROVED).ColumnWidth = 9
    ws.Columns(COL_D_NOTES).ColumnWidth = 24
    ws.Columns(COL_D_CURRENT).WrapText = True
    ws.Columns(COL_D_DRAFT).WrapText = True
    ws.Columns(COL_D_SUBMIT).WrapText = True
    ws.Columns(COL_D_NOTES).WrapText = True

    ' SUBMIT is the only column that reaches a slide, so it is the only one
    ' that looks like an input. ORIGINAL and AI DRAFT are shaded as reference.
    ws.Columns(COL_D_CURRENT).Interior.Color = RGB(242, 242, 242)
    ws.Columns(COL_D_DRAFT).Interior.Color = RGB(242, 242, 242)
    ws.Range(ws.Cells(DRAFT_FIRST_ROW, COL_D_SUBMIT), _
             ws.Cells(r - 1, COL_D_SUBMIT)).Interior.Color = RGB(255, 249, 219)
    ws.Range(ws.Cells(DRAFT_FIRST_ROW, COL_D_APPROVED), _
             ws.Cells(r - 1, COL_D_APPROVED)).Interior.Color = RGB(255, 249, 219)

    ' Top-align, or a 500-character ORIGINAL centres itself against a one-line
    ' SUBMIT and the two stop reading as the same row.
    ws.Cells.VerticalAlignment = -4160          ' xlTop

    ' The header row and the identifying columns stay put while you scroll --
    ' at 43 rows and 10 columns you otherwise lose track of which project a
    ' 350-character paragraph belongs to.
    '
    ' ws.Application, NOT a bare ActiveWindow. This code runs inside
    ' PowerPoint's VBA host and drives Excel over COM, so an unqualified
    ' ActiveWindow is POWERPOINT'S window -- it would either raise or silently
    ' freeze panes on a slide. Same class of mistake as reading Worksheets(1).
    '
    ' Wrapped, because freezing panes is cosmetic and must never be the reason
    ' a drafting sheet fails to build.
    On Error Resume Next
    Dim xlApp As Object
    Set xlApp = ws.Application
    ws.Activate
    xlApp.ActiveWindow.FreezePanes = False
    ws.Cells(DRAFT_FIRST_ROW, COL_D_CURRENT).Select
    xlApp.ActiveWindow.FreezePanes = True
    On Error GoTo 0

    ' The AI prompt, on the sheet rather than in a console window that closes.
    ' Placed to the RIGHT of the grid so it is available to copy without
    ' pushing the instructions a person needs off the top of the screen.
    ws.Cells(DRAFT_INTRO_ROW, COL_D_PROMPT).Font.Bold = True
    ' The prompt comes from the field's OWN spec row when there is one. A field
    ' with no row still drafts -- on generic guidance that says so -- because
    ' blocking the work until somebody writes a style guide would be paperwork
    ' standing in front of delivery.
    Dim g As FieldGuidance
    ' guidance arrives As Variant (it is Optional); LookupGuidance takes As Object
    ' ByRef, and VBA will not coerce Variant -> Object across a ByRef boundary --
    ' it is a COMPILE error, not a runtime one, so the unit tests could never see
    ' it: they call LookupGuidance directly with an already-typed Object.
    ' Hand it a typed local instead.
    Dim guidanceWs As Object
    If IsObject(guidance) Then
        Set guidanceWs = guidance
        g = FieldSpec.LookupGuidance(guidanceWs, fieldId)
    Else
        g.FieldId = fieldId
    End If
    ws.Cells(2, COL_D_PROMPT).Value = "'" & FieldSpec.PromptFrom(g)
    ws.Cells(DRAFT_INTRO_ROW, COL_D_PROMPT).Value = _
        "PROMPT TO GIVE COPILOT (copy this cell)" & IIf(g.Found, "", "  --  GENERIC, no Field Spec row")
    ws.Columns(COL_D_PROMPT).ColumnWidth = 70
    ws.Cells(2, COL_D_PROMPT).WrapText = True

    ws.Rows(DRAFT_HEADER_ROW).RowHeight = 30

    ' ROW HEIGHTS ARE SET, NEVER LEFT TO AUTOFIT.
    '
    ' Three wrapped columns of 350-500 characters at width 52 autofit to well
    ' over 100 points each, so one project fills the screen and the sheet reads
    ' as a stack of essays rather than a grid. Rohan, 2026-08-01: "something
    ' weird happening with row height".
    '
    ' Fixed instead: enough to see the shape of a paragraph and compare three
    ' of them, not enough to read one in full. Reading one in full is what
    ' clicking the cell is for -- the formula bar holds the whole thing, and a
    ' row can always be dragged taller for a moment.
    ws.Rows(DRAFT_INTRO_ROW).RowHeight = 14
    Dim ir As Long
    For ir = 2 To 7
        ws.Rows(ir).RowHeight = 11
    Next ir
    ws.Rows(8).RowHeight = 6
    ws.Rows(DRAFT_HEADER_ROW).RowHeight = 30
    If r > DRAFT_FIRST_ROW Then
        ws.Range(ws.Rows(DRAFT_FIRST_ROW), ws.Rows(r - 1)).RowHeight = 52
    End If

    ws.Cells(DRAFT_INTRO_ROW, COL_D_LAYOUT).Value = DRAFT_LAYOUT_VERSION
    ws.Cells(DRAFT_INTRO_ROW, COL_D_LAYOUT).Font.Color = RGB(190, 190, 190)
    ws.Columns(COL_D_LAYOUT).ColumnWidth = 4

    WriteDraftingSheet = written & " row(s) written for " & fieldId & _
        IIf(restored > 0, ", " & restored & " existing note(s)/draft(s) preserved", "") & "."

    ' Said out loud, because the alternative is a person discovering by absence
    ' that a rebuild dropped their drafting. Not raised: the rebuild itself is
    ' correct and the register is untouched, so this is news, not a failure.
    If Not layoutMatches And Not isNewSheet Then
        WriteDraftingSheet = WriteDraftingSheet & vbCrLf & _
            "NOTE: this sheet was built by layout " & sheetLayout & _
            " and the tool is now on layout " & DRAFT_LAYOUT_VERSION & ". Nothing was" & vbCrLf & _
            "carried across -- the columns have been renumbered, so the old contents" & vbCrLf & _
            "would have been read into the wrong meanings. Anything you had drafted is" & vbCrLf & _
            "in the .bak beside the workbook, not lost."
    End If
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
                              Optional dryRun As Boolean = True, _
                              Optional sourcesWs As Variant) As String
    Dim report As String
    report = IIf(dryRun, "=== PREVIEW: publish " & fieldId & " ===", "=== Publish " & fieldId & " ===") & vbCrLf

    ' TYPED LOCAL, not the Variant straight through. Passing an Optional
    ' Variant into a ByRef `As Object` parameter is a COMPILE error in VBA and
    ' cost a whole session on 2026-08-01 -- the unit tests could not see it
    ' because it only exists where two modules meet.
    Dim knownSources As Object
    If IsObject(sourcesWs) Then
        Dim srcSheet As Object
        Set srcSheet = sourcesWs
        Set knownSources = Sources.KnownSourceIds(srcSheet)
    End If
    Dim badRefs As String

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

        ' SUBMIT, NEVER THE AI COLUMN. This is the entire point of splitting
        ' them: text Copilot produced has no route to a slide until a person
        ' has moved it into SUBMIT and ticked it. Reading COL_D_DRAFT here
        ' would silently collapse the two back into one and undo the control
        ' without changing anything visible on the sheet.
        Dim draft As String
        draft = CStr(ws.Cells(r, COL_D_SUBMIT).Value)
        Dim tick As Boolean
        tick = ReviewQueue.IsApprovalMark(CStr(ws.Cells(r, COL_D_APPROVED).Value))

        ' Source IDs are checked on every row that is being sent, and REPORTED
        ' rather than enforced. A wrong ID means the provenance record is wrong,
        ' which is worth saying out loud -- but refusing to publish a quarter's
        ' text over a typo in a reference column would be the tool getting in
        ' the way of the work it exists to do.
        If tick And Not knownSources Is Nothing Then
            Dim rowBad As String
            rowBad = Sources.UnknownRefs(CStr(ws.Cells(r, COL_D_SOURCES).Value), knownSources)
            If rowBad <> "" Then
                badRefs = badRefs & "  " & ent & " refers to " & rowBad & _
                    " -- not on the Sources sheet" & vbCrLf
            End If
        End If

        If Trim(draft) = "" Then
            If tick Then
                skippedEmpty = skippedEmpty + 1
                If Trim(CStr(ws.Cells(r, COL_D_DRAFT).Value)) <> "" Then
                    report = report & "  SKIPPED " & ent & " -- ticked, and there IS an AI draft, " & _
                        "but SUBMIT is empty. Run Copy AI to Submit, or type it yourself." & vbCrLf
                Else
                    report = report & "  SKIPPED " & ent & " -- ticked but SUBMIT is empty" & vbCrLf
                End If
            End If
        ElseIf Not tick Then
            skippedNoTick = skippedNoTick + 1
        Else
            Dim encoded As String
            encoded = Replace(draft, vbCrLf, "||")
            encoded = Replace(encoded, vbCr, "||")
            encoded = Replace(encoded, vbLf, "||")
            encoded = Replace(encoded, Chr(11), "||")

            ' THIS GUARD USED TO BE UNREACHABLE. It tested for vbCr and vbLf
            ' four lines after Replace had removed both, so the branch could
            ' never run -- an always-false guard reading as care taken, the same
            ' shape as IsToolOwnedSheet's 13-versus-14 and the two guards the
            ' project's own zettel was written about.
            '
            ' Now checks what could ACTUALLY survive: any remaining control
            ' character, including the Unicode line and paragraph separators
            ' (U+2028/U+2029) that Office produces and that Replace above does
            ' not touch. A register value containing one breaks the round-trip
            ' silently.
            If HasControlCharacter(encoded) Then
                failed = failed + 1
                report = report & "  FAILED " & ent & " -- a line break or control character survived conversion" & vbCrLf
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

    If badRefs <> "" Then
        report = report & vbCrLf & "SOURCE REFERENCES THAT DO NOT RESOLVE:" & vbCrLf & badRefs & _
            "The text still published. The provenance record for those rows is wrong." & vbCrLf
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

' Copy AI DRAFT into SUBMIT, for rows where SUBMIT is still empty.
'
' FILLS ONLY BLANKS, DELIBERATELY. The obvious implementation overwrites
' SUBMIT wholesale, and it would destroy an afternoon's editing the second time
' somebody ran it -- silently, because a re-run that looks identical to the
' first is exactly what nobody re-checks. Anything already in SUBMIT is a
' decision a person made and this function is not entitled to it.
'
' Reports the skipped count rather than hiding it: "23 copied" and "23 copied,
' 8 left alone because you had already written something" are different
' situations and a person needs to know which one happened.
Public Function CopyAiToSubmit(ws As Object) As String
    Dim copied As Long, keptExisting As Long, noAi As Long

    Dim r As Long
    r = DRAFT_FIRST_ROW
    Do While Trim(CStr(ws.Cells(r, COL_D_ENTITY).Value)) <> ""
        Dim ai As String, submitted As String
        ai = CStr(ws.Cells(r, COL_D_DRAFT).Value)
        submitted = CStr(ws.Cells(r, COL_D_SUBMIT).Value)

        If Trim(ai) = "" Then
            noAi = noAi + 1
        ElseIf Trim(submitted) <> "" Then
            keptExisting = keptExisting + 1
        Else
            ws.Cells(r, COL_D_SUBMIT).Value = "'" & ai
            ws.Cells(r, COL_D_SUBCHARS).Value = Len(ai)
            copied = copied + 1
        End If
        r = r + 1
    Loop

    CopyAiToSubmit = "Copy AI -> Submit: " & copied & " copied, " & _
        keptExisting & " left alone (you had already written something there), " & _
        noAi & " with no AI draft." & vbCrLf & _
        "Nothing was overwritten. Edit column G, tick column I, then publish." & vbCrLf
End Function

' Refresh the SUBMIT character counts without touching any text. Cheap, and it
' means the length column is never stale after hand-editing.
Public Function RefreshSubmitCounts(ws As Object) As String
    Dim n As Long
    Dim r As Long
    r = DRAFT_FIRST_ROW
    Do While Trim(CStr(ws.Cells(r, COL_D_ENTITY).Value)) <> ""
        Dim s As String
        s = CStr(ws.Cells(r, COL_D_SUBMIT).Value)
        If Trim(s) = "" Then
            ws.Cells(r, COL_D_SUBCHARS).Value = ""
        Else
            ws.Cells(r, COL_D_SUBCHARS).Value = Len(s)
            n = n + 1
        End If
        r = r + 1
    Loop
    RefreshSubmitCounts = n & " row(s) have SUBMIT text."
End Function

' Any control character that would break a register round-trip.
'
' Deliberately includes U+2028 LINE SEPARATOR and U+2029 PARAGRAPH SEPARATOR:
' Office emits both, and the Replace chain that strips vbCrLf/vbCr/vbLf/Chr(11)
' does not touch them. Tab is allowed -- it is legal in a cell and harmless in
' the register.
Private Function HasControlCharacter(value As String) As Boolean
    Dim i As Long
    For i = 1 To Len(value)
        Dim c As Long
        c = AscW(Mid(value, i, 1))
        If c < 32 And c <> 9 Then
            HasControlCharacter = True
            Exit Function
        End If
        If c = 8232 Or c = 8233 Then
            HasControlCharacter = True
            Exit Function
        End If
    Next i
End Function
