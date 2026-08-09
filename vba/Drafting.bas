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
' READ, WRITE, TICK -- ADJACENT. Reordered 2026-08-08.
'
' ORIGINAL was column C and SUBMIT column G, with Chars, Sources and AI Draft
' between them. Reading C then typing in G is the action repeated once per
' project -- 43 times for one field -- and it meant crossing four columns each
' time, or scrolling sideways on a narrower screen. Everything else on the sheet
' is consulted occasionally; these three are the work.
'
' PROMPT stays at column 12 so Copilot's prompt is still in cell L2 -- that
' address is written in the on-sheet instructions and in the toolbar tooltip.
' LEFT TO RIGHT IS THE ORDER YOU DO IT IN. Rohan, 2026-08-10: "fix the column
' order in the drafting sheet to make more sense re workflow".
'
' Layout 3 had the columns in the order they were ADDED, not the order they are
' used: you read C, jumped right to G for sources, back to F for the AI draft,
' left again to D to edit, then E to tick. The sheet's own instructions listed
' the steps 1-5 and the columns went 3, 7, 6, 4, 5.
'
' Layout 4 puts them in step order, so working the sheet is a walk rightwards:
'   C read what the slide says now
'   D name the sources you are drafting from
'   E the AI writes here (never published)
'   F your words (this is what publishes)
'   G tick to approve
'   H notes back to the tool
Public Const COL_D_ENTITY As Long = 1
Public Const COL_D_NAME As Long = 2
Public Const COL_D_CURRENT As Long = 3      ' C  ORIGINAL -- read this first
Public Const COL_D_SOURCES As Long = 4      ' D  what you drafted FROM
Public Const COL_D_DRAFT As Long = 5        ' E  AI draft, never published
Public Const COL_D_SUBMIT As Long = 6       ' F  your text -- this publishes
Public Const COL_D_APPROVED As Long = 7     ' G  the tick
Public Const COL_D_CHARS As Long = 8
Public Const COL_D_SUBCHARS As Long = 9
Public Const COL_D_NOTES As Long = 10      ' J  notes back to the tool
Public Const COL_D_LAYOUT As Long = 11
Public Const COL_D_PERIOD As Long = 13
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
Public Const DRAFT_LAYOUT_VERSION As Long = 4

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
' WHERE A COLUMN USED TO LIVE.
'
' Bumping DRAFT_LAYOUT_VERSION protects a person from a renumbered sheet being
' read with the wrong column numbers -- "column 7 meaning tick on Monday and
' SUBMIT on Tuesday". It protects by REFUSING to carry anything, which is safe
' and lossy: every unpublished draft, note and source citation on the sheet is
' dropped on the next rebuild.
'
' That is too expensive to pay for a reordering. So a KNOWN older layout is read
' through its own column map instead of being abandoned. An UNKNOWN one still
' returns 0 and is still refused -- a layout nobody has described cannot be read
' safely, and guessing is the failure this whole mechanism exists to prevent.
Private Function ColumnInLayout(layoutVersion As Long, which As String) As Long
    Select Case layoutVersion
        Case DRAFT_LAYOUT_VERSION
            Select Case which
                Case "CURRENT":  ColumnInLayout = COL_D_CURRENT
                Case "SOURCES":  ColumnInLayout = COL_D_SOURCES
                Case "DRAFT":    ColumnInLayout = COL_D_DRAFT
                Case "SUBMIT":   ColumnInLayout = COL_D_SUBMIT
                Case "APPROVED": ColumnInLayout = COL_D_APPROVED
                Case "NOTES":    ColumnInLayout = COL_D_NOTES
            End Select

        ' Layout 3: columns sat in the order they were added, not used.
        ' C ORIGINAL, D SUBMIT, E tick, F AI draft, G sources, J notes.
        ' Layout 3: columns sat in the order they were added, not used.
        ' C ORIGINAL, D SUBMIT, E tick, F AI draft, G sources, J notes.
        Case 3
            Select Case which
                Case "CURRENT":  ColumnInLayout = 3
                Case "SOURCES":  ColumnInLayout = 7
                Case "DRAFT":    ColumnInLayout = 6
                Case "SUBMIT":   ColumnInLayout = 4
                Case "APPROVED": ColumnInLayout = 5
                Case "NOTES":    ColumnInLayout = 10
            End Select

        Case Else
            ColumnInLayout = 0
    End Select
End Function

Public Function WriteDraftingSheet(ws As Object, reg As Sheet, fieldId As String, _
                                   Optional guidance As Variant, _
                                   Optional periodStamp As String = "", _
                                   Optional cadence As Object = Nothing, _
                                   Optional srcWs As Object = Nothing) As String
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
    ' Readable, not merely identical -- a known older layout is migrated below.
    layoutMatches = (ColumnInLayout(sheetLayout, "SUBMIT") > 0)

    ' AND THE PERIOD MUST MATCH TOO.
    '
    ' The sheet is per FIELD, not per period, and a rebuild carries SUBMIT and
    ' SOURCES forward by EntityCode. So rolling the deck to a new quarter left
    ' last quarter's text sitting in column G with last quarter's source IDs --
    ' approvals are cleared, but one tick republishes stale prose as though it
    ' were this quarter's, and CopyAiToSubmit would decline to overwrite it
    ' ("left alone -- you had already written something there").
    '
    ' Silent, plausible, and in a funder-facing deck. Found by review
    ' 2026-08-01, the same evening the layout guard was added for exactly this
    ' shape of failure one axis over.
    '
    ' THE TWO EMPTY CASES ARE NOT SYMMETRIC, and an earlier version of this
    ' comment claimed they were:
    '   - Sheet has no stamp, caller names a period -> treated as a MISMATCH.
    '     A sheet that cannot say which quarter it belongs to is not evidence
    '     that it belongs to this one, and dropping work is the safe direction.
    '   - Caller names no period -> carry across as before. The caller did not
    '     say, so there is nothing to compare; guessing "mismatch" here would
    '     make every driver that omits the argument destroy drafting silently.
    ' RefreshDraftingSheets refuses outright on an undeclared deck period
    ' (DraftingUI.bas:223), so the UI path can never reach the second case.
    '
    ' THE DROP IS PER ROW, NOT PER SHEET, and that distinction is the whole of
    ' this guard's correctness.
    '
    ' Variability is a property of the ROW. The register carries `Quarter = ALL`
    ' rows into every period as entity-static (Register.bas:232), and for such a
    ' row LAST QUARTER'S TEXT IS THIS QUARTER'S TEXT -- dropping it is pure loss
    ' with no safety bought, because there is no stale value to republish.
    '
    ' The first version of this guard dropped everything on the sheet, defended
    ' by "a drafting sheet only holds Prose fields, so it is all quarterly". That
    ' conflates two different axes and is wrong:
    '     FieldSpec.Kind (Controlled/Prose/Static)  = HOW a value is produced
    '     Register Quarter (a period, or ALL)       = WHEN it applies
    ' A project description is prose AND static -- Round 5 §3 classes ABOUT_BODY,
    ' the flagship prose field, as entity-static, and Rohan confirmed 2026-08-02
    ' that he writes it once and edits it rarely. So the destroyed-work case was
    ' not hypothetical; it was the main field, every rollover.
    '
    ' `cadence` is the register's own answer, keyed EntityCode & Chr(1) & FieldID:
    ' True = this value came from a period row (drop the drafting on rollover),
    ' False = it came from an ALL row (keep it). Absent or unknown drops, because
    ' the old blunt behaviour is the safe direction when nobody can say.
    Dim sheetPeriod As String
    On Error Resume Next
    sheetPeriod = Trim(CStr(ws.Cells(DRAFT_INTRO_ROW, COL_D_PERIOD).Value))
    On Error GoTo 0

    Dim periodMatches As Boolean
    If periodStamp = "" Then
        periodMatches = True            ' caller did not say; do not guess
    Else
        periodMatches = (StrComp(sheetPeriod, periodStamp, vbTextCompare) = 0)
    End If

    ' THE LAYOUT GUARD STILL BLOCKS EVERYTHING, and must. An old layout is read
    ' with column numbers that have since been renumbered, so there is no such
    ' thing as a row that can be safely rescued from it -- unlike a rollover,
    ' where the columns are fine and only the CONTENT'S shelf life is in
    ' question. Two different failures; only one of them is per row.

    Dim isNewSheet As Boolean
    isNewSheet = (Trim(CStr(ws.Cells(DRAFT_FIRST_ROW, COL_D_ENTITY).Value)) = "")

    ' A SHEET THAT CANNOT SAY WHICH QUARTER IT IS FROM IS TREATED AS A ROLLOVER.
    ' Every drafting sheet built before this stamp existed is in that state. It
    ' cannot be shown to be current, so its quarterly rows go and its ALL rows
    ' stay -- which is strictly better than the two alternatives of assuming it
    ' is current (republishes stale prose) or dropping the lot (throws away
    ' entity-static work that was never at risk).
    Dim periodChanged As Boolean
    periodChanged = (Not periodMatches) And (Not isNewSheet)

    Dim droppedQuarterly As Long, keptStatic As Long

    Dim r As Long
    r = DRAFT_FIRST_ROW
    If layoutMatches Then
        On Error Resume Next
        Do While Trim(CStr(ws.Cells(r, COL_D_ENTITY).Value)) <> ""
            Dim oldKey As String
            oldKey = Trim(CStr(ws.Cells(r, COL_D_ENTITY).Value))

            ' On a rollover, this row survives only if the register says its
            ' value came from a Quarter = ALL row.
            Dim carryThisRow As Boolean
            carryThisRow = True
            If periodChanged Then
                carryThisRow = False
                If Not cadence Is Nothing Then
                    Dim cadKey As String
                    cadKey = oldKey & Chr(1) & fieldId
                    ' CBool of the stored flag: True means period-specific.
                    ' Exists() first -- reading a missing key from a Dictionary
                    ' ADDS it, and silently growing the register's own cadence
                    ' map from inside the drafting sheet would be a fine way to
                    ' make a later sync disagree with itself.
                    If cadence.Exists(cadKey) Then carryThisRow = Not CBool(cadence(cadKey))
                End If
                If carryThisRow Then keptStatic = keptStatic + 1 Else droppedQuarterly = droppedQuarterly + 1
            End If

            If carryThisRow Then
                If Trim(CStr(ws.Cells(r, ColumnInLayout(sheetLayout, "DRAFT")).Value)) <> "" Then keptDraft(oldKey) = CStr(ws.Cells(r, ColumnInLayout(sheetLayout, "DRAFT")).Value)
                If Trim(CStr(ws.Cells(r, ColumnInLayout(sheetLayout, "NOTES")).Value)) <> "" Then keptNotes(oldKey) = CStr(ws.Cells(r, ColumnInLayout(sheetLayout, "NOTES")).Value)
                If Trim(CStr(ws.Cells(r, ColumnInLayout(sheetLayout, "SUBMIT")).Value)) <> "" Then keptSubmit(oldKey) = CStr(ws.Cells(r, ColumnInLayout(sheetLayout, "SUBMIT")).Value)
                If Trim(CStr(ws.Cells(r, ColumnInLayout(sheetLayout, "SOURCES")).Value)) <> "" Then keptSources(oldKey) = CStr(ws.Cells(r, ColumnInLayout(sheetLayout, "SOURCES")).Value)
            End If
            r = r + 1
        Loop
        On Error GoTo 0
    End If

    ws.Cells.Clear

    ' --- instructions, ON the sheet, for the person ----------------------
    ws.Cells(DRAFT_INTRO_ROW, 1).Value = "WHAT TO DO ON THIS SHEET  --  " & fieldId
    ws.Cells(DRAFT_INTRO_ROW, 1).Font.Bold = True
    ws.Cells(DRAFT_INTRO_ROW, 1).Font.Size = 9

    ws.Cells(2, 1).Value = "STEP 1   Read column C -- what the slide says today."
    ' COLUMN G, NOT E. Step 5 below sends the tick to E, so this line named one
    ' column for two things inside a single instruction block -- and E is the tick,
    ' which is the consent gate. Stale since 3de4be8 moved SUBMIT to D and the tick
    ' to E; that commit updated the header row and the toolbar tooltip and left
    ' every prose instruction pointing at the old layout.
    ws.Cells(3, 1).Value = "STEP 2   Name your sources in column D. Add new ones on the Sources sheet first."
    ws.Cells(4, 1).Value = "STEP 3   Ask Copilot for a draft -- the prompt is in cell L2. It writes into column E. E is NEVER published."
    ws.Cells(5, 1).Value = "STEP 4   Press '" & CommandBarUI.CAP_SYNC_NOW & "' to copy E into F, then EDIT column F until you are happy. F is what gets sent."
    ws.Cells(6, 1).Value = "STEP 5   Type  Y  in column G, save and CLOSE the file, then press '" & CommandBarUI.CAP_SYNC_NOW & "' again."

    ws.Cells(7, 1).Value = "Only column F is published -- nothing the AI writes reaches a slide unless you have moved it into F and ticked it. " & _
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
    ws.Cells(DRAFT_HEADER_ROW, COL_D_CURRENT).Value = "C   ORIGINAL -- what the slide says now (read-only)"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_CHARS).Value = "Chars"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_SOURCES).Value = "D   SOURCES -- IDs from the Sources sheet, e.g. S01,S03"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_DRAFT).Value = "E   AI DRAFT -- Copilot writes here. NEVER published"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_SUBMIT).Value = "F   SUBMIT -- your words. THIS is what gets sent"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_SUBCHARS).Value = "Chars"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_APPROVED).Value = "G   APPROVE -- type Y"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_NOTES).Value = "J   NOTES -- back to the tool (optional)"
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
            ' CLEARED EVERY ITERATION, and that line is the whole defect.
            '
            ' VBA's Dim does not scope to a loop -- these are procedure-scoped and
            ' keep the PREVIOUS entity's value on any iteration where the If below
            ' does not fire. So the first project with a value for this field had
            ' its text copied into column C for every project after it that had
            ' none.
            '
            ' Found 2026-08-09 on the first real run of three new fields: the
            ' register held STRATEGIC_ALIGNMENT_BODY for exactly one project, and
            ' the drafting sheet showed that project's 1,113 characters against
            ' FORTY of them. The three rows ABOVE it were blank, which is the
            ' signature -- nothing had been assigned yet.
            '
            ' Column C is labelled "what the slide says now" and is what a person
            ' and Copilot are both told to stay close to. Forty projects would
            ' have been drafted against one project's story, and nothing
            ' downstream could have noticed: the values are real, just attributed
            ' to the wrong entity.
            '
            ' Rohan asked "is that the old q?" about this exact hazard in
            ' RibbonUI earlier the same day. It was safe there. It was not here.
            Dim current As String
            Dim projName As String
            current = ""
            projName = ""
            If vals.Exists(fieldId) Then current = CStr(vals(fieldId))
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
    ' THE CITED SOURCES GO INTO THE PROMPT. Built AFTER the rows are written,
    ' because it reads column G off the sheet it has just laid out -- including
    ' the citations carried across from the previous build.
    Dim citedBlock As String
    If Not srcWs Is Nothing Then
        citedBlock = Sources.CitedBlockFor(srcWs, ws, COL_D_SOURCES, DRAFT_FIRST_ROW)
    End If
    ws.Cells(2, COL_D_PROMPT).Value = "'" & FieldSpec.PromptFrom(g, citedBlock)
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

    ws.Cells(DRAFT_INTRO_ROW, COL_D_PERIOD).Value = periodStamp
    ws.Cells(DRAFT_INTRO_ROW, COL_D_PERIOD).Font.Color = RGB(190, 190, 190)
    ws.Columns(COL_D_PERIOD).ColumnWidth = 9
    ws.Cells(DRAFT_INTRO_ROW, COL_D_LAYOUT).Value = DRAFT_LAYOUT_VERSION
    ws.Cells(DRAFT_INTRO_ROW, COL_D_LAYOUT).Font.Color = RGB(190, 190, 190)
    ws.Columns(COL_D_LAYOUT).ColumnWidth = 4

    WriteDraftingSheet = written & " row(s) written for " & fieldId & _
        IIf(restored > 0, ", " & restored & " existing note(s)/draft(s) preserved", "") & "."

    ' Said out loud, because the alternative is a person discovering by absence
    ' that a rebuild dropped their drafting. Not raised: the rebuild itself is
    ' correct and the register is untouched, so this is news, not a failure.
    ' ONE LINE, NOT A PARAGRAPH. MsgBox caps its prompt near 1024 characters and
    ' truncates SILENTLY past it -- so every line of rationale here costs a line
    ' of a real warning somewhere else in the same dialog, and the warning is the
    ' part that gets dropped. Rohan on the previous version, 2026-08-08:
    ' "illegible, too long, and the user has no idea what is going on."
    '
    ' The old text also explained itself with "their register row is Quarter =
    ' ALL" -- the sentinel retired on 2026-08-03.
    If periodChanged Then
        ' NAMES BOTH PERIODS. Shortening this to the new period alone was a real
        ' regression, caught by the test that exists for it: a note saying work
        ' was cleared, without saying which period it was cleared FROM, does not
        ' explain the thing the reader is looking at.
        ' NO WORKBOOK BACKUP EXISTS, AND THIS USED TO SAY THERE WAS ONE.
        '
        ' Until 2026-08-08 both messages in this function ended "Previous sheet
        ' saved in the .bak beside the workbook." No code in this add-in has ever
        ' written a workbook .bak. The only .bak taken anywhere is
        ' ReviewQueue.BackupBeforeWrite, which does pres.SaveCopyAs -- the DECK,
        ' at Apply time, named .r13-<stamp>.bak.pptx. A deck backup sitting in the
        ' same folder is exactly what made the claim survive a glance.
        '
        ' ReviewQueue's own comment states the rule this broke: "A REPORTED BACKUP
        ' THAT IS NOT ON DISK IS WORSE THAN NO BACKUP: it is the reason you feel
        ' safe running the destructive write that follows."
        '
        ' Says what survives before what does not, because the register really does
        ' still hold anything published and that is the actionable half.
        WriteDraftingSheet = WriteDraftingSheet & vbCrLf & _
            "  Rebuilt " & IIf(sheetPeriod = "", "(no period recorded)", sheetPeriod) & _
            " -> " & periodStamp & ": " & droppedQuarterly & " row(s) cleared for redrafting" & _
            IIf(keptStatic > 0, ", " & keptStatic & " carried over", "") & _
            ". Anything already published is safe in the register. Drafts, source IDs " & _
            "and notes that were NOT published are gone -- this workbook is not backed up."
    ElseIf Not layoutMatches And Not isNewSheet Then
        ' Same false claim as above, same fix -- see the comment there. Fixed in
        ' both places at once because this is one defect with two call sites, and
        ' fixing only where it was noticed is how the truncation bug came back
        ' four times.
        WriteDraftingSheet = WriteDraftingSheet & vbCrLf & _
            "  Rebuilt on a new sheet layout: nothing carried across. Anything already " & _
            "published is safe in the register. Drafts, source IDs and notes that were " & _
            "NOT published are gone -- this workbook is not backed up."
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
' PUBLISHES INTO THE WIDE SHEET, which is why `period` is required and not
' Optional.
'
' Until 2026-08-05 this wrote the LONG register -- it located EntityCode/FieldID/
' Value/Status columns and set Value + Status = Approved on the matching row.
' Sync Now had already moved to the wide sheet, so publish and sync were reading
' and writing two different files and the quarterly loop did not close: text
' published here could never reach a slide.
'
' Three things go with the move, none of them optional:
'
'   - STATUS HAS NOWHERE TO GO. The wide sheet has no Status column, and adding
'     one would make it a FIELD -- every field column on this sheet is something
'     that lands on a slide, so a "Status" column would be injected as slide
'     text. The tick in the drafting sheet's APPROVED column is now the only
'     consent gate, which is what it always actually was; Status was a second
'     copy of the same decision, recorded after the fact.
'   - CharCount and UpdatedDate go for the same reason -- they were long-register
'     bookkeeping columns and would become slide fields here. The drafting sheet
'     already shows character counts.
'   - A MISSING ROW IS STILL REFUSED. UpsertRow would happily CREATE a row for a
'     slide that has no row in this period, which would mean publishing invents
'     slides that were never onboarded. The read below is what makes the refusal
'     possible, so it happens once, up front, rather than per row.
Public Function PublishDrafts(ws As Object, regWs As Object, fieldId As String, _
                              period As String, _
                              Optional dryRun As Boolean = True, _
                              Optional sourcesWs As Variant) As String
    Dim report As String
    report = IIf(dryRun, "=== PREVIEW: publish " & fieldId & " ===", "=== Publish " & fieldId & " ===") & vbCrLf

    ' TYPED LOCAL, not the Variant straight through. Passing an Optional
    ' Variant into a ByRef `As Object` parameter is a COMPILE error in VBA and
    ' cost a whole session on 2026-08-01 -- the unit tests could not see it
    ' because it only exists where two modules meet.
    Dim knownSources As Object
    Dim sourceApplies As Object
    If IsObject(sourcesWs) Then
        Dim srcSheet As Object
        Set srcSheet = sourcesWs
        Set knownSources = Sources.KnownSourceIds(srcSheet)
        Set sourceApplies = Sources.SourceApplicability(srcSheet)
    End If
    Dim badRefs As String
    Dim wrongPeriodRefs As String

    ' Refused rather than defaulted, exactly as UpsertRow refuses a blank period:
    ' a publish with no period would have to guess which quarter it is writing,
    ' and the wrong guess overwrites a real quarter's text with no trace.
    If Trim$(period) = "" Then
        PublishDrafts = report & "STOPPED: no period given. Publish writes into one " & _
            "quarter's rows and cannot pick which."
        Exit Function
    End If

    ' ONE READ, THROUGH THE GUARDED READER THE SYNC USES. Not a convenience:
    ' reading the sheet the same way Sync Now reads it is what makes "published"
    ' and "will reach a slide" the same claim. It also inherits both refusals --
    ' a repeated slide, and a sheet whose rows are all stamped some other period
    ' (which would otherwise let every row report NO REGISTER ROW and read as a
    ' drafting-sheet problem rather than a wrong-period one).
    Dim problem As String
    Dim reg As Sheet
    reg = ExcelOutput.ReadSheetForDeckPeriod(regWs, period, problem)
    If problem <> "" Then
        PublishDrafts = report & "STOPPED: " & problem
        Exit Function
    End If

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

        ' A source that names a period, cited from a different period's text.
        ' Publishing Q1F27 while citing last quarter's progress report is a
        ' provenance error the existence check cannot see: the ID is perfectly
        ' real, it just documents the wrong quarter.
        If tick And Not sourceApplies Is Nothing Then
            Dim rowWrongPeriod As String
            rowWrongPeriod = Sources.RefsForOtherPeriod( _
                CStr(ws.Cells(r, COL_D_SOURCES).Value), sourceApplies, period)
            If rowWrongPeriod <> "" Then
                wrongPeriodRefs = wrongPeriodRefs & "  " & ent & " cites " & rowWrongPeriod & _
                    " while publishing " & period & vbCrLf
            End If
        End If

        If Trim(draft) = "" Then
            If tick Then
                skippedEmpty = skippedEmpty + 1
                If Trim(CStr(ws.Cells(r, COL_D_DRAFT).Value)) <> "" Then
                    report = report & "  SKIPPED " & ent & " -- ticked, and there IS an AI draft, " & _
                        "but SUBMIT is empty. Press '" & CommandBarUI.CAP_SYNC_NOW & "' to copy it across, or type into D yourself." & vbCrLf
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
                ' Existence is asked of the READ, not of the sheet directly, so
                ' publish and sync agree by construction about which rows are
                ' this period's. A row for this slide in ANOTHER period is not a
                ' row here, and must not be written to.
                If Not reg.Rows.Exists(ent) Then
                    noRow = noRow + 1
                    report = report & "  NO REGISTER ROW for " & ent & " in " & period & vbCrLf
                Else
                    If Not dryRun Then
                        ' One field at a time. UpsertRow MERGES into the row --
                        ' it writes only the keys handed to it -- so publishing
                        ' ABOUT_BODY cannot disturb the other fields sitting in
                        ' that same row.
                        Dim oneField As Object
                        Set oneField = CreateObject("Scripting.Dictionary")
                        oneField(fieldId) = encoded
                        ExcelOutput.UpsertRow regWs, ent, oneField, period
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

    If wrongPeriodRefs <> "" Then
        report = report & vbCrLf & "SOURCES FROM ANOTHER PERIOD:" & vbCrLf & wrongPeriodRefs & _
            "The text still published. Either the citation is wrong, or that source's " & _
            "'Applies to' cell should say " & Sources.APPLIES_ALL & "." & vbCrLf
    End If

    PublishDrafts = report
End Function

' FindRegisterRow was deleted 2026-08-05 with the move to the wide sheet. It
' located a row by (EntityCode, FieldID) in the long register -- a second way of
' addressing a register, living alongside ExcelOutput's. Two addressing schemes
' for one sheet is the specific failure this codebase has paid for more than
' once (the sheet read by tab position, the register read by index), so it goes
' rather than sitting here unused waiting to be picked up again.

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

    ' NOTHING TO DO IS SAID OUT LOUD. Rohan, 2026-08-08: "copy ai to submit
    ' didnt do anything". It was right -- column F was empty, so there was
    ' nothing to copy -- but "0 copied, 0 left alone, 43 with no AI draft" reads
    ' as a machine shrugging. A tool that does nothing must say why, or the
    ' person is left deciding whether it is broken.
    If copied = 0 And keptExisting = 0 Then
        CopyAiToSubmit = "Nothing to copy: there are no AI drafts on this sheet yet." & vbCrLf & _
            "Column F (AI DRAFT) is empty for all " & noAi & " row(s)." & vbCrLf & vbCrLf & _
            "Paste Copilot's text into column F first -- the prompt is in cell L2 -- " & _
            "or just type into column D (SUBMIT) yourself." & vbCrLf
        Exit Function
    End If

    CopyAiToSubmit = "Copy AI -> Submit: " & copied & " copied, " & _
        keptExisting & " left alone (you had already written something there), " & _
        noAi & " with no AI draft." & vbCrLf & _
        "Nothing was overwritten. Edit column D, tick column E, then publish." & vbCrLf
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
