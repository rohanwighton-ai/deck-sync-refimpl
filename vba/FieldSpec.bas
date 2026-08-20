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

Public Const COL_SPEC_FIELDID As Long = 1
Public Const COL_SPEC_KIND As Long = 2
Public Const COL_SPEC_PURPOSE As Long = 3
Public Const COL_SPEC_VOICE As Long = 4
Public Const COL_SPEC_LENGTH As Long = 5
Public Const COL_SPEC_OWNJOB As Long = 6
Public Const COL_SPEC_DONOT As Long = 7

' The vocabulary a Controlled field is allowed to take, comma-separated.
'
' Evidence for why this exists, from Rohan's own register on 2026-08-01:
' PROJECT_STATUS held "Not Started" 10 times and "Not started" 5 times. Same
' status, two spellings, and nothing anywhere objected -- which is how a
' controlled field stops being controlled and a filter silently misses a third
' of its rows.
'
' Empty means unconstrained. A Prose field leaves it blank forever.
Public Const COL_SPEC_ALLOWED As Long = 8

' THE RULES THAT APPLY TO EVERY FIELD, ON THE SHEET RATHER THAN IN CODE.
'
' Three clauses used to be hardcoded in PromptFrom -- that column C is the
' standard, that the workbook is the sole source of truth, that gaps get asked
' about rather than invented. Meanwhile the index sheet told the reader this
' sheet was "the instructions the AI is given. Yours, not the tool's." That was
' false for a third of the prompt, and the third a person would most want to
' change. Rohan, 2026-08-01.
'
' Held in one cell rather than a row per clause: it is prose the AI reads
' verbatim, and splitting prose into a grid makes it harder to edit, not
' easier. The per-field columns are a grid because they are compared field to
' field; this is not.
' HOW THE FIELD'S CONTENT IS PLACED, as opposed to what it says.
'
' Rohan, 2026-08-10: "per field, on the field spec row, same principle for text
' and can extend to anchor positions for variable object placement (deliverable
' icon graphics etc)."
'
' A third axis, and it must not collapse into the other two -- this project has
' already paid once for a word doing two jobs:
'   Kind      how the content is DECIDED   (Controlled / Prose / Static)
'   FieldType what the content IS          (Text / Picture / Shape)
'   Behaviour how it is PLACED             (fill the frame / fit inside / ...)
'
' Derived from his own deck rather than invented: the project banner is cropped
' to fill its frame, while the logos and article thumbnails beside it sit
' uncropped at their own proportions. Two pictures, two behaviours, on one
' slide -- so it cannot be a global setting.
'
' The code owns the vocabulary and the sheet owns the assignment, the same split
' as Kind and the Sources period list. Picked, never typed.
Public Const COL_SPEC_BEHAVIOUR As Long = 10

Public Const BEHAVIOUR_FILL As String = "Fill the frame"
Public Const BEHAVIOUR_FIT As String = "Fit inside"
Public Const BEHAVIOUR_ASIS As String = "Leave as is"

' WHAT THE FIELD RENDERS AS -- the axis the comment above calls FieldType.
'
' NOT named `FieldType` in code, deliberately: BatchOnboardFlow already has a
' `fieldType` meaning text/number/currency/date, a hint about the CONTENT. Two
' things called FieldType is the exact defect the axes comment above warns
' about, and this project has already lost a feature to one word doing two
' jobs. `Renders as` is the sheet's word and RENDER_* is the code's.
'
' WHY IT MUST BE DECLARED AND CANNOT BE DERIVED, which is the whole reason this
' column exists when the sync router happily reads the shape instead:
'
'   at WRITE time  the shape answers "what am I" -- a picture is a picture, a
'                  shape with a `.track` sibling is a bar. InjectField reads it
'                  and cannot be wrong, because it IS the thing being written.
'   at MARK time   there is no shape yet, or there are two anonymous
'                  rectangles. Nothing can answer, so someone has to say.
'
' That is intent versus reality, not two answers to one question -- and
' comparing them is precisely what FieldWiring's check is for. Blank means
' Text: every field that existed before this column was added is text, and a
' blank that meant "unknown" would make every old sheet report a problem.
Public Const COL_SPEC_RENDERS As Long = 11

Public Const RENDER_TEXT As String = "Text"
Public Const RENDER_PICTURE As String = "Picture"
Public Const RENDER_PROGRESS As String = "Progress bar"

' FIX-LIST P4, 2026-08-19: the fourth render kind, for a field that is not
' ONE shape but several -- HIGHLIGHTS_BODY is the specimen (three shapes per
' slide, one field). Its own purpose is to be checked FOR, not computed from:
' `ExcelOutput.MissingRegisterColumns` excludes any field marked this way
' from the bundled "add a column for each" prompt, because
' `AddRegisterColumns` can only ever create exactly one column, which is the
' wrong structure for a Slots field -- the same "a real architectural
' decision made silently by a Yes on a bundled prompt" this project has
' already refused once for Derived fields.
Public Const RENDER_SLOTS As String = "Slots"

Public Const COL_SPEC_GLOBAL As Long = 9
Public Const SPEC_GLOBAL_ROW As Long = 2

' HOW REPORTED LAST TIME SHOULD BE USED FOR THIS FIELD -- the axis that
' replaced the one-size "match voice, never repeat words" global instruction
' added 2026-08-20 alongside the Sources/REPORTED LAST TIME column swap. That
' single rule was right for FRESH fields and wrong for CARRY ones:
' DELIVERABLES_BODY's own recipe already forbids re-sifting an unchanged
' value once one exists, and a global "every row needs a genuine attempt this
' quarter" told it the opposite. Four treatments, not one:
'   CARRY        -- unchanged unless a source changed (most Standing prose)
'   FRESH        -- new words every quarter (KEY_EVENTS_BODY, PROGRESS_BODY)
'   PART-FROZEN  -- carried items keep their exact words; only new items are
'                   written (HIGHLIGHTS_BODY, whose bullets carry a quarter tag)
'   DIFF         -- no voice relationship; history is a change detector only
'                   (STRATEGIC_LINKAGES)
' The DEFINITIONS live in GLOBAL RULES (DefaultGlobalRules, below), because
' that is the sheet's own editable prose -- same split as every other axis on
' this sheet: the code owns the vocabulary, the sheet owns which value
' applies to a given field and can edit what each value means. PICKED, NEVER
' TYPED, so this gets the same dropdown treatment as Behaviour and Renders as.
Public Const COL_SPEC_HISTORY As Long = 13
Public Const COL_SPEC_HISTORY_NOTES As Long = 14

Public Const HIST_CARRY As String = "CARRY"
Public Const HIST_FRESH As String = "FRESH"
Public Const HIST_PARTFROZEN As String = "PART-FROZEN"
Public Const HIST_DIFF As String = "DIFF"

' The heading HistoryPreamble sends a drafter to, and the heading
' DefaultGlobalRules writes. It was the same string typed in both places --
' a second copy of a machine-knowable fact, and the cross-reference silently
' breaks if either is edited. One constant so they cannot disagree.
Public Const HISTORY_ANCHOR As String = "REPORTED LAST TIME IS A SOURCE"

Public Const SPEC_HEADER_ROW As Long = 1
Public Const SPEC_FIRST_ROW As Long = 2

Public Type FieldGuidance
    Found As Boolean
    FieldId As String
    Kind As String          ' Controlled / Prose / Static -- see ReviewQueue
    Allowed As String       ' comma-separated vocabulary, Controlled fields only
    Purpose As String
    Voice As String
    Length As String
    OwnJob As String
    DoNot As String
    GlobalRules As String   ' shared by every field -- see COL_SPEC_GLOBAL
    Behaviour As String     ' how the content is PLACED -- see COL_SPEC_BEHAVIOUR
    History As String       ' CARRY / FRESH / PART-FROZEN / DIFF -- see COL_SPEC_HISTORY
    HistoryNotes As String  ' optional field-specific addition -- see COL_SPEC_HISTORY_NOTES
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
    Do While Trim(CStr(ws.Cells(r, COL_SPEC_FIELDID).Value)) <> ""
        existing(UCase(Trim(CStr(ws.Cells(r, COL_SPEC_FIELDID).Value)))) = True
        r = r + 1
    Loop
    On Error GoTo 0

    ws.Cells(SPEC_HEADER_ROW, COL_SPEC_FIELDID).Value = "FieldID"
    ws.Cells(SPEC_HEADER_ROW, COL_SPEC_KIND).Value = "Kind (Controlled/Prose/Static/Derived)"
    ws.Cells(SPEC_HEADER_ROW, COL_SPEC_PURPOSE).Value = "Purpose -- the question this field answers"
    ws.Cells(SPEC_HEADER_ROW, COL_SPEC_VOICE).Value = "Voice"
    ws.Cells(SPEC_HEADER_ROW, COL_SPEC_LENGTH).Value = "Length"
    ws.Cells(SPEC_HEADER_ROW, COL_SPEC_OWNJOB).Value = "Own-job test"
    ws.Cells(SPEC_HEADER_ROW, COL_SPEC_DONOT).Value = "Do NOT"
    ws.Cells(SPEC_HEADER_ROW, COL_SPEC_ALLOWED).Value = "Allowed values (Controlled fields -- comma separated)"
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
    ' THE THREE PANELS THAT HAD NO RECIPE, transcribed from the CRC prompt
    ' library's Prompt 18 (locked definitions, in use at work). Added 2026-08-09.
    '
    ' WHY THESE THREE, AND WHY AHEAD OF EVERYTHING ELSE ON THE LIST. A quarter by
    ' hand costs Rohan about three weeks of nights and MOST OF IT IS RE-DECIDING --
    ' re-deriving what each section should say, every quarter, from memory. A
    ' recipe removes exactly that, and sync cannot touch it. Strategic Alignment
    ' and Problem are the two hardest sections to re-decide (the "so what" and the
    ' "why"), and Progress is the one that changes most.
    '
    ' The three of them plus ABOUT_BODY are the set Prompt 18 warns are "most
    ' prone to bleeding into each other", which is why each Own-job test names the
    ' neighbours it must not stray into. That test is the part doing the work: it
    ' is what a person checks against at 11pm, not the prose above it.
    If Not existing.Exists("STRATEGIC_ALIGNMENT_BODY") Then
        SeedRow ws, r, "STRATEGIC_ALIGNMENT_BODY", "Prose", _
            "The ""so what"" for SAAFE and Australia. Connects the project's outcomes to SAAFE's strategic objectives, the declared linkage codes, and the broader sector or AMR benefit.", _
            "Taciturn, interpretive, professional. Forward-looking and value-oriented. No promotional language, no hedging, no padding. Assumes AMR literacy.", _
            "About two short paragraphs. Target 600-800 characters.", _
            "Does it say why this MATTERS to SAAFE and Australia -- without describing the technology (that is ABOUT_BODY) or setting out the sector gap (that is PROBLEM_BODY)?", _
            "Describe the approach or technology in detail. Restate the sector gap. Present an INFERRED linkage code as a declared one -- if it cannot be confirmed, write [TBC] and say so. Invent codes, figures, organisations or outcomes not in the workbook."
        r = r + 1: added = added + 1
    End If
    If Not existing.Exists("PROBLEM_BODY") Then
        SeedRow ws, r, "PROBLEM_BODY", "Prose", _
            "The ""why"". The need this project addresses: the sector challenge, why current approaches are inadequate, and briefly why this approach is a candidate.", _
            "Taciturn and factual. Present tense. Assumes AMR literacy. No promotional language.", _
            "A few sentences. Target 300-500 characters.", _
            "Does it establish a need that EXISTS WHETHER OR NOT THIS PROJECT HAPPENS -- rather than describing the project's own activity or claiming its value?", _
            "Describe what the project does (that is ABOUT_BODY). Claim strategic value (that is STRATEGIC_ALIGNMENT_BODY). Overstate the gap. Invent facts not in the workbook."
        r = r + 1: added = added + 1
    End If
    If Not existing.Exists("PROGRESS_BODY") Then
        SeedRow ws, r, "PROGRESS_BODY", "Prose", _
            "What the most recent reporting quarter amounts to, interpreted rather than listed. Opens with a header carrying the quarter label, e.g. ""Last reported quarter update - Q4F26"".", _
            "Interpretive and taciturn. States material constraints or blockers plainly rather than softening them. No promotional language.", _
            "Two or three sentences, or short bullets, after the quarter-labelled header.", _
            "Does it INTERPRET the quarter rather than list activity -- and does every claim carry the quarter it was reported against?", _
            "Dump a raw activity list. Restate standing facts about the project (that is ABOUT_BODY). Report an event as this quarter's when the workbook does not say so. GUESS a quarter tag -- leave it untagged and flag it instead."
        r = r + 1: added = added + 1
    End If
    ' THE ELAPSED-TIME BAR. Kind = Derived (ExcelOutput.KIND_DERIVED) -- computed
    ' fresh every sync from START_DATE/END_DATE, never drafted, never a register
    ' column of its own. Rohan, 2026-08-09: "time elapsed bar autoshapes that
    ' move with the clock regardless of progress." Purpose/Voice/etc are still
    ' filled in for consistency with every other row on this sheet, even though
    ' a Derived field is never drafted and none of this guidance is ever read by
    ' a person or an AI -- SeedRow has no optional parameters, and inventing a
    ' second row shape for one Kind would be a bigger change than filling five
    ' cells nothing reads.
    If Not existing.Exists("TIMELINE_ELAPSED") Then
        SeedRow ws, r, "TIMELINE_ELAPSED", "Derived", _
            "How far the project is through its declared timeline, as a fraction of today's date between START_DATE and END_DATE.", _
            "N/A -- computed, never drafted.", _
            "N/A.", _
            "N/A.", _
            "N/A -- this field is never drafted; SyncOperations.ElapsedFraction computes it directly from START_DATE and END_DATE at sync time."
        r = r + 1: added = added + 1
    End If
    ' THE STATUS BADGE. Kind = Derived, same class as TIMELINE_ELAPSED above --
    ' computed fresh every sync from PROJECT_STATUS/SCHEDULE_STATUS, never
    ' drafted, never a register column of its own. Rohan wrote the real
    ' derivation rule directly onto the live Field Spec sheet's row (a
    ' priority table combining lifecycle stage and schedule health into one
    ' word), addressed to "Claude Code" by name -- this SeedRow default won't
    ' touch that row (SeedRow only fires when the FieldID is missing), it
    ' only gives a FRESH workbook the same field with equivalent guidance.
    ' Full logic: SyncOperations.DeriveStatusBadge.
    If Not existing.Exists("STATUS_BADGE") Then
        SeedRow ws, r, "STATUS_BADGE", "Derived", _
            "The single status word shown in the header bar. Combines lifecycle stage (PROJECT_STATUS) with schedule health (SCHEDULE_STATUS) so the badge never shows two things at once.", _
            "N/A -- computed, never drafted.", _
            "N/A.", _
            "N/A.", _
            "N/A -- this field is never drafted; SyncOperations.DeriveStatusBadge computes it directly from PROJECT_STATUS and SCHEDULE_STATUS at sync time."
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

    ' TIMELINE_ELAPSED renders as a progress bar, not text -- set explicitly
    ' rather than relying on the blank-fills-to-Text pass below, same as
    ' PROJECT_STATUS's Allowed-values pass immediately after this one.
    Dim er As Long
    er = SPEC_FIRST_ROW
    Do While Trim(CStr(ws.Cells(er, COL_SPEC_FIELDID).Value)) <> ""
        If StrComp(Trim(CStr(ws.Cells(er, COL_SPEC_FIELDID).Value)), "TIMELINE_ELAPSED", vbTextCompare) = 0 Then
            If Trim(CStr(ws.Cells(er, COL_SPEC_RENDERS).Value)) = "" Then
                ws.Cells(er, COL_SPEC_RENDERS).Value = RENDER_PROGRESS
            End If
        End If
        er = er + 1
    Loop

    ' Seeded onto whatever row PROJECT_STATUS is on, only when that cell is
    ' still empty -- like every other cell on this sheet, an edit is the
    ' owner's and survives a rebuild.
    Dim vr As Long
    vr = SPEC_FIRST_ROW
    Do While Trim(CStr(ws.Cells(vr, COL_SPEC_FIELDID).Value)) <> ""
        If StrComp(Trim(CStr(ws.Cells(vr, COL_SPEC_FIELDID).Value)), "PROJECT_STATUS", vbTextCompare) = 0 Then
            If Trim(CStr(ws.Cells(vr, COL_SPEC_ALLOWED).Value)) = "" Then
                ws.Cells(vr, COL_SPEC_ALLOWED).Value = "In Progress,Not Started,Project Closed"
            End If
        End If
        vr = vr + 1
    Loop
    ws.Columns(COL_SPEC_ALLOWED).ColumnWidth = 40
    ws.Columns(COL_SPEC_ALLOWED).WrapText = True

    ' The global clauses. Seeded once with what used to be hardcoded, then
    ' never touched again -- like every other row on this sheet, an edit here
    ' is the owner's and must survive a rebuild.
    ws.Cells(SPEC_HEADER_ROW, COL_SPEC_BEHAVIOUR).Value = "Behaviour  --  how the content is PLACED (pictures and objects)"
    ws.Columns(COL_SPEC_BEHAVIOUR).ColumnWidth = 18

    ' WHAT IT RENDERS AS. Read when the field is TAGGED, not when it is written
    ' -- at tag time a bar is two anonymous rectangles and nothing else can say
    ' which is which.
    ws.Cells(SPEC_HEADER_ROW, COL_SPEC_RENDERS).Value = _
        "Renders as  --  what the field IS on the slide. A '" & RENDER_PROGRESS & _
        "' is tagged as a PAIR: the bar and its track."
    ws.Columns(COL_SPEC_RENDERS).ColumnWidth = 16

    ws.Cells(SPEC_HEADER_ROW, COL_SPEC_HISTORY).Value = _
        "History treatment -- " & HIST_CARRY & " / " & HIST_FRESH & " / " & HIST_PARTFROZEN & _
        " / " & HIST_DIFF & " (controlled; definitions in GLOBAL RULES)"
    ws.Columns(COL_SPEC_HISTORY).ColumnWidth = 16
    ws.Cells(SPEC_HEADER_ROW, COL_SPEC_HISTORY_NOTES).Value = _
        "History notes -- optional, field-specific addition to the declared treatment. Leave blank if none."
    ws.Columns(COL_SPEC_HISTORY_NOTES).ColumnWidth = 46
    ws.Columns(COL_SPEC_HISTORY_NOTES).WrapText = True

    ' Blank is filled with Text rather than left empty. Every field that existed
    ' before this column is text, and a blank meaning "unknown" would make every
    ' pre-existing sheet report a problem it does not have. An owner who wants
    ' something else picks it from the dropdown.
    Dim rr As Long
    rr = SPEC_FIRST_ROW
    Do While Trim(CStr(ws.Cells(rr, COL_SPEC_FIELDID).Value)) <> ""
        If Trim(CStr(ws.Cells(rr, COL_SPEC_RENDERS).Value)) = "" Then
            ws.Cells(rr, COL_SPEC_RENDERS).Value = RENDER_TEXT
        End If
        rr = rr + 1
    Loop
    ws.Cells(SPEC_HEADER_ROW, COL_SPEC_GLOBAL).Value = "GLOBAL RULES  --  added to EVERY field's prompt. Edit freely."
    If Trim(CStr(ws.Cells(SPEC_GLOBAL_ROW, COL_SPEC_GLOBAL).Value)) = "" Then
        ws.Cells(SPEC_GLOBAL_ROW, COL_SPEC_GLOBAL).Value = "'" & DefaultGlobalRules()
    End If
    ws.Columns(COL_SPEC_GLOBAL).ColumnWidth = 60
    ws.Columns(COL_SPEC_GLOBAL).WrapText = True
    ws.Rows(SPEC_GLOBAL_ROW).RowHeight = 90

    ' 8pt, matching every other sheet the tools write.
    ws.Cells.Font.Size = 8
    ws.Cells.VerticalAlignment = -4160        ' xlTop
    ws.Columns(COL_SPEC_FIELDID).ColumnWidth = 20
    ws.Columns(COL_SPEC_KIND).ColumnWidth = 16
    Dim c As Long
    For c = COL_SPEC_PURPOSE To COL_SPEC_DONOT
        ws.Columns(c).ColumnWidth = 46
        ws.Columns(c).WrapText = True
    Next c

    WriteSpecSheet = "Field Spec: " & existing.count & " row(s) kept, " & added & " seeded."
End Function

Private Sub SeedRow(ws As Object, r As Long, fieldId As String, kind As String, _
                    purpose As String, voice As String, length As String, _
                    ownJob As String, doNot As String)
    ws.Cells(r, COL_SPEC_FIELDID).Value = fieldId
    ws.Cells(r, COL_SPEC_KIND).Value = kind
    ws.Cells(r, COL_SPEC_PURPOSE).Value = "'" & purpose
    ws.Cells(r, COL_SPEC_VOICE).Value = "'" & voice
    ws.Cells(r, COL_SPEC_LENGTH).Value = "'" & length
    ws.Cells(r, COL_SPEC_OWNJOB).Value = "'" & ownJob
    ws.Cells(r, COL_SPEC_DONOT).Value = "'" & doNot
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

    ' Read regardless of whether the field has a row of its own: the global
    ' rules apply to every field, including the ones running unguided.
    On Error Resume Next
    g.GlobalRules = Trim(CStr(ws.Cells(SPEC_GLOBAL_ROW, COL_SPEC_GLOBAL).Value))
    On Error GoTo 0

    Dim r As Long
    r = SPEC_FIRST_ROW
    On Error Resume Next
    Do While Trim(CStr(ws.Cells(r, COL_SPEC_FIELDID).Value)) <> ""
        If StrComp(Trim(CStr(ws.Cells(r, COL_SPEC_FIELDID).Value)), fieldId, vbTextCompare) = 0 Then
            g.Found = True
            g.Kind = Trim(CStr(ws.Cells(r, COL_SPEC_KIND).Value))
            g.Purpose = Trim(CStr(ws.Cells(r, COL_SPEC_PURPOSE).Value))
            g.Voice = Trim(CStr(ws.Cells(r, COL_SPEC_VOICE).Value))
            g.Length = Trim(CStr(ws.Cells(r, COL_SPEC_LENGTH).Value))
            g.OwnJob = Trim(CStr(ws.Cells(r, COL_SPEC_OWNJOB).Value))
            g.DoNot = Trim(CStr(ws.Cells(r, COL_SPEC_DONOT).Value))
            g.Allowed = Trim(CStr(ws.Cells(r, COL_SPEC_ALLOWED).Value))
                ' Per ROW, unlike GlobalRules which comes from a fixed cell.
                g.Behaviour = Trim(CStr(ws.Cells(r, COL_SPEC_BEHAVIOUR).Value))
                g.History = Trim(CStr(ws.Cells(r, COL_SPEC_HISTORY).Value))
                g.HistoryNotes = Trim(CStr(ws.Cells(r, COL_SPEC_HISTORY_NOTES).Value))
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
' citedBlock is APPENDED AFTER the global rules, deliberately.
'
' The rules end with "the workbook is the sole source of truth", and the sheet's
' own copy of that wording wins over the built-in one. So widening the built-in
' would do nothing on any workbook that has its own -- the exact "editing the
' cell appears to do nothing" failure the comment below warns about, inverted.
' Coming last, the sources block states the wider rule itself and cannot be
' silently overridden by either copy.
' What to tell the drafter about REPORTED LAST TIME before the field's own
' guidance, keyed off the field's declared History treatment (COL_SPEC_
' HISTORY) rather than a hardcoded column letter. This used to be a single
' sentence naming a specific column ("the existing text in column C") for
' EVERY field regardless of what that field actually needed from its own
' history -- and it went stale the moment Sources and REPORTED LAST TIME
' swapped columns on 2026-08-20, because nothing here derived it. Found by
' Claude (chat) asking why a per-field bug report didn't match either
' Purpose or Voice: it wasn't in either, because it was never on the sheet at
' all -- it was hardcoded in this function, applying to all thirteen fields
' at once, not just the one it was first noticed on.
Private Function HistoryPreamble(g As FieldGuidance) As String
    Dim s As String
    If Trim(g.History) = "" Then
        s = "Read the workbook, including REPORTED LAST TIME and the cited SOURCES."
    Else
        s = "This field's declared History treatment is " & g.History & _
            " -- see " & HISTORY_ANCHOR & " in the rules below for what that means."
        If Trim(g.HistoryNotes) <> "" Then s = s & vbCrLf & g.HistoryNotes
    End If
    HistoryPreamble = s
End Function

Public Function PromptFrom(g As FieldGuidance, Optional citedBlock As String = "") As String
    Dim s As String
    s = HistoryPreamble(g) & vbCrLf & vbCrLf

    Dim submitCol As String
    submitCol = Chr$(64 + Drafting.COL_D_SUBMIT) & " (SUBMIT)"

    If Not g.Found Then
        s = s & "NOTE: no Field Spec row exists for " & g.FieldId & ", so this is generic" & vbCrLf & _
            "guidance. Add a row to the 'Field Spec' sheet to say how this field" & vbCrLf & _
            "should be written." & vbCrLf & vbCrLf & _
            "Write an updated version of " & g.FieldId & " for each project into " & submitCol & " ONLY." & vbCrLf & vbCrLf
    Else
        s = s & "WHAT THIS FIELD IS FOR" & vbCrLf & g.Purpose & vbCrLf & vbCrLf & _
            "VOICE" & vbCrLf & g.Voice & vbCrLf & vbCrLf & _
            "LENGTH" & vbCrLf & g.Length & vbCrLf & vbCrLf & _
            "BEFORE YOU WRITE EACH ROW, ASK" & vbCrLf & g.OwnJob & vbCrLf & vbCrLf & _
            "DO NOT" & vbCrLf & g.DoNot & vbCrLf & vbCrLf & _
            "Write into " & submitCol & " ONLY. Leave every other column untouched." & vbCrLf & vbCrLf
    End If

    ' The sheet's wording wins when it has any. The built-in is a fallback for
    ' a workbook that predates the global cell, never an override of it --
    ' otherwise editing the cell would appear to do nothing, which is the worst
    ' possible behaviour for a control whose whole promise is "yours, not the
    ' tool's".
    If Trim(g.GlobalRules) <> "" Then
        s = s & g.GlobalRules
    Else
        s = s & DefaultGlobalRules()
    End If

    If Trim(citedBlock) <> "" Then s = s & citedBlock

    PromptFrom = s
End Function

' What the global clauses say when nobody has edited them. Seeded onto the
' sheet by WriteSpecSheet and used as the fallback by PromptFrom, so the two
' can never drift apart -- they were one hardcoded string in two conceptual
' places before, which is how the sheet came to claim ownership of text it did
' not hold.
' REWRITTEN AGAIN 2026-08-20, hours after the first rewrite. That first
' version replaced "column C is the standard, leave the row blank if it
' already does the job" with "every row needs a genuine attempt this
' quarter, even where nothing changed" -- correct for FRESH fields
' (KEY_EVENTS_BODY, PROGRESS_BODY) and wrong for CARRY ones:
' DELIVERABLES_BODY's own recipe already forbids re-sifting an unchanged
' value once one exists, and MSn_LABEL's four-word compression is decided
' ONCE, not re-decided every quarter on a funder-facing slide. A single rule
' cannot be right for both, so this now states all four treatments and lets
' each field's declared History treatment (COL_SPEC_HISTORY) pick which one
' applies -- caught by Claude (chat) reviewing the first rewrite before it
' propagated into thirteen fields' worth of individually-tailored prompts.
'
' FRESH extended 2026-08-20 (same night): the original wording only asked for
' non-repetition and closed loops -- functional continuity, not narrative
' continuity. Rohan's own framing: a fresh quarter should be free to continue
' a real thread already in progress, not read as a cold restart every time.
' Deliberately understated, not "always frame as a story" -- this is
' Commonwealth-program reporting, most of it plainly factual, and the first
' draft of this clause used narrative language (build-up, turning point,
' resolution) that risked reading drama into milestone comments nobody
' intended as one. Softened same session, before it ever reached the live
' register: pick up a thread if it's genuinely there, never manufacture one.
Public Function DefaultGlobalRules() As String
    ' Built as several statements, not one -- VBA caps a single continued
    ' statement at 25 physical lines, and the four treatment definitions
    ' together ran past it as one assignment.
    Dim s As String
    s = HISTORY_ANCHOR & ". It is what this exact field said for" & vbCrLf & _
        "this exact project last quarter -- a source for voice, continuity and" & vbCrLf & _
        "scope, never for facts. Facts come from the cited SOURCES only. How you" & vbCrLf & _
        "use it depends on this field's declared History treatment, stated above:" & vbCrLf & vbCrLf

    s = s & HIST_CARRY & " -- this field describes something that does not change" & vbCrLf & _
        "quarterly. Last quarter's text is your starting point and, where the" & vbCrLf & _
        "sources are unchanged, your answer. Do not rewrite it for freshness." & vbCrLf & _
        "Rewrite only where a source has actually changed, and say in NOTES what" & vbCrLf & _
        "changed. An unchanged field is a correct outcome, not a skipped one." & vbCrLf & vbCrLf

    s = s & HIST_FRESH & " -- this quarter gets its own words. Read REPORTED LAST" & vbCrLf & _
        "TIME to learn two things: what has already been reported (do not" & vbCrLf & _
        "report it again as if new), and what was left open or unresolved (say" & vbCrLf & _
        "what happened to it). Match its register; reuse none of its sentences. If a" & vbCrLf & _
        "real thread genuinely carries from one quarter to the next, it's fine to" & vbCrLf & _
        "pick it up rather than restate it cold -- but most reporting is not a story" & vbCrLf & _
        "and should not be written like one. This is about not ignoring an obvious" & vbCrLf & _
        "continuation, never about inventing one." & vbCrLf & vbCrLf

    s = s & HIST_PARTFROZEN & " -- items carried over from a previous quarter keep" & vbCrLf & _
        "their exact original wording, because they carry the quarter tag they" & vbCrLf & _
        "were reported under. Rewording a carried item makes its tag false." & vbCrLf & _
        "Write new items fresh; the judgement this quarter is which items are" & vbCrLf & _
        "carried and which are displaced." & vbCrLf & vbCrLf

    s = s & HIST_DIFF & " -- no voice relationship. Compare against REPORTED LAST" & vbCrLf & _
        "TIME only to detect change. Any difference is either a formal" & vbCrLf & _
        "variation or an error, and must be traced to the source before it is" & vbCrLf & _
        "submitted." & vbCrLf & vbCrLf

    s = s & "The workbook is the sole source of truth. Do not introduce facts," & vbCrLf & _
        "figures, organisations or outcomes that are not in it. Where something" & vbCrLf & _
        "needed is missing or ambiguous, say so in " & Chr$(64 + Drafting.COL_D_NOTES) & _
        " (NOTES) and ask -- do not infer or fill the gap."

    DefaultGlobalRules = s
End Function

' Every field's declared Kind, read in ONE pass, keyed by upper-case FieldID.
'
' Built as a map rather than a per-field lookup on purpose. Its caller
' (ReviewQueue.AssignBatches) asks for a kind twice per queue item inside two
' loops; resolving each of those against the sheet directly would be the same
' rescan-from-row-1-inside-a-loop shape this project has already paid to fix
' twice (FIX-LIST items W and AB). One pass, then dictionary lookups.
'
' Returns an empty Dictionary rather than Nothing when the sheet cannot be
' read, so a caller's fallback path is taken per field instead of the whole
' mechanism failing at once.
Public Function KindMap(ws As Object) As Object
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")
    Set KindMap = m

    On Error GoTo Done
    If ws Is Nothing Then Exit Function

    Dim r As Long
    Dim fid As String
    r = SPEC_FIRST_ROW
    Do While Trim(CStr(ws.Cells(r, COL_SPEC_FIELDID).Value)) <> ""
        fid = UCase$(Trim(CStr(ws.Cells(r, COL_SPEC_FIELDID).Value)))
        If Not m.Exists(fid) Then
            m.Add fid, Trim(CStr(ws.Cells(r, COL_SPEC_KIND).Value))
        End If
        r = r + 1
    Loop

Done:
End Function

' Does the GLOBAL RULES cell still define what the prompts send a drafter to?
'
' The cell is the owner's, deliberately: PromptFrom prefers it over
' DefaultGlobalRules whenever it holds anything, so that editing it is never a
' no-op. But WriteSpecSheet seeds it ONLY when blank and nothing reconciles it
' afterwards, so a cell seeded before a rules change goes on winning long after
' the code has moved. Not hypothetical -- on 2026-08-20 the live cell still
' held the pre-treatment wording: it named "Column C" for last quarter's text
' (layout 7 had made C the SOURCES column) and "column J" for NOTES (which is
' I), and contained none of the four treatments, while every field's prompt was
' telling the drafter to read those treatment definitions in the rules below.
' Thirteen prompts cross-referencing a section that was not there.
'
' This does NOT overwrite the owner's text, and must not. It answers one
' structural question: does the heading the code POINTS AT, and the treatment
' each field actually DECLARES, exist in the cell being pointed into. No
' threshold and no judgement in that -- a cross-reference either resolves or it
' does not, which is why this can fail for a real reason and cannot fire on a
' cell that is merely worded differently from the default.
'
' Returns "" when sound, and when the cell is blank -- WriteSpecSheet reseeds a
' blank cell from DefaultGlobalRules, so blank is self-healing and is the
' recovery path: clear the cell, run the chain, get the current wording back.
Public Function GlobalRulesProblem(ws As Object) As String
    Dim rules As String
    Dim used As Object
    Dim missing As String
    Dim r As Long
    Dim h As String
    Dim k As Variant

    On Error GoTo Fail

    rules = Trim(CStr(ws.Cells(SPEC_GLOBAL_ROW, COL_SPEC_GLOBAL).Value))
    If rules = "" Then Exit Function

    Set used = CreateObject("Scripting.Dictionary")
    r = SPEC_FIRST_ROW
    Do While Trim(CStr(ws.Cells(r, COL_SPEC_FIELDID).Value)) <> ""
        h = UCase$(Trim(CStr(ws.Cells(r, COL_SPEC_HISTORY).Value)))
        If h <> "" Then
            If Not used.Exists(h) Then used.Add h, True
        End If
        r = r + 1
    Loop

    ' No field declares a treatment, so no prompt cross-references the anchor.
    If used.Count = 0 Then Exit Function

    If InStr(1, rules, HISTORY_ANCHOR, vbTextCompare) = 0 Then
        missing = HISTORY_ANCHOR
    End If

    For Each k In used.Keys
        If InStr(1, rules, CStr(k), vbTextCompare) = 0 Then
            If missing <> "" Then missing = missing & ", "
            missing = missing & CStr(k)
        End If
    Next k

    If missing = "" Then Exit Function

    GlobalRulesProblem = "GLOBAL RULES on Field Spec does not define: " & missing & _
        ". Every field prompt sends the drafter there for it. Clear that cell and" & _
        " run the chain again to restore the current wording, or add the text yourself."
    Exit Function

Fail:
    ' Never report clean on a check that could not run -- that is the shape this
    ' project has been burned by more than once.
    GlobalRulesProblem = "Could not check GLOBAL RULES on Field Spec: " & Err.Description
End Function

' Put a real dropdown on every Controlled field's Value cell in the register,
' and report anything already sitting there that the vocabulary does not allow.
'
' VALIDATES BUT DOES NOT CORRECT. "Not started" is almost certainly meant to be
' "Not Started" -- and almost certainly is not a licence for a tool to rewrite
' the record without being asked. It is reported; a person decides. Same posture
' as unresolved source IDs.
'
' Excel's own dropdown rather than a checker of our own: a constraint that can
' only be violated is worth less than one that is awkward to violate, and the
' cheapest place to stop a typo is the cell it would be typed into.
Public Function ApplyControlledValidation(regWs As Object, specWs As Object, Optional ByRef outOfVocabulary As Long) As String
    If regWs Is Nothing Or specWs Is Nothing Then
        ApplyControlledValidation = "Validation: skipped (no register or no Field Spec)."
        Exit Function
    End If

    ' Vocabulary per field, read once.
    Dim vocab As Object
    Set vocab = CreateObject("Scripting.Dictionary")
    Dim r As Long
    r = SPEC_FIRST_ROW
    Do While Trim(CStr(specWs.Cells(r, COL_SPEC_FIELDID).Value)) <> ""
        Dim allowed As String
        allowed = Trim(CStr(specWs.Cells(r, COL_SPEC_ALLOWED).Value))
        If allowed <> "" Then
            vocab(UCase(Trim(CStr(specWs.Cells(r, COL_SPEC_FIELDID).Value)))) = allowed
        End If
        r = r + 1
    Loop
    If vocab.count = 0 Then
        ApplyControlledValidation = "Validation: no field declares an allowed-value list."
        Exit Function
    End If

    ' A FIELD IS A COLUMN. This walked ROWS keyed by FieldID/Value until
    ' 2026-08-05 -- the long register's shape. Against the wide sheet it found
    ' no such headers and returned "register has no FieldID/Value column",
    ' which is a sentence, not an error: the caller printed it and carried on,
    ' so every controlled field lost its dropdown AND its out-of-vocabulary
    ' check the moment the register went wide, and nothing said so.
    '
    ' It FAILED SOFT, which is why it survived. The whole point of this function
    ' is to catch a value drifting out of its vocabulary; a version that
    ' silently checks nothing reports exactly what a clean register reports.
    Dim cInstance As Long, cQuarter As Long
    ExcelOutput.LocateStructuralColumns regWs, cInstance, cQuarter

    Dim lastCol As Long, lastRow As Long
    lastCol = ExcelOutput.LastUsedColumn(regWs)
    lastRow = ExcelOutput.LastUsedRow(regWs)

    ' A sheet with a header row and no data rows is legal -- a freshly created
    ' register. There is nothing to validate and nothing wrong.
    If lastCol = 0 Or lastRow < 2 Then
        ApplyControlledValidation = "Validation: no data rows on the register yet."
        Exit Function
    End If

    Dim applied As Long, controlledCols As Long, offending As String, offendingCount As Long
    Dim c As Long
    For c = 1 To lastCol
        ' Structural columns are never fields and must never be given a
        ' vocabulary -- writing a dropdown onto Instance ID or Quarter would
        ' constrain the row's identity or its period.
        If c <> cInstance And c <> cQuarter Then
            Dim fid As String
            fid = UCase(Trim(CStr(regWs.Cells(1, c).Value)))

            If vocab.Exists(fid) Then
                controlledCols = controlledCols + 1

                ' The whole column's data cells in one call, rather than cell by
                ' cell. Must never break a caller: a workbook opened read-only,
                ' or a protected sheet, will refuse validation.
                Dim rng As Object
                Set rng = regWs.Range(regWs.Cells(2, c), regWs.Cells(lastRow, c))

                On Error Resume Next
                rng.Validation.Delete
                rng.Validation.Add 3, 1, 1, CStr(vocab(fid))   ' xlValidateList, xlValidAlertStop, xlBetween
                rng.Validation.IgnoreBlank = True
                rng.Validation.InCellDropdown = True
                If Err.Number = 0 Then applied = applied + (lastRow - 1)
                Err.Clear
                On Error GoTo 0

                Dim rr As Long
                For rr = 2 To lastRow
                    Dim current As String
                    current = Trim(CStr(regWs.Cells(rr, c).Value))
                    If current <> "" Then
                        If Not InVocabulary(current, CStr(vocab(fid))) Then
                            ' Named by SLIDE and FIELD, not by row number. A row
                            ' number is worthless on a sheet where the same
                            ' project appears once per period -- the reader has
                            ' to go and look up which row that was.
                            Dim who As String
                            who = Trim(CStr(regWs.Cells(rr, cInstance).Value))
                            If cQuarter > 0 Then
                                who = who & " (" & Trim(CStr(regWs.Cells(rr, cQuarter).Value)) & ")"
                            End If
                            offending = offending & "  " & who & "  " & fid & " = """ & current & """" & vbCrLf
                            offendingCount = offendingCount + 1
                        End If
                    End If
                Next rr
            End If
        End If
    Next c

    ' SAYS WHEN IT DID NOTHING, and why. "dropdown on 0 cell(s)" reads as
    ' success to anyone skimming, and reading it that way is exactly how this
    ' function stayed broken. No controlled column on the register is a real
    ' state worth naming -- it means the vocabulary exists on the Field Spec
    ' but no column on the register carries that field.
    If controlledCols = 0 Then
        ApplyControlledValidation = "Validation: NO column on the register matches a field " & _
            "with an allowed-value list, so nothing was checked."
        Exit Function
    End If

    ApplyControlledValidation = "Validation: dropdown on " & applied & " cell(s) across " & _
        controlledCols & " controlled column(s)."
    If offending <> "" Then
        ' COUNTED IN FULL, LISTED IN PART. This went into a MsgBox that truncates
        ' near 1024 characters without saying so, so an unbounded list did not
        ' just read badly -- past the cap it pushed everything after it out of
        ' the dialog entirely. The count is the part that must always survive.
        Dim shown As String, lines As Long, pos As Long, nextPos As Long
        pos = 1
        Do While pos <= Len(offending) And lines < 5
            nextPos = InStr(pos, offending, vbCrLf)
            If nextPos = 0 Then Exit Do
            shown = shown & Mid$(offending, pos, nextPos - pos + 2)
            lines = lines + 1
            pos = nextPos + 2
        Loop

        Dim total As Long
        total = offendingCount
        outOfVocabulary = total

        ApplyControlledValidation = ApplyControlledValidation & vbCrLf & _
            total & " value(s) outside the allowed list (left exactly as they are):" & vbCrLf & shown
        If total > lines Then
            ApplyControlledValidation = ApplyControlledValidation & _
                "  ... and " & (total - lines) & " more -- see the register." & vbCrLf
        End If
    End If
End Function

' Case-SENSITIVE on purpose. "Not started" against a vocabulary of "Not Started"
' is precisely the drift this exists to surface -- matching case-insensitively
' here would make the checker agree with the thing it was written to catch.
Private Function InVocabulary(value As String, allowed As String) As Boolean
    Dim parts As Variant
    parts = Split(allowed, ",")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If Trim(CStr(parts(i))) = Trim(value) Then
            InVocabulary = True
            Exit Function
        End If
    Next i
End Function

' PICKED, NEVER TYPED -- the same rule as periods and allowed values, and for
' the same reason: a behaviour matched by string that nobody can spell wrong is
' worth more than one that reads well in a comment.
'
' FAILS LOUD, like Sources.ApplyPeriodValidation. A dropdown that silently did
' not apply would read as care taken and stop anyone re-checking.
Public Function ApplyBehaviourValidation(ws As Object) As String
    If ws Is Nothing Then
        ApplyBehaviourValidation = "Behaviour validation: skipped (no Field Spec sheet)."
        Exit Function
    End If

    Dim listText As String
    listText = BEHAVIOUR_FILL & "," & BEHAVIOUR_FIT & "," & BEHAVIOUR_ASIS

    Dim lastRow As Long
    lastRow = SPEC_FIRST_ROW
    Do While Trim(CStr(ws.Cells(lastRow, COL_SPEC_FIELDID).Value)) <> ""
        lastRow = lastRow + 1
    Loop
    If lastRow <= SPEC_FIRST_ROW Then
        ApplyBehaviourValidation = "Behaviour validation: no field rows yet."
        Exit Function
    End If

    Dim rng As Object
    Set rng = ws.Range(ws.Cells(SPEC_FIRST_ROW, COL_SPEC_BEHAVIOUR), ws.Cells(lastRow - 1, COL_SPEC_BEHAVIOUR))

    On Error Resume Next
    Err.Clear
    rng.Validation.Delete
    rng.Validation.Add 3, 1, 1, listText      ' xlValidateList, xlValidAlertStop, xlBetween
    If Err.Number <> 0 Then
        Dim e As String
        e = Err.Description
        On Error GoTo 0
        ApplyBehaviourValidation = "Behaviour validation NOT APPLIED (" & e & _
            ") -- the column still works, it just will not offer the list."
        Exit Function
    End If
    On Error GoTo 0

    ApplyBehaviourValidation = "Behaviour: list applied to " & (lastRow - SPEC_FIRST_ROW) & " field row(s)."
End Function

' PICKED, NEVER TYPED, same as Behaviour and Renders as. Fails loud for the
' same reason: a dropdown that silently did not apply reads as care taken.
Public Function ApplyHistoryValidation(ws As Object) As String
    If ws Is Nothing Then
        ApplyHistoryValidation = "History validation: skipped (no Field Spec sheet)."
        Exit Function
    End If

    Dim listText As String
    listText = HIST_CARRY & "," & HIST_FRESH & "," & HIST_PARTFROZEN & "," & HIST_DIFF

    Dim lastRow As Long
    lastRow = SPEC_FIRST_ROW
    Do While Trim(CStr(ws.Cells(lastRow, COL_SPEC_FIELDID).Value)) <> ""
        lastRow = lastRow + 1
    Loop
    If lastRow <= SPEC_FIRST_ROW Then
        ApplyHistoryValidation = "History validation: no field rows yet."
        Exit Function
    End If

    Dim rng As Object
    Set rng = ws.Range(ws.Cells(SPEC_FIRST_ROW, COL_SPEC_HISTORY), ws.Cells(lastRow - 1, COL_SPEC_HISTORY))

    On Error Resume Next
    Err.Clear
    rng.Validation.Delete
    rng.Validation.Add 3, 1, 1, listText      ' xlValidateList, xlValidAlertStop, xlBetween
    If Err.Number <> 0 Then
        Dim e As String
        e = Err.Description
        On Error GoTo 0
        ApplyHistoryValidation = "History validation NOT APPLIED (" & e & _
            ") -- the column still works, it just will not offer the list."
        Exit Function
    End If
    On Error GoTo 0

    ApplyHistoryValidation = "History: list applied to " & (lastRow - SPEC_FIRST_ROW) & " field row(s)."
End Function

' What a field is declared to render as, from the sheet.
'
' Blank is Text, and that is a decision rather than a fallback: every field that
' predates this column is text, so a blank meaning "unknown" would turn every
' existing workbook into a sheet full of warnings about nothing.
'
' An UNRECOGNISED value is a different matter and is REPORTED through `note`,
' never absorbed. Excel validation is a help to the person, not a guarantee to
' the code -- Sources.ApplyPeriodValidation already documents five paths where
' the dropdown does not get applied and the run carries on -- so a value this
' code does not know can and will arrive here.
Public Function RendersAsFor(specWs As Object, fieldId As String, Optional ByRef note As String) As String
    note = ""
    RendersAsFor = RENDER_TEXT
    If specWs Is Nothing Then Exit Function
    If Trim(fieldId) = "" Then Exit Function

    Dim want As String
    want = UCase(Trim(fieldId))

    Dim r As Long
    r = SPEC_FIRST_ROW
    Do While Trim(CStr(specWs.Cells(r, COL_SPEC_FIELDID).Value)) <> ""
        If UCase(Trim(CStr(specWs.Cells(r, COL_SPEC_FIELDID).Value))) = want Then
            Dim raw As String
            raw = Trim(CStr(specWs.Cells(r, COL_SPEC_RENDERS).Value))
            If raw = "" Then Exit Function

            Select Case LCase(raw)
                Case LCase(RENDER_TEXT):     RendersAsFor = RENDER_TEXT
                Case LCase(RENDER_PICTURE):  RendersAsFor = RENDER_PICTURE
                Case LCase(RENDER_PROGRESS): RendersAsFor = RENDER_PROGRESS
                Case LCase(RENDER_SLOTS):    RendersAsFor = RENDER_SLOTS
                Case Else
                    note = "'" & fieldId & "' has 'Renders as' = '" & raw & _
                        "', which is not one of " & RENDER_TEXT & " / " & RENDER_PICTURE & _
                        " / " & RENDER_PROGRESS & " / " & RENDER_SLOTS & ". Treated as " & RENDER_TEXT & "."
                    RendersAsFor = RENDER_TEXT
            End Select
            Exit Function
        End If
        r = r + 1
    Loop
End Function

' PICKED, NEVER TYPED, same as Behaviour and periods. Fails loud for the same
' reason: a dropdown that silently did not apply reads as care taken.
Public Function ApplyRendersValidation(ws As Object) As String
    If ws Is Nothing Then
        ApplyRendersValidation = "Renders-as validation: skipped (no Field Spec sheet)."
        Exit Function
    End If

    Dim listText As String
    listText = RENDER_TEXT & "," & RENDER_PICTURE & "," & RENDER_PROGRESS & "," & RENDER_SLOTS

    Dim lastRow As Long
    lastRow = SPEC_FIRST_ROW
    Do While Trim(CStr(ws.Cells(lastRow, COL_SPEC_FIELDID).Value)) <> ""
        lastRow = lastRow + 1
    Loop
    If lastRow <= SPEC_FIRST_ROW Then
        ApplyRendersValidation = "Renders-as validation: no field rows yet."
        Exit Function
    End If

    Dim rng As Object
    Set rng = ws.Range(ws.Cells(SPEC_FIRST_ROW, COL_SPEC_RENDERS), ws.Cells(lastRow - 1, COL_SPEC_RENDERS))

    On Error Resume Next
    Err.Clear
    rng.Validation.Delete
    rng.Validation.Add 3, 1, 1, listText      ' xlValidateList, xlValidAlertStop, xlBetween
    If Err.Number <> 0 Then
        Dim e As String
        e = Err.Description
        On Error GoTo 0
        ApplyRendersValidation = "Renders-as validation NOT APPLIED (" & e & _
            ") -- the column still works, it just will not offer the list."
        Exit Function
    End If
    On Error GoTo 0

    ApplyRendersValidation = "Renders as: list applied to " & (lastRow - SPEC_FIRST_ROW) & " field row(s)."
End Function
