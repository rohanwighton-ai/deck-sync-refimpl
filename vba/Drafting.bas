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
Public Const COL_D_PREV As Long = 4         ' D  REPORTED LAST TIME -- see below
Public Const COL_D_SOURCES As Long = 5      ' E  what you drafted FROM
Public Const COL_D_DRAFT As Long = 6        ' F  AI draft, never published
Public Const COL_D_SUBMIT As Long = 7       ' G  your text -- this publishes
Public Const COL_D_APPROVED As Long = 8     ' H  the tick
Public Const COL_D_CHARS As Long = 9
Public Const COL_D_SUBCHARS As Long = 10
Public Const COL_D_NOTES As Long = 11      ' K  notes back to the tool
Public Const COL_D_LAYOUT As Long = 12
Public Const COL_D_PROMPT As Long = 13
Public Const COL_D_PERIOD As Long = 14

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
' 5 adds COL_D_PREV -- "REPORTED LAST TIME". Rohan, 2026-08-14: "whenever a 1/4
' changes at the top, the ferries belonging to that system deal with information
' for the new quarter. The last previous set that runs move info into 'reported
' last time' column."
'
' So the outgoing quarter's SUBMIT is carried SIDEWAYS at the moment the update
' notices the period changed, rather than being destroyed and looked up again
' from somewhere else. That is what makes the rollover safe: nothing has to be
' preserved across a gap, because no gap is opened.
Public Const DRAFT_LAYOUT_VERSION As Long = 5

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
' REBUILDING PRESERVES EVERYTHING A HUMAN TYPED -- draft, submitted text, source
' IDs, notes AND the tick. A rebuild happens when entities are added or the
' exemplar moves on, and losing half-written drafts to a refresh would make the
' sheet untrustworthy in exactly the way the modal prompts it replaces were.
'
' THE TICK USED TO BE THE ONE EXCEPTION, and it was wrong. The reason given was
' that "an approval is against a specific pairing of exemplar and draft, and a
' rebuild MAY have changed the exemplar" -- true, and it does not license an
' unconditional clear. Three things settle it (Rohan, 2026-08-14, asking why
' these sheets are rebuilt at all):
'
'   1. THE FREQUENCY WAS THE DEFECT, NOT THE CLEAR. There is exactly one moment
'      the work columns must reset -- a period change, so last quarter's text
'      cannot be republished as this quarter's. This function ALREADY computes
'      that moment (`periodChanged`, and per row `carryThisRow`) for the other
'      four columns. The tick was cleared outside that decision, so the code
'      worked out the answer and then ignored it for one column.
'
'   2. THE TICK APPROVES SUBMIT, NOT THE EXEMPLAR. It is carried under the same
'      `carryThisRow` test as SUBMIT and therefore cannot outlive the words it
'      was given for. If ORIGINAL has moved underneath it, the person's own
'      wording is still their own wording -- and that difference is shown to
'      them again, per row, at the review grid.
'
'   3. THE SLIDE'S GATE IS SOMEWHERE ELSE ENTIRELY. This tick governs sheet ->
'      REGISTER. Nothing reaches a slide on the strength of it: ReviewQueue
'      gates register -> slide with a per-row ChangeHash, revalidated at apply,
'      which is the proper instrument for "the before-and-after I approved no
'      longer exists". Wiping every tick here bought no safety the review grid
'      was not already providing, and it cost the publish path entirely.
'
' What it cost: publish reads the tick (PublishDraftsForField, step 4 of the
' Sync Now chain) and RefreshDraftingSheets rebuilds the sheet (step 3). A tick
' could never survive to be read, so no drafted text has ever reached the
' register through the tool. The 43 PROGRESS_BODY values of 2026-08-14 were put
' in by hand over COM for this reason.
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
                Case "PREV":     ColumnInLayout = COL_D_PREV
                Case "SOURCES":  ColumnInLayout = COL_D_SOURCES
                Case "DRAFT":    ColumnInLayout = COL_D_DRAFT
                Case "SUBMIT":   ColumnInLayout = COL_D_SUBMIT
                Case "APPROVED": ColumnInLayout = COL_D_APPROVED
                Case "NOTES":    ColumnInLayout = COL_D_NOTES
            End Select

        ' Layout 4: as layout 5 but with no REPORTED LAST TIME column, so
        ' everything from SOURCES rightwards sat one column to the left.
        ' "PREV" is absent here on purpose -- it returns 0 and the migration
        ' writes the new column empty, which is correct for a sheet that has
        ' never seen a quarter turn.
        Case 4
            Select Case which
                Case "CURRENT":  ColumnInLayout = 3
                Case "SOURCES":  ColumnInLayout = 4
                Case "DRAFT":    ColumnInLayout = 5
                Case "SUBMIT":   ColumnInLayout = 6
                Case "APPROVED": ColumnInLayout = 7
                Case "NOTES":    ColumnInLayout = 10
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


' WHERE EACH LAYOUT KEEPS ITS OWN STAMPS.
'
' Needed because the stamps moved when the columns did, and the code that reads
' them cannot use the current layout's positions -- that is how every drafting
' sheet was wiped on 2026-08-14. Historical positions are literals because they
' are historical FACTS: layout 3 and 4 are frozen and can never change again.
' The current one is derived, so a future bump moves it in one place.
Public Function LayoutStampColumn(ByVal layoutVersion As Long) As Long
    Select Case layoutVersion
        Case DRAFT_LAYOUT_VERSION: LayoutStampColumn = COL_D_LAYOUT
        Case Else:                 LayoutStampColumn = 11      ' layouts 3 and 4
    End Select
End Function

Public Function PeriodStampColumn(ByVal layoutVersion As Long) As Long
    Select Case layoutVersion
        Case DRAFT_LAYOUT_VERSION: PeriodStampColumn = COL_D_PERIOD
        Case Else:                 PeriodStampColumn = 13      ' layouts 3 and 4
    End Select
End Function


' THE LAYOUT VERSION, FOUND BY LOOKING FOR IT.
'
' The pure half, so the rule can be tested without a workbook. `vals` is the
' intro row's values across the stamp-bearing columns, in ascending column order.
'
' SEARCHED RIGHT TO LEFT, and that is not arbitrary: a sheet part-way through a
' bump can carry BOTH a stale old stamp and the new one, and the newer position
' is always further right. Left to right would find the stale 4 and migrate a
' layout-5 sheet as though it were layout 4 -- relabelling every column, which
' is the failure that looks like content rather than like loss.
'
' A period ("Q4F26") and a prompt are not numeric, so only a stamp can match.
' Empty is numeric in VBA but Val("") is 0, and 0 is below the floor.
Public Function DetectLayoutFromRow(vals As Variant) As Long
    Dim i As Long, n As Double
    For i = UBound(vals) To LBound(vals) Step -1
        If IsNumeric(vals(i)) Then
            n = Val(CStr(vals(i)))
            If n >= 1 And n <= DRAFT_LAYOUT_VERSION And n = Int(n) Then
                DetectLayoutFromRow = CLng(n)
                Exit Function
            End If
        End If
    Next i
End Function

Public Function DetectSheetLayout(ws As Object) As Long
    Dim vals(1 To 7) As Variant
    Dim c As Long
    On Error Resume Next
    For c = 10 To 16
        vals(c - 9) = ws.Cells(DRAFT_INTRO_ROW, c).Value
    Next c
    On Error GoTo 0
    DetectSheetLayout = DetectLayoutFromRow(vals)
End Function


' MIGRATE A KNOWN OLDER LAYOUT INTO THE CURRENT COLUMNS, IN PLACE.
'
' This is the job the five kept* carry dictionaries used to do as a SIDE EFFECT
' of clearing the sheet and rebuilding it. Deleting them (2026-08-14) deleted
' the migration with them, and nothing noticed until the suite ran: a layout-3
' sheet was left physically untouched and then READ with layout-4 numbers, so a
' person's SUBMIT text was read as a source ID and the AI draft landed in the
' column that publishes. The sheet was then stamped forward, so it could never
' self-correct. Worse than losing the work, because it looks like content.
'
' In place and per row, with every value read BEFORE any value is written: the
' old and new positions overlap -- 3 -> 4 is a permutation of the same six
' columns -- so writing as it goes would overwrite a cell still needed.
'
' Every old position is emptied between the read and the write, so a future bump
' that REMOVES a column cannot leave a stale value stranded under a new heading.
Private Sub MigrateSheetLayout(ws As Object, fromLayout As Long)
    Dim names As Variant
    names = Array("CURRENT", "PREV", "SOURCES", "DRAFT", "SUBMIT", "APPROVED", "NOTES")

    Dim vals() As Variant
    ReDim vals(0 To UBound(names))

    Dim r As Long, i As Long, oldCol As Long, newCol As Long
    r = DRAFT_FIRST_ROW

    Do While Trim(CStr(ws.Cells(r, COL_D_ENTITY).Value)) <> ""
        For i = 0 To UBound(names)
            vals(i) = Empty
            oldCol = ColumnInLayout(fromLayout, CStr(names(i)))
            If oldCol > 0 Then vals(i) = ws.Cells(r, oldCol).Value
        Next i

        For i = 0 To UBound(names)
            oldCol = ColumnInLayout(fromLayout, CStr(names(i)))
            If oldCol > 0 Then ws.Cells(r, oldCol).ClearContents
        Next i

        For i = 0 To UBound(names)
            newCol = ColumnInLayout(DRAFT_LAYOUT_VERSION, CStr(names(i)))
            If newCol > 0 Then ws.Cells(r, newCol).Value = vals(i)
        Next i

        r = r + 1
    Loop
End Sub


' NOTHING IS DESTROYED, ONLY SUPERSEDED.
'
' Rohan, 2026-08-10: "we can't have drafting risk losing information."
'
' Rebuilding a drafting sheet discards whatever it cannot carry across -- on a
' layout mismatch that is EVERYTHING, and on a period change it is every
' quarterly row. Both were announced but still lost, and announcing a loss is
' not preventing one.
'
' So the sheet is COPIED ASIDE before it is rebuilt. Migration is hard and can
' fail silently; a copy cannot. The old sheet sits in the workbook to be read,
' compared, or copied out of by hand.
'
' Only when there is something to lose: an empty sheet, or one whose work all
' carries across, is rebuilt in place as before. A copy per run regardless would
' bury the workbook in tabs and train someone to ignore them.
'
' Excel caps a sheet name at 31 characters, so the field id is truncated -- the
' sheet's own contents identify it, and the timestamp makes it findable.

' What became of the work that could not be carried across.
'
' THE OLD WORDING WAS "gone -- this workbook is not backed up", which was true
' and is now false, and a message that overstates a loss trains someone to
' ignore the one that does not. A FAILED copy says so plainly instead of
' inheriting the reassuring half of the sentence.
Private Function ParkedNote(parkedName As String, parkFailed As Boolean) As String
    If parkFailed Then
        ParkedNote = " COULD NOT SAVE A COPY of the previous sheet, so unpublished drafts, " & _
            "source IDs and notes on it are gone. Anything already published is safe in " & _
            "the register."
    ElseIf parkedName <> "" Then
        ParkedNote = " The previous sheet was kept as '" & parkedName & "' -- nothing was lost. " & _
            "Delete that tab once you have taken what you need from it."
    Else
        ParkedNote = " Anything already published is safe in the register."
    End If
End Function


' KEEPS THE TWO MOST RECENT ARCHIVES PER FIELD.
'
' Unbounded archives are their own failure: a workbook with thirty hidden tabs
' is one nobody reads, and "nothing was lost" stops meaning anything when the
' thing it was lost into is unfindable. Two is enough to recover from a mistake
' and from the rebuild that followed it.
'
' Matched on the field id the name ends with, so archives of one field never
' prune another's.
Private Sub PruneParked(wb As Object, fieldId As String)
    On Error Resume Next
    Dim keepNewest As Long
    keepNewest = 2

    Dim names As Collection
    Set names = New Collection

    Dim sh As Object
    For Each sh In wb.Sheets
        If Left(sh.Name, 6) = "SAVED " Then
            If InStr(sh.Name, Left(fieldId, 12)) > 0 Then names.Add sh.Name
        End If
    Next sh

    ' The timestamp is mmdd-hhnn inside the name, so alphabetical order IS
    ' chronological order within a year -- no date parsing needed.
    Do While names.Count > keepNewest
        Dim oldest As String, i As Long
        oldest = names(1)
        For i = 2 To names.Count
            If names(i) < oldest Then oldest = names(i)
        Next i
        ' FIX-LIST item AS, 2026-08-17/18 night. Same defect, same fix, as
        ' DraftingLobby.bas's own Lobby-sheet rebuild (see that file's
        ' comment on this exact pattern, 2026-08-16): a plain .Delete here
        ' raises Excel's native "permanently delete this sheet?" alert,
        ' unsuppressed, once per field on every period-change rollover --
        ' confirmed live by screenshot. wb.Application, NOT the bare
        ' Application -- this VBA project is hosted in PowerPoint, so a bare
        ' Application reference resolves to PowerPoint.Application, not the
        ' Excel instance that actually owns wb (the same trap
        ' DraftingLobby.bas already documents having hit once).
        wb.Application.DisplayAlerts = False
        wb.Sheets(oldest).Delete
        wb.Application.DisplayAlerts = True
        Dim j As Long
        For j = 1 To names.Count
            If names(j) = oldest Then
                names.Remove j
                Exit For
            End If
        Next j
    Loop
    On Error GoTo 0
End Sub

Private Function ParkSheetCopy(ws As Object, fieldId As String, ByRef parkedName As String) As Boolean
    parkedName = ""
    On Error GoTo Failed

    Dim stamp As String
    stamp = Format(Now, "mmdd-hhnn")

    Dim base As String
    base = "SAVED " & stamp & " " & fieldId
    If Len(base) > 31 Then base = Left(base, 31)

    ' A second rebuild in the same minute must not collide with the first.
    Dim candidate As String, n As Long
    candidate = base
    Do While WorkbookBridge.WorksheetExists(ws.Parent, candidate) And n < 20
        n = n + 1
        Dim suffix As String
        suffix = "~" & n
        candidate = Left(base, 31 - Len(suffix)) & suffix
    Loop

    ws.Copy After:=ws
    Dim copied As Object
    Set copied = ws.Parent.Sheets(ws.Index + 1)
    copied.Name = candidate

    ' NOT A WORKSPACE, AND SAID SO ON THE SHEET ITSELF. A copy that looks like a
    ' drafting sheet is a second place to type, and typing there is silent --
    ' publish reads the live sheet only. Rohan, 2026-08-10: "how do you stop it
    ' confusing point of truth?" By making it impossible to work in by accident,
    ' and by saying what it is to anyone who goes looking.
    copied.Cells(1, 1).Value = "ARCHIVE of " & fieldId & " taken " & _
        Format(Now, "d mmm yyyy hh:nn") & " -- THE TOOL NEVER READS THIS SHEET. " & _
        "Copy anything you still want into the live drafting sheet, then delete this tab."
    copied.Cells(1, 1).Font.Bold = True
    copied.Visible = 0                       ' xlSheetHidden

    PruneParked ws.Parent, fieldId

    parkedName = candidate
    ParkSheetCopy = True
    Exit Function

Failed:
    ' A COPY THAT FAILED MUST NOT LOOK LIKE ONE THAT WORKED. The caller decides
    ' what to do about it; silently carrying on would be the loss this exists to
    ' prevent, wearing a reassuring message.
    parkedName = ""
    ParkSheetCopy = False
End Function

Public Function WriteDraftingSheet(ws As Object, reg As Sheet, fieldId As String, _
                                   Optional guidance As Variant, _
                                   Optional periodStamp As String = "", _
                                   Optional srcWs As Object = Nothing, _
                                   Optional familySeed As Long = -1) As String
    ' A REBUILD MUST NOT COST A PERSON THEIR WORK. Everything a human or an AI
    ' put on this sheet is carried across: the AI draft, the SUBMIT text they
    ' edited, the source IDs they assigned, and their notes. Only ORIGINAL and
    ' the character counts are re-derived, because only those come from the
    ' register. Losing a column here would be silent and would cost an evening.
    ' NOTHING IS CARRIED, BECAUSE NOTHING IS DESTROYED.
    '
    ' Five Scripting.Dictionaries used to be declared here -- keptDraft,
    ' keptNotes, keptSubmit, keptSources, keptApproved -- to hold a copy of every
    ' human column while ws.Cells.Clear wiped the sheet, then write them back.
    ' Their existence WAS the defect, not their contents: a rebuild that has to
    ' carry a person's work across a gap is a rebuild that opened the gap.
    '
    ' An existing project's row is now updated where it sits and its typed
    ' columns are not written at all.
    '
    ' THE INVARIANT WAS FIRST WRITTEN AS "if these dictionaries ever come back,
    ' so has the clear", AND THAT WAS WRONG. Rohan, 2026-08-14: "can't they just
    ' sate the need then run on update?" -- and they can. Holding values in a
    ' variable was never the defect; opening a gap was. MigrateSheetLayout is a
    ' carry, and the rollover ferry below is another, and both are correct
    ' precisely because they run ON the update instead of around a destroy.
    '
    ' The real invariant, and the only one worth guarding: NO ws.Cells.Clear ON
    ' THE NORMAL PATH. Naming a mechanism as the proxy for a hazard forbids
    ' useful code and still misses the hazard when it arrives by another route.

    ' Only carry work across if the sheet was written by THIS layout. A sheet
    ' from an older layout is read with column numbers that have since been
    ' renumbered, which silently relabels a person's data rather than losing
    ' it -- the worse of the two failures, because it looks like content.
    Dim sheetLayout As Long
    Dim layoutMatches As Boolean
    ' FOUND, NOT ASSUMED. This used to read COL_D_LAYOUT -- the CURRENT layout's
    ' stamp column -- which is a bootstrap error: you need the layout to know
    ' where the layout stamp is. On a layout-4 sheet it read column 12, which
    ' holds the PROMPT there, got 0, and 0 is an unknown layout: layoutMatches
    ' went False, ws.Cells.Clear fired on the whole sheet, and the migration
    ' block below -- guarded by `If layoutMatches` -- never ran at all.
    '
    ' Cost, 2026-08-14 19:11: every drafting sheet wiped on a real workbook.
    ' 129 drafted paragraphs, 43 approve ticks and 75 notes destroyed across
    ' KEY_EVENTS_BODY, PROGRESS_BODY and HIGHLIGHTS_BODY. Recovered from a
    ' backup and from the parked copies, which is the only reason this reads as
    ' a defect rather than as a fortnight of lost evenings.
    '
    ' The comment 60 lines below ALREADY KNEW the columns moved -- "layout 4 kept
    ' the version in column 11 and the period in 13; layout 5 uses 12 and 14" --
    ' and used that knowledge only to WIPE the intro rows, never to READ them.
    sheetLayout = DetectSheetLayout(ws)
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
    ' THE PER-ROW CADENCE MACHINERY IS GONE, AND THE QUESTION IT ANSWERED NO
    ' LONGER EXISTS.
    '
    ' It asked, per row, whether a value was quarterly (drop it on a rollover) or
    ' entity-static (keep it), because for a static row LAST QUARTER'S TEXT IS
    ' THIS QUARTER'S TEXT and dropping it was pure loss. That was a real problem
    ' while a rollover DESTROYED. It stops being one when the rollover FERRIES:
    ' the text is one column away, in plain sight, for every row alike, and the
    ' person can copy it back in a keystroke if it still stands.
    '
    ' So a distinction that needed a register lookup, an extra argument threaded
    ' through every caller, and a fallback rule for "nobody can say" is replaced
    ' by moving the text where the person can see it. That is the trade worth
    ' remembering: the machinery existed to decide something on a person's
    ' behalf, and showing them instead deleted the decision.
    Dim sheetPeriod As String
    ' READ AT THE SHEET'S OWN LAYOUT'S POSITION, not the current one. Same
    ' bootstrap error as the layout stamp above and a worse consequence: a
    ' layout-4 sheet's period sits in column 13, reading column 14 returned "",
    ' and a blank period reads as A QUARTER TURN -- which clears every work
    ' column, per row, by design. Two independent routes to the same wipe.
    On Error Resume Next
    sheetPeriod = Trim(CStr(ws.Cells(DRAFT_INTRO_ROW, PeriodStampColumn(sheetLayout)).Value))
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

    ' ============================================================
    ' THE ROLLOVER REFUSAL IS GONE. IT WAS A DEADLOCK.
    ' ============================================================
    '
    ' It refused a rollover whenever the sheet held typed work, and told the
    ' person to "publish this sheet's work before rebuilding". PublishDrafts only
    ' READS SUBMIT and APPROVED -- it never clears them -- and the only
    ' ClearContents for those columns sat INSIDE the branch the refusal had
    ' already exited past. Nothing else in the tool clears or deletes a drafting
    ' sheet. So once a sheet held typed work, which is the wanted state, it could
    ' never be rolled forward, and the instruction it gave could not be followed
    ' by any means. Landed 2026-08-13 in ddf867b; never hit, because no rollover
    ' was attempted between then and its removal.
    '
    ' The refusal existed to stop last quarter's prose being republished as this
    ' quarter's. The ferry below achieves that WITHOUT a refusal, because the text
    ' moves out of SUBMIT rather than being destroyed -- so there is no loss to
    ' refuse on behalf of. A guard whose only remedy is unreachable is not a
    ' guard; it is a stop.
    Dim ferried As Long
    ' WORK ABOUT TO BE DISCARDED BY A LAYOUT MISMATCH, counted so it can be
    ' SAID rather than silently lost. See the block below.
    Dim strandedRows As Long, strandedDetail As String

    Dim r As Long
    r = DRAFT_FIRST_ROW

    ' A LAYOUT MISMATCH USED TO DROP EVERYTHING WITHOUT SAYING SO.
    '
    ' Carry-over runs only `If layoutMatches`, and when it does, work follows its
    ' PROJECT KEY -- so reordering or adding projects is safe. When it does not
    ' match, every draft, note, submission and citation on the sheet was simply
    ' gone: no refusal, no count, no line in the report.
    '
    ' Rohan, 2026-08-10: "I think the register and drafting sheets sound a bit
    ' fragile?" The durable sheets are guarded and the disposable ones are
    ' rebuilt, which is the right shape -- the fragility was all in the
    ' transitions, and this was the one that failed QUIETLY. Counted here,
    ' reported below, and still discarded: the columns genuinely cannot be
    ' located on a sheet whose layout is unknown, so the honest act is to say
    ' what is being lost, not to pretend it can be saved.
    If Not layoutMatches Then
        Dim sr As Long
        sr = DRAFT_FIRST_ROW
        On Error Resume Next
        Do While Trim(CStr(ws.Cells(sr, COL_D_ENTITY).Value)) <> ""
            Dim anyText As Boolean
            anyText = False
            Dim c As Long
            For c = 3 To 12
                If Trim(CStr(ws.Cells(sr, c).Value)) <> "" Then anyText = True
            Next c
            If anyText Then
                strandedRows = strandedRows + 1
                If strandedRows <= 5 Then
                    If strandedDetail <> "" Then strandedDetail = strandedDetail & ", "
                    strandedDetail = strandedDetail & Trim(CStr(ws.Cells(sr, COL_D_ENTITY).Value))
                End If
            End If
            sr = sr + 1
        Loop
        On Error GoTo 0
    End If

    ' PARK A COPY BEFORE ANYTHING IS CLEARED. Two cases lose work: a layout
    ' mismatch loses everything, and a period change clears the quarterly rows.
    ' Both are now preserved rather than merely announced.
    Dim parkedName As String, parkFailed As Boolean
    If strandedRows > 0 Then
        If Not ParkSheetCopy(ws, fieldId, parkedName) Then parkFailed = True
    End If

    ' THE FIVE kept* CARRY DICTIONARIES USED TO BE FILLED HERE, AND THEY ARE GONE.
    '
    ' This block read every row's DRAFT, SUBMIT, SOURCES, NOTES and APPROVED into
    ' memory so they could be written back after ws.Cells.Clear destroyed them.
    ' It existed for one reason: to survive the clear. With the clear confined to
    ' a layout migration, there is nothing to survive -- an existing row is now
    ' updated where it sits and its typed columns are never written at all.
    '
    ' Deleted rather than left unreachable. A harvest that still runs is a second
    ' copy of a person's work held in memory during the write, and the whole
    ' lesson of 2026-08-14 is that the copy is the risk. It also double-counted
    ' the rollover, which the row loop's ferry now owns outright.

    ' PARKING IS STILL CONDITIONAL, NOT UNCONDITIONAL -- CORRECTED 2026-08-17.
    '
    ' This comment used to claim parking ran "unconditional on 'there was a
    ' sheet here'". It doesn't: parking below is gated on `Not layoutMatches`,
    ' on `layoutMatches And sheetLayout <> DRAFT_LAYOUT_VERSION`, and on
    ' `periodChanged` -- three real conditions, all further down this
    ' function. The ORDINARY case -- same layout, same period, the path every
    ' routine "1. Set up my quarter" press actually takes -- has NO backup
    ' taken before it writes, same as when this comment was first written.
    '
    ' Cost, 2026-08-14: a rebuild left TPL_KEY_EVENTS_BODY with 23 of 43 rows and
    ' TPL_PROGRESS_BODY with 37 of 43, losing 27 drafted paragraphs, and there was
    ' no archive to recover them from -- the layout matched, so nothing parked.
    ' Recovered only because a whole-file backup happened to exist.
    '
    ' The in-place rewrite (this comment's own neighbour, "THE CLEAR IS A
    ' MIGRATION TOOL") already shrank the blast radius for this exact case --
    ' a mid-write failure now leaves unreached rows untouched rather than
    ' destroyed, which is a real, different mitigation for the same scenario.
    ' Whether the ordinary path should ALSO get a cheap pre-write backup is
    ' still open -- not decided here, not implemented here. Found stale while
    ' reading this function for FIX-LIST item AD's speed work
    ' (DRAFTING-SPEED-STRATEGY.md); fixing what the comment claims, not what
    ' the code does, matches this project's "a description of a machine fact
    ' goes stale, so don't leave a wrong one in place" rule.
    ' ============================================================
    ' THE CLEAR IS A MIGRATION TOOL. IT IS NOT THE NORMAL PATH.
    ' ============================================================
    '
    ' Destroying the sheet is defensible in exactly ONE situation: the columns
    ' have been renumbered, so the old cells cannot be located and there is no
    ' such thing as reading them safely. That has happened three times in this
    ' tool's life (DRAFT_LAYOUT_VERSION is 4).
    '
    ' It was previously the path taken EVERY time, for every reason -- a new
    ' project joining mid-quarter destroyed all 43 rows to add 1. Rohan asked
    ' why on 1, 9, 10, 13 and 14 August; each asking got a patch that made the
    ' destruction more survivable (carry the tick, park a copy, refuse on a
    ' rollover) and none removed it. On 2026-08-14 a mid-write failure left two
    ' sheets holding 23 and 37 of 43 rows, and 27 drafted paragraphs were gone.
    '
    ' In place, that same failure leaves the rows it has not reached UNTOUCHED.
    ' The blast radius stops being the whole sheet.
    If Not layoutMatches Then
        If parkedName = "" And Not isNewSheet Then
            If Not ParkSheetCopy(ws, fieldId, parkedName) Then parkFailed = True
        End If
        ws.Cells.Clear
    End If

    ' A KNOWN OLDER LAYOUT IS MIGRATED, NOT ABANDONED AND NOT MISREAD.
    '
    ' Must run BEFORE the row loop touches anything through the CURRENT column
    ' numbers, and before the stamp below declares the sheet already current.
    ' Parked first for the same reason every other rewrite is: this moves a
    ' person's typed work between columns, and a copy costs nothing.
    If layoutMatches And sheetLayout <> DRAFT_LAYOUT_VERSION And Not isNewSheet Then
        If parkedName = "" Then
            If Not ParkSheetCopy(ws, fieldId, parkedName) Then parkFailed = True
        End If
        MigrateSheetLayout ws, sheetLayout
    End If

    ' PARK BEFORE A QUARTER TURN AS WELL. The ferry preserves SUBMIT by moving
    ' it, but the AI draft, the source citations and the notes ARE cleared -- and
    ' citations are the slowest column on the sheet to rebuild, being the control
    ' on the generative step rather than decoration. A copy costs nothing.
    If periodChanged And parkedName = "" Then
        If Not ParkSheetCopy(ws, fieldId, parkedName) Then parkFailed = True
    End If

    ' STAMPS FIRST, BEFORE ANY ROW. They used to be written at the END of this
    ' function, so a failure part-way through the row loop left a sheet with rows
    ' and NO layout version and NO period -- which reads on the next rebuild as
    ' "unknown layout", and an unknown layout carries nothing across. A crash
    ' mid-write therefore armed the NEXT run to discard every surviving draft.
    '
    ' Written first, the same crash leaves a sheet that correctly declares what
    ' it is, so the next rebuild reads it, carries the partial work, and repairs
    ' it. Found 2026-08-14 on two sheets that had rows, no stamp and no prompt.
    ' THE INTRO ROWS ARE WIPED BEFORE THEY ARE REWRITTEN, so a layout bump cannot
    ' strand an old stamp under a new heading. Layout 4 kept the version in
    ' column 11 and the period in 13; layout 5 uses 12 and 14, so without this
    ' the old value would sit there forever, under NOTES, looking like data.
    ' Everything in rows 1-2 from column 2 rightwards is tool-written and is
    ' rewritten immediately below -- both stamps are already read into
    ' sheetPeriod and sheetLayout well before this point.
    ws.Range(ws.Cells(DRAFT_INTRO_ROW, 2), ws.Cells(DRAFT_INTRO_ROW + 1, 20)).ClearContents

    ws.Cells(DRAFT_INTRO_ROW, COL_D_PERIOD).Value = periodStamp
    ws.Cells(DRAFT_INTRO_ROW, COL_D_LAYOUT).Value = DRAFT_LAYOUT_VERSION

    ' --- instructions, ON the sheet, for the person ----------------------
    ws.Cells(DRAFT_INTRO_ROW, 1).Value = "WHAT TO DO ON THIS SHEET  --  " & fieldId
    ws.Cells(DRAFT_INTRO_ROW, 1).Font.Bold = True
    ws.Cells(DRAFT_INTRO_ROW, 1).Font.Size = 9

    ws.Cells(2, 1).Value = "STEP 1   Read column " & Chr$(64 + COL_D_CURRENT) & " -- what the slide says today. Column " & _
                           Chr$(64 + COL_D_PREV) & " is what was reported last quarter: match its voice, do not repeat it."
    ' COLUMN G, NOT E. Step 5 below sends the tick to E, so this line named one
    ' column for two things inside a single instruction block -- and E is the tick,
    ' which is the consent gate. Stale since 3de4be8 moved SUBMIT to D and the tick
    ' to E; that commit updated the header row and the toolbar tooltip and left
    ' every prose instruction pointing at the old layout.
    ws.Cells(3, 1).Value = "STEP 2   Name your sources in column " & Chr$(64 + COL_D_SOURCES) & ". Add new ones on the Sources sheet first."
    ws.Cells(4, 1).Value = "STEP 3   Ask Copilot for a draft -- the prompt is in cell L2. It writes into column " & Chr$(64 + COL_D_DRAFT) & ". That column is NEVER published."
    ws.Cells(5, 1).Value = "STEP 4   Press '" & CommandBarUI.CAP_SET_UP_QUARTER & "' to copy " & Chr$(64 + COL_D_DRAFT) & " into " & Chr$(64 + COL_D_SUBMIT) & ", then EDIT column " & Chr$(64 + COL_D_SUBMIT) & " until you are happy. That is what gets sent."
    ws.Cells(6, 1).Value = "STEP 5   Type  Y  in column " & Chr$(64 + COL_D_APPROVED) & ", save and CLOSE the file, then press '" & CommandBarUI.CAP_PUT_ON_SLIDES & "' again."

    ws.Cells(7, 1).Value = "Only column " & Chr$(64 + COL_D_SUBMIT) & " is published -- nothing the AI writes reaches a slide unless you have moved it there and ticked it. " & _
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
    ws.Cells(DRAFT_HEADER_ROW, COL_D_CURRENT).Value = _
        Chr$(64 + COL_D_CURRENT) & "   ORIGINAL -- what the slide says now (read-only)"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_PREV).Value = _
        Chr$(64 + COL_D_PREV) & "   REPORTED LAST TIME -- for style and continuity (read-only)"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_CHARS).Value = "Chars"
    ' EVERY COLUMN LETTER IS DERIVED, NEVER TYPED. These were literals -- "D
    ' SOURCES", "E AI DRAFT", "J NOTES" -- and adding one column at position 4
    ' made all six of them name the wrong column while still reading as correct.
    ' A heading that lies about where to type is worse than no heading.
    ws.Cells(DRAFT_HEADER_ROW, COL_D_SOURCES).Value = _
        Chr$(64 + COL_D_SOURCES) & "   SOURCES -- IDs from the Sources sheet, e.g. S01,S03"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_DRAFT).Value = _
        Chr$(64 + COL_D_DRAFT) & "   AI DRAFT -- Copilot writes here. NEVER published"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_SUBMIT).Value = _
        Chr$(64 + COL_D_SUBMIT) & "   SUBMIT -- your words. THIS is what gets sent"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_SUBCHARS).Value = "Chars"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_APPROVED).Value = _
        Chr$(64 + COL_D_APPROVED) & "   APPROVE -- type Y"
    ws.Cells(DRAFT_HEADER_ROW, COL_D_NOTES).Value = _
        Chr$(64 + COL_D_NOTES) & "   NOTES -- back to the tool (optional)"
    ws.Rows(DRAFT_HEADER_ROW).Font.Bold = True
    ws.Rows(DRAFT_HEADER_ROW).WrapText = True

    Dim written As Long, restored As Long

    ' A PROJECT'S ROW IS FOUND BY ITS CODE, NOT BY ITS POSITION.
    '
    ' Writing by position is what made a clear necessary in the first place: if
    ' row 10 is "whatever the register lists first", then every row has to be
    ' rewritten whenever the register's order or membership changes, and
    ' rewriting every row means first destroying every row. Indexing by project
    ' code breaks that chain -- an existing project keeps its row and its typed
    ' work is never touched; a new one is appended after the last.
    Dim rowOf As Object
    Set rowOf = CreateObject("Scripting.Dictionary")
    Dim appendAt As Long
    appendAt = DRAFT_FIRST_ROW
    If layoutMatches Then
        Do While Trim(CStr(ws.Cells(appendAt, COL_D_ENTITY).Value)) <> ""
            rowOf(Trim(CStr(ws.Cells(appendAt, COL_D_ENTITY).Value))) = appendAt
            appendAt = appendAt + 1
        Loop
    End If
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

            ' KEEP THE ROW THIS PROJECT ALREADY HAS. A new project takes the next
            ' free row at the bottom, which is the only write that moves anything.
            Dim isNewRow As Boolean
            If layoutMatches And rowOf.Exists(key) Then
                r = CLng(rowOf(key))
                isNewRow = False
            Else
                r = appendAt
                appendAt = appendAt + 1
                isNewRow = True
            End If

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
            ' ================================================================
            ' THE PERSON'S COLUMNS ARE NOT WRITTEN HERE. THAT IS THE FIX.
            ' ================================================================
            '
            ' Sources, AI draft, SUBMIT, the tick and notes are left exactly as
            ' the person left them. There is no restore step because there was no
            ' destruction to recover from -- which is why the five kept* carry
            ' dictionaries are now dead on this path. They only ever existed to
            ' ferry work across ws.Cells.Clear.
            '
            ' A rollover is the ONE time these must go: text written for last
            ' quarter must not be republishable as this quarter's. That decision
            ' is per ROW and the register already answers it -- exactly the same
            ' rule the old carry used, applied to a clear instead of a copy.
            If Not isNewRow Then
                restored = restored + 1

                ' THE FERRY. The quarter turned, so the last thing the outgoing
                ' quarter's pass does is carry its SUBMIT one column sideways
                ' into REPORTED LAST TIME, and hand the working columns to the
                ' new quarter empty.
                '
                ' The text is not looked up from anywhere afterwards, which is
                ' what makes this cheap: the sheet already holds it, and the
                ' update already knows the period changed. Nothing has to know
                ' which quarter precedes which -- and quarter labels are free
                ' text with no ordering, so nothing could have known.
                '
                ' It survives unpublished work too. A draft that never reached a
                ' slide is still what this project said last time, which is
                ' exactly the style-and-narrative material the next draft wants.
                If periodChanged Then
                    ws.Cells(r, COL_D_PREV).Value = "'" & _
                        Replace(CStr(ws.Cells(r, COL_D_SUBMIT).Value), "||", vbLf)

                    ws.Cells(r, COL_D_DRAFT).ClearContents
                    ws.Cells(r, COL_D_SUBMIT).ClearContents
                    ws.Cells(r, COL_D_SUBCHARS).ClearContents
                    ws.Cells(r, COL_D_APPROVED).ClearContents
                    ws.Cells(r, COL_D_SOURCES).ClearContents
                    ws.Cells(r, COL_D_NOTES).ClearContents
                    ferried = ferried + 1
                End If
            End If
            written = written + 1
        End If
    Next k

    ' `r` IS NOW A CURSOR INTO THE GRID, NOT A COUNTER. It ends holding whichever
    ' row the LAST register entry happened to occupy, which is no longer the
    ' bottom of the sheet. Everything below reads it as "one past the last row"
    ' -- row heights, and ApplyDraftingLook's banding extent. Handing it the
    ' cursor would stripe and size an arbitrary prefix of the grid.
    r = appendAt

    ' 8pt throughout. At 11pt three text columns of 350+ characters do not fit
    ' on a screen together, and the whole point of ORIGINAL / AI / SUBMIT is
    ' reading them side by side.
    ws.Cells.Font.Size = 8
    ws.Cells(DRAFT_INTRO_ROW, 1).Font.Size = 9

    ws.Columns(COL_D_ENTITY).ColumnWidth = 11
    ws.Columns(COL_D_NAME).ColumnWidth = 30
    ws.Columns(COL_D_CURRENT).ColumnWidth = 52
    ws.Columns(COL_D_PREV).ColumnWidth = 52
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

    ' The look goes on LAST, over the top of the per-column settings above, so
    ' the sheet reads as one thing rather than as the sum of its edits.
    ApplyDraftingLook ws, r - 1, fieldId, familySeed

    ' VALUES ARE WRITTEN IMMEDIATELY AFTER THE CLEAR, not here -- see the comment
    ' there. Only the cosmetics are left at this end, because they are the part
    ' it is safe to lose to a mid-write failure.
    ws.Cells(DRAFT_INTRO_ROW, COL_D_PERIOD).Font.Color = RGB(190, 190, 190)
    ws.Columns(COL_D_PERIOD).ColumnWidth = 9
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
        ' FIX-LIST 1d, REMOVED 2026-08-14 rather than repaired. This called
        ' ParkSheetCopy *here* -- hundreds of lines after ws.Cells.Clear -- so the
        ' copy it took was of the already-cleared, already-rebuilt sheet, and
        ' ParkedNote then reported "nothing was lost" over an archive that did not
        ' contain the thing it was taken to preserve. The park now runs before the
        ' clear, unconditionally, which makes this site both unreachable and wrong.

    WriteDraftingSheet = WriteDraftingSheet & vbCrLf & _
            "  Quarter turned " & IIf(sheetPeriod = "", "(no period recorded)", sheetPeriod) & _
            " -> " & periodStamp & ": " & ferried & " row(s) moved into '" & _
            Chr$(64 + COL_D_PREV) & " REPORTED LAST TIME' and cleared for redrafting." & _
            ParkedNote(parkedName, parkFailed)
    ElseIf Not layoutMatches And Not isNewSheet Then
        ' Same false claim as above, same fix -- see the comment there. Fixed in
        ' both places at once because this is one defect with two call sites, and
        ' fixing only where it was noticed is how the truncation bug came back
        ' four times.
        WriteDraftingSheet = WriteDraftingSheet & vbCrLf & _
            "  Rebuilt on a new sheet layout: nothing carried across." & _
            IIf(strandedRows > 0, " " & strandedRows & " row(s) had unpublished work" & _
                IIf(strandedDetail <> "", " (" & strandedDetail & _
                    IIf(strandedRows > 5, ", ...", "") & ")", "") & ".", "") & _
            ParkedNote(parkedName, parkFailed)
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
' THE SUMMARY LINE AND THE "IS THERE ANYTHING TO DO" TEST, BUILT FROM ONE PLACE.
'
' The publish stage stopped to ask "write these into the register?" over a
' preview reading all zeros, then reported it had written nothing, then saved.
' Skipping that needs a test for "nothing to do" -- and a test written
' separately from the sentence is a second copy of the same fact, which drifts.
' So the sentence is built here and the test reads the SAME counts.
Public Function PublishSummaryLine(published As Long, dryRun As Boolean, _
                                   skippedNoTick As Long, skippedEmpty As Long, _
                                   noRow As Long, failed As Long) As String
    PublishSummaryLine = "Summary: " & published & IIf(dryRun, " would be published", " published") & _
        ", " & skippedNoTick & " drafted but not ticked, " & skippedEmpty & " ticked but empty, " & _
        noRow & " with no register row, " & failed & " failed"
End Function

' True when a publish report found no work of any kind -- dry preview OR a
' real wet-mode result, both.
'
' Reads the report the person is shown, so the question asked and the numbers
' displayed cannot disagree -- if the summary ever says something happened, this
' will not claim otherwise.
'
' FIX-LIST item AF, real bug found writing PublishOneFieldForChain's own
' test tonight, not a hypothetical: this used to check ONLY against
' PublishSummaryLine(0, True, ...) -- the DRY-RUN phrasing ("0 would be
' published"). A wet-mode summary line says "0 published" (no "would be"),
' which that check could never match, so calling this on a real (non-dry)
' result would always report something-happened even when nothing did.
' Every caller before tonight happened to only ever pass this a dry
' preview, so the gap was never exercised -- until a genuine wet-only
' caller (PublishOneFieldForChain, replacing a dry-then-wet pair to halve
' the register-scan cost per field) needed it to work on real output too.
' Checked both phrasings rather than picking one, since existing callers
' still legitimately pass dry previews and must keep working exactly as
' before.
Public Function NothingToPublish(previewReport As String) As Boolean
    NothingToPublish = (InStr(previewReport, PublishSummaryLine(0, True, 0, 0, 0, 0)) > 0) Or _
                        (InStr(previewReport, PublishSummaryLine(0, False, 0, 0, 0, 0)) > 0)
End Function

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
                        "but SUBMIT is empty. Press '" & CommandBarUI.CAP_SET_UP_QUARTER & "' to copy it across, or type into D yourself." & vbCrLf
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

    report = report & vbCrLf & PublishSummaryLine(published, dryRun, skippedNoTick, _
        skippedEmpty, noRow, failed) & vbCrLf

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
        ' COLUMN LETTERS DERIVED, NEVER TYPED. Written as D/E/F when the layout
        ' put them there, and left saying so when layout 4 moved them -- so the
        ' message told a person to type into the sources column while the sheet
        ' beside it said the opposite. Seen on screen 2026-08-10.
        CopyAiToSubmit = "Nothing to copy: there are no AI drafts on this sheet yet." & vbCrLf & _
            "Column " & Chr$(64 + COL_D_DRAFT) & " (AI DRAFT) is empty for all " & noAi & " row(s)." & vbCrLf & vbCrLf & _
            "Paste Copilot's text into column " & Chr$(64 + COL_D_DRAFT) & " first -- the prompt is in cell L2 -- " & _
            "or just type into column " & Chr$(64 + COL_D_SUBMIT) & " (SUBMIT) yourself." & vbCrLf
        Exit Function
    End If

    CopyAiToSubmit = "Copy AI -> Submit: " & copied & " copied, " & _
        keptExisting & " left alone (you had already written something there), " & _
        noAi & " with no AI draft." & vbCrLf & _
        "Nothing was overwritten. Edit column " & Chr$(64 + COL_D_SUBMIT) & _
        ", tick column " & Chr$(64 + COL_D_APPROVED) & ", then publish." & vbCrLf
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

' ---------------------------------------------------------------------
' THE LOOK. Rohan, 2026-08-10: "clearer formatting transport tycoon style
' headers in the worksheet ... make body text and line work to suit".
'
' The point is not decoration. This sheet is a WORKBENCH with a direction of
' travel -- read C, gather D, machine writes E, you write F, you tick G -- and
' the formatting's whole job is to make that direction visible before a word is
' read. So the columns are colour-coded by ROLE, not prettified individually:
'
'   grey    you may read this but not change it      (ORIGINAL)
'   slate   small inputs you supply                  (SOURCES)
'   amber   a machine wrote this, it never publishes (AI DRAFT)
'   cream   YOUR words, this is what publishes       (SUBMIT)
'   green   the consent                              (APPROVE)
'
' The heavy rule sits to the LEFT OF SUBMIT, because that is the real boundary
' on this sheet: everything left of it is material, everything right of it is
' yours and goes to the deck.
'
' EVERY EXCEL CONSTANT IS A NUMERIC LITERAL. This module is driven from
' PowerPoint, where xlEdgeLeft, xlContinuous and friends do not resolve -- the
' same gotcha ExcelOutput hit on 2026-07-25 with xlUp. A named constant here
' would be Empty, which is 0, which silently means something else.
' A COLOUR FAMILY PER FIELD, so you can tell which sheet you are on from the tab
' strip. Rohan, 2026-08-10, with 13 tabs in the workbook and 8 fields coming.
'
' THE ROLE COLOURS DO NOT CHANGE. Grey means read-only, amber means a machine
' wrote it, cream means yours and publishes, green means consent -- on every
' sheet, always. If amber meant "AI draft" here and "sources" there, the role
' encoding would stop being readable, and it is doing more work than the sheet
' identity is. So the family colours the TITLE BAR, the HEADER ROW, the TAB and
' the small SOURCES column, and nothing else.
'
' Stable from the field name rather than from creation order, so a field keeps
' its colour across rebuilds, machines and quarters. Two fields can land on the
' same family -- harmless, they are still adjacent to their own name in a 40pt
' header -- and that is preferred over a lookup table someone has to maintain.
' POSITION FIRST, HASH ONLY AS A FALLBACK.
'
' A hash of the field name was the first attempt and it collided on the two
' fields that matter most: ABOUT_BODY and PROGRESS_BODY both landed on teal,
' which is precisely the pair a person flips between. Caught by listing the
' eight real field names against the function before shipping it, not after.
'
' The sheets are built in a known order, so passing that position guarantees
' eight distinct families for the first eight fields -- which covers the five
' in use and the three coming. Colours follow sheet order, so reordering the
' Field Spec reshuffles them; that is honest rather than surprising, because
' the tab strip reorders too.
'
' seed < 0 means "caller did not say", and only then does the name get hashed.
Private Function FamilyIndex(fieldId As String, seed As Long) As Long
    If seed >= 0 Then
        FamilyIndex = seed Mod 8
        Exit Function
    End If
    Dim h As Long, i As Long
    For i = 1 To Len(fieldId)
        h = (h * 31 + Asc(Mid$(fieldId, i, 1))) Mod 100000
    Next i
    FamilyIndex = h Mod 8
End Function

Private Function FamilyDark(fieldId As String, seed As Long) As Long
    Select Case FamilyIndex(fieldId, seed)
        Case 0: FamilyDark = RGB(46, 74, 48)      ' green
        Case 1: FamilyDark = RGB(40, 64, 92)      ' blue
        Case 2: FamilyDark = RGB(112, 58, 40)     ' rust
        Case 3: FamilyDark = RGB(74, 52, 92)      ' purple
        Case 4: FamilyDark = RGB(30, 78, 78)      ' teal
        Case 5: FamilyDark = RGB(82, 78, 34)      ' olive
        Case 6: FamilyDark = RGB(96, 42, 58)      ' maroon
        Case Else: FamilyDark = RGB(58, 62, 70)   ' slate
    End Select
End Function

Private Function FamilyTint(fieldId As String, seed As Long) As Long
    Select Case FamilyIndex(fieldId, seed)
        Case 0: FamilyTint = RGB(222, 233, 221)
        Case 1: FamilyTint = RGB(219, 229, 240)
        Case 2: FamilyTint = RGB(243, 226, 216)
        Case 3: FamilyTint = RGB(232, 224, 240)
        Case 4: FamilyTint = RGB(214, 234, 234)
        Case 5: FamilyTint = RGB(236, 236, 210)
        Case 6: FamilyTint = RGB(240, 220, 228)
        Case Else: FamilyTint = RGB(226, 228, 233)
    End Select
End Function

Private Sub ApplyDraftingLook(ws As Object, lastRow As Long, fieldId As String, familySeed As Long)
    Const EDGE_LEFT As Long = 7, EDGE_TOP As Long = 8
    Const EDGE_BOTTOM As Long = 9, EDGE_RIGHT As Long = 10
    Const INSIDE_V As Long = 11, INSIDE_H As Long = 12
    Const CONTINUOUS As Long = 1, THIN As Long = 2, MEDIUM As Long = -4138

    Dim INK As Long, RULE As Long, FAM As Long, FAMLIGHT As Long
    INK = RGB(38, 46, 38)
    RULE = RGB(120, 124, 112)
    FAM = FamilyDark(fieldId, familySeed)
    FAMLIGHT = FamilyTint(fieldId, familySeed)

    ' The tab strip becomes the index: you find the sheet before you open it.
    On Error Resume Next
    ws.Tab.Color = FAM
    On Error GoTo 0

    ' --- the instruction panel: one block, not seven loose lines ------
    With ws.Range(ws.Cells(DRAFT_INTRO_ROW, 1), ws.Cells(7, COL_D_NOTES))
        .Interior.Color = RGB(233, 229, 210)
        .Font.Color = INK
    End With
    With ws.Range(ws.Cells(DRAFT_INTRO_ROW, 1), ws.Cells(7, COL_D_NOTES)).Borders(EDGE_BOTTOM)
        .LineStyle = CONTINUOUS: .Weight = MEDIUM: .Color = RULE
    End With
    ws.Range(ws.Cells(2, 1), ws.Cells(6, 1)).Font.Bold = True

    ' Title bar, reversed out -- the one thing you see before anything else.
    With ws.Range(ws.Cells(DRAFT_INTRO_ROW, 1), ws.Cells(DRAFT_INTRO_ROW, COL_D_NOTES))
        .Interior.Color = FAM
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    ' --- the header row: chunky, reversed out, ruled underneath -------
    With ws.Range(ws.Cells(DRAFT_HEADER_ROW, 1), ws.Cells(DRAFT_HEADER_ROW, COL_D_NOTES))
        .Interior.Color = FAM
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 8
        .VerticalAlignment = -4108          ' xlCenter
        .WrapText = True
    End With
    With ws.Range(ws.Cells(DRAFT_HEADER_ROW, 1), ws.Cells(DRAFT_HEADER_ROW, COL_D_NOTES)).Borders(EDGE_BOTTOM)
        .LineStyle = CONTINUOUS: .Weight = MEDIUM: .Color = INK
    End With

    If lastRow < DRAFT_FIRST_ROW Then Exit Sub

    ' --- the body: role colour per column ----------------------------
    ws.Range(ws.Cells(DRAFT_FIRST_ROW, COL_D_ENTITY), ws.Cells(lastRow, COL_D_NAME)).Interior.Color = RGB(245, 244, 238)
    ws.Range(ws.Cells(DRAFT_FIRST_ROW, COL_D_CURRENT), ws.Cells(lastRow, COL_D_CURRENT)).Interior.Color = RGB(226, 226, 221)
    ws.Range(ws.Cells(DRAFT_FIRST_ROW, COL_D_SOURCES), ws.Cells(lastRow, COL_D_SOURCES)).Interior.Color = FAMLIGHT
    ws.Range(ws.Cells(DRAFT_FIRST_ROW, COL_D_DRAFT), ws.Cells(lastRow, COL_D_DRAFT)).Interior.Color = RGB(250, 238, 205)
    ws.Range(ws.Cells(DRAFT_FIRST_ROW, COL_D_SUBMIT), ws.Cells(lastRow, COL_D_SUBMIT)).Interior.Color = RGB(255, 252, 235)
    ws.Range(ws.Cells(DRAFT_FIRST_ROW, COL_D_APPROVED), ws.Cells(lastRow, COL_D_APPROVED)).Interior.Color = RGB(206, 226, 202)
    ws.Range(ws.Cells(DRAFT_FIRST_ROW, COL_D_NOTES), ws.Cells(lastRow, COL_D_NOTES)).Interior.Color = RGB(245, 244, 238)

    ' Body text: one size, top-aligned, prose wrapped. Set on the block so a
    ' later column tweak cannot leave one column reading differently.
    With ws.Range(ws.Cells(DRAFT_FIRST_ROW, 1), ws.Cells(lastRow, COL_D_NOTES))
        .Font.Size = 8
        .Font.Color = INK
        .VerticalAlignment = -4160          ' xlTop
    End With
    ws.Range(ws.Cells(DRAFT_FIRST_ROW, COL_D_APPROVED), ws.Cells(lastRow, COL_D_APPROVED)).HorizontalAlignment = -4108
    ws.Range(ws.Cells(DRAFT_FIRST_ROW, COL_D_APPROVED), ws.Cells(lastRow, COL_D_APPROVED)).Font.Bold = True

    ' --- line work ----------------------------------------------------
    With ws.Range(ws.Cells(DRAFT_HEADER_ROW, 1), ws.Cells(lastRow, COL_D_NOTES))
        .Borders(INSIDE_V).LineStyle = CONTINUOUS
        .Borders(INSIDE_V).Weight = THIN
        .Borders(INSIDE_V).Color = RULE
        .Borders(INSIDE_H).LineStyle = CONTINUOUS
        .Borders(INSIDE_H).Weight = THIN
        .Borders(INSIDE_H).Color = RGB(198, 200, 190)
        .Borders(EDGE_LEFT).LineStyle = CONTINUOUS
        .Borders(EDGE_LEFT).Weight = MEDIUM
        .Borders(EDGE_LEFT).Color = INK
        .Borders(EDGE_RIGHT).LineStyle = CONTINUOUS
        .Borders(EDGE_RIGHT).Weight = MEDIUM
        .Borders(EDGE_RIGHT).Color = INK
        .Borders(EDGE_TOP).LineStyle = CONTINUOUS
        .Borders(EDGE_TOP).Weight = MEDIUM
        .Borders(EDGE_TOP).Color = INK
        .Borders(EDGE_BOTTOM).LineStyle = CONTINUOUS
        .Borders(EDGE_BOTTOM).Weight = MEDIUM
        .Borders(EDGE_BOTTOM).Color = INK
    End With

    ' THE PUBLISH BOUNDARY, drawn heavy. Left of it is material; F and G are
    ' yours and reach the deck. If one line on this sheet has to be noticed,
    ' it is this one.
    With ws.Range(ws.Cells(DRAFT_HEADER_ROW, COL_D_SUBMIT), ws.Cells(lastRow, COL_D_SUBMIT)).Borders(EDGE_LEFT)
        .LineStyle = CONTINUOUS: .Weight = MEDIUM: .Color = FAM
    End With
End Sub
