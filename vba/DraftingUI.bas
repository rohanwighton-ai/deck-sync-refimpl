Attribute VB_Name = "DraftingUI"
Option Explicit

' BUTTONS FOR THE DRAFTING HALF.
'
' Everything in Drafting.bas, FieldSpec.bas and Sources.bas was reachable only
' from vba\tools\field_e2e.ps1 -- a PowerShell test harness. Which meant the
' drafting loop could be demonstrated but not USED: somebody had to type
' commands in a terminal between every step.
'
' That is not a packaging detail. TRACKER item 10 asks whether producing a real
' quarter SAVED TIME, and that question cannot be answered honestly while a
' developer is driving the steps. Rohan, 2026-08-01: "Before you do that please
' make sure I have a UI to make it happen."
'
' These are deliberately thin. Every one of them resolves its inputs the same
' way RibbonUI.SyncPreviewCore does -- active presentation, DeckRegistry for the
' paired workbook, WorkbookBridge to open it -- and then calls the SAME
' functions the harness calls. A button that resolved its inputs differently
' from the tested path could disagree with it about what happens, which is the
' reason SyncPreview was written to mirror SyncNow line for line.
'
' Excel is left VISIBLE here, unlike the harness which runs it hidden. Drafting
' is work a person does IN Excel; the button's job is to put them in front of
' the right sheet, not to do something invisible and report on it.


' ---------------------------------------------------------------------
' ONE REPORT AT THE END, NOT A DIALOG PER STAGE.
'
' Rohan, 2026-08-09, after the first real run of the two-button build:
' "with a full sync subsequent changes shouldn't trigger same hellish popup
' chain?" Ten dialogs to do one thing, and about five of them were a stage
' announcing that it had done nothing.
'
' The cause was that every stage owned its own MsgBox. So an INFORMATIONAL
' message now goes through Say: shown immediately when the stage is run on its
' own, appended to a buffer when a chain is collecting. DECISIONS and INPUTS
' are deliberately NOT routed through here -- a question that does not block is
' not a question, and the write-authorising gates must stay exactly where they
' are.
'
' EXPLICIT BEGIN/END rather than a quiet-mode flag. A mode is a thing someone
' forgets to turn off; a Begin with a matching End in the caller's error handler
' fails closed. BeginCollecting resets the buffer, so a chain that died halfway
' cannot leak its half-report into the next run.
Private mReport As String
Private mCollecting As Boolean

' THE FIELD THIS CHAIN RUN IS ABOUT, asked once and reused. Two stages need it
' -- Copy AI to Submit and Publish -- and asking the same question twice in one
' run is its own defect. Cleared by BeginCollecting for the same fail-closed
' reason the buffer is: a chain that died halfway must not answer for the next.
Private mChainField As String

Public Sub BeginCollecting()
    mReport = ""
    mChainField = ""
    mCollecting = True
End Sub

' Returns everything the stages said, and ALWAYS stops collecting -- callers put
' this in their error handler as well as their happy path.
Public Function EndCollecting() As String
    EndCollecting = mReport
    mReport = ""
    mCollecting = False
End Function

' PUBLIC AND PURE, so this is testable without a live presentation or the
' mCollecting/mChainField state that only exists mid-chain. Say() is the only
' real caller; the label logic does not need to live inside it to be correct.
'
' 2026-08-15: `2. Put it on the slides` runs 13 fields through this chain in
' one press (a deliberate design -- see PublishAllDraftedFields's header, "a
' person is not asked thirteen times"), and every field's Copy/Publish block
' landed in the SAME dialog under the SAME two fixed headers, with the field
' name buried in prose or missing entirely for the "nothing to do" case.
' Rohan, reading the result: "this msg makes zero sense" -- correctly, since
' two consecutive blocks reporting "0 rows" and "38 rows" with no visible
' field label read as self-contradictory. The one-press LOOP is right and
' stays; only the label was missing.
Public Function ChainBlockHeader(caption As String, chainField As String) As String
    ChainBlockHeader = caption
    If chainField <> "" Then ChainBlockHeader = ChainBlockHeader & " (" & chainField & ")"
End Function

Private Sub Say(text As String, style As VbMsgBoxStyle, caption As String)
    If mCollecting Then
        If mReport <> "" Then mReport = mReport & vbCrLf & vbCrLf
        mReport = mReport & "-- " & ChainBlockHeader(caption, mChainField) & " --" & vbCrLf & text
    Else
        ' NOT Say. This line was rewritten into a call to its own procedure by a
        ' scripted MsgBox -> Say replacement that did not exclude the body of Say
        ' itself, and shipped as "Run-time error 28: Out of stack space" the
        ' first time a stage ran outside a chain. Second instance of that exact
        ' mistake in one session; check_vba_static.py now refuses it.
        MsgBox text, style, caption
    End If
End Sub

' The register worksheet for this deck.
'
' ONE SHAPE, NOT TWO. This comment used to say the opposite: that the e2e rig's
' sheet named "Register" and the live pairing's per-slide-type registered name
' were both real, so it looked for the first and fell back to the second.
' Rather than declaring one wrong, it declared neither -- and a workbook with
' both put publish on one sheet and Sync Now on the other, each reporting
' success. There is one answer: the name the slide type is REGISTERED against.
' The rig was never a second shape, it just happened to register the name
' "Register", which is exactly why it could not show the bug.
'
' Returns Nothing rather than raising, because every caller here wants to say
' something useful to a person rather than show them an error dialog.
Private Function ResolveRegisterSheet(pres As Object, wb As Object, ByRef problem As String) As Object
    problem = ""

    Dim types() As String
    types = DeckRegistry.ListRegisteredTypes(pres)

    Dim lo As Long, hi As Long
    On Error Resume Next
    lo = LBound(types): hi = UBound(types)
    If Err.Number <> 0 Then
        On Error GoTo 0
        problem = "this deck has no slide type registered yet, so there is no register to draft " & _
                  "against. Onboard a slide type first."
        Exit Function
    End If
    On Error GoTo 0

    ' ONE SLIDE TYPE PER DECK. Not a simplification -- the architecture. A child
    ' deck carries one slide type and is where a human works; a composite deck is
    ' a build output, generated one-way and never hand-edited (DECISIONS
    ' 2026-07-26 and 2026-07-30). Drafting only ever runs on a child deck.
    '
    ' So more than one registered type is not a choice to offer, it is a deck in
    ' a state the model says cannot happen -- most likely two onboardings against
    ' the same file. Refused rather than resolved: this used to take types(LBound),
    ' the first in whatever order the custom properties enumerate, and draft
    ' against that type's register without saying which it had picked.
    Dim chosen As String
    If hi > lo Then
        Dim names As String, i As Long
        For i = lo To hi
            names = names & vbCrLf & "  - " & types(i)
        Next i
        problem = "this deck has " & (hi - lo + 1) & " slide types registered:" & names & vbCrLf & vbCrLf & _
            "A deck that gets drafted carries exactly one. Two usually means the deck " & _
            "was onboarded twice. Sort out which type this deck is before drafting, " & _
            "because each type has its own register sheet and picking the wrong one " & _
            "puts a quarter's writing in the wrong place."
        Exit Function
    End If
    chosen = types(lo)

    ' THE SAME ANSWER SYNC USES. This function used to prefer a sheet literally
    ' named "Register" and fall back to the registered name only if none
    ' existed -- so a workbook holding both put publish on one sheet and Sync
    ' Now on another, each reporting success. See
    ' WorkbookBridge.WorksheetForSlideType for the full account.
    Set ResolveRegisterSheet = WorkbookBridge.WorksheetForSlideType(pres, wb, chosen, problem)
End Function

' Shared opening move for all three buttons: get the deck, its workbook, and
' its register sheet, or explain to the person exactly which of those is
' missing. Returns False when it has already shown the message.
Private Function Resolve(caption As String, ByRef pres As Object, ByRef wb As Object, _
                         ByRef regWs As Object) As Boolean
    Set pres = Application.ActivePresentation

    Dim workbookPath As String
    workbookPath = DeckRegistry.GetWorkbookPath(pres)
    If workbookPath = "" Then
        Say "This deck has no paired workbook yet." & vbCrLf & vbCrLf & _
               "Onboard a slide type first -- drafting reads the register, and " & _
               "there is no register until the deck and a workbook are paired.", _
               vbExclamation, caption
        Exit Function
    End If

    Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath)
    If wb Is Nothing Then
        Say "Could not open the paired workbook at:" & vbCrLf & workbookPath, vbCritical, caption
        Exit Function
    End If

    Dim problem As String
    Set regWs = ResolveRegisterSheet(pres, wb, problem)
    If regWs Is Nothing Then
        Say "Could not work out which register sheet to use." & vbCrLf & vbCrLf & _
               problem, vbCritical, caption
        Exit Function
    End If

    Resolve = True
End Function

' Which field to draft. Read from the Field Spec sheet, because that is where
' fields declare what they ARE -- and Kind tells us which ones are worth
' drafting at all.
'
' PROSE FIELDS ARE OFFERED FIRST AND MARKED. Static and Controlled fields do not
' need a drafting sheet: their values are known, not written, and they are
' edited in the register or fed from upstream. Offering all forty equally would
' invite somebody to build forty drafting sheets, which is the failure this
' distinction exists to prevent.
' CASE-INSENSITIVE, AND IT NAMES WHAT IT MEANT.
'
' 2026-08-08: Rohan typed "About_Body" -- Office capitalises the first letter of
' an input box by habit -- and the tool answered "There is no drafting sheet for
' About_Body yet. Press '" & CommandBarUI.CAP_SET_UP_QUARTER & "' -- it builds them." Following that advice would
' have created a SECOND sheet, TPL_About_Body, alongside TPL_ABOUT_BODY: two
' drafting sheets for one field, diverging quietly.
'
' Asking someone to defeat their own autocorrect is not a specification. The
' typed text is matched against the Field Spec ignoring case and the CANONICAL
' FieldID is returned, so everything downstream still deals in exact IDs.
Private Function CanonicalFieldId(wb As Object, typed As String) As String
    CanonicalFieldId = typed
    If Trim(typed) = "" Then Exit Function
    If Not WorkbookBridge.WorksheetExists(wb, FieldSpec.SPEC_SHEET_NAME) Then Exit Function

    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(wb, FieldSpec.SPEC_SHEET_NAME)

    Dim r As Long
    r = FieldSpec.SPEC_FIRST_ROW
    Do While Trim(CStr(ws.Cells(r, FieldSpec.COL_S_FIELDID).Value)) <> ""
        Dim fid As String
        fid = Trim(CStr(ws.Cells(r, FieldSpec.COL_S_FIELDID).Value))
        If StrComp(fid, Trim(typed), vbTextCompare) = 0 Then
            CanonicalFieldId = fid
            Exit Function
        End If
        r = r + 1
    Loop
End Function

' Is this FieldID on the Field Spec sheet at all? A workbook with no Field Spec
' cannot answer, so it says yes rather than blocking a legitimate run.
Private Function FieldIsKnown(wb As Object, fieldId As String) As Boolean
    If Not WorkbookBridge.WorksheetExists(wb, FieldSpec.SPEC_SHEET_NAME) Then
        FieldIsKnown = True
        Exit Function
    End If

    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(wb, FieldSpec.SPEC_SHEET_NAME)

    Dim r As Long
    r = FieldSpec.SPEC_FIRST_ROW
    Do While Trim(CStr(ws.Cells(r, FieldSpec.COL_S_FIELDID).Value)) <> ""
        If StrComp(Trim(CStr(ws.Cells(r, FieldSpec.COL_S_FIELDID).Value)), fieldId, vbTextCompare) = 0 Then
            FieldIsKnown = True
            Exit Function
        End If
        r = r + 1
    Loop
End Function

' WHICH FIELD THIS STAGE IS ACTING ON.
'
' Standalone, the answer is on screen: ActiveDraftField reads whichever TPL_
' sheet Excel is showing, and asking would be asking a question the person has
' already answered by looking at it.
'
' INSIDE THE CHAIN IT IS THE OPPOSITE, and this is the defect this function
' exists to close. RefreshDraftingSheets ends with `ShowSheet wb, firstSheet`,
' where firstSheet is the FIRST `Kind = Prose` row on the Field Spec -- so the
' following stages inherited a sheet the chain had just chosen for itself, and
' could only ever act on that one field whatever the person intended.
'
' On the real deck that field was ABOUT_BODY, 0 submitted and 0 approved, while
' KEY_EVENTS_BODY held 43 submitted and 39 approved and could not be reached from
' the toolbar at all -- there is no field picker, deliberately. Every Sync Now
' published nothing, reported "0 would be published", and finished quietly. That
' is a large part of why no drafted value had ever reached a slide.
'
' Found 2026-08-13 by pressing the button on the real deck. 192 tests pass and
' none of them asks whether a person can CAUSE a given field to publish -- they
' test that publishing works once called. Same shape as the picture injector and
' the progress bars: built, tested, and behind a locked door.
'
' TWO CALL SITES, FIXED TOGETHER. CopyAiDraftsToSubmit had the identical line and
' the identical consequence. Fixing only where it was noticed is how the
' quarter-ordering defect cost a second failed run in the same hour.
'
' Rejected fix: reorder the Field Spec so the wanted field comes first. Rohan,
' on being offered it: "why are you having to move register rows manually?
' Worries me that the code won't work when it needs to." Right on both counts --
' it makes which field reaches a slide depend on spreadsheet row order, and it
' does not exist on the work machine, where a quarter must run from buttons.
' EVERY DRAFTED FIELD, ASKING NOTHING. 2026-08-14.
'
' Rohan, on being asked which field: "it shouldn't have to ask" -- and he is
' right, twice over. His own walkthrough of the tool never once mentions choosing
' a field; the entire per-field selection UX is absent from how he thinks about
' this. And it is absent for a good reason: the review queue downstream is ALREADY
' matrix-shaped, because ChangeHash is keyed per entity AND field. The question
' existed only because publish was written one field at a time, standing in front
' of a machine that handles all of them perfectly well.
'
' So the picker is not improved, it is DELETED from the chain. That removes the
' question, its wording, and the Excel-focus problem the click version had --
' none of which should have been on screen.
'
' Each field still publishes through exactly the same path, with the same tick
' gate. Nothing is approved here that would not have been approved one at a time;
' the difference is that a person is not asked thirteen times.
'
' AskForField survives for the standalone buttons, where choosing really is the
' point and the answer is not "all of them".
Public Sub PublishAllDraftedFields(caption As String)
    Dim pres As Object, wb As Object, regWs As Object
    If Not Resolve(caption, pres, wb, regWs) Then Exit Sub

    Dim list As String
    list = ProseFields(wb)
    If Trim(list) = "" Then
        Say "There are no Prose fields on the Field Spec sheet, so there is nothing to publish.", _
            vbInformation, caption
        Exit Sub
    End If

    Dim parts() As String, i As Long
    parts = Split(list, ",")

    For i = LBound(parts) To UBound(parts)
        mChainField = Trim(parts(i))
        If mChainField <> "" Then
            ' Copy first, publish second, and BOTH per field: CopyAiToSubmit
            ' never overwrites a row that already has your words, so running it
            ' here is safe and saves a separate press. It used to run in the
            ' quarter-setup chain, BEFORE Copilot had written anything -- which
            ' is why it asked a question about drafts that did not exist yet.
            CopyAiDraftsToSubmit
            PublishDraftsForField
        End If
    Next i

    mChainField = ""
End Sub

Private Function FieldForRun(caption As String, wb As Object) As String
    If mCollecting Then
        If mChainField = "" Then mChainField = AskForField(caption, wb)
        FieldForRun = mChainField
        Exit Function
    End If

    FieldForRun = ActiveDraftField(wb)
    If FieldForRun = "" Then FieldForRun = AskForField(caption, wb)
End Function

Private Function AskForField(caption As String, wb As Object) As String
    If Not WorkbookBridge.WorksheetExists(wb, FieldSpec.SPEC_SHEET_NAME) Then
        AskForField = Trim(InputBox("Which field? (" & caption & ")" & vbCrLf & vbCrLf & _
            "(No 'Field Spec' sheet in this workbook yet, so there is no list to " & _
            "choose from -- type the FieldID, e.g. ABOUT_BODY.)", caption))
        AskForField = CanonicalFieldId(wb, AskForField)
        Exit Function
    End If

    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(wb, FieldSpec.SPEC_SHEET_NAME)

    Dim prose As String, other As String
    Dim r As Long
    r = FieldSpec.SPEC_FIRST_ROW
    Do While Trim(CStr(ws.Cells(r, FieldSpec.COL_S_FIELDID).Value)) <> ""
        Dim fid As String, kind As String
        fid = Trim(CStr(ws.Cells(r, FieldSpec.COL_S_FIELDID).Value))
        kind = Trim(CStr(ws.Cells(r, FieldSpec.COL_S_KIND).Value))
        If StrComp(kind, "Prose", vbTextCompare) = 0 Then
            prose = prose & "    " & fid & vbCrLf
        Else
            other = other & "    " & fid & "  (" & kind & ")" & vbCrLf
        End If
        r = r + 1
    Loop

    Dim msg As String
    msg = "Which field? (" & caption & ")" & vbCrLf & vbCrLf
    If prose <> "" Then
        msg = msg & "WORTH DRAFTING -- the words are the work:" & vbCrLf & prose & vbCrLf
    End If
    If other <> "" Then
        msg = msg & "These do NOT need a drafting sheet. Their values are known, not" & vbCrLf & _
                    "written -- edit them in the register instead:" & vbCrLf & other & vbCrLf
    End If
    msg = msg & "Type the FieldID -- capitals do not matter."

    ' CLICKING IS THE WAY IN; TYPING IS THE FALLBACK, not the other way round.
    Dim clicked As String
    clicked = PickFieldByClicking(caption, wb, ws)
    If clicked <> "" Then
        AskForField = clicked
        Exit Function
    End If

    AskForField = CanonicalFieldId(wb, Trim(InputBox(msg, caption)))
End Function


' PICK A FIELD BY CLICKING IT. Returns "" if the person cancels.
'
' Replaces a typed InputBox that required an EXACT match against a 30-item list
' printed in its own prompt -- a list long enough that it pushed the text box off
' the bottom of the screen, so the thing you had to type into was not on screen
' with the thing you had to type. Rohan, 2026-08-14: "I don't want to do any
' secret hidden typing, I need to select the field by clicking on it."
'
' Excel's InputBox Type:=8 is a RANGE picker: it collapses to a bar, the person
' clicks any cell, and it hands back a Range. Chosen over a UserForm because
' build_ppam.ps1 imports .bas modules only -- a .frm would change the build
' pipeline -- and over a CommandBars popup because this one also PUTS THE LIST IN
' FRONT OF THEM, which the typed version only pretended to do.
'
' Any cell in the row counts. Requiring the FieldID column would be the same
' exact-target problem one step smaller.
Private Function PickFieldByClicking(caption As String, wb As Object, ws As Object) As String
    ShowSheet wb, FieldSpec.SPEC_SHEET_NAME
    BringExcelToFront wb

    ' A LOOP, NOT RECURSION. The first version called itself on a miss, passing
    ' the same three arguments -- which check_vba_static.py correctly refused as
    ' unbounded. A person clicking the wrong row should get another go, and that
    ' is a loop; re-entering the procedure to express "try again" was borrowing a
    ' mechanism that carries a stack with it for no reason.
    Dim pick As Object
    Dim picked As String

    Do
        Set pick = Nothing
        On Error Resume Next
        Set pick = wb.Application.InputBox( _
            Prompt:="Click any cell in the row of the field you want, then press OK." & vbCrLf & vbCrLf & _
                    "(Cancel to type the name instead.)", _
            Title:=caption, Type:=8)
        On Error GoTo 0

        ' Cancelled. "" here means "did not choose", which the caller reads as a
        ' reason to offer typing -- distinct from "chose something invalid".
        If pick Is Nothing Then Exit Function

        picked = ""
        On Error Resume Next
        If pick.Row >= FieldSpec.SPEC_FIRST_ROW Then
            picked = Trim(CStr(ws.Cells(pick.Row, FieldSpec.COL_S_FIELDID).Value))
        End If
        On Error GoTo 0

        If picked <> "" Then Exit Do

        ' A row above the list, or below its last entry, is a miss rather than a
        ' choice. Saying so and asking again beats returning "" and having the
        ' run silently skip the field.
        Say "That row is not a field." & vbCrLf & vbCrLf & _
            "Click a cell in one of the listed field rows.", vbExclamation, caption
    Loop

    PickFieldByClicking = CanonicalFieldId(wb, picked)
End Function

' Put the person in front of the sheet they just asked for. The whole point of
' a button over a script is that it ends with you looking at the thing.
Private Sub ShowSheet(wb As Object, sheetName As String)
    On Error Resume Next
    wb.Application.Visible = True
    wb.Activate
    wb.Worksheets(sheetName).Activate
    wb.Worksheets(sheetName).Range("A1").Select
    On Error GoTo 0
End Sub


' PUT EXCEL'S WINDOW IN FRONT, not merely visible.
'
' Seen live 2026-08-14, first press of addin86: the click-a-cell field picker
' opened CORRECTLY and over POWERPOINT, because ShowSheet above sets
' Application.Visible and activates the sheet but never raises Excel's window.
' The range picker was live and pointing at a grid the person could not see -- a
' picker you cannot look at is worse than the typed box it replaced, since at
' least that one admitted it wanted typing.
'
' AppActivate matches on the START of a window title, and Excel's title begins
' with the workbook name ("register-wide.xlsx - Excel"), so activating on
' Application.Caption alone does not reliably match. Setting the caption first
' makes the title start with a string we chose, so the match is deterministic;
' it is restored immediately afterwards.
'
' Entirely best-effort. If it fails the picker still works, it is just behind --
' and Cancel still falls back to typing, so nobody is stuck either way.
Private Sub BringExcelToFront(wb As Object)
    Const MARKER As String = "Deck Sync"

    On Error Resume Next
    wb.Application.Visible = True
    If wb.Application.WindowState = 2 Then wb.Application.WindowState = -4143  ' xlMinimized -> xlNormal

    Dim previous As String
    previous = wb.Application.Caption
    wb.Application.Caption = MARKER
    AppActivate MARKER
    wb.Application.Caption = previous
    On Error GoTo 0
End Sub

' Every field on the Field Spec sheet whose Kind is Prose, comma-separated.
'
' PROSE ONLY, and that restriction is what makes automatic generation safe.
' Static and Controlled fields have known values, not written ones -- generating
' a drafting sheet per field would produce forty tabs, thirty-five of which
' would be a surface for work nobody should do there.
Private Function ProseFields(wb As Object) As String
    If Not WorkbookBridge.WorksheetExists(wb, FieldSpec.SPEC_SHEET_NAME) Then Exit Function

    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(wb, FieldSpec.SPEC_SHEET_NAME)

    Dim out As String
    Dim r As Long
    r = FieldSpec.SPEC_FIRST_ROW
    Do While Trim(CStr(ws.Cells(r, FieldSpec.COL_S_FIELDID).Value)) <> ""
        If StrComp(Trim(CStr(ws.Cells(r, FieldSpec.COL_S_KIND).Value)), "Prose", vbTextCompare) = 0 Then
            out = out & IIf(out = "", "", ",") & Trim(CStr(ws.Cells(r, FieldSpec.COL_S_FIELDID).Value))
        End If
        r = r + 1
    Loop
    ProseFields = out
End Function

' ---------------------------------------------------------------------------
' BUTTON: Refresh Drafting Sheets
' ---------------------------------------------------------------------------
'
' NO FIELD PICKER, DELIBERATELY. Rohan, 2026-08-01: "Shouldn't drafting just
' appear? does it need buttons?" He is right. Building a drafting sheet is not
' a decision -- it is maintenance, and asking a person which field they would
' like a sheet for makes the tool's housekeeping their problem.
'
' The two operations that DO need a button are the ones that need consent:
' Copy AI -> Submit writes into the column they own, and Publish puts the word
' "Approved" in their mouth. Those ask. This does not.
'
' Safe to run at any time: WriteDraftingSheet preserves every drafted, submitted
' and sourced cell, and refuses to carry anything across a layout change.
Public Sub RefreshDraftingSheets()
    Const CAP As String = "Refresh Drafting Sheets"
    On Error GoTo Failed

    Dim pres As Object, wb As Object, regWs As Object
    If Not Resolve(CAP, pres, wb, regWs) Then Exit Sub

    Dim specWs As Object
    Set specWs = WorkbookBridge.GetOrAddWorksheet(wb, FieldSpec.SPEC_SHEET_NAME)
    FieldSpec.WriteSpecSheet specWs

    Dim srcWs As Object
    Set srcWs = WorkbookBridge.GetOrAddWorksheet(wb, Sources.SOURCES_SHEET_NAME)
    Sources.WriteSourcesSheet srcWs

    ' The period gets PICKED on the Sources sheet, from the periods the register
    ' actually holds. Called here because this runs on every drafting build, so
    ' a period added since last time is offered next time without a separate step.
    ' EVERY DECLARED FIELD NEEDS SOMEWHERE TO LIVE.
    '
    ' A register column is only ever created as a side effect of WRITING a value
    ' -- UpsertRow appends one when a field is published from a drafting sheet or
    ' harvested from a tagged shape. A `Given` field is neither: nobody drafts
    ' it, and it is typed straight into the register. So declaring one in the
    ' Field Spec used to give it a recipe, a source, and nowhere to be entered.
    '
    ' This runs here because this IS the "make the workbook match the Field Spec"
    ' operation -- it already rebuilds the spec sheet, the sources sheet and the
    ' drafting sheets from it. The register was the one thing left out.
    '
    ' OFFERS, NEVER SILENTLY ADDS. Columns are cheap but not free: a mistyped
    ' FieldID would appear as a real column and look authoritative. Naming them
    ' first is what makes a typo visible while it is still one keystroke to fix.
    Dim missingCols As String
    missingCols = ExcelOutput.MissingRegisterColumns(specWs, regWs)
    If missingCols <> "" Then
        Dim colCount As Long
        colCount = UBound(Split(missingCols, ",")) - LBound(Split(missingCols, ",")) + 1
        If MsgBox( _
            colCount & " field(s) on the Field Spec have no column in the register, " & _
            "so there is nowhere to enter them:" & vbCrLf & vbCrLf & _
            "    " & Replace(missingCols, ",", ", ") & vbCrLf & vbCrLf & _
            "Add a column for each?" & vbCrLf & vbCrLf & _
            "Yes -- add the headers now (no values, just the columns)." & vbCrLf & _
            "No  -- leave them; they stay unenterable until they have a column." & vbCrLf & vbCrLf & _
            "Derived fields are deliberately not listed -- they are computed, " & _
            "never stored.", _
            vbYesNo + vbQuestion, CAP) = vbYes Then
            Dim addedCols As String
            addedCols = ExcelOutput.AddRegisterColumns(regWs, missingCols)
            If addedCols = "" Then
                Say "No register columns were added -- none of the headers could be written.", _
                    vbExclamation, CAP
            Else
                Say "Added register column(s): " & addedCols, vbInformation, CAP
            End If
        End If
    End If

    Dim srcValidation As String
    srcValidation = Sources.ApplyPeriodValidation(srcWs, regWs)

    Dim fields As String
    fields = ProseFields(wb)
    If fields = "" Then
        Say "No field on the 'Field Spec' sheet is marked Kind = Prose, so there " & _
               "is nothing to draft." & vbCrLf & vbCrLf & _
               "Static and Controlled fields are edited in the register, not drafted.", _
               vbInformation, CAP
        Exit Sub
    End If

    Dim period As String
    period = DeckRegistry.GetDeckPeriod(pres)
    If period = "" Then
        Say "This deck does not declare a period, so the register cannot be " & _
               "filtered to a quarter." & vbCrLf & vbCrLf & _
               "Roll the deck forward (or set its period) before drafting.", vbExclamation, CAP
        Exit Sub
    End If

    ' THE SAME READ SYNC NOW USES. This called Register.ReadRegisterAllStatuses
    ' until 2026-08-05 -- the LONG register -- while Sync Now read the wide
    ' sheet. Two files, so nothing drafted here could reach a slide.
    '
    ' The slide-type argument goes with it, and is not replaced. Type separation
    ' on the wide sheet is the WORKSHEET: each slide type has its own, registered
    ' per deck. Filtering within a sheet was the long register's answer to having
    ' every type in one table. Worth noting what that argument had become -- it
    ' was the literal "q", the rig's old type name, renamed to "project-status"
    ' in 80fe9af. Every row would have been rejected as the wrong type.
    Dim problem As String
    Dim reg As Sheet
    reg = ExcelOutput.ReadSheetForDeckPeriod(regWs, period, problem)
    If problem <> "" Then
        Say "Cannot build drafting sheets from this register." & vbCrLf & vbCrLf & problem & _
               vbCrLf & vbCrLf & "Nothing was written.", vbExclamation, CAP
        Exit Sub
    End If

    Dim parts As Variant
    parts = Split(fields, ",")

    Dim report As String, firstSheet As String
    ' A REFUSAL IS THE HEADLINE, NOT A LINE IN THE LOG. WriteDraftingSheet was
    ' changed on 2026-08-13 to REFUSE rather than discard when a rollover would
    ' destroy typed work -- but this routine still concatenated the refusal into
    ' `report`, sent it to the Run Log sheet, and then said "drafting sheets are
    ' ready. Workbook saved." So the one outcome a person must act on was the one
    ' outcome the dialog denied. With all seven sheets stamped Q4F26 against a
    ' Q3F26 deck, EVERY sheet refuses and the modal still reads as success.
    ' Same "reports success without confirming the effect" shape this project has
    ' now fixed five times; counted here so it can be SAID.
    Dim refusedCount As Long, refusedFields As String
    Dim i As Long
    Dim draftOrder As String
    Dim seedIndex As Long
    seedIndex = 0
    For i = LBound(parts) To UBound(parts)
        Dim fid As String
        fid = Trim(CStr(parts(i)))
        ' The colour family comes from POSITION, so the eight fields get eight
        ' distinct families. A name hash collided ABOUT_BODY with PROGRESS_BODY.
        Dim sName As String
        sName = Drafting.DraftSheetNameFor(fid)
        If firstSheet = "" Then firstSheet = sName
        If draftOrder <> "" Then draftOrder = draftOrder & vbLf
        draftOrder = draftOrder & sName

        Dim ws As Object
        Set ws = WorkbookBridge.GetOrAddWorksheet(wb, sName)
        ' THE CADENCE PARAMETER IS GONE (2026-08-14). A quarter turn now FERRIES
        ' last quarter's SUBMIT into the REPORTED LAST TIME column rather than
        ' deciding per row whether to drop it, so there is nothing left to ask
        ' the register about.
        '
        ' The comment that stood here said "No cadence argument" while the call
        ' below passed a bare positional `Nothing` into that very slot. Removing
        ' the parameter then bound `Nothing` to srcWs and srcWs to a Long, and
        ' the compile died with "ByRef argument type mismatch". The comment was
        ' read instead of the code, and it was wrong in the one way that mattered:
        ' it described the argument BY NAME for a call that passes POSITIONALLY.
        Dim fieldReport As String
        fieldReport = Drafting.WriteDraftingSheet(ws, reg, fid, specWs, period, srcWs, seedIndex)
        ' Matched on the prefix WriteDraftingSheet returns, which is its contract
        ' for "nothing was changed" -- not on the prose after it, which is written
        ' for a person and will be reworded.
        If Left$(fieldReport, 8) = "REFUSED " Then
            refusedCount = refusedCount + 1
            If refusedFields <> "" Then refusedFields = refusedFields & ", "
            refusedFields = refusedFields & fid
        End If
        report = report & fid & ": " & fieldReport & vbCrLf
        seedIndex = seedIndex + 1
    Next i

    WorkbookBridge.ArrangeTabs wb, draftOrder
    WorkbookBridge.WriteWorkbookIndex wb
    WorkbookBridge.FormatRegisterSheet regWs

    ' Dropdowns on the controlled fields, and a report of anything already in
    ' the register that the vocabulary does not allow.
    Dim valNote As String
    Dim outOfVocab As Long
    valNote = FieldSpec.ApplyControlledValidation(regWs, specWs, outOfVocab)

    ' THE OTHER TWO DROPDOWNS, which had no caller at all.
    '
    ' ApplyBehaviourValidation was written, tested by nothing, and invoked from
    ' nowhere -- so the Behaviour column has never once offered its list. Found
    ' 2026-08-10 while wiring the Renders-as column beside it; the same
    ' built-and-unreachable shape as the injectors and the marking route, in a
    ' third place. Both are applied here because this is the one routine that
    ' rebuilds the Field Spec sheet.
    valNote = valNote & vbCrLf & FieldSpec.ApplyBehaviourValidation(specWs)
    valNote = valNote & vbCrLf & FieldSpec.ApplyRendersValidation(specWs)

    ShowSheet wb, firstSheet

    ' THE DETAIL GOES ON A SHEET; THE DIALOG KEEPS FOUR LINES.
    '
    ' Trimming the wording was tried first and was not enough -- Rohan, twice:
    ' "illegible, too long", then "still pretty hard to understand". The container
    ' was the problem. A modal cannot be scrolled, kept, or returned to, and
    ' MsgBox silently truncates past ~1024 characters, so the more the report had
    ' to say the less of it survived.
    WorkbookBridge.WriteRunLog wb, _
        "Drafting sheets rebuilt for " & period, _
        report & vbCrLf & valNote & vbCrLf & srcValidation

    ' THE REFUSAL GOES FIRST SO TRUNCATION EATS THE GUIDANCE, NOT THE WARNING.
    ' MsgBox caps near 1024 characters and truncates silently, so ordering is the
    ' only guarantee that the actionable half survives.
    Dim msg As String
    If refusedCount > 0 Then
        msg = refusedCount & " drafting sheet(s) were NOT rebuilt." & vbCrLf & vbCrLf & _
              refusedFields & vbCrLf & vbCrLf & _
              "They hold writing for a different quarter than the deck declares (" & period & "). " & _
              "Nothing on them was changed." & vbCrLf & vbCrLf & _
              "Set the deck's quarter to match, or publish their work first." & vbCrLf & vbCrLf
    End If

    msg = msg & period & " -- " & IIf(refusedCount > 0, "the remaining drafting sheets are ready.", _
              "drafting sheets are ready.") & vbCrLf & vbCrLf & _
          "Your wording goes in column " & Chr$(64 + Drafting.COL_D_SUBMIT) & " (SUBMIT). Type Y in column " & _
              Chr$(64 + Drafting.COL_D_APPROVED) & " to approve." & vbCrLf & _
          "Column " & Chr$(64 + Drafting.COL_D_CURRENT) & " is what the slide says now. " & _
              "Copilot's prompt is in cell L2." & vbCrLf & _
          "Nothing reaches a slide until you publish and apply."

    If outOfVocab > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              outOfVocab & " value(s) are not in their allowed list. Nothing was changed."
    End If

    ' SAVED, AND SAID SO. Sync had this defect and was fixed on 2026-08-08;
    ' drafting had it too and was missed, because the fix went in where the
    ' failure was found rather than everywhere the same shape existed. A rebuild
    ' that renumbers columns and is not written to disk is worse than one that
    ' never ran: Excel holds the new layout, the file holds the old.
    Dim saveProblem As String
    saveProblem = WorkbookBridge.SaveWorkbookVerified(wb)

    msg = msg & vbCrLf & vbCrLf & "Full detail is on the '" & _
          WorkbookBridge.RUN_LOG_SHEET_NAME & "' sheet."
    If saveProblem = "" Then
        msg = msg & vbCrLf & "Workbook saved."
    Else
        msg = msg & vbCrLf & vbCrLf & saveProblem
    End If

    ' An information icon on a run that refused work reads as "all done".
    Say msg, IIf(refusedCount > 0, vbExclamation, vbInformation), CAP
    Exit Sub

Failed:
    Say "Could not refresh the drafting sheets." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, CAP
End Sub


' The field whose drafting sheet is currently in front of the person, or ""
' if they are not looking at one.
'
' Asking "which field?" when Excel is already showing TPL_ABOUT_BODY is the
' same species of nuisance as asking which field to build a sheet for. Falls
' back to asking only when the answer genuinely is not on screen.
Private Function ActiveDraftField(wb As Object) As String
    On Error Resume Next
    Dim n As String
    n = CStr(wb.Application.ActiveSheet.Name)
    On Error GoTo 0
    If Left(n, 4) <> "TPL_" Then Exit Function

    ' Recovered by matching the sheet NAME the tool itself would generate --
    ' never by slicing "TPL_" off the front, because DraftSheetNameFor
    ' sanitises and hash-disambiguates long FieldIDs, so the sheet name is not
    ' always the field name with a prefix.
    Dim parts As Variant
    parts = Split(ProseFields(wb), ",")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If StrComp(Drafting.DraftSheetNameFor(Trim(CStr(parts(i)))), n, vbTextCompare) = 0 Then
            ActiveDraftField = Trim(CStr(parts(i)))
            Exit Function
        End If
    Next i
End Function

' ---------------------------------------------------------------------------
' BUTTON: Copy AI drafts into Submit
' ---------------------------------------------------------------------------
Private Sub CopyAiDraftsToSubmit()
    Const CAP As String = "Copy AI to Submit"
    On Error GoTo Failed

    Dim pres As Object, wb As Object, regWs As Object
    If Not Resolve(CAP, pres, wb, regWs) Then Exit Sub

    Dim fieldId As String
    fieldId = FieldForRun(CAP, wb)
    If fieldId = "" Then Exit Sub

    Dim sheetName As String
    sheetName = Drafting.DraftSheetNameFor(fieldId)
    If Not WorkbookBridge.WorksheetExists(wb, sheetName) Then
        ' "Press '" & CommandBarUI.CAP_SET_UP_QUARTER & "' first" is the RIGHT advice for a real field with no
        ' sheet yet, and the WRONG advice for a typo -- following it would build a
        ' sheet for a FieldID that does not exist. So the two cases are separated.
        If StrComp(CanonicalFieldId(wb, fieldId), fieldId, vbBinaryCompare) <> 0 Or _
           Not FieldIsKnown(wb, fieldId) Then
            Say "There is no field called '" & fieldId & "' on the Field Spec sheet." & vbCrLf & vbCrLf & _
                   "Check the spelling against the list, then try again. Nothing was created.", _
                   vbExclamation, CAP
        Else
            Say "There is no drafting sheet for " & fieldId & " yet." & vbCrLf & vbCrLf & _
                   "Press '" & CommandBarUI.CAP_SET_UP_QUARTER & "' -- it builds them.", vbExclamation, CAP
        End If
        Exit Sub
    End If

    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(wb, sheetName)

    Dim note As String
    note = Drafting.CopyAiToSubmit(ws) & Drafting.RefreshSubmitCounts(ws)

    Dim copySaveProblem As String
    copySaveProblem = WorkbookBridge.SaveWorkbookVerified(wb)
    If copySaveProblem = "" Then
        note = note & vbCrLf & "Workbook saved."
    Else
        note = note & vbCrLf & vbCrLf & copySaveProblem
    End If

    ShowSheet wb, sheetName
    Say note, vbInformation, CAP
    Exit Sub

Failed:
    Say "Could not copy the AI drafts." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, CAP
End Sub

' ---------------------------------------------------------------------------
' BUTTON: Publish Drafts
' ---------------------------------------------------------------------------
'
' PREVIEWS FIRST, ALWAYS, AND ASKS. Publishing writes Approved into the
' register, and Approved is the register's word for "a person read this and
' meant it". A button that wrote that without showing what it was about to
' write would be putting those words in your mouth.
Private Sub PublishDraftsForField()
    Const CAP As String = "Publish Drafts"
    On Error GoTo Failed

    Dim pres As Object, wb As Object, regWs As Object
    If Not Resolve(CAP, pres, wb, regWs) Then Exit Sub

    Dim fieldId As String
    fieldId = FieldForRun(CAP, wb)
    If fieldId = "" Then Exit Sub

    Dim sheetName As String
    sheetName = Drafting.DraftSheetNameFor(fieldId)
    If Not WorkbookBridge.WorksheetExists(wb, sheetName) Then
        Say "There is no drafting sheet for " & fieldId & " yet." & vbCrLf & vbCrLf & _
               "Press '" & CommandBarUI.CAP_SET_UP_QUARTER & "' -- it builds them.", vbExclamation, CAP
        Exit Sub
    End If

    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(wb, sheetName)

    Dim srcWs As Object
    Set srcWs = WorkbookBridge.GetOrAddWorksheet(wb, Sources.SOURCES_SHEET_NAME)

    ' The deck says which quarter it is, exactly as it does for drafting and for
    ' sync. Asked here rather than defaulted, because publish now writes into one
    ' period's rows and there is no safe guess -- writing the wrong period
    ' overwrites a real quarter's text.
    Dim period As String
    period = DeckRegistry.GetDeckPeriod(pres)
    If period = "" Then
        Say "This deck does not declare a period, so there is no way to know " & _
               "which quarter's rows to publish into." & vbCrLf & vbCrLf & _
               "Press '" & CommandBarUI.CAP_SET_UP_QUARTER & "' -- it sets the quarter first.", vbExclamation, CAP
        Exit Sub
    End If

    ' Said BEFORE the confirmation, not after the write. A warning that arrives
    ' with the result is a warning about something already done.
    ' Covers read-only AND macro-enabled: both mean this write will not stick,
    ' and both used to fail silently.
    Dim macroWarn As String
    macroWarn = WorkbookBridge.WriteBlockedReason(wb)
    If macroWarn <> "" Then macroWarn = macroWarn & vbCrLf & vbCrLf

    ' IS THIS EVEN OUR REGISTER? Asked here because this is the write path into
    ' it. Until 2026-08-14 nothing anywhere compared the workbook's DeckReference
    ' GUID against this deck's identity, so a deck pointed at a stranger's
    ' register wrote every field into it and reported success at every stage.
    '
    ' REFUSES rather than warns, and only this one case does. A wrong register is
    ' not a thing to be waved through with an OK -- there is no reading of "yes"
    ' that is correct, which is the same test that condemned the invariant
    ' prompts. Unstamped registers are NOT caught here; see PairingProblem.
    Dim pairNote As String
    pairNote = DeckRegistry.PairingProblem(pres, wb)
    If pairNote <> "" Then
        Say pairNote, vbCritical, CAP
        Exit Sub
    End If

    Dim preview As String
    preview = Drafting.PublishDrafts(ws, regWs, fieldId, period, True, srcWs)

    WorkbookBridge.WriteRunLog wb, "Publish " & fieldId & " -- preview", preview

    ' NOTHING TO PUBLISH IS NOT A DECISION, SO DO NOT ASK.
    '
    ' Seen live 2026-08-10: this stopped the whole chain to ask "write these into
    ' the register?" over a preview reading `0 published, 0 drafted but not
    ' ticked, 0 ticked but empty, 0 with no register row, 0 failed`, then
    ' reported that it had written nothing, then SAVED the workbook. Three
    ' interruptions for a stage with no work in it -- exactly the "about five of
    ' them a stage announcing it did nothing" Rohan named as the biggest thing
    ' between this tool and one he would use willingly.
    '
    ' Detected from the preview's own counts rather than a separate calculation,
    ' so the question and the report can never disagree about whether there was
    ' anything to do.
    If Drafting.NothingToPublish(preview) Then
        Say "Nothing to publish for " & fieldId & " in " & period & "." & vbCrLf & vbCrLf & _
            "Type your wording into the SUBMIT column and tick APPROVE, then run this again." & _
            IIf(macroWarn <> "", vbCrLf & vbCrLf & macroWarn, ""), vbInformation, CAP
        Exit Sub
    End If

    ' THE "WRITE THESE INTO THE REGISTER?" PROMPT IS GONE. 2026-08-14.
    '
    ' THE REGISTER IS NOT THE DECK. Nothing reaches a slide from it without the
    ' review tick, the workbook is backed up, and the write is reversible -- so
    ' this gate guarded nothing that was not already guarded, while standing in
    ' the middle of the loop the whole tool exists to make repeatable.
    '
    ' The preview it displayed is NOT lost: WriteRunLog above records it every
    ' run, under "Publish <field> -- preview", where it can be read after the
    ' fact instead of only in the two seconds before it is dismissed.
    Dim result As String
    result = Drafting.PublishDrafts(ws, regWs, fieldId, period, False, srcWs)

    ' SAVE. THE BUTTON PATH DID NOT.
    '
    ' PublishDrafts writes the field's value into the register in memory. The
    ' harness path (E2EField.PublishDraftSheet) has always called wb.Save; this
    ' button never did. So "published: 12" was true of Excel's buffer and of
    ' nothing on disk -- and a person who believes they have published is
    ' exactly the person who clicks "Don't Save" on the way out.
    '
    ' Found by consultant review 2026-08-01 and verified before fixing. The
    ' register is the record; writing to it without committing is the one place
    ' this tool cannot afford to be optimistic.
    '
    ' Reported, not assumed: the result says where it saved, and says loudly if
    ' the save did not take -- a read-only workbook (see FIRST-REAL-RUN finding
    ' 14) fails here silently otherwise.
    ' Err.Number = 0 IS NOT EVIDENCE THE FILE MOVED. It was the check here while
    ' the other three write paths in this module (:400, :487, :767) all used
    ' SaveWorkbookVerified, which compares the file's DateLastModified before and
    ' after. The gap matters most in exactly the case this block exists to catch:
    ' WorkbookBridge documents a macro-enabled/managed-policy save that is SILENT
    ' -- no error raised -- so Err.Number stays 0 and the register reports "SAVED"
    ' against a file nothing was written to. This is the button that writes the
    ' record, and it was the one still trusting the writer.
    Dim savedOk As Boolean
    Dim saveProblemText As String
    saveProblemText = WorkbookBridge.SaveWorkbookVerified(wb)
    savedOk = (saveProblemText = "")

    If savedOk Then
        result = result & vbCrLf & vbCrLf & "Register SAVED to:" & vbCrLf & wb.FullName
    Else
        result = result & vbCrLf & vbCrLf & _
            "!! THE REGISTER COULD NOT BE SAVED !!" & vbCrLf & _
            "The rows above are in Excel's memory and NOT on disk. Do not close " & _
            "Excel without saving." & vbCrLf & vbCrLf & _
            saveProblemText & vbCrLf & vbCrLf & _
            "Most likely the file is open read-only, or somewhere you cannot write:" & vbCrLf & _
            wb.FullName
    End If

    ShowSheet wb, WorkbookBridge.REGISTER_SHEET_NAME

    ' THE STEP THAT TOLD YOU TO PRESS THE NEXT BUTTON NOW OFFERS TO.
    '
    ' This used to end with "Now press '" & CommandBarUI.CAP_SET_UP_QUARTER & "'" -- the tool admitting the
    ' boundary was artificial. Nothing happens between writing Approved into the
    ' register and looking at what that would do to slides, so there is no
    ' decision for a button to mark. Offered rather than done, because it opens
    ' the deck and a person may not want that yet.
    WorkbookBridge.WriteRunLog wb, "Publish " & fieldId & " -- published", result

    ' INSIDE THE CHAIN THIS QUESTION IS ALREADY ANSWERED, so it is not asked.
    '
    ' The chain's own plan -- agreed two dialogs earlier -- says step 5 shows
    ' every slide change and asks before writing any of it. Asking here whether
    ' to preview is the same question a second time, and a chain that asks twice
    ' about one thing is how a person learns to click through both.
    '
    ' Pressed on its own, publish has no step 5 to follow it, so the offer
    ' stands. mCollecting is exactly "am I running inside the chain".
    If mCollecting Then
        Say RibbonUI.CapReport(result), vbInformation, CAP
        Exit Sub
    End If

    If MsgBox(RibbonUI.CapReport(result) & vbCrLf & vbCrLf & _
              "Show what this would change on the slides?" & vbCrLf & _
              "Preview only -- reads the deck, writes nothing.", _
              vbYesNo + vbQuestion, CAP) = vbYes Then
        RibbonUI.SyncPreview
    End If
    Exit Sub

Failed:
    Say "Could not publish." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, CAP
End Sub

' ---------------------------------------------------------------------------
' BUTTON: 0. Start a Quarter
' ---------------------------------------------------------------------------
'
' THE LOOP HAD NO BEGINNING. The toolbar ran 1..4, but the quarter starts before
' step 1: somebody has to tell the deck it is now FY27Q1. DeckRegistry.SetDeckPeriod
' was reachable only from a harness module and a Python script -- so the first act
' of a new quarter required a developer in WSL.
'
' Worse, skipping it FAILED QUIETLY. RefreshDraftingSheets filters the register by
' the deck's period and reports which period it used in a MsgBox AFTER writing the
' sheets. Draft a whole quarter against last quarter's rows and nothing objects.
'
' Found by product review 2026-08-01. Deliberately does NOT create the new
' period's register rows -- that is a bigger decision (which rows carry forward,
' at what Status) and guessing at it would be worse than not doing it. This tells
' the deck what quarter it is, verifies the write landed, and says plainly what is
' still missing.
Public Sub StartQuarter()
    Const CAP As String = "Start a Quarter"
    On Error GoTo Failed

    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim current As String
    current = DeckRegistry.GetDeckPeriod(pres)

    Dim typed As String
    typed = Trim(InputBox( _
        "This deck currently reports its period as:" & vbCrLf & vbCrLf & _
        "    " & IIf(current = "", "(none set)", current) & vbCrLf & vbCrLf & _
        IIf(current = "", _
            "Type the period it should now be, worded exactly as your slides and " & _
            "register word it.", _
            "Type the period it should now be, in the same form as the one above.") & _
        vbCrLf & vbCrLf & _
        "Nothing else changes: no slide is touched and no register row is created.", _
        CAP, current))
    If typed = "" Then Exit Sub

    If StrComp(typed, current, vbTextCompare) = 0 Then
        Say "The deck already reports " & current & ". Nothing changed.", vbInformation, CAP
        Exit Sub
    End If

    If MsgBox("Change this deck's period?" & vbCrLf & vbCrLf & _
              "    from:  " & IIf(current = "", "(none)", current) & vbCrLf & _
              "    to:    " & typed & vbCrLf & vbCrLf & _
              "After this, the drafting sheets and every sync will read " & typed & _
              " rows only.", vbYesNo + vbQuestion, CAP) <> vbYes Then Exit Sub

    ' VERIFIED AGAINST THE FILE'S OWN BYTES, and it saves the deck itself.
    '
    ' This used to write the property, read it straight back through the same
    ' Presentation object, and report success -- which reads PowerPoint's cache,
    ' so it confirmed its own write whether or not anything reached disk. On
    ' 2026-08-08 it reported success against a file that had not been touched
    ' for three days.
    Dim problem As String
    problem = DeckRegistry.SetDeckPeriodVerified(pres, typed, 4)

    Dim readBack As String
    readBack = typed

    If problem = "" Then
        ' USER-FACING TEXT SAYS WHAT TO DO, NOT WHAT THIS PROJECT HAS LEARNED.
        ' Rohan, 2026-08-08, on seeing "this project has lost it that way before"
        ' in a modal: the reasoning belongs in the code and the repo. A dialog
        ' narrating its own history reads as the tool talking about itself
        ' instead of telling you the next action.
        ' No "save the deck" instruction any more: the verified write saves it,
        ' and confirms the value in the saved file before saying this.
        Say "Deck period is now " & readBack & ", confirmed in the saved file." & vbCrLf & vbCrLf & _
               "STILL TO DO: the register needs rows for " & typed & ". Without " & _
               "them the drafting sheets will be empty." & vbCrLf & vbCrLf & _
               "'" & CommandBarUI.CAP_SET_UP_QUARTER & "' does this next -- no separate step. It copies the previous period's " & _
               "rows and stamps them " & typed & ", one row per slide.", vbInformation, CAP
    Else
        Say problem, vbCritical, CAP
    End If
    Exit Sub

Failed:
    Say "Could not set the period." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, CAP
End Sub

' ---------------------------------------------------------------------------
' BUTTON: Roll Forward
' ---------------------------------------------------------------------------
'
' ExcelOutput.RollForwardPeriod has existed since the wide model landed and had
' NO BUTTON -- so Start a Quarter told you to copy 43 rows by hand in Excel and
' retype the Quarter cell on every copy. The function that does it correctly was
' sitting right there, unreachable from the UI.
'
' THE DECK'S PERIOD IS THE DESTINATION, not something typed twice. You have just
' told the deck what quarter it is; asking again invites the two spellings to
' disagree, and this project has already established that two spellings of one
' quarter match nothing and report as a clean run of zero rows.
Public Sub RollForwardUI()
    Const CAP As String = "Roll Forward"
    On Error GoTo Failed

    Dim pres As Object, wb As Object, regWs As Object
    If Not Resolve(CAP, pres, wb, regWs) Then Exit Sub

    Dim toPeriod As String
    toPeriod = DeckRegistry.GetDeckPeriod(pres)
    If Trim$(toPeriod) = "" Then
        Say "This deck does not say what period it is." & vbCrLf & vbCrLf & _
               "Set the deck's quarter first -- '" & CommandBarUI.CAP_SET_UP_QUARTER & "' does that. Rolling forward copies rows INTO " & _
               "the period the deck declares, so it cannot run without one.", _
               vbExclamation, CAP
        Exit Sub
    End If

    ' DON'T ASK A QUESTION WHOSE EVERY ANSWER IS REFUSED.
    '
    ' RollForwardPeriod refuses when the destination already holds rows, because
    ' copying again would duplicate every project. That guard is right and stays.
    ' What was wrong is that it fired AFTER a modal and a free-text prompt: on the
    ' real deck, Q4F26 already had 43 rows, so "which period should they be copied
    ' FROM?" had no answer that could succeed. Checked here instead, and reported
    ' through Say -- which inside the Sync Now chain means a line in the run report
    ' rather than a dialog at all.
    Dim already As Long
    already = ExcelOutput.PeriodRowCount(regWs, toPeriod)
    If already > 0 Then
        Say toPeriod & " already has " & already & " row(s), so there is nothing to " & _
               "roll forward." & vbCrLf & vbCrLf & _
               "Rolling forward again would duplicate every project. Nothing was changed.", _
               vbInformation, CAP
        Exit Sub
    End If

    ' PICK THE SOURCE QUARTER BY CLICKING A ROW OF IT. 2026-08-14.
    '
    ' This was a free-text InputBox, and that was a DATA HAZARD rather than mere
    ' friction: periods are free text matched EXACTLY, so "Q3F26 " or "q3f26" or a
    ' quarter that simply is not in the register produces a clean run that copies
    ' nothing, reported as success. Reading the period out of a row the person
    ' pointed at makes a typo impossible by construction instead of by validation
    ' -- the value can only be one the register already holds.
    '
    ' The confirm MsgBox that followed is gone with it. Choosing the row IS the
    ' choice, the destination is stated below before anything is written, and the
    ' duplicate guard above has already refused the one case that could do harm.
    Dim qCol As Long
    qCol = ExcelOutput.QuarterColumn(regWs)
    If qCol = 0 Then
        Say "This register has no '" & ExcelOutput.QUARTER_HEADER & "' column, so it holds " & _
            "no periods to roll forward from.", vbExclamation, CAP
        Exit Sub
    End If

    ShowSheet wb, regWs.Name
    BringExcelToFront wb

    Dim srcPick As Object
    Dim fromPeriod As String

    Do
        Set srcPick = Nothing
        On Error Resume Next
        Set srcPick = wb.Application.InputBox( _
            Prompt:="Copy the register's rows INTO " & toPeriod & "." & vbCrLf & vbCrLf & _
                    "Click any cell in a row of the quarter you want to copy FROM, " & _
                    "then press OK." & vbCrLf & vbCrLf & _
                    "Those rows are duplicated and stamped " & toPeriod & ". The originals " & _
                    "are left exactly as they are, and no slide is touched.", _
            Title:=CAP, Type:=8)
        On Error GoTo 0

        If srcPick Is Nothing Then Exit Sub

        fromPeriod = ""
        On Error Resume Next
        If srcPick.Row > 1 Then fromPeriod = Trim$(CStr(regWs.Cells(srcPick.Row, qCol).Value))
        On Error GoTo 0

        If fromPeriod <> "" And StrComp(fromPeriod, toPeriod, vbTextCompare) <> 0 Then Exit Do

        If fromPeriod = "" Then
            Say "That row has no quarter in it." & vbCrLf & vbCrLf & _
                "Click a cell in one of the register's data rows.", vbExclamation, CAP
        Else
            Say "That row is already " & toPeriod & " -- the quarter you are copying INTO." & _
                vbCrLf & vbCrLf & "Click a row of the quarter you want to copy FROM.", _
                vbExclamation, CAP
        End If
    Loop

    ' FREEZE THE QUARTER BEING ROLLED OUT OF, BEFORE ANYTHING IS WRITTEN.
    '
    ' First half of file-per-quarter. It only CREATES a file, so it cannot damage
    ' the register, and a failure here is reported and does not stop the run --
    ' today's roll-forward only appends, so a missing archive loses nothing.
    ' WHEN THE PRUNE IS BUILT THIS MUST BECOME A HARD GATE: `If archiveProblem <> ""
    ' Then Say ... : Exit Sub`. Pruning the old period's rows without a verified
    ' archive is the destructive step the archive exists to make safe.
    Dim archiveProblem As String
    archiveProblem = WorkbookBridge.ArchiveWorkbookForPeriod(wb, fromPeriod)

    Dim outcome As String
    outcome = ExcelOutput.RollForwardPeriod(regWs, fromPeriod, toPeriod)

    If archiveProblem = "" Then
        outcome = outcome & vbCrLf & vbCrLf & fromPeriod & " archived as its own file " & _
            "beside the register."
    Else
        outcome = outcome & vbCrLf & vbCrLf & "ARCHIVE NOT WRITTEN -- the roll forward " & _
            "still ran, and nothing was lost, because rolling forward only adds rows." & _
            vbCrLf & archiveProblem
    End If

    ' Rolling forward writes a whole period's rows. Leaving them unsaved would
    ' lose an entire quarter's worth of register on a crash, silently.
    Dim rollSaveProblem As String
    rollSaveProblem = WorkbookBridge.SaveWorkbookVerified(wb)
    If rollSaveProblem = "" Then
        outcome = outcome & vbCrLf & vbCrLf & "Workbook saved."
    Else
        outcome = outcome & vbCrLf & vbCrLf & rollSaveProblem
    End If

    Say outcome, vbInformation, CAP
    Exit Sub

Failed:
    Say "Could not roll the period forward." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, CAP
End Sub

' ---------------------------------------------------------------------------
' BUTTON: Repoint Workbook
' ---------------------------------------------------------------------------
'
' DeckRegistry.RepointWorkbook has existed with no button. GetWorkbookPath's
' sibling fallback covers the common case -- a workbook that moved WITH its deck
' -- so this has never bitten. It becomes the first support call the moment the
' two are separated, with no self-service fix.
Public Sub RepointWorkbookUI()
    ' Named for the button that reaches it. This said CAP_SET_UP_QUARTER while its
    ' only caller was the sync path, and the moment it got its own button on
    ' 2026-08-15 the dialog started lying about which button had been pressed --
    ' the same defect already open against the apply confirmation. A caption
    ' constant belongs to the entry point, not to whichever one existed first.
    Const CAP As String = CommandBarUI.CAP_REPOINT_WORKBOOK
    On Error GoTo Failed

    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim current As String
    current = DeckRegistry.GetWorkbookPath(pres)

    Dim typed As String
    typed = Trim(InputBox( _
        "This deck is paired with:" & vbCrLf & vbCrLf & _
        "    " & IIf(current = "", "(nothing)", current) & vbCrLf & vbCrLf & _
        "Type the full path of the workbook it should use instead." & vbCrLf & vbCrLf & _
        "Easiest way to get it: open the workbook, File > Info, Copy path.", _
        CAP, current))
    If typed = "" Then Exit Sub

    ' VERIFIED AGAINST THE FILE'S OWN BYTES, and it saves the deck itself.
    '
    ' This used to call RepointWorkbook, read the value straight back through
    ' GetWorkbookPath on the same Presentation object, report success, and then ask
    ' the person to save the deck. Two faults in one: the read-back is PowerPoint's
    ' cache, so it confirmed the write regardless of what reached disk; and
    ' GetWorkbookPath falls back to a workbook beside the deck, so it could report a
    ' path that was never written at all.
    Dim problem As String
    problem = DeckRegistry.SetWorkbookPathVerified(pres, typed, 4)

    If problem <> "" Then
        Say problem, vbCritical, CAP
        Exit Sub
    End If

    ' AND STAMP THE OTHER END. Until 2026-08-14 a repoint updated the deck's path
    ' and left the workbook's DeckReference untouched, so the old workbook still
    ' claimed this deck and the new one claimed nothing. The GUID is the half of
    ' the pairing that survives a file being moved, which is exactly the failure
    ' this deck's OneDrive paths keep producing -- so it is the half that must not
    ' be allowed to go stale.
    Dim stampNote As String
    stampNote = DeckRegistry.StampPairing(pres, WorkbookBridge.OpenOrGetWorkbook(typed))
    If stampNote <> "" Then stampNote = vbCrLf & vbCrLf & _
        "The path was saved, but this deck's identity could not be written into " & _
        "the workbook:" & vbCrLf & vbCrLf & stampNote

    ' A PATH IS NOT THE PAIRING. The deck also stores DeckSyncType:<type> =
    ' slideID|worksheetName, and this never touched it -- so repointing at a
    ' workbook whose register sheet is named differently left LookupType asking
    ' for the old name, GetOrAddWorksheet CREATING it blank, and sync reporting
    ' success over zero rows. Reported as a repair that had half worked.
    '
    ' Rohan's test, 2026-08-09: "if a relink / repair / repoint button will
    ' genuinely work, fine, otherwise people just start from the top". So it has
    ' to say when it has not genuinely worked.
    Dim linkNote As String
    linkNote = WorksheetLinkProblem(pres, typed)

    If linkNote = "" Then
        Say "Paired workbook is now:" & vbCrLf & vbCrLf & typed & vbCrLf & vbCrLf & _
               "Confirmed in the saved file, and this deck's slide type still finds " & _
               "its sheet there. The deck has been saved for you." & stampNote, _
               IIf(stampNote = "", vbInformation, vbExclamation), CAP
    Else
        Say "The path was changed and confirmed on disk:" & vbCrLf & vbCrLf & typed & vbCrLf & vbCrLf & _
               "BUT THE PAIRING IS NOT REPAIRED." & vbCrLf & vbCrLf & linkNote & vbCrLf & vbCrLf & _
               "Do not sync until this is resolved -- a missing sheet gets created " & _
               "empty, and the run would report success having written nothing. " & _
               "Onboard this deck's slide type again to rebuild the link." & stampNote, _
               vbExclamation, CAP
    End If
    Exit Sub

Failed:
    Say "Could not repoint the workbook." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, CAP
End Sub

' Does every slide type this deck registers still find its worksheet in the
' workbook now paired to it? Returns "" when the link is sound, otherwise the
' sentence to show.
'
' OPENS THE WORKBOOK READ-ONLY VIA THE NORMAL BRIDGE and asks whether the sheet
' EXISTS -- never GetOrAddWorksheet, which would create the very sheet whose
' absence is the thing being tested, and turn the check into a repair that
' hides the fault it found.
Private Function WorksheetLinkProblem(pres As Object, workbookPath As String) As String
    WorksheetLinkProblem = ""

    Dim types() As String
    types = DeckRegistry.ListRegisteredTypes(pres)

    Dim lo As Long, hi As Long, hasTypes As Boolean
    On Error Resume Next
    lo = LBound(types): hi = UBound(types)
    hasTypes = (Err.Number = 0)
    On Error GoTo 0
    If Not hasTypes Then
        WorksheetLinkProblem = "This deck registers no slide type, so there is no sheet to find."
        Exit Function
    End If

    Dim wb As Object
    Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath)
    If wb Is Nothing Then
        WorksheetLinkProblem = "The new workbook could not be opened, so the link could not be checked."
        Exit Function
    End If

    Dim missing As String
    Dim i As Long
    For i = lo To hi
        Dim templateSld As Object
        Dim wsName As String
        If DeckRegistry.LookupType(pres, types(i), templateSld, wsName) Then
            If Not WorkbookBridge.WorksheetExists(wb, wsName) Then
                If missing <> "" Then missing = missing & vbCrLf
                missing = missing & "  '" & types(i) & "' expects a sheet named '" & wsName & "' -- not in this workbook."
            End If
        End If
    Next i

    If missing <> "" Then
        WorksheetLinkProblem = "This deck's slide type points at a sheet the new workbook " & _
            "does not have:" & vbCrLf & missing
    End If
End Function
