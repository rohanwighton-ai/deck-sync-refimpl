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
        MsgBox "This deck has no paired workbook yet." & vbCrLf & vbCrLf & _
               "Onboard a slide type first -- drafting reads the register, and " & _
               "there is no register until the deck and a workbook are paired.", _
               vbExclamation, caption
        Exit Function
    End If

    Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath)
    If wb Is Nothing Then
        MsgBox "Could not open the paired workbook at:" & vbCrLf & workbookPath, vbCritical, caption
        Exit Function
    End If

    Dim problem As String
    Set regWs = ResolveRegisterSheet(pres, wb, problem)
    If regWs Is Nothing Then
        MsgBox "Could not work out which register sheet to use." & vbCrLf & vbCrLf & _
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
' About_Body yet. Run '1. Drafting Sheets' first." Following that advice would
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

    AskForField = CanonicalFieldId(wb, Trim(InputBox(msg, caption)))
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
    Dim srcValidation As String
    srcValidation = Sources.ApplyPeriodValidation(srcWs, regWs)

    Dim fields As String
    fields = ProseFields(wb)
    If fields = "" Then
        MsgBox "No field on the 'Field Spec' sheet is marked Kind = Prose, so there " & _
               "is nothing to draft." & vbCrLf & vbCrLf & _
               "Static and Controlled fields are edited in the register, not drafted.", _
               vbInformation, CAP
        Exit Sub
    End If

    Dim period As String
    period = DeckRegistry.GetDeckPeriod(pres)
    If period = "" Then
        MsgBox "This deck does not declare a period, so the register cannot be " & _
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
        MsgBox "Cannot build drafting sheets from this register." & vbCrLf & vbCrLf & problem & _
               vbCrLf & vbCrLf & "Nothing was written.", vbExclamation, CAP
        Exit Sub
    End If

    Dim parts As Variant
    parts = Split(fields, ",")

    Dim report As String, firstSheet As String
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim fid As String
        fid = Trim(CStr(parts(i)))
        Dim sName As String
        sName = Drafting.DraftSheetNameFor(fid)
        If firstSheet = "" Then firstSheet = sName

        Dim ws As Object
        Set ws = WorkbookBridge.GetOrAddWorksheet(wb, sName)
        ' No cadence argument. Cadence was the long register's Quarter = ALL
        ' sentinel, and the wide sheet has no such row -- every row states one
        ' period. Omitting it means a rollover drops the drafting sheet's
        ' carried-over work, which is the safe direction and no longer the
        ' damaging one: RollForwardPeriod COPIES last period's rows, so the
        ' previous text arrives as this sheet's ORIGINAL column instead. The
        ' protection moved; it did not disappear.
        report = report & fid & ": " & Drafting.WriteDraftingSheet(ws, reg, fid, specWs, period) & vbCrLf
    Next i

    WorkbookBridge.WriteWorkbookIndex wb
    WorkbookBridge.FormatRegisterSheet regWs

    ' Dropdowns on the controlled fields, and a report of anything already in
    ' the register that the vocabulary does not allow.
    Dim valNote As String
    Dim outOfVocab As Long
    valNote = FieldSpec.ApplyControlledValidation(regWs, specWs, outOfVocab)

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

    Dim msg As String
    msg = period & " -- drafting sheets are ready." & vbCrLf & vbCrLf & _
          "Your wording goes in column D (SUBMIT). Type Y in column E to approve." & vbCrLf & _
          "Column C is what the slide says now. Copilot's prompt is in cell L2." & vbCrLf & _
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

    MsgBox msg, vbInformation, CAP
    Exit Sub

Failed:
    MsgBox "Could not refresh the drafting sheets." & vbCrLf & vbCrLf & _
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
Public Sub CopyAiDraftsToSubmit()
    Const CAP As String = "Copy AI to Submit"
    On Error GoTo Failed

    Dim pres As Object, wb As Object, regWs As Object
    If Not Resolve(CAP, pres, wb, regWs) Then Exit Sub

    Dim fieldId As String
    fieldId = ActiveDraftField(wb)
    If fieldId = "" Then fieldId = AskForField(CAP, wb)
    If fieldId = "" Then Exit Sub

    Dim sheetName As String
    sheetName = Drafting.DraftSheetNameFor(fieldId)
    If Not WorkbookBridge.WorksheetExists(wb, sheetName) Then
        ' "Run Drafting Sheets first" is the RIGHT advice for a real field with no
        ' sheet yet, and the WRONG advice for a typo -- following it would build a
        ' sheet for a FieldID that does not exist. So the two cases are separated.
        If StrComp(CanonicalFieldId(wb, fieldId), fieldId, vbBinaryCompare) <> 0 Or _
           Not FieldIsKnown(wb, fieldId) Then
            MsgBox "There is no field called '" & fieldId & "' on the Field Spec sheet." & vbCrLf & vbCrLf & _
                   "Check the spelling against the list, then try again. Nothing was created.", _
                   vbExclamation, CAP
        Else
            MsgBox "There is no drafting sheet for " & fieldId & " yet." & vbCrLf & vbCrLf & _
                   "Run '1. Drafting Sheets' first.", vbExclamation, CAP
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
    MsgBox note, vbInformation, CAP
    Exit Sub

Failed:
    MsgBox "Could not copy the AI drafts." & vbCrLf & vbCrLf & _
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
Public Sub PublishDraftsForField()
    Const CAP As String = "Publish Drafts"
    On Error GoTo Failed

    Dim pres As Object, wb As Object, regWs As Object
    If Not Resolve(CAP, pres, wb, regWs) Then Exit Sub

    Dim fieldId As String
    fieldId = ActiveDraftField(wb)
    If fieldId = "" Then fieldId = AskForField(CAP, wb)
    If fieldId = "" Then Exit Sub

    Dim sheetName As String
    sheetName = Drafting.DraftSheetNameFor(fieldId)
    If Not WorkbookBridge.WorksheetExists(wb, sheetName) Then
        MsgBox "There is no drafting sheet for " & fieldId & " yet." & vbCrLf & vbCrLf & _
               "Run '1. Drafting Sheets' first.", vbExclamation, CAP
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
        MsgBox "This deck does not declare a period, so there is no way to know " & _
               "which quarter's rows to publish into." & vbCrLf & vbCrLf & _
               "Run '0. Start a Quarter' first.", vbExclamation, CAP
        Exit Sub
    End If

    ' Said BEFORE the confirmation, not after the write. A warning that arrives
    ' with the result is a warning about something already done.
    ' Covers read-only AND macro-enabled: both mean this write will not stick,
    ' and both used to fail silently.
    Dim macroWarn As String
    macroWarn = WorkbookBridge.WriteBlockedReason(wb)
    If macroWarn <> "" Then macroWarn = macroWarn & vbCrLf & vbCrLf

    Dim preview As String
    preview = Drafting.PublishDrafts(ws, regWs, fieldId, period, True, srcWs)

    WorkbookBridge.WriteRunLog wb, "Publish " & fieldId & " -- preview", preview
    If MsgBox(RibbonUI.CapReport(macroWarn & preview) & vbCrLf & vbCrLf & _
              "Write these into the register for " & period & "?" & vbCrLf & vbCrLf & _
              "This does NOT touch any slide -- run Sync Now or Preview Sync " & _
              "afterwards to get them onto the deck.", _
              vbYesNo + vbQuestion, CAP) <> vbYes Then
        MsgBox "Nothing was published.", vbInformation, CAP
        Exit Sub
    End If

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
    ' This used to end with "Now run Preview Sync" -- the tool admitting the
    ' boundary was artificial. Nothing happens between writing Approved into the
    ' register and looking at what that would do to slides, so there is no
    ' decision for a button to mark. Offered rather than done, because it opens
    ' the deck and a person may not want that yet.
    WorkbookBridge.WriteRunLog wb, "Publish " & fieldId & " -- published", result
    If MsgBox(RibbonUI.CapReport(result) & vbCrLf & vbCrLf & _
              "Show what this would change on the slides?" & vbCrLf & _
              "Preview only -- reads the deck, writes nothing.", _
              vbYesNo + vbQuestion, CAP) = vbYes Then
        RibbonUI.SyncPreview
    End If
    Exit Sub

Failed:
    MsgBox "Could not publish." & vbCrLf & vbCrLf & _
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
        MsgBox "The deck already reports " & current & ". Nothing changed.", vbInformation, CAP
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
        MsgBox "Deck period is now " & readBack & ", confirmed in the saved file." & vbCrLf & vbCrLf & _
               "STILL TO DO: the register needs rows for " & typed & ". Without " & _
               "them the drafting sheets will be empty." & vbCrLf & vbCrLf & _
               "Press 'Roll Forward' on the toolbar. It copies the previous period's " & _
               "rows and stamps them " & typed & ", one row per slide.", vbInformation, CAP
    Else
        MsgBox problem, vbCritical, CAP
    End If
    Exit Sub

Failed:
    MsgBox "Could not set the period." & vbCrLf & vbCrLf & _
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
        MsgBox "This deck does not say what period it is." & vbCrLf & vbCrLf & _
               "Press '0. Start a Quarter' first. Roll Forward copies rows INTO " & _
               "the period the deck declares, so it cannot run without one.", _
               vbExclamation, CAP
        Exit Sub
    End If

    Dim fromPeriod As String
    fromPeriod = Trim(InputBox( _
        "Copy the register's rows INTO " & toPeriod & "." & vbCrLf & vbCrLf & _
        "Which period should they be copied FROM?" & vbCrLf & vbCrLf & _
        "Every project's row is duplicated and stamped " & toPeriod & ". Last " & _
        "period's rows are left exactly as they are -- nothing is moved or " & _
        "overwritten, and no slide is touched.", CAP))
    If fromPeriod = "" Then Exit Sub

    If MsgBox("Copy every " & fromPeriod & " row into " & toPeriod & "?", _
              vbQuestion + vbOKCancel, CAP) <> vbOK Then Exit Sub

    Dim outcome As String
    outcome = ExcelOutput.RollForwardPeriod(regWs, fromPeriod, toPeriod)

    ' Rolling forward writes a whole period's rows. Leaving them unsaved would
    ' lose an entire quarter's worth of register on a crash, silently.
    Dim rollSaveProblem As String
    rollSaveProblem = WorkbookBridge.SaveWorkbookVerified(wb)
    If rollSaveProblem = "" Then
        outcome = outcome & vbCrLf & vbCrLf & "Workbook saved."
    Else
        outcome = outcome & vbCrLf & vbCrLf & rollSaveProblem
    End If

    MsgBox outcome, vbInformation, CAP
    Exit Sub

Failed:
    MsgBox "Could not roll the period forward." & vbCrLf & vbCrLf & _
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

    If problem = "" Then
        MsgBox "Paired workbook is now:" & vbCrLf & vbCrLf & typed & vbCrLf & vbCrLf & _
               "Confirmed in the saved file -- the deck has been saved for you." & vbCrLf & vbCrLf & _
               "Keep the deck and its workbook in the SAME FOLDER -- then the pairing " & _
               "repairs itself when either moves.", vbInformation, CAP
    Else
        MsgBox problem, vbCritical, CAP
    End If
    Exit Sub

Failed:
    MsgBox "Could not repoint the workbook." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, CAP
End Sub
