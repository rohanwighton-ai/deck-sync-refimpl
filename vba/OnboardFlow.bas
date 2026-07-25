Attribute VB_Name = "OnboardFlow"
Option Explicit

' Implements specs/ribbon-ui.md's "Onboard New Slide Type" flow, following
' onboard-slide-type.md's six steps exactly (per ribbon-ui.md's own
' requirement). The only genuinely new logic is the phase-gate review this
' module builds -- Discovery, Onboarding.ConfirmFieldMatch, ExcelOutput's
' CreateSheet/UpsertRow, and InjectPrimitive (the verify-the-link step) are
' all existing, already-tested calls, unchanged.
'
' InputBox-based, not a UserForm ListBox -- same posture as
' ResolveFields.bas, for the same reason (see its header comment): no
' established, real-Office-verified precedent yet for authoring a .frm/.frx
' pair in this project. A real Office session is available this pass
' (2026-07-26, WSL host), so this constraint could be lifted, but building
' and proving a first-ever UserForm was deliberately kept out of this
' already-large pass; flagged as a real, deferred upgrade, not an oversight.
'
' Design call this module makes that no spec pins down: whether the example
' slide, once tagged, ALSO becomes the type's first live instance (carries
' slide_type + instance_key, not just field-level role tags) or stays
' purely structural. Chosen: it becomes instance #1 too. Reasoning: Step 5
' requires a seed Data-sheet row, and ExcelOutput.ReadSheet deliberately
' excludes blank-Instance-ID rows (DeckAdoption.bas's own finding) -- a
' seed row with no instance_key would sit permanently invisible to
' RunRoutineSync's normal case-3 dispatch, needing new "keyless row"
' handling `DeckAdoption.bas` already had to invent for a different reason.
' Giving the template its own instance_key avoids inventing that a second
' time: it becomes a completely ordinary tagged instance like any other,
' and RunSync.GatherInstances/DuplicateAndTag need no special-casing to
' treat it correctly (a template that happens to also be instance #1 is not
' a contradiction anywhere else in this codebase).

Public Type PendingField
    FieldCandidate As Candidate
    Shape As Object
    ProposedName As String
    HarvestedValue As String
    Excluded As Boolean
    IsPeriodKey As Boolean
End Type

Public Type OnboardingResult
    Ok As Boolean
    SlideType As String
    InstanceKey As String
    FieldCount As Long
    VerifyReport As String
    ErrorMessage As String
End Type

' ---------------------------------------------------------------------
' Step 1-2: duplicate + discover -- thin, both calls already exist
' ---------------------------------------------------------------------

' `exampleSld` must already be the working copy (Step 1's duplication is
' the caller's job -- PromptOnboardNewSlideType below does it before calling
' this -- kept separate so this stays testable against a slide TestRunner.bas
' creates directly, no live Selection needed).
Public Function PlanOnboarding(exampleSld As Object) As PendingField()
    Dim allCandidates() As Candidate
    Dim allShapes() As Object
    allCandidates = Discovery.DiscoverSlideWithShapes(exampleSld, allShapes)

    Dim results() As PendingField
    Dim n As Long
    n = 0

    Dim i As Long, lo As Long, hi As Long, hasCandidates As Boolean
    On Error Resume Next
    lo = LBound(allCandidates)
    hi = UBound(allCandidates)
    hasCandidates = (Err.Number = 0)
    On Error GoTo 0

    If hasCandidates Then
        For i = lo To hi
            If Onboarding.IsCandidateField(allCandidates(i)) Then
                n = n + 1
                ReDim Preserve results(1 To n)
                results(n).FieldCandidate = allCandidates(i)
                Set results(n).Shape = allShapes(i)
                results(n).ProposedName = SuggestFieldName(allShapes(i), n)
                If allCandidates(i).HasText Then
                    results(n).HarvestedValue = allShapes(i).TextFrame.TextRange.Text
                Else
                    results(n).HarvestedValue = ""
                End If
                results(n).Excluded = False
                results(n).IsPeriodKey = False
            End If
        Next i
    End If

    PlanOnboarding = results
End Function

' Shape.Name if it already follows the ph_ convention (input-contract.md's
' unique_named_shapes rule), lower-cased and de-spaced; otherwise a
' positional fallback ("ph_field1", "ph_field2", ...) -- always a proposal
' the Review step can rename, never assumed final.
Private Function SuggestFieldName(shp As Object, ordinal As Long) As String
    Dim rawName As String
    rawName = LCase(Trim(shp.Name))
    rawName = Replace(rawName, " ", "_")

    If Left(rawName, 3) = "ph_" And Len(rawName) > 3 Then
        SuggestFieldName = rawName
    Else
        SuggestFieldName = "ph_field" & ordinal
    End If
End Function

' ---------------------------------------------------------------------
' Step 4: the review phase gate -- pure logic pieces, exercised directly by
' TestRunner.bas (mirrors ResolveFields.bas's split).
' ---------------------------------------------------------------------

Public Function BuildFieldReviewPrompt(fields() As PendingField, index As Long) As String
    Dim f As PendingField
    f = fields(index)

    Dim lo As Long, hi As Long
    lo = LBound(fields): hi = UBound(fields)

    Dim s As String
    s = "Field " & index & " of " & hi & " -- shape '" & f.Shape.Name & "'" & vbCrLf
    s = s & "Harvested value: '" & f.HarvestedValue & "'" & vbCrLf & vbCrLf
    s = s & "Enter the field name to use, or:" & vbCrLf
    s = s & "  (leave blank to accept '" & f.ProposedName & "')" & vbCrLf
    s = s & "  SKIP to exclude this field"
    BuildFieldReviewPrompt = s
End Function

' Applies one review answer to `fields(index)` in place. "" keeps the
' proposed name; "SKIP" (case-insensitive) excludes the field; anything
' else becomes the field's name verbatim (Review's whole point is letting a
' human override discovery's guess, so no validation against a naming
' convention is imposed here).
Public Sub ApplyFieldReviewAnswer(ByRef fields() As PendingField, index As Long, answer As String)
    If UCase(Trim(answer)) = "SKIP" Then
        fields(index).Excluded = True
    ElseIf Trim(answer) <> "" Then
        fields(index).ProposedName = Trim(answer)
    End If
End Sub

' Builds the period-key picker prompt, listing only non-excluded fields
' (excluding a field and then marking it the period-key would be
' self-contradictory, so the excluded ones are never offered).
Public Function BuildPeriodKeyPrompt(fields() As PendingField) As String
    Dim s As String
    s = "Which field is this type's period-key (e.g. quarter/date)?" & vbCrLf
    s = s & "Enter its number, or leave blank if this type is evergreen (never versions by period):" & vbCrLf

    Dim lo As Long, hi As Long, i As Long
    lo = LBound(fields): hi = UBound(fields)
    For i = lo To hi
        If Not fields(i).Excluded Then
            s = s & i & ") " & fields(i).ProposedName & " (current: '" & fields(i).HarvestedValue & "')" & vbCrLf
        End If
    Next i

    BuildPeriodKeyPrompt = s
End Function

' Sets IsPeriodKey on exactly the field `answer` (a 1-based index string)
' names, clearing it on every other field first. Returns False (no field
' marked) for a blank answer -- the type is evergreen -- or for an answer
' that names an excluded or out-of-range field, which the caller should
' report as an error rather than silently ignore.
Public Function ApplyPeriodKeyAnswer(ByRef fields() As PendingField, answer As String) As Boolean
    Dim lo As Long, hi As Long, i As Long
    lo = LBound(fields): hi = UBound(fields)
    For i = lo To hi
        fields(i).IsPeriodKey = False
    Next i

    If Trim(answer) = "" Then
        ApplyPeriodKeyAnswer = False
        Exit Function
    End If

    If Not IsNumeric(answer) Then
        ApplyPeriodKeyAnswer = False
        Exit Function
    End If

    Dim idx As Long
    idx = CLng(answer)
    If idx < lo Or idx > hi Then
        ApplyPeriodKeyAnswer = False
        Exit Function
    End If
    If fields(idx).Excluded Then
        ApplyPeriodKeyAnswer = False
        Exit Function
    End If

    fields(idx).IsPeriodKey = True
    ApplyPeriodKeyAnswer = True
End Function

' The seed instance_key: the period-key field's harvested value if one was
' marked (sanitized -- Slide.Tags values are plain strings, but a stable,
' readable key beats an opaque one for a human later inspecting tags), or a
' fixed key for an evergreen type, which by definition only ever has one
' instance.
Public Function DeriveSeedInstanceKey(fields() As PendingField) As String
    Dim lo As Long, hi As Long, i As Long
    lo = LBound(fields): hi = UBound(fields)
    For i = lo To hi
        If fields(i).IsPeriodKey Then
            Dim raw As String
            raw = Trim(fields(i).HarvestedValue)
            raw = Replace(raw, " ", "-")
            If raw = "" Then raw = "period-1"
            DeriveSeedInstanceKey = raw
            Exit Function
        End If
    Next i

    DeriveSeedInstanceKey = "evergreen"
End Function

' ---------------------------------------------------------------------
' Step 5: commit -- the only step that writes anything
' ---------------------------------------------------------------------

' Writes role tags (non-excluded fields only), tags `exampleSld` as both
' the type's template and its first live instance (see header comment for
' why), registers it in DeckRegistry, builds/extends the Data-sheet columns
' via UpsertRow, and seeds the row. Does not run Step 6's verification --
' VerifyOnboarding below is separate so a caller can report each half
' distinctly, matching ribbon-ui.md's "reports pass/fail per field."
Public Function CommitOnboarding(pres As Object, exampleSld As Object, fields() As PendingField, slideType As String, ws As Object, workbookDeckId As String) As OnboardingResult
    Dim result As OnboardingResult
    result.SlideType = slideType

    Dim lo As Long, hi As Long, i As Long
    lo = LBound(fields): hi = UBound(fields)

    Dim harvested As Object
    Set harvested = CreateObject("Scripting.Dictionary")

    For i = lo To hi
        If Not fields(i).Excluded Then
            Onboarding.ConfirmFieldMatch fields(i).Shape, fields(i).ProposedName
            harvested(fields(i).ProposedName) = fields(i).HarvestedValue
            result.FieldCount = result.FieldCount + 1
        End If
    Next i

    Dim instanceKey As String
    instanceKey = DeriveSeedInstanceKey(fields)
    result.InstanceKey = instanceKey

    exampleSld.Tags.Add "slide_type", slideType
    exampleSld.Tags.Add "instance_key", instanceKey

    DeckRegistry.RegisterType pres, slideType, exampleSld, ws.Name

    If IsEmpty(ws.Cells(1, 1).Value) Then
        ExcelOutput.CreateSheet ws, workbookDeckId
    End If
    ExcelOutput.UpsertRow ws, instanceKey, harvested

    result.Ok = True
    CommitOnboarding = result
End Function

' Step 6: run InjectPrimitive on every committed field against its own
' harvested value -- proving the tag actually resolves to the shape that
' fed it, not just that a tag and a cell happen to currently agree. Per the
' workflow doc: every field must hit the no-op path (Verified=True,
' Written=False); anything else is a bug in this onboarding pass, not
' something a later sync silently corrects.
Public Function VerifyOnboarding(exampleSld As Object, fields() As PendingField) As String
    Dim report As String
    Dim lo As Long, hi As Long, i As Long
    lo = LBound(fields): hi = UBound(fields)

    Dim allOk As Boolean
    allOk = True

    For i = lo To hi
        If Not fields(i).Excluded Then
            Dim r As InjectResult
            r = InjectPrimitive.InjectPrimitive(exampleSld, fields(i).ProposedName, fields(i).HarvestedValue)
            If r.Verified And Not r.Written Then
                report = report & "  OK    " & fields(i).ProposedName & vbCrLf
            Else
                allOk = False
                report = report & "  FAIL  " & fields(i).ProposedName & ": " & r.ErrorMessage & vbCrLf
            End If
        End If
    Next i

    If allOk Then
        VerifyOnboarding = "All fields verified (no-op path):" & vbCrLf & report
    Else
        VerifyOnboarding = "VERIFICATION FAILED -- do not treat this type as onboarded:" & vbCrLf & report
    End If
End Function

' ---------------------------------------------------------------------
' Ribbon entry point -- the only interactive piece; not automatable
' (TestRunner.bas cannot drive a live InputBox chain, same constraint as
' every other UserForm/InputBox flow in this project). See
' SPIKE_NOTES_OnboardFlow.md for the manual verification recipe.
' ---------------------------------------------------------------------

Public Function PromptOnboardNewSlideType() As String
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim sel As Object
    Set sel = Application.ActiveWindow.Selection
    If sel.Type <> ppSelectionSlides Or sel.SlideRange.count < 1 Then
        PromptOnboardNewSlideType = "Select 1-2 example slides first."
        Exit Function
    End If

    Dim original As Object
    Set original = sel.SlideRange(1)

    Dim slideType As String
    slideType = InputBox("Name for this new slide type (e.g. 'quarterly-update'):", "Onboard New Slide Type")
    If Trim(slideType) = "" Then
        PromptOnboardNewSlideType = "Cancelled -- no type name given."
        Exit Function
    End If

    ' Step 1: duplicate immediately; every remaining step works only on the
    ' copy, never the original the user had selected.
    Dim dupRange As Object
    Set dupRange = original.Duplicate
    Dim workingCopy As Object
    Set workingCopy = dupRange(1)

    Dim fields() As PendingField
    fields = PlanOnboarding(workingCopy)

    Dim lo As Long, hi As Long, hasFields As Boolean
    On Error Resume Next
    lo = LBound(fields): hi = UBound(fields)
    hasFields = (Err.Number = 0)
    On Error GoTo 0

    If Not hasFields Then
        PromptOnboardNewSlideType = "No candidate fields found on the selected slide (no text or picture shapes) -- nothing to onboard."
        Exit Function
    End If

    Dim i As Long
    For i = lo To hi
        Dim answer As String
        answer = InputBox(BuildFieldReviewPrompt(fields, i), "Onboard New Slide Type -- Review Field " & i & " of " & hi)
        ApplyFieldReviewAnswer fields, i, answer
    Next i

    Dim pkAnswer As String
    pkAnswer = InputBox(BuildPeriodKeyPrompt(fields), "Onboard New Slide Type -- Period Key")
    ApplyPeriodKeyAnswer fields, pkAnswer

    Dim confirmMsg As String
    confirmMsg = "Onboard type '" & slideType & "' with these fields?" & vbCrLf
    For i = lo To hi
        If Not fields(i).Excluded Then
            confirmMsg = confirmMsg & "  " & fields(i).ProposedName & IIf(fields(i).IsPeriodKey, " (period-key)", "") & " = '" & fields(i).HarvestedValue & "'" & vbCrLf
        End If
    Next i
    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Confirm Onboarding") <> vbYes Then
        PromptOnboardNewSlideType = "Cancelled at confirmation -- nothing written."
        Exit Function
    End If

    ' Establish (or reuse) the deck-workbook pairing before committing --
    ' Step 5 needs a live worksheet to write into.
    Dim workbookPath As String
    workbookPath = DeckRegistry.GetWorkbookPath(pres)
    Dim wb As Object
    If workbookPath = "" Then
        workbookPath = InputBox("This deck has no paired workbook yet." & vbCrLf & "Enter a full path for the new Data workbook (.xlsx):", "Onboard New Slide Type -- Pair Workbook")
        If Trim(workbookPath) = "" Then
            PromptOnboardNewSlideType = "Cancelled -- no workbook path given."
            Exit Function
        End If
        Set wb = WorkbookBridge.CreateWorkbook(workbookPath)
        If wb Is Nothing Then
            PromptOnboardNewSlideType = "Could not create workbook at: " & workbookPath
            Exit Function
        End If
        DeckRegistry.SetWorkbookPath pres, workbookPath
    Else
        Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath)
        If wb Is Nothing Then
            PromptOnboardNewSlideType = "Could not open the paired workbook at: " & workbookPath
            Exit Function
        End If
    End If

    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(wb, WorkbookBridge.SanitizeSheetName(slideType))

    Dim deckId As String
    deckId = DeckRegistry.GetOrCreateDeckId(pres)

    Dim commitResult As OnboardingResult
    commitResult = CommitOnboarding(pres, workingCopy, fields, slideType, ws, deckId)

    Dim verifyReport As String
    verifyReport = VerifyOnboarding(workingCopy, fields)

    PromptOnboardNewSlideType = "Onboarded '" & slideType & "' (" & commitResult.FieldCount & " field(s), seed instance '" & commitResult.InstanceKey & "')." & vbCrLf & vbCrLf & verifyReport
End Function
