Attribute VB_Name = "VerifyRealDeck"
Option Explicit

' One-off, read-only diagnostic -- NOT part of the shipped add-in (not in
' build_ppam.ps1's module list, not wired to CommandBarUI). Exists to
' answer one question after the 2026-07-26 "Linked: 0 / FAILED
' verification: 46" incident (see SPIKE_NOTES_BatchOnboardFlow.md): did the
' underlying Shape.Tags/text writes from that run actually land correctly
' on the real deck, even though CommitBatch's own verification step
' reported failure for every slide (root cause: InjectPrimitive.bas's
' pre-fix FindShapeByRoleTag couldn't see into groups -- the writes always
' worked, only the read-back that CommitBatch used to confirm them
' couldn't find what it had just written).
'
' Opens both real files READ-ONLY (Presentations.Open ReadOnly:=True,
' Workbooks.Open ReadOnly, and closes both without saving) and never
' writes to either. Cross-checks purely by reading: for every slide with a
' slide_type/instance_key (Resolve.ResolveSlideInstance, the same read
' path RunSync/DeckAdoption already trust), walks its shapes recursively
' (mirrors InjectPrimitive.bas's WalkForRoleTag) collecting role-tagged
' shapes' current text, and compares that against the matching row in the
' paired Data workbook, filtered to the deck's own current period
' (ExcelOutput.ReadSheetForDeckPeriod, the same read path RibbonUI.SyncNow
' already trusts) field-by-field.
'
' CORRECTED 2026-08-21: this used to call the UNFILTERED ExcelOutput.
' ReadSheet -- same bug as SyncRealDeck.bas/HiddenFixCheck.bas had ("first
' row wins" on a duplicate instance ID, and Q3F26 rows sit above Q4F26 rows
' in the register). The irony: this tool exists specifically to be trusted
' when something looks wrong after a sync, and a mother-hound audit found
' it would compare a live Q4F26 slide against its Q3F26 register row --
' either a false MISMATCH on any field that legitimately changed between
' quarters, or worse, a false "slideOk" if the two happened to coincide.
' A safety-net tool with this bug is more dangerous than no tool, because
' it's the one thing a person reaches for specifically when suspicious.
Public Function VerifyRealDeck(deckPath As String, workbookPath As String) As String
    Dim report As String
    Dim detail As String

    Dim pres As Object
    Set pres = Application.Presentations.Open(deckPath, msoTrue, msoFalse, msoFalse)

    Dim xl As Object, wb As Object, ws As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(workbookPath, 0, True) ' positional: UpdateLinks:=0, ReadOnly:=True -- named args aren't reliable on a late-bound Object
    Set ws = WorkbookBridge.RegisterOrFirstDataSheet(wb)

    Dim deckPeriod As String
    deckPeriod = DeckRegistry.GetDeckPeriod(pres)

    Dim sheet As ExcelOutput.Sheet
    Dim readProblem As String
    sheet = ExcelOutput.ReadSheetForDeckPeriod(ws, deckPeriod, readProblem)
    If readProblem <> "" Then
        VerifyRealDeck = "REFUSED: could not read the register for period '" & deckPeriod & "': " & readProblem
        wb.Close False
        xl.Quit
        pres.Close
        Exit Function
    End If

    Dim slidesChecked As Long, slidesOk As Long, slidesMissingTags As Long, slidesNoWorkbookRow As Long
    Dim fieldsChecked As Long, fieldsMissingShape As Long, fieldsMismatch As Long, fieldsSkippedByDesign As Long

    Dim sld As Object
    For Each sld In pres.Slides
        slidesChecked = slidesChecked + 1

        Dim inst As Resolve.SlideInstance
        inst = Resolve.ResolveSlideInstance(sld)

        If Not inst.HasTypeTag Or Not inst.HasInstanceKey Then
            slidesMissingTags = slidesMissingTags + 1
            detail = detail & "Slide " & sld.SlideIndex & ": MISSING slide_type and/or instance_key tag" & vbCrLf
        ElseIf Not sheet.Rows.Exists(inst.InstanceKey) Then
            slidesNoWorkbookRow = slidesNoWorkbookRow + 1
            detail = detail & "Slide " & sld.SlideIndex & " (instance_key=" & inst.InstanceKey & "): no matching row in workbook" & vbCrLf
        Else
            Dim roleTags As Object
            Set roleTags = CreateObject("Scripting.Dictionary")
            CollectRoleTags sld.Shapes, roleTags

            Dim rowValues As Object
            Set rowValues = sheet.Rows(inst.InstanceKey)

            Dim slideOk As Boolean
            slideOk = True

            Dim fieldName As Variant
            For Each fieldName In rowValues.Keys
                fieldsChecked = fieldsChecked + 1
                If IsExpectedToCarryNoOwnRoleTag(CStr(fieldName)) Then
                    ' NOT A GAP, A WRONG QUESTION -- mother-hound audit,
                    ' 2026-08-22. This register column was never meant to
                    ' answer "does any shape carry this exact role tag",
                    ' either because MilestoneDevice.IsColumnForThisDevice
                    ' says so (MS1_LABEL etc. -- addressed by shape NAME
                    ' inside the milestone group, not a role tag; see that
                    ' function's own header, which already fixed the
                    ' identical false-alarm class once for
                    ' FieldWiring.ScanFieldWiring: "Reported as '21 fields on
                    ' the register that no slide carries' on every run before
                    ' this existed") or because it is a source-only column
                    ' for a Derived field (PROJECT_STATUS/SCHEDULE_STATUS ->
                    ' STATUS_BADGE, FieldSpec.bas's own STATUS_BADGE row:
                    ' "computes it directly from PROJECT_STATUS and
                    ' SCHEDULE_STATUS at sync time" -- the badge is tagged
                    ' and checked; its two register-only sources never are).
                    ' 624 of this tool's first-ever 686 findings (91%) were
                    ' this exact false alarm, on all 43 real slides, before
                    ' this exclusion existed.
                    fieldsSkippedByDesign = fieldsSkippedByDesign + 1
                ElseIf Not roleTags.Exists(CStr(fieldName)) Then
                    fieldsMissingShape = fieldsMissingShape + 1
                    slideOk = False
                    detail = detail & "Slide " & sld.SlideIndex & " (" & inst.InstanceKey & "): field '" & fieldName & "' -- no shape on the slide carries this role tag" & vbCrLf
                ElseIf CStr(roleTags(fieldName)) <> CStr(rowValues(fieldName)) Then
                    fieldsMismatch = fieldsMismatch + 1
                    slideOk = False
                    detail = detail & "Slide " & sld.SlideIndex & " (" & inst.InstanceKey & "): field '" & fieldName & "' -- shape text does not match workbook value" & vbCrLf
                End If
            Next fieldName

            If slideOk Then slidesOk = slidesOk + 1
        End If
    Next sld

    wb.Close False
    xl.Quit
    pres.Close

    report = "=== Real Deck Verification Report ===" & vbCrLf & _
        "Deck: " & deckPath & vbCrLf & _
        "Workbook: " & workbookPath & vbCrLf & _
        "Run at: " & Now & vbCrLf & vbCrLf & _
        "Slides checked: " & slidesChecked & vbCrLf & _
        "Slides fully OK (all fields present + matching): " & slidesOk & vbCrLf & _
        "Slides missing slide_type/instance_key tags: " & slidesMissingTags & vbCrLf & _
        "Slides with no matching workbook row: " & slidesNoWorkbookRow & vbCrLf & _
        "Fields checked (across all OK-tagged slides): " & fieldsChecked & vbCrLf & _
        "Fields skipped, no role tag expected by design (milestone-device slots, Derived-field sources): " & fieldsSkippedByDesign & vbCrLf & _
        "Fields with no tagged shape found: " & fieldsMissingShape & vbCrLf & _
        "Fields with a text/value mismatch: " & fieldsMismatch & vbCrLf & vbCrLf & _
        "--- Per-slide detail (only non-OK slides/fields listed) ---" & vbCrLf & detail

    VerifyRealDeck = report
End Function

' A register column this tool must not ask "does any shape carry this exact
' role tag" about, because the answer is always no by design, not by defect.
' Two closed categories, both documented at the one call site above:
'   - MilestoneDevice.IsColumnForThisDevice: MS<n>_LABEL/_DATE/_DONE, owned
'     by the milestone group and addressed by shape name, never role-tagged
'     individually.
'   - Source-only columns for a Kind=Derived field (FieldSpec.bas): the
'     Derived field itself IS role-tagged and IS checked (e.g. STATUS_BADGE);
'     only its raw register inputs are exempt. Hardcoded here, not derived
'     from FieldSpec, because FieldSpec's Kind=Derived column names the
'     OUTPUT field, not its inputs, and there is no structured "feeds into"
'     column to read this from -- if a new Derived field with register-only
'     sources is ever added, its sources need adding here too.
Private Function IsExpectedToCarryNoOwnRoleTag(fieldName As String) As Boolean
    If MilestoneDevice.IsColumnForThisDevice(fieldName) Then
        IsExpectedToCarryNoOwnRoleTag = True
        Exit Function
    End If

    Select Case UCase(Trim(fieldName))
        Case "PROJECT_STATUS", "SCHEDULE_STATUS" ' feed STATUS_BADGE only
            IsExpectedToCarryNoOwnRoleTag = True
    End Select
End Function

' Same recursion shape as InjectPrimitive.bas's WalkForRoleTag, but
' collects every role-tagged shape's text instead of searching for one
' specific tag -- this diagnostic needs the whole set per slide, not a
' single lookup.
Private Sub CollectRoleTags(shapesColl As Object, ByRef roleTags As Object)
    Dim shp As Object
    For Each shp In shapesColl
        If shp.Type = msoGroup Then
            CollectRoleTags shp.GroupItems, roleTags
        Else
            Dim r As String
            r = ""
            On Error Resume Next
            r = shp.Tags("role")
            On Error GoTo 0
            If r <> "" And shp.HasTextFrame Then
                roleTags(r) = shp.TextFrame.TextRange.Text
            End If
        End If
    Next shp
End Sub
