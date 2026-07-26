Attribute VB_Name = "BatchOnboardFlow"
Option Explicit

' Onboards a new slide type AND bulk-links a whole batch of its already-
' existing instances in one pass, driven by an editable Excel review grid
' instead of InputBox chains -- built 2026-07-26 specifically because a
' real, richly-designed deck (see test-fixtures/SOURCE.md's
' crc-real-deck-redacted.pptx entry) discovers 60-90 raw candidate shapes
' per slide, and walking that many one InputBox at a time (OnboardFlow.bas's
' existing flow) is unusable. Also closes specs/onboarding.md's own
' documented-but-never-built "boilerplate-vs-varying pre-filter" gap, in a
' different shape than that spec sketched: rather than a silent geometric
' clustering filter, this surfaces the guess AND lets a human freely
' override it in a real, editable table (Rohan's own framing, 2026-07-26 --
' more precise than an algorithm guessing alone, and Excel is a far more
' capable editing surface than anything hand-built here would be).
'
' Reuses proven primitives throughout, no new matching/scoring logic:
' Discovery.DiscoverSlideWithShapes (candidate enumeration), Matching.Match
' (geometry-based correspondence -- exactly its tier-2 scoring, just run
' template-vs-every-other-slide instead of template-vs-one-new-instance),
' Onboarding.ConfirmFieldMatch (the tag write), ExcelOutput.CreateSheet/
' UpsertRow, InjectPrimitive (verify-the-link), DeckRegistry, WorkbookBridge.
' AdoptFlow.ValidateAdoptionSelection is reused directly for selection
' validation/deck-order sorting -- identical requirement, no need to
' duplicate it.
'
' Per-field/per-slide correspondence and harvested values (Object
' shapes/String text -- never a UDT) are stored in Scripting.Dictionary,
' keyed by a composite "fieldIndex|slideIndex" string. A REAL, confirmed
' restriction found the hard way this pass (2026-07-26, needed a Fable-
' model agent's live-tested diagnosis after a long confusing debugging
' session -- see SPIKE_NOTES_BatchOnboardFlow.md for the full account):
' a Candidate() array (Discovery.bas's Public Type) can NEVER be stored as
' a Dictionary/Variant value -- VBA's compiler rejects it outright ("Only
' user-defined types defined in public object modules can be coerced to or
' from a variant or passed to late-bound functions"), since Candidate is
' declared in a standard module, not a class. This is DIFFERENT from
' DeckAdoption.bas's own AdoptionSlidePlan comment (which flags a UDT
' *containing* an array member, inside an array-of-that-UDT, as untested
' territory) -- this restriction is broader: ANY UDT array, even a plain
' local variable, can't be boxed into a Variant-backed container
' (Dictionary, Collection item, Variant array) at all. BuildBatchPlan below
' works around this by flattening all other-slides' Candidate()/Object()/
' Boolean() data into plain, properly-typed arrays (never boxed into a
' Dictionary) instead of caching them per-slide in one.
'
' Unlike OnboardFlow.bas, this flow does NOT duplicate the template slide
' first -- that invariant exists for onboard-slide-type.md's "hand over an
' example, it becomes the master template for future creates" authoring
' workflow. This flow's job is the opposite: retroactively tag REAL,
' already-existing slides in place (the same semantics DeckAdoption.bas
' already uses, which also never duplicates) -- every selected slide,
' including whichever one is chosen as the template, is tagged where it
' already sits.

' Every module-level declaration (Const, Type) in this module lives here,
' above ALL Functions/Subs -- NOT stylistic. Real, confirmed VBA compiler
' quirk (2026-07-26, found by bisecting minimal repros after a long
' debugging session): if a Public Function/Sub appears textually BEFORE a
' Public Type in the same standard module, that Type fails to resolve when
' referenced from a DIFFERENT module ("Compile error: User-defined type
' not defined"), even though the function itself compiles fine and
' same-module usage of the Type works. The COL_* Consts hit a second,
' apparently-related symptom of the same underlying rule when left
' declared after BuildBatchPlan ("Only comments may appear after End Sub,
' End Function, or End Property") -- moving them up here too resolved it.
' Confirmed via multiple minimal, isolated repros (not guesses): every
' module-level declaration in this project should now go above every
' Function/Sub, not just Types. See SPIKE_NOTES_BatchOnboardFlow.md and
' AGENTS.md's Known Patterns for the full account.
Private Const FIELD_KEY_SEP As String = "|"
Private Const COL_FIELD_ID As Long = 1
Private Const COL_FIELD_NAME As Long = 2
Private Const COL_SUGGESTED As Long = 3
Private Const COL_INCLUDE As Long = 4
Private Const COL_TEMPLATE_VALUE As Long = 5
Private Const COL_SAMPLE_OTHER_VALUES As Long = 6
Public Type BatchOnboardPlan
    FieldCount As Long
    FieldNames As Object          ' Dictionary: fieldIndex -> String (proposed, editable)
    FieldTemplateShapes As Object ' Dictionary: fieldIndex -> Object (the template's own shape)
    FieldSuggestIdentical As Object ' Dictionary: fieldIndex -> Boolean
    FieldInclude As Object        ' Dictionary: fieldIndex -> Boolean (set after grid review)
    Correspondence As Object      ' Dictionary: "field|slide" -> Object (the shape), absent if not found
    HarvestedText As Object       ' Dictionary: "field|slide" -> String
End Type

Public Type BatchCommitResult
    LinkedCount As Long
    SkippedCount As Long
    FailedVerificationCount As Long
    FailedVerificationLabels() As String
End Type

' ---------------------------------------------------------------------
' Pure-ish logic: classification -- exercised directly by TestRunner.bas
' ---------------------------------------------------------------------

' True if every non-empty value in `values` (a Collection of String) is
' identical -- the batch-wide "identical everywhere" signal that makes this
' a real, computed suggestion rather than a blind guess (unlike onboarding
' a type from a single example, where no such batch exists to diff against
' at all). An empty/1-item collection is trivially "identical" (nothing to
' differ from) -- callers should not call this for a field found on zero
' slides in the first place.
Public Function AllValuesIdentical(values As Collection) As Boolean
    If values.count <= 1 Then
        AllValuesIdentical = True
        Exit Function
    End If

    Dim first As String
    first = CStr(values(1))

    Dim v As Variant
    For Each v In values
        If CStr(v) <> first Then
            AllValuesIdentical = False
            Exit Function
        End If
    Next v

    AllValuesIdentical = True
End Function

' Proposes a default field name the same way OnboardFlow.bas's
' SuggestFieldName does (reuses the shape's own ph_ name if it already has
' one, else a positional fallback) -- kept here rather than promoting that
' Private function, since this module's ordinal is a field index across a
' filtered candidate list, not literally the same concept.
Public Function SuggestBatchFieldName(shp As Object, ordinal As Long) As String
    Dim rawName As String
    rawName = LCase(Trim(shp.Name))
    rawName = Replace(rawName, " ", "_")

    If Left(rawName, 3) = "ph_" And Len(rawName) > 3 Then
        SuggestBatchFieldName = rawName
    Else
        SuggestBatchFieldName = "ph_field" & ordinal
    End If
End Function

' ---------------------------------------------------------------------
' Correspondence + harvesting across the batch
' ---------------------------------------------------------------------

' Everything computed before the Excel grid is shown, keyed by field index
' (1-based, position in the filtered template candidate list) and by
' composite "fieldIndex|slideIndex" (slideIndex 0 = the template itself,
' 1..N = otherSlides() in the same order the caller supplies).

Private Function FieldSlideKey(fieldIdx As Long, slideIdx As Long) As String
    FieldSlideKey = CStr(fieldIdx) & FIELD_KEY_SEP & CStr(slideIdx)
End Function

' Builds the full BatchOnboardPlan: discovers the template's candidate fields,
' then for each one, finds its best-corresponding shape on every other
' slide via Matching.Match (pure tier-2 geometry scoring -- nothing is
' tagged yet, so tier-1 trust-the-tag never applies here) and harvests its
' current text. A shape already claimed by an earlier field on the same
' slide is excluded from later fields' candidate pool on that slide, so two
' fields can never both grab the same shape just because they scored
' similarly.
Public Function BuildBatchPlan(templateSld As Object, otherSlides() As Object) As BatchOnboardPlan
    Dim plan As BatchOnboardPlan
    Set plan.FieldNames = CreateObject("Scripting.Dictionary")
    Set plan.FieldTemplateShapes = CreateObject("Scripting.Dictionary")
    Set plan.FieldSuggestIdentical = CreateObject("Scripting.Dictionary")
    Set plan.FieldInclude = CreateObject("Scripting.Dictionary")
    Set plan.Correspondence = CreateObject("Scripting.Dictionary")
    Set plan.HarvestedText = CreateObject("Scripting.Dictionary")

    Dim templateCandidates() As Candidate
    Dim templateShapes() As Object
    templateCandidates = Discovery.DiscoverSlideWithShapes(templateSld, templateShapes)

    Dim tLo As Long, tHi As Long, hasTemplateCandidates As Boolean
    On Error Resume Next
    tLo = LBound(templateCandidates): tHi = UBound(templateCandidates)
    hasTemplateCandidates = (Err.Number = 0)
    On Error GoTo 0

    If Not hasTemplateCandidates Then
        plan.FieldCount = 0
        BuildBatchPlan = plan
        Exit Function
    End If

    ' Discover every other slide's candidates ONCE (not once per field),
    ' flattened into single arrays across ALL other-slides rather than
    ' cached per-slide in a Dictionary. Real bug found and fixed 2026-07-25
    ' (diagnosed with a Fable-model agent after a long, confusing live-
    ' debugging session -- see SPIKE_NOTES_BatchOnboardFlow.md for the full
    ' account): a Candidate() array can NEVER legally be stored as a
    ' Dictionary/Variant value -- VBA's compiler rejects it ("Compile
    ' error: Only user-defined types defined in public object modules can
    ' be coerced to or from a variant or passed to late-bound functions"),
    ' since Candidate is a Public Type in a standard module (Discovery.bas),
    ' not a class. This is a DIFFERENT, narrower restriction than the
    ' already-documented "array-of-UDT-as-a-member-of-another-UDT" risk
    ' this module's own header comment warns about -- that one is about a
    ' UDT containing an array; this one is about ANY UDT array (even a
    ' plain local variable) being boxed into a Variant-backed container
    ' like a Dictionary, Collection item, or Variant array. otherShapesBySlide/
    ' otherAvailableBySlide (Object()/Boolean() arrays, not UDTs) were ALSO
    ' wrong here in a related but distinct way -- `Set` was used on them,
    ' which is invalid syntax for arrays (arrays aren't object references)
    ' -- moot now since this rewrite drops Dictionaries for this caching
    ' entirely. otherSlideCandStart(s)/otherSlideCandCount(s) (plain Long
    ' arrays -- primitives, no restriction) record where each slide's own
    ' candidates sit within the one flattened Candidate()/Object()/Boolean()
    ' triple, so each slide is still discovered only once.
    Dim oLo As Long, oHi As Long, hasOtherSlides As Boolean
    On Error Resume Next
    oLo = LBound(otherSlides): oHi = UBound(otherSlides)
    hasOtherSlides = (Err.Number = 0)
    On Error GoTo 0

    Dim allOtherCandidates() As Candidate
    Dim allOtherShapes() As Object
    Dim allOtherAvailable() As Boolean
    Dim otherSlideCandStart() As Long
    Dim otherSlideCandCount() As Long
    Dim totalOtherCandCount As Long
    totalOtherCandCount = 0

    Dim s As Long
    If hasOtherSlides Then
        ReDim otherSlideCandStart(oLo To oHi)
        ReDim otherSlideCandCount(oLo To oHi)

        For s = oLo To oHi
            Dim oc() As Candidate
            Dim os_() As Object
            oc = Discovery.DiscoverSlideWithShapes(otherSlides(s), os_)

            Dim ocLo As Long, ocHi As Long, hasOc As Boolean
            On Error Resume Next
            ocLo = LBound(oc): ocHi = UBound(oc): hasOc = (Err.Number = 0)
            On Error GoTo 0

            Dim thisCount As Long
            thisCount = 0
            If hasOc Then thisCount = ocHi - ocLo + 1

            otherSlideCandStart(s) = totalOtherCandCount + 1
            otherSlideCandCount(s) = thisCount

            If thisCount > 0 Then
                Dim newTotal As Long
                newTotal = totalOtherCandCount + thisCount
                ReDim Preserve allOtherCandidates(1 To newTotal)
                ReDim Preserve allOtherShapes(1 To newTotal)
                ReDim Preserve allOtherAvailable(1 To newTotal)

                Dim ci As Long, destIdx As Long
                destIdx = totalOtherCandCount
                For ci = ocLo To ocHi
                    destIdx = destIdx + 1
                    allOtherCandidates(destIdx) = oc(ci)
                    Set allOtherShapes(destIdx) = os_(ci)
                    allOtherAvailable(destIdx) = True
                Next ci
                totalOtherCandCount = newTotal
            End If
        Next s
    End If

    Dim fieldIdx As Long
    fieldIdx = 0

    Dim i As Long
    For i = tLo To tHi
        If Onboarding.IsCandidateField(templateCandidates(i)) Then
            fieldIdx = fieldIdx + 1
            plan.FieldNames(fieldIdx) = SuggestBatchFieldName(templateShapes(i), fieldIdx)
            Set plan.FieldTemplateShapes(fieldIdx) = templateShapes(i)
            plan.FieldInclude(fieldIdx) = True ' default: link it (safer default than silently dropping a real field)

            Dim templateValue As String
            If templateCandidates(i).HasText Then
                templateValue = templateShapes(i).TextFrame.TextRange.Text
            Else
                templateValue = "" ' pictures: not harvested as text, same limitation as every other module here
            End If
            plan.HarvestedText(FieldSlideKey(fieldIdx, 0)) = templateValue

            Dim allValues As Collection
            Set allValues = New Collection
            allValues.Add templateValue

            If hasOtherSlides Then
                For s = oLo To oHi
                    Dim sStart As Long, sCount As Long
                    sStart = otherSlideCandStart(s)
                    sCount = otherSlideCandCount(s)

                    If sCount > 0 Then
                        ' Build the still-available subset (within this
                        ' slide's own slice of the flattened arrays) to
                        ' score against -- claimed shapes (matched to an
                        ' earlier field on this same slide) are excluded so
                        ' two fields can never both grab the same shape.
                        Dim pool() As Candidate
                        Dim poolOrigIdx() As Long
                        Dim poolN As Long
                        poolN = 0
                        Dim a As Long
                        For a = sStart To sStart + sCount - 1
                            If allOtherAvailable(a) Then
                                poolN = poolN + 1
                                ReDim Preserve pool(1 To poolN)
                                ReDim Preserve poolOrigIdx(1 To poolN)
                                pool(poolN) = allOtherCandidates(a)
                                poolOrigIdx(poolN) = a
                            End If
                        Next a

                        If poolN > 0 Then
                            Dim m As MatchResult
                            m = Matching.Match(pool, templateCandidates(i))
                            If m.HasCandidate Then
                                Dim origIdx As Long
                                origIdx = poolOrigIdx(m.CandidateIndex)
                                allOtherAvailable(origIdx) = False

                                Set plan.Correspondence(FieldSlideKey(fieldIdx, s)) = allOtherShapes(origIdx)
                                Dim thisValue As String
                                If allOtherCandidates(origIdx).HasText Then
                                    thisValue = allOtherShapes(origIdx).TextFrame.TextRange.Text
                                Else
                                    thisValue = ""
                                End If
                                plan.HarvestedText(FieldSlideKey(fieldIdx, s)) = thisValue
                                allValues.Add thisValue
                            End If
                            ' No HasCandidate: this field simply has no
                            ' corresponding shape on this slide -- absent
                            ' from Correspondence/HarvestedText, handled at
                            ' commit time as "skip this field for this slide".
                        End If
                    End If
                Next s
            End If

            plan.FieldSuggestIdentical(fieldIdx) = AllValuesIdentical(allValues)
            ' Suggested default Include mirrors the classification: proven
            ' identical everywhere -> suggest skipping (decoration);
            ' anything else -> suggest linking, the safer default per
            ' Rohan's own framing (2026-07-26): missing a real field is
            ' worse than harmlessly linking a rarely-changing one.
            If plan.FieldSuggestIdentical(fieldIdx) Then
                plan.FieldInclude(fieldIdx) = False
            End If
        End If
    Next i

    plan.FieldCount = fieldIdx
    BuildBatchPlan = plan
End Function

' ---------------------------------------------------------------------
' Excel review grid -- the actual editable table. A fresh, throwaway
' workbook (not the paired Data workbook -- keeping review scratch state
' completely separate from real synced data avoids any risk of confusing
' the two), closed without saving once read back.
' ---------------------------------------------------------------------


Public Function WriteReviewGrid(ws As Object, plan As BatchOnboardPlan, otherSlideCount As Long) As Object
    ws.Cells(1, COL_FIELD_ID).Value = "Field ID (do not edit)"
    ws.Cells(1, COL_FIELD_NAME).Value = "Field Name"
    ws.Cells(1, COL_SUGGESTED).Value = "Suggested"
    ws.Cells(1, COL_INCLUDE).Value = "Include (Y/N)"
    ws.Cells(1, COL_TEMPLATE_VALUE).Value = "Template's Current Value"
    ws.Cells(1, COL_SAMPLE_OTHER_VALUES).Value = "Sample Values From Other Slides"

    Dim r As Long
    Dim fieldIdx As Long
    For fieldIdx = 1 To plan.FieldCount
        r = fieldIdx + 1
        ws.Cells(r, COL_FIELD_ID).Value = fieldIdx
        ws.Cells(r, COL_FIELD_NAME).Value = plan.FieldNames(fieldIdx)
        ws.Cells(r, COL_SUGGESTED).Value = IIf(plan.FieldSuggestIdentical(fieldIdx), "Decoration (identical everywhere)", "Field (varies)")
        ws.Cells(r, COL_INCLUDE).Value = IIf(plan.FieldInclude(fieldIdx), "Y", "N")
        ws.Cells(r, COL_TEMPLATE_VALUE).Value = plan.HarvestedText(FieldSlideKey(fieldIdx, 0))

        Dim samples As String
        Dim shown As Long
        shown = 0
        Dim s As Long
        For s = 1 To otherSlideCount
            Dim key As String
            key = FieldSlideKey(fieldIdx, s)
            If plan.HarvestedText.Exists(key) And shown < 3 Then
                If samples <> "" Then samples = samples & " | "
                samples = samples & plan.HarvestedText(key)
                shown = shown + 1
            End If
        Next s
        ws.Cells(r, COL_SAMPLE_OTHER_VALUES).Value = samples
    Next fieldIdx

    ws.Columns(COL_FIELD_NAME).ColumnWidth = 20
    ws.Columns(COL_SUGGESTED).ColumnWidth = 26
    ws.Columns(COL_TEMPLATE_VALUE).ColumnWidth = 30
    ws.Columns(COL_SAMPLE_OTHER_VALUES).ColumnWidth = 40
    ws.Rows(1).Font.Bold = True

    Set WriteReviewGrid = ws
End Function

' Reads the (possibly human-edited) grid back, updating `plan`'s
' FieldNames/FieldInclude in place -- ByRef since this is exactly the same
' "mutate the caller's structure" shape PlanAdoption's own out-params use,
' not a fresh return value. "Y"/"N" is read case-insensitively, trimmed;
' anything else is treated as "N" (never silently include on an
' unrecognized answer). A blank Field Name falls back to the existing
' proposed name rather than accepting an empty field name.
Public Sub ReadReviewGrid(ws As Object, ByRef plan As BatchOnboardPlan)
    Dim fieldIdx As Long
    For fieldIdx = 1 To plan.FieldCount
        Dim r As Long
        r = fieldIdx + 1

        Dim newName As String
        newName = Trim(CStr(ws.Cells(r, COL_FIELD_NAME).Value))
        If newName <> "" Then
            plan.FieldNames(fieldIdx) = newName
        End If

        Dim includeAnswer As String
        includeAnswer = UCase(Trim(CStr(ws.Cells(r, COL_INCLUDE).Value)))
        plan.FieldInclude(fieldIdx) = (includeAnswer = "Y")
    Next fieldIdx
End Sub

' ---------------------------------------------------------------------
' Commit -- the only part that writes anything. Slide index 0 means the
' template itself; 1..N means otherSlides(1)..otherSlides(N) (that array
' is always 1-based here -- both callers of this module build it via
' AdoptFlow.ValidateAdoptionSelection, which ReDims 1-based, so this is a
' documented internal invariant, not re-validated defensively per slide).
' ---------------------------------------------------------------------


Private Function SlideForIndex(s As Long, templateSld As Object, otherSlides() As Object) As Object
    If s = 0 Then
        Set SlideForIndex = templateSld
    Else
        Set SlideForIndex = otherSlides(s)
    End If
End Function

' confirmedKeys: Dictionary, slide index (0=template, 1..N=otherSlides) ->
' instance key String; "" or absent means skip that slide entirely this
' pass (no tags, no row -- never guessed). The template (slide 0) must have
' a real key, since it defines the type going forward; a caller supplying
' "" there is a genuine error, not silently handled here.
Public Function CommitBatch(plan As BatchOnboardPlan, templateSld As Object, otherSlides() As Object, otherSlideCount As Long, slideType As String, ws As Object, confirmedKeys As Object) As BatchCommitResult
    Dim result As BatchCommitResult

    Dim s As Long
    For s = 0 To otherSlideCount
        Dim instanceKey As String
        instanceKey = ""
        If confirmedKeys.Exists(s) Then instanceKey = CStr(confirmedKeys(s))

        If instanceKey = "" Then
            result.SkippedCount = result.SkippedCount + 1
        Else
            Dim sld As Object
            Set sld = SlideForIndex(s, templateSld, otherSlides)

            Dim harvested As Object
            Set harvested = CreateObject("Scripting.Dictionary")

            Dim fieldIdx As Long
            For fieldIdx = 1 To plan.FieldCount
                If plan.FieldInclude(fieldIdx) Then
                    Dim key As String
                    key = FieldSlideKey(fieldIdx, s)
                    If plan.HarvestedText.Exists(key) Then
                        Dim shp As Object
                        If s = 0 Then
                            Set shp = plan.FieldTemplateShapes(fieldIdx)
                        Else
                            Set shp = plan.Correspondence(key)
                        End If
                        Onboarding.ConfirmFieldMatch shp, plan.FieldNames(fieldIdx)
                        harvested(plan.FieldNames(fieldIdx)) = plan.HarvestedText(key)
                    End If
                    ' No correspondence found for this field on this slide:
                    ' silently skipped for this slide only, same "missing
                    ' field, never forced" posture SlideDuplication.bas's
                    ' MissingFieldCount already established -- not tracked
                    ' as a distinct count here to keep this report simple,
                    ' visible instead via the sample-values column a human
                    ' already reviewed in the grid.
                End If
            Next fieldIdx

            sld.Tags.Add "slide_type", slideType
            sld.Tags.Add "instance_key", instanceKey

            ExcelOutput.UpsertRow ws, instanceKey, harvested

            If VerifyBatchLink(sld, harvested) Then
                result.LinkedCount = result.LinkedCount + 1
            Else
                result.FailedVerificationCount = result.FailedVerificationCount + 1
                ReDim Preserve result.FailedVerificationLabels(1 To result.FailedVerificationCount)
                result.FailedVerificationLabels(result.FailedVerificationCount) = "Slide " & sld.SlideIndex & " (" & sld.Name & ")"
            End If
        End If
    Next s

    CommitBatch = result
End Function

' Same no-op round-trip DeckAdoption.bas's own (Private) VerifyLink already
' does -- duplicated rather than reused across modules since that one is
' Private and this project's convention is small, focused modules, not a
' shared-utility module for a 4-line check.
Private Function VerifyBatchLink(sld As Object, harvested As Object) As Boolean
    Dim fieldName As Variant
    For Each fieldName In harvested.Keys
        Dim r As InjectResult
        r = InjectPrimitive.InjectPrimitive(sld, CStr(fieldName), CStr(harvested(fieldName)))
        If Not r.Found Or r.Written Or Not r.Verified Then
            VerifyBatchLink = False
            Exit Function
        End If
    Next fieldName
    VerifyBatchLink = True
End Function

' ---------------------------------------------------------------------
' Ribbon entry point
' ---------------------------------------------------------------------

Public Function PromptBatchOnboardType() As String
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim slides() As Object
    Dim selErr As String
    selErr = AdoptFlow.ValidateAdoptionSelection(Application.ActiveWindow.Selection, slides)
    If selErr <> "" Then
        PromptBatchOnboardType = selErr
        Exit Function
    End If

    Dim lo As Long, hi As Long
    lo = LBound(slides): hi = UBound(slides)
    If (hi - lo + 1) < 2 Then
        PromptBatchOnboardType = "Select at least 2 slides of the same type (a template plus at least one more instance)."
        Exit Function
    End If

    Dim slideType As String
    slideType = InputBox("Name for this new slide type (e.g. 'quarterly-update'):", "Bulk Onboard Type")
    If Trim(slideType) = "" Then
        PromptBatchOnboardType = "Cancelled -- no type name given."
        Exit Function
    End If

    Dim templateSld As Object
    Set templateSld = slides(lo)

    Dim otherSlideCount As Long
    otherSlideCount = hi - lo ' hi-lo+1 total slides, minus the template
    Dim otherSlides() As Object
    ReDim otherSlides(1 To otherSlideCount)
    Dim i As Long
    For i = 1 To otherSlideCount
        Set otherSlides(i) = slides(lo + i)
    Next i

    Dim plan As BatchOnboardPlan
    plan = BuildBatchPlan(templateSld, otherSlides)

    If plan.FieldCount = 0 Then
        PromptBatchOnboardType = "No candidate fields found on the template slide (no text or picture shapes) -- nothing to onboard."
        Exit Function
    End If

    ' Scratch review workbook -- separate from the real paired Data
    ' workbook, closed without saving once read back.
    Dim xl As Object
    Set xl = WorkbookBridge.GetExcelApp()
    Dim reviewWb As Object
    Set reviewWb = xl.Workbooks.Add()
    Dim reviewWs As Object
    Set reviewWs = reviewWb.Worksheets(1)
    reviewWs.Name = "Field Review"
    WriteReviewGrid reviewWs, plan, otherSlideCount
    reviewWb.Activate

    If MsgBox("Review and edit the 'Field Review' sheet in Excel now -- rename fields, flip Include Y/N, exclude anything wrong." & vbCrLf & vbCrLf & "Click Yes when you're done editing, No to cancel.", vbYesNo + vbQuestion, "Bulk Onboard Type -- Review") <> vbYes Then
        reviewWb.Saved = True
        reviewWb.Close
        PromptBatchOnboardType = "Cancelled at review -- nothing written."
        Exit Function
    End If

    ReadReviewGrid reviewWs, plan
    reviewWb.Saved = True
    reviewWb.Close

    ' Instance-key prompts: template first (required), then each other
    ' slide (blank = skip, never guessed) -- same InputBox convention
    ' AdoptFlow.bas already established for exactly this decision.
    Dim confirmedKeys As Object
    Set confirmedKeys = CreateObject("Scripting.Dictionary")

    Dim templatePrompt As String
    templatePrompt = "Instance key for the template slide (Slide " & templateSld.SlideIndex & ") -- required, this slide defines the type:"
    Dim templateKey As String
    templateKey = InputBox(templatePrompt, "Bulk Onboard Type -- Instance Key")
    If Trim(templateKey) = "" Then
        PromptBatchOnboardType = "Cancelled -- the template slide must have an instance key."
        Exit Function
    End If
    confirmedKeys(0) = Trim(templateKey)

    For i = 1 To otherSlideCount
        Dim prompt As String
        prompt = "Instance key for Slide " & otherSlides(i).SlideIndex & " (leave blank to skip this slide this pass):"
        confirmedKeys(i) = InputBox(prompt, "Bulk Onboard Type -- Instance Key")
    Next i

    ' Establish (or reuse) the deck-workbook pairing.
    Dim workbookPath As String
    workbookPath = DeckRegistry.GetWorkbookPath(pres)
    Dim wb As Object
    If workbookPath = "" Then
        workbookPath = InputBox("This deck has no paired workbook yet." & vbCrLf & "Enter a full path for the new Data workbook (.xlsx):", "Bulk Onboard Type -- Pair Workbook")
        If Trim(workbookPath) = "" Then
            PromptBatchOnboardType = "Cancelled -- no workbook path given."
            Exit Function
        End If
        Set wb = WorkbookBridge.CreateWorkbook(workbookPath)
        If wb Is Nothing Then
            PromptBatchOnboardType = "Could not create workbook at: " & workbookPath
            Exit Function
        End If
        DeckRegistry.SetWorkbookPath pres, workbookPath
    Else
        Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath)
        If wb Is Nothing Then
            PromptBatchOnboardType = "Could not open the paired workbook at: " & workbookPath
            Exit Function
        End If
    End If

    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(wb, WorkbookBridge.SanitizeSheetName(slideType))
    If IsEmpty(ws.Cells(1, 1).Value) Then
        ExcelOutput.CreateSheet ws, DeckRegistry.GetOrCreateDeckId(pres)
    End If

    DeckRegistry.RegisterType pres, slideType, templateSld, ws.Name

    Dim commitResult As BatchCommitResult
    commitResult = CommitBatch(plan, templateSld, otherSlides, otherSlideCount, slideType, ws, confirmedKeys)

    Dim report As String
    report = "Linked: " & commitResult.LinkedCount & vbCrLf & _
        "Skipped (no instance key given): " & commitResult.SkippedCount & vbCrLf & _
        "FAILED verification: " & commitResult.FailedVerificationCount

    If commitResult.FailedVerificationCount > 0 Then
        Dim m As Long
        report = report & vbCrLf & "Failed slides (harvest bug this pass -- fix before re-running):"
        For m = 1 To commitResult.FailedVerificationCount
            report = report & vbCrLf & "  " & commitResult.FailedVerificationLabels(m)
        Next m
    End If

    PromptBatchOnboardType = report
End Function

Public Sub BatchOnboardType()
    Dim report As String
    report = PromptBatchOnboardType()
    If report <> "" Then
        RibbonUI.ShowSyncResult "Bulk Onboard Type", report
    End If
End Sub
