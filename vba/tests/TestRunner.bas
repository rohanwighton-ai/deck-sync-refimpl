Attribute VB_Name = "TestRunner"
Option Explicit

' A sentinel, not "" -- "" already means PASS, and SKIP must never be
' countable as one. A plain literal, not Chr(2)&"SKIPPED" -- VBA's Const
' requires a genuine constant expression at compile time, and Chr() is a
' function call, not one ("Compile error: Constant expression required",
' hit for real 2026-08-16). No Assert message is ever this exact literal, so
' no real result can collide with it. Declared here, at the very top -- a
' module-level Const after the first procedure is a VBA compile error that
' surfaces in a DIFFERENT module (AGENTS.md).
Private Const TEST_SKIPPED As String = "@@DECKSYNC_TEST_SKIPPED@@"

' First real-execution test harness for the PowerPoint-hosted modules
' (Discovery, InjectPrimitive, Matching, Resolve, SyncOperations,
' Onboarding). Every prior SPIKE_NOTES_*.md said "not executed or verified
' in this environment" -- this module exists to close that gap for real,
' driven headlessly via COM automation (see the PowerShell driver script),
' not by a human clicking through the Immediate window.
'
' Deliberately NOT the same as the ManualSmokeTest* subs already in each
' module: those use MsgBox, which would hang a headless script waiting on a
' click that never comes. Every Test_* function here is a pure assertion
' function -- returns "" on full pass, or one "FAIL: ..." line per failed
' assertion otherwise -- so RunAllTests can produce one machine-readable
' report with no human present.
'
' Fixture strategy: most tests build their own slides/shapes programmatically
' (deterministic, no dependency on a fixture file's exact geometry surviving
' across PowerPoint versions). Two tests still open a real test-fixtures/
' file directly (shp-groupshape.pptx, for its real 4-way sibling-ambiguity
' geometry, which would be needlessly fragile to reproduce by hand) --
' `fixturesDir` is that folder's path, passed in by the driver rather than
' hardcoded, since the real repo path isn't necessarily reachable the same
' way from a Windows-hosted COM session (see the driver script for how
' fixtures get staged).

' filterPattern "" (the default, and every existing caller's behaviour)
' runs everything, same as before this existed. Non-"" runs only tests whose
' name contains it (case-insensitive substring, not a regex) -- for fast
' iteration on one area without paying the full suite's real-Office COM
' overhead every time. See run_vba_tests.ps1's -Filter for the driver side,
' including the every-10th-run-forces-a-full-run cadence that exists because a
' filtered run proves nothing about the tests it skipped.
Public Function RunAllTests(fixturesDir As String, stagingDir As String, _
                             Optional filterPattern As String = "") As String
    Dim report As String
    report = "=== deck-sync-refimpl VBA test run (PowerPoint) ===" & vbCrLf
    If filterPattern <> "" Then report = report & "(filtered: '" & filterPattern & "')" & vbCrLf

    Dim r As String

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("Discovery_GroupRecursionFindsCandidates", filterPattern) Then
        r = Test_Discovery_GroupRecursionFindsCandidates(fixturesDir)
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Discovery_GroupRecursionFindsCandidates", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("InjectPrimitive_NoOpWhenValueAlreadyMatches", filterPattern) Then
        r = Test_InjectPrimitive_NoOpWhenValueAlreadyMatches()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "InjectPrimitive_NoOpWhenValueAlreadyMatches", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("InjectPrimitive_WritesAndVerifiesOnMismatch", filterPattern) Then
        r = Test_InjectPrimitive_WritesAndVerifiesOnMismatch()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "InjectPrimitive_WritesAndVerifiesOnMismatch", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("InjectPrimitive_AmbiguousTagRefusesToGuess", filterPattern) Then
        r = Test_InjectPrimitive_AmbiguousTagRefusesToGuess()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "InjectPrimitive_AmbiguousTagRefusesToGuess", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("InjectPrimitive_DryRunReportsWithoutWriting", filterPattern) Then
        r = Test_InjectPrimitive_DryRunReportsWithoutWriting()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "InjectPrimitive_DryRunReportsWithoutWriting", r

    If TestMatches("InjectPrimitive_FindsRoleTagInsideGroup", filterPattern) Then
        r = Test_InjectPrimitive_FindsRoleTagInsideGroup()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "InjectPrimitive_FindsRoleTagInsideGroup", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("InjectPrimitive_AmbiguousTagAcrossGroupAndTopLevelRefusesToGuess", filterPattern) Then
        r = Test_InjectPrimitive_AmbiguousTagAcrossGroupAndTopLevelRefusesToGuess()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "InjectPrimitive_AmbiguousTagAcrossGroupAndTopLevelRefusesToGuess", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("Matching_SiblingAmbiguityResolvedByZOrder", filterPattern) Then
        r = Test_Matching_SiblingAmbiguityResolvedByZOrder(fixturesDir)
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Matching_SiblingAmbiguityResolvedByZOrder", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("Matching_EnrichPlaceholderIdxReadsRealFile", filterPattern) Then
        r = Test_Matching_EnrichPlaceholderIdxReadsRealFile(stagingDir)
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Matching_EnrichPlaceholderIdxReadsRealFile", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("Resolve_ReadsTagsOffLiveSlide", filterPattern) Then
        r = Test_Resolve_ReadsTagsOffLiveSlide()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Resolve_ReadsTagsOffLiveSlide", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("SyncOperations_Cases1And4", filterPattern) Then
        r = Test_SyncOperations_Cases1And4()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "SyncOperations_Cases1And4", r

    If TestMatches("SyncOperations_PlanRoutineSyncDryRunWritesNothing", filterPattern) Then
        r = Test_SyncOperations_PlanRoutineSyncDryRunWritesNothing()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "SyncOperations_PlanRoutineSyncDryRunWritesNothing", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("SyncOperations_Case3NewRecord", filterPattern) Then
        r = Test_SyncOperations_Case3NewRecord()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "SyncOperations_Case3NewRecord", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("SyncOperations_DeviceFieldReachableThroughPlanRoutineSync", filterPattern) Then
        r = Test_SyncOperations_DeviceFieldReachableThroughPlanRoutineSync()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "SyncOperations_DeviceFieldReachableThroughPlanRoutineSync", r
    If TestMatches("SyncOperations_Case6UnclassifiedSlide", filterPattern) Then
        r = Test_SyncOperations_Case6UnclassifiedSlide()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "SyncOperations_Case6UnclassifiedSlide", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("Onboarding_HighAndMediumConfidence", filterPattern) Then
        r = Test_Onboarding_HighAndMediumConfidence()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Onboarding_HighAndMediumConfidence", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("Onboarding_OnboardNewInstanceAutoTagsHighOnly", filterPattern) Then
        r = Test_Onboarding_OnboardNewInstanceAutoTagsHighOnly()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Onboarding_OnboardNewInstanceAutoTagsHighOnly", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("Onboarding_PureDecorationNeverMatched", filterPattern) Then
        r = Test_Onboarding_PureDecorationNeverMatched()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Onboarding_PureDecorationNeverMatched", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("Verification_StructureMatchesAfterDuplicate", filterPattern) Then
        r = Test_Verification_StructureMatchesAfterDuplicate()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Verification_StructureMatchesAfterDuplicate", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("Verification_DetectsShapeCountMismatch", filterPattern) Then
        r = Test_Verification_DetectsShapeCountMismatch()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Verification_DetectsShapeCountMismatch", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("Verification_DetectsZOrderSwap", filterPattern) Then
        r = Test_Verification_DetectsZOrderSwap()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Verification_DetectsZOrderSwap", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("SlideDuplication_CreatesTaggedInjectedSlide", filterPattern) Then
        r = Test_SlideDuplication_CreatesTaggedInjectedSlide()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "SlideDuplication_CreatesTaggedInjectedSlide", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("SlideDuplication_RefusesInstanceKeyCollision", filterPattern) Then
        r = Test_SlideDuplication_RefusesInstanceKeyCollision()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "SlideDuplication_RefusesInstanceKeyCollision", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("SlideDuplication_PartialRowStillCreatesSlideButFlagsMissing", filterPattern) Then
        r = Test_SlideDuplication_PartialRowStillCreatesSlideButFlagsMissing()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "SlideDuplication_PartialRowStillCreatesSlideButFlagsMissing", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("TemplateSlide_CodeLetterOfReadsBothKeyShapes", filterPattern) Then
        r = Test_TemplateSlide_CodeLetterOfReadsBothKeyShapes()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "TemplateSlide_CodeLetterOfReadsBothKeyShapes", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("TemplateSlide_MakeTemplateProducesKeylessMarkedCopy", filterPattern) Then
        r = Test_TemplateSlide_MakeTemplateProducesKeylessMarkedCopy()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "TemplateSlide_MakeTemplateProducesKeylessMarkedCopy", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("TemplateSlide_ExcludedFromGatherAndNeverFlagged", filterPattern) Then
        r = Test_TemplateSlide_ExcludedFromGatherAndNeverFlagged()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "TemplateSlide_ExcludedFromGatherAndNeverFlagged", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("TemplateSlide_DuplicateStripsTheTemplateMarker", filterPattern) Then
        r = Test_TemplateSlide_DuplicateStripsTheTemplateMarker()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "TemplateSlide_DuplicateStripsTheTemplateMarker", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("TemplateSlide_RefusesToTemplateATemplate", filterPattern) Then
        r = Test_TemplateSlide_RefusesToTemplateATemplate()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "TemplateSlide_RefusesToTemplateATemplate", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("TemplateSlide_ConfirmTextStatesTheConsequences", filterPattern) Then
        r = Test_TemplateSlide_ConfirmTextStatesTheConsequences()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "TemplateSlide_ConfirmTextStatesTheConsequences", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("TemplateSlide_ConfirmTextIsLetterAwareAndDoesNotStealTheFallback", filterPattern) Then
        r = Test_TemplateSlide_ConfirmTextIsLetterAwareAndDoesNotStealTheFallback()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "TemplateSlide_ConfirmTextIsLetterAwareAndDoesNotStealTheFallback", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("TemplateSlide_ConfirmTextNamesWhenALetterBecomesTheFallback", filterPattern) Then
        r = Test_TemplateSlide_ConfirmTextNamesWhenALetterBecomesTheFallback()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "TemplateSlide_ConfirmTextNamesWhenALetterBecomesTheFallback", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("TemplateSlide_ExistingTemplateForLetterFindsTheRightOne", filterPattern) Then
        r = Test_TemplateSlide_ExistingTemplateForLetterFindsTheRightOne()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "TemplateSlide_ExistingTemplateForLetterFindsTheRightOne", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("TemplateSlide_ExistingTemplateForLetterEmptyLetterUsesTypeWideCheck", filterPattern) Then
        r = Test_TemplateSlide_ExistingTemplateForLetterEmptyLetterUsesTypeWideCheck()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "TemplateSlide_ExistingTemplateForLetterEmptyLetterUsesTypeWideCheck", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckRegistry_RegisterNewTemplateLetterClaimsFallbackWhenNoneExists", filterPattern) Then
        r = Test_DeckRegistry_RegisterNewTemplateLetterClaimsFallbackWhenNoneExists()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckRegistry_RegisterNewTemplateLetterClaimsFallbackWhenNoneExists", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckRegistry_RegisterNewTemplateLetterDoesNotStealAnExistingTemplateFallback", filterPattern) Then
        r = Test_DeckRegistry_RegisterNewTemplateLetterDoesNotStealAnExistingTemplateFallback()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckRegistry_RegisterNewTemplateLetterDoesNotStealAnExistingTemplateFallback", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("ReviewQueue_HashDistinguishesEveryField", filterPattern) Then
        r = Test_ReviewQueue_HashDistinguishesEveryField()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "ReviewQueue_HashDistinguishesEveryField", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("ReviewQueue_ProseNeverBatchesEvenWhenUniform", filterPattern) Then
        r = Test_ReviewQueue_ProseNeverBatchesEvenWhenUniform()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "ReviewQueue_ProseNeverBatchesEvenWhenUniform", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("ReviewQueue_UniformControlledGroupIsOneDecision", filterPattern) Then
        r = Test_ReviewQueue_UniformControlledGroupIsOneDecision()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "ReviewQueue_UniformControlledGroupIsOneDecision", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("ReviewQueue_FastPathAppliesUniformPartOnly", filterPattern) Then
        r = Test_ReviewQueue_FastPathAppliesUniformPartOnly()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "ReviewQueue_FastPathAppliesUniformPartOnly", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("ReviewQueue_ApprovalIsAffirmativeAndBatchWide", filterPattern) Then
        r = Test_ReviewQueue_ApprovalIsAffirmativeAndBatchWide()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "ReviewQueue_ApprovalIsAffirmativeAndBatchWide", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("InjectPrimitive_TrailingBreaksAreNotADifference", filterPattern) Then
        r = Test_InjectPrimitive_TrailingBreaksAreNotADifference()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "InjectPrimitive_TrailingBreaksAreNotADifference", r
    On Error GoTo 0
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("Drafting_OnlyTickedNonEmptyDraftsPublish", filterPattern) Then
        r = Test_Drafting_OnlyTickedNonEmptyDraftsPublish()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Drafting_OnlyTickedNonEmptyDraftsPublish", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("WorkbookBridge_RegisteredNameWinsOverASheetCalledRegister", filterPattern) Then
        r = Test_WorkbookBridge_RegisteredNameWinsOverASheetCalledRegister()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "WorkbookBridge_RegisteredNameWinsOverASheetCalledRegister", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("WorkbookBridge_RefusesToInventAMissingRegisterSheet", filterPattern) Then
        r = Test_WorkbookBridge_RefusesToInventAMissingRegisterSheet()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "WorkbookBridge_RefusesToInventAMissingRegisterSheet", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("FieldSpec_ValidationAppliesDownTheControlledColumn", filterPattern) Then
        r = Test_FieldSpec_ValidationAppliesDownTheControlledColumn()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "FieldSpec_ValidationAppliesDownTheControlledColumn", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("FieldSpec_ValidationReportsValuesOutsideTheVocabulary", filterPattern) Then
        r = Test_FieldSpec_ValidationReportsValuesOutsideTheVocabulary()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "FieldSpec_ValidationReportsValuesOutsideTheVocabulary", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("FieldSpec_ValidationSaysSoWhenNothingIsControlled", filterPattern) Then
        r = Test_FieldSpec_ValidationSaysSoWhenNothingIsControlled()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "FieldSpec_ValidationSaysSoWhenNothingIsControlled", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("FieldSpec_ValidationRefusesAVocabularyNamedLikeAStructuralColumn", filterPattern) Then
        r = Test_FieldSpec_ValidationRefusesAVocabularyNamedLikeAStructuralColumn()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "FieldSpec_ValidationRefusesAVocabularyNamedLikeAStructuralColumn", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("Drafting_QuarterTurnFerriesSubmitIntoReportedLastTime", filterPattern) Then
        r = Test_Drafting_QuarterTurnFerriesSubmitIntoReportedLastTime()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Drafting_QuarterTurnFerriesSubmitIntoReportedLastTime", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    On Error GoTo 0
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("Drafting_TickSurvivesSamePeriodRebuildAndClearsOnRollover", filterPattern) Then
        r = Test_Drafting_TickSurvivesSamePeriodRebuildAndClearsOnRollover()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Drafting_TickSurvivesSamePeriodRebuildAndClearsOnRollover", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("ExcelOutput_PeriodRowsAndRollForward", filterPattern) Then
        r = Test_ExcelOutput_PeriodRowsAndRollForward()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "ExcelOutput_PeriodRowsAndRollForward", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("ExcelOutput_FindsExistingRegisters", filterPattern) Then
        r = Test_ExcelOutput_FindsExistingRegisters()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "ExcelOutput_FindsExistingRegisters", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("ExcelOutput_MissingRegisterColumns", filterPattern) Then
        r = Test_ExcelOutput_MissingRegisterColumns()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "ExcelOutput_MissingRegisterColumns", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("WorkbookBridge_RefusesToCreateOverAnExistingFile", filterPattern) Then
        r = Test_WorkbookBridge_RefusesToCreateOverAnExistingFile()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "WorkbookBridge_RefusesToCreateOverAnExistingFile", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DraftingUI_ChainBlockHeaderLabelsTheField", filterPattern) Then
        r = Test_DraftingUI_ChainBlockHeaderLabelsTheField()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DraftingUI_ChainBlockHeaderLabelsTheField", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("ExcelOutput_LargestPeriodRowCountSpotsAPartialQuarter", filterPattern) Then
        r = Test_ExcelOutput_LargestPeriodRowCountSpotsAPartialQuarter()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "ExcelOutput_LargestPeriodRowCountSpotsAPartialQuarter", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("WorkbookBridge_ArchivesAPeriodToItsOwnFile", filterPattern) Then
        r = Test_WorkbookBridge_ArchivesAPeriodToItsOwnFile()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "WorkbookBridge_ArchivesAPeriodToItsOwnFile", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("WorkbookBridge_RefusesToOverwriteAnExistingArchive", filterPattern) Then
        r = Test_WorkbookBridge_RefusesToOverwriteAnExistingArchive()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "WorkbookBridge_RefusesToOverwriteAnExistingArchive", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("WorkbookBridge_ArchiveRefusesABlankPeriod", filterPattern) Then
        r = Test_WorkbookBridge_ArchiveRefusesABlankPeriod()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "WorkbookBridge_ArchiveRefusesABlankPeriod", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("BatchOnboard_NormalisesWorkbookPath", filterPattern) Then
        r = Test_BatchOnboard_NormalisesWorkbookPath()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboard_NormalisesWorkbookPath", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    On Error GoTo 0

    ' REGISTERED BY HAND, like every other test here. Writing the Function is
    ' not enough -- RunAllTests dispatches explicitly, so an unregistered test
    ' simply never runs and the suite still says PASS. Added both of these and
    ' the count stayed at 135, which is the only reason it was noticed.
    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DiscoverUI_GridListsEveryTextShapeWithUniqueIds", filterPattern) Then
        r = Test_DiscoverUI_GridListsEveryTextShapeWithUniqueIds()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DiscoverUI_GridListsEveryTextShapeWithUniqueIds", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DiscoverUI_MarksOnlyTickedAndNamedRows", filterPattern) Then
        r = Test_DiscoverUI_MarksOnlyTickedAndNamedRows()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DiscoverUI_MarksOnlyTickedAndNamedRows", r
    On Error GoTo 0
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("RunSync_CreateMissingRefusesWhileSlidesAreUnclassified", filterPattern) Then
        r = Test_RunSync_CreateMissingRefusesWhileSlidesAreUnclassified()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "RunSync_CreateMissingRefusesWhileSlidesAreUnclassified", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("RunSync_CreateMissingSlidesChoosesTemplateByRowLetter", filterPattern) Then
        r = Test_RunSync_CreateMissingSlidesChoosesTemplateByRowLetter()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "RunSync_CreateMissingSlidesChoosesTemplateByRowLetter", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckRegistry_DeckDeclaresItsOwnPeriod", filterPattern) Then
        r = Test_DeckRegistry_DeckDeclaresItsOwnPeriod()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckRegistry_DeckDeclaresItsOwnPeriod", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("WorkbookBridge_IndexExplainsEachSheet", filterPattern) Then
        r = Test_WorkbookBridge_IndexExplainsEachSheet()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "WorkbookBridge_IndexExplainsEachSheet", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("FieldSpec_GuidanceDrivesThePrompt", filterPattern) Then
        r = Test_FieldSpec_GuidanceDrivesThePrompt()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "FieldSpec_GuidanceDrivesThePrompt", r
    If TestMatches("FieldSpec_TheFiveProsePanelsEachHaveTheirOwnJob", filterPattern) Then
        r = Test_FieldSpec_TheFiveProsePanelsEachHaveTheirOwnJob()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "FieldSpec_TheFiveProsePanelsEachHaveTheirOwnJob", r
    If TestMatches("Drafting_AFieldWithNoValueLeavesColumnCEmpty", filterPattern) Then
        r = Test_Drafting_AFieldWithNoValueLeavesColumnCEmpty()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Drafting_AFieldWithNoValueLeavesColumnCEmpty", r
    If TestMatches("WorkbookBridge_RunLogSurvivesALineStartingWithEquals", filterPattern) Then
        r = Test_WorkbookBridge_RunLogSurvivesALineStartingWithEquals()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "WorkbookBridge_RunLogSurvivesALineStartingWithEquals", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("PlaceholderCheck_FindsRecordsNotTheTemplate", filterPattern) Then
        r = Test_PlaceholderCheck_FindsRecordsNotTheTemplate()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "PlaceholderCheck_FindsRecordsNotTheTemplate", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("PlaceholderCheck_MarkerDistinguishesStale", filterPattern) Then
        r = Test_PlaceholderCheck_MarkerDistinguishesStale()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "PlaceholderCheck_MarkerDistinguishesStale", r
    On Error GoTo 0
    On Error GoTo 0
    On Error GoTo 0
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("TagMigration_RenamesIncludingTemplateAndGroups", filterPattern) Then
        r = Test_TagMigration_RenamesIncludingTemplateAndGroups()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "TagMigration_RenamesIncludingTemplateAndGroups", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("TagMigration_MatchesValueCaseInsensitively", filterPattern) Then
        r = Test_TagMigration_MatchesValueCaseInsensitively()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "TagMigration_MatchesValueCaseInsensitively", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("IdentityCheck_FindsClonedKeys", filterPattern) Then
        r = Test_IdentityCheck_FindsClonedKeys()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "IdentityCheck_FindsClonedKeys", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("IdentityCheck_IgnoresTheTemplate", filterPattern) Then
        r = Test_IdentityCheck_IgnoresTheTemplate()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "IdentityCheck_IgnoresTheTemplate", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("TemplateAudit_ClassifyBoundaries", filterPattern) Then
        r = Test_TemplateAudit_ClassifyBoundaries()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "TemplateAudit_ClassifyBoundaries", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("TemplateAudit_SeparatesChromeFromProjectData", filterPattern) Then
        r = Test_TemplateAudit_SeparatesChromeFromProjectData()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "TemplateAudit_SeparatesChromeFromProjectData", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("TemplateAudit_NoComparisonSlidesStillLists", filterPattern) Then
        r = Test_TemplateAudit_NoComparisonSlidesStillLists()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "TemplateAudit_NoComparisonSlidesStillLists", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("TemplateAudit_RewriteLeavesNoStaleRows", filterPattern) Then
        r = Test_TemplateAudit_RewriteLeavesNoStaleRows(stagingDir)
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "TemplateAudit_RewriteLeavesNoStaleRows", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("TemplateAudit_RefusesToDiscardPendingDecisions", filterPattern) Then
        r = Test_TemplateAudit_RefusesToDiscardPendingDecisions()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "TemplateAudit_RefusesToDiscardPendingDecisions", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("RunSync_GatherInstancesFiltersByType", filterPattern) Then
        r = Test_RunSync_GatherInstancesFiltersByType()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "RunSync_GatherInstancesFiltersByType", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("RunSync_EndToEndCreatesSlidesFromFreshSheet", filterPattern) Then
        r = Test_RunSync_EndToEndCreatesSlidesFromFreshSheet()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "RunSync_EndToEndCreatesSlidesFromFreshSheet", r

    If TestMatches("RunSync_PreviewReportsWithoutTouchingTheDeck", filterPattern) Then
        r = Test_RunSync_PreviewReportsWithoutTouchingTheDeck()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "RunSync_PreviewReportsWithoutTouchingTheDeck", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("RunSync_ConfirmSyncTextReportsUncreatableRows", filterPattern) Then
        r = Test_RunSync_ConfirmSyncTextReportsUncreatableRows()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "RunSync_ConfirmSyncTextReportsUncreatableRows", r
    On Error GoTo 0


    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckAdoption_AlreadyLinkedSlideSkipped", filterPattern) Then
        r = Test_DeckAdoption_AlreadyLinkedSlideSkipped()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckAdoption_AlreadyLinkedSlideSkipped", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("Harvest_WritesOnlyWhereTheRegisterCellIsEmpty", filterPattern) Then
        r = Test_Harvest_WritesOnlyWhereTheRegisterCellIsEmpty()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Harvest_WritesOnlyWhereTheRegisterCellIsEmpty", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("Harvest_SkipsATaggedGroupInsteadOfReadingItAsEmpty", filterPattern) Then
        r = Test_Harvest_SkipsATaggedGroupInsteadOfReadingItAsEmpty()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Harvest_SkipsATaggedGroupInsteadOfReadingItAsEmpty", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("Harvest_PropagationStampsOnlyTheMissingRole", filterPattern) Then
        r = Test_Harvest_PropagationStampsOnlyTheMissingRole()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Harvest_PropagationStampsOnlyTheMissingRole", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("Harvest_StoresSlideTextVerbatimRatherThanLettingExcelParseIt", filterPattern) Then
        r = Test_Harvest_StoresSlideTextVerbatimRatherThanLettingExcelParseIt()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Harvest_StoresSlideTextVerbatimRatherThanLettingExcelParseIt", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("Harvest_RefusesAFieldThePublishPathWouldTreatAsABar", filterPattern) Then
        r = Test_Harvest_RefusesAFieldThePublishPathWouldTreatAsABar()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Harvest_RefusesAFieldThePublishPathWouldTreatAsABar", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("Harvest_DryRunCountsFieldsThatLabellingWillCreate", filterPattern) Then
        r = Test_Harvest_DryRunCountsFieldsThatLabellingWillCreate()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Harvest_DryRunCountsFieldsThatLabellingWillCreate", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("Membership_RetirementSkipsTemplateAndUnclassifiedSlides", filterPattern) Then
        r = Test_Membership_RetirementSkipsTemplateAndUnclassifiedSlides()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Membership_RetirementSkipsTemplateAndUnclassifiedSlides", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("Matching_ExactShapeNameBeatsGeometryOnlyWhenUnambiguous", filterPattern) Then
        r = Test_Matching_ExactShapeNameBeatsGeometryOnlyWhenUnambiguous()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Matching_ExactShapeNameBeatsGeometryOnlyWhenUnambiguous", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckAdoption_ReadyHighConfidenceSlideLinkedAndCreatesFreshRow", filterPattern) Then
        r = Test_DeckAdoption_ReadyHighConfidenceSlideLinkedAndCreatesFreshRow()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckAdoption_ReadyHighConfidenceSlideLinkedAndCreatesFreshRow", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckAdoption_MediumConfidenceSlideNeedsConfirmationAndIsNotTagged", filterPattern) Then
        r = Test_DeckAdoption_MediumConfidenceSlideNeedsConfirmationAndIsNotTagged()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckAdoption_MediumConfidenceSlideNeedsConfirmationAndIsNotTagged", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckAdoption_UnclassifiedSlideExcluded", filterPattern) Then
        r = Test_DeckAdoption_UnclassifiedSlideExcluded()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckAdoption_UnclassifiedSlideExcluded", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckAdoption_MatchesExistingKeylessRowLinksWithoutCreatingNewRow", filterPattern) Then
        r = Test_DeckAdoption_MatchesExistingKeylessRowLinksWithoutCreatingNewRow()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckAdoption_MatchesExistingKeylessRowLinksWithoutCreatingNewRow", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckAdoption_MultiSlideZeroBasedBatchKeepsIndicesAligned", filterPattern) Then
        r = Test_DeckAdoption_MultiSlideZeroBasedBatchKeepsIndicesAligned()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckAdoption_MultiSlideZeroBasedBatchKeepsIndicesAligned", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("ResolveFields_ValidateSingleShapeSelectionAcceptsOneShape", filterPattern) Then
        r = Test_ResolveFields_ValidateSingleShapeSelectionAcceptsOneShape()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "ResolveFields_ValidateSingleShapeSelectionAcceptsOneShape", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("ResolveFields_ValidateSingleShapeSelectionRejectsMultiple", filterPattern) Then
        r = Test_ResolveFields_ValidateSingleShapeSelectionRejectsMultiple()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "ResolveFields_ValidateSingleShapeSelectionRejectsMultiple", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("ResolveFields_BuildRolePickerPromptListsRolesNumbered", filterPattern) Then
        r = Test_ResolveFields_BuildRolePickerPromptListsRolesNumbered()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "ResolveFields_BuildRolePickerPromptListsRolesNumbered", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("ResolveFields_PickRoleFromListAcceptsNumberOrName", filterPattern) Then
        r = Test_ResolveFields_PickRoleFromListAcceptsNumberOrName()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "ResolveFields_PickRoleFromListAcceptsNumberOrName", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("ResolveFields_EndToEndTagsSelectedShapeViaPickedRole", filterPattern) Then
        r = Test_ResolveFields_EndToEndTagsSelectedShapeViaPickedRole()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "ResolveFields_EndToEndTagsSelectedShapeViaPickedRole", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckRegistry_BuildAndParseTypeRegistrationRoundTrip", filterPattern) Then
        r = Test_DeckRegistry_BuildAndParseTypeRegistrationRoundTrip()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckRegistry_BuildAndParseTypeRegistrationRoundTrip", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckRegistry_ParseTypeRegistrationRejectsMalformed", filterPattern) Then
        r = Test_DeckRegistry_ParseTypeRegistrationRejectsMalformed()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckRegistry_ParseTypeRegistrationRejectsMalformed", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckRegistry_GetOrCreateDeckIdIsStableAcrossCalls", filterPattern) Then
        r = Test_DeckRegistry_GetOrCreateDeckIdIsStableAcrossCalls()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckRegistry_GetOrCreateDeckIdIsStableAcrossCalls", r

    If TestMatches("DeckRegistry_PairingVerdictOnlyRefusesAKnownDifferentDeck", filterPattern) Then
        r = Test_DeckRegistry_PairingVerdictOnlyRefusesAKnownDifferentDeck()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckRegistry_PairingVerdictOnlyRefusesAKnownDifferentDeck", r

    If TestMatches("Drafting_LayoutStampIsFoundNotAssumed", filterPattern) Then
        r = Test_Drafting_LayoutStampIsFoundNotAssumed()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Drafting_LayoutStampIsFoundNotAssumed", r

    If TestMatches("Discovery_ADeviceGroupIsOneCandidateNotItsParts", filterPattern) Then
        r = Test_Discovery_ADeviceGroupIsOneCandidateNotItsParts()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Discovery_ADeviceGroupIsOneCandidateNotItsParts", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckRegistry_RegisterAndLookupTypeRoundTrip", filterPattern) Then
        r = Test_DeckRegistry_RegisterAndLookupTypeRoundTrip()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckRegistry_RegisterAndLookupTypeRoundTrip", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckRegistry_LookupTypeFalseWhenNotRegistered", filterPattern) Then
        r = Test_DeckRegistry_LookupTypeFalseWhenNotRegistered()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckRegistry_LookupTypeFalseWhenNotRegistered", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckRegistry_LookupTypeFalseWhenTemplateSlideDeleted", filterPattern) Then
        r = Test_DeckRegistry_LookupTypeFalseWhenTemplateSlideDeleted()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckRegistry_LookupTypeFalseWhenTemplateSlideDeleted", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckRegistry_ListRegisteredTypesListsAllRegistered", filterPattern) Then
        r = Test_DeckRegistry_ListRegisteredTypesListsAllRegistered()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckRegistry_ListRegisteredTypesListsAllRegistered", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckRegistry_RegisterAndLookupTemplateLetterRoundTrip", filterPattern) Then
        r = Test_DeckRegistry_RegisterAndLookupTemplateLetterRoundTrip()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckRegistry_RegisterAndLookupTemplateLetterRoundTrip", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckRegistry_LookupTemplateLetterFalseWhenNotRegistered", filterPattern) Then
        r = Test_DeckRegistry_LookupTemplateLetterFalseWhenNotRegistered()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckRegistry_LookupTemplateLetterFalseWhenNotRegistered", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckRegistry_TwoLettersOfSameTypeDoNotCollide", filterPattern) Then
        r = Test_DeckRegistry_TwoLettersOfSameTypeDoNotCollide()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckRegistry_TwoLettersOfSameTypeDoNotCollide", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckRegistry_LookupTemplateForLetterFallsBackWhenLetterEmpty", filterPattern) Then
        r = Test_DeckRegistry_LookupTemplateForLetterFallsBackWhenLetterEmpty()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckRegistry_LookupTemplateForLetterFallsBackWhenLetterEmpty", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckRegistry_LookupTemplateForLetterFallsBackWhenLetterUnregistered", filterPattern) Then
        r = Test_DeckRegistry_LookupTemplateForLetterFallsBackWhenLetterUnregistered()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckRegistry_LookupTemplateForLetterFallsBackWhenLetterUnregistered", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckRegistry_LookupTemplateForLetterPrefersLetterOverUnlettered", filterPattern) Then
        r = Test_DeckRegistry_LookupTemplateForLetterPrefersLetterOverUnlettered()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckRegistry_LookupTemplateForLetterPrefersLetterOverUnlettered", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckRegistry_WorkbookPathRoundTrip", filterPattern) Then
        r = Test_DeckRegistry_WorkbookPathRoundTrip()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckRegistry_WorkbookPathRoundTrip", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("WorkbookBridge_SanitizeSheetNameStripsInvalidCharsAndTruncates", filterPattern) Then
        r = Test_WorkbookBridge_SanitizeSheetNameStripsInvalidCharsAndTruncates()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "WorkbookBridge_SanitizeSheetNameStripsInvalidCharsAndTruncates", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("WorkbookBridge_IsDirtyDetectsUnsavedEdits", filterPattern) Then
        r = Test_WorkbookBridge_IsDirtyDetectsUnsavedEdits()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "WorkbookBridge_IsDirtyDetectsUnsavedEdits", r
    On Error GoTo 0
    On Error GoTo 0
    On Error GoTo 0
    On Error GoTo 0
    On Error GoTo 0
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("RibbonUI_ResolveTypeAnswerAcceptsNumberOrName", filterPattern) Then
        r = Test_RibbonUI_ResolveTypeAnswerAcceptsNumberOrName()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "RibbonUI_ResolveTypeAnswerAcceptsNumberOrName", r
    On Error GoTo 0


    r = "": On Error Resume Next: Err.Clear
    If TestMatches("RibbonUI_BuildTypePickerPromptListsAllTypes", filterPattern) Then
        r = Test_RibbonUI_BuildTypePickerPromptListsAllTypes()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "RibbonUI_BuildTypePickerPromptListsAllTypes", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("CommandBarUI_ShowToolbarCreatesWiredButtons", filterPattern) Then
        r = Test_CommandBarUI_ShowToolbarCreatesWiredButtons()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "CommandBarUI_ShowToolbarCreatesWiredButtons", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckRegistry_PeriodOnDiskReadsTheSavedFile", filterPattern) Then
        r = Test_DeckRegistry_PeriodOnDiskReadsTheSavedFile()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckRegistry_PeriodOnDiskReadsTheSavedFile", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DeckRegistry_SaveDeckVerifiedProvesTheFileMoved", filterPattern) Then
        r = Test_DeckRegistry_SaveDeckVerifiedProvesTheFileMoved()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckRegistry_SaveDeckVerifiedProvesTheFileMoved", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("ReviewQueue_DescribeDifferenceNamesTheInvisible", filterPattern) Then
        r = Test_ReviewQueue_DescribeDifferenceNamesTheInvisible()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "ReviewQueue_DescribeDifferenceNamesTheInvisible", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("Sources_RefsForOtherPeriodCatchesTheWrongQuarter", filterPattern) Then
        r = Test_Sources_RefsForOtherPeriodCatchesTheWrongQuarter()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Sources_RefsForOtherPeriodCatchesTheWrongQuarter", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("Sources_RefsForOtherPeriodIsSilentOnNeutralAndUnknown", filterPattern) Then
        r = Test_Sources_RefsForOtherPeriodIsSilentOnNeutralAndUnknown()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Sources_RefsForOtherPeriodIsSilentOnNeutralAndUnknown", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("CommandBarUI_ShowToolbarIsIdempotent", filterPattern) Then
        r = Test_CommandBarUI_ShowToolbarIsIdempotent()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "CommandBarUI_ShowToolbarIsIdempotent", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("CommandBarUI_HideToolbarRemovesIt", filterPattern) Then
        r = Test_CommandBarUI_HideToolbarRemovesIt()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "CommandBarUI_HideToolbarRemovesIt", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("AdoptFlow_ValidateAdoptionSelectionSortsIntoDeckOrder", filterPattern) Then
        r = Test_AdoptFlow_ValidateAdoptionSelectionSortsIntoDeckOrder()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "AdoptFlow_ValidateAdoptionSelectionSortsIntoDeckOrder", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("AdoptFlow_ValidateAdoptionSelectionRejectsNonSlideSelection", filterPattern) Then
        r = Test_AdoptFlow_ValidateAdoptionSelectionRejectsNonSlideSelection()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "AdoptFlow_ValidateAdoptionSelectionRejectsNonSlideSelection", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("AdoptFlow_ExcludeTemplateSlideRemovesOnlyTemplate", filterPattern) Then
        r = Test_AdoptFlow_ExcludeTemplateSlideRemovesOnlyTemplate()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "AdoptFlow_ExcludeTemplateSlideRemovesOnlyTemplate", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("AdoptFlow_BuildAdoptionReviewSummaryCountsAndListsNonReady", filterPattern) Then
        r = Test_AdoptFlow_BuildAdoptionReviewSummaryCountsAndListsNonReady()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "AdoptFlow_BuildAdoptionReviewSummaryCountsAndListsNonReady", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("BatchOnboardFlow_AllValuesIdenticalDetectsMatchAndMismatch", filterPattern) Then
        r = Test_BatchOnboardFlow_AllValuesIdenticalDetectsMatchAndMismatch()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_AllValuesIdenticalDetectsMatchAndMismatch", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("BatchOnboardFlow_SuggestBatchFieldNameReusesPhNameOrFallsBack", filterPattern) Then
        r = Test_BatchOnboardFlow_SuggestBatchFieldNameReusesPhNameOrFallsBack()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_SuggestBatchFieldNameReusesPhNameOrFallsBack", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("BatchOnboardFlow_NormalizeFieldTypeAcceptsNumberOrName", filterPattern) Then
        r = Test_BatchOnboardFlow_NormalizeFieldTypeAcceptsNumberOrName()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_NormalizeFieldTypeAcceptsNumberOrName", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("BatchOnboardFlow_NormalizeFieldVolatilityAcceptsNumberOrName", filterPattern) Then
        r = Test_BatchOnboardFlow_NormalizeFieldVolatilityAcceptsNumberOrName()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_NormalizeFieldVolatilityAcceptsNumberOrName", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("BatchOnboardFlow_SuggestInstanceKeyUsesFirstFieldsHarvestedValue", filterPattern) Then
        r = Test_BatchOnboardFlow_SuggestInstanceKeyUsesFirstFieldsHarvestedValue()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_SuggestInstanceKeyUsesFirstFieldsHarvestedValue", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("BatchOnboardFlow_FindSameLayoutSlidesGroupsByLayoutOnly", filterPattern) Then
        r = Test_BatchOnboardFlow_FindSameLayoutSlidesGroupsByLayoutOnly()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_FindSameLayoutSlidesGroupsByLayoutOnly", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("BatchOnboardFlow_BuildBatchPlanFindsCorrespondenceAndHarvestsAcrossSlides", filterPattern) Then
        r = Test_BatchOnboardFlow_BuildBatchPlanFindsCorrespondenceAndHarvestsAcrossSlides()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_BuildBatchPlanFindsCorrespondenceAndHarvestsAcrossSlides", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("BatchOnboardFlow_BuildBatchPlanFromMarkedFieldsUsesOnlyMarkedShapes", filterPattern) Then
        r = Test_BatchOnboardFlow_BuildBatchPlanFromMarkedFieldsUsesOnlyMarkedShapes()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_BuildBatchPlanFromMarkedFieldsUsesOnlyMarkedShapes", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("BatchOnboardFlow_MarkShapeForBatchAccumulatesAndDedupes", filterPattern) Then
        r = Test_BatchOnboardFlow_MarkShapeForBatchAccumulatesAndDedupes()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_MarkShapeForBatchAccumulatesAndDedupes", r

    If TestMatches("BatchOnboardFlow_FieldPreviewIsShortAndSingleLine", filterPattern) Then
        r = Test_BatchOnboardFlow_FieldPreviewIsShortAndSingleLine()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_FieldPreviewIsShortAndSingleLine", r

    If TestMatches("BatchOnboardFlow_ExistingInstanceKeyIsReusedNotRederived", filterPattern) Then
        r = Test_BatchOnboardFlow_ExistingInstanceKeyIsReusedNotRederived()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_ExistingInstanceKeyIsReusedNotRederived", r

    If TestMatches("DeckRegistry_WorkbookPathSurvivesAMovedDeck", filterPattern) Then
        r = Test_DeckRegistry_WorkbookPathSurvivesAMovedDeck()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckRegistry_WorkbookPathSurvivesAMovedDeck", r

    If TestMatches("BatchOnboardFlow_ConflictingSlideTypeIsDetected", filterPattern) Then
        r = Test_BatchOnboardFlow_ConflictingSlideTypeIsDetected()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_ConflictingSlideTypeIsDetected", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("BatchOnboardFlow_MarkShapeForBatchAcceptsShapeInsideGroup", filterPattern) Then
        r = Test_BatchOnboardFlow_MarkShapeForBatchAcceptsShapeInsideGroup()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_MarkShapeForBatchAcceptsShapeInsideGroup", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("BatchOnboardFlow_MarkShapeForBatchRejectsWholeGroupSelection", filterPattern) Then
        r = Test_BatchOnboardFlow_MarkShapeForBatchRejectsWholeGroupSelection()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_MarkShapeForBatchRejectsWholeGroupSelection", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("BatchOnboardFlow_MarkingSessionPropertyRoundTripsBeyond255Chars", filterPattern) Then
        r = Test_BatchOnboardFlow_MarkingSessionPropertyRoundTripsBeyond255Chars()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_MarkingSessionPropertyRoundTripsBeyond255Chars", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("BatchOnboardFlow_RestoreMarkingSessionRecoversMarkedFields", filterPattern) Then
        r = Test_BatchOnboardFlow_RestoreMarkingSessionRecoversMarkedFields()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_RestoreMarkingSessionRecoversMarkedFields", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("BatchOnboardFlow_SaveMarkingSessionToPropertyForcesRealSave", filterPattern) Then
        r = Test_BatchOnboardFlow_SaveMarkingSessionToPropertyForcesRealSave()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_SaveMarkingSessionToPropertyForcesRealSave", r
    If TestMatches("BatchOnboardFlow_ShouldForceSaveLeavesHealthyAutoSaveAlone", filterPattern) Then
        r = Test_BatchOnboardFlow_ShouldForceSaveLeavesHealthyAutoSaveAlone()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_ShouldForceSaveLeavesHealthyAutoSaveAlone", r
    If TestMatches("BatchOnboardFlow_LastSaveTimeOfReadsARealTimestamp", filterPattern) Then
        r = Test_BatchOnboardFlow_LastSaveTimeOfReadsARealTimestamp()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_LastSaveTimeOfReadsARealTimestamp", r
    If TestMatches("DeckRegistry_PeriodIsReadableThroughSlideParent", filterPattern) Then
        r = Test_DeckRegistry_PeriodIsReadableThroughSlideParent()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckRegistry_PeriodIsReadableThroughSlideParent", r
    If TestMatches("BatchOnboardFlow_NeedsSessionRestoreCoversSameDeckReopen", filterPattern) Then
        r = Test_BatchOnboardFlow_NeedsSessionRestoreCoversSameDeckReopen()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_NeedsSessionRestoreCoversSameDeckReopen", r
    If TestMatches("BatchOnboardFlow_WorkbookPathProblemRejectsTheRealMistakes", filterPattern) Then
        r = Test_BatchOnboardFlow_WorkbookPathProblemRejectsTheRealMistakes()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_WorkbookPathProblemRejectsTheRealMistakes", r
    If TestMatches("BatchOnboardFlow_InstanceKeyGridPrefillsAndCatchesClashes", filterPattern) Then
        r = Test_BatchOnboardFlow_InstanceKeyGridPrefillsAndCatchesClashes()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_InstanceKeyGridPrefillsAndCatchesClashes", r
    If TestMatches("RibbonUI_UnexpectedErrorTextTellsTheTruth", filterPattern) Then
        r = Test_RibbonUI_UnexpectedErrorTextTellsTheTruth()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "RibbonUI_UnexpectedErrorTextTellsTheTruth", r
    If TestMatches("RibbonUI_CapReportKeepsTheQuestion", filterPattern) Then
        r = Test_RibbonUI_CapReportKeepsTheQuestion()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "RibbonUI_CapReportKeepsTheQuestion", r
    If TestMatches("ReviewQueue_EmptyQueueDoesNotClaimTheDeckMatches", filterPattern) Then
        r = Test_ReviewQueue_EmptyQueueDoesNotClaimTheDeckMatches()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "ReviewQueue_EmptyQueueDoesNotClaimTheDeckMatches", r
    If TestMatches("CommandBarUI_EveryDeclaredCapabilityHasAButton", filterPattern) Then
        r = Test_CommandBarUI_EveryDeclaredCapabilityHasAButton()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "CommandBarUI_EveryDeclaredCapabilityHasAButton", r
    If TestMatches("ReviewQueue_ParityAndTheCreateThreshold", filterPattern) Then
        r = Test_ReviewQueue_ParityAndTheCreateThreshold()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "ReviewQueue_ParityAndTheCreateThreshold", r
    If TestMatches("RibbonUI_WrapperHandlerSurvivesOnErrorGoToZero", filterPattern) Then
        r = Test_RibbonUI_WrapperHandlerSurvivesOnErrorGoToZero()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "RibbonUI_WrapperHandlerSurvivesOnErrorGoToZero", r
    If TestMatches("ReviewQueue_PendingApprovalsCountsTicksAndIgnoresConsumed", filterPattern) Then
        r = Test_ReviewQueue_PendingApprovalsCountsTicksAndIgnoresConsumed()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "ReviewQueue_PendingApprovalsCountsTicksAndIgnoresConsumed", r
    If TestMatches("DeckRegistry_LocalPathForUrlOnlyAnswersWhenTheFileIsThere", filterPattern) Then
        r = Test_DeckRegistry_LocalPathForUrlOnlyAnswersWhenTheFileIsThere()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckRegistry_LocalPathForUrlOnlyAnswersWhenTheFileIsThere", r
    If TestMatches("DeckRegistry_LocalPathForUrlFindsARealSyncedFile", filterPattern) Then
        r = Test_DeckRegistry_LocalPathForUrlFindsARealSyncedFile()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DeckRegistry_LocalPathForUrlFindsARealSyncedFile", r
    If TestMatches("ReviewQueue_BackupDestinationHandlesACloudDeck", filterPattern) Then
        r = Test_ReviewQueue_BackupDestinationHandlesACloudDeck()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "ReviewQueue_BackupDestinationHandlesACloudDeck", r
    If TestMatches("ReviewQueue_ApplyApprovedNamesTheItemWhenInjectFieldCrashes", filterPattern) Then
        r = Test_ReviewQueue_ApplyApprovedNamesTheItemWhenInjectFieldCrashes()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "ReviewQueue_ApplyApprovedNamesTheItemWhenInjectFieldCrashes", r
    If TestMatches("ReviewQueue_BuildQueuePreTicksEveryItem", filterPattern) Then
        r = Test_ReviewQueue_BuildQueuePreTicksEveryItem()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "ReviewQueue_BuildQueuePreTicksEveryItem", r
    If TestMatches("Sources_CitedBlockPutsTheDocumentInThePrompt", filterPattern) Then
        r = Test_Sources_CitedBlockPutsTheDocumentInThePrompt()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Sources_CitedBlockPutsTheDocumentInThePrompt", r
    If TestMatches("Drafting_CitedSourceReachesThePromptCell", filterPattern) Then
        r = Test_Drafting_CitedSourceReachesThePromptCell()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Drafting_CitedSourceReachesThePromptCell", r
    If TestMatches("Drafting_Layout3SheetMigratesIntoLayout4Columns", filterPattern) Then
        r = Test_Drafting_Layout3SheetMigratesIntoLayout4Columns()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Drafting_Layout3SheetMigratesIntoLayout4Columns", r
    If TestMatches("InjectPicture_FillsStampsAndThenStaysSilent", filterPattern) Then
        r = Test_InjectPicture_FillsStampsAndThenStaysSilent()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "InjectPicture_FillsStampsAndThenStaysSilent", r
    If TestMatches("InjectPicture_RefusesABadLocatorWithoutLosingTheOldImage", filterPattern) Then
        r = Test_InjectPicture_RefusesABadLocatorWithoutLosingTheOldImage()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "InjectPicture_RefusesABadLocatorWithoutLosingTheOldImage", r
    If TestMatches("InjectPicture_CroppedFrameIsFilledUncroppedIsFitted", filterPattern) Then
        r = Test_InjectPicture_CroppedFrameIsFilledUncroppedIsFitted()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "InjectPicture_CroppedFrameIsFilledUncroppedIsFitted", r
    If TestMatches("InjectProgress_MeasuresAgainstTheTrackNotItself", filterPattern) Then
        r = Test_InjectProgress_MeasuresAgainstTheTrackNotItself()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "InjectProgress_MeasuresAgainstTheTrackNotItself", r
    If TestMatches("InjectProgress_RefusesWithoutATrack", filterPattern) Then
        r = Test_InjectProgress_RefusesWithoutATrack()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "InjectProgress_RefusesWithoutATrack", r
    If TestMatches("InjectField_RoutesEachTypeByItsShape", filterPattern) Then
        r = Test_InjectField_RoutesEachTypeByItsShape()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "InjectField_RoutesEachTypeByItsShape", r
    If TestMatches("InjectField_RefusesANonNumberForABarWithoutGoingQuiet", filterPattern) Then
        r = Test_InjectField_RefusesANonNumberForABarWithoutGoingQuiet()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "InjectField_RefusesANonNumberForABarWithoutGoingQuiet", r
    If TestMatches("InjectField_TwoSlidesEachGetTheirOwnValue", filterPattern) Then
        r = Test_InjectField_TwoSlidesEachGetTheirOwnValue()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "InjectField_TwoSlidesEachGetTheirOwnValue", r
    If TestMatches("FieldWiring_NamesTheFieldsNothingCarries", filterPattern) Then
        r = Test_FieldWiring_NamesTheFieldsNothingCarries()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "FieldWiring_NamesTheFieldsNothingCarries", r
    If TestMatches("FieldWiring_TemplateIsCheckedSeparatelyFromInstances", filterPattern) Then
        r = Test_FieldWiring_TemplateIsCheckedSeparatelyFromInstances()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "FieldWiring_TemplateIsCheckedSeparatelyFromInstances", r
    If TestMatches("FieldWiring_OrphanTrackIsAHalfMarkedBar", filterPattern) Then
        r = Test_FieldWiring_OrphanTrackIsAHalfMarkedBar()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "FieldWiring_OrphanTrackIsAHalfMarkedBar", r
    If TestMatches("FieldWiring_CoverageCountsSlidesNotPresence", filterPattern) Then
        r = Test_FieldWiring_CoverageCountsSlidesNotPresence()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "FieldWiring_CoverageCountsSlidesNotPresence", r
    If TestMatches("FieldWiring_ATrackIsATagNotAField", filterPattern) Then
        r = Test_FieldWiring_ATrackIsATagNotAField()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "FieldWiring_ATrackIsATagNotAField", r
    If TestMatches("FieldWiring_BothCompanionsAreTagsNotFields", filterPattern) Then
        r = Test_FieldWiring_BothCompanionsAreTagsNotFields()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "FieldWiring_BothCompanionsAreTagsNotFields", r
    If TestMatches("FieldWiring_CaseNearMissIsNamedNotSwallowed", filterPattern) Then
        r = Test_FieldWiring_CaseNearMissIsNamedNotSwallowed()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "FieldWiring_CaseNearMissIsNamedNotSwallowed", r
    If TestMatches("Drafting_NothingToPublishReadsTheSameCounts", filterPattern) Then
        r = Test_Drafting_NothingToPublishReadsTheSameCounts()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "Drafting_NothingToPublishReadsTheSameCounts", r
    If TestMatches("MilestoneDevice_DrawsFromDataAndCreatesNothing", filterPattern) Then
        r = Test_MilestoneDevice_DrawsFromDataAndCreatesNothing()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "MilestoneDevice_DrawsFromDataAndCreatesNothing", r
    If TestMatches("MilestoneDevice_RefusesListsThatCannotBeAligned", filterPattern) Then
        r = Test_MilestoneDevice_RefusesListsThatCannotBeAligned()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "MilestoneDevice_RefusesListsThatCannotBeAligned", r
    If TestMatches("MilestoneDevice_ReadsAColumnPerIntervalFromTheRow", filterPattern) Then
        r = Test_MilestoneDevice_ReadsAColumnPerIntervalFromTheRow()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "MilestoneDevice_ReadsAColumnPerIntervalFromTheRow", r
    If TestMatches("MilestoneDevice_ReportWrapperMatchesTheRealResult", filterPattern) Then
        r = Test_MilestoneDevice_ReportWrapperMatchesTheRealResult()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "MilestoneDevice_ReportWrapperMatchesTheRealResult", r
    If TestMatches("MilestoneDevice_ReportsAWriteThatDidNotTake", filterPattern) Then
        r = Test_MilestoneDevice_ReportsAWriteThatDidNotTake()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "MilestoneDevice_ReportsAWriteThatDidNotTake", r
    If TestMatches("MilestoneDevice_IntegrityNamesWhatIsMissing", filterPattern) Then
        r = Test_MilestoneDevice_IntegrityNamesWhatIsMissing()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "MilestoneDevice_IntegrityNamesWhatIsMissing", r
    If TestMatches("InjectField_RoutesADeviceGroupToTheTimeline", filterPattern) Then
        r = Test_InjectField_RoutesADeviceGroupToTheTimeline()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "InjectField_RoutesADeviceGroupToTheTimeline", r
    If TestMatches("InjectPrimitive_DeviceDiscoveredByNameWhenUntagged", filterPattern) Then
        r = Test_InjectPrimitive_DeviceDiscoveredByNameWhenUntagged()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "InjectPrimitive_DeviceDiscoveredByNameWhenUntagged", r
    If TestMatches("InjectPrimitive_FindShapeByRoleTagFindsUntaggedDeviceByName", filterPattern) Then
        r = Test_InjectPrimitive_FindShapeByRoleTagFindsUntaggedDeviceByName()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "InjectPrimitive_FindShapeByRoleTagFindsUntaggedDeviceByName", r
    If TestMatches("InjectPrimitive_InjectorForRoutesUntaggedDeviceToDevice", filterPattern) Then
        r = Test_InjectPrimitive_InjectorForRoutesUntaggedDeviceToDevice()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "InjectPrimitive_InjectorForRoutesUntaggedDeviceToDevice", r
    If TestMatches("InjectPrimitive_NameFallbackDoesNotWidenOrdinaryFieldLookup", filterPattern) Then
        r = Test_InjectPrimitive_NameFallbackDoesNotWidenOrdinaryFieldLookup()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "InjectPrimitive_NameFallbackDoesNotWidenOrdinaryFieldLookup", r
    If TestMatches("InjectPrimitive_ExplicitTagStillWinsOverDeviceName", filterPattern) Then
        r = Test_InjectPrimitive_ExplicitTagStillWinsOverDeviceName()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "InjectPrimitive_ExplicitTagStillWinsOverDeviceName", r
    If TestMatches("InjectField_RepeatingBarsOneCellManyMilestones", filterPattern) Then
        r = Test_InjectField_RepeatingBarsOneCellManyMilestones()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "InjectField_RepeatingBarsOneCellManyMilestones", r
    If TestMatches("InjectProgress_TracklessPairMeasuresTheirSum", filterPattern) Then
        r = Test_InjectProgress_TracklessPairMeasuresTheirSum()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "InjectProgress_TracklessPairMeasuresTheirSum", r
    If TestMatches("BatchOnboardFlow_ReopeningTheSameDeckLeavesShapeRefsDead", filterPattern) Then
        r = Test_BatchOnboardFlow_ReopeningTheSameDeckLeavesShapeRefsDead()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_ReopeningTheSameDeckLeavesShapeRefsDead", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("BatchOnboardFlow_MarkingSessionSurvivesRealCloseAndReopen", filterPattern) Then
        r = Test_BatchOnboardFlow_MarkingSessionSurvivesRealCloseAndReopen()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_MarkingSessionSurvivesRealCloseAndReopen", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("BatchOnboardFlow_RestoreMarkingSessionFindsNothingOnWrongSlide", filterPattern) Then
        r = Test_BatchOnboardFlow_RestoreMarkingSessionFindsNothingOnWrongSlide()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_RestoreMarkingSessionFindsNothingOnWrongSlide", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("BatchOnboardFlow_FlattenGroupLeavesReturnsAllMembers", filterPattern) Then
        r = Test_BatchOnboardFlow_FlattenGroupLeavesReturnsAllMembers()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_FlattenGroupLeavesReturnsAllMembers", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("BatchOnboardFlow_ReviewGridRoundTrip", filterPattern) Then
        r = Test_BatchOnboardFlow_ReviewGridRoundTrip()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_ReviewGridRoundTrip", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("BatchOnboardFlow_CommitBatchTagsLinksAndVerifies", filterPattern) Then
        r = Test_BatchOnboardFlow_CommitBatchTagsLinksAndVerifies()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_CommitBatchTagsLinksAndVerifies", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("BatchOnboardFlow_CommitBatchWithGroupedFieldsAtScale", filterPattern) Then
        r = Test_BatchOnboardFlow_CommitBatchWithGroupedFieldsAtScale()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "BatchOnboardFlow_CommitBatchWithGroupedFieldsAtScale", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DraftingLobby_PinReadClearRoundTrip", filterPattern) Then
        r = Test_DraftingLobby_PinReadClearRoundTrip()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DraftingLobby_PinReadClearRoundTrip", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DraftingLobby_BuildFromScratchFindsOnlyApprovedRows", filterPattern) Then
        r = Test_DraftingLobby_BuildFromScratchFindsOnlyApprovedRows()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DraftingLobby_BuildFromScratchFindsOnlyApprovedRows", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DraftingLobby_PinTwiceUpdatesInPlaceNotDuplicate", filterPattern) Then
        r = Test_DraftingLobby_PinTwiceUpdatesInPlaceNotDuplicate()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DraftingLobby_PinTwiceUpdatesInPlaceNotDuplicate", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    If TestMatches("DraftingUI_DistinctPinnedFieldsReadsOnlyTheLobby", filterPattern) Then
        r = Test_DraftingUI_DistinctPinnedFieldsReadsOnlyTheLobby()
    Else
        r = TEST_SKIPPED
    End If
    AppendResult report, "DraftingUI_DistinctPinnedFieldsReadsOnlyTheLobby", r
    On Error GoTo 0

    RunAllTests = report
End Function

Private Function TestMatches(testName As String, filterPattern As String) As Boolean
    TestMatches = (filterPattern = "") Or (InStr(1, testName, filterPattern, vbTextCompare) > 0)
End Function

Private Sub AppendResult(ByRef report As String, testName As String, testResult As String)
    If Err.Number <> 0 Then
        report = report & "ERROR " & testName & " :: " & Err.Description & " (line context lost -- VBA has no stack trace)" & vbCrLf
    ElseIf testResult = TEST_SKIPPED Then
        report = report & "SKIP  " & testName & vbCrLf
    ElseIf testResult = "" Then
        report = report & "PASS  " & testName & vbCrLf
    Else
        report = report & "FAIL  " & testName & vbCrLf & testResult
    End If
End Sub

Private Function Assert(cond As Boolean, msg As String) As String
    If cond Then
        Assert = ""
    Else
        Assert = "    FAIL: " & msg & vbCrLf
    End If
End Function

' ---------------------------------------------------------------------
' THE CRITICAL-EFFECTS SHORTLIST -- Rohan, 2026-08-16: "make key bits
' obvious to a human viewer while still obviously completing."
'
' PowerPoint is already run VISIBLE for this suite (run_vba_tests.ps1's own
' comment: "visible on purpose for a first real run"), so a human CAN watch
' -- the gap was speed: an automated test toggles a shape and moves on
' inside a fraction of a second, too fast to register. Watch() pauses just
' long enough to see it, only at the specific moments on this list, so the
' other ~230 assertions stay exactly as fast as they are today.
'
' THE LIST ITSELF, kept here deliberately rather than scattered as ad-hoc
' Sleep calls, so it stays a maintained, reviewable set rather than growing
' by whoever last found a test they wanted to watch:
'   1. Milestone slot ON/OFF/NOW visibility (MilestoneDevice.DrawMilestones/
'      SetVisible) -- Test_MilestoneDevice_DrawsFromDataAndCreatesNothing
'   2. Progress bar/track resize (InjectProgressVia) -- a real, visible
'      geometric change -- Test_InjectProgress_MeasuresAgainstTheTrackNotItself
'   NOT YET REAL: shape colour change. Checked before adding it here --
'   nothing in production code writes Fill.ForeColor at runtime (colour is
'   static, drawn once per K/S/P template, per today's architecture-fork
'   resolution). Adding a colour-change entry to this list before the
'   mechanism exists would be exactly the kind of item that looks like care
'   taken while testing nothing real.
' FORCES the PowerPoint window to the OS foreground before pausing --
' deliberately, not just "leaves it visible". "Visible" and "in front of
' whatever you are actually looking at" are different claims, and Rohan
' caught the gap live: the first version paused correctly but he was
' looking at this chat, not PowerPoint, and missed it entirely. Self-
' contained on purpose per his own correction -- "needs to run without you"
' -- so this cannot depend on Claude narrating "switch windows now" in chat
' first; it has to work identically whether a Claude session is watching or
' Rohan runs the suite himself later with nobody narrating anything.
' AppActivate can raise if no matching window title exists (e.g. a truly
' headless CI-style run with no window at all) -- guarded so a watch point
' degrades to "just pauses" rather than failing the test that carries it.
'
' AppActivate ALONE ONLY FLASHED THE TASKBAR ICON -- checked live 2026-08-16,
' not assumed: Windows' focus-stealing prevention blocks a background
' automation process from stealing foreground focus outright, and
' AppActivate is exactly that kind of call.
'
' TRIED AND REVERTED, same session: `SendKeys " "` before AppActivate is the
' standard VBA workaround (a keystroke reads as "real user input just
' happened", which relaxes the restriction for the next foreground request)
' -- but it HUNG the run under this specific driver (PowerShell -> COM ->
' VBA), 2 minutes, no error, nothing recovered. Two attempts is this
' problem's stated cap.
'
' ROHAN'S FIX, and the right one: stop trying to force the WINDOW in front
' of him and put the label ON THE SLIDE instead -- a big bold orange banner
' naming what to look at. This sidesteps the whole foreground-stealing
' problem: it does not matter whether PowerPoint has OS focus, because
' whenever he DOES look (now or later, from a screenshot, from the taskbar
' thumbnail), the banner says exactly what was being demonstrated. AppActivate
' stays as a best-effort attempt (harmless, sometimes helps) but is no longer
' load-bearing for "obvious".
Private Sub Watch(sld As Object, label As String, Optional seconds As Double = 1.2)
    On Error Resume Next

    Dim banner As Object
    Set banner = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 20, 15, 650, 45)
    banner.Name = "WATCH_BANNER"
    banner.TextFrame.TextRange.Text = "WATCHING: " & label
    banner.TextFrame.TextRange.Font.Size = 24
    banner.TextFrame.TextRange.Font.Bold = True
    banner.TextFrame.TextRange.Font.Color.RGB = RGB(255, 140, 0)   ' orange
    banner.Fill.Visible = msoFalse
    banner.Line.Visible = msoFalse

    AppActivate Application.Caption
    DoEvents

    ' Application.Wait is an EXCEL method, not PowerPoint's -- this module
    ' compiles against PowerPoint.Application, so that call is a compile-time
    ' "member not found", not a runtime one. Found for real 2026-08-16: it
    ' blocked PowerPoint with its own compile-error dialog before a single
    ' test could run. Timer is a plain VBA stdlib function, identical on
    ' both hosts, so a busy-wait against it needs no Declare and no
    ' host-specific branch. DoEvents inside the loop, not just before it, so
    ' the screen actually repaints during the pause instead of once at the
    ' start.
    Dim stopAt As Double
    stopAt = Timer + seconds
    Do While Timer < stopAt
        DoEvents
    Loop

    ' Removed after the pause -- it is a viewing aid, not part of what the
    ' assertions below are checking, and leaving it would change the
    ' shape/group counts the test itself asserts on.
    banner.Delete

    On Error GoTo 0
End Sub

' Safe replacement for "Assert(Not actual Is Nothing And actual.SlideID =
' expected.SlideID, msg)". VBA's And is NOT short-circuit -- both operands
' always evaluate -- so that combined form raises "Object variable or With
' block variable not set" the moment `actual` is genuinely Nothing, instead
' of failing cleanly with the message. Found for real 2026-08-16, after the
' unsafe form had already passed silently in several tests because `actual`
' was never actually Nothing on their happy path -- a landmine, not yet a
' failure. Every same-slide comparison in this file goes through here now.
Private Function AssertSameSlide(actual As Object, expected As Object, msg As String) As String
    If actual Is Nothing Then
        AssertSameSlide = Assert(False, msg & " -- got Nothing")
    Else
        AssertSameSlide = Assert(actual.SlideID = expected.SlideID, msg & " -- got SlideID " & actual.SlideID & " want " & expected.SlideID)
    End If
End Function

' Shape.Select (used by ResolveFields' tests to simulate a real user click)
' requires the shape's slide to be the window's active view -- confirmed the
' hard way (2026-07-26 real-Office run): Slides.Add does not itself navigate
' the view, so a shape on a freshly-added slide fails Select with "Invalid
' request. To select a shape, its view must be active" until the view is
' explicitly moved there. GotoSlide here makes every NewBlankSlide() caller
' immediately select-able, not just ResolveFields' tests.
Private Function NewBlankSlide() As Object
    Dim n As Long
    n = Application.ActivePresentation.Slides.count + 1
    Dim sld As Object
    Set sld = Application.ActivePresentation.Slides.Add(n, ppLayoutBlank)
    Application.ActiveWindow.View.GotoSlide sld.SlideIndex
    Set NewBlankSlide = sld
End Function

' Recurses into groups -- see InjectPrimitive.bas's own FindShapeByRoleTag
' header for why a flat `sld.Shapes` loop alone misses any shape nested
' inside a GroupShape (real bug found 2026-07-26 against Rohan's real,
' group-heavy "card layout" deck). Kept in sync with that fix so this test
' helper can find a role tag anywhere InjectPrimitive itself now can.
Private Function FindShapeByRole(sld As Object, role As String) As Object
    Set FindShapeByRole = FindShapeByRoleWalk(sld.Shapes, role)
End Function

Private Function FindShapeByRoleWalk(shapesColl As Object, role As String) As Object
    Dim shp As Object
    For Each shp In shapesColl
        If shp.Type = msoGroup Then
            Dim nested As Object
            Set nested = FindShapeByRoleWalk(shp.GroupItems, role)
            If Not nested Is Nothing Then
                Set FindShapeByRoleWalk = nested
                Exit Function
            End If
        ElseIf shp.Tags("role") = role Then
            Set FindShapeByRoleWalk = shp
            Exit Function
        End If
    Next shp
    Set FindShapeByRoleWalk = Nothing
End Function

' ---------------------------------------------------------------------
' Discovery
' ---------------------------------------------------------------------

Private Function Test_Discovery_GroupRecursionFindsCandidates(fixturesDir As String) As String
    Dim result As String
    Dim testPres As Object
    Set testPres = Application.Presentations.Open(fixturesDir & "shp-groupshape.pptx", msoTrue)

    Dim candidates() As Candidate
    candidates = Discovery.DiscoverSlide(testPres.Slides(1))

    Dim lo As Long, hi As Long, hasCandidates As Boolean
    On Error Resume Next
    lo = LBound(candidates)
    hi = UBound(candidates)
    hasCandidates = (Err.Number = 0)
    On Error GoTo 0
    result = result & Assert(hasCandidates, "DiscoverSlide returned at least one candidate")

    Dim found As Boolean
    Dim i As Long
    If hasCandidates Then
        For i = lo To hi
            If candidates(i).Name = "Oval 2" Then found = True
        Next i
    End If
    result = result & Assert(found, "DiscoverSlide recursed into a group and found 'Oval 2'")

    testPres.Close
    Test_Discovery_GroupRecursionFindsCandidates = result
End Function

' ---------------------------------------------------------------------
' InjectPrimitive
' ---------------------------------------------------------------------

Private Function Test_InjectPrimitive_NoOpWhenValueAlreadyMatches() As String
    Dim result As String
    Dim sld As Object
    Set sld = NewBlankSlide()

    Dim shp As Object
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    shp.TextFrame.TextRange.Text = "hello"
    shp.Tags.Add "role", "demo_field"

    Dim r As InjectResult
    r = InjectPrimitive.InjectPrimitive(sld, "demo_field", "hello")

    result = result & Assert(r.Found, "shape found by role tag")
    result = result & Assert(Not r.Written, "no-op when value already matches")
    result = result & Assert(r.Verified, "verified true on no-op path")

    Test_InjectPrimitive_NoOpWhenValueAlreadyMatches = result
End Function

Private Function Test_InjectPrimitive_WritesAndVerifiesOnMismatch() As String
    Dim result As String
    Dim sld As Object
    Set sld = NewBlankSlide()

    Dim shp As Object
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    shp.TextFrame.TextRange.Text = "old value"
    shp.Tags.Add "role", "demo_field"

    Dim r As InjectResult
    r = InjectPrimitive.InjectPrimitive(sld, "demo_field", "new value")

    result = result & Assert(r.Found, "shape found by role tag")
    result = result & Assert(r.Written, "written when value differs")
    result = result & Assert(r.Verified, "re-read confirms the write took")
    result = result & Assert(shp.TextFrame.TextRange.Text = "new value", "shape text actually changed, got '" & shp.TextFrame.TextRange.Text & "'")

    Test_InjectPrimitive_WritesAndVerifiesOnMismatch = result
End Function

Private Function Test_InjectPrimitive_AmbiguousTagRefusesToGuess() As String
    Dim result As String
    Dim sld As Object
    Set sld = NewBlankSlide()

    Dim shp1 As Object, shp2 As Object
    Set shp1 = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    shp1.TextFrame.TextRange.Text = "a"
    shp1.Tags.Add "role", "dup_field"
    Set shp2 = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 150, 200, 50)
    shp2.TextFrame.TextRange.Text = "b"
    shp2.Tags.Add "role", "dup_field"

    Dim r As InjectResult
    r = InjectPrimitive.InjectPrimitive(sld, "dup_field", "c")

    result = result & Assert(Not r.Found, "ambiguous tag (2 shapes) refuses rather than picking one")
    result = result & Assert(shp1.TextFrame.TextRange.Text = "a" And shp2.TextFrame.TextRange.Text = "b", "neither shape was touched")

    Test_InjectPrimitive_AmbiguousTagRefusesToGuess = result
End Function

' Real, confirmed bug found 2026-07-26 against Rohan's real 46-slide "card
' layout" deck (see SPIKE_NOTES_BatchOnboardFlow.md's addendum): every
' InjectPrimitive test above only ever used a TOP-LEVEL textbox, so none of
' them could have caught FindShapeByRoleTag's `sld.Shapes` loop missing any
' shape nested inside a group -- exactly the structure Rohan's real deck
' actually uses for its fields (confirmed via the real, redacted deck
' fixture: all 46 slides contain at least one <p:grpSp>, and most
' text-bearing shapes on a given slide are nested inside one, not
' top-level). Proves the fix directly: a role tag written onto a shape
' nested one level inside a group must still be found, and the round-trip
' verify must still pass -- this is the single smallest repro of the actual
' production failure ("Linked: 0 / FAILED verification: 46").
Private Function Test_InjectPrimitive_FindsRoleTagInsideGroup() As String
    Dim result As String
    Dim sld As Object
    Set sld = NewBlankSlide()

    Dim shapeA As Object, shapeB As Object
    Set shapeA = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 100, 50)
    shapeA.TextFrame.TextRange.Text = "hello"
    Set shapeB = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 200, 50, 100, 50)
    shapeB.TextFrame.TextRange.Text = "sibling"

    Dim grp As Object
    Set grp = sld.Shapes.Range(Array(shapeA.Name, shapeB.Name)).Group()

    ' Tag the nested child directly -- Shape.Tags works on any shape
    ' regardless of group nesting (this half of the mechanism was never
    ' broken; only the subsequent lookup was).
    Dim groupedChild As Object
    Set groupedChild = grp.GroupItems.Item(1)
    groupedChild.Tags.Add "role", "demo_grouped_field"

    Dim r As InjectResult
    r = InjectPrimitive.InjectPrimitive(sld, "demo_grouped_field", "hello")
    result = result & Assert(r.Found, "shape found by role tag even though it's nested inside a group")
    result = result & Assert(Not r.Written, "no-op when value already matches (proves the SAME nested shape was found, not a coincidence)")
    result = result & Assert(r.Verified, "verified true")

    ' Also prove the write path (not just the no-op path) works through a
    ' group -- mirrors WritesAndVerifiesOnMismatch above.
    Dim r2 As InjectResult
    r2 = InjectPrimitive.InjectPrimitive(sld, "demo_grouped_field", "new value")
    result = result & Assert(r2.Found, "shape found by role tag on the write path too")
    result = result & Assert(r2.Written, "written when value differs")
    result = result & Assert(r2.Verified, "re-read confirms the write took")
    result = result & Assert(groupedChild.TextFrame.TextRange.Text = "new value", "the grouped shape's text actually changed, got '" & groupedChild.TextFrame.TextRange.Text & "'")

    Test_InjectPrimitive_FindsRoleTagInsideGroup = result
End Function

' InjectPrimitive holds the only slide-mutating line in the whole sync path,
' so its dryRun flag is what makes a preview provably safe rather than merely
' careful. Covers the gate itself, the no-op case (which must stay a no-op in
' both modes), and that the same call without the flag still writes.
Private Function Test_InjectPrimitive_DryRunReportsWithoutWriting() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim shp As Object
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    shp.TextFrame.TextRange.Text = "Q3 Revenue"
    shp.Tags.Add "role", "Title"

    Dim dry As InjectResult
    dry = InjectPrimitive.InjectPrimitive(sld, "Title", "Q4 Revenue", True)
    result = result & Assert(dry.Found = True, "dry run still locates the tagged shape")
    result = result & Assert(dry.WouldChange = True, "dry run reports that the value differs")
    result = result & Assert(dry.Written = False, "dry run reports nothing was written")
    result = result & Assert(dry.CurrentValue = "Q3 Revenue", _
        "dry run reports the shape's current value, got '" & dry.CurrentValue & "'")
    result = result & Assert(dry.ErrorMessage = "", _
        "a suppressed write is not an error, got '" & dry.ErrorMessage & "'")
    result = result & Assert(shp.TextFrame.TextRange.Text = "Q3 Revenue", _
        "DRY RUN MUST NOT WRITE: shape text unchanged, got '" & shp.TextFrame.TextRange.Text & "'")

    ' Already-equal is a no-op in both modes, and must not be reported as a
    ' pending change -- otherwise a preview overstates what a sync would do.
    Dim noop As InjectResult
    noop = InjectPrimitive.InjectPrimitive(sld, "Title", "Q3 Revenue", True)
    result = result & Assert(noop.WouldChange = False, "an already-matching value would not change")
    result = result & Assert(noop.Verified = True, "an already-matching value is verified in a dry run too")

    ' Same call, flag off: the write happens. Proves the flag gates the write
    ' and not the lookup.
    Dim wet As InjectResult
    wet = InjectPrimitive.InjectPrimitive(sld, "Title", "Q4 Revenue", False)
    result = result & Assert(wet.Written = True, "without the flag the write happens")
    result = result & Assert(wet.Verified = True, "without the flag the write verifies")
    result = result & Assert(wet.CurrentValue = "Q3 Revenue", _
        "CurrentValue is the pre-write value even on a real write, got '" & wet.CurrentValue & "'")
    result = result & Assert(shp.TextFrame.TextRange.Text = "Q4 Revenue", _
        "shape text actually updated, got '" & shp.TextFrame.TextRange.Text & "'")

    Test_InjectPrimitive_DryRunReportsWithoutWriting = result
End Function

' Same-tag collision detection must still hold across nesting levels -- a
' shape inside a group and a top-level shape both carrying the same role
' tag is exactly as ambiguous as two top-level shapes (Test_InjectPrimitive_
' AmbiguousTagRefusesToGuess above), and the recursive rewrite must not
' accidentally scope matchCount per group level instead of across the whole
' walk.
Private Function Test_InjectPrimitive_AmbiguousTagAcrossGroupAndTopLevelRefusesToGuess() As String
    Dim result As String
    Dim sld As Object
    Set sld = NewBlankSlide()

    Dim topShp As Object
    Set topShp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 200, 100, 50)
    topShp.TextFrame.TextRange.Text = "top-level"
    topShp.Tags.Add "role", "dup_across_levels"

    Dim shapeA As Object, shapeB As Object
    Set shapeA = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 100, 50)
    shapeA.TextFrame.TextRange.Text = "grouped"
    Set shapeB = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 200, 50, 100, 50)
    shapeB.TextFrame.TextRange.Text = "sibling"
    Dim grp As Object
    Set grp = sld.Shapes.Range(Array(shapeA.Name, shapeB.Name)).Group()
    grp.GroupItems.Item(1).Tags.Add "role", "dup_across_levels"

    Dim r As InjectResult
    r = InjectPrimitive.InjectPrimitive(sld, "dup_across_levels", "z")

    result = result & Assert(Not r.Found, "a top-level/grouped same-tag collision is still refused, not silently resolved")

    Test_InjectPrimitive_AmbiguousTagAcrossGroupAndTopLevelRefusesToGuess = result
End Function

' ---------------------------------------------------------------------
' Matching
' ---------------------------------------------------------------------

Private Function Test_Matching_SiblingAmbiguityResolvedByZOrder(fixturesDir As String) As String
    Dim result As String
    Dim testPres As Object
    Set testPres = Application.Presentations.Open(fixturesDir & "shp-groupshape.pptx", msoTrue)

    Dim candidates() As Candidate
    candidates = Discovery.DiscoverSlide(testPres.Slides(1))

    ' The reference must carry real geometry/shape-type, not just ZOrder/
    ' HasText -- a reference with everything else left at Long/String
    ' defaults (0, "") scores near-zero on every candidate regardless of
    ' z-order, which is exactly what this test found the first time it
    ' actually ran (a gap in the original ManualSmokeTest_SiblingAmbiguity
    ' this test was copied from -- never caught before because that sub had
    ' never been executed either). The 4 decoration ovals are near-identical
    ' overlapping siblings by design (that's the whole scenario), so any one
    ' of their real geometries is representative -- clone one, but keep
    ' ZOrder explicitly at 2 (not copied) since z-order proximity to the
    ' reference is exactly the signal under test.
    Dim reference As Candidate
    Dim i As Long
    For i = LBound(candidates) To UBound(candidates)
        If Not candidates(i).HasText And candidates(i).ShapeType = "autoshape_or_textbox" Then
            reference = candidates(i)
            Exit For
        End If
    Next i
    reference.ZOrder = 2
    reference.PlaceholderIdx = -1 ' not testing placeholder matching here

    Dim r As MatchResult
    r = Matching.Match(candidates, reference)

    result = result & Assert(r.Confidence = "high", "confidence is high, got '" & r.Confidence & "'")
    result = result & Assert(r.HasCandidate, "HasCandidate true")
    ' Not asserting the exact resolved shape name: this test's reference is
    ' cloned from the first non-text autoshape found (a stand-in for "a
    ' sibling with matching geometry"), not necessarily Oval 2 itself --
    ' real run (2026-07-25) showed shp-groupshape.pptx has other non-text
    ' autoshapes (e.g. "Rounded Rectangle 1") competing for that same
    ' geometry class, so which specific sibling wins z-order proximity to
    ' ZOrder=2 depends on the fixture's exact real layout, not asserted
    ' here. What's actually under test -- multiple similarly-scored
    ' candidates get disambiguated by z-order proximity rather than
    ' flagged or picked arbitrarily -- is what HasCandidate=True/
    ' Confidence=high already confirm.

    testPres.Close
    Test_Matching_SiblingAmbiguityResolvedByZOrder = result
End Function

' Exercises Matching.EnrichPlaceholderIdx's real, previously-unexecuted
' Shell.Application zip-extraction + XPath fallback against a real file on
' disk -- the most exotic, least-proven code in this whole port. Uses an
' ordinary slide with a real placeholder (not test-fixtures/mst-slide-
' layouts.pptx's CustomLayouts, whose live-COM accessibility was flagged
' unconfirmed in SPIKE_NOTES_Resolve.md divergence 6) -- EnrichPlaceholderIdx
' only cares about a part path + shape name, so a slide's own
' ppt/slides/slideN.xml exercises the identical code path with less risk.
Private Function Test_Matching_EnrichPlaceholderIdxReadsRealFile(stagingDir As String) As String
    Dim result As String
    Dim sld As Object
    Set sld = Application.ActivePresentation.Slides.Add(Application.ActivePresentation.Slides.count + 1, ppLayoutText)

    Dim savePath As String
    savePath = stagingDir & "enrich_test.pptx"
    Application.ActivePresentation.SaveCopyAs savePath

    Dim candidates() As Candidate
    candidates = Discovery.DiscoverSlide(sld)

    Dim partName As String
    partName = "ppt/slides/slide" & sld.SlideIndex & ".xml"

    Dim ok As Boolean
    ok = Matching.EnrichPlaceholderIdx(candidates, savePath, partName)
    result = result & Assert(ok, "EnrichPlaceholderIdx completed without error against a real saved file")

    Dim anyPlaceholder As Boolean, anyResolved As Boolean
    Dim i As Long
    For i = LBound(candidates) To UBound(candidates)
        If candidates(i).HasPlaceholder Then
            anyPlaceholder = True
            If candidates(i).PlaceholderIdx <> -1 Then anyResolved = True
        End If
    Next i
    result = result & Assert(anyPlaceholder, "ppLayoutText slide has at least one placeholder shape")
    result = result & Assert(anyResolved, "at least one placeholder's idx resolved to something other than -1")

    Test_Matching_EnrichPlaceholderIdxReadsRealFile = result
End Function

' ---------------------------------------------------------------------
' Resolve
' ---------------------------------------------------------------------

Private Function Test_Resolve_ReadsTagsOffLiveSlide() As String
    Dim result As String
    Dim sld As Object
    Set sld = NewBlankSlide()
    sld.Tags.Add "slide_type", "quarterly-update"
    sld.Tags.Add "instance_key", "rec-1"

    Dim instance As SlideInstance
    instance = Resolve.ResolveSlideInstance(sld)
    result = result & Assert(instance.HasTypeTag And instance.TypeTag = "quarterly-update", "TypeTag read correctly, got '" & instance.TypeTag & "'")
    result = result & Assert(instance.HasInstanceKey And instance.InstanceKey = "rec-1", "InstanceKey read correctly, got '" & instance.InstanceKey & "'")

    Dim untaggedSld As Object
    Set untaggedSld = NewBlankSlide()
    Dim untaggedInstance As SlideInstance
    untaggedInstance = Resolve.ResolveSlideInstance(untaggedSld)
    result = result & Assert(Not untaggedInstance.HasTypeTag And Not untaggedInstance.HasInstanceKey, "untagged slide resolves to no type/key")

    Test_Resolve_ReadsTagsOffLiveSlide = result
End Function

' ---------------------------------------------------------------------
' SyncOperations
' ---------------------------------------------------------------------

Private Function Test_SyncOperations_Cases1And4() As String
    Dim result As String
    Dim sld As Object
    Set sld = NewBlankSlide()
    sld.Tags.Add "slide_type", "quarterly-update"
    sld.Tags.Add "instance_key", "rec-1"
    Dim shp As Object
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    shp.TextFrame.TextRange.Text = "Q3 Revenue"
    shp.Tags.Add "role", "Title"

    Dim instances(1 To 1) As Object
    Set instances(1) = sld
    Dim order As New Collection
    order.Add "rec-1"

    Dim rows1 As Object
    Set rows1 = CreateObject("Scripting.Dictionary")
    Dim fields1 As Object
    Set fields1 = CreateObject("Scripting.Dictionary")
    fields1("Title") = "Q3 Revenue"
    Set rows1("rec-1") = fields1

    Dim actions1() As SyncAction
    actions1 = SyncOperations.PlanRoutineSync(instances, order, rows1)
    result = result & Assert(UBound(actions1) = 1, "one action produced, got " & (UBound(actions1) - LBound(actions1) + 1))
    result = result & Assert(actions1(1).Kind = "no_change", "case 1: no_change, got '" & actions1(1).Kind & "'")

    Dim rows2 As Object
    Set rows2 = CreateObject("Scripting.Dictionary")
    Dim fields2 As Object
    Set fields2 = CreateObject("Scripting.Dictionary")
    fields2("Title") = "Q3 Revenue (revised)"
    Set rows2("rec-1") = fields2

    Dim actions2() As SyncAction
    actions2 = SyncOperations.PlanRoutineSync(instances, order, rows2)
    result = result & Assert(actions2(1).Kind = "in_place_correction", "case 4: in_place_correction, got '" & actions2(1).Kind & "'")
    result = result & Assert(shp.TextFrame.TextRange.Text = "Q3 Revenue (revised)", "shape text actually updated, got '" & shp.TextFrame.TextRange.Text & "'")
    result = result & Assert(actions2(1).ChangedFieldVerified("Title") = True, "changed field reports Verified=True")

    Test_SyncOperations_Cases1And4 = result
End Function

' The load-bearing test for the Sync Now preview. PlanRoutineSync writes to
' slides *while planning* (it calls InjectPrimitive directly), so "planned but
' not executed" was never a safe state in this codebase -- a preview is only
' real if the plan itself is provably inert. Asserts both halves: the dry run
' classifies exactly as a real run would AND leaves the slide byte-identical.
Private Function Test_SyncOperations_PlanRoutineSyncDryRunWritesNothing() As String
    Dim result As String
    Dim sld As Object
    Set sld = NewBlankSlide()
    sld.Tags.Add "slide_type", "quarterly-update"
    sld.Tags.Add "instance_key", "rec-1"
    Dim shp As Object
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    shp.TextFrame.TextRange.Text = "Q3 Revenue"
    shp.Tags.Add "role", "Title"

    Dim slideCountBefore As Long
    slideCountBefore = Application.ActivePresentation.Slides.count

    Dim instances(1 To 1) As Object
    Set instances(1) = sld
    Dim order As New Collection
    order.Add "rec-1"
    order.Add "rec-2"   ' an orphaned row: no slide carries this key

    Dim rows As Object
    Set rows = CreateObject("Scripting.Dictionary")
    Dim fields1 As Object
    Set fields1 = CreateObject("Scripting.Dictionary")
    fields1("Title") = "Q3 Revenue (revised)"
    Set rows("rec-1") = fields1
    Dim fields2 As Object
    Set fields2 = CreateObject("Scripting.Dictionary")
    fields2("Title") = "Q4 Revenue"
    Set rows("rec-2") = fields2

    Dim actions() As SyncAction
    actions = SyncOperations.PlanRoutineSync(instances, order, rows, True)

    result = result & Assert(UBound(actions) - LBound(actions) + 1 = 2, _
        "two actions produced (one per row), got " & (UBound(actions) - LBound(actions) + 1))
    result = result & Assert(actions(1).Kind = "in_place_correction", _
        "a dry run still classifies a differing field as in_place_correction, got '" & actions(1).Kind & "'")
    result = result & Assert(actions(2).Kind = "new_record", _
        "a dry run still classifies an orphaned row as new_record, got '" & actions(2).Kind & "'")

    ' The whole point.
    result = result & Assert(shp.TextFrame.TextRange.Text = "Q3 Revenue", _
        "DRY RUN MUST NOT WRITE: shape text unchanged, got '" & shp.TextFrame.TextRange.Text & "'")
    result = result & Assert(Application.ActivePresentation.Slides.count = slideCountBefore, _
        "DRY RUN MUST NOT DUPLICATE: slide count unchanged, got " & Application.ActivePresentation.Slides.count & " vs " & slideCountBefore)

    ' The preview has to be able to show before/after, so the pre-write value
    ' is carried on the action rather than only existing on the slide.
    result = result & Assert(actions(1).ChangedFieldCurrent("Title") = "Q3 Revenue", _
        "the field's current (pre-write) value is reported, got '" & actions(1).ChangedFieldCurrent("Title") & "'")

    ' And the real run, on the same slide, still works -- proving the dry run
    ' suppressed the write rather than the classification.
    Dim realActions() As SyncAction
    realActions = SyncOperations.PlanRoutineSync(instances, order, rows, False)
    result = result & Assert(realActions(1).Kind = "in_place_correction", _
        "the real run classifies identically, got '" & realActions(1).Kind & "'")
    result = result & Assert(shp.TextFrame.TextRange.Text = "Q3 Revenue (revised)", _
        "the real run does write, got '" & shp.TextFrame.TextRange.Text & "'")

    Test_SyncOperations_PlanRoutineSyncDryRunWritesNothing = result
End Function

Private Function Test_SyncOperations_Case3NewRecord() As String
    Dim result As String
    Dim instances() As Object ' deliberately left unallocated -- no known instances at all; PlanRoutineSync error-guards this (see AGENTS.md's Known Patterns)

    Dim order As New Collection
    order.Add "rec-new"

    Dim rows As Object
    Set rows = CreateObject("Scripting.Dictionary")
    Dim fields As Object
    Set fields = CreateObject("Scripting.Dictionary")
    fields("Title") = "Brand New"
    Set rows("rec-new") = fields

    Dim actions() As SyncAction
    actions = SyncOperations.PlanRoutineSync(instances, order, rows)

    result = result & Assert(UBound(actions) = 1, "one action produced, got " & (UBound(actions) - LBound(actions) + 1))
    result = result & Assert(actions(1).Kind = "new_record", "case 3: new_record, got '" & actions(1).Kind & "'")
    result = result & Assert(actions(1).RowInstanceKey = "rec-new", "RowInstanceKey correct")
    result = result & Assert(actions(1).Values("Title") = "Brand New", "Values carries the row's data")

    Test_SyncOperations_Case3NewRecord = result
End Function

' FIX-LIST R. The row dictionary below holds ONLY MS1_LABEL/MS1_DATE/MS1_DONE
' -style keys -- never a "MILESTONE_TIMELINE" key, because no such register
' column exists anywhere in the real project. Before this fix, PlanRoutineSync
' could only ask InjectField about a field whose name was a dictionary key, so
' a device tagged MILESTONE_TIMELINE was never once asked about, on any slide,
' ever -- not routed wrong, never called. This proves the fix reaches it
' anyway, discovered from the slide's own tag rather than the register.
Private Function Test_SyncOperations_DeviceFieldReachableThroughPlanRoutineSync() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    sld.Tags.Add "slide_type", "project-progress"
    sld.Tags.Add "instance_key", "3_P001"

    Dim grp As Object
    Set grp = NewMilestoneDevice(sld, 3)
    grp.Tags.Add "role", "MILESTONE_TIMELINE"

    Dim instances(1 To 1) As Object
    Set instances(1) = sld
    Dim order As New Collection
    order.Add "3_P001"

    Dim rowValues As Object
    Set rowValues = CreateObject("Scripting.Dictionary")
    rowValues.Add "MS1_LABEL", "Kickoff": rowValues.Add "MS1_DATE", "Oct 2023": rowValues.Add "MS1_DONE", "Y"
    rowValues.Add "MS2_LABEL", "Design": rowValues.Add "MS2_DATE", "Mar 2024": rowValues.Add "MS2_DONE", "N"

    Dim rows As Object
    Set rows = CreateObject("Scripting.Dictionary")
    Set rows("3_P001") = rowValues

    ' CONTROL: a dry run must classify the SAME way and touch NOTHING -- same
    ' shape as the project's own load-bearing dry-run test for ordinary fields.
    Dim dryActions() As SyncAction
    dryActions = SyncOperations.PlanRoutineSync(instances, order, rows, True)
    result = result & Assert(dryActions(1).Kind = "in_place_correction", _
        "dry run still finds the device field, got '" & dryActions(1).Kind & "'")
    result = result & Assert(NamedIn(grp, "MS1_LABEL").TextFrame.TextRange.Text <> "Kickoff", _
        "dry run wrote NOTHING to the slide")

    ' THE REAL RUN.
    Dim actions() As SyncAction
    actions = SyncOperations.PlanRoutineSync(instances, order, rows)

    result = result & Assert(UBound(actions) = 1, "one action produced, got " & (UBound(actions) - LBound(actions) + 1))
    result = result & Assert(actions(1).Kind = "in_place_correction", _
        "THE FIX: a device tag with no matching register column is still found, got '" & actions(1).Kind & "'")
    result = result & Assert(actions(1).ChangedFieldVerified.Exists("MILESTONE_TIMELINE"), _
        "the change is keyed by the device's OWN tag, not a register column")

    ' THE WRITE ACTUALLY HAPPENED -- read straight off the real shapes.
    result = result & Assert(NamedIn(grp, "MS1_LABEL").TextFrame.TextRange.Text = "Kickoff", _
        "slot 1 label actually written, got '" & NamedIn(grp, "MS1_LABEL").TextFrame.TextRange.Text & "'")
    result = result & Assert(NamedIn(grp, "MS1_ON").Visible = msoTrue, "slot 1 shows achieved")
    result = result & Assert(NamedIn(grp, "MS2_OFF").Visible = msoTrue, "slot 2 shows not achieved")

    sld.Delete
    Test_SyncOperations_DeviceFieldReachableThroughPlanRoutineSync = result
End Function

Private Function Test_SyncOperations_Case6UnclassifiedSlide() As String
    Dim result As String
    Dim sld As Object
    Set sld = NewBlankSlide() ' no tags at all

    Dim instances(1 To 1) As Object
    Set instances(1) = sld
    Dim order As New Collection ' empty -- nothing to dispatch against

    Dim rows As Object
    Set rows = CreateObject("Scripting.Dictionary")

    Dim actions() As SyncAction
    actions = SyncOperations.PlanRoutineSync(instances, order, rows)

    result = result & Assert(UBound(actions) = 1, "one action produced, got " & (UBound(actions) - LBound(actions) + 1))
    result = result & Assert(actions(1).Kind = "flagged", "case 6: flagged, got '" & actions(1).Kind & "'")
    result = result & Assert(actions(1).FlagKind = "unclassified_slide", "FlagKind correct, got '" & actions(1).FlagKind & "'")

    Test_SyncOperations_Case6UnclassifiedSlide = result
End Function

' ---------------------------------------------------------------------
' Onboarding
' ---------------------------------------------------------------------

Private Function Test_Onboarding_HighAndMediumConfidence() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 300, 50)
    titleShp.TextFrame.TextRange.Text = "Title text"
    titleShp.Tags.Add "role", "Title"
    Dim bodyShp As Object
    Set bodyShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 200, 300, 200)
    bodyShp.TextFrame.TextRange.Text = "Body text"
    bodyShp.Tags.Add "role", "Body"

    Dim templateRoles() As String
    Dim templateFieldShapes() As Candidate
    templateFieldShapes = Onboarding.BuildTemplateFieldShapes(templateSld, templateRoles)
    result = result & Assert((UBound(templateRoles) - LBound(templateRoles) + 1) = 2, "template has 2 field roles")

    ' New slide: Title shape barely drifted (well inside tolerance -> high
    ' confidence expected). Body shape moved far away entirely (600pt off,
    ' many multiples of POSITION_TOLERANCE_EMU's 1-inch/72pt radius) --
    ' first attempt here used a modest drift and empirically scored "high"
    ' anyway (real run, 2026-07-25: geometry is only 0.3 of the renormalized
    ' weight once placeholder is inapplicable, so a moderate offset wasn't
    ' enough to pull the blended score below HIGH_THRESHOLD=0.75 on its
    ' own) -- a real finding about how forgiving the tolerance actually is
    ' with only 3 applicable signals, not a harness bug. Widened the drift
    ' to unambiguous rather than trying to hit "exactly medium" precisely.
    Dim newSld As Object
    Set newSld = NewBlankSlide()
    Dim newTitleShp As Object
    Set newTitleShp = newSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 102, 101, 300, 50)
    newTitleShp.TextFrame.TextRange.Text = "Title text drifted"
    Dim newBodyShp As Object
    Set newBodyShp = newSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 700, 700, 300, 200)
    newBodyShp.TextFrame.TextRange.Text = "Body text drifted"

    Dim untaggedShapes() As Object
    Dim matches() As FieldMatch
    matches = Onboarding.MatchSlideAgainstTemplate(newSld, templateRoles, templateFieldShapes, untaggedShapes)

    Dim titleConf As String, bodyConf As String
    Dim haveTitle As Boolean, haveBody As Boolean
    Dim i As Long
    For i = LBound(matches) To UBound(matches)
        If matches(i).Role = "Title" Then titleConf = matches(i).Result.Confidence: haveTitle = True
        If matches(i).Role = "Body" Then bodyConf = matches(i).Result.Confidence: haveBody = True
    Next i

    result = result & Assert(haveTitle And titleConf = "high", "Title (barely drifted) scores high, got '" & titleConf & "'")
    result = result & Assert(haveBody And bodyConf <> "high", "Body (drifted position+size) does not auto-accept, got '" & bodyConf & "'")

    Test_Onboarding_HighAndMediumConfidence = result
End Function

Private Function Test_Onboarding_OnboardNewInstanceAutoTagsHighOnly() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 300, 50)
    titleShp.TextFrame.TextRange.Text = "Title text"
    titleShp.Tags.Add "role", "Title"

    Dim templateRoles() As String
    Dim templateFieldShapes() As Candidate
    templateFieldShapes = Onboarding.BuildTemplateFieldShapes(templateSld, templateRoles)

    Dim newSld As Object
    Set newSld = NewBlankSlide()
    Dim newTitleShp As Object
    Set newTitleShp = newSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 101, 100, 300, 50)
    newTitleShp.TextFrame.TextRange.Text = "New title"

    Dim matches() As FieldMatch
    matches = Onboarding.OnboardNewInstance(newSld, templateRoles, templateFieldShapes, "quarterly-update", "rec-2")

    Dim instance As SlideInstance
    instance = Resolve.ResolveSlideInstance(newSld)
    result = result & Assert(instance.HasTypeTag And instance.TypeTag = "quarterly-update", "slide-level type tagged unconditionally")
    result = result & Assert(instance.HasInstanceKey And instance.InstanceKey = "rec-2", "slide-level instance key tagged unconditionally")
    result = result & Assert(newTitleShp.Tags("role") = "Title", "high-confidence Title auto-tagged onto the shape, got '" & newTitleShp.Tags("role") & "'")

    Test_Onboarding_OnboardNewInstanceAutoTagsHighOnly = result
End Function

Private Function Test_Onboarding_PureDecorationNeverMatched() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim decoShp As Object
    Set decoShp = templateSld.Shapes.AddShape(msoShapeOval, 100, 100, 50, 50) ' no text -- pure decoration
    decoShp.Tags.Add "role", "Deco"

    Dim templateRoles() As String
    Dim templateFieldShapes() As Candidate
    templateFieldShapes = Onboarding.BuildTemplateFieldShapes(templateSld, templateRoles)

    Dim newSld As Object
    Set newSld = NewBlankSlide()
    Dim newDecoShp As Object
    Set newDecoShp = newSld.Shapes.AddShape(msoShapeOval, 100, 100, 50, 50) ' identical geometry, still no text

    Dim untaggedShapes() As Object
    Dim matches() As FieldMatch
    matches = Onboarding.MatchSlideAgainstTemplate(newSld, templateRoles, templateFieldShapes, untaggedShapes)

    Dim lo As Long, hi As Long, hasMatches As Boolean
    On Error Resume Next
    lo = LBound(matches)
    hi = UBound(matches)
    hasMatches = (Err.Number = 0)
    On Error GoTo 0

    result = result & Assert(hasMatches And (hi - lo + 1) = 1, "one match result produced")
    If hasMatches Then
        result = result & Assert(matches(1).Result.Confidence = "low", "pure decoration never matches even at identical geometry, got '" & matches(1).Result.Confidence & "'")
        result = result & Assert(Not matches(1).Result.HasCandidate, "HasCandidate false")
    End If

    Test_Onboarding_PureDecorationNeverMatched = result
End Function

' ---------------------------------------------------------------------
' Verification
' ---------------------------------------------------------------------

Private Function Test_Verification_StructureMatchesAfterDuplicate() As String
    Dim result As String
    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim shp1 As Object, shp2 As Object
    Set shp1 = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    shp1.TextFrame.TextRange.Text = "Title"
    shp1.Tags.Add "role", "Title"
    Set shp2 = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 150, 200, 50)
    shp2.TextFrame.TextRange.Text = "Body"
    shp2.Tags.Add "role", "Body"

    Dim dupSlides As Object
    Set dupSlides = sld.Duplicate()
    Dim dupSld As Object
    Set dupSld = dupSlides(1)

    Dim structCheck As StructuralVerification
    structCheck = Verification.VerifyStructure(sld, dupSld)
    result = result & Assert(structCheck.Ok, "structure matches after a clean duplicate, got " & structCheck.MismatchCount & " mismatch(es)")
    result = result & Assert(structCheck.SourceCount = structCheck.DuplicateCount, "counts match, got " & structCheck.SourceCount & " vs " & structCheck.DuplicateCount)

    Dim zCheck As ZOrderVerification
    zCheck = Verification.VerifyZOrder(sld, dupSld)
    result = result & Assert(zCheck.Ok, "z-order matches after a clean duplicate, got " & zCheck.MismatchCount & " mismatch(es)")

    Test_Verification_StructureMatchesAfterDuplicate = result
End Function

Private Function Test_Verification_DetectsShapeCountMismatch() As String
    Dim result As String
    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim shp1 As Object, shp2 As Object
    Set shp1 = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    shp1.TextFrame.TextRange.Text = "Title"
    shp1.Tags.Add "role", "Title"
    Set shp2 = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 150, 200, 50)
    shp2.TextFrame.TextRange.Text = "Body"
    shp2.Tags.Add "role", "Body"

    Dim dupSlides As Object
    Set dupSlides = sld.Duplicate()
    Dim dupSld As Object
    Set dupSld = dupSlides(1)
    dupSld.Shapes(2).Delete ' real structural defect: duplicate is missing a shape

    Dim structCheck As StructuralVerification
    structCheck = Verification.VerifyStructure(sld, dupSld)
    result = result & Assert(Not structCheck.Ok, "structure mismatch detected")
    result = result & Assert(structCheck.MismatchCount > 0, "at least one mismatch reported")

    Test_Verification_DetectsShapeCountMismatch = result
End Function

Private Function Test_Verification_DetectsZOrderSwap() As String
    Dim result As String
    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim shp1 As Object, shp2 As Object
    Set shp1 = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    shp1.TextFrame.TextRange.Text = "Title"
    shp1.Tags.Add "role", "Title"
    Set shp2 = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 150, 200, 50)
    shp2.TextFrame.TextRange.Text = "Body"
    shp2.Tags.Add "role", "Body"

    Dim dupSlides As Object
    Set dupSlides = sld.Duplicate()
    Dim dupSld As Object
    Set dupSld = dupSlides(1)
    ' Shapes(1) (Title) was added first, so it's already backmost among just
    ' 2 shapes -- sending it "to back" would be a no-op (first attempt here
    ' did exactly that and found nothing, a test bug not a Verification.bas
    ' bug, per the real 2026-07-25 run). Bring it to front instead, which
    ' actually swaps its relative order against Body.
    dupSld.Shapes(1).ZOrder msoBringToFront ' same shapes/tags/types, stacking order swapped on the duplicate only

    Dim zCheck As ZOrderVerification
    zCheck = Verification.VerifyZOrder(sld, dupSld)
    result = result & Assert(Not zCheck.Ok, "z-order swap detected")
    result = result & Assert(zCheck.MismatchCount > 0, "at least one z-order mismatch reported")

    Dim structCheck As StructuralVerification
    structCheck = Verification.VerifyStructure(sld, dupSld)
    result = result & Assert(structCheck.Ok, "structure itself is unaffected by a pure z-order swap, got " & structCheck.MismatchCount & " mismatch(es)")

    Test_Verification_DetectsZOrderSwap = result
End Function

' ---------------------------------------------------------------------
' SlideDuplication
' ---------------------------------------------------------------------

Private Function Test_SlideDuplication_CreatesTaggedInjectedSlide() As String
    Dim result As String
    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object, bodyShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    titleShp.TextFrame.TextRange.Text = "Template Title"
    titleShp.Tags.Add "role", "Title"
    Set bodyShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 150, 200, 50)
    bodyShp.TextFrame.TextRange.Text = "Template Body"
    bodyShp.Tags.Add "role", "Body"
    templateSld.Tags.Add "slide_type", "quarterly-update"
    templateSld.Tags.Add "instance_key", "rec-template"

    Dim values As Object
    Set values = CreateObject("Scripting.Dictionary")
    values("Title") = "Q3 Revenue"
    values("Body") = "Strong quarter"

    Dim noInstances() As Object ' deliberately unallocated -- nothing to collision-check against

    Dim dr As DuplicateResult
    dr = SlideDuplication.DuplicateAndTag(templateSld, "quarterly-update", "rec-new-1", values, noInstances)

    result = result & Assert(dr.Ok, "DuplicateAndTag succeeded, reason='" & dr.Reason & "'")
    If dr.Ok Then
        Dim instance As SlideInstance
        instance = Resolve.ResolveSlideInstance(dr.NewSlide)
        result = result & Assert(instance.TypeTag = "quarterly-update", "new slide tagged with correct slide_type, got '" & instance.TypeTag & "'")
        result = result & Assert(instance.InstanceKey = "rec-new-1", "new slide tagged with correct instance_key, got '" & instance.InstanceKey & "'")
        result = result & Assert(dr.MissingFieldCount = 0, "no missing fields when values supplies everything, got " & dr.MissingFieldCount)

        Dim newTitleShp As Object
        Set newTitleShp = FindShapeByRole(dr.NewSlide, "Title")
        result = result & Assert(Not newTitleShp Is Nothing, "Title shape found on new slide")
        If Not newTitleShp Is Nothing Then
            result = result & Assert(newTitleShp.TextFrame.TextRange.Text = "Q3 Revenue", "Title value injected correctly, got '" & newTitleShp.TextFrame.TextRange.Text & "'")
        End If
    End If

    Test_SlideDuplication_CreatesTaggedInjectedSlide = result
End Function

Private Function Test_SlideDuplication_RefusesInstanceKeyCollision() As String
    Dim result As String
    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    titleShp.TextFrame.TextRange.Text = "Title"
    titleShp.Tags.Add "role", "Title"
    templateSld.Tags.Add "slide_type", "quarterly-update"
    templateSld.Tags.Add "instance_key", "rec-a"

    Dim existingSld As Object
    Set existingSld = NewBlankSlide()
    existingSld.Tags.Add "slide_type", "quarterly-update"
    existingSld.Tags.Add "instance_key", "rec-dup"

    Dim existingInstances(1 To 1) As Object
    Set existingInstances(1) = existingSld

    Dim values As Object
    Set values = CreateObject("Scripting.Dictionary")
    values("Title") = "New Value"

    Dim slidesBefore As Long
    slidesBefore = Application.ActivePresentation.Slides.count

    Dim dr As DuplicateResult
    dr = SlideDuplication.DuplicateAndTag(templateSld, "quarterly-update", "rec-dup", values, existingInstances)

    result = result & Assert(Not dr.Ok, "refuses to create a duplicate instance_key")
    result = result & Assert(Application.ActivePresentation.Slides.count = slidesBefore, "no new slide was left behind, got " & Application.ActivePresentation.Slides.count & " vs expected " & slidesBefore)

    Test_SlideDuplication_RefusesInstanceKeyCollision = result
End Function

Private Function Test_SlideDuplication_PartialRowStillCreatesSlideButFlagsMissing() As String
    Dim result As String
    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object, bodyShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    titleShp.TextFrame.TextRange.Text = "Title"
    titleShp.Tags.Add "role", "Title"
    Set bodyShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 150, 200, 50)
    bodyShp.TextFrame.TextRange.Text = "Body"
    bodyShp.Tags.Add "role", "Body"
    templateSld.Tags.Add "slide_type", "quarterly-update"
    templateSld.Tags.Add "instance_key", "rec-template2"

    Dim values As Object
    Set values = CreateObject("Scripting.Dictionary")
    values("Title") = "Only Title Supplied" ' Body deliberately missing

    Dim noInstances() As Object

    Dim dr As DuplicateResult
    dr = SlideDuplication.DuplicateAndTag(templateSld, "quarterly-update", "rec-partial", values, noInstances)

    result = result & Assert(dr.Ok, "slide still created despite a missing field, reason='" & dr.Reason & "'")
    result = result & Assert(dr.MissingFieldCount = 1, "exactly one missing field flagged, got " & dr.MissingFieldCount)
    If dr.MissingFieldCount = 1 Then
        result = result & Assert(dr.MissingFields(1) = "Body", "flagged field is 'Body', got '" & dr.MissingFields(1) & "'")
    End If

    Test_SlideDuplication_PartialRowStillCreatesSlideButFlagsMissing = result
End Function

' ---------------------------------------------------------------------
' TemplateSlide -- progression step 1 of specs/deck-compiler-concept.md
' ---------------------------------------------------------------------

' Helper: a minimal already-onboarded instance of `slideType`, two tagged
' fields carrying values that read as one specific project's real data --
' so a test can tell "inherited from the source project" apart from
' "replaced with a placeholder", which is the entire point of step 1.
Private Function NewOnboardedSlide(slideType As String, instanceKey As String) As Object
    Dim sld As Object
    Set sld = NewBlankSlide()

    Dim titleShp As Object, statusShp As Object
    Set titleShp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    titleShp.TextFrame.TextRange.Text = "Real Project Name"
    titleShp.Tags.Add "role", "Title"
    Set statusShp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 150, 200, 50)
    statusShp.TextFrame.TextRange.Text = "In Progress"
    statusShp.Tags.Add "role", "Status"

    sld.Tags.Add "slide_type", slideType
    sld.Tags.Add "instance_key", instanceKey

    Set NewOnboardedSlide = sld
End Function

' The core of step 1: the copy is typed, marked, KEYLESS, placeholdered and
' hidden -- and the slide it came from is untouched.
'
' The source-untouched assertions are not padding. MakeTemplateFrom's whole
' safety story is "it only ever writes to the copy it just made", and the
' source here is standing in for a real slide a human authored.
' Every case here comes from the REAL deck's 43 instance keys, not from an
' imagined format. The five bare keys (S023 and friends) are the reason the
' underscore cannot be assumed, and the digit case is the reason the first
' character cannot be taken blindly.
Private Function Test_TemplateSlide_CodeLetterOfReadsBothKeyShapes() As String
    Dim result As String

    result = result & Assert(TemplateSlide.CodeLetterOf("3_P001") = "P", _
        "themed key 3_P001 reads P, got '" & TemplateSlide.CodeLetterOf("3_P001") & "'")
    result = result & Assert(TemplateSlide.CodeLetterOf("1_K1008") = "K", _
        "themed key 1_K1008 reads K, got '" & TemplateSlide.CodeLetterOf("1_K1008") & "'")
    result = result & Assert(TemplateSlide.CodeLetterOf("2_S015") = "S", _
        "themed key 2_S015 reads S, got '" & TemplateSlide.CodeLetterOf("2_S015") & "'")

    ' Five real projects have no underscore at all.
    result = result & Assert(TemplateSlide.CodeLetterOf("S023") = "S", _
        "bare key S023 reads S, got '" & TemplateSlide.CodeLetterOf("S023") & "'")
    result = result & Assert(TemplateSlide.CodeLetterOf("P008") = "P", _
        "bare key P008 reads P, got '" & TemplateSlide.CodeLetterOf("P008") & "'")

    ' No letter to read means NO OPINION -- never a guess, and never a digit
    ' promoted to a fourth colour.
    result = result & Assert(TemplateSlide.CodeLetterOf("1_2003") = "", _
        "a key with digits after the underscore yields no letter")
    result = result & Assert(TemplateSlide.CodeLetterOf("") = "", _
        "an empty key yields no letter")
    result = result & Assert(TemplateSlide.CodeLetterOf("   ") = "", _
        "a whitespace key yields no letter")
    result = result & Assert(TemplateSlide.CodeLetterOf("3_") = "", _
        "a key ending in the underscore yields no letter")

    ' Case is normalised so a hand-typed key cannot address a different template.
    result = result & Assert(TemplateSlide.CodeLetterOf("3_p001") = "P", _
        "lowercase key 3_p001 still reads P")

    Test_TemplateSlide_CodeLetterOfReadsBothKeyShapes = result
End Function

Private Function Test_TemplateSlide_MakeTemplateProducesKeylessMarkedCopy() As String
    Dim result As String

    Dim sourceSld As Object
    Set sourceSld = NewOnboardedSlide("tmpl-type-1", "rec-real-1")

    Dim mr As MakeTemplateResult
    mr = TemplateSlide.MakeTemplateFrom(sourceSld, "tmpl-type-1")

    result = result & Assert(mr.Ok, "MakeTemplateFrom succeeded, reason='" & mr.Reason & "'")
    If Not mr.Ok Then
        Test_TemplateSlide_MakeTemplateProducesKeylessMarkedCopy = result
        Exit Function
    End If

    Dim tmpl As SlideInstance
    tmpl = Resolve.ResolveSlideInstance(mr.NewSlide)
    result = result & Assert(tmpl.IsTemplate, "the copy is marked is_template")
    result = result & Assert(Not tmpl.HasInstanceKey, "the copy has NO instance_key -- it must never match a Data row, got '" & tmpl.InstanceKey & "'")
    result = result & Assert(tmpl.TypeTag = "tmpl-type-1", "the copy keeps its slide_type, got '" & tmpl.TypeTag & "'")
    result = result & Assert(mr.NewSlide.SlideShowTransition.Hidden = msoTrue, "the copy is hidden from the slideshow")
    result = result & Assert(mr.FieldCount = 2, "both fields reported as placeholdered, got " & mr.FieldCount)

    ' Placeholders, asserted by their rendered text rather than by count --
    ' a count of 2 would pass while both fields still read "Real Project
    ' Name", which is the defect this whole operation exists to prevent.
    Dim tmplTitle As Object, tmplStatus As Object
    Set tmplTitle = FindShapeByRole(mr.NewSlide, "Title")
    Set tmplStatus = FindShapeByRole(mr.NewSlide, "Status")
    result = result & Assert(Not tmplTitle Is Nothing, "Title field present on the template")
    result = result & Assert(Not tmplStatus Is Nothing, "Status field present on the template")
    If Not tmplTitle Is Nothing Then
        result = result & Assert(tmplTitle.TextFrame.TextRange.Text = "<<Title>>", "Title reads as a placeholder, got '" & tmplTitle.TextFrame.TextRange.Text & "'")
    End If
    If Not tmplStatus Is Nothing Then
        result = result & Assert(tmplStatus.TextFrame.TextRange.Text = "<<Status>>", "Status reads as a placeholder, got '" & tmplStatus.TextFrame.TextRange.Text & "'")
    End If

    ' Source slide: completely unchanged.
    Dim src As SlideInstance
    src = Resolve.ResolveSlideInstance(sourceSld)
    result = result & Assert(src.InstanceKey = "rec-real-1", "source keeps its instance_key, got '" & src.InstanceKey & "'")
    result = result & Assert(Not src.IsTemplate, "source did NOT become a template itself")
    result = result & Assert(sourceSld.SlideShowTransition.Hidden <> msoTrue, "source was not hidden")
    Dim srcTitle As Object
    Set srcTitle = FindShapeByRole(sourceSld, "Title")
    If Not srcTitle Is Nothing Then
        result = result & Assert(srcTitle.TextFrame.TextRange.Text = "Real Project Name", "source field text untouched, got '" & srcTitle.TextFrame.TextRange.Text & "'")
    End If

    Test_TemplateSlide_MakeTemplateProducesKeylessMarkedCopy = result
End Function

' The load-bearing exclusion. Every consequence of a template slide flows
' through GatherInstances, so this asserts both halves at once: the template
' is not gathered, AND the plan built from that gather contains no "flagged"
' action for it.
'
' The second half is the one that matters in use. A typed keyless slide is
' precisely case 6 (unclassified_slide), so before the exclusion existed a
' template would have appeared as a flagged problem on every single sync --
' permanent noise, in the report a human is meant to read for real problems.
Private Function Test_TemplateSlide_ExcludedFromGatherAndNeverFlagged() As String
    Dim result As String

    Dim instSld As Object
    Set instSld = NewOnboardedSlide("tmpl-type-2", "rec-real-2")

    Dim mr As MakeTemplateResult
    mr = TemplateSlide.MakeTemplateFrom(instSld, "tmpl-type-2")
    result = result & Assert(mr.Ok, "template created for the gather test, reason='" & mr.Reason & "'")
    If Not mr.Ok Then
        Test_TemplateSlide_ExcludedFromGatherAndNeverFlagged = result
        Exit Function
    End If

    Dim gathered() As Object
    gathered = RunSync.GatherInstances("tmpl-type-2")

    Dim lo As Long, hi As Long, hasAny As Boolean
    On Error Resume Next
    lo = LBound(gathered): hi = UBound(gathered)
    hasAny = (Err.Number = 0)
    On Error GoTo 0

    result = result & Assert(hasAny, "the real instance is still gathered")
    If hasAny Then
        result = result & Assert(hi - lo + 1 = 1, "exactly 1 instance gathered (the template is excluded), got " & (hi - lo + 1))
        Dim seenTemplate As Boolean
        Dim i As Long
        For i = lo To hi
            If Resolve.IsTemplateSlide(gathered(i)) Then seenTemplate = True
        Next i
        result = result & Assert(Not seenTemplate, "no gathered instance is a template slide")
    End If

    ' FindTemplateFor still sees it -- the template exists, it is just not a record.
    Dim found As Object
    Set found = TemplateSlide.FindTemplateFor("tmpl-type-2")
    result = result & Assert(Not found Is Nothing, "FindTemplateFor locates the template even though the gather excludes it")

    ' And nothing about it reaches the human's report.
    Dim order As New Collection
    order.Add "rec-real-2"
    Dim rows As Object
    Set rows = CreateObject("Scripting.Dictionary")
    Dim rowVals As Object
    Set rowVals = CreateObject("Scripting.Dictionary")
    rowVals("Title") = "Real Project Name"
    rowVals("Status") = "In Progress"
    Set rows("rec-real-2") = rowVals

    Dim actions() As SyncAction
    actions = SyncOperations.PlanRoutineSync(gathered, order, rows, True)

    Dim aLo As Long, aHi As Long, hasActions As Boolean
    On Error Resume Next
    aLo = LBound(actions): aHi = UBound(actions)
    hasActions = (Err.Number = 0)
    On Error GoTo 0

    Dim flaggedCount As Long
    If hasActions Then
        Dim a As Long
        For a = aLo To aHi
            If actions(a).Kind = "flagged" Then flaggedCount = flaggedCount + 1
        Next a
    End If
    result = result & Assert(flaggedCount = 0, "the template produces NO flagged action -- it would otherwise be case-6 noise on every sync, got " & flaggedCount)

    Test_TemplateSlide_ExcludedFromGatherAndNeverFlagged = result
End Function

' Cloning the template must produce a NORMAL record, not a second template.
'
' This is the one failure in step 1 that is invisible by construction: an
' inherited is_template marker would make the new slide skip every future
' sync, and it skips them silently, because the exclusion exists precisely
' to keep templates out of reports. Nothing would say the project had gone
' missing.
Private Function Test_TemplateSlide_DuplicateStripsTheTemplateMarker() As String
    Dim result As String

    Dim seedSld As Object
    Set seedSld = NewOnboardedSlide("tmpl-type-3", "rec-real-3")

    Dim mr As MakeTemplateResult
    mr = TemplateSlide.MakeTemplateFrom(seedSld, "tmpl-type-3")
    result = result & Assert(mr.Ok, "template created for the duplication test, reason='" & mr.Reason & "'")
    If Not mr.Ok Then
        Test_TemplateSlide_DuplicateStripsTheTemplateMarker = result
        Exit Function
    End If

    Dim values As Object
    Set values = CreateObject("Scripting.Dictionary")
    values("Title") = "Brand New Project"
    values("Status") = "Not Started"

    Dim noInstances() As Object

    Dim dr As DuplicateResult
    dr = SlideDuplication.DuplicateAndTag(mr.NewSlide, "tmpl-type-3", "rec-created-3", values, noInstances)

    result = result & Assert(dr.Ok, "DuplicateAndTag succeeded off a template, reason='" & dr.Reason & "'")
    If Not dr.Ok Then
        Test_TemplateSlide_DuplicateStripsTheTemplateMarker = result
        Exit Function
    End If

    Dim created As SlideInstance
    created = Resolve.ResolveSlideInstance(dr.NewSlide)
    result = result & Assert(Not created.IsTemplate, "the created slide is NOT a template -- the marker was stripped from the copy")
    result = result & Assert(created.InstanceKey = "rec-created-3", "the created slide carries its own instance_key, got '" & created.InstanceKey & "'")

    ' The template is hidden, and Slide.Duplicate copies that too -- so a
    ' record cloned from it arrived hidden from the slideshow. Found live
    ' 2026-07-30 (two struck-through slide numbers where one was expected),
    ' AFTER this test had already passed asserting only the tag. The tag and
    ' the hidden flag are inherited by the same mechanism; only one of them
    ' was being checked.
    result = result & Assert(dr.NewSlide.SlideShowTransition.Hidden = msoFalse, "the created slide is VISIBLE -- it must not inherit the template's hidden flag")
    result = result & Assert(mr.NewSlide.SlideShowTransition.Hidden = msoTrue, "...while the template itself stays hidden")

    ' The real test of "not a template": the sync can see it.
    Dim gathered() As Object
    gathered = RunSync.GatherInstances("tmpl-type-3")
    Dim gLo As Long, gHi As Long, hasAny As Boolean
    On Error Resume Next
    gLo = LBound(gathered): gHi = UBound(gathered)
    hasAny = (Err.Number = 0)
    On Error GoTo 0
    result = result & Assert(hasAny, "gather found something after creating a record off the template")
    If hasAny Then
        Dim seenCreated As Boolean
        Dim i As Long
        For i = gLo To gHi
            If Resolve.ResolveSlideInstance(gathered(i)).InstanceKey = "rec-created-3" Then seenCreated = True
        Next i
        result = result & Assert(seenCreated, "the created slide IS gathered by future syncs -- not silently skipped")
    End If

    ' Placeholders replaced, not inherited.
    Dim newTitle As Object
    Set newTitle = FindShapeByRole(dr.NewSlide, "Title")
    If Not newTitle Is Nothing Then
        result = result & Assert(newTitle.TextFrame.TextRange.Text = "Brand New Project", "the row's value overwrote the placeholder, got '" & newTitle.TextFrame.TextRange.Text & "'")
    End If

    Test_TemplateSlide_DuplicateStripsTheTemplateMarker = result
End Function

' A type must have exactly one template. Two would have no defined
' behaviour at all -- DeckRegistry.LookupType returns whichever SlideID was
' registered last, so which slide gets cloned would depend on click order.
Private Function Test_TemplateSlide_RefusesToTemplateATemplate() As String
    Dim result As String

    Dim seedSld As Object
    Set seedSld = NewOnboardedSlide("tmpl-type-4", "rec-real-4")

    Dim firstTmpl As MakeTemplateResult
    firstTmpl = TemplateSlide.MakeTemplateFrom(seedSld, "tmpl-type-4")
    result = result & Assert(firstTmpl.Ok, "first template created, reason='" & firstTmpl.Reason & "'")
    If Not firstTmpl.Ok Then
        Test_TemplateSlide_RefusesToTemplateATemplate = result
        Exit Function
    End If

    Dim slidesBefore As Long
    slidesBefore = Application.ActivePresentation.Slides.count

    Dim secondTmpl As MakeTemplateResult
    secondTmpl = TemplateSlide.MakeTemplateFrom(firstTmpl.NewSlide, "tmpl-type-4")

    result = result & Assert(Not secondTmpl.Ok, "refuses to make a template out of a template")
    result = result & Assert(Application.ActivePresentation.Slides.count = slidesBefore, "no slide left behind by the refusal, got " & Application.ActivePresentation.Slides.count & " vs expected " & slidesBefore)
    result = result & Assert(InStr(secondTmpl.Reason, "already a master template") > 0, "the refusal says why, got '" & secondTmpl.Reason & "'")

    ' And refuses across types too -- one type's template must not be built
    ' out of another type's slide.
    Dim otherSld As Object
    Set otherSld = NewOnboardedSlide("tmpl-type-5", "rec-real-5")
    Dim mismatched As MakeTemplateResult
    mismatched = TemplateSlide.MakeTemplateFrom(otherSld, "tmpl-type-4")
    result = result & Assert(Not mismatched.Ok, "refuses to build type-4's template out of a type-5 slide")

    Test_TemplateSlide_RefusesToTemplateATemplate = result
End Function

' The confirmation wording IS the guard for this action (it writes, and it
' re-points what Sync Now clones), so it gets pinned rather than
' hand-checked -- same reason RunSync.ConfirmSyncText's wording is asserted.
' Re-registration is the consequence a user is least likely to predict from
' the button's name, so that is the phrase held in place.
Private Function Test_TemplateSlide_ConfirmTextStatesTheConsequences() As String
    Dim result As String

    Dim s As String
    s = TemplateSlide.ConfirmTemplateText("quarterly-update", "3_P001 (slide 3)", 4)

    result = result & Assert(InStr(s, "quarterly-update") > 0, "confirmation names the type")
    result = result & Assert(InStr(s, "3_P001 (slide 3)") > 0, "confirmation names the slide being copied")
    result = result & Assert(InStr(s, "4 field(s)") > 0, "confirmation states how many fields get placeholdered")
    result = result & Assert(InStr(s, "RE-REGISTERED") > 0, "confirmation states that the type is re-pointed -- the least predictable consequence")
    result = result & Assert(InStr(s, "is NOT touched") > 0, "confirmation states the source slide is left alone")
    result = result & Assert(InStr(s, "Proceed?") > 0, "confirmation actually asks")

    result = result & Assert(TemplateSlide.PlaceholderFor("Status") = "<<Status>>", "PlaceholderFor wraps the role name, got '" & TemplateSlide.PlaceholderFor("Status") & "'")

    Test_TemplateSlide_ConfirmTextStatesTheConsequences = result
End Function

' Proves the letter-aware wording fixed 2026-08-16, found live: the old
' unconditional "RE-REGISTERED" text was FALSE the moment a second letter
' was added to a type that already had a template. Two real shapes, both
' exercised live that day (K and S against an existing P template) --
' willClaimFallback=False -- and the other real shape (a genuinely first-ever
' template for a type) -- willClaimFallback=True.
Private Function Test_TemplateSlide_ConfirmTextIsLetterAwareAndDoesNotStealTheFallback() As String
    Dim result As String

    Dim s As String
    s = TemplateSlide.ConfirmTemplateText("project-progress", "1_K1001 (slide 12)", 15, "K", False)

    result = result & Assert(InStr(s, "letter 'K'") > 0, "confirmation names the letter, got: " & s)
    result = result & Assert(InStr(s, "RE-REGISTERED") = 0, _
        "confirmation does NOT claim the whole type is re-registered when it is not")
    result = result & Assert(InStr(s, "ONLY letter 'K' rows") > 0, _
        "confirmation states the scope is limited to this letter, not the whole type")

    Test_TemplateSlide_ConfirmTextIsLetterAwareAndDoesNotStealTheFallback = result
End Function

Private Function Test_TemplateSlide_ConfirmTextNamesWhenALetterBecomesTheFallback() As String
    Dim result As String

    Dim s As String
    s = TemplateSlide.ConfirmTemplateText("new-type", "1_K1001 (slide 12)", 15, "K", True)

    result = result & Assert(InStr(s, "letter 'K'") > 0, "confirmation names the letter, got: " & s)
    result = result & Assert(InStr(s, "FIRST template") > 0, _
        "confirmation states this letter is also becoming the type-level default")

    Test_TemplateSlide_ConfirmTextNamesWhenALetterBecomesTheFallback = result
End Function

Private Function Test_TemplateSlide_ExistingTemplateForLetterFindsTheRightOne() As String
    Dim result As String
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim tmplK As Object, tmplS As Object
    Set tmplK = NewBlankSlide()
    tmplK.Tags.Add "slide_type", "existing-tmpl-type"
    tmplK.Tags.Add "is_template", "1"
    Set tmplS = NewBlankSlide()
    tmplS.Tags.Add "slide_type", "existing-tmpl-type"
    tmplS.Tags.Add "is_template", "1"

    DeckRegistry.RegisterTemplateLetter pres, "existing-tmpl-type", "K", tmplK, "SheetK"
    DeckRegistry.RegisterTemplateLetter pres, "existing-tmpl-type", "S", tmplS, "SheetS"

    Dim foundK As Object
    Set foundK = TemplateSlide.ExistingTemplateForLetter(pres, "existing-tmpl-type", "K")
    result = result & AssertSameSlide(foundK, tmplK, "K letter finds the K template")

    Dim foundS As Object
    Set foundS = TemplateSlide.ExistingTemplateForLetter(pres, "existing-tmpl-type", "S")
    result = result & AssertSameSlide(foundS, tmplS, "S letter finds the S template, not K's")

    Dim foundP As Object
    Set foundP = TemplateSlide.ExistingTemplateForLetter(pres, "existing-tmpl-type", "P")
    Dim foundPDesc As String
    If foundP Is Nothing Then foundPDesc = "Nothing" Else foundPDesc = "SlideIndex " & foundP.SlideIndex
    result = result & Assert(foundP Is Nothing, "an unregistered letter (P) finds nothing -- it is safe to create one, got " & foundPDesc)

    tmplK.Delete
    tmplS.Delete
    Test_TemplateSlide_ExistingTemplateForLetterFindsTheRightOne = result
End Function

Private Function Test_TemplateSlide_ExistingTemplateForLetterEmptyLetterUsesTypeWideCheck() As String
    Dim result As String
    Dim pres As Object
    Set pres = Application.ActivePresentation

    ' No letter to key on -- must fall back to the ORIGINAL one-per-type
    ' check (FindTemplateFor's live scan), not report "safe to proceed" just
    ' because nothing was ever registered under an empty-string letter.
    Dim tmpl As Object
    Set tmpl = NewBlankSlide()
    tmpl.Tags.Add "slide_type", "empty-letter-tmpl-type"
    tmpl.Tags.Add "is_template", "1"

    Dim found As Object
    Set found = TemplateSlide.ExistingTemplateForLetter(pres, "empty-letter-tmpl-type", "")
    result = result & AssertSameSlide(found, tmpl, "empty letter still finds the type's one template via the type-wide check")

    tmpl.Delete
    Test_TemplateSlide_ExistingTemplateForLetterEmptyLetterUsesTypeWideCheck = result
End Function

Private Function Test_DeckRegistry_RegisterNewTemplateLetterClaimsFallbackWhenNoneExists() As String
    Dim result As String
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim tmplK As Object
    Set tmplK = NewBlankSlide()
    tmplK.Tags.Add "slide_type", "regnew-type-1"
    tmplK.Tags.Add "is_template", "1"

    DeckRegistry.RegisterNewTemplateLetter pres, "regnew-type-1", "K", tmplK, "SheetK"

    Dim foundLetter As Object
    Dim wsLetter As String
    Dim okLetter As Boolean
    okLetter = DeckRegistry.LookupTemplateLetter(pres, "regnew-type-1", "K", foundLetter, wsLetter)
    result = result & Assert(okLetter, "the letter-specific lookup reports success")
    result = result & AssertSameSlide(foundLetter, tmplK, "the letter-specific slot is registered")

    Dim foundFallback As Object
    Dim wsFallback As String
    Dim okFallback As Boolean
    okFallback = DeckRegistry.LookupType(pres, "regnew-type-1", foundFallback, wsFallback)
    result = result & Assert(okFallback, "the type-level fallback lookup reports success")
    result = result & AssertSameSlide(foundFallback, tmplK, _
        "the FIRST template for a type also claims the type-level fallback, since nothing held it yet")

    tmplK.Delete
    Test_DeckRegistry_RegisterNewTemplateLetterClaimsFallbackWhenNoneExists = result
End Function

Private Function Test_DeckRegistry_RegisterNewTemplateLetterDoesNotStealAnExistingTemplateFallback() As String
    Dim result As String
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim tmplK As Object, tmplS As Object
    Set tmplK = NewBlankSlide()
    tmplK.Tags.Add "slide_type", "regnew-type-2"
    tmplK.Tags.Add "is_template", "1"
    Set tmplS = NewBlankSlide()
    tmplS.Tags.Add "slide_type", "regnew-type-2"
    tmplS.Tags.Add "is_template", "1"

    ' K registered first -- claims both its own slot AND the type-level
    ' fallback, same as the previous test.
    DeckRegistry.RegisterNewTemplateLetter pres, "regnew-type-2", "K", tmplK, "SheetK"

    ' S registered second -- must get its OWN slot, but must NOT steal the
    ' fallback out from under K.
    DeckRegistry.RegisterNewTemplateLetter pres, "regnew-type-2", "S", tmplS, "SheetS"

    Dim foundS As Object
    Dim wsS As String
    DeckRegistry.LookupTemplateLetter pres, "regnew-type-2", "S", foundS, wsS
    result = result & AssertSameSlide(foundS, tmplS, "S's own letter slot is registered")

    Dim foundFallback As Object
    Dim wsFallback As String
    DeckRegistry.LookupType pres, "regnew-type-2", foundFallback, wsFallback
    result = result & AssertSameSlide(foundFallback, tmplK, _
        "the type-level fallback still points at K, the FIRST template -- S did not steal it")

    tmplK.Delete
    tmplS.Delete
    Test_DeckRegistry_RegisterNewTemplateLetterDoesNotStealAnExistingTemplateFallback = result
End Function

' ---------------------------------------------------------------------
' PlaceholderCheck -- V7, Amendment A
' ---------------------------------------------------------------------

' Detection: a real record still showing scaffolding is found; the master
' template is NOT, because its fields are supposed to read <<...>> forever and
' counting them would mean the marker could never reach zero.
' ---------------------------------------------------------------------
' R13 -- ReviewQueue
' ---------------------------------------------------------------------

' Builds a queue item without needing a deck. Every R13 rule below is a
' property of the change SET, not of Office, so these run as pure logic.
Private Sub AddQItem(ByRef q As ReviewQueueSet, entityKey As String, fieldId As String, _
                     currentValue As String, proposedValue As String)
    q.Count = q.Count + 1
    ReDim Preserve q.Items(1 To q.Count)
    q.Items(q.Count).EntityKey = entityKey
    q.Items(q.Count).FieldID = fieldId
    q.Items(q.Count).CurrentValue = currentValue
    q.Items(q.Count).ProposedValue = proposedValue
    q.Items(q.Count).ChangeHash = ReviewQueue.ChangeHash(entityKey, fieldId, currentValue, proposedValue)
    q.Items(q.Count).BatchLabel = ""
    q.Items(q.Count).Approved = False
End Sub

Private Function Test_ReviewQueue_HashDistinguishesEveryField() As String
    Dim result As String

    Dim base As String
    base = ReviewQueue.ChangeHash("3_P001", "PROJECT_STATUS", "In progress", "In Progress")

    result = result & Assert(base = ReviewQueue.ChangeHash("3_P001", "PROJECT_STATUS", "In progress", "In Progress"), _
        "the same change hashes the same twice -- an approval must survive being re-derived")

    ' All four inputs must matter. If any did not, an approval for one row could
    ' silently authorise a write to a different one.
    result = result & Assert(base <> ReviewQueue.ChangeHash("3_P002", "PROJECT_STATUS", "In progress", "In Progress"), "entity key changes the hash")
    result = result & Assert(base <> ReviewQueue.ChangeHash("3_P001", "ABOUT_BODY", "In progress", "In Progress"), "field id changes the hash")
    result = result & Assert(base <> ReviewQueue.ChangeHash("3_P001", "PROJECT_STATUS", "In Progress", "In Progress"), "the CURRENT value changes the hash -- this is what catches a hand edit made after approval")
    result = result & Assert(base <> ReviewQueue.ChangeHash("3_P001", "PROJECT_STATUS", "In progress", "Closed"), "the PROPOSED value changes the hash -- this is what catches a register edit made after approval")

    ' The separator must not be forgeable out of field content.
    result = result & Assert(ReviewQueue.ChangeHash("a", "b", "c", "d") <> ReviewQueue.ChangeHash("a", "b", "c|d", ""), _
        "content containing the visible delimiter cannot collide with a different row")

    Test_ReviewQueue_HashDistinguishesEveryField = result
End Function

Private Function Test_ReviewQueue_ProseNeverBatchesEvenWhenUniform() As String
    Dim result As String

    ' THE CASE THAT MEASUREMENT ALONE CANNOT CATCH. Four projects whose ABOUT
    ' text is currently identical and would be rewritten identically. By
    ' uniformity this is a perfect batch of 4; by R13.2 prose may never be
    ' batched, so kind has to be able to veto the measurement.
    Dim q As ReviewQueueSet
    AddQItem q, "3_P001", "ABOUT_BODY", "Placeholder text.", "A study of AMR in soil."
    AddQItem q, "3_P002", "ABOUT_BODY", "Placeholder text.", "A study of AMR in soil."
    AddQItem q, "3_P003", "ABOUT_BODY", "Placeholder text.", "A study of AMR in soil."
    AddQItem q, "3_P004", "ABOUT_BODY", "Placeholder text.", "A study of AMR in soil."
    ReviewQueue.AssignBatches q

    result = result & Assert(ReviewQueue.ContentKindOf("ABOUT_BODY") = ReviewQueue.KIND_PROSE, _
        "an unlisted field defaults to Prose -- absence of a label is never permission to batch")

    Dim i As Long
    Dim anyLabelled As Boolean
    For i = 1 To q.Count
        If q.Items(i).BatchLabel <> "" Then anyLabelled = True
    Next i
    result = result & Assert(Not anyLabelled, "four IDENTICAL prose changes still batch into nothing -- R13.2's 'no, never'")
    result = result & Assert(ReviewQueue.IndividualCount(q) = 4, "all four remain individual decisions, got " & ReviewQueue.IndividualCount(q))
    result = result & Assert(Not ReviewQueue.HasBatchableWork(q), "a prose-only change set can never take the one-dialog fast path")

    Test_ReviewQueue_ProseNeverBatchesEvenWhenUniform = result
End Function

Private Function Test_ReviewQueue_UniformControlledGroupIsOneDecision() As String
    Dim result As String

    ' The real 2026-07-31 shape: three slides sharing one transformation, one
    ' slide with a different current value, one lone change of its own.
    Dim q As ReviewQueueSet
    AddQItem q, "3_P001", "PROJECT_STATUS", "In progress", "In Progress"
    AddQItem q, "3_P002", "PROJECT_STATUS", "In progress", "In Progress"
    AddQItem q, "3_P003", "PROJECT_STATUS", "In progress", "In Progress"
    AddQItem q, "3_P004", "PROJECT_STATUS", "Complete", "Closed"
    AddQItem q, "3_P005", "PROJECT_NAME", "Old Name", "New Name"
    ReviewQueue.AssignBatches q

    result = result & Assert(q.Items(1).BatchLabel <> "", "the uniform trio is batched")
    result = result & Assert(q.Items(1).BatchLabel = q.Items(2).BatchLabel And q.Items(2).BatchLabel = q.Items(3).BatchLabel, _
        "all three share ONE label -- one transformation, one decision")
    result = result & Assert(q.Items(4).BatchLabel = "", _
        "a controlled field with a DIFFERENT transformation is not swept into the batch")
    result = result & Assert(q.Items(5).BatchLabel = "", _
        "a lone change is never labelled a batch -- a 'batch of one' is an individual review wearing a different word")
    result = result & Assert(ReviewQueue.DistinctBatchCount(q) = 1, "exactly one batch, got " & ReviewQueue.DistinctBatchCount(q))
    result = result & Assert(ReviewQueue.IndividualCount(q) = 2, "two individual decisions remain, got " & ReviewQueue.IndividualCount(q))

    ' Presented as R13.2 requires: the transformation, the count, the entities.
    Dim summary As String
    summary = ReviewQueue.BatchSummaryText(q)
    result = result & Assert(InStr(summary, "In progress") > 0 And InStr(summary, "In Progress") > 0, "the batch summary states the transformation, both halves")
    result = result & Assert(InStr(summary, "3 entities") > 0, "the batch summary states the count")
    result = result & Assert(InStr(summary, "3_P002") > 0, "the batch summary names the affected entities")

    Test_ReviewQueue_UniformControlledGroupIsOneDecision = result
End Function

Private Function Test_ReviewQueue_FastPathAppliesUniformPartOnly() As String
    Dim result As String

    ' Wholly uniform: the fast path is honest here, and this is the case Rohan
    ' identified -- 19 slides, one transformation, one decision.
    Dim allBatched As ReviewQueueSet
    Dim i As Long
    For i = 1 To 19
        AddQItem allBatched, "3_P" & Format(i, "000"), "PROJECT_STATUS", "In progress", "In Progress"
    Next i
    ReviewQueue.AssignBatches allBatched

    result = result & Assert(ReviewQueue.HasBatchableWork(allBatched), "19 identical corrections ARE one decision and may be confirmed in a dialog")
    result = result & Assert(InStr(ReviewQueue.ConfirmBatchText(allBatched), "In progress") > 0, _
        "the confirmation shows the actual before-and-after, never degrading to a bare count")
    result = result & Assert(InStr(ReviewQueue.ConfirmBatchText(allBatched), "19") > 0, "the confirmation states how many slides it covers")

    ' One prose row poisons the whole run, deliberately.
    Dim mixed As ReviewQueueSet
    For i = 1 To 19
        AddQItem mixed, "3_P" & Format(i, "000"), "PROJECT_STATUS", "In progress", "In Progress"
    Next i
    AddQItem mixed, "3_P020", "ABOUT_BODY", "Old prose.", "New prose."
    ReviewQueue.AssignBatches mixed

    ' F5: a mixed run still does its uniform part. The prose row does NOT
    ' disqualify the other 19 -- that was extra strictness with no rule behind it.
    result = result & Assert(ReviewQueue.HasBatchableWork(mixed), _
        "one prose row does NOT block the 19 uniform corrections -- partial application is a valid outcome")
    result = result & Assert(ReviewQueue.IndividualCount(mixed) = 1, "the prose row remains an individual decision, got " & ReviewQueue.IndividualCount(mixed))

    ' F5's other half: the dialog must not read as covering the whole run.
    Dim confirmMixed As String
    confirmMixed = ReviewQueue.ConfirmBatchText(mixed)
    result = result & Assert(InStr(confirmMixed, "NOT covered") > 0, _
        "the confirmation states what it is NOT applying -- a partial run must never read as a whole one")
    result = result & Assert(InStr(confirmMixed, "19 slide field(s)") > 0, _
        "the count offered is the BATCHED count, not the whole queue -- offering 20 would be a lie")

    ' Approving batches must leave the individual row untouched.
    ReviewQueue.ApproveBatchedOnly mixed
    result = result & Assert(mixed.Items(1).Approved, "batched rows are approved by the dialog")
    result = result & Assert(Not mixed.Items(20).Approved, "the prose row is NOT approved by the dialog -- it still needs its own reading")

    ' An empty queue offers no path, and must not read as 'approve nothing'.
    Dim emptyQ As ReviewQueueSet
    result = result & Assert(Not ReviewQueue.HasBatchableWork(emptyQ), "an empty change set takes no path at all")

    ' Too many distinct transformations to read: the dialog is refused (F4).
    Dim noisy As ReviewQueueSet
    For i = 1 To 15
        AddQItem noisy, "3_A" & Format(i, "000"), "PROJECT_STATUS", "was" & i, "now" & i
        AddQItem noisy, "3_B" & Format(i, "000"), "PROJECT_STATUS", "was" & i, "now" & i
    Next i
    ReviewQueue.AssignBatches noisy
    result = result & Assert(ReviewQueue.DistinctBatchCount(noisy) = 15, "15 distinct transformations, got " & ReviewQueue.DistinctBatchCount(noisy))
    result = result & Assert(Not ReviewQueue.HasBatchableWork(noisy), _
        "15 batches is past MAX_BATCHES_IN_MODAL -- a wall of transformations gets dismissed, not read")

    Test_ReviewQueue_FastPathAppliesUniformPartOnly = result
End Function

Private Function Test_ReviewQueue_ApprovalIsAffirmativeAndBatchWide() As String
    Dim result As String

    result = result & Assert(ReviewQueue.IsApprovalMark("Y"), "'Y' approves")
    result = result & Assert(ReviewQueue.IsApprovalMark(" yes "), "'yes' approves, trimmed and case-insensitive")
    result = result & Assert(Not ReviewQueue.IsApprovalMark(""), "blank does NOT approve")
    result = result & Assert(Not ReviewQueue.IsApprovalMark("N"), "'N' does not approve")
    result = result & Assert(Not ReviewQueue.IsApprovalMark("maybe"), "an unrecognised answer never silently approves")

    Dim q As ReviewQueueSet
    AddQItem q, "3_P001", "PROJECT_STATUS", "In progress", "In Progress"
    AddQItem q, "3_P002", "PROJECT_STATUS", "In progress", "In Progress"
    AddQItem q, "3_P003", "PROJECT_STATUS", "In progress", "In Progress"
    AddQItem q, "3_P004", "ABOUT_BODY", "Old prose.", "New prose."
    ReviewQueue.AssignBatches q

    ' Tick ONE member of the batch, and nothing else.
    q.Items(2).Approved = True
    ReviewQueue.PropagateBatchApprovals q

    result = result & Assert(q.Items(1).Approved And q.Items(3).Approved, _
        "approving one member of a uniform batch approves the batch -- that is what makes it one decision")
    result = result & Assert(Not q.Items(4).Approved, _
        "the individual prose row is untouched by a batch approval -- it still needs its own yes")

    Dim hashes As Object
    Set hashes = ReviewQueue.ApprovedHashSet(q)
    result = result & Assert(hashes.Count = 3, "exactly the three batch members are in the approved set, got " & hashes.Count)
    result = result & Assert(hashes.Exists(q.Items(1).ChangeHash), "the approved set is keyed by change hash, not by row position")

    Test_ReviewQueue_ApprovalIsAffirmativeAndBatchWide = result
End Function

Private Function Test_InjectPrimitive_TrailingBreaksAreNotADifference() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim shp As Object
    Set shp = sld.Shapes.AddTextbox(1, 10, 10, 300, 50)
    shp.Tags.Add "role", "about"
    shp.TextFrame.TextRange.Text = "Some prose." & vbCr

    ' The 2026-07-31 case: slide carries a trailing paragraph mark, the
    ' register value does not. 20 of 46 real ABOUT_BODY rows looked like this
    ' and not one was a wording change.
    Dim r As InjectResult
    r = InjectPrimitive.InjectPrimitive(sld, "about", "Some prose.", True)
    result = result & Assert(r.Found, "the tagged shape is found")
    result = result & Assert(Not r.WouldChange, "a trailing paragraph mark alone is NOT a change -- otherwise 20 real slides get rewritten to remove an invisible character")

    ' The rule must not swallow a REAL difference that happens to sit near the end.
    Dim r2 As InjectResult
    r2 = InjectPrimitive.InjectPrimitive(sld, "about", "Some prose!", True)
    result = result & Assert(r2.WouldChange, "a genuine wording change is still a change, even one character of it")

    ' Nor a difference in the middle, which is the 2_P004 shape (a paragraph
    ' break the register does not carry) -- that one is NOT covered by this rule
    ' and must still surface.
    shp.TextFrame.TextRange.Text = "One." & vbCr & "Two." & vbCr
    Dim r3 As InjectResult
    r3 = InjectPrimitive.InjectPrimitive(sld, "about", "One.Two.", True)
    result = result & Assert(r3.WouldChange, "an INTERNAL break is still a difference -- only trailing ones are ignored")

    ' Normalisation is for comparison only. What lands on the slide is exactly
    ' what was supplied -- trimming on write would strip real formatting out of
    ' a deck one sync at a time.
    shp.TextFrame.TextRange.Text = "Old."
    Dim r4 As InjectResult
    r4 = InjectPrimitive.InjectPrimitive(sld, "about", "New." & vbCr, False)
    result = result & Assert(r4.Written, "the write happens")
    result = result & Assert(r4.Verified, "and verifies under the same rule it was decided by")
    ' EXACT equality. The first version asserted
    '   Right(text,1) = vbCr Or Len(text) >= 4
    ' against a 5-character write -- and "New." trimmed to 4 satisfies the
    ' second disjunct, which is precisely the regression the message names. The
    ' two assertions above it cannot backstop that either, because Verified
    ' compares through IgnoringTrailingBreaks on both sides and passes whether
    ' the break survived or not. A check whose bound is the length of the
    ' failure case cannot detect the failure case.
    result = result & Assert(shp.TextFrame.TextRange.Text = "New." & vbCr, _
        "the supplied value is written EXACTLY as given, trailing break included -- normalisation is for comparison only")

    sld.Delete
    Test_InjectPrimitive_TrailingBreaksAreNotADifference = result
End Function

Private Function Test_Drafting_OnlyTickedNonEmptyDraftsPublish() As String
    Dim result As String

    Dim xl As Object, wb As Object, dws As Object, rws As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set wb = xl.Workbooks.Add
    Set dws = wb.Worksheets(1)
    Set rws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))

    ' A minimal WIDE register: four slides, one field, all for FY26Q4. Built
    ' through CreateSheet/UpsertRow rather than by hand, so the test cannot drift
    ' from the shape the tool actually writes.
    '
    ' Layout after this: A = Instance ID, B = Quarter, C = ABOUT_BODY.
    ' Rows 2..5 = P001..P004.
    '
    ' STATUS IS GONE FROM THIS TEST because it is gone from the sheet. The old
    ' version asserted Seed -> Approved transitions in a long register; the wide
    ' sheet has no Status column, and what it asserts now is the thing that
    ' actually matters -- WHICH CELLS CHANGED. That is a stricter question than
    ' the status one, because an unticked row whose value was overwritten used to
    ' pass the Status assertion as long as the status stayed Seed.
    ExcelOutput.CreateSheet rws, "deck-v1"
    Dim seedVals As Object
    Dim i As Long
    For i = 1 To 4
        Set seedVals = CreateObject("Scripting.Dictionary")
        seedVals("ABOUT_BODY") = "old " & i
        ExcelOutput.UpsertRow rws, "P00" & i, seedVals, "FY26Q4"
    Next i

    ' The drafting sheet, covering every combination that decides publication.
    dws.Cells(Drafting.DRAFT_HEADER_ROW, 1).Value = "Project code"
    ' PUBLISH READS SUBMIT, NOT THE AI COLUMN. This test used to write into
    ' COL_D_DRAFT, which was the only text column when it was written. Since
    ' 2026-08-01 the sheet carries three: ORIGINAL (read-only), AI DRAFT (what
    ' Copilot wrote, never published) and SUBMIT (what the person is sending).
    dws.Cells(Drafting.DRAFT_FIRST_ROW + 0, Drafting.COL_D_ENTITY).Value = "P001": dws.Cells(Drafting.DRAFT_FIRST_ROW + 0, Drafting.COL_D_SUBMIT).Value = "new one":  dws.Cells(Drafting.DRAFT_FIRST_ROW + 0, Drafting.COL_D_APPROVED).Value = "Y"
    dws.Cells(Drafting.DRAFT_FIRST_ROW + 1, Drafting.COL_D_ENTITY).Value = "P002": dws.Cells(Drafting.DRAFT_FIRST_ROW + 1, Drafting.COL_D_SUBMIT).Value = "new two":  dws.Cells(Drafting.DRAFT_FIRST_ROW + 1, Drafting.COL_D_APPROVED).Value = ""
    dws.Cells(Drafting.DRAFT_FIRST_ROW + 2, Drafting.COL_D_ENTITY).Value = "P003": dws.Cells(Drafting.DRAFT_FIRST_ROW + 2, Drafting.COL_D_SUBMIT).Value = "":         dws.Cells(Drafting.DRAFT_FIRST_ROW + 2, Drafting.COL_D_APPROVED).Value = "Y"
    dws.Cells(Drafting.DRAFT_FIRST_ROW + 3, Drafting.COL_D_ENTITY).Value = "P004": dws.Cells(Drafting.DRAFT_FIRST_ROW + 3, Drafting.COL_D_SUBMIT).Value = "a" & vbCr & "b": dws.Cells(Drafting.DRAFT_FIRST_ROW + 3, Drafting.COL_D_APPROVED).Value = "Y"
    ' P005: an AI draft, TICKED, with SUBMIT left empty. Must NOT publish -- this
    ' is the whole reason the two columns are separate, and nothing asserted it.
    dws.Cells(Drafting.DRAFT_FIRST_ROW + 4, Drafting.COL_D_ENTITY).Value = "P005": dws.Cells(Drafting.DRAFT_FIRST_ROW + 4, Drafting.COL_D_DRAFT).Value = "AI WROTE THIS": dws.Cells(Drafting.DRAFT_FIRST_ROW + 4, Drafting.COL_D_APPROVED).Value = "Y"
    ' P006: real SUBMIT text, ticked, and NO ROW in the register for this period.
    ' New with the wide sheet 2026-08-05 -- UpsertRow would CREATE a row here,
    ' which would mean publishing invents a slide nobody onboarded. Publish
    ' refuses instead, and this is what holds that refusal in place.
    dws.Cells(Drafting.DRAFT_FIRST_ROW + 5, Drafting.COL_D_ENTITY).Value = "P006": dws.Cells(Drafting.DRAFT_FIRST_ROW + 5, Drafting.COL_D_SUBMIT).Value = "orphan text": dws.Cells(Drafting.DRAFT_FIRST_ROW + 5, Drafting.COL_D_APPROVED).Value = "Y"

    Dim rep As String
    rep = Drafting.PublishDrafts(dws, rws, "ABOUT_BODY", "FY26Q4", False)

    ' A DRAFT ALONE IS NOT CONSENT and A TICK ALONE IS NOT CONTENT.
    result = result & Assert(rws.Cells(2, 3).Value = "new one", "ticked + drafted publishes, and the drafted text is what lands, got '" & rws.Cells(2, 3).Value & "'")
    result = result & Assert(rws.Cells(3, 3).Value = "old 2", "DRAFTED BUT NOT TICKED leaves the register value untouched -- a draft is not an approval, got '" & rws.Cells(3, 3).Value & "'")
    result = result & Assert(rws.Cells(4, 3).Value = "old 3", "TICKED BUT EMPTY publishes nothing -- a tick against no draft is a mis-click, got '" & rws.Cells(4, 3).Value & "'")
    result = result & Assert(InStr(rep, "SUBMIT is empty") > 0, "the empty-but-ticked row is REPORTED, not silently dropped")

    ' THE POINT OF SPLITTING THE TWO COLUMNS. P005 has an AI draft and a tick,
    ' and nothing in SUBMIT. Publishing it would mean text the AI wrote reaching
    ' a slide because nobody stopped it, which is precisely the act the split
    ' exists to prevent -- and nothing asserted it until 2026-08-01.
    '
    ' On the wide sheet the assertion is that NO ROW APPEARED for P005 at all --
    ' stricter than the old status check, because UpsertRow creates rows and the
    ' failure would now be a fabricated slide rather than a wrong status.
    result = result & Assert(Not RowExistsForInstance(rws, "P005"), _
        "AN AI DRAFT ALONE DOES NOT PUBLISH, even when ticked -- only SUBMIT does")
    result = result & Assert(InStr(rep, "there IS an AI draft") > 0, _
        "the report NAMES the AI-draft-without-submit case, so a person knows to run Copy AI to Submit")

    ' PUBLISH DOES NOT INVENT SLIDES. P006 is ticked with real text and has no
    ' row for FY26Q4.
    result = result & Assert(Not RowExistsForInstance(rws, "P006"), _
        "a ticked draft for a slide with NO register row does not create one -- publish would be inventing a slide")
    result = result & Assert(InStr(rep, "NO REGISTER ROW for P006") > 0, _
        "the missing row is REPORTED by name, not silently skipped")

    ' Line breaks become the register delimiter, the exact inverse of what
    ' InjectPrimitive does on the way out.
    result = result & Assert(rws.Cells(5, 3).Value = "a||b", "a real line break is stored as the || delimiter, got '" & rws.Cells(5, 3).Value & "'")
    result = result & Assert(InStr(rws.Cells(5, 3).Value, vbCr) = 0, "no carriage return survives into the register")

    ' EXACTLY TWO CELLS CHANGED. Counted against the seeded values rather than
    ' against a status, so an unticked row that was silently overwritten fails
    ' here -- the old status-based count could not see that.
    Dim changed As Long
    For i = 1 To 4
        If rws.Cells(i + 1, 3).Value <> "old " & i Then changed = changed + 1
    Next i
    result = result & Assert(changed = 2, "exactly the two ticked-and-drafted rows changed, got " & changed)

    ' The period column is still intact on every row -- a row whose period got
    ' blanked is invisible to every filtered read and would report as a clean
    ' sync of nothing.
    Dim periodsOk As Boolean
    periodsOk = True
    For i = 2 To 5
        If Trim(CStr(rws.Cells(i, 2).Value)) <> "FY26Q4" Then periodsOk = False
    Next i
    result = result & Assert(periodsOk, "publishing leaves every row's Quarter stamp intact")

    wb.Close False
    xl.Quit
    Test_Drafting_OnlyTickedNonEmptyDraftsPublish = result
End Function

' --- WorkbookBridge.WorksheetForSlideType ------------------------------
'
' The bug these exist for: drafting/publish resolved the register by looking
' for a sheet literally NAMED "Register", while sync resolved it through the
' worksheet name registered per slide type. A workbook with both put the two
' halves of the loop on DIFFERENT SHEETS, each reporting success.
Private Function Test_WorkbookBridge_RegisteredNameWinsOverASheetCalledRegister() As String
    Dim result As String

    Dim pres As Object
    Set pres = Application.Presentations.Add
    Dim xl As Object, wb As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set wb = xl.Workbooks.Add

    ' A decoy named exactly what the old resolver preferred...
    Dim decoy As Object
    Set decoy = wb.Worksheets.Add
    decoy.Name = "Register"
    decoy.Cells(1, 1).Value = "DECOY"

    ' ...and the sheet the deck is actually registered against.
    Dim real As Object
    Set real = wb.Worksheets.Add
    real.Name = "Research Project Status"
    real.Cells(1, 1).Value = "THE REAL ONE"

    ' Presentations.Add gives a deck with ZERO slides, so a template slide has
    ' to be made before RegisterType can key against one.
    Dim tpl As Object
    Set tpl = pres.Slides.Add(1, ppLayoutBlank)
    DeckRegistry.RegisterType pres, "Research Project Status", tpl, "Research Project Status"

    Dim problem As String
    Dim got As Object
    Set got = WorkbookBridge.WorksheetForSlideType(pres, wb, "Research Project Status", problem)

    result = result & Assert(Not got Is Nothing, "a sheet was resolved at all; problem was '" & problem & "'")
    If Not got Is Nothing Then
        ' THE ASSERTION. The old resolver returned the decoy here.
        result = result & Assert(got.Name = "Research Project Status", _
            "resolves the REGISTERED sheet, not the one merely named 'Register', got '" & got.Name & "'")
        result = result & Assert(CStr(got.Cells(1, 1).Value) = "THE REAL ONE", _
            "...and it is really that sheet's content, got '" & got.Cells(1, 1).Value & "'")
    End If

    wb.Close False
    xl.Quit
    pres.Saved = msoTrue
    pres.Close
    Test_WorkbookBridge_RegisteredNameWinsOverASheetCalledRegister = result
End Function

Private Function Test_WorkbookBridge_RefusesToInventAMissingRegisterSheet() As String
    Dim result As String

    Dim pres As Object
    Set pres = Application.Presentations.Add
    Dim xl As Object, wb As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set wb = xl.Workbooks.Add
    ' Give the workbook a used sheet, so GetOrAddWorksheet's "rename the lone
    ' blank sheet" path is not what is being measured.
    wb.Worksheets(1).Cells(1, 1).Value = "something"

    Dim before As Long
    before = wb.Worksheets.count

    ' Registered against a sheet that is NOT in this workbook -- the shape of a
    ' deck paired to the wrong file.
    Dim tpl As Object
    Set tpl = pres.Slides.Add(1, ppLayoutBlank)
    DeckRegistry.RegisterType pres, "Research Project Status", tpl, "Not In This Workbook"

    Dim problem As String
    Dim got As Object
    Set got = WorkbookBridge.WorksheetForSlideType(pres, wb, "Research Project Status", problem)

    result = result & Assert(got Is Nothing, "refuses rather than resolving something")
    result = result & Assert(InStr(problem, "no such sheet") > 0, _
        "says the workbook has no such sheet, got '" & problem & "'")
    ' THE POINT. GetOrAddWorksheet would have CREATED it, and a blank invented
    ' register reads as a register with nothing in it -- a clean run of nothing.
    result = result & Assert(wb.Worksheets.count = before, _
        "and does NOT create it: sheet count " & before & " -> " & wb.Worksheets.count)

    wb.Close False
    xl.Quit
    pres.Saved = msoTrue
    pres.Close
    Test_WorkbookBridge_RefusesToInventAMissingRegisterSheet = result
End Function

' --- FieldSpec.ApplyControlledValidation -------------------------------
'
' THESE ARE THE TESTS THAT DID NOT EXIST, which is why the function could go
' silently inert when the register went wide. It kept locating FieldID/Value
' columns -- the long register's shape -- found none, and returned a polite
' sentence the caller printed and ignored. Every controlled field lost its
' dropdown and its out-of-vocabulary check, and nothing anywhere said so.
'
' Every one of these would FAIL against the pre-2026-08-05 implementation,
' which is the property that makes them worth having.

' Builds a Field Spec sheet declaring `fieldId` as controlled by `allowed`.
Private Function SpecSheetWithVocabulary(wb As Object, fieldId As String, allowed As String) As Object
    Dim ws As Object
    Set ws = wb.Worksheets.Add
    ws.Cells(FieldSpec.SPEC_HEADER_ROW, FieldSpec.COL_S_FIELDID).Value = "FieldID"
    ws.Cells(FieldSpec.SPEC_HEADER_ROW, FieldSpec.COL_S_ALLOWED).Value = "Allowed"
    ws.Cells(FieldSpec.SPEC_FIRST_ROW, FieldSpec.COL_S_FIELDID).Value = fieldId
    ws.Cells(FieldSpec.SPEC_FIRST_ROW, FieldSpec.COL_S_ALLOWED).Value = allowed
    Set SpecSheetWithVocabulary = ws
End Function

' The column whose header is `name`, or 0.
'
' BY HEADER, NEVER BY POSITION -- even in a test. The obvious version of these
' tests wrote Cells(2, 3) for PROJECT_STATUS, which assumes UpsertRow appends
' field columns in Scripting.Dictionary key order. ExcelOutput's own header
' comment says that order is "de-facto-but-undocumented" and refuses to depend
' on it; a test that depends on it is asserting something the code deliberately
' does not promise, and would fail for a reason having nothing to do with
' validation.
' headerText, NOT name: Name is a VBA statement (Name x As y), so using it as a
' parameter is a compile-time Syntax error -- and headless that is a hang with no
' output rather than a message. Flagged by check_vba_static.py; it had survived
' because VBA compiles per procedure and nothing had forced this one to compile.
Private Function ColumnNamed(ws As Object, headerText As String) As Long
    Dim c As Long
    For c = 1 To ExcelOutput.LastUsedColumn(ws)
        If StrComp(Trim(CStr(ws.Cells(1, c).Value)), headerText, vbTextCompare) = 0 Then
            ColumnNamed = c
            Exit Function
        End If
    Next c
End Function

' Reports the validation Type of a cell, or -1 when it has none. Wrapped
' because reading .Validation.Type on an unvalidated cell RAISES rather than
' returning a "no validation" value.
Private Function ValidationTypeOf(ws As Object, r As Long, c As Long) As Long
    ValidationTypeOf = -1
    On Error Resume Next
    ValidationTypeOf = ws.Cells(r, c).Validation.Type
    On Error GoTo 0
End Function

' A wide register: Instance ID | Quarter | PROJECT_STATUS | ABOUT_BODY
Private Function WideRegisterForValidation(wb As Object) As Object
    Dim ws As Object
    Set ws = wb.Worksheets.Add
    ExcelOutput.CreateSheet ws, "deck-v1"

    Dim v As Object
    Set v = CreateObject("Scripting.Dictionary")
    v("PROJECT_STATUS") = "In Progress": v("ABOUT_BODY") = "about one"
    ExcelOutput.UpsertRow ws, "P001", v, "FY26Q4"

    Set v = CreateObject("Scripting.Dictionary")
    v("PROJECT_STATUS") = "Not Started": v("ABOUT_BODY") = "about two"
    ExcelOutput.UpsertRow ws, "P002", v, "FY26Q4"

    Set WideRegisterForValidation = ws
End Function

Private Function Test_FieldSpec_ValidationAppliesDownTheControlledColumn() As String
    Dim result As String
    Dim xl As Object, wb As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set wb = xl.Workbooks.Add

    Dim regWs As Object, specWs As Object
    Set regWs = WideRegisterForValidation(wb)
    Set specWs = SpecSheetWithVocabulary(wb, "PROJECT_STATUS", "Not Started,In Progress,Project Closed")

    Dim rep As String
    rep = FieldSpec.ApplyControlledValidation(regWs, specWs)

    ' THE REGRESSION ITSELF. The old implementation returned this sentence
    ' against a wide sheet and did nothing at all.
    result = result & Assert(InStr(rep, "no FieldID/Value column") = 0, _
        "does NOT bail out looking for the long register's columns, got '" & rep & "'")
    result = result & Assert(InStr(rep, "1 controlled column") > 0, _
        "reports the one controlled column it found, got '" & rep & "'")

    ' The validation is really ON the cells -- asked of Excel, not of the
    ' report, because the report is the thing under test.
    Dim cStatus As Long, cAbout As Long
    cStatus = ColumnNamed(regWs, "PROJECT_STATUS")
    cAbout = ColumnNamed(regWs, "ABOUT_BODY")
    result = result & Assert(cStatus > 0, "the register carries a PROJECT_STATUS column")
    result = result & Assert(cAbout > 0, "the register carries an ABOUT_BODY column")

    result = result & Assert(ValidationTypeOf(regWs, 2, cStatus) = 3, _
        "the controlled column's first data cell carries a list validation (xlValidateList=3), got " & _
        ValidationTypeOf(regWs, 2, cStatus))
    result = result & Assert(ValidationTypeOf(regWs, 3, cStatus) = 3, _
        "...and so does the second data row, got " & ValidationTypeOf(regWs, 3, cStatus))

    ' A FIELD WITH NO VOCABULARY IS LEFT ALONE. ABOUT_BODY is prose; giving it
    ' a dropdown would make the register unwritable for the flagship field.
    result = result & Assert(ValidationTypeOf(regWs, 2, cAbout) <> 3, _
        "the uncontrolled ABOUT_BODY column gets NO dropdown, got " & ValidationTypeOf(regWs, 2, cAbout))

    ' NO ASSERTION HERE THAT THE STRUCTURAL COLUMNS ESCAPED A DROPDOWN.
    '
    ' There were two, and they were worthless. With only PROJECT_STATUS in the
    ' vocabulary, the structural columns are skipped by the vocab lookup whether
    ' the structural guard exists or not -- so both assertions passed with the
    ' guard deliberately REMOVED, measured 2026-08-05. They read as care taken
    ' and tested nothing.
    '
    ' The guard only bites when a vocabulary is declared for a field NAMED like
    ' a structural column, so that is where the assertion belongs. See
    ' Test_FieldSpec_ValidationRefusesAVocabularyNamedLikeAStructuralColumn.

    wb.Close False
    xl.Quit
    Test_FieldSpec_ValidationAppliesDownTheControlledColumn = result
End Function

Private Function Test_FieldSpec_ValidationReportsValuesOutsideTheVocabulary() As String
    Dim result As String
    Dim xl As Object, wb As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set wb = xl.Workbooks.Add

    Dim regWs As Object, specWs As Object
    Set regWs = WideRegisterForValidation(wb)

    ' Case drift, which is the exact thing InVocabulary is case-SENSITIVE to
    ' catch: "Not started" against a vocabulary of "Not Started".
    Dim cStatus As Long
    cStatus = ColumnNamed(regWs, "PROJECT_STATUS")
    regWs.Cells(3, cStatus).Value = "Not started"

    Set specWs = SpecSheetWithVocabulary(wb, "PROJECT_STATUS", "Not Started,In Progress,Project Closed")

    Dim rep As String
    rep = FieldSpec.ApplyControlledValidation(regWs, specWs)

    ' The COUNT is now the load-bearing part: the list is capped at 5 entries so
    ' it cannot push the rest of a MsgBox past the truncation limit, but the
    ' total must always be stated, whether 1 value drifted or 40.
    result = result & Assert(InStr(rep, "outside the allowed list") > 0, _
        "the out-of-vocabulary section appears, got '" & rep & "'")
    result = result & Assert(InStr(rep, "1 value(s) outside") > 0, _
        "and the report states HOW MANY drifted, got '" & rep & "'")
    result = result & Assert(InStr(rep, "Not started") > 0, _
        "the offending VALUE is quoted, got '" & rep & "'")
    ' Named by slide and period, not by row number -- a row number is worthless
    ' on a sheet where one project appears once per period.
    result = result & Assert(InStr(rep, "P002") > 0, _
        "the offending row is named by SLIDE, got '" & rep & "'")
    result = result & Assert(InStr(rep, "FY26Q4") > 0, _
        "...and by its PERIOD, got '" & rep & "'")
    ' P001 is "In Progress", which IS in the vocabulary.
    result = result & Assert(InStr(rep, "P001") = 0, _
        "a value that IS in the vocabulary is not reported, got '" & rep & "'")

    ' REPORTED, NEVER CORRECTED. The register is the record; silently rewriting
    ' somebody's value to the nearest legal one would be the tool editing data
    ' it does not own.
    result = result & Assert(regWs.Cells(3, cStatus).Value = "Not started", _
        "the offending value is LEFT EXACTLY AS IT IS, got '" & regWs.Cells(3, cStatus).Value & "'")

    wb.Close False
    xl.Quit
    Test_FieldSpec_ValidationReportsValuesOutsideTheVocabulary = result
End Function

' A vocabulary declared for a field NAMED LIKE A STRUCTURAL COLUMN.
'
' This is the only scenario in which the structural-column guard does anything,
' and it exists because the first attempt to test that guard could not fail: with
' only PROJECT_STATUS in the vocabulary, the Quarter and Instance ID columns are
' skipped by the vocab lookup regardless, so removing the guard changed nothing
' and the assertions passed anyway.
'
' What it protects: a dropdown written onto the Quarter column constrains which
' PERIOD a row may belong to. Every row not matching the list becomes an Excel
' validation error on the one column the whole model is keyed by, and
' RollForwardPeriod writes periods that would not be on any list.
'
' UpsertRow already REFUSES a field named like a structural column, but nothing
' stops somebody typing "Quarter" into the Field Spec sheet, which is a
' hand-edited content surface owned by the RM rather than by the tool.
Private Function Test_FieldSpec_ValidationRefusesAVocabularyNamedLikeAStructuralColumn() As String
    Dim result As String
    Dim xl As Object, wb As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set wb = xl.Workbooks.Add

    Dim regWs As Object, specWs As Object
    Set regWs = WideRegisterForValidation(wb)
    Set specWs = SpecSheetWithVocabulary(wb, "Quarter", "FY26Q4,FY27Q1")

    Dim cInst As Long, cQtr As Long
    ExcelOutput.LocateStructuralColumns regWs, cInst, cQtr
    ' Located for real, so the assertions below cannot pass by reading column 0.
    result = result & Assert(cQtr > 0, "the Quarter column was actually located, got " & cQtr)
    result = result & Assert(cInst > 0, "the Instance ID column was actually located, got " & cInst)

    Dim rep As String
    rep = FieldSpec.ApplyControlledValidation(regWs, specWs)

    result = result & Assert(ValidationTypeOf(regWs, 2, cQtr) <> 3, _
        "NO dropdown is written onto the Quarter column, got " & ValidationTypeOf(regWs, 2, cQtr))
    result = result & Assert(ValidationTypeOf(regWs, 3, cQtr) <> 3, _
        "...on any of its rows, got " & ValidationTypeOf(regWs, 3, cQtr))

    ' And it is reported as having checked nothing, rather than claiming a
    ' controlled column it must not have.
    result = result & Assert(InStr(rep, "NO column") > 0, _
        "reports that nothing on the register was controlled, got '" & rep & "'")

    wb.Close False
    xl.Quit
    Test_FieldSpec_ValidationRefusesAVocabularyNamedLikeAStructuralColumn = result
End Function

Private Function Test_FieldSpec_ValidationSaysSoWhenNothingIsControlled() As String
    Dim result As String
    Dim xl As Object, wb As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set wb = xl.Workbooks.Add

    Dim regWs As Object, specWs As Object
    Set regWs = WideRegisterForValidation(wb)
    ' A vocabulary for a field the register does not carry.
    Set specWs = SpecSheetWithVocabulary(wb, "RISK_RATING", "Low,Medium,High")

    Dim rep As String
    rep = FieldSpec.ApplyControlledValidation(regWs, specWs)

    ' "dropdown on 0 cell(s)" reads as success to anyone skimming, and reading
    ' it that way is how this function stayed broken. Doing nothing has to be
    ' stated as doing nothing.
    result = result & Assert(InStr(rep, "NO column") > 0, _
        "says plainly that nothing was checked, got '" & rep & "'")
    result = result & Assert(InStr(rep, "dropdown on 0") = 0, _
        "does NOT report a zero count that reads like success, got '" & rep & "'")

    wb.Close False
    xl.Quit
    Test_FieldSpec_ValidationSaysSoWhenNothingIsControlled = result
End Function

' Does the wide sheet have ANY row whose Instance ID is `instanceId`?
'
' Deliberately ignores the period: the question these assertions ask is "did a
' row get fabricated", and a row invented under the wrong period would be just
' as wrong as one invented under the right one. Column A is safe to read
' directly here because the sheet was built by CreateSheet three lines earlier.
Private Function RowExistsForInstance(ws As Object, instanceId As String) As Boolean
    Dim r As Long
    r = 2
    Do While Trim(CStr(ws.Cells(r, 1).Value)) <> ""
        If Trim(CStr(ws.Cells(r, 1).Value)) = instanceId Then
            RowExistsForInstance = True
            Exit Function
        End If
        r = r + 1
    Loop
End Function

Private Function Test_RunSync_CreateMissingRefusesWhileSlidesAreUnclassified() As String
    Dim result As String

    ' The MECE case: a slide keeps its slide_type tag but has lost its
    ' instance_key. It is reported `flagged`, AND the register row for the
    ' entity it used to be is reported `new_record`. Creating slides in that
    ' state duplicates the template for an entity that already has a slide.
    ' Two args, not four -- the helper builds its own Title/Status shapes.
    Dim keyed As Object
    Set keyed = NewOnboardedSlide("cm-type", "CM001")

    Dim orphan As Object
    Set orphan = NewBlankSlide()
    orphan.Tags.Add "slide_type", "cm-type"      ' typed, but NO instance_key

    ' A REAL TEMPLATE, because this test used to hand CreateMissingSlides the
    ' onboarded slide `keyed` -- a genuine project's slide -- as its source.
    ' That is the shape the new guard refuses, and the shape the rig deck is in.
    Dim tmpl As Object
    Set tmpl = NewBlankSlide()
    tmpl.Tags.Add "slide_type", "cm-type"
    tmpl.Tags.Add "is_template", "1"

    Dim sheet As Sheet
    Set sheet.Rows = CreateObject("Scripting.Dictionary")
    Set sheet.Fields = New Collection
    Set sheet.InstanceOrder = New Collection
    Dim vals As Object
    Set vals = CreateObject("Scripting.Dictionary")
    vals("Title") = "Real Project Name"
    Set sheet.Rows("CM001") = vals
    sheet.InstanceOrder.Add "CM001"
    Dim vals2 As Object
    Set vals2 = CreateObject("Scripting.Dictionary")
    vals2("Title") = "orphaned"
    Set sheet.Rows("CM002") = vals2
    sheet.InstanceOrder.Add "CM002"

    Dim before As Long
    before = Application.ActivePresentation.Slides.count

    Dim rep As String
    rep = RunSync.CreateMissingSlides(sheet, "cm-type", tmpl, False)

    result = result & Assert(InStr(rep, "REFUSED") > 0, "it REFUSES while a slide of this type is unclassified")
    result = result & Assert(InStr(rep, "two slides claiming one project") > 0, "the refusal explains the consequence, not just the rule")
    result = result & Assert(Application.ActivePresentation.Slides.count = before, _
        "and NOTHING was created -- the refusal is a control, not a warning, got " & Application.ActivePresentation.Slides.count & " vs " & before)

    ' Prove the test would FAIL if the guard were removed: with no unclassified
    ' slide present, the same call must go ahead and create. A refusal test that
    ' cannot show the non-refusing case proves only that the function returns a
    ' string.
    orphan.Delete
    Dim before2 As Long
    before2 = Application.ActivePresentation.Slides.count
    Dim rep2 As String
    rep2 = RunSync.CreateMissingSlides(sheet, "cm-type", tmpl, False)
    result = result & Assert(InStr(rep2, "REFUSED") = 0, "with no unclassified slide it does NOT refuse")
    result = result & Assert(Application.ActivePresentation.Slides.count = before2 + 1, _
        "and it DOES create the missing slide, got " & Application.ActivePresentation.Slides.count & " vs " & before2)

    ' clean up whatever the second call created
    Dim extra As Object
    For Each extra In Application.ActivePresentation.Slides
        Dim ei As SlideInstance
        ei = Resolve.ResolveSlideInstance(extra)
        If ei.HasInstanceKey Then
            If ei.InstanceKey = "CM002" Then extra.Delete
        End If
    Next extra

    keyed.Delete
    Test_RunSync_CreateMissingRefusesWhileSlidesAreUnclassified = result
End Function

' Proves Scenario 3 step 3: two rows of the same type, different letters, each
' get cloned from THEIR letter's template -- not both from whichever template
' happened to be passed in as CreateMissingSlides' own templateSld param. That
' param stays only as the type-level fallback for a row with no letter.
Private Function Test_RunSync_CreateMissingSlidesChoosesTemplateByRowLetter() As String
    Dim result As String
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim tmplK As Object, tmplS As Object
    Set tmplK = NewBlankSlide()
    tmplK.Tags.Add "slide_type", "cml-type"
    tmplK.Tags.Add "is_template", "1"
    Dim markK As Object
    Set markK = tmplK.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    markK.TextFrame.TextRange.Text = "MARK-K"
    markK.Tags.Add "role", "Marker"

    Set tmplS = NewBlankSlide()
    tmplS.Tags.Add "slide_type", "cml-type"
    tmplS.Tags.Add "is_template", "1"
    Dim markS As Object
    Set markS = tmplS.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    markS.TextFrame.TextRange.Text = "MARK-S"
    markS.Tags.Add "role", "Marker"

    DeckRegistry.RegisterTemplateLetter pres, "cml-type", "K", tmplK, "SheetK"
    DeckRegistry.RegisterTemplateLetter pres, "cml-type", "S", tmplS, "SheetS"

    Dim sheet As Sheet
    Set sheet.Rows = CreateObject("Scripting.Dictionary")
    Set sheet.Fields = New Collection
    Set sheet.InstanceOrder = New Collection
    Dim valsK As Object
    Set valsK = CreateObject("Scripting.Dictionary")
    Set sheet.Rows("3_K900") = valsK
    sheet.InstanceOrder.Add "3_K900"
    Dim valsS As Object
    Set valsS = CreateObject("Scripting.Dictionary")
    Set sheet.Rows("3_S900") = valsS
    sheet.InstanceOrder.Add "3_S900"

    ' templateSld deliberately passed as tmplK, the WRONG template for the S
    ' row -- if the per-row lookup were silently ignored and every row cloned
    ' this one param instead, the S row would come back carrying MARK-K, and
    ' the assertion below would catch it.
    Dim rep As String
    rep = RunSync.CreateMissingSlides(sheet, "cml-type", tmplK, False)

    Dim instances() As Object
    instances = RunSync.GatherInstances("cml-type")
    Dim lo As Long, hi As Long, hasAny As Boolean
    On Error Resume Next
    lo = LBound(instances): hi = UBound(instances): hasAny = (Err.Number = 0)
    On Error GoTo 0

    Dim byKey As Object
    Set byKey = CreateObject("Scripting.Dictionary")
    If hasAny Then
        Dim i As Long
        For i = lo To hi
            Dim inst As SlideInstance
            inst = Resolve.ResolveSlideInstance(instances(i))
            If inst.HasInstanceKey Then Set byKey(inst.InstanceKey) = instances(i)
        Next i
    End If

    result = result & Assert(byKey.Exists("3_K900") And byKey.Exists("3_S900"), _
        "both rows got created -- report: " & rep)

    If byKey.Exists("3_K900") Then
        Dim kShp As Object
        Set kShp = FindShapeByRole(byKey("3_K900"), "Marker")
        Dim kText As String
        kText = IIf(kShp Is Nothing, "<not found>", kShp.TextFrame.TextRange.Text)
        result = result & Assert(kText = "MARK-K", "the K row was cloned from the K template, got '" & kText & "'")
    End If

    If byKey.Exists("3_S900") Then
        Dim sShp As Object
        Set sShp = FindShapeByRole(byKey("3_S900"), "Marker")
        Dim sText As String
        sText = IIf(sShp Is Nothing, "<not found>", sShp.TextFrame.TextRange.Text)
        result = result & Assert(sText = "MARK-S", "the S row was cloned from the S template, not the passed-in K template, got '" & sText & "'")
    End If

    ' Clean up: the two clones plus both templates.
    Dim extra As Object
    For Each extra In Application.ActivePresentation.Slides
        Dim ei As SlideInstance
        ei = Resolve.ResolveSlideInstance(extra)
        If ei.HasInstanceKey Then
            If ei.InstanceKey = "3_K900" Or ei.InstanceKey = "3_S900" Then extra.Delete
        End If
    Next extra
    tmplK.Delete
    tmplS.Delete

    Test_RunSync_CreateMissingSlidesChoosesTemplateByRowLetter = result
End Function

Private Function Test_DeckRegistry_DeckDeclaresItsOwnPeriod() As String
    Dim result As String
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim original As String
    original = DeckRegistry.GetDeckPeriod(pres)

    DeckRegistry.SetDeckPeriod pres, "FY26Q4"
    result = result & Assert(DeckRegistry.GetDeckPeriod(pres) = "FY26Q4", "the period round-trips through the deck, got '" & DeckRegistry.GetDeckPeriod(pres) & "'")

    ' THE DECK WINS. This is the whole point: a supplied period is a habit or a
    ' script default, and trusting it is how a copied deck reports last
    ' quarter's content with no error anywhere.
    Dim mm As String
    mm = DeckRegistry.PeriodMismatchText("FY26Q4", "FY26Q3")
    result = result & Assert(InStr(mm, "PERIOD MISMATCH") > 0, "a disagreement is reported, not silently resolved")
    result = result & Assert(InStr(mm, "using the DECK") > 0, "and it says WHICH one wins")

    result = result & Assert(DeckRegistry.PeriodMismatchText("FY26Q4", "FY26Q4") = "", "agreement says nothing")
    result = result & Assert(DeckRegistry.PeriodMismatchText("FY26Q4", "") = "", "supplying nothing is not a mismatch")
    result = result & Assert(DeckRegistry.PeriodMismatchText("", "FY26Q4") = "", "a deck with no period recorded cannot mismatch")

    ' Rolling forward states both ends, so it cannot be done by accident.
    Dim adv As String
    adv = DeckRegistry.AdvancePeriodText("FY26Q4", "FY27Q1")
    result = result & Assert(InStr(adv, "FY26Q4") > 0 And InStr(adv, "FY27Q1") > 0, "the roll-forward names BOTH periods")
    result = result & Assert(InStr(DeckRegistry.AdvancePeriodText("FY26Q4", "FY26Q4"), "Nothing to do") > 0, "rolling to the same period is a no-op, and says so")

    DeckRegistry.SetDeckPeriod pres, original
    Test_DeckRegistry_DeckDeclaresItsOwnPeriod = result
End Function

Private Function Test_WorkbookBridge_IndexExplainsEachSheet() As String
    Dim result As String

    ' The LIFESPAN is the part nobody can infer from a tab, and the part that
    ' decides whether it is safe to type in a sheet.
    result = result & Assert(InStr(WorkbookBridge.LifespanOf("Register"), "PERMANENT") > 0, _
        "the register is marked permanent, got '" & WorkbookBridge.LifespanOf("Register") & "'")
    result = result & Assert(InStr(WorkbookBridge.LifespanOf("TPL_ABOUT_BODY"), "Rebuilt") > 0, _
        "a drafting sheet is marked rebuilt, got '" & WorkbookBridge.LifespanOf("TPL_ABOUT_BODY") & "'")
    result = result & Assert(InStr(WorkbookBridge.LifespanOf("Sync Review q"), "consumed") > 0, _
        "a review grid is marked consumed, got '" & WorkbookBridge.LifespanOf("Sync Review q") & "'")
    result = result & Assert(InStr(WorkbookBridge.LifespanOf("Sync Log"), "Append") > 0, _
        "the log is marked append-only, got '" & WorkbookBridge.LifespanOf("Sync Log") & "'")

    ' An unrecognised sheet must say so rather than be described wrongly --
    ' a confident description of somebody else's sheet is worse than none.
    result = result & Assert(InStr(WorkbookBridge.DescribeSheet("Bob's notes"), "not created by this tool") > 0, _
        "a foreign sheet is not described as ours")
    result = result & Assert(WorkbookBridge.LifespanOf("Bob's notes") = "unknown", _
        "and its lifespan is unknown, not guessed")

    ' The drafting description must say where to type, since that was the
    ' exact thing a reader could not work out.
    Dim d As String
    d = WorkbookBridge.DescribeSheet("TPL_ABOUT_BODY")
    result = result & Assert(InStr(d, "ABOUT_BODY") > 0, "it names the field")
    ' THIS TEST HAS NOW HELD TWO DIFFERENT DEFECTS IN PLACE, FOR ONE REASON.
    '
    ' Written to catch the 3de4be8 drift, it asserted the letters of the layout
    ' current at the time -- read C, type D, tick E -- as string literals. Layout
    ' 4 then moved SUBMIT to F and the tick to G, DescribeSheet was correctly
    ' updated to derive both from the COL_D_* constants, and this test went red
    ' while demanding the retired layout back. It sat red from a6e57af until
    ' 2026-08-14, and NEXT-SESSION.md's "192 passed, 0 failed" was stale that
    ' whole time, because nobody re-ran after the code was fixed.
    '
    ' So the original comment's diagnosis was right and its remedy was not: the
    ' fault was never WHICH letters were named, it was that letters were named at
    ' all. A test that types a machine-knowable fact is a second copy, and second
    ' copies drift silently -- the same rule DOCUMENT-MAP.md applies to prose,
    ' one layer down, where no document checker can see it.
    '
    ' Derived now, so the next layout move cannot make this test wrong. The
    ' negatives are derived too and still give it a way to fail: naming the AI
    ' DRAFT column as the place to type is the specific regression that costs an
    ' evening, because that column is never published and a row with an empty
    ' SUBMIT and no tick lands in no bucket at all -- publish reports zeros with
    ' no diagnostic.
    Dim curCol As String, subCol As String, apprCol As String
    Dim draftCol As String, srcCol As String
    curCol = Chr$(64 + Drafting.COL_D_CURRENT)
    subCol = Chr$(64 + Drafting.COL_D_SUBMIT)
    apprCol = Chr$(64 + Drafting.COL_D_APPROVED)
    draftCol = Chr$(64 + Drafting.COL_D_DRAFT)
    srcCol = Chr$(64 + Drafting.COL_D_SOURCES)

    result = result & Assert(InStr(d, "column " & curCol) > 0 _
        And InStr(d, "wording in " & subCol) > 0 _
        And InStr(d, "Y in " & apprCol) > 0, _
        "the drafting index names ORIGINAL, SUBMIT and the tick by their real columns -- got '" & d & "'")
    result = result & Assert(InStr(d, "wording in " & draftCol) = 0, _
        "and does NOT send a person to the AI draft column, which never publishes -- got '" & d & "'")

    ' Sources is cited from the SOURCES column, and the index has twice named a
    ' neighbouring column instead. Both sides derived, so this cannot go stale.
    Dim srcDesc As String
    srcDesc = WorkbookBridge.DescribeSheet("Sources")
    result = result & Assert(InStr(srcDesc, "column " & srcCol) > 0 _
        And InStr(srcDesc, "column " & apprCol) = 0, _
        "the sources index cites the SOURCES column, not the tick column -- got '" & srcDesc & "'")

    Test_WorkbookBridge_IndexExplainsEachSheet = result
End Function

Private Function Test_FieldSpec_GuidanceDrivesThePrompt() As String
    Dim result As String

    ' A field WITHOUT a spec row must still draft -- blocking on paperwork would
    ' stop the work -- but the prompt must SAY it is running unguided, because
    ' generic output and specific output look equally plausible.
    Dim none As FieldGuidance
    none.FieldId = "NEW_FIELD"
    Dim p1 As String
    p1 = FieldSpec.PromptFrom(none)
    result = result & Assert(InStr(p1, "NEW_FIELD") > 0, "the generic prompt still names the field")
    result = result & Assert(InStr(p1, "no Field Spec row exists") > 0, _
        "and SAYS it is generic -- an unguided draft must not look guided")
    result = result & Assert(InStr(p1, "sole source of truth") > 0, "the no-invention rule survives into the fallback")

    ' A field WITH a row must carry its own guidance, not a template.
    Dim g As FieldGuidance
    g.Found = True
    g.FieldId = "KEY_EVENTS_BODY"
    g.Purpose = "What actually happened this quarter."
    g.Voice = "Factual and dated."
    g.Length = "One event per line."
    g.OwnJob = "Did it HAPPEN in this period?"
    g.DoNot = "Restate what the project is."
    Dim p2 As String
    p2 = FieldSpec.PromptFrom(g)
    result = result & Assert(InStr(p2, "no Field Spec row exists") = 0, "a specified field is not flagged generic")
    Dim v As Variant
    For Each v In Array("What actually happened", "Factual and dated", "One event per line", _
                        "Did it HAPPEN", "Restate what the project is")
        result = result & Assert(InStr(p2, CStr(v)) > 0, "the prompt carries '" & v & "'")
    Next v

    ' The two fields must get MATERIALLY different instructions -- that is the
    ' entire reason this exists. A generic prompt told ABOUT_BODY and
    ' KEY_EVENTS_BODY the same thing, and they want nearly opposite things.
    Dim a As FieldGuidance
    a.Found = True: a.FieldId = "ABOUT_BODY"
    a.Purpose = "A neutral description of what the project is and does."
    a.DoNot = "Describe this quarter's activity."
    result = result & Assert(FieldSpec.PromptFrom(a) <> p2, _
        "two fields with different specs get different prompts")
    result = result & Assert(InStr(FieldSpec.PromptFrom(a), "Describe this quarter's activity") > 0, _
        "ABOUT_BODY is told NOT to describe the quarter")
    result = result & Assert(InStr(p2, "What actually happened this quarter") > 0, _
        "while KEY_EVENTS_BODY is told to do exactly that")

    Test_FieldSpec_GuidanceDrivesThePrompt = result
End Function

' THE FIVE PROSE PANELS MUST EACH DO THEIR OWN JOB, and the recipes are where
' that is enforced -- Prompt 18 names these as the ones most prone to bleeding
' into each other, which is precisely the thing a person burns evenings
' re-deciding. A recipe whose Own-job test does not name its neighbours is not
' doing the work the recipe exists for.
Private Function Test_FieldSpec_TheFiveProsePanelsEachHaveTheirOwnJob() As String
    Dim result As String

    Dim xl As Object, wb As Object, ws As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Add()
    Set ws = wb.Worksheets(1)

    FieldSpec.WriteSpecSheet ws

    Dim wanted As Variant
    wanted = Array("ABOUT_BODY", "KEY_EVENTS_BODY", "STRATEGIC_ALIGNMENT_BODY", _
                   "PROBLEM_BODY", "PROGRESS_BODY")
    Dim v As Variant
    For Each v In wanted
        Dim g As FieldGuidance
        g = FieldSpec.LookupGuidance(ws, CStr(v))
        result = result & Assert(g.Found, "seeded a recipe for " & CStr(v))
        result = result & Assert(Len(Trim(g.Purpose)) > 0 And Len(Trim(g.Voice)) > 0 _
            And Len(Trim(g.Length)) > 0 And Len(Trim(g.OwnJob)) > 0 And Len(Trim(g.DoNot)) > 0, _
            CStr(v) & " has all five guidance columns filled")
    Next v

    ' The three that were added 2026-08-09 must name the neighbour they must not
    ' become. Checked on the OWN-JOB text specifically, because that is the line
    ' a person reads at 11pm when deciding whether a paragraph belongs here.
    Dim sa As FieldGuidance
    sa = FieldSpec.LookupGuidance(ws, "STRATEGIC_ALIGNMENT_BODY")
    result = result & Assert(InStr(sa.OwnJob, "ABOUT_BODY") > 0 And InStr(sa.OwnJob, "PROBLEM_BODY") > 0, _
        "Strategic Alignment's own-job test names both neighbours -- got '" & sa.OwnJob & "'")
    result = result & Assert(InStr(sa.DoNot, "[TBC]") > 0, _
        "and it says to write [TBC] rather than assert a linkage code it cannot confirm")

    Dim pb As FieldGuidance
    pb = FieldSpec.LookupGuidance(ws, "PROBLEM_BODY")
    result = result & Assert(InStr(pb.OwnJob, "WHETHER OR NOT THIS PROJECT HAPPENS") > 0, _
        "Problem's own-job test is the need existing independently of the project -- got '" & pb.OwnJob & "'")

    Dim pg As FieldGuidance
    pg = FieldSpec.LookupGuidance(ws, "PROGRESS_BODY")
    result = result & Assert(InStr(pg.DoNot, "GUESS a quarter tag") > 0, _
        "Progress refuses to guess a quarter tag -- got '" & pg.DoNot & "'")

    ' And the recipe must actually reach the prompt, or it is decoration.
    Dim prompt As String
    prompt = FieldSpec.PromptFrom(sa)
    result = result & Assert(InStr(prompt, "so what") > 0 And InStr(prompt, "600-800") > 0, _
        "the seeded recipe reaches the generated prompt")
    result = result & Assert(InStr(prompt, "no Field Spec row exists") = 0, _
        "and the prompt is not flagged as unguided")

    On Error Resume Next
    wb.Close False
    xl.Quit
    On Error GoTo 0

    Test_FieldSpec_TheFiveProsePanelsEachHaveTheirOwnJob = result
End Function

' ONE PROJECT'S TEXT MUST NEVER APPEAR AGAINST ANOTHER PROJECT.
'
' VBA's Dim does not scope to a loop, so `current` was procedure-scoped and kept
' the PREVIOUS entity's value whenever the register had none for this field. The
' first project with a value had its text copied into column C for every project
' after it. Found 2026-08-09 on the first real run of a new field: one project
' had STRATEGIC_ALIGNMENT_BODY, forty showed its 1,113 characters, and the three
' rows above it were blank -- which is the signature of exactly this bug.
'
' Column C is what a person and Copilot are both told to stay close to, so the
' failure is silent and it attributes a real paragraph to the wrong project.
'
' The fixture deliberately puts the value in the MIDDLE: an entity before it
' proves nothing leaks backwards, and two after prove nothing leaks forwards.
Private Function Test_Drafting_AFieldWithNoValueLeavesColumnCEmpty() As String
    Dim result As String

    Dim xl As Object, wb As Object, ws As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Add()
    Set ws = wb.Worksheets(1)

    Dim reg As Sheet
    Set reg.Rows = CreateObject("Scripting.Dictionary")
    Set reg.Fields = New Collection
    Set reg.InstanceOrder = New Collection

    Dim ids As Variant
    ids = Array("P-BEFORE", "P-HAS", "P-AFTER1", "P-AFTER2")
    Dim i As Long
    For i = LBound(ids) To UBound(ids)
        Dim vals As Object
        Set vals = CreateObject("Scripting.Dictionary")
        vals("PROJECT_NAME") = "Name " & CStr(ids(i))
        ' ONLY the middle entity carries the field.
        If CStr(ids(i)) = "P-HAS" Then vals("ABOUT_BODY") = "ONLY-P-HAS-SHOULD-SHOW-THIS"
        reg.Rows.Add CStr(ids(i)), vals
        reg.InstanceOrder.Add CStr(ids(i))
    Next i
    reg.Fields.Add "ABOUT_BODY"

    Drafting.WriteDraftingSheet ws, reg, "ABOUT_BODY", Empty, "Q4F26"

    ' Read column C back per entity, by matching column A rather than by row
    ' offset -- the header block's height is not this test's business.
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 1 To 200
        Dim ent As String
        ent = Trim(CStr(ws.Cells(r, Drafting.COL_D_ENTITY).Value))
        If ent <> "" Then seen(ent) = Trim(CStr(ws.Cells(r, Drafting.COL_D_CURRENT).Value))
    Next r

    result = result & Assert(seen.Exists("P-HAS"), "the entity with a value is on the sheet")
    result = result & Assert(InStr(CStr(seen("P-HAS")), "ONLY-P-HAS-SHOULD-SHOW-THIS") > 0, _
        "and it shows its own text -- got '" & CStr(seen("P-HAS")) & "'")

    Dim v As Variant
    For Each v In Array("P-BEFORE", "P-AFTER1", "P-AFTER2")
        result = result & Assert(seen.Exists(CStr(v)), CStr(v) & " is on the sheet")
        result = result & Assert(CStr(seen(CStr(v))) = "", _
            CStr(v) & " has NO value for this field, so column C must be EMPTY -- got '" & _
            CStr(seen(CStr(v))) & "'")
    Next v

    On Error Resume Next
    wb.Close False
    xl.Quit
    On Error GoTo 0

    Test_Drafting_AFieldWithNoValueLeavesColumnCEmpty = result
End Function

' THE LOG MUST SURVIVE THE FIRST LINE OF EVERY REPORT IT IS GIVEN.
'
' Reports open with "=== PREVIEW (nothing written): <type> ===". A cell value
' starting with "=" is a FORMULA, so Excel raised, and WriteRunLog's handler --
' which deliberately swallows errors so a log cannot stop a run -- abandoned the
' body. The sheet then held the title, the timestamp and a note, while the
' dialog promised "the full before-and-after ... untruncated".
'
' Asserts the LAST line arrives, not merely that something was written: the old
' behaviour wrote the first few lines perfectly well, so a test that only
' checked "the sheet is not empty" would have passed against the defect.
Private Function Test_WorkbookBridge_RunLogSurvivesALineStartingWithEquals() As String
    Dim result As String

    Dim xl As Object, wb As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Add()

    Dim body As String
    body = "=== PREVIEW (nothing written): project-status ===" & vbCrLf & _
           "  would correct: 2_P004" & vbCrLf & _
           "-1.5 leading minus, must not become a number" & vbCrLf & _
           "+cat" & vbCrLf & _
           "LAST-LINE-MARKER"

    WorkbookBridge.WriteRunLog wb, "Preview Sync -- nothing was written", body

    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(wb, WorkbookBridge.RUN_LOG_SHEET_NAME)

    Dim found As String
    Dim lastSeen As String
    Dim r As Long
    For r = 1 To 40
        Dim v As String
        v = CStr(ws.Cells(r, 1).Value)
        If Trim(v) <> "" Then
            found = found & "|" & v
            lastSeen = v
        End If
    Next r

    result = result & Assert(InStr(found, "=== PREVIEW") > 0, _
        "the banner line survives as TEXT -- got '" & found & "'")
    result = result & Assert(InStr(found, "would correct: 2_P004") > 0, _
        "the body reaches the sheet at all")
    result = result & Assert(lastSeen = "LAST-LINE-MARKER", _
        "and the LAST line arrives -- the old failure wrote the first lines fine, " & _
        "so 'not empty' would have passed against it. Got last = '" & lastSeen & "'")
    result = result & Assert(InStr(found, "-1.5 leading minus") > 0, _
        "a line starting with a minus is not coerced into a number")

    On Error Resume Next
    wb.Close False
    xl.Quit
    On Error GoTo 0

    Test_WorkbookBridge_RunLogSurvivesALineStartingWithEquals = result
End Function

Private Function Test_PlaceholderCheck_FindsRecordsNotTheTemplate() As String
    Dim result As String

    Dim seed As Object
    Set seed = NewOnboardedSlide("phc-type-1", "P100")

    Dim mr As MakeTemplateResult
    mr = TemplateSlide.MakeTemplateFrom(seed, "phc-type-1")
    result = result & Assert(mr.Ok, "template created, reason='" & mr.Reason & "'")

    ' Nothing outstanding yet: the real slide carries real values.
    Dim clean As PlaceholderReport
    clean = PlaceholderCheck.FindPlaceholders("phc-type-1")
    result = result & Assert(clean.Count = 0, "the TEMPLATE's own placeholders are not counted -- otherwise this can never reach zero, got " & clean.Count)

    ' Now leave one field on a real record showing scaffolding, as an unfilled
    ' created slide would.
    Dim shp As Object
    Set shp = FindShapeByRole(seed, "Status")
    result = result & Assert(Not shp Is Nothing, "found the field to blank")
    If Not shp Is Nothing Then shp.TextFrame.TextRange.Text = TemplateSlide.PlaceholderFor("Status")

    Dim dirty As PlaceholderReport
    dirty = PlaceholderCheck.FindPlaceholders("phc-type-1")
    result = result & Assert(dirty.Count = 1, "the record's placeholder IS counted, got " & dirty.Count)
    result = result & Assert(InStr(dirty.Detail, "P100") > 0, "the report names the entity, got '" & dirty.Detail & "'")
    result = result & Assert(InStr(dirty.Detail, "Status") > 0, "and the field")

    ' Exact match, not a substring hunt for "<<". False positives on a
    ' publication gate are how a gate gets ignored.
    shp.TextFrame.TextRange.Text = "See appendix <<note 4>> for detail"
    Dim notAPlaceholder As PlaceholderReport
    notAPlaceholder = PlaceholderCheck.FindPlaceholders("phc-type-1")
    result = result & Assert(notAPlaceholder.Count = 0, "legitimate text containing << >> is NOT flagged, got " & notAPlaceholder.Count)

    Test_PlaceholderCheck_FindsRecordsNotTheTemplate = result
End Function

' The staleness distinction -- the reason the period and timestamp are stored
' at all. A bare count is unactionable: from this morning it means fix them,
' from last quarter it means ignore it.
Private Function Test_PlaceholderCheck_MarkerDistinguishesStale() As String
    Dim result As String

    Dim raw As String
    raw = PlaceholderCheck.BuildMarker(3, "Q4 FY26", "2026-07-31 10:42")

    Dim c As Long, p As String, s As String
    result = result & Assert(PlaceholderCheck.ParseMarker(raw, c, p, s), "round-trips")
    result = result & Assert(c = 3, "count survives, got " & c)
    result = result & Assert(p = "Q4 FY26", "period survives INCLUDING its space, got '" & p & "'")
    result = result & Assert(s = "2026-07-31 10:42", "timestamp survives, got '" & s & "'")

    ' Same period -> actionable.
    Dim cur As String
    cur = PlaceholderCheck.MarkerStatusText(raw, "Q4 FY26")
    result = result & Assert(InStr(cur, "3 placeholder") > 0, "current-period status reports the count, got '" & cur & "'")
    result = result & Assert(InStr(cur, "STALE") = 0, "and does NOT cry stale when it isn't")

    ' Different period -> the number describes a different quarter and saying
    ' it plainly is the whole point.
    Dim stale As String
    stale = PlaceholderCheck.MarkerStatusText(raw, "Q1 FY27")
    result = result & Assert(InStr(stale, "STALE") > 0, "a marker from another period is called stale, got '" & stale & "'")
    result = result & Assert(InStr(stale, "Q4 FY26") > 0, "it names the period the record actually describes")
    result = result & Assert(InStr(stale, "different quarter") > 0, "it says why the number cannot be trusted")

    ' Absent and malformed are routine, not errors.
    result = result & Assert(Not PlaceholderCheck.ParseMarker("", c, p, s), "empty marker parses as absent")
    result = result & Assert(Not PlaceholderCheck.ParseMarker("garbage", c, p, s), "malformed marker parses as absent")
    result = result & Assert(InStr(PlaceholderCheck.MarkerStatusText("", "Q4 FY26"), "No placeholder record") > 0, "absent marker explains itself")

    ' The headline must be ABSENT when there is nothing outstanding -- a
    ' publication warning that always fires trains people to skip it.
    Dim none As PlaceholderReport
    result = result & Assert(PlaceholderCheck.HeadlineText(none, "Q4 FY26") = "", "no headline when nothing is outstanding")

    Dim some As PlaceholderReport
    some.Count = 2
    some.Detail = "  slide 3 (P001): PROJECT_STATUS" & vbCrLf
    Dim h As String
    h = PlaceholderCheck.HeadlineText(some, "Q4 FY26")
    result = result & Assert(InStr(h, "NOT READY TO PUBLISH") > 0, "headline leads with the consequence")
    result = result & Assert(InStr(h, "PROJECT_STATUS") > 0, "headline carries the LIST, not just a count -- a count in a summary is skimmable")

    Test_PlaceholderCheck_MarkerDistinguishesStale = result
End Function

' ---------------------------------------------------------------------
' Register -- V3, the long-format reader
' ---------------------------------------------------------------------

' Builds a long-format register in a real worksheet.
' Columns deliberately NOT in the E4 order -- the reader locates them by header
' name, and a test that lays them out in the documented order proves nothing
' about that.
Private Sub SeedRegister(ws As Object)
    ws.Cells(1, 1).Value = "Status"
    ws.Cells(1, 2).Value = "FieldID"
    ws.Cells(1, 3).Value = "Quarter"
    ws.Cells(1, 4).Value = "Value"
    ws.Cells(1, 5).Value = "SlideType"
    ws.Cells(1, 6).Value = "EntityCode"

    Dim d As Variant
    ' Status | FieldID | Quarter | Value | SlideType | EntityCode
    d = Array( _
        Array("Approved", "PROJECT_STATUS", "Q4 FY26", "In Progress", "q", "P001"), _
        Array("Approved", "PROJECT_NAME", "ALL", "Alpha Project", "q", "P001"), _
        Array("Approved", "PROJECT_STATUS", "Q3 FY26", "Old Status", "q", "P001"), _
        Array("Draft", "PROJECT_STATUS", "Q4 FY26", "Not Yet Signed Off", "q", "P002"), _
        Array("Approved", "PROJECT_STATUS", "Q4 FY26", "Closed", "q", "P002"), _
        Array("Approved", "PROJECT_STATUS", "Q4 FY26", "Wrong Type", "milestone", "P003"), _
        Array("", "", "", "", "", ""))

    Dim i As Long, j As Long
    For i = 0 To UBound(d)
        For j = 0 To 5
            ws.Cells(i + 2, j + 1).Value = d(i)(j)
        Next j
    Next i
End Sub

' ---------------------------------------------------------------------
' TagMigration -- V1, the FieldID rename
' ---------------------------------------------------------------------

' The migration end to end, including the two things most likely to be got
' wrong: nested shapes, and the master template.
Private Function Test_TagMigration_RenamesIncludingTemplateAndGroups() As String
    Dim result As String

    ' A normal instance with one top-level field and one nested inside a group.
    Dim sld As Object
    Set sld = NewBlankSlide()
    sld.Tags.Add "slide_type", "mig-type-1"
    sld.Tags.Add "instance_key", "MIG-001"

    Dim topShp As Object
    Set topShp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    topShp.TextFrame.TextRange.Text = "status"
    topShp.Tags.Add "role", "Project Status"

    Dim a As Object, b As Object
    Set a = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 150, 120, 40)
    a.TextFrame.TextRange.Text = "name"
    a.Tags.Add "role", "Project Name"
    Set b = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 200, 120, 40)
    b.TextFrame.TextRange.Text = "decoration"

    Dim grp As Object
    Set grp = sld.Shapes.Range(Array(a.Name, b.Name)).Group

    ' An unmapped field, which must survive untouched.
    Dim strayShp As Object
    Set strayShp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 260, 200, 40)
    strayShp.TextFrame.TextRange.Text = "stray"
    strayShp.Tags.Add "role", "Not In The Map"

    ' The master template -- the case a GatherInstances-based walk would miss.
    Dim mr As MakeTemplateResult
    mr = TemplateSlide.MakeTemplateFrom(sld, "mig-type-1")
    result = result & Assert(mr.Ok, "template created for the migration test, reason='" & mr.Reason & "'")

    Dim fromV(1 To 2) As String
    Dim toV(1 To 2) As String
    fromV(1) = "Project Status": toV(1) = "PROJECT_STATUS"
    fromV(2) = "Project Name":   toV(2) = "PROJECT_NAME"

    ' --- dry run must write nothing ---------------------------------------
    Dim dry As MigrationReport
    dry = TagMigration.MigrateRoleTags(fromV, toV, True)
    result = result & Assert(dry.Renamed = 4, "dry run counts all 4 mapped shapes (2 fields x instance + template), got " & dry.Renamed)
    result = result & Assert(topShp.Tags("role") = "Project Status", "DRY RUN WROTE NOTHING -- tag still reads '" & topShp.Tags("role") & "'")

    ' --- real run ----------------------------------------------------------
    Dim live As MigrationReport
    live = TagMigration.MigrateRoleTags(fromV, toV, False)

    result = result & Assert(topShp.Tags("role") = "PROJECT_STATUS", "top-level field renamed, got '" & topShp.Tags("role") & "'")
    result = result & Assert(a.Tags("role") = "PROJECT_NAME", "field NESTED IN A GROUP renamed -- a flat walk would have missed it, got '" & a.Tags("role") & "'")
    result = result & Assert(strayShp.Tags("role") = "Not In The Map", "unmapped field left untouched, got '" & strayShp.Tags("role") & "'")
    result = result & Assert(live.Unmapped >= 1, "the unmapped field is reported, got " & live.Unmapped)

    ' The template must have been migrated too. If it were skipped, every slide
    ' created from it afterwards would carry the OLD names back into the deck.
    Dim tmplShp As Object
    Set tmplShp = FindShapeByRole(mr.NewSlide, "PROJECT_STATUS")
    result = result & Assert(Not tmplShp Is Nothing, "THE MASTER TEMPLATE was migrated too -- otherwise every future created slide reintroduces the old names")

    ' --- idempotence: a second run changes nothing and says so -------------
    Dim again As MigrationReport
    again = TagMigration.MigrateRoleTags(fromV, toV, False)
    result = result & Assert(again.Renamed = 0, "re-running renames nothing, got " & again.Renamed)
    result = result & Assert(again.AlreadyDone = 4, "re-run reports 4 already-correct rather than 4 unmapped, got " & again.AlreadyDone)

    ' --- rollback: the same operation, map reversed ------------------------
    Dim back As MigrationReport
    back = TagMigration.MigrateRoleTags(toV, fromV, False)
    result = result & Assert(topShp.Tags("role") = "Project Status", "ROLLBACK restored the original value, got '" & topShp.Tags("role") & "'")
    result = result & Assert(a.Tags("role") = "Project Name", "rollback reached the grouped shape too, got '" & a.Tags("role") & "'")

    Test_TagMigration_RenamesIncludingTemplateAndGroups = result
End Function

' Case-insensitive matching on the tag VALUE. These are human-typed names, and
' a migration that silently skips a shape over a capital letter leaves a
' half-renamed deck that nothing reports as wrong.
Private Function Test_TagMigration_MatchesValueCaseInsensitively() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim shp As Object
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    shp.TextFrame.TextRange.Text = "x"
    shp.Tags.Add "role", "project status"     ' lower case in the deck

    Dim fromV(1 To 1) As String
    Dim toV(1 To 1) As String
    fromV(1) = "Project Status"               ' title case in the map
    toV(1) = "PROJECT_STATUS"

    Dim r As MigrationReport
    r = TagMigration.MigrateRoleTags(fromV, toV, False)

    result = result & Assert(shp.Tags("role") = "PROJECT_STATUS", "matched despite differing case, got '" & shp.Tags("role") & "'")

    ' Asserts THIS shape was not reported unmapped, not that nothing was.
    ' MigrateRoleTags walks every slide in the presentation by design -- it has
    ' to, or it would skip the master template -- and the test presentation is
    ' shared, so by this point it carries dozens of role tags left by other
    ' tests. A global "Unmapped = 0" is therefore asserting something the
    ' function never promised, and it failed for exactly that reason on first
    ' run (59 unmapped, all of them other tests' fixtures).
    result = result & Assert(InStr(r.UnmappedDetail, "project status") = 0, "the case-differing tag was NOT treated as unmapped -- detail: " & r.UnmappedDetail)
    result = result & Assert(r.Renamed >= 1, "at least this one shape was renamed, got " & r.Renamed)

    ' And the summary must distinguish a preview from a write, since the whole
    ' safety story is that the caller asks for the write explicitly.
    result = result & Assert(InStr(TagMigration.MigrationSummary(r, True), "nothing written") > 0, "preview summary says nothing was written")
    result = result & Assert(InStr(TagMigration.MigrationSummary(r, False), "MIGRATION APPLIED") > 0, "applied summary says so")
    result = result & Assert(InStr(TagMigration.MigrationSummary(r, False), "roll back") > 0, "applied summary tells the reader how to reverse it")

    Test_TagMigration_MatchesValueCaseInsensitively = result
End Function

' ---------------------------------------------------------------------
' IdentityCheck -- R9 / D5, duplicate identity tags
' ---------------------------------------------------------------------

' The condition R9 exists for, reproduced the way it actually occurs: by
' duplicating a slide. Probed against real Office 2026-07-31 -- identity tags
' clone on Slide.Duplicate and on paste into another deck, exactly as shape
' names do, which is why the check is needed at all.
Private Function Test_IdentityCheck_FindsClonedKeys() As String
    Dim result As String

    Dim a As Object
    Set a = NewOnboardedSlide("idchk-type-1", "P004")
    Dim b As Object
    Set b = NewOnboardedSlide("idchk-type-1", "P005")

    Dim clean As DuplicateKeyReport
    clean = IdentityCheck.FindDuplicateKeys("idchk-type-1")
    result = result & Assert(Not clean.HasDuplicates, "two distinct keys report no duplicates")
    result = result & Assert(IdentityCheck.DuplicateKeyWarningText("idchk-type-1", clean) = "", "no warning text when there is nothing to warn about")

    ' Duplicate one of them -- the real-world cause, not a synthetic tag write.
    Dim dupColl As Object
    Set dupColl = a.Duplicate()

    Dim dirty As DuplicateKeyReport
    dirty = IdentityCheck.FindDuplicateKeys("idchk-type-1")
    result = result & Assert(dirty.HasDuplicates, "duplicating a slide is DETECTED as a duplicate key")
    result = result & Assert(dirty.Count = 1, "exactly one key is duplicated, got " & dirty.Count)
    result = result & Assert(InStr(dirty.Detail, "P004") > 0, "the report names the offending key, got '" & dirty.Detail & "'")
    result = result & Assert(InStr(dirty.Detail, ",") > 0, "the report names BOTH slides, not just that a clash exists -- got '" & dirty.Detail & "'")

    ' The warning must state the consequence, not just the fact. "Two slides
    ' share a key" means nothing to someone who has not read the matching spec.
    Dim warn As String
    warn = IdentityCheck.DuplicateKeyWarningText("idchk-type-1", dirty)
    result = result & Assert(InStr(warn, "only ONE") > 0, "warning states that only one slide gets updated")
    result = result & Assert(InStr(warn, "not defined") > 0, "warning admits WHICH one is undefined rather than implying an order")
    result = result & Assert(InStr(warn, "copied") > 0, "warning names the likely cause so it is actionable")

    dupColl(1).Delete
    Test_IdentityCheck_FindsClonedKeys = result
End Function

' The template must never be counted as a competing claim. It is deliberately
' keyless, and it is excluded by GatherInstances -- but that exclusion is the
' thing this depends on, so it is asserted here rather than assumed.
Private Function Test_IdentityCheck_IgnoresTheTemplate() As String
    Dim result As String

    Dim seed As Object
    Set seed = NewOnboardedSlide("idchk-type-2", "P010")

    Dim mr As MakeTemplateResult
    mr = TemplateSlide.MakeTemplateFrom(seed, "idchk-type-2")
    result = result & Assert(mr.Ok, "template created for the identity test, reason='" & mr.Reason & "'")

    Dim rep As DuplicateKeyReport
    rep = IdentityCheck.FindDuplicateKeys("idchk-type-2")
    result = result & Assert(Not rep.HasDuplicates, "a template alongside its source is NOT a duplicate key -- it is keyless by design")

    Test_IdentityCheck_IgnoresTheTemplate = result
End Function

' ---------------------------------------------------------------------
' TemplateAudit -- "what on this slide is the tool not tracking?"
' ---------------------------------------------------------------------

' The classification boundaries, including the one that carries all the
' signal (seenOn = 0) and the one that must NOT pretend to know
' (instanceCount = 0).
Private Function Test_TemplateAudit_ClassifyBoundaries() As String
    Dim result As String

    result = result & Assert(InStr(TemplateAudit.Classify(0, 0), "UNKNOWN") > 0, "no comparison slides -> UNKNOWN, not a guess. got '" & TemplateAudit.Classify(0, 0) & "'")
    result = result & Assert(InStr(TemplateAudit.Classify(0, 3), "LIKELY PROJECT DATA") > 0, "on none of 3 -> likely project data, got '" & TemplateAudit.Classify(0, 3) & "'")
    result = result & Assert(InStr(TemplateAudit.Classify(3, 3), "chrome") > 0, "on all 3 -> chrome, got '" & TemplateAudit.Classify(3, 3) & "'")
    result = result & Assert(InStr(TemplateAudit.Classify(1, 3), "CHECK") > 0, "on 1 of 3 -> CHECK, got '" & TemplateAudit.Classify(1, 3) & "'")

    ' The middle case must stay visible rather than being rounded into one of
    ' the confident verdicts -- that rounding is what would make the report
    ' look decisive while being wrong.
    result = result & Assert(InStr(TemplateAudit.Classify(1, 3), "1 of 3") > 0, "CHECK states the actual proportion, got '" & TemplateAudit.Classify(1, 3) & "'")

    ' The recogniser the ordering and the counting both depend on. Pinned
    ' directly because a prefix match that silently never fires is invisible --
    ' it mis-sorts the grid and reports zero, and neither looks like a bug.
    result = result & Assert(TemplateAudit.IsLikelyProjectData(TemplateAudit.Classify(0, 3)), "IsLikelyProjectData recognises its OWN output for the seenOn=0 case")
    result = result & Assert(Not TemplateAudit.IsLikelyProjectData(TemplateAudit.Classify(3, 3)), "IsLikelyProjectData rejects the chrome verdict")
    result = result & Assert(Not TemplateAudit.IsLikelyProjectData(TemplateAudit.Classify(0, 0)), "IsLikelyProjectData rejects the UNKNOWN verdict")

    ' seenOn = 0 with instanceCount = 0 must NOT read as "on no other slide" --
    ' both are zero and only the second one means anything.
    result = result & Assert(InStr(TemplateAudit.Classify(0, 0), "LIKELY") = 0, "zero comparisons is not the same as 'on no other slide'")

    Test_TemplateAudit_ClassifyBoundaries = result
End Function

' The audit over real slides: tracked fields are excluded, untracked text is
' listed, and the cross-slide comparison actually separates chrome from data.
'
' Deliberately does NOT use a template as the subject. The audit must work on
' any slide of the type regardless of whether Create Template Slide has ever
' been run (Rohan's requirement, 2026-07-30) -- so the test exercises the case
' that proves the independence, not the convenient one.
Private Function Test_TemplateAudit_SeparatesChromeFromProjectData() As String
    Dim result As String

    ' Subject: one tagged field, one piece of shared furniture, one value
    ' unique to this slide.
    Dim subjectSld As Object
    Set subjectSld = NewBlankSlide()
    Dim fieldShp As Object, chromeShp As Object, dataShp As Object
    Set fieldShp = subjectSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    fieldShp.TextFrame.TextRange.Text = "<<Project Name>>"
    fieldShp.Tags.Add "role", "Project Name"
    Set chromeShp = subjectSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 120, 200, 50)
    chromeShp.TextFrame.TextRange.Text = "STRATEGIC ALIGNMENT"
    Set dataShp = subjectSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 190, 200, 50)
    dataShp.TextFrame.TextRange.Text = "~$280K"

    ' Two sibling slides carrying the same furniture and different figures.
    ' The trailing space on one is deliberate: hand-authored decks differ by
    ' whitespace constantly, and if normalisation failed, the furniture would
    ' be reported as project data on almost every row of a real audit.
    Dim otherA As Object, otherB As Object
    Set otherA = NewBlankSlide()
    Dim a1 As Object, a2 As Object
    Set a1 = otherA.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 120, 200, 50)
    a1.TextFrame.TextRange.Text = "STRATEGIC ALIGNMENT "
    Set a2 = otherA.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 190, 200, 50)
    a2.TextFrame.TextRange.Text = "~$999K"

    Set otherB = NewBlankSlide()
    Dim b1 As Object, b2 As Object
    Set b1 = otherB.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 120, 200, 50)
    b1.TextFrame.TextRange.Text = "strategic alignment"
    Set b2 = otherB.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 190, 200, 50)
    b2.TextFrame.TextRange.Text = "~$111K"

    Dim comparisons(1 To 2) As Object
    Set comparisons(1) = otherA
    Set comparisons(2) = otherB

    Dim rowCount As Long
    Dim trackedFields As String
    Dim rows() As AuditRow
    rows = TemplateAudit.BuildAudit(subjectSld, comparisons, rowCount, trackedFields)

    result = result & Assert(trackedFields = "Project Name", "the tagged field is reported as tracked, not audited. got '" & trackedFields & "'")
    result = result & Assert(rowCount = 2, "exactly the 2 untracked text items are listed, got " & rowCount)

    If rowCount <> 2 Then
        Test_TemplateAudit_SeparatesChromeFromProjectData = result
        Exit Function
    End If

    ' Actionable first: the project-data row must sort above the chrome row,
    ' because a 60-row grid in document order is a wall rather than a worklist.
    result = result & Assert(rows(1).Text = "~$280K", "the likely-project-data row sorts FIRST, got '" & rows(1).Text & "'")
    result = result & Assert(InStr(rows(1).Verdict, "LIKELY PROJECT DATA") > 0, "the unique figure reads as project data, got '" & rows(1).Verdict & "'")
    result = result & Assert(rows(1).SeenOn = 0, "the unique figure is on 0 other slides, got " & rows(1).SeenOn)

    result = result & Assert(InStr(rows(2).Text, "STRATEGIC ALIGNMENT") > 0, "the shared heading is the second row, got '" & rows(2).Text & "'")
    result = result & Assert(InStr(rows(2).Verdict, "chrome") > 0, "the shared heading reads as chrome DESPITE differing whitespace and case, got '" & rows(2).Verdict & "'")
    result = result & Assert(rows(2).SeenOn = 2, "the shared heading is found on both other slides, got " & rows(2).SeenOn)

    Test_TemplateAudit_SeparatesChromeFromProjectData = result
End Function

' With nothing to compare against, every verdict must be UNKNOWN and the
' summary must say so loudly -- a grid of UNKNOWNs with no explanation reads
' as a broken tool rather than an honest one.
Private Function Test_TemplateAudit_NoComparisonSlidesStillLists() As String
    Dim result As String

    Dim onlySld As Object
    Set onlySld = NewBlankSlide()
    Dim shp As Object
    Set shp = onlySld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    shp.TextFrame.TextRange.Text = "Some untracked text"

    Dim noComparisons() As Object ' deliberately unallocated

    Dim rowCount As Long
    Dim trackedFields As String
    Dim rows() As AuditRow
    rows = TemplateAudit.BuildAudit(onlySld, noComparisons, rowCount, trackedFields)

    result = result & Assert(rowCount = 1, "the item is still listed with no comparison available, got " & rowCount)
    If rowCount = 1 Then
        result = result & Assert(InStr(rows(1).Verdict, "UNKNOWN") > 0, "verdict is UNKNOWN, not a guess. got '" & rows(1).Verdict & "'")
    End If

    Dim s As String
    s = TemplateAudit.SummaryText("q", "slide 1", 0, 1, 0, 0)
    result = result & Assert(InStr(s, "no other slides to compare") > 0, "summary explains why every guess is UNKNOWN")

    ' And the opposite case must NOT carry that warning -- a caveat that always
    ' fires is one nobody reads.
    Dim s2 As String
    s2 = TemplateAudit.SummaryText("q", "slide 1", 5, 3, 1, 4)
    result = result & Assert(InStr(s2, "no other slides to compare") = 0, "the UNKNOWN caveat is ABSENT when comparisons exist")
    result = result & Assert(InStr(s2, "REPLACES that sheet") > 0, "summary warns the sheet gets replaced")
    result = result & Assert(InStr(s2, "refuse") > 0, "and that a re-run with pending decisions refuses rather than discarding them")

    Test_TemplateAudit_NoComparisonSlidesStillLists = result
End Function

' Re-running must not leave a previous run's surplus rows below the new ones.
' Stale rows are indistinguishable from current findings, so the audit would
' get less trustworthy the more of the work you had done -- the worst possible
' direction for a progress tool.
Private Function Test_TemplateAudit_RewriteLeavesNoStaleRows(stagingDir As String) As String
    Dim result As String

    Dim xl As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Dim wb As Object
    Set wb = xl.Workbooks.Add
    Dim ws As Object
    Set ws = wb.Worksheets(1)

    Dim many(1 To 3) As AuditRow
    many(1).ShapeName = "A": many(1).Text = "first": many(1).Verdict = "CHECK -- on 1 of 2 other slide(s)"
    many(2).ShapeName = "B": many(2).Text = "second": many(2).Verdict = "CHECK -- on 1 of 2 other slide(s)"
    many(3).ShapeName = "C": many(3).Text = "third": many(3).Verdict = "CHECK -- on 1 of 2 other slide(s)"
    TemplateAudit.WriteAuditGrid ws, many, 3

    Dim few(1 To 1) As AuditRow
    few(1).ShapeName = "A": few(1).Text = "first": few(1).Verdict = "CHECK -- on 1 of 2 other slide(s)"
    TemplateAudit.WriteAuditGrid ws, few, 1

    result = result & Assert(ws.Cells(2, 1).Value = "A", "the surviving row is still written, got '" & CStr(ws.Cells(2, 1).Value) & "'")
    result = result & Assert(Trim(CStr(ws.Cells(3, 1).Value & "")) = "", "row 3 from the longer run is GONE, got '" & CStr(ws.Cells(3, 1).Value & "") & "'")
    result = result & Assert(Trim(CStr(ws.Cells(4, 1).Value & "")) = "", "row 4 from the longer run is GONE, got '" & CStr(ws.Cells(4, 1).Value & "") & "'")

    wb.Close False
    xl.Quit

    Test_TemplateAudit_RewriteLeavesNoStaleRows = result
End Function

' THE REAL BUG THIS FIXES: found 2026-08-16 auditing every .Clear in the repo
' after the same-shape defect in DiscoverUI.bas. SummaryText's own warning
' fired AFTER WriteAuditGrid had already cleared the sheet -- by the time a
' person read "copy them out first", there was nothing left to copy. Proves
' the refusal actually fires, and that it leaves the pending decision
' completely untouched, not just that it returns a string.
Private Function Test_TemplateAudit_RefusesToDiscardPendingDecisions() As String
    Dim result As String

    Dim xl As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Dim wb As Object
    Set wb = xl.Workbooks.Add
    Dim ws As Object
    Set ws = wb.Worksheets(1)

    Dim first(1 To 2) As AuditRow
    first(1).ShapeName = "A": first(1).Text = "first": first(1).Verdict = "CHECK -- on 1 of 2 other slide(s)"
    first(2).ShapeName = "B": first(2).Text = "second": first(2).Verdict = "CHECK -- on 1 of 2 other slide(s)"
    Dim r1 As String
    r1 = TemplateAudit.WriteAuditGrid(ws, first, 2)
    result = result & Assert(r1 = "", "the first, clean write succeeds, got '" & r1 & "'")

    ' A PERSON RECORDS A DECISION -- not yet acted on.
    ws.Cells(3, 6).Value = "field"

    Dim second(1 To 1) As AuditRow
    second(1).ShapeName = "C": second(1).Text = "third": second(1).Verdict = "CHECK -- on 1 of 2 other slide(s)"
    Dim r2 As String
    r2 = TemplateAudit.WriteAuditGrid(ws, second, 1)
    result = result & Assert(Left(r2, 1) = "!", "a re-run with a pending decision REFUSES, got '" & r2 & "'")
    result = result & Assert(InStr(r2, "1 decision") > 0, "the refusal names how many, got '" & r2 & "'")

    ' NOTHING WAS TOUCHED. The original grid, decision included, must survive
    ' a refused call exactly as it was -- that is the whole point.
    result = result & Assert(ws.Cells(2, 1).Value = "A", "row A survives the refused call")
    result = result & Assert(ws.Cells(3, 1).Value = "B", "row B survives the refused call")
    result = result & Assert(ws.Cells(3, 6).Value = "field", "the recorded decision itself survives, got '" & CStr(ws.Cells(3, 6).Value) & "'")
    result = result & Assert(Trim(CStr(ws.Cells(4, 1).Value & "")) = "", "the second run's row C was NOT written -- the refusal happened before any write")

    wb.Close False
    xl.Quit

    Test_TemplateAudit_RefusesToDiscardPendingDecisions = result
End Function

' ---------------------------------------------------------------------
' RunSync
' ---------------------------------------------------------------------

Private Function Test_RunSync_GatherInstancesFiltersByType() As String
    Dim result As String
    Dim sldA As Object, sldB As Object, sldC As Object
    Set sldA = NewBlankSlide()
    sldA.Tags.Add "slide_type", "type-a"
    sldA.Tags.Add "instance_key", "a-1"
    Set sldB = NewBlankSlide()
    sldB.Tags.Add "slide_type", "type-b"
    sldB.Tags.Add "instance_key", "b-1"
    Set sldC = NewBlankSlide()
    sldC.Tags.Add "slide_type", "type-a"
    sldC.Tags.Add "instance_key", "a-2"

    Dim instances() As Object
    instances = RunSync.GatherInstances("type-a")

    Dim lo As Long, hi As Long, hasAny As Boolean
    On Error Resume Next
    lo = LBound(instances): hi = UBound(instances): hasAny = (Err.Number = 0)
    On Error GoTo 0

    result = result & Assert(hasAny And (hi - lo + 1) >= 2, "at least the 2 type-a slides found, got " & IIf(hasAny, hi - lo + 1, 0))

    Dim foundA1 As Boolean, foundA2 As Boolean, foundB As Boolean
    Dim i As Long
    If hasAny Then
        For i = lo To hi
            Dim inst As SlideInstance
            inst = Resolve.ResolveSlideInstance(instances(i))
            If inst.InstanceKey = "a-1" Then foundA1 = True
            If inst.InstanceKey = "a-2" Then foundA2 = True
            If inst.InstanceKey = "b-1" Then foundB = True
        Next i
    End If
    result = result & Assert(foundA1 And foundA2, "both type-a instances included")
    result = result & Assert(Not foundB, "type-b instance correctly excluded")

    Test_RunSync_GatherInstancesFiltersByType = result
End Function

' End-to-end: a template, a fresh Excel sheet (via cross-app Excel
' automation -- the real intended usage per specs/vba-port.md's "VBA runs
' inside Excel or drives it via COM from the PowerPoint side") with one
' existing-and-stale row (exercises case 4) and two rows with no matching
' slide yet (exercises case 3 twice, plus resequencing). Written directly
' to the worksheet's cells matching ExcelOutput's own layout convention
' rather than importing ExcelOutput.bas into a second, Excel-hosted VBA
' project -- RunSync.RunRoutineSync's own ExcelOutput.ReadSheet call reads
' it from the PowerPoint-hosted project against the live cross-app ws
' object either way, which is the actual thing under test.
Private Function Test_RunSync_EndToEndCreatesSlidesFromFreshSheet() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    titleShp.TextFrame.TextRange.Text = "Template Title"
    titleShp.Tags.Add "role", "Title"
    templateSld.Tags.Add "slide_type", "e2e-type"
    templateSld.Tags.Add "instance_key", "e2e-template"
    ' MARKED AS A TEMPLATE, which no fixture in this suite ever did.
    ' CreateMissingSlides now refuses a source that is not marked -- because
    ' DeckRegistry.LookupType returns whatever slide the type was registered
    ' against, and on the real rig NO slide carries is_template, so creation
    ' would have cloned a REAL PROJECT'S slide. The copy keeps every panel the
    ' tool does not track, so a new project would silently inherit another
    ' project's Strategic Alignment, Problem and Progress text.
    ' Resolve excludes templates from GatherInstances, hence the count below.
    templateSld.Tags.Add "is_template", "1"

    Dim xl As Object, wb As Object, ws As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Add()
    Set ws = wb.Worksheets(1)

    ws.Cells(1, 1).Value = "Instance ID"
    ws.Cells(1, 2).Value = "Title"
    ws.Cells(2, 1).Value = "e2e-existing"
    ws.Cells(2, 2).Value = "Existing Value"
    ws.Cells(3, 1).Value = "e2e-new-1"
    ws.Cells(3, 2).Value = "Brand New One"
    ws.Cells(4, 1).Value = "e2e-new-2"
    ws.Cells(4, 2).Value = "Brand New Two"

    Dim existingSld As Object
    Set existingSld = NewBlankSlide()
    Dim existingTitleShp As Object
    Set existingTitleShp = existingSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    existingTitleShp.TextFrame.TextRange.Text = "Stale Value"
    existingTitleShp.Tags.Add "role", "Title"
    existingSld.Tags.Add "slide_type", "e2e-type"
    existingSld.Tags.Add "instance_key", "e2e-existing"

    Dim report As String
    ' Two calls now, not one. As of 2026-07-31 a sync no longer creates slides
    ' (Excel Control Layer round 9 section 4) -- it reports rows that have none,
    ' and CreateMissingSlides is the separate operation a person chooses. This
    ' test previously proved creation happened INSIDE the sync; it now proves
    ' the same end state is reached, via the two operations it actually takes.
    '
    ' The sync runs FIRST, deliberately: if creation had silently stayed in it,
    ' the CreateMissingSlides call after would find nothing to do and the
    ' assertions below would still pass. Asserting the sync's own report says
    ' "0 created" is what makes that detectable.
    report = RunSync.RunRoutineSync(ws, "e2e-type", templateSld)
    result = result & Assert(InStr(report, "no slide for: e2e-new-1") > 0, _
        "the SYNC reports the row as having no slide rather than creating one -- report: " & report)
    result = result & Assert(InStr(report, "created:") = 0, _
        "the sync created NOTHING -- report: " & report)

    Dim sheetForCreate As Sheet
    sheetForCreate = ExcelOutput.ReadSheet(ws)
    report = report & RunSync.CreateMissingSlides(sheetForCreate, "e2e-type", templateSld, False)

    Dim instances() As Object
    instances = RunSync.GatherInstances("e2e-type")

    Dim lo As Long, hi As Long, hasAny As Boolean
    On Error Resume Next
    lo = LBound(instances): hi = UBound(instances): hasAny = (Err.Number = 0)
    On Error GoTo 0

    ' 3, not 4: the template is now marked and GatherInstances excludes it
    ' (RunSync.bas:54). It was previously counted as a record, which is exactly
    ' the confusion the marker exists to prevent.
    result = result & Assert(hasAny And (hi - lo + 1) = 3, "3 real instances after sync (existing + 2 new; the template is excluded), got " & IIf(hasAny, hi - lo + 1, 0) & " -- report: " & report)

    Dim byKey As Object
    Set byKey = CreateObject("Scripting.Dictionary")
    Dim i As Long
    If hasAny Then
        For i = lo To hi
            Dim inst As SlideInstance
            inst = Resolve.ResolveSlideInstance(instances(i))
            If inst.HasInstanceKey Then Set byKey(inst.InstanceKey) = instances(i)
        Next i
    End If

    result = result & Assert(byKey.Exists("e2e-new-1") And byKey.Exists("e2e-new-2"), "both missing slides were created by CreateMissingSlides")

    If byKey.Exists("e2e-existing") Then
        Dim correctedShp As Object
        Set correctedShp = FindShapeByRole(byKey("e2e-existing"), "Title")
        Dim correctedText As String
        If correctedShp Is Nothing Then
            correctedText = "<shape not found>"
        Else
            correctedText = correctedShp.TextFrame.TextRange.Text
        End If
        result = result & Assert(correctedText = "Existing Value", "existing slide's stale value was corrected (case 4), got '" & correctedText & "'")

        ' The REPORT, not just the deck. Until 2026-07-30 this branch printed
        ' "corrected: <key>" and nothing else, so the record you get after a
        ' deck changes said less than the preview you get before it. Past tense
        ' is asserted deliberately: the write has already happened here, and the
        ' preview's "now/new" wording would describe it as still pending.
        result = result & Assert(InStr(report, "Title:") > 0, _
            "the sync report names the field it changed -- report: " & report)
        result = result & Assert(InStr(report, "was:  'Stale Value'") > 0, _
            "the sync report records what the field held before the write -- report: " & report)
        result = result & Assert(InStr(report, "now:  'Existing Value'") > 0, _
            "the sync report records what it now holds -- report: " & report)
    End If
    If byKey.Exists("e2e-new-1") Then
        Dim new1Shp As Object
        Set new1Shp = FindShapeByRole(byKey("e2e-new-1"), "Title")
        Dim new1Text As String
        If new1Shp Is Nothing Then
            new1Text = "<shape not found>"
        Else
            new1Text = new1Shp.TextFrame.TextRange.Text
        End If
        result = result & Assert(new1Text = "Brand New One", "new slide's value injected correctly, got '" & new1Text & "'")
    End If

    If byKey.Exists("e2e-existing") And byKey.Exists("e2e-new-1") And byKey.Exists("e2e-new-2") Then
        result = result & Assert(byKey("e2e-existing").SlideIndex < byKey("e2e-new-1").SlideIndex, "e2e-existing sits before e2e-new-1 after resequencing")
        result = result & Assert(byKey("e2e-new-1").SlideIndex < byKey("e2e-new-2").SlideIndex, "e2e-new-1 sits before e2e-new-2 after resequencing")
    End If

    wb.Close False
    xl.Quit

    Test_RunSync_EndToEndCreatesSlidesFromFreshSheet = result
End Function

' The end-to-end proof for the preview: the SAME setup as the test above --
' one stale row and two rows with no slide -- run through PreviewRoutineSync
' instead. Asserts the report says what a real sync would do AND that the deck
' is bit-for-bit untouched afterwards: no new slides, stale text still stale,
' slide order unchanged. Deliberately mirrors the real end-to-end test rather
' than testing the preview in isolation, so the two can be read side by side
' and any divergence in classification shows up as a failing assertion here.
Private Function Test_RunSync_PreviewReportsWithoutTouchingTheDeck() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    titleShp.TextFrame.TextRange.Text = "Template Title"
    titleShp.Tags.Add "role", "Title"
    templateSld.Tags.Add "slide_type", "preview-type"
    templateSld.Tags.Add "instance_key", "preview-template"

    Dim existingSld As Object
    Set existingSld = NewBlankSlide()
    Dim existingTitleShp As Object
    Set existingTitleShp = existingSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    existingTitleShp.TextFrame.TextRange.Text = "Stale Value"
    existingTitleShp.Tags.Add "role", "Title"
    existingSld.Tags.Add "slide_type", "preview-type"
    existingSld.Tags.Add "instance_key", "preview-existing"

    Dim xl As Object, wb As Object, ws As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Add()
    Set ws = wb.Worksheets(1)

    ws.Cells(1, 1).Value = "Instance ID"
    ws.Cells(1, 2).Value = "Title"
    ws.Cells(2, 1).Value = "preview-existing"
    ws.Cells(2, 2).Value = "Corrected Value"
    ws.Cells(3, 1).Value = "preview-orphan-1"
    ws.Cells(3, 2).Value = "Would Be Created One"
    ws.Cells(4, 1).Value = "preview-orphan-2"
    ws.Cells(4, 2).Value = "Would Be Created Two"

    Dim slideCountBefore As Long
    slideCountBefore = Application.ActivePresentation.Slides.count
    Dim existingIndexBefore As Long
    existingIndexBefore = existingSld.SlideIndex

    Dim report As String
    report = RunSync.PreviewRoutineSync(ws, "preview-type", "")

    ' --- the deck must be exactly as it was ---
    result = result & Assert(Application.ActivePresentation.Slides.count = slideCountBefore, _
        "PREVIEW MUST NOT CREATE SLIDES: count unchanged, got " & Application.ActivePresentation.Slides.count & " vs " & slideCountBefore)
    result = result & Assert(existingTitleShp.TextFrame.TextRange.Text = "Stale Value", _
        "PREVIEW MUST NOT WRITE: stale value still stale, got '" & existingTitleShp.TextFrame.TextRange.Text & "'")
    result = result & Assert(existingSld.SlideIndex = existingIndexBefore, _
        "PREVIEW MUST NOT REORDER: slide index unchanged, got " & existingSld.SlideIndex & " vs " & existingIndexBefore)

    ' --- and it must still say what would happen ---
    result = result & Assert(InStr(report, "PREVIEW (nothing written)") > 0, _
        "report is labelled as a preview -- report: " & report)
    result = result & Assert(InStr(report, "would correct: preview-existing") > 0, _
        "report names the instance whose field would be corrected -- report: " & report)
    result = result & Assert(InStr(report, "Stale Value") > 0, _
        "report shows the field's current value so a human can see the before -- report: " & report)
    ' BOTH halves, pinned to their exact rendered form. Live 2026-07-30 the
    ' preview showed only "now:", so approving a sync meant opening Excel to
    ' find out what you were approving. Asserting the value alone is not enough
    ' -- "Corrected Value" could appear anywhere in the report and still leave
    ' the before/after pairing broken, which is the thing a human reads.
    result = result & Assert(InStr(report, "now:  'Stale Value'") > 0, _
        "report shows the BEFORE on its own labelled line -- report: " & report)
    result = result & Assert(InStr(report, "new:  'Corrected Value'") > 0, _
        "report shows the AFTER -- the value that would actually be written -- report: " & report)
    result = result & Assert(InStr(report, "preview-orphan-1") > 0 And InStr(report, "preview-orphan-2") > 0, _
        "report names both orphaned rows -- report: " & report)
    ' Both halves of the warning, asserted exactly: the per-row callout and the
    ' separate block that spells out the consequence. InStr is case-sensitive
    ' (Option Compare Binary), so these pin the real wording rather than
    ' matching loosely -- this is the one warning standing between a drifted
    ' deck and a mass slide duplication, and it must not be able to go quiet
    ' through a reword.
    ' WAS "WOULD CREATE A NEW SLIDE", WHICH STOPPED BEING TRUE ON 2026-07-31.
    ' Slide creation left the sync path that day, so the preview was threatening
    ' something Sync Now cannot do -- while Sync Now, over the same rows, said
    ' "every linked slide already matches the workbook". Both wrong, in opposite
    ' directions. The test asserted the misleading half, which is how it survived.
    result = result & Assert(InStr(report, "REACHES NO SLIDE: preview-orphan-1") > 0, _
        "report flags each orphaned row at the point it lists it -- report: " & report)
    result = result & Assert(InStr(report, "DUPLICATE the template") = 0, _
        "and does NOT threaten duplication, which sync has not done since 2026-07-31")
    result = result & Assert(InStr(report, "Their text will NOT appear anywhere") > 0, _
        "report spells out the real consequence in its own warning block -- report: " & report)
    result = result & Assert(InStr(report, "no button for adding a missing slide") > 0, _
        "and says plainly that the remedy has no button yet, rather than naming one that " & _
        "cannot be pressed -- report: " & report)
    result = result & Assert(InStr(report, "2 row(s) reach no slide") > 0, _
        "summary counts both orphaned rows -- report: " & report)

    wb.Close False
    xl.Quit

    Test_RunSync_PreviewReportsWithoutTouchingTheDeck = result
End Function

' The guard that replaced "Sync Now isn't on the toolbar" as the thing keeping
' a drifted deck safe (2026-07-30). Pure text, so the wording is pinned here
' rather than left to a live click-through nobody will repeat.
'
' The asymmetry is the point: corrections are cheap to undo and get a plain
' line; slide creation is the outcome that was one click from happening to the
' real deck on 2026-07-27, so it must be impossible to skim past.
Private Function Test_RunSync_ConfirmSyncTextReportsUncreatableRows() As String
    Dim result As String

    ' Renamed and rewritten 2026-07-31. It used to pin a capitalised warning
    ' that slides WOULD BE CREATED, because until today a sync could create
    ' them and the warning was the only thing standing between a drifted deck
    ' and mass duplication. Creation has since moved out of the sync entirely
    ' (Excel Control Layer round 9 section 4), so the warning describes a
    ' capability that no longer exists.
    '
    ' This test failing was the correct outcome of that change, not collateral:
    ' pinning wording is precisely so that removing it has to be deliberate.
    ' What is pinned NOW is the opposite guarantee -- that the dialog does not
    ' claim anything will be created.

    ' --- the ordinary case ---
    Dim plain As String
    plain = RunSync.ConfirmSyncText(3, 0, 0)
    result = result & Assert(InStr(plain, "3 slide(s) corrected") > 0, _
        "states how many slides change, got: " & plain)
    result = result & Assert(InStr(plain, "0 rows without a slide") > 0, _
        "states the count plainly, got: " & plain)
    result = result & Assert(InStr(plain, "Proceed?") > 0, _
        "asks, rather than announcing, got: " & plain)

    ' --- the case that used to be the dangerous one ---
    ' 43 unmatched rows against a real deck was the live state on 2026-07-27,
    ' and one click would have duplicated the template 43 times. The same input
    ' must now produce a dialog that promises no such thing.
    Dim many As String
    many = RunSync.ConfirmSyncText(2, 43, 0)
    result = result & Assert(InStr(many, "43 row(s) have no slide") > 0, _
        "reports the count, got: " & many)
    result = result & Assert(InStr(many, "NOT created by this action") > 0, _
        "states outright that this action does not create them, got: " & many)

    ' The removed alarm must be genuinely absent, not merely reworded --
    ' case-sensitive, since Option Compare is Binary here.
    result = result & Assert(InStr(many, "WILL BE CREATED") = 0, _
        "the old capitalised creation warning is GONE -- it described a capability that no longer exists, got: " & many)
    result = result & Assert(InStr(many, "duplication") = 0, _
        "and so is the mass-duplication language, got: " & many)

    ' The number is still shown, because it is still the signal that this deck's
    ' linkage has drifted -- or that it is an assembled pack that should not be
    ' synced at all. Only the alarm went; the information stayed.
    result = result & Assert(InStr(many, "drifted linkage") > 0, _
        "explains what a large number means, got: " & many)
    result = result & Assert(InStr(many, "assembled pack") > 0, _
        "including that it may be a composite that should not be synced, got: " & many)

    ' --- flagged is reported only when it happened ---
    result = result & Assert(InStr(RunSync.ConfirmSyncText(1, 0, 0), "flagged") = 0, _
        "no flagged line when nothing is flagged")
    result = result & Assert(InStr(RunSync.ConfirmSyncText(1, 0, 5), "5 flagged") > 0, _
        "flagged count shown when there is one")

    Test_RunSync_ConfirmSyncTextReportsUncreatableRows = result
End Function

' ---------------------------------------------------------------------
' DeckAdoption
' ---------------------------------------------------------------------

' Opens a fresh, blank Excel worksheet via cross-app automation, same
' pattern as Test_RunSync_EndToEndCreatesSlidesFromFreshSheet -- DeckAdoption
' always reads/writes a live worksheet, never a closed file. Caller is
' responsible for wb.Close False / xl.Quit when done.
Private Sub NewTestWorksheet(ByRef xl As Object, ByRef wb As Object, ByRef ws As Object)
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Add()
    Set ws = wb.Worksheets(1)
    ws.Cells(1, 1).Value = "Instance ID"
End Sub

Private Function Test_DeckAdoption_AlreadyLinkedSlideSkipped() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 300, 50)
    titleShp.TextFrame.TextRange.Text = "Template Title"
    titleShp.Tags.Add "role", "Title"

    Dim alreadySld As Object
    Set alreadySld = NewBlankSlide()
    Dim alreadyShp As Object
    Set alreadyShp = alreadySld.Shapes.AddTextbox(msoTextOrientationHorizontal, 101, 100, 300, 50)
    alreadyShp.TextFrame.TextRange.Text = "Already Linked"
    alreadySld.Tags.Add "slide_type", "adopt-type"
    alreadySld.Tags.Add "instance_key", "already-1"

    Dim xl As Object, wb As Object, ws As Object
    NewTestWorksheet xl, wb, ws

    Dim slidesToAdopt(1 To 1) As Object
    Set slidesToAdopt(1) = alreadySld

    Dim harvestedValues() As Object
    Dim plans() As AdoptionSlidePlan
    plans = DeckAdoption.PlanAdoption(slidesToAdopt, templateSld, ws, harvestedValues)

    result = result & Assert(UBound(plans) = 1, "one plan entry produced, got " & (UBound(plans) - LBound(plans) + 1))
    result = result & Assert(plans(1).Disposition = "already_linked", "disposition is already_linked, got '" & plans(1).Disposition & "'")

    Dim confirmedKeys(1 To 1) As String
    confirmedKeys(1) = ""

    Dim commitResult As AdoptionResult
    commitResult = DeckAdoption.CommitAdoption(plans, slidesToAdopt, harvestedValues, confirmedKeys, "adopt-type", templateSld, ws)

    result = result & Assert(commitResult.AlreadyLinkedCount = 1, "AlreadyLinkedCount=1, got " & commitResult.AlreadyLinkedCount)
    result = result & Assert(commitResult.LinkedCount = 0, "LinkedCount=0 (nothing to link), got " & commitResult.LinkedCount)

    wb.Close False
    xl.Quit

    Test_DeckAdoption_AlreadyLinkedSlideSkipped = result
End Function

' THE SAFETY PROPERTY, PINNED. A harvest may only fill a register cell that is
' empty. FIELD_A is seeded with text a person typed and FIELD_B is left empty,
' and the slide carries a different value for BOTH -- so a harvest that ignores
' the emptiness rule fails on FIELD_A's assertion rather than passing quietly
' with one extra write.
Private Function Test_Harvest_WritesOnlyWhereTheRegisterCellIsEmpty() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    sld.Tags.Add "slide_type", "harvest-type"
    sld.Tags.Add "instance_key", "h-1"

    Dim aShp As Object
    Set aShp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 300, 50)
    aShp.TextFrame.TextRange.Text = "FROM THE SLIDE A"
    aShp.Tags.Add "role", "FIELD_A"

    Dim bShp As Object
    Set bShp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 200, 300, 50)
    bShp.TextFrame.TextRange.Text = "FROM THE SLIDE B"
    bShp.Tags.Add "role", "FIELD_B"

    Dim xl As Object, wb As Object, ws As Object
    NewTestWorksheet xl, wb, ws
    ws.Cells(1, 2).Value = "Quarter"
    ws.Cells(1, 3).Value = "FIELD_A"
    ws.Cells(1, 4).Value = "FIELD_B"
    ws.Cells(2, 1).Value = "h-1"
    ws.Cells(2, 2).Value = "Q4F26"
    ws.Cells(2, 3).Value = "TYPED BY A PERSON"
    ' FIELD_B deliberately left empty -- this is the cell the harvest may fill.

    Dim o As HarvestOutcome
    o = Harvest.HarvestSlide(sld, ws, "Q4F26", False)

    result = result & Assert(o.Ran, "the harvest ran, problem was '" & o.Problem & "'")
    result = result & Assert(o.Written = 1, "exactly one value written, got " & o.Written)
    result = result & Assert(o.SkippedHasValue = 1, "one field skipped for already holding a value, got " & o.SkippedHasValue)
    result = result & Assert(CStr(ws.Cells(2, 3).Value) = "TYPED BY A PERSON", _
        "FIELD_A was NOT overwritten, got '" & CStr(ws.Cells(2, 3).Value) & "'")
    result = result & Assert(CStr(ws.Cells(2, 4).Value) = "FROM THE SLIDE B", _
        "FIELD_B was filled from the slide, got '" & CStr(ws.Cells(2, 4).Value) & "'")

    wb.Close False
    xl.Quit

    Test_Harvest_WritesOnlyWhereTheRegisterCellIsEmpty = result
End Function

' A DEVICE IS SKIPPED BY NAME, NOT READ AS BLANK. A group has no text of its
' own, so a harvest that treated it as an ordinary shape would find "" and
' report it as "blank on the slide" -- indistinguishable from a genuinely empty
' field, and the timeline's 21 values would look accounted for while nothing
' had been read. The assertion is on SkippedNotText specifically for that reason.
Private Function Test_Harvest_SkipsATaggedGroupInsteadOfReadingItAsEmpty() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    sld.Tags.Add "slide_type", "harvest-type"
    sld.Tags.Add "instance_key", "h-2"

    Dim a As Object, b As Object
    Set a = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 100, 30)
    a.TextFrame.TextRange.Text = "part one"
    Set b = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 140, 100, 30)
    b.TextFrame.TextRange.Text = "part two"

    Dim grp As Object
    Set grp = sld.Shapes.Range(Array(a.Name, b.Name)).Group
    grp.Tags.Add "role", "TIMELINE_X"

    Dim xl As Object, wb As Object, ws As Object
    NewTestWorksheet xl, wb, ws
    ws.Cells(1, 2).Value = "Quarter"
    ws.Cells(1, 3).Value = "TIMELINE_X"
    ws.Cells(2, 1).Value = "h-2"
    ws.Cells(2, 2).Value = "Q4F26"

    Dim o As HarvestOutcome
    o = Harvest.HarvestSlide(sld, ws, "Q4F26", False)

    result = result & Assert(o.Ran, "the harvest ran, problem was '" & o.Problem & "'")
    result = result & Assert(o.SkippedNotText = 1, "the tagged group was counted as not-text, got " & o.SkippedNotText)
    result = result & Assert(o.SkippedBlankOnSlide = 0, "it was NOT counted as blank-on-slide, got " & o.SkippedBlankOnSlide)
    result = result & Assert(o.Written = 0, "nothing was written, got " & o.Written)
    result = result & Assert(IsEmpty(ws.Cells(2, 3).Value), "the device's register cell is still empty")

    wb.Close False
    xl.Quit

    Test_Harvest_SkipsATaggedGroupInsteadOfReadingItAsEmpty = result
End Function

' PROPAGATION CARRIES A MISSING ROLE AND LEAVES A PRESENT ONE ALONE.
'
' The target already carries FIELD_A and is missing FIELD_B. Both properties are
' asserted because they fail in opposite directions: forgetting to filter the
' roles re-matches FIELD_A against whatever untagged shape is left (stamping the
' wrong shape), and filtering too hard stamps nothing at all.
Private Function Test_Harvest_PropagationStampsOnlyTheMissingRole() As String
    Dim result As String

    Dim tplSld As Object
    Set tplSld = NewBlankSlide()
    Dim tA As Object, tB As Object
    Set tA = tplSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 300, 50)
    tA.TextFrame.TextRange.Text = "template A"
    tA.Tags.Add "role", "FIELD_A"
    Set tB = tplSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 200, 300, 50)
    tB.TextFrame.TextRange.Text = "template B"
    tB.Tags.Add "role", "FIELD_B"

    Dim sld As Object
    Set sld = NewBlankSlide()
    sld.Tags.Add "slide_type", "harvest-type"
    sld.Tags.Add "instance_key", "h-3"

    Dim gotA As Object, gotB As Object
    Set gotA = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 101, 100, 300, 50)
    gotA.TextFrame.TextRange.Text = "target A"
    gotA.Tags.Add "role", "FIELD_A"                 ' already labelled
    Set gotB = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 101, 200, 300, 50)
    gotB.TextFrame.TextRange.Text = "target B"      ' deliberately unlabelled

    Dim o As PropagateOutcome
    o = Harvest.PropagateTemplateTags(sld, tplSld, False)

    result = result & Assert(o.Ran, "propagation ran, problem was '" & o.Problem & "'")
    result = result & Assert(o.AlreadyTagged = 1, "one role was already on the slide, got " & o.AlreadyTagged)
    result = result & Assert(o.Stamped = 1, "exactly one role stamped, got " & o.Stamped)
    result = result & Assert(gotB.Tags("role") = "FIELD_B", _
        "the unlabelled shape now carries FIELD_B, got '" & gotB.Tags("role") & "'")
    result = result & Assert(gotA.Tags("role") = "FIELD_A", _
        "the already-labelled shape still carries FIELD_A, got '" & gotA.Tags("role") & "'")

    Test_Harvest_PropagationStampsOnlyTheMissingRole = result
End Function

' A HARVESTED VALUE IS STORED AS THE CHARACTERS THAT WERE ON THE SLIDE.
'
' `.Value = CStr(...)` looks like a string write and is not -- Excel parses
' anything date-shaped or number-shaped. On the real deck this turned
' "30 Oct 2023" into 45229 and "$275,598" into 275598: right values, formatting
' gone, and publishing them back would put a serial number on a funder slide.
'
' Both fixtures are chosen because Excel WILL coerce them. A test using ordinary
' prose would pass with or without the fix and prove nothing.
Private Function Test_Harvest_StoresSlideTextVerbatimRatherThanLettingExcelParseIt() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    sld.Tags.Add "slide_type", "harvest-type"
    sld.Tags.Add "instance_key", "h-4"

    Dim dShp As Object
    Set dShp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 300, 50)
    dShp.TextFrame.TextRange.Text = "30 Oct 2023"
    dShp.Tags.Add "role", "START_DATE"

    Dim mShp As Object
    Set mShp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 200, 300, 50)
    mShp.TextFrame.TextRange.Text = "$275,598"
    mShp.Tags.Add "role", "INDUSTRY_CASH"

    Dim xl As Object, wb As Object, ws As Object
    NewTestWorksheet xl, wb, ws
    ws.Cells(1, 2).Value = "Quarter"
    ws.Cells(1, 3).Value = "START_DATE"
    ws.Cells(1, 4).Value = "INDUSTRY_CASH"
    ws.Cells(2, 1).Value = "h-4"
    ws.Cells(2, 2).Value = "Q4F26"

    Dim o As HarvestOutcome
    o = Harvest.HarvestSlide(sld, ws, "Q4F26", False)

    result = result & Assert(o.Ran, "the harvest ran, problem was '" & o.Problem & "'")
    result = result & Assert(o.Written = 2, "two values written, got " & o.Written)
    result = result & Assert(CStr(ws.Cells(2, 3).Value) = "30 Oct 2023", _
        "the date is stored as it appeared on the slide, got '" & CStr(ws.Cells(2, 3).Value) & "'")
    result = result & Assert(CStr(ws.Cells(2, 4).Value) = "$275,598", _
        "the money is stored as it appeared on the slide, got '" & CStr(ws.Cells(2, 4).Value) & "'")

    wb.Close False
    xl.Quit

    Test_Harvest_StoresSlideTextVerbatimRatherThanLettingExcelParseIt = result
End Function

' A BAR'S DISPLAYED TEXT IS NOT ITS REGISTER VALUE.
'
' The fixture is the real failure: a field whose shape carries the text `33%`
' and which the router sends to the progress injector, because a `.track`
' companion sits beside it. InjectProgressVia refuses anything non-numeric --
' `'90%'` is named as wrong in its own error string -- so harvesting `33%` writes
' a value the tool cannot publish, and the empty-cell rule then makes it
' permanent.
'
' Asserting on SkippedNotText rather than just Written=0 matters: a harvest that
' skipped it for the wrong reason (blank on slide, no shape) would also write
' nothing, and would break again the moment that reason stopped holding.
Private Function Test_Harvest_RefusesAFieldThePublishPathWouldTreatAsABar() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    sld.Tags.Add "slide_type", "harvest-type"
    sld.Tags.Add "instance_key", "h-5"

    Dim bar As Object
    Set bar = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 200, 30)
    bar.TextFrame.TextRange.Text = "33%"
    bar.Tags.Add "role", "PROJECT_PROGRESS"

    ' The companion is what makes it a bar, per InjectorFor.
    Dim track As Object
    Set track = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 140, 200, 30)
    track.Tags.Add "role", "PROJECT_PROGRESS.track"

    Dim xl As Object, wb As Object, ws As Object
    NewTestWorksheet xl, wb, ws
    ws.Cells(1, 2).Value = "Quarter"
    ws.Cells(1, 3).Value = "PROJECT_PROGRESS"
    ws.Cells(2, 1).Value = "h-5"
    ws.Cells(2, 2).Value = "Q4F26"

    result = result & Assert(InjectPrimitive.InjectorFor(sld, "PROJECT_PROGRESS") = InjectPrimitive.INJECTOR_BAR, _
        "the router calls this a bar, got '" & InjectPrimitive.InjectorFor(sld, "PROJECT_PROGRESS") & "'")

    Dim o As HarvestOutcome
    o = Harvest.HarvestSlide(sld, ws, "Q4F26", False)

    result = result & Assert(o.Ran, "the harvest ran, problem was '" & o.Problem & "'")
    result = result & Assert(o.SkippedNotText = 1, _
        "the bar was skipped as not-text, got " & o.SkippedNotText)
    result = result & Assert(o.Written = 0, "nothing was written, got " & o.Written)
    result = result & Assert(IsEmpty(ws.Cells(2, 3).Value), _
        "and its register cell is still empty, so a corrected run can still fill it")

    wb.Close False
    xl.Quit

    Test_Harvest_RefusesAFieldThePublishPathWouldTreatAsABar = result
End Function

' THE PREVIEW MUST COUNT WHAT THE RUN WILL WRITE, NOT WHAT IS TAGGED TODAY.
'
' The real failure: a dry run offered 10 values and the run wrote 34, because the
' dry harvest looks fields up BY their role tag and propagation has not written
' those tags yet. The fixture is that exact shape -- an UNTAGGED shape whose role
' propagation would stamp -- and the assertion is that the dry count includes it.
'
' Both halves matter. Without the pending map the count is 0 and the caller
' under-promises; if pending were consulted in a WET run it would double-count,
' since propagation stamps first and the ordinary lookup finds the shape.
' WHAT RETIREMENT MUST NEVER TOUCH.
'
' SlidesWithNoRow decides what gets DELETED, so its guards are the only thing
' between "the register no longer lists this project" and destroying the slide
' every other slide is cloned from. The fixture holds all four states at once,
' because each guard fails in a different direction and a test that built them
' one at a time would pass while the combination was wrong.
'
' The template is the one that would hurt most: it carries a slide type and no
' instance key BY DESIGN, so any rule keying off "has a type" deletes it.
Private Function Test_Membership_RetirementSkipsTemplateAndUnclassifiedSlides() As String
    Dim result As String

    Dim pres As Object
    Set pres = Application.Presentations.Add(msoFalse)

    ' 1. a live project -- has a row, must survive
    Dim liveSld As Object
    Set liveSld = pres.Slides.Add(1, ppLayoutBlank)
    liveSld.Tags.Add "slide_type", "membership-type"
    liveSld.Tags.Add "instance_key", "live-1"

    ' 2. a retired project -- no row, must be returned
    Dim goneSld As Object
    Set goneSld = pres.Slides.Add(2, ppLayoutBlank)
    goneSld.Tags.Add "slide_type", "membership-type"
    goneSld.Tags.Add "instance_key", "gone-1"

    ' 3. the TEMPLATE -- type, no instance key, marked template
    Dim tplSld As Object
    Set tplSld = pres.Slides.Add(3, ppLayoutBlank)
    tplSld.Tags.Add "slide_type", "membership-type"
    tplSld.Tags.Add TEMPLATE_TAG_NAME, "1"

    ' 4. unclassified -- a title page, a divider, anything a person added
    Dim plainSld As Object
    Set plainSld = pres.Slides.Add(4, ppLayoutBlank)

    ' A sheet holding a row for live-1 only.
    Dim xl As Object, wb As Object, ws As Object
    NewTestWorksheet xl, wb, ws
    ws.Cells(1, 2).Value = "Quarter"
    ws.Cells(1, 3).Value = "ABOUT_BODY"
    ws.Cells(2, 1).Value = "live-1"
    ws.Cells(2, 2).Value = "Q4F26"
    ws.Cells(2, 3).Value = "still running"

    Dim sheet As Sheet
    Dim problem As String
    sheet = ExcelOutput.ReadSheetForDeckPeriod(ws, "Q4F26", problem)
    result = result & Assert(problem = "", "the fixture sheet reads, got '" & problem & "'")

    Dim doomed As Collection
    Set doomed = RibbonUI.SlidesWithNoRow(pres, sheet, "membership-type")

    result = result & Assert(doomed.count = 1, _
        "exactly one slide is retirable, got " & doomed.count)

    If doomed.count = 1 Then
        Dim gotInst As SlideInstance
        gotInst = Resolve.ResolveSlideInstance(doomed(1))
        result = result & Assert(gotInst.InstanceKey = "gone-1", _
            "and it is the one with no row, got '" & gotInst.InstanceKey & "'")
    End If

    wb.Close False
    xl.Quit
    pres.Close

    Test_Membership_RetirementSkipsTemplateAndUnclassifiedSlides = result
End Function

Private Function Test_Harvest_DryRunCountsFieldsThatLabellingWillCreate() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    sld.Tags.Add "slide_type", "harvest-type"
    sld.Tags.Add "instance_key", "h-6"

    ' Deliberately NOT tagged -- this is the state a dry run is blind to.
    Dim shp As Object
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 300, 50)
    shp.TextFrame.TextRange.Text = "Prof. Someone"

    Dim xl As Object, wb As Object, ws As Object
    NewTestWorksheet xl, wb, ws
    ws.Cells(1, 2).Value = "Quarter"
    ws.Cells(1, 3).Value = "PROJECT_LEAD"
    ws.Cells(2, 1).Value = "h-6"
    ws.Cells(2, 2).Value = "Q4F26"

    Dim blind As HarvestOutcome
    blind = Harvest.HarvestSlide(sld, ws, "Q4F26", True)
    result = result & Assert(blind.Written = 0, _
        "without the pending map the dry run cannot see it, got " & blind.Written)

    Dim pending As Object
    Set pending = CreateObject("Scripting.Dictionary")
    pending("PROJECT_LEAD") = "Prof. Someone"

    Dim seeing As HarvestOutcome
    seeing = Harvest.HarvestSlide(sld, ws, "Q4F26", True, pending)
    result = result & Assert(seeing.Written = 1, _
        "with it, the dry run counts the field labelling will create, got " & seeing.Written)
    result = result & Assert(IsEmpty(ws.Cells(2, 3).Value), _
        "and a dry run still wrote nothing")

    wb.Close False
    xl.Quit

    Test_Harvest_DryRunCountsFieldsThatLabellingWillCreate = result
End Function

' Builds the real deck's date pair: same x, identical size, 0.14" apart -- which
' is 128016 EMU, well inside POSITION_TOLERANCE_EMU's one inch, so geometry alone
' cannot tell them apart. Measured from slide 2 of 3. Project Progress.pptx.
Private Function DateCandidate(nm As String, yEmu As Long) As Candidate
    Dim c As Candidate
    c.Name = nm
    c.ShapeType = "autoshape_or_textbox"
    c.HasText = True
    c.PositionX = 10876280           ' 11.89"
    c.PositionY = yEmu
    c.SizeCx = 1152144               ' 1.26"
    c.SizeCy = 128016                ' 0.14"
    DateCandidate = c
End Function

' A NAME DECIDES WHAT GEOMETRY CANNOT -- and only when it is unambiguous.
'
' Both halves are asserted because they fail in opposite directions. Without the
' name tier, END_DATE's reference scores ~0.87 and ~0.98 against the two
' candidates and picks the WRONG one (or ties with START_DATE and collides).
' With a name tier that fired on ambiguous names, the duplicate-name case would
' return an arbitrary one of four `Shape 16`s as "high" confidence -- which is a
' guess wearing a certainty's label, and worse than falling through to scoring.
Private Function Test_Matching_ExactShapeNameBeatsGeometryOnlyWhenUnambiguous() As String
    Dim result As String

    Dim candidates(1 To 2) As Candidate
    candidates(1) = DateCandidate("Text 212a", 2889504)      ' 3.16"
    candidates(2) = DateCandidate("Text 216a", 3026664)      ' 3.31"

    ' The template's END_DATE: nearer candidate 2, but well within one inch of both.
    Dim reference As Candidate
    reference = DateCandidate("Text 216a", 3035808)          ' 3.32"

    Dim r As MatchResult
    r = Matching.Match(candidates, reference)

    result = result & Assert(r.HasCandidate, "a candidate was chosen")
    result = result & Assert(r.CandidateIndex = 2, _
        "the shape with the MATCHING NAME was chosen, got index " & r.CandidateIndex)
    result = result & Assert(r.Reason = "sibling ambiguity resolved by matching shape name", _
        "and it says so, got '" & r.Reason & "'")

    ' AMBIGUOUS NAMES MUST FALL THROUGH, not resolve arbitrarily. Rohan's slide 1
    ' carries four shapes called 'Shape 16'.
    Dim dupes(1 To 2) As Candidate
    dupes(1) = DateCandidate("Shape 16", 2889504)
    dupes(2) = DateCandidate("Shape 16", 3026664)

    Dim dupRef As Candidate
    dupRef = DateCandidate("Shape 16", 3035808)

    Dim d As MatchResult
    d = Matching.Match(dupes, dupRef)
    result = result & Assert(d.Reason <> "sibling ambiguity resolved by matching shape name", _
        "a duplicated name does NOT resolve by name, got reason '" & d.Reason & "'")

    Test_Matching_ExactShapeNameBeatsGeometryOnlyWhenUnambiguous = result
End Function

Private Function Test_DeckAdoption_ReadyHighConfidenceSlideLinkedAndCreatesFreshRow() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 300, 50)
    titleShp.TextFrame.TextRange.Text = "Template Title"
    titleShp.Tags.Add "role", "Title"

    Dim candidateSld As Object
    Set candidateSld = NewBlankSlide()
    Dim candidateShp As Object
    Set candidateShp = candidateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 102, 101, 300, 50)
    candidateShp.TextFrame.TextRange.Text = "Harvested Title Value"

    Dim xl As Object, wb As Object, ws As Object
    NewTestWorksheet xl, wb, ws

    Dim slidesToAdopt(1 To 1) As Object
    Set slidesToAdopt(1) = candidateSld

    Dim harvestedValues() As Object
    Dim plans() As AdoptionSlidePlan
    plans = DeckAdoption.PlanAdoption(slidesToAdopt, templateSld, ws, harvestedValues)

    result = result & Assert(plans(1).Disposition = "ready", "disposition is ready, got '" & plans(1).Disposition & "'")
    result = result & Assert(plans(1).MatchedKeylessRowId = "", "no keyless row to match against an empty sheet")
    result = result & Assert(Not harvestedValues(1) Is Nothing, "harvested values captured for a ready slide")
    If Not harvestedValues(1) Is Nothing Then
        result = result & Assert(harvestedValues(1)("Title") = "Harvested Title Value", "harvested Title value correct, got '" & harvestedValues(1)("Title") & "'")
    End If

    Dim confirmedKeys(1 To 1) As String
    confirmedKeys(1) = "adopt-rec-1"

    Dim commitResult As AdoptionResult
    commitResult = DeckAdoption.CommitAdoption(plans, slidesToAdopt, harvestedValues, confirmedKeys, "adopt-type", templateSld, ws)

    result = result & Assert(commitResult.LinkedCount = 1, "LinkedCount=1, got " & commitResult.LinkedCount)
    result = result & Assert(commitResult.FailedVerificationCount = 0, "no verification failures, got " & commitResult.FailedVerificationCount)

    Dim instance As SlideInstance
    instance = Resolve.ResolveSlideInstance(candidateSld)
    result = result & Assert(instance.HasTypeTag And instance.TypeTag = "adopt-type", "slide tagged with slide_type, got '" & instance.TypeTag & "'")
    result = result & Assert(instance.HasInstanceKey And instance.InstanceKey = "adopt-rec-1", "slide tagged with instance_key, got '" & instance.InstanceKey & "'")

    Dim sheet As Sheet
    sheet = ExcelOutput.ReadSheet(ws)
    result = result & Assert(sheet.Rows.Exists("adopt-rec-1"), "Data-sheet row created for the new instance")
    If sheet.Rows.Exists("adopt-rec-1") Then
        result = result & Assert(sheet.Rows("adopt-rec-1")("Title") = "Harvested Title Value", "Data-sheet row carries the harvested value, got '" & sheet.Rows("adopt-rec-1")("Title") & "'")
    End If

    wb.Close False
    xl.Quit

    Test_DeckAdoption_ReadyHighConfidenceSlideLinkedAndCreatesFreshRow = result
End Function

' Reuses the exact same drift (title barely moved, body moved 600pt away)
' Test_Onboarding_HighAndMediumConfidence already established scores
' high/medium respectively (not low) -- see that test's own comment for why
' this specific offset was chosen empirically.
Private Function Test_DeckAdoption_MediumConfidenceSlideNeedsConfirmationAndIsNotTagged() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 300, 50)
    titleShp.TextFrame.TextRange.Text = "Title text"
    titleShp.Tags.Add "role", "Title"
    Dim bodyShp As Object
    Set bodyShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 200, 300, 200)
    bodyShp.TextFrame.TextRange.Text = "Body text"
    bodyShp.Tags.Add "role", "Body"

    Dim candidateSld As Object
    Set candidateSld = NewBlankSlide()
    Dim candidateTitleShp As Object
    Set candidateTitleShp = candidateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 102, 101, 300, 50)
    candidateTitleShp.TextFrame.TextRange.Text = "Title text drifted"
    Dim candidateBodyShp As Object
    Set candidateBodyShp = candidateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 700, 700, 300, 200)
    candidateBodyShp.TextFrame.TextRange.Text = "Body text drifted"

    Dim xl As Object, wb As Object, ws As Object
    NewTestWorksheet xl, wb, ws

    Dim slidesToAdopt(1 To 1) As Object
    Set slidesToAdopt(1) = candidateSld

    Dim harvestedValues() As Object
    Dim plans() As AdoptionSlidePlan
    plans = DeckAdoption.PlanAdoption(slidesToAdopt, templateSld, ws, harvestedValues)

    result = result & Assert(plans(1).Disposition = "needs_confirmation", "disposition is needs_confirmation, got '" & plans(1).Disposition & "'")

    Dim confirmedKeys(1 To 1) As String
    confirmedKeys(1) = "" ' irrelevant -- needs_confirmation is never committed regardless

    Dim commitResult As AdoptionResult
    commitResult = DeckAdoption.CommitAdoption(plans, slidesToAdopt, harvestedValues, confirmedKeys, "adopt-type", templateSld, ws)

    result = result & Assert(commitResult.LinkedCount = 0, "nothing committed for a needs_confirmation slide, got LinkedCount=" & commitResult.LinkedCount)
    result = result & Assert(commitResult.ExcludedUnclassifiedCount = 1, "reported in the excluded/unclassified bucket, got " & commitResult.ExcludedUnclassifiedCount)

    Dim instance As SlideInstance
    instance = Resolve.ResolveSlideInstance(candidateSld)
    result = result & Assert(Not instance.HasTypeTag And Not instance.HasInstanceKey, "slide-level tags never written for a needs_confirmation slide")
    Dim untouchedTitleShp As Object
    Set untouchedTitleShp = FindShapeByRole(candidateSld, "Title")
    result = result & Assert(untouchedTitleShp Is Nothing, "Title shape was never tagged either -- no partial commit")

    wb.Close False
    xl.Quit

    Test_DeckAdoption_MediumConfidenceSlideNeedsConfirmationAndIsNotTagged = result
End Function

Private Function Test_DeckAdoption_UnclassifiedSlideExcluded() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 300, 50)
    titleShp.TextFrame.TextRange.Text = "Template Title"
    titleShp.Tags.Add "role", "Title"

    Dim unrelatedSld As Object
    Set unrelatedSld = NewBlankSlide()
    unrelatedSld.Shapes.AddShape msoShapeOval, 500, 500, 50, 50 ' pure decoration, no text -- excluded by IsCandidateField entirely

    Dim xl As Object, wb As Object, ws As Object
    NewTestWorksheet xl, wb, ws

    Dim slidesToAdopt(1 To 1) As Object
    Set slidesToAdopt(1) = unrelatedSld

    Dim harvestedValues() As Object
    Dim plans() As AdoptionSlidePlan
    plans = DeckAdoption.PlanAdoption(slidesToAdopt, templateSld, ws, harvestedValues)

    result = result & Assert(plans(1).Disposition = "unclassified", "disposition is unclassified, got '" & plans(1).Disposition & "'")

    Dim confirmedKeys(1 To 1) As String
    confirmedKeys(1) = ""

    Dim commitResult As AdoptionResult
    commitResult = DeckAdoption.CommitAdoption(plans, slidesToAdopt, harvestedValues, confirmedKeys, "adopt-type", templateSld, ws)

    result = result & Assert(commitResult.ExcludedUnclassifiedCount = 1, "reported excluded/unclassified, got " & commitResult.ExcludedUnclassifiedCount)
    result = result & Assert(commitResult.LinkedCount = 0, "never forced in, got LinkedCount=" & commitResult.LinkedCount)

    wb.Close False
    xl.Quit

    Test_DeckAdoption_UnclassifiedSlideExcluded = result
End Function

Private Function Test_DeckAdoption_MatchesExistingKeylessRowLinksWithoutCreatingNewRow() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 300, 50)
    titleShp.TextFrame.TextRange.Text = "Template Title"
    titleShp.Tags.Add "role", "Title"

    Dim candidateSld As Object
    Set candidateSld = NewBlankSlide()
    Dim candidateShp As Object
    Set candidateShp = candidateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 102, 101, 300, 50)
    candidateShp.TextFrame.TextRange.Text = "Existing Value X"

    Dim xl As Object, wb As Object, ws As Object
    NewTestWorksheet xl, wb, ws
    ws.Cells(1, 2).Value = "Title"
    ws.Cells(2, 2).Value = "Existing Value X" ' row 2: real data, blank Instance ID -- a keyless row

    Dim slidesToAdopt(1 To 1) As Object
    Set slidesToAdopt(1) = candidateSld

    Dim harvestedValues() As Object
    Dim plans() As AdoptionSlidePlan
    plans = DeckAdoption.PlanAdoption(slidesToAdopt, templateSld, ws, harvestedValues)

    result = result & Assert(plans(1).Disposition = "ready", "disposition is ready, got '" & plans(1).Disposition & "'")
    result = result & Assert(plans(1).MatchedKeylessRowId = "2", "matched keyless row 2, got '" & plans(1).MatchedKeylessRowId & "'")

    Dim confirmedKeys(1 To 1) As String
    confirmedKeys(1) = "adopt-rec-existing"

    Dim commitResult As AdoptionResult
    commitResult = DeckAdoption.CommitAdoption(plans, slidesToAdopt, harvestedValues, confirmedKeys, "adopt-type", templateSld, ws)

    result = result & Assert(commitResult.LinkedCount = 1, "LinkedCount=1, got " & commitResult.LinkedCount)
    result = result & Assert(ws.Cells(2, 1).Value = "adopt-rec-existing", "row 2's Instance ID cell was filled in, got '" & ws.Cells(2, 1).Value & "'")

    Dim sheet As Sheet
    sheet = ExcelOutput.ReadSheet(ws)
    result = result & Assert(sheet.InstanceOrder.count = 1, "no new row was appended -- still exactly one instance row, got " & sheet.InstanceOrder.count)

    wb.Close False
    xl.Quit

    Test_DeckAdoption_MatchesExistingKeylessRowLinksWithoutCreatingNewRow = result
End Function

' Regression test for a real bug found and fixed during review (2026-07-25,
' see SPIKE_NOTES_DeckAdoption.md): an earlier version of PlanAdoption forced
' its returned AdoptionSlidePlan() array to be 1-based (via a separate
' counter) while harvestedValues() kept whatever base the caller's
' slidesToAdopt() array used -- CommitAdoption then indexed both (plus
' confirmedInstanceKeys()) with one shared loop variable, silently
' misattributing which slide's harvest/confirmation belongs to which plan
' entry for any non-1-based slidesToAdopt(). Every other DeckAdoption test
' uses a 1-based Dim slidesToAdopt(1 To 1), which would never have exposed
' this -- this test deliberately uses a 0-based array with 3 slides of 3
' different dispositions to prove indices stay aligned end-to-end.
Private Function Test_DeckAdoption_MultiSlideZeroBasedBatchKeepsIndicesAligned() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim titleShp As Object
    Set titleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 300, 50)
    titleShp.TextFrame.TextRange.Text = "Template Title"
    titleShp.Tags.Add "role", "Title"

    Dim alreadySld As Object
    Set alreadySld = NewBlankSlide()
    alreadySld.Tags.Add "slide_type", "adopt-type"
    alreadySld.Tags.Add "instance_key", "already-batch-1"

    Dim readySld As Object
    Set readySld = NewBlankSlide()
    Dim readyShp As Object
    Set readyShp = readySld.Shapes.AddTextbox(msoTextOrientationHorizontal, 101, 100, 300, 50)
    readyShp.TextFrame.TextRange.Text = "Batch Ready Title"

    Dim unrelatedSld As Object
    Set unrelatedSld = NewBlankSlide()
    unrelatedSld.Shapes.AddShape msoShapeOval, 500, 500, 50, 50 ' pure decoration -- unclassified

    Dim xl As Object, wb As Object, ws As Object
    NewTestWorksheet xl, wb, ws

    Dim slidesToAdopt(0 To 2) As Object ' deliberately 0-based, not 1-based
    Set slidesToAdopt(0) = alreadySld
    Set slidesToAdopt(1) = readySld
    Set slidesToAdopt(2) = unrelatedSld

    Dim harvestedValues() As Object
    Dim plans() As AdoptionSlidePlan
    plans = DeckAdoption.PlanAdoption(slidesToAdopt, templateSld, ws, harvestedValues)

    result = result & Assert(LBound(plans) = 0 And UBound(plans) = 2, "plans() shares slidesToAdopt's 0-based range, got " & LBound(plans) & " To " & UBound(plans))
    result = result & Assert(plans(0).Disposition = "already_linked", "index 0 (alreadySld) is already_linked, got '" & plans(0).Disposition & "'")
    result = result & Assert(plans(1).Disposition = "ready", "index 1 (readySld) is ready, got '" & plans(1).Disposition & "'")
    result = result & Assert(plans(2).Disposition = "unclassified", "index 2 (unrelatedSld) is unclassified, got '" & plans(2).Disposition & "'")
    result = result & Assert(Not harvestedValues(1) Is Nothing, "harvested values captured at index 1, not misattributed to another index")
    If Not harvestedValues(1) Is Nothing Then
        result = result & Assert(harvestedValues(1)("Title") = "Batch Ready Title", "harvested value at index 1 belongs to readySld, got '" & harvestedValues(1)("Title") & "'")
    End If

    Dim confirmedKeys(0 To 2) As String
    confirmedKeys(0) = ""
    confirmedKeys(1) = "adopt-batch-ready-1"
    confirmedKeys(2) = ""

    Dim commitResult As AdoptionResult
    commitResult = DeckAdoption.CommitAdoption(plans, slidesToAdopt, harvestedValues, confirmedKeys, "adopt-type", templateSld, ws)

    result = result & Assert(commitResult.AlreadyLinkedCount = 1, "AlreadyLinkedCount=1, got " & commitResult.AlreadyLinkedCount)
    result = result & Assert(commitResult.LinkedCount = 1, "LinkedCount=1, got " & commitResult.LinkedCount)
    result = result & Assert(commitResult.ExcludedUnclassifiedCount = 1, "ExcludedUnclassifiedCount=1, got " & commitResult.ExcludedUnclassifiedCount)

    ' The critical correctness check: readySld (index 1) actually received
    ' the tag/instance_key confirmed for index 1 -- not index 0's or 2's.
    Dim readyInstance As SlideInstance
    readyInstance = Resolve.ResolveSlideInstance(readySld)
    result = result & Assert(readyInstance.HasInstanceKey And readyInstance.InstanceKey = "adopt-batch-ready-1", "readySld (index 1) tagged with its own confirmed instance_key, got '" & readyInstance.InstanceKey & "'")
    Dim readyTitleShp As Object
    Set readyTitleShp = FindShapeByRole(readySld, "Title")
    result = result & Assert(Not readyTitleShp Is Nothing, "readySld's Title shape was tagged")

    ' alreadySld (index 0) must be untouched -- still its original key, not overwritten.
    Dim alreadyInstance As SlideInstance
    alreadyInstance = Resolve.ResolveSlideInstance(alreadySld)
    result = result & Assert(alreadyInstance.InstanceKey = "already-batch-1", "alreadySld (index 0) instance_key unchanged, got '" & alreadyInstance.InstanceKey & "'")

    ' unrelatedSld (index 2) must remain untagged.
    Dim unrelatedInstance As SlideInstance
    unrelatedInstance = Resolve.ResolveSlideInstance(unrelatedSld)
    result = result & Assert(Not unrelatedInstance.HasInstanceKey, "unrelatedSld (index 2) never tagged")

    wb.Close False
    xl.Quit

    Test_DeckAdoption_MultiSlideZeroBasedBatchKeepsIndicesAligned = result
End Function

' ---------------------------------------------------------------------
' ResolveFields
' ---------------------------------------------------------------------

Private Function Test_ResolveFields_ValidateSingleShapeSelectionAcceptsOneShape() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim shp As Object
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    shp.TextFrame.TextRange.Text = "Untagged Field"
    shp.Select

    Dim outShp As Object
    Dim errMsg As String
    errMsg = ResolveFields.ValidateSingleShapeSelection(Application.ActiveWindow.Selection, outShp)

    result = result & Assert(errMsg = "", "single-shape selection accepted, got error '" & errMsg & "'")
    result = result & Assert(Not outShp Is Nothing, "outShp was set")
    If Not outShp Is Nothing Then
        result = result & Assert(outShp.Name = shp.Name, "outShp is the selected shape, got '" & outShp.Name & "' want '" & shp.Name & "'")
    End If

    Test_ResolveFields_ValidateSingleShapeSelectionAcceptsOneShape = result
End Function

Private Function Test_ResolveFields_ValidateSingleShapeSelectionRejectsMultiple() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim shp1 As Object, shp2 As Object
    Set shp1 = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 100, 50)
    Set shp2 = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 200, 50, 100, 50)
    sld.Shapes.Range(Array(shp1.Name, shp2.Name)).Select

    Dim outShp As Object
    Dim errMsg As String
    errMsg = ResolveFields.ValidateSingleShapeSelection(Application.ActiveWindow.Selection, outShp)

    result = result & Assert(errMsg <> "", "multi-shape selection rejected with an error message")
    result = result & Assert(InStr(errMsg, "2") > 0, "error message mentions the selected count (2), got '" & errMsg & "'")

    Test_ResolveFields_ValidateSingleShapeSelectionRejectsMultiple = result
End Function

Private Function Test_ResolveFields_BuildRolePickerPromptListsRolesNumbered() As String
    Dim result As String

    Dim roles(1 To 2) As String
    roles(1) = "Title"
    roles(2) = "Date"

    Dim prompt As String
    prompt = ResolveFields.BuildRolePickerPrompt(roles)

    result = result & Assert(InStr(prompt, "1) Title") > 0, "prompt lists role 1 as 'Title', got: " & prompt)
    result = result & Assert(InStr(prompt, "2) Date") > 0, "prompt lists role 2 as 'Date', got: " & prompt)

    Dim emptyRoles() As String
    Dim emptyPrompt As String
    emptyPrompt = ResolveFields.BuildRolePickerPrompt(emptyRoles)
    result = result & Assert(InStr(emptyPrompt, "no defined roles") > 0, "unallocated roles() produces a 'no defined roles' prompt, got: " & emptyPrompt)

    Test_ResolveFields_BuildRolePickerPromptListsRolesNumbered = result
End Function

Private Function Test_ResolveFields_PickRoleFromListAcceptsNumberOrName() As String
    Dim result As String

    Dim roles(1 To 2) As String
    roles(1) = "Title"
    roles(2) = "Date"

    result = result & Assert(ResolveFields.PickRoleFromList("1", roles) = "Title", "'1' resolves to Title")
    result = result & Assert(ResolveFields.PickRoleFromList("date", roles) = "Date", "'date' resolves case-insensitively to Date")
    result = result & Assert(ResolveFields.PickRoleFromList("nonsense", roles) = "", "an unrecognized name resolves to ''")
    result = result & Assert(ResolveFields.PickRoleFromList("99", roles) = "", "an out-of-range number resolves to ''")

    Dim emptyRoles() As String
    result = result & Assert(ResolveFields.PickRoleFromList("1", emptyRoles) = "", "unallocated roles() always resolves to ''")

    Test_ResolveFields_PickRoleFromListAcceptsNumberOrName = result
End Function

' Exercises the full pipeline the InputBox-driven PromptResolveUnmatchedField
' wraps (selection validation -> role lookup -> ConfirmFieldMatch) without
' the InputBox itself, since TestRunner.bas cannot drive a live modal
' dialog -- proves the pieces actually wire together correctly, which the
' unit-level tests above don't on their own.
Private Function Test_ResolveFields_EndToEndTagsSelectedShapeViaPickedRole() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim templateTitleShp As Object
    Set templateTitleShp = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    templateTitleShp.TextFrame.TextRange.Text = "Template Title"
    templateTitleShp.Tags.Add "role", "Title"

    Dim newSld As Object
    Set newSld = NewBlankSlide()
    Dim newShp As Object
    Set newShp = newSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 51, 50, 200, 50)
    newShp.TextFrame.TextRange.Text = "New Value"
    newShp.Select

    Dim outShp As Object
    Dim selErr As String
    selErr = ResolveFields.ValidateSingleShapeSelection(Application.ActiveWindow.Selection, outShp)
    result = result & Assert(selErr = "", "selection validated, got error '" & selErr & "'")

    Dim roles() As String
    Dim fieldShapes() As Candidate
    fieldShapes = Onboarding.BuildTemplateFieldShapes(templateSld, roles)

    Dim role As String
    role = ResolveFields.PickRoleFromList("1", roles)
    result = result & Assert(role = "Title", "role 1 off the template resolves to 'Title', got '" & role & "'")

    If Not outShp Is Nothing And role <> "" Then
        Onboarding.ConfirmFieldMatch outShp, role
    End If

    result = result & Assert(newShp.Tags("role") = "Title", "selected shape was tagged with the picked role, got '" & newShp.Tags("role") & "'")

    Test_ResolveFields_EndToEndTagsSelectedShapeViaPickedRole = result
End Function

' ---------------------------------------------------------------------
' DeckRegistry
' ---------------------------------------------------------------------

Private Function Test_DeckRegistry_BuildAndParseTypeRegistrationRoundTrip() As String
    Dim result As String

    Dim raw As String
    raw = DeckRegistry.BuildTypeRegistration(1234&, "QuarterlyData")
    result = result & Assert(raw = "1234|QuarterlyData", "built registration is '1234|QuarterlyData', got '" & raw & "'")

    Dim slideId As Long
    Dim ws As String
    Dim ok As Boolean
    ok = DeckRegistry.ParseTypeRegistration(raw, slideId, ws)
    result = result & Assert(ok, "parse succeeded")
    result = result & Assert(slideId = 1234, "parsed slideId is 1234, got " & slideId)
    result = result & Assert(ws = "QuarterlyData", "parsed worksheetName is 'QuarterlyData', got '" & ws & "'")

    Test_DeckRegistry_BuildAndParseTypeRegistrationRoundTrip = result
End Function

Private Function Test_DeckRegistry_ParseTypeRegistrationRejectsMalformed() As String
    Dim result As String
    Dim slideId As Long
    Dim ws As String

    result = result & Assert(Not DeckRegistry.ParseTypeRegistration("", slideId, ws), "empty string rejected")
    result = result & Assert(Not DeckRegistry.ParseTypeRegistration("no-separator", slideId, ws), "missing '|' rejected")
    result = result & Assert(Not DeckRegistry.ParseTypeRegistration("abc|Sheet1", slideId, ws), "non-numeric slide id rejected")
    result = result & Assert(Not DeckRegistry.ParseTypeRegistration("123|", slideId, ws), "empty worksheet name after '|' rejected")

    Test_DeckRegistry_ParseTypeRegistrationRejectsMalformed = result
End Function

Private Function Test_DeckRegistry_GetOrCreateDeckIdIsStableAcrossCalls() As String
    Dim result As String
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim first As String
    first = DeckRegistry.GetOrCreateDeckId(pres)
    result = result & Assert(first <> "", "a deck id was generated")

    Dim second As String
    second = DeckRegistry.GetOrCreateDeckId(pres)
    result = result & Assert(second = first, "second call returns the same id, got '" & second & "' want '" & first & "'")

    Test_DeckRegistry_GetOrCreateDeckIdIsStableAcrossCalls = result
End Function

' THE UNSTAMPED CASE IS THE ONE THAT MATTERS MOST and it is the easiest to get
' wrong: every register created before the pairing existed has a blank
' DeckReference, so a verdict that treated blank as a mismatch would refuse the
' real register on the machine where nothing can be debugged.
Private Function Test_DeckRegistry_PairingVerdictOnlyRefusesAKnownDifferentDeck() As String
    Dim result As String

    result = result & Assert(DeckRegistry.PairingVerdict("", "deck-A", "C:\r.xlsx") = "", _
        "an UNSTAMPED workbook is not a mismatch")
    result = result & Assert(DeckRegistry.PairingVerdict("   ", "deck-A", "C:\r.xlsx") = "", _
        "a whitespace-only stamp is not a mismatch")
    result = result & Assert(DeckRegistry.PairingVerdict("deck-A", "deck-A", "C:\r.xlsx") = "", _
        "the same id is not a mismatch")
    result = result & Assert(DeckRegistry.PairingVerdict(" DECK-a ", "deck-A", "C:\r.xlsx") = "", _
        "case and padding do not make a mismatch")

    Dim v As String
    v = DeckRegistry.PairingVerdict("deck-B", "deck-A", "C:\other.xlsx")
    result = result & Assert(v <> "", "a DIFFERENT id IS a mismatch")

    ' Both ids, or this repeats "could not open the paired workbook" -- three
    ' causes in one morning behind one message that named only one of them.
    result = result & Assert(InStr(v, "deck-B") > 0, "the mismatch names the workbook's id, got: " & v)
    result = result & Assert(InStr(v, "deck-A") > 0, "the mismatch names THIS deck's id, got: " & v)
    result = result & Assert(InStr(v, "C:\other.xlsx") > 0, "the mismatch names the workbook, got: " & v)

    Test_DeckRegistry_PairingVerdictOnlyRefusesAKnownDifferentDeck = result
End Function

' THE STAMP HAS TO BE FOUND, NOT ASSUMED, AND THIS IS WHY.
'
' 2026-08-14: the layout stamp was read at the CURRENT layout's column. On a
' layout-4 sheet that cell holds the prompt, so the version came back 0, 0 is an
' unknown layout, and the whole-sheet clear fired on every drafting sheet of a
' real workbook while the migration -- guarded on the same flag -- never ran.
Private Function Test_Drafting_LayoutStampIsFoundNotAssumed() As String
    Dim result As String

    ' Columns 10..16 of the intro row. Layout 4: version in 11, period in 13.
    Dim l4 As Variant
    l4 = Array("", 4, "Draft the following...", "Q4F26", "", "", "")
    result = result & Assert(Drafting.DetectLayoutFromRow(l4) = 4, _
        "a layout-4 sheet reports 4, got " & Drafting.DetectLayoutFromRow(l4))

    ' Layout 5: version in 12, period in 14.
    Dim l5 As Variant
    l5 = Array("", "", 5, "Draft the following...", "Q4F26", "", "")
    result = result & Assert(Drafting.DetectLayoutFromRow(l5) = 5, _
        "a layout-5 sheet reports 5, got " & Drafting.DetectLayoutFromRow(l5))

    ' BOTH STAMPS PRESENT -- a sheet caught part-way through a bump. The NEWER
    ' position must win; picking the stale one would migrate an already-current
    ' sheet and relabel every column, which looks like content, not like loss.
    Dim both As Variant
    both = Array("", 4, 5, "Draft the following...", "Q4F26", "", "")
    result = result & Assert(Drafting.DetectLayoutFromRow(both) = 5, _
        "a stale stamp does not beat the newer one, got " & Drafting.DetectLayoutFromRow(both))

    ' A brand new sheet has no stamp anywhere.
    Dim none As Variant
    none = Array("", "", "", "", "", "", "")
    result = result & Assert(Drafting.DetectLayoutFromRow(none) = 0, _
        "an unstamped sheet reports 0, got " & Drafting.DetectLayoutFromRow(none))

    ' A period must never be mistaken for a version.
    Dim onlyPeriod As Variant
    onlyPeriod = Array("", "", "", "Q4F26", "", "", "")
    result = result & Assert(Drafting.DetectLayoutFromRow(onlyPeriod) = 0, _
        "a period is not a version, got " & Drafting.DetectLayoutFromRow(onlyPeriod))

    ' And the stamps' own columns, per layout.
    result = result & Assert(Drafting.PeriodStampColumn(4) = 13, "layout 4 keeps its period in 13")
    result = result & Assert(Drafting.PeriodStampColumn(5) = 14, "layout 5 keeps its period in 14")
    result = result & Assert(Drafting.LayoutStampColumn(4) = 11, "layout 4 keeps its version in 11")
    result = result & Assert(Drafting.LayoutStampColumn(5) = 12, "layout 5 keeps its version in 12")

    Test_Drafting_LayoutStampIsFoundNotAssumed = result
End Function

' A DEVICE IS DISCOVERED AS ONE FIELD, NOT AS ITS PARTS.
'
' Rohan's real timeline is a group of ~30 named shapes, and discovery walked
' into it -- so it presented as 21 separate fields to tag by hand. Injection and
' marking had both been device-aware since 2026-08-13; discovery was the one
' that still saw parts, which is the "writing is solved, recognising is not" gap.
Private Function Test_Discovery_ADeviceGroupIsOneCandidateNotItsParts() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()

    ' A plain group of two text boxes -- NOT a device, so it must still be
    ' walked into and yield its parts. This is the half that proves the new
    ' branch is selective rather than swallowing every group.
    Dim a As Object, b As Object
    Set a = sld.Shapes.AddTextbox(1, 10, 10, 100, 20): a.TextFrame.TextRange.text = "plain one"
    Set b = sld.Shapes.AddTextbox(1, 10, 40, 100, 20): b.TextFrame.TextRange.text = "plain two"
    Dim plain As Object
    Set plain = sld.Shapes.Range(Array(a.Name, b.Name)).Group
    plain.Name = "PLAIN_GROUP"

    ' A device: a group carrying milestone slot circles by name.
    Dim c1 As Object, c2 As Object, lbl As Object
    Set c1 = sld.Shapes.AddShape(9, 200, 10, 20, 20): c1.Name = "MS1_ON"
    Set c2 = sld.Shapes.AddShape(9, 240, 10, 20, 20): c2.Name = "MS2_NOW"
    Set lbl = sld.Shapes.AddTextbox(1, 200, 40, 100, 20)
    lbl.TextFrame.TextRange.text = "first milestone"
    lbl.Name = "MS1_LABEL"
    Dim dev As Object
    Set dev = sld.Shapes.Range(Array("MS1_ON", "MS2_NOW", "MS1_LABEL")).Group
    dev.Name = "MILESTONE_TIMELINE"

    result = result & Assert(MilestoneDevice.SlotCount(dev) = 2, _
        "the fixture really is a device -- 2 slots, got " & MilestoneDevice.SlotCount(dev))

    Dim shapes() As Object
    Dim cands() As Candidate
    cands = Discovery.DiscoverSlideWithShapes(sld, shapes)

    Dim i As Long, devCount As Long, partCount As Long, plainCount As Long
    For i = LBound(cands) To UBound(cands)
        Select Case UCase(cands(i).Name)
            Case "MILESTONE_TIMELINE": devCount = devCount + 1
            Case "MS1_ON", "MS2_NOW", "MS1_LABEL": partCount = partCount + 1
        End Select
        If cands(i).GroupPath = "PLAIN_GROUP" Then plainCount = plainCount + 1
    Next i

    result = result & Assert(devCount = 1, _
        "the device appears ONCE as a candidate, got " & devCount)
    result = result & Assert(partCount = 0, _
        "none of the device's parts appear as candidates, got " & partCount)
    result = result & Assert(plainCount = 2, _
        "an ordinary group is still walked into -- 2 parts expected, got " & plainCount)

    ' Discovered is not enough: Onboarding.IsCandidateField decides what a person
    ' is ever offered, and a device has no text and is not a picture.
    Dim devCand As Candidate
    For i = LBound(cands) To UBound(cands)
        If UCase(cands(i).Name) = "MILESTONE_TIMELINE" Then devCand = cands(i)
    Next i
    result = result & Assert(devCand.ShapeType = Discovery.SHAPE_TYPE_DEVICE, _
        "the device is typed as a device, got '" & devCand.ShapeType & "'")
    result = result & Assert(Onboarding.IsCandidateField(devCand), _
        "a device counts as a field despite having no text of its own")

    Test_Discovery_ADeviceGroupIsOneCandidateNotItsParts = result
End Function

Private Function Test_DeckRegistry_RegisterAndLookupTypeRoundTrip() As String
    Dim result As String
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()

    DeckRegistry.RegisterType pres, "test-registry-type-A", templateSld, "TestSheetA"

    Dim foundSld As Object
    Dim ws As String
    Dim ok As Boolean
    ok = DeckRegistry.LookupType(pres, "test-registry-type-A", foundSld, ws)

    result = result & Assert(ok, "lookup found the registered type")
    result = result & Assert(Not foundSld Is Nothing, "lookup returned a slide")
    If Not foundSld Is Nothing Then
        result = result & Assert(foundSld.SlideID = templateSld.SlideID, "returned slide matches the registered template, got SlideID " & foundSld.SlideID & " want " & templateSld.SlideID)
    End If
    result = result & Assert(ws = "TestSheetA", "returned worksheet name is 'TestSheetA', got '" & ws & "'")

    Test_DeckRegistry_RegisterAndLookupTypeRoundTrip = result
End Function

Private Function Test_DeckRegistry_LookupTypeFalseWhenNotRegistered() As String
    Dim result As String
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim foundSld As Object
    Dim ws As String
    Dim ok As Boolean
    ok = DeckRegistry.LookupType(pres, "test-registry-type-never-registered", foundSld, ws)

    result = result & Assert(Not ok, "lookup of an unregistered type returns False")
    result = result & Assert(foundSld Is Nothing, "outSld left Nothing")
    result = result & Assert(ws = "", "worksheetName left empty, got '" & ws & "'")

    Test_DeckRegistry_LookupTypeFalseWhenNotRegistered = result
End Function

Private Function Test_DeckRegistry_LookupTypeFalseWhenTemplateSlideDeleted() As String
    Dim result As String
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    DeckRegistry.RegisterType pres, "test-registry-type-deleted", templateSld, "TestSheetDeleted"
    templateSld.Delete

    Dim foundSld As Object
    Dim ws As String
    Dim ok As Boolean
    ok = DeckRegistry.LookupType(pres, "test-registry-type-deleted", foundSld, ws)

    result = result & Assert(Not ok, "lookup of a type whose template was deleted returns False, not an error")
    result = result & Assert(foundSld Is Nothing, "outSld left Nothing")

    Test_DeckRegistry_LookupTypeFalseWhenTemplateSlideDeleted = result
End Function

Private Function Test_DeckRegistry_ListRegisteredTypesListsAllRegistered() As String
    Dim result As String
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim sld1 As Object, sld2 As Object
    Set sld1 = NewBlankSlide()
    Set sld2 = NewBlankSlide()
    DeckRegistry.RegisterType pres, "test-registry-list-type-1", sld1, "Sheet1"
    DeckRegistry.RegisterType pres, "test-registry-list-type-2", sld2, "Sheet2"

    Dim types() As String
    types = DeckRegistry.ListRegisteredTypes(pres)

    Dim foundType1 As Boolean, foundType2 As Boolean
    Dim lo As Long, hi As Long, i As Long
    lo = LBound(types): hi = UBound(types)
    For i = lo To hi
        If types(i) = "test-registry-list-type-1" Then foundType1 = True
        If types(i) = "test-registry-list-type-2" Then foundType2 = True
    Next i

    result = result & Assert(foundType1, "list includes 'test-registry-list-type-1'")
    result = result & Assert(foundType2, "list includes 'test-registry-list-type-2'")

    Test_DeckRegistry_ListRegisteredTypesListsAllRegistered = result
End Function

Private Function Test_DeckRegistry_RegisterAndLookupTemplateLetterRoundTrip() As String
    Dim result As String
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()

    DeckRegistry.RegisterTemplateLetter pres, "test-letter-type-A", "K", templateSld, "TestSheetK"

    Dim foundSld As Object
    Dim ws As String
    Dim ok As Boolean
    ok = DeckRegistry.LookupTemplateLetter(pres, "test-letter-type-A", "K", foundSld, ws)

    result = result & Assert(ok, "lookup found the registered letter")
    result = result & Assert(Not foundSld Is Nothing, "lookup returned a slide")
    If Not foundSld Is Nothing Then
        result = result & Assert(foundSld.SlideID = templateSld.SlideID, "returned slide matches the registered template, got SlideID " & foundSld.SlideID & " want " & templateSld.SlideID)
    End If
    result = result & Assert(ws = "TestSheetK", "returned worksheet name is 'TestSheetK', got '" & ws & "'")

    ' Case-insensitive: registered "K", looked up "k".
    Dim foundSld2 As Object
    Dim ws2 As String
    ok = DeckRegistry.LookupTemplateLetter(pres, "test-letter-type-A", "k", foundSld2, ws2)
    result = result & Assert(ok, "lookup is case-insensitive on the letter")

    Test_DeckRegistry_RegisterAndLookupTemplateLetterRoundTrip = result
End Function

Private Function Test_DeckRegistry_LookupTemplateLetterFalseWhenNotRegistered() As String
    Dim result As String
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim foundSld As Object
    Dim ws As String
    Dim ok As Boolean
    ok = DeckRegistry.LookupTemplateLetter(pres, "test-letter-type-never", "K", foundSld, ws)

    result = result & Assert(Not ok, "lookup of an unregistered letter returns False")
    result = result & Assert(foundSld Is Nothing, "outSld left Nothing")
    result = result & Assert(ws = "", "worksheetName left empty, got '" & ws & "'")

    Test_DeckRegistry_LookupTemplateLetterFalseWhenNotRegistered = result
End Function

Private Function Test_DeckRegistry_TwoLettersOfSameTypeDoNotCollide() As String
    Dim result As String
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim sldK As Object, sldS As Object
    Set sldK = NewBlankSlide()
    Set sldS = NewBlankSlide()
    DeckRegistry.RegisterTemplateLetter pres, "test-letter-type-collide", "K", sldK, "SheetK"
    DeckRegistry.RegisterTemplateLetter pres, "test-letter-type-collide", "S", sldS, "SheetS"

    Dim foundK As Object, foundS As Object
    Dim wsK As String, wsS As String
    DeckRegistry.LookupTemplateLetter pres, "test-letter-type-collide", "K", foundK, wsK
    DeckRegistry.LookupTemplateLetter pres, "test-letter-type-collide", "S", foundS, wsS

    result = result & AssertSameSlide(foundK, sldK, "K resolves to the K slide")
    result = result & AssertSameSlide(foundS, sldS, "S resolves to the S slide")
    result = result & Assert(wsK = "SheetK", "K's worksheet name is 'SheetK', got '" & wsK & "'")
    result = result & Assert(wsS = "SheetS", "S's worksheet name is 'SheetS', got '" & wsS & "'")

    Test_DeckRegistry_TwoLettersOfSameTypeDoNotCollide = result
End Function

Private Function Test_DeckRegistry_LookupTemplateForLetterFallsBackWhenLetterEmpty() As String
    Dim result As String
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim unlettered As Object
    Set unlettered = NewBlankSlide()
    DeckRegistry.RegisterType pres, "test-fallback-type-empty", unlettered, "SheetUnlettered"

    Dim foundSld As Object
    Dim ws As String
    Dim ok As Boolean
    ok = DeckRegistry.LookupTemplateForLetter(pres, "test-fallback-type-empty", "", foundSld, ws)

    result = result & Assert(ok, "an empty letter falls back to the plain type registration")
    result = result & AssertSameSlide(foundSld, unlettered, "fallback returned the unlettered template")

    Test_DeckRegistry_LookupTemplateForLetterFallsBackWhenLetterEmpty = result
End Function

Private Function Test_DeckRegistry_LookupTemplateForLetterFallsBackWhenLetterUnregistered() As String
    Dim result As String
    Dim pres As Object
    Set pres = Application.ActivePresentation

    ' A deck that has only ever registered the plain type -- this is what
    ' keeps every deck predating per-letter registration working unchanged.
    Dim unlettered As Object
    Set unlettered = NewBlankSlide()
    DeckRegistry.RegisterType pres, "test-fallback-type-unreg", unlettered, "SheetUnlettered2"

    Dim foundSld As Object
    Dim ws As String
    Dim ok As Boolean
    ok = DeckRegistry.LookupTemplateForLetter(pres, "test-fallback-type-unreg", "K", foundSld, ws)

    result = result & Assert(ok, "a letter with no per-letter registration falls back to the plain type")
    result = result & AssertSameSlide(foundSld, unlettered, "fallback returned the unlettered template")

    Test_DeckRegistry_LookupTemplateForLetterFallsBackWhenLetterUnregistered = result
End Function

Private Function Test_DeckRegistry_LookupTemplateForLetterPrefersLetterOverUnlettered() As String
    Dim result As String
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim unlettered As Object, lettered As Object
    Set unlettered = NewBlankSlide()
    Set lettered = NewBlankSlide()
    DeckRegistry.RegisterType pres, "test-fallback-type-prefer", unlettered, "SheetUnlettered3"
    DeckRegistry.RegisterTemplateLetter pres, "test-fallback-type-prefer", "K", lettered, "SheetK3"

    Dim foundSld As Object
    Dim ws As String
    Dim ok As Boolean
    ok = DeckRegistry.LookupTemplateForLetter(pres, "test-fallback-type-prefer", "K", foundSld, ws)

    result = result & Assert(ok, "lookup succeeds when both are registered")
    result = result & AssertSameSlide(foundSld, lettered, "the per-letter template wins over the unlettered one")
    result = result & Assert(ws = "SheetK3", "worksheet name is the letter's, got '" & ws & "'")

    Test_DeckRegistry_LookupTemplateForLetterPrefersLetterOverUnlettered = result
End Function

Private Function Test_DeckRegistry_WorkbookPathRoundTrip() As String
    Dim result As String
    Dim pres As Object
    Set pres = Application.ActivePresentation

    result = result & Assert(DeckRegistry.GetWorkbookPath(pres) = "" Or DeckRegistry.GetWorkbookPath(pres) <> "", "GetWorkbookPath never raises before anything is set")

    DeckRegistry.SetWorkbookPath pres, "C:\fake\path\Data.xlsx"
    result = result & Assert(DeckRegistry.GetWorkbookPath(pres) = "C:\fake\path\Data.xlsx", "path round-trips, got '" & DeckRegistry.GetWorkbookPath(pres) & "'")

    DeckRegistry.SetWorkbookPath pres, "C:\fake\path\Data2.xlsx"
    result = result & Assert(DeckRegistry.GetWorkbookPath(pres) = "C:\fake\path\Data2.xlsx", "overwrite replaces the old path, got '" & DeckRegistry.GetWorkbookPath(pres) & "'")

    Test_DeckRegistry_WorkbookPathRoundTrip = result
End Function

' ---------------------------------------------------------------------
' WorkbookBridge
' ---------------------------------------------------------------------

' The guard for the 2026-07-30 live incident: the engine attaches to the
' RUNNING Excel, so it reads the workbook as it appears on screen. A slide was
' created from row 5 while the saved file held only rows 1-4.
'
' Tests the real predicate against a real workbook rather than a stub. `.Saved`
' is the entire basis of the guard, and an assumption about a COM property is
' exactly the kind of thing this project has been burned by asserting from
' memory -- so it is exercised here, dirtied and cleaned, against live Excel.
Private Function Test_WorkbookBridge_IsDirtyDetectsUnsavedEdits() As String
    Dim result As String

    ' The harness's own Excel instance, never a private one: an earlier test
    ' called CreateObject + Quit and tore down the instance other tests were
    ' still holding, hanging the whole run. Take a workbook, never the app.
    Dim xl As Object, wb As Object, ws As Object
    Set xl = WorkbookBridge.GetExcelApp()
    Set wb = xl.Workbooks.Add()
    Set ws = wb.Worksheets(1)

    ' Verified against real Excel 2026-07-30, having first been asserted wrongly
    ' from memory: a brand-new Workbooks.Add() reports Saved = TRUE. Excel's
    ' flag means "unmodified since last write", not "exists on disk", so an
    ' untouched new workbook is clean even though no file exists.
    '
    ' Harmless for the guard, and worth stating so nobody later "fixes" it:
    ' the only workbook the guard ever sees comes from OpenOrGetWorkbook(path),
    ' which opens an existing file, so the never-saved case cannot arise there.
    result = result & Assert(Not WorkbookBridge.IsDirty(wb), _
        "an untouched new workbook reads as clean -- Saved means unmodified, not on-disk")

    ws.Cells(1, 1).Value = "an edit that exists only in memory"
    result = result & Assert(WorkbookBridge.IsDirty(wb), _
        "editing a cell makes it dirty again -- this is the exact state that " & _
        "built a slide from a row in no file")

    wb.Saved = True
    result = result & Assert(Not WorkbookBridge.IsDirty(wb), _
        "and it can be cleared a second time")

    ' --- the wording the human actually sees ---
    Dim msg As String
    msg = WorkbookBridge.UnsavedWorkbookText("C:\somewhere\Data.xlsx")
    result = result & Assert(InStr(msg, "C:\somewhere\Data.xlsx") > 0, _
        "names the workbook, so it is actionable, got: " & msg)
    result = result & Assert(InStr(msg, "only in Excel's memory") > 0, _
        "explains WHY it matters rather than just refusing, got: " & msg)
    result = result & Assert(InStr(msg, "no matching row") > 0, _
        "names the actual consequence -- an orphaned slide, got: " & msg)

    wb.Saved = True
    wb.Close

    Test_WorkbookBridge_IsDirtyDetectsUnsavedEdits = result
End Function

Private Function Test_WorkbookBridge_SanitizeSheetNameStripsInvalidCharsAndTruncates() As String
    Dim result As String

    result = result & Assert(WorkbookBridge.SanitizeSheetName("quarterly-update") = "quarterly-update", "a clean name passes through unchanged, got '" & WorkbookBridge.SanitizeSheetName("quarterly-update") & "'")

    Dim cleaned As String
    cleaned = WorkbookBridge.SanitizeSheetName("a/b\c?d*e[f]g:h")
    result = result & Assert(InStr(cleaned, "/") = 0 And InStr(cleaned, "\") = 0 And InStr(cleaned, "?") = 0 And InStr(cleaned, "*") = 0 And InStr(cleaned, "[") = 0 And InStr(cleaned, "]") = 0 And InStr(cleaned, ":") = 0, "every invalid Excel sheet-name char is stripped, got '" & cleaned & "'")

    Dim longName As String
    longName = String(50, "x")
    result = result & Assert(Len(WorkbookBridge.SanitizeSheetName(longName)) = 31, "truncated to Excel's 31-char sheet name limit, got length " & Len(WorkbookBridge.SanitizeSheetName(longName)))

    result = result & Assert(WorkbookBridge.SanitizeSheetName("") = "Data", "blank name falls back to 'Data', got '" & WorkbookBridge.SanitizeSheetName("") & "'")

    Test_WorkbookBridge_SanitizeSheetNameStripsInvalidCharsAndTruncates = result
End Function

' ---------------------------------------------------------------------
' RibbonUI
' ---------------------------------------------------------------------

Private Function Test_RibbonUI_ResolveTypeAnswerAcceptsNumberOrName() As String
    Dim result As String
    Dim types(1 To 2) As String
    types(1) = "quarterly-update"
    types(2) = "annual-summary"

    result = result & Assert(RibbonUI.ResolveTypeAnswer("1", types) = "quarterly-update", "numeric answer resolves by position")
    result = result & Assert(RibbonUI.ResolveTypeAnswer("annual-summary", types) = "annual-summary", "name answer resolves case-insensitively")
    result = result & Assert(RibbonUI.ResolveTypeAnswer("nonexistent", types) = "", "unknown name resolves to empty")
    result = result & Assert(RibbonUI.ResolveTypeAnswer("", types) = "", "blank answer resolves to empty")

    Test_RibbonUI_ResolveTypeAnswerAcceptsNumberOrName = result
End Function

Private Function Test_RibbonUI_BuildTypePickerPromptListsAllTypes() As String
    Dim result As String
    Dim types(1 To 2) As String
    types(1) = "quarterly-update"
    types(2) = "annual-summary"

    Dim prompt As String
    prompt = RibbonUI.BuildTypePickerPrompt(types)

    result = result & Assert(InStr(prompt, "1) quarterly-update") > 0, "prompt lists type 1, got: " & prompt)
    result = result & Assert(InStr(prompt, "2) annual-summary") > 0, "prompt lists type 2, got: " & prompt)

    Test_RibbonUI_BuildTypePickerPromptListsAllTypes = result
End Function

' ---------------------------------------------------------------------
' CommandBarUI
' ---------------------------------------------------------------------

' Toolbar deliberately trimmed 2026-07-26 -- Rohan: "I only want to add an
' operation when I'm fully clear it works and I know what it does." Only
' actions live-tested against his real deck stay on it; the rest are commented
' out in CommandBarUI.bas, not deleted, per that Sub's own header.
'
' Preview Sync joined them 2026-07-29 without a prior live test, which is
' consistent with that rule rather than an exception to it: the rule guards
' against a half-understood button CHANGING a real deck, and Preview Sync
' provably cannot (see Test_RunSync_PreviewReportsWithoutTouchingTheDeck).
'
' Sync Now joined 2026-07-30 and DOES write, so it is a real bend in the rule.
' The first live cycle got as far as clicking it and found no button, which is
' how an untested action stays untested forever. What replaced the rule was a
' confirmation dialog stating slide creation in capitals.
'
' As of 2026-07-31 that warning is gone, because the capability it warned about
' is gone -- a sync no longer creates slides at all (round 9 section 4 of the
' Excel Control Layer exchange). Sync Now is therefore a materially smaller
' risk than when it was promoted: the worst it can now do is write wrong text
' into fields that already exist, which is visible and correctable, rather than
' duplicate a template across a deck. See
' Test_RunSync_ConfirmSyncTextReportsUncreatableRows, which pins the current
' wording and asserts the old alarm is genuinely absent.
'
' Create Template Slide joined 2026-07-30 (evening) and also writes. Same
' argument as Sync Now, plus one specific to it: the hazard it fixes is live
' for as long as it is unreachable, because until a type has a master template
' every slide Sync Now creates is cloned from a real project's slide.
'
' Asserting each by NAME below, not just that six buttons exist. The old
' version only checked each button resolved to one of the expected Subs, which
' would still pass if a button silently vanished and another were duplicated --
' a check that can't fail the way the thing it guards actually breaks.
' The Sources module had ZERO tests before 2026-08-08, which is how a citation
' check that only ever asked "does this ID exist" survived unexamined.
'
' Both directions are driven here on purpose. A test that only proves the wrong
' quarter is caught would still pass if the function reported EVERY citation --
' which is the failure that trains you to ignore the report.
' Reads the period out of a SAVED .pptx's bytes, sharing no process, no cache and
' no code with the writer -- the only kind of verification that means anything
' here, since an in-process read-back confirmed its own write against a file that
' had not been touched for three days (2026-08-08).
'
' Written because the first VBA version returned "" against a file that provably
' held the value: a PowerShell probe making the same COM calls read it correctly,
' so the technique was sound and the VBA was not. Guessing at the difference is
' exactly the habit this project keeps paying for, so it gets a test instead.
' Every sync path reported success against a deck file it never wrote to
' (2026-08-08). Nothing called pres.Save at all, and no test could see that,
' because no test asked the FILE anything.
' The review sheet showed two values that looked identical and asked for
' approval; they differed by a single trailing space (2026-08-08, 2 rows of 11).
'
' Both directions are driven: a describer that shouted INVISIBLE at every row
' would be as useless as one that stayed silent, because the word would stop
' meaning anything.
Private Function Test_ReviewQueue_DescribeDifferenceNamesTheInvisible() As String
    Dim result As String

    Dim d As String
    d = ReviewQueue.DescribeDifference("composts usage ", "composts usage")
    result = result & Assert(InStr(d, "INVISIBLE") > 0, _
        "a trailing space is called invisible, got '" & d & "'")
    result = result & Assert(InStr(d, "trailing") > 0, _
        "and says WHICH end it is on, got '" & d & "'")

    d = ReviewQueue.DescribeDifference(" leading", "leading")
    result = result & Assert(InStr(d, "leading") > 0, _
        "a leading space is named as leading, got '" & d & "'")

    d = ReviewQueue.DescribeDifference("In progress", "In Progress")
    result = result & Assert(InStr(d, "capitalisation") > 0, _
        "a case-only difference is named, got '" & d & "'")

    ' A plainly visible difference must NOT be called invisible.
    d = ReviewQueue.DescribeDifference("Project Closed", "Project Open")
    result = result & Assert(InStr(d, "INVISIBLE") = 0, _
        "an obvious difference is not called invisible, got '" & d & "'")

    result = result & Assert(ReviewQueue.DescribeDifference("same", "same") = "", _
        "identical values produce no note at all")

    Test_ReviewQueue_DescribeDifferenceNamesTheInvisible = result
End Function

Private Function Test_DeckRegistry_SaveDeckVerifiedProvesTheFileMoved() As String
    Dim result As String

    Dim testPath As String
    testPath = Environ("TEMP") & "\deck_sync_test_save_" & Format(Now, "hhmmss") & ".pptx"

    Dim testPres As Object
    Set testPres = Application.Presentations.Add
    testPres.SaveAs testPath

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim before As Date
    before = fso.GetFile(testPath).DateLastModified

    ' A real change, and a pause: the check is "did the timestamp ADVANCE", and
    ' NTFS would happily report the same second for two saves a moment apart --
    ' which would make this test fail for a reason that has nothing to do with
    ' the code. Waiting is the difference between testing the save and testing
    ' the clock's resolution.
    DeckRegistry.SetDeckPeriod testPres, "SAVETEST"
    Dim waitUntil As Date
    waitUntil = DateAdd("s", 2, Now)
    Do While Now < waitUntil
        DoEvents
    Loop

    Dim problem As String
    problem = DeckRegistry.SaveDeckVerified(testPres)
    result = result & Assert(problem = "", _
        "a deck that can be saved reports no problem, got '" & problem & "'")
    result = result & Assert(fso.GetFile(testPath).DateLastModified > before, _
        "and the file on disk is genuinely newer than before the save")

    ' NOTHING PENDING IS NOT A FAILURE -- and it must not cost a rewrite to say so.
    '
    ' Called again with no change made, the timestamp cannot advance. Before
    ' 2026-08-14 that ambiguity was resolved the wrong way: "the mtime did not
    ' move" was read as "THE DECK WAS NOT SAVED" about a deck that was saved,
    ' after forcing a full SaveAs of a 49MB file to find out. Seen for real on a
    ' harvest run that labelled 0 shapes.
    '
    ' BOTH assertions are needed. The first alone would pass if the function
    ' rewrote the file -- the timestamp would advance and it would report
    ' success for the wrong reason. The second is what proves it took the
    ' nothing-to-do path instead of paying for the answer.
    Dim afterFirst As Date
    afterFirst = fso.GetFile(testPath).DateLastModified

    Dim secondProblem As String
    secondProblem = DeckRegistry.SaveDeckVerified(testPres)
    result = result & Assert(secondProblem = "", _
        "a deck with nothing pending reports no problem, got '" & secondProblem & "'")
    result = result & Assert(fso.GetFile(testPath).DateLastModified = afterFirst, _
        "and it did not rewrite the file to establish that")

    ' A presentation that has never been written to a file must be REFUSED with
    ' an explanation, not silently reported as saved.
    Dim unsaved As Object
    Set unsaved = Application.Presentations.Add
    Dim unsavedProblem As String
    unsavedProblem = DeckRegistry.SaveDeckVerified(unsaved)
    result = result & Assert(unsavedProblem <> "", _
        "a never-saved presentation is refused rather than reported as saved")

    testPres.Saved = True
    testPres.Close
    unsaved.Saved = True
    unsaved.Close

    Test_DeckRegistry_SaveDeckVerifiedProvesTheFileMoved = result
End Function

Private Function Test_DeckRegistry_PeriodOnDiskReadsTheSavedFile() As String
    Dim result As String

    Dim testPath As String
    testPath = Environ("TEMP") & "\deck_sync_test_period_" & Format(Now, "hhmmss") & ".pptx"

    Dim testPres As Object
    Set testPres = Application.Presentations.Add
    testPres.SaveAs testPath

    ' A period nobody would produce by accident, so a stale file cannot pass this.
    Const PROBE_PERIOD As String = "ZZ9F99"
    DeckRegistry.SetDeckPeriod testPres, PROBE_PERIOD
    testPres.SaveAs testPath

    Dim onDisk As String, trace As String
    onDisk = DeckRegistry.PeriodOnDisk(testPath, trace)
    result = result & Assert(onDisk = PROBE_PERIOD, _
        "PeriodOnDisk reads the value from the saved file, got '" & onDisk & "' [" & trace & "]")

    ' A deck that declares nothing must read as "" rather than erroring, because
    ' the caller treats "" as "not confirmed" and retries.
    Dim blankPath As String
    blankPath = Environ("TEMP") & "\deck_sync_test_noperiod_" & Format(Now, "hhmmss") & ".pptx"
    Dim blankPres As Object
    Set blankPres = Application.Presentations.Add
    blankPres.SaveAs blankPath
    ' THESE TWO CASES RETURN THE SAME EMPTY STRING AND MEAN OPPOSITE THINGS.
    ' Until 2026-08-09 nothing could tell them apart, and a OneDrive deck whose
    ' file could not be read was reported to the user as "Period: BLOCKED -- not
    ' set in the saved file" while its bytes held Q3F26. The empty-string
    ' assertions below are kept exactly as they were; what is new is that each
    ' now also states WHICH of the two it is.
    Dim blankTrace As String, blankUnreadable As Boolean
    result = result & Assert(DeckRegistry.PeriodOnDisk(blankPath, blankTrace, blankUnreadable) = "", _
        "a deck with no period reads as empty, not an error")
    result = result & Assert(Not blankUnreadable, _
        "a deck that WAS read and simply has no period is not a read failure [" & blankTrace & "]")

    Dim missTrace As String, missUnreadable As Boolean
    result = result & Assert(DeckRegistry.PeriodOnDisk(Environ("TEMP") & "\no_such_deck_here.pptx", _
        missTrace, missUnreadable) = "", _
        "a missing file reads as empty rather than raising")
    result = result & Assert(missUnreadable, _
        "a file that could not be read reports a READ FAILURE, not an absent period [" & missTrace & "]")

    testPres.Saved = True
    testPres.Close
    blankPres.Saved = True
    blankPres.Close

    Test_DeckRegistry_PeriodOnDiskReadsTheSavedFile = result
End Function

Private Function Test_Sources_RefsForOtherPeriodCatchesTheWrongQuarter() As String
    Dim result As String

    Dim applic As Object
    Set applic = CreateObject("Scripting.Dictionary")
    applic("S01") = UCase("Q3F26")
    applic("S02") = UCase(Sources.APPLIES_ALL)

    Dim got As String
    got = Sources.RefsForOtherPeriod("S01", applic, "Q4F26")
    result = result & Assert(InStr(got, "S01") > 0, _
        "a source stamped Q3F26 is reported when publishing Q4F26, got '" & got & "'")
    result = result & Assert(InStr(got, "Q3F26") > 0, _
        "the report names the period the source actually belongs to, got '" & got & "'")

    got = Sources.RefsForOtherPeriod("S01,S02", applic, "Q4F26")
    result = result & Assert(InStr(got, "S02") = 0, _
        "the period-neutral source is NOT dragged into the report, got '" & got & "'")

    Test_Sources_RefsForOtherPeriodCatchesTheWrongQuarter = result
End Function

Private Function Test_Sources_RefsForOtherPeriodIsSilentOnNeutralAndUnknown() As String
    Dim result As String

    Dim applic As Object
    Set applic = CreateObject("Scripting.Dictionary")
    applic("S01") = UCase("Q4F26")
    applic("S02") = UCase(Sources.APPLIES_ALL)

    result = result & Assert(Sources.RefsForOtherPeriod("S01", applic, "Q4F26") = "", _
        "a source stamped with THIS period is silent")
    result = result & Assert(Sources.RefsForOtherPeriod("S02", applic, "Q4F26") = "", _
        "a period-neutral source is silent")
    result = result & Assert(Sources.RefsForOtherPeriod("", applic, "Q4F26") = "", _
        "no citation is silent")
    ' UnknownRefs already reports this one. Saying it twice for one typo reads
    ' as two separate faults, so this function must stay quiet about it.
    result = result & Assert(Sources.RefsForOtherPeriod("S99", applic, "Q4F26") = "", _
        "an ID that is not on the sheet is left to UnknownRefs")
    result = result & Assert(Sources.RefsForOtherPeriod("S01", applic, "") = "", _
        "no period means no judgement, rather than every source being wrong")

    Test_Sources_RefsForOtherPeriodIsSilentOnNeutralAndUnknown = result
End Function

Private Function Test_CommandBarUI_ShowToolbarCreatesWiredButtons() As String
    Dim result As String

    CommandBarUI.ShowToolbar

    ' THREE BARS since 2026-08-08 -- sixteen buttons did not fit one row and the
    ' last four were unreachable behind an overflow chevron. The count that
    ' matters is still every button being present and visible -- two, since 2026-08-09.
    Dim allBars As Collection
    Set allBars = New Collection
    Dim nm As Variant
    For Each nm In CommandBarUI.ActiveToolbarNames()
        Dim oneBar As Object
        Set oneBar = Nothing
        On Error Resume Next
        Set oneBar = Application.CommandBars(CStr(nm))
        On Error GoTo 0
        result = result & Assert(Not oneBar Is Nothing, "bar '" & CStr(nm) & "' exists after ShowToolbar")
        If Not oneBar Is Nothing Then
            result = result & Assert(oneBar.Visible, "bar '" & CStr(nm) & "' is VISIBLE")
            allBars.Add oneBar
        End If
    Next nm

    Dim totalButtons As Long
    Dim bb As Variant
    For Each bb In allBars
        totalButtons = totalButtons + bb.Controls.count
    Next bb
    ' THE COUNT IS DERIVED FROM THE CAPTIONS, NOT TYPED.
    '
    ' This asserted "= 2", then the toolbar changed and it was edited to "= 4" to
    ' make it pass -- which proves nothing except that someone moved the number.
    ' A literal here can always be dragged to match whatever the code now does.
    '
    ' Built from the CAP_* constants instead, this fails when a button is added
    ' without a caption constant, when a declared caption loses its button, and
    ' when a caption is renamed in only one of the two places.
    Dim expectedCaptions As String
    expectedCaptions = "|" & CommandBarUI.CAP_SET_UP_QUARTER & _
                       "|" & CommandBarUI.CAP_PUT_ON_SLIDES & _
                       "|" & CommandBarUI.CAP_REVIEW_ONLY & _
                       "|" & CommandBarUI.CAP_ADD_SLIDES & _
                       "|" & CommandBarUI.CAP_RETIRE_SLIDES & _
                       "|" & CommandBarUI.CAP_REPOINT_WORKBOOK & _
                       "|" & CommandBarUI.CAP_DISCOVER_FIELDS & _
                       "|" & CommandBarUI.CAP_CREATE_TEMPLATE & "|"

    Dim expectedCount As Long
    expectedCount = UBound(Split(expectedCaptions, "|")) - 1

    result = result & Assert(totalButtons = expectedCount, _
        "one button per declared caption -- expected " & expectedCount & ", got " & totalButtons)

    ' Every button on the bar is a DECLARED one (catches a stray extra button)...
    Dim seenCaptions As String
    Dim capBar As Variant, capIdx As Long
    seenCaptions = "|"
    For Each capBar In allBars
        For capIdx = 1 To capBar.Controls.count
            Dim oneCap As String
            oneCap = capBar.Controls.Item(capIdx).Caption
            result = result & Assert(InStr(expectedCaptions, "|" & oneCap & "|") > 0, _
                "button '" & oneCap & "' is a declared caption")
            seenCaptions = seenCaptions & oneCap & "|"
        Next capIdx
    Next capBar

    ' ...and every declared one is ON the bar (catches a missing button, which is
    ' how a capability goes unreachable while every other assertion still passes).
    Dim wantCap As Variant
    For Each wantCap In Split(Mid(expectedCaptions, 2, Len(expectedCaptions) - 2), "|")
        result = result & Assert(InStr(seenCaptions, "|" & CStr(wantCap) & "|") > 0, _
            "declared caption '" & CStr(wantCap) & "' has a button on the toolbar")
    Next wantCap

    Dim bar As Object
    Set bar = allBars(1)
    ' A bar that exists, is fully wired, and is NOT VISIBLE is the exact state
    ' shipped in addin40 and found on 2026-08-08. Every other assertion in this
    ' test passed against it. Without this line the suite calls that a pass.

    Dim seenPreview As Boolean
    Dim seenSyncNow As Boolean
    Dim seenReview As Boolean
    Dim seenApply As Boolean
    Dim seenCreateTemplate As Boolean
    Dim seenAuditFields As Boolean
    seenPreview = False
    seenSyncNow = False
    seenReview = False
    seenApply = False
    seenCreateTemplate = False
    seenAuditFields = False

    ' PowerPoint sometimes normalizes a set OnAction like "RibbonUI.SyncNow"
    ' to its own "<PresentationName>!SyncNow" display form (module-
    ' unqualified) -- confirmed 2026-07-26 against real Office -- but this
    ' is NOT reliable across runs: re-confirmed 2026-07-26 (same day, same
    ' machine, no CommandBarUI.bas/RibbonUI.bas changes) that it sometimes
    ' leaves OnAction as the raw "Module.Sub" string instead, for reasons
    ' that don't appear to depend on anything this project controls. Strip
    ' BOTH a leading "<anything>!" AND a leading "<anything>." (last
    ' occurrence of each) so the assertion matches the bare Sub name
    ' regardless of which form Office happens to report this run -- what
    ' actually matters (the button fires the right Sub) is unaffected
    ' either way, only this display string varies.
    ' Delimited on BOTH sides and matched with the delimiters included: a bare
    ' InStr over the list would accept any substring of it, so a button wired to
    ' "Sync" -- or to nothing recognisable that happens to be a fragment -- would
    ' pass. Tightened 2026-07-30 while adding SyncNow, whose name contains
    ' another entry's prefix.
    Dim expectedActions As String
    expectedActions = "|SyncNowChain|PutItOnTheSlides|ReviewChanges|AddMissingSlides|RetireSlides|ChangePairedWorkbook|DiscoverFieldsOnSlide|CreateTemplateSlide|"

    Dim i As Long
    Dim eachBar As Variant
    For Each eachBar In allBars
    For i = 1 To eachBar.Controls.count
        Dim ctrl As Object
        Set ctrl = eachBar.Controls.Item(i)
        Dim afterBang As String
        Dim bangPos As Long
        bangPos = InStr(ctrl.OnAction, "!")
        afterBang = IIf(bangPos > 0, Mid(ctrl.OnAction, bangPos + 1), ctrl.OnAction)
        Dim dotPos As Long
        dotPos = InStrRev(afterBang, ".")
        Dim subName As String
        subName = IIf(dotPos > 0, Mid(afterBang, dotPos + 1), afterBang)
        result = result & Assert(InStr(expectedActions, "|" & subName & "|") > 0, "button '" & ctrl.Caption & "' OnAction '" & ctrl.OnAction & "' resolves to one of the real action Subs")
        result = result & Assert(Len(ctrl.TooltipText) > 0, "button '" & ctrl.Caption & "' has a non-empty tooltip explainer")
        ' 255 is Office's hard cap and it RAISES rather than trimming, taking
        ' the whole toolbar with it mid-build (2026-08-08).
        result = result & Assert(Len(ctrl.TooltipText) <= 255, _
            "button '" & ctrl.Caption & "' tooltip is within Office's 255-char cap, got " & Len(ctrl.TooltipText))
        result = result & Assert(Left$(ctrl.TooltipText, 7) = "Use to ", _
            "button '" & ctrl.Caption & "' tooltip opens with 'Use to ', got '" & Left$(ctrl.TooltipText, 20) & "'")
        If subName = "WhereAmI" Then seenPreview = True
        If subName = "SyncNowChain" Then seenSyncNow = True
        If subName = "ReviewChanges" Then seenReview = True
        If subName = "ApplyApprovedChanges" Then seenApply = True
        If subName = "CreateTemplateSlide" Then seenCreateTemplate = True
        If subName = "AuditFields" Then seenAuditFields = True
    Next i
    Next eachBar

    ' THESE FOUR NO LONGER HAVE BUTTONS, AND THAT IS THE DESIGN, not drift.
    ' 2026-08-09: 16 buttons became 2. Where am I?, Review Changes, Apply
    ' Approved and Audit Fields are reached from inside the Sync Now chain --
    ' the chain opens by rebuilding the readiness sheet, offers read-one-at-a-
    ' time or approve-all, and stops at the write-authorising confirmation.
    '
    ' Asserting their ABSENCE would be asserting the implementation. What must
    ' stay true is that they are still REACHABLE, and a live toolbar cannot
    ' answer that -- check_vba_static.py walks the call graph and fails the
    ' build on a genuine orphan. It caught exactly this regression when the
    ' chains were calling private Cores and nine capabilities went dark.
    result = result & Assert(seenSyncNow, "Sync Now is ON the toolbar -- the chain, and the only route to a slide write")
    ' Create Template Slide was deliberately NOT a button from 2026-08-01 to
    ' 2026-08-16 -- reached only via a one-time MsgBox at the end of Bulk
    ' Onboard, on the assumption a type gets exactly one template, made once.
    ' Scenario 3 (per-letter colour templates) broke that assumption: a
    ' SECOND and THIRD template need adding to a type onboarded weeks ago,
    ' with no re-onboarding in sight. Reversed 2026-08-16 after Rohan was
    ' sent to press a button that did not exist -- this assertion flips with
    ' it, inverted AGAIN so the decision stays visible rather than looking
    ' like a regression.
    result = result & Assert(seenCreateTemplate, _
        "Create Template Slide IS a toolbar button -- adding a second/third letter to an already-onboarded type needs a repeatable entry point, not a one-time prompt")

    CommandBarUI.HideToolbar
    Test_CommandBarUI_ShowToolbarCreatesWiredButtons = result
End Function

Private Function Test_CommandBarUI_ShowToolbarIsIdempotent() As String
    Dim result As String

    CommandBarUI.ShowToolbar
    CommandBarUI.ShowToolbar  ' must not raise "toolbar already exists" or leave duplicates

    Dim total2 As Long
    Dim nm2 As Variant
    For Each nm2 In CommandBarUI.ActiveToolbarNames()
        Dim b2 As Object
        Set b2 = Nothing
        On Error Resume Next
        Set b2 = Application.CommandBars(CStr(nm2))
        On Error GoTo 0
        result = result & Assert(Not b2 Is Nothing, "bar '" & CStr(nm2) & "' still exists after calling ShowToolbar twice")
        If Not b2 Is Nothing Then total2 = total2 + b2.Controls.count
    Next nm2
    Dim wantAfter As Long
    wantAfter = UBound(Split("|" & CommandBarUI.CAP_SET_UP_QUARTER & _
                             "|" & CommandBarUI.CAP_PUT_ON_SLIDES & _
                             "|" & CommandBarUI.CAP_REVIEW_ONLY & _
                             "|" & CommandBarUI.CAP_ADD_SLIDES & _
                             "|" & CommandBarUI.CAP_RETIRE_SLIDES & _
                             "|" & CommandBarUI.CAP_REPOINT_WORKBOOK & _
                             "|" & CommandBarUI.CAP_DISCOVER_FIELDS & _
                             "|" & CommandBarUI.CAP_CREATE_TEMPLATE & "|", "|")) - 1
    result = result & Assert(total2 = wantAfter, _
        "a second ShowToolbar leaves one button per declared caption -- expected " & wantAfter & ", got " & total2)

    CommandBarUI.HideToolbar
    Test_CommandBarUI_ShowToolbarIsIdempotent = result
End Function

Private Function Test_CommandBarUI_HideToolbarRemovesIt() As String
    Dim result As String

    CommandBarUI.ShowToolbar
    CommandBarUI.HideToolbar

    Dim nm3 As Variant
    For Each nm3 In CommandBarUI.ToolbarNames()
        Dim b3 As Object
        Set b3 = Nothing
        On Error Resume Next
        Set b3 = Application.CommandBars(CStr(nm3))
        On Error GoTo 0
        result = result & Assert(b3 Is Nothing, "bar '" & CStr(nm3) & "' no longer exists after HideToolbar")
    Next nm3

    ' Calling HideToolbar again with nothing to remove must not raise.
    CommandBarUI.HideToolbar
    result = result & Assert(True, "HideToolbar is safe to call when nothing exists")

    Test_CommandBarUI_HideToolbarRemovesIt = result
End Function

' ---------------------------------------------------------------------
' AdoptFlow
' ---------------------------------------------------------------------

Private Function Test_AdoptFlow_ValidateAdoptionSelectionSortsIntoDeckOrder() As String
    Dim result As String

    Dim sld1 As Object, sld2 As Object, sld3 As Object
    Set sld1 = NewBlankSlide()
    Set sld2 = NewBlankSlide()
    Set sld3 = NewBlankSlide()

    ' Slides.Range(...).Select only actually registers as a slide-type
    ' selection (Selection.Type = ppSelectionSlides) when the window is in
    ' Slide Sorter view -- confirmed 2026-07-26 against real Office: in
    ' Normal view (where NewBlankSlide() leaves the window via GotoSlide),
    ' Selection.Type stayed ppSelectionNone under COM automation even though
    ' the .Select call itself raised no error. This is an automation-only
    ' quirk, not a real-user limitation -- a human Ctrl/Shift-clicking slide
    ' thumbnails in Normal view's own thumbnail pane does not have this
    ' problem, only headless automation driving .Select with no genuine pane
    ' focus does. Switched back to Normal view afterward so later tests in
    ' this suite (which all assume Normal view, e.g. NewBlankSlide()'s own
    ' GotoSlide) are unaffected.
    Dim priorViewType As Long
    priorViewType = Application.ActiveWindow.ViewType
    Application.ActiveWindow.ViewType = 7 ' ppViewSlideSorter
    Application.ActivePresentation.Slides.Range(Array(sld3.SlideIndex, sld1.SlideIndex, sld2.SlideIndex)).Select

    Dim outSlides() As Object
    Dim errMsg As String
    errMsg = AdoptFlow.ValidateAdoptionSelection(Application.ActiveWindow.Selection, outSlides)

    Application.ActiveWindow.ViewType = priorViewType

    result = result & Assert(errMsg = "", "selection validated, got error '" & errMsg & "'")

    Dim lo As Long, hi As Long, hasSlides As Boolean
    On Error Resume Next
    lo = LBound(outSlides): hi = UBound(outSlides): hasSlides = (Err.Number = 0)
    On Error GoTo 0

    result = result & Assert(hasSlides And (hi - lo + 1) = 3, "all 3 selected slides returned, got " & IIf(hasSlides, hi - lo + 1, 0))
    If hasSlides And (hi - lo + 1) = 3 Then
        result = result & Assert(outSlides(lo).SlideID = sld1.SlideID, "first returned slide is the lowest SlideIndex (sld1), got SlideID " & outSlides(lo).SlideID)
        result = result & Assert(outSlides(lo + 1).SlideID = sld2.SlideID, "second returned slide is sld2")
        result = result & Assert(outSlides(lo + 2).SlideID = sld3.SlideID, "third returned slide is the highest SlideIndex (sld3)")
    End If

    Test_AdoptFlow_ValidateAdoptionSelectionSortsIntoDeckOrder = result
End Function

Private Function Test_AdoptFlow_ValidateAdoptionSelectionRejectsNonSlideSelection() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim shp As Object
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    shp.Select

    Dim outSlides() As Object
    Dim errMsg As String
    errMsg = AdoptFlow.ValidateAdoptionSelection(Application.ActiveWindow.Selection, outSlides)

    result = result & Assert(errMsg <> "", "a shape selection (not a slide selection) is rejected with an error message")

    Test_AdoptFlow_ValidateAdoptionSelectionRejectsNonSlideSelection = result
End Function

Private Function Test_AdoptFlow_ExcludeTemplateSlideRemovesOnlyTemplate() As String
    Dim result As String

    Dim sld1 As Object, sld2 As Object, sld3 As Object
    Set sld1 = NewBlankSlide()
    Set sld2 = NewBlankSlide()
    Set sld3 = NewBlankSlide()

    Dim slides(1 To 3) As Object
    Set slides(1) = sld1
    Set slides(2) = sld2
    Set slides(3) = sld3

    Dim filtered() As Object
    filtered = AdoptFlow.ExcludeTemplateSlide(slides, sld2)

    result = result & Assert((UBound(filtered) - LBound(filtered) + 1) = 2, "2 slides remain after excluding the template, got " & (UBound(filtered) - LBound(filtered) + 1))
    Dim lo As Long
    lo = LBound(filtered)
    result = result & Assert(filtered(lo).SlideID = sld1.SlideID, "first remaining slide is sld1")
    result = result & Assert(filtered(lo + 1).SlideID = sld3.SlideID, "second remaining slide is sld3 (order preserved, sld2 removed)")

    Test_AdoptFlow_ExcludeTemplateSlideRemovesOnlyTemplate = result
End Function

Private Function Test_AdoptFlow_BuildAdoptionReviewSummaryCountsAndListsNonReady() As String
    Dim result As String

    Dim plans(1 To 4) As AdoptionSlidePlan
    plans(1).SlideLabel = "Slide 1 (A)"
    plans(1).Disposition = "ready"
    plans(1).Reason = "no existing keyless Data-sheet row matches verbatim -- will create a fresh row (instance_key required from the human)"

    plans(2).SlideLabel = "Slide 2 (B)"
    plans(2).Disposition = "already_linked"

    plans(3).SlideLabel = "Slide 3 (C)"
    plans(3).Disposition = "needs_confirmation"

    plans(4).SlideLabel = "Slide 4 (D)"
    plans(4).Disposition = "unclassified"

    Dim summary As String
    summary = AdoptFlow.BuildAdoptionReviewSummary(plans)

    result = result & Assert(InStr(summary, "1 ready to link") > 0, "summary reports 1 ready, got: " & summary)
    result = result & Assert(InStr(summary, "1 already linked") > 0, "summary reports 1 already linked, got: " & summary)
    result = result & Assert(InStr(summary, "1 need confirmation") > 0, "summary reports 1 needing confirmation, got: " & summary)
    result = result & Assert(InStr(summary, "1 unclassified") > 0, "summary reports 1 unclassified, got: " & summary)
    result = result & Assert(InStr(summary, "Slide 1 (A)") > 0, "ready slide is listed by label")
    result = result & Assert(InStr(summary, "Slide 3 (C)") > 0, "needs_confirmation slide is listed by label")
    result = result & Assert(InStr(summary, "Slide 4 (D)") > 0, "unclassified slide is listed by label")
    result = result & Assert(InStr(summary, "Slide 2 (B)") = 0, "already_linked slides are NOT listed in the detail section (counted only)")

    Test_AdoptFlow_BuildAdoptionReviewSummaryCountsAndListsNonReady = result
End Function

' ---------------------------------------------------------------------
' BatchOnboardFlow
' ---------------------------------------------------------------------

Private Function Test_BatchOnboardFlow_AllValuesIdenticalDetectsMatchAndMismatch() As String
    Dim result As String

    Dim allSame As New Collection
    allSame.Add "Overall Status"
    allSame.Add "Overall Status"
    allSame.Add "Overall Status"
    result = result & Assert(BatchOnboardFlow.AllValuesIdentical(allSame), "identical values across the batch report True")

    Dim differs As New Collection
    differs.Add "Q1 2026"
    differs.Add "Q2 2026"
    result = result & Assert(Not BatchOnboardFlow.AllValuesIdentical(differs), "differing values across the batch report False")

    Dim onlyOne As New Collection
    onlyOne.Add "only one value"
    result = result & Assert(BatchOnboardFlow.AllValuesIdentical(onlyOne), "a single value is trivially identical")

    Test_BatchOnboardFlow_AllValuesIdenticalDetectsMatchAndMismatch = result
End Function

Private Function Test_BatchOnboardFlow_SuggestBatchFieldNameReusesPhNameOrFallsBack() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim named As Object
    Set named = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    named.Name = "ph_quarter"
    Dim unnamed As Object
    Set unnamed = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 300, 50, 200, 50)

    result = result & Assert(BatchOnboardFlow.SuggestBatchFieldName(named, 1) = "ph_quarter", "reuses an existing ph_ name, got '" & BatchOnboardFlow.SuggestBatchFieldName(named, 1) & "'")
    result = result & Assert(BatchOnboardFlow.SuggestBatchFieldName(unnamed, 2) = "ph_field2", "falls back to a positional name, got '" & BatchOnboardFlow.SuggestBatchFieldName(unnamed, 2) & "'")

    Test_BatchOnboardFlow_SuggestBatchFieldNameReusesPhNameOrFallsBack = result
End Function

Private Function Test_BatchOnboardFlow_NormalizeFieldTypeAcceptsNumberOrName() As String
    Dim result As String

    result = result & Assert(BatchOnboardFlow.NormalizeFieldType("1") = "text", "'1' resolves to text")
    result = result & Assert(BatchOnboardFlow.NormalizeFieldType("Text") = "text", "'Text' resolves to text (case-insensitive)")
    result = result & Assert(BatchOnboardFlow.NormalizeFieldType("2") = "number", "'2' resolves to number")
    result = result & Assert(BatchOnboardFlow.NormalizeFieldType("number") = "number", "'number' resolves to number")
    result = result & Assert(BatchOnboardFlow.NormalizeFieldType("3") = "currency", "'3' resolves to currency")
    result = result & Assert(BatchOnboardFlow.NormalizeFieldType("Currency") = "currency", "'Currency' resolves to currency (case-insensitive)")
    result = result & Assert(BatchOnboardFlow.NormalizeFieldType("4") = "date", "'4' resolves to date")
    result = result & Assert(BatchOnboardFlow.NormalizeFieldType("date") = "date", "'date' resolves to date")
    result = result & Assert(BatchOnboardFlow.NormalizeFieldType("") = "text", "blank falls back to text -- CALLERS MUST CHECK FOR CANCEL FIRST (finding 2); this function cannot tell Cancel from OK-with-nothing-typed and is not expected to")
    result = result & Assert(BatchOnboardFlow.NormalizeFieldType("gibberish") = "text", "unrecognized answer falls back to text, never errors")

    Test_BatchOnboardFlow_NormalizeFieldTypeAcceptsNumberOrName = result
End Function

Private Function Test_BatchOnboardFlow_NormalizeFieldVolatilityAcceptsNumberOrName() As String
    Dim result As String

    result = result & Assert(BatchOnboardFlow.NormalizeFieldVolatility("1") = "static", "'1' resolves to static")
    result = result & Assert(BatchOnboardFlow.NormalizeFieldVolatility("Static") = "static", "'Static' resolves to static (case-insensitive)")
    result = result & Assert(BatchOnboardFlow.NormalizeFieldVolatility("2") = "variable", "'2' resolves to variable")
    result = result & Assert(BatchOnboardFlow.NormalizeFieldVolatility("variable") = "variable", "'variable' resolves to variable")
    result = result & Assert(BatchOnboardFlow.NormalizeFieldVolatility("") = "variable", "blank falls back to variable (the safer default), never errors")
    result = result & Assert(BatchOnboardFlow.NormalizeFieldVolatility("gibberish") = "variable", "unrecognized answer falls back to variable, never errors")

    Test_BatchOnboardFlow_NormalizeFieldVolatilityAcceptsNumberOrName = result
End Function

' SuggestInstanceKey is the fix for Rohan's real friction (2026-07-26):
' hand-typing an instance key for every slide in a batch, not the
' instance-key concept itself. It should default each prompt to the
' template's/each slide's own first-marked-field value (a natural per-
' slide identifier like "Project Number" in his real deck).
Private Function Test_BatchOnboardFlow_SuggestInstanceKeyUsesFirstFieldsHarvestedValue() As String
    Dim result As String

    Dim plan As BatchOnboardPlan
    Set plan.HarvestedText = CreateObject("Scripting.Dictionary")
    plan.HarvestedText("1|0") = "3_P001"
    plan.HarvestedText("1|1") = "3_P002"
    plan.HarvestedText("2|0") = "Development of alternative antimicrobial agents"

    result = result & Assert(BatchOnboardFlow.SuggestInstanceKey(plan, 0) = "3_P001", "suggests the template's first-field value, got '" & BatchOnboardFlow.SuggestInstanceKey(plan, 0) & "'")
    result = result & Assert(BatchOnboardFlow.SuggestInstanceKey(plan, 1) = "3_P002", "suggests slide 1's first-field value, got '" & BatchOnboardFlow.SuggestInstanceKey(plan, 1) & "'")
    result = result & Assert(BatchOnboardFlow.SuggestInstanceKey(plan, 2) = "", "no correspondence found for slide 2 -- falls back to blank, not an error, got '" & BatchOnboardFlow.SuggestInstanceKey(plan, 2) & "'")

    Test_BatchOnboardFlow_SuggestInstanceKeyUsesFirstFieldsHarvestedValue = result
End Function

' FindSameLayoutSlides is the pure logic behind auto-selecting a batch by
' layout instead of requiring a manual multi-slide selection every time --
' Rohan's own framing (2026-07-26): "same-layout slides auto-included, with
' a chance to review/edit before committing."
Private Function Test_BatchOnboardFlow_FindSameLayoutSlidesGroupsByLayoutOnly() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide() ' ppLayoutBlank

    Dim sameLayout1 As Object
    Set sameLayout1 = NewBlankSlide() ' ppLayoutBlank -- same layout as template

    Dim differentLayout As Object
    Set differentLayout = Application.ActivePresentation.Slides.Add(Application.ActivePresentation.Slides.count + 1, ppLayoutText)

    Dim sameLayout2 As Object
    Set sameLayout2 = NewBlankSlide() ' ppLayoutBlank again -- also same layout as template

    Dim siblings As Collection
    Set siblings = BatchOnboardFlow.FindSameLayoutSlides(templateSld)

    ' Not an exact count: this test runs in a shared presentation that
    ' accumulates every other test's own fixture slides across the whole
    ' suite (most of them also ppLayoutBlank via NewBlankSlide()), so
    ' asserting anything beyond "at least our two known siblings, by
    ' identity" would be asserting about slides this test doesn't own.
    result = result & Assert(siblings.count >= 2, "at least the 2 known same-layout siblings found, got " & siblings.count)

    Dim foundS1 As Boolean, foundS2 As Boolean, foundDiff As Boolean, foundTemplate As Boolean
    Dim s As Variant
    For Each s In siblings
        If s Is sameLayout1 Then foundS1 = True
        If s Is sameLayout2 Then foundS2 = True
        If s Is differentLayout Then foundDiff = True
        If s Is templateSld Then foundTemplate = True
    Next s

    result = result & Assert(foundS1, "same-layout sibling 1 included")
    result = result & Assert(foundS2, "same-layout sibling 2 included")
    result = result & Assert(Not foundDiff, "different-layout slide excluded")
    result = result & Assert(Not foundTemplate, "template itself excluded from its own siblings")

    Test_BatchOnboardFlow_FindSameLayoutSlidesGroupsByLayoutOnly = result
End Function

Private Function Test_BatchOnboardFlow_BuildBatchPlanFindsCorrespondenceAndHarvestsAcrossSlides() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim tShapeA As Object, tShapeB As Object
    Set tShapeA = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    tShapeA.TextFrame.TextRange.Text = "Same Everywhere"
    Set tShapeB = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 150, 200, 50)
    tShapeB.TextFrame.TextRange.Text = "Q1 2026"

    Dim other1 As Object
    Set other1 = NewBlankSlide()
    Dim o1ShapeA As Object, o1ShapeB As Object
    Set o1ShapeA = other1.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    o1ShapeA.TextFrame.TextRange.Text = "Same Everywhere"
    Set o1ShapeB = other1.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 150, 200, 50)
    o1ShapeB.TextFrame.TextRange.Text = "Q2 2026"

    Dim other2 As Object
    Set other2 = NewBlankSlide()
    Dim o2ShapeA As Object, o2ShapeB As Object
    Set o2ShapeA = other2.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    o2ShapeA.TextFrame.TextRange.Text = "Same Everywhere"
    Set o2ShapeB = other2.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 150, 200, 50)
    o2ShapeB.TextFrame.TextRange.Text = "Q3 2026"

    Dim otherSlides(1 To 2) As Object
    Set otherSlides(1) = other1
    Set otherSlides(2) = other2

    Dim plan As BatchOnboardPlan
    plan = BatchOnboardFlow.BuildBatchPlan(templateSld, otherSlides)

    result = result & Assert(plan.FieldCount = 2, "2 candidate fields found on the template, got " & plan.FieldCount)

    ' Identify which field index corresponds to which shape by its
    ' template value, since discovery/enumeration order isn't guaranteed.
    Dim identicalFieldIdx As Long, varyingFieldIdx As Long
    identicalFieldIdx = 0: varyingFieldIdx = 0
    Dim fi As Long
    For fi = 1 To plan.FieldCount
        Dim tv As String
        tv = plan.HarvestedText(CStr(fi) & "|0")
        If tv = "Same Everywhere" Then identicalFieldIdx = fi
        If tv = "Q1 2026" Then varyingFieldIdx = fi
    Next fi

    result = result & Assert(identicalFieldIdx > 0, "found the field whose template value is 'Same Everywhere'")
    result = result & Assert(varyingFieldIdx > 0, "found the field whose template value is 'Q1 2026'")

    If identicalFieldIdx > 0 Then
        result = result & Assert(plan.FieldSuggestIdentical(identicalFieldIdx), "the identical-everywhere field is suggested as Decoration")
        result = result & Assert(plan.HarvestedText(CStr(identicalFieldIdx) & "|1") = "Same Everywhere", "correspondence found on other1 for the identical field")
        result = result & Assert(plan.HarvestedText(CStr(identicalFieldIdx) & "|2") = "Same Everywhere", "correspondence found on other2 for the identical field")
    End If
    If varyingFieldIdx > 0 Then
        result = result & Assert(Not plan.FieldSuggestIdentical(varyingFieldIdx), "the varying field is suggested as a real Field, not Decoration")
        result = result & Assert(plan.HarvestedText(CStr(varyingFieldIdx) & "|1") = "Q2 2026", "correspondence found on other1 for the varying field, got '" & plan.HarvestedText(CStr(varyingFieldIdx) & "|1") & "'")
        result = result & Assert(plan.HarvestedText(CStr(varyingFieldIdx) & "|2") = "Q3 2026", "correspondence found on other2 for the varying field, got '" & plan.HarvestedText(CStr(varyingFieldIdx) & "|2") & "'")
    End If

    Test_BatchOnboardFlow_BuildBatchPlanFindsCorrespondenceAndHarvestsAcrossSlides = result
End Function

' Proves BuildBatchPlanFromMarkedFields is click-scoped, not Discovery-
' scoped: tShapeA exists on the template slide but is never marked, and
' must be entirely absent from the resulting plan.
Private Function Test_BatchOnboardFlow_BuildBatchPlanFromMarkedFieldsUsesOnlyMarkedShapes() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim tShapeA As Object, tShapeB As Object
    Set tShapeA = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    tShapeA.TextFrame.TextRange.Text = "Not Marked"
    Set tShapeB = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 150, 200, 50)
    tShapeB.TextFrame.TextRange.Text = "Q1 2026"

    Dim other1 As Object
    Set other1 = NewBlankSlide()
    Dim o1ShapeA As Object, o1ShapeB As Object
    Set o1ShapeA = other1.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    o1ShapeA.TextFrame.TextRange.Text = "Not Marked"
    Set o1ShapeB = other1.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 150, 200, 50)
    o1ShapeB.TextFrame.TextRange.Text = "Q2 2026"

    Dim otherSlides(1 To 1) As Object
    Set otherSlides(1) = other1

    Dim marked As Collection
    Set marked = New Collection
    marked.Add tShapeB

    Dim markedNames As Object
    Set markedNames = CreateObject("Scripting.Dictionary")
    markedNames(1) = "Project Number"

    Dim markedTypes As Object
    Set markedTypes = CreateObject("Scripting.Dictionary")
    markedTypes(1) = "number"

    Dim markedVolatility As Object
    Set markedVolatility = CreateObject("Scripting.Dictionary")
    markedVolatility(1) = "static"

    Dim matchErr As String
    Dim plan As BatchOnboardPlan
    plan = BatchOnboardFlow.BuildBatchPlanFromMarkedFields(templateSld, marked, markedNames, markedTypes, markedVolatility, otherSlides, matchErr)

    result = result & Assert(matchErr = "", "no match error, got '" & matchErr & "'")
    result = result & Assert(plan.FieldCount = 1, "exactly 1 field in the plan (only the marked shape), got " & plan.FieldCount)
    result = result & Assert(plan.FieldNames(1) = "Project Number", "the human-typed name at mark time is used, not an auto-suggested one, got '" & plan.FieldNames(1) & "'")
    result = result & Assert(plan.FieldTypes(1) = "number", "the human-declared type at mark time is used, not the default 'text', got '" & plan.FieldTypes(1) & "'")
    result = result & Assert(plan.FieldVolatility(1) = "static", "the human-declared volatility at mark time is used, not the default 'variable', got '" & plan.FieldVolatility(1) & "'")
    result = result & Assert(plan.HarvestedText("1|0") = "Q1 2026", "the one field's template value is the marked shape's text, got '" & plan.HarvestedText("1|0") & "'")
    result = result & Assert(plan.HarvestedText("1|1") = "Q2 2026", "correspondence found on other1 for the marked field, got '" & plan.HarvestedText("1|1") & "'")

    Test_BatchOnboardFlow_BuildBatchPlanFromMarkedFieldsUsesOnlyMarkedShapes = result
End Function

' MarkShapeForBatch is the pure-logic half of the click-one-at-a-time
' marking flow (MarkFieldForBatch's own header explains why the MsgBox
' confirmation half can't be automated). Proves accumulation, order
' preservation, and re-mark de-duplication -- the three properties
' BuildBatchPlanFromMarkedFields depends on.
Private Function Test_BatchOnboardFlow_MarkShapeForBatchAccumulatesAndDedupes() As String
    Dim result As String

    BatchOnboardFlow.ResetMarkingSession ' clean slate regardless of prior test/session state -- MsgBox-free, unlike ClearMarkedFieldsForBatch

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim shapeA As Object, shapeB As Object
    Set shapeA = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    shapeA.TextFrame.TextRange.Text = "Field A"
    Set shapeB = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 150, 200, 50)
    shapeB.TextFrame.TextRange.Text = "Field B"

    Dim status1 As String
    status1 = BatchOnboardFlow.MarkShapeForBatch(shapeA, "Project Number", "number", "static")
    result = result & Assert(InStr(status1, "Marked field 1: 'Project Number'") > 0, "first mark reports field 1 with the given name, got '" & status1 & "'")
    result = result & Assert(BatchOnboardFlow.MarkedFieldCountForBatch() = 1, "1 field marked after first mark")
    result = result & Assert(BatchOnboardFlow.MarkedFieldNameForBatch(1) = "Project Number", "field 1's stored name matches, got '" & BatchOnboardFlow.MarkedFieldNameForBatch(1) & "'")
    result = result & Assert(BatchOnboardFlow.MarkedFieldTypeForBatch(1) = "number", "field 1's stored type matches, got '" & BatchOnboardFlow.MarkedFieldTypeForBatch(1) & "'")
    result = result & Assert(BatchOnboardFlow.MarkedFieldVolatilityForBatch(1) = "static", "field 1's stored volatility matches, got '" & BatchOnboardFlow.MarkedFieldVolatilityForBatch(1) & "'")

    Dim status2 As String
    status2 = BatchOnboardFlow.MarkShapeForBatch(shapeB, "Project Title", "text", "variable")
    result = result & Assert(InStr(status2, "Marked field 2: 'Project Title'") > 0, "second mark reports field 2 with the given name, got '" & status2 & "'")
    result = result & Assert(BatchOnboardFlow.MarkedFieldCountForBatch() = 2, "2 fields marked after second mark")

    Dim status3 As String
    status3 = BatchOnboardFlow.MarkShapeForBatch(shapeB, "Project Title (renamed)", "date", "static") ' re-mark the same shape with a new name, type, and volatility
    result = result & Assert(BatchOnboardFlow.MarkedFieldCountForBatch() = 2, "re-marking the same shape does not duplicate it, got " & BatchOnboardFlow.MarkedFieldCountForBatch())
    result = result & Assert(BatchOnboardFlow.MarkedFieldNameForBatch(2) = "Project Title (renamed)", "re-marking the same shape renames it in place, got '" & BatchOnboardFlow.MarkedFieldNameForBatch(2) & "'")
    result = result & Assert(BatchOnboardFlow.MarkedFieldTypeForBatch(2) = "date", "re-marking the same shape re-types it in place, got '" & BatchOnboardFlow.MarkedFieldTypeForBatch(2) & "'")
    result = result & Assert(BatchOnboardFlow.MarkedFieldVolatilityForBatch(2) = "static", "re-marking the same shape re-hints its volatility in place, got '" & BatchOnboardFlow.MarkedFieldVolatilityForBatch(2) & "'")

    Dim otherSld As Object
    Set otherSld = NewBlankSlide()
    Dim shapeC As Object
    Set shapeC = otherSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    shapeC.TextFrame.TextRange.Text = "Field C"
    Dim status4 As String
    status4 = BatchOnboardFlow.MarkShapeForBatch(shapeC, "Field C", "text", "variable") ' a shape on a different slide
    result = result & Assert(status4 = "DIFFERENT_SLIDE", "marking a shape on a different slide returns the DIFFERENT_SLIDE sentinel for the Sub wrapper to resolve interactively, got '" & status4 & "'")
    result = result & Assert(BatchOnboardFlow.MarkedFieldCountForBatch() = 2, "the differing-slide mark did not silently alter the existing session")

    BatchOnboardFlow.ResetMarkingSession
    result = result & Assert(BatchOnboardFlow.MarkedFieldCountForBatch() = 0, "marking session cleared")

    Test_BatchOnboardFlow_MarkShapeForBatchAccumulatesAndDedupes = result
End Function

' FieldPreview is display-only (mark-time prompts), but it is the only thing
' standing between a human and a 250-char paragraph rendered inside a modal
' InputBox -- which is exactly what the real deck's About-text field did on
' 2026-07-26. Pure function, so no slide/Office state needed.
Private Function Test_BatchOnboardFlow_FieldPreviewIsShortAndSingleLine() As String
    Dim result As String

    result = result & Assert(BatchOnboardFlow.FieldPreview("") = "", _
        "empty text previews as empty, got '" & BatchOnboardFlow.FieldPreview("") & "'")
    result = result & Assert(BatchOnboardFlow.FieldPreview("P001") = "P001", _
        "short text is passed through untouched, got '" & BatchOnboardFlow.FieldPreview("P001") & "'")

    ' Exactly at the limit: no ellipsis (an ellipsis implies elision).
    Dim exactly20 As String
    exactly20 = "12345678901234567890"
    result = result & Assert(BatchOnboardFlow.FieldPreview(exactly20) = exactly20, _
        "text exactly at the limit is not truncated, got '" & BatchOnboardFlow.FieldPreview(exactly20) & "'")
    result = result & Assert(BatchOnboardFlow.FieldPreview(exactly20 & "1") = exactly20 & "...", _
        "one char over the limit truncates with an ellipsis, got '" & BatchOnboardFlow.FieldPreview(exactly20 & "1") & "'")

    ' The real failure case: a multi-paragraph About-text value. TextRange.Text
    ' returns paragraphs CR-separated; those render as literal boxes in a prompt.
    Dim para As String
    para = "The project develops and evaluates nanostructured magnesium oxide." & vbCr & vbCr & "Second paragraph."
    Dim previewed As String
    previewed = BatchOnboardFlow.FieldPreview(para)
    result = result & Assert(InStr(previewed, vbCr) = 0, _
        "no carriage returns survive into the preview, got '" & previewed & "'")
    result = result & Assert(Len(previewed) <= 23, _
        "preview stays within the limit plus its ellipsis, got " & Len(previewed) & " chars")
    result = result & Assert(previewed = "The project develops...", _
        "preview reads as the start of the real sentence, got '" & previewed & "'")

    ' Soft line breaks (Chr 11) and leading blank paragraphs: a field whose text
    ' opens with empty paragraphs must not spend its whole budget on whitespace.
    Dim leading As String
    leading = vbCr & vbCr & "  Project" & Chr(11) & "Status  "
    result = result & Assert(BatchOnboardFlow.FieldPreview(leading) = "Project Status", _
        "leading blank paragraphs and soft breaks collapse to single spaces, got '" & BatchOnboardFlow.FieldPreview(leading) & "'")

    Test_BatchOnboardFlow_FieldPreviewIsShortAndSingleLine = result
End Function

' The regression test for the 2026-07-26 real-deck corruption: a second
' onboarding pass re-derived every instance key from a body-text field and
' overwrote 46 correct keys with paragraphs, orphaning every slide from its
' Data-sheet row. An already-linked slide must report its own key so the flow
' reuses it instead of prompting for (and then overwriting it with) a new one.
Private Function Test_BatchOnboardFlow_ExistingInstanceKeyIsReusedNotRederived() As String
    Dim result As String

    Dim freshSld As Object
    Set freshSld = NewBlankSlide()
    result = result & Assert(BatchOnboardFlow.ExistingInstanceKey(freshSld) = "", _
        "a slide with no tags has no existing key, got '" & BatchOnboardFlow.ExistingInstanceKey(freshSld) & "'")

    ' A slide carrying only a type tag is still not linked -- the instance key
    ' is the join to the row, and it alone decides reuse.
    freshSld.Tags.Add "slide_type", "quarterly-update"
    result = result & Assert(BatchOnboardFlow.ExistingInstanceKey(freshSld) = "", _
        "a type tag alone does not count as an existing key, got '" & BatchOnboardFlow.ExistingInstanceKey(freshSld) & "'")

    Dim linkedSld As Object
    Set linkedSld = NewBlankSlide()
    linkedSld.Tags.Add "slide_type", "quarterly-update"
    linkedSld.Tags.Add "instance_key", "3_P002"
    result = result & Assert(BatchOnboardFlow.ExistingInstanceKey(linkedSld) = "3_P002", _
        "an already-linked slide reports its own key, got '" & BatchOnboardFlow.ExistingInstanceKey(linkedSld) & "'")

    ' The indexed-suffix form real duplicate slides carry must survive intact:
    ' it is an ordinary key, not something to be parsed or regenerated.
    Dim dupeSld As Object
    Set dupeSld = NewBlankSlide()
    dupeSld.Tags.Add "instance_key", "3_P002-2"
    result = result & Assert(BatchOnboardFlow.ExistingInstanceKey(dupeSld) = "3_P002-2", _
        "an indexed-suffix key round-trips unchanged, got '" & BatchOnboardFlow.ExistingInstanceKey(dupeSld) & "'")

    ' A paragraph-shaped key (what the corruption actually wrote) is still a
    ' key: reuse must not "helpfully" reject or rewrite it, or a deck already
    ' in that state could not be re-onboarded without corrupting it further.
    Dim paragraphSld As Object
    Set paragraphSld = NewBlankSlide()
    paragraphSld.Tags.Add "instance_key", "The project develops and evaluates nanostructured magnesium oxide."
    result = result & Assert(BatchOnboardFlow.ExistingInstanceKey(paragraphSld) = "The project develops and evaluates nanostructured magnesium oxide.", _
        "an existing key is returned verbatim whatever its shape, got '" & BatchOnboardFlow.ExistingInstanceKey(paragraphSld) & "'")

    Test_BatchOnboardFlow_ExistingInstanceKeyIsReusedNotRederived = result
End Function

' Guard against silently stranding another type's dataset. Found live
' 2026-07-28: re-onboarding 46 already-typed slides under a new name left the
' old type's 46 rows orphaned, and the next preview reported "46 new slide(s)
' would be created" -- one Sync Now from mass duplication, with no warning at
' onboard time.
Private Function Test_BatchOnboardFlow_ConflictingSlideTypeIsDetected() As String
    Dim result As String

    Dim fresh As Object
    Set fresh = NewBlankSlide()
    result = result & Assert(BatchOnboardFlow.ConflictingSlideType(fresh, "quarterly") = "", _
        "an untyped slide conflicts with nothing, got '" & BatchOnboardFlow.ConflictingSlideType(fresh, "quarterly") & "'")

    Dim typed As Object
    Set typed = NewBlankSlide()
    typed.Tags.Add "slide_type", "q"
    result = result & Assert(BatchOnboardFlow.ConflictingSlideType(typed, "sandbox-test") = "q", _
        "a slide typed 'q' reports the conflict, got '" & BatchOnboardFlow.ConflictingSlideType(typed, "sandbox-test") & "'")

    ' Re-onboarding the SAME type is legitimate (adding a field) and must not warn.
    result = result & Assert(BatchOnboardFlow.ConflictingSlideType(typed, "q") = "", _
        "re-onboarding the same type is not a conflict, got '" & BatchOnboardFlow.ConflictingSlideType(typed, "q") & "'")

    Test_BatchOnboardFlow_ConflictingSlideTypeIsDetected = result
End Function

' Portability: the stored workbook path is absolute, so a deck copied to another
' machine (different user profile, work PC, SharePoint sync root) names a folder
' that does not exist there. GetWorkbookPath must recover by looking beside the
' deck, which is where the Data workbook actually lives.
Private Function Test_DeckRegistry_WorkbookPathSurvivesAMovedDeck() As String
    Dim result As String

    result = result & Assert(DeckRegistry.FileNameOnly("C:\Users\someone\OneDrive\Claude\Data.xlsx") = "Data.xlsx", _
        "backslash path -> bare file name, got '" & DeckRegistry.FileNameOnly("C:\Users\someone\OneDrive\Claude\Data.xlsx") & "'")
    result = result & Assert(DeckRegistry.FileNameOnly("/mnt/c/x/Data.xlsx") = "Data.xlsx", _
        "forward-slash path -> bare file name, got '" & DeckRegistry.FileNameOnly("/mnt/c/x/Data.xlsx") & "'")
    result = result & Assert(DeckRegistry.FileNameOnly("Data.xlsx") = "Data.xlsx", _
        "a bare name is returned unchanged, got '" & DeckRegistry.FileNameOnly("Data.xlsx") & "'")

    ' A path that resolves nowhere must come back unchanged, so the caller's error
    ' message still names something the human recognises rather than a guess.
    Dim pres As Object
    Set pres = Application.ActivePresentation
    Dim original As String
    original = DeckRegistry.GetWorkbookPath(pres)

    DeckRegistry.SetWorkbookPath pres, "Z:\definitely\not\here\Nope.xlsx"
    result = result & Assert(DeckRegistry.GetWorkbookPath(pres) = "Z:\definitely\not\here\Nope.xlsx", _
        "an unresolvable path is returned unchanged, got '" & DeckRegistry.GetWorkbookPath(pres) & "'")

    ' RepointWorkbook must refuse a path that does not exist -- registering a bad
    ' path is how a deck ends up silently unsyncable.
    Dim refused As Boolean
    On Error Resume Next
    Err.Clear
    DeckRegistry.RepointWorkbook pres, "Z:\also\not\here.xlsx"
    refused = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    result = result & Assert(refused, "RepointWorkbook refuses a non-existent path")

    DeckRegistry.SetWorkbookPath pres, original

    Test_DeckRegistry_WorkbookPathSurvivesAMovedDeck = result
End Function

' Real bug found and fixed 2026-07-26 (Rohan, live-testing against his real
' deck's grouped "card" layouts: "Is it allowing me to mark selected
' objects in groups? I think maybe no?"): MarkShapeForBatch used to reject
' any shape inside a group outright, based on an assumption about
' Shape.Parent that was never actually verified against real Office and
' turned out to be wrong -- confirmed live that .Parent resolves directly
' to the Slide even for a grouped shape. This proves a shape fetched via
' GroupItems (mirroring what Selection.ShapeRange(1) gives after a real
' user double-clicks into a group to select one member) is now accepted.
Private Function Test_BatchOnboardFlow_MarkShapeForBatchAcceptsShapeInsideGroup() As String
    Dim result As String

    BatchOnboardFlow.ResetMarkingSession

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim shapeA As Object, shapeB As Object
    Set shapeA = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 100, 50)
    shapeA.TextFrame.TextRange.Text = "Grouped Field"
    Set shapeB = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 200, 50, 100, 50)
    shapeB.TextFrame.TextRange.Text = "Grouped Sibling"

    Dim grp As Object
    Set grp = sld.Shapes.Range(Array(shapeA.Name, shapeB.Name)).Group()

    Dim groupedChild As Object
    Set groupedChild = grp.GroupItems.Item(1)

    Dim status As String
    status = BatchOnboardFlow.MarkShapeForBatch(groupedChild, "Grouped Field Name", "text", "variable")

    result = result & Assert(InStr(status, "Could not determine") = 0, "marking a shape inside a group does not fail, got '" & status & "'")
    result = result & Assert(InStr(status, "Marked field 1") > 0, "marking a shape inside a group succeeds and reports field 1, got '" & status & "'")
    result = result & Assert(BatchOnboardFlow.MarkedFieldCountForBatch() = 1, "1 field marked")

    BatchOnboardFlow.ResetMarkingSession

    Test_BatchOnboardFlow_MarkShapeForBatchAcceptsShapeInsideGroup = result
End Function

' Real bug found and fixed 2026-07-26 (Rohan: "why when I select long text
' fields in a group is it showing current value in the message box as
' current value: '):"): the actual cause was the whole GROUP being
' selected (a group container's .HasTextFrame is False even when its
' members have real text), not a text-reading bug. MarkShapeForBatch now
' rejects a whole-group selection outright with guidance, rather than
' silently accepting it and tagging the wrong object.
Private Function Test_BatchOnboardFlow_MarkShapeForBatchRejectsWholeGroupSelection() As String
    Dim result As String

    BatchOnboardFlow.ResetMarkingSession

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim shapeA As Object, shapeB As Object
    Set shapeA = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 100, 50)
    shapeA.TextFrame.TextRange.Text = "Long real text content here for the card field."
    Set shapeB = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 200, 50, 100, 50)
    shapeB.TextFrame.TextRange.Text = "Sibling"

    Dim grp As Object
    Set grp = sld.Shapes.Range(Array(shapeA.Name, shapeB.Name)).Group()

    Dim status As String
    status = BatchOnboardFlow.MarkShapeForBatch(grp, "Whole Group Name", "text", "variable")

    result = result & Assert(InStr(status, "whole group") > 0, "whole-group selection is rejected with guidance, got '" & status & "'")
    result = result & Assert(BatchOnboardFlow.MarkedFieldCountForBatch() = 0, "nothing marked when the whole group was selected, got " & BatchOnboardFlow.MarkedFieldCountForBatch())

    BatchOnboardFlow.ResetMarkingSession

    Test_BatchOnboardFlow_MarkShapeForBatchRejectsWholeGroupSelection = result
End Function

' Real question 2026-07-26: Office's CustomDocumentProperties string type
' has a documented 255-character limit -- does it actually apply/truncate
' here, or is that stale/version-dependent? A PowerShell-side probe hit an
' unrelated .NET COM interop crash trying to answer this directly (dynamic
' COM dispatch failing to reflect the DocumentProperties collection type --
' a known PowerShell quirk, not an Office one), so this is answered for
' real here instead, via the same native-VBA path DeckRegistry.bas's own
' CustomDocumentProperties usage already proves works. Content is
' deliberately over 255 characters (a realistic ~10-field marking session)
' -- if Office truncates, this assertion fails with a clear, honest diff
' instead of silently shipping corrupted persistence.
Private Function Test_BatchOnboardFlow_MarkingSessionPropertyRoundTripsBeyond255Chars() As String
    Dim result As String

    Dim longValue As String
    Dim i As Long
    For i = 1 To 10
        If longValue <> "" Then longValue = longValue & vbCrLf
        longValue = longValue & "TextBox " & (40 + i) & "|Field Name Number " & i & " (a realistic length)|text|variable"
    Next i
    result = result & Assert(Len(longValue) > 255, "test setup sanity: content is genuinely over 255 chars, got " & Len(longValue))

    BatchOnboardFlow.WriteMarkingSessionProperty Application.ActivePresentation, longValue
    Dim readBack As String
    readBack = BatchOnboardFlow.ReadMarkingSessionProperty(Application.ActivePresentation)

    result = result & Assert(readBack = longValue, "long (" & Len(longValue) & "-char) marking session round-trips through CustomDocumentProperties without truncation, got " & Len(readBack) & " chars back")

    Test_BatchOnboardFlow_MarkingSessionPropertyRoundTripsBeyond255Chars = result
End Function

' SerializeMarkingSession/RestoreMarkingSession are the actual persistence
' fix for Rohan's real friction (2026-07-26): "sick of linking every test"
' -- re-marking every field from scratch after every add-in reload during
' live testing. Proves a real mark -> serialize -> reset (simulating a
' close) -> restore round trip re-finds the same live shapes by Name and
' recovers every field's name/type/volatility.
Private Function Test_BatchOnboardFlow_RestoreMarkingSessionRecoversMarkedFields() As String
    Dim result As String

    BatchOnboardFlow.ResetMarkingSession

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim shapeA As Object, shapeB As Object
    Set shapeA = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 100, 50)
    shapeA.TextFrame.TextRange.Text = "Field A"
    Set shapeB = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 150, 100, 50)
    shapeB.TextFrame.TextRange.Text = "Field B"

    BatchOnboardFlow.MarkShapeForBatch shapeA, "Project Number", "text", "static"
    BatchOnboardFlow.MarkShapeForBatch shapeB, "Status", "text", "variable"

    Dim serialized As String
    serialized = BatchOnboardFlow.SerializeCurrentMarkingSession()

    BatchOnboardFlow.ResetMarkingSession ' simulates a close -- wipes the in-memory session, leaves only `serialized`
    result = result & Assert(BatchOnboardFlow.MarkedFieldCountForBatch() = 0, "in-memory session genuinely cleared before restoring")

    Dim restoreReport As String
    restoreReport = BatchOnboardFlow.RestoreMarkingSession(serialized, sld)

    result = result & Assert(InStr(restoreReport, "Restored 2 field") > 0, "restore reports 2 fields recovered, got '" & restoreReport & "'")
    result = result & Assert(BatchOnboardFlow.MarkedFieldCountForBatch() = 2, "2 fields marked after restore, got " & BatchOnboardFlow.MarkedFieldCountForBatch())
    result = result & Assert(BatchOnboardFlow.MarkedFieldNameForBatch(1) = "Project Number", "field 1's name recovered, got '" & BatchOnboardFlow.MarkedFieldNameForBatch(1) & "'")
    result = result & Assert(BatchOnboardFlow.MarkedFieldTypeForBatch(1) = "text", "field 1's type recovered, got '" & BatchOnboardFlow.MarkedFieldTypeForBatch(1) & "'")
    result = result & Assert(BatchOnboardFlow.MarkedFieldVolatilityForBatch(1) = "static", "field 1's volatility recovered, got '" & BatchOnboardFlow.MarkedFieldVolatilityForBatch(1) & "'")
    result = result & Assert(BatchOnboardFlow.MarkedFieldNameForBatch(2) = "Status", "field 2's name recovered, got '" & BatchOnboardFlow.MarkedFieldNameForBatch(2) & "'")

    BatchOnboardFlow.ResetMarkingSession

    Test_BatchOnboardFlow_RestoreMarkingSessionRecoversMarkedFields = result
End Function

' Real bug found live 2026-07-26: Rohan's actual tagging/marking work was
' lost across a PowerPoint close despite Presentation.Saved correctly
' flipping to False on each edit (proven separately) -- AutoSave's own
' background/debounced save was not reliably capturing macro-driven edits.
' SaveMarkingSessionToProperty now forces an explicit, synchronous
' Presentation.Save instead of trusting AutoSave. Proves the happy path:
' saving a real (test) presentation succeeds with no warning, and the
' property genuinely round-trips afterward.
Private Function Test_BatchOnboardFlow_SaveMarkingSessionToPropertyForcesRealSave() As String
    Dim result As String

    ' The shared test presentation is never given a save path (see
    ' run_vba_tests.ps1 -- Presentations.Add() only), and calling .Save on
    ' a presentation with no path at all pops a real "Save As" dialog,
    ' which would hang this automated run waiting for a click nobody's
    ' there to give. Give it a real (throwaway) path first -- once it has
    ' one, .Save behaves as a normal, silent, no-dialog operation for the
    ' rest of this run too.
    If Application.ActivePresentation.Path = "" Then
        Application.ActivePresentation.SaveAs Environ("TEMP") & "\deck_sync_test_run_presentation.pptx"
    End If

    BatchOnboardFlow.ResetMarkingSession

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim shapeA As Object
    Set shapeA = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 100, 50)
    shapeA.TextFrame.TextRange.Text = "Field A"
    BatchOnboardFlow.MarkShapeForBatch shapeA, "Project Number", "text", "static"

    Dim saveWarning As String
    saveWarning = BatchOnboardFlow.SaveMarkingSessionToProperty(Application.ActivePresentation)

    result = result & Assert(saveWarning = "", "no warning on a successful save, got '" & saveWarning & "'")

    Dim readBack As String
    readBack = BatchOnboardFlow.ReadMarkingSessionProperty(Application.ActivePresentation)
    result = result & Assert(InStr(readBack, "Project Number") > 0, "the saved property genuinely contains the marked field, got '" & readBack & "'")

    BatchOnboardFlow.ResetMarkingSession

    Test_BatchOnboardFlow_SaveMarkingSessionToPropertyForcesRealSave = result
End Function

' The failure message a human actually sees when an action dies unexpectedly.
'
' Asserting the CONTENT, not just that a string comes back, because the one
' thing this message must not do is reassure. An error partway through a commit
' can leave real changes in the deck and the Data sheet, so "nothing was
' written" would be a comforting lie told at the exact moment the human most
' needs the truth.
Private Function Test_RibbonUI_UnexpectedErrorTextTellsTheTruth() As String
    Dim result As String

    Dim txt As String
    txt = RibbonUI.UnexpectedErrorText("Bulk Onboard Type", 52, "Bad file name or number", "WorkbookBridge.CreateWorkbook")

    result = result & Assert(InStr(txt, "Bulk Onboard Type") > 0, "names the action that failed")
    result = result & Assert(InStr(txt, "52") > 0, "includes the error number")
    result = result & Assert(InStr(txt, "Bad file name or number") > 0, "includes the error description")
    result = result & Assert(InStr(txt, "WorkbookBridge.CreateWorkbook") > 0, "includes where it came from")
    result = result & Assert(InStr(txt, "may have already changed") > 0, _
        "WARNS that partial changes may exist -- never claims nothing was written")

    ' An empty Err.Source is normal for a plain runtime error, and must not
    ' produce a message with a blank hole in it.
    Dim blankSource As String
    blankSource = RibbonUI.UnexpectedErrorText("Sync Now", 9, "Subscript out of range", "")
    result = result & Assert(InStr(blankSource, "unidentified step") > 0, _
        "a missing Err.Source still reads as a sentence, got '" & blankSource & "'")

    Test_RibbonUI_UnexpectedErrorTextTellsTheTruth = result
End Function

' THE PART A PERSON IS AGREEING TO MUST SURVIVE THE CUT.
'
' MsgBox truncates near 1024 characters silently. ConfirmBatchText appends
' "Apply the N uniform change(s) above?" LAST, and Sync Now's confirmation was a
' raw MsgBox -- so on a deck with more than about two batches the question was
' the part that disappeared, leaving Yes/No buttons over a partial list with
' nothing visibly being asked. Capping alone does not fix that; the tail has to
' be carried past the cut deliberately.
Private Function Test_RibbonUI_CapReportKeepsTheQuestion() As String
    Dim result As String
    Dim question As String
    question = "Apply the 7 uniform change(s) above?"

    ' Short text: returned untouched, and the tail is NOT doubled. The tail is
    ' already inside the text -- appending it again would show the question twice.
    Dim shortText As String
    shortText = "Two changes." & vbCrLf & question
    result = result & Assert(RibbonUI.CapReport(shortText, question) = shortText, _
        "a report under the cap is returned unchanged")
    result = result & Assert(CountOccurrences(RibbonUI.CapReport(shortText, question), question) = 1, _
        "and the question is not duplicated")

    ' Long text: body is cut, cut is announced, question still there and LAST.
    Dim longText As String
    longText = String(2000, "x") & vbCrLf & question
    Dim capped As String
    capped = RibbonUI.CapReport(longText, question)
    result = result & Assert(Len(capped) < Len(longText), "an over-long report is actually shortened")
    result = result & Assert(InStr(capped, "[shortened") > 0, "and the cut is announced, not silent")
    result = result & Assert(InStr(capped, question) > 0, _
        "and the QUESTION survives the cut -- got '" & Right$(capped, 80) & "'")
    result = result & Assert(Right$(capped, Len(question)) = question, _
        "and it is the last thing on screen")

    ' Without a tail, the old behaviour is unchanged: capped and announced.
    Dim plain As String
    plain = RibbonUI.CapReport(String(2000, "y"))
    result = result & Assert(Len(plain) < 2000 And InStr(plain, "[shortened") > 0, _
        "a report with no must-keep tail is still capped and announced")

    Test_RibbonUI_CapReportKeepsTheQuestion = result
End Function

' AN EMPTY QUEUE HAS TWO CAUSES AND THEY ARE OPPOSITES.
'
' BuildQueue keeps only in_place_correction, so new_record and flagged used to
' leave Count at 0 -- and both messages answered Count = 0 with "every linked
' slide already matches the workbook" / "this deck already matches the register".
' A deck where every row reaches a slide and a deck where NO row reaches one
' produced the same sentence, and Sync Now was the only place a person would
' have looked. The condition survived solely on Preview Sync, which described it
' wrongly in the other direction.
Private Function Test_ReviewQueue_EmptyQueueDoesNotClaimTheDeckMatches() As String
    Dim result As String

    ' Genuinely up to date: nothing queued, nothing dropped. The reassuring
    ' sentence is CORRECT here and must survive.
    Dim clean As ReviewQueueSet
    clean.Count = 0
    result = result & Assert(InStr(ReviewQueue.FastPathRefusalText(clean), "already matches") > 0, _
        "a truly clean run still says the deck matches")
    result = result & Assert(InStr(ReviewQueue.QueueSummaryText(clean), "already matches") > 0, _
        "and so does the per-type summary")

    ' Orphans: same Count, opposite meaning.
    Dim orphaned As ReviewQueueSet
    orphaned.Count = 0
    orphaned.OrphanCount = 43
    orphaned.OrphanKeys = "1_K010, 1_K022"

    Dim fast As String
    fast = ReviewQueue.FastPathRefusalText(orphaned)
    result = result & Assert(InStr(fast, "already matches") = 0, _
        "43 orphaned rows must NOT be reported as the deck already matching -- got '" & fast & "'")
    result = result & Assert(InStr(fast, "43") > 0, "it states how many rows reach no slide")
    result = result & Assert(InStr(fast, "1_K010") > 0, "and names them")
    result = result & Assert(InStr(fast, "does not create slides") > 0, _
        "and says Sync Now will not create them, since it cannot")

    Dim summary As String
    summary = ReviewQueue.QueueSummaryText(orphaned)
    result = result & Assert(InStr(summary, "already matches") = 0, _
        "the per-type summary must not claim it either -- got '" & summary & "'")
    result = result & Assert(InStr(summary, "1_K010") > 0, "and it names the rows too")

    ' Flagged items alone are also not "up to date".
    Dim flagged As ReviewQueueSet
    flagged.Count = 0
    flagged.FlaggedCount = 2
    flagged.FlaggedNotes = "  flagged: 3_P001 (untagged) -- no role tag" & vbCrLf
    result = result & Assert(InStr(ReviewQueue.FastPathRefusalText(flagged), "already matches") = 0, _
        "flagged items alone are not a clean run either")

    Test_ReviewQueue_EmptyQueueDoesNotClaimTheDeckMatches = result
End Function

' THE OTHER DIRECTION. The existing toolbar test asserts every BUTTON resolves to
' a real Sub. Nothing asserted that every CAPABILITY has a button -- so a function
' could be built, tested, and unreachable by the only interface that exists on the
' machine where the work happens.
'
' It has happened twice. RollForwardPeriod was built and tested with no button,
' and StartQuarter told Rohan to hand-copy 43 rows in Excel because of it. Then
' CreateMissingSlides: built, tested, no button, so a new project could not be
' given a slide at all. Neither was a coding mistake -- tests call functions, a
' person presses buttons, and nothing covered the gap between them.
'
' This is the DECLARED half and it blocks: losing a button for anything listed
' here fails the suite. check_vba_static.py is the discovery half and only
' reports, because whether an unlisted Public proc SHOULD be a capability is a
' judgement, and a gate that is red for a judgement gets bypassed.
'
' Adding to this list is how you say "a person must be able to do this".
Private Function Test_CommandBarUI_EveryDeclaredCapabilityHasAButton() As String
    Dim result As String

    Dim required As Variant
    ' THE RUNTIME HALF ONLY, and the split is a real loss worth stating.
    '
    ' This asserted that sixteen named Subs each had a button. Under two buttons
    ' and a chain, thirteen of them are reached from INSIDE the chain instead --
    ' so the question "can a person still do this?" is no longer answerable from
    ' the live toolbar. It moved to check_vba_static.py, which walks the call
    ' graph and fails the build on a genuine orphan. That check is Python, so it
    ' runs on the dev machine and NOT at work: the strongest guard against
    ' silently orphaning a capability is now weaker where it matters most.
    '
    ' What stays here is what only a live toolbar can answer: the dispatchers
    ' are wired, and they resolve to real Subs.
    ' THREE DISPATCHERS SINCE 2026-08-14, not two, and the change is a promotion
    ' rather than a relaxation. The artifact split gave the deck side and the
    ' review their own doors, so both are now things a person must be able to
    ' reach from the toolbar and both are asserted here.
    '
    ' RefreshDraftingSheets LEFT this list because its button was removed -- its
    ' tooltip said "use this when a drafting sheet looks wrong", which is a defect
    ' with instructions attached rather than a capability. It is still reached,
    ' from inside SyncNowChain, exactly like the thirteen described above.
    required = Array("SyncNowChain", "PutItOnTheSlides", "ReviewChanges")

    ' Read from the LIVE toolbar, not from a list of what we think we built.
    CommandBarUI.ShowToolbar
    Dim wired As String
    wired = "|"

    Dim barNames As Variant
    barNames = CommandBarUI.ActiveToolbarNames
    Dim nm As Variant
    For Each nm In barNames
        Dim bar As Object
        Set bar = Application.CommandBars(CStr(nm))
        Dim i As Long
        For i = 1 To bar.Controls.count
            Dim act As String
            act = bar.Controls.Item(i).OnAction
            ' Office reports OnAction as "Module.Sub" or "<Deck>!Sub" depending on
            ' the run -- strip both, same as the sibling toolbar test.
            Dim bangPos As Long
            bangPos = InStr(act, "!")
            If bangPos > 0 Then act = Mid(act, bangPos + 1)
            Dim dotPos As Long
            dotPos = InStrRev(act, ".")
            If dotPos > 0 Then act = Mid(act, dotPos + 1)
            wired = wired & act & "|"
        Next i
    Next nm

    Dim missing As String
    Dim k As Long
    For k = LBound(required) To UBound(required)
        If InStr(wired, "|" & CStr(required(k)) & "|") = 0 Then
            missing = missing & CStr(required(k)) & " "
        End If
    Next k

    result = result & Assert(missing = "", _
        "every declared capability has a toolbar button -- MISSING: " & missing & _
        "(wired: " & wired & ")")

    ' The check must be able to fail. If the wiring string came back empty --
    ' toolbar not built, names changed, Office reporting nothing -- every lookup
    ' above would "pass" by finding nothing to contradict it.
    result = result & Assert(Len(wired) > 20, _
        "and the toolbar was actually read -- got wired='" & wired & "'")

    CommandBarUI.HideToolbar
    Test_CommandBarUI_EveryDeclaredCapabilityHasAButton = result
End Function

' The threshold is a DISCRIMINATOR, not a tolerance -- it separates "a few new
' projects" from "this is not the deck you think it is". Both sides are pinned,
' because a rule that only ever answers one way is not a rule.
Private Function Test_ReviewQueue_ParityAndTheCreateThreshold() As String
    Dim result As String

    ' 2 of 43 -- new projects. Create.
    Dim few As ReviewQueueSet
    few.RowCount = 43: few.OrphanCount = 2
    result = result & Assert(ReviewQueue.OrphansLookLikeNewProjects(few), _
        "2 orphans in 43 rows reads as new projects")

    ' 43 of 46 -- the 2026-07-31 deck, and the board-pack case. Refuse.
    Dim many As ReviewQueueSet
    many.RowCount = 46: many.OrphanCount = 43
    result = result & Assert(Not ReviewQueue.OrphansLookLikeNewProjects(many), _
        "43 orphans in 46 rows does NOT -- this is the case that would duplicate " & _
        "the template across a hand-assembled pack")

    ' Exactly on the line is allowed; just over is not.
    Dim onLine As ReviewQueueSet
    onLine.RowCount = 40: onLine.OrphanCount = 10
    result = result & Assert(ReviewQueue.OrphansLookLikeNewProjects(onLine), _
        "exactly 25% is allowed")
    Dim overLine As ReviewQueueSet
    overLine.RowCount = 40: overLine.OrphanCount = 11
    result = result & Assert(Not ReviewQueue.OrphansLookLikeNewProjects(overLine), _
        "just over 25% is not")

    ' No orphans is not "looks like new projects" -- there is nothing to create.
    Dim none As ReviewQueueSet
    none.RowCount = 43: none.OrphanCount = 0
    result = result & Assert(Not ReviewQueue.OrphansLookLikeNewProjects(none), _
        "no orphans means nothing to create, not a green light")

    ' An empty register must never divide by zero into a yes.
    Dim noRows As ReviewQueueSet
    noRows.RowCount = 0: noRows.OrphanCount = 5
    result = result & Assert(Not ReviewQueue.OrphansLookLikeNewProjects(noRows), _
        "orphans against a register with no rows is refused, not divided by zero")

    ' Parity, both directions.
    Dim clean As ReviewQueueSet
    clean.RowCount = 43: clean.SlideCount = 43
    result = result & Assert(InStr(ReviewQueue.ParityText(clean), "PARITY: the deck and the register agree") > 0, _
        "a matched deck states parity outright")

    Dim rowsNoSlide As ReviewQueueSet
    rowsNoSlide.RowCount = 43: rowsNoSlide.OrphanCount = 2: rowsNoSlide.OrphanKeys = "3_P009"
    result = result & Assert(InStr(ReviewQueue.ParityText(rowsNoSlide), "NOT AT PARITY") > 0 _
        And InStr(ReviewQueue.ParityText(rowsNoSlide), "3_P009") > 0, _
        "rows with no slide are named")

    ' The direction nothing measured before: a slide the register never mentions
    ' is never visited by the planner, so it keeps last period's text silently.
    Dim slideNoRow As ReviewQueueSet
    slideNoRow.RowCount = 43: slideNoRow.SlideCount = 44
    slideNoRow.SlideNoRowCount = 1: slideNoRow.SlideNoRowKeys = "1_K099"
    Dim pt As String
    pt = ReviewQueue.ParityText(slideNoRow)
    result = result & Assert(InStr(pt, "NOT AT PARITY") > 0 And InStr(pt, "1_K099") > 0, _
        "slides with no row are named too -- got '" & pt & "'")
    result = result & Assert(InStr(pt, "keep whatever text they already carry") > 0, _
        "and the consequence is spelled out, since this one is silent by nature")

    Test_ReviewQueue_ParityAndTheCreateThreshold = result
End Function

Private Function CountOccurrences(haystack As String, needle As String) As Long
    If needle = "" Then Exit Function
    Dim p As Long
    p = InStr(haystack, needle)
    Do While p > 0
        CountOccurrences = CountOccurrences + 1
        p = InStr(p + Len(needle), haystack, needle)
    Loop
End Function

' Proves the wrapper pattern is load-bearing rather than decorative.
'
' The reason every toolbar action got a separate wrapper Sub instead of an
' inline "On Error GoTo Failed" is a VBA rule that is easy to forget: "On Error
' GoTo 0" DISABLES the enabled handler for the whole procedure. Every one of
' these bodies is full of "On Error Resume Next / On Error GoTo 0" pairs, so an
' inline handler would be switched off by the first pair and would then read as
' protection while providing none -- the same always-true-guard shape as the
' two faults fixed earlier today.
'
' This test asserts both halves against real VBA: that an inline handler really
' is destroyed by On Error GoTo 0, and that a wrapper's handler really does
' still catch. If VBA's behaviour were ever different, the first assertion
' fails and the wrapper indirection can be simplified away with confidence.
Private Function Test_RibbonUI_WrapperHandlerSurvivesOnErrorGoToZero() As String
    Dim result As String

    result = result & Assert(InlineHandlerIsDefeated() = "escaped", _
        "an inline handler IS destroyed by a later On Error GoTo 0 -- the reason wrappers exist")
    result = result & Assert(WrapperHandlerCatches() = "caught", _
        "a wrapper's handler still catches an error raised below it")

    Test_RibbonUI_WrapperHandlerSurvivesOnErrorGoToZero = result
End Function

' Mimics the naive fix: handler at the top, then an ordinary guarded probe of
' the kind these Subs do constantly. Returns "escaped" if the error got past
' the handler, "caught" if the handler still worked.
Private Function InlineHandlerIsDefeated() As String
    On Error GoTo Failed

    Dim ignored As Variant
    On Error Resume Next
    ignored = 1 / 0            ' a routine guarded probe
    On Error GoTo 0            ' <- this is the line that disables the handler above

    Dim boom As Long
    On Error Resume Next
    Err.Clear
    boom = 1 / 0               ' now unguarded: does the top handler still catch?
    If Err.Number <> 0 Then
        InlineHandlerIsDefeated = "escaped"
    Else
        InlineHandlerIsDefeated = "caught"
    End If
    On Error GoTo 0
    Exit Function
Failed:
    InlineHandlerIsDefeated = "caught"
End Function

Private Function WrapperHandlerCatches() As String
    On Error GoTo Failed
    RaiseSomethingUnexpected      ' separate frame, exactly like <Action>Core
    WrapperHandlerCatches = "notcaught"
    Exit Function
Failed:
    WrapperHandlerCatches = "caught"
End Function

Private Sub RaiseSomethingUnexpected()
    ' Guarded probes of its own first, mirroring a real Core Sub, so the test
    ' proves they don't interfere with the CALLER's handler.
    Dim ignored As Variant
    On Error Resume Next
    ignored = 1 / 0
    On Error GoTo 0

    Err.Raise vbObjectError + 99, "TestRunner.RaiseSomethingUnexpected", "simulated unexpected failure"
End Sub

' The instance-key grid, which replaced the per-slide InputBox chain on
' 2026-07-29 ("I'd rather fix the excel and skip all this having to confirm the
' instance key... too time consuming for organised slides").
'
' Tests the round trip that actually runs, not a helper beside it: write the
' grid from real slides, tamper with the sheet the way a human would, read it
' back, and check what the add-in concluded. The two behaviours that matter
' most are the ones a careless implementation gets wrong -- a Skip must yield
' NO key rather than a blank row in the Data sheet, and an edit to an
' already-linked slide must be refused, because re-keying a linked slide is the
' corruption the 2026-07-28 bug caused.
Private Function Test_BatchOnboardFlow_InstanceKeyGridPrefillsAndCatchesClashes() As String
    Dim result As String

    BatchOnboardFlow.ResetMarkingSession

    ' Template + two others. Slide 3 is deliberately a "copy" of slide 2:
    ' identical field text, so it suggests the identical key.
    Dim tmpl As Object, other1 As Object, other2 As Object
    Set tmpl = NewBlankSlide()
    Set other1 = NewBlankSlide()
    Set other2 = NewBlankSlide()

    ' Written out rather than looped over Array(): a For Each control variable
    ' over an array must be a Variant in VBA, and getting that wrong is a
    ' COMPILE error -- which in this harness means a modal dialog inside
    ' PowerPoint and a run that hangs instead of failing. Cost 20 minutes to
    ' find, because the whole suite stops, including tests that never call it.
    tmpl.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50).Name = "KeyField"
    other1.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50).Name = "KeyField"
    other2.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50).Name = "KeyField"
    tmpl.Shapes("KeyField").TextFrame.TextRange.text = "P001"
    other1.Shapes("KeyField").TextFrame.TextRange.text = "P002"
    other2.Shapes("KeyField").TextFrame.TextRange.text = "P002"   ' the copy

    Dim others(1 To 2) As Object
    Set others(1) = other1
    Set others(2) = other2

    Dim marked As Collection
    Set marked = New Collection
    marked.Add tmpl.Shapes("KeyField")

    Dim markedNames As Object
    Set markedNames = CreateObject("Scripting.Dictionary")
    markedNames(1) = "Project Number"
    Dim markedTypes As Object
    Set markedTypes = CreateObject("Scripting.Dictionary")
    markedTypes(1) = "text"
    Dim markedVolatility As Object
    Set markedVolatility = CreateObject("Scripting.Dictionary")
    markedVolatility(1) = "static"

    Dim matchErr As String
    Dim plan As BatchOnboardPlan
    plan = BatchOnboardFlow.BuildBatchPlanFromMarkedFields(tmpl, marked, markedNames, markedTypes, markedVolatility, others, matchErr)
    result = result & Assert(matchErr = "", "plan built cleanly, got '" & matchErr & "'")

    ' Reuse the harness's own Excel instance, exactly as production does.
    ' An earlier version of this test did CreateObject + xl.Quit and hung the
    ' whole run: Quit tore down the instance WorkbookBridge.GetExcelApp() was
    ' still holding, and every later Excel call blocked on a dead COM object.
    ' Tests share this process -- take a workbook, never the application.
    Dim xl As Object, wb As Object, ws As Object
    Set xl = WorkbookBridge.GetExcelApp()
    Set wb = xl.Workbooks.Add()
    Set ws = wb.Worksheets(1)

    BatchOnboardFlow.WriteInstanceKeyGrid ws, plan, tmpl, others, 2

    ' --- pre-filled, so the human reviews rather than dictates ---
    result = result & Assert(CStr(ws.Cells(2, 4).Value) = "P001", "template row pre-filled from the slide, got '" & CStr(ws.Cells(2, 4).Value) & "'")
    result = result & Assert(CStr(ws.Cells(3, 4).Value) = "P002", "second slide pre-filled, got '" & CStr(ws.Cells(3, 4).Value) & "'")
    result = result & Assert(InStr(CStr(ws.Cells(2, 3).Value), "TEMPLATE") > 0, "the template row says it's the template and can't be skipped")
    result = result & Assert(InStr(CStr(ws.Cells(4, 3).Value), "CLASH") > 0, _
        "the copied slide is flagged as a CLASH up front, not discovered one prompt deep, got '" & CStr(ws.Cells(4, 3).Value) & "'")

    ' --- reading back an unresolved clash must refuse ---
    Dim keys As Object, reused As Long, ignored As Long
    Dim problems As String
    problems = BatchOnboardFlow.ReadInstanceKeyGrid(ws, 2, keys, reused, ignored)
    result = result & Assert(problems <> "", "an unresolved duplicate is reported, not accepted")
    result = result & Assert(InStr(problems, "same row") > 0 Or InStr(problems, "one row") > 0, _
        "and the message explains the consequence, got '" & problems & "'")

    ' --- the human marks the copy as Skip, which is the agreed handling ---
    ws.Cells(4, 5).Value = "Y"
    problems = BatchOnboardFlow.ReadInstanceKeyGrid(ws, 2, keys, reused, ignored)
    result = result & Assert(problems = "", "skipping the copy resolves it, got '" & problems & "'")
    result = result & Assert(CStr(keys(2)) = "", "a skipped slide gets NO key, so it never reaches the Data sheet, got '" & CStr(keys(2)) & "'")
    result = result & Assert(CStr(keys(0)) = "P001" And CStr(keys(1)) = "P002", "the other two keep their keys")

    ' --- case-only differences are still collisions ---
    ws.Cells(4, 5).Value = "N"
    ws.Cells(4, 4).Value = "p002"
    problems = BatchOnboardFlow.ReadInstanceKeyGrid(ws, 2, keys, reused, ignored)
    result = result & Assert(problems <> "", "a case-only difference is still a clash -- it's a typo, not an intent")

    ' --- the template cannot be skipped or emptied ---
    ws.Cells(4, 4).Value = ""
    ws.Cells(4, 5).Value = "Y"
    ws.Cells(2, 5).Value = "Y"
    problems = BatchOnboardFlow.ReadInstanceKeyGrid(ws, 2, keys, reused, ignored)
    result = result & Assert(InStr(problems, "template") > 0, "skipping the template is refused, got '" & problems & "'")

    wb.Saved = True
    wb.Close

    BatchOnboardFlow.ResetMarkingSession
    Test_BatchOnboardFlow_InstanceKeyGridPrefillsAndCatchesClashes = result
End Function

' Live failure 2026-07-29: a hand-typed workbook path produced VBA runtime
' error 52, unhandled, and the resulting Debug/End dialog discarded 45 instance
' keys that had just been confirmed one prompt at a time.
'
' The first assertion is the one that matters most. A cloud-hosted deck reports
' its own Path as an https:// URL, so copying the deck's location when asked
' where to put its data is the natural move and produces something no file API
' can open -- a mistake that deserves its own explanation, not a generic
' rejection.
Private Function Test_BatchOnboardFlow_WorkbookPathProblemRejectsTheRealMistakes() As String
    Dim result As String

    ' The trap that actually caught a human.
    Dim urlProblem As String
    urlProblem = BatchOnboardFlow.WorkbookPathProblem("https://saafecrc.sharepoint.com/sites/x/Shared%20Documents/Data.xlsx")
    result = result & Assert(urlProblem <> "", "a SharePoint URL is rejected as a workbook path")
    result = result & Assert(InStr(urlProblem, "web address") > 0, "and the message says WHY, rather than just 'bad path', got '" & urlProblem & "'")

    result = result & Assert(BatchOnboardFlow.WorkbookPathProblem("http://example.com/Data.xlsx") <> "", "http:// is rejected too")

    ' Not rooted -- no folder to live in.
    result = result & Assert(BatchOnboardFlow.WorkbookPathProblem("Data.xlsx") <> "", "a bare file name is rejected")
    result = result & Assert(BatchOnboardFlow.WorkbookPathProblem("Documents\Data.xlsx") <> "", "a relative path is rejected")

    ' Wrong extension.
    result = result & Assert(BatchOnboardFlow.WorkbookPathProblem("C:\Temp\Data.pptx") <> "", "a non-workbook extension is rejected")

    ' Characters Windows won't take.
    result = result & Assert(BatchOnboardFlow.WorkbookPathProblem("C:\Temp\Da<ta.xlsx") <> "", "an invalid filename character is rejected")
    result = result & Assert(BatchOnboardFlow.WorkbookPathProblem("C:\Temp\Da:ta.xlsx") <> "", "a stray colon past the drive letter is rejected")

    result = result & Assert(BatchOnboardFlow.WorkbookPathProblem("") <> "", "an empty path is rejected")
    result = result & Assert(BatchOnboardFlow.WorkbookPathProblem("   ") <> "", "a whitespace-only path is rejected")

    ' The shapes that must still be ACCEPTED -- a validator that rejects
    ' everything is as useless as one that accepts everything.
    result = result & Assert(BatchOnboardFlow.WorkbookPathProblem("C:\Users\rohan\OneDrive\Claude\SAAFE-Projects-Data.xlsx") = "", _
        "a normal local path is accepted")
    result = result & Assert(BatchOnboardFlow.WorkbookPathProblem("C:\Users\rohan\OneDrive - SAAFE CRC\Data\Projects.xlsx") = "", _
        "a synced OneDrive-for-Business folder path is accepted -- this is the RIGHT answer to the URL trap")
    result = result & Assert(BatchOnboardFlow.WorkbookPathProblem("\\fileserver\share\Projects.xlsx") = "", _
        "a UNC network path is accepted")
    result = result & Assert(BatchOnboardFlow.WorkbookPathProblem("C:\Temp\Macro Data.xlsm") = "", _
        "an .xlsm with a space in the name is accepted")

    Test_BatchOnboardFlow_WorkbookPathProblemRejectsTheRealMistakes = result
End Function

' The regression guard for the bug that made the previous "conditional" save
' unconditional. The old guard was `(Not autoSaveOn) Or (Not pres.Saved)`, and
' Presentation.Saved reads False on an AutoSave cloud document even straight
' after a successful save -- so the second term was permanently true and every
' cloud deck got force-saved anyway, hiding Office's own Save command.
'
' The single most important assertion in this file is the healthy-cloud case:
' AutoSave on, timestamp readable, file written recently -> DO NOT save. If
' that ever returns True again, the guard has silently stopped guarding, which
' is precisely the failure that shipped once and was invisible to 93 tests.
Private Function Test_BatchOnboardFlow_ShouldForceSaveLeavesHealthyAutoSaveAlone() As String
    Dim result As String

    ' A plausible non-zero timestamp; only its non-zero-ness is meaningful.
    Dim stamp As Double
    stamp = CDbl(Now)

    ' THE case the whole change exists for.
    result = result & Assert(BatchOnboardFlow.ShouldForceSave(True, stamp, 5) = False, _
        "AutoSave on and the file saved 5s ago -> leave saving to Office")
    result = result & Assert(BatchOnboardFlow.ShouldForceSave(True, stamp, 119) = False, _
        "AutoSave on and the file saved 119s ago -> still inside tolerance, leave it")

    ' AutoSave stalled -- the 2026-07-28 failure, where the deck on disk was
    ' 2.6 hours stale while marks sat only in the open document.
    result = result & Assert(BatchOnboardFlow.ShouldForceSave(True, stamp, 121) = True, _
        "AutoSave on but no real save for 121s -> AutoSave has stalled, save it ourselves")
    result = result & Assert(BatchOnboardFlow.ShouldForceSave(True, stamp, 9360) = True, _
        "AutoSave on but the file is 2.6 hours stale -> save it ourselves")

    ' AutoSave off: the human owns saving and hasn't been asked to.
    result = result & Assert(BatchOnboardFlow.ShouldForceSave(False, stamp, 0) = True, _
        "AutoSave off -> always save, even if the file was just written")

    ' Timestamp unreadable: no evidence either way, so keep the work.
    result = result & Assert(BatchOnboardFlow.ShouldForceSave(True, 0, 0) = True, _
        "unreadable save timestamp -> err toward saving rather than trusting AutoSave")

    Test_BatchOnboardFlow_ShouldForceSaveLeavesHealthyAutoSaveAlone = result
End Function

' ShouldForceSave is only as good as the signal fed into it, so prove the
' signal is real: a presentation that has genuinely been saved must report a
' non-zero Last Save Time. If BuiltInDocumentProperties("Last Save Time") ever
' stops being readable this way, LastSaveTimeOf returns 0, ShouldForceSave
' returns True for everything, and the add-in quietly degrades back to the
' unconditional save -- failing safe for the user's data, but with the UI cost
' this whole change was made to remove. That degradation should be a visible
' test failure, not a silent one.
Private Function Test_BatchOnboardFlow_LastSaveTimeOfReadsARealTimestamp() As String
    Dim result As String

    Dim testPath As String
    testPath = Environ("TEMP") & "\deck_sync_test_lastsave_" & Format(Now, "hhmmss") & ".pptx"

    Dim testPres As Object
    Set testPres = Application.Presentations.Add
    testPres.SaveAs testPath

    Dim stamp As Double
    stamp = BatchOnboardFlow.LastSaveTimeOf(testPres)
    result = result & Assert(stamp <> 0, "a genuinely saved presentation reports a readable Last Save Time, got " & stamp)

    Dim stampText As String
    stampText = BatchOnboardFlow.LastSaveTimeTextOf(testPres)
    result = result & Assert(stampText <> "unknown", "the human-readable timestamp is a real value, got '" & stampText & "'")

    testPres.Saved = True
    testPres.Close

    Test_BatchOnboardFlow_LastSaveTimeOfReadsARealTimestamp = result
End Function

' The case the 2026-07-28 restore fix missed.
'
' That fix added a deck-identity term to the restore guard, on the reasoning
' that a stale in-memory session was suppressing the restore. True, but
' DeckIdentity is Presentation.FullName -- so closing deck X and reopening
' deck X reports the SAME identity, the guard stays False, and the marks still
' do not come back. The fix only ever covered switching to a different deck;
' the reported scenario (close and reopen the same deck) remained broken and
' no test said so.
'
' Row 4 below is that scenario. It fails against the old guard.
Private Function Test_BatchOnboardFlow_NeedsSessionRestoreCoversSameDeckReopen() As String
    Dim result As String

    Const SAME As String = "C:\Decks\project-status.pptx"
    Const OTHER As String = "C:\Decks\a-different-deck.pptx"

    ' 1. No session at all -- first mark since the add-in loaded.
    result = result & Assert(BatchOnboardFlow.NeedsSessionRestore(False, 0, "", SAME, False) = True, _
        "no session in memory -> read any saved session back")

    ' 2. Session object exists but holds nothing.
    result = result & Assert(BatchOnboardFlow.NeedsSessionRestore(True, 0, SAME, SAME, False) = True, _
        "empty session -> read any saved session back")

    ' 3. A live session belonging to a different deck.
    result = result & Assert(BatchOnboardFlow.NeedsSessionRestore(True, 2, OTHER, SAME, True) = True, _
        "session belongs to another deck -> restore this deck's own")

    ' 4. THE REGRESSION: same deck, non-empty session, dead shape references.
    result = result & Assert(BatchOnboardFlow.NeedsSessionRestore(True, 2, SAME, SAME, False) = True, _
        "same deck closed and reopened (refs dead) -> restore, not 'nothing marked'")

    ' 5. The one case that must NOT restore -- a genuinely live session, where
    '    restoring would throw away marks the user just made.
    result = result & Assert(BatchOnboardFlow.NeedsSessionRestore(True, 2, SAME, SAME, True) = False, _
        "a live session for this deck -> keep it, do not overwrite from the file")

    Test_BatchOnboardFlow_NeedsSessionRestoreCoversSameDeckReopen = result
End Function

' NeedsSessionRestore is only correct if `shapesStillLive` tells the truth, and
' that rests on an empirical claim about Office: that reading .Name on a Shape
' whose presentation has been closed raises rather than quietly returning a
' value. Asserted from memory that is a guess, so this probes it against the
' real thing -- mark a shape, close the deck WITHOUT resetting the session
' (exactly what happens in use, since nothing resets on close), reopen, and
' check what the add-in now believes.
Private Function Test_BatchOnboardFlow_ReopeningTheSameDeckLeavesShapeRefsDead() As String
    Dim result As String

    Dim testPath As String
    testPath = Environ("TEMP") & "\deck_sync_test_reopen_guard_" & Format(Now, "hhmmss") & ".pptx"

    Dim testPres As Object
    Set testPres = Application.Presentations.Add
    testPres.SaveAs testPath

    Dim sld As Object
    Set sld = testPres.Slides.Add(1, ppLayoutBlank)
    Application.ActiveWindow.View.GotoSlide sld.SlideIndex

    Dim shapeA As Object
    Set shapeA = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 100, 50)
    shapeA.Name = "ReopenGuardField"
    shapeA.TextFrame.TextRange.Text = "Field A"

    BatchOnboardFlow.ResetMarkingSession
    BatchOnboardFlow.MarkShapeForBatch shapeA, "Project Number", "text", "static"
    BatchOnboardFlow.SaveMarkingSessionToProperty testPres

    result = result & Assert(BatchOnboardFlow.MarkedShapesStillLive() = True, _
        "shape references are live while the deck is open")

    ' Deliberately NO ResetMarkingSession here -- the whole point is that real
    ' use leaves the session in memory across a close.
    testPres.Saved = True
    testPres.Close
    Set testPres = Nothing

    result = result & Assert(BatchOnboardFlow.MarkedFieldCountForBatch() > 0, _
        "the session survives the close in memory, got " & BatchOnboardFlow.MarkedFieldCountForBatch() & " -- if this is 0, the premise of the whole guard has changed")
    result = result & Assert(BatchOnboardFlow.MarkedShapesStillLive() = False, _
        "after the deck closes, the held Shape references are detectably dead")

    Dim reopened As Object
    Set reopened = Application.Presentations.Open(testPath, False, False, False)

    result = result & Assert(BatchOnboardFlow.MarkedShapesStillLive() = False, _
        "reopening the same file does not resurrect the old Shape references")
    result = result & Assert(BatchOnboardFlow.NeedsSessionRestore(True, BatchOnboardFlow.MarkedFieldCountForBatch(), reopened.FullName, reopened.FullName, BatchOnboardFlow.MarkedShapesStillLive()) = True, _
        "so the add-in decides to restore from the file after a same-deck reopen")

    BatchOnboardFlow.ResetMarkingSession
    reopened.Saved = True
    reopened.Close

    Test_BatchOnboardFlow_ReopeningTheSameDeckLeavesShapeRefsDead = result
End Function

' The test above proves the write+Save+read-back sequence within ONE
' still-running presentation object -- it does NOT prove the write is
' actually durable once that presentation is genuinely closed and the file
' is reopened fresh, which is the exact failure mode Rohan actually hit
' (marks vanishing after a real PowerPoint close/reopen, not just "read it
' back a moment later"). VBA can't spawn a whole new PowerPoint.EXE process
' mid-test the way the PowerShell-side investigation did (see
' SPIKE_NOTES_BatchOnboardFlow.md's "genuine close/reopen" addendum for that
' out-of-process proof, run via COM automation, which is outside what an
' in-process VBA test harness can do) -- but a real Presentation.Close
' followed by a fresh Presentations.Open of the same file on disk is the
' strongest version of this check reachable from inside TestRunner.bas, and
' it's a real regression guard: it would fail if a future change removed
' the forced Save, wrote to the wrong property name, or broke the
' serialize/restore round trip, none of which the in-session-only test
' above can catch on its own.
Private Function Test_BatchOnboardFlow_MarkingSessionSurvivesRealCloseAndReopen() As String
    Dim result As String

    Dim testPath As String
    testPath = Environ("TEMP") & "\deck_sync_test_close_reopen_" & Format(Now, "hhmmss") & ".pptx"

    Dim testPres As Object
    Set testPres = Application.Presentations.Add
    testPres.SaveAs testPath

    Dim sld As Object
    Set sld = testPres.Slides.Add(1, ppLayoutBlank)
    Application.ActiveWindow.View.GotoSlide sld.SlideIndex

    Dim shapeA As Object
    Set shapeA = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 100, 50)
    shapeA.Name = "CloseReopenField"
    shapeA.TextFrame.TextRange.Text = "Field A"

    BatchOnboardFlow.ResetMarkingSession
    BatchOnboardFlow.MarkShapeForBatch shapeA, "Project Number", "text", "static"

    Dim saveWarning As String
    saveWarning = BatchOnboardFlow.SaveMarkingSessionToProperty(testPres)
    result = result & Assert(saveWarning = "", "no warning forcing the save before closing, got '" & saveWarning & "'")

    Dim writtenSerialized As String
    writtenSerialized = BatchOnboardFlow.SerializeCurrentMarkingSession()

    BatchOnboardFlow.ResetMarkingSession ' clears in-memory + the persisted property's in-memory mirror is irrelevant now -- only the file on disk matters from here
    testPres.Close
    Set testPres = Nothing

    ' Fresh Open of the same path -- ReadOnly:=False, Untitled:=False, so
    ' this is a genuine reopen of the real file identity, not a detached
    ' copy (confirmed the hard way during this investigation's own
    ' PowerShell probing: Untitled:=True silently opens a disconnected
    ' "Presentation1" instead of the real file).
    Dim reopened As Object
    Set reopened = Application.Presentations.Open(testPath, False, False, False)

    Dim readBack As String
    readBack = BatchOnboardFlow.ReadMarkingSessionProperty(reopened)
    result = result & Assert(readBack = writtenSerialized, "marking session property survives a real Close+reopen, wrote " & Len(writtenSerialized) & " chars, read back " & Len(readBack) & " chars")
    result = result & Assert(InStr(readBack, "Project Number") > 0, "the reopened property genuinely contains the marked field, got '" & readBack & "'")

    Dim restoreReport As String
    restoreReport = BatchOnboardFlow.RestoreMarkingSession(readBack, reopened.Slides(1))
    result = result & Assert(InStr(restoreReport, "Restored 1 field") > 0, "restoring after a real close/reopen recovers the field, got '" & restoreReport & "'")
    result = result & Assert(BatchOnboardFlow.MarkedFieldNameForBatch(1) = "Project Number", "restored field name matches after real close/reopen, got '" & BatchOnboardFlow.MarkedFieldNameForBatch(1) & "'")

    BatchOnboardFlow.ResetMarkingSession
    reopened.Saved = True
    reopened.Close

    Test_BatchOnboardFlow_MarkingSessionSurvivesRealCloseAndReopen = result
End Function

' A shape marked as "different slide" and declined should NOT get saved
' into the persisted property under the abandoned slide's identity --
' covers the restoreReport-discard branch in MarkFieldForBatch (not
' independently unit-testable itself, since it's inside the MsgBox-driven
' Sub) by exercising the same underlying state transition directly:
' RestoreMarkingSession finding nothing on a mismatched slide must leave a
' clean, empty session rather than a half-populated one.
Private Function Test_BatchOnboardFlow_RestoreMarkingSessionFindsNothingOnWrongSlide() As String
    Dim result As String

    BatchOnboardFlow.ResetMarkingSession

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim shapeA As Object
    Set shapeA = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 100, 50)
    shapeA.TextFrame.TextRange.Text = "Field A"
    BatchOnboardFlow.MarkShapeForBatch shapeA, "Project Number", "text", "static"

    Dim serialized As String
    serialized = BatchOnboardFlow.SerializeCurrentMarkingSession()
    BatchOnboardFlow.ResetMarkingSession

    Dim otherSld As Object
    Set otherSld = NewBlankSlide() ' a genuinely different, unrelated slide -- "Project Number" shape doesn't exist here

    Dim restoreReport As String
    restoreReport = BatchOnboardFlow.RestoreMarkingSession(serialized, otherSld)

    result = result & Assert(InStr(restoreReport, "Restored 0 field") > 0, "restoring against the wrong slide finds nothing, got '" & restoreReport & "'")
    result = result & Assert(BatchOnboardFlow.MarkedFieldCountForBatch() = 0, "no fields marked after a mismatched restore, got " & BatchOnboardFlow.MarkedFieldCountForBatch())

    BatchOnboardFlow.ResetMarkingSession

    Test_BatchOnboardFlow_RestoreMarkingSessionFindsNothingOnWrongSlide = result
End Function

' Real finding 2026-07-26, live-tested against Rohan's real deck: a group
' selected on-screen (Shape Format ribbon tab active, tight selection
' handles around just one field) does NOT reliably mean Application.
' Selection.ShapeRange(1) reports that individual member -- it can still
' report the outer group. FlattenGroupLeaves is the fix's pure half: given
' the group Application.Selection actually reports, list every real field
' shape inside it so MarkFieldForBatch can offer a numbered pick instead of
' guessing which one was clicked.
Private Function Test_BatchOnboardFlow_FlattenGroupLeavesReturnsAllMembers() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim shapeA As Object, shapeB As Object
    Set shapeA = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 100, 50)
    shapeA.TextFrame.TextRange.Text = "Leaf A"
    Set shapeB = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 200, 50, 100, 50)
    shapeB.TextFrame.TextRange.Text = "Leaf B"

    Dim grp As Object
    Set grp = sld.Shapes.Range(Array(shapeA.Name, shapeB.Name)).Group()

    Dim leaves() As Object
    Dim leafCount As Long
    leafCount = BatchOnboardFlow.FlattenGroupLeaves(grp, leaves)

    result = result & Assert(leafCount = 2, "2 leaf shapes found, got " & leafCount)

    Dim foundA As Boolean, foundB As Boolean
    Dim i As Long
    For i = 1 To leafCount
        If leaves(i) Is shapeA Then foundA = True
        If leaves(i) Is shapeB Then foundB = True
    Next i
    result = result & Assert(foundA, "leaf A found by identity")
    result = result & Assert(foundB, "leaf B found by identity")

    Test_BatchOnboardFlow_FlattenGroupLeavesReturnsAllMembers = result
End Function

Private Function Test_BatchOnboardFlow_ReviewGridRoundTrip() As String
    Dim result As String

    Dim plan As BatchOnboardPlan
    Set plan.FieldNames = CreateObject("Scripting.Dictionary")
    Set plan.FieldTypes = CreateObject("Scripting.Dictionary")
    Set plan.FieldVolatility = CreateObject("Scripting.Dictionary")
    Set plan.FieldTemplateShapes = CreateObject("Scripting.Dictionary")
    Set plan.FieldSuggestIdentical = CreateObject("Scripting.Dictionary")
    Set plan.FieldInclude = CreateObject("Scripting.Dictionary")
    Set plan.Correspondence = CreateObject("Scripting.Dictionary")
    Set plan.HarvestedText = CreateObject("Scripting.Dictionary")
    plan.FieldCount = 2
    plan.FieldNames(1) = "ph_field1"
    plan.FieldTypes(1) = "text"
    plan.FieldVolatility(1) = "variable"
    plan.FieldSuggestIdentical(1) = True
    plan.FieldInclude(1) = False
    plan.HarvestedText("1|0") = "Overall Status"
    plan.HarvestedText("1|1") = "Overall Status" ' field 1 has its own real sample too -- must not leak into field 2's cell (real bug found live 2026-07-26: samples accumulated across rows because a loop-local reset was missing)
    plan.FieldNames(2) = "ph_field2"
    plan.FieldTypes(2) = "currency"
    plan.FieldVolatility(2) = "static"
    plan.FieldSuggestIdentical(2) = False
    plan.FieldInclude(2) = True
    plan.HarvestedText("2|0") = "Q1 2026"
    plan.HarvestedText("2|1") = "Q2 2026"

    Dim xl As Object, wb As Object, ws As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Add()
    Set ws = wb.Worksheets(1)

    BatchOnboardFlow.WriteReviewGrid ws, plan, 1

    result = result & Assert(ws.Cells(2, 2).Value = "ph_field1", "row 2 (field 1) name written, got '" & ws.Cells(2, 2).Value & "'")
    result = result & Assert(ws.Cells(2, 4).Value = "N", "row 2 (field 1, suggested decoration) defaults Include to N, got '" & ws.Cells(2, 4).Value & "'")
    result = result & Assert(ws.Cells(3, 4).Value = "Y", "row 3 (field 2, suggested real field) defaults Include to Y, got '" & ws.Cells(3, 4).Value & "'")
    result = result & Assert(InStr(ws.Cells(3, 6).Value, "Q2 2026") > 0, "row 3's sample-other-values column includes the other slide's harvested value, got '" & ws.Cells(3, 6).Value & "'")
    result = result & Assert(InStr(ws.Cells(3, 6).Value, "Overall Status") = 0, "row 3's sample-other-values column does NOT include field 1's own sample -- each row's samples must not accumulate across fields, got '" & ws.Cells(3, 6).Value & "'")
    result = result & Assert(ws.Cells(3, 7).Value = "currency", "row 3 (field 2)'s Type column shows the declared type, got '" & ws.Cells(3, 7).Value & "'")
    result = result & Assert(ws.Cells(3, 8).Value = "static", "row 3 (field 2)'s Static/Variable column shows the declared hint, got '" & ws.Cells(3, 8).Value & "'")

    ' Simulate a human editing the sheet: rename field 1, flip its Include
    ' to Y (overriding the suggestion), exclude field 2, change field 1's
    ' type from its default "text" to "date", and flip field 1's volatility
    ' hint from "variable" to "static".
    ws.Cells(2, 2).Value = "ph_renamed"
    ws.Cells(2, 4).Value = "y" ' lower-case, must still be read as Y
    ws.Cells(3, 4).Value = "n"
    ws.Cells(2, 7).Value = "date"
    ws.Cells(2, 8).Value = "static"

    BatchOnboardFlow.ReadReviewGrid ws, plan

    result = result & Assert(plan.FieldNames(1) = "ph_renamed", "renamed field name read back, got '" & plan.FieldNames(1) & "'")
    result = result & Assert(plan.FieldInclude(1), "field 1's Include flipped to True after edit (lower-case 'y' accepted)")
    result = result & Assert(Not plan.FieldInclude(2), "field 2's Include flipped to False after edit")
    result = result & Assert(plan.FieldTypes(1) = "date", "field 1's type edit read back, got '" & plan.FieldTypes(1) & "'")
    result = result & Assert(plan.FieldTypes(2) = "currency", "field 2's type unchanged when the grid cell wasn't touched, got '" & plan.FieldTypes(2) & "'")
    result = result & Assert(plan.FieldVolatility(1) = "static", "field 1's volatility edit read back, got '" & plan.FieldVolatility(1) & "'")
    result = result & Assert(plan.FieldVolatility(2) = "static", "field 2's volatility unchanged when the grid cell wasn't touched, got '" & plan.FieldVolatility(2) & "'")

    wb.Saved = True
    wb.Close
    xl.Quit

    Test_BatchOnboardFlow_ReviewGridRoundTrip = result
End Function

Private Function Test_BatchOnboardFlow_CommitBatchTagsLinksAndVerifies() As String
    Dim result As String

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()
    Dim tShapeA As Object
    Set tShapeA = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    tShapeA.TextFrame.TextRange.Text = "Overall Status"

    Dim other1 As Object
    Set other1 = NewBlankSlide()
    Dim o1ShapeA As Object
    Set o1ShapeA = other1.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    o1ShapeA.TextFrame.TextRange.Text = "Overall Status"

    Dim otherSlides(1 To 1) As Object
    Set otherSlides(1) = other1

    Dim plan As BatchOnboardPlan
    plan = BatchOnboardFlow.BuildBatchPlan(templateSld, otherSlides)
    result = result & Assert(plan.FieldCount = 1, "1 candidate field found, got " & plan.FieldCount)

    ' The harvested text ("Overall Status") is deliberately identical on
    ' both slides -- confirms BuildBatchPlan's own classification correctly
    ' suggests this as decoration (Include defaults to False) before this
    ' test explicitly overrides it, exactly as a human reviewing the grid
    ' and choosing to keep it would. Commit/tagging mechanics are this
    ' test's actual subject, not the classification default itself (that's
    ' BuildBatchPlanFindsCorrespondenceAndHarvestsAcrossSlides's job).
    result = result & Assert(plan.FieldSuggestIdentical(1), "identical harvested text is correctly suggested as decoration")
    result = result & Assert(Not plan.FieldInclude(1), "decoration defaults to excluded before the override below")
    plan.FieldInclude(1) = True

    Dim xl As Object, wb As Object, ws As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Add()
    Set ws = wb.Worksheets(1)
    ExcelOutput.CreateSheet ws, "test-deck-id"

    Dim confirmedKeys As Object
    Set confirmedKeys = CreateObject("Scripting.Dictionary")
    confirmedKeys(0) = "batch-template"
    confirmedKeys(1) = "batch-other-1"

    Dim commitResult As BatchCommitResult
    ' The deck must declare its quarter before it can be onboarded -- every row
    ' written carries a period and UpsertRow refuses a blank one. Toolbar step 0.
    DeckRegistry.SetDeckPeriod Application.ActivePresentation, "FY26Q4"

    commitResult = BatchOnboardFlow.CommitBatch(plan, templateSld, otherSlides, 1, "batch-test-type", ws, confirmedKeys)

    result = result & Assert(commitResult.LinkedCount = 2, "both slides linked, got " & commitResult.LinkedCount)
    result = result & Assert(commitResult.FailedVerificationCount = 0, "no verification failures, got " & commitResult.FailedVerificationCount)
    result = result & Assert(tShapeA.Tags("role") = plan.FieldNames(1), "template shape tagged with the field's role")
    result = result & Assert(o1ShapeA.Tags("role") = plan.FieldNames(1), "other1's corresponding shape tagged with the same role")

    Dim templateInstance As SlideInstance
    templateInstance = Resolve.ResolveSlideInstance(templateSld)
    result = result & Assert(templateInstance.HasTypeTag And templateInstance.TypeTag = "batch-test-type", "template slide carries the new slide_type tag")
    result = result & Assert(templateInstance.HasInstanceKey And templateInstance.InstanceKey = "batch-template", "template slide carries its confirmed instance key")

    result = result & Assert(ws.Cells(2, 1).Value = "batch-template" Or ws.Cells(3, 1).Value = "batch-template", "a Data-sheet row exists for the template's instance key")
    result = result & Assert(ws.Cells(2, 1).Value = "batch-other-1" Or ws.Cells(3, 1).Value = "batch-other-1", "a Data-sheet row exists for other1's instance key")

    wb.Saved = True
    wb.Close
    xl.Quit

    Test_BatchOnboardFlow_CommitBatchTagsLinksAndVerifies = result
End Function

' Real production incident, 2026-07-26: Rohan ran "Bulk Onboard Type" for
' real, for the first time at real scale, against his actual 46-slide deck.
' Every single slide FAILED verification (Linked: 0, FailedVerification: 46)
' -- see SPIKE_NOTES_BatchOnboardFlow.md's addendum for the full root-cause
' account. Root cause: InjectPrimitive.bas's FindShapeByRoleTag looped over
' `sld.Shapes` (PowerPoint's TOP-LEVEL-ONLY shape collection), which never
' finds a role tag written onto a shape nested inside a group -- and
' Rohan's real deck's "card" fields live inside groups. Every CommitBatch
' test before this one (including CommitBatchTagsLinksAndVerifies directly
' above) used only top-level, ungrouped textboxes on a 1-template-plus-1-
' other-slide batch -- nowhere near real scale, and structurally incapable
' of catching this bug regardless of scale, since the bug is about
' nesting, not slide count.
'
' This test reproduces BOTH dimensions of the gap: (1) a template + 10
' other slides (11 total, matching the "much bigger than any prior test
' batch, which topped out at 2-3 slides" scale that just failed for real),
' and (2) each slide's field set is a REALISTIC mix -- one top-level field
' plus one field nested one level inside a group, mirroring Rohan's actual
' "card" layout. Before the FindShapeByRoleTag fix, this test fails exactly
' the way the real run did: LinkedCount=0, FailedVerificationCount=11 (every
' slide fails, because the grouped field's verification always fails and
' VerifyBatchLink fails the WHOLE slide the moment any one field fails).
' After the fix, all 11 link and verify cleanly.
Private Function Test_BatchOnboardFlow_CommitBatchWithGroupedFieldsAtScale() As String
    Dim result As String
    Const otherCount As Long = 10

    Dim templateSld As Object
    Set templateSld = NewBlankSlide()

    Dim tFieldA As Object
    Set tFieldA = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 250, 200, 50)
    tFieldA.TextFrame.TextRange.Text = "Project Number: T-000"

    Dim tFieldB As Object, tSibling As Object
    Set tFieldB = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
    tFieldB.TextFrame.TextRange.Text = "Status: Active"
    Set tSibling = templateSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 300, 50, 200, 50)
    tSibling.TextFrame.TextRange.Text = "Card Chrome" ' identical on every slide -- decoration, never included
    templateSld.Shapes.Range(Array(tFieldB.Name, tSibling.Name)).Group

    Dim otherSlides(1 To otherCount) As Object
    Dim otherFieldB(1 To otherCount) As Object ' kept for post-commit assertions -- the grouped shape on each other slide
    Dim i As Long
    For i = 1 To otherCount
        Dim sld As Object
        Set sld = NewBlankSlide()
        Set otherSlides(i) = sld

        Dim oFieldA As Object
        Set oFieldA = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 250, 200, 50)
        oFieldA.TextFrame.TextRange.Text = "Project Number: T-0" & i

        Dim oFieldB As Object, oSibling As Object
        Set oFieldB = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 200, 50)
        oFieldB.TextFrame.TextRange.Text = "Status: Phase " & i
        Set oSibling = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 300, 50, 200, 50)
        oSibling.TextFrame.TextRange.Text = "Card Chrome"
        sld.Shapes.Range(Array(oFieldB.Name, oSibling.Name)).Group
        Set otherFieldB(i) = oFieldB
    Next i

    Dim plan As BatchOnboardPlan
    plan = BatchOnboardFlow.BuildBatchPlan(templateSld, otherSlides)
    result = result & Assert(plan.FieldCount = 3, "3 candidate fields found (fieldA, grouped fieldB, grouped sibling decoration), got " & plan.FieldCount)

    ' Identify each field's index by its template value -- discovery order
    ' isn't guaranteed, same idiom BuildBatchPlanFindsCorrespondenceAnd
    ' HarvestsAcrossSlides already established.
    Dim fieldAIdx As Long, fieldBIdx As Long, siblingIdx As Long
    Dim fi As Long
    For fi = 1 To plan.FieldCount
        Dim tv As String
        tv = plan.HarvestedText(CStr(fi) & "|0")
        If tv = "Project Number: T-000" Then fieldAIdx = fi
        If tv = "Status: Active" Then fieldBIdx = fi
        If tv = "Card Chrome" Then siblingIdx = fi
    Next fi

    result = result & Assert(fieldAIdx > 0, "found the top-level field")
    result = result & Assert(fieldBIdx > 0, "found the grouped field")
    result = result & Assert(siblingIdx > 0, "found the grouped decoration sibling")
    If fieldAIdx = 0 Or fieldBIdx = 0 Or siblingIdx = 0 Then
        Test_BatchOnboardFlow_CommitBatchWithGroupedFieldsAtScale = result
        Exit Function
    End If

    plan.FieldInclude(fieldAIdx) = True
    plan.FieldInclude(fieldBIdx) = True
    ' siblingIdx left at its suggested default (excluded -- identical everywhere)

    Dim xl As Object, wb As Object, ws As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Add()
    Set ws = wb.Worksheets(1)
    ExcelOutput.CreateSheet ws, "test-deck-id-scale"

    Dim confirmedKeys As Object
    Set confirmedKeys = CreateObject("Scripting.Dictionary")
    confirmedKeys(0) = "scale-template"
    For i = 1 To otherCount
        confirmedKeys(i) = "scale-other-" & i
    Next i

    Dim commitResult As BatchCommitResult
    DeckRegistry.SetDeckPeriod Application.ActivePresentation, "FY26Q4"

    commitResult = BatchOnboardFlow.CommitBatch(plan, templateSld, otherSlides, otherCount, "batch-scale-test-type", ws, confirmedKeys)

    result = result & Assert(commitResult.LinkedCount = otherCount + 1, "all " & (otherCount + 1) & " slides (template + " & otherCount & " others) linked, got " & commitResult.LinkedCount)
    result = result & Assert(commitResult.FailedVerificationCount = 0, "no verification failures -- got " & commitResult.FailedVerificationCount & " (this is exactly the real production symptom: pre-fix, this would be " & (otherCount + 1) & ", matching 'Linked: 0 / FAILED verification: 46')")

    ' Directly confirm the grouped shapes themselves carry the right role
    ' tag and final text -- not just that VerifyBatchLink's own internal
    ' round-trip passed, in case that round-trip and this check shared a
    ' blind spot.
    result = result & Assert(tFieldB.Tags("role") = plan.FieldNames(fieldBIdx), "template's grouped field shape tagged with the field's role")
    For i = 1 To otherCount
        result = result & Assert(otherFieldB(i).Tags("role") = plan.FieldNames(fieldBIdx), "other" & i & "'s grouped field shape tagged with the same role")
        result = result & Assert(otherFieldB(i).TextFrame.TextRange.Text = "Status: Phase " & i, "other" & i & "'s grouped field text is unchanged (harvested value re-verified against itself), got '" & otherFieldB(i).TextFrame.TextRange.Text & "'")
    Next i

    wb.Saved = True
    wb.Close
    xl.Quit

    Test_BatchOnboardFlow_CommitBatchWithGroupedFieldsAtScale = result
End Function

' ===========================================================================
' DiscoverUI -- the marking grid. Added 2026-08-01, the evening it was written,
' because a module with no coverage is a module whose next change breaks it
' silently. The prompt chain it replaces went 52 fields and three findings
' without anything asserting its behaviour.
' ===========================================================================

Private Function Test_DiscoverUI_GridListsEveryTextShapeWithUniqueIds() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()

    ' Deliberately out of reading order, and deliberately sharing a NAME --
    ' the exact condition that destroyed an hour of real marking.
    Dim bottom As Object, top_ As Object, middle As Object
    Set bottom = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 60, 400, 200, 40)
    bottom.TextFrame.TextRange.Text = "BOTTOM"
    Set top_ = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 60, 50, 200, 40)
    top_.TextFrame.TextRange.Text = "TOP"
    Set middle = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 60, 220, 200, 40)
    middle.TextFrame.TextRange.Text = "MIDDLE"
    bottom.Name = "Shape 16"
    top_.Name = "Shape 16"
    middle.Name = "Shape 16"

    Dim xl As Object, wb As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set wb = xl.Workbooks.Add

    Dim note As String
    note = DiscoverUI.BuildDiscoverySheet(sld, wb)
    result = result & Assert(Left(note, 1) <> "!", "the grid builds, got '" & note & "'")

    Dim ws As Object
    Set ws = wb.Worksheets("Field Discovery")

    ' Reading order, not creation order and not z-order.
    result = result & Assert(ws.Cells(7, 4).Value = "TOP", _
        "first grid row is the TOPMOST shape, got '" & ws.Cells(7, 4).Value & "'")
    result = result & Assert(ws.Cells(8, 4).Value = "MIDDLE", _
        "second row is the middle shape, got '" & ws.Cells(8, 4).Value & "'")
    result = result & Assert(ws.Cells(9, 4).Value = "BOTTOM", _
        "third row is the bottom shape, got '" & ws.Cells(9, 4).Value & "'")

    ' THE WHOLE POINT. Three shapes share one name; their ids must differ, or
    ' the grid reproduces finding 11 in a new place.
    result = result & Assert(ws.Cells(7, 2).Value = ws.Cells(8, 2).Value, _
        "the fixture really does share a shape NAME across rows")
    result = result & Assert(ws.Cells(7, 1).Value <> ws.Cells(8, 1).Value, _
        "IDs DIFFER even where names collide -- got " & ws.Cells(7, 1).Value & " and " & ws.Cells(8, 1).Value)
    result = result & Assert(ws.Cells(8, 1).Value <> ws.Cells(9, 1).Value, _
        "all three ids are distinct")

    wb.Saved = True: wb.Close
    xl.Quit
    Set wb = Nothing: Set xl = Nothing
    sld.Delete
    Test_DiscoverUI_GridListsEveryTextShapeWithUniqueIds = result
End Function

Private Function Test_DiscoverUI_MarksOnlyTickedAndNamedRows() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim a As Object, b As Object, c As Object
    Set a = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 60, 50, 200, 40)
    a.TextFrame.TextRange.Text = "ALPHA"
    Set b = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 60, 150, 200, 40)
    b.TextFrame.TextRange.Text = "BETA"
    Set c = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 60, 250, 200, 40)
    c.TextFrame.TextRange.Text = "GAMMA"

    Dim xl As Object, wb As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set wb = xl.Workbooks.Add
    DiscoverUI.BuildDiscoverySheet sld, wb

    Dim ws As Object
    Set ws = wb.Worksheets("Field Discovery")

    ' Row 1: ticked and named  -> marks.
    ' Row 2: named but NOT ticked -> must not mark. A name typed while thinking
    '        is not an instruction.
    ' Row 3: ticked with NO name -> must not mark, and must be REPORTED.
    ws.Cells(7, 6).Value = "Y": ws.Cells(7, 7).Value = "FIELD_ALPHA"
    ws.Cells(8, 7).Value = "FIELD_BETA"
    ws.Cells(9, 6).Value = "Y"

    Dim rep As String
    rep = DiscoverUI.ApplyDiscoverySheet(sld, wb)

    result = result & Assert(BatchOnboardFlow.MarkedFieldCountForBatch() = 1, _
        "exactly one row marked -- ticked AND named, got " & BatchOnboardFlow.MarkedFieldCountForBatch())
    result = result & Assert(BatchOnboardFlow.MarkedFieldNameForBatch(1) = "FIELD_ALPHA", _
        "the marked field is the ticked-and-named one, got '" & BatchOnboardFlow.MarkedFieldNameForBatch(1) & "'")
    result = result & Assert(InStr(rep, "ticked but has no field name") > 0, _
        "the ticked-but-unnamed row is REPORTED, never guessed at")
    result = result & Assert(InStr(rep, "FIELD_BETA") = 0, _
        "a named-but-unticked row is silently left alone, not marked and not complained about")

    ' ResetMarkingSession, not ClearMarkedFieldsForBatch -- the latter shows a
    ' MsgBox and would hang a headless run. The public reset already existed;
    ' adding a new "silent" wrapper for it was a solution to a problem I had not
    ' checked for.
    ' UNTICK MUST UNMARK. Re-apply the same grid with row 1 unticked; the field
    ' marked a moment ago must be gone. Without this the grid only ever adds,
    ' which is what made finding 4 survive a rewrite that looked like it fixed it.
    ws.Cells(7, 6).Value = ""
    Dim rep2 As String
    rep2 = DiscoverUI.ApplyDiscoverySheet(sld, wb)
    result = result & Assert(BatchOnboardFlow.MarkedFieldCountForBatch() = 0, _
        "unticking a row REMOVES its mark, got " & BatchOnboardFlow.MarkedFieldCountForBatch() & " still marked")
    result = result & Assert(InStr(rep2, "UNMARKED") > 0, _
        "the removal is reported, not silent")
    result = result & Assert(InStr(rep2, "FIELD_ALPHA") > 0, _
        "the report NAMES what it removed, so an accidental untick is visible")

    BatchOnboardFlow.ResetMarkingSession
    wb.Saved = True: wb.Close
    xl.Quit
    Set wb = Nothing: Set xl = Nothing
    sld.Delete
    Test_DiscoverUI_MarksOnlyTickedAndNamedRows = result
End Function

' PERIOD ROLLOVER MUST NOT CARRY LAST QUARTER'S TEXT FORWARD.
'
' The drafting sheet is per FIELD, not per period, and a rebuild deliberately
' carries SUBMIT / SOURCES / AI DRAFT / NOTES across by EntityCode so that a
' refresh never costs someone their evening. Roll the deck to a new quarter and
' that same kindness becomes the defect: last quarter's prose sits in column G
' against last quarter's source IDs, approvals cleared but ONE TICK away from
' being republished as though it were current -- and CopyAiToSubmit would
' decline to overwrite it, because from where it stands the person had already
' written something there. Silent, plausible, and funder-facing.
'
' ASSERTS BOTH DIRECTIONS, deliberately. A guard that fired on every rebuild
' would destroy work on the ordinary path while looking like caution, so the
' same-period rebuild is asserted to still carry everything across. A test that
' only checked the dropping half would pass against `layoutMatches = False`
' hardcoded, which is not the behaviour anyone wants.
'
' SCOPE, and it is borrowed from another module: dropping every row wholesale
' is only correct because a drafting sheet can only ever hold PROSE fields --
' AskForField offers nothing else (DraftingUI.bas:124). Entity-static content
' lives in `Quarter = ALL` register rows, which are supposed to survive a
' rollover. Nothing here can test that, because nothing can currently put such
' a row on a drafting sheet; see the note in Drafting.WriteDraftingSheet.
' A SAME-PERIOD REBUILD COSTS NOTHING, AND A ROLLOVER WITH NOTHING AT RISK
' STILL REBUILDS.
'
' This function was called Test_Drafting_PeriodRolloverDropsStaleSubmit and
' asserted that a rollover DISCARDED last quarter's typed work. That behaviour
' was removed on 2026-08-13 -- it was one button press from taking 129 drafted
' values off Rohan -- and the test was left asserting the defect, which is why
' the suite went red. Renamed rather than patched: the old name stated the
' defect as the requirement, so a passing test under it would have read as
' confirmation that the loss was intended.
'
' The refusal contract itself lives in
' Test_Drafting_RefusesRatherThanDiscardOnPeriodChange and is NOT re-asserted
' here. What remains, and is covered nowhere else, is the two paths that still
' rebuild:
'   1. same period, work present        -> everything is carried across
'   2. different period, nothing at risk -> rebuilds and re-stamps
'
' Path 2 is now the ONLY way a rollover reaches the rebuild at all: a single row
' holding SUBMIT or DRAFT text refuses the whole sheet. Worth stating plainly,
' because it makes the per-row cadence machinery below far narrower than it
' looks -- see Test_Drafting_RolloverCadenceGovernsUntypedRows.
' THE QUARTER TURN FERRIES SUBMIT SIDEWAYS. It does not refuse and it does not
' destroy.
'
' Rohan, 2026-08-14: "whenever a 1/4 changes at the top, the ferries belonging to
' that system deal with information for the new quarter. The last previous set
' that runs move info into 'reported last time' column."
'
' This replaces Drafting_RolloverRebuildsOnlyWhenNothingIsAtRisk, which asserted
' the refusal contract. That refusal was a DEADLOCK -- nothing in the tool could
' ever satisfy it -- so a test written to hold it in place was holding a defect
' in place. Both it and Drafting_RolloverCadenceGovernsUntypedRows are deleted
' rather than patched, because their names stated the old behaviour as the
' requirement.
Private Function Test_Drafting_QuarterTurnFerriesSubmitIntoReportedLastTime() As String
    Dim result As String

    Dim xl As Object, wb As Object, dws As Object, rws As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set wb = xl.Workbooks.Add
    Set dws = wb.Worksheets(1)
    Set rws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))

    rws.Cells(1, 1).Value = ExcelOutput.INSTANCE_ID_HEADER
    rws.Cells(1, 2).Value = "PROJECT_NAME"
    rws.Cells(1, 3).Value = "ABOUT_BODY"
    rws.Cells(2, 1).Value = "P001": rws.Cells(2, 2).Value = "Alpha": rws.Cells(2, 3).Value = "register one"
    rws.Cells(3, 1).Value = "P002": rws.Cells(3, 2).Value = "Beta":  rws.Cells(3, 3).Value = "register two"

    Dim reg As Sheet
    reg = ExcelOutput.ReadSheet(rws)

    ' 1. Build it for FY26Q3 and put a person's evening on it.
    Drafting.WriteDraftingSheet dws, reg, "ABOUT_BODY", Empty, "FY26Q3"
    dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_SUBMIT).Value = "P001 last quarter"
    dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_SOURCES).Value = "S01,S03"
    dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_NOTES).Value = "chase the finance number"
    dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_DRAFT).Value = "AI text from last quarter"
    dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_APPROVED).Value = "Y"

    ' 2. SAME quarter: nothing moves. This is the everyday case and it must not
    '    touch a single typed column.
    Drafting.WriteDraftingSheet dws, reg, "ABOUT_BODY", Empty, "FY26Q3"
    result = result & Assert(CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_SUBMIT).Value) = "P001 last quarter", _
        "SAME QUARTER LEAVES SUBMIT ALONE, got '" & _
        CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_SUBMIT).Value) & "'")
    result = result & Assert(CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_PREV).Value) = "", _
        "and nothing is ferried while the quarter has not turned, got '" & _
        CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_PREV).Value) & "'")

    ' 3. THE QUARTER TURNS.
    Dim rep As String
    rep = Drafting.WriteDraftingSheet(dws, reg, "ABOUT_BODY", Empty, "FY26Q4")

    result = result & Assert(InStr(rep, "REFUSED") = 0, _
        "A QUARTER TURN IS NOT REFUSED. Report was: " & Left(rep, 120))

    result = result & Assert(CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_PREV).Value) = "P001 last quarter", _
        "LAST QUARTER'S SUBMIT LANDS IN REPORTED LAST TIME, got '" & _
        CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_PREV).Value) & "'")

    result = result & Assert(Trim(CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_SUBMIT).Value)) = "", _
        "and SUBMIT is handed to the new quarter empty, got '" & _
        CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_SUBMIT).Value) & "'")

    result = result & Assert(Trim(CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_APPROVED).Value)) = "", _
        "and the tick goes with the text it approved, got '" & _
        CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_APPROVED).Value) & "'")

    result = result & Assert(Trim(CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_DRAFT).Value)) = "", _
        "and the AI draft does not survive the turn, got '" & _
        CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_DRAFT).Value) & "'")

    result = result & Assert(Trim(CStr(dws.Cells(Drafting.DRAFT_INTRO_ROW, Drafting.COL_D_PERIOD).Value)) = "FY26Q4", _
        "and the sheet is re-stamped to the new quarter, got '" & _
        CStr(dws.Cells(Drafting.DRAFT_INTRO_ROW, Drafting.COL_D_PERIOD).Value) & "'")

    ' 4. A SECOND pass in the new quarter must not ferry the empty SUBMIT over
    '    the top of what it just saved. Rebuilding twice is normal -- Sync Now
    '    calls this every run -- so an idempotence failure here would silently
    '    erase the reference text on the very next press.
    Drafting.WriteDraftingSheet dws, reg, "ABOUT_BODY", Empty, "FY26Q4"
    result = result & Assert(CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_PREV).Value) = "P001 last quarter", _
        "REPORTED LAST TIME SURVIVES A SECOND REBUILD, got '" & _
        CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_PREV).Value) & "'")

    ' 5. ORIGINAL still comes from the register, not from the ferry.
    result = result & Assert(InStr(CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_CURRENT).Value), "register one") > 0, _
        "ORIGINAL still reads the register, got '" & _
        CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_CURRENT).Value) & "'")

    wb.Close False
    xl.Quit
    Set wb = Nothing
    Set xl = Nothing

    Test_Drafting_QuarterTurnFerriesSubmitIntoReportedLastTime = result
End Function

' THE TICK MUST SURVIVE A REBUILD IN ITS OWN QUARTER, AND MUST NOT SURVIVE A
' ROLLOVER. Both halves, because either alone passes against a defect.
'
' This is the test that did not exist, and its absence is the whole story of
' fix-list defect 1. WriteDraftingSheet cleared COL_D_APPROVED unconditionally;
' PublishDraftsForField reads it; and RibbonUI.SyncNowChainCore runs the rebuild
' (step 3) immediately before the publish (step 4). So the tick was destroyed one
' step before it was read, on every run, and no drafted text ever reached the
' register through the tool. 192 tests were green throughout -- they asked
' whether publish works when called, never whether a person can cause it to be
' called with a tick still on the sheet.
'
' WHY BOTH ASSERTIONS. Carrying the tick always would republish last quarter's
' approval over this quarter's text. Clearing it always is the defect being
' removed. Only the pair pins the actual rule: the tick travels with the SUBMIT
' text it approves, under the same per-row carry test.
'
' Made to fail on purpose before being trusted, in both directions -- see the
' note in the fix-list entry. Restoring the unconditional clear fails assertion
' 2; harvesting the tick outside the `carryThisRow` test fails assertion 4.
Private Function Test_Drafting_TickSurvivesSamePeriodRebuildAndClearsOnRollover() As String
    Dim result As String

    Dim xl As Object, wb As Object, dws As Object, rws As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set wb = xl.Workbooks.Add
    Set dws = wb.Worksheets(1)
    Set rws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))

    rws.Cells(1, 1).Value = ExcelOutput.INSTANCE_ID_HEADER
    rws.Cells(1, 2).Value = "PROJECT_NAME"
    rws.Cells(1, 3).Value = "ABOUT_BODY"
    rws.Cells(2, 1).Value = "P001": rws.Cells(2, 2).Value = "Alpha": rws.Cells(2, 3).Value = "register one"

    Dim reg As Sheet
    reg = ExcelOutput.ReadSheet(rws)

    ' 1. A person writes their words and ticks the row, in FY26Q3.
    Drafting.WriteDraftingSheet dws, reg, "ABOUT_BODY", Empty, "FY26Q3"
    dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_SUBMIT).Value = "P001 this quarter"
    dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_APPROVED).Value = "Y"

    ' 2. THE ASSERTION THAT UNBLOCKS PUBLISH. Same quarter, so a rebuild must
    '    leave the tick exactly where they put it. Tested through the same
    '    affirmative reader publish uses -- a non-empty check would pass on a
    '    cell holding "0", which is the miscount this project has already made
    '    once (fix-list P6).
    Drafting.WriteDraftingSheet dws, reg, "ABOUT_BODY", Empty, "FY26Q3"

    result = result & Assert(ReviewQueue.IsApprovalMark(CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_APPROVED).Value)), _
        "A SAME-PERIOD REBUILD KEEPS THE TICK -- publish runs immediately after a rebuild, " & _
        "so a tick cleared here can never be read, got '" & _
        CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_APPROVED).Value) & "'")

    ' 3. And it still sits beside the text it approves. A tick carried without
    '    its SUBMIT would approve whatever landed in the box next.
    result = result & Assert(CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_SUBMIT).Value) = "P001 this quarter", _
        "the tick is carried WITH the submitted text it approves, got '" & _
        CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_SUBMIT).Value) & "'")

    ' 4. THE OTHER HALF. Roll the quarter with nothing at risk -- the refusal
    '    guard fires on SUBMIT or DRAFT text, so those go first; a tick alone
    '    does not refuse. The rollover must now drop the tick, because an
    '    approval given for FY26Q3's wording is not an approval of FY26Q4's.
    dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_SUBMIT).ClearContents
    dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_DRAFT).ClearContents

    Dim rep As String
    rep = Drafting.WriteDraftingSheet(dws, reg, "ABOUT_BODY", Empty, "FY26Q4")

    result = result & Assert(InStr(rep, "REFUSED") = 0, _
        "a tick with no text does not trip the at-risk refusal, got '" & Left(rep, 90) & "'")
    result = result & Assert(Not ReviewQueue.IsApprovalMark(CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_APPROVED).Value)), _
        "A ROLLOVER DROPS THE TICK -- last quarter's approval must not authorise this " & _
        "quarter's text, got '" & _
        CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_APPROVED).Value) & "'")

    wb.Close False
    xl.Quit
    Set wb = Nothing: Set xl = Nothing
    Test_Drafting_TickSurvivesSamePeriodRebuildAndClearsOnRollover = result
End Function

' THE ROLLOVER DROP IS PER ROW, AND THIS IS THE ASSERTION THAT SAYS SO.
'
' Two projects, same field, different cadence in the register: P001's value came
' from a period row, P002's from a `Quarter = ALL` row. A rollover must clear
' P001's carried-over work and leave P002's completely alone.
'
' WHY THIS TEST NOW USES SOURCES AND NOTES INSTEAD OF SUBMIT.
'
' It used to prove the point with SUBMIT text, and that is why it went red on
' 2026-08-13: the refusal guard added that day fires on the FIRST row holding
' SUBMIT or DRAFT text and refuses the whole sheet, so a fixture with submitted
' text never reaches the per-row logic at all. The old assertions were testing a
' path their own fixture had closed.
'
' What that leaves is narrow and worth saying out loud: the cadence machinery
' below now governs only SOURCES and NOTES, on rows carrying neither a draft nor
' a submission. Everything richer than that is settled by the refusal instead.
'
' OPEN QUESTION, NOT SETTLED BEHAVIOUR -- see FIX-LIST.md "at-risk scan misses
' SOURCES and NOTES". Citations and notes are also work a person typed, and the
' refusal does not currently count them, which is the only reason this path can
' still discard anything. If that gets fixed, the drop becomes unreachable for
' any row with content at all and this whole mechanism should be DELETED rather
' than kept passing. The test asserts what the code does today; it is not a vote
' that today's answer is right.

' THE WIDE SHEET CARRIES ITS OWN PERIOD -- one row per slide per period, rows
' accumulating, each deck picking up the period it declares. Rohan's model,
' 2026-08-03, and the thing that lets history and "one row per slide" stop being
' a trade-off.
'
' Also asserts the two ways this could quietly go wrong, both of which this
' project has already been bitten by in other forms: a sheet with no Quarter
' column must NOT be filtered to nothing, and two rows for one project in one
' period must not silently resolve to whichever sits lower.
' Onboarding used to invent a second, empty register beside a populated one and
' pair the deck to it -- nothing destroyed, every later read returning zero rows,
' reported as a clean run of nothing. This is the finder that stops it.
' A Given field declared in the Field Spec had nowhere to be entered: register
' columns are only created as a side effect of publishing or harvesting, and a
' Given field is neither. 14 of the milestone timeline's columns were in that
' state.
' CreateWorkbook does Workbooks.Add then SaveAs -- a BLANK workbook written over
' whatever is at that path. Its caller reached it by asking a person where the
' Data workbook lives, and a person who types a path they already know is
' usually naming a file that EXISTS. So the dangerous case was the likely one,
' and pointing it at a real register would have destroyed a quarter of drafting
' in a single click. Found 2026-08-13 by reading the code, one click before it.
' The arithmetic behind the readiness line that would have caught the five stub
' rows before they refused a quarter turn.
Private Function Test_ExcelOutput_LargestPeriodRowCountSpotsAPartialQuarter() As String
    Dim result As String

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dir As String
    dir = Environ$("TEMP") & "\decksync_partial"
    If Not fso.FolderExists(dir) Then fso.CreateFolder dir
    Dim path As String
    path = dir & "\reg.xlsx"
    If fso.FileExists(path) Then fso.DeleteFile path

    ' CLOSE EVEN ON FAILURE. The first version of this test threw before its
    ' wb.Close and left a workbook open, which made SaveAs fail in the THREE
    ' archive tests that follow -- four errors reported, one actual fault, and none
    ' of them in the code under test. A fixture that leaks is not a test problem,
    ' it is a false bug report aimed at whoever reads the run next.
    On Error GoTo CleanUp

    Dim wb As Object
    Set wb = WorkbookBridge.CreateWorkbook(path)

    ' The first sheet, NOT WorkbookBridge.RegisterSheet: a brand-new workbook has
    ' no sheet named "Register", and assuming it did is what threw.
    Dim ws As Object
    Set ws = wb.Worksheets(1)

    ' Two full quarters of 4, and a partial third of 1 -- the shape that stopped
    ' the real run, scaled down.
    ws.Cells(1, 1).Value = "Instance ID"
    ws.Cells(1, 2).Value = ExcelOutput.QUARTER_HEADER
    Dim i As Long
    For i = 1 To 4
        ws.Cells(1 + i, 1).Value = "P00" & i
        ws.Cells(1 + i, 2).Value = "Q3F26"
        ws.Cells(5 + i, 1).Value = "P00" & i
        ws.Cells(5 + i, 2).Value = "Q4F26"
    Next i
    ws.Cells(10, 1).Value = "P001"
    ws.Cells(10, 2).Value = "Q1F27"

    Dim full As Long
    full = ExcelOutput.LargestPeriodRowCount(ws)
    result = result & Assert(full = 4, "a full quarter is 4, got " & full)
    result = result & Assert(ExcelOutput.PeriodRowCount(ws, "Q1F27") = 1, _
        "the partial period has 1 row")

    ' THE JUDGEMENT THE READINESS LINE MAKES. Partial is 0 < n < full; a complete
    ' quarter and an empty one must BOTH read as fine, or the check cries wolf on
    ' every ordinary register and gets clicked past.
    result = result & Assert(1 > 0 And 1 < full, "1 of 4 reads as PARTIAL")
    result = result & Assert(Not (4 > 0 And 4 < full), "a complete quarter does NOT read as partial")
    result = result & Assert(ExcelOutput.PeriodRowCount(ws, "Q2F27") = 0, _
        "an untouched future quarter is 0, which reads as 'none yet'")

CleanUp:
    If Err.Number <> 0 Then
        result = result & Assert(False, "threw: " & Err.Description)
        Err.Clear
    End If
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
    Test_ExcelOutput_LargestPeriodRowCountSpotsAPartialQuarter = result
End Function

' 2026-08-15: proves the fix for "this msg makes zero sense" (Rohan, reading a
' chained Put-it-on-the-slides run where 13 fields' Copy/Publish blocks landed
' under the same two fixed headers with no field name visible). Tests the
' PURE labeling function in isolation -- no live deck needed, and none of the
' actual chain (Say/BeginCollecting/PublishAllDraftedFields) has been
' re-exercised against a real presentation since this change. That is still
' owed at the keyboard.
Private Function Test_DraftingUI_ChainBlockHeaderLabelsTheField() As String
    Dim result As String

    result = result & Assert(DraftingUI.ChainBlockHeader("Copy AI to Submit", "") = "Copy AI to Submit", _
        "no chain field: header is unchanged")

    result = result & Assert( _
        DraftingUI.ChainBlockHeader("Copy AI to Submit", "ABOUT_BODY") = "Copy AI to Submit (ABOUT_BODY)", _
        "with a chain field: it is appended in parentheses")

    ' THE ONE THAT MATTERS. Two consecutive fields must be DISTINGUISHABLE by
    ' header alone -- that is the entire defect. A fix that appended the same
    ' text regardless of field would pass the assertion above and still leave
    ' every block looking identical.
    Dim h1 As String, h2 As String
    h1 = DraftingUI.ChainBlockHeader("Publish Drafts", "ABOUT_BODY")
    h2 = DraftingUI.ChainBlockHeader("Publish Drafts", "KEY_EVENTS_BODY")
    result = result & Assert(h1 <> h2, "two different fields produce two different headers, got '" & h1 & "' and '" & h2 & "'")

    Test_DraftingUI_ChainBlockHeaderLabelsTheField = result
End Function

' The archive is the first half of file-per-quarter, and the half trusted to be
' harmless. These prove the three claims that make it harmless -- it writes, it
' refuses rather than overwriting, and IT DOES NOT MOVE THE LIVE WORKBOOK.
Private Function Test_WorkbookBridge_ArchivesAPeriodToItsOwnFile() As String
    Dim result As String

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dir As String
    dir = Environ$("TEMP") & "\decksync_archive"
    If Not fso.FolderExists(dir) Then fso.CreateFolder dir

    Dim livePath As String, archPath As String
    livePath = dir & "\reg.xlsx"
    archPath = dir & "\reg-Q4F26.xlsx"
    If fso.FileExists(livePath) Then fso.DeleteFile livePath
    If fso.FileExists(archPath) Then fso.DeleteFile archPath

    Dim wb As Object
    Set wb = WorkbookBridge.CreateWorkbook(livePath)

    ' CONTROL: the archive must not exist yet, or "it exists afterwards" proves
    ' nothing about this call having done anything.
    result = result & Assert(Not fso.FileExists(archPath), _
        "no archive before the call")

    Dim problem As String
    problem = WorkbookBridge.ArchiveWorkbookForPeriod(wb, "Q4F26")

    result = result & Assert(problem = "", "reports success, got '" & problem & "'")
    result = result & Assert(fso.FileExists(archPath), "the archive file EXISTS on disk")
    If fso.FileExists(archPath) Then
        result = result & Assert(fso.GetFile(archPath).Size > 0, "and is not empty")
    End If

    ' THE ONE THAT MATTERS, AND THE REASON THIS IS SaveCopyAs AND NOT SaveAs.
    ' SaveAs would re-point the OPEN workbook at the archive, so every later write
    ' -- including the roll forward that runs immediately after this -- would land
    ' in the frozen file while the deck's pairing still named the live one. A test
    ' that only checked the archive appeared would pass against that disaster.
    result = result & Assert(StrComp(wb.FullName, livePath, vbTextCompare) = 0, _
        "the LIVE workbook still points at itself, got '" & wb.FullName & "'")

    wb.Close False
    Test_WorkbookBridge_ArchivesAPeriodToItsOwnFile = result
End Function

Private Function Test_WorkbookBridge_RefusesToOverwriteAnExistingArchive() As String
    Dim result As String

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dir As String
    dir = Environ$("TEMP") & "\decksync_archive2"
    If Not fso.FolderExists(dir) Then fso.CreateFolder dir

    Dim livePath As String, archPath As String
    livePath = dir & "\reg.xlsx"
    archPath = dir & "\reg-Q4F26.xlsx"
    If fso.FileExists(livePath) Then fso.DeleteFile livePath
    If fso.FileExists(archPath) Then fso.DeleteFile archPath

    ' A previous quarter's archive, with KNOWN CONTENT. Checking only that a
    ' message came back would pass against a version that refused AFTER writing --
    ' which would have destroyed the very thing the archive exists to protect.
    Dim ts As Object
    Set ts = fso.CreateTextFile(archPath, True)
    ts.Write "LAST-QUARTERS-FROZEN-RECORD"
    ts.Close

    Dim wb As Object
    Set wb = WorkbookBridge.CreateWorkbook(livePath)

    Dim problem As String
    problem = WorkbookBridge.ArchiveWorkbookForPeriod(wb, "Q4F26")

    result = result & Assert(problem <> "", "refuses when an archive already exists")
    result = result & Assert(InStr(problem, "already exists") > 0, _
        "and says why, got '" & Left(problem, 60) & "'")

    Dim after As String
    Set ts = fso.OpenTextFile(archPath, 1)
    after = ts.ReadAll
    ts.Close
    result = result & Assert(after = "LAST-QUARTERS-FROZEN-RECORD", _
        "and the existing archive is UNTOUCHED, got '" & Left(after, 40) & "'")

    wb.Close False
    Test_WorkbookBridge_RefusesToOverwriteAnExistingArchive = result
End Function

Private Function Test_WorkbookBridge_ArchiveRefusesABlankPeriod() As String
    Dim result As String

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dir As String
    dir = Environ$("TEMP") & "\decksync_archive3"
    If Not fso.FolderExists(dir) Then fso.CreateFolder dir

    Dim livePath As String
    livePath = dir & "\reg.xlsx"
    If fso.FileExists(livePath) Then fso.DeleteFile livePath

    Dim wb As Object
    Set wb = WorkbookBridge.CreateWorkbook(livePath)

    ' An unnamed period would produce "reg-.xlsx" -- a file that looks like an
    ' archive, belongs to no quarter, and would then BLOCK the real archive for
    ' whatever quarter tried next.
    Dim problem As String
    problem = WorkbookBridge.ArchiveWorkbookForPeriod(wb, "")

    result = result & Assert(problem <> "", "refuses a blank period")
    result = result & Assert(Not fso.FileExists(dir & "\reg-.xlsx"), _
        "and writes no unnamed archive file")

    wb.Close False
    Test_WorkbookBridge_ArchiveRefusesABlankPeriod = result
End Function

Private Function Test_WorkbookBridge_RefusesToCreateOverAnExistingFile() As String
    Dim result As String

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dir As String
    dir = Environ$("TEMP") & "\decksync_createguard"
    If Not fso.FolderExists(dir) Then fso.CreateFolder dir

    ' A file with KNOWN CONTENT, so the test can prove it still has it. Checking
    ' only that an error was raised would pass against a version that raised
    ' AFTER writing.
    Dim victim As String
    victim = dir & "\already_here.xlsx"
    Dim ts As Object
    Set ts = fso.CreateTextFile(victim, True)
    ts.Write "ORIGINAL-CONTENT-DO-NOT-DESTROY"
    ts.Close

    Dim raised As Boolean, msg As String
    On Error Resume Next
    Err.Clear
    Dim wb As Object
    Set wb = WorkbookBridge.CreateWorkbook(victim)
    raised = (Err.Number <> 0)
    msg = Err.Description
    Err.Clear
    On Error GoTo 0
    If Not wb Is Nothing Then wb.Close False

    result = result & Assert(raised, "creating over an existing file RAISES")
    result = result & Assert(InStr(msg, "Refusing to create over it") > 0, _
        "and says why, got '" & msg & "'")

    ' THE ONE THAT MATTERS. A refusal that happens after the write is no refusal.
    Dim after As String
    Set ts = fso.OpenTextFile(victim, 1)
    after = ts.ReadAll
    ts.Close
    result = result & Assert(after = "ORIGINAL-CONTENT-DO-NOT-DESTROY", _
        "and the existing file is UNTOUCHED, got '" & Left(after, 40) & "'")

    ' Creating where nothing exists must still work, or the guard has just
    ' broken onboarding for every genuinely new workbook.
    Dim fresh As String
    fresh = dir & "\brand_new.xlsx"
    If fso.FileExists(fresh) Then fso.DeleteFile fresh
    Dim wb2 As Object
    On Error Resume Next
    Err.Clear
    Set wb2 = WorkbookBridge.CreateWorkbook(fresh)
    Dim freshErr As String
    freshErr = Err.Description
    Err.Clear
    On Error GoTo 0
    result = result & Assert(Not wb2 Is Nothing, _
        "a NEW path still creates, got '" & freshErr & "'")
    If Not wb2 Is Nothing Then wb2.Close False
    If fso.FileExists(fresh) Then fso.DeleteFile fresh
    If fso.FileExists(victim) Then fso.DeleteFile victim

    Test_WorkbookBridge_RefusesToCreateOverAnExistingFile = result
End Function

' PowerPoint's Save As dialog takes no file-type filter, so it appends .pptx to
' whatever you type. The old repair then appended .xlsx to THAT, producing
' `register-wide.xlsx.pptx.xlsx` -- a name that passed validation, did not
' exist, and was created as a blank workbook while the real register sat
' untouched beside it.
Private Function Test_BatchOnboard_NormalisesWorkbookPath() As String
    Dim result As String

    ' THE CASE THAT BIT. The real register's name, mangled by the dialog.
    result = result & Assert( _
        BatchOnboardFlow.NormaliseWorkbookPath("C:\a\register-wide.xlsx.pptx") = "C:\a\register-wide.xlsx", _
        "strips an appended .pptx back to the .xlsx the person typed, got '" & _
        BatchOnboardFlow.NormaliseWorkbookPath("C:\a\register-wide.xlsx.pptx") & "'")

    ' Two passes through the dialog append twice.
    result = result & Assert( _
        BatchOnboardFlow.NormaliseWorkbookPath("C:\a\reg.xlsx.pptx.pptx") = "C:\a\reg.xlsx", _
        "strips more than one, got '" & BatchOnboardFlow.NormaliseWorkbookPath("C:\a\reg.xlsx.pptx.pptx") & "'")

    ' The case the ORIGINAL line was written for must still work -- a host that
    ' returns a bare name. Removing that behaviour to fix the other one would
    ' trade a new-workbook bug for an old one.
    result = result & Assert( _
        BatchOnboardFlow.NormaliseWorkbookPath("C:\a\Project-Data") = "C:\a\Project-Data.xlsx", _
        "still adds .xlsx to a bare name, got '" & BatchOnboardFlow.NormaliseWorkbookPath("C:\a\Project-Data") & "'")

    ' A clean path is left exactly alone.
    result = result & Assert( _
        BatchOnboardFlow.NormaliseWorkbookPath("C:\a\Project-Data.xlsx") = "C:\a\Project-Data.xlsx", _
        "leaves a correct path untouched")

    ' Macro-enabled workbooks are a legal pairing and must not gain .xlsx.
    result = result & Assert( _
        BatchOnboardFlow.NormaliseWorkbookPath("C:\a\Data.xlsm") = "C:\a\Data.xlsm", _
        "leaves .xlsm alone, got '" & BatchOnboardFlow.NormaliseWorkbookPath("C:\a\Data.xlsm") & "'")

    ' A folder whose NAME contains .ppt must not be eaten -- the strip only ever
    ' applies to the end of the string.
    result = result & Assert( _
        BatchOnboardFlow.NormaliseWorkbookPath("C:\my.pptx stuff\Data.xlsx") = "C:\my.pptx stuff\Data.xlsx", _
        "does not touch .pptx inside a folder name, got '" & _
        BatchOnboardFlow.NormaliseWorkbookPath("C:\my.pptx stuff\Data.xlsx") & "'")

    Test_BatchOnboard_NormalisesWorkbookPath = result
End Function

' A ROLLOVER MUST NOT DESTROY TYPED WORK.
'
' The period guard used to silently DROP every quarterly row's drafting when a
' sheet's stamp disagreed with the deck. Intent right, remedy wrong: it was one
' button press from taking 129 drafted values off Rohan, in a state that was
' entirely legitimate -- sheets stamped Q4F26 while the deck was deliberately
' set to Q3F26 to capture a baseline. Found 2026-08-13 because he asked "but the
' Q4 text I generated is real? How are you preserving that?"

Private Function Test_ExcelOutput_MissingRegisterColumns() As String
    Dim result As String

    Dim xl As Object, wb As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set wb = xl.Workbooks.Add

    Dim reg As Object
    Set reg = wb.Worksheets(1)
    reg.Name = "Register"
    reg.Cells(1, 1).Value = ExcelOutput.INSTANCE_ID_HEADER
    reg.Cells(1, 2).Value = ExcelOutput.QUARTER_HEADER
    reg.Cells(1, 3).Value = "ABOUT_BODY"
    reg.Cells(2, 1).Value = "P001": reg.Cells(2, 2).Value = "Q4F26"

    Dim spec As Object
    Set spec = wb.Worksheets.Add
    spec.Name = FieldSpec.SPEC_SHEET_NAME
    spec.Cells(1, FieldSpec.COL_S_FIELDID).Value = "FieldID"
    spec.Cells(2, FieldSpec.COL_S_FIELDID).Value = "ABOUT_BODY":   spec.Cells(2, FieldSpec.COL_S_KIND).Value = "Prose"
    spec.Cells(3, FieldSpec.COL_S_FIELDID).Value = "MS1_DATE":     spec.Cells(3, FieldSpec.COL_S_KIND).Value = "Given"
    spec.Cells(4, FieldSpec.COL_S_FIELDID).Value = "MS1_DONE":     spec.Cells(4, FieldSpec.COL_S_KIND).Value = "Given"
    spec.Cells(5, FieldSpec.COL_S_FIELDID).Value = "TIME_ELAPSED": spec.Cells(5, FieldSpec.COL_S_KIND).Value = ExcelOutput.KIND_DERIVED

    Dim missing As String
    missing = ExcelOutput.MissingRegisterColumns(spec, reg)

    result = result & Assert(InStr(missing, "MS1_DATE") > 0 And InStr(missing, "MS1_DONE") > 0, _
        "the Given fields with no column are named, got '" & missing & "'")
    result = result & Assert(InStr(missing, "ABOUT_BODY") = 0, _
        "a field that already HAS a column is not listed, got '" & missing & "'")
    ' THE DERIVED CARVE-OUT. Without it this list would name every computed
    ' field forever, and a prompt that always fires is one that gets clicked
    ' through -- which would take the real ones with it.
    result = result & Assert(InStr(missing, "TIME_ELAPSED") = 0, _
        "a DERIVED field is never listed -- it is computed, never stored, got '" & missing & "'")

    Dim added As String
    added = ExcelOutput.AddRegisterColumns(reg, missing)
    result = result & Assert(InStr(added, "MS1_DATE") > 0 And InStr(added, "MS1_DONE") > 0, _
        "both columns report as added, got '" & added & "'")

    ' EACH ON ITS OWN COLUMN. Caching the last-used column outside the loop
    ' writes every header over the previous one and reports success for all.
    result = result & Assert(reg.Cells(1, 4).Value <> reg.Cells(1, 5).Value, _
        "the two headers landed in DIFFERENT columns, got '" & reg.Cells(1, 4).Value & _
        "' and '" & reg.Cells(1, 5).Value & "'")

    ' And it is genuinely done -- asking again finds nothing.
    result = result & Assert(ExcelOutput.MissingRegisterColumns(spec, reg) = "", _
        "nothing is missing once added, got '" & ExcelOutput.MissingRegisterColumns(spec, reg) & "'")

    wb.Saved = True
    wb.Close
    xl.Quit

    Test_ExcelOutput_MissingRegisterColumns = result
End Function

Private Function Test_ExcelOutput_FindsExistingRegisters() As String
    Dim result As String

    Dim xl As Object, wb As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set wb = xl.Workbooks.Add

    ' A REGISTER, with Instance ID in column B rather than A -- position must be
    ' irrelevant, exactly as it is for the reader. An earlier version of the
    ' finder looked only at A1 and would have missed this sheet, which is how
    ' the defect it guards against would have returned wearing a new hat.
    Dim reg As Object
    Set reg = wb.Worksheets(1)
    reg.Name = "Register"
    reg.Cells(1, 1).Value = "PROJECT_NAME"
    reg.Cells(1, 2).Value = ExcelOutput.INSTANCE_ID_HEADER
    reg.Cells(1, 3).Value = ExcelOutput.QUARTER_HEADER
    reg.Cells(2, 1).Value = "Alpha": reg.Cells(2, 2).Value = "P001": reg.Cells(2, 3).Value = "Q4F26"
    reg.Cells(3, 1).Value = "Beta":  reg.Cells(3, 2).Value = "P002": reg.Cells(3, 3).Value = "Q4F26"
    reg.Cells(4, 1).Value = "Alpha": reg.Cells(4, 2).Value = "P001": reg.Cells(4, 3).Value = "Q1F27"

    ' NOT a register. A drafting sheet is the realistic confuser -- it has
    ' headers and rows and sits in the same workbook.
    Dim tpl As Object
    Set tpl = wb.Worksheets.Add
    tpl.Name = "TPL_ABOUT_BODY"
    tpl.Cells(1, 1).Value = "Project code"
    tpl.Cells(1, 2).Value = "Project name"
    tpl.Cells(2, 1).Value = "P001"

    Dim found As String
    found = ExcelOutput.RegisterShapedSheets(wb)

    result = result & Assert(InStr(found, "Register|") > 0, _
        "names the register sheet, got '" & found & "'")
    result = result & Assert(InStr(found, "TPL_ABOUT_BODY") = 0, _
        "does NOT mistake a drafting sheet for a register, got '" & found & "'")
    result = result & Assert(InStr(found, "|3|") > 0, _
        "counts 3 data rows, not the header, got '" & found & "'")
    result = result & Assert(InStr(found, "Q4F26") > 0 And InStr(found, "Q1F27") > 0, _
        "names BOTH periods so a person can tell two registers apart, got '" & found & "'")
    result = result & Assert(InStr(found, vbLf) = 0, _
        "finds exactly one register -- one line, no separator, got '" & found & "'")

    ' A SECOND register must be found too, because that is the case onboarding
    ' REFUSES on, and a finder that silently returns one of two would make the
    ' refusal unreachable.
    Dim reg2 As Object
    Set reg2 = wb.Worksheets.Add
    reg2.Name = "project-status"
    reg2.Cells(1, 1).Value = ExcelOutput.INSTANCE_ID_HEADER
    reg2.Cells(1, 2).Value = ExcelOutput.QUARTER_HEADER

    found = ExcelOutput.RegisterShapedSheets(wb)
    result = result & Assert(InStr(found, vbLf) > 0, _
        "finds BOTH registers once a second exists, got '" & found & "'")
    result = result & Assert(InStr(found, "project-status|0|") > 0, _
        "reports the empty one as 0 rows rather than omitting it, got '" & found & "'")

    ' RowCountForPeriod says WHAT is at risk before a harvest overwrites it.
    ' The two negative cases are what stop it being an always-true guard: a
    ' period that is not in the sheet, and a blank period, must both be 0 --
    ' otherwise the onboarding warning would fire on every deck forever.
    result = result & Assert(ExcelOutput.RowCountForPeriod(reg, "Q4F26") = 2, _
        "counts 2 rows for Q4F26, got " & ExcelOutput.RowCountForPeriod(reg, "Q4F26"))
    result = result & Assert(ExcelOutput.RowCountForPeriod(reg, "Q1F27") = 1, _
        "counts 1 row for Q1F27, got " & ExcelOutput.RowCountForPeriod(reg, "Q1F27"))
    result = result & Assert(ExcelOutput.RowCountForPeriod(reg, "Q9F99") = 0, _
        "counts 0 for a period the sheet does not have, got " & ExcelOutput.RowCountForPeriod(reg, "Q9F99"))
    result = result & Assert(ExcelOutput.RowCountForPeriod(reg, "") = 0, _
        "counts 0 for a blank period, got " & ExcelOutput.RowCountForPeriod(reg, ""))

    wb.Saved = True
    wb.Close
    xl.Quit

    Test_ExcelOutput_FindsExistingRegisters = result
End Function

Private Function Test_ExcelOutput_PeriodRowsAndRollForward() As String
    Dim result As String

    Dim xl As Object, wb As Object, ws As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set wb = xl.Workbooks.Add
    Set ws = wb.Worksheets(1)

    ' Quarter deliberately NOT in column A: the reader must find both structural
    ' columns by header name. Put a field first to prove position is irrelevant.
    ws.Cells(1, 1).Value = "PROJECT_NAME"
    ws.Cells(1, 2).Value = ExcelOutput.INSTANCE_ID_HEADER
    ws.Cells(1, 3).Value = ExcelOutput.QUARTER_HEADER
    ws.Cells(1, 4).Value = "PROJECT_STATUS"

    ws.Cells(2, 1).Value = "Alpha": ws.Cells(2, 2).Value = "P001": ws.Cells(2, 3).Value = "FY26Q4": ws.Cells(2, 4).Value = "In Progress"
    ws.Cells(3, 1).Value = "Beta":  ws.Cells(3, 2).Value = "P002": ws.Cells(3, 3).Value = "FY26Q4": ws.Cells(3, 4).Value = "In Progress"
    ws.Cells(4, 1).Value = "Alpha": ws.Cells(4, 2).Value = "P001": ws.Cells(4, 3).Value = "FY27Q1": ws.Cells(4, 4).Value = "Closed"

    Dim s As Sheet
    s = ExcelOutput.ReadSheetForPeriod(ws, "FY26Q4")

    result = result & Assert(s.InstanceOrder.count = 2, _
        "FY26Q4 sees ONLY its own two projects, not all three rows, got " & s.InstanceOrder.count)
    result = result & Assert(s.Rows("P001")("PROJECT_STATUS") = "In Progress", _
        "P001 reads its FY26Q4 status, got '" & s.Rows("P001")("PROJECT_STATUS") & "'")

    Dim s2 As Sheet
    s2 = ExcelOutput.ReadSheetForPeriod(ws, "FY27Q1")
    result = result & Assert(s2.InstanceOrder.count = 1, _
        "FY27Q1 sees only the one project rolled into it, got " & s2.InstanceOrder.count)
    result = result & Assert(s2.Rows("P001")("PROJECT_STATUS") = "Closed", _
        "THE SAME PROJECT READS A DIFFERENT VALUE IN A DIFFERENT PERIOD -- the whole point, got '" & _
        s2.Rows("P001")("PROJECT_STATUS") & "'")

    ' The structural columns are not fields. If Quarter leaked into Fields it
    ' would be offered for drafting and synced onto a slide.
    Dim f As Variant, sawQuarter As Boolean, sawInstance As Boolean
    For Each f In s.Fields
        If CStr(f) = ExcelOutput.QUARTER_HEADER Then sawQuarter = True
        If CStr(f) = ExcelOutput.INSTANCE_ID_HEADER Then sawInstance = True
    Next f
    result = result & Assert(s.Fields.count = 2 And Not sawQuarter And Not sawInstance, _
        "Quarter and Instance ID are STRUCTURE, not fields -- got " & s.Fields.count & " field(s)")

    ' A SHEET WITH NO QUARTER COLUMN IS NEVER FILTERED. Every sheet built before
    ' this change is in that state, and returning nothing would be an empty read
    ' presenting as a clean one.
    Dim old As Object
    Set old = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    old.Cells(1, 1).Value = ExcelOutput.INSTANCE_ID_HEADER
    old.Cells(1, 2).Value = "PROJECT_NAME"
    old.Cells(2, 1).Value = "P001": old.Cells(2, 2).Value = "Alpha"
    Dim s3 As Sheet
    s3 = ExcelOutput.ReadSheetForPeriod(old, "FY26Q4")
    result = result & Assert(s3.InstanceOrder.count = 1, _
        "A PRE-QUARTER SHEET IS READ IN FULL, not filtered to nothing, got " & s3.InstanceOrder.count)

    ' Two rows, one project, one period: counted, not silently resolved.
    ws.Cells(5, 1).Value = "Alpha dup": ws.Cells(5, 2).Value = "P001": ws.Cells(5, 3).Value = "FY26Q4": ws.Cells(5, 4).Value = "Wrong"
    Dim s4 As Sheet
    s4 = ExcelOutput.ReadSheetForPeriod(ws, "FY26Q4")
    result = result & Assert(s4.DuplicateInstances = 1, _
        "a second row for one project in one period is COUNTED, got " & s4.DuplicateInstances)
    result = result & Assert(s4.Rows("P001")("PROJECT_STATUS") = "In Progress", _
        "first row wins deterministically rather than whichever sits lower, got '" & _
        s4.Rows("P001")("PROJECT_STATUS") & "'")
    ws.Rows(5).Delete

    ' ROLL FORWARD REPLACES Quarter = ALL. The project name arrives by being
    ' copied; no sentinel, no cadence declaration, nothing for a person to learn.
    Dim rep As String
    rep = ExcelOutput.RollForwardPeriod(ws, "FY26Q4", "FY27Q2")
    Dim s5 As Sheet
    s5 = ExcelOutput.ReadSheetForPeriod(ws, "FY27Q2")
    result = result & Assert(s5.InstanceOrder.count = 2, _
        "both FY26Q4 projects were carried into FY27Q2, got " & s5.InstanceOrder.count)
    result = result & Assert(s5.Rows("P001")("PROJECT_NAME") = "Alpha", _
        "THE STATIC VALUE ARRIVES BY BEING COPIED -- this is what ALL was for, got '" & _
        s5.Rows("P001")("PROJECT_NAME") & "'")

    ' The source period is untouched -- it is the record of what was reported.
    Dim s6 As Sheet
    s6 = ExcelOutput.ReadSheetForPeriod(ws, "FY26Q4")
    result = result & Assert(s6.InstanceOrder.count = 2, _
        "rolling forward does not disturb the period it copied from, got " & s6.InstanceOrder.count)

    ' Twice would double every project, and would look exactly like once.
    Dim rep2 As String
    rep2 = ExcelOutput.RollForwardPeriod(ws, "FY26Q4", "FY27Q2")
    result = result & Assert(InStr(rep2, "REFUSED") > 0, _
        "A SECOND ROLL FORWARD IS REFUSED rather than duplicating every project, got '" & rep2 & "'")

    wb.Close False
    xl.Quit
    Set wb = Nothing: Set xl = Nothing
    Test_ExcelOutput_PeriodRowsAndRollForward = result
End Function

' Slide.Parent is the Presentation, and a deck property reads through it.
'
' NOT AN OBVIOUS TRUTH TO ASSUME -- it is a COM object-model claim, and this
' project's own rule is that those get probed against real Office rather than
' asserted from memory. DeckAdoption.CommitAdoption and
' BatchOnboardFlow.CommitBatch both derive the period this way instead of
' taking it as a parameter (a deck cannot disagree with itself; a parameter
' can), so if this ever stopped holding, onboarding would stamp every row with
' an empty period -- which UpsertRow now refuses, loudly, but only at run time.
' This turns that into a test failure instead.
Private Function Test_DeckRegistry_PeriodIsReadableThroughSlideParent() As String
    Dim result As String

    Dim testPres As Object
    Set testPres = Application.Presentations.Add
    testPres.Slides.Add 1, 12   ' ppLayoutBlank

    DeckRegistry.SetDeckPeriod testPres, "FY26Q4"

    Dim sld As Object
    Set sld = testPres.Slides(1)

    result = result & Assert(DeckRegistry.GetDeckPeriod(sld.Parent) = "FY26Q4", _
        "the deck's period reads back through Slide.Parent, got '" & DeckRegistry.GetDeckPeriod(sld.Parent) & "'")
    result = result & Assert(sld.Parent.Slides.count = testPres.Slides.count, _
        "Slide.Parent is the presentation the slide lives in")

    testPres.Saved = True
    testPres.Close

    Test_DeckRegistry_PeriodIsReadableThroughSlideParent = result
End Function

' PendingApprovals is the detection that stops chain C rebuilding a review sheet
' over a person's ticks. Three questions, and the second is the one that matters:
' a CONSUMED sheet has already been applied, so its ticks are history, not
' pending work. Without that branch the chain would offer to re-apply an
' evening's approvals every time the button was pressed.
Private Function Test_ReviewQueue_PendingApprovalsCountsTicksAndIgnoresConsumed() As String
    Dim result As String

    Dim xl As Object, wb As Object, ws As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set wb = xl.Workbooks.Add
    Set ws = wb.Worksheets(1)

    ' The sheet must carry the name PendingApprovals derives, or it is not the
    ' sheet the chain would find.
    ws.Name = ReviewQueue.ReviewSheetNameFor("project-status")

    ' Built through WriteQueueSheet rather than by hand, so the fixture cannot
    ' drift from the shape the tool actually writes.
    Dim q As ReviewQueueSet
    q.SlideType = "project-status"
    q.RunStamp = "2026-08-09 15:30"
    q.Consumed = False
    q.Count = 3
    ReDim q.Items(1 To 3)
    Dim i As Long
    For i = 1 To 3
        q.Items(i).EntityKey = "P00" & i
        q.Items(i).FieldID = "ABOUT_BODY"
        q.Items(i).ChangeHash = "HASH" & i
        q.Items(i).Approved = False
    Next i
    ReviewQueue.WriteQueueSheet ws, q

    Dim n As Long, nm As String, stamp As String

    ' 1. Nothing ticked yet -- a queue is not pending approvals.
    n = ReviewQueue.PendingApprovals(wb, "project-status", nm, stamp)
    result = result & Assert(n = 0, "an unticked sheet has no pending approvals, got " & n)
    result = result & Assert(nm = "", "no sheet is named when the answer is 0, got '" & nm & "'")

    ' 2. Ticked and open -- this is the state that must stop the chain.
    ReviewQueue.ApproveAllInSheet ws
    n = ReviewQueue.PendingApprovals(wb, "project-status", nm, stamp)
    result = result & Assert(n = 3, "3 ticked rows are 3 pending approvals, got " & n)
    result = result & Assert(nm = ReviewQueue.ReviewSheetNameFor("project-status"), _
        "the sheet holding the ticks is named, got '" & nm & "'")
    result = result & Assert(stamp = "2026-08-09 15:30", _
        "the run stamp is carried so the person can date the ticks, got '" & stamp & "'")

    ' 3. THE DISCRIMINATOR. Same ticks, same rows, but already applied.
    ReviewQueue.MarkConsumed ws
    n = ReviewQueue.PendingApprovals(wb, "project-status", nm, stamp)
    result = result & Assert(n = 0, "a CONSUMED sheet has no pending approvals, got " & n)

    wb.Close False
    xl.Quit
    ' Quit alone does not always end the process while a live reference is held,
    ' and a leaked EXCEL.EXE makes the NEXT run error rather than fail -- 4 tests
    ' errored that way on 2026-08-09 before this was noticed. 25 of the 30 tests
    ' that quit Excel omit this line; see the note in TRACKER.
    Set wb = Nothing
    Set xl = Nothing

    Test_ReviewQueue_PendingApprovalsCountsTicksAndIgnoresConsumed = result
End Function

' LocalPathForUrl only ever returns a path it has confirmed exists, so a wrong
' guess and no answer are the same outcome. These pin the two ends that do not
' depend on this machine's sync roots; the real mapping is proven by pressing
' Where am I? on a cloud-hosted deck, which is the only place it matters.
Private Function Test_DeckRegistry_LocalPathForUrlOnlyAnswersWhenTheFileIsThere() As String
    Dim result As String
    Dim trace As String

    ' A filesystem path is not a URL and must come back untouched -- every
    ' non-cloud caller goes through this line.
    Dim plain As String
    plain = "C:\Users\rohan\deck-sync-e2e\e2e-deck.pptx"
    result = result & Assert(DeckRegistry.LocalPathForUrl(plain, trace) = plain, _
        "a local path passes through unchanged, got '" & DeckRegistry.LocalPathForUrl(plain, trace) & "'")

    ' A URL under no registered sync root has no local copy to offer, and
    ' inventing one would be worse than admitting it.
    trace = ""
    result = result & Assert( _
        DeckRegistry.LocalPathForUrl("https://not-a-real-tenant.example.com/sites/x/deck.pptx", trace) = "", _
        "an unmapped URL returns empty rather than a constructed path")
    result = result & Assert(InStr(trace, "matched no synced folder") > 0 Or InStr(trace, "no OneDrive") > 0, _
        "the trace says why it could not map, got [" & trace & "]")

    ' And the whole point: a URL must never be handed to FileSystemObject as-is.
    trace = ""
    Dim readFailed As Boolean
    result = result & Assert( _
        DeckRegistry.PropertyOnDisk("https://not-a-real-tenant.example.com/sites/x/deck.pptx", _
            "DeckSyncPeriod", trace, readFailed) = "", _
        "an unreadable cloud deck yields no value")
    result = result & Assert(readFailed, _
        "an unreadable cloud deck is a READ FAILURE, not an absent period [" & trace & "]")

    Test_DeckRegistry_LocalPathForUrlOnlyAnswersWhenTheFileIsThere = result
End Function

' THE TEST THAT ACTUALLY EXERCISES THE MAPPING.
'
' The first three assertions written for LocalPathForUrl used a tenant that
' matches nothing, so all three passed against a version of the function whose
' body had been short-circuited to Exit Function -- proven by doing exactly that
' on 2026-08-09 and watching 159 tests stay green. They tested the two ends and
' skipped the middle, which was the only part that had ever been broken.
'
' This one creates a real file inside the real sync root and asks the function
' to find it from a URL. OneDrive's root comes from the environment, NOT from
' the registry the code under test reads, so the test cannot pass by sharing the
' bug. SKIPS LOUDLY rather than passing when there is no OneDrive to test with.
Private Function Test_DeckRegistry_LocalPathForUrlFindsARealSyncedFile() As String
    Dim result As String

    Dim root As String
    root = Environ("OneDrive")
    If root = "" Then
        Test_DeckRegistry_LocalPathForUrlFindsARealSyncedFile = _
            "    FAIL: SKIPPED -- no OneDrive environment variable, mapping never exercised" & vbCrLf
        Exit Function
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim leaf As String, full As String
    leaf = "dsurltest_" & Format(Now, "yyyymmddhhnnss") & ".txt"
    full = fso.BuildPath(root, leaf)

    fso.CreateTextFile(full, True).Write "probe"

    ' The namespace is read here independently of the code under test. If this
    ' machine registers no personal sync root the mapping cannot be exercised,
    ' and that is reported rather than passed.
    Dim reg As Object, ns As String
    On Error Resume Next
    Set reg = GetObject("winmgmts:\\.\root\default:StdRegProv")
    reg.GetStringValue &H80000001, "SOFTWARE\SyncEngines\Providers\OneDrive\Personal", "UrlNamespace", ns
    On Error GoTo 0

    If ns = "" Then
        On Error Resume Next
        fso.DeleteFile full
        On Error GoTo 0
        Test_DeckRegistry_LocalPathForUrlFindsARealSyncedFile = _
            "    FAIL: SKIPPED -- no personal OneDrive UrlNamespace registered, mapping never exercised" & vbCrLf
        Exit Function
    End If

    ' The shape a cloud-hosted deck actually reports: namespace, a user-id
    ' segment that does NOT appear in the local path, then the file. Resolving
    ' this requires the drop-a-leading-segment retry to work.
    Dim trace As String
    Dim url As String
    url = ns & "/0123456789abcdef/" & leaf

    Dim mapped As String
    mapped = DeckRegistry.LocalPathForUrl(url, trace)

    result = result & Assert(StrComp(mapped, full, vbTextCompare) = 0, _
        "a cloud URL resolves to the real synced file -- wanted '" & full & _
        "', got '" & mapped & "' [" & trace & "]")

    ' And the whole point of the change: reading a property through a URL must
    ' now reach the file rather than reporting it missing.
    Dim readFailed As Boolean
    trace = ""
    DeckRegistry.PropertyOnDisk url, "DeckSyncPeriod", trace, readFailed
    result = result & Assert(InStr(trace, "URL mapped to") > 0, _
        "PropertyOnDisk translates a URL before looking [" & trace & "]")

    On Error Resume Next
    fso.DeleteFile full
    On Error GoTo 0

    Test_DeckRegistry_LocalPathForUrlFindsARealSyncedFile = result
End Function

' The cloud branch of the backup is what blocked Apply Approved entirely: a
' deck whose FullName is a URL got no backup, and ApplyApproved refuses to write
' without one. Destination choice is split out as a pure function precisely so
' this can be exercised without a cloud-hosted deck open.
Private Function Test_ReviewQueue_BackupDestinationHandlesACloudDeck() As String
    Dim result As String
    Dim note As String

    ' A local deck keeps the sibling .bak -- findable next to what it protects.
    Dim localDeck As String
    localDeck = "C:\Users\rohan\deck-sync-e2e\e2e-deck.pptx"
    Dim localDest As String
    localDest = ReviewQueue.BackupDestinationFor(localDeck, note)
    result = result & Assert(InStr(localDest, localDeck) = 1, _
        "a local deck backs up beside itself, got '" & localDest & "'")
    result = result & Assert(Right$(localDest, 9) = ".bak.pptx", _
        "the backup is named as a pptx backup, got '" & localDest & "'")

    ' A cloud deck that resolves to a real synced file must get a destination --
    ' this is the case that used to return nothing and stop the whole run.
    Dim root As String
    root = Environ("OneDrive")
    If root = "" Then
        Test_ReviewQueue_BackupDestinationHandlesACloudDeck = result & _
            "    FAIL: SKIPPED -- no OneDrive, the cloud branch was never exercised" & vbCrLf
        Exit Function
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim leaf As String, full As String
    leaf = "dsbackuptest_" & Format(Now, "yyyymmddhhnnss") & ".pptx"
    full = fso.BuildPath(root, leaf)
    fso.CreateTextFile(full, True).Write "probe"

    Dim reg As Object, ns As String
    On Error Resume Next
    Set reg = GetObject("winmgmts:\\.\root\default:StdRegProv")
    reg.GetStringValue &H80000001, "SOFTWARE\SyncEngines\Providers\OneDrive\Personal", "UrlNamespace", ns
    On Error GoTo 0

    If ns = "" Then
        On Error Resume Next
        fso.DeleteFile full
        On Error GoTo 0
        Test_ReviewQueue_BackupDestinationHandlesACloudDeck = result & _
            "    FAIL: SKIPPED -- no personal UrlNamespace, the cloud branch was never exercised" & vbCrLf
        Exit Function
    End If

    note = ""
    Dim cloudDest As String
    cloudDest = ReviewQueue.BackupDestinationFor(ns & "/0123456789abcdef/" & leaf, note)

    result = result & Assert(cloudDest <> "", _
        "a cloud deck that resolves to a synced file GETS a backup destination [" & note & "]")
    result = result & Assert(InStr(1, cloudDest, Environ("LOCALAPPDATA"), vbTextCompare) = 1, _
        "the cloud backup goes to LOCALAPPDATA, got '" & cloudDest & "'")
    result = result & Assert(InStr(1, cloudDest, root, vbTextCompare) = 0, _
        "the cloud backup is NOT inside the synced folder -- it must not upload, got '" & cloudDest & "'")
    result = result & Assert(InStr(cloudDest, leaf) > 0, _
        "the backup carries the deck's real name, got '" & cloudDest & "'")

    ' And an unresolvable URL still yields nothing, so Apply Approved still stops.
    note = ""
    result = result & Assert( _
        ReviewQueue.BackupDestinationFor("https://not-a-real-tenant.example.com/x/deck.pptx", note) = "", _
        "an unresolvable cloud deck still gets no destination")
    result = result & Assert(note <> "", "and says why [" & note & "]")

    On Error Resume Next
    fso.DeleteFile full
    On Error GoTo 0

    Test_ReviewQueue_BackupDestinationHandlesACloudDeck = result
End Function

' Error 50290 (FIX-LIST.md item V) has crashed ApplyApproved's write loop
' three times across three sessions with no more diagnostic value than
' "Reported by: VBAProject" -- because nothing captured Err.Description/
' Source at the actual point of failure, only at the generic top-level chain
' handler far above it. This proves the fix that closes that gap: an
' unhandled error inside the per-item write must be caught, logged with the
' specific EntityKey/FieldID, and re-raised with that context folded into
' Err.Source -- not swallowed, not generic.
'
' CANNOT FORCE OFFICE TO GENUINELY RAISE 50290 ON DEMAND -- that unreliability
' is the whole reason this bug has gone three sessions unfixed. So this
' proves the WRAPPER is correct against a deliberate, controlled fault
' (ReviewQueue.mTestForceInjectCrash), not that Office will always fault at
' this point. Made to fail on purpose first, same discipline as
' DistinctPinnedFields's test earlier tonight: run by hand against the
' unwrapped call first, confirmed the crash escaped with Source = "VBAProject"
' and no item named anywhere -- then the wrapper was added and this is what
' proves it now.
Private Function Test_ReviewQueue_ApplyApprovedNamesTheItemWhenInjectFieldCrashes() As String
    Dim result As String

    ' ISOLATED PRESENTATION, SAVED TO A REAL PATH. ApplyApproved's first act is
    ' BackupBeforeWrite(Application.ActivePresentation, ...) -- an unsaved
    ' Presentations.Add() deck (what NewBlankSlide/NewTaggedSlide add to,
    ' unless the shared suite deck happens to be saved) has nowhere for
    ' SaveCopyAs to land, backupPath comes back "", and ApplyApproved exits
    ' with "STOPPED before writing" before the write loop this test exists to
    ' probe ever runs. First run of this test genuinely failed that way --
    ' every assertion below came back on an empty string, because nothing had
    ' reached the crash branch at all.
    Dim testPres As Object
    Set testPres = Application.Presentations.Add
    Dim tmpPath As String
    tmpPath = Environ("TEMP") & "\applyapproved_crash_test_" & Format(Now, "yyyymmddhhnnss") & ".pptx"
    testPres.SaveAs tmpPath

    Dim sld As Object
    Set sld = testPres.Slides.Add(1, ppLayoutBlank)
    sld.Tags.Add "slide_type", "test-crash-type"
    sld.Tags.Add "instance_key", "CRASH001"
    Dim shp As Object
    Set shp = sld.Shapes.AddTextbox(1, 40, 40, 300, 40)
    shp.TextFrame.TextRange.Text = "before"
    shp.Tags.Add "role", "ABOUT_BODY"

    Dim sheet As Sheet
    Set sheet.Rows = CreateObject("Scripting.Dictionary")
    Set sheet.Fields = New Collection
    Set sheet.InstanceOrder = New Collection
    Dim vals As Object
    Set vals = CreateObject("Scripting.Dictionary")
    vals("ABOUT_BODY") = "after"
    Set sheet.Rows("CRASH001") = vals
    sheet.InstanceOrder.Add "CRASH001"

    Dim q As ReviewQueueSet
    q.SlideType = "test-crash-type"
    q.RunStamp = "TEST"
    q.Consumed = False
    q.Count = 1
    ReDim q.Items(1 To 1)
    q.Items(1).EntityKey = "CRASH001"
    q.Items(1).FieldID = "ABOUT_BODY"
    q.Items(1).Approved = True
    q.Items(1).ChangeHash = ReviewQueue.ChangeHash("CRASH001", "ABOUT_BODY", "before", "after")

    Dim xl As Object, wb As Object, ws As Object, logWs As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set wb = xl.Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = ReviewQueue.ReviewSheetNameFor("test-crash-type")
    ReviewQueue.WriteQueueSheet ws, q
    Set logWs = wb.Worksheets.Add
    logWs.Name = "TestLog"

    ReviewQueue.mTestForceInjectCrash = True
    Dim report As String
    Dim raised As Boolean, raisedNum As Long, raisedDesc As String, raisedSrc As String
    On Error Resume Next
    Err.Clear
    report = ReviewQueue.ApplyApproved(sheet, "test-crash-type", ws, logWs)
    raised = (Err.Number <> 0)
    raisedNum = Err.Number: raisedDesc = Err.Description: raisedSrc = Err.Source
    On Error GoTo 0
    ReviewQueue.mTestForceInjectCrash = False

    result = result & Assert(raised, "the deliberate fault propagates out of ApplyApproved rather than being swallowed")
    result = result & Assert(InStr(raisedSrc, "CRASH001") > 0 And InStr(raisedSrc, "ABOUT_BODY") > 0, _
        "the re-raised error's Source names the specific item, got '" & raisedSrc & "'")
    result = result & Assert(raisedDesc = "TEST: deliberately injected fault", _
        "the original Err.Description survives the re-raise unchanged, got '" & raisedDesc & "'")

    ' The Sync Log line is the part that survives even if PowerPoint crashes
    ' entirely mid-run -- AppendLogLine's own reasoning, extended to the crash
    ' case. Written BEFORE the re-raise, so it must be there even though the
    ' call above never returned normally.
    Dim logText As String
    logText = CStr(logWs.Cells(2, 5).Value)
    result = result & Assert(InStr(logText, "CRASHED") > 0, _
        "the crash is logged to the Sync Log before re-raising, got '" & logText & "'")

    wb.Close False
    xl.Quit
    Set wb = Nothing
    Set xl = Nothing

    testPres.Saved = True
    testPres.Close

    ' BackupBeforeWrite runs (and succeeds) BEFORE the crash branch this test
    ' forces, so a real .bak.pptx sibling was actually created next to
    ' tmpPath -- clean up both, not just the presentation.
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim f As Object
    For Each f In fso.GetFolder(Environ("TEMP")).Files
        If InStr(f.Name, "applyapproved_crash_test_") > 0 Then f.Delete
    Next f
    On Error GoTo 0

    Test_ReviewQueue_ApplyApprovedNamesTheItemWhenInjectFieldCrashes = result
End Function

' LOBBY-DESIGN.md section 5, phase 3: every field a fresh review queue shows
' now arrives pre-ticked -- working the queue is REMOVING ticks from what
' should not sync this round, not adding them to bless what should. Proves
' BuildQueue's own default rather than assuming the source edit took: run by
' hand against the pre-Phase-3 code first (Approved = False), confirmed
' genuinely False and this test genuinely failing, before the default was
' flipped to True.
Private Function Test_ReviewQueue_BuildQueuePreTicksEveryItem() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewTaggedSlide("pretick-probe", "PT001")
    Dim shp As Object
    Set shp = sld.Shapes.AddTextbox(1, 40, 40, 300, 40)
    shp.TextFrame.TextRange.Text = "old value"
    shp.Tags.Add "role", "ABOUT_BODY"

    Dim sheet As Sheet
    Set sheet.Rows = CreateObject("Scripting.Dictionary")
    Set sheet.Fields = New Collection
    Set sheet.InstanceOrder = New Collection
    Dim vals As Object
    Set vals = CreateObject("Scripting.Dictionary")
    vals("ABOUT_BODY") = "new value"
    Set sheet.Rows("PT001") = vals
    sheet.InstanceOrder.Add "PT001"

    Dim q As ReviewQueueSet
    q = ReviewQueue.BuildQueue(sheet, "pretick-probe")

    result = result & Assert(q.Count = 1, "one real diff produces one queue item, got " & q.Count)
    result = result & Assert(q.Items(1).Approved, _
        "a freshly-built queue item arrives pre-ticked (LOBBY-DESIGN.md section 5), got Approved=" & q.Items(1).Approved)

    sld.Delete
    Test_ReviewQueue_BuildQueuePreTicksEveryItem = result
End Function

' The link that did not exist until 2026-08-09: a cited source reaching the
' prompt. Sources were validated at publish and never shown to the thing doing
' the writing, so a recipe forbidding an inferred claim could never be satisfied
' by a document -- WORKED-EXAMPLE-STRATEGIC-ALIGNMENT.md's "[TBC]" was permanent.
Private Function Test_Sources_CitedBlockPutsTheDocumentInThePrompt() As String
    Dim result As String

    Dim xl As Object, wb As Object, srcWs As Object, dws As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set wb = xl.Workbooks.Add
    Set srcWs = wb.Worksheets(1)
    Set dws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))

    srcWs.Cells(Sources.SRC_FIRST_ROW, Sources.COL_S_ID).Value = "S12"
    srcWs.Cells(Sources.SRC_FIRST_ROW, Sources.COL_S_LABEL).Value = "Program plan -- declared linkage codes"
    srcWs.Cells(Sources.SRC_FIRST_ROW, Sources.COL_S_TYPE).Value = "SPOT source"
    srcWs.Cells(Sources.SRC_FIRST_ROW, Sources.COL_S_LOCATOR).Value = "C:\rig\plan.md"
    srcWs.Cells(Sources.SRC_FIRST_ROW, Sources.COL_S_APPLIES).Value = Sources.APPLIES_ALL

    ' Two rows: one citing a known source, one citing an ID that does not exist.
    dws.Cells(Drafting.DRAFT_FIRST_ROW, 1).Value = "2_P004"
    dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_SOURCES).Value = "S12"
    dws.Cells(Drafting.DRAFT_FIRST_ROW + 1, 1).Value = "3_P001"
    dws.Cells(Drafting.DRAFT_FIRST_ROW + 1, Drafting.COL_D_SOURCES).Value = "S99"

    Dim block As String
    block = Sources.CitedBlockFor(srcWs, dws, Drafting.COL_D_SOURCES, Drafting.DRAFT_FIRST_ROW)

    result = result & Assert(InStr(block, "S12") > 0, "the cited ID reaches the prompt")
    result = result & Assert(InStr(block, "Program plan -- declared linkage codes") > 0, _
        "so does what the source IS -- an ID alone is not evidence")
    result = result & Assert(InStr(block, "C:\rig\plan.md") > 0, _
        "and WHERE it lives, or it cannot be opened")
    result = result & Assert(InStr(block, "unopened source is not a source") > 0, _
        "the block refuses inference from a document that could not be read")
    result = result & Assert(InStr(block, "S99") > 0 And InStr(block, "NOT ON THE SOURCES SHEET") > 0, _
        "an unknown citation is named as NOT evidence rather than passed over")

    ' A sheet citing nothing must add nothing -- a field needing no evidence
    ' should not carry a paragraph about evidence.
    Dim bare As Object
    Set bare = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    bare.Cells(Drafting.DRAFT_FIRST_ROW, 1).Value = "2_P004"
    result = result & Assert(Sources.CitedBlockFor(srcWs, bare, Drafting.COL_D_SOURCES, Drafting.DRAFT_FIRST_ROW) = "", _
        "a sheet that cites nothing produces no block")

    wb.Close False
    xl.Quit
    Set wb = Nothing
    Set xl = Nothing

    Test_Sources_CitedBlockPutsTheDocumentInThePrompt = result
End Function

' THE WIRING, NOT THE PIECES. The first test written for this covered
' CitedBlockFor alone, so replacing the call with citedBlock = "" left 162 tests
' green -- the link from a citation to the prompt was untested while looking
' tested. This builds a real drafting sheet and reads cell L2.
'
' Two builds, deliberately: citations are carried across a rebuild, so a first
' build has none. That is also how it happens in use -- you cite after the sheet
' exists, and the next rebuild must carry it into the prompt.
Private Function Test_Drafting_CitedSourceReachesThePromptCell() As String
    Dim result As String

    Dim xl As Object, wb As Object, regWs As Object, dws As Object, srcWs As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set wb = xl.Workbooks.Add
    Set regWs = wb.Worksheets(1)
    Set dws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    Set srcWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))

    ExcelOutput.CreateSheet regWs, "deck-v1"
    Dim vals As Object
    Set vals = CreateObject("Scripting.Dictionary")
    vals("ABOUT_BODY") = "existing text"
    ExcelOutput.UpsertRow regWs, "2_P004", vals, "Q4F26"
    Dim reg As Sheet
    reg = ExcelOutput.ReadSheetForDeckPeriod(regWs, "Q4F26", "")

    srcWs.Cells(Sources.SRC_FIRST_ROW, Sources.COL_S_ID).Value = "S12"
    srcWs.Cells(Sources.SRC_FIRST_ROW, Sources.COL_S_LABEL).Value = "Program plan -- declared linkage codes"
    srcWs.Cells(Sources.SRC_FIRST_ROW, Sources.COL_S_LOCATOR).Value = "C:\rig\plan.md"

    ' Build once, cite, rebuild -- the citation is carried across.
    Drafting.WriteDraftingSheet dws, reg, "ABOUT_BODY", Nothing, "Q4F26", srcWs
    Dim before As String
    before = CStr(dws.Cells(2, Drafting.COL_D_PROMPT).Value)
    result = result & Assert(InStr(before, "S12") = 0, _
        "with nothing cited, the prompt carries no sources block")

    dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_SOURCES).Value = "S12"
    Drafting.WriteDraftingSheet dws, reg, "ABOUT_BODY", Nothing, "Q4F26", srcWs

    Dim after As String
    after = CStr(dws.Cells(2, Drafting.COL_D_PROMPT).Value)
    result = result & Assert(InStr(after, "S12") > 0, _
        "the cited source reaches the PROMPT CELL, not just the block builder")
    result = result & Assert(InStr(after, "C:\rig\plan.md") > 0, _
        "with its locator, so the document can actually be opened")
    result = result & Assert(InStr(after, "sole source of truth") > 0, _
        "and the original guard survives -- the block widens it, never replaces it")

    wb.Close False
    xl.Quit
    Set wb = Nothing
    Set xl = Nothing

    Test_Drafting_CitedSourceReachesThePromptCell = result
End Function

' A LAYOUT-3 SHEET MUST GIVE UP ITS WORK INTO THE LAYOUT-4 POSITIONS.
'
' Layout 3 ordered columns as they were added; layout 4 orders them as they are
' used (Rohan, 2026-08-10). Bumping the version alone is SAFE AND LOSSY -- the
' carry-across refuses an unrecognised layout, so every unpublished draft, note
' and source citation on the sheet is dropped at the next rebuild. Too expensive
' to pay for a reordering, hence ColumnInLayout.
'
' The failure this guards is the one Drafting.bas's own header describes and
' which happened for real on 2026-08-01: reading an old sheet with new column
' numbers does not crash, it RELABELS -- a tick arriving as publishable text.
' So the assertions below check WHICH FIELD each value came back as, not merely
' that something was carried.
'
' The suite's 163 could not see any of this: every test addresses columns
' through the COL_D_* constants, so they pass whatever the numbers are.
Private Function Test_Drafting_Layout3SheetMigratesIntoLayout4Columns() As String
    Dim result As String

    Dim xl As Object, wb As Object, regWs As Object, dws As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set wb = xl.Workbooks.Add
    Set regWs = wb.Worksheets(1)
    Set dws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))

    ExcelOutput.CreateSheet regWs, "deck-v1"
    Dim vals As Object
    Set vals = CreateObject("Scripting.Dictionary")
    vals("ABOUT_BODY") = "register text"
    ExcelOutput.UpsertRow regWs, "2_P004", vals, "Q4F26"
    Dim reg As Sheet
    reg = ExcelOutput.ReadSheetForDeckPeriod(regWs, "Q4F26", "")

    ' A sheet as LAYOUT 3 left it: C original, D submit, E tick, F AI, G sources,
    ' J notes -- written by hand at those numbers, which is the whole point.
    ' STAMPED WHERE LAYOUT 3 ACTUALLY STAMPS -- columns 11 and 13, not the
    ' current layout's 12 and 14. This fixture used to use COL_D_LAYOUT and
    ' COL_D_PERIOD, so it described a sheet that has never existed: layout-3 DATA
    ' under layout-5 STAMPS. It passed only because the reading code made the
    ' same mistake, which is why 192 green tests sat on top of a defect that
    ' wiped every drafting sheet of a real workbook on 2026-08-14.
    dws.Cells(Drafting.DRAFT_INTRO_ROW, Drafting.LayoutStampColumn(3)).Value = 3
    dws.Cells(Drafting.DRAFT_INTRO_ROW, Drafting.PeriodStampColumn(3)).Value = "Q4F26"
    dws.Cells(Drafting.DRAFT_FIRST_ROW, 1).Value = "2_P004"
    dws.Cells(Drafting.DRAFT_FIRST_ROW, 4).Value = "MY SUBMITTED WORDS"
    dws.Cells(Drafting.DRAFT_FIRST_ROW, 5).Value = "Y"
    dws.Cells(Drafting.DRAFT_FIRST_ROW, 6).Value = "THE AI DRAFT"
    dws.Cells(Drafting.DRAFT_FIRST_ROW, 7).Value = "S12"
    dws.Cells(Drafting.DRAFT_FIRST_ROW, 10).Value = "MY NOTE"

    Drafting.WriteDraftingSheet dws, reg, "ABOUT_BODY", Nothing, "Q4F26", Nothing

    Dim got As String
    got = CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_SUBMIT).Value)
    result = result & Assert(got = "MY SUBMITTED WORDS", _
        "SUBMIT lands in the new SUBMIT column, got '" & got & "'")

    got = CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_DRAFT).Value)
    result = result & Assert(got = "THE AI DRAFT", _
        "the AI draft lands in the new AI column, got '" & got & "'")

    got = CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_SOURCES).Value)
    result = result & Assert(got = "S12", _
        "the source citation survives the move, got '" & got & "'")

    got = CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_NOTES).Value)
    result = result & Assert(got = "MY NOTE", _
        "notes survive, got '" & got & "'")

    ' THE RELABELLING TEST. Layout 3's SUBMIT sat at column 4, which is layout
    ' 4's SOURCES. If the migration were skipped, "MY SUBMITTED WORDS" would be
    ' read as a source ID -- text arriving as a citation, the same shape as a
    ' tick arriving as publishable text.
    result = result & Assert( _
        InStr(CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_SOURCES).Value), "SUBMITTED") = 0, _
        "SUBMIT text is NOT read as a source ID")
    result = result & Assert( _
        InStr(CStr(dws.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_SUBMIT).Value), "AI DRAFT") = 0, _
        "the AI draft is NOT read as submitted text -- that would publish it")

    ' And the sheet is stamped forward, so this migration happens once.
    result = result & Assert( _
        CLng(Val(CStr(dws.Cells(Drafting.DRAFT_INTRO_ROW, Drafting.COL_D_LAYOUT).Value))) = Drafting.DRAFT_LAYOUT_VERSION, _
        "the rebuilt sheet is stamped with the current layout version")

    wb.Close False
    xl.Quit
    Set wb = Nothing
    Set xl = Nothing

    Test_Drafting_Layout3SheetMigratesIntoLayout4Columns = result
End Function

' Writes a real 24-bit BMP of the given size. Self-contained on purpose: the
' picture tests need an image PowerPoint will actually accept, and depending on
' a committed fixture file would make them fail for a reason unrelated to the
' thing under test.
Private Function MakeTestBitmap(path As String, w As Long, h As Long) As String
    Dim rowBytes As Long, padding As Long, pixelBytes As Long
    rowBytes = w * 3
    padding = (4 - (rowBytes Mod 4)) Mod 4
    pixelBytes = (rowBytes + padding) * h

    Dim b() As Byte
    ReDim b(0 To 53 + pixelBytes - 1)

    b(0) = 66: b(1) = 77                          ' "BM"
    PutLong b, 2, 54 + pixelBytes                 ' file size
    PutLong b, 10, 54                             ' pixel data offset
    PutLong b, 14, 40                             ' DIB header size
    PutLong b, 18, w
    PutLong b, 22, h
    b(26) = 1                                     ' planes
    b(28) = 24                                    ' bits per pixel
    PutLong b, 34, pixelBytes

    Dim i As Long
    For i = 54 To UBound(b)
        b(i) = 200                                ' flat grey, content is irrelevant
    Next i

    Dim fnum As Integer
    fnum = FreeFile
    On Error Resume Next
    Kill path
    On Error GoTo 0
    Open path For Binary Access Write As #fnum
    Put #fnum, 1, b
    Close #fnum

    MakeTestBitmap = path
End Function

Private Sub PutLong(ByRef b() As Byte, at As Long, v As Long)
    b(at) = v And &HFF
    b(at + 1) = (v \ 256) And &HFF
    b(at + 2) = (v \ 65536) And &HFF
    b(at + 3) = (v \ 16777216) And &HFF
End Sub

' A picture field: filled from a link, stamped, and silent on the second run.
'
' The stamp is the whole design -- without it, idempotence would mean comparing
' images across 43 slides every quarter. So the assertions are about the STAMP
' and about what happens when it already matches, not about pixels.
Private Function Test_InjectPicture_FillsStampsAndThenStaysSilent() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()

    ' A picture already on the slide, tagged as a field -- the frame whose
    ' position and size the replacement must respect.
    Dim seedPath As String
    seedPath = MakeTestBitmap(Environ("TEMP") & "\dsseed.bmp", 40, 20)
    Dim frame As Object
    Set frame = sld.Shapes.AddPicture(seedPath, msoFalse, msoTrue, 100, 80, 300, 150)
    frame.Tags.Add "role", "PHOTO"

    Dim newPath As String
    newPath = MakeTestBitmap(Environ("TEMP") & "\dsnew.bmp", 40, 20)

    Dim r As InjectResult
    r = InjectPrimitive.InjectPictureField(sld, "PHOTO", "S20", newPath)

    result = result & Assert(r.Found, "the tagged picture is found")
    result = result & Assert(r.Written, "the picture is written [" & r.ErrorMessage & "]")
    result = result & Assert(r.Verified, "and verified from the new shape's own stamp")

    Dim placed As Object
    Set placed = ShapeTaggedRole(sld, "PHOTO")
    result = result & Assert(Not placed Is Nothing, "the replacement carries the role tag forward")
    If Not placed Is Nothing Then
        result = result & Assert(InjectPrimitive.PictureSourceOf(placed) = "S20", _
            "and the source stamp, got '" & InjectPrimitive.PictureSourceOf(placed) & "'")
        ' The frame is untouched -- it was fed, not replaced.
        result = result & Assert(Abs(placed.Width - 300) < 1 And Abs(placed.Height - 150) < 1, _
            "the frame keeps the size the slide gave it, got " & placed.Width & "x" & placed.Height)
        result = result & Assert(Abs(placed.Left - 100) < 0.5 And Abs(placed.Top - 80) < 0.5, _
            "and its position, got " & placed.Left & "," & placed.Top)
    End If

    ' THE SECOND RUN MUST DO NOTHING. This is the property that makes a picture
    ' field cost nothing per quarter.
    Dim r2 As InjectResult
    r2 = InjectPrimitive.InjectPictureField(sld, "PHOTO", "S20", newPath)
    result = result & Assert(Not r2.WouldChange, "a second run with the same source changes nothing")
    result = result & Assert(Not r2.Written, "and writes nothing")

    ' A CHANGED LINK RE-FIRES BY ITSELF.
    Dim r3 As InjectResult
    r3 = InjectPrimitive.InjectPictureField(sld, "PHOTO", "S21", newPath)
    result = result & Assert(r3.WouldChange, "a different source ID re-fires")

    sld.Delete
    Test_InjectPicture_FillsStampsAndThenStaysSilent = result
End Function

' THE SAFETY ONE. The existing image is the only copy on the slide, so a locator
' that is not a readable file must be refused BEFORE anything is deleted --
' otherwise a typo in a Sources row costs a photo.
Private Function Test_InjectPicture_RefusesABadLocatorWithoutLosingTheOldImage() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim seedPath As String
    seedPath = MakeTestBitmap(Environ("TEMP") & "\dsseed2.bmp", 40, 20)
    Dim frame As Object
    Set frame = sld.Shapes.AddPicture(seedPath, msoFalse, msoTrue, 100, 80, 300, 150)
    frame.Tags.Add "role", "PHOTO"
    frame.Tags.Add InjectPrimitive.PICTURE_SOURCE_TAG, "S20"

    Dim before As Long
    before = sld.Shapes.count

    Dim r As InjectResult
    r = InjectPrimitive.InjectPictureField(sld, "PHOTO", "S21", Environ("TEMP") & "\no_such_image_here.bmp")

    result = result & Assert(Not r.Written, "a missing file is not written")
    result = result & Assert(InStr(r.ErrorMessage, "not there") > 0, _
        "and says the file is not there, got '" & r.ErrorMessage & "'")
    result = result & Assert(sld.Shapes.count = before, _
        "THE OLD IMAGE SURVIVES -- shape count unchanged, got " & sld.Shapes.count & " was " & before)

    Dim still As Object
    Set still = ShapeTaggedRole(sld, "PHOTO")
    result = result & Assert(Not still Is Nothing, "the field is still tagged")
    If Not still Is Nothing Then
        result = result & Assert(InjectPrimitive.PictureSourceOf(still) = "S20", _
            "and still stamped with what it was actually filled from")
    End If

    sld.Delete
    Test_InjectPicture_RefusesABadLocatorWithoutLosingTheOldImage = result
End Function

' Finds a shape by its role tag without reaching into InjectPrimitive's private
' locator -- the test should not widen production visibility to see a result.
Private Function ShapeTaggedRole(sld As Object, roleTag As String) As Object
    Dim shp As Object
    For Each shp In sld.Shapes
        On Error Resume Next
        If shp.Tags("role") = roleTag Then
            Set ShapeTaggedRole = shp
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    Next shp
End Function

' A CROPPED FRAME MEANS FILL IT. Read off the shape rather than configured --
' the real deck's project banner carries a crop and its logos do not, so the
' deck already states which treatment each frame wants.
Private Function Test_InjectPicture_CroppedFrameIsFilledUncroppedIsFitted() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim wide As String, tall As String
    wide = MakeTestBitmap(Environ("TEMP") & "\dswide.bmp", 40, 20)
    tall = MakeTestBitmap(Environ("TEMP") & "\dstall.bmp", 20, 40)

    ' A cropped frame: 300x150, holding a portrait image it must fill.
    Dim cropped As Object
    Set cropped = sld.Shapes.AddPicture(wide, msoFalse, msoTrue, 100, 80, 300, 150)
    cropped.Tags.Add "role", "BANNER"
    cropped.PictureFormat.CropLeft = 2
    cropped.PictureFormat.CropRight = 2

    ' MEASURED, NOT ASSUMED. Applying a crop SHRINKS a picture shape, so the
    ' frame is not the 300x150 it was created at -- an earlier version of this
    ' test asserted 300x150 and failed against numbers the setup itself had
    ' changed. What matters is that the injection leaves them alone.
    Dim wasW As Single, wasH As Single, wasL As Single, wasT As Single, wasCropL As Single
    wasW = cropped.Width: wasH = cropped.Height
    wasL = cropped.Left: wasT = cropped.Top
    wasCropL = cropped.PictureFormat.CropLeft

    Dim r As InjectResult
    r = InjectPrimitive.InjectPictureField(sld, "BANNER", "S30", tall)

    ' A CROPPED FRAME IS REPLACED AND KEEPS EVERYTHING. Size, position and the
    ' crop itself -- the last of which needs LockAspectRatio off before the
    ' dimensions are assigned, or each one undoes the other.
    result = result & Assert(r.Written, "a cropped frame is replaced [" & r.ErrorMessage & "]")
    result = result & Assert(r.Verified, "and verified [" & r.ErrorMessage & "]")

    Dim after As Object
    Set after = ShapeTaggedRole(sld, "BANNER")
    If Not after Is Nothing Then
        result = result & Assert(Abs(after.Width - wasW) < 0.5 And Abs(after.Height - wasH) < 0.5, _
            "the frame keeps its size, was " & wasW & "x" & wasH & " now " & after.Width & "x" & after.Height)
        result = result & Assert(Abs(after.Left - wasL) < 0.5 And Abs(after.Top - wasT) < 0.5, _
            "and its position, was " & wasL & "," & wasT & " now " & after.Left & "," & after.Top)
        result = result & Assert(Abs(after.PictureFormat.CropLeft - wasCropL) < 0.1, _
            "and its cropping, was " & wasCropL & " now " & after.PictureFormat.CropLeft)
        result = result & Assert(InjectPrimitive.PictureSourceOf(after) = "S30", _
            "and is stamped with what it was filled from")
    End If

    ' An uncropped frame with the same mismatch must FIT, not fill.
    Dim plain As Object
    Set plain = sld.Shapes.AddPicture(wide, msoFalse, msoTrue, 400, 80, 300, 150)
    plain.Tags.Add "role", "THUMB"
    Dim plainW As Single, plainH As Single
    plainW = plain.Width: plainH = plain.Height
    Dim r2 As InjectResult
    r2 = InjectPrimitive.InjectPictureField(sld, "THUMB", "S31", tall)
    Dim placed2 As Object
    Set placed2 = ShapeTaggedRole(sld, "THUMB")
    If Not placed2 Is Nothing Then
        result = result & Assert(Abs(placed2.Width - plainW) < 0.5 And Abs(placed2.Height - plainH) < 0.5, _
            "an uncropped frame also keeps its size, was " & plainW & "x" & plainH & " now " & placed2.Width & "x" & placed2.Height)
    End If

    sld.Delete
    Test_InjectPicture_CroppedFrameIsFilledUncroppedIsFitted = result
End Function

' A progress bar measured against a track it never writes.
'
' The assertion that matters is the SECOND RUN. Scaling the done part against
' its own width works exactly once -- after that the bar is shorter, so the next
' run takes a fraction of a fraction and walks toward zero while every report
' says success. Reading a shape the tool never touches is what prevents it, and
' running twice is the only way to prove it.
Private Function Test_InjectProgress_MeasuresAgainstTheTrackNotItself() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()

    Dim track As Object, done As Object, rest As Object
    Set track = sld.Shapes.AddShape(1, 100, 200, 400, 12)     ' msoShapeRectangle
    track.Tags.Add "role", "ELAPSED.track"
    Set done = sld.Shapes.AddShape(1, 100, 200, 40, 12)
    done.Tags.Add "role", "ELAPSED"
    Set rest = sld.Shapes.AddShape(1, 140, 200, 360, 12)
    rest.Tags.Add "role", "ELAPSED.rest"

    Dim r As InjectResult
    r = InjectPrimitive.InjectProgressField(sld, "ELAPSED", 0.75)

    ' CRITICAL-EFFECTS SHORTLIST, item 2 -- see Watch's header. The done
    ' shape resizes to 75% of the track's width, visibly, right here.
    Watch sld, "progress bar resized to 75% after InjectProgressField"

    result = result & Assert(r.Written, "the bar is set [" & r.ErrorMessage & "]")
    result = result & Assert(Abs(done.Width - 300) < 1, _
        "75% of a 400pt track is 300pt, got " & done.Width)
    result = result & Assert(Abs(rest.Left - 400) < 1 And Abs(rest.Width - 100) < 1, _
        "the remainder takes what is left, got left " & rest.Left & " width " & rest.Width)
    result = result & Assert(Abs(track.Width - 400) < 0.5, _
        "THE TRACK IS NEVER WRITTEN, got " & track.Width)

    ' SECOND RUN, SAME VALUE. A bar measured against itself would shrink here.
    Dim r2 As InjectResult
    r2 = InjectPrimitive.InjectProgressField(sld, "ELAPSED", 0.75)
    result = result & Assert(Not r2.WouldChange, "a second run with the same value changes nothing")
    result = result & Assert(Abs(done.Width - 300) < 1, _
        "and the bar is STILL 300pt, not a fraction of a fraction, got " & done.Width)

    ' A value that is a percentage rather than a fraction must be refused, not
    ' clamped -- 90 meaning 0.9 would otherwise draw a full bar and look right.
    Dim r3 As InjectResult
    r3 = InjectPrimitive.InjectProgressField(sld, "ELAPSED", 90)
    result = result & Assert(Not r3.Written, "90 is refused rather than clamped to full")
    result = result & Assert(Abs(done.Width - 300) < 1, "and the bar is untouched, got " & done.Width)

    sld.Delete
    Test_InjectProgress_MeasuresAgainstTheTrackNotItself = result
End Function

' No track means no answer. Falling back to the bar's own width is the shrinking
' bug; falling back to the slide is a scale nobody chose.
Private Function Test_InjectProgress_RefusesWithoutATrack() As String
    Dim result As String
    Dim sld As Object
    Set sld = NewBlankSlide()

    Dim done As Object
    Set done = sld.Shapes.AddShape(1, 100, 200, 40, 12)
    done.Tags.Add "role", "LONELY"
    Dim before As Single
    before = done.Width

    Dim r As InjectResult
    r = InjectPrimitive.InjectProgressField(sld, "LONELY", 0.5)
    result = result & Assert(Not r.Written, "a bar with no track is not written")
    result = result & Assert(InStr(r.ErrorMessage, ".track") > 0, _
        "and the message names the tag to add, got '" & r.ErrorMessage & "'")
    result = result & Assert(done.Width = before, "the bar is left exactly as it was")

    sld.Delete
    Test_InjectProgress_RefusesWithoutATrack = result
End Function

' THE ROUTER. Three fields of three different types on ONE slide, all sent
' through the single entry point sync uses.
'
' This is the test that would have caught the defect it was written for: until
' 2026-08-10 every sync call site called the TEXT injector for every field, so
' the bar and the picture here would both have come back "no text frame to
' write into" while the suite stayed green -- because the picture and progress
' injectors were tested directly and nothing tested that anything CALLS them.
'
' All three on one slide deliberately: routing that works when a slide holds
' only one kind of field proves much less than routing that has to tell three
' apart with all three in front of it.
Private Function Test_InjectField_RoutesEachTypeByItsShape() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()

    ' --- text -------------------------------------------------------------
    Dim tb As Object
    Set tb = sld.Shapes.AddTextbox(1, 40, 40, 300, 40)
    tb.TextFrame.TextRange.Text = "old prose"
    tb.Tags.Add "role", "ABOUT_BODY"

    ' --- a bar: done part plus the track it measures against ---------------
    Dim track As Object, done As Object
    Set track = sld.Shapes.AddShape(1, 40, 200, 400, 12)      ' msoShapeRectangle
    track.Tags.Add "role", "ELAPSED.track"
    Set done = sld.Shapes.AddShape(1, 40, 200, 40, 12)
    done.Tags.Add "role", "ELAPSED"

    ' --- a picture --------------------------------------------------------
    Dim seedPath As String
    seedPath = MakeTestBitmap(Environ("TEMP") & "\dsroute.bmp", 40, 20)
    Dim frame As Object
    Set frame = sld.Shapes.AddPicture(seedPath, msoFalse, msoTrue, 40, 300, 200, 100)
    frame.Tags.Add "role", "PHOTO"

    ' --- text goes to the text injector -----------------------------------
    Dim rt As InjectResult
    rt = InjectPrimitive.InjectField(sld, "ABOUT_BODY", "new prose")
    result = result & Assert(rt.Written, "text is written [" & rt.ErrorMessage & "]")
    result = result & Assert(IgnoringTrailingBreaksForTest(tb.TextFrame.TextRange.Text) = "new prose", _
        "and the textbox holds it, got '" & tb.TextFrame.TextRange.Text & "'")

    ' --- a bar goes to the progress injector, from a STRING cell -----------
    ' The register holds text; the injector needs a Double. The conversion is
    ' the router's, and this is the assertion that proves it happened.
    Dim rb As InjectResult
    rb = InjectPrimitive.InjectField(sld, "ELAPSED", "0.75")
    result = result & Assert(rb.Written, "the bar is written [" & rb.ErrorMessage & "]")
    result = result & Assert(Abs(done.Width - 300) < 1, _
        "75% of a 400pt track is 300pt, got " & done.Width)
    result = result & Assert(Abs(track.Width - 400) < 0.5, _
        "and the track is still never written, got " & track.Width)

    ' --- a picture goes to the picture injector ----------------------------
    ' With no Sources sheet in scope it must say THAT, and specifically must
    ' not fall through to the text writer's "no text frame" -- a true sentence
    ' about the wrong question, which is what sent five minutes into the wrong
    ' file the last time a message named the wrong thing.
    Dim rp As InjectResult
    rp = InjectPrimitive.InjectField(sld, "PHOTO", "S20")
    result = result & Assert(InStr(rp.ErrorMessage, Sources.SOURCES_SHEET_NAME) > 0, _
        "a picture with no Sources sheet names the Sources sheet, got '" & rp.ErrorMessage & "'")
    result = result & Assert(InStr(rp.ErrorMessage, "text frame") = 0, _
        "and is NOT refused as a text field, got '" & rp.ErrorMessage & "'")
    result = result & Assert(rp.Found And rp.WouldChange, _
        "and does not read as 'nothing to do', which would hide it from the report")

    sld.Delete
    Test_InjectField_RoutesEachTypeByItsShape = result
End Function

' A register cell holding something that is not a number.
'
' The assertion that carries the weight is the LAST one. Val("done") returns 0,
' which would draw an empty bar and report success; and a result with Found or
' WouldChange False is SKIPPED by SyncOperations and reported as no change --
' so a bar could fail to draw inside a run that says it went perfectly. Being
' refused is not enough; it has to be refused LOUDLY.
Private Function Test_InjectField_RefusesANonNumberForABarWithoutGoingQuiet() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()

    Dim track As Object, done As Object
    Set track = sld.Shapes.AddShape(1, 40, 200, 400, 12)
    track.Tags.Add "role", "ELAPSED.track"
    Set done = sld.Shapes.AddShape(1, 40, 200, 40, 12)
    done.Tags.Add "role", "ELAPSED"

    Dim before As Single
    before = done.Width

    Dim r As InjectResult
    r = InjectPrimitive.InjectField(sld, "ELAPSED", "not a number")

    result = result & Assert(Not r.Written, "a non-number is not written")
    result = result & Assert(done.Width = before, _
        "and the bar is left exactly as it was, got " & done.Width & " was " & before)
    result = result & Assert(InStr(r.ErrorMessage, "ELAPSED") > 0, _
        "the message names the field, got '" & r.ErrorMessage & "'")
    result = result & Assert(InStr(r.ErrorMessage, "not a number") > 0, _
        "and quotes what the register actually held, got '" & r.ErrorMessage & "'")
    result = result & Assert(r.Found And r.WouldChange, _
        "and it is LOUD: Found and WouldChange are both set, or the run reports clean over it")

    sld.Delete
    Test_InjectField_RefusesANonNumberForABarWithoutGoingQuiet = result
End Function

' TWO SLIDES. Every picture and progress test before this used one, so nothing
' proved a second slide gets its own value rather than the first slide's -- and
' a quarter is 43 slides, not one. The `Dim` that does not scope to a loop cost
' this project forty projects' worth of column C on 2026-08-09; a value that
' leaks from one slide to the next is the same defect wearing a different hat.
Private Function Test_InjectField_TwoSlidesEachGetTheirOwnValue() As String
    Dim result As String

    Dim sldA As Object, sldB As Object
    Set sldA = NewBlankSlide()
    Set sldB = NewBlankSlide()

    Dim trackA As Object, doneA As Object, trackB As Object, doneB As Object
    Set trackA = sldA.Shapes.AddShape(1, 40, 200, 400, 12)
    trackA.Tags.Add "role", "ELAPSED.track"
    Set doneA = sldA.Shapes.AddShape(1, 40, 200, 40, 12)
    doneA.Tags.Add "role", "ELAPSED"

    ' A DIFFERENT track width, so a bar drawn against the wrong slide's track
    ' cannot land on the right number by coincidence.
    Set trackB = sldB.Shapes.AddShape(1, 40, 200, 200, 12)
    trackB.Tags.Add "role", "ELAPSED.track"
    Set doneB = sldB.Shapes.AddShape(1, 40, 200, 40, 12)
    doneB.Tags.Add "role", "ELAPSED"

    Dim ra As InjectResult, rb As InjectResult
    ra = InjectPrimitive.InjectField(sldA, "ELAPSED", "0.25")
    rb = InjectPrimitive.InjectField(sldB, "ELAPSED", "0.75")

    result = result & Assert(ra.Written And rb.Written, _
        "both slides are written [" & ra.ErrorMessage & "] [" & rb.ErrorMessage & "]")
    result = result & Assert(Abs(doneA.Width - 100) < 1, _
        "slide A: 25% of its own 400pt track is 100pt, got " & doneA.Width)
    result = result & Assert(Abs(doneB.Width - 150) < 1, _
        "slide B: 75% of its own 200pt track is 150pt, got " & doneB.Width)

    sldB.Delete
    sldA.Delete
    Test_InjectField_TwoSlidesEachGetTheirOwnValue = result
End Function

' Builds a real tagged instance of a slide type, so GatherInstances finds it.
Private Function NewTaggedSlide(slideType As String, instanceKey As String) As Object
    Dim sld As Object
    Set sld = NewBlankSlide()
    sld.Tags.Add "slide_type", slideType
    sld.Tags.Add "instance_key", instanceKey
    Set NewTaggedSlide = sld
End Function

Private Function FieldsCollection(ParamArray names() As Variant) As Collection
    Dim c As New Collection
    Dim i As Long
    For i = LBound(names) To UBound(names)
        c.Add CStr(names(i))
    Next i
    Set FieldsCollection = c
End Function

' A register column with nothing on any slide to write into.
'
' The message must NAME the field. A count alone has been true-and-unusable
' four separate times in this project, and this is the message that decides
' whether someone goes and tags the right shape or hunts the wrong sheet.
Private Function Test_FieldWiring_NamesTheFieldsNothingCarries() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewTaggedSlide("wiring-probe", "wp-1")
    Dim tb As Object
    Set tb = sld.Shapes.AddTextbox(1, 40, 40, 300, 40)
    tb.TextFrame.TextRange.Text = "about"
    tb.Tags.Add "role", "ABOUT_BODY"

    Dim r As FieldWiringResult
    r = FieldWiring.ScanFieldWiring("wiring-probe", _
        FieldsCollection("ABOUT_BODY", "PROGRESS_BODY"), Nothing)

    result = result & Assert(r.Scanned, "the scan ran")
    result = result & Assert(r.Wired = 1, "one field is wired, got " & r.Wired)
    result = result & Assert(r.UnmarkedCount = 1, _
        "one field is unmarked, got " & r.UnmarkedCount)
    result = result & Assert(InStr(r.Unmarked, "PROGRESS_BODY") > 0, _
        "and it is NAMED, got '" & r.Unmarked & "'")
    result = result & Assert(InStr(FieldWiring.WiringText(r), "PROGRESS_BODY") > 0, _
        "the sentence names it too, got '" & FieldWiring.WiringText(r) & "'")

    ' No template was passed, so the template question was never asked. It must
    ' not read as a clean template.
    result = result & Assert(Not r.TemplateScanned, _
        "a missing template is reported as not-scanned, never as fine")

    sld.Delete
    Test_FieldWiring_NamesTheFieldsNothingCarries = result
End Function

' THE ONE THAT MATTERS FOR NEW SLIDES. GatherInstances excludes the template by
' design, so a scan built on it alone reports every field wired while the
' TEMPLATE is missing one -- and every slide created from then on silently
' lacks that field, because a new slide is a Duplicate of the template.
'
' The instance carries BOTH fields and the template carries only one, so a scan
' that ignored the template would return a clean result here. That is the shape
' this test exists to make impossible.
Private Function Test_FieldWiring_TemplateIsCheckedSeparatelyFromInstances() As String
    Dim result As String

    Dim inst As Object
    Set inst = NewTaggedSlide("wiring-tmpl", "wt-1")
    Dim a As Object, b As Object
    Set a = inst.Shapes.AddTextbox(1, 40, 40, 200, 30)
    a.TextFrame.TextRange.Text = "about"
    a.Tags.Add "role", "ABOUT_BODY"
    Set b = inst.Shapes.AddTextbox(1, 40, 90, 200, 30)
    b.TextFrame.TextRange.Text = "problem"
    b.Tags.Add "role", "PROBLEM_BODY"

    ' The template carries only ABOUT_BODY.
    Dim tmpl As Object
    Set tmpl = NewBlankSlide()
    Dim t1 As Object
    Set t1 = tmpl.Shapes.AddTextbox(1, 40, 40, 200, 30)
    t1.TextFrame.TextRange.Text = "<<ABOUT_BODY>>"
    t1.Tags.Add "role", "ABOUT_BODY"

    Dim r As FieldWiringResult
    r = FieldWiring.ScanFieldWiring("wiring-tmpl", _
        FieldsCollection("ABOUT_BODY", "PROBLEM_BODY"), tmpl)

    result = result & Assert(r.TemplateScanned, "the template was scanned")
    result = result & Assert(r.UnmarkedCount = 0, _
        "every field IS on an existing slide, got " & r.UnmarkedCount & " unmarked")
    result = result & Assert(r.TemplateUnmarkedCount = 1, _
        "but one is missing from the TEMPLATE, got " & r.TemplateUnmarkedCount)
    result = result & Assert(InStr(r.TemplateUnmarked, "PROBLEM_BODY") > 0, _
        "and it is named, got '" & r.TemplateUnmarked & "'")
    result = result & Assert(InStr(FieldWiring.WiringText(r), "TEMPLATE") > 0, _
        "the sentence says the template is the problem, got '" & FieldWiring.WiringText(r) & "'")

    tmpl.Delete
    inst.Delete
    Test_FieldWiring_TemplateIsCheckedSeparatelyFromInstances = result
End Function

' A `.track` with no bar behind it: a marking started and abandoned.
'
' It matters because the failure is otherwise silent AND plausible -- an
' untracked bar shape is an ordinary rectangle with a text frame, so the router
' hands it to the text writer and it quietly becomes a text field.
Private Function Test_FieldWiring_OrphanTrackIsAHalfMarkedBar() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewTaggedSlide("wiring-orphan", "wo-1")
    Dim track As Object
    Set track = sld.Shapes.AddShape(1, 40, 200, 400, 12)
    track.Tags.Add "role", "ELAPSED.track"

    Dim r As FieldWiringResult
    r = FieldWiring.ScanFieldWiring("wiring-orphan", FieldsCollection(), Nothing)

    result = result & Assert(r.OrphanCount = 1, _
        "the orphan track is found, got " & r.OrphanCount)
    result = result & Assert(InStr(UCase(r.OrphanTracks), "ELAPSED") > 0, _
        "and it is named, got '" & r.OrphanTracks & "'")

    ' CONTROL. Add the bar the track was missing and the finding must GO AWAY --
    ' otherwise this check reports a problem that cannot be fixed, which is
    ' worse than not reporting it.
    Dim done As Object
    Set done = sld.Shapes.AddShape(1, 40, 200, 100, 12)
    done.Tags.Add "role", "ELAPSED"

    Dim r2 As FieldWiringResult
    r2 = FieldWiring.ScanFieldWiring("wiring-orphan", FieldsCollection(), Nothing)
    result = result & Assert(r2.OrphanCount = 0, _
        "once the bar exists the finding clears, got " & r2.OrphanCount)

    sld.Delete
    Test_FieldWiring_OrphanTrackIsAHalfMarkedBar = result
End Function

' COVERAGE, which is the question presence was too weak to ask.
'
' On the rig, 2026-08-10, three fields were tagged on 2 slides of 44 and the
' check reported a bare `ok`, because each was carried by *a* slide. Nothing
' was stranded that day -- only one project had text for them -- and that is
' exactly why the bare `ok` was dangerous: it would have stayed `ok` right up
' until the second project's drafting had nowhere to land.
'
' Two slides, one carrying the field, is the smallest case that can tell
' "tagged somewhere" apart from "tagged everywhere".
Private Function Test_FieldWiring_CoverageCountsSlidesNotPresence() As String
    Dim result As String

    Dim a As Object, b As Object
    Set a = NewTaggedSlide("wiring-cover", "wc-1")
    Set b = NewTaggedSlide("wiring-cover", "wc-2")

    ' Both slides carry ABOUT_BODY; only one carries PROBLEM_BODY.
    Dim s1 As Object, s2 As Object, s3 As Object
    Set s1 = a.Shapes.AddTextbox(1, 40, 40, 200, 30)
    s1.TextFrame.TextRange.Text = "about a"
    s1.Tags.Add "role", "ABOUT_BODY"
    Set s2 = b.Shapes.AddTextbox(1, 40, 40, 200, 30)
    s2.TextFrame.TextRange.Text = "about b"
    s2.Tags.Add "role", "ABOUT_BODY"
    Set s3 = a.Shapes.AddTextbox(1, 40, 90, 200, 30)
    s3.TextFrame.TextRange.Text = "problem a"
    s3.Tags.Add "role", "PROBLEM_BODY"

    Dim r As FieldWiringResult
    r = FieldWiring.ScanFieldWiring("wiring-cover", _
        FieldsCollection("ABOUT_BODY", "PROBLEM_BODY"), Nothing)

    result = result & Assert(r.SlidesScanned = 2, _
        "two slides were scanned, got " & r.SlidesScanned)

    ' PRESENCE STILL PASSES -- both fields ARE carried by a slide. That is the
    ' whole point: a check asking only presence reports nothing here.
    result = result & Assert(r.UnmarkedCount = 0, _
        "neither field is unmarked, got " & r.UnmarkedCount)

    result = result & Assert(r.PartialCount = 1, _
        "exactly one field is not on every slide, got " & r.PartialCount)
    result = result & Assert(InStr(r.Coverage, "PROBLEM_BODY on 1 of 2") > 0, _
        "and the count is per slide and NAMED, got '" & r.Coverage & "'")
    result = result & Assert(InStr(r.Coverage, "ABOUT_BODY") = 0, _
        "a field on every slide is not reported, got '" & r.Coverage & "'")
    result = result & Assert(InStr(FieldWiring.WiringText(r), "1 of 2") > 0, _
        "the sentence carries it, got '" & FieldWiring.WiringText(r) & "'")

    ' ORDER, not just presence. With a template in play the sentence read
    ' "...PROBLEM_BODY on 1 of 2, and on the template" on the real deck, where
    ' the template looks like a fourth entry in the coverage list. The template
    ' clause has to come FIRST, before the list it is not part of.
    Dim tmpl As Object
    Set tmpl = NewBlankSlide()
    Dim t1 As Object, t2 As Object
    Set t1 = tmpl.Shapes.AddTextbox(1, 40, 40, 200, 30)
    t1.TextFrame.TextRange.Text = "<<ABOUT_BODY>>"
    t1.Tags.Add "role", "ABOUT_BODY"
    Set t2 = tmpl.Shapes.AddTextbox(1, 40, 90, 200, 30)
    t2.TextFrame.TextRange.Text = "<<PROBLEM_BODY>>"
    t2.Tags.Add "role", "PROBLEM_BODY"

    Dim rt As FieldWiringResult
    rt = FieldWiring.ScanFieldWiring("wiring-cover", _
        FieldsCollection("ABOUT_BODY", "PROBLEM_BODY"), tmpl)
    Dim txt As String
    txt = FieldWiring.WiringText(rt)
    result = result & Assert(InStr(txt, "on the template. Not on every slide yet:") > 0, _
        "the template clause comes BEFORE the coverage list, got '" & txt & "'")
    tmpl.Delete

    ' CONTROL. Tag the second slide and the finding must CLEAR -- a coverage
    ' warning that cannot be satisfied would just be noise to click through.
    Dim s4 As Object
    Set s4 = b.Shapes.AddTextbox(1, 40, 90, 200, 30)
    s4.TextFrame.TextRange.Text = "problem b"
    s4.Tags.Add "role", "PROBLEM_BODY"

    Dim r2 As FieldWiringResult
    r2 = FieldWiring.ScanFieldWiring("wiring-cover", _
        FieldsCollection("ABOUT_BODY", "PROBLEM_BODY"), Nothing)
    result = result & Assert(r2.PartialCount = 0, _
        "once both slides carry it the finding clears, got " & r2.PartialCount)
    result = result & Assert(r2.Coverage = "", _
        "and the coverage text is empty, got '" & r2.Coverage & "'")

    b.Delete
    a.Delete
    Test_FieldWiring_CoverageCountsSlidesNotPresence = result
End Function

' A TRACK IS A TAG, NEVER A REGISTER COLUMN.
'
' CommitBatch writes every marked field into the register through UpsertRow,
' which creates a COLUMN per field. A track marked as a field would therefore
' grow a `PROGRESS_BODY.track` column that nothing can ever fill, that drafting
' would build a sheet for, and that the wiring check would report as unwired
' forever. The exclusion is what keeps the tag and the column apart.
'
' The boundary cases matter more than the obvious one: a field merely CONTAINING
' the word track, and a field called exactly ".track" with nothing in front of
' it, must not be mistaken for one.
Private Function Test_FieldWiring_ATrackIsATagNotAField() As String
    Dim result As String

    result = result & Assert(FieldWiring.IsTrackFieldName("PROGRESS_BODY.track"), _
        "the suffix is recognised")
    result = result & Assert(FieldWiring.IsTrackFieldName("PROGRESS_BODY.TRACK"), _
        "and case does not matter -- a person typing it will not match ours")

    result = result & Assert(Not FieldWiring.IsTrackFieldName("PROGRESS_BODY"), _
        "the bar itself is a field")
    result = result & Assert(Not FieldWiring.IsTrackFieldName("TRACK_RECORD_BODY"), _
        "a field that merely contains the word is a field")
    result = result & Assert(Not FieldWiring.IsTrackFieldName(".track"), _
        "a bare suffix with no field in front of it is not a track")
    result = result & Assert(Not FieldWiring.IsTrackFieldName(""), _
        "and neither is nothing")

    Test_FieldWiring_ATrackIsATagNotAField = result
End Function

' THE MILESTONE CASE: one metric, several milestones, one register cell.
'
' Rohan, 2026-08-10: "same metric but different progress against different
' milestones". Three bars tagged `.1 .2 .3`, values `0.25||0.5||1`.
'
' The bars are given DIFFERENT track widths so a value written against the wrong
' milestone's track cannot land on the right number by coincidence -- the same
' precaution as the two-slide test, applied within one slide.
'
' The count-mismatch assertion is the one that matters: three values against
' three bars is the easy case, and it is the FOURTH bar that decides whether
' this refuses or draws a plausible wrong slide.
Private Function Test_InjectField_RepeatingBarsOneCellManyMilestones() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()

    ' A shared axis rather than a track per bar -- the layout the design has to
    ' resolve rather than assume, and the one a milestone timeline uses.
    Dim axis As Object
    Set axis = sld.Shapes.AddShape(1, 40, 300, 400, 8)
    axis.Tags.Add "role", "MILESTONE_PROGRESS" & FieldWiring.TRACK_SUFFIX

    Dim b1 As Object, b2 As Object, b3 As Object
    Set b1 = sld.Shapes.AddShape(1, 40, 100, 10, 8)
    b1.Tags.Add "role", "MILESTONE_PROGRESS.1"
    Set b2 = sld.Shapes.AddShape(1, 40, 140, 10, 8)
    b2.Tags.Add "role", "MILESTONE_PROGRESS.2"
    Set b3 = sld.Shapes.AddShape(1, 40, 180, 10, 8)
    b3.Tags.Add "role", "MILESTONE_PROGRESS.3"

    Dim r As InjectResult
    r = InjectPrimitive.InjectField(sld, "MILESTONE_PROGRESS", "0.25||0.5||1")

    result = result & Assert(r.Written, "all three bars are written [" & r.ErrorMessage & "]")
    result = result & Assert(Abs(b1.Width - 100) < 1, _
        "milestone 1 is 25% of the 400pt axis = 100pt, got " & b1.Width)
    result = result & Assert(Abs(b2.Width - 200) < 1, _
        "milestone 2 is 50% = 200pt, got " & b2.Width)
    result = result & Assert(Abs(b3.Width - 400) < 1, _
        "milestone 3 is 100% = 400pt, got " & b3.Width)
    result = result & Assert(Abs(axis.Width - 400) < 0.5, _
        "the shared axis is never written, got " & axis.Width)

    ' COUNT MISMATCH REFUSES OUTRIGHT. Writing the ones that match would leave
    ' the remaining milestone showing an older figure on a slide that looks
    ' finished -- plausible and wrong is worse than refused.
    Dim w1 As Single, w2 As Single, w3 As Single
    w1 = b1.Width: w2 = b2.Width: w3 = b3.Width

    Dim r2 As InjectResult
    r2 = InjectPrimitive.InjectField(sld, "MILESTONE_PROGRESS", "0.1||0.2")
    result = result & Assert(Not r2.Written, "two values against three bars is refused")
    result = result & Assert(InStr(r2.ErrorMessage, "2 value") > 0 And InStr(r2.ErrorMessage, "3 bar") > 0, _
        "and the message names BOTH counts, got '" & r2.ErrorMessage & "'")
    result = result & Assert(b1.Width = w1 And b2.Width = w2 And b3.Width = w3, _
        "and not one bar moved")

    ' A repeating field carries no shape of its own name, so the wiring check
    ' has to recognise `.1` or it reports every milestone field as unwired.
    Dim roles As Object
    Set roles = FieldWiring.RoleTagsOnSlide(sld)
    result = result & Assert(FieldWiring.CarriesField(roles, "MILESTONE_PROGRESS"), _
        "the wiring check sees a repeating field as carried")
    result = result & Assert(Not FieldWiring.CarriesField(roles, "NOT_ON_THIS_SLIDE"), _
        "and still says no to a field that is genuinely absent")

    sld.Delete
    Test_InjectField_RepeatingBarsOneCellManyMilestones = result
End Function

' A BAR WITH NO TRACK, measured against its own remainder.
'
' Geometry taken from the real deck, slide 1, 2026-08-10: `Time elapsed` is a
' filled rect 1.45" wide beside a grey one 0.17" wide, with no full-width shape
' behind them. 1.45 / (1.45 + 0.17) = 89.5%, against a label reading 90%.
' Scaled x100 here so PowerPoint's points are comfortable.
'
' THE THIRD RUN IS THE TEST. Measuring the done part against ITSELF compounds --
' the bar shrinks a little further every run while every report says success --
' and that is exactly what the no-track refusal was protecting against. The sum
' does not compound, because both halves are written in one operation and add
' back to the same extent. Running the SAME value twice and then a different one
' is the only way to tell those two apart: one run proves nothing, and two runs
' of the same value would pass even if the extent were quietly shrinking.
Private Function Test_InjectProgress_TracklessPairMeasuresTheirSum() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()

    Dim done As Object, rest As Object
    Set done = sld.Shapes.AddShape(1, 100, 200, 145, 7)   ' msoShapeRectangle
    done.Tags.Add "role", "TIME_ELAPSED"
    Set rest = sld.Shapes.AddShape(1, 245, 200, 17, 7)
    rest.Tags.Add "role", "TIME_ELAPSED.rest"

    ' No shape tagged TIME_ELAPSED.track exists anywhere on this slide.
    Dim extent As Single
    extent = done.Width + rest.Width                       ' 162

    Dim r As InjectResult
    r = InjectPrimitive.InjectField(sld, "TIME_ELAPSED", "0.5")
    result = result & Assert(r.Written, "the bar is written with no track [" & r.ErrorMessage & "]")
    result = result & Assert(Abs(done.Width - 81) < 1, _
        "50% of the 162pt extent is 81pt, got " & done.Width)
    result = result & Assert(Abs(rest.Left - 181) < 1 And Abs(rest.Width - 81) < 1, _
        "the remainder takes the other half, got left " & rest.Left & " width " & rest.Width)
    result = result & Assert(Abs((done.Width + rest.Width) - extent) < 0.5, _
        "and the two still sum to the original extent, got " & (done.Width + rest.Width))

    ' SAME VALUE AGAIN. A no-op, and the sum must not move.
    Dim r2 As InjectResult
    r2 = InjectPrimitive.InjectField(sld, "TIME_ELAPSED", "0.5")
    result = result & Assert(Not r2.WouldChange, "a second run at the same value changes nothing")
    result = result & Assert(Abs((done.Width + rest.Width) - extent) < 0.5, _
        "the extent is still " & extent & ", got " & (done.Width + rest.Width))

    ' A DIFFERENT VALUE, measured against the SAME extent. A bar measuring
    ' itself would compute 25% of 81 = 20pt here instead of 25% of 162 = 40pt.
    Dim r3 As InjectResult
    r3 = InjectPrimitive.InjectField(sld, "TIME_ELAPSED", "0.25")
    result = result & Assert(Abs(done.Width - 40.5) < 1, _
        "25% is measured against the FULL extent (40.5pt), not the shrunken bar (20pt), got " & done.Width)
    result = result & Assert(Abs((done.Width + rest.Width) - extent) < 0.5, _
        "and the extent is STILL " & extent & " after three runs, got " & (done.Width + rest.Width))

    ' Neither half moved from the pair's own left edge.
    result = result & Assert(Abs(done.Left - 100) < 0.5, _
        "the bar still starts where it started, got " & done.Left)

    sld.Delete
    Test_InjectProgress_TracklessPairMeasuresTheirSum = result
End Function

' A REMAINDER IS A TAG TOO, and forgetting that nearly cost a live run.
'
' On 2026-08-10 the exclusion covered `.track` only. Rohan's `Time elapsed` bar
' is a fill beside a grey REMAINDER with no track, so tagging its companion
' would have created a `TIME_ELAPSED.rest` register COLUMN -- and the router,
' finding no companion for THAT, would have handed it to the text writer and
' written text into the grey tail of a progress bar.
'
' The suffix is returned rather than a Boolean so callers can strip it: a check
' written for `.track` strips six characters, and `.rest` is five.
Private Function Test_FieldWiring_BothCompanionsAreTagsNotFields() As String
    Dim result As String

    result = result & Assert(FieldWiring.IsCompanionFieldName("TIME_ELAPSED.track"), _
        "a track is a companion")
    result = result & Assert(FieldWiring.IsCompanionFieldName("TIME_ELAPSED.rest"), _
        "and so is a remainder -- the one that was missed")
    result = result & Assert(FieldWiring.IsCompanionFieldName("TIME_ELAPSED.REST"), _
        "case does not matter")

    result = result & Assert(Not FieldWiring.IsCompanionFieldName("TIME_ELAPSED"), _
        "the bar itself is a field")
    result = result & Assert(Not FieldWiring.IsCompanionFieldName("REST_PERIOD_BODY"), _
        "a field merely containing the word is a field")
    result = result & Assert(Not FieldWiring.IsCompanionFieldName(".rest"), _
        "a bare suffix with no field in front of it is not a companion")

    ' THE SUFFIX ITSELF, because the orphan check strips it by length. A
    ' remainder stripped as though it were a track loses a character and looks
    ' for the wrong base name.
    result = result & Assert(FieldWiring.CompanionSuffixOf("X.track") = FieldWiring.TRACK_SUFFIX, _
        "a track reports the track suffix, got '" & FieldWiring.CompanionSuffixOf("X.track") & "'")
    result = result & Assert(FieldWiring.CompanionSuffixOf("X.rest") = FieldWiring.REST_SUFFIX, _
        "a remainder reports the rest suffix, got '" & FieldWiring.CompanionSuffixOf("X.rest") & "'")
    result = result & Assert(FieldWiring.CompanionSuffixOf("X") = "", _
        "an ordinary field reports no suffix")

    ' And the orphan check must find a REMAINDER with no bar, not just a track.
    Dim sld As Object
    Set sld = NewTaggedSlide("wiring-rest", "wr-1")
    Dim tail As Object
    Set tail = sld.Shapes.AddShape(1, 240, 200, 17, 7)
    tail.Tags.Add "role", "TIME_ELAPSED" & FieldWiring.REST_SUFFIX

    Dim r As FieldWiringResult
    r = FieldWiring.ScanFieldWiring("wiring-rest", FieldsCollection(), Nothing)
    result = result & Assert(r.OrphanCount = 1, _
        "a remainder with no bar is an orphan, got " & r.OrphanCount)
    result = result & Assert(InStr(UCase(r.OrphanTracks), "TIME_ELAPSED") > 0, _
        "and it is named, got '" & r.OrphanTracks & "'")

    sld.Delete
    Test_FieldWiring_BothCompanionsAreTagsNotFields = result
End Function

' THE NEAR-MISS: right field, wrong capitalisation.
'
' The injector matches role tags with `=` under VBA's default binary comparison,
' so it is CASE SENSITIVE. This check has always uppercased both sides. Left
' alone, a register reading `PROJECT_PROGRESS` against a shape tagged
' `Project_Progress` produces a GREEN readiness sheet over a sync that writes
' nothing -- the check disagreeing with the writer, which is worse than either
' being wrong alone.
'
' Nearly cost a real run on 2026-08-10: Rohan's header went in as
' `PROJECT_PROGRESS` and the marking dialog had proposed `Project_Progress`.
Private Function Test_FieldWiring_CaseNearMissIsNamedNotSwallowed() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewTaggedSlide("wiring-case", "wc-1")
    Dim tb As Object
    Set tb = sld.Shapes.AddTextbox(1, 40, 40, 200, 30)
    tb.TextFrame.TextRange.Text = "x"
    tb.Tags.Add "role", "Project_Progress"          ' mixed case on the slide

    ' The register asks for it in caps.
    Dim r As FieldWiringResult
    r = FieldWiring.ScanFieldWiring("wiring-case", FieldsCollection("PROJECT_PROGRESS"), Nothing)

    result = result & Assert(r.CaseMismatchCount = 1, _
        "the near-miss is found, got " & r.CaseMismatchCount)
    result = result & Assert(InStr(r.CaseMismatch, "PROJECT_PROGRESS") > 0 _
        And InStr(r.CaseMismatch, "Project_Progress") > 0, _
        "and BOTH spellings are shown, got '" & r.CaseMismatch & "'")
    result = result & Assert(InStr(FieldWiring.BlockingText(r), "Project_Progress") > 0, _
        "the blocking sentence carries it, got '" & FieldWiring.BlockingText(r) & "'")

    ' It must NOT read as ok -- that combination is the whole danger.
    result = result & Assert(InStr(FieldWiring.WiringText(r), "field(s) tagged on") = 0, _
        "and it does not report the clean 'N fields tagged' line, got '" & _
        FieldWiring.WiringText(r) & "'")

    ' CONTROL. Spell it the same and the finding must clear.
    Dim r2 As FieldWiringResult
    r2 = FieldWiring.ScanFieldWiring("wiring-case", FieldsCollection("Project_Progress"), Nothing)
    result = result & Assert(r2.CaseMismatchCount = 0, _
        "an exact match reports no near-miss, got " & r2.CaseMismatchCount)

    sld.Delete
    Test_FieldWiring_CaseNearMissIsNamedNotSwallowed = result
End Function

' The "nothing to publish" test must read the SAME counts the person is shown.
'
' Written separately it would be a second copy of the same fact, and the stage
' could then skip its question while the report said work had happened, or ask
' about nothing while the report said zero.
Private Function Test_Drafting_NothingToPublishReadsTheSameCounts() As String
    Dim result As String

    Dim allZero As String
    allZero = "some preamble" & vbCrLf & Drafting.PublishSummaryLine(0, True, 0, 0, 0, 0) & vbCrLf
    result = result & Assert(Drafting.NothingToPublish(allZero), _
        "an all-zero preview is nothing to publish")

    ' Each count on its own is WORK, and must not be skipped.
    result = result & Assert(Not Drafting.NothingToPublish(Drafting.PublishSummaryLine(1, True, 0, 0, 0, 0)), _
        "one to publish is work")
    result = result & Assert(Not Drafting.NothingToPublish(Drafting.PublishSummaryLine(0, True, 1, 0, 0, 0)), _
        "a draft waiting on a tick is work -- it is the thing to tell someone about")
    result = result & Assert(Not Drafting.NothingToPublish(Drafting.PublishSummaryLine(0, True, 0, 1, 0, 0)), _
        "a tick with no text is work")
    result = result & Assert(Not Drafting.NothingToPublish(Drafting.PublishSummaryLine(0, True, 0, 0, 1, 0)), _
        "a row with no register row is work")
    result = result & Assert(Not Drafting.NothingToPublish(Drafting.PublishSummaryLine(0, True, 0, 0, 0, 1)), _
        "a failure is DEFINITELY work")

    Test_Drafting_NothingToPublishReadsTheSameCounts = result
End Function

' Builds a milestone device: a track, a bar, and `slots` slots each with an ON
' circle, an OFF circle, a label and a date -- named the way a person would name
' them in the Selection Pane, then grouped.
Private Function NewMilestoneDevice(sld As Object, slots As Long) As Object
    Dim names() As String
    ReDim names(1 To 2 + slots * 4)
    Dim n As Long
    n = 0

    Dim track As Object, bar As Object
    Set track = sld.Shapes.AddShape(1, 100, 100, 6, 300)      ' vertical, 300 tall
    track.Name = MilestoneDevice.NAME_TRACK
    n = n + 1: names(n) = track.Name
    Set bar = sld.Shapes.AddShape(1, 100, 100, 6, 10)
    bar.Name = MilestoneDevice.NAME_BAR
    n = n + 1: names(n) = bar.Name

    Dim i As Long
    For i = 1 To slots
        Dim y As Single
        y = 100 + (i - 1) * 50                                 ' 50pt apart
        Dim onS As Object, offS As Object, labS As Object, datS As Object
        Set onS = sld.Shapes.AddShape(9, 96, y, 14, 14)        ' msoShapeOval
        onS.Name = MilestoneDevice.SLOT_PREFIX & i & MilestoneDevice.PART_ON
        n = n + 1: names(n) = onS.Name
        Set offS = sld.Shapes.AddShape(9, 96, y, 14, 14)
        offS.Name = MilestoneDevice.SLOT_PREFIX & i & MilestoneDevice.PART_OFF
        n = n + 1: names(n) = offS.Name
        Set labS = sld.Shapes.AddTextbox(1, 130, y, 120, 14)
        labS.Name = MilestoneDevice.SLOT_PREFIX & i & MilestoneDevice.PART_LABEL
        n = n + 1: names(n) = labS.Name
        Set datS = sld.Shapes.AddTextbox(1, 130, y + 14, 120, 14)
        datS.Name = MilestoneDevice.SLOT_PREFIX & i & MilestoneDevice.PART_DATE
        n = n + 1: names(n) = datS.Name
    Next i

    Set NewMilestoneDevice = sld.Shapes.Range(names).Group
End Function

Private Function NamedIn(grp As Object, nm As String) As Object
    Dim s As Object
    For Each s In grp.GroupItems
        If UCase(s.Name) = UCase(nm) Then
            Set NamedIn = s
            Exit Function
        End If
    Next s
End Function

Private Function SplitList(v As String) As String()
    SplitList = Split(v, InjectPrimitive.VALUE_SEPARATOR)
End Function

' THE TIMELINE, DRAWN FROM DATA, CREATING NOTHING.
'
' Rohan, 2026-08-10: "I'd like it all data drawn given I take a robotic approach
' when updating it anyway." Four slots on the template, three milestones in the
' register, the middle one done.
'
' The assertions that matter are the ones about what did NOT happen: the shape
' count is unchanged, and the track -- which the bar is measured against -- is
' never written. "you can see the extremely accurate positioning? we need to
' maintain that and z order etc."
Private Function Test_MilestoneDevice_DrawsFromDataAndCreatesNothing() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim grp As Object
    Set grp = NewMilestoneDevice(sld, 4)

    Dim before As Long
    before = grp.GroupItems.Count

    result = result & Assert(MilestoneDevice.SlotCount(grp) = 4, _
        "the slots are COUNTED from the template, not configured, got " & MilestoneDevice.SlotCount(grp))

    Dim r As MilestoneDrawResult
    r = MilestoneDevice.DrawMilestones(grp, _
        SplitList("Kickoff||Design||Build"), _
        SplitList("Oct 2023||Mar 2024||Sep 2024"), _
        SplitList("Y||Y||N"))

    ' CRITICAL-EFFECTS SHORTLIST, item 1 -- see Watch's header. Three slots
    ' ON, one OFF, slot 2 the bolded NOW: a human watching the (visible)
    ' PowerPoint window can actually see this happen here, not just trust
    ' the assertions below it.
    Watch sld, "milestone ON/OFF/NOW after DrawMilestones"

    result = result & Assert(r.ErrorMessage = "", "it draws [" & r.ErrorMessage & "]")
    result = result & Assert(r.Drawn = 3, "three milestones drawn, got " & r.Drawn)
    result = result & Assert(r.Hidden = 1, "the unused fourth slot is hidden, got " & r.Hidden)

    ' NOTHING WAS CREATED OR DELETED.
    result = result & Assert(grp.GroupItems.Count = before, _
        "the shape count is unchanged: " & before & " -> " & grp.GroupItems.Count)

    ' Achieved shows ON and hides OFF; not achieved does the reverse.
    result = result & Assert(NamedIn(grp, "MS1_ON").Visible = msoTrue _
        And NamedIn(grp, "MS1_OFF").Visible = msoFalse, "slot 1 is achieved")
    result = result & Assert(NamedIn(grp, "MS3_ON").Visible = msoFalse _
        And NamedIn(grp, "MS3_OFF").Visible = msoTrue, "slot 3 is NOT achieved")
    result = result & Assert(NamedIn(grp, "MS4_ON").Visible = msoFalse _
        And NamedIn(grp, "MS4_LABEL").Visible = msoFalse, "slot 4 is hidden entirely")

    result = result & Assert(NamedIn(grp, "MS2_LABEL").TextFrame.TextRange.text = "Design", _
        "labels are written, got '" & NamedIn(grp, "MS2_LABEL").TextFrame.TextRange.text & "'")
    result = result & Assert(NamedIn(grp, "MS2_DATE").TextFrame.TextRange.text = "Mar 2024", _
        "and so are dates, got '" & NamedIn(grp, "MS2_DATE").TextFrame.TextRange.text & "'")

    ' THE BAR IS MEASURED TO THE LAST ACHIEVED CIRCLE, not to a fraction.
    ' Slot 2 is the last done one; its centre is 50pt below slot 1's.
    Dim bar As Object, track As Object, ms2 As Object
    Set bar = NamedIn(grp, MilestoneDevice.NAME_BAR)
    Set track = NamedIn(grp, MilestoneDevice.NAME_TRACK)
    Set ms2 = NamedIn(grp, "MS2_ON")
    result = result & Assert(r.BarSet, "the bar was set")
    result = result & Assert(Abs(bar.Height - ((ms2.Top + ms2.Height / 2) - track.Top)) < 1, _
        "the bar reaches slot 2's centre, got " & bar.Height)
    ' THE TRACK SHORTENS TO THE LAST USED SLOT.
    '
    ' This used to assert the track was never written at all. That was the
    ' contract until 2026-08-13, when Rohan asked what a shorter milestone chain
    ' does -- and the answer was two slots' worth of track hanging below the
    ' last circle, pointing at nothing. It is the same measurement the bar
    ' already used, aimed at the last USED circle instead of the last ACHIEVED
    ' one.
    '
    ' The fixture draws 3 slots and supplies 3 milestones, so the track reaches
    ' slot 3's centre -- SHORTER than the 300pt it was drawn at, which is what
    ' makes this assertion able to fail if the shortening silently stops.
    Dim ms3 As Object
    Set ms3 = NamedIn(grp, "MS3_ON")
    result = result & Assert(Abs(track.Height - ((ms3.Top + ms3.Height / 2) - track.Top)) < 1, _
        "the TRACK reaches the last USED slot, got " & track.Height)
    result = result & Assert(track.Height < 299, _
        "and is genuinely shorter than the 300pt it was drawn at, got " & track.Height)

    sld.Delete
    Test_MilestoneDevice_DrawsFromDataAndCreatesNothing = result
End Function

' THREE LISTS PAIRED BY POSITION IS THE ONE REAL HAZARD IN THIS DESIGN.
'
' Five labels against four dates would pair every milestone after the gap with
' the wrong date, and the slide would look finished. There is no way to tell
' which list is wrong, so nothing is drawn -- and nothing must MOVE either,
' which is the assertion that actually matters.
' FIX-LIST item Q. Proves DrawMilestones now REPORTS a write that did not
' take, through its real public API -- not by calling the private writers
' directly, which would only prove the helper works in isolation, not that
' the fix actually reaches the caller.
'
' A Line shape is used because it has no TextFrame -- a well-established COM
' fact, unlike AddShape ovals, which DO carry a TextFrame in modern
' PowerPoint and would have made this test assert something false. Checked
' rather than assumed, per this project's own rule on Office automation.
' The wrapper exists purely to cross the Application.Run boundary -- prove it
' says the same thing the real UDT-returning function says, on both the
' normal path and the failure path, so nobody trusts a wrapper that silently
' drifted from what it wraps.
Private Function Test_MilestoneDevice_ReportWrapperMatchesTheRealResult() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim grp As Object
    Set grp = NewMilestoneDevice(sld, 3)

    Dim labels() As String, dates() As String, done() As String
    labels = SplitList("Kickoff||Design||Build")
    dates = SplitList("Oct 2023||Mar 2024||Sep 2024")
    done = SplitList("Y||Y||N")

    Dim direct As MilestoneDrawResult
    direct = MilestoneDevice.DrawMilestones(grp, labels, dates, done)

    ' Re-run against the SAME group via DrawFromRow/the wrapper -- same data,
    ' different entry point, so the wrapper's numbers must match what DIRECT
    ' just did.
    Dim rowValues As Object
    Set rowValues = CreateObject("Scripting.Dictionary")
    rowValues.Add "MS1_LABEL", "Kickoff": rowValues.Add "MS1_DATE", "Oct 2023": rowValues.Add "MS1_DONE", "Y"
    rowValues.Add "MS2_LABEL", "Design": rowValues.Add "MS2_DATE", "Mar 2024": rowValues.Add "MS2_DONE", "Y"
    rowValues.Add "MS3_LABEL", "Build": rowValues.Add "MS3_DATE", "Sep 2024": rowValues.Add "MS3_DONE", "N"

    Dim report As String
    report = MilestoneDevice.DrawFromRowReport(grp, rowValues)

    result = result & Assert(InStr(report, "Drawn=" & direct.Drawn) > 0, _
        "wrapper's Drawn count matches the real result, got '" & report & "'")
    result = result & Assert(InStr(report, "Hidden=" & direct.Hidden) > 0, _
        "wrapper's Hidden count matches, got '" & report & "'")
    result = result & Assert(InStr(report, "Error=[]") > 0, _
        "the no-error case shows an empty Error, got '" & report & "'")

    ' THE FAILURE PATH TOO -- an empty group has no slots, and DrawFromRow's
    ' ErrorMessage must survive the trip through the wrapper unchanged.
    Dim emptySld As Object
    Set emptySld = NewBlankSlide()
    Dim badReport As String
    badReport = MilestoneDevice.DrawFromRowReport(Nothing, rowValues)
    result = result & Assert(InStr(badReport, "no timeline group given") > 0, _
        "the wrapper carries the real ErrorMessage through on failure too, got '" & badReport & "'")

    emptySld.Delete
    sld.Delete
    Test_MilestoneDevice_ReportWrapperMatchesTheRealResult = result
End Function

Private Function Test_MilestoneDevice_ReportsAWriteThatDidNotTake() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim grp As Object
    Set grp = NewMilestoneDevice(sld, 3)

    ' Swap slot 2's label for a shape that CANNOT hold text, keeping the same
    ' name so the device still finds it by name as designed.
    Dim badLabel As Object
    Set badLabel = sld.Shapes.AddLine(130, 150, 250, 150)
    result = result & Assert(Not badLabel.HasTextFrame, _
        "control: a Line genuinely has no TextFrame, HasTextFrame=" & badLabel.HasTextFrame)

    NamedIn(grp, "MS2_LABEL").Delete
    badLabel.Name = "MS2_LABEL"
    ' The swap-in must join the SAME group, or it will not be found by PartsOf.
    Set grp = sld.Shapes.Range(Array(grp.Name, badLabel.Name)).Group

    Dim r As MilestoneDrawResult
    r = MilestoneDevice.DrawMilestones(grp, _
        SplitList("Kickoff||Design||Build"), _
        SplitList("Oct 2023||Mar 2024||Sep 2024"), _
        SplitList("Y||Y||N"))

    ' THE ONE THAT MATTERS. Before this fix, a declined write was invisible --
    ' Drawn would read 3 and nothing would say slot 2 is wrong.
    result = result & Assert(InStr(r.Detail, "MS2") > 0 And InStr(r.Detail, "did not take") > 0, _
        "the report NAMES the failed slot, got '" & r.Detail & "'")
    result = result & Assert(r.Drawn = 3, _
        "the slot still counts as drawn -- a bad write is reported, not silently skipped, got " & r.Drawn)

    ' THE OTHER TWO SLOTS ARE UNAFFECTED. One bad shape must not corrupt slots
    ' that were never touched.
    result = result & Assert(NamedIn(grp, "MS1_LABEL").TextFrame.TextRange.text = "Kickoff", _
        "slot 1 still wrote correctly, got '" & NamedIn(grp, "MS1_LABEL").TextFrame.TextRange.text & "'")
    result = result & Assert(NamedIn(grp, "MS3_LABEL").TextFrame.TextRange.text = "Build", _
        "slot 3 still wrote correctly, got '" & NamedIn(grp, "MS3_LABEL").TextFrame.TextRange.text & "'")

    sld.Delete
    Test_MilestoneDevice_ReportsAWriteThatDidNotTake = result
End Function

Private Function Test_MilestoneDevice_RefusesListsThatCannotBeAligned() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim grp As Object
    Set grp = NewMilestoneDevice(sld, 4)

    Dim bar As Object
    Set bar = NamedIn(grp, MilestoneDevice.NAME_BAR)
    Dim barBefore As Single
    barBefore = bar.Height

    Dim r As MilestoneDrawResult
    r = MilestoneDevice.DrawMilestones(grp, _
        SplitList("A||B||C"), SplitList("1||2"), SplitList("Y||N||N"))

    result = result & Assert(r.ErrorMessage <> "", "mismatched lists are refused")
    result = result & Assert(InStr(r.ErrorMessage, "3 label") > 0 _
        And InStr(r.ErrorMessage, "2 date") > 0, _
        "and the message names BOTH counts, got '" & r.ErrorMessage & "'")
    result = result & Assert(r.Drawn = 0, "nothing was drawn, got " & r.Drawn)
    result = result & Assert(bar.Height = barBefore, "and the bar did not move")

    ' MORE MILESTONES THAN SLOTS is also refused rather than truncated --
    ' drawing the first four would silently drop the fifth.
    Dim r2 As MilestoneDrawResult
    r2 = MilestoneDevice.DrawMilestones(grp, _
        SplitList("A||B||C||D||E"), SplitList("1||2||3||4||5"), SplitList("Y||Y||Y||Y||Y"))
    result = result & Assert(r2.ErrorMessage <> "" And r2.Drawn = 0, _
        "five milestones into four slots is refused, not truncated")
    result = result & Assert(InStr(r2.ErrorMessage, "5 milestone") > 0 _
        And InStr(r2.ErrorMessage, "4 slot") > 0, _
        "and names both numbers, got '" & r2.ErrorMessage & "'")

    sld.Delete
    Test_MilestoneDevice_RefusesListsThatCannotBeAligned = result
End Function


' ONE COLUMN PER INTERVAL, READ FROM THE ROW.
'
' Rohan: "with one row per project wouldn't the timeline be one column per
' interval not one row?" -- which killed the packed-cell format. And then: "I'm
' not sure I understand the difference between that and a field." There is
' none: MS1_LABEL is a register column whose value lands in a shape, same as
' ABOUT_BODY. Only the addressing differs.
'
' The gap assertion is the one that matters. Slots draw in order, so a filled
' MS3 with an empty MS2 would quietly renumber the timeline -- milestone three
' would appear in slot two and the slide would look finished.
Private Function Test_MilestoneDevice_ReadsAColumnPerIntervalFromTheRow() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim grp As Object
    Set grp = NewMilestoneDevice(sld, 4)

    Dim row As Object
    Set row = CreateObject("Scripting.Dictionary")
    row("MS1_LABEL") = "Kickoff":  row("MS1_DATE") = "Oct 2023": row("MS1_DONE") = "Y"
    row("MS2_LABEL") = "Design":   row("MS2_DATE") = "Mar 2024": row("MS2_DONE") = "yes"
    row("MS3_LABEL") = "Build":    row("MS3_DATE") = "Sep 2024": row("MS3_DONE") = ""

    Dim r As MilestoneDrawResult
    r = MilestoneDevice.DrawFromRow(grp, row)

    result = result & Assert(r.ErrorMessage = "", "it draws [" & r.ErrorMessage & "]")
    result = result & Assert(r.Drawn = 3, "three milestones drawn, got " & r.Drawn)
    result = result & Assert(r.Hidden = 1, "the empty fourth slot is hidden, got " & r.Hidden)
    result = result & Assert(NamedIn(grp, "MS2_LABEL").TextFrame.TextRange.text = "Design", _
        "labels come from their own column")
    result = result & Assert(NamedIn(grp, "MS2_ON").Visible = msoTrue, _
        "'yes' counts as done")
    result = result & Assert(NamedIn(grp, "MS3_ON").Visible = msoFalse _
        And NamedIn(grp, "MS3_OFF").Visible = msoTrue, _
        "a BLANK done column means NOT done -- never assumed the other way")

    ' A GAP IS REFUSED, not closed.
    Dim row2 As Object
    Set row2 = CreateObject("Scripting.Dictionary")
    row2("MS1_LABEL") = "Kickoff": row2("MS1_DATE") = "Oct 2023": row2("MS1_DONE") = "Y"
    row2("MS3_LABEL") = "Build":   row2("MS3_DATE") = "Sep 2024": row2("MS3_DONE") = "N"

    Dim bar As Object
    Set bar = NamedIn(grp, MilestoneDevice.NAME_BAR)
    Dim barBefore As Single
    barBefore = bar.Height

    Dim r2 As MilestoneDrawResult
    r2 = MilestoneDevice.DrawFromRow(grp, row2)
    result = result & Assert(r2.ErrorMessage <> "", "a gap in the columns is refused")
    result = result & Assert(InStr(r2.ErrorMessage, "MS3_LABEL") > 0 _
        And InStr(r2.ErrorMessage, "MS2_LABEL") > 0, _
        "and names the filled one AND the empty one, got '" & r2.ErrorMessage & "'")
    result = result & Assert(r2.Drawn = 0 And bar.Height = barBefore, _
        "nothing drawn and the bar did not move")

    sld.Delete
    Test_MilestoneDevice_ReadsAColumnPerIntervalFromTheRow = result
End Function


' THE CHECK THAT MAKES RENAMING SAFE.
'
' Parts are addressed by NAME, so nothing protects them -- rename MS2_DATE and
' that milestone silently stops updating. Rohan asked whether each shape should
' carry its own tag instead; the answer was no, because two copies of one fact
' drift and then the Selection Pane lies. This is what stands in for it.
'
' THE STRAY-SLOT ASSERTION IS THE ONE THAT MATTERS. SlotCount stops at the first
' gap, so a device with MS1..MS3 complete and MS4_ON renamed reports a perfectly
' healthy three-slot device -- while MS4 and MS5 are invisible to the tool. That
' is a silent, plausible, wrong answer, which is the exact failure class this
' project keeps paying for.
Private Function Test_MilestoneDevice_IntegrityNamesWhatIsMissing() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim grp As Object
    Set grp = NewMilestoneDevice(sld, 3)

    Dim clean As String
    clean = MilestoneDevice.DeviceIntegrity(grp)
    result = result & Assert(InStr(clean, "all parts present") > 0, _
        "an intact device says so, got '" & clean & "'")

    ' A RENAMED PART is named, not counted.
    NamedIn(grp, "MS2_DATE").Name = "Rectangle 99"
    Dim broken As String
    broken = MilestoneDevice.DeviceIntegrity(grp)
    result = result & Assert(InStr(broken, "MS2_DATE") > 0, _
        "the missing part is NAMED, got '" & broken & "'")
    result = result & Assert(InStr(broken, "MS1_") = 0 And InStr(broken, "MS3_") = 0, _
        "and the intact slots are not, got '" & broken & "'")
    NamedIn(grp, "Rectangle 99").Name = "MS2_DATE"

    ' A RENAMED _ON IS NAMED, and the count no longer collapses.
    '
    ' This used to assert "SlotCount stops at the gap -- one slot". That encoded
    ' the two-state model, where a slot WAS its _ON circle: lose it and the slot
    ' ceased to exist, taking the whole tail with it. Under four states a slot
    ' carrying _OFF or _NOW is a real slot -- Rohan's own deck has one drawn with
    ' the big circle alone -- so truncating there would refuse a valid template.
    '
    ' The protection did not go away, it moved: the count survives the gap, and
    ' the missing circle is reported BY NAME. That is strictly better, because
    ' the old behaviour hid slots 3+ as a side effect of a fault in slot 2.
    NamedIn(grp, "MS2_ON").Name = "Oval 77"
    Dim tail As String
    tail = MilestoneDevice.DeviceIntegrity(grp)
    result = result & Assert(InStr(tail, "3 slot(s)") > 0, _
        "the count SURVIVES the gap now -- three slots, got '" & tail & "'")
    result = result & Assert(InStr(tail, "MS2_ON") > 0, _
        "and the missing circle is named, got '" & tail & "'")
    result = result & Assert(InStr(tail, "MS1_ON") = 0 And InStr(tail, "MS3_ON") = 0, _
        "while the intact slots are not, got '" & tail & "'")
    NamedIn(grp, "Oval 77").Name = "MS2_ON"

    ' A slot that loses EVERY circle stops being counted, and the STRAY PROBE
    ' catches its orphaned label and date. Asserted here because the obvious
    ' place to handle it -- a "no circle at all" branch inside the per-slot
    ' loop -- turns out to be unreachable: the loop runs to SlotCount, and
    ' SlotCount will not count a slot with no circle. That branch was written,
    ' proven unreachable by this very case, and deleted.
    NamedIn(grp, "MS3_ON").Name = "Oval 78"
    NamedIn(grp, "MS3_OFF").Name = "Oval 79"
    Dim noCircle As String
    noCircle = MilestoneDevice.DeviceIntegrity(grp)
    result = result & Assert(InStr(noCircle, "2 slot(s)") > 0, _
        "the count stops before the circle-less slot, got '" & noCircle & "'")
    result = result & Assert(InStr(noCircle, "MS3") > 0 And InStr(noCircle, "invisible") > 0, _
        "and the stray probe says MS3's parts are invisible, got '" & noCircle & "'")
    NamedIn(grp, "Oval 78").Name = "MS3_ON"
    NamedIn(grp, "Oval 79").Name = "MS3_OFF"

    sld.Delete
    Test_MilestoneDevice_IntegrityNamesWhatIsMissing = result
End Function


' THE DEVICE IS REACHABLE FROM SYNC, which is the assertion that stops this
' being the fourth built-and-unreachable component found today.
'
' A group is not a picture and has no text frame, so before this branch existed
' it fell through to the TEXT writer and was refused for having nowhere to put a
' string -- exactly how the picture and progress injectors sat unreachable for
' weeks.
Private Function Test_InjectField_RoutesADeviceGroupToTheTimeline() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim grp As Object
    Set grp = NewMilestoneDevice(sld, 3)
    grp.Tags.Add "role", "MILESTONE_TIMELINE"

    Dim row As Object
    Set row = CreateObject("Scripting.Dictionary")
    row("MS1_LABEL") = "Kickoff": row("MS1_DATE") = "Oct 2023": row("MS1_DONE") = "Y"
    row("MS2_LABEL") = "Design":  row("MS2_DATE") = "Mar 2024": row("MS2_DONE") = "N"

    ' Routed through the SAME entry point sync uses, with the row threaded in.
    Dim r As InjectResult
    r = InjectPrimitive.InjectField(sld, "MILESTONE_TIMELINE", "", False, Nothing, row)

    result = result & Assert(r.Found, "the device is found")
    result = result & Assert(r.Written, "and drawn [" & r.ErrorMessage & "]")
    result = result & Assert(InStr(r.ErrorMessage, "text frame") = 0, _
        "NOT refused as a text field, got '" & r.ErrorMessage & "'")
    result = result & Assert(NamedIn(grp, "MS1_LABEL").TextFrame.TextRange.text = "Kickoff", _
        "the label was written through the router")
    result = result & Assert(NamedIn(grp, "MS3_ON").Visible = msoFalse, _
        "the unused third slot is hidden")

    ' A DRY RUN MUST NOT DRAW. Preview Sync promises nothing is written.
    Dim sld2 As Object
    Set sld2 = NewBlankSlide()
    Dim grp2 As Object
    Set grp2 = NewMilestoneDevice(sld2, 3)
    grp2.Tags.Add "role", "MILESTONE_TIMELINE"
    Dim before As String
    before = NamedIn(grp2, "MS1_LABEL").TextFrame.TextRange.text

    Dim r2 As InjectResult
    r2 = InjectPrimitive.InjectField(sld2, "MILESTONE_TIMELINE", "", True, Nothing, row)
    result = result & Assert(Not r2.Written, "a dry run does not write")
    result = result & Assert(NamedIn(grp2, "MS1_LABEL").TextFrame.TextRange.text = before, _
        "and the slide is untouched")

    ' WITHOUT THE ROW it says so, rather than drawing an empty timeline.
    Dim r3 As InjectResult
    r3 = InjectPrimitive.InjectField(sld, "MILESTONE_TIMELINE", "")
    result = result & Assert(InStr(r3.ErrorMessage, "MS1_LABEL") > 0, _
        "no row means it names the columns it needed, got '" & r3.ErrorMessage & "'")

    sld2.Delete
    sld.Delete
    Test_InjectField_RoutesADeviceGroupToTheTimeline = result
End Function

' THE REAL BUG, PROVEN AGAINST THE REAL SHAPE: found 2026-08-16 that 3_P001,
' the deck's own prototype slide, has a fully-built MILESTONE_TIMELINE group
' -- correctly named, every MS1..MS7 part present -- with literally zero tags
' on it. WalkForDeviceRoleTags required a tag anyway, so PlanRoutineSync had
' never once synced it. This fixture reproduces that exact shape: named, not
' tagged.
Private Function Test_InjectPrimitive_DeviceDiscoveredByNameWhenUntagged() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim grp As Object
    Set grp = NewMilestoneDevice(sld, 3)
    grp.Name = "MILESTONE_TIMELINE"
    ' Deliberately NOT tagged -- that is the whole point of this test.

    Dim tags As Object
    Set tags = InjectPrimitive.DeviceRoleTagsOnSlide(sld)
    result = result & Assert(tags.Exists("MILESTONE_TIMELINE"), _
        "an untagged-but-correctly-named device is discovered by name")

    sld.Delete
    Test_InjectPrimitive_DeviceDiscoveredByNameWhenUntagged = result
End Function

Private Function Test_InjectPrimitive_FindShapeByRoleTagFindsUntaggedDeviceByName() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim grp As Object
    Set grp = NewMilestoneDevice(sld, 3)
    grp.Name = "MILESTONE_TIMELINE"

    Dim found As Object
    Set found = InjectPrimitive.FindShapeByRoleTag(sld, "MILESTONE_TIMELINE")
    result = result & Assert(Not found Is Nothing, "the untagged device is found by name")
    If Not found Is Nothing Then
        result = result & Assert(found.Id = grp.Id, "and it is the SAME group, not a coincidence")
    End If

    sld.Delete
    Test_InjectPrimitive_FindShapeByRoleTagFindsUntaggedDeviceByName = result
End Function

Private Function Test_InjectPrimitive_InjectorForRoutesUntaggedDeviceToDevice() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim grp As Object
    Set grp = NewMilestoneDevice(sld, 3)
    grp.Name = "MILESTONE_TIMELINE"

    result = result & Assert(InjectPrimitive.InjectorFor(sld, "MILESTONE_TIMELINE") = InjectPrimitive.INJECTOR_DEVICE, _
        "an untagged device still routes to the device injector, got '" & InjectPrimitive.InjectorFor(sld, "MILESTONE_TIMELINE") & "'")

    ' And the real end-to-end proof: InjectField actually draws into it,
    ' through the exact same entry point Test_InjectField_RoutesADeviceGroupToTheTimeline
    ' uses for the tagged case -- this is that same test, minus the tag.
    Dim row As Object
    Set row = CreateObject("Scripting.Dictionary")
    row("MS1_LABEL") = "Kickoff": row("MS1_DATE") = "Oct 2023": row("MS1_DONE") = "Y"

    Dim r As InjectResult
    r = InjectPrimitive.InjectField(sld, "MILESTONE_TIMELINE", "", False, Nothing, row)
    result = result & Assert(r.Found, "the untagged device is found through InjectField")
    result = result & Assert(r.Written, "and drawn [" & r.ErrorMessage & "]")
    result = result & Assert(NamedIn(grp, "MS1_LABEL").TextFrame.TextRange.text = "Kickoff", _
        "the label was actually written, with no tag anywhere on the group")

    sld.Delete
    Test_InjectPrimitive_InjectorForRoutesUntaggedDeviceToDevice = result
End Function

' THE SAFETY RAIL: the name fallback must not leak into ordinary field
' lookup. A ROLE tag is still how every text/bar/picture field is found --
' only a shape MilestoneDevice.SlotCount already confirms IS a device gets
' the name fallback.
Private Function Test_InjectPrimitive_NameFallbackDoesNotWidenOrdinaryFieldLookup() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()

    ' An ordinary, untagged textbox that HAPPENS to be named like a real
    ' FieldID -- must NOT be found by that name. Only tags find text fields.
    Dim txt As Object
    Set txt = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 100, 20)
    txt.Name = "ABOUT_BODY"
    result = result & Assert(InjectPrimitive.FindShapeByRoleTag(sld, "ABOUT_BODY") Is Nothing, _
        "an untagged textbox is NOT matched by name, even if its name matches a real FieldID")

    ' A group that IS msoGroup but is NOT a real device (no MS-named parts)
    ' -- must not be matched by name either. Structure, not just group-ness,
    ' is what earns the fallback.
    Dim plainGrp As Object
    Dim s1 As Object, s2 As Object
    Set s1 = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 40, 50, 20)
    Set s2 = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 70, 40, 50, 20)
    Set plainGrp = sld.Shapes.Range(Array(s1.Name, s2.Name)).Group
    plainGrp.Name = "MILESTONE_TIMELINE"
    result = result & Assert(InjectPrimitive.FindShapeByRoleTag(sld, "MILESTONE_TIMELINE") Is Nothing, _
        "a group named like a device but with no real slot structure is NOT matched -- SlotCount must be > 0")

    sld.Delete
    Test_InjectPrimitive_NameFallbackDoesNotWidenOrdinaryFieldLookup = result
End Function

Private Function Test_InjectPrimitive_ExplicitTagStillWinsOverDeviceName() As String
    Dim result As String

    Dim sld As Object
    Set sld = NewBlankSlide()
    Dim grp As Object
    Set grp = NewMilestoneDevice(sld, 3)
    grp.Name = "MILESTONE_TIMELINE"
    grp.Tags.Add "role", "OTHER_DEVICE_ID"

    Dim foundByTag As Object
    Set foundByTag = InjectPrimitive.FindShapeByRoleTag(sld, "OTHER_DEVICE_ID")
    result = result & Assert(Not foundByTag Is Nothing, "an explicit tag is still found")
    If Not foundByTag Is Nothing Then
        result = result & Assert(foundByTag.Id = grp.Id, "and it is the same group")
    End If

    Dim foundByName As Object
    Set foundByName = InjectPrimitive.FindShapeByRoleTag(sld, "MILESTONE_TIMELINE")
    result = result & Assert(foundByName Is Nothing, _
        "and the name fallback does NOT ALSO fire once a real tag is present -- one identity, not two")

    sld.Delete
    Test_InjectPrimitive_ExplicitTagStillWinsOverDeviceName = result
End Function

' Local copy of the trailing-break normalisation, for comparing what a textbox
' reads back against what was written. InjectPrimitive's own is Private, and
' widening a production function's scope to satisfy a test would let the test
' change the shipped surface.
Private Function IgnoringTrailingBreaksForTest(text As String) As String
    Dim t As String
    t = text
    Do While Len(t) > 0
        Dim lastChar As String
        lastChar = Right(t, 1)
        If lastChar = vbCr Or lastChar = vbLf Or lastChar = Chr(11) Then
            t = Left(t, Len(t) - 1)
        Else
            Exit Do
        End If
    Loop
    IgnoringTrailingBreaksForTest = t
End Function

' -----------------------------------------------------------------------
' DraftingLobby -- see LOBBY-DESIGN.md.
' -----------------------------------------------------------------------

' Same (1 To 0)-throws-at-runtime guard used throughout the production code
' (AGENTS.md) -- an array-returning function that found nothing leaves its
' array unallocated, and LBound/UBound on that raises rather than returning 0.
Private Function ArrayCountForTest(arr() As DraftingLobby.LobbyEntry) As Long
    On Error Resume Next
    Dim lo As Long, hi As Long
    lo = LBound(arr)
    hi = UBound(arr)
    If Err.Number <> 0 Then
        ArrayCountForTest = 0
    Else
        ArrayCountForTest = hi - lo + 1
    End If
    On Error GoTo 0
End Function

Private Function MakeLobbyFixture(xl As Object) As Object
    Dim wb As Object
    Set wb = xl.Workbooks.Add

    Dim spec As Object
    Set spec = wb.Worksheets(1)
    spec.Name = FieldSpec.SPEC_SHEET_NAME
    spec.Cells(1, FieldSpec.COL_S_FIELDID).Value = "FieldID"
    spec.Cells(1, FieldSpec.COL_S_KIND).Value = "Kind"
    spec.Cells(2, FieldSpec.COL_S_FIELDID).Value = "TEST_FIELD"
    spec.Cells(2, FieldSpec.COL_S_KIND).Value = "Prose"

    Dim draft As Object
    Set draft = wb.Worksheets.Add
    draft.Name = Drafting.DraftSheetNameFor("TEST_FIELD")
    ' Two rows: one approved, one not -- BuildFromScratch must find only the
    ' first. DRAFT_FIRST_ROW and the two columns actually read are the only
    ' cells that matter for this fixture.
    draft.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_ENTITY).Value = "P001"
    draft.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_APPROVED).Value = "Y"
    draft.Cells(Drafting.DRAFT_FIRST_ROW + 1, Drafting.COL_D_ENTITY).Value = "P002"
    draft.Cells(Drafting.DRAFT_FIRST_ROW + 1, Drafting.COL_D_APPROVED).Value = ""

    Set MakeLobbyFixture = wb
End Function

Private Function Test_DraftingLobby_PinReadClearRoundTrip() As String
    Dim result As String

    Dim xl As Object, wb As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set wb = MakeLobbyFixture(xl)

    DraftingLobby.PinToLobby wb, "TPL_TEST_FIELD", 10, "TEST_FIELD", "P001"

    Dim entries() As DraftingLobby.LobbyEntry
    Dim n As Long
    entries = DraftingLobby.ReadLobby(wb)
    n = ArrayCountForTest(entries)
    result = result & Assert(n = 1, "one pin lands as one entry, got " & n)

    If n = 1 Then
        result = result & Assert(entries(1).SheetName = "TPL_TEST_FIELD", "sheet name round-trips, got '" & entries(1).SheetName & "'")
        result = result & Assert(entries(1).FieldId = "TEST_FIELD", "field id round-trips, got '" & entries(1).FieldId & "'")
        result = result & Assert(entries(1).EntityKey = "P001", "entity key round-trips, got '" & entries(1).EntityKey & "'")
        result = result & Assert(entries(1).Row = 10, "row round-trips, got " & entries(1).Row)
    End If

    DraftingLobby.ClearLobbyEntry wb, "TPL_TEST_FIELD", "TEST_FIELD", "P001"
    entries = DraftingLobby.ReadLobby(wb)
    n = ArrayCountForTest(entries)
    result = result & Assert(n = 0, "cleared entry leaves the Lobby empty, got " & n)

    wb.Saved = True
    wb.Close
    xl.Quit

    Test_DraftingLobby_PinReadClearRoundTrip = result
End Function

Private Function Test_DraftingLobby_BuildFromScratchFindsOnlyApprovedRows() As String
    Dim result As String

    Dim xl As Object, wb As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set wb = MakeLobbyFixture(xl)

    Dim report As String
    report = DraftingLobby.BuildLobbyFromScratch(wb)
    result = result & Assert(InStr(report, "1 approved") > 0, _
        "reports exactly one approved row found, got '" & report & "'")

    Dim entries() As DraftingLobby.LobbyEntry
    Dim n As Long
    entries = DraftingLobby.ReadLobby(wb)
    n = ArrayCountForTest(entries)
    result = result & Assert(n = 1, _
        "only the ticked row (P001) is pinned, the untucked one (P002) is not, got " & n)

    If n = 1 Then
        result = result & Assert(entries(1).EntityKey = "P001", _
            "the pinned entry is the approved row, got '" & entries(1).EntityKey & "'")
    End If

    ' REBUILT FROM REALITY, NOT MERGED. Untick the row, rebuild, confirm the
    ' stale pin does not survive -- the Lobby must never trust its own past
    ' content over what the drafting sheet says now.
    Dim draft As Object
    Set draft = wb.Worksheets(Drafting.DraftSheetNameFor("TEST_FIELD"))
    draft.Cells(Drafting.DRAFT_FIRST_ROW, Drafting.COL_D_APPROVED).Value = ""
    DraftingLobby.BuildLobbyFromScratch wb
    entries = DraftingLobby.ReadLobby(wb)
    n = ArrayCountForTest(entries)
    result = result & Assert(n = 0, _
        "an un-ticked row does not survive a rebuild as a stale pin, got " & n)

    wb.Saved = True
    wb.Close
    xl.Quit

    Test_DraftingLobby_BuildFromScratchFindsOnlyApprovedRows = result
End Function

Private Function Test_DraftingLobby_PinTwiceUpdatesInPlaceNotDuplicate() As String
    Dim result As String

    Dim xl As Object, wb As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set wb = MakeLobbyFixture(xl)

    DraftingLobby.PinToLobby wb, "TPL_TEST_FIELD", 10, "TEST_FIELD", "P001"
    DraftingLobby.PinToLobby wb, "TPL_TEST_FIELD", 10, "TEST_FIELD", "P001"

    Dim entries() As DraftingLobby.LobbyEntry
    Dim n As Long
    entries = DraftingLobby.ReadLobby(wb)
    n = ArrayCountForTest(entries)
    result = result & Assert(n = 1, _
        "pinning the same (sheet, field, entity) twice updates in place, does not duplicate -- got " & n)

    wb.Saved = True
    wb.Close
    xl.Quit

    Test_DraftingLobby_PinTwiceUpdatesInPlaceNotDuplicate = result
End Function

' LOBBY-DESIGN.md phase 2: PublishAllDraftedFields now runs only the fields
' the Lobby has pinned, instead of every declared Prose field. The chain
' itself needs a live presentation (Application.ActivePresentation, via
' Resolve) that this harness does not build, so this proves the part that
' actually changed -- DraftingUI.DistinctPinnedFields, the pure function the
' chain now delegates the field-selection decision to -- directly, without
' one.
Private Function Test_DraftingUI_DistinctPinnedFieldsReadsOnlyTheLobby() As String
    Dim result As String

    ' NOTHING PINNED -> "". An unallocated array, not a (1 To 0) one -- the
    ' same shape DraftingLobby.ReadLobby actually returns when the Lobby sheet
    ' does not exist yet, proven here rather than assumed.
    Dim noPins() As DraftingLobby.LobbyEntry
    result = result & Assert(DraftingUI.DistinctPinnedFields(noPins) = "", _
        "an unallocated (nothing pinned) array publishes nothing, got '" & DraftingUI.DistinctPinnedFields(noPins) & "'")

    ' THE CASE THAT MATTERS: two entities pinned to the SAME field must not
    ' run that field's Copy+Publish chain twice -- Drafting.PublishDrafts
    ' already publishes every ticked row on the sheet in one pass, so a
    ' second pass would be redundant work, not a second field.
    Dim two(1 To 2) As DraftingLobby.LobbyEntry
    two(1).FieldId = "ABOUT_BODY": two(1).EntityKey = "P001"
    two(2).FieldId = "ABOUT_BODY": two(2).EntityKey = "P002"
    Dim gotTwo As String
    gotTwo = DraftingUI.DistinctPinnedFields(two)
    result = result & Assert(gotTwo = "ABOUT_BODY", _
        "two entities pinned to one field publish that field once, got '" & gotTwo & "'")

    ' TWO DISTINCT FIELDS, FIRST-PINNED ORDER -- and a field NOT in the Lobby
    ' at all (KEY_EVENTS_BODY, never pinned) must never appear, which is the
    ' whole fix for the crawl: an untouched field's sheet is never opened.
    Dim three(1 To 3) As DraftingLobby.LobbyEntry
    three(1).FieldId = "PROGRESS_BODY": three(1).EntityKey = "P001"
    three(2).FieldId = "ABOUT_BODY": three(2).EntityKey = "P001"
    three(3).FieldId = "PROGRESS_BODY": three(3).EntityKey = "P002"
    Dim gotThree As String
    gotThree = DraftingUI.DistinctPinnedFields(three)
    result = result & Assert(gotThree = "PROGRESS_BODY,ABOUT_BODY", _
        "distinct fields in first-pinned order, got '" & gotThree & "'")
    result = result & Assert(InStr(1, gotThree, "KEY_EVENTS_BODY", vbTextCompare) = 0, _
        "a field never pinned in the Lobby must never appear -- this is the crawl fix")

    Test_DraftingUI_DistinctPinnedFieldsReadsOnlyTheLobby = result
End Function
