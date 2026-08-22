Attribute VB_Name = "DeckAdoption"
Option Explicit

' Implements specs/deck-adoption.md: bulk retroactive linking of a deck's
' already-populated, untagged slides -- the one-time bootstrap that captures
' N real existing records into the Data sheet and tags them, before ongoing
' sync starts governing them. Engine layer only, per the spec's own
' Non-goals ("the UI itself... a future pass") -- no ribbon/form here, same
' boundary every other module drew before its own ribbon wiring landed
' (specs/ribbon-ui.md).
'
' Deliberately reuses every existing primitive unchanged, per the spec's own
' framing: Onboarding.BuildTemplateFieldShapes/MatchSlideAgainstTemplate for
' scoring (no new matching logic), Onboarding.OnboardNewInstance for tagging
' (it already does exactly "tag slide identity + auto-accept every
' high-confidence field," which is precisely what a "ready" slide here
' needs), ExcelOutput.UpsertRow for the Data-sheet write, and
' InjectPrimitive.InjectPrimitive for the verify-the-link round trip. The
' one genuinely new piece is instance-key resolution against pre-existing
' keyless Data-sheet rows (ReadKeylessRows/FindMatchingKeylessRow below) --
' the spec calls this out explicitly as new decision logic, not present
' anywhere else in this project.
'
' See SPIKE_NOTES_DeckAdoption.md for the full divergence list, the
' plan/commit phase-gate design, and the manual verification recipe.

' ---------------------------------------------------------------------
' Types
' ---------------------------------------------------------------------

' One entry per slide in scope (excluding the template). Deliberately
' scalar-only (no nested array/UDT members) -- this project has never
' stored a dynamic array of one UDT type as a member of another UDT that
' itself lives in a dynamic array, and this module doesn't originate that
' risk; every other array-of-UDT type here (SyncAction, FieldMatch,
' MatchResult) is the same shape: scalars plus at most a single Object
' member. See SPIKE_NOTES_DeckAdoption.md.
Public Type AdoptionSlidePlan
    SlideId As Long             ' Slide.SlideID -- stable identity for reporting/debugging
    SlideLabel As String        ' "Slide N (Name)" -- human-readable, for the phase-gate review and final report
    Disposition As String       ' "already_linked" | "ready" | "needs_confirmation" | "unclassified"
    MatchedKeylessRowId As String ' "" unless Disposition="ready" and a pre-existing keyless Data-sheet row matched verbatim; the raw row number (as a string) to link into rather than append fresh
    Reason As String
End Type

' Batch commit report, per specs/deck-adoption.md's "same shape as
' ribbon-ui.md's shared result form" requirement: counts plus per-slide
' labels in each bucket, so a caller (the eventual UI, or a test) can locate
' every flagged item by slide name/index. Multiple String() members on a
' single (non-array-of) UDT instance mirrors SlideDuplication.DuplicateResult's
' already-proven MissingFields() pattern -- not the untested nested-array
' construct AdoptionSlidePlan's own comment avoids.
Public Type AdoptionResult
    LinkedCount As Long
    LinkedLabels() As String
    AlreadyLinkedCount As Long
    AlreadyLinkedLabels() As String
    ExcludedUnclassifiedCount As Long
    ExcludedUnclassifiedLabels() As String
    FailedVerificationCount As Long
    FailedVerificationLabels() As String
End Type

' ---------------------------------------------------------------------
' Plan (the phase gate) -- nothing is written by this function.
' ---------------------------------------------------------------------

' `slidesToAdopt` is every slide in the user's selection except whichever
' one (if any) was picked as the template -- gathering/excluding is the
' caller's job, per this project's standing convention (RunSync.
' GatherInstances is the one exception, where the gather target is a live
' object-model scan of the whole deck by tag; here the "gather" is a raw UI
' selection, deliberately left to the caller so this function stays
' headlessly testable). MUST be supplied in current deck order (SlideIndex
' ascending) -- this function does not sort, and relies on that order for
' "row order bootstraps from deck order" when appending fresh rows in
' CommitAdoption. `templateSld` must already be onboarded (tagged fields) --
' the greenfield "pick one slide, run it through onboard-slide-type.md"
' path is explicitly this spec's own Non-goal territory for this module.
'
' `harvestedValues` is an out-parameter, parallel to `slidesToAdopt` (same
' index) -- a Scripting.Dictionary (fieldName -> harvested value) for every
' "ready" slide, unset (Nothing) otherwise. Parallel arrays, not a
' Dictionary keyed by slide, for the same reason every other module here
' uses them: a UDT (or a live Slide Object used as a Dictionary key) has no
' clean Dictionary-key story, and Object arrays indexed in lockstep are
' this project's established idiom (Discovery.DiscoverSlideWithShapes,
' Onboarding.MatchSlideAgainstTemplate's own untaggedShapes()).
'
' The returned AdoptionSlidePlan() array shares `slidesToAdopt`'s exact
' index range (LBound/UBound), not forced to 1-based -- CommitAdoption's own
' `confirmedInstanceKeys` argument must be built by the caller over that
' same range, since plans(i)/slidesToAdopt(i)/harvestedValues(i)/
' confirmedInstanceKeys(i) are all read with one shared `i`.
Public Function PlanAdoption(slidesToAdopt() As Object, templateSld As Object, ws As Object, ByRef harvestedValues() As Object) As AdoptionSlidePlan()
    Dim templateRoles() As String
    Dim templateFieldShapes() As Candidate
    templateFieldShapes = Onboarding.BuildTemplateFieldShapes(templateSld, templateRoles)

    Dim sheet As Sheet
    sheet = ExcelOutput.ReadSheet(ws)

    Dim keylessRows As Object
    Set keylessRows = ReadKeylessRows(ws, sheet.Fields)
    Dim usedKeylessRowIds As Object
    Set usedKeylessRowIds = CreateObject("Scripting.Dictionary")

    Dim plans() As AdoptionSlidePlan

    Dim lo As Long, hi As Long, hasSlides As Boolean
    On Error Resume Next
    lo = LBound(slidesToAdopt)
    hi = UBound(slidesToAdopt)
    hasSlides = (Err.Number = 0)
    On Error GoTo 0

    If Not hasSlides Then
        PlanAdoption = plans ' genuinely unallocated -- nothing in scope; see AGENTS.md's (1 To 0) restriction
        Exit Function
    End If

    ' `plans`/`harvestedValues` are allocated over the SAME index range as
    ' `slidesToAdopt` (lo To hi), not forced to 1-based -- CommitAdoption
    ' indexes plans(i)/slidesToAdopt(i)/harvestedValues(i)/
    ' confirmedInstanceKeys(i) with one shared `i`, so every one of those
    ' four arrays must share this exact index space or writes get silently
    ' misattributed to the wrong slide. An earlier version of this function
    ' used a separate 1-based counter for `plans` (via ReDim Preserve on
    ' each append) while `harvestedValues` kept `slidesToAdopt`'s own bounds
    ' -- harmless only because every test built a 1-based slidesToAdopt();
    ' found and fixed during review (2026-07-25) before any caller could
    ' hit it with a non-1-based array. See SPIKE_NOTES_DeckAdoption.md.
    ReDim plans(lo To hi)
    ReDim harvestedValues(lo To hi)

    Dim i As Long
    For i = lo To hi
        Dim sld As Object
        Set sld = slidesToAdopt(i)

        plans(i).SlideId = sld.SlideID
        plans(i).SlideLabel = "Slide " & sld.SlideIndex & " (" & sld.Name & ")"

        ' Idempotent skip: a slide already carrying instance_key + type tag
        ' is already linked -- touch nothing, report only. Reruns (adopting
        ' slides missed on a first pass) must be safe.
        Dim existing As SlideInstance
        existing = Resolve.ResolveSlideInstance(sld)
        If existing.HasInstanceKey And existing.HasTypeTag Then
            plans(i).Disposition = "already_linked"
            plans(i).Reason = "already carries slide_type/instance_key tags -- skipped, not re-adopted"
            GoTo ContinueLoop
        End If

        ' Match against the template using matching.md's existing tier-2
        ' scoring, unchanged -- no new matching logic, applied per-slide
        ' across this batch instead of to a single incoming slide.
        Dim untaggedShapes() As Object
        Dim matches() As FieldMatch
        matches = Onboarding.MatchSlideAgainstTemplate(sld, templateRoles, templateFieldShapes, untaggedShapes)

        Dim mLo As Long, mHi As Long, hasMatches As Boolean
        On Error Resume Next
        mLo = LBound(matches)
        mHi = UBound(matches)
        hasMatches = (Err.Number = 0)
        On Error GoTo 0

        ' Confidence dispatch, aggregated to one disposition per slide (the
        ' spec's own phase-gate review is per-slide, not per-field) -- a
        ' judgment call this spec doesn't fully pin down on its own, same
        ' posture as this project's other documented design choices (e.g.
        ' RunSync's resequencing anchor). "ready" requires EVERY template
        ' role to score high on this slide; any role scoring medium (but at
        ' least one role matching at all) is "needs_confirmation"; zero
        ' roles matching anything is "unclassified". See
        ' SPIKE_NOTES_DeckAdoption.md.
        Dim allHigh As Boolean, anyMatch As Boolean
        allHigh = hasMatches
        anyMatch = False

        Dim harvested As Object
        Set harvested = CreateObject("Scripting.Dictionary")

        If hasMatches Then
            Dim j As Long
            For j = mLo To mHi
                If matches(j).Result.Confidence = "high" And matches(j).Result.HasCandidate Then
                    anyMatch = True
                    If matches(j).Role = SyncOperations.SUBTITLE_COMPOSITE_FIELD Then
                        ' SUBTITLE_A's shape displays a middot-joined composite
                        ' of four register columns (SyncOperations.
                        ' ComposeSubtitleLine), not its own raw value --
                        ' harvesting the rendered text here would write the
                        ' whole composite into this one column, corrupting it
                        ' for the next real sync. Same refusal Harvest.bas and
                        ' E2EField.bas already apply for the same reason; a
                        ' third live call site with the identical blind spot,
                        ' found by mother-hound's 2026-08-21 kennel run. Still
                        ' counts as a confident structural match (the shape IS
                        ' in the right place) -- only the value harvest is
                        ' unsafe, so the field is simply left unwritten rather
                        ' than the slide downgraded to unclassified.
                    Else
                        ' Harvest current values verbatim from every matched
                        ' field -- the same technique InjectPrimitive.bas already
                        ' uses to read a shape's current value.
                        '
                        ' FIX-LIST item, 2026-08-22 (mother-hound kennel
                        ' survey, finding #5). This was an unguarded
                        ' .TextFrame.TextRange.Text read -- a picture-typed
                        ' shape has no TextFrame at all, so a high-confidence
                        ' picture-field match raised a hard VBA runtime error
                        ' here during bulk onboarding harvest, where the
                        ' sibling read two lines away in InjectPrimitive.bas
                        ' already guards this exact case. A picture shape's
                        ' "current value" is its picsrc TAG stamp, the same
                        ' thing VerifyLink and a real sync both compare
                        ' against -- not its (nonexistent) text.
                        Dim matchedShp As Object
                        Set matchedShp = untaggedShapes(matches(j).Result.CandidateIndex)
                        If InjectPrimitive.IsPictureShape(matchedShp) Then
                            harvested(matches(j).Role) = InjectPrimitive.PictureSourceOf(matchedShp)
                        Else
                            harvested(matches(j).Role) = matchedShp.TextFrame.TextRange.Text
                        End If
                    End If
                Else
                    allHigh = False
                    If matches(j).Result.Confidence = "medium" Then anyMatch = True
                End If
            Next j
        End If

        If Not anyMatch Then
            plans(i).Disposition = "unclassified"
            plans(i).Reason = "no template field scored medium or high confidence on this slide -- excluded, never forced in"
        ElseIf allHigh Then
            plans(i).Disposition = "ready"
        Else
            plans(i).Disposition = "needs_confirmation"
            plans(i).Reason = "at least one template field needs a human-confirmed match (Onboarding.ConfirmFieldMatch) before this slide can be adopted"
        End If

        If plans(i).Disposition = "ready" Then
            Set harvestedValues(i) = harvested

            ' Instance-key resolution against existing keyless Data-sheet
            ' rows -- the one piece of genuinely new decision logic this
            ' spec introduces. Zero-or-multiple verbatim matches always
            ' falls back to a fresh row; never guessed.
            Dim matchedRowId As String
            matchedRowId = FindMatchingKeylessRow(keylessRows, harvested, usedKeylessRowIds)
            If matchedRowId <> "" Then
                plans(i).MatchedKeylessRowId = matchedRowId
                usedKeylessRowIds(matchedRowId) = True
                plans(i).Reason = "matches existing keyless Data-sheet row " & matchedRowId & " verbatim -- will link to it rather than creating a fresh row (instance_key still required from the human)"
            Else
                plans(i).Reason = "no existing keyless Data-sheet row matches verbatim -- will create a fresh row (instance_key required from the human)"
            End If
        End If

ContinueLoop:
    Next i

    PlanAdoption = plans
End Function

' ---------------------------------------------------------------------
' Commit -- the only function in this module that writes anything.
' ---------------------------------------------------------------------

' Commits every "ready" plan entry whose slide the human confirmed an
' instance_key for at the phase gate (`confirmedInstanceKeys`, parallel to
' `plans`/`slidesToAdopt`/`harvestedValues` -- "" means not yet confirmed, so
' that slide is skipped this pass rather than guessed). All four arrays are
' indexed with one shared loop variable, so they must share the exact same
' index range -- `plans`/`harvestedValues` already do, since PlanAdoption
' allocates both over `slidesToAdopt`'s own LBound/UBound (not forced to
' 1-based); `confirmedInstanceKeys` must be built by the caller over that
' same range. "already_linked" is reported,
' untouched. "needs_confirmation"/"unclassified" slides are NOT partially
' committed here -- per the spec, a needs_confirmation slide must first be
' resolved via the existing Onboarding.ConfirmFieldMatch (a human picking
' the ambiguous shape) and then re-planned in a later pass, exactly the same
' "flag, never force, never partially commit" posture matching.md itself
' requires; this function does not invent a partial-field-write path.
'
' Reuses Onboarding.OnboardNewInstance directly for the actual tag-write --
' it already does exactly what a "ready" slide needs (tag slide-level
' identity unconditionally, auto-accept every high-confidence field match),
' since PlanAdoption already confirmed every field is high-confidence for a
' "ready" slide. No new tagging logic here.
Public Function CommitAdoption(plans() As AdoptionSlidePlan, slidesToAdopt() As Object, harvestedValues() As Object, confirmedInstanceKeys() As String, slideType As String, templateSld As Object, ws As Object) As AdoptionResult
    Dim result As AdoptionResult

    ' Resolved once here, not threaded through the caller's signature --
    ' same reasoning DeckRegistry.GetDeckPeriod's own comment gives a few
    ' lines below: ws.Parent is the one workbook this whole commit is
    ' already operating against, so it cannot disagree with itself.
    ' FIX-LIST item, 2026-08-22: needed so VerifyLink's picture-field
    ' check (routed through InjectField below) can resolve a picsrc stamp
    ' the same way every other real write path already does.
    Dim srcWs As Object
    Set srcWs = Nothing
    If WorkbookBridge.WorksheetExists(ws.Parent, Sources.SOURCES_SHEET_NAME) Then
        Set srcWs = WorkbookBridge.GetOrAddWorksheet(ws.Parent, Sources.SOURCES_SHEET_NAME)
    End If

    Dim templateRoles() As String
    Dim templateFieldShapes() As Candidate
    templateFieldShapes = Onboarding.BuildTemplateFieldShapes(templateSld, templateRoles)

    Dim lo As Long, hi As Long, hasPlans As Boolean
    On Error Resume Next
    lo = LBound(plans)
    hi = UBound(plans)
    hasPlans = (Err.Number = 0)
    On Error GoTo 0

    If Not hasPlans Then
        CommitAdoption = result
        Exit Function
    End If

    Dim i As Long
    For i = lo To hi
        Select Case plans(i).Disposition
            Case "already_linked"
                result.AlreadyLinkedCount = result.AlreadyLinkedCount + 1
                ReDim Preserve result.AlreadyLinkedLabels(1 To result.AlreadyLinkedCount)
                result.AlreadyLinkedLabels(result.AlreadyLinkedCount) = plans(i).SlideLabel

            Case "ready"
                If confirmedInstanceKeys(i) = "" Then
                    result.ExcludedUnclassifiedCount = result.ExcludedUnclassifiedCount + 1
                    ReDim Preserve result.ExcludedUnclassifiedLabels(1 To result.ExcludedUnclassifiedCount)
                    result.ExcludedUnclassifiedLabels(result.ExcludedUnclassifiedCount) = plans(i).SlideLabel & " (instance_key not confirmed -- skipped, never guessed)"
                Else
                    Dim sld As Object
                    Set sld = slidesToAdopt(i)
                    Dim instanceKey As String
                    instanceKey = confirmedInstanceKeys(i)

                    Onboarding.OnboardNewInstance sld, templateRoles, templateFieldShapes, slideType, instanceKey

                    ' Row order bootstraps from deck order: a genuinely
                    ' fresh row is appended by UpsertRow in call order, so
                    ' the caller supplying `slidesToAdopt` in deck order
                    ' (this function's own documented precondition) is what
                    ' makes that hold -- this function itself does not sort.
                    If plans(i).MatchedKeylessRowId <> "" Then
                        ws.Cells(CLng(plans(i).MatchedKeylessRowId), 1).Value = instanceKey
                    End If
                    ' Period from the deck the template lives in (Slide.Parent
                    ' is that Presentation -- probed against real PowerPoint
                    ' 2026-08-04, and asserted by a test so it stays true),
                    ' rather than threaded through this function's signature.
                    ' The deck cannot disagree with itself; a parameter can.
                    ExcelOutput.UpsertRow ws, instanceKey, harvestedValues(i), _
                        DeckRegistry.GetDeckPeriod(templateSld.Parent)

                    ' Verify the link, not just the write -- the same
                    ' inject_primitive no-op round trip onboard-slide-
                    ' type.md's Step 6 already mandates for the template,
                    ' generalized across this batch.
                    If VerifyLink(sld, harvestedValues(i), srcWs) Then
                        result.LinkedCount = result.LinkedCount + 1
                        ReDim Preserve result.LinkedLabels(1 To result.LinkedCount)
                        result.LinkedLabels(result.LinkedCount) = plans(i).SlideLabel
                    Else
                        ' A harvest bug in THIS pass, not something a later
                        ' sync should silently "correct" -- flagged, not
                        ' rolled back (no undo primitive exists anywhere in
                        ' this project; see SPIKE_NOTES_DeckAdoption.md).
                        result.FailedVerificationCount = result.FailedVerificationCount + 1
                        ReDim Preserve result.FailedVerificationLabels(1 To result.FailedVerificationCount)
                        result.FailedVerificationLabels(result.FailedVerificationCount) = plans(i).SlideLabel
                    End If
                End If

            Case Else ' "needs_confirmation" or "unclassified"
                result.ExcludedUnclassifiedCount = result.ExcludedUnclassifiedCount + 1
                ReDim Preserve result.ExcludedUnclassifiedLabels(1 To result.ExcludedUnclassifiedCount)
                result.ExcludedUnclassifiedLabels(result.ExcludedUnclassifiedCount) = plans(i).SlideLabel & " (" & plans(i).Disposition & ")"
        End Select
    Next i

    CommitAdoption = result
End Function

' FIX-LIST item, 2026-08-22 (mother-hound kennel survey, finding #4). This
' called InjectPrimitive.InjectPrimitive (the plain TEXT writer) directly --
' CV/CW's own router-bypass class, opposite failure direction. For a
' picture-typed field, InjectPrimitive returns Found=True, Written=False,
' Verified=False (no text frame to write into), which unconditionally
' failed VerifyLink's `Not r.Found Or r.Written Or Not r.Verified` gate --
' so a correctly-linked picture field was ALWAYS reported as a failed
' link, never a corrupted one. Now routes through InjectField, the same
' type-aware router CV threaded into slide creation; srcWs resolves a
' picture field's picsrc stamp comparison the same way it does everywhere
' else.
Private Function VerifyLink(sld As Object, harvested As Object, Optional srcWs As Object = Nothing) As Boolean
    Dim fieldName As Variant
    For Each fieldName In harvested.Keys
        Dim r As InjectResult
        r = InjectPrimitive.InjectField(sld, CStr(fieldName), CStr(harvested(fieldName)), False, srcWs)
        If Not r.Found Or r.Written Or Not r.Verified Then
            VerifyLink = False
            Exit Function
        End If
    Next fieldName
    VerifyLink = True
End Function

' ---------------------------------------------------------------------
' Keyless-row resolution -- the genuinely new decision logic this spec
' introduces (not present in onboarding.md or sync-operations.md, which
' never handle a Data-sheet row that predates any linked slide).
' ---------------------------------------------------------------------

' ExcelOutput.ReadSheet deliberately excludes rows with a blank Instance ID
' cell from Sheet.Rows/InstanceOrder (its own read loop only includes
' `If instanceId <> ""`) -- correct for every existing caller (RunSync et
' al. never care about keyless rows), but it leaves exactly the gap this
' function closes: reading rows that have real data but no key yet (e.g. a
' user hand-typed rows ahead of running this). Returns a Scripting.
' Dictionary: CStr(row number) -> Scripting.Dictionary(fieldName -> value),
' one entry per non-empty keyless row.
Private Function ReadKeylessRows(ws As Object, fields As Collection) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim lastCol As Long
    lastCol = fields.count + 1 ' column A (Instance ID) + one column per known field

    Dim ur As Object
    Set ur = ws.UsedRange
    Dim lastRow As Long
    lastRow = ur.Row + ur.Rows.count - 1

    Dim r As Long
    For r = 2 To lastRow
        If IsEmpty(ws.Cells(r, 1).Value) Or CStr(ws.Cells(r, 1).Value) = "" Then
            Dim rowValues As Object
            Set rowValues = CreateObject("Scripting.Dictionary")
            Dim hasAnyValue As Boolean
            hasAnyValue = False

            Dim c As Long
            For c = 2 To lastCol
                If Not IsEmpty(ws.Cells(r, c).Value) Then
                    rowValues(CStr(fields(c - 1))) = CStr(ws.Cells(r, c).Value)
                    hasAnyValue = True
                End If
            Next c

            If hasAnyValue Then
                Set result(CStr(r)) = rowValues
            End If
        End If
    Next r

    Set ReadKeylessRows = result
End Function

' A slide matches a keyless row only if EVERY one of that row's own
' populated fields equals the slide's harvested value for that same field
' exactly (a field the row never had is not evidence either way; a field
' the slide harvested but the row lacks is not evidence either way -- only
' fields the row actually carries are the "non-key fields" the spec means).
' Zero or more than one such match always falls back to a fresh row rather
' than guessing which to merge into -- `usedRowIds` prevents two different
' slides in the same batch from both claiming the same keyless row.
Private Function FindMatchingKeylessRow(keylessRows As Object, harvested As Object, usedRowIds As Object) As String
    Dim matchId As String
    Dim matchCount As Long
    matchCount = 0

    Dim rowId As Variant
    For Each rowId In keylessRows.Keys
        If Not usedRowIds.Exists(CStr(rowId)) Then
            If RowValuesMatchHarvested(keylessRows(rowId), harvested) Then
                matchCount = matchCount + 1
                matchId = CStr(rowId)
            End If
        End If
    Next rowId

    If matchCount = 1 Then
        FindMatchingKeylessRow = matchId
    Else
        FindMatchingKeylessRow = "" ' zero or ambiguous -- never guess
    End If
End Function

Private Function RowValuesMatchHarvested(rowValues As Object, harvested As Object) As Boolean
    Dim fieldName As Variant
    For Each fieldName In rowValues.Keys
        If Not harvested.Exists(fieldName) Then
            RowValuesMatchHarvested = False
            Exit Function
        End If
        If CStr(harvested(fieldName)) <> CStr(rowValues(fieldName)) Then
            RowValuesMatchHarvested = False
            Exit Function
        End If
    Next fieldName
    RowValuesMatchHarvested = True
End Function
