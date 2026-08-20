Attribute VB_Name = "SyncOperations"
Option Explicit

' VBA port of src/sync_operations.py's plan_routine_sync()/
' plan_period_rollover(), per specs/vba-port.md's port order (module 4,
' together with Resolve.bas). Dispatches per specs/sync-operations.md:
' cases 1 (no_change), 3 (new_record), and 4 (in_place_correction) from
' PlanRoutineSync; case 6 (unclassified_slide) is folded into the same
' function, matching the Python original. Cases 5/7 are non-goals per
' specs/sync-operations.md and are not produced anywhere in this file.
'
' CASE 2 (period_rollover) NO LONGER EXISTS HERE. It duplicated a slide
' inside the deck to represent the next period -- the deck-accumulates
' model, rejected 2026-08-03. A period now gets its own deck file and last
' period's stays as the record, so rolling forward copies ROWS
' (ExcelOutput.RollForwardPeriod), never slides. Removed 2026-08-05.
'
' See SPIKE_NOTES_Resolve.md for the full divergence list -- most notably,
' this port skips resolve.py's separate field_shapes pre-resolution step
' and calls InjectPrimitive.InjectPrimitive() directly per field, relying
' on its own native tag lookup as the single source of truth for "does
' this instance have a shape for this field."

Public Type SyncAction
    Kind As String              ' "no_change" | "in_place_correction" | "new_record" | "flagged"
    InstanceKey As String       ' no_change / in_place_correction
    ChangedFieldVerified As Object ' in_place_correction: Scripting.Dictionary fieldName -> Boolean
    ChangedFieldError As Object    ' in_place_correction: Scripting.Dictionary fieldName -> String (ErrorMessage, "" if none)
    ChangedFieldCurrent As Object  ' in_place_correction: Scripting.Dictionary fieldName -> String, the slide's text BEFORE the write (populated in both modes; the whole point of a dry run is being able to show before/after)
    ChangedFieldNew As Object      ' in_place_correction: Scripting.Dictionary fieldName -> String, the Data-sheet value that WOULD be / was written. The "after" half of the line above. Added 2026-07-30 after the first live preview reported "now: 'Project Closed'" and never said what it would become -- leaving the human to go and read Excel, which is the errand a preview exists to save. The intent was already documented here; only the implementation was missing.
    ' NOTE: was a single "ChangedFields fieldName -> InjectResult" Dictionary until
    ' 2026-07-25 -- a UDT cannot be assigned to a Variant in VBA (compile-time "Invalid
    ' use of type"), so a Dictionary can never hold an InjectResult value directly. Split
    ' into two Dictionaries of primitives instead (Found/Written are always True for any
    ' field that reaches `changed` in the first place, per PlanRoutineSync's own gate
    ' below, so only Verified/ErrorMessage are worth carrying per field). See
    ' SPIKE_NOTES_Onboarding.md, which is where this was first found (in Onboarding.bas's
    ' own design), and AGENTS.md's Known Patterns.
    RowInstanceKey As String    ' new_record: the Data-sheet row's instance key
    Values As Object            ' new_record: Scripting.Dictionary fieldName -> String value
    Subject As String           ' flagged: a readable handle on the flagged slide (its SlideID)
    FlagKind As String          ' flagged: always "unclassified_slide" here (cases 5/7 are non-goals)
    Reason As String
End Type


' Dispatch cases 1/3/4/6 across `instances` (live Slide objects already
' believed to belong to one type -- gathering that set is the caller's job,
' matching sync_operations.py's own non-goal of not discovering instances
' itself) against a Data-sheet already read into memory:
'   dataRows     - Scripting.Dictionary: instance id (String) -> Scripting.
'                  Dictionary of fieldName (String) -> value (String)
'   instanceOrder - Collection of instance-id strings, in the Data-sheet's
'                  row order
' `instances` must be a 1-based array when non-empty; a genuinely
' unallocated array (never ReDim'd) represents "no items" -- see this
' function's own comment on why (1 To 0) can't be used for that instead.
'
' Module 6 (Excel-side reads) is not yet ported (per specs/vba-port.md's
' port order, it's step 6, after onboarding) -- this function accepts the
' shape that reader will eventually produce rather than reading a worksheet
' itself, mirroring sync_operations.py's own separation: it never touches a
' file directly either, excel_output.py reads the Sheet it operates on.
' `dryRun` is threaded straight through to InjectPrimitive, which is the only
' thing here that mutates a slide. With it set, this function is a pure read:
' it classifies every instance exactly as a real run would, and reports the
' in-place corrections it WOULD make (with each field's current value), without
' making any of them. Note that "planning" writing at all is the surprising
' part -- PlanRoutineSync calls InjectPrimitive directly rather than returning
' actions for a caller to execute, so before this flag existed there was no way
' to look at a routine sync without performing it.
' -----------------------------------------------------------------------
' THE ELAPSED-TIME BAR -- Kind = Derived, per ExcelOutput.KIND_DERIVED.
' -----------------------------------------------------------------------
'
' Rohan, 2026-08-09: "timeline will move towards end but not necessarily
' every quarter, vs time elapsed bar autoshapes that move with the clock
' regardless of progress." A pure function of the date, computed fresh
' every run from START_DATE/END_DATE -- never stored, per KIND_DERIVED's
' own header: a stored copy of a computed value is exactly the drift a
' Derived field exists to prevent.
'
' A SECOND DERIVED FIELD ARRIVED (STATUS_BADGE, 2026-08-19) -- THIS IS THE
' MOMENT THIS COMMENT ITSELF CALLED FOR. It used to say "the only Derived
' field that exists is this one... if a second is ever added, THAT is the
' moment to replace this with a real mechanism." PlanRoutineSync below now
' loops over a small list of known derived tags (see the loop's own header)
' instead of a single hardcoded TIMELINE_ELAPSED block -- still not a
' Field-Spec-driven discovery walk (each derived field's VALUE computation
' is genuinely different -- date math vs. a status lookup -- so there is
' nothing left to generalise there), but the shape-finding/injection/
' bookkeeping around it is shared instead of copy-pasted a second time.
Public Const TIMELINE_ELAPSED_TAG As String = "TIMELINE_ELAPSED"

' STATUS_BADGE -- see DeriveStatusBadge's own header for the derivation
' rule (Rohan wrote it directly onto the live Field Spec sheet, addressed
' to "Claude Code" by name; this constant and DeriveStatusBadge are that
' request, built).
Public Const STATUS_BADGE_TAG As String = "STATUS_BADGE"

' KEY_EVENTS_HEADER -- the bold header line above KEY_EVENTS_BODY. Rohan,
' 2026-08-21: "Key events header is really any important status like
' project closed" -- so unlike PROGRESS_HEADER (Kind=Given, a plain
' directly-editable register column, same pattern as SECTOR/TRL), this one
' is a genuine lookup: show PROJECT_STATUS's own value, but only when it
' ISN'T the default "In Progress" -- an "In Progress" project has nothing
' noteworthy to headline, so the header stays blank rather than stating the
' unremarkable case. See DeriveKeyEventsHeader.
Public Const KEY_EVENTS_HEADER_TAG As String = "KEY_EVENTS_HEADER"

' SUBTITLE_A -- NOT a Derived field in DerivedFieldTags()'s sense, and does
' not belong in that array. It is Kind=Given with a real register column of
' its own, unlike TIMELINE_ELAPSED/STATUS_BADGE which have none -- so the
' ordinary per-field loop in PlanRoutineSync already claims it, and adding it
' to DerivedFieldTags() would never fire (the derived loop's own guard skips
' any tag the ordinary loop already claimed). What IS shared with the Derived
' pattern is the underlying problem: the shape's role tag ("SUBTITLE_A") and
' the register column of the SAME name hold two different things -- one raw
' input, one composite of four. Field Spec's own Behaviour column (row 12,
' column J) states the rule: "SUBTITLE_A - SUBTITLE_B - SECTOR - TRL joined
' with a middot separator." Handled as an explicit substitution at both real
' write call sites (PlanRoutineSync's ordinary loop, ReviewQueue.
' ApplyApproved's "register's value NOW" branch) rather than a general
' mechanism -- SUBTITLE_A is the only field on this sheet whose displayed
' value is a computed join of several OTHER register columns, so a second
' instance is what would justify generalising this into one.
Public Const SUBTITLE_COMPOSITE_FIELD As String = "SUBTITLE_A"

' MODULE-LEVEL DECLARATION, DELIBERATELY UP HERE with the Consts: a bare
' module-level variable below the first procedure compiles quietly wrong and
' reports its error in a DIFFERENT module -- hit twice on 2026-08-17, once
' building the ORIGINAL version of this very fix (AGENTS.md, Known Patterns).
' Moved down to here, still above every procedure, when SUBTITLE_COMPOSITE_
' FIELD's block landed above it and nearly reintroduced the exact defect this
' comment already warns about -- caught by the compiler, not by re-reading.
Public mTestForcePlanCrash As Boolean

' The four inputs, in the declared order, never re-typed elsewhere -- see
' ComposeSubtitleLine.
Private Function SubtitleComponentFields() As Variant
    SubtitleComponentFields = Array("SUBTITLE_A", "SUBTITLE_B", "SECTOR", "TRL")
End Function

' Builds the one line actually shown under the title, from whatever of the
' four inputs the row actually has. Field Spec's own words: "The separators
' are chrome" -- an empty segment contributes NEITHER text nor a dangling
' middot, so two present segments read as "A · D", never "A ·  · · D".
' Chr$(183), not a literal middle-dot character, so the byte is unambiguous
' regardless of how this .bas file's own encoding is read or re-saved --
' this project has already been burned once by a line-ending assumption
' silently changing a file's meaning (AppEvents.cls importing as the wrong
' component type), and a raw non-ASCII glyph in source is the same class of
' risk for one character instead of a whole file.
Public Function ComposeSubtitleLine(rowValues As Object) As String
    Dim fields As Variant
    fields = SubtitleComponentFields()

    Dim parts() As String
    ReDim parts(UBound(fields))
    Dim n As Long
    n = 0

    Dim f As Variant
    For Each f In fields
        Dim v As String
        v = ""
        If Not rowValues Is Nothing Then
            If rowValues.Exists(CStr(f)) Then v = Trim(CStr(rowValues(CStr(f))))
        End If
        If v <> "" Then
            parts(n) = v
            n = n + 1
        End If
    Next f

    Dim s As String
    Dim i As Long
    For i = 0 To n - 1
        If s <> "" Then s = s & " " & Chr$(183) & " "
        s = s & parts(i)
    Next i
    ComposeSubtitleLine = s
End Function

' TEST-ONLY HOOK, same shape and same reason as ReviewQueue.
' mTestForceInjectCrash (its comment carries the full history). Error 50290
' (FIX-LIST.md item V) hit a FOURTH call site on 2026-08-19 -- this time
' during the queue-BUILD phase (RibbonUI.BuildAllQueuesCore -> ReviewQueue.
' BuildQueue -> PlanRoutineSync), before ApplyApproved's per-item loop was
' ever reached, and once again the top-level dialog said nothing more
' specific than "VBAProject" because nothing captured Err.Description/Source
' at the actual point of failure. The fault cannot be summoned from Office on
' demand, so this flag lets a test simulate one deterministically inside
' PlanRoutineSync's per-field probe, proving the per-item capture actually
' captures and enriches rather than merely compiling. Set only by
' Test_SyncOperations_PlanRoutineSyncNamesTheItemWhenProbeCrashes; never
' read by anything reachable from a button.
'
' Returns a string fraction 0-1 (e.g. "0.42"), clamped, or "" if either
' date is missing or unparseable -- refusing rather than drawing a wrong
' bar, the same rule InjectProgressVia already applies to an out-of-range
' register value.
Public Function ElapsedFraction(startText As String, endText As String) As String
    If Trim(startText) = "" Or Trim(endText) = "" Then Exit Function

    Dim startDate As Date, endDate As Date
    On Error GoTo BadDate
    startDate = CDate(Trim(startText))
    endDate = CDate(Trim(endText))
    On Error GoTo 0

    If endDate <= startDate Then Exit Function

    Dim frac As Double
    frac = (Date - startDate) / (endDate - startDate)
    If frac < 0 Then frac = 0
    If frac > 1 Then frac = 1

    ElapsedFraction = Format(frac, "0.####")
    Exit Function

BadDate:
    ElapsedFraction = ""
End Function

' THE STATUS BADGE. Rohan wrote this derivation rule directly onto the live
' Field Spec sheet's STATUS_BADGE row (row 49, 2026-08-19), addressed to
' "Claude Code" by name -- the priority order and wording below are his,
' transcribed exactly, not reinterpreted:
'
'   PRIORITY ORDER, highest wins. Show exactly ONE word, never two joined.
'   1. PROJECT_STATUS = Not Started              -> "Not Started"
'   2. PROJECT_STATUS = Project Closed            -> "Project Closed"
'   3. In Progress AND SCHEDULE_STATUS = Delayed  -> "Delayed"
'   4. In Progress AND SCHEDULE_STATUS = At Risk  -> "At Risk"
'   5. In Progress AND anything else              -> "In Progress"
'   Rationale: schedule health is meaningless before a project starts or
'   after it closes, so lifecycle stage wins outright at both ends.
'
' "ONE WORD" means one label, not literally one English word -- the point
' (Rohan's own wording) is that the badge never shows two things joined
' together, not that "Project Closed" is somehow one word.
'
' UNVERIFIED, PER ROHAN'S OWN FLAG ON THE SHEET, NOT SILENTLY RESOLVED HERE:
' whether SCHEDULE_STATUS is meaningful per-project or per-milestone-row is
' still open against the source tracker. This function reads it as ONE
' value per project, because that is the only shape the Register actually
' stores it in today (one SCHEDULE_STATUS column, confirmed live 2026-08-19)
' -- not a resolution of the open question, just the only input available.
'
' PROJECT_STATUS is Kind=Controlled (FieldSpec.bas, fixed vocabulary), so an
' unrecognised value here means the vocabulary and this derivation have
' drifted apart, not a case to guess a badge for. Refuses rather than
' invents a sixth word -- the same "refuse rather than draw a wrong bar"
' instinct as ElapsedFraction immediately above.
Public Function DeriveStatusBadge(projectStatus As String, scheduleStatus As String) As String
    Dim ps As String, ss As String
    ps = Trim(projectStatus)
    ss = Trim(scheduleStatus)

    If StrComp(ps, "Not Started", vbTextCompare) = 0 Then
        DeriveStatusBadge = "Not Started"
        Exit Function
    End If

    If StrComp(ps, "Project Closed", vbTextCompare) = 0 Then
        DeriveStatusBadge = "Project Closed"
        Exit Function
    End If

    If StrComp(ps, "In Progress", vbTextCompare) = 0 Then
        If StrComp(ss, "Delayed", vbTextCompare) = 0 Then
            DeriveStatusBadge = "Delayed"
        ElseIf StrComp(ss, "At Risk", vbTextCompare) = 0 Then
            DeriveStatusBadge = "At Risk"
        Else
            DeriveStatusBadge = "In Progress"
        End If
        Exit Function
    End If

    DeriveStatusBadge = ""
End Function

' Simple lookup, not a priority ladder like DeriveStatusBadge -- Rohan's own
' framing was "any important status like project closed", and PROJECT_STATUS
' is the one field that already carries that. "In Progress" is the default,
' unremarkable state every normally-running project sits in most of the
' time, so it headlines nothing; anything else (Not Started, Project Closed,
' or a future vocabulary addition) is exactly the kind of thing worth
' flagging above KEY_EVENTS_BODY. PROJECT_STATUS is Kind=Controlled, so an
' unrecognised value still passes through here unfiltered -- there is no
' invented sixth word to guess, unlike DeriveStatusBadge which manufactures
' a badge from two inputs.
Public Function DeriveKeyEventsHeader(projectStatus As String) As String
    Dim ps As String
    ps = Trim(projectStatus)
    If ps = "" Then Exit Function
    If StrComp(ps, "In Progress", vbTextCompare) = 0 Then Exit Function
    DeriveKeyEventsHeader = ps
End Function

' THE DERIVED-FIELD LIST. Small and explicit on purpose -- see the comment
' on TIMELINE_ELAPSED_TAG above for why this isn't a Field-Spec-driven
' discovery walk. Adding a fourth derived field means adding its tag here
' and a Case to ComputeDerivedValue below; nothing else changes.
Public Function DerivedFieldTags() As Variant
    DerivedFieldTags = Array(TIMELINE_ELAPSED_TAG, STATUS_BADGE_TAG, KEY_EVENTS_HEADER_TAG)
End Function

' One place that knows how to compute EACH derived field's value from the
' row it was given. Returns "" for "cannot be computed from what's here" --
' the caller (PlanRoutineSync) already treats "" as skip, the same refusal
' shape ElapsedFraction and DeriveStatusBadge both already use on their own.
Public Function ComputeDerivedValue(fieldId As String, rowValues As Object) As String
    Select Case fieldId
        Case TIMELINE_ELAPSED_TAG
            Dim startVal As String, endVal As String
            startVal = "": endVal = ""
            If rowValues.Exists("START_DATE") Then startVal = CStr(rowValues("START_DATE"))
            If rowValues.Exists("END_DATE") Then endVal = CStr(rowValues("END_DATE"))
            ComputeDerivedValue = ElapsedFraction(startVal, endVal)

        Case STATUS_BADGE_TAG
            Dim psVal As String, ssVal As String
            psVal = "": ssVal = ""
            If rowValues.Exists("PROJECT_STATUS") Then psVal = CStr(rowValues("PROJECT_STATUS"))
            If rowValues.Exists("SCHEDULE_STATUS") Then ssVal = CStr(rowValues("SCHEDULE_STATUS"))
            ComputeDerivedValue = DeriveStatusBadge(psVal, ssVal)

        Case KEY_EVENTS_HEADER_TAG
            Dim kehVal As String
            kehVal = ""
            If rowValues.Exists("PROJECT_STATUS") Then kehVal = CStr(rowValues("PROJECT_STATUS"))
            ComputeDerivedValue = DeriveKeyEventsHeader(kehVal)

        Case Else
            ComputeDerivedValue = ""
    End Select
End Function

Public Function PlanRoutineSync(instances() As Object, instanceOrder As Collection, dataRows As Object, _
                                Optional dryRun As Boolean = False, _
                                Optional logWs As Object = Nothing, _
                                Optional runStamp As String = "", _
                                Optional srcWs As Object = Nothing) As SyncAction()
    Dim actions() As SyncAction
    Dim n As Long
    n = 0

    ' `instances` may be genuinely unallocated -- zero known instances of
    ' this type exist yet (e.g. before anything of this type has ever been
    ' onboarded). ReDim-to-(1 To 0) throws at runtime (a real, confirmed VBA
    ' restriction, see AGENTS.md's Known Patterns), so this can't be
    ' represented the way Discovery.bas's own "(1 To 0) convention" this
    ' comment used to describe assumed -- error-guard instead of assuming
    ' LBound/UBound are always safe to call.
    Dim hasInstances As Boolean
    Dim lo As Long, hi As Long
    On Error Resume Next
    lo = LBound(instances)
    hi = UBound(instances)
    hasInstances = (Err.Number = 0)
    On Error GoTo 0

    ' Captured per COM call site below -- the crash capture for FIX-LIST item
    ' V's fourth occurrence (see GuardedPlanProbe's header for the pattern).
    Dim planErrNum As Long, planErrDesc As String, planErrSrc As String

    Dim resolved() As SlideInstance
    Dim i As Long
    If hasInstances Then
        ReDim resolved(lo To hi)
        For i = lo To hi
            ' Per-slide COM work (tag reads) -- trapped like every other
            ' per-item COM call in this function.
            On Error Resume Next
            Err.Clear
            resolved(i) = Resolve.ResolveSlideInstance(instances(i))
            planErrNum = Err.Number: planErrDesc = Err.Description: planErrSrc = Err.Source
            On Error GoTo 0
            If planErrNum <> 0 Then
                Dim crashSlideLabel As String
                crashSlideLabel = "slide " & i & " of " & hi
                ' SlideID read is itself COM -- best-effort only, never let
                ' the label lookup mask the error being reported.
                On Error Resume Next
                crashSlideLabel = crashSlideLabel & " (SlideID " & instances(i).SlideID & ")"
                On Error GoTo 0
                ReviewQueue.LogAndReraiseCrash logWs, runStamp, "SyncOperations.PlanRoutineSync", _
                    "", "", "resolving tags of " & crashSlideLabel, planErrNum, planErrDesc, planErrSrc
            End If
            If Not resolved(i).HasTypeTag Or Not resolved(i).HasInstanceKey Then
                n = n + 1
                ReDim Preserve actions(1 To n)
                actions(n).Kind = "flagged"
                actions(n).Subject = "SlideID " & instances(i).SlideID
                actions(n).FlagKind = "unclassified_slide"
                actions(n).Reason = "no recognized type tag / persistent instance key -- flagged for reclassification, not guessed"
            End If
        Next i
    End If

    Dim knownByKey As Object
    Set knownByKey = CreateObject("Scripting.Dictionary")
    If hasInstances Then
        For i = lo To hi
            If resolved(i).HasInstanceKey Then
                knownByKey(resolved(i).InstanceKey) = i
            End If
        Next i
    End If

    Dim instanceId As Variant
    For Each instanceId In instanceOrder
        Dim key As String
        key = CStr(instanceId)

        Dim rowValues As Object
        If dataRows.Exists(key) Then
            Set rowValues = dataRows(key)
        Else
            Set rowValues = CreateObject("Scripting.Dictionary")
        End If

        If Not knownByKey.Exists(key) Then
            n = n + 1
            ReDim Preserve actions(1 To n)
            actions(n).Kind = "new_record"
            actions(n).RowInstanceKey = key
            Set actions(n).Values = CloneStringDict(rowValues)
            actions(n).Reason = "no known slide instance carries this row's instance key"
        Else
            Dim idx As Long
            idx = knownByKey(key)
            Dim instanceSlide As Object
            Set instanceSlide = instances(idx)

            Dim changedVerified As Object, changedError As Object, changedCurrent As Object
            Dim changedNew As Object
            Set changedVerified = CreateObject("Scripting.Dictionary")
            Set changedError = CreateObject("Scripting.Dictionary")
            Set changedCurrent = CreateObject("Scripting.Dictionary")
            Set changedNew = CreateObject("Scripting.Dictionary")

            Dim fieldName As Variant
            For Each fieldName In rowValues.Keys
                Dim sourceValue As String
                ' SUBTITLE_A's shape shows a composite, not its own raw
                ' register value -- see SUBTITLE_COMPOSITE_FIELD's own note.
                If CStr(fieldName) = SUBTITLE_COMPOSITE_FIELD Then
                    sourceValue = ComposeSubtitleLine(rowValues)
                Else
                    sourceValue = rowValues(fieldName)
                End If

                Dim r As InjectResult
                ' Routed by shape type. No Sources sheet is passed: PlanRoutineSync
                ' is handed rows, not a workbook, so a picture field here reports
                ' that it could not be resolved rather than being written wrongly.
                ' Progress bars need nothing extra and work on this path.
                '
                ' GUARDED, same as the device/elapsed-bar probes below --
                ' this is the MAIN field probe, hit for every ordinary
                ' register column on every slide, and until now it was the
                ' one call in this function with no crash capture at all.
                ' Deliberately left unwrapped through the rest of tonight's
                ' fix so Test_SyncOperations_PlanRoutineSyncNamesTheItemWhen
                ' ProbeCrashes could prove it fails without this wrapping --
                ' confirmed 2026-08-19: raw Err.Source, nothing logged. This
                ' is that fix, applied.
                r = GuardedPlanProbe(instanceSlide, key, CStr(fieldName), sourceValue, dryRun, rowValues, logWs, runStamp, srcWs)

                ' r.Found = False covers both "no shape carries this
                ' field's tag" (skip -- matches resolve.py's
                ' field_shapes.get() returning None -> skip) and "more
                ' than one shape carries it" (ambiguous -- also skipped
                ' here, not separately flagged). Deliberately conflating
                ' those two Python-distinguishable situations into one
                ' skip outcome: disambiguating structural drift like a
                ' duplicate role tag is case-7 (deck_side_conflict)
                ' adjacent territory, a non-goal per
                ' specs/sync-operations.md. See SPIKE_NOTES_Resolve.md.
                ' A dry run never sets Written (nothing was written), so the
                ' gate has to accept WouldChange too -- otherwise a preview
                ' would report every instance as "no_change", which is the
                ' most dangerous possible wrong answer for a preview to give.
                If r.Found And (r.Written Or r.WouldChange) Then
                    changedVerified(fieldName) = r.Verified
                    changedError(fieldName) = r.ErrorMessage
                    changedCurrent(fieldName) = r.CurrentValue
                    ' sourceValue, not anything read back off the slide: in a dry
                    ' run nothing was written, so the slide cannot be asked what
                    ' the new text is. The Data sheet is the only source for it.
                    changedNew(fieldName) = sourceValue
                End If
            Next fieldName

            ' DEVICE FIELDS ARE NOT REGISTER COLUMNS. FIX-LIST R, 2026-08-15:
            ' the loop above can only ask about a field whose name is a
            ' register column, and a device's identity tag (MILESTONE_TIMELINE)
            ' is not one -- its data lives across 21 separate columns instead.
            ' Discovered by walking the slide's own shapes, same test
            ' InjectorFor already uses to decide a tag routes to a device, so
            ' this asks about exactly the devices actually on this slide.
            Dim deviceTags As Object
            ' The scan walks every shape on the slide -- per-item COM work,
            ' trapped like the probes.
            On Error Resume Next
            Err.Clear
            Set deviceTags = InjectPrimitive.DeviceRoleTagsOnSlide(instanceSlide)
            planErrNum = Err.Number: planErrDesc = Err.Description: planErrSrc = Err.Source
            On Error GoTo 0
            If planErrNum <> 0 Then
                ReviewQueue.LogAndReraiseCrash logWs, runStamp, "SyncOperations.PlanRoutineSync", _
                    key, "", "scanning for device tags", planErrNum, planErrDesc, planErrSrc
            End If
            Dim devTag As Variant
            For Each devTag In deviceTags.Keys
                If Not changedVerified.Exists(devTag) Then
                    Dim rd As InjectResult
                    ' NOTE this probe is NOT read-only even when dryRun is
                    ' True: InjectDeviceVia calls MilestoneDevice.
                    ' DeviceIntegrity(grp) unconditionally (InjectPrimitive.
                    ' bas), so a dry run still does real COM work against the
                    ' device group -- one reason a mid-PLAN crash was always
                    ' plausible here.
                    rd = GuardedPlanProbe(instanceSlide, key, CStr(devTag), "", dryRun, rowValues, logWs, runStamp, srcWs)
                    If rd.Found And (rd.Written Or rd.WouldChange) Then
                        changedVerified(devTag) = rd.Verified
                        changedError(devTag) = rd.ErrorMessage
                        changedCurrent(devTag) = rd.CurrentValue
                        ' No single "new value" exists for a device -- it is
                        ' redrawn from many columns at once, not one cell.
                        ' rd.ErrorMessage carries DrawFromRow's own Detail
                        ' (drawn/hidden counts, any per-slot note) on success.
                        changedNew(devTag) = "(redrawn from its register columns)"
                    End If
                End If
            Next devTag

            ' DERIVED FIELDS -- computed here, at sync time, from whatever
            ' source columns each one needs (see ComputeDerivedValue), never
            ' read from or written to a register column of their own. One
            ' shared loop over DerivedFieldTags() rather than a block per
            ' field -- see TIMELINE_ELAPSED_TAG's own comment for why this
            ' changed 2026-08-19.
            Dim derivedTag As Variant
            For Each derivedTag In DerivedFieldTags()
                Dim derivedTagStr As String
                derivedTagStr = CStr(derivedTag)
                If Not changedVerified.Exists(derivedTagStr) Then
                    Dim derivedShp As Object
                    ' Another whole-slide shape walk -- trapped like the scan above.
                    On Error Resume Next
                    Err.Clear
                    Set derivedShp = InjectPrimitive.FindShapeByRoleTag(instanceSlide, derivedTagStr)
                    planErrNum = Err.Number: planErrDesc = Err.Description: planErrSrc = Err.Source
                    On Error GoTo 0
                    If planErrNum <> 0 Then
                        ReviewQueue.LogAndReraiseCrash logWs, runStamp, "SyncOperations.PlanRoutineSync", _
                            key, derivedTagStr, "locating the " & derivedTagStr & " shape", _
                            planErrNum, planErrDesc, planErrSrc
                    End If
                    If Not derivedShp Is Nothing Then
                        Dim derivedVal As String
                        derivedVal = ComputeDerivedValue(derivedTagStr, rowValues)
                        If derivedVal <> "" Then
                            Dim dre As InjectResult
                            dre = GuardedPlanProbe(instanceSlide, key, derivedTagStr, derivedVal, dryRun, rowValues, logWs, runStamp, srcWs)
                            If dre.Found And (dre.Written Or dre.WouldChange) Then
                                changedVerified(derivedTagStr) = dre.Verified
                                changedError(derivedTagStr) = dre.ErrorMessage
                                changedCurrent(derivedTagStr) = dre.CurrentValue
                                changedNew(derivedTagStr) = derivedVal
                            End If
                        End If
                    End If
                End If
            Next derivedTag

            n = n + 1
            ReDim Preserve actions(1 To n)
            If changedVerified.count > 0 Then
                actions(n).Kind = "in_place_correction"
                actions(n).InstanceKey = key
                Set actions(n).ChangedFieldVerified = changedVerified
                Set actions(n).ChangedFieldError = changedError
                Set actions(n).ChangedFieldCurrent = changedCurrent
                Set actions(n).ChangedFieldNew = changedNew
            Else
                actions(n).Kind = "no_change"
                actions(n).InstanceKey = key
            End If
        End If
    Next instanceId

    PlanRoutineSync = actions
End Function

' One guarded call to InjectPrimitive.InjectField on behalf of PlanRoutineSync,
' trapping locally so a mid-plan COM fault (Error 50290, FIX-LIST item V --
' its fourth occurrence landed in this planning chain, 2026-08-19) is captured
' at the actual point of failure: the Sync Log line is written BEFORE the
' re-raise (it must survive PowerPoint dying entirely), and Err.Source gains
' the specific instance/field so the top-level dialog names the item instead
' of "VBAProject". Same pattern as ReviewQueue.ApplyApproved's per-item traps,
' one layer earlier in the chain. The phase string distinguishes a dry probe
' from PlanRoutineSync's write mode (dryRun:=False is the real Sync Now path);
' WHICH probe it was is already told by the fieldId itself (a device tag, the
' elapsed bar's tag, or an ordinary register column).
'
' The test hook fires here, in place of the real call, for the same reason
' ReviewQueue.mTestForceInjectCrash fires inside ApplyApproved's trap: Office
' cannot be made to raise the real fault on demand, so proving this capture
' works at all requires a deterministic stand-in at the exact wrapped site.
Private Function GuardedPlanProbe(instanceSlide As Object, key As String, fieldId As String, _
                                  sourceValue As String, dryRun As Boolean, rowValues As Object, _
                                  logWs As Object, runStamp As String, _
                                  Optional srcWs As Object = Nothing) As InjectResult
    Dim errNum As Long, errDesc As String, errSrc As String
    On Error Resume Next
    Err.Clear
    If mTestForcePlanCrash Then
        Err.Raise 12346, "InjectPrimitive.InjectField", "TEST: deliberately injected fault"
    Else
        ' srcWs THREADED THROUGH, NOT HARDCODED NOTHING. A picture field's
        ' proposed value is a source-ID citation, not a value InjectorFor can
        ' resolve on its own -- without the real Sources sheet here,
        ' InjectPictureVia can only ever report "Sources sheet was not
        ' available here" and force WouldChange=True, so a picture field
        ' could never correctly report itself unchanged during planning, even
        ' once it is genuinely synced. Found 2026-08-18 wiring the first real
        ' picture field (PROJECT_PHOTO) -- ApplyApproved already resolved
        ' srcWs correctly; PlanRoutineSync's own dry-run probe, one call
        ' earlier in the same chain, never did.
        GuardedPlanProbe = InjectPrimitive.InjectField(instanceSlide, fieldId, sourceValue, dryRun, srcWs, rowValues)
    End If
    errNum = Err.Number: errDesc = Err.Description: errSrc = Err.Source
    On Error GoTo 0
    If errNum <> 0 Then
        ReviewQueue.LogAndReraiseCrash logWs, runStamp, "SyncOperations.PlanRoutineSync", _
            key, fieldId, IIf(dryRun, "planning dry probe", "sync write"), _
            errNum, errDesc, errSrc
    End If
End Function

Private Function CloneStringDict(source As Object) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In source.Keys
        result(k) = source(k)
    Next k
    Set CloneStringDict = result
End Function

' ---------------------------------------------------------------------
' Manual smoke test -- not a real test harness, same as every other module
' here. See SPIKE_NOTES_Resolve.md for the full recipe and expected
' values, cross-checked against tests/test_resolve.py's already-proven
' end-to-end results.
' ---------------------------------------------------------------------

' Run with slide 1 of the active presentation already tagged (via the
' Immediate window, before running this):
'   Application.ActivePresentation.Slides(1).Tags.Add "slide_type", "quarterly-update"
'   Application.ActivePresentation.Slides(1).Tags.Add "instance_key", "rec-1"
'   Application.ActivePresentation.Slides(1).Shapes(1).Tags.Add "role", "Title"
' (pick the actual title shape index/name for your slide if it isn't
' Shapes(1)). Seeds the Data-sheet row with the shape's own current text,
' so this should report "no_change" first, then "in_place_correction"
' after you change TITLE_TEXT below to something else and re-run.
Public Sub ManualSmokeTest_NoChangeThenInPlaceCorrection()
    Const TITLE_TEXT As String = "" ' leave blank to seed from the shape's current text (expect no_change)

    Dim sld As Object
    Set sld = Application.ActivePresentation.Slides(1)

    Dim instances(1 To 1) As Object
    Set instances(1) = sld

    Dim order As New Collection
    order.Add "rec-1"

    Dim rowsDict As Object
    Set rowsDict = CreateObject("Scripting.Dictionary")
    Dim rec1Fields As Object
    Set rec1Fields = CreateObject("Scripting.Dictionary")
    If TITLE_TEXT = "" Then
        rec1Fields("Title") = sld.Shapes(1).TextFrame.TextRange.Text
    Else
        rec1Fields("Title") = TITLE_TEXT
    End If
    Set rowsDict("rec-1") = rec1Fields

    Dim actions() As SyncAction
    actions = PlanRoutineSync(instances, order, rowsDict)

    Dim msg As String
    msg = "actions=" & (UBound(actions) - LBound(actions) + 1) & " Kind=" & actions(1).Kind & _
        " InstanceKey=" & actions(1).InstanceKey
    Debug.Print msg
    MsgBox msg & " (expected: no_change when TITLE_TEXT is blank, in_place_correction otherwise)"
End Sub
