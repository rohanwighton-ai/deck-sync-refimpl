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
Private Const MARKING_SESSION_PROPERTY_NAME As String = "DeckSyncMarkingSession"
Private Const COL_FIELD_ID As Long = 1
Private Const COL_FIELD_NAME As Long = 2
Private Const COL_SUGGESTED As Long = 3
Private Const COL_INCLUDE As Long = 4
Private Const COL_TEMPLATE_VALUE As Long = 5
Private Const COL_SAMPLE_OTHER_VALUES As Long = 6
Private Const COL_TYPE As Long = 7
Private Const COL_VOLATILITY As Long = 8
' How much of a field's current text to show a human at mark time. Short on
' purpose: this is "enough to recognise which shape I'm looking at" while the
' slide is on screen in front of you, not an identifier. Measured against the
' real 46-slide deck 2026-07-27: 20 chars and 30 chars separate exactly the
' same number of fields (42/46 for About text), so a longer preview buys no
' distinguishing power and only makes the prompt harder to read. Never used as
' a field name, a role tag, or an instance key -- text-derived identifiers are
' unstable under ordinary editing, which is what corrupted the real deck's
' linkage on 2026-07-26.
Private Const FIELD_PREVIEW_CHARS As Long = 20

' How long the deck's last REAL save may go unchanged before the add-in stops
' trusting AutoSave and saves for itself.
'
' 120 seconds, not 60: "Last Save Time" is stored at minute resolution, so a
' deck saved 40 seconds ago can legitimately read as ~100 seconds old. Two
' minutes clears that granularity, and it is still three orders of magnitude
' below the real failure this guard exists to catch (2026-07-28: AutoSave on,
' file on disk 2.6 HOURS stale, both marks sitting only in the open document).
' Erring long is the cheap direction -- an unnecessary wait costs nothing,
' an unnecessary save costs Office's own Save command (see the note in
' SaveMarkingSessionToProperty).
Private Const AUTOSAVE_STALE_SECONDS As Double = 120
Public Type BatchOnboardPlan
    FieldCount As Long
    FieldNames As Object          ' Dictionary: fieldIndex -> String (proposed, editable)
    FieldTypes As Object          ' Dictionary: fieldIndex -> String ("text"|"number"|"currency"|"date")
    FieldVolatility As Object     ' Dictionary: fieldIndex -> String ("static"|"variable") -- human-declared hint only, not wired into sync behavior yet (see NormalizeFieldVolatility's own header)
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

' Marking-session state for MarkFieldForBatch/ClearMarkedFieldsForBatch --
' persists across separate macro invocations while the add-in stays loaded
' (each toolbar click is its own synchronous macro run; VBA has no way to
' pause mid-run and wait for a canvas click without a WithEvents class
' module, which this project has never used -- see MarkFieldForBatch's own
' header for why a click-then-button-press loop was chosen instead).
Private markedShapes As Collection
Private markedNames As Object ' Scripting.Dictionary: 1-based position in markedShapes -> String (the human-typed field name)
Private markedTypes As Object ' Scripting.Dictionary: 1-based position in markedShapes -> String ("text"|"number"|"currency"|"date")
Private markedVolatility As Object ' Scripting.Dictionary: 1-based position in markedShapes -> String ("static"|"variable")
Private markedSlideId As Long
' Which DECK the in-memory session belongs to. Closing a presentation does NOT
' unload the add-in, so module state (and its now-dead Shape references)
' outlives the document it came from. Without this, the restore guard sees a
' non-empty session, concludes one is already loaded, and never reads the saved
' one back -- so a user who marks fields, closes, and reopens is told nothing
' was marked while their marks sit intact in the file. Confirmed live
' 2026-07-28: MarkedFieldCountForBatch still returned 1 after a full
' close/reopen, and the saved property read back perfectly at the same moment.
Private markedDeckId As String

' The last "Last Save Time" value this module saw, and the LOCAL clock reading
' at the moment it saw it. Together they answer the only question that matters
' here: "has the file actually been written since the last time we looked, and
' if not, how long has it been?"
'
' Deliberately a comparison of the property against ITSELF over time, never
' against Now. The absolute value of Last Save Time cannot safely be subtracted
' from the local clock -- if Office ever hands it back in UTC, a machine in
' Adelaide (UTC+9:30) would read every freshly-saved deck as 9.5 hours stale and
' the guard below would silently degrade back into an unconditional save, which
' is the exact bug it replaces. Equality of two readings plus a local elapsed
' time is immune to that.
Private lastObservedSaveStamp As Double
Private lastObservedSaveAt As Double

' ---------------------------------------------------------------------
' Marking-session persistence -- survives a PowerPoint close/reopen, built
' 2026-07-26 per Rohan's own framing ("sick of linking every test" --
' re-marking every field from scratch after every add-in reload during
' live testing was real, repeated friction). Stored as a CustomDocument-
' Property on the presentation (same mechanism DeckRegistry.bas already
' uses, and already proven against real Office in that module's own
' passing tests) -- travels with the deck once saved, same as any other
' document property. A live Shape reference can't itself survive a close,
' so only each field's Name is persisted; restoring re-finds the shape by
' Name via Discovery (already recurses into groups). Any field whose name
' can no longer be found (renamed, deleted) is skipped, not guessed --
' reported by count so it's never silently fewer fields than expected.
' ---------------------------------------------------------------------

' Pure -- exercised directly by TestRunner.bas. One field per line
' ("Name|FieldName|Type|Volatility"), first line the marked slide's
' SlideID, vbCrLf-joined -- same small delimited-string convention
' DeckRegistry.bas's own BuildTypeRegistration already established.
Public Function SerializeMarkingSession(slideId As Long, shapes As Collection, names As Object, types As Object, volatility As Object) As String
    Dim lines As String
    lines = CStr(slideId)
    Dim i As Long
    For i = 1 To shapes.count
        lines = lines & vbCrLf & shapes(i).Name & FIELD_KEY_SEP & names(i) & FIELD_KEY_SEP & types(i) & FIELD_KEY_SEP & volatility(i)
    Next i
    SerializeMarkingSession = lines
End Function

' Restores a session serialized by SerializeMarkingSession against
' `templateSld` -- populates the module-level marking-session state
' directly (same shape a real restore needs, not a fresh return value).
' Returns a human-readable summary; never raises on a malformed/missing
' string, matching this module's "readable status, not an exception"
' convention throughout.
Public Function RestoreMarkingSession(serialized As String, templateSld As Object) As String
    If Trim(serialized) = "" Then
        RestoreMarkingSession = "Nothing to restore."
        Exit Function
    End If

    Dim lines() As String
    lines = Split(serialized, vbCrLf)

    Dim lo As Long, hi As Long, hasLines As Boolean
    On Error Resume Next
    lo = LBound(lines): hi = UBound(lines)
    hasLines = (Err.Number = 0)
    On Error GoTo 0
    If Not hasLines Or hi < 1 Then
        RestoreMarkingSession = "Nothing to restore."
        Exit Function
    End If

    Dim slideId As Long
    slideId = CLng(lines(0))

    Dim allCandidates() As Candidate
    Dim allShapes() As Object
    allCandidates = Discovery.DiscoverSlideWithShapes(templateSld, allShapes)

    Dim aLo As Long, aHi As Long, hasAny As Boolean
    On Error Resume Next
    aLo = LBound(allShapes): aHi = UBound(allShapes)
    hasAny = (Err.Number = 0)
    On Error GoTo 0

    Set markedShapes = New Collection
    Set markedNames = CreateObject("Scripting.Dictionary")
    Set markedTypes = CreateObject("Scripting.Dictionary")
    Set markedVolatility = CreateObject("Scripting.Dictionary")
    markedSlideId = slideId
    markedDeckId = DeckIdentity(templateSld.Parent)

    Dim restoredCount As Long, missingCount As Long
    restoredCount = 0
    missingCount = 0

    Dim li As Long
    For li = 1 To hi
        If Trim(lines(li)) <> "" Then
            Dim parts() As String
            parts = Split(lines(li), FIELD_KEY_SEP)

            Dim foundShp As Object
            Set foundShp = Nothing
            If hasAny Then
                Dim ai As Long
                For ai = aLo To aHi
                    If allShapes(ai).Name = parts(0) Then
                        Set foundShp = allShapes(ai)
                        Exit For
                    End If
                Next ai
            End If

            If foundShp Is Nothing Then
                missingCount = missingCount + 1
            Else
                markedShapes.Add foundShp
                Dim posIdx As Long
                posIdx = markedShapes.count
                markedNames(posIdx) = parts(1)
                markedTypes(posIdx) = parts(2)
                markedVolatility(posIdx) = parts(3)
                restoredCount = restoredCount + 1
            End If
        End If
    Next li

    RestoreMarkingSession = "Restored " & restoredCount & " field(s) from a previous session." & IIf(missingCount > 0, " (" & missingCount & " could not be re-found -- shape renamed or deleted since marking.)", "")
End Function

Public Sub WriteMarkingSessionProperty(pres As Object, value As String)
    Dim prop As Object
    On Error Resume Next
    Set prop = pres.CustomDocumentProperties(MARKING_SESSION_PROPERTY_NAME)
    On Error GoTo 0
    If prop Is Nothing Then
        pres.CustomDocumentProperties.Add Name:=MARKING_SESSION_PROPERTY_NAME, _
            LinkToContent:=False, Type:=msoPropertyTypeString, Value:=value
    Else
        prop.Value = value
    End If
End Sub

Public Function ReadMarkingSessionProperty(pres As Object) As String
    Dim prop As Object
    On Error Resume Next
    Set prop = pres.CustomDocumentProperties(MARKING_SESSION_PROPERTY_NAME)
    On Error GoTo 0
    If Not prop Is Nothing Then ReadMarkingSessionProperty = CStr(prop.Value)
End Function

' Reads the deck's own "Last Save Time" -- the timestamp Office stamps when the
' file is ACTUALLY written -- as a Double, or 0 when it cannot be read.
'
' This exists because Presentation.Saved cannot do this job. Saved is a dirty
' flag, and on an AutoSave-connected cloud document it is confirmed
' untrustworthy: it read False on 2026-07-26 immediately after a fully
' successful, durable save, and False again on 2026-07-28 when the file
' genuinely was 2.6 hours stale. A signal that reads the same in both the
' healthy and the broken case cannot distinguish them -- which is precisely how
' the previous "conditional" force-save ended up firing unconditionally on
' every cloud deck.
'
' Returning 0 on failure is the deliberate "assume the worst" value: callers
' treat an unreadable timestamp as stale, so a deck whose save state cannot be
' established gets saved rather than trusted.
Public Function LastSaveTimeOf(pres As Object) As Double
    Dim raw As Variant

    On Error Resume Next
    raw = pres.BuiltInDocumentProperties("Last Save Time").Value
    If Err.Number <> 0 Then
        On Error GoTo 0
        LastSaveTimeOf = 0
        Exit Function
    End If
    On Error GoTo 0

    ' A never-saved deck reports this as empty rather than raising.
    On Error Resume Next
    LastSaveTimeOf = CDbl(CDate(raw))
    If Err.Number <> 0 Then LastSaveTimeOf = 0
    On Error GoTo 0
End Function

' Should the add-in save the deck itself, or leave it to AutoSave?
'
' Pulled out as a pure function on purpose. The bug this replaces -- a guard
' that was permanently true and so never guarded anything -- was invisible
' because the decision was three inline terms inside a routine that needs a
' real cloud-hosted Presentation to exercise. As a pure function it can be
' tested at every interesting combination without Office, PowerPoint, or a
' OneDrive deck, which is exactly the gap the 2026-07-28 live session exposed:
' 93 tests that verified mechanisms and none that verified this outcome.
'
' Three reasons to save for ourselves, in order of confidence:
'   1. AutoSave is off -- the human owns saving and hasn't been asked to.
'   2. The save timestamp can't be read at all (stamp = 0) -- no evidence
'      either way, so err toward keeping the work.
'   3. The file demonstrably hasn't been written in AUTOSAVE_STALE_SECONDS
'      despite AutoSave claiming to be on -- AutoSave has stalled.
' In the healthy cloud case none of these hold, this returns False, and Office
' keeps ownership of saving along with its native Save command.
Public Function ShouldForceSave(autoSaveOn As Boolean, lastSaveStamp As Double, secondsSinceRealSave As Double) As Boolean
    If Not autoSaveOn Then
        ShouldForceSave = True
    ElseIf lastSaveStamp = 0 Then
        ShouldForceSave = True
    ElseIf secondsSinceRealSave > AUTOSAVE_STALE_SECONDS Then
        ShouldForceSave = True
    Else
        ShouldForceSave = False
    End If
End Function

' Human-readable form of the same timestamp, for the confirmation message.
' Reported as the deck's own raw value rather than as "x minutes ago" on
' purpose: an elapsed-time phrasing would need the local-clock subtraction that
' LastSaveTimeOf's own note rules out, and the literal timestamp is a stronger
' trust signal anyway -- the human can compare it against the clock themselves.
Public Function LastSaveTimeTextOf(pres As Object) As String
    Dim stamp As Double
    stamp = LastSaveTimeOf(pres)
    If stamp = 0 Then
        LastSaveTimeTextOf = "unknown"
    Else
        LastSaveTimeTextOf = Format(CDate(stamp), "yyyy-mm-dd hh:nn")
    End If
End Function

' Troubleshooting utility, added 2026-07-26 -- reads the deck's own
' BuiltInDocumentProperties("Last Save Time") plus Saved/FullName via
' native VBA (callable through Application.Run from PowerShell), since
' PowerShell's own direct COM dispatch is confirmed unreliable for
' DocumentProperties-family collections (see SPIKE_NOTES's addenda) --
' BuiltInDocumentProperties is the same COM family as
' CustomDocumentProperties, so it's expected to hit the identical bug if
' read the same (unreliable) way. This gives a real, trustworthy way to
' check whether AutoSave/manual Ctrl+S is actually updating the deck's own
' save timestamp, independent of anything this add-in writes itself.
Public Function GetSaveDiagnostic() As String
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim lastSaveTime As String
    On Error Resume Next
    lastSaveTime = CStr(pres.BuiltInDocumentProperties("Last Save Time").Value)
    If Err.Number <> 0 Then lastSaveTime = "(could not read: " & Err.Description & ")"
    On Error GoTo 0

    GetSaveDiagnostic = "FullName: " & pres.FullName & vbCrLf & _
        "Saved: " & pres.Saved & vbCrLf & _
        "Last Save Time: " & lastSaveTime
End Function

Private Sub ClearMarkingSessionProperty(pres As Object)
    On Error Resume Next
    pres.CustomDocumentProperties(MARKING_SESSION_PROPERTY_NAME).Delete
    On Error GoTo 0
End Sub

' Saves the current in-memory marking session to the presentation's
' CustomDocumentProperty -- called after every successful mark, so the
' session is always as durable as the deck itself (survives a close once
' the human saves the .pptx, same as any other document property; this
' does not force a save of its own -- see this module's own note on
' AutoSave not reliably picking up Tag-only writes, a related but distinct
' finding from earlier this session).
' Returns "" on success, or a warning message if the save itself failed --
' never raises, same "readable status, not an exception" convention as the
' rest of this module, so the caller can surface a real problem instead of
' silently continuing as if it definitely worked.
'
' Saves for itself only when AutoSave is demonstrably not doing the job --
' judged by the deck's real Last Save Time, not by its dirty flag. See the
' guard below for why, and LastSaveTimeOf for why Presentation.Saved cannot
' be used as that signal.
'
' Microsoft's own guidance (learn.microsoft.com, "How AutoSave impacts add-ins
' and macros") treats AutoSave and macro-owned saving as mutually exclusive:
' the only documented lever it offers an add-in is turning AutoSave OFF via
' AutoSaveOn = False. There is no sanctioned idiom for forcing a save
' underneath a live AutoSave, which is why this code watches for AutoSave
' failing rather than trying to cooperate with it.
Public Function SaveMarkingSessionToProperty(pres As Object) As String
    Dim serialized As String
    serialized = SerializeCurrentMarkingSession()
    If serialized = "" Then Exit Function
    WriteMarkingSessionProperty pres, serialized

    ' DO NOT fight AutoSave. On a cloud-hosted document with AutoSave on,
    ' PowerPoint owns saving, and calling pres.Save underneath it desyncs
    ' Office's own save state: the manual Save command disappears, Ctrl+S
    ' becomes a no-op, and the "Saved" indicator stops being trustworthy.
    ' Confirmed live 2026-07-28 against a real OneDrive-hosted deck --
    ' Rohan's report, and the reason this guard exists: "the app's ability to
    ' demonstrate a clean save and to manually save disappears."
    '
    ' The forced save was originally added for data loss that this project's
    ' own investigation (SPIKE_NOTES_BatchOnboardFlow.md, Finding 2) then
    ' repeatedly FAILED to reproduce -- a 59MB cloud deck, property written,
    ' .Save called, PowerPoint force-killed with zero delay, marks still came
    ' back every time. A fix for a phantom problem was causing a real one.
    '
    ' Property writes are persisted by AutoSave like any other change; that a
    ' marking session survives a genuine close/reopen on a cloud document was
    ' directly verified the same day (37 chars, separators intact). So when
    ' AutoSave is on, write the property and let Office save it. Force the
    ' save only when AutoSave is genuinely off and the human owns saving.
    '
    ' Presentation.AutoSaveOn is read defensively: if the property does not
    ' exist (older PowerPoint), autoSaveOn stays False and the previous
    ' force-save behaviour applies unchanged.
    Dim autoSaveOn As Boolean
    autoSaveOn = False
    On Error Resume Next
    autoSaveOn = pres.AutoSaveOn
    On Error GoTo 0

    ' Decide whether AutoSave is actually keeping up, using the deck's real
    ' save timestamp rather than its dirty flag.
    '
    ' The previous version of this guard tested "Not pres.Saved" and was
    ' therefore not a guard at all: Saved reads False on an AutoSave-connected
    ' cloud document whether or not the save succeeded (see LastSaveTimeOf), so
    ' the condition was permanently true and every cloud deck got the forced
    ' save the guard was written to avoid -- reintroducing the regression Rohan
    ' reported on 2026-07-28, "the app's ability to demonstrate a clean save and
    ' to manually save disappears."
    '
    ' Last Save Time moves only when the file is genuinely written, so watching
    ' it change is a real observation of AutoSave doing its job.
    Dim stamp As Double
    stamp = LastSaveTimeOf(pres)
    If stamp <> lastObservedSaveStamp Then
        ' The file has been written since we last looked -- AutoSave (or the
        ' human) is alive. Reset the clock we measure staleness against.
        lastObservedSaveStamp = stamp
        lastObservedSaveAt = CDbl(Now)
    End If

    Dim secondsSinceRealSave As Double
    secondsSinceRealSave = (CDbl(Now) - lastObservedSaveAt) * 86400#

    If ShouldForceSave(autoSaveOn, stamp, secondsSinceRealSave) Then
        On Error Resume Next
        Err.Clear
        pres.Save
        If Err.Number <> 0 Then
            SaveMarkingSessionToProperty = "WARNING: could not save the deck (" & Err.Description & ") -- your marks may not survive closing PowerPoint. Save manually."
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0

        ' Re-observe after our own save, so the next mark measures staleness
        ' from this write rather than immediately forcing another one.
        Dim savedStamp As Double
        savedStamp = LastSaveTimeOf(pres)
        If savedStamp <> 0 Then lastObservedSaveStamp = savedStamp
        lastObservedSaveAt = CDbl(Now)
    End If

    ' Verifies the WRITE, not the save. Worth being precise about, because an
    ' earlier version of this comment claimed to be a "closed-loop" check on
    ' durability and it is not: ReadMarkingSessionProperty reads the same
    ' in-memory Presentation object that was just written, so it succeeds
    ' whether or not a single byte reached disk. It cannot detect an unsaved
    ' file, and it never could.
    '
    ' It still earns its place -- a failed or truncated property write is a
    ' real failure mode and this catches it -- but durability is established by
    ' the Last Save Time guard above, and the honest trust signal for the human
    ' is the deck's own save timestamp, which the caller reports.
    Dim readBack As String
    readBack = ReadMarkingSessionProperty(pres)
    If readBack <> serialized Then
        SaveMarkingSessionToProperty = "WARNING: the deck was saved but the marking session did not read back correctly afterward (wrote " & Len(serialized) & " chars, read back " & Len(readBack) & ") -- your marks may not survive closing PowerPoint. Save manually and re-mark if needed."
    End If
End Function

' Serializes the CURRENT in-memory marking session (module-level state),
' wrapping SerializeMarkingSession -- the version TestRunner.bas and
' SaveMarkingSessionToProperty can both call without either reaching into
' private module state directly. "" if nothing is currently marked.
' Stable identity for the deck an in-memory marking session belongs to.
' FullName works for both local paths and the cloud URLs OneDrive-hosted decks
' report (https://d.docs.live.net/...), so it identifies a document rather than
' assuming a filesystem.
Private Function DeckIdentity(pres As Object) As String
    On Error Resume Next
    DeckIdentity = pres.FullName
    On Error GoTo 0
End Function

' Why a typed workbook path is rejected, or "" if it looks usable.
'
' Pure, so every rejection reason is testable without Office. Written after a
' live run on 2026-07-29 died on a hand-typed path with VBA runtime error 52
' ("Bad file name or number") -- raw, unhandled, Debug/End, discarding 45
' instance keys the human had just confirmed one prompt at a time.
'
' The first check is the one that actually matters. A cloud-hosted deck reports
' its own Path as an https:// URL, so copying the deck's location -- the
' obvious move when asked where to put its data -- produces something no file
' API can open. That deserves its own message rather than a generic "bad path",
' because the fix (use the synced local folder instead) is not guessable from
' the error.
Public Function WorkbookPathProblem(candidate As String) As String
    Dim p As String
    p = Trim(candidate)

    If p = "" Then
        WorkbookPathProblem = "No path given."
        Exit Function
    End If

    If LCase(Left(p, 7)) = "http://" Or LCase(Left(p, 8)) = "https://" Then
        WorkbookPathProblem = "That's a web address, not a file path." & vbCrLf & vbCrLf & _
            "Cloud-hosted decks report their location as an https:// URL, so copying the deck's own path gives you something Excel can't open. Use the synced folder on this PC instead -- something like C:\Users\<you>\OneDrive - <org>\...\Data.xlsx"
        Exit Function
    End If

    ' A drive letter or a UNC share -- anything else has no folder to live in.
    Dim rooted As Boolean
    rooted = (Left(p, 2) = "\\")
    If Not rooted Then
        If Len(p) >= 3 Then
            If Mid(p, 2, 2) = ":\" Then rooted = True
        End If
    End If
    If Not rooted Then
        WorkbookPathProblem = "That isn't a full path. It needs to start with a drive (C:\...) or a network share (\\server\...)."
        Exit Function
    End If

    If LCase(Right(p, 5)) <> ".xlsx" And LCase(Right(p, 5)) <> ".xlsm" Then
        WorkbookPathProblem = "The Data workbook needs to end in .xlsx"
        Exit Function
    End If

    ' Characters Windows will not accept in a file name. The drive colon is
    ' already past us, so any remaining colon is invalid too.
    Dim bad As Variant, i As Long
    bad = Array("<", ">", """", "|", "?", "*")
    For i = LBound(bad) To UBound(bad)
        If InStr(p, CStr(bad(i))) > 0 Then
            WorkbookPathProblem = "That path contains a character Windows won't allow in a file name: " & CStr(bad(i))
            Exit Function
        End If
    Next i
    If InStr(3, p, ":") > 0 Then
        WorkbookPathProblem = "That path contains a stray colon."
        Exit Function
    End If

    WorkbookPathProblem = ""
End Function

' Which already-confirmed slide is using this instance key, or -1 if it's free.
' Index 0 is the template slide; 1..N index into otherSlides.
'
' The instance key is the identity that links a slide to its row in the Data
' sheet, and nothing checked for collisions until now (the gap was even named
' in ConflictingSlideType's header as "the missing duplicate-key guard"). Two
' slides given the same key both resolve to one row: at onboard the second
' slide's harvested values overwrite the first's, and from then on one row
' feeds two slides. Nothing errors. The deck simply starts showing numbers
' that belong to a different project, which is the worst class of bug this
' tool can have -- a reporting deck that is confidently wrong.
'
' Comparison is trimmed and case-insensitive. Two keys differing only by case
' or trailing space are a typo in every realistic scenario, and the cost of
' being wrong is asymmetric: refusing a deliberate case-distinct pair costs
' one rename, while accepting an accidental one costs silently merged data.
Public Function IndexUsingInstanceKey(confirmedKeys As Object, candidate As String) As Long
    IndexUsingInstanceKey = -1

    Dim needle As String
    needle = LCase(Trim(candidate))
    If needle = "" Then Exit Function ' blank means "skip this slide" -- never a clash

    Dim k As Variant
    For Each k In confirmedKeys.Keys
        If LCase(Trim(CStr(confirmedKeys(k)))) = needle Then
            IndexUsingInstanceKey = CLng(k)
            Exit Function
        End If
    Next k
End Function

' Human-facing name for a confirmedKeys index -- 0 is the template, 1..N index
' into otherSlides. Clash messages are useless with array indices in them; the
' human is looking at slide numbers.
Private Function SlideLabelForKeyIndex(idx As Long, templateSld As Object, otherSlides() As Object) As String
    SlideLabelForKeyIndex = "another slide"
    On Error Resume Next
    If idx = 0 Then
        SlideLabelForKeyIndex = "the template slide (Slide " & templateSld.SlideIndex & ")"
    Else
        SlideLabelForKeyIndex = "Slide " & otherSlides(idx).SlideIndex
    End If
    On Error GoTo 0
End Function

' Asks the human where the Data workbook should live, preferring a real file
' browser over a typed path.
'
' Verified 2026-07-29 against real PowerPoint rather than assumed: all four
' FileDialog modes are available in this host (1 File Open, 2 File Save,
' 3 Browse, 4 Browse), so the Save As browser is the normal path.
'
' The typed-path prompt is kept as a fallback purely because this add-in is
' meant to travel to work machines whose Office policy is unknown. There is
' deliberately no middle fallback: an earlier version picked a folder via
' FileDialog(4) and asked only for the file name, which reads like defence in
' depth and is not -- mode 4 comes off the same object as mode 2, so it is
' unavailable in exactly the situations mode 2 is. A fallback that fails
' whenever the thing it backs up fails is decoration.
Private Function AskForWorkbookPath(pres As Object) As String
    Dim suggested As String
    suggested = "SAAFE-Projects-Data.xlsx"

    ' Layer 1: a genuine Save As dialog -- no syntax in human hands at all.
    Dim fd As Object
    On Error Resume Next
    Set fd = Application.FileDialog(2) ' msoFileDialogSaveAs
    On Error GoTo 0
    If Not fd Is Nothing Then
        Dim picked As String
        picked = ""
        On Error Resume Next
        fd.Title = "Where should this deck's Data workbook live?"
        fd.InitialFileName = suggested
        If fd.Show = -1 Then picked = fd.SelectedItems(1)
        On Error GoTo 0
        If picked <> "" Then
            ' Some hosts return the name without the extension.
            If LCase(Right(picked, 5)) <> ".xlsx" And LCase(Right(picked, 5)) <> ".xlsm" Then picked = picked & ".xlsx"
            AskForWorkbookPath = picked
            Exit Function
        End If
        ' Shown and cancelled is a real answer -- don't fall through to a
        ' second prompt the human didn't ask for.
        AskForWorkbookPath = ""
        Exit Function
    End If

    ' Fallback: the original typed-path prompt, now with the trap named in the
    ' prompt itself rather than discovered through a runtime error.
    AskForWorkbookPath = InputBox( _
        "This deck has no paired workbook yet." & vbCrLf & vbCrLf & _
        "Enter a FULL path for the new Data workbook, e.g." & vbCrLf & _
        "C:\Users\you\OneDrive\Claude\SAAFE-Projects-Data.xlsx" & vbCrLf & vbCrLf & _
        "Not a web address -- an https:// link won't work here.", _
        "Bulk Onboard Type -- Pair Workbook", suggested)
End Function

' WorkbookBridge.CreateWorkbook raises rather than returning Nothing on a bad
' path (live, 2026-07-29: runtime error 52 straight to a Debug/End dialog).
' Wrapping it here means a bad path is a sentence the human can act on and a
' retry, not a dead run.
Private Function CreateWorkbookSafely(path As String, ByRef errText As String) As Object
    errText = ""
    On Error Resume Next
    Err.Clear
    Set CreateWorkbookSafely = WorkbookBridge.CreateWorkbook(path)
    If Err.Number <> 0 Then
        errText = "Windows rejected that path (" & Err.Description & ")."
        Set CreateWorkbookSafely = Nothing
    ElseIf CreateWorkbookSafely Is Nothing Then
        errText = "Could not create a workbook there."
    End If
    On Error GoTo 0
End Function

Private Function OpenWorkbookSafely(path As String, ByRef errText As String) As Object
    errText = ""
    On Error Resume Next
    Err.Clear
    Set OpenWorkbookSafely = WorkbookBridge.OpenOrGetWorkbook(path)
    If Err.Number <> 0 Then
        errText = "Could not open the paired workbook (" & Err.Description & ")."
        Set OpenWorkbookSafely = Nothing
    ElseIf OpenWorkbookSafely Is Nothing Then
        errText = "Could not open the paired workbook at: " & path
    End If
    On Error GoTo 0
End Function

' Establish (or reuse) the deck-workbook pairing, re-prompting on a bad path
' instead of dying. Returns Nothing only when the human cancels or gives up.
Private Function ResolveDataWorkbook(pres As Object, ByRef outPath As String, ByRef cancelMsg As String) As Object
    Set ResolveDataWorkbook = Nothing
    cancelMsg = ""
    outPath = ""

    Dim existing As String
    existing = DeckRegistry.GetWorkbookPath(pres)
    If existing <> "" Then
        Dim openErr As String
        Dim wbExisting As Object
        Set wbExisting = OpenWorkbookSafely(existing, openErr)
        If wbExisting Is Nothing Then
            cancelMsg = openErr & vbCrLf & "Paired path: " & existing
            Exit Function
        End If
        outPath = existing
        Set ResolveDataWorkbook = wbExisting
        Exit Function
    End If

    ' Bounded, so a dialog that somehow always fails can't trap the human in a
    ' modal loop with no way out but Task Manager.
    Dim attempt As Long
    For attempt = 1 To 5
        Dim candidate As String
        candidate = AskForWorkbookPath(pres)
        If Trim(candidate) = "" Then
            cancelMsg = "Cancelled -- no workbook path given. Nothing was written."
            Exit Function
        End If

        Dim problem As String
        problem = WorkbookPathProblem(candidate)
        If problem = "" Then
            Dim createErr As String
            Dim wbNew As Object
            Set wbNew = CreateWorkbookSafely(candidate, createErr)
            If Not wbNew Is Nothing Then
                DeckRegistry.SetWorkbookPath pres, candidate
                outPath = candidate
                Set ResolveDataWorkbook = wbNew
                Exit Function
            End If
            problem = createErr
        End If

        If MsgBox(problem & vbCrLf & vbCrLf & "Try a different location?", vbYesNo + vbExclamation, "Bulk Onboard Type -- Pair Workbook") <> vbYes Then
            cancelMsg = "Cancelled at the workbook step. Nothing was written."
            Exit Function
        End If
    Next attempt

    cancelMsg = "Gave up after 5 attempts at a workbook location. Nothing was written."
End Function

' Are the Shape references in the in-memory session still attached to a live
' document?
'
' This is the check that catches closing and reopening THE SAME deck, which
' the deck-identity comparison structurally cannot: DeckIdentity is the file's
' FullName, so a reopened deck reports the identical value and the session
' looks current when every Shape in it is dead. That is the exact scenario
' reported on 2026-07-28 ("marks never restored on reopen"), and the deck-id
' guard added that day only ever covered switching to a DIFFERENT deck.
'
' Counting cannot substitute for this. markedShapes is an ordinary VBA
' Collection -- it survives the document perfectly well and .Count keeps
' answering 1 long after the shape it holds has ceased to exist. Touching a
' property is the only operation that actually asks Office whether the
' reference is real, and a dead one raises there.
' Public only so TestRunner can assert the dead-reference probe against real
' Office rather than trusting that a stale Shape raises the way it's assumed to.
Public Function MarkedShapesStillLive() As Boolean
    MarkedShapesStillLive = False
    If markedShapes Is Nothing Then Exit Function
    If markedShapes.count = 0 Then Exit Function

    Dim probe As String
    On Error Resume Next
    Err.Clear
    probe = markedShapes(1).Name
    MarkedShapesStillLive = (Err.Number = 0)
    On Error GoTo 0
End Function

' Should a marking session be read back from the deck rather than reusing
' what's in memory?
'
' Pure, and separated from the probing above, for the same reason
' ShouldForceSave is: the previous version of this decision was four inline
' terms inside a MsgBox-driven Sub that no test could reach, so a condition
' that never fired in the most common real scenario looked fine indefinitely.
' Every branch below is a case that has actually occurred in live use.
Public Function NeedsSessionRestore(hasSession As Boolean, markedCount As Long, sessionDeckId As String, currentDeckId As String, shapesStillLive As Boolean) As Boolean
    If Not hasSession Then
        NeedsSessionRestore = True          ' nothing marked yet this add-in load
    ElseIf markedCount = 0 Then
        NeedsSessionRestore = True          ' session exists but is empty
    ElseIf sessionDeckId <> currentDeckId Then
        NeedsSessionRestore = True          ' the session belongs to a different deck
    ElseIf Not shapesStillLive Then
        NeedsSessionRestore = True          ' same deck, but closed and reopened -- refs are dead
    Else
        NeedsSessionRestore = False         ' a genuinely live session for this deck
    End If
End Function

' Returns the slide_type this slide ALREADY belongs to, if that differs from
' `newType` -- otherwise "". Onboarding slides that already carry a different
' registered type silently re-tags them, which strands the old type's entire
' Data sheet: its rows keep existing but no slide answers to them any more, and
' PlanRoutineSync classifies a row with no slide as new_record. One Sync Now
' later, that is a duplicated template slide for every orphaned row.
'
' Found live 2026-07-28: a practice bulk-onboard re-typed all 46 slides from
' "q" to "sandbox-test", and the very next preview reported "46 new slide(s)
' would be created" against the untouched q worksheet. Nothing had warned at
' onboard time. Same class as the missing duplicate-key guard.
Public Function ConflictingSlideType(sld As Object, newType As String) As String
    Dim inst As SlideInstance
    inst = Resolve.ResolveSlideInstance(sld)
    If inst.HasTypeTag Then
        If inst.TypeTag <> newType Then ConflictingSlideType = inst.TypeTag
    End If
End Function

Public Function SerializeCurrentMarkingSession() As String
    If markedShapes Is Nothing Then Exit Function
    If markedShapes.count = 0 Then Exit Function
    SerializeCurrentMarkingSession = SerializeMarkingSession(markedSlideId, markedShapes, markedNames, markedTypes, markedVolatility)
End Function

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

' One-line, human-recognisable preview of a field's current text, for prompts
' shown at mark time. Collapses the separators a multi-paragraph field carries
' -- TextRange.Text returns paragraphs CR-separated and soft breaks as Chr(11),
' both of which render as literal boxes inside an InputBox -- then truncates to
' FIELD_PREVIEW_CHARS. Display only: never a name, tag, or key.
Public Function FieldPreview(text As String) As String
    Dim s As String
    s = Replace(text, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, Chr(11), " ")

    ' Collapse runs of spaces so a field whose text starts with several empty
    ' paragraphs doesn't spend its whole preview budget on whitespace.
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    s = Trim(s)

    If Len(s) > FIELD_PREVIEW_CHARS Then
        s = RTrim(Left(s, FIELD_PREVIEW_CHARS)) & "..."
    End If
    FieldPreview = s
End Function

' A slide that already carries an instance_key tag is ALREADY LINKED: that key
' is the join between it and its Data-sheet row, and re-deriving one from field
' text is what orphaned 46 real slides from their rows on 2026-07-26. The
' second onboarding pass on that deck marked a body-text field first, so
' SuggestInstanceKey proposed a whole paragraph for every slide, CommitBatch
' overwrote each correct key with it, and UpsertRow -- finding no row under the
' new key -- appended 43 duplicate rows that nothing pointed at. Adding a field
' to an already-onboarded type must never re-key a slide.
'
' Returns "" for a genuinely new slide, which still gets the normal prompt.
Public Function ExistingInstanceKey(sld As Object) As String
    Dim inst As SlideInstance
    inst = Resolve.ResolveSlideInstance(sld)
    If inst.HasInstanceKey Then ExistingInstanceKey = inst.InstanceKey
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
' Human-driven field marking -- the click-based replacement for Discovery's
' whole-slide auto-enumeration
' ---------------------------------------------------------------------

' Recursively flattens `grp`'s members into a 1-based Object array
' (ReDim'd here), skipping intermediate group containers and returning only
' real leaf shapes -- mirrors Discovery.Walk's own group recursion
' (Discovery.bas), but returns live Shape references directly since no
' Candidate metadata is needed here. Used by MarkFieldForBatch's group-
' member picker (see that Sub's own header for why: Application.Selection
' can report the outer group even when the UI shows an individual member
' selected, so the human is asked to pick explicitly from this list rather
' than the tool guessing which one was clicked). Returns the leaf count.
Public Function FlattenGroupLeaves(grp As Object, ByRef leaves() As Object) As Long
    Dim count As Long
    count = 0
    Dim item As Object
    For Each item In grp.GroupItems
        If item.Type = msoGroup Then
            Dim subLeaves() As Object
            Dim subCount As Long
            subCount = FlattenGroupLeaves(item, subLeaves)
            Dim j As Long
            For j = 1 To subCount
                count = count + 1
                ReDim Preserve leaves(1 To count)
                Set leaves(count) = subLeaves(j)
            Next j
        Else
            count = count + 1
            ReDim Preserve leaves(1 To count)
            Set leaves(count) = item
        End If
    Next item
    FlattenGroupLeaves = count
End Function

' Pure-ish marking logic -- exercised directly by TestRunner.bas, same
' "thin interactive Sub + testable helper" split ResolveFields.bas already
' established. Adds `shp` (with
' `fieldName`, e.g. "Project Number" -- named at mark time per Rohan's
' direct feedback live-testing this against his real deck, rather than only
' renamable later in the Excel review grid) to the current marking session
' (module-level markedShapes/markedNames/markedSlideId, created on first
' use). Re-marking an already-marked shape renames it in place instead of
' duplicating it -- the natural way to fix a typo or a bad first guess
' without clearing the whole session. Returns a human-readable status to
' show the user, or the sentinel "DIFFERENT_SLIDE" when `shp` is on a
' different slide than the session already in progress -- resolving that is
' a genuine interactive decision (clear and restart, or cancel), left to
' MarkFieldForBatch's own MsgBox rather than handled here, since this
' function must stay callable without a live dialog for tests to drive it.
' `fieldType` is one of "text"|"number"|"currency"|"date" (see
' NormalizeFieldType) -- captured as metadata (shown/editable in the
' Field Review grid via WriteReviewGrid/ReadReviewGrid) but deliberately
' NOT applied to the actual Data-sheet cell format: see ExcelOutput.
' UpsertRow's own header for a real sync-correctness risk found 2026-07-26
' that ruled that out (a typed Excel value's read-back can differ from the
' exact harvested string, which risks a routine sync silently rewriting
' slide text into a reformatted version). Every field is still linked and
' synced exactly as harvested regardless of its declared type.
' `fieldVolatility` is "static"|"variable" (see NormalizeFieldVolatility) --
' a human-declared hint only, captured 2026-07-26 per Rohan's own framing
' ("I know what's variable vs static", so the tool should record his
' declaration rather than try to infer it). Deliberately not wired into any
' sync behavior yet (no locking, no flagging) -- whether a field is
' genuinely static can only really be known after watching it across real
' periods, and Rohan explicitly agreed to defer that until there's real
' data to act on rather than a Day-1 guess.
Public Function MarkShapeForBatch(shp As Object, fieldName As String, fieldType As String, fieldVolatility As String) As String
    ' Shape.Parent resolves directly to the containing Slide regardless of
    ' whether `shp` is top-level or nested inside a group -- confirmed live
    ' 2026-07-26 (a prior version of this function wrongly assumed .Parent
    ' returned the enclosing GroupShape for a nested shape and rejected
    ' those entirely; Rohan's real deck uses grouped "card" layouts for its
    ' fields, so that assumption -- never actually verified against real
    ' Office -- was a real, blocking bug, not a deliberate scope limit. A
    ' shape's immediate group container, if any, is exposed separately via
    ' .ParentGroup, which this function has no reason to use). Discovery.
    ' Walk already recurses into groups when building the Candidate list
    ' BuildBatchPlanFromMarkedFields matches a marked Shape back against
    ' (by object identity), so no other code needed to change for grouped
    ' fields to work end-to-end.
    ' A whole group container (msoGroup) has no text of its own -- its
    ' .HasTextFrame is False even when every member inside it has real text
    ' -- confirmed live 2026-07-26 after Rohan reported the "current value"
    ' preview showing empty for long text fields inside a grouped "card"
    ' layout: the actual cause was the whole GROUP being selected (one
    ' click), not the individual field shape (which needs a second click on
    ' the same spot to drill in). Silently accepting a group here would tag
    ' the wrong object entirely -- InjectPrimitive/verification expect the
    ' real text-bearing shape, not its container -- so it's rejected with
    ' guidance instead.
    If shp.Type = msoGroup Then
        MarkShapeForBatch = "That's the whole group, not a single field -- click the same spot again to select just the field shape inside it, then run this again."
        Exit Function
    End If

    Dim thisSlideId As Long
    thisSlideId = 0
    On Error Resume Next
    thisSlideId = shp.Parent.SlideID
    On Error GoTo 0
    If thisSlideId = 0 Then
        MarkShapeForBatch = "Could not determine this shape's slide -- select a real field shape and try again."
        Exit Function
    End If

    If markedShapes Is Nothing Then
        Set markedShapes = New Collection
        Set markedNames = CreateObject("Scripting.Dictionary")
        Set markedTypes = CreateObject("Scripting.Dictionary")
        Set markedVolatility = CreateObject("Scripting.Dictionary")
    End If

    If markedShapes.Count > 0 And thisSlideId <> markedSlideId Then
        MarkShapeForBatch = "DIFFERENT_SLIDE"
        Exit Function
    End If

    markedSlideId = thisSlideId
    markedDeckId = DeckIdentity(shp.Parent.Parent)

    Dim existingIdx As Long
    existingIdx = 0
    Dim i As Long
    For i = 1 To markedShapes.Count
        If markedShapes(i) Is shp Then existingIdx = i
    Next i

    Dim posIdx As Long
    If existingIdx > 0 Then
        posIdx = existingIdx
    Else
        markedShapes.Add shp
        posIdx = markedShapes.Count
    End If
    markedNames(posIdx) = fieldName
    markedTypes(posIdx) = fieldType
    markedVolatility(posIdx) = fieldVolatility

    MarkShapeForBatch = "Marked field " & posIdx & ": '" & fieldName & "' (" & fieldType & ", " & fieldVolatility & ", shape '" & shp.Name & "')." & vbCrLf & vbCrLf & "Click the next field's shape and run this again, or run 'Bulk Onboard Type' when done."
End Function

' Resolves an InputBox answer to a canonical field type -- accepts either
' the numbered-list index or the type name itself (case-insensitive), same
' idiom as ResolveFields.PickRoleFromList. Never errors: an unrecognized or
' blank answer falls back to "text", the safe no-op formatting choice
' (harmless here, unlike guessing at real harvested data, since this only
' ever affects a bonus NumberFormat, never the synced value itself).
Public Function NormalizeFieldType(answer As String) As String
    Select Case Trim(LCase(answer))
        Case "1", "text": NormalizeFieldType = "text"
        Case "2", "number": NormalizeFieldType = "number"
        Case "3", "currency": NormalizeFieldType = "currency"
        Case "4", "date": NormalizeFieldType = "date"
        Case Else: NormalizeFieldType = "text"
    End Select
End Function

' Same idiom as NormalizeFieldType. Falls back to "variable" (not "static")
' on a blank/unrecognized answer -- the safer default here too: an
' incorrectly-"variable"-tagged field is just a slightly-less-useful hint,
' while an incorrectly-"static"-tagged one is the kind of mistake that
' would matter if this hint ever gets wired into real behavior later.
Public Function NormalizeFieldVolatility(answer As String) As String
    Select Case Trim(LCase(answer))
        Case "1", "static": NormalizeFieldVolatility = "static"
        Case "2", "variable": NormalizeFieldVolatility = "variable"
        Case Else: NormalizeFieldVolatility = "variable"
    End Select
End Function

' Read-only testability hooks -- TestRunner.bas has no other way to observe
' this module's private marking-session state. Return 0/"" if no session is
' active or `position` is out of range.
Public Function MarkedFieldCountForBatch() As Long
    If markedShapes Is Nothing Then
        MarkedFieldCountForBatch = 0
    Else
        MarkedFieldCountForBatch = markedShapes.Count
    End If
End Function

Public Function MarkedFieldNameForBatch(position As Long) As String
    If markedNames Is Nothing Then Exit Function
    If markedNames.Exists(position) Then MarkedFieldNameForBatch = CStr(markedNames(position))
End Function

Public Function MarkedFieldTypeForBatch(position As Long) As String
    If markedTypes Is Nothing Then Exit Function
    If markedTypes.Exists(position) Then MarkedFieldTypeForBatch = CStr(markedTypes(position))
End Function

Public Function MarkedFieldVolatilityForBatch(position As Long) As String
    If markedVolatility Is Nothing Then Exit Function
    If markedVolatility.Exists(position) Then MarkedFieldVolatilityForBatch = CStr(markedVolatility(position))
End Function

' Click one field's shape on your template slide, then run this (toolbar
' button "Mark Field for Batch"). Repeats: click the next field, run again.
' VBA's InputBox/MsgBox are fully modal (block the whole application), so
' there is no way to show a "click now" prompt and have the user click the
' canvas while it's up -- each mark is therefore its own separate toolbar
' click, not a loop inside one macro run. Reuses ResolveFields.
' ValidateSingleShapeSelection unchanged (same "exactly one shape selected"
' requirement).
Public Sub MarkFieldForBatch()
    Dim shp As Object
    Dim selErr As String
    selErr = ResolveFields.ValidateSingleShapeSelection(Application.ActiveWindow.Selection, shp)
    If selErr <> "" Then
        RibbonUI.ShowSyncResult "Mark Field for Batch", selErr
        Exit Sub
    End If

    Dim pres As Object
    Set pres = Application.ActivePresentation

    ' Transparently pick up a marking session saved before a prior close,
    ' if nothing's in memory yet -- see this module's persistence header.
    ' Restored against `shp`'s own slide (not necessarily the persisted
    ' session's slide): if they don't match, RestoreMarkingSession simply
    ' finds nothing (0 shapes re-found on the wrong slide) and the normal
    ' marking flow below proceeds fresh, exactly as if nothing had been
    ' saved -- no separate mismatch-handling needed here.
    Dim restoreReport As String
    restoreReport = ""

    ' Restore when there is no USABLE session for THIS deck -- not merely when
    ' the collection has never been created. Closing a presentation does not
    ' unload the add-in, so a stale session (holding dead Shape references from
    ' the closed document) survives and used to suppress the restore entirely.
    ' See NeedsSessionRestore for each case, and MarkedShapesStillLive for why
    ' liveness has to be probed rather than inferred from the count or the
    ' deck's identity.
    Dim needRestore As Boolean
    needRestore = NeedsSessionRestore( _
        Not (markedShapes Is Nothing), _
        MarkedFieldCountForBatch(), _
        markedDeckId, _
        DeckIdentity(pres), _
        MarkedShapesStillLive())

    If needRestore Then
        Dim savedSession As String
        savedSession = ReadMarkingSessionProperty(pres)
        If Trim(savedSession) <> "" Then
            Dim restoreTemplateSld As Object
            Set restoreTemplateSld = Nothing
            On Error Resume Next
            Set restoreTemplateSld = shp.Parent
            On Error GoTo 0
            If Not restoreTemplateSld Is Nothing Then
                restoreReport = RestoreMarkingSession(savedSession, restoreTemplateSld)
                If markedShapes Is Nothing Then
                    restoreReport = ""
                ElseIf markedShapes.count = 0 Then
                    restoreReport = "" ' nothing usable restored (different slide, or shapes gone) -- don't clutter the confirmation with a no-op report
                End If
            End If
        End If
    End If

    ' A group selected on-screen doesn't reliably mean Application.Selection
    ' reports the individual member -- confirmed live 2026-07-26 (diagnostic
    ' build): even a real click-to-drill-in on the canvas (Shape Format
    ' ribbon tab active, tight selection handles around just the one field)
    ' can still leave Application.ActiveWindow.Selection.ShapeRange(1)
    ' reporting the outer group, not the shape the UI visually shows
    ' selected. Rather than keep guessing at click techniques, resolve it
    ' explicitly: flatten the group's members (recursing through any nested
    ' groups) and let the human pick the intended field by number from a
    ' list with a text preview -- same "numbered list, pick by number"
    ' idiom ResolveFields.BuildRolePickerPrompt already established.
    If shp.Type = msoGroup Then
        Dim leaves() As Object
        Dim leafCount As Long
        leafCount = FlattenGroupLeaves(shp, leaves)

        If leafCount = 0 Then
            RibbonUI.ShowSyncResult "Mark Field for Batch", "The group '" & shp.Name & "' has no field shapes inside it."
            Exit Sub
        End If

        Dim pickerPrompt As String
        pickerPrompt = "PowerPoint reported the whole group ('" & shp.Name & "') as selected, not a single field inside it -- pick which one you meant:" & vbCrLf
        Dim k As Long
        For k = 1 To leafCount
            Dim preview As String
            preview = ""
            If leaves(k).HasTextFrame Then
                If leaves(k).TextFrame.HasText Then preview = FieldPreview(leaves(k).TextFrame.TextRange.Text)
            End If
            pickerPrompt = pickerPrompt & k & ") " & leaves(k).Name & IIf(preview <> "", " -- '" & preview & "'", "") & vbCrLf
        Next k

        Dim pickAnswer As String
        pickAnswer = InputBox(pickerPrompt, "Mark Field for Batch -- Pick the Field")
        Dim pickedIdx As Long
        pickedIdx = 0
        If IsNumeric(pickAnswer) Then
            Dim candIdx As Long
            candIdx = CLng(pickAnswer)
            If candIdx >= 1 And candIdx <= leafCount Then pickedIdx = candIdx
        End If

        If pickedIdx = 0 Then
            RibbonUI.ShowSyncResult "Mark Field for Batch", "Cancelled -- no valid field number chosen."
            Exit Sub
        End If

        Set shp = leaves(pickedIdx)
    End If

    ' Re-marking an already-marked shape prefills its existing name (a quick
    ' rename); otherwise suggest one the same way OnboardFlow/BuildBatchPlan
    ' already do (reuse the shape's own ph_ name if it has one, else a
    ' positional fallback) -- either way the human can overwrite it with
    ' something real (e.g. "Project Number") right here, while looking at
    ' the field, instead of only being renamable later in the Excel grid.
    Dim existingIdx As Long
    existingIdx = 0
    If Not markedShapes Is Nothing Then
        Dim i As Long
        For i = 1 To markedShapes.Count
            If markedShapes(i) Is shp Then existingIdx = i
        Next i
    End If

    Dim defaultName As String
    If existingIdx > 0 Then
        defaultName = MarkedFieldNameForBatch(existingIdx)
    Else
        Dim nextOrdinal As Long
        nextOrdinal = 1
        If Not markedShapes Is Nothing Then nextOrdinal = markedShapes.Count + 1
        defaultName = SuggestBatchFieldName(shp, nextOrdinal)
    End If

    ' Previewed, not shown whole: an About-text field's real value is a 250+
    ' char paragraph, which made this prompt a wall of text (and rendered its
    ' paragraph CRs as literal boxes) on the real deck.
    Dim currentValue As String
    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then currentValue = FieldPreview(shp.TextFrame.TextRange.Text)
    End If

    Dim typedName As String
    typedName = InputBox("Field name for this shape (current value: '" & currentValue & "'):", "Mark Field for Batch", defaultName)
    If Trim(typedName) = "" Then
        RibbonUI.ShowSyncResult "Mark Field for Batch", "Cancelled -- no field name given."
        Exit Sub
    End If

    Dim defaultType As String
    If existingIdx > 0 Then
        defaultType = MarkedFieldTypeForBatch(existingIdx)
    Else
        defaultType = "text"
    End If

    Dim typedType As String
    typedType = InputBox("Field type -- reference metadata shown in the Field Review grid (never changes what's synced):" & vbCrLf & "1) Text  2) Number  3) Currency  4) Date", "Mark Field for Batch", defaultType)
    Dim fieldType As String
    fieldType = NormalizeFieldType(typedType)

    Dim defaultVolatility As String
    If existingIdx > 0 Then
        defaultVolatility = MarkedFieldVolatilityForBatch(existingIdx)
    Else
        defaultVolatility = "variable"
    End If

    Dim typedVolatility As String
    typedVolatility = InputBox("Does this field usually stay the same, or change each period? Hint only -- reference metadata shown in the Field Review grid (never locks or changes what's synced):" & vbCrLf & "1) Static  2) Variable", "Mark Field for Batch", defaultVolatility)
    Dim fieldVolatility As String
    fieldVolatility = NormalizeFieldVolatility(typedVolatility)

    Dim status As String
    status = MarkShapeForBatch(shp, Trim(typedName), fieldType, fieldVolatility)

    If status = "DIFFERENT_SLIDE" Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox(markedShapes.Count & " field(s) already marked on a different slide." & vbCrLf & "Clear them and start a new marking session on this slide?", vbYesNo + vbQuestion, "Mark Field for Batch")
        If resp <> vbYes Then Exit Sub
        Set markedShapes = New Collection
        Set markedNames = CreateObject("Scripting.Dictionary")
        Set markedTypes = CreateObject("Scripting.Dictionary")
        Set markedVolatility = CreateObject("Scripting.Dictionary")
        restoreReport = "" ' the restored session (if any) was for the slide we just abandoned -- don't report it alongside a fresh one
        status = MarkShapeForBatch(shp, Trim(typedName), fieldType, fieldVolatility)
    End If

    ' Persist the session after every successful mark, so it is always as
    ' durable as the deck itself. Whether that costs an explicit save or is
    ' left to AutoSave is SaveMarkingSessionToProperty's decision, not this
    ' caller's -- see its header.
    Dim saveWarning As String
    saveWarning = ""
    If InStr(status, "Marked field") > 0 Then
        saveWarning = SaveMarkingSessionToProperty(pres)

        ' State success POSITIVELY, not just failure. On an AutoSave cloud
        ' document Office's own indicators cannot be trusted (Presentation.Saved
        ' desyncs, the manual Save command is hidden), so the human has no native
        ' way to confirm their marks are stored.
        '
        ' Report the deck's own last-save timestamp rather than asserting
        ' "AutoSave will persist them". That assertion was a promise this code
        ' could not keep -- it was printed on exactly the deck later found 2.6
        ' hours stale. A timestamp is checkable: if it is minutes old, AutoSave
        ' is working; if it is hours old, the human can see that for themselves
        ' and hit Ctrl+S.
        If saveWarning = "" Then
            saveWarning = "Marks stored in the deck (write confirmed by read-back)." & vbCrLf & _
                "Deck last saved to disk: " & BatchOnboardFlow.LastSaveTimeTextOf(pres) & "."
        End If
    End If

    If restoreReport <> "" Then status = restoreReport & vbCrLf & vbCrLf & status
    If saveWarning <> "" Then status = status & vbCrLf & vbCrLf & saveWarning
    RibbonUI.ShowSyncResult "Mark Field for Batch", status
End Sub

' Pure reset -- exercised directly by TestRunner.bas so tests never have to
' call the MsgBox-showing Sub below just to reach a clean slate (that
' mistake shipped once already in this pass: an earlier version of
' Test_BatchOnboardFlow_MarkShapeForBatchAccumulatesAndDedupes called
' ClearMarkedFieldsForBatch directly, which popped a real modal MsgBox
' during what was supposed to be an unattended automated run -- it only
' "passed" because Rohan was watching PowerPoint live and clicked through
' it by hand. Fixed by giving tests this MsgBox-free path instead).
Public Sub ResetMarkingSession()
    Set markedShapes = Nothing
    Set markedNames = Nothing
    Set markedTypes = Nothing
    Set markedVolatility = Nothing
    markedSlideId = 0
    markedDeckId = ""
    ' Also clears the persisted CustomDocumentProperty (see this module's
    ' persistence header) so a deliberate reset never silently resurrects
    ' on the next reopen. On Error Resume Next inside ClearMarkingSession
    ' Property already makes this a safe no-op if nothing was ever saved,
    ' or if there's no active presentation (test contexts always have one).
    ClearMarkingSessionProperty Application.ActivePresentation
End Sub

' Discards the current marking session (misclick recovery) -- no confirm
' prompt, since re-marking is cheap and a silent no-op when nothing was
' marked yet is the least surprising behavior.
Public Sub ClearMarkedFieldsForBatch()
    ResetMarkingSession
    RibbonUI.ShowSyncResult "Clear Marked Fields", "Marked fields cleared."
End Sub

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

' Builds the full BatchOnboardPlan by discovering the template's candidate
' fields via Discovery + Onboarding.IsCandidateField (every text/picture
' shape on the slide), then delegating to BuildBatchPlanFromCandidates for
' the actual cross-slide correspondence/harvesting. Kept for the tests that
' already exercise it directly; the live "Bulk Onboard Type" ribbon entry
' point uses BuildBatchPlanFromMarkedFields instead (see that function's own
' header) -- Discovery-based auto-enumeration produced an 87-row, unreviewable
' grid on Rohan's real deck (60-90 candidate shapes/slide), so the live flow
' now requires a human to click each field shape first rather than walking
' every shape on the slide.
Public Function BuildBatchPlan(templateSld As Object, otherSlides() As Object) As BatchOnboardPlan
    Dim allCandidates() As Candidate
    Dim allShapes() As Object
    allCandidates = Discovery.DiscoverSlideWithShapes(templateSld, allShapes)

    Dim aLo As Long, aHi As Long, hasAny As Boolean
    On Error Resume Next
    aLo = LBound(allCandidates): aHi = UBound(allCandidates)
    hasAny = (Err.Number = 0)
    On Error GoTo 0

    If Not hasAny Then
        Dim emptyPlan As BatchOnboardPlan
        BuildBatchPlan = emptyPlan
        Exit Function
    End If

    Dim templateCandidates() As Candidate
    Dim templateShapes() As Object
    Dim n As Long
    n = 0
    Dim i As Long
    For i = aLo To aHi
        If Onboarding.IsCandidateField(allCandidates(i)) Then
            n = n + 1
            ReDim Preserve templateCandidates(1 To n)
            ReDim Preserve templateShapes(1 To n)
            templateCandidates(n) = allCandidates(i)
            Set templateShapes(n) = allShapes(i)
        End If
    Next i

    BuildBatchPlan = BuildBatchPlanFromCandidates(templateCandidates, templateShapes, otherSlides)
End Function

' Same cross-slide correspondence/harvesting as BuildBatchPlan, but the
' template's field list is exactly the human-clicked shapes recorded by
' MarkFieldForBatch (in click order) instead of every text/picture shape
' Discovery finds -- built 2026-07-26 per Rohan's direct feedback on the
' first live run against his real deck ("very hard for a human to interpret
' ... better human led by selection then you find the matching value and
' position"). Discovery.DiscoverSlideWithShapes is still run once, purely to
' get a Candidate for each marked Shape (no public per-shape Candidate
' builder exists in Discovery.bas) -- matched back to the marked Shape by
' object identity (`Is`), never by name or position, since names collide and
' positions can legitimately match by coincidence. `markedFieldNames`/
' `markedFieldTypes`/`markedFieldVolatility` (Scripting.Dictionary: 1-based
' position in markedFields -> String) are the human-typed name and declared
' type/volatility given at mark time (MarkFieldForBatch/MarkShapeForBatch)
' -- override BuildBatchPlanFromCandidates' own SuggestBatchFieldName guess
' and "text"/"variable" defaults rather than leaving every field only
' editable later in the Excel review grid.
Public Function BuildBatchPlanFromMarkedFields(templateSld As Object, markedFields As Collection, markedFieldNames As Object, markedFieldTypes As Object, markedFieldVolatility As Object, otherSlides() As Object, ByRef matchErr As String) As BatchOnboardPlan
    matchErr = ""

    Dim allCandidates() As Candidate
    Dim allShapes() As Object
    allCandidates = Discovery.DiscoverSlideWithShapes(templateSld, allShapes)

    Dim aLo As Long, aHi As Long, hasAny As Boolean
    On Error Resume Next
    aLo = LBound(allCandidates): aHi = UBound(allCandidates)
    hasAny = (Err.Number = 0)
    On Error GoTo 0

    Dim templateCandidates() As Candidate
    Dim templateShapes() As Object
    Dim n As Long
    n = 0

    Dim marked As Variant
    For Each marked In markedFields
        Dim foundIdx As Long
        foundIdx = 0
        If hasAny Then
            Dim i As Long
            For i = aLo To aHi
                If allShapes(i) Is marked Then
                    foundIdx = i
                    Exit For
                End If
            Next i
        End If

        If foundIdx = 0 Then
            matchErr = "Marked shape '" & marked.Name & "' could not be re-found on the template slide (was it deleted or moved into/out of a group after marking?). Clear marked fields and mark again."
            Dim emptyPlan2 As BatchOnboardPlan
            BuildBatchPlanFromMarkedFields = emptyPlan2
            Exit Function
        End If

        n = n + 1
        ReDim Preserve templateCandidates(1 To n)
        ReDim Preserve templateShapes(1 To n)
        templateCandidates(n) = allCandidates(foundIdx)
        Set templateShapes(n) = allShapes(foundIdx)
    Next marked

    Dim plan As BatchOnboardPlan
    plan = BuildBatchPlanFromCandidates(templateCandidates, templateShapes, otherSlides)

    Dim idx As Long
    For idx = 1 To plan.FieldCount
        If markedFieldNames.Exists(idx) Then
            plan.FieldNames(idx) = CStr(markedFieldNames(idx))
        End If
        If markedFieldTypes.Exists(idx) Then
            plan.FieldTypes(idx) = CStr(markedFieldTypes(idx))
        End If
        If markedFieldVolatility.Exists(idx) Then
            plan.FieldVolatility(idx) = CStr(markedFieldVolatility(idx))
        End If
    Next idx

    BuildBatchPlanFromMarkedFields = plan
End Function

' The shared correspondence engine both BuildBatchPlan and
' BuildBatchPlanFromMarkedFields delegate to: for each already-chosen
' template field (position in templateCandidates()/templateShapes(), no
' further filtering here), finds its best-corresponding shape on every other
' slide via Matching.Match (pure tier-2 geometry scoring -- nothing is
' tagged yet, so tier-1 trust-the-tag never applies here) and harvests its
' current text. A shape already claimed by an earlier field on the same
' slide is excluded from later fields' candidate pool on that slide, so two
' fields can never both grab the same shape just because they scored
' similarly.
Private Function BuildBatchPlanFromCandidates(templateCandidates() As Candidate, templateShapes() As Object, otherSlides() As Object) As BatchOnboardPlan
    Dim plan As BatchOnboardPlan
    Set plan.FieldNames = CreateObject("Scripting.Dictionary")
    Set plan.FieldTypes = CreateObject("Scripting.Dictionary")
    Set plan.FieldVolatility = CreateObject("Scripting.Dictionary")
    Set plan.FieldTemplateShapes = CreateObject("Scripting.Dictionary")
    Set plan.FieldSuggestIdentical = CreateObject("Scripting.Dictionary")
    Set plan.FieldInclude = CreateObject("Scripting.Dictionary")
    Set plan.Correspondence = CreateObject("Scripting.Dictionary")
    Set plan.HarvestedText = CreateObject("Scripting.Dictionary")

    Dim tLo As Long, tHi As Long, hasTemplateCandidates As Boolean
    On Error Resume Next
    tLo = LBound(templateCandidates): tHi = UBound(templateCandidates)
    hasTemplateCandidates = (Err.Number = 0)
    On Error GoTo 0

    If Not hasTemplateCandidates Then
        plan.FieldCount = 0
        BuildBatchPlanFromCandidates = plan
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
        fieldIdx = fieldIdx + 1
            plan.FieldNames(fieldIdx) = SuggestBatchFieldName(templateShapes(i), fieldIdx)
            plan.FieldTypes(fieldIdx) = "text" ' BuildBatchPlan (Discovery-based) has no type info; BuildBatchPlanFromMarkedFields overrides this with the human-declared type
            plan.FieldVolatility(fieldIdx) = "variable" ' same reasoning as FieldTypes above
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
    Next i

    plan.FieldCount = fieldIdx
    BuildBatchPlanFromCandidates = plan
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
    ws.Cells(1, COL_TYPE).Value = "Type (text/number/currency/date)"
    ws.Cells(1, COL_VOLATILITY).Value = "Static/Variable (hint only)"

    Dim r As Long
    Dim fieldIdx As Long
    For fieldIdx = 1 To plan.FieldCount
        r = fieldIdx + 1
        ws.Cells(r, COL_FIELD_ID).Value = fieldIdx
        ws.Cells(r, COL_FIELD_NAME).Value = plan.FieldNames(fieldIdx)
        ws.Cells(r, COL_SUGGESTED).Value = IIf(plan.FieldSuggestIdentical(fieldIdx), "Decoration (identical everywhere)", "Field (varies)")
        ws.Cells(r, COL_INCLUDE).Value = IIf(plan.FieldInclude(fieldIdx), "Y", "N")
        ws.Cells(r, COL_TEMPLATE_VALUE).Value = plan.HarvestedText(FieldSlideKey(fieldIdx, 0))
        ws.Cells(r, COL_TYPE).Value = plan.FieldTypes(fieldIdx)
        ws.Cells(r, COL_VOLATILITY).Value = plan.FieldVolatility(fieldIdx)

        ' samples must be reset every iteration -- VBA's Dim inside a loop
        ' does NOT re-initialize the variable each pass (it's the same
        ' procedure-scoped variable throughout, Dim is only elaborated
        ' once); without this explicit reset, real bug found live 2026-07-26
        ' against Rohan's real deck, every row's samples accumulated on top
        ' of every prior field's, so field 9's cell ended up containing all
        ' 9 fields' samples concatenated together.
        Dim samples As String
        samples = ""
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
    ws.Columns(COL_TYPE).ColumnWidth = 16
    ws.Columns(COL_VOLATILITY).ColumnWidth = 20
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

        plan.FieldTypes(fieldIdx) = NormalizeFieldType(CStr(ws.Cells(r, COL_TYPE).Value))
        plan.FieldVolatility(fieldIdx) = NormalizeFieldVolatility(CStr(ws.Cells(r, COL_VOLATILITY).Value))
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
                    ' plan.FieldTypes(fieldIdx) (captured at mark time,
                    ' shown/editable in the Field Review grid) is
                    ' deliberately NOT passed to UpsertRow -- see that
                    ' function's own header for the real sync-correctness
                    ' risk found 2026-07-26 that ruled out applying it here.
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

' Pure logic -- exercised directly by TestRunner.bas. Every other slide in
' `templateSld`'s presentation sharing its CustomLayout (by Name -- same
' robustness reasoning as MarkShapeForBatch's shape-identity comparison:
' cheap, predictable, no risk of a stale wrapper-object comparison), in
' deck order, excluding templateSld itself. Layout, not any content/
' geometry heuristic -- this project's whole "child deck per slide type"
' authoring model (DECISIONS.md, 2026-07-26) means every slide built from
' the same layout genuinely IS the same type, a much safer signal than
' guessing from shape positions the way cross-slide field correspondence
' already does for a DIFFERENT purpose.
Public Function FindSameLayoutSlides(templateSld As Object) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim templateLayoutName As String
    templateLayoutName = templateSld.CustomLayout.Name

    Dim sld As Object
    For Each sld In templateSld.Parent.Slides
        If sld.SlideID <> templateSld.SlideID Then
            If sld.CustomLayout.Name = templateLayoutName Then result.Add sld
        End If
    Next sld

    Set FindSameLayoutSlides = result
End Function

' Interactive wrapper around FindSameLayoutSlides -- switches to Slide
' Sorter (the one view Slides.Range(...).Select is confirmed to register
' reliably in -- see AdoptFlow's own SPIKE_NOTES for the Normal-view
' finding this sidesteps rather than risks) and selects the template plus
' every same-layout sibling, then asks the human to confirm. Returns "" on
' confirmed proceed (the caller re-reads Application.ActiveWindow.Selection
' itself, same as the manual-selection path) or a message to show/return
' on decline/nothing-found -- never raises, same "readable status, not an
' exception" convention as every other entry point here.
Private Function AutoSelectSameLayoutSlides(templateSld As Object) As String
    Dim siblings As Collection
    Set siblings = FindSameLayoutSlides(templateSld)

    If siblings.count = 0 Then
        AutoSelectSameLayoutSlides = "No other slides share the template's layout ('" & templateSld.CustomLayout.Name & "') -- select at least one instance slide yourself (Slide Sorter or Ctrl-click, template first) and run 'Bulk Onboard Type' again."
        Exit Function
    End If

    Dim indices() As Long
    ReDim indices(1 To siblings.count + 1)
    indices(1) = templateSld.SlideIndex
    Dim i As Long
    Dim summary As String
    For i = 1 To siblings.count
        indices(i + 1) = siblings(i).SlideIndex
        If summary <> "" Then summary = summary & ", "
        summary = summary & siblings(i).SlideIndex
    Next i

    Application.ActiveWindow.ViewType = ppViewSlideSorter
    templateSld.Parent.Slides.Range(indices).Select

    Dim resp As VbMsgBoxResult
    resp = MsgBox("Auto-selected " & siblings.count & " other slide(s) using the same layout as your template (slide " & templateSld.SlideIndex & ", layout '" & templateSld.CustomLayout.Name & "'): " & summary & vbCrLf & vbCrLf & "Click Yes to continue with this batch, or No to select a different batch yourself (Slide Sorter/Ctrl-click, template first) and run 'Bulk Onboard Type' again.", vbYesNo + vbQuestion, "Bulk Onboard Type -- Auto-Selected Slides")

    If resp <> vbYes Then
        AutoSelectSameLayoutSlides = "Cancelled -- adjust the slide selection yourself and run 'Bulk Onboard Type' again."
    End If
End Function

' Suggests a default instance key for `slideIdx` (0 = template, 1..N =
' otherSlides) by reusing the FIRST marked field's harvested value for
' that slide -- typically a natural per-slide identifier (e.g. "Project
' Number"), since it's usually the field marked first. Cuts the instance-
' key prompts down to a single OK click in the common case; the human can
' still edit it for genuine ambiguity (e.g. a manually duplicated slide
' needing an extra "-Q3"/"-Q4" suffix). Returns "" if unavailable, falling
' back to today's blank-prompt behavior.
Public Function SuggestInstanceKey(plan As BatchOnboardPlan, slideIdx As Long) As String
    Dim key As String
    key = FieldSlideKey(1, slideIdx)
    If plan.HarvestedText.Exists(key) Then SuggestInstanceKey = CStr(plan.HarvestedText(key))
End Function

Public Function PromptBatchOnboardType() As String
    Dim pres As Object
    Set pres = Application.ActivePresentation

    If markedShapes Is Nothing Then
        PromptBatchOnboardType = "No fields marked yet. Click each field's shape on your template slide and run 'Mark Field for Batch' for each one before running this."
        Exit Function
    End If
    If markedShapes.Count = 0 Then
        PromptBatchOnboardType = "No fields marked yet. Click each field's shape on your template slide and run 'Mark Field for Batch' for each one before running this."
        Exit Function
    End If

    Dim markedTemplateSld As Object
    Set markedTemplateSld = pres.Slides.FindBySlideID(markedSlideId)
    If markedTemplateSld Is Nothing Then
        PromptBatchOnboardType = "The slide your marked fields are on no longer exists. Clear marked fields and mark again."
        Exit Function
    End If

    ' If the human hasn't already made an explicit multi-slide selection
    ' themselves, auto-detect and select every other slide sharing the
    ' template's layout instead of requiring one -- Rohan's own framing
    ' (2026-07-26): "same-layout slides auto-included, with a chance to
    ' review/edit before committing." An explicit selection already in
    ' place (a deliberately different batch, or a retry after declining
    ' the auto-selection) is always used as-is, never overridden.
    Dim hasExplicitSelection As Boolean
    hasExplicitSelection = False
    If Application.ActiveWindow.Selection.Type = ppSelectionSlides Then
        If Application.ActiveWindow.Selection.SlideRange.count >= 2 Then hasExplicitSelection = True
    End If

    If Not hasExplicitSelection Then
        Dim autoErr As String
        autoErr = AutoSelectSameLayoutSlides(markedTemplateSld)
        If autoErr <> "" Then
            PromptBatchOnboardType = autoErr
            Exit Function
        End If
    End If

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

    If markedSlideId <> templateSld.SlideID Then
        PromptBatchOnboardType = "Marked fields are on a different slide than Slide " & templateSld.SlideIndex & " (the earliest slide in your selection, used as the template). Include the slide you marked fields on, and make sure it is the first slide -- by deck order -- in your selection."
        Exit Function
    End If

    Dim otherSlideCount As Long
    otherSlideCount = hi - lo ' hi-lo+1 total slides, minus the template
    Dim otherSlides() As Object
    ReDim otherSlides(1 To otherSlideCount)
    Dim i As Long
    For i = 1 To otherSlideCount
        Set otherSlides(i) = slides(lo + i)
    Next i

    Dim matchErr As String
    Dim plan As BatchOnboardPlan
    plan = BuildBatchPlanFromMarkedFields(templateSld, markedShapes, markedNames, markedTypes, markedVolatility, otherSlides, matchErr)

    If matchErr <> "" Then
        PromptBatchOnboardType = matchErr
        Exit Function
    End If

    If plan.FieldCount = 0 Then
        PromptBatchOnboardType = "No candidate fields found -- nothing to onboard."
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

    ' Pair the workbook BEFORE asking for a single instance key.
    '
    ' This used to sit after the key loop, and that ordering is what turned one
    ' mistyped path into a genuinely expensive failure on 2026-07-29: 45 keys
    ' confirmed one prompt at a time, then the riskiest step in the flow ran for
    ' the first time, raised an unhandled error 52, and the End button threw all
    ' of it away. A path browser stops that particular typo; doing the fragile
    ' step first is what stops the whole class, because now the only thing at
    ' risk when it fails is the click that started the run.
    '
    ' General rule this flow should keep following: collect cheap-to-redo input
    ' after the steps that can fail, never before.
    Dim workbookPath As String
    Dim wb As Object
    Dim pairMsg As String
    Set wb = ResolveDataWorkbook(pres, workbookPath, pairMsg)
    If wb Is Nothing Then
        PromptBatchOnboardType = pairMsg
        Exit Function
    End If

    ' Instance-key prompts: template first (required), then each other
    ' slide (blank = skip, never guessed) -- same InputBox convention
    ' AdoptFlow.bas already established for exactly this decision. Each
    ' prompt is pre-filled with a suggested key (see SuggestInstanceKey) so
    ' confirming is a single OK click in the common case -- Rohan's own
    ' feedback (2026-07-26): hand-typing every key, not the instance-key
    ' concept itself, was the actual friction.
    Dim confirmedKeys As Object
    Set confirmedKeys = CreateObject("Scripting.Dictionary")

    ' A blank suggestion (SuggestInstanceKey returning "") looks visually
    ' IDENTICAL to every other prompt's pre-filled box -- real gap found
    ' live 2026-07-26: Rohan, having learned to trust the pre-fill and just
    ' click OK, hit a genuinely blank suggestion on the required template
    ' prompt and silently cancelled the whole run without realizing that
    ' one prompt was different. Making the missing-suggestion case visually
    ' distinct (extra warning text) so a blind click-through is far less
    ' likely, rather than relying on the human to notice an empty textbox
    ' that otherwise looks the same as a pre-filled one.
    ' Already-linked slides keep their own key and are never prompted for --
    ' see ExistingInstanceKey's header for the real corruption this prevents.
    ' Reused SILENTLY rather than pre-filling the prompt: re-onboarding this
    ' deck would otherwise mean 46 confirmation clicks, which is precisely the
    ' friction that drove this whole batch flow. The tradeoff is that a key
    ' cannot be CHANGED from here once set; the count is reported below so a
    ' reuse is never invisible.
    Dim reusedCount As Long

    Dim templateExisting As String
    templateExisting = ExistingInstanceKey(templateSld)
    If templateExisting <> "" Then
        confirmedKeys(0) = templateExisting
        reusedCount = reusedCount + 1
    Else
        Dim templateSuggestion As String
        templateSuggestion = SuggestInstanceKey(plan, 0)

        Dim templatePrompt As String
        templatePrompt = "Instance key for the template slide (Slide " & templateSld.SlideIndex & ") -- required, this slide defines the type:"
        If templateSuggestion = "" Then
            templatePrompt = templatePrompt & vbCrLf & vbCrLf & "(No suggested value available -- type one yourself. Leaving this blank cancels the whole run.)"
        End If

        Dim templateKey As String
        templateKey = InputBox(templatePrompt, "Bulk Onboard Type -- Instance Key", templateSuggestion)
        If Trim(templateKey) = "" Then
            PromptBatchOnboardType = "Cancelled -- the template slide must have an instance key."
            Exit Function
        End If
        confirmedKeys(0) = Trim(templateKey)
    End If

    ' Keys already written into slides are reused silently, but a collision
    ' AMONG them is reported rather than fixed here -- those keys are already in
    ' the deck, so a clash is pre-existing damage (exactly what the 2026-07-28
    ' re-derivation bug produced), and quietly re-prompting would hide it. The
    ' human needs to know their deck has two slides claiming one row.
    Dim reusedClashes As String
    reusedClashes = ""

    For i = 1 To otherSlideCount
        Dim otherExisting As String
        otherExisting = ExistingInstanceKey(otherSlides(i))
        If otherExisting <> "" Then
            Dim reusedClashIdx As Long
            reusedClashIdx = IndexUsingInstanceKey(confirmedKeys, otherExisting)
            If reusedClashIdx >= 0 Then
                reusedClashes = reusedClashes & vbCrLf & "  Slide " & otherSlides(i).SlideIndex & _
                    " and " & SlideLabelForKeyIndex(reusedClashIdx, templateSld, otherSlides) & _
                    " both already carry the key '" & otherExisting & "'"
            End If
            confirmedKeys(i) = otherExisting
            reusedCount = reusedCount + 1
        Else
            Dim otherSuggestion As String
            otherSuggestion = SuggestInstanceKey(plan, i)

            Dim prompt As String
            prompt = "Instance key for Slide " & otherSlides(i).SlideIndex & " (leave blank to skip this slide this pass):"
            If otherSuggestion = "" Then
                prompt = prompt & vbCrLf & vbCrLf & "(No suggested value available -- leaving this blank will skip this slide entirely this pass.)"
            End If

            ' Re-prompt on a clash rather than accepting it. Bounded, so a
            ' human who cannot find a free key can still get out -- and the way
            ' out is a blank, which skips the slide and leaves the deck
            ' untouched rather than merging it onto someone else's row.
            Dim proposed As String
            Dim tries As Long
            Dim thisPrompt As String
            thisPrompt = prompt
            For tries = 1 To 5
                proposed = Trim(InputBox(thisPrompt, "Bulk Onboard Type -- Instance Key", otherSuggestion))
                If proposed = "" Then Exit For ' blank = skip this slide

                Dim clashIdx As Long
                clashIdx = IndexUsingInstanceKey(confirmedKeys, proposed)
                If clashIdx < 0 Then Exit For ' free -- take it

                thisPrompt = "'" & proposed & "' is already used by " & _
                    SlideLabelForKeyIndex(clashIdx, templateSld, otherSlides) & "." & vbCrLf & vbCrLf & _
                    "Two slides sharing an instance key both point at the same row in the Data sheet, so one slide's numbers would quietly overwrite the other's." & vbCrLf & vbCrLf & _
                    prompt
                otherSuggestion = proposed
            Next tries

            ' Fell out of the loop still clashing -- refuse rather than merge.
            If proposed <> "" Then
                If IndexUsingInstanceKey(confirmedKeys, proposed) >= 0 Then proposed = ""
            End If

            confirmedKeys(i) = proposed
        End If
    Next i

    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(wb, WorkbookBridge.SanitizeSheetName(slideType))
    If IsEmpty(ws.Cells(1, 1).Value) Then
        ExcelOutput.CreateSheet ws, DeckRegistry.GetOrCreateDeckId(pres)
    End If

    DeckRegistry.RegisterType pres, slideType, templateSld, ws.Name

    ' Refuse to silently strand another type's dataset -- see
    ' ConflictingSlideType's header for what this prevents.
    Dim cflType As String, cflCount As Long, cflIdx As Long, cflOne As String
    cflType = "": cflCount = 0
    cflOne = ConflictingSlideType(templateSld, slideType)
    If cflOne <> "" Then
        cflCount = cflCount + 1
        cflType = cflOne
    End If
    For cflIdx = 1 To otherSlideCount
        cflOne = ConflictingSlideType(otherSlides(cflIdx), slideType)
        If cflOne <> "" Then
            cflCount = cflCount + 1
            If cflType = "" Then cflType = cflOne
        End If
    Next cflIdx

    If cflCount > 0 Then
        If MsgBox(cflCount & " of these slides already belong to slide type '" & cflType & "'." & vbCrLf & vbCrLf & _
            "Onboarding them as '" & slideType & "' re-tags them. '" & cflType & "' keeps its rows but loses its slides, and a later Sync Now would DUPLICATE the template slide once for EVERY orphaned row." & vbCrLf & vbCrLf & _
            "Continue anyway?", vbYesNo + vbExclamation, "Bulk Onboard Type -- Slides Already Have a Type") <> vbYes Then
            PromptBatchOnboardType = "Cancelled -- these slides already belong to type '" & cflType & "'. Nothing changed."
            Exit Function
        End If
    End If

    Dim commitResult As BatchCommitResult
    commitResult = CommitBatch(plan, templateSld, otherSlides, otherSlideCount, slideType, ws, confirmedKeys)

    ResetMarkingSession

    ' Force explicit, synchronous saves of both the deck (real role tags just
    ' written) and the Data workbook (real harvested values just written).
    '
    ' UNCONDITIONAL here, unlike SaveMarkingSessionToProperty's staleness-gated
    ' save, and that difference is deliberate rather than an oversight. The
    ' marking path runs after every single click, so a forced save there is
    ' both frequent and cheap to defer -- worth gating, to leave Office's own
    ' Save command working. This path runs once, on an explicit human "commit",
    ' and writes the real linked data. Waiting up to AUTOSAVE_STALE_SECONDS to
    ' find out whether AutoSave felt like handling it is the wrong trade for
    ' the one write in this add-in that actually matters.
    Dim pptSaveWarning As String, xlSaveWarning As String
    pptSaveWarning = "": xlSaveWarning = ""
    On Error Resume Next
    Err.Clear
    pres.Save
    If Err.Number <> 0 Then pptSaveWarning = "WARNING: could not save the deck (" & Err.Description & ") -- save it manually now."
    Err.Clear
    wb.Save
    If Err.Number <> 0 Then xlSaveWarning = "WARNING: could not save the Data workbook (" & Err.Description & ") -- save it manually now."
    On Error GoTo 0

    ' Same permanent closed-loop verification as SaveMarkingSessionToProperty
    ' (see that function's own header for why Err.Number/Saved alone aren't
    ' enough): if the deck save itself didn't already warn, confirm the type
    ' registration this commit just wrote (DeckRegistry.RegisterType, above)
    ' genuinely reads back through DeckRegistry's own public lookup -- this
    ' is the real, durable link between the deck and its Data sheet, so it
    ' matters more here than anywhere else in this flow.
    If pptSaveWarning = "" Then
        Dim verifyTemplateSld As Object, verifyWsName As String
        If Not DeckRegistry.LookupType(pres, slideType, verifyTemplateSld, verifyWsName) Then
            pptSaveWarning = "WARNING: the deck was saved but its '" & slideType & "' type registration did not read back correctly afterward -- save it manually now and re-check before closing PowerPoint."
        ElseIf verifyWsName <> ws.Name Then
            pptSaveWarning = "WARNING: the deck was saved but its '" & slideType & "' type registration read back a different worksheet name (expected '" & ws.Name & "', got '" & verifyWsName & "') -- save it manually now and re-check before closing PowerPoint."
        End If
    End If

    Dim report As String
    report = "Linked: " & commitResult.LinkedCount & vbCrLf & _
        "Skipped (no instance key given): " & commitResult.SkippedCount & vbCrLf & _
        "FAILED verification: " & commitResult.FailedVerificationCount

    ' Never leave a silent reuse invisible -- this is the difference between
    ' "added a field to slides that keep their existing rows" and "re-keyed
    ' the deck", and the human needs to be able to tell those apart.
    If reusedCount > 0 Then
        report = report & vbCrLf & "Kept existing instance key (already linked, not re-keyed): " & reusedCount
    End If

    ' Pre-existing duplicates among keys already written into the deck. Not
    ' something this run caused or can safely fix, but the human is now looking
    ' at a deck where two slides claim one row of data -- silence here would let
    ' that sit until the numbers went wrong on a slide nobody was watching.
    If reusedClashes <> "" Then
        report = report & vbCrLf & vbCrLf & _
            "WARNING -- duplicate instance keys ALREADY in this deck (not created by this run):" & reusedClashes & vbCrLf & _
            "Each pair shares one row in the Data sheet, so one slide's values overwrite the other's. Re-key one slide of each pair."
    End If

    If commitResult.FailedVerificationCount > 0 Then
        Dim m As Long
        report = report & vbCrLf & "Failed slides (harvest bug this pass -- fix before re-running):"
        For m = 1 To commitResult.FailedVerificationCount
            report = report & vbCrLf & "  " & commitResult.FailedVerificationLabels(m)
        Next m
    End If

    If pptSaveWarning <> "" Then report = report & vbCrLf & vbCrLf & pptSaveWarning
    If xlSaveWarning <> "" Then report = report & vbCrLf & vbCrLf & xlSaveWarning

    PromptBatchOnboardType = report
End Function

Public Sub BatchOnboardType()
    Dim report As String
    report = PromptBatchOnboardType()
    If report <> "" Then
        RibbonUI.ShowSyncResult "Bulk Onboard Type", report
    End If
End Sub
