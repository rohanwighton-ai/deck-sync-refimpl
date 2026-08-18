Attribute VB_Name = "InjectPrimitive"
Option Explicit

' VBA port of src/verification.py's inject_primitive, scoped to a single
' primitive write. Ports the semantics, not the implementation: the Python
' side hand-rolls OOXML zip/XML surgery to read/write a shape's text and to
' resolve a shape by its hidden identity tag (custDataLst -> Tags Part,
' reverse-engineered from ECMA-376 since python-pptx has no support for it
' -- see src/identity_tags.py). VBA/the PowerPoint object model already
' implements that exact mechanism natively as Shape.Tags, so this port does
' not re-derive any XML; it uses the object model directly. See
' SPIKE_NOTES.md for the full list of deliberate divergences from the
' Python semantics and why each one is safe.

Public Type InjectResult
    Found As Boolean       ' False if no shape (or >1 shapes) carried the tag
    Written As Boolean     ' False if the current value already matched (no-op)
    Verified As Boolean    ' current value re-read from the shape equals source_value
    ErrorMessage As String ' non-empty iff Found=False or Verified=False
    WouldChange As Boolean ' the shape's text differs from sourceValue (set in BOTH modes)
    CurrentValue As String ' the shape's text as found, before any write -- for previews
    ' Geometry, because "the text is right" is not the same as "the slide is
    ' right" on a deck where position and size carry meaning. GeometryMoved is
    ' True when writing the text shifted or resized the shape (autofit doing its
    ' job); GeometryRestored is False when it could not be put back, which is a
    ' slide that now differs from the one that was approved.
    GeometryMoved As Boolean
    GeometryRestored As Boolean
End Type

' The register's line-break delimiter (R6). Excel holds no multi-line cells --
' every line break in a value is written as "||" and converted to a real break
' here, at the point of injection.
'
' Not an edge case. Measured across the real 46-slide deck 2026-07-31:
' KEY_EVENTS_BODY is multi-line on 46 of 46 slides, median 5 paragraphs;
' ABOUT_BODY on 3 of 46.
Public Const LINE_BREAK_DELIMITER As String = "||"

' PowerPoint's in-paragraph line break is vbCr (Chr 13), NOT vbCrLf. Writing
' vbCrLf into a TextRange leaves the Lf as a stray glyph on the slide.
Private Const PPT_LINE_BREAK As String = vbCr

' Locate the single shape on `sld` whose Tags("role") equals `identityTag`.
' Mirrors identity_tags.py's read_shape_tags/upsert_shape_tags convention:
' the identity tag lives under the fixed tag name "role" (specs/identity-
' tags.md: "Shape-level tags (role)"), not under the shape's visible Name --
' real-deck validation upstream found 5 shapes on one real slide all named
' the literal string "Rounded Rectangle 1", which is exactly why tag-based
' lookup exists instead of name-based lookup.
'
' Returns Nothing if zero or more than one shape carries the tag -- multiple
' matches are a data-integrity problem (tags are supposed to be unique per
' role on a slide) and are never silently resolved by picking the first one,
' consistent with the source project's stated philosophy of proving links
' rather than assuming them.
'
' RECURSES into groups (delegates to the private Walk helper below) -- real,
' confirmed bug found 2026-07-26 against Rohan's real 46-slide deck (see
' SPIKE_NOTES_BatchOnboardFlow.md's addendum for the full account): `sld.
' Shapes` is PowerPoint's TOP-LEVEL-ONLY shape collection -- it does NOT
' include shapes nested inside a GroupShape's GroupItems, unlike Discovery.
' Walk (Discovery.bas), which explicitly recurses into groups when building
' its Candidate list. Since Rohan's real deck's "card" layout fields live
' inside grpSp groups (confirmed structurally: every one of the real deck's
' 46 slides contains at least one group, and most of a slide's text-bearing
' shapes are nested inside one, not top-level), a role tag written onto a
' grouped field via Onboarding.ConfirmFieldMatch (Shape.Tags.Add, which
' works on any shape regardless of nesting) was being written correctly but
' could never be found again by this function -- Found=False for every
' grouped field, on every slide, deterministically. This is what actually
' turned "Linked: 0 / FAILED verification: 46" into a uniform, whole-batch
' failure rather than a handful of edge cases: the write always worked, the
' read-back never could, for exactly the shapes this deck actually uses.
' What a picture field was last filled from, stamped on the shape itself.
Public Const PICTURE_SOURCE_TAG As String = "picsrc"

' ONE FIELD, MANY VALUES -- the milestone case.
'
' Rohan, 2026-08-10: some slides carry several progress bars, "same metric but
' different progress against different milestones". Five bars is five values for
' one concept, and a register column holds one value per slide, so the values
' share a cell: `0.9||0.5||0.2`, against shapes tagged `<field>.1`, `.2`, `.3`.
'
' SAME CHARACTERS AS LINE_BREAK_DELIMITER, DELIBERATELY DERIVED FROM IT rather
' than typed again -- but named separately, because it is doing a different job
' and this project has already paid for one token meaning two things. In a text
' field `||` is a line break; in a repeating progress field it separates values.
' They never collide: a progress value is a number and cannot contain a break.
' Deriving means a future change to one is a deliberate choice about both,
' instead of a silent divergence.
Public Const VALUE_SEPARATOR As String = LINE_BREAK_DELIMITER

Public Const INJECTOR_DEVICE As String = "device"
Public Const INJECTOR_PICTURE As String = "picture"
Public Const INJECTOR_REPEATING As String = "repeating-bar"
Public Const INJECTOR_BAR As String = "bar"
Public Const INJECTOR_TEXT As String = "text"

' FIX-LIST R, 2026-08-15. A device's identity lives on the SLIDE (a tagged
' group), and the register may hold nothing under that name at all --
' MILESTONE_TIMELINE is the case that found this: 21 real data columns
' (MS1_LABEL..MS7_DONE) and zero rows containing a cell literally named
' "MILESTONE_TIMELINE". SyncOperations.PlanRoutineSync walks the register's
' OWN column headers as field identities, so a device tag that is not a
' column was never once asked about -- not routed wrong, never called.
' Confirmed live: seeded real milestone data, ran the real sync button
' against a real deck, nothing changed, confirmed from the slide's own
' shapes afterward.
'
' This is the other half of the fix: find every device-routed tag by walking
' the slide's shapes, the same way InjectorFor already decides a tag routes
' to INJECTOR_DEVICE (msoGroup + MilestoneDevice.SlotCount > 0), so the sync
' loop has something to call PlanRoutineSync with even when no register
' column shares the tag's name.
'
' Deduplicated by tag, not by shape: two groups sharing one tag is the
' ambiguous case FindShapeByRoleTag already refuses downstream (Nothing on
' 2+ matches), so this only needs to ask ONCE per distinct tag.
Public Function DeviceRoleTagsOnSlide(sld As Object) As Object
    Dim tags As Object
    Set tags = CreateObject("Scripting.Dictionary")
    WalkForDeviceRoleTags sld.Shapes, tags
    Set DeviceRoleTagsOnSlide = tags
End Function

' A DEVICE IS ADDRESSED BY NAME, NOT BY TAG -- and requiring a tag here made
' the mechanism unreachable on every real slide in the deck, found 2026-08-16.
' MilestoneDevice.bas's own header is explicit about the convention this
' module has to match: "they are simply addressed by name... this is what
' stands in for the robustness a tag would have given" (Rohan, 2026-08-10).
' The prototype slide (3_P001) has a fully-built, correctly-named
' MILESTONE_TIMELINE group with zero tags on it -- by design, not omission --
' and this walk required a tag anyway, so PlanRoutineSync had never once
' synced it. A tag still WINS when present (an explicit override stays
' possible), it is just no longer REQUIRED for a structurally-confirmed
' device (msoGroup + a real slot structure).
Private Sub WalkForDeviceRoleTags(shapesColl As Object, ByRef tags As Object)
    Dim shp As Object
    For Each shp In shapesColl
        If shp.Type = msoGroup Then
            If MilestoneDevice.SlotCount(shp) > 0 Then
                Dim roleVal As String
                roleVal = shp.Tags("role")
                If roleVal = "" Then roleVal = shp.Name
                If roleVal <> "" Then
                    If Not tags.Exists(roleVal) Then tags.Add roleVal, True
                End If
            End If
            WalkForDeviceRoleTags shp.GroupItems, tags
        End If
    Next shp
End Sub

' PUBLIC since 2026-08-14, for Harvest.bas. Deliberately shared rather than
' copied: this walk carries two properties a second copy would lose within a
' week -- it TESTS a group as well as recursing into it (see WalkForRoleTag's
' header for the day that was an ElseIf and made every device unfindable), and
' it returns Nothing on TWO matches as well as none, so an ambiguous tag can
' never be silently resolved to whichever shape came first.
Public Function FindShapeByRoleTag(sld As Object, identityTag As String) As Object
    ' THE FAST PATH. ShapeAddressBook.bas's own header has the full story --
    ' compressed here: a shape's name and location never change between
    ' runs (Slide.Duplicate copies both from the template, and nothing in
    ' this codebase renames a shape afterwards), so a name once discovered
    ' by the full walk below is remembered and tried FIRST next time,
    ' verified by the same tag check the walk itself uses, never trusted
    ' blindly. Costs one cheap Lookup and, on a hit, one native
    ' Shapes(name) access instead of walking the whole slide.
    Dim slideType As String
    Dim resolved As SlideInstance
    resolved = Resolve.ResolveSlideInstance(sld)
    If resolved.HasTypeTag Then
        slideType = resolved.TypeTag
        Dim cachedName As String
        cachedName = ShapeAddressBook.Lookup(slideType, identityTag)

        ' NO_SHAPE_MARKER can only arrive from a LEGACY sheet row (the
        ' pre-AT era persisted misses to the sheet; ShapeAddressBook.Lookup
        ' no longer produces it itself). It is a hint that expired with its
        ' design, not a verdict: treating it as one was the 2026-08-18
        ' TIMELINE_ELAPSED incident (see mSlideTagIndex's header in
        ' ShapeAddressBook.bas). Ignore it and let the per-slide index
        ' below answer instead.
        If cachedName <> "" And cachedName <> ShapeAddressBook.NO_SHAPE_MARKER Then
            Dim candidate As Object
            On Error Resume Next
            Set candidate = sld.Shapes(cachedName)
            On Error GoTo 0
            If Not candidate Is Nothing Then
                If ShapeHasRoleTag(candidate, identityTag) Then
                    Set FindShapeByRoleTag = candidate
                    Exit Function
                End If
            End If
        End If
    End If

    ' CONFIRMED ABSENCE, PER SLIDE. If a full walk of THIS slide has already
    ' run this session, the identity keys it saw are on record, and a tag
    ' not among them is confirmed absent here with zero COM calls -- for
    ' this slide only. The old per-TYPE absent marker let one slide's
    ' genuine lack of a shape veto another slide's confirmed presence of it
    ' (the 2026-08-18 TIMELINE_ELAPSED "dropped: changed since approval"
    ' incident -- ShapeAddressBook.mSlideTagIndex's header has the full
    ' story). A recorded "present" is deliberately NOT trusted the same
    ' way: it only licenses the walk below, whose own tag check remains the
    ' authority, so a stale "present" costs a walk while a false "absent"
    ' can never happen for a slide the walk actually saw.
    Dim slideKey As String
    slideKey = ShapeAddressBook.SlideKeyFor(sld)
    Dim slideTags As Object
    Set slideTags = ShapeAddressBook.SlideTagsFor(slideKey)
    If Not slideTags Is Nothing Then
        If Not slideTags.Exists(identityTag) Then
            Set FindShapeByRoleTag = Nothing
            Exit Function
        End If
    End If

    ' THE SLOW PATH. Cache miss (nothing recorded yet) or a stale entry
    ' (candidate gone, or its tag no longer matches) both land here -- the
    ' same full, ambiguity-checked walk this function has always done. The
    ' walk now also collects EVERY identity key it passes, so one walk
    ' indexes the whole slide as a byproduct: that is what bounds the
    ' session's walk count by the slide count instead of by
    ' (type x absent-field x suffix variant).
    Dim match As Object
    Dim matchCount As Long
    matchCount = 0
    Dim tagsPresent As Object
    Set tagsPresent = CreateObject("Scripting.Dictionary")

    WalkForRoleTag sld.Shapes, identityTag, match, matchCount, tagsPresent
    ShapeAddressBook.RecordSlideTags slideKey, tagsPresent

    If matchCount = 1 Then
        Set FindShapeByRoleTag = match
        ' SELF-HEALS HERE. Whatever the walk just found -- for the first
        ' time, or correcting a stale entry -- is recorded so the NEXT call
        ' for this (slide type, field) takes the fast path instead. No
        ' separate discovery pass anywhere; this is the only place the book
        ' is ever written.
        If slideType <> "" Then ShapeAddressBook.Record slideType, identityTag, match.Name
    Else
        Set FindShapeByRoleTag = Nothing
        ' matchCount = 2+ (two shapes claiming one tag) still refuses on
        ' every call: the tag IS in tagsPresent, so the index never
        ' short-circuits it, and every probe re-walks and re-refuses --
        ' ambiguity stays loud, exactly as before.
    End If
End Function

' Recursive helper -- same group-recursion shape Discovery.bas's own Walk
' establishes (Discovery.Walk shp.GroupItems on msoGroup), kept private and
' local to this module rather than promoted/shared, matching this project's
' "small, focused modules" convention (InjectPrimitive.bas's own header
' notes the same reasoning for not sharing VerifyBatchLink). ByRef match/
' matchCount accumulate across the whole recursive walk so a same-tag
' collision between a top-level shape and a nested one (or between two
' shapes in different groups) is still caught, not silently missed by
' scoping matchCount per group level.
' `tagsPresent` collects every identity key this walk passes -- the role tag
' of any shape carrying one, plus the NAME of any untagged group that
' MilestoneDevice.SlotCount confirms as a device. The collection rule MUST
' mirror ShapeHasRoleTag exactly (same tag read, same name-fallback gate),
' because the per-slide index built from it answers "could ShapeHasRoleTag
' ever return True for tag T on this slide" -- collect less and a real
' shape becomes invisible for the session, collect more and the only cost
' is a wasted walk that answers correctly.
Private Sub WalkForRoleTag(shapesColl As Object, identityTag As String, ByRef match As Object, ByRef matchCount As Long, ByRef tagsPresent As Object)
    Dim shp As Object
    For Each shp In shapesColl
        Dim roleVal As String
        roleVal = shp.Tags("role")
        If roleVal <> "" Then
            tagsPresent(roleVal) = True
        ElseIf shp.Type = msoGroup Then
            ' The device name-fallback, same gate as ShapeHasRoleTag: only a
            ' structurally-confirmed device (untagged group with a real slot
            ' structure) is findable by its name.
            If MilestoneDevice.SlotCount(shp) > 0 Then tagsPresent(shp.Name) = True
        End If

        ' A GROUP IS TESTED **AND** RECURSED INTO. This was an ElseIf until
        ' 2026-08-10, so a group carrying a role tag could never be found by
        ' anything -- the walk stepped straight past it into its members. That
        ' made a tagged DEVICE unreachable no matter how the router dispatched,
        ' and it was invisible until the first test that tagged a group.
        '
        ' Safe to widen: nothing tagged a group before now, because the marking
        ' flow refused one outright.
        If ShapeHasRoleTag(shp, identityTag) Then
            matchCount = matchCount + 1
            Set match = shp
        End If
        If shp.Type = msoGroup Then
            WalkForRoleTag shp.GroupItems, identityTag, match, matchCount, tagsPresent
        End If
    Next shp
End Sub

' HOW MANY shapes carry this role, which FindShapeByRoleTag cannot tell you --
' it collapses none and two into the same Nothing, correctly for its own purpose
' and uselessly for a caller deciding whether to ADD a tag. Harvest.bas needs the
' difference: stamping a role onto a slide that already has two of them makes a
' broken slide worse, and stamping onto one that has none is the whole job.
Public Function CountShapesWithRoleTag(sld As Object, identityTag As String) As Long
    Dim match As Object
    Dim matchCount As Long
    matchCount = 0
    ' Throwaway collector: this caller wants the count, not the index, and
    ' deliberately does not record it -- counting is a pre-stamp check
    ' (Harvest), and recording a pre-stamp snapshot would be exactly the
    ' stale-absence NoteRoleTagAdded exists to prevent.
    Dim discard As Object
    Set discard = CreateObject("Scripting.Dictionary")
    WalkForRoleTag sld.Shapes, identityTag, match, matchCount, discard
    CountShapesWithRoleTag = matchCount
End Function

' Shape.Tags(name) returns "" both when the tag is absent and when it is
' present with an empty string value -- a real VBA API quirk with no
' equivalent in the Python side's explicit dict-based tag storage (identity_
' tags.py's _parse_tag_list returns a real {} when a tag is missing). Since
' an empty-string role value is not a meaningful identity tag in this
' project's scheme, treating "absent" and "empty" the same way here is a
' safe simplification, not a correctness gap -- documented in SPIKE_NOTES.md.
' The name-fallback for a device (see WalkForDeviceRoleTags's header) has to
' land here too, not just in discovery -- InjectorFor and every other caller
' re-find the shape by whatever identityTag discovery handed back, and if
' this function still demanded a tag, a device found by name upstream would
' vanish again the moment anything tried to locate it a second time.
'
' Scoped STRICTLY to a structurally-confirmed device (msoGroup with a real
' slot structure, no tag already present) so this cannot widen matching for
' an ordinary text/bar/picture field that happens to share a shape's name --
' only a shape MilestoneDevice.SlotCount already recognises as a device gets
' the name fallback; everything else still requires an exact tag match,
' unchanged.
Private Function ShapeHasRoleTag(shp As Object, identityTag As String) As Boolean
    If identityTag = "" Then
        ShapeHasRoleTag = False
        Exit Function
    End If
    If shp.Tags("role") = identityTag Then
        ShapeHasRoleTag = True
        Exit Function
    End If
    If shp.Type = msoGroup Then
        If shp.Tags("role") = "" Then
            If MilestoneDevice.SlotCount(shp) > 0 Then
                ShapeHasRoleTag = (shp.Name = identityTag)
                Exit Function
            End If
        End If
    End If
    ShapeHasRoleTag = False
End Function

' Port of inject_primitive(path, part_name, shape, source_value). Takes a
' live Slide object and the identity tag instead of a file path + z_order,
' since the calling context here is a shape already open in the PowerPoint
' object model, not a closed .pptx being surgically edited byte-for-byte.
'
' Semantics preserved from the Python original:
'   1. No-op if the shape's current value already equals source_value.
'   2. Otherwise write source_value, then re-read the value back and
'      confirm it matches, rather than assuming the write call succeeded.
'   3. Error (Found=False) if the shape has no text frame / no text runs to
'      write into, mirroring inject_primitive's raise on a shape with no
'      text runs.
'
' Semantics NOT preserved (see SPIKE_NOTES.md for why each is judged safe
' for this spike's scope):
'   - Hash comparison (SHA-256) is replaced with direct VBA string equality.
'   - "Verified" re-reads the value from the live Shape object, which only
'     proves the object-model state changed -- it does not prove the write
'     persisted to the underlying OOXML part the way the Python version's
'     re-open-the-zip-from-disk check does. Manual verification recipe
'     below closes and reopens the file to close this gap by hand.
' `dryRun` makes this a pure read: the shape is located and compared, and the
' result reports what WOULD happen (WouldChange / CurrentValue), but nothing is
' written. This is the primitive the whole Sync Now preview is built on -- the
' write lives here and nowhere else, so gating it here is what makes a preview
' provably safe rather than merely careful. Deliberately a parameter on the
' existing function rather than a parallel "compare" function: a preview whose
' logic can drift from the real thing is worse than no preview.
' Trailing paragraph marks do not count as a difference.
'
' RULE ADDED 2026-07-31, on Rohan's decision, and it changes comparison
' semantics for EVERY field -- recorded here rather than buried at a call site
' because of that blast radius.
'
' The evidence: ABOUT_BODY measured against the real 46-slide deck showed 22
' pending changes after the register's encoding was repaired, and 20 of them
' were this and nothing else -- `TextRange.Text` returns the shape's trailing
' paragraph mark, and the harvested register value does not carry it. Not one
' was a wording change. Applying them would have rewritten 20 slides of real
' prose to remove a character nobody can see, and no human reading that list
' would ever have approved it.
'
' NORMALISES FOR COMPARISON ONLY. What gets WRITTEN is still `sourceValue`
' exactly as supplied -- this decides whether to write, never what to write.
' The distinction matters: silently trimming on write would slowly strip real
' formatting out of a deck one sync at a time.
'
' vbCr, vbLf and Chr(11) are all treated as trailing whitespace: PowerPoint
' returns paragraphs CR-separated and soft line breaks as Chr(11) (AGENTS.md),
' so which one lands at the end depends on how the text was authored.
Private Function IgnoringTrailingBreaks(text As String) As String
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
    IgnoringTrailingBreaks = t
End Function

' ===========================================================================
' THE ROUTER -- the one entry point sync uses, so a field's TYPE is decided in
' exactly one place.
'
' Until 2026-08-10 every sync call site called InjectPrimitive -- the TEXT
' writer -- for every field. InjectPictureField and InjectProgressField were
' unit-tested and called by NOTHING but their own tests, so a picture or a
' progress bar was handed to the text writer and refused as "no text frame to
' write into": two tested components and no feature. 169 passing tests said
' nothing about it, because a unit test asks "does this behave when called"
' and nothing asked "can a person cause it to be called".
'
' THE TYPE COMES FROM THE SHAPE, NOT FROM A COLUMN. The template already
' decides what a field is by what it puts on the slide; a second answer in the
' workbook could disagree with it silently, which is the collision `Kind`
' already has (FIX-LIST item 4). Derived, so it cannot drift.
'
'   the tagged shape is a picture             -> InjectPictureField
'   a sibling tagged `<field>.track` exists   -> InjectProgressField
'   anything else                             -> InjectPrimitive (text)
'
' Picture beats track when a slide somehow offers both, rather than being
' reported as ambiguous: the decision has to be deterministic, because an
' injector chosen differently on the preview run and the write run would show
' one thing and do another.
'
' srcWs is the Sources worksheet and it is OPTIONAL, because only the picture
' branch needs it -- a bar reads its whole value from the register cell. Where
' it is absent a picture field SAYS SO instead of falling through to the text
' writer, whose "no text frame" is a true sentence about the wrong question.
' WHICH INJECTOR CLAIMS THIS FIELD ON THIS SLIDE -- a FOURTH axis, and it needed
' its own word because the other three are taken and mean different things:
'
'   Kind      Controlled/Prose/Static/Derived  -- how a value is reviewed (ReviewQueue)
'   FieldType text/number/currency/date        -- how the register cell is formatted
'   Behaviour fill/fit/as-is                   -- how a picture is placed
'   Injector  (here)                           -- what will be WRITTEN INTO, on this slide
'
' Overloading any of those would be the "a word doing two jobs" defect this
' project has already paid for twice.
'
' EXTRACTED FROM InjectField 2026-08-15 rather than copied. Harvest.bas needs the
' same answer -- a field whose injector is a BAR must never be given the string
' the slide displays ("33%"), because InjectProgressVia refuses anything
' non-numeric and the harvest's empty-cell rule would make that unfixable. Two
' copies of this decision would drift within a week; the router below now uses
' this function, so harvest and publish cannot disagree about the same field.
Public Function InjectorFor(sld As Object, identityTag As String) As String
    Dim shp As Object
    Set shp = FindShapeByRoleTag(sld, identityTag)

    ' A DEVICE IS A GROUP ON PURPOSE -- a milestone timeline is ONE field made
    ' of fifteen shapes, tagged once, its parts found by NAME inside it. Checked
    ' before everything else because a group is not a picture, has no text
    ' frame, and would otherwise fall through to the text writer and be refused
    ' for having nowhere to put a string.
    If Not shp Is Nothing Then
        If shp.Type = msoGroup Then
            If MilestoneDevice.SlotCount(shp) > 0 Then
                InjectorFor = INJECTOR_DEVICE
                Exit Function
            End If
        End If
    End If

    If Not shp Is Nothing Then
        If IsPictureShape(shp) Then
            InjectorFor = INJECTOR_PICTURE
            Exit Function
        End If
    End If

    ' REPEATING BEFORE SINGLE. A slide carrying `<field>.1` is the milestone
    ' case -- one metric measured against several milestones -- and it is
    ' checked before the single-bar branch because such a slide may ALSO carry a
    ' shared `<field>.track` axis, which would otherwise capture it and inject
    ' one value into a bar that does not exist.
    If Not FindShapeByRoleTag(sld, identityTag & ".1") Is Nothing Then
        InjectorFor = INJECTOR_REPEATING
        Exit Function
    End If

    ' EITHER COMPANION MAKES IT A BAR -- a track, or a remainder.
    '
    ' The done part is an ordinary rectangle and looks like any other shape, so
    ' the discriminator has to be something beside it. Keying on the TRACK alone
    ' was wrong the moment the trackless pair existed: Rohan's `Time elapsed`
    ' bar is a filled rect and a grey one with no track anywhere, so it fell
    ' through to the TEXT writer, which wrote "0.5" into the rectangle as a
    ' string and reported success. Caught by the trackless test, invisible to
    ' every unit test of the injector itself -- the injector was never the thing
    ' that was wrong.
    '
    ' Asking for a companion rather than the done part also means a bar whose
    ' done part has been deleted still routes here and gets InjectProgressField's
    ' specific message, instead of being told it has no text frame.
    If Not FindShapeByRoleTag(sld, identityTag & ".track") Is Nothing _
       Or Not FindShapeByRoleTag(sld, identityTag & FieldWiring.REST_SUFFIX) Is Nothing Then
        InjectorFor = INJECTOR_BAR
        Exit Function
    End If

    InjectorFor = INJECTOR_TEXT
End Function

Public Function InjectField(sld As Object, identityTag As String, sourceValue As String, _
                            Optional dryRun As Boolean = False, _
                            Optional srcWs As Object = Nothing, _
                            Optional rowValues As Object = Nothing) As InjectResult
    ' ONE DECISION, made in one place. See InjectorFor's header.
    Select Case InjectorFor(sld, identityTag)
        Case INJECTOR_DEVICE
            InjectField = InjectDeviceVia(FindShapeByRoleTag(sld, identityTag), identityTag, rowValues, dryRun)
        Case INJECTOR_PICTURE
            InjectField = InjectPictureVia(sld, identityTag, sourceValue, srcWs, dryRun)
        Case INJECTOR_REPEATING
            InjectField = InjectRepeatingProgress(sld, identityTag, sourceValue, dryRun)
        Case INJECTOR_BAR
            InjectField = InjectProgressVia(sld, identityTag, sourceValue, dryRun)
        Case Else
            InjectField = InjectPrimitive(sld, identityTag, sourceValue, dryRun)
    End Select
End Function

' A register cell is text; a bar needs a number. Converting is the router's
' job, and REFUSING is the important half of it.
'
' Val("done") returns 0 and CDbl("done") raises. Val's answer is the dangerous
' one: it would draw an empty bar and report success, which is the exact
' failure InjectProgressField's own out-of-range check exists to prevent (a
' register holding 90 for 90% drawing a full bar and looking right).
Private Function InjectProgressVia(sld As Object, identityTag As String, _
                                   sourceValue As String, dryRun As Boolean) As InjectResult
    Dim result As InjectResult

    If Not IsNumeric(Trim(sourceValue)) Then
        ' Found and WouldChange are both True deliberately. Found=False is a
        ' SKIP to SyncOperations and WouldChange=False is a NO CHANGE -- either
        ' one would swallow this in silence, which is how the tool would come
        ' to report a clean run over a bar it never drew.
        result.Found = True
        result.WouldChange = True
        result.ErrorMessage = "progress field " & identityTag & " needs a number between 0 and 1" & _
            " and the register holds '" & sourceValue & "'" & _
            " -- 90% belongs in the register as 0.9, not 90 and not '90%'."
        InjectProgressVia = result
        Exit Function
    End If

    InjectProgressVia = InjectProgressField(sld, identityTag, CDbl(Trim(sourceValue)), dryRun)
End Function

' SEVERAL BARS, ONE CELL -- the milestone case.
'
' `0.9||0.5||0.2` against shapes tagged `<field>.1`, `.2`, `.3`.
'
' A COUNT MISMATCH IS REFUSED, NOT TRUNCATED, and that is the whole reason this
' is a function rather than a loop at the call site. Three values against five
' bars could be "write three and leave two", which draws a slide that is
' PLAUSIBLE and wrong -- two milestones silently showing last quarter's
' progress. Writing nothing and saying which counts disagree is the only
' honest answer, and it is the same rule as the out-of-range check: a bar that
' cannot be drawn correctly is not drawn at all.
'
' The track for bar N is `<field>.N.track` if the slide carries one, otherwise
' the shared `<field>.track`. Both are real layouts -- milestones may each have
' their own track, or all sit against one timeline axis -- so it resolves
' rather than assumes, and says which it looked for when neither is there.
Public Function InjectRepeatingProgress(sld As Object, identityTag As String, _
                                        sourceValue As String, _
                                        Optional dryRun As Boolean = False) As InjectResult
    Dim result As InjectResult
    result.Found = True
    result.WouldChange = True

    Dim bars As Long
    bars = 0
    Do While Not FindShapeByRoleTag(sld, identityTag & "." & (bars + 1)) Is Nothing
        bars = bars + 1
    Loop

    Dim parts() As String
    parts = Split(sourceValue, VALUE_SEPARATOR)
    Dim vals As Long
    vals = UBound(parts) - LBound(parts) + 1

    If vals <> bars Then
        result.ErrorMessage = identityTag & ": the register holds " & vals & " value(s) (" & _
            sourceValue & ") but the slide has " & bars & " bar(s) tagged " & identityTag & _
            ".1.." & identityTag & "." & bars & ". Nothing was drawn -- writing the ones " & _
            "that match would leave the rest showing an older figure."
        InjectRepeatingProgress = result
        Exit Function
    End If

    ' Reported as one value so a preview and a change-hash see the whole field,
    ' not one bar of it.
    Dim currents As String, wroteAll As Boolean, anyChange As Boolean
    wroteAll = True

    Dim i As Long
    For i = 1 To bars
        Dim barTag As String, trackTag As String
        barTag = identityTag & "." & i
        trackTag = barTag & FieldWiring.TRACK_SUFFIX
        If FindShapeByRoleTag(sld, trackTag) Is Nothing Then
            trackTag = identityTag & FieldWiring.TRACK_SUFFIX
        End If

        Dim one As InjectResult
        If Not IsNumeric(Trim(parts(LBound(parts) + i - 1))) Then
            one.ErrorMessage = barTag & " needs a number between 0 and 1, and the register holds '" & _
                Trim(parts(LBound(parts) + i - 1)) & "'"
        Else
            one = InjectProgressField(sld, barTag, CDbl(Trim(parts(LBound(parts) + i - 1))), _
                                      dryRun, trackTag)
        End If

        If currents <> "" Then currents = currents & VALUE_SEPARATOR
        currents = currents & one.CurrentValue
        If one.WouldChange Then anyChange = True
        If Not (one.Written Or one.Verified) Then
            wroteAll = False
            If result.ErrorMessage <> "" Then result.ErrorMessage = result.ErrorMessage & "; "
            result.ErrorMessage = result.ErrorMessage & one.ErrorMessage
        End If
    Next i

    result.CurrentValue = currents
    result.WouldChange = anyChange
    result.Written = wroteAll And Not dryRun
    result.Verified = wroteAll
    InjectRepeatingProgress = result
End Function


' A milestone device, driven from the row.
'
' DRY RUN READS AND REPORTS WITHOUT DRAWING. A preview that silently drew the
' timeline would write to the deck during "Preview Sync -- nothing written",
' which is the promise that surface is built on.
Private Function InjectDeviceVia(grp As Object, identityTag As String, _
                                 rowValues As Object, dryRun As Boolean) As InjectResult
    Dim result As InjectResult
    result.Found = True

    If rowValues Is Nothing Then
        result.WouldChange = True
        result.ErrorMessage = identityTag & " is a milestone timeline and needs its whole " & _
            "register row (the " & MilestoneDevice.SLOT_PREFIX & "1" & MilestoneDevice.COL_LABEL & _
            " columns), which was not available here."
        InjectDeviceVia = result
        Exit Function
    End If

    result.CurrentValue = MilestoneDevice.DeviceIntegrity(grp)

    If dryRun Then
        result.WouldChange = True
        result.Verified = True
        InjectDeviceVia = result
        Exit Function
    End If

    Dim drawn As MilestoneDrawResult
    drawn = MilestoneDevice.DrawFromRow(grp, rowValues)

    If drawn.ErrorMessage <> "" Then
        result.WouldChange = True
        result.ErrorMessage = drawn.ErrorMessage
        InjectDeviceVia = result
        Exit Function
    End If

    result.Written = True
    result.Verified = True
    result.WouldChange = True
    result.ErrorMessage = drawn.Detail
    InjectDeviceVia = result
End Function

' A picture field's register cell holds a SOURCE ID, not a path -- the stamp
' the injector leaves behind is that ID, and the path is whatever the Sources
' sheet currently says it is. So a photo that moves is re-pointed in one row
' rather than in forty register cells.
Private Function InjectPictureVia(sld As Object, identityTag As String, _
                                  sourceId As String, srcWs As Object, _
                                  dryRun As Boolean) As InjectResult
    Dim result As InjectResult

    If srcWs Is Nothing Then
        result.Found = True
        result.WouldChange = True
        result.ErrorMessage = "picture field " & identityTag & " is filled from a source, and the '" & _
            Sources.SOURCES_SHEET_NAME & "' sheet was not available here, so its image could not be found."
        InjectPictureVia = result
        Exit Function
    End If

    Dim known As Boolean
    Dim locator As String
    locator = Sources.LocatorFor(srcWs, sourceId, known)

    ' The two failures are told apart because they send the person to
    ' different places: a citation typo is fixed in the register, a blank
    ' locator on a real row is fixed on the Sources sheet.
    If Not known Then
        result.Found = True
        result.WouldChange = True
        result.ErrorMessage = "picture field " & identityTag & " cites source '" & Trim(sourceId) & _
            "', which is not on the '" & Sources.SOURCES_SHEET_NAME & "' sheet."
        InjectPictureVia = result
        Exit Function
    End If

    InjectPictureVia = InjectPictureField(sld, identityTag, Trim(sourceId), locator, dryRun)
End Function
' ===========================================================================

Public Function InjectPrimitive(sld As Object, identityTag As String, sourceValue As String, Optional dryRun As Boolean = False) As InjectResult
    Dim result As InjectResult
    Dim shp As Object

    ' Converted FIRST, before anything is located or compared, so every
    ' comparison below is made in the slide's own terms.
    '
    ' The ordering is the correctness of it, not a tidiness preference. The
    ' no-op check compares the slide's current text against sourceValue -- if
    ' the value still held "||" while the slide held a real line break, the two
    ' could never be equal, so the field would be rewritten on every single
    ' sync and reported as corrected every time. A permanent phantom
    ' correction, on the field that is multi-line on 46 of 46 slides.
    sourceValue = Replace(sourceValue, LINE_BREAK_DELIMITER, PPT_LINE_BREAK)

    Set shp = FindShapeByRoleTag(sld, identityTag)
    If shp Is Nothing Then
        result.Found = False
        result.Written = False
        result.Verified = False
        result.ErrorMessage = "no single shape found tagged role=" & identityTag & _
            " (zero matches, or more than one -- ambiguous tag, refusing to guess)"
        InjectPrimitive = result
        Exit Function
    End If

    result.Found = True

    If Not shp.HasTextFrame Then
        result.Written = False
        result.Verified = False
        result.ErrorMessage = "shape tagged role=" & identityTag & " has no text frame to write into"
        InjectPrimitive = result
        Exit Function
    End If

    Dim currentValue As String
    currentValue = shp.TextFrame.TextRange.Text
    result.CurrentValue = currentValue
    result.WouldChange = (IgnoringTrailingBreaks(currentValue) <> IgnoringTrailingBreaks(sourceValue))

    If Not result.WouldChange Then
        result.Written = False
        result.Verified = True
        result.ErrorMessage = ""
        InjectPrimitive = result
        Exit Function
    End If

    ' The one and only place this module mutates a slide -- everything above is
    ' a read. A dry run stops here, reporting the pending change rather than
    ' making it. Verified stays False because nothing was written, so nothing
    ' could be verified: callers must not read it as "this write failed".
    If dryRun Then
        result.Written = False
        result.Verified = False
        result.ErrorMessage = ""
        InjectPrimitive = result
        Exit Function
    End If

    ' POSITION AND SIZE ARE PART OF THE SLIDE, AND WRITING TEXT CAN MOVE THEM.
    '
    ' Rohan, 2026-08-01: "Position is critical to these slides, and size."
    ' Nothing here had ever touched geometry -- but PowerPoint does. A shape
    ' with autofit on resizes when its text gets shorter or longer, and one
    ' with "resize shape to fit text" will move its own bottom edge. So a sync
    ' that only ever assigns .Text can still shift the layout of every slide it
    ' touches, silently, and the existing re-read check would call every one of
    ' those a clean verified write, because the TEXT is right.
    '
    ' Captured before, restored after, and the restore is then RE-READ. An
    ' autofit shape can refuse to keep a width it has been given, so assuming
    ' the restore worked would be the same mistake in a new place.
    '
    ' SCOPE, AND DO NOT GENERALISE THIS. Rohan, 2026-08-01: "Some geometry
    ' movement allowed ie the circles on the timeline changing shape /
    ' formatting based on different points of achievement."
    '
    ' Correct, and the distinction is the whole point:
    '   geometry moving as a SIDE EFFECT of writing text   -> a bug, restore it
    '   geometry or formatting changing as the PAYLOAD     -> the feature
    '
    ' This function is the TEXT injector. Everything it does is the first case,
    ' so restoring is unconditionally right HERE. A bar-part or milestone-marker
    ' injector deliberately moves and restyles shapes -- that is what it is for
    ' -- and must NOT reuse this guard. Anyone tempted to "make injection
    ' consistent" by adding geometry restore there would be deleting the
    ' feature.
    Dim hadL As Single, hadT As Single, hadW As Single, hadH As Single
    Dim haveGeom As Boolean
    On Error Resume Next
    hadL = shp.Left: hadT = shp.Top: hadW = shp.Width: hadH = shp.Height
    haveGeom = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    ' Assigning .Text replaces the entire text-range contents in one run,
    ' which matches _set_text's "write the first run, clear the rest" net
    ' effect (formatting of the original first run is retained by
    ' PowerPoint; any additional runs are removed rather than left stale).
    shp.TextFrame.TextRange.Text = sourceValue

    If haveGeom Then
        Dim movedBy As Single
        movedBy = Abs(shp.Left - hadL) + Abs(shp.Top - hadT) + _
                  Abs(shp.Width - hadW) + Abs(shp.Height - hadH)

        ' A twentieth of a point is below anything visible and above the noise
        ' of PowerPoint's own float handling. Zero would fire on every write.
        If movedBy > 0.05 Then
            result.GeometryMoved = True
            On Error Resume Next
            shp.Left = hadL: shp.Top = hadT
            shp.Width = hadW: shp.Height = hadH
            On Error GoTo 0

            Dim stillOff As Single
            stillOff = Abs(shp.Left - hadL) + Abs(shp.Top - hadT) + _
                       Abs(shp.Width - hadW) + Abs(shp.Height - hadH)
            result.GeometryRestored = (stillOff <= 0.05)
        Else
            result.GeometryRestored = True
        End If
    Else
        ' No geometry readable (a placeholder mid-edit, a protected shape).
        ' Reported as not-restored rather than quietly true: unknown is not the
        ' same as fine, and this is the field where that distinction matters.
        result.GeometryRestored = False
    End If

    Dim reReadValue As String
    reReadValue = shp.TextFrame.TextRange.Text

    result.Written = True
    ' Same rule as the pre-write check. If these disagreed, a value whose
    ' only difference was a trailing break would be judged "needs writing"
    ' by one and "write did not take" by the other -- a permanent failure
    ' on a field that is actually correct.
    result.Verified = (IgnoringTrailingBreaks(reReadValue) = IgnoringTrailingBreaks(sourceValue))
    If Not result.Verified Then
        result.ErrorMessage = "wrote " & sourceValue & " but re-read " & reReadValue & " -- write did not take"
    Else
        result.ErrorMessage = ""
    End If

    InjectPrimitive = result
End Function

' Manual smoke-test entry point -- not a real test harness (no VBA unit-test
' framework wired up here, see SPIKE_NOTES.md). Run this from the VBA IDE
' (F5) against a slide that already has a shape tagged role="demo_field",
' per the manual verification recipe in SPIKE_NOTES.md.
Public Sub ManualSmokeTest()
    Dim sld As Object
    Set sld = Application.ActivePresentation.Slides(1)

    Dim r As InjectResult
    r = InjectPrimitive(sld, "demo_field", "hello from VBA " & Now)

    Dim msg As String
    msg = "Found=" & r.Found & " Written=" & r.Written & " Verified=" & r.Verified
    If r.ErrorMessage <> "" Then msg = msg & " Error=" & r.ErrorMessage
    MsgBox msg
End Sub

' ---------------------------------------------------------------------
' PICTURE FIELDS
' ---------------------------------------------------------------------
'
' Rohan, 2026-08-10: project images "don't change quarterly, set once at project
' start, I'd still like them set from link rather than played around with
' manually."
'
' So this is a GIVEN field filled from a link, not a quarterly sync. Three
' consequences shape the whole implementation:
'
' 1. WHETHER A FIELD IS A PICTURE IS DERIVED, NEVER DECLARED. The tagged shape
'    either is a picture or it is not. No Field Spec column, nothing to
'    configure, and nothing that can disagree with the deck -- the failure mode
'    a declared type would add is a register that says Picture over a text box.
'
' 2. THE REGISTER CELL HOLDS A SOURCE ID, not a path. An image is
'    evidence-shaped: it came from somewhere, that somewhere has an owner, and
'    Sources already gives one row per thing with a locator and a period
'    binding. A path per row would be 43 spellings of one folder with nothing
'    checking any of them.
'
' 3. IDEMPOTENCE COMES FROM A STAMP, NOT FROM COMPARING IMAGES. The applied
'    source ID is written into the new shape's own tags, so a later run compares
'    two short strings. It fires once at project start, stays silent forever
'    after, and re-fires by itself if the link is ever changed. Comparing pixels
'    every quarter across 43 slides is the thing this avoids.
'
' ASPECT RATIO: FIT INSIDE, CENTRED, NEVER DISTORT, ALWAYS REPORT. Rohan has not
' picked between letterbox / crop / refuse, so the default is the only one that
' cannot lose information or lie about the subject: the image keeps its own
' proportions, sits centred in the frame the slide already defines, and a
' mismatch is REPORTED so the choice can be made against real photos rather than
' in the abstract. Cropping silently decides what matters in someone's picture.

Public Function IsPictureShape(shp As Object) As Boolean
    On Error Resume Next
    IsPictureShape = (shp.Type = msoPicture) Or (shp.Type = msoLinkedPicture)
    On Error GoTo 0
End Function

' What this shape was last filled from, or "" if never.
Public Function PictureSourceOf(shp As Object) As String
    On Error Resume Next
    PictureSourceOf = shp.Tags(PICTURE_SOURCE_TAG)
    On Error GoTo 0
End Function

Public Function InjectPictureField(sld As Object, identityTag As String, _
                                   sourceId As String, locator As String, _
                                   Optional dryRun As Boolean = False) As InjectResult
    Dim result As InjectResult
    Dim shp As Object

    Set shp = FindShapeByRoleTag(sld, identityTag)
    If shp Is Nothing Then
        result.ErrorMessage = "no single shape found tagged role=" & identityTag
        InjectPictureField = result
        Exit Function
    End If
    result.Found = True
    result.CurrentValue = PictureSourceOf(shp)

    ' ALREADY FILLED FROM THIS SOURCE. The whole point of the stamp.
    If StrComp(result.CurrentValue, sourceId, vbTextCompare) = 0 Then
        result.WouldChange = False
        result.Verified = True
        InjectPictureField = result
        Exit Function
    End If

    result.WouldChange = True

    If Trim(sourceId) = "" Then
        result.ErrorMessage = "no source ID given for picture field " & identityTag
        InjectPictureField = result
        Exit Function
    End If

    ' A LOCATOR THAT IS NOT A READABLE FILE IS NOT A PICTURE. Refused before
    ' anything is deleted -- the old image is the only copy on the slide, and
    ' removing it before knowing the replacement exists is how a deck loses a
    ' photo to a typo in a Sources row.
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Trim(locator) = "" Then
        result.ErrorMessage = "source " & sourceId & " has no locator, so there is no image to place"
        InjectPictureField = result
        Exit Function
    End If
    If Not fso.FileExists(locator) Then
        result.ErrorMessage = "source " & sourceId & " points at a file that is not there: " & locator
        InjectPictureField = result
        Exit Function
    End If

    If dryRun Then
        result.Written = False
        result.Verified = True
        InjectPictureField = result
        Exit Function
    End If

    ' FEED THE SHAPE, DO NOT REPLACE IT. Rohan, 2026-08-10: "shouldn't the shape
    ' settings handle it you just feed it".
    '
    ' He is right, and the first implementation proved it the hard way. Deleting
    ' the old shape and inserting a new one meant recreating everything the
    ' shape already knew -- position, size, crop, z-order, tags -- and PowerPoint
    ' recomputes a picture's geometry when crops change, so a 300x150 frame came
    ' back 284x150 at the wrong origin. Two passes of arithmetic did not fix it,
    ' because the arithmetic was never the problem: replacing the shape was.
    '
    ' Fill.UserPicture was PROBED against real PowerPoint before being relied on
    ' -- it works on a picture shape (type 13), not only on an autoshape. The
    ' shape keeps its own settings because nothing touches them.
    ' TWO ROUTES, CHOSEN BY WHAT THE TEMPLATE DID -- not by arithmetic.
    '
    ' Probed against real PowerPoint: Fill.UserPicture and cropping are mutually
    ' exclusive on a picture shape. After feeding, CropLeft reads the no-crop
    ' sentinel and CANNOT be set again. So an uncropped frame can be fed in
    ' place, and a cropped frame cannot -- its framing would be lost with no way
    ' back.
    '
    ' For a cropped frame the shape is rebuilt and the TEMPLATE'S OWN CROP VALUES
    ' are applied verbatim. That is not geometry logic: nothing is computed from
    ' the image's proportions, nothing is fitted or filled. The template said how
    ' this frame is framed, and those numbers are carried across unchanged.
    Dim cL As Single, cR As Single, cT As Single, cB As Single
    On Error Resume Next
    cL = shp.PictureFormat.CropLeft
    cR = shp.PictureFormat.CropRight
    cT = shp.PictureFormat.CropTop
    cB = shp.PictureFormat.CropBottom
    On Error GoTo 0

    Dim isCropped As Boolean
    isCropped = (cL > 0.01) Or (cR > 0.01) Or (cT > 0.01) Or (cB > 0.01)

    If Not isCropped Then
        ' FED IN PLACE. Nothing about the shape changes but its image.
        On Error Resume Next
        Err.Clear
        shp.Fill.UserPicture locator
        If Err.Number <> 0 Then
            Dim eFeed As String
            eFeed = Err.Description
            On Error GoTo 0
            result.ErrorMessage = "could not place " & locator & " (" & eFeed & ")"
            InjectPictureField = result
            Exit Function
        End If
        shp.Tags.Delete PICTURE_SOURCE_TAG
        shp.Tags.Add PICTURE_SOURCE_TAG, sourceId
        On Error GoTo 0
        result.Written = True
        result.Verified = (StrComp(PictureSourceOf(shp), sourceId, vbTextCompare) = 0)
        If Not result.Verified Then
            result.ErrorMessage = "the picture was placed but its source stamp did not stick"
        End If
        InjectPictureField = result
        Exit Function
    End If

    ' A CROPPED FRAME IS REPLACED, CARRYING ITS RECORDED GEOMETRY ACROSS.
    '
    ' Fill.UserPicture cannot be used here -- probed: after feeding, the crop
    ' reads the no-crop sentinel and cannot be set again. So the shape is
    ' rebuilt and its size, position, z-order and CROP are applied from values
    ' captured before anything changed.
    '
    ' LOCKASPECTRATIO MUST BE OFF FIRST, and that single line is what five
    ' earlier attempts were missing. With it on, setting Width makes PowerPoint
    ' rescale Height and setting Height rescales Width back -- probed: a 284x150
    ' target lands on 150x150, and no ordering, re-assertion or arithmetic fixes
    ' it, because nothing was ever wrong with the numbers. Rohan: "you can
    ' replace a pic right? if needed rec size and position and order".
    Dim fL As Single, fT As Single, fW As Single, fH As Single, fZ As Long
    fL = shp.Left: fT = shp.Top: fW = shp.Width: fH = shp.Height
    On Error Resume Next
    fZ = shp.ZOrderPosition
    On Error GoTo 0

    Dim newShp As Object
    On Error Resume Next
    Err.Clear
    Set newShp = sld.Shapes.AddPicture(locator, msoFalse, msoTrue, fL, fT, -1, -1)
    If Err.Number <> 0 Or newShp Is Nothing Then
        Dim eAdd As String
        eAdd = Err.Description
        On Error GoTo 0
        result.ErrorMessage = "could not place " & locator & " (" & eAdd & ")"
        InjectPictureField = result
        Exit Function
    End If

    newShp.LockAspectRatio = msoFalse

    ' The template's framing, verbatim -- nothing computed from the image.
    newShp.PictureFormat.CropLeft = cL
    newShp.PictureFormat.CropRight = cR
    newShp.PictureFormat.CropTop = cT
    newShp.PictureFormat.CropBottom = cB

    newShp.Left = fL
    newShp.Top = fT
    newShp.Width = fW
    newShp.Height = fH

    newShp.Tags.Add "role", identityTag
    newShp.Tags.Add PICTURE_SOURCE_TAG, sourceId
    On Error GoTo 0
    ' Keeps the per-slide tag index honest about the stamp (a no-op here in
    ' practice -- the deleted original carried the same tag -- but every
    ' path that adds a role tag notifies, so the rule has no exceptions).
    ShapeAddressBook.NoteRoleTagAdded newShp, identityTag

    shp.Delete
    On Error Resume Next
    If fZ > 0 Then
        Do While newShp.ZOrderPosition > fZ
            newShp.ZOrder 3          ' msoSendBackward
        Loop
    End If
    On Error GoTo 0

    Set shp = newShp

    result.Written = True
    result.Verified = (StrComp(PictureSourceOf(shp), sourceId, vbTextCompare) = 0)
    If Not result.Verified Then
        result.ErrorMessage = "the picture was placed but its source stamp did not stick -- a later run would replace it again"
    ElseIf isCropped Then
        If Abs(shp.Width - fW) > 0.5 Or Abs(shp.Height - fH) > 0.5 Then
            result.Verified = False
            result.ErrorMessage = "the image was placed but the frame ended " & _
                shp.Width & "x" & shp.Height & " instead of " & fW & "x" & fH & _
                " -- check this slide before trusting it"
        End If
    End If

    InjectPictureField = result
End Function

' ---------------------------------------------------------------------
' PROGRESS BARS: A TRACK AND ITS PARTS, ALL GEOMETRY READ OFF THE SLIDE
' ---------------------------------------------------------------------
'
' Rohan, 2026-08-10: "do the track pair approach, i have used two shapes b4 to
' show progress remainder. use position etc from slide."
'
' A progress field is a PAIR (or a trio), not one shape:
'
'   role = FIELD           the DONE part -- the bit that grows
'   role = FIELD.track     the full extent. Read, never written.
'   role = FIELD.rest      the REMAINDER, optional. Written if present.
'
' THE TRACK IS THE MEASUREMENT AND IT IS NEVER TOUCHED. Scaling the done part
' against its own previous width would work exactly once: after the first run
' the bar is shorter, so the next run scales a fraction of a fraction and the
' bar walks toward zero while every report says success. Reading a shape the
' tool never writes removes that entirely -- and it means nudging or resizing
' the bar in the deck fixes itself, because the slide is the authority on
' position, not the register.
'
' Suffix convention on the role tag rather than new schema: a field's parts are
' discoverable from its own name, and a template that carries them carries the
' whole behaviour with it.
'
' HORIZONTAL ONLY, deliberately, and it says so when it cannot help: a vertical
' bar is the same arithmetic on Top/Height, but guessing the axis from which
' dimension happens to be larger would silently do the wrong thing to a square.
' `trackTag` overrides which shape is measured against, and exists for the
' repeating case: five milestone bars may each carry their own track, or may all
' sit against ONE shared axis. Both are real layouts and the caller resolves
' which, so this function never has to guess. Empty means the usual
' `<field>.track`.
Public Function InjectProgressField(sld As Object, identityTag As String, _
                                    fraction As Double, _
                                    Optional dryRun As Boolean = False, _
                                    Optional trackTag As String = "") As InjectResult
    Dim result As InjectResult

    Dim useTrack As String
    useTrack = trackTag
    If useTrack = "" Then useTrack = identityTag & FieldWiring.TRACK_SUFFIX

    Dim doneShp As Object, trackShp As Object, restShp As Object
    Set doneShp = FindShapeByRoleTag(sld, identityTag)
    Set trackShp = FindShapeByRoleTag(sld, useTrack)
    Set restShp = FindShapeByRoleTag(sld, identityTag & FieldWiring.REST_SUFFIX)

    If doneShp Is Nothing Then
        result.ErrorMessage = "no single shape tagged role=" & identityTag
        InjectProgressField = result
        Exit Function
    End If
    result.Found = True

    ' WHERE THE 100% MARK COMES FROM. Two ways, and never a third.
    '
    ' 1. A TRACK. A shape this code never writes, so it cannot drift. Robust,
    '    and the right answer when the template can carry one.
    '
    ' 2. A DONE + REST PAIR WITH NO TRACK -- the extent is their SUM.
    '
    ' The pair case is how Rohan's decks are already authored. Measured from
    ' slide 1 on 2026-08-10: `Time elapsed` is a 1.45" filled rect beside a
    ' 0.17" grey one, and 1.45 / (1.45 + 0.17) = 89.5% against a label reading
    ' 90%. There is no full-width shape anywhere behind them.
    '
    ' THIS IS NOT THE SHRINKING-BAR BUG, and the distinction is the whole
    ' justification. Measuring the done part against ITSELF compounds: the bar
    ' is shorter after each run, so the next run takes a fraction of a fraction
    ' and walks to zero. The SUM does not compound, because both halves are
    ' written together in one operation and always add back to the same extent.
    ' 1.45 + 0.17 and 0.81 + 0.81 are both 1.62.
    '
    ' WHAT IT COSTS, stated because it is real: the sum is only invariant while
    ' nothing moves one half on its own. Nudge the remainder in the deck and the
    ' extent silently changes, and nothing here can tell -- there is no
    ' authority left to check against. A track cannot fail that way, so a
    ' template that can carry one should.
    Dim extentLeft As Single, extentWidth As Single
    If Not trackShp Is Nothing Then
        extentLeft = trackShp.Left
        extentWidth = trackShp.Width
    ElseIf Not restShp Is Nothing Then
        extentLeft = doneShp.Left
        If restShp.Left < extentLeft Then extentLeft = restShp.Left
        extentWidth = doneShp.Width + restShp.Width
    Else
        result.ErrorMessage = "no shape tagged role=" & useTrack & ", and no " & _
            identityTag & ".rest either -- a progress bar needs either a track to " & _
            "measure against or a remainder to share its extent with. Tag the " & _
            "full-width shape behind the bar, or the grey part beside it."
        InjectProgressField = result
        Exit Function
    End If

    ' OUT OF RANGE IS REPORTED, NOT CLAMPED SILENTLY. A register cell holding
    ' 90 when it meant 0.9 would otherwise draw a full bar and look correct.
    Dim f As Double
    f = fraction
    If f < 0 Or f > 1 Then
        result.ErrorMessage = "progress value " & fraction & " is not between 0 and 1" & _
            " -- the bar was NOT changed. A percentage belongs in the register as 0.9, not 90."
        InjectProgressField = result
        Exit Function
    End If

    Dim wantLeft As Single, wantWidth As Single
    wantLeft = extentLeft
    wantWidth = extentWidth * f

    ' ROUNDED, NOT RAW. `doneShp.Width` is a Single, and a value like 127.2757
    ' sits right at the edge of Single's ~7-significant-digit precision -- CStr
    ' on the raw value can produce a different last digit across two reads of
    ' the exact same, untouched shape, which is invisible to a person and fatal
    ' to a hash comparison expecting byte-identical strings. Found 2026-08-16:
    ' the elapsed-time bar's own review-queue approval was dropped as "changed
    ' since you approved it" on every attempt, never once written, because the
    ' build-time read and the apply-time read of an UNCHANGED shape produced
    ' two different strings. Same defect class as item N (a formatting-
    ' sensitive value reported as a real change) -- rounding to hundredths is
    ' well past the precision that matters on a slide and stable across reads.
    result.CurrentValue = Format(doneShp.Width, "0.00")
    result.WouldChange = (Abs(doneShp.Width - wantWidth) > 0.5) Or (Abs(doneShp.Left - wantLeft) > 0.5)

    If Not result.WouldChange Then
        result.Verified = True
        InjectProgressField = result
        Exit Function
    End If
    If dryRun Then
        result.Verified = True
        InjectProgressField = result
        Exit Function
    End If

    On Error Resume Next
    Err.Clear
    doneShp.Left = wantLeft
    doneShp.Width = wantWidth

    ' The remainder takes what is left of the extent. WRITTEN IN THE SAME
    ' OPERATION as the done part, always -- that is what keeps their sum
    ' invariant and makes the trackless case safe. Writing one without the
    ' other is the only way to turn this into the shrinking-bar bug.
    If Not restShp Is Nothing Then
        restShp.Left = wantLeft + wantWidth
        restShp.Width = extentWidth - wantWidth
    End If

    If Err.Number <> 0 Then
        Dim e As String
        e = Err.Description
        On Error GoTo 0
        result.ErrorMessage = "could not resize the bar (" & e & ")"
        InjectProgressField = result
        Exit Function
    End If
    On Error GoTo 0

    ' VERIFIED FROM THE SHAPE, not from the absence of an error.
    result.Written = True
    result.Verified = (Abs(doneShp.Width - wantWidth) <= 0.5)
    If Not result.Verified Then
        result.ErrorMessage = "the bar was set to " & wantWidth & " but reads " & doneShp.Width
    End If

    InjectProgressField = result
End Function
