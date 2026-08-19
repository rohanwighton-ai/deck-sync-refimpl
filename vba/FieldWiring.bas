Attribute VB_Name = "FieldWiring"
Option Explicit

' IS EVERY FIELD ON THE REGISTER ACTUALLY ATTACHED TO A SHAPE?
'
' Added 2026-08-10, after the third built-but-unreachable finding of one
' morning. The two before it were code-level -- an injector nothing called, a
' marking gate that refused pictures. This one is the same shape at the level a
' PERSON meets it: `BatchOnboardFlow.MarkFieldForBatch` has exactly one call
' site (`RibbonUI.bas`, inside `SyncNowChainCore`) and it sits behind
' `If Not hasTypes`, so on a deck whose slide type is already registered there
' is NO WAY to mark a new field at all. Every field added to the rig after its
' first setup was tagged by Claude over COM, because nothing on the toolbar
' could do it -- which is exactly the dependence Rohan named as the thing that
' must not exist.
'
' A register column with no shape behind it is not a small problem. Sync reads
' the register's columns, so the field is carried all the way to the injector
' and refused there, once per slide, in a report nobody reads to the bottom.
' The work of drafting it was already done by then.
'
' DERIVED, NEVER REMEMBERED. Nothing is stored about which fields are wired:
' the answer is recomputed by looking at the slides, so it cannot go stale the
' way a flag set at setup time would. That is the same reason the router reads
' the SHAPE rather than a column -- see InjectPrimitive.InjectField.
'
' The orphan-track check is the half-marked bar. A progress field is a pair,
' `FIELD` and `FIELD.track`, and a `.track` with no `FIELD` behind it is a
' marking that was started and abandoned. It matters because the failure is
' otherwise SILENT AND PLAUSIBLE: an untracked bar shape is an ordinary
' rectangle with a text frame, so the router sends it to the text writer and it
' quietly becomes a text field. Catching it here is cheaper than catching it on
' a slide.

' EXISTING SLIDES AND THE TEMPLATE ARE TWO QUESTIONS, NOT ONE.
'
' `RunSync.GatherInstances` excludes the template by design (`Not
' resolved.IsTemplate`), and a first version of this module inherited that
' without noticing -- so it would have reported every field wired while the
' TEMPLATE was missing one, and every slide created from then on would have
' quietly lacked that field. New slides are `Slide.Duplicate` copies of the
' template and inherit its SHAPE tags (probed against real PowerPoint
' 2026-08-10, including a suffixed `.track` value on a shape inside a group),
' which is precisely why the template is the slide that decides the future.
'
' Different blast radius, so different lines: a field missing on one existing
' slide costs that slide, and a field missing on the template costs every slide
' not yet made.
' PRESENCE IS TOO EASY A QUESTION, and the first version of this module only
' asked it. On the rig, 2026-08-10: five fields were tagged on all 44 slides
' and THREE were tagged on two -- slide 4 and the template -- and the check
' reported `ok`, because each of those three was carried by *a* slide.
'
' So the count is per slide: `PROBLEM_BODY on 2 of 43`. That is visible BEFORE
' the second project is drafted, which is the moment it starts to matter -- a
' field on 2 of 43 slides costs nothing while only one project has text for it,
' and costs the second project's work the moment someone writes it.
'
' A field part-way across a deck is a REAL state, not an error: it is what a
' deck being extended looks like. So this is reported and never blocks. A
' check that fires on a normal state gets clicked through, and then it is not
' there for the times it matters.
' THE SUFFIXES ARE THE MARKERS, defined once.
'
' A progress field is two or three shapes bound by a naming convention, and
' every one of those relationships is invisible on the slide. Several rules turn
' on the names: the register-exclusion at commit (a companion is tagged, never a
' column), the orphan check below, and the injector's own lookups. Copies of
' ".track" scattered about would drift the first time anyone renamed one, and
' the failure would be silent -- a bar that quietly becomes a text field.
'
' TWO COMPANIONS, NOT ONE, and missing that cost a live run on 2026-08-10.
' Marking asked only ever for a `.track`, so Rohan's `Time elapsed` bar -- a
' fill beside a 0.17" grey REMAINDER, with no track anywhere -- was about to
' have its remainder tagged as the track. The extent would then have been the
' tail's own width and the bar would have been drawn at 90% OF THE TAIL. The
' trackless fill/rest mode had shipped two hours earlier and the marking flow
' knew nothing about it.
Public Const TRACK_SUFFIX As String = ".track"
Public Const REST_SUFFIX As String = ".rest"

Public Type FieldWiringResult
    Unmarked As String          ' register fields NO existing slide carries
    UnmarkedCount As Long
    TemplateUnmarked As String  ' register fields the TEMPLATE does not carry
    TemplateUnmarkedCount As Long
    OrphanTracks As String      ' `X.track` present with no `X`, either place
    OrphanCount As Long
    CaseMismatch As String      ' register name vs the differently-cased tag
    CaseMismatchCount As Long
    Coverage As String          ' "FIELD on X of Y", only when X < Y
    PartialCount As Long
    MissingDetail As String     ' "FIELD missing on: key1, key2" per partial field
    Wired As Long               ' register fields that resolve on some slide
    DeviceOwnedCount As Long    ' register fields owned by a device (e.g.
                                 ' MilestoneDevice's MS1_LABEL..MSn_DONE) --
                                 ' never individually role-tagged by design, so
                                 ' asking "does any slide carry this exact
                                 ' field name" is the wrong question for them.
                                 ' Counted here instead of Unmarked/Wired.
    SlidesScanned As Long
    Scanned As Boolean          ' False = could not look; never report a pass
    TemplateScanned As Boolean  ' False = no template found to look at
End Type

' Does this set of role tags carry `fieldName` at all?
'
' A REPEATING FIELD NEVER CARRIES ITS OWN NAME. The milestone case tags shapes
' `<field>.1`, `.2`, `.3` and nothing is tagged `<field>` -- so a check that
' asked only for the bare name would report every repeating field as unwired,
' on every slide, forever. `.1` is the marker: a repeating field always has a
' first one.
Public Function CarriesField(roleSet As Object, fieldName As String) As Boolean
    If roleSet Is Nothing Then Exit Function
    Dim n As String
    n = UCase(Trim(fieldName))
    If n = "" Then Exit Function
    CarriesField = roleSet.Exists(n) Or roleSet.Exists(n & ".1")
End Function

' Which companion suffix this name ends in, or "" if it is an ordinary field.
'
' Returning the SUFFIX rather than a Boolean is what lets every caller strip it
' correctly without knowing which one matched -- the orphan check needs the base
' name, and hardcoding a length there is how ".rest" would have been mis-stripped
' by a check written for ".track".
Public Function CompanionSuffixOf(fieldName As String) As String
    Dim n As String
    n = UCase(Trim(fieldName))

    If Len(n) > Len(TRACK_SUFFIX) Then
        If Right(n, Len(TRACK_SUFFIX)) = UCase(TRACK_SUFFIX) Then
            CompanionSuffixOf = TRACK_SUFFIX
            Exit Function
        End If
    End If
    If Len(n) > Len(REST_SUFFIX) Then
        If Right(n, Len(REST_SUFFIX)) = UCase(REST_SUFFIX) Then
            CompanionSuffixOf = REST_SUFFIX
            Exit Function
        End If
    End If
End Function

' A companion is a TAG, not a field: it carries no value, and must never become
' a register column. A `.rest` that became one would be worse than useless --
' the router would find no companion for IT, hand it to the text writer, and
' write text into the grey remainder of a progress bar.
Public Function IsCompanionFieldName(fieldName As String) As Boolean
    IsCompanionFieldName = (CompanionSuffixOf(fieldName) <> "")
End Function

' Kept for the track-specific question -- the orphan check reports a track with
' no bar differently from a remainder with no bar, because the remedies differ.
Public Function IsTrackFieldName(fieldName As String) As Boolean
    IsTrackFieldName = (CompanionSuffixOf(fieldName) = TRACK_SUFFIX)
End Function


' Every distinct `role` tag value carried by any shape on any instance of
' `slideType`, groups walked into.
'
' Returns Nothing rather than an empty dictionary when the deck cannot be
' walked, because "no tags found" and "did not look" must not render the same
' -- Readiness rule 1, and the defect FastPathRefusalText earned it for.
Public Function RoleTagsInDeck(slideType As String, ByRef slidesScanned As Long) As Object
    slidesScanned = 0

    Dim instances() As Object
    instances = RunSync.GatherInstances(slideType)

    Dim lo As Long, hi As Long, hasAny As Boolean
    On Error Resume Next
    lo = LBound(instances): hi = UBound(instances)
    hasAny = (Err.Number = 0)
    On Error GoTo 0
    If Not hasAny Then
        Set RoleTagsInDeck = Nothing
        Exit Function
    End If

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = lo To hi
        WalkForRoles instances(i).Shapes, seen
        slidesScanned = slidesScanned + 1
    Next i

    Set RoleTagsInDeck = seen
End Function

' Same group recursion as PlaceholderCheck.WalkForPlaceholders and
' InjectPrimitive.WalkForRoleTag. Real decks nest fields inside grouped "card"
' layouts, and a flat walk under-reports -- which here would mean reporting a
' field as unmarked when it is tagged, and sending someone to re-tag a shape
' that is already correct.
'
' A GROUP IS TESTED **AND** RECURSED INTO, same fix as InjectPrimitive.
' WalkForRoleTag (2026-08-10) and the same reason: this was an ElseIf, so a
' device's own role tag (MILESTONE_TIMELINE, stamped on the group shape
' itself, not on any child) was invisible to this scanner -- masked so far
' because nothing has queried a device's own tag through this dictionary yet,
' but latent for the same reason the InjectPrimitive copy of this bug was
' real: a group carrying a role tag could never be found by anything.
Private Sub WalkForRoles(shapesColl As Object, ByRef seen As Object)
    Dim shp As Object
    For Each shp In shapesColl
        Dim role As String
        role = ""
        On Error Resume Next
        role = shp.Tags("role")
        On Error GoTo 0
        ' KEY uppercased, VALUE the original casing. The injector matches
        ' role tags with `=` under VBA's default binary comparison, so it is
        ' CASE SENSITIVE, while this check has always uppercased both sides.
        ' Keeping the original is what lets the near-miss be reported by name
        ' instead of the two quietly disagreeing.
        If role <> "" Then seen(UCase(Trim(role))) = Trim(role)

        If shp.Type = msoGroup Then
            WalkForRoles shp.GroupItems, seen
        End If
    Next shp
End Sub

' instance key -> the set of role tags on THAT slide, plus the union in `seen`.
'
' Keyed by the RAW instance key, not an upper-cased one, because the caller
' matches it against `Sheet.Rows`, and ReviewQueue already matches raw keys
' against that same dictionary. Two different conventions for one join is how a
' lookup silently finds nothing.
'
' Returns Nothing when the deck could not be walked -- "no instances" and "did
' not look" must not render the same.
Public Function RolesByInstance(slideType As String, ByRef seen As Object) As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim instances() As Object
    instances = RunSync.GatherInstances(slideType)

    Dim lo As Long, hi As Long, hasAny As Boolean
    On Error Resume Next
    lo = LBound(instances): hi = UBound(instances)
    hasAny = (Err.Number = 0)
    On Error GoTo 0
    If Not hasAny Then
        Set RolesByInstance = Nothing
        Exit Function
    End If

    Dim byKey As Object
    Set byKey = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = lo To hi
        Dim inst As SlideInstance
        inst = Resolve.ResolveSlideInstance(instances(i))

        Dim rs As Object
        Set rs = RolesForSlideCached(instances(i))

        Dim k As Variant
        For Each k In rs.Keys
            seen(k) = True
        Next k

        ' An unkeyed slide still has to be COUNTED, or coverage would read
        ' "3 of 2" and look like a bug in the counter rather than a slide
        ' missing its identity.
        Dim key As String
        If inst.HasInstanceKey Then
            key = inst.InstanceKey
        Else
            key = "<unkeyed slide " & instances(i).SlideIndex & ">"
        End If
        If Not byKey.Exists(key) Then byKey.Add key, rs
    Next i

    Set RolesByInstance = byKey
End Function

' Every distinct `role` tag on ONE slide. Used for the template, which
' GatherInstances deliberately excludes.
Public Function RoleTagsOnSlide(sld As Object) As Object
    If sld Is Nothing Then
        Set RoleTagsOnSlide = Nothing
        Exit Function
    End If
    Set RoleTagsOnSlide = RolesForSlideCached(sld)
End Function

' CACHE-FIRST, SAME SHARED INDEX InjectPrimitive.FindShapeByRoleTag ALREADY
' BUILDS AND READS -- ShapeAddressBook.bas's own header. Before this,
' RolesByInstance/RoleTagsOnSlide did a live, uncached `shp.Tags("role")`
' walk of every shape on every slide of a type, every single call -- and
' this scan runs from "1. Set up my quarter"'s own chain
' (OfferMarkingForUnwiredFields), which a person presses every session, not
' once ever. Measured live 2026-08-19: on a 44-slide deck this is the exact
' cost class ShapeAddressBook exists to eliminate (items W/Y/AT), just never
' given the same treatment here.
'
' ADAPTED, NOT SHARED DIRECTLY -- the two caches disagree on shape.
' ShapeAddressBook.SlideTagsFor keys by the tag's OWN case (built from
' InjectPrimitive.WalkForRoleTag's `tagsPresent(roleVal) = True`, plus a
' NAME-based entry for an untagged device group found via SlotCount's
' fallback -- see that function's header). This module's own `seen`/`rs`
' dictionaries key by UPPERCASE, value the original case, because
' CarriesField queries by `UCase(fieldName)` and the case-mismatch check
' (line ~370) needs the original case to compare against. Converting is one
' cheap in-memory pass over a per-slide set (a handful of entries), not a
' second COM walk -- the whole point.
'
' Cache MISS still does the exact same full walk as before (WalkForRoles),
' then records the result so the NEXT call -- by this scan, or by the
' injector's own FindShapeByRoleTag -- takes the fast path instead. Shared
' cache, shared benefit, same self-healing contract ShapeAddressBook
' documents: a stale or wrong entry is never trusted blindly by the
' injector side, and this side only ever reads what a real walk produced.
Private Function RolesForSlideCached(sld As Object) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim slideKey As String
    slideKey = ShapeAddressBook.SlideKeyFor(sld)

    Dim cached As Object
    If slideKey <> "" Then Set cached = ShapeAddressBook.SlideTagsFor(slideKey)

    If Not cached Is Nothing Then
        Dim ck As Variant
        For Each ck In cached.Keys
            Dim original As String
            original = CStr(ck)
            result(UCase(Trim(original))) = original
        Next ck
        Set RolesForSlideCached = result
        Exit Function
    End If

    ' MISS -- the existing, correct, ambiguity-checked walk, unchanged.
    WalkForRoles sld.Shapes, result

    If slideKey <> "" Then
        Dim toCache As Object
        Set toCache = CreateObject("Scripting.Dictionary")
        Dim rk As Variant
        For Each rk In result.Keys
            toCache(CStr(result(rk))) = True
        Next rk
        ShapeAddressBook.RecordSlideTags slideKey, toCache
    End If

    Set RolesForSlideCached = result
End Function

' `fields` is the register's own field collection (ExcelOutput.Sheet.Fields), so
' the question asked is exactly "does the spreadsheet name something the slides
' do not carry" -- not a question about a list maintained somewhere else.
'
' `templateSld` may be Nothing; that is reported as not-scanned rather than as a
' clean template, because "there is no template to check" and "the template is
' fine" must never render the same.
Public Function ScanFieldWiring(slideType As String, fields As Collection, _
                                templateSld As Object) As FieldWiringResult
    Dim result As FieldWiringResult

    Dim tmpl As Object
    Set tmpl = RoleTagsOnSlide(templateSld)
    result.TemplateScanned = Not (tmpl Is Nothing)

    ' PER SLIDE, not merged. The union answers "is this tagged anywhere", which
    ' is the question that reported ok over a field present on 2 slides of 43.
    Dim rolesByKey As Object
    Dim seen As Object
    Set rolesByKey = RolesByInstance(slideType, seen)

    If rolesByKey Is Nothing Then
        result.Scanned = False
        ScanFieldWiring = result
        Exit Function
    End If

    result.Scanned = True
    result.SlidesScanned = rolesByKey.Count

    Dim f As Variant
    If Not fields Is Nothing Then
        For Each f In fields
            Dim fieldName As String
            fieldName = UCase(Trim(CStr(f)))
            If fieldName <> "" Then
                ' DEVICE-OWNED COLUMNS ARE NOT AN ORDINARY FIELD QUESTION.
                ' Skip the whole carrier/unmarked/case-mismatch machinery for
                ' them -- see MilestoneDevice.IsColumnForThisDevice's header
                ' for why "does a slide role-tag carry this exact name" is the
                ' wrong question to ask of MS1_LABEL..MSn_DONE. This is what
                ' used to report as "21 fields on the register that no slide
                ' carries" every single run.
                If MilestoneDevice.IsColumnForThisDevice(fieldName) Then
                    result.DeviceOwnedCount = result.DeviceOwnedCount + 1
                    GoTo NextField
                End If

                Dim carriers As Long
                carriers = 0
                Dim missingKeys As String
                Dim missingCount As Long
                missingKeys = ""
                missingCount = 0

                ' NAMED, BUT BOUNDED. A deck with many gaps on one field
                ' (the realistic case -- see MissingDetail's own header)
                ' can have dozens of missing slide keys; listing every one
                ' is how a modal downstream ends up silently mid-word
                ' truncated (MsgBox's own real, undocumented ~1024-char
                ' limit -- confirmed live 2026-08-19, same failure class
                ' FieldSpec.ApplyControlledValidation was already fixed
                ' for). The COUNT stays exact regardless; only the NAMED
                ' list is capped.
                Const MAX_NAMED_MISSING_KEYS As Long = 6

                Dim k As Variant
                For Each k In rolesByKey.Keys
                    If CarriesField(rolesByKey(k), fieldName) Then
                        carriers = carriers + 1
                    Else
                        missingCount = missingCount + 1
                        If missingCount <= MAX_NAMED_MISSING_KEYS Then
                            If missingKeys <> "" Then missingKeys = missingKeys & ", "
                            missingKeys = missingKeys & CStr(k)
                        End If
                    End If
                Next k
                If missingCount > MAX_NAMED_MISSING_KEYS Then
                    missingKeys = missingKeys & ", and " & _
                        (missingCount - MAX_NAMED_MISSING_KEYS) & " more"
                End If

                ' THE NEAR-MISS. `PROJECT_PROGRESS` on the register against a
                ' shape tagged `Project_Progress` passes this check (it
                ' uppercases) and FAILS at write time (the injector does not).
                ' That combination is a green readiness sheet over a sync that
                ' writes nothing -- this project's defining defect, arriving
                ' from the check rather than the writer.
                Dim actualTag As String
                actualTag = ""
                Dim kk As Variant
                For Each kk In rolesByKey.Keys
                    If rolesByKey(kk).Exists(fieldName) Then
                        actualTag = CStr(rolesByKey(kk)(fieldName))
                        Exit For
                    End If
                Next kk
                If actualTag = "" And result.TemplateScanned Then
                    If tmpl.Exists(fieldName) Then actualTag = CStr(tmpl(fieldName))
                End If

                If actualTag <> "" Then
                    If StrComp(actualTag, CStr(f), vbBinaryCompare) <> 0 Then
                        result.CaseMismatchCount = result.CaseMismatchCount + 1
                        If result.CaseMismatch <> "" Then result.CaseMismatch = result.CaseMismatch & ", "
                        result.CaseMismatch = result.CaseMismatch & _
                            "register '" & CStr(f) & "' vs slide '" & actualTag & "'"
                    End If
                End If

                If carriers > 0 Then
                    result.Wired = result.Wired + 1
                Else
                    result.UnmarkedCount = result.UnmarkedCount + 1
                    If result.Unmarked <> "" Then result.Unmarked = result.Unmarked & ", "
                    result.Unmarked = result.Unmarked & CStr(f)
                End If

                If carriers > 0 And carriers < result.SlidesScanned Then
                    result.PartialCount = result.PartialCount + 1
                    If result.Coverage <> "" Then result.Coverage = result.Coverage & ", "
                    result.Coverage = result.Coverage & CStr(f) & " on " & carriers & _
                        " of " & result.SlidesScanned

                    ' NAMES THE SLIDES, NOT JUST THE COUNT -- Rohan, 2026-08-19:
                    ' "which field and which slides", one consolidated modal.
                    ' `missingKeys` (built above, in the same pass, no second
                    ' walk) lists the instance keys that do NOT carry this
                    ' field -- the actionable direction ("go fix these"), not
                    ' the larger set that already has it.
                    If result.MissingDetail <> "" Then result.MissingDetail = result.MissingDetail & vbCrLf
                    result.MissingDetail = result.MissingDetail & CStr(f) & " missing on: " & missingKeys
                End If

                If result.TemplateScanned Then
                    If Not CarriesField(tmpl, fieldName) Then
                        result.TemplateUnmarkedCount = result.TemplateUnmarkedCount + 1
                        If result.TemplateUnmarked <> "" Then _
                            result.TemplateUnmarked = result.TemplateUnmarked & ", "
                        result.TemplateUnmarked = result.TemplateUnmarked & CStr(f)
                    End If
                End If
            End If
NextField:
        Next f
    End If

    ' A `.track` whose field is missing. Checked against the tags actually on
    ' the slides rather than against the register, because a bar's track is
    ' never a register column -- it is measured, never written.
    CollectOrphanTracks seen, "", result
    If result.TemplateScanned Then CollectOrphanTracks tmpl, "template ", result

    ScanFieldWiring = result
End Function

Private Sub CollectOrphanTracks(seen As Object, prefix As String, ByRef result As FieldWiringResult)
    Dim k As Variant
    For Each k In seen.Keys
        Dim tag As String
        tag = CStr(k)
        Dim suffix As String
        suffix = CompanionSuffixOf(tag)
        If suffix <> "" Then
            Dim baseName As String
            baseName = Left(tag, Len(tag) - Len(suffix))
            If Not seen.Exists(baseName) Then
                result.OrphanCount = result.OrphanCount + 1
                If result.OrphanTracks <> "" Then result.OrphanTracks = result.OrphanTracks & ", "
                result.OrphanTracks = result.OrphanTracks & prefix & tag
            End If
        End If
    Next k
End Sub

' The sentence the readiness sheet shows. Built here so the wording and the
' counts cannot drift apart, and so every count NAMES WHAT IT COUNTED -- the
' fix-list's item 1a, which has now been true-and-unusable four separate times.
Public Function WiringText(r As FieldWiringResult) As String
    If Not r.Scanned Then
        WiringText = "could not read the slides"
        Exit Function
    End If

    If r.UnmarkedCount = 0 And r.OrphanCount = 0 And r.TemplateUnmarkedCount = 0 _
       And r.CaseMismatchCount = 0 Then
        Dim ok As String
        ok = r.Wired & " field(s) tagged on " & r.SlidesScanned & " slide(s)"

        ' THE TEMPLATE CLAUSE GOES HERE, BEFORE THE COVERAGE LIST. Appended
        ' after it, the sentence read "...PROGRESS_BODY on 1 of 43, and on the
        ' template", which parses as a fourth entry in the coverage list rather
        ' than the separate reassurance it is. Seen on the real deck 2026-08-10.
        '
        ' Stated either way: a silent omission would be the same defect the
        ' template split exists to fix -- a clean-looking line that never looked
        ' at the slide the future is made from.

        ' NOT A PROBLEM TODAY, AND WORTH KNOWING BEFORE DRAFTING. A field on 1
        ' of 43 slides strands nothing while only that one project has text for
        ' it; it strands the SECOND project's work, after that work is done.
        If r.TemplateScanned Then
            ok = ok & ", and on the template"
        Else
            ok = ok & " -- NO TEMPLATE was checked"
        End If
        If r.PartialCount > 0 Then
            ok = ok & ". Not on every slide yet: " & r.Coverage
        End If

        WiringText = ok
        Exit Function
    End If

    Dim s As String
    If r.CaseMismatchCount > 0 Then
        s = r.CaseMismatchCount & " field(s) spelled differently on the slide -- these will NOT " & _
            "match when syncing, because tags are matched exactly: " & r.CaseMismatch
    End If
    If r.UnmarkedCount > 0 Then
        If s <> "" Then s = s & ". "
        s = s & r.UnmarkedCount & " field(s) on the register that no slide carries: " & r.Unmarked
    End If
    If r.PartialCount > 0 Then
        If s <> "" Then s = s & ". "
        s = s & "Not on every slide: " & r.Coverage
    End If
    If r.TemplateUnmarkedCount > 0 Then
        If s <> "" Then s = s & ". "
        s = s & r.TemplateUnmarkedCount & " field(s) missing from the TEMPLATE, so every " & _
            "new slide will lack them: " & r.TemplateUnmarked
    End If
    If r.OrphanCount > 0 Then
        If s <> "" Then s = s & ". "
        s = s & r.OrphanCount & " progress track(s) with no bar: " & r.OrphanTracks
    End If
    WiringText = s
End Function

' WHAT ACTUALLY BLOCKS, for a dialog that is about to stop someone.
'
' WiringText carries everything, including partial coverage -- which is
' deliberately NOT a problem and was never meant to prompt. Printing the whole
' sentence in the blocking dialog put three informational "on 1 of 43" lines
' beside the one real blocker, under a heading saying syncing would carry those
' fields and refuse them. That is true of the unwired field and false of the
' other three, whose rows are blank so nothing is carried at all.
'
' So the prompt gets this, and the START HERE sheet keeps the full picture.
Public Function BlockingText(r As FieldWiringResult) As String
    Dim s As String
    If r.CaseMismatchCount > 0 Then
        s = r.CaseMismatchCount & " field(s) spelled differently on the slide -- these will NOT " & _
            "match when syncing: " & r.CaseMismatch
    End If
    If r.UnmarkedCount > 0 Then
        If s <> "" Then s = s & vbCrLf
        s = s & r.UnmarkedCount & " field(s) on the register that no slide carries: " & r.Unmarked
    End If
    If r.TemplateUnmarkedCount > 0 Then
        If s <> "" Then s = s & vbCrLf
        s = s & r.TemplateUnmarkedCount & " field(s) missing from the TEMPLATE, so every new " & _
            "slide will lack them: " & r.TemplateUnmarked
    End If
    If r.OrphanCount > 0 Then
        If s <> "" Then s = s & vbCrLf
        s = s & r.OrphanCount & " progress bar part(s) with no bar: " & r.OrphanTracks
    End If
    ' NAMES THE SLIDES for the one case where naming them is actionable --
    ' see MissingDetail's own header comment. Unmarked/TemplateUnmarked
    ' don't get the same treatment: "no slide carries this" already implies
    ' "every slide", and "missing from the template" already names the one
    ' place that matters.
    If r.PartialCount > 0 And r.MissingDetail <> "" Then
        If s <> "" Then s = s & vbCrLf
        s = s & r.MissingDetail
    End If
    BlockingText = s
End Function

' COUNTS ONLY, NO NAMES -- for a dialog that has to share its budget with
' other content, not read it alone. Rohan, 2026-08-19, choosing this over
' BlockingText's full field/slide detail once the two turned out not to
' fit together: "folded [into the combined 'Set up my quarter' summary],
' I think" -- confirmed live the same night that BlockingText's full
' output, however tightly capped on its own, still gets chopped by the
' OUTER cap wrapping that combined dialog's other three sections, because
' those alone already use most of the shared character budget. This
' trades the field/slide NAMES for something that reliably fits instead
' of something that reliably gets cut mid-word.
Public Function CoverageSummaryLine(r As FieldWiringResult) As String
    Dim parts As String
    If r.UnmarkedCount > 0 Then
        parts = AddCommaPart(parts, r.UnmarkedCount & " not on any slide")
    End If
    If r.TemplateUnmarkedCount > 0 Then
        parts = AddCommaPart(parts, r.TemplateUnmarkedCount & " missing from the template")
    End If
    If r.PartialCount > 0 Then
        parts = AddCommaPart(parts, r.PartialCount & " partially covered")
    End If
    If r.CaseMismatchCount > 0 Then
        parts = AddCommaPart(parts, r.CaseMismatchCount & " spelled differently on the slide")
    End If
    If r.OrphanCount > 0 Then
        parts = AddCommaPart(parts, r.OrphanCount & " progress bar part(s) with no bar")
    End If
    If parts = "" Then Exit Function
    CoverageSummaryLine = "field(s): " & parts & "."
End Function

Private Function AddCommaPart(existing As String, addition As String) As String
    If existing = "" Then
        AddCommaPart = addition
    Else
        AddCommaPart = existing & ", " & addition
    End If
End Function
