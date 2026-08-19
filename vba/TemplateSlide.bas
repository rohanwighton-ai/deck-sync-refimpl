Attribute VB_Name = "TemplateSlide"
Option Explicit

' Progression step 1 of specs/deck-compiler-concept.md: give each slide type
' a MASTER TEMPLATE SLIDE that is never a real project.
'
' The hazard this closes. Until now the slide a new record gets cloned from
' was whichever real slide happened to be registered at onboarding time
' (DeckRegistry.RegisterType stores its SlideID; RunSync.RunRoutineSync
' clones it for every unmatched Data row). Every field the sync knows about
' gets overwritten on the copy -- but everything it does NOT know about does
' not: figures, chart data, speaker notes, any text that was never tagged as
' a field. So a brand-new project's slide arrives carrying a different
' project's numbers, and it looks entirely plausible, because the tagged
' fields are all correct. That is the worst shape a reporting-tool defect
' can take.
'
' Why a whole operation rather than "add a slide by hand". Three things have
' to be true at once for a template to be safe, and none of them is visible
' in the PowerPoint UI:
'   1. tagged is_template, so RunSync.GatherInstances excludes it (otherwise
'      it is reported as case 6 unclassified_slide on every single sync)
'   2. NOT tagged instance_key, so it can never match a Data row
'   3. registered as its type's template, so cloning actually uses it
' Getting one of the three wrong fails quietly. Doing it in code means the
' three cannot come apart.
'
' D3 in the concept spec ("one deck per slide type") is what makes this
' natural rather than awkward: one deck, one type, exactly one template.
'
' NOT built here, deliberately: creating the template as part of onboarding.
' Every existing deck is already onboarded, so a retrofit operation is
' needed regardless, and it is the same operation -- a new deck can run it
' immediately after its first onboard. Folding it into the onboarding flows
' (three of them: OnboardFlow, BatchOnboardFlow, DeckAdoption) before it has
' been used once would be building the second version first.

' Declarations first, procedures after. VBA requires it -- a Type declared
' below the module's first Function is a compile error, and it does NOT
' report itself here: it surfaces as "User-defined type not defined" at
' whichever OTHER module referenced the Type, which reads as a missing
' module rather than a misplaced declaration. Cost one full 8-minute suite
' run on 2026-07-30 to learn, and was then identified in seconds from a
' screenshot of the VBE.
Public Type MakeTemplateResult
    Ok As Boolean
    Reason As String            ' populated on refusal/failure
    NewSlide As Object          ' the created template slide, if Ok
    FieldCount As Long          ' fields blanked to placeholders
End Type

' The text every field on the template carries. ASCII on purpose -- .bas
' files are imported by COM automation and non-ASCII (guillemets were the
' first choice) is an encoding risk for zero benefit.
'
' Placeholder text rather than empty fields, because the two fail
' differently when a template leaks into a real deck by mistake: an empty
' field reads as a real slide with missing data and invites someone to fill
' it in, while "<<Project Name>>" cannot be read as anything but scaffolding.
' It is also the whole of the template's visible signposting -- see
' MakeTemplateFrom's note on why no marker SHAPE is added.
Public Function PlaceholderFor(role As String) As String
    PlaceholderFor = "<<" & role & ">>"
End Function

' Turns `sourceSld` (a real, already-onboarded instance of `slideType`) into
' the type's master template, by COPYING it -- the source slide is left
' completely untouched and stays an ordinary record. Nothing here edits a
' slide a human authored; it only ever writes to the copy it just made.
'
' Does NOT register the result. Registration is the caller's move (see
' RibbonUI.CreateTemplateSlideCore) so this function stays testable against
' a throwaway presentation with no registry entries at all.
Public Function MakeTemplateFrom(sourceSld As Object, slideType As String) As MakeTemplateResult
    Dim result As MakeTemplateResult

    ' --- Refuse to template a template ------------------------------------
    ' Idempotency guard with teeth: running this twice would otherwise
    ' produce a second template, and a type with two templates has no
    ' defined behaviour -- LookupType returns whichever SlideID was
    ' registered last, so which one gets cloned depends on click order.
    If Resolve.IsTemplateSlide(sourceSld) Then
        result.Ok = False
        result.Reason = "that slide is already a master template -- a type must have exactly one, so there is nothing to do."
        MakeTemplateFrom = result
        Exit Function
    End If

    Dim sourceInstance As SlideInstance
    sourceInstance = Resolve.ResolveSlideInstance(sourceSld)

    If Not sourceInstance.HasTypeTag Then
        result.Ok = False
        result.Reason = "that slide carries no slide_type tag -- only an already-onboarded slide can be turned into its type's template."
        MakeTemplateFrom = result
        Exit Function
    End If

    If sourceInstance.TypeTag <> slideType Then
        result.Ok = False
        result.Reason = "that slide is type '" & sourceInstance.TypeTag & "', not '" & slideType & _
            "' -- refusing to build one type's template out of another type's slide."
        MakeTemplateFrom = result
        Exit Function
    End If

    ' --- Copy -------------------------------------------------------------
    Dim newSlides As Object
    Set newSlides = sourceSld.Duplicate()
    Dim newSld As Object
    Set newSld = newSlides(1)

    ' --- Identity: typed, marked, and deliberately KEYLESS -----------------
    ' Order matters. is_template goes on BEFORE instance_key comes off, so
    ' there is no window in which this slide is a typed keyless instance --
    ' which is the state PlanRoutineSync flags as unclassified. Nothing runs
    ' concurrently here, so the window is theoretical; the ordering costs
    ' nothing and means a reader never has to work that out.
    newSld.Tags.Add Resolve.TEMPLATE_TAG_NAME, "1"
    newSld.Tags.Add "slide_type", slideType

    On Error Resume Next
    newSld.Tags.Delete "instance_key"
    On Error GoTo 0

    ' Postcondition on both halves, because both can fail silently and in
    ' opposite directions: a template that kept its instance_key would
    ' double-claim a real Data row, and one that never took the marker would
    ' be flagged as unclassified forever. Neither is visible in PowerPoint.
    ' This guard can genuinely fail -- Tags.Delete's behaviour on this build
    ' is not established (see SlideDuplication.bas's matching note).
    Dim check As SlideInstance
    check = Resolve.ResolveSlideInstance(newSld)
    If check.HasInstanceKey Or Not check.IsTemplate Then
        newSld.Delete
        result.Ok = False
        result.Reason = "could not give the copy a clean template identity" & _
            " (is_template=" & check.IsTemplate & ", instance_key='" & check.InstanceKey & "')" & _
            " -- refusing to leave a half-marked slide in the deck. Copy removed."
        MakeTemplateFrom = result
        Exit Function
    End If

    ' --- Replace every field's value with a placeholder --------------------
    ' Reuses the same two functions the sync itself uses to find and write a
    ' type's fields (BuildTemplateFieldShapes + InjectPrimitive), rather
    ' than walking shapes a second way -- if the sync can find a field, so
    ' can this, by construction.
    '
    ' Untagged content is NOT touched, and cannot be: this operation has no
    ' way to know which figure or sentence belongs to the source project.
    ' That is the honest limit of step 1 -- it guarantees new slides no
    ' longer inherit another project's FIELD values, and it makes any
    ' remaining inherited content visible on a slide whose fields all read
    ' "<<...>>", where before it was camouflaged by correct-looking data.
    Dim templateRoles() As String
    Dim templateFieldShapes() As Candidate
    templateFieldShapes = Onboarding.BuildTemplateFieldShapes(newSld, templateRoles)

    Dim rLo As Long, rHi As Long, hasRoles As Boolean
    On Error Resume Next
    rLo = LBound(templateRoles)
    rHi = UBound(templateRoles)
    hasRoles = (Err.Number = 0)
    On Error GoTo 0

    Dim fieldCount As Long
    fieldCount = 0
    If hasRoles Then
        Dim r As Long
        For r = rLo To rHi
            InjectPrimitive.InjectPrimitive newSld, templateRoles(r), PlaceholderFor(templateRoles(r))
            fieldCount = fieldCount + 1
        Next r
    End If

    ' --- Refuse a template that already carries a name collision -----------
    ' The one guarantee this operation can actually make (see
    ' FindDuplicateShapeName's own header) -- checked here, on the COPY,
    ' after duplication so it also catches a collision introduced by the
    ' duplicate itself, not just one inherited from the source.
    Dim collidingName As String
    collidingName = FindDuplicateShapeName(newSld.Shapes)
    If collidingName <> "" Then
        newSld.Delete
        result.Ok = False
        result.Reason = "two shapes on this slide are both named '" & collidingName & "' -- " & _
            "refusing to register a template where a shape name is ambiguous. Rename one of " & _
            "them on the SOURCE slide (Text 216a on real slide 27 is exactly this problem) and " & _
            "try again. Copy removed."
        MakeTemplateFrom = result
        Exit Function
    End If

    ' --- Keep it out of the presented deck ---------------------------------
    ' Hidden, not deleted-at-export: the template has to stay in the file
    ' (it is what gets cloned) but must never appear in a slideshow or a
    ' handout. This is the one thing about it a human WILL see in the
    ' PowerPoint UI -- a greyed, struck-through slide number -- which is
    ' useful signposting on top of being correct.
    newSld.SlideShowTransition.Hidden = msoTrue

    ' --- Park it at the end ------------------------------------------------
    ' Last, not first, and the reason is ResequenceByRowOrder. It packs a
    ' type's instances into contiguous positions starting at the lowest
    ' index any KEYED slide occupies, so a template sitting among them gets
    ' shuffled to wherever the packing pushes it -- correct, but arbitrary,
    ' and it moves on every sync that reorders anything. Parked at the end
    ' it sits outside that range permanently and stays put.
    newSld.MoveTo Application.ActivePresentation.Slides.count

    result.FieldCount = fieldCount
    result.Ok = True
    Set result.NewSlide = newSld
    MakeTemplateFrom = result
End Function

' What the user is agreeing to, before anything is written. This operation
' adds a slide AND re-points the type's registration, so it changes what a
' future Sync Now will clone -- which is a bigger consequence than "one new
' slide" sounds, and the reason it is spelled out rather than summarised.
'
' A pure function, like RunSync.ConfirmSyncText, so a test can pin the
' wording. C3's lesson (2026-07-30) was that a writing action reaching the
' toolbar without a confirmation is one click from damage; the wording IS
' the guard here, so it is worth asserting rather than hand-checking.
' letter/willClaimFallback let this stay accurate after Scenario 3 (per-letter
' templates) without guessing: found live 2026-08-16 that the unconditional
' "'<type>' RE-REGISTERED" wording is FALSE the moment a type already has one
' template and a second letter is being added -- only that letter's own slot
' gets registered; the existing fallback is deliberately left alone
' (DeckRegistry.RegisterNewTemplateLetter). Caught by Rohan reading the actual
' dialog before clicking through it, not by this pinned test, which still
' passed the whole time -- it only ever exercised the letter="" case.
'
' willClaimFallback is the caller's job to compute (same predicate
' RegisterNewTemplateLetter itself uses: does a real template already hold
' the type-level slot), not this function's -- it only WRITES text from a
' fact it is handed, it does not go looking for the fact itself.
Public Function ConfirmTemplateText(slideType As String, sourceLabel As String, fieldCount As Long, _
                                     Optional letter As String = "", Optional willClaimFallback As Boolean = False) As String
    Dim registrationLine As String
    If letter = "" Then
        registrationLine = "    '" & slideType & "' RE-REGISTERED to clone this new slide from now on"
    ElseIf willClaimFallback Then
        registrationLine = "    '" & slideType & "' letter '" & letter & "' registered to clone this new slide from now on" & vbCrLf & _
            "    -- the FIRST template for this type, so it ALSO becomes the default for any row with no letter"
    Else
        registrationLine = "    '" & slideType & "' letter '" & letter & "' registered to clone this new slide from now on" & vbCrLf & _
            "    -- ONLY letter '" & letter & "' rows; other letters keep cloning from their own templates"
    End If

    Dim s As String
    s = "This will change the deck." & vbCrLf & vbCrLf & _
        "Type:   " & slideType & vbCrLf & _
        "Copy of: " & sourceLabel & vbCrLf & vbCrLf & _
        "    1 new slide, added at the END and hidden from the slideshow" & vbCrLf & _
        "    its " & fieldCount & " field(s) replaced with " & PlaceholderFor("placeholders") & vbCrLf & _
        registrationLine & vbCrLf & vbCrLf & _
        "The slide it is copied from is NOT touched -- it stays an" & vbCrLf & _
        "ordinary project record." & vbCrLf & vbCrLf & _
        "Why: new slides are currently cloned from a real project's" & vbCrLf & _
        "slide, so anything the sync does not manage (figures, notes," & vbCrLf & _
        "untagged text) arrives on the new slide belonging to that" & vbCrLf & _
        "project, looking correct." & vbCrLf & vbCrLf & _
        "Proceed?"
    ConfirmTemplateText = s
End Function

' Which slide is `slideType`'s master template, if it has one. Scans rather
' than reading the registry on purpose: the registry answers "what will be
' cloned", this answers "what IS a template", and step 1 exists precisely
' because those two were allowed to be different. A caller wanting to know
' whether they have drifted apart needs both.
'
' Returns Nothing when the type has no template yet -- the normal state for
' every deck onboarded before today, not an error.
' WHICH TEMPLATE A PROJECT WANTS, DERIVED FROM WHAT THE REGISTER ALREADY HOLDS.
'
' Rohan, 2026-08-15: "P projects are green and normal research project, orange
' projects are kickstart projects (K), S projects are PhD (student) project,
' purple. You can tell by the letter in their project code after the underscore."
'
' So the colour is not a new fact to store -- it IS the code letter, and this
' invents no column, no tag vocabulary and no term. `Kind` was NOT reused: it is
' already an axis word here (Controlled/Prose/Static/Derived), and a word doing
' two jobs is the defect that has already cost this project a feature.
'
' TWO KEY SHAPES EXIST IN THE REAL DECK and a parser assuming one would silently
' mishandle five projects. Counted from the deck's own tags, 43 keys:
'   38 of the form <theme>_<letter><digits>   e.g. 3_P001
'    5 with NO underscore at all              e.g. S023, P008, S009, S021, S022
' Letter counts K=15, P=11, S=17 -- which match Rohan's slide ranges 12-26,
' 1-11 and 27-43 exactly, so the letter and the colour are the same fact
' confirmed from two independent directions.
'
' Returns "" rather than guessing when there is no letter to read. "" means
' "no opinion, use the type's unlettered template", which is what keeps today's
' single-template decks working unchanged.
' THE PROMISE THIS PROJECT CAN ACTUALLY KEEP. `Matching.bas`'s name tie-break
' (see its own header, `fdee2e6`) needs a real slide's shape name to match the
' template's exactly -- and nothing in the shipped add-in ever renames a
' shape, by design (the alternative is guessing which shape is which and
' risking a silent wrong answer, worse than refusing). So the add-in cannot
' promise "any instance heals itself." What it CAN promise is "nothing built
' from a clean template drifts into this" -- but only if the template itself
' is verified clean at the one moment that matters: when it is created.
'
' FIX-LIST item C (slide 27's `Text 216a` collision) is exactly the failure
' this closes off at the source: a shape squatting on a name another field
' needs. That collision was never possible to prevent on an already-drifted
' instance; it is entirely possible to prevent on every template made from
' here on.
'
' Recurses into groups (the same reason FormattingAudit.CollectSpecimens
' does) -- a colliding name inside a group is exactly as capable of stealing
' the name InjectPrimitive's rebuild path needs as a top-level shape.
' Returns the first repeated name found, or "" if every name is unique.
' Blank names (an unnamed placeholder) are not a collision with each other --
' PowerPoint does not use "" as an addressing key the way a real name is used
' here, and requiring every shape to carry a distinct name would fail
' templates this project does not touch.
Private Function FindDuplicateShapeName(shapesColl As Object) As String
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    FindDuplicateShapeName = FindDuplicateShapeNameCore(shapesColl, seen)
End Function

Private Function FindDuplicateShapeNameCore(shapesColl As Object, ByRef seen As Object) As String
    Dim shp As Object
    For Each shp In shapesColl
        If shp.Type = msoGroup Then
            Dim inGroup As String
            inGroup = FindDuplicateShapeNameCore(shp.GroupItems, seen)
            If inGroup <> "" Then
                FindDuplicateShapeNameCore = inGroup
                Exit Function
            End If
        ElseIf Trim$(shp.Name) <> "" Then
            If seen.Exists(shp.Name) Then
                FindDuplicateShapeNameCore = shp.Name
                Exit Function
            End If
            seen.Add shp.Name, True
        End If
    Next shp
End Function

Public Function CodeLetterOf(instanceKey As String) As String
    Dim s As String
    s = Trim(instanceKey)
    If s = "" Then Exit Function

    Dim p As Long
    p = InStrRev(s, "_")
    If p > 0 Then s = Mid(s, p + 1)
    If s = "" Then Exit Function

    Dim ch As String
    ch = UCase(Left(s, 1))

    ' A letter or nothing. A key like "1_2003" has no variant to read, and
    ' returning its first digit would silently invent a fourth colour.
    If ch < "A" Or ch > "Z" Then Exit Function

    CodeLetterOf = ch
End Function

' The guard CreateTemplateSlideCore enforces before making a new template:
' one per type/letter pair, never a silent second one (Scenario 3 step 4).
' Returns the BLOCKING template slide if one already exists for this exact
' type+letter, or -- when letter is "" (CodeLetterOf's own "no opinion"
' convention, for a deck with no letter axis at all, or a key like "1_2003"
' with nothing to read) -- for the type overall, the same one-per-type check
' this generalises. Returns Nothing when it is safe to proceed.
Public Function ExistingTemplateForLetter(pres As Object, slideType As String, letter As String) As Object
    If letter <> "" Then
        Dim existing As Object
        Dim ws As String
        DeckRegistry.LookupTemplateLetter pres, slideType, letter, existing, ws
        Set ExistingTemplateForLetter = existing
    Else
        Set ExistingTemplateForLetter = FindTemplateFor(slideType)
    End If
End Function

Public Function FindTemplateFor(slideType As String) As Object
    Dim sld As Object
    For Each sld In Application.ActivePresentation.Slides
        Dim resolved As SlideInstance
        resolved = Resolve.ResolveSlideInstance(sld)
        If resolved.IsTemplate And resolved.HasTypeTag And resolved.TypeTag = slideType Then
            Set FindTemplateFor = sld
            Exit Function
        End If
    Next sld
End Function
