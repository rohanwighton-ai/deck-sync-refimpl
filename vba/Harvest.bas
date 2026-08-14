Attribute VB_Name = "Harvest"
Option Explicit

' READING VALUES OFF A SLIDE INTO THE REGISTER -- the direction this tool did
' not have.
'
' Why it exists. Adoption (DeckAdoption.PlanAdoption) harvests, but only from a
' slide it has never seen: its first act is to skip any slide already carrying
' slide_type + instance_key, BEFORE any matching runs (DeckAdoption.bas:149).
' Every slide in a live deck is already linked, so on a real deck adoption
' harvests nothing, by construction rather than by accident. That was found on
' 2026-08-14 while trying to pull 29 newly-tagged fields off 43 linked slides,
' and the honest summary was that the capability did not exist.
'
' What it does. For a LINKED slide, for each field the register already has a
' column for, find the shape carrying that role tag and copy its text into the
' row for that slide and period.
'
' THE SAFETY PROPERTY IS STRUCTURAL, NOT CAREFUL: a cell is written ONLY when
' the register has nothing there. Not "only when it looks stale", not "only
' when the user confirms" -- only when empty. So a harvest cannot destroy
' authored text no matter how wrong its matching is, and re-running it is safe
' by construction. This is the same rule that made the 27-paragraph restore on
' 2026-08-14 survivable without a wholesale file restore, promoted from a
' one-off tactic to the shape of the operation.
'
' Emptiness is read STRUCTURALLY. ExcelOutput.ReadSheetForPeriod puts a field
' in a row's dictionary only when the cell Is Not IsEmpty, so "the register has
' nothing here" is `Not rowValues.Exists(field)` -- it is not a comparison
' against "", which would treat a deliberately-blanked cell as never-filled.
'
' WHAT IT DOES NOT DO, NAMED RATHER THAN DISCOVERED LATER:
'
' 1. IT HARVESTS ONLY WHAT PUBLISH WOULD WRITE AS TEXT, asked via
'    InjectPrimitive.InjectorFor -- the SAME decision the router makes, so the
'    two can never disagree about a field.
'
'    The slide shows a formatted VIEW of some fields, not the value the register
'    holds. `PROJECT_PROGRESS` displays `33%`; InjectProgressVia refuses anything
'    non-numeric and names `'90%'` as wrong in those exact words. Harvesting what
'    is displayed would write a value the tool itself cannot publish -- and the
'    empty-cell rule above then makes it PERMANENT, because the cell is no longer
'    empty for a corrected run to fill. Caught on the real deck 2026-08-15 at the
'    dialog, before anything was written.
'
'    A milestone timeline is one tagged group standing for 21 register fields,
'    and reading it back means recovering each slot's on/now/off state from shape
'    visibility -- GAP 3, unbuilt. A picture's register cell holds a source id,
'    not a caption. All are reported by name, never silently passed over.
'
'    REFUSAL, NOT CONVERSION. `33%` -> `0.33` is guessable; `33% (est.)` and
'    `as at Q3` are not, and a wrong guess here cannot be corrected. Refusing
'    leaves the cell empty, which is recoverable.
' 2. It never creates a register column. UpsertRow would happily append one;
'    this asks only for fields the sheet already has, so a typo in a role tag
'    shows up as "no column for this role" instead of quietly inventing one.
' 3. It never invents a row. A slide with no instance key is reported and left.

Public Type HarvestOutcome
    Ran As Boolean
    Problem As String
    InstanceKey As String
    Written As Long
    SkippedHasValue As Long
    SkippedBlankOnSlide As Long
    SkippedNoShape As Long
    SkippedNotText As Long
    Detail As String
End Type

Public Type PropagateOutcome
    Ran As Boolean
    Problem As String
    Stamped As Long
    AlreadyTagged As Long
    NoConfidentMatch As Long
    Collided As Long
    Detail As String
    ' SEPARATE FROM Detail, and that separation is the fix. Both used to be
    ' appended to Detail, so a caller wanting "what was refused" got "what was
    ' refused AND everything that succeeded" -- which is how a run reporting 16
    ' correct stamps printed them all underneath a "Refused" header on
    ' 2026-08-14. A caller cannot un-mix two things that arrived in one string.
    Collisions As String
End Type

' CARRYING THE TEMPLATE'S ROLE TAGS ONTO A SLIDE THAT IS ALREADY LINKED.
'
' The missing middle link. Tagging the template is what a person does; every
' other slide of that type then needs the same shape identified, and until
' 2026-08-14 the only code that did that was adoption -- which refuses a linked
' slide before it matches anything. So a field tagged on the template after
' first setup could never reach the other slides through the tool at all.
'
' Reuses Onboarding.MatchSlideAgainstTemplate unchanged, with one thing done
' by the caller because the matcher does not do it: MatchSlideAgainstTemplate
' loops EVERY template role, while filtering the target slide to UNTAGGED
' shapes. Ask it about a role this slide already carries and it will happily
' score that role against some other shape -- the right answer is already
' tagged and invisible to it. So the roles are filtered to the missing ones
' FIRST, and the parallel templateFieldShapes array is filtered with them.
'
' THREE REFUSALS, each for a state that would otherwise corrupt a slide:
'
' 1. A role already on this slide is left alone -- counted, never re-stamped.
'    CountShapesWithRoleTag is used rather than FindShapeByRoleTag because the
'    latter reports "none" and "two" identically, and those need opposite
'    treatment: stamp the first, never touch the second.
' 2. Only "high" confidence stamps. Medium is what adoption sends to a human,
'    and a tag written on a guess is worse than no tag -- it is silent, and the
'    value it later publishes goes onto the wrong shape.
' 3. TWO ROLES MATCHING THE SAME SHAPE IS A COLLISION, NOT A RACE. Matching.Match
'    is called per role independently and nothing stops two roles picking one
'    shape; a shape carries ONE role tag, so the second write would overwrite the
'    first and the loser would vanish without a message. Both are refused and
'    reported by name.
Public Function PropagateTemplateTags(sld As Object, templateSld As Object, _
                                      dryRun As Boolean) As PropagateOutcome
    Dim outcome As PropagateOutcome
    outcome.Ran = False

    If sld Is Nothing Or templateSld Is Nothing Then
        outcome.Problem = "No slide or no template slide."
        PropagateTemplateTags = outcome
        Exit Function
    End If

    Dim templateRoles() As String
    Dim templateFieldShapes() As Candidate
    templateFieldShapes = Onboarding.BuildTemplateFieldShapes(templateSld, templateRoles)

    Dim rLo As Long, rHi As Long, hasRoles As Boolean
    On Error Resume Next
    rLo = LBound(templateRoles): rHi = UBound(templateRoles)
    hasRoles = (Err.Number = 0)
    On Error GoTo 0
    If Not hasRoles Then
        outcome.Problem = "The template slide carries no role tags."
        PropagateTemplateTags = outcome
        Exit Function
    End If

    ' Filter to the roles THIS slide is missing, keeping the two arrays parallel.
    Dim wantRoles() As String
    Dim wantShapes() As Candidate
    Dim n As Long
    n = 0

    Dim i As Long
    For i = rLo To rHi
        If InjectPrimitive.CountShapesWithRoleTag(sld, templateRoles(i)) > 0 Then
            outcome.AlreadyTagged = outcome.AlreadyTagged + 1
        Else
            n = n + 1
            ReDim Preserve wantRoles(1 To n)
            ReDim Preserve wantShapes(1 To n)
            wantRoles(n) = templateRoles(i)
            wantShapes(n) = templateFieldShapes(i)
        End If
    Next i

    If n = 0 Then
        outcome.Ran = True
        PropagateTemplateTags = outcome
        Exit Function
    End If

    Dim untaggedShapes() As Object
    Dim matches() As FieldMatch
    matches = Onboarding.MatchSlideAgainstTemplate(sld, wantRoles, wantShapes, untaggedShapes)

    Dim mLo As Long, mHi As Long, hasMatches As Boolean
    On Error Resume Next
    mLo = LBound(matches): mHi = UBound(matches)
    hasMatches = (Err.Number = 0)
    On Error GoTo 0
    If Not hasMatches Then
        outcome.Ran = True
        outcome.NoConfidentMatch = n
        PropagateTemplateTags = outcome
        Exit Function
    End If

    ' PASS A -- which shape each role wants, and who wants the same one.
    Dim claims As Object
    Set claims = CreateObject("Scripting.Dictionary")
    Dim j As Long
    For j = mLo To mHi
        If matches(j).Result.Confidence = "high" And matches(j).Result.HasCandidate Then
            Dim k As String
            k = CStr(matches(j).Result.CandidateIndex)
            If claims.Exists(k) Then
                claims(k) = claims(k) & "|" & matches(j).Role
            Else
                claims(k) = matches(j).Role
            End If
        End If
    Next j

    ' PASS B -- stamp the uncontested ones.
    For j = mLo To mHi
        If Not (matches(j).Result.Confidence = "high" And matches(j).Result.HasCandidate) Then
            outcome.NoConfidentMatch = outcome.NoConfidentMatch + 1
        Else
            Dim key As String
            key = CStr(matches(j).Result.CandidateIndex)
            If InStr(claims(key), "|") > 0 Then
                outcome.Collided = outcome.Collided + 1
                outcome.Collisions = outcome.Collisions & "  " & Replace(claims(key), "|", " and ") & _
                                     " both matched one shape -- neither stamped" & vbCrLf
            Else
                Dim target As Object
                Set target = untaggedShapes(matches(j).Result.CandidateIndex)
                outcome.Stamped = outcome.Stamped + 1
                outcome.Detail = outcome.Detail & "  " & matches(j).Role & " -> shape '" & _
                                 target.Name & "' = '" & _
                                 BatchOnboardFlow.FieldPreview(ShapeText(target)) & "'" & vbCrLf
                ' Tags.Add on an existing name REPLACES rather than duplicating
                ' (TagMigration.bas:12), and this shape has no role tag anyway --
                ' it came out of the untagged list.
                If Not dryRun Then target.Tags.Add "role", matches(j).Role
            End If
        End If
    Next j

    outcome.Ran = True
    PropagateTemplateTags = outcome
End Function

' `dryRun` reports exactly what a real run would write and touches nothing.
' It is the first argument a person should use on a deck of 43 slides.
Public Function HarvestSlide(sld As Object, ws As Object, period As String, _
                             dryRun As Boolean) As HarvestOutcome
    Dim outcome As HarvestOutcome
    outcome.Ran = False

    If sld Is Nothing Or ws Is Nothing Then
        outcome.Problem = "No slide or no register sheet."
        HarvestSlide = outcome
        Exit Function
    End If

    If Trim$(period) = "" Then
        outcome.Problem = "The deck has no period set, so there is no row to write into."
        HarvestSlide = outcome
        Exit Function
    End If

    Dim inst As SlideInstance
    inst = Resolve.ResolveSlideInstance(sld)
    If Not inst.HasInstanceKey Then
        outcome.Problem = "Slide " & sld.SlideIndex & " has no instance key -- it is not linked to a register row."
        HarvestSlide = outcome
        Exit Function
    End If
    outcome.InstanceKey = inst.InstanceKey

    Dim problem As String
    Dim sh As Sheet
    sh = ExcelOutput.ReadSheetForDeckPeriod(ws, period, problem)
    If problem <> "" Then
        outcome.Problem = problem
        HarvestSlide = outcome
        Exit Function
    End If

    ' A DUPLICATE ROW IS A REFUSAL, NOT A WARNING. Two rows for this instance in
    ' this period means the reader picked one and the writer would pick the
    ' other's neighbour; harvesting into that is how a value lands somewhere
    ' nothing reads. ReadSheetForPeriod counts them precisely so a caller can
    ' refuse, and this one does.
    If sh.DuplicateInstances > 0 Then
        outcome.Problem = sh.DuplicateInstances & " duplicate instance row(s) in " & period & _
                          " -- fix those before harvesting, or a value will be written into the row nothing reads."
        HarvestSlide = outcome
        Exit Function
    End If

    Dim rowValues As Object
    If sh.Rows.Exists(inst.InstanceKey) Then
        Set rowValues = sh.Rows(inst.InstanceKey)
    Else
        Set rowValues = CreateObject("Scripting.Dictionary")
    End If

    Dim toWrite As Object
    Set toWrite = CreateObject("Scripting.Dictionary")

    Dim fieldName As Variant
    For Each fieldName In sh.Fields
        Dim fname As String
        fname = Trim$(CStr(fieldName))
        If fname <> "" Then
            If rowValues.Exists(fname) Then
                outcome.SkippedHasValue = outcome.SkippedHasValue + 1
            Else
                Dim shp As Object
                Set shp = InjectPrimitive.FindShapeByRoleTag(sld, fname)
                If shp Is Nothing Then
                    ' Nothing means none OR more than one -- both are "do not
                    ' guess", and both are ordinary on a slide whose register
                    ' has more columns than the slide has tagged shapes.
                    outcome.SkippedNoShape = outcome.SkippedNoShape + 1
                ElseIf ShapeIsNotHarvestableText(sld, shp, fname) <> "" Then
                    ' HARVEST ONLY WHAT PUBLISH WOULD WRITE AS TEXT.
                    '
                    ' The slide shows a formatted VIEW of some fields, not the
                    ' value the register holds. `PROJECT_PROGRESS` reads "33%" on
                    ' the slide and InjectProgressVia refuses anything
                    ' non-numeric, naming '90%' as wrong in those exact words --
                    ' so harvesting the displayed string writes a value the tool
                    ' itself cannot publish. A device has no text of its own, and
                    ' a picture's register cell holds a source id, not a caption.
                    '
                    ' WHY THIS IS A REFUSAL AND NOT A CONVERSION: "33%" -> 0.33 is
                    ' guessable, but 'as at Q3' or '33% (est.)' is not, and a
                    ' wrong guess here is PERMANENT -- the empty-cell rule means a
                    ' corrected harvest can never overwrite it. Refusing names the
                    ' field and leaves the cell empty, which is recoverable.
                    outcome.SkippedNotText = outcome.SkippedNotText + 1
                    outcome.Detail = outcome.Detail & "  not read (" & _
                        ShapeIsNotHarvestableText(sld, shp, fname) & "): " & fname & vbCrLf
                Else
                    Dim v As String
                    v = ShapeText(shp)
                    If Trim$(v) = "" Then
                        outcome.SkippedBlankOnSlide = outcome.SkippedBlankOnSlide + 1
                    Else
                        toWrite(fname) = v
                        outcome.Written = outcome.Written + 1
                        outcome.Detail = outcome.Detail & "  " & fname & " = '" & _
                                         BatchOnboardFlow.FieldPreview(v) & "'" & vbCrLf
                    End If
                End If
            End If
        End If
    Next fieldName

    ' asText:=True -- the harvest's contract is "what is on the slide", and
    ' `.Value = CStr(...)` is not a string write: Excel parses "30 Oct 2023" into
    ' a serial and "$275,598" into 275598. Proven on the real deck 2026-08-14.
    If outcome.Written > 0 And Not dryRun Then
        ExcelOutput.UpsertRow ws, inst.InstanceKey, toWrite, period, True
    End If

    outcome.Ran = True
    HarvestSlide = outcome
End Function

' Returns "" if this shape's text IS the register's value, or a short reason if
' it is not. TWO INDEPENDENT FACTS, deliberately not collapsed into one:
'
' 1. WHAT PUBLISH WOULD DO WITH IT -- asked via InjectPrimitive.InjectorFor, the
'    same decision the router makes, so harvest and publish can never disagree
'    about a field. A bar's slide text is `33%` and its register value is a
'    fraction; a picture's cell holds a source id, not a caption.
'
' 2. WHETHER THE SHAPE HAS TEXT OF ITS OWN AT ALL. A group container does not,
'    even when every member inside it does (confirmed live 2026-07-26,
'    BatchOnboardFlow.MarkShapeForBatch's header). This is NOT implied by (1):
'    the router calls a group a DEVICE only when it carries milestone slots, so
'    a tagged group without them routes to the text writer -- and reading it
'    would report "blank on the slide", which is false. It has no text to be
'    blank. Found by Harvest_SkipsATaggedGroupInsteadOfReadingItAsEmpty going red
'    on 2026-08-15 when the injector check replaced the older group check.
Private Function ShapeIsNotHarvestableText(sld As Object, shp As Object, fname As String) As String
    If shp.Type = msoGroup Then
        ShapeIsNotHarvestableText = "a group has no text of its own"
        Exit Function
    End If

    Dim inj As String
    inj = InjectPrimitive.InjectorFor(sld, fname)
    If inj <> InjectPrimitive.INJECTOR_TEXT Then
        ShapeIsNotHarvestableText = inj & ", not text"
        Exit Function
    End If

    ShapeIsNotHarvestableText = ""
End Function

' A shape with no text frame raises rather than returning "", and a harvest that
' dies on one odd shape has read nothing from the other forty-two.
Private Function ShapeText(shp As Object) As String
    Dim result As String
    result = ""
    On Error Resume Next
    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then result = shp.TextFrame.TextRange.Text
    End If
    On Error GoTo 0
    ShapeText = result
End Function

' The one-line summary. A function of the counts alone so it cannot drift away
' from the detail beneath it -- same rule as Readiness.Headline.
Public Function HarvestSummary(o As HarvestOutcome, dryRun As Boolean) As String
    If Not o.Ran Then
        HarvestSummary = "Nothing harvested: " & o.Problem
        Exit Function
    End If

    Dim verb As String
    If dryRun Then verb = "would be written" Else verb = "written"

    HarvestSummary = o.InstanceKey & ": " & o.Written & " value(s) " & verb & "." & vbCrLf & _
        "  already had a value (left alone): " & o.SkippedHasValue & vbCrLf & _
        "  blank on the slide:               " & o.SkippedBlankOnSlide & vbCrLf & _
        "  no single shape carries the tag:  " & o.SkippedNoShape & vbCrLf & _
        "  on the slide but not text (bar/device/picture): " & o.SkippedNotText
End Function
