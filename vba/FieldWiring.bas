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
Public Type FieldWiringResult
    Unmarked As String          ' register fields NO existing slide carries
    UnmarkedCount As Long
    TemplateUnmarked As String  ' register fields the TEMPLATE does not carry
    TemplateUnmarkedCount As Long
    OrphanTracks As String      ' `X.track` present with no `X`, either place
    OrphanCount As Long
    Coverage As String          ' "FIELD on X of Y", only when X < Y
    PartialCount As Long
    Wired As Long               ' register fields that resolve on some slide
    SlidesScanned As Long
    Scanned As Boolean          ' False = could not look; never report a pass
    TemplateScanned As Boolean  ' False = no template found to look at
End Type

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
Private Sub WalkForRoles(shapesColl As Object, ByRef seen As Object)
    Dim shp As Object
    For Each shp In shapesColl
        If shp.Type = msoGroup Then
            WalkForRoles shp.GroupItems, seen
        Else
            Dim role As String
            role = ""
            On Error Resume Next
            role = shp.Tags("role")
            On Error GoTo 0
            If role <> "" Then seen(UCase(Trim(role))) = True
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
        Set rs = CreateObject("Scripting.Dictionary")
        WalkForRoles instances(i).Shapes, rs

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
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    If sld Is Nothing Then
        Set RoleTagsOnSlide = Nothing
        Exit Function
    End If
    WalkForRoles sld.Shapes, seen
    Set RoleTagsOnSlide = seen
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
                Dim carriers As Long
                carriers = 0

                Dim k As Variant
                For Each k In rolesByKey.Keys
                    If rolesByKey(k).Exists(fieldName) Then carriers = carriers + 1
                Next k

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
                End If

                If result.TemplateScanned Then
                    If Not tmpl.Exists(fieldName) Then
                        result.TemplateUnmarkedCount = result.TemplateUnmarkedCount + 1
                        If result.TemplateUnmarked <> "" Then _
                            result.TemplateUnmarked = result.TemplateUnmarked & ", "
                        result.TemplateUnmarked = result.TemplateUnmarked & CStr(f)
                    End If
                End If
            End If
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
        If Len(tag) > 6 Then
            If Right(tag, 6) = ".TRACK" Then
                Dim baseName As String
                baseName = Left(tag, Len(tag) - 6)
                If Not seen.Exists(baseName) Then
                    result.OrphanCount = result.OrphanCount + 1
                    If result.OrphanTracks <> "" Then result.OrphanTracks = result.OrphanTracks & ", "
                    result.OrphanTracks = result.OrphanTracks & prefix & tag
                End If
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

    If r.UnmarkedCount = 0 And r.OrphanCount = 0 And r.TemplateUnmarkedCount = 0 Then
        Dim ok As String
        ok = r.Wired & " field(s) tagged on " & r.SlidesScanned & " slide(s)"
        ' NOT A PROBLEM TODAY, AND WORTH KNOWING BEFORE DRAFTING. A field on 2
        ' of 43 slides strands nothing while only that one project has text for
        ' it; it strands the SECOND project's work, after that work is done.
        If r.PartialCount > 0 Then
            ok = ok & ". Not on every slide yet: " & r.Coverage
        End If
        ' The template is stated either way. A silent omission here would be the
        ' same defect the split exists to fix -- a clean-looking line that never
        ' looked at the slide the future is made from.
        If r.TemplateScanned Then
            ok = ok & ", and on the template"
        Else
            ok = ok & " -- NO TEMPLATE was checked"
        End If
        WiringText = ok
        Exit Function
    End If

    Dim s As String
    If r.UnmarkedCount > 0 Then
        s = r.UnmarkedCount & " field(s) on the register that no slide carries: " & r.Unmarked
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
