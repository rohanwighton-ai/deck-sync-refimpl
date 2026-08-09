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

Private Function FindShapeByRoleTag(sld As Object, identityTag As String) As Object
    Dim match As Object
    Dim matchCount As Long
    matchCount = 0

    WalkForRoleTag sld.Shapes, identityTag, match, matchCount

    If matchCount = 1 Then
        Set FindShapeByRoleTag = match
    Else
        Set FindShapeByRoleTag = Nothing
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
Private Sub WalkForRoleTag(shapesColl As Object, identityTag As String, ByRef match As Object, ByRef matchCount As Long)
    Dim shp As Object
    For Each shp In shapesColl
        If shp.Type = msoGroup Then
            WalkForRoleTag shp.GroupItems, identityTag, match, matchCount
        ElseIf ShapeHasRoleTag(shp, identityTag) Then
            matchCount = matchCount + 1
            Set match = shp
        End If
    Next shp
End Sub

' Shape.Tags(name) returns "" both when the tag is absent and when it is
' present with an empty string value -- a real VBA API quirk with no
' equivalent in the Python side's explicit dict-based tag storage (identity_
' tags.py's _parse_tag_list returns a real {} when a tag is missing). Since
' an empty-string role value is not a meaningful identity tag in this
' project's scheme, treating "absent" and "empty" the same way here is a
' safe simplification, not a correctness gap -- documented in SPIKE_NOTES.md.
Private Function ShapeHasRoleTag(shp As Object, identityTag As String) As Boolean
    ShapeHasRoleTag = (shp.Tags("role") = identityTag) And (identityTag <> "")
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

    ' The frame the slide already defines. Captured BEFORE anything changes.
    Dim fL As Single, fT As Single, fW As Single, fH As Single, fZ As Long
    fL = shp.Left: fT = shp.Top: fW = shp.Width: fH = shp.Height
    On Error Resume Next
    fZ = shp.ZOrderPosition
    On Error GoTo 0

    Dim newShp As Object
    On Error Resume Next
    ' -1, -1 = insert at the image's NATIVE size, so its true proportions are
    ' known before anything is scaled.
    Set newShp = sld.Shapes.AddPicture(locator, msoFalse, msoTrue, fL, fT, -1, -1)
    If Err.Number <> 0 Or newShp Is Nothing Then
        Dim e As String
        e = Err.Description
        On Error GoTo 0
        result.ErrorMessage = "could not place " & locator & " (" & e & ")"
        InjectPictureField = result
        Exit Function
    End If
    On Error GoTo 0

    ' FIT INSIDE, CENTRED, PROPORTIONS KEPT.
    '
    ' THE NATIVE SIZE IS CAPTURED BEFORE ANYTHING IS ASSIGNED, and the aspect
    ' lock is OFF while assigning. With the lock ON, setting Width makes
    ' PowerPoint recompute Height underneath you, so a second line reading
    ' .Height multiplies a value that has already changed -- which put a 40x20
    ' image into a 300x150 frame at 300x300, twice as tall as the frame it was
    ' supposed to fit inside. Caught by the test on its first run.
    Dim natW As Single, natH As Single
    natW = newShp.Width
    natH = newShp.Height

    Dim scaleFactor As Single
    If natW <= 0 Or natH <= 0 Then
        scaleFactor = 1
    ElseIf (fW / natW) < (fH / natH) Then
        scaleFactor = fW / natW
    Else
        scaleFactor = fH / natH
    End If

    Dim ratioNote As String
    If natH > 0 And fH > 0 Then
        If Abs((natW / natH) - (fW / fH)) > 0.02 Then
            ratioNote = " (image proportions differ from the frame -- fitted inside and centred, nothing cropped)"
        End If
    End If

    newShp.LockAspectRatio = msoFalse
    newShp.Width = natW * scaleFactor
    newShp.Height = natH * scaleFactor
    newShp.Left = fL + (fW - newShp.Width) / 2
    newShp.Top = fT + (fH - newShp.Height) / 2

    ' The tags travel, or the field stops being a field.
    newShp.Tags.Add "role", identityTag
    newShp.Tags.Add PICTURE_SOURCE_TAG, sourceId

    shp.Delete
    On Error Resume Next
    If fZ > 0 Then
        Do While newShp.ZOrderPosition > fZ
            newShp.ZOrder 3          ' msoSendBackward
        Loop
    End If
    On Error GoTo 0

    ' VERIFIED FROM THE NEW SHAPE, not from the fact that no error was raised.
    result.Written = True
    result.Verified = (StrComp(PictureSourceOf(newShp), sourceId, vbTextCompare) = 0)
    If Not result.Verified Then
        result.ErrorMessage = "the picture was placed but its source stamp did not stick -- a later run would replace it again"
    Else
        result.ErrorMessage = ratioNote
    End If

    InjectPictureField = result
End Function
