Attribute VB_Name = "E2EField"
Option Explicit

' ONE FIELD, taken end to end on a copy of the real 46-slide deck.
'
' Generalised from E2EAboutBody on 2026-07-31 once TWO fields had been through
' it -- a parameter added after the second case, not an abstraction guessed at
' before the first. The field name was hardcoded while only ABOUT_BODY existed,
' which was correct then and became misleading the moment KEY_EVENTS_BODY needed
' the same treatment.
'
' Deliberately modelled line for line on E2EFirstField, which is the only path
' that has ever actually delivered a field. Same order, same verification style,
' same save. Where this differs from that one, it is because ABOUT_BODY differs
' from PROJECT_STATUS -- not because a better structure suggested itself.
'
' WHY THIS RATHER THAN FINISHING ReviewQueue. A delivery check on 2026-07-31
' found the field count had sat at 1 all day while ~1,700 lines of R13 machinery
' accumulated -- generalising the review gate before a second field existed to
' prove it against. R13's requirement is real and its logic is written and
' tested; what it does not need yet is to be the thing standing between here and
' field 2. The gate here is the narrowest thing that satisfies R13: every
' change is listed with before and after, and nothing is written unless the
' caller passes apply explicitly.
'
' ABOUT_BODY is prose and entity-static (Quarter = ALL), so under the flow rules
' every one of its changes is an INDIVIDUAL decision -- there is no batching to
' exercise here and none is attempted.
'
' A COUNT OF ZERO IS NEVER SELF-EXPLANATORY, which cost a wrong conclusion
' earlier the same day: a queue of 0 was read as "nothing differs" when it could
' equally have meant "no tagged shape was found and every field was silently
' skipped". Those produce identical output from the planner. So this reports
' found / not-found / would-change separately, always.

' Application.Run UDT warm-up probe -- see E2EFirstField's Ping ladder. In a
' freshly Imported project a Public Function only resolves once the cross-module
' Public UDTs it declares have been touched by an earlier Application.Run.

Private Function FindByRole(shapesColl As Object, role As String) As Object
    Dim shp As Object
    For Each shp In shapesColl
        If shp.Type = msoGroup Then
            Dim inner As Object
            Set inner = FindByRole(shp.GroupItems, role)
            If Not inner Is Nothing Then
                Set FindByRole = inner
                Exit Function
            End If
        ElseIf shp.Tags("role") = role Then
            Set FindByRole = shp
            Exit Function
        End If
    Next shp
End Function

' mode: "migrate" renames role tags to FieldIDs and saves. Touches no slide
'                 TEXT -- it renames the labels the tool matches on. A one-off
'                 structural step, not part of a sync.
'       "dryrun"  lists what would change and writes nothing.
'       "apply"   writes, then verifies by re-reading the deck, then saves.
'
' MIGRATE IS SEPARATE BECAUSE A DRY RUN CANNOT PREVIEW THROUGH IT. With the old
' tags still in place every field lookup misses, so a dry run reports 46 shapes
' NOT FOUND and zero changes -- an honest report of a deck that is not ready,
' but useless as a preview of what apply would do. Measured on this rig
' 2026-07-31: renamed 230, already 0.
Public Function RunField(deckPath As String, registerPath As String, _
                         period As String, mode As String, fieldId As String) As String
    Dim r As String
    Dim doWrite As Boolean, doMigrate As Boolean
    doWrite = (LCase(Trim(mode)) = "apply")
    doMigrate = (LCase(Trim(mode)) = "migrate")

    Dim pres As Object
    Set pres = Application.Presentations.Open(deckPath, msoFalse, msoFalse, msoTrue)
    pres.Windows(1).Activate

    ' THE DECK'S OWN PERIOD WINS over whatever was supplied. A supplied period
    ' is a habit or a script default; the deck's property was written when
    ' somebody rolled it forward deliberately. Getting this backwards is how a
    ' deck copied to start next quarter reports last quarter's content, with no
    ' error anywhere.
    Dim mismatch As String
    mismatch = DeckRegistry.PeriodMismatchText(DeckRegistry.GetDeckPeriod(pres), period)
    If DeckRegistry.GetDeckPeriod(pres) <> "" Then period = DeckRegistry.GetDeckPeriod(pres)

    r = "Deck:     " & deckPath & vbCrLf & _
        "Register: " & registerPath & vbCrLf & _
        "Period:   " & period & IIf(DeckRegistry.GetDeckPeriod(pres) <> "", "  (from the deck)", "  (SUPPLIED -- this deck declares no period)") & vbCrLf & _
        "Mode:     " & IIf(doMigrate, "MIGRATE TAGS (renames labels, saves, no text change)", IIf(doWrite, "APPLY (will write and save)", "DRY RUN (writes nothing)")) & vbCrLf & _
        "Slides:   " & pres.Slides.count & vbCrLf & vbCrLf & _
        IIf(mismatch <> "", mismatch & vbCrLf & vbCrLf, "")

    ' --- Tag migration, exactly as the delivering run did it ---------------
    ' Idempotent: already-correct tags count as AlreadyDone, not as work. This
    ' step is why the earlier R13 run found nothing -- it was skipped, so every
    ' field lookup missed and every slide reported no_change.
    Dim fromV(1 To 5) As String
    Dim toV(1 To 5) As String
    fromV(1) = "Project Status": toV(1) = "PROJECT_STATUS"
    fromV(2) = "Project Name":   toV(2) = "PROJECT_NAME"
    fromV(3) = "Project number": toV(3) = "PROJECT_CODE"
    ' LITERAL, never the fieldId parameter. This map is the deck's fixed
    ' old-name -> FieldID translation and has nothing to do with which field
    ' this run is processing. It briefly read `toV(4) = fieldId` -- collateral
    ' from a blanket rename when this module was generalised -- which would have
    ' renamed every "About text" tag to whatever field was passed, corrupting
    ' the deck's tagging on any migrate run that was not for ABOUT_BODY.
    fromV(4) = "About text":     toV(4) = "ABOUT_BODY"
    fromV(5) = "events text":    toV(5) = "KEY_EVENTS_BODY"

    Dim mig As MigrationReport
    mig = TagMigration.MigrateRoleTags(fromV, toV, Not doMigrate)
    r = r & "--- tag migration (" & IIf(doMigrate, "LIVE", "dry") & ") ---" & vbCrLf & _
        "  scanned:  " & mig.Scanned & vbCrLf & _
        "  renamed:  " & mig.Renamed & vbCrLf & _
        "  already:  " & mig.AlreadyDone & vbCrLf & _
        "  unmapped: " & mig.Unmapped & vbCrLf & vbCrLf

    If doMigrate Then
        pres.Save
        RunField = r & "Tags migrated and deck SAVED. No slide text was changed." & vbCrLf & _
            "Re-run with -Mode dryrun to preview " & fieldId & "." & vbCrLf
        Exit Function
    End If

    ' --- Register ----------------------------------------------------------
    Dim xl As Object, wb As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(registerPath, 0, True)

    Dim reg As RegisterRead
    reg = Register.ReadRegister(wb.Worksheets(1), period, "q")
    r = r & "--- register ---" & vbCrLf & _
        "  rows seen: " & reg.RowsSeen & "   accepted: " & reg.Accepted & vbCrLf & _
        "  missing columns: '" & reg.MissingColumns & "'" & vbCrLf & _
        "  " & Register.ReadDiagnostic(reg, period) & vbCrLf & vbCrLf

    If reg.MissingColumns <> "" Or reg.Accepted = 0 Then
        wb.Close False: xl.Quit
        pres.Saved = msoTrue: pres.Close
        RunField = r & "STOPPED: nothing usable read from the register."
        Exit Function
    End If

    ' --- The gate: every ABOUT_BODY change, with before and after ----------
    ' R13.1 in its narrowest honest form. Listed before anything is written,
    ' and in dry run this is the whole output.
    Dim keyToSlide As Object
    Set keyToSlide = CreateObject("Scripting.Dictionary")
    Dim sld As Object
    For Each sld In pres.Slides
        Dim inst As SlideInstance
        inst = Resolve.ResolveSlideInstance(sld)
        If inst.HasInstanceKey And Not inst.IsTemplate Then Set keyToSlide(inst.InstanceKey) = sld
    Next sld

    Dim nFound As Long, nNotFound As Long, nSame As Long, nDiffer As Long
    ' Two DIFFERENT causes, kept apart. "no register row/slide: 46" was one
    ' number covering "this entity has no slide" and "this field was filtered
    ' out of the register" -- and after the Seed/Approved split the second is
    ' the normal, correct state while the first is a broken link.
    Dim nNoSlide As Long, nNotInRegister As Long
    Dim changes As String

    Dim k As Variant
    For Each k In reg.Data.Rows.Keys
        If Not keyToSlide.Exists(CStr(k)) Then
            nNoSlide = nNoSlide + 1
        ElseIf Not reg.Data.Rows(k).Exists(fieldId) Then
            nNotInRegister = nNotInRegister + 1
        Else
            Dim want As String
            want = CStr(reg.Data.Rows(k)(fieldId))

            Dim shp As Object
            Set shp = FindByRole(keyToSlide(CStr(k)).Shapes, fieldId)
            If shp Is Nothing Then
                nNotFound = nNotFound + 1
            Else
                nFound = nFound + 1
                ' DECIDED BY InjectPrimitive ITSELF, dry, rather than by a
                ' comparison written here. This module reimplemented the
                ' comparison at first (|| conversion by hand, then a plain =),
                ' which meant the preview could disagree with the writer about
                ' what differs -- and it did, the moment F11 (trailing
                ' paragraph marks are not a difference) landed in the injector
                ' and not here. A preview that resolves its inputs differently
                ' from the real thing is worse than no preview.
                Dim probe As InjectResult
                probe = InjectPrimitive.InjectPrimitive(keyToSlide(CStr(k)), fieldId, want, True)

                Dim wantSlideForm As String
                wantSlideForm = Replace(want, "||", Chr(13))
                Dim have As String
                have = probe.CurrentValue

                If Not probe.WouldChange Then
                    nSame = nSame + 1
                Else
                    nDiffer = nDiffer + 1
                    ' CHARACTER CODES, NOT RENDERED TEXT. A difference that looks
                    ' identical on screen is the whole hazard: an em-dash, a
                    ' non-breaking hyphen or a trailing space all render
                    ' invisibly, and the transport between here and a terminal
                    ' can itself mangle non-ASCII -- so rendered text cannot
                    ' distinguish "the data is corrupt" from "my console is".
                    ' Codes can.
                    Dim dpos As Long, maxLen As Long
                    maxLen = Len(have)
                    If Len(wantSlideForm) > maxLen Then maxLen = Len(wantSlideForm)
                    dpos = 0
                    Dim ci As Long
                    For ci = 1 To maxLen
                        If Mid(have, ci, 1) <> Mid(wantSlideForm, ci, 1) Then
                            dpos = ci
                            Exit For
                        End If
                    Next ci

                    changes = changes & "  " & k & _
                        "  slideLen=" & Len(have) & " regLen=" & Len(wantSlideForm) & _
                        " firstDiffAt=" & dpos
                    ' NOT IIf. VBA's IIf evaluates BOTH branches regardless of
                    ' the condition, so `IIf(dpos <= Len(x), AscW(Mid(x,dpos,1)), -1)`
                    ' still calls AscW on an empty string when the guard is
                    ' False -- run-time error 5, which headless is a modal
                    ' dialog and a hung run. A guard that cannot prevent the
                    ' thing it guards against is worse than no guard: it reads
                    ' as care taken. Hit live 2026-07-31.
                    If dpos > 0 Then
                        Dim sc As Long, rc As Long
                        sc = -1: rc = -1
                        If dpos <= Len(have) Then sc = AscW(Mid(have, dpos, 1))
                        If dpos <= Len(wantSlideForm) Then rc = AscW(Mid(wantSlideForm, dpos, 1))
                        changes = changes & "  slideChr=" & sc & " regChr=" & rc
                    End If
                    changes = changes & vbCrLf
                End If
            End If
        End If
    Next k

    r = r & "--- " & fieldId & ", what is actually there ---" & vbCrLf & _
        "  shape found on slide:  " & nFound & vbCrLf & _
        "  SHAPE NOT FOUND:       " & nNotFound & vbCrLf & _
        "  already correct:       " & nSame & vbCrLf & _
        "  WOULD CHANGE:          " & nDiffer & vbCrLf & _
        "  entity has no slide:   " & nNoSlide & vbCrLf & _
        "  not writable (held back by Status): " & nNotInRegister & vbCrLf & vbCrLf

    If nDiffer > 0 Then
        r = r & "--- the changes, before and after ---" & vbCrLf & changes & vbCrLf
    End If

    If Not doWrite Then
        wb.Close False: xl.Quit
        pres.Saved = msoTrue: pres.Close
        RunField = r & "DRY RUN -- nothing was written to the deck." & vbCrLf
        Exit Function
    End If

    ' --- Write, one field, through the real injector -----------------------
    Dim wrote As Long, failed As Long
    For Each k In reg.Data.Rows.Keys
        If keyToSlide.Exists(CStr(k)) Then
            If reg.Data.Rows(k).Exists(fieldId) Then
                Dim res As InjectResult
                res = InjectPrimitive.InjectPrimitive(keyToSlide(CStr(k)), fieldId, _
                        CStr(reg.Data.Rows(k)(fieldId)), False)
                If res.Found And res.Written Then
                    If res.Verified Then wrote = wrote + 1 Else failed = failed + 1
                End If
            End If
        End If
    Next k

    wb.Close False
    xl.Quit

    r = r & "--- write ---" & vbCrLf & _
        "  written and verified: " & wrote & vbCrLf & _
        "  failed verification:  " & failed & vbCrLf & vbCrLf

    ' --- Verify by re-reading the DECK, not by trusting the report ---------
    Dim vMatch As Long, vMiss As Long
    For Each k In reg.Data.Rows.Keys
        If keyToSlide.Exists(CStr(k)) Then
            Dim vshp As Object
            Set vshp = FindByRole(keyToSlide(CStr(k)).Shapes, fieldId)
            ' Guarded on Exists FIRST. Scripting.Dictionary ADDS a key when you
            ' read a missing one, returning Empty -- so an entity with no
            ' approved row for THIS field would silently compare the slide
            ' against "" and be counted a mismatch. Harmless with one field in
            ' the register; wrong the moment there are three, which is now.
            If Not vshp Is Nothing And reg.Data.Rows(k).Exists(fieldId) Then
                ' THROUGH InjectPrimitive, not a hand-rolled comparison.
                '
                ' This block previously did `slideText = Replace(value,"||",CR)`
                ' and compared with `=`. The moment F11 landed (trailing
                ' paragraph marks are not a difference) the verifier and the
                ' gate were applying DIFFERENT rules, and the same run reported
                ' "45 already correct" and "19 mismatched" about one deck.
                '
                ' A verifier that reimplements the thing it verifies is not a
                ' second opinion, it is a second implementation -- and when they
                ' disagree neither number means anything. Ask the writer.
                Dim vprobe As InjectResult
                vprobe = InjectPrimitive.InjectPrimitive(keyToSlide(CStr(k)), fieldId, _
                            CStr(reg.Data.Rows(k)(fieldId)), True)
                If vprobe.Found And Not vprobe.WouldChange Then vMatch = vMatch + 1 Else vMiss = vMiss + 1
            End If
        End If
    Next k

    r = r & "--- VERIFIED BY RE-READING THE DECK ---" & vbCrLf & _
        "  slides matching the register: " & vMatch & vbCrLf & _
        "  mismatched:                   " & vMiss & vbCrLf

    pres.Save
    r = r & vbCrLf & "Deck saved." & vbCrLf

    RunField = r
End Function

' Repair named register rows FROM the slide, correctly encoded.
'
' This is the seeding operation done right, and it is the population side
' writing back to the synthesis side -- so it is deliberately NOT something any
' sync does, and it takes an explicit list of entities rather than "everything
' that differs". Seeding is not approving, and a routine that could silently
' reseed a whole register would erase the difference between the two.
'
' The case it exists for, measured 2026-07-31: 2_P004's slide holds a blank
' line between paragraphs; the register's "||" expanded to a single break, so
' the two disagreed by 2 characters at position 309. The SLIDE is right there
' and the register under-represents it -- applying would have flattened a real
' blank line out of real prose to satisfy a comparison.
'
' Every real line break becomes "||", which is the register's own convention
' (R6) and the inverse of what InjectPrimitive does on the way in.
Public Function ReseedFromSlides(deckPath As String, registerPath As String, _
                                 period As String, entityList As String, fieldId As String) As String
    Dim r As String

    Dim pres As Object
    Set pres = Application.Presentations.Open(deckPath, msoTrue, msoFalse, msoTrue)
    pres.Windows(1).Activate

    Dim wanted As Object
    Set wanted = CreateObject("Scripting.Dictionary")
    Dim parts() As String
    parts = Split(entityList, ",")
    Dim pi As Long
    For pi = LBound(parts) To UBound(parts)
        If Trim(parts(pi)) <> "" Then wanted(Trim(parts(pi))) = True
    Next pi

    r = "Reseeding " & wanted.count & " entity(ies) FROM the slides." & vbCrLf & vbCrLf

    Dim keyToSlide As Object
    Set keyToSlide = CreateObject("Scripting.Dictionary")
    Dim sld As Object
    For Each sld In pres.Slides
        Dim inst As SlideInstance
        inst = Resolve.ResolveSlideInstance(sld)
        If inst.HasInstanceKey And Not inst.IsTemplate Then Set keyToSlide(inst.InstanceKey) = sld
    Next sld

    Dim xl As Object, wb As Object, ws As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(registerPath)      ' read-WRITE: this repairs it
    Set ws = wb.Worksheets(1)

    ' Columns by header name, never by position.
    Dim cEntity As Long, cField As Long, cValue As Long
    Dim c As Long
    For c = 1 To 20
        Select Case Trim(CStr(ws.Cells(1, c).Value))
            Case "EntityCode": cEntity = c
            Case "FieldID":    cField = c
            Case "Value":      cValue = c
        End Select
    Next c

    If cEntity = 0 Or cField = 0 Or cValue = 0 Then
        wb.Close False: xl.Quit
        pres.Saved = msoTrue: pres.Close
        ReseedFromSlides = r & "STOPPED: could not locate EntityCode/FieldID/Value by header."
        Exit Function
    End If

    Dim fixed As Long, skipped As Long
    Dim rowN As Long
    rowN = 2
    Do While Trim(CStr(ws.Cells(rowN, cEntity).Value)) <> ""
        Dim ent As String
        ent = Trim(CStr(ws.Cells(rowN, cEntity).Value))
        If wanted.Exists(ent) And Trim(CStr(ws.Cells(rowN, cField).Value)) = fieldId Then
            If keyToSlide.Exists(ent) Then
                Dim shp As Object
                Set shp = FindByRole(keyToSlide(ent).Shapes, fieldId)
                If shp Is Nothing Then
                    skipped = skipped + 1
                Else
                    Dim slideText As String
                    slideText = shp.TextFrame.TextRange.Text
                    Dim encoded As String
                    encoded = Replace(slideText, Chr(13), "||")
                    encoded = Replace(encoded, Chr(11), "||")
                    ws.Cells(rowN, cValue).Value = encoded
                    fixed = fixed + 1
                    r = r & "  " & ent & ": reseeded, slideLen=" & Len(slideText) & _
                        " encodedLen=" & Len(encoded) & vbCrLf
                End If
            Else
                skipped = skipped + 1
            End If
        End If
        rowN = rowN + 1
    Loop

    wb.Save
    wb.Close False
    xl.Quit
    pres.Saved = msoTrue
    pres.Close

    r = r & vbCrLf & "reseeded: " & fixed & "   skipped: " & skipped & vbCrLf & _
        "Register updated. Deck was opened READ-ONLY and not changed." & vbCrLf
    ReseedFromSlides = r
End Function

' Does the harvest round-trip? Reads the TSV that DumpFieldValues produced and
' asks the INJECTOR whether writing each value back would change anything.
'
' This settles the open "||" question directly instead of inferring it. The
' harvest encodes every real line break as "||"; InjectPrimitive converts every
' "||" back to one break. If those two are exact inverses, replaying the harvest
' is a no-op on all 46 slides. If they are not, this says so on the field where
' it matters most -- KEY_EVENTS_BODY is multi-line on 46 of 46, median 5
' paragraphs, so a one-character asymmetry shows up everywhere.
'
' Deliberately reads the TSV and NOT the register: this is a question about
' harvest fidelity, not about approval, and routing it through Status would
' conflate the two things that were just separated.
Public Function VerifyHarvest(deckPath As String, tsvPath As String, fieldId As String) As String
    Dim r As String

    Dim pres As Object
    Set pres = Application.Presentations.Open(deckPath, msoTrue, msoFalse, msoTrue)
    pres.Windows(1).Activate

    Dim keyToSlide As Object
    Set keyToSlide = CreateObject("Scripting.Dictionary")
    Dim sld As Object
    For Each sld In pres.Slides
        Dim inst As SlideInstance
        inst = Resolve.ResolveSlideInstance(sld)
        If inst.HasInstanceKey And Not inst.IsTemplate Then Set keyToSlide(inst.InstanceKey) = sld
    Next sld

    ' ADODB.Stream, not FileSystemObject. FSO's OpenTextFile only offers ASCII,
    ' system-default or UTF-16 -- there is no UTF-8 option. The TSV is UTF-8
    ' with a BOM (written by PowerShell), so reading it as UTF-16 returned zero
    ' usable rows: every line came back as mojibake and no FieldID matched.
    '
    ' It reported "rows in TSV: 0", which is the only reason this was obvious
    ' rather than silent -- the same count printed without that line would have
    ' read as "nothing differs, all clean".
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2          ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open
    stream.LoadFromFile tsvPath

    Dim nRows As Long, nNoSlide As Long, nNotFound As Long, nSame As Long, nDiffer As Long
    Dim detail As String

    Do While Not stream.EOS
        Dim line As String
        line = stream.ReadText(-2)      ' adReadLine
        Dim parts() As String
        parts = Split(line, vbTab)
        If UBound(parts) >= 2 Then
            If Trim(parts(1)) = fieldId Then
                nRows = nRows + 1
                Dim ent As String
                ent = Trim(parts(0))
                If Not keyToSlide.Exists(ent) Then
                    nNoSlide = nNoSlide + 1
                Else
                    Dim probe As InjectResult
                    probe = InjectPrimitive.InjectPrimitive(keyToSlide(ent), fieldId, parts(2), True)
                    If Not probe.Found Then
                        nNotFound = nNotFound + 1
                    ElseIf probe.WouldChange Then
                        nDiffer = nDiffer + 1
                        Dim breaksSlide As Long, breaksValue As Long
                        breaksSlide = Len(probe.CurrentValue) - Len(Replace(probe.CurrentValue, Chr(13), ""))
                        breaksValue = (Len(parts(2)) - Len(Replace(parts(2), "||", ""))) / 2
                        detail = detail & "  " & ent & "  slideLen=" & Len(probe.CurrentValue) & _
                            " breaksOnSlide=" & breaksSlide & " ||inValue=" & breaksValue & vbCrLf
                    Else
                        nSame = nSame + 1
                    End If
                End If
            End If
        End If
    Loop
    stream.Close

    r = "--- harvest round-trip: " & fieldId & " ---" & vbCrLf & _
        "  rows in TSV:        " & nRows & vbCrLf & _
        "  no slide:           " & nNoSlide & vbCrLf & _
        "  shape not found:    " & nNotFound & vbCrLf & _
        "  ROUND-TRIPS EXACTLY:" & nSame & vbCrLf & _
        "  WOULD CHANGE:       " & nDiffer & vbCrLf
    If nDiffer > 0 Then r = r & vbCrLf & detail

    pres.Saved = msoTrue
    pres.Close
    VerifyHarvest = r
End Function

' Delete named entities: their slides AND their register rows, in one pass.
'
' Both halves together on purpose. Deleting the slides alone would leave orphan
' register rows, which the planner classifies as new_record -- and a large
' new_record count is the exact signal that means "this deck's linkage has
' drifted", so it would look like a fault rather than a tidy-up. Deleting the
' rows alone would leave slides nothing can address.
'
' RM ruling: 3_P002-2, 2_P004-2, 1_P006-2 are duplicates and go. They are also
' the only instances of the overloaded key syntax -- "project number plus a
' disambiguator" sharing one namespace with real project numbers -- so removing
' them narrows the identity problem while the GUID redesign is still pending.
'
' REFUSES rather than does less than asked: an entity named here but not found
' stops the whole operation. Silently deleting two of three requested slides and
' reporting success is how a deck ends up in a state nobody predicted.
Public Function DeleteEntities(deckPath As String, registerPath As String, entityList As String) As String
    Dim r As String

    Dim wanted As Object
    Set wanted = CreateObject("Scripting.Dictionary")
    Dim parts() As String
    parts = Split(entityList, ",")
    Dim pi As Long
    For pi = LBound(parts) To UBound(parts)
        If Trim(parts(pi)) <> "" Then wanted(Trim(parts(pi))) = True
    Next pi

    Dim pres As Object
    Set pres = Application.Presentations.Open(deckPath, msoFalse, msoFalse, msoTrue)
    pres.Windows(1).Activate

    r = "Deleting " & wanted.count & " entity(ies)." & vbCrLf & _
        "Slides before: " & pres.Slides.count & vbCrLf & vbCrLf

    ' Locate first, delete second. Deleting while enumerating renumbers the
    ' collection underneath the loop.
    Dim targets As Collection
    Set targets = New Collection
    Dim found As Object
    Set found = CreateObject("Scripting.Dictionary")

    Dim sld As Object
    For Each sld In pres.Slides
        Dim inst As SlideInstance
        inst = Resolve.ResolveSlideInstance(sld)
        If inst.HasInstanceKey Then
            If wanted.Exists(inst.InstanceKey) And Not inst.IsTemplate Then
                targets.Add sld
                found(inst.InstanceKey) = True
            End If
        End If
    Next sld

    Dim missing As String
    Dim k As Variant
    For Each k In wanted.Keys
        If Not found.Exists(k) Then missing = missing & " " & k
    Next k

    If missing <> "" Then
        pres.Saved = msoTrue
        pres.Close
        DeleteEntities = r & "REFUSED: no slide found for:" & missing & vbCrLf & _
            "Nothing was deleted. Check the keys before re-running."
        Exit Function
    End If

    Dim i As Long
    For i = targets.count To 1 Step -1
        r = r & "  deleted slide for " & Resolve.ResolveSlideInstance(targets(i)).InstanceKey & vbCrLf
        targets(i).Delete
    Next i

    r = r & "Slides after: " & pres.Slides.count & vbCrLf & vbCrLf

    ' Register rows for those entities, every field.
    Dim xl As Object, wb As Object, ws As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(registerPath)
    Set ws = wb.Worksheets(1)

    Dim cEntity As Long, c As Long
    For c = 1 To 20
        If Trim(CStr(ws.Cells(1, c).Value)) = "EntityCode" Then cEntity = c
    Next c

    Dim removed As Long
    If cEntity = 0 Then
        r = r & "WARNING: could not locate EntityCode column -- register NOT changed." & vbCrLf
    Else
        ' Bottom-up, for the same reason as the slides.
        Dim lastRow As Long
        lastRow = 2
        Do While Trim(CStr(ws.Cells(lastRow, cEntity).Value)) <> ""
            lastRow = lastRow + 1
        Loop
        Dim rowN As Long
        For rowN = lastRow - 1 To 2 Step -1
            If wanted.Exists(Trim(CStr(ws.Cells(rowN, cEntity).Value))) Then
                ws.Rows(rowN).Delete
                removed = removed + 1
            End If
        Next rowN
        wb.Save
    End If
    wb.Close False
    xl.Quit

    r = r & "Register rows removed: " & removed & vbCrLf

    pres.Save
    r = r & "Deck saved." & vbCrLf
    DeleteEntities = r
End Function

' Drafting entry points. Workbook-only -- no deck is opened, because drafting is
' work you should be able to do on a laptop with no deck in front of you.
Public Function BuildDraftSheet(registerPath As String, period As String, fieldId As String) As String
    Dim xl As Object, wb As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(registerPath)

    Dim reg As RegisterRead
    reg = Register.ReadRegisterAllStatuses(wb.Worksheets(1), period, "q")

    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(wb, Drafting.DraftSheetNameFor(fieldId))

    Dim r As String
    r = Drafting.WriteDraftingSheet(ws, reg.Data, fieldId) & vbCrLf & vbCrLf & _
        "--- prompt to paste above the sheet ---" & vbCrLf & _
        Drafting.DraftingPromptFor(fieldId) & vbCrLf

    wb.Save
    wb.Close False
    xl.Quit
    BuildDraftSheet = r
End Function

Public Function PublishDraftSheet(registerPath As String, fieldId As String, mode As String) As String
    Dim xl As Object, wb As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(registerPath)

    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(wb, Drafting.DraftSheetNameFor(fieldId))

    Dim r As String
    r = Drafting.PublishDrafts(ws, wb.Worksheets(1), fieldId, (LCase(Trim(mode)) <> "apply"))

    wb.Save
    wb.Close False
    xl.Quit
    PublishDraftSheet = r
End Function

' Roll the deck forward. Explicit, with the from-and-to stated.
Public Function SetPeriod(deckPath As String, newPeriod As String) As String
    Dim pres As Object
    Set pres = Application.Presentations.Open(deckPath, msoFalse, msoFalse, msoTrue)
    pres.Windows(1).Activate

    Dim old As String
    old = DeckRegistry.GetDeckPeriod(pres)

    Dim r As String
    r = DeckRegistry.AdvancePeriodText(old, newPeriod) & vbCrLf & vbCrLf

    ' Where the deck was opened from and whether it is writable. Kept rather
    ' than removed after the debugging that added it: a save reporting success
    ' on a READ-ONLY presentation is the failure mode that cost most of this,
    ' and "ReadOnly: 0" in the output is how the next person rules it out in a
    ' second instead of an hour.
    r = r & "  opened:   " & pres.fullName & vbCrLf & _
            "  ReadOnly: " & pres.ReadOnly & vbCrLf

    DeckRegistry.SetDeckPeriod pres, newPeriod
    On Error Resume Next
    pres.Save
    If Err.Number <> 0 Then r = r & "  *** Save RAISED: " & Err.Number & " " & Err.Description & vbCrLf
    On Error GoTo 0
    pres.Close

    ' VERIFY BY REOPENING. The first version read the value back in the SAME
    ' session and reported "Deck period is now: X" -- which proves the write
    ' happened and says nothing whatever about the save. It reported success
    ' while the file on disk still held the previous period, and the next run
    ' silently used the old quarter. Exactly the shape of every other failure
    ' this session: something looked like it worked and had not.
    Dim check As Object
    Set check = Application.Presentations.Open(deckPath, msoTrue, msoFalse, msoTrue)
    Dim onDisk As String
    onDisk = DeckRegistry.GetDeckPeriod(check)
    check.Close

    If StrComp(onDisk, newPeriod, vbTextCompare) = 0 Then
        r = r & "Deck period is now: " & onDisk & "  (verified by reopening the file)" & vbCrLf
    Else
        r = r & "*** DID NOT PERSIST ***" & vbCrLf & _
            "    asked for: " & newPeriod & vbCrLf & _
            "    on disk:   " & IIf(onDisk = "", "<nothing>", onDisk) & vbCrLf & _
            "The write and the save both reported success. Do not trust the next" & vbCrLf & _
            "run's period until this is resolved." & vbCrLf
    End If

    SetPeriod = r
End Function

' PROBE, not the shipped path. Tests candidate fixes for the open
' CustomDocumentProperties persistence bug (AGENTS.md, 2026-07-31) against a
' SINGLE write, one process lifetime. The caller (field_e2e.ps1) is what
' supplies "Office fully closed between each" -- this function does not loop
' or retry across process boundaries, because the whole point is that an
' in-process check cannot be trusted. The offline docProps/custom.xml read is
' the only thing that judges pass/fail.
'
' variant:
'   "save"       -- baseline, identical to SetPeriod's own write (Save).
'   "saveas"     -- SaveAs to the SAME path instead of Save. Untried step #1
'                   from AGENTS.md. Hypothesis: Save on a large deck (this rig
'                   is ~49MB) can be an incremental/partial rewrite that
'                   doesn't always regenerate docProps/custom.xml; SaveAs
'                   forces a full package rewrite.
'   "hiddenopen" -- open with WithWindow:=msoFalse (untried step #2), then
'                   plain Save.
'   "doublesave" -- call Save twice in a row before Close.
Public Function SetPeriodVariant(deckPath As String, newPeriod As String, saveVariant As String) As String
    Dim withWindow As Long
    withWindow = IIf(LCase(saveVariant) = "hiddenopen", msoFalse, msoTrue)

    Dim pres As Object
    Set pres = Application.Presentations.Open(deckPath, msoFalse, msoFalse, withWindow)
    On Error Resume Next
    pres.Windows(1).Activate   ' no Windows collection item when WithWindow:=False; ignore
    On Error GoTo 0

    Dim old As String
    old = DeckRegistry.GetDeckPeriod(pres)

    Dim r As String
    r = "variant=" & saveVariant & "  from='" & old & "' to='" & newPeriod & "'" & vbCrLf & _
        "  opened:   " & pres.fullName & vbCrLf & _
        "  ReadOnly: " & pres.ReadOnly & vbCrLf

    DeckRegistry.SetDeckPeriod pres, newPeriod

    On Error Resume Next
    Select Case LCase(saveVariant)
        Case "saveas"
            pres.SaveAs pres.fullName, ppSaveAsDefault
        Case "doublesave"
            pres.Save
            pres.Save
        Case Else   ' "save", "hiddenopen"
            pres.Save
    End Select
    If Err.Number <> 0 Then r = r & "  *** write RAISED: " & Err.Number & " " & Err.Description & vbCrLf
    On Error GoTo 0

    r = r & "  pres.Saved after write: " & pres.Saved & vbCrLf

    pres.Close

    SetPeriodVariant = r
End Function
