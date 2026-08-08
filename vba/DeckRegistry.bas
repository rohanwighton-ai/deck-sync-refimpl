Attribute VB_Name = "DeckRegistry"
Option Explicit

' The deck's own reporting period (D4). See GetDeckPeriod below for why this
' is stored on the deck rather than supplied per run.
Public Const PROP_DECK_PERIOD As String = "DeckSyncPeriod"

' Implements specs/deck-registry.md: the missing lookup a one-click ribbon
' button needs and no prior module provides -- "for this open deck, which
' workbook is paired with it, and where does each known slide type's
' template/worksheet live." Every existing engine entry point
' (RunSync.RunRoutineSync, DeckAdoption's templateSld
' param) takes these as caller-supplied parameters, correct for a developer
' at the VBE but not for a button with no caller to ask.
'
' Storage is Presentation.CustomDocumentProperties, the same mechanism
' ExcelOutput.bas already uses for Workbook.CustomDocumentProperties
' (WriteDeckReference/ReadDeckReference) -- confirmed via
' `grep -rn WriteDeckReference vba/*.bas` that those are never actually
' called by anything, so this module is also the first real caller closing
' that half of input-contract.md's deck_workbook_pairing rule.

Private Const DECK_ID_PROPERTY_NAME As String = "DeckSyncId"
Private Const WORKBOOK_PATH_PROPERTY_NAME As String = "DeckSyncWorkbookPath"
Private Const TYPE_PROPERTY_PREFIX As String = "DeckSyncType:"
' Literal, NOT Application.PathSeparator: that property exists on Excel's
' Application object, not PowerPoint's, and this module is PowerPoint-hosted --
' using it is a hard compile error ("Method or data member not found"), confirmed
' 2026-07-28 against real Office. Same cross-application trap AGENTS.md already
' records for Excel's xlToLeft/xlUp constants, hit from the opposite direction.
' This add-in is Windows-only (.ppam, PowerPoint COM), so "\" is not a portability
' compromise.
Private Const PATH_SEP As String = "\"

' ---------------------------------------------------------------------
' Pure logic -- no CustomDocumentProperties access, testable without a
' live Presentation. Mirrors ResolveFields.bas's split between interactive
' entry points and pure helpers.
' ---------------------------------------------------------------------

' `templateSlideId` is a Slide.SlideID (stable across reorder/insert/delete
' of *other* slides), never a SlideIndex (which shifts). "|" is safe as a
' separator: worksheet names cannot contain it (Excel's own naming rules
' disallow it), and a numeric SlideID cannot contain it either.
Public Function BuildTypeRegistration(templateSlideId As Long, worksheetName As String) As String
    BuildTypeRegistration = CStr(templateSlideId) & "|" & worksheetName
End Function

' Returns False (leaving both ByRef args at their zero-values) if `raw` is
' empty or malformed -- a genuinely-absent or corrupted registration is a
' routine, expected state for a caller to handle, not an error to raise.
Public Function ParseTypeRegistration(raw As String, ByRef templateSlideId As Long, ByRef worksheetName As String) As Boolean
    templateSlideId = 0
    worksheetName = ""

    If raw = "" Then
        ParseTypeRegistration = False
        Exit Function
    End If

    Dim sepPos As Long
    sepPos = InStr(raw, "|")
    If sepPos = 0 Or sepPos = Len(raw) Then
        ParseTypeRegistration = False
        Exit Function
    End If

    Dim idPart As String
    idPart = Left(raw, sepPos - 1)
    If Not IsNumeric(idPart) Then
        ParseTypeRegistration = False
        Exit Function
    End If

    templateSlideId = CLng(idPart)
    worksheetName = Mid(raw, sepPos + 1)
    ParseTypeRegistration = True
End Function

' ---------------------------------------------------------------------
' CustomDocumentProperties access
' ---------------------------------------------------------------------

Private Function ReadStringProperty(pres As Object, propertyName As String) As String
    Dim prop As Object
    On Error Resume Next
    Set prop = pres.CustomDocumentProperties(propertyName)
    On Error GoTo 0

    If prop Is Nothing Then
        ReadStringProperty = ""
    Else
        ReadStringProperty = CStr(prop.Value)
    End If
End Function

Private Sub WriteStringProperty(pres As Object, propertyName As String, value As String)
    Dim prop As Object
    On Error Resume Next
    Set prop = pres.CustomDocumentProperties(propertyName)
    On Error GoTo 0

    ' DELETE AND RE-ADD, never assign to an existing property.
    '
    ' Measured 2026-07-31 against real PowerPoint: `prop.Value = value` on an
    ' EXISTING custom document property reports success, survives a read-back in
    ' the same session, survives Save reporting success -- and the old value is
    ' still on disk when the file is reopened. Adding a property that does not
    ' yet exist persists correctly, which is why this went unnoticed: every
    ' first write worked, and only updates were lost.
    '
    ' The failure is silent in the worst way. A deck rolled forward to a new
    ' quarter kept the old one, every subsequent run read the old quarter from
    ' the deck and believed it, and the tool would have synced last quarter's
    ' content while reporting the new period back to whoever asked.
    '
    ' Affects every setting stored this way -- the workbook path and the deck id
    ' as much as the period. Any deck ever re-pointed at a different workbook
    ' was at risk of silently keeping the old path.
    If Not prop Is Nothing Then
        On Error Resume Next
        pres.CustomDocumentProperties(propertyName).Delete
        On Error GoTo 0
    End If

    pres.CustomDocumentProperties.Add Name:=propertyName, _
        LinkToContent:=False, Type:=msoPropertyTypeString, Value:=value
End Sub

' Reads DeckSyncId, generating and persisting one via Scriptlet.TypeLib's
' Guid property (the standard VBA GUID idiom -- VBA has no native GUID
' generator) the first time a deck is registered at all. Stable for the
' life of the deck once written; never regenerated on subsequent calls.
Public Function GetOrCreateDeckId(pres As Object) As String
    Dim existing As String
    existing = ReadStringProperty(pres, DECK_ID_PROPERTY_NAME)

    If existing <> "" Then
        GetOrCreateDeckId = existing
        Exit Function
    End If

    Dim newId As String
    newId = Mid(CreateObject("Scriptlet.TypeLib").Guid, 2, 36)
    WriteStringProperty pres, DECK_ID_PROPERTY_NAME, newId
    GetOrCreateDeckId = newId
End Function

' The stored path is absolute (it is whatever the machine that onboarded the deck
' saw), which makes a deck non-portable: copy deck + workbook to another machine --
' a different user profile, a work PC, a SharePoint/OneDrive-for-Business sync root
' -- and the stored path names a folder that does not exist there. Every entry point
' then dead-ends at "Could not open the paired workbook", with no way to re-point it
' short of re-onboarding.
'
' Fix applied HERE rather than at the six call sites so every caller inherits it:
' if the stored path does not resolve, look for a file of the same name beside the
' DECK. That is the normal arrangement (deck and its Data workbook live together and
' get copied together), so it recovers the overwhelmingly common case automatically
' and silently.
'
' Falls back to returning the stored path unchanged when neither resolves, so the
' existing "could not open <path>" message still names something a human recognises
' rather than a guessed path they have never seen.
Public Function GetWorkbookPath(pres As Object) As String
    Dim stored As String
    stored = ReadStringProperty(pres, WORKBOOK_PATH_PROPERTY_NAME)
    If stored = "" Then Exit Function

    If FileExists(stored) Then
        GetWorkbookPath = stored
        Exit Function
    End If

    Dim beside As String
    beside = SiblingOfDeck(pres, FileNameOnly(stored))
    If beside <> "" Then
        ' Dir() cannot test an http(s) URL, so a correctly-built cloud sibling
        ' would be constructed and then rejected as non-existent, making the
        ' whole lookup useless for OneDrive/SharePoint decks. Workbooks.Open
        ' DOES accept a URL, so hand the URL back and let the open attempt be
        ' the real test; the caller already reports a clean "could not open the
        ' paired workbook at: <path>" if it fails.
        '
        ' Corrected 2026-07-29: this comment used to say Dir() "returns "" for
        ' any URL". It does not -- it RAISES runtime error 52. Probed directly
        ' against real Office: an https:// or http:// path raises 52, while a
        ' merely non-existent local path returns "". The conclusion above is
        ' unaffected (FileExists guards the call, so a raise still lands as
        ' False), but the stated mechanism was wrong, and the same wrong belief
        ' left an UNGUARDED Dir() in WorkbookBridge.CreateWorkbook that killed a
        ' live run this day. A comment that misstates why something works is how
        ' the next person writes the unguarded version.
        If IsUrl(beside) Then
            GetWorkbookPath = beside
            Exit Function
        ElseIf FileExists(beside) Then
            GetWorkbookPath = beside
            Exit Function
        End If
    End If

    GetWorkbookPath = stored
End Function

' Re-point the deck at a workbook that has genuinely moved (renamed, or filed
' somewhere other than beside the deck). The automatic sibling lookup above covers
' the common case; this is the escape hatch for the rest, so a moved workbook is
' never a dead end requiring re-onboarding.
Public Sub RepointWorkbook(pres As Object, newPath As String)
    If Not FileExists(newPath) And Not IsUrl(newPath) Then
        Err.Raise vbObjectError + 2, "DeckRegistry.RepointWorkbook", _
            "refusing to register a workbook path that does not exist: " & newPath
    End If
    SetWorkbookPath pres, newPath
End Sub

Public Function IsUrl(path As String) As Boolean
    IsUrl = (LCase(Left(path, 5)) = "http:" Or LCase(Left(path, 6)) = "https:")
End Function

Private Function FileExists(path As String) As Boolean
    If Trim(path) = "" Then Exit Function
    On Error Resume Next
    FileExists = (Dir(path) <> "")
    On Error GoTo 0
End Function

' Bare file name from a full path. Handles both separators -- a path stored on one
' machine and read on another is not guaranteed to use the local convention.
Public Function FileNameOnly(path As String) As String
    Dim i As Long, ch As String
    FileNameOnly = path
    For i = Len(path) To 1 Step -1
        ch = Mid(path, i, 1)
        If ch = "\" Or ch = "/" Then
            FileNameOnly = Mid(path, i + 1)
            Exit Function
        End If
    Next i
End Function

' `fileName` resolved against the deck's own folder. Returns "" for a presentation
' that has never been saved (Path is empty), which is a real state during testing.
Private Function SiblingOfDeck(pres As Object, fileName As String) As String
    Dim deckDir As String
    On Error Resume Next
    deckDir = pres.path
    On Error GoTo 0
    If deckDir = "" Or fileName = "" Then Exit Function

    ' A OneDrive/SharePoint-hosted deck reports Path as a URL, not a filesystem
    ' path -- confirmed live 2026-07-28: "https://d.docs.live.net/<id>/Claude/
    ' test-sandbox". Joining that with a backslash produces a path that resolves
    ' nowhere, so the sibling lookup would fail for exactly the cloud-hosted case
    ' that matters most (decks on SharePoint at work). Match the separator to the
    ' location's own convention.
    If LCase(Left(deckDir, 5)) = "http:" Or LCase(Left(deckDir, 6)) = "https:" Then
        SiblingOfDeck = deckDir & "/" & fileName
    Else
        SiblingOfDeck = deckDir & PATH_SEP & fileName
    End If
End Function

Public Sub SetWorkbookPath(pres As Object, path As String)
    WriteStringProperty pres, WORKBOOK_PATH_PROPERTY_NAME, path
End Sub

Public Sub RegisterType(pres As Object, slideType As String, templateSld As Object, worksheetName As String)
    WriteStringProperty pres, TYPE_PROPERTY_PREFIX & slideType, _
        BuildTypeRegistration(templateSld.SlideID, worksheetName)
End Sub

' False (both ByRef args left Nothing/"") if `slideType` was never
' registered, or if its stored SlideID no longer resolves to a live slide
' (its template was deleted) -- never raises, since a ribbon button needs
' to distinguish "not onboarded yet" from a real error, not catch an
' exception to find out.
Public Function LookupType(pres As Object, slideType As String, ByRef templateSld As Object, ByRef worksheetName As String) As Boolean
    Set templateSld = Nothing
    worksheetName = ""

    Dim raw As String
    raw = ReadStringProperty(pres, TYPE_PROPERTY_PREFIX & slideType)

    Dim slideId As Long
    Dim ws As String
    If Not ParseTypeRegistration(raw, slideId, ws) Then
        LookupType = False
        Exit Function
    End If

    On Error Resume Next
    Set templateSld = pres.Slides.FindBySlideID(slideId)
    On Error GoTo 0

    If templateSld Is Nothing Then
        worksheetName = ""
        LookupType = False
        Exit Function
    End If

    worksheetName = ws
    LookupType = True
End Function

' Every registered type's name (the part after the "DeckSyncType:" prefix),
' for the New Period picker's type dropdown and any other "what types does
' this deck know about" need. Order is whatever CustomDocumentProperties'
' own iteration order is (insertion order in practice, but not a documented
' guarantee) -- callers needing a stable display order should sort.
Public Function ListRegisteredTypes(pres As Object) As String()
    Dim results() As String
    Dim n As Long
    n = 0

    Dim prop As Object
    For Each prop In pres.CustomDocumentProperties
        If Left(prop.Name, Len(TYPE_PROPERTY_PREFIX)) = TYPE_PROPERTY_PREFIX Then
            n = n + 1
            ReDim Preserve results(1 To n)
            results(n) = Mid(prop.Name, Len(TYPE_PROPERTY_PREFIX) + 1)
        End If
    Next prop

    ListRegisteredTypes = results
End Function

' Manual smoke-test entry point -- not a real test harness, same as every
' other module here. Registers a throwaway type against slide 1 of the
' active presentation and reads it back.
Public Sub ManualSmokeTest()
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim deckId As String
    deckId = GetOrCreateDeckId(pres)

    RegisterType pres, "manual-smoke-test-type", pres.Slides(1), "Sheet1"

    Dim templateSld As Object
    Dim ws As String
    Dim found As Boolean
    found = LookupType(pres, "manual-smoke-test-type", templateSld, ws)

    Dim msg As String
    msg = "DeckId=" & deckId & vbCrLf & _
        "LookupType found=" & found & " templateSld.SlideID=" & IIf(found, templateSld.SlideID, "n/a") & " ws=" & ws
    Debug.Print msg
    MsgBox msg
End Sub

' ---------------------------------------------------------------------
' The deck's own reporting period (D4)
' ---------------------------------------------------------------------

' THE DECK DECLARES ITS OWN PERIOD. Until 2026-07-31 the period was supplied as
' an argument on every run, which meant nothing in the deck knew what quarter it
' was for -- and the register filter is
' `Status = Approved AND (Quarter = <deck period> OR Quarter = ALL)`, so the
' argument silently decided which quarter's content landed.
'
' Why an argument is not good enough, and this is the specific hazard the
' WORKPLAN already recorded as undefended: starting next quarter means COPYING
' this deck. The copy carries every tag, every registered type, and -- with an
' argument-supplied period -- absolutely nothing that says which quarter it is.
' Type the old period by habit, or let a script default, and the new quarter's
' deck reports last quarter's numbers with no error anywhere. `SlideID` cannot
' help: the probe showed it is preserved on cross-deck paste.
'
' Stored, so a copy inherits the WRONG period rather than no period -- which is
' worse in principle and better in practice, because AdvancePeriod below makes
' rolling forward an explicit act, and PeriodMismatchText makes disagreeing with
' the caller loud.

Public Function GetDeckPeriod(pres As Object) As String
    GetDeckPeriod = ReadStringProperty(pres, PROP_DECK_PERIOD)
End Function

Public Sub SetDeckPeriod(pres As Object, period As String)
    WriteStringProperty pres, PROP_DECK_PERIOD, period
End Sub

' ---------------------------------------------------------------------------
' Setting the period RELIABLY, on top of a write that is not reliable.
' ---------------------------------------------------------------------------
'
' 2026-08-08, on the rig: Rohan set the period, the dialog said "Deck period is
' now Q4F26", he saved, he closed PowerPoint -- and the file's timestamp was
' still THREE DAYS OLD. Nothing reached disk. The old code verified by reading
' the property back through the same Presentation object, which reads
' PowerPoint's cache, so it confirmed its own write regardless of the file.
'
' That is the trap set_deck_period.py was written against and its rule is
' explicit: verification "must stay OUT of process ... an in-process reopen
' shares PowerPoint's cache with the writer and cannot be trusted in either
' direction". The same rule has to hold inside the add-in, because at work
' there is no Python and that script does not exist there.
'
' A .pptx IS A ZIP, and Windows' Shell can read inside one without Office. So
' the file is copied, docProps/custom.xml extracted from the COPY, and the
' value read from its bytes -- sharing no process, no cache and no code with
' the writer. Probed against this 49MB deck before being relied on: extraction
' returned in well under a second.
'
' SaveAs rather than Save: Save performs an incremental rewrite that does not
' always regenerate docProps/custom.xml, while SaveAs forces a full rewrite.
' That is the measured direction, not a guarantee -- hence write, verify,
' retry, and fail LOUDLY rather than return a wrong answer.
' Any custom document property, read from the SAVED FILE.
'
' Generalised from PeriodOnDisk on 2026-08-08 so onboarding can verify the slide
' type registration the same way -- it was confirming that registration through
' DeckRegistry.LookupType(pres, ...), which reads the live Presentation object,
' i.e. PowerPoint's cache. That check passes against a deck that never saved,
' which is the exact defect Start a Quarter had that morning, sitting in the flow
' this module's own comment calls "the one write in this add-in that matters".
Public Function PropertyOnDisk(deckPath As String, propertyName As String, Optional ByRef trace As String) As String
    On Error GoTo Failed

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    trace = "fso ok"
    If Not fso.FileExists(deckPath) Then
        trace = trace & " | file missing: " & deckPath
        Exit Function
    End If

    Dim work As String
    ' .Path spelled out: GetSpecialFolder returns a Folder object, and relying on
    ' its default property to coerce to a string is the kind of implicit
    ' behaviour that reads fine and breaks somewhere else.
    work = fso.BuildPath(fso.GetSpecialFolder(2).Path, "dsverify_" & Format(Now, "hhnnss") & Int(Rnd * 10000))
    fso.CreateFolder work
    trace = trace & " | work=" & work

    Dim zipPath As String, outDir As String
    zipPath = fso.BuildPath(work, "deck.zip")
    outDir = fso.BuildPath(work, "out")
    fso.CreateFolder outDir
    fso.CopyFile deckPath, zipPath
    trace = trace & " | copied"

    Dim sh As Object
    Set sh = CreateObject("Shell.Application")

    ' VARIANT, NOT STRING, and this is the whole bug that made the first version
    ' return "" against a file that provably held the value. Shell.Application's
    ' Namespace takes a Variant; handed a String-typed VBA variable it returns
    ' NOTHING rather than raising, so the next call died with error 91 and the
    ' function reported "not on disk" for a perfectly good file. The identical
    ' COM calls from PowerShell worked, which is exactly why the technique looked
    ' proven and the code still failed.
    Dim zipVar As Variant, outVar As Variant, propsVar As Variant
    zipVar = zipPath
    outVar = outDir

    Dim zipNs As Object
    Set zipNs = sh.Namespace(zipVar)
    If zipNs Is Nothing Then
        trace = trace & " | Namespace(zip) returned Nothing"
        GoTo Cleanup
    End If

    Dim props As Object
    Set props = zipNs.ParseName("docProps")
    If props Is Nothing Then
        trace = trace & " | docProps NOT FOUND"
        GoTo Cleanup
    End If
    trace = trace & " | docProps=" & props.Path

    propsVar = props.Path
    Dim propsNs As Object
    Set propsNs = sh.Namespace(propsVar)
    If propsNs Is Nothing Then
        trace = trace & " | Namespace(docProps) returned Nothing"
        GoTo Cleanup
    End If

    Dim customFile As Object
    Set customFile = propsNs.ParseName("custom.xml")
    If customFile Is Nothing Then
        trace = trace & " | custom.xml NOT FOUND"
        GoTo Cleanup
    End If
    trace = trace & " | found custom.xml"

    ' 16 = "yes to all". CopyHere is ASYNCHRONOUS, so the file is waited for
    ' rather than assumed -- reading it immediately is a race, and a race that
    ' usually wins is worse than one that always loses.
    sh.Namespace(outVar).CopyHere customFile, 16

    Dim extracted As String
    extracted = fso.BuildPath(outDir, "custom.xml")

    Dim waited As Long
    Do While Not fso.FileExists(extracted) And waited < 100
        WaitAMoment
        waited = waited + 1
    Loop
    If Not fso.FileExists(extracted) Then
        trace = trace & " | EXTRACT TIMED OUT after " & waited
        GoTo Cleanup
    End If
    trace = trace & " | extracted"

    Dim xml As String
    xml = fso.OpenTextFile(extracted, 1).ReadAll

    Dim atProp As Long
    atProp = InStr(1, xml, propertyName, vbTextCompare)
    If atProp = 0 Then
        trace = trace & " | property name not in xml (len=" & Len(xml) & ")"
        GoTo Cleanup
    End If

    Dim openTag As Long, closeTag As Long
    openTag = InStr(atProp, xml, "<vt:lpwstr>")
    If openTag = 0 Then GoTo Cleanup
    openTag = openTag + Len("<vt:lpwstr>")
    closeTag = InStr(openTag, xml, "</vt:lpwstr>")
    If closeTag = 0 Then GoTo Cleanup

    PropertyOnDisk = Mid$(xml, openTag, closeTag - openTag)

Cleanup:
    On Error Resume Next
    fso.DeleteFolder work, True
    On Error GoTo 0
    Exit Function

Failed:
    trace = trace & " | ERROR " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not fso Is Nothing Then fso.DeleteFolder work, True
    On Error GoTo 0
End Function

' Saves the deck and CONFIRMS the file changed, or says why not.
'
' 2026-08-08, measured on the rig: Apply Approved reported "16 written, 0 failed",
' took a backup, re-checked every change against its slide -- and the deck file
' was untouched. Nothing in RibbonUI ever called pres.Save. Sixteen slide fields
' of work sat in volatile memory behind a dialog that said it had succeeded, and
' the workbook's review sheet with them.
'
' Verified by the file's MODIFICATION TIME advancing, which no in-process read
' can fake. That is a weaker claim than reading the content back and it is the
' honest one: it proves the file was rewritten, not that any particular value is
' in it. Every field written here has already been re-read from its own shape by
' the caller; what was missing was evidence the file moved at all.
'
' Falls back to SaveAs, which forces a full package rewrite: plain Save performs
' an incremental one that this project has measured losing writes on this deck.
Public Function SaveDeckVerified(pres As Object) As String
    On Error GoTo Failed

    Dim path As String
    path = pres.FullName

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then
        SaveDeckVerified = "The deck has never been saved to a file, so there is " & _
            "nothing to save into. Use File > Save As first."
        Exit Function
    End If

    Dim before As Date
    before = fso.GetFile(path).DateLastModified

    On Error Resume Next
    pres.Save
    Err.Clear
    On Error GoTo 0

    If fso.GetFile(path).DateLastModified > before Then Exit Function     ' "" = saved

    ' Save reported nothing and moved nothing. Force the full rewrite.
    Dim retryErr As String
    On Error Resume Next
    pres.SaveAs path, 24                          ' ppSaveAsOpenXMLPresentation
    If Err.Number <> 0 Then retryErr = "Error " & Err.Number & ": " & Err.Description
    Err.Clear
    On Error GoTo 0

    If fso.GetFile(path).DateLastModified > before Then Exit Function     ' "" = saved

    SaveDeckVerified = "THE DECK WAS NOT SAVED." & vbCrLf & vbCrLf & _
        path & vbCrLf & vbCrLf & _
        "The changes are in PowerPoint's memory only. Save the deck yourself " & _
        "before closing it, and check the change is there afterwards." & _
        IIf(retryErr = "", "", vbCrLf & vbCrLf & retryErr)
    Exit Function

Failed:
    SaveDeckVerified = "Could not save the deck." & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description
End Function


' The deck's period, from the saved file. Kept as its own name because it is the
' one every caller asks for.
Public Function PeriodOnDisk(deckPath As String, Optional ByRef trace As String) As String
    PeriodOnDisk = PropertyOnDisk(deckPath, PROP_DECK_PERIOD, trace)
End Function

' Returns "" when the period is confirmed on disk, otherwise a message saying
' exactly what happened. NEVER returns "" on an unverified write.
Public Function SetDeckPeriodVerified(pres As Object, period As String, ByVal attempts As Long) As String
    If attempts < 1 Then attempts = 1

    Dim path As String
    path = pres.FullName

    Dim n As Long
    For n = 1 To attempts
        On Error Resume Next
        WriteStringProperty pres, PROP_DECK_PERIOD, period
        pres.SaveAs path, 24            ' ppSaveAsOpenXMLPresentation -- forces a full rewrite
        Dim writeErr As String
        writeErr = ""
        If Err.Number <> 0 Then writeErr = "Error " & Err.Number & ": " & Err.Description
        Err.Clear
        On Error GoTo 0

        If PeriodOnDisk(path) = period Then Exit Function      ' "" = confirmed

        If n = attempts Then
            SetDeckPeriodVerified = "THE PERIOD DID NOT REACH THE FILE after " & attempts & _
                " attempt(s)." & vbCrLf & vbCrLf & _
                "Asked for: " & period & vbCrLf & _
                "On disk:   " & IIf(PeriodOnDisk(path) = "", "(nothing)", PeriodOnDisk(path)) & _
                IIf(writeErr = "", "", vbCrLf & writeErr) & vbCrLf & vbCrLf & _
                "Do not draft or sync until this is right. Everything downstream " & _
                "filters on the period, and a deck whose period is wrong reads as " & _
                "a clean run of zero rows."
        End If
    Next n
End Function

' A short pause that does not need a Windows API declaration -- kept separate
' so the wait in PeriodOnDisk reads as a wait rather than as arithmetic.
Private Sub WaitAMoment()
    ' NOT named "until": Until is a VBA keyword and using it as a variable is a
    ' compile error -- which stops the WHOLE project, not just this Sub.
    Dim deadline As Date
    deadline = DateAdd("s", 1, Now)
    Do While Now < deadline
        DoEvents
    Loop
End Sub

' Rolling forward is an EXPLICIT act with a stated from-and-to, never inferred.
' Returns the text a caller should show; does not itself decide anything.
Public Function AdvancePeriodText(oldPeriod As String, newPeriod As String) As String
    If oldPeriod = "" Then
        AdvancePeriodText = "This deck has no period recorded." & vbCrLf & _
            "Setting it to '" & newPeriod & "'."
    ElseIf StrComp(oldPeriod, newPeriod, vbTextCompare) = 0 Then
        AdvancePeriodText = "This deck is already '" & newPeriod & "'. Nothing to do."
    Else
        AdvancePeriodText = "Roll this deck forward?" & vbCrLf & vbCrLf & _
            "    from:  " & oldPeriod & vbCrLf & _
            "    to:    " & newPeriod & vbCrLf & vbCrLf & _
            "Every quarterly field will then be read from the '" & newPeriod & "'" & vbCrLf & _
            "rows of the register. Entity-static fields are unaffected."
    End If
End Function

' What to say when a caller supplies a period that disagrees with the deck.
'
' The DECK wins, always. A caller-supplied period is a habit or a script
' default; the deck's property was written by somebody rolling it forward on
' purpose. Returning "" means they agree and there is nothing to say.
Public Function PeriodMismatchText(deckPeriod As String, suppliedPeriod As String) As String
    If suppliedPeriod = "" Then Exit Function
    If deckPeriod = "" Then Exit Function
    If StrComp(deckPeriod, suppliedPeriod, vbTextCompare) = 0 Then Exit Function

    PeriodMismatchText = "PERIOD MISMATCH -- using the DECK's period, not the one supplied." & vbCrLf & _
        "    deck says:     " & deckPeriod & vbCrLf & _
        "    you supplied:  " & suppliedPeriod & vbCrLf & _
        "The deck's period was written when somebody rolled it forward on purpose." & vbCrLf & _
        "A supplied period is usually a habit or a script default, which is exactly" & vbCrLf & _
        "how a copied deck comes to report last quarter's content."
End Function
