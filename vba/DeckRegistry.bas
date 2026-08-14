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
' readFailed DISTINGUISHES "THE FILE SAYS NO" FROM "I COULD NOT ASK".
'
' Both used to return "", and on 2026-08-09 that cost the readiness sheet its
' credibility: a OneDrive-hosted deck reported "Period: BLOCKED -- not set in
' the saved file" while the file's bytes held Q3F26. Seven of the eight exit
' paths below are failures to READ, and every one of them was reported to the
' user as a confident statement about the deck's contents.
'
' The sheet's own rule 1 (Readiness.bas:32) requires three states, not two: a
' check that could not run must say so. It cannot honour that while its only
' input is an empty string.
'
' DEFAULTS TO TRUE, deliberately. Any path added later that forgets to clear it
' reports CANNOT TELL, which is recoverable; the alternative default reports a
' confident wrong answer, which is what this whole function exists to prevent.
' A cloud-hosted deck reports its FullName as a URL, and FileSystemObject cannot
' read one. Windows keeps the mapping from cloud URL to local synced folder in
' HKCU\SOFTWARE\SyncEngines\Providers\OneDrive: each sync root has a
' UrlNamespace ("https://d.docs.live.net") and a MountPoint
' ("C:\Users\rohan\OneDrive").
'
' THE REMAINDER DOES NOT MAP STRAIGHT ACROSS. Personal OneDrive puts a user-id
' segment after the namespace that does not appear in the local path; business
' and SharePoint namespaces are longer and have no such segment. Rather than
' encode either convention -- and get it wrong on the one that matters, at work,
' where nothing can be debugged -- this tries the candidate and CONFIRMS IT BY
' THE FILE BEING THERE, dropping one leading segment and retrying if not. It
' returns only a path that exists, so a wrong guess is indistinguishable from no
' answer, which is the safe direction.
'
' This reads the LOCAL SYNCED FILE, which is genuinely on disk and shares no
' state with PowerPoint -- the property this whole function exists to have. The
' local copy is also what PowerPoint actually wrote; the upload is OneDrive's
' problem afterwards.
Public Function LocalPathForUrl(url As String, Optional ByRef trace As String) As String
    LocalPathForUrl = ""

    If LCase$(Left$(url, 5)) <> "http:" And LCase$(Left$(url, 6)) <> "https:" Then
        LocalPathForUrl = url                     ' already a filesystem path
        Exit Function
    End If

    Dim decoded As String
    decoded = Replace(url, "%20", " ")
    decoded = Replace(decoded, "%27", "'")
    decoded = Replace(decoded, "%28", "(")
    decoded = Replace(decoded, "%29", ")")
    decoded = Replace(decoded, "%26", "&")

    Dim reg As Object
    On Error Resume Next
    Set reg = GetObject("winmgmts:\\.\root\default:StdRegProv")
    On Error GoTo 0
    If reg Is Nothing Then
        trace = trace & " | no registry provider -- cannot map URL"
        Exit Function
    End If

    Const HKCU As Long = &H80000001
    Const BASE As String = "SOFTWARE\SyncEngines\Providers\OneDrive"

    Dim keys As Variant
    On Error Resume Next
    If reg.EnumKey(HKCU, BASE, keys) <> 0 Then
        On Error GoTo 0
        trace = trace & " | no OneDrive sync roots registered"
        Exit Function
    End If
    On Error GoTo 0
    If IsEmpty(keys) Then Exit Function
    If Not IsArray(keys) Then Exit Function

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        Dim ns As String, mp As String
        ns = "": mp = ""
        On Error Resume Next
        reg.GetStringValue HKCU, BASE & "\" & keys(i), "UrlNamespace", ns
        reg.GetStringValue HKCU, BASE & "\" & keys(i), "MountPoint", mp
        On Error GoTo 0

        If ns <> "" And mp <> "" Then
            If StrComp(Left$(decoded, Len(ns)), ns, vbTextCompare) = 0 Then
                Dim rest As String
                rest = Mid$(decoded, Len(ns) + 1)
                rest = Replace(rest, "/", "\")
                Do While Left$(rest, 1) = "\"
                    rest = Mid$(rest, 2)
                Loop

                ' Try the whole remainder, then with one leading segment
                ' dropped, then two. Personal OneDrive needs one dropped (the
                ' user id); business usually needs none.
                Dim attempt As Long
                For attempt = 0 To 2
                    Dim candidate As String
                    candidate = fso.BuildPath(mp, rest)
                    If fso.FileExists(candidate) Then
                        trace = trace & " | URL mapped to " & candidate
                        LocalPathForUrl = candidate
                        Exit Function
                    End If
                    Dim cut As Long
                    cut = InStr(rest, "\")
                    If cut = 0 Then Exit For
                    rest = Mid$(rest, cut + 1)
                Next attempt
            End If
        End If
    Next i

    trace = trace & " | URL matched no synced folder that holds this file"
End Function

Public Function PropertyOnDisk(deckPath As String, propertyName As String, _
                               Optional ByRef trace As String, _
                               Optional ByRef readFailed As Boolean) As String
    On Error GoTo Failed
    readFailed = True

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    trace = "fso ok"

    ' A cloud-hosted deck arrives here as a URL. Translate before looking, or
    ' every check below reports a perfectly good deck as empty -- which is what
    ' the readiness sheet did on 2026-08-09.
    If LCase$(Left$(deckPath, 4)) = "http" Then
        Dim mapped As String
        mapped = LocalPathForUrl(deckPath, trace)
        If mapped = "" Then
            trace = trace & " | cloud deck, no local copy found: " & deckPath
            Exit Function
        End If
        deckPath = mapped
    End If
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

    ' THE ZIP OPENED. From here the file has been read, and everything below is
    ' a statement about its CONTENTS. A deck created with no custom properties
    ' never gets a docProps/custom.xml part written at all, so "not found" here
    ' is an honest "this deck declares nothing", not a failure to look -- the
    ' distinction the blank-deck fixture in TestRunner exists to hold.
    readFailed = False

    Dim props As Object
    Set props = zipNs.ParseName("docProps")
    If props Is Nothing Then
        trace = trace & " | docProps NOT FOUND -- deck declares no properties"
        GoTo Cleanup
    End If
    trace = trace & " | docProps=" & props.Path

    propsVar = props.Path
    Dim propsNs As Object
    Set propsNs = sh.Namespace(propsVar)
    If propsNs Is Nothing Then
        trace = trace & " | Namespace(docProps) returned Nothing"
        readFailed = True          ' found the folder, could not open it -- a failure to look
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
        readFailed = True          ' custom.xml exists and we never got to read it
        GoTo Cleanup
    End If
    trace = trace & " | extracted"

    Dim xml As String
    xml = fso.OpenTextFile(extracted, 1).ReadAll

    ' THE READ SUCCEEDED. From here an empty answer is a fact about the deck,
    ' not a failure to look -- this is the only place that can honestly say so.
    readFailed = False

    Dim atProp As Long
    atProp = InStr(1, xml, propertyName, vbTextCompare)
    If atProp = 0 Then
        trace = trace & " | property name not in xml (len=" & Len(xml) & ")"
        GoTo Cleanup
    End If

    ' Present but unreadable is NOT the same as absent: the property is in the
    ' file and this code cannot parse it, which is a defect to surface rather
    ' than a deck to send someone off to reconfigure.
    Dim openTag As Long, closeTag As Long
    openTag = InStr(atProp, xml, "<vt:lpwstr>")
    If openTag = 0 Then
        trace = trace & " | '" & propertyName & "' present but no <vt:lpwstr> value"
        readFailed = True
        GoTo Cleanup
    End If
    openTag = openTag + Len("<vt:lpwstr>")
    closeTag = InStr(openTag, xml, "</vt:lpwstr>")
    If closeTag = 0 Then
        trace = trace & " | '" & propertyName & "' value never closed"
        readFailed = True
        GoTo Cleanup
    End If

    PropertyOnDisk = Mid$(xml, openTag, closeTag - openTag)

Cleanup:
    On Error Resume Next
    fso.DeleteFolder work, True
    On Error GoTo 0
    Exit Function

Failed:
    trace = trace & " | ERROR " & Err.Number & ": " & Err.Description
    ' Set explicitly rather than relying on the entry default: the zip-opened
    ' branch clears it partway through, so an error raised AFTER that point
    ' would otherwise be reported as a clean read of an empty value.
    readFailed = True
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

    ' A CLOUD-HOSTED DECK REPORTS ITS FullName AS AN https:// URL, AND
    ' FileSystemObject ANSWERS False FOR ONE RATHER THAN RAISING. So the guard
    ' below used to read "the deck has never been saved to a file" about a 49MB
    ' deck that had been saved all evening, and -- far worse -- Exit Function
    ' before pres.Save was ever called. The one function written to guarantee the
    ' save was the one that skipped it, silently, on exactly the configuration the
    ' work machine is always in. Measured 2026-08-13 from the Run Log's own
    ' "---- NOT SAVED ----" block, which printed the URL inside the sentence
    ' claiming no file existed. AutoSave being on was the only reason the evening's
    ' KEY_EVENTS publish survived; with it off, the period write had already failed
    ' four times the same way.
    '
    ' LocalPathForUrl resolves the URL to the local synced file and returns "" when
    ' it cannot. Note the ORDER: this is used ONLY for the existence and
    ' modification-time checks. pres.Save still targets the document itself, which
    ' is what PowerPoint knows how to write; the local copy is merely where the
    ' evidence lands.
    Dim checkPath As String
    checkPath = LocalPathForUrl(path)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' THREE STATES, NOT TWO -- Readiness.bas:32's rule, and the reason this
    ' function exists at all. "I could not check" must never be reported as either
    ' "saved" or "never existed". The save is ATTEMPTED regardless, because a deck
    ' we cannot verify is still a deck that must be written.
    If checkPath = "" Or Not fso.FileExists(checkPath) Then
        On Error Resume Next
        pres.Save
        Dim blindErr As String
        If Err.Number <> 0 Then blindErr = "Error " & Err.Number & ": " & Err.Description
        Err.Clear
        On Error GoTo 0

        SaveDeckVerified = "THE SAVE COULD NOT BE VERIFIED." & vbCrLf & vbCrLf & _
            path & vbCrLf & vbCrLf & _
            "The deck was told to save, but its file could not be located on this " & _
            "PC to confirm the bytes moved. Open the deck's folder and check the " & _
            "modified time yourself before closing it." & _
            IIf(blindErr = "", "", vbCrLf & vbCrLf & blindErr)
        Exit Function
    End If

    Dim before As Date
    before = fso.GetFile(checkPath).DateLastModified

    On Error Resume Next
    pres.Save
    Err.Clear
    On Error GoTo 0

    If fso.GetFile(checkPath).DateLastModified > before Then Exit Function     ' "" = saved

    ' Save reported nothing and moved nothing. Force the full rewrite.
    Dim retryErr As String
    On Error Resume Next
    pres.SaveAs path, 24                          ' ppSaveAsOpenXMLPresentation
    If Err.Number <> 0 Then retryErr = "Error " & Err.Number & ": " & Err.Description
    Err.Clear
    On Error GoTo 0

    If fso.GetFile(checkPath).DateLastModified > before Then Exit Function     ' "" = saved

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
Public Function PeriodOnDisk(deckPath As String, Optional ByRef trace As String, _
                             Optional ByRef readFailed As Boolean) As String
    PeriodOnDisk = PropertyOnDisk(deckPath, PROP_DECK_PERIOD, trace, readFailed)
End Function

' Returns "" when the period is confirmed on disk, otherwise a message saying
' exactly what happened. NEVER returns "" on an unverified write.
Public Function SetDeckPeriodVerified(pres As Object, period As String, ByVal attempts As Long) As String
    If attempts < 1 Then attempts = 1

    Dim path As String
    path = pres.FullName

    Dim n As Long
    For n = 1 To attempts
        ' PLAIN SAVE FIRST, SaveAs ONLY IF THAT DID NOT LAND.
        '
        ' SaveAs was chosen because a plain Save is incremental and does not
        ' always regenerate docProps/custom.xml -- true, and the reason this
        ' used SaveAs alone. But on Rohan's real 49MB deck, opened from a
        ' OneDrive-synced folder, `SaveAs path, 24` returns WITHOUT RAISING and
        ' writes nothing at all: four attempts, file untouched, and the error
        ' text this function appends stayed empty because Err was never set.
        '
        ' Plain Save works there -- PowerPoint's own saves landed on that file
        ' all morning. So try the one that works, verify, and escalate to the
        ' full rewrite only when the cheap write did not reach the disk. Both
        ' are verified the same way, so neither is trusted on its own.
        On Error Resume Next
        WriteStringProperty pres, PROP_DECK_PERIOD, period
        pres.Save
        Dim writeErr As String
        writeErr = ""
        If Err.Number <> 0 Then writeErr = "Save -- Error " & Err.Number & ": " & Err.Description
        Err.Clear
        On Error GoTo 0

        If PeriodOnDisk(path) = period Then Exit Function      ' "" = confirmed

        On Error Resume Next
        pres.SaveAs path, 24            ' ppSaveAsOpenXMLPresentation -- forces a full rewrite
        If Err.Number <> 0 Then
            If writeErr <> "" Then writeErr = writeErr & vbCrLf
            writeErr = writeErr & "SaveAs -- Error " & Err.Number & ": " & Err.Description
        End If
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

' The paired workbook path as it is stored IN THE SAVED FILE, or "" if absent.
'
' Deliberately NOT GetWorkbookPath. That resolver falls back to a workbook sitting
' beside the deck when the stored path no longer exists, which is right for opening
' the pair and WRONG for verification: it can return a path that was never written
' and read as confirmation of the write. A verifier must answer "what is actually
' recorded", not "what would we open".
Public Function WorkbookPathOnDisk(deckPath As String, Optional ByRef trace As String, _
                                   Optional ByRef readFailed As Boolean) As String
    WorkbookPathOnDisk = PropertyOnDisk(deckPath, WORKBOOK_PATH_PROPERTY_NAME, trace, readFailed)
End Function

' Returns "" when the pairing is confirmed in the saved file, otherwise a message
' saying exactly what happened. NEVER returns "" on an unverified write.
'
' Same shape and same reason as SetDeckPeriodVerified above. RepointWorkbookUI used
' to write the property, read it straight back through the same Presentation object,
' report "Paired workbook is now: <path>", and then tell the person to save the deck
' themselves -- which is the 2026-08-08 Start-a-Quarter defect exactly, 130 lines
' below the fixed version, with the remedy already sitting in this module. An
' in-process read returns PowerPoint's cache and confirms the write whether or not a
' byte moved; the instruction to "save it so this survives" is the very step that
' silently did not happen when it was reported for the period.
Public Function SetWorkbookPathVerified(pres As Object, newPath As String, ByVal attempts As Long) As String
    If attempts < 1 Then attempts = 1

    Dim path As String
    path = pres.FullName

    Dim n As Long
    For n = 1 To attempts
        ' SAME ORDER AS SaveDeckVerified AND SetDeckPeriodVerified: plain Save
        ' first, full rewrite only if that did not land. SaveAs alone returns
        ' clean and writes nothing on a large deck in a synced folder, which is
        ' how the period write failed four times in a row while PowerPoint's own
        ' saves were landing on the same file all morning.
        On Error Resume Next
        RepointWorkbook pres, newPath
        pres.Save
        Dim writeErr As String
        writeErr = ""
        If Err.Number <> 0 Then writeErr = "Save -- Error " & Err.Number & ": " & Err.Description
        Err.Clear
        On Error GoTo 0

        If StrComp(WorkbookPathOnDisk(path), newPath, vbTextCompare) = 0 Then Exit Function

        On Error Resume Next
        pres.SaveAs path, 24            ' ppSaveAsOpenXMLPresentation -- forces a full rewrite
        If Err.Number <> 0 Then
            If writeErr <> "" Then writeErr = writeErr & vbCrLf
            writeErr = writeErr & "SaveAs -- Error " & Err.Number & ": " & Err.Description
        End If
        Err.Clear
        On Error GoTo 0

        If StrComp(WorkbookPathOnDisk(path), newPath, vbTextCompare) = 0 Then Exit Function

        If n = attempts Then
            SetWorkbookPathVerified = "THE PAIRING DID NOT REACH THE FILE after " & attempts & _
                " attempt(s)." & vbCrLf & vbCrLf & _
                "Asked for: " & newPath & vbCrLf & _
                "On disk:   " & IIf(WorkbookPathOnDisk(path) = "", "(nothing)", WorkbookPathOnDisk(path)) & _
                IIf(writeErr = "", "", vbCrLf & writeErr) & vbCrLf & vbCrLf & _
                "Do not draft or sync until this is right. A deck pointed at the wrong " & _
                "workbook reads every field from the wrong register and reports success."
        End If
    Next n
End Function


' ---------------------------------------------------------------------
' THE PAIRING, BOTH WAYS
' ---------------------------------------------------------------------
'
' The design was always two-sided and only one side was ever maintained:
'
'   deck  --> workbook   DeckSyncWorkbookPath, a PATH. Repointed by the person,
'                        verified against the deck's own bytes above.
'   workbook --> deck    DeckReference, this deck's DeckSyncId GUID. Written by
'                        ExcelOutput.CreateSheet at onboarding and never again.
'
' Path one way, IDENTITY the other, and that asymmetry is deliberate: a path
' breaks the moment OneDrive moves a file, a GUID does not. specs/deck-registry.md
' says this "directly closes input-contract.md's cross-wiring risk".
'
' It did not close it. Found 2026-08-14: the GUID was written and NEVER READ for
' its purpose -- the only readers were one struct assignment and a MsgBox in a
' demo whose expected value is the literal string "deck-v1". So the tool opened
' whatever the stored path pointed at and began writing, with no one ever asking
' the workbook whether it belonged to this deck. SetWorkbookPathVerified's own
' error text already names the consequence: "A deck pointed at the wrong workbook
' reads every field from the wrong register and reports success."
'
' Written-but-never-read is this project's signature defect, and it is the same
' shape as the tested picture injector, the tested progress bars and the tested
' publish path -- machinery that works perfectly and that nothing can reach.

' Stamp this deck's identity into the workbook, so a repoint updates BOTH ends.
'
' Returns "" when the stamp is believed to have reached the file, else why not.
'
' The read-back here goes through the same Workbook object that was just written,
' so it confirms the OBJECT, not the bytes -- the cache-answering-for-the-file
' trap this project has been caught by twice. SaveWorkbookVerified is what
' actually crosses the boundary; that is why its result, not the read-back, is
' what this function reports on.
Public Function StampPairing(pres As Object, wb As Object) As String
    If wb Is Nothing Then
        StampPairing = "The workbook could not be opened, so the pairing could not be stamped into it."
        Exit Function
    End If

    Dim deckId As String
    deckId = GetOrCreateDeckId(pres)

    On Error Resume Next
    ExcelOutput.WriteDeckReference wb, deckId
    If Err.Number <> 0 Then
        StampPairing = "Could not write this deck's identity into the workbook -- Error " & _
            Err.Number & ": " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    StampPairing = WorkbookBridge.SaveWorkbookVerified(wb)
End Function


' Why this workbook does not belong to this deck, or "" if it does.
'
' Deliberately a QUESTION THE CALLER ASKS rather than a refusal inside
' OpenOrGetWorkbook -- the same shape, and for the same reason, as
' WorkbookBridge.WriteBlockedReason: some paths legitimately open a register only
' to read it, and the caller that is about to WRITE is the one that has to ask.
'
' An UNSTAMPED workbook is not a mismatch. Every register created before this
' pairing existed has a blank DeckReference, and refusing those would strand a
' person on the one machine where they cannot debug anything. Blank means
' "unknown, stamp it"; a DIFFERENT GUID means "this is someone else's register".
Public Function PairingProblem(pres As Object, wb As Object) As String
    If wb Is Nothing Then Exit Function

    Dim stamped As String
    On Error Resume Next
    stamped = ExcelOutput.ReadDeckReference(wb)
    On Error GoTo 0

    Dim where As String
    where = ""
    On Error Resume Next
    where = wb.FullName
    On Error GoTo 0

    PairingProblem = PairingVerdict(stamped, GetOrCreateDeckId(pres), where)
End Function


' The DECISION, with no CustomDocumentProperties access, so it can be tested
' without Excel and without a deck -- same split as BuildTypeRegistration above.
'
' Three cases and they are not symmetric:
'   blank stamp     -> "", every register predating the pairing looks like this
'   same GUID       -> ""
'   different GUID  -> the sentence, naming BOTH ids
'
' Naming both is not decoration. "Could not open the paired workbook" appeared
' three times in one morning with three different causes because the message
' named one thing and discarded the rest; a mismatch that shows only "wrong
' workbook" would repeat that exactly.
Public Function PairingVerdict(stampedId As String, deckId As String, workbookPath As String) As String
    If Trim(stampedId) = "" Then Exit Function
    If StrComp(Trim(stampedId), Trim(deckId), vbTextCompare) = 0 Then Exit Function

    PairingVerdict = _
        "THIS WORKBOOK BELONGS TO A DIFFERENT DECK." & vbCrLf & vbCrLf & _
        "Workbook: " & IIf(workbookPath = "", "(unknown path)", workbookPath) & vbCrLf & _
        "  it says it belongs to deck:  " & Trim(stampedId) & vbCrLf & _
        "  this deck is:                " & Trim(deckId) & vbCrLf & vbCrLf & _
        "Writing into it would put this deck's content into another deck's " & _
        "register, and every stage after that would report success." & vbCrLf & vbCrLf & _
        "Use Repoint Workbook to pair this deck with the right file, or open the " & _
        "deck that owns this one."
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
