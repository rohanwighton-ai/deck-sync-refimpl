Attribute VB_Name = "DeckRegistry"
Option Explicit

' Implements specs/deck-registry.md: the missing lookup a one-click ribbon
' button needs and no prior module provides -- "for this open deck, which
' workbook is paired with it, and where does each known slide type's
' template/worksheet live." Every existing engine entry point
' (RunSync.RunRoutineSync/RunPeriodRollover, DeckAdoption's templateSld
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

    If prop Is Nothing Then
        pres.CustomDocumentProperties.Add Name:=propertyName, _
            LinkToContent:=False, Type:=msoPropertyTypeString, Value:=value
    Else
        prop.Value = value
    End If
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
        If FileExists(beside) Then
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
    If Not FileExists(newPath) Then
        Err.Raise vbObjectError + 2, "DeckRegistry.RepointWorkbook", _
            "refusing to register a workbook path that does not exist: " & newPath
    End If
    SetWorkbookPath pres, newPath
End Sub

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
    SiblingOfDeck = deckDir & PATH_SEP & fileName
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
