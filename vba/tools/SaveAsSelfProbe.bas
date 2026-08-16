Attribute VB_Name = "SaveAsSelfProbe"
Option Explicit

' EXPERIMENTAL, STANDALONE PROBE -- NOT WIRED TO ANY BUTTON, NOT PART OF THE
' PRODUCTION BUILD. Self-contained deliberately (own copy of IsUrl/
' LocalPathForUrl/PropertyOnDisk, verbatim from DeckRegistry.bas) so this can
' be imported into a throwaway scratch presentation without pulling in
' DeckRegistry's other cross-module dependencies (ExcelOutput, WorkbookBridge,
' RunSync) -- a probe validates the TECHNIQUE in isolation, not the whole
' module. See zettel 20260808-a-probe-validates-the-technique-not-the-
' implementation.
'
' WHAT THIS TESTS. FIX-LIST P's cloud branch never attempts `SaveAs path, 24`
' on a cloud-hosted deck -- it was measured 2026-08-15 midday to raise
' 0x80CD1001 and leave the document read-only. But that measurement was taken
' BEFORE the same day's afternoon fix to PropertyOnDisk's ByRef->ByVal bug,
' under which reading the file (which every verifier does right before
' deciding whether to escalate) silently rewrote the caller's `path` variable
' from the true https:// URL to a local mapped path. So the SaveAs that got
' measured was "save this cloud-open document to a DIFFERENT local path", not
' "save it to itself" -- a genuinely different and more dangerous operation.
' Nobody has re-measured a clean SaveAs-to-self since the fix. This does.

' Copied verbatim from DeckRegistry.bas -- do not let this drift from that copy
' if both survive past this probe; it exists only to make this module import
' cleanly on its own.
Public Function IsUrl(path As String) As Boolean
    IsUrl = (LCase(Left(path, 5)) = "http:" Or LCase(Left(path, 6)) = "https:")
End Function

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

' ByVal ON deckPath, deliberately -- see DeckRegistry.bas's comment on the same
' function for why this is load-bearing and not stylistic.
Public Function PropertyOnDisk(ByVal deckPath As String, propertyName As String, _
                               Optional ByRef trace As String, _
                               Optional ByRef readFailed As Boolean) As String
    On Error GoTo Failed
    readFailed = True

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    trace = "fso ok"

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

    Dim zipVar As Variant, outVar As Variant, propsVar As Variant
    zipVar = zipPath
    outVar = outDir

    Dim zipNs As Object
    Set zipNs = sh.Namespace(zipVar)
    If zipNs Is Nothing Then
        trace = trace & " | Namespace(zip) returned Nothing"
        GoTo Cleanup
    End If

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
        readFailed = True
        GoTo Cleanup
    End If

    Dim customFile As Object
    Set customFile = propsNs.ParseName("custom.xml")
    If customFile Is Nothing Then
        trace = trace & " | custom.xml NOT FOUND"
        GoTo Cleanup
    End If
    trace = trace & " | found custom.xml"

    sh.Namespace(outVar).CopyHere customFile, 16

    Dim extracted As String
    extracted = fso.BuildPath(outDir, "custom.xml")

    Dim waited As Long
    Do While Not fso.FileExists(extracted) And waited < 100
        DoEvents
        waited = waited + 1
    Loop
    If Not fso.FileExists(extracted) Then
        trace = trace & " | EXTRACT TIMED OUT after " & waited
        readFailed = True
        GoTo Cleanup
    End If
    trace = trace & " | extracted"

    Dim xml As String
    xml = fso.OpenTextFile(extracted, 1).ReadAll

    readFailed = False

    Dim atProp As Long
    atProp = InStr(1, xml, propertyName, vbTextCompare)
    If atProp = 0 Then
        trace = trace & " | property name not in xml (len=" & Len(xml) & ")"
        GoTo Cleanup
    End If

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
    readFailed = True
    On Error Resume Next
    If Not fso Is Nothing Then fso.DeleteFolder work, True
    On Error GoTo 0
End Function

Private Sub WriteStringProperty(pres As Object, propertyName As String, value As String)
    Dim prop As Object
    On Error Resume Next
    Set prop = pres.CustomDocumentProperties(propertyName)
    On Error GoTo 0

    If Not prop Is Nothing Then
        On Error Resume Next
        pres.CustomDocumentProperties(propertyName).Delete
        On Error GoTo 0
    End If

    pres.CustomDocumentProperties.Add Name:=propertyName, _
        LinkToContent:=False, Type:=msoPropertyTypeString, Value:=value
End Sub

' THE ACTUAL TEST. Active SaveAs-to-self on every attempt -- no passive wait
' at all. `path` comes from pres.FullName ONCE and is never touched again in
' this function (nothing here reassigns it, ByRef or otherwise), so on a
' cloud-hosted presentation `pres.SaveAs path, 24` targets the document's own
' URL, unmangled. Returns "" when the write is confirmed FROM THE RAW FILE
' BYTES (via PropertyOnDisk, which copies and unzips the saved package --
' never trusts PowerPoint's own object cache). Also reports whether the
' presentation came out flagged read-only, which is the specific damage
' FIX-LIST P is worried this operation causes.
Public Function ProbeSaveAsToSelfVerified(pres As Object, propertyName As String, _
                                          value As String, ByVal attempts As Long) As String
    If attempts < 1 Then attempts = 1

    Dim path As String
    path = pres.FullName

    Dim n As Long
    For n = 1 To attempts
        WriteStringProperty pres, propertyName, value

        On Error Resume Next
        pres.SaveAs path, 24        ' ppSaveAsOpenXMLPresentation -- to SELF
        Dim writeErr As String
        writeErr = ""
        If Err.Number <> 0 Then writeErr = "SaveAs -- Error " & Err.Number & ": " & Err.Description
        Err.Clear
        On Error GoTo 0

        Dim nowReadOnly As Boolean
        On Error Resume Next
        nowReadOnly = pres.ReadOnly
        On Error GoTo 0

        If PropertyOnDisk(path, propertyName) = value Then
            ProbeSaveAsToSelfVerified = "" & IIf(nowReadOnly, "(landed, but presentation is now ReadOnly)", "")
            Exit Function
        End If

        If n = attempts Then
            ProbeSaveAsToSelfVerified = "DID NOT REACH THE FILE after " & attempts & " attempt(s)." & vbCrLf & _
                "Asked for: " & value & vbCrLf & _
                "On disk:   " & IIf(PropertyOnDisk(path, propertyName) = "", "(nothing)", PropertyOnDisk(path, propertyName)) & _
                " | ReadOnly now: " & nowReadOnly & _
                IIf(writeErr = "", "", vbCrLf & writeErr)
        End If
    Next n
End Function
