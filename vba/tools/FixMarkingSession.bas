Attribute VB_Name = "FixMarkingSession"
Option Explicit

' RESCUE A MARKING SESSION WITHOUT REDOING IT.
'
' Written 2026-08-01, when Rohan marked a slide's worth of fields on his real
' deck, one of them an icon, and Bulk Onboard refused the whole batch. The only
' route the add-in offers is Clear Marked Fields, which discards ALL of them --
' "that marking took friggin ages surely it can be copied rather than clearing
' it?" It can. The session is not just in memory.
'
' Marks live in the presentation's custom document property
' "DeckSyncMarkingSession", one line per field:
'
'     ShapeName|FieldName|Type|Volatility
'
' and RestoreMarkingSession matches shapes back BY NAME. So dropping one line
' unmarks exactly one field and leaves the rest intact.
'
' PASTE THIS INTO THE PRESENTATION'S OWN VBA PROJECT (Alt+F11, Insert >
' Module), not the add-in's. Run ListMarks first. Delete the module afterwards;
' nothing here needs to persist.

Private Const SESSION_PROP As String = "DeckSyncMarkingSession"

' ---------------------------------------------------------------------------
' Run this FIRST. Shows every marked field, numbered, and changes nothing.
' ---------------------------------------------------------------------------
Public Sub ListMarks()
    Dim raw As String
    raw = ReadSession()
    If raw = "" Then
        MsgBox "No marking session on this presentation.", vbInformation, "Marks"
        Exit Sub
    End If

    Dim lines() As String
    lines = Split(raw, vbCrLf)

    Dim out As String, n As Long
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If Trim(lines(i)) <> "" Then
            n = n + 1
            Dim p() As String
            p = Split(lines(i), "|")
            out = out & n & ".  shape: " & p(0)
            If UBound(p) >= 1 Then out = out & "   field: " & p(1)
            If UBound(p) >= 2 Then out = out & "   (" & p(2) & ")"
            out = out & vbCrLf
        End If
    Next i

    MsgBox n & " marked field(s):" & vbCrLf & vbCrLf & out & vbCrLf & _
           "To drop one, run RemoveMark with its SHAPE name.", vbInformation, "Marks"
End Sub

' ---------------------------------------------------------------------------
' Drop one field by shape name. Everything else is kept.
'
' Edit the shape name below, then run. Case-insensitive, and it reports how
' many lines it removed so a typo shows up as "removed 0" rather than as a
' silent no-op that looks like success.
' ---------------------------------------------------------------------------
Public Sub RemoveMark()
    Const SHAPE_TO_DROP As String = "Graphic 285"      ' <-- edit this

    Dim raw As String
    raw = ReadSession()
    If raw = "" Then
        MsgBox "No marking session on this presentation -- nothing to remove.", vbExclamation, "Marks"
        Exit Sub
    End If

    Dim lines() As String
    lines = Split(raw, vbCrLf)

    Dim kept As String, removed As Long, keptCount As Long
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If Trim(lines(i)) <> "" Then
            Dim p() As String
            p = Split(lines(i), "|")
            If StrComp(Trim(p(0)), Trim(SHAPE_TO_DROP), vbTextCompare) = 0 Then
                removed = removed + 1
            Else
                kept = kept & vbCrLf & lines(i)
                keptCount = keptCount + 1
            End If
        End If
    Next i

    If removed = 0 Then
        MsgBox "Removed 0 -- no marked field has the shape name '" & SHAPE_TO_DROP & "'." & vbCrLf & vbCrLf & _
               "Nothing was changed. Run ListMarks and copy the name exactly.", vbExclamation, "Marks"
        Exit Sub
    End If

    WriteSession kept

    MsgBox "Removed " & removed & " mark(s) for '" & SHAPE_TO_DROP & "'." & vbCrLf & _
           keptCount & " mark(s) kept." & vbCrLf & vbCrLf & _
           "NOW: save, close the deck, and reopen it. The add-in holds the old " & _
           "list in memory until then, and reopening is what makes it read this " & _
           "corrected one back.", vbInformation, "Marks"
End Sub

Private Function ReadSession() As String
    On Error Resume Next
    ReadSession = CStr(ActivePresentation.CustomDocumentProperties(SESSION_PROP).Value)
    On Error GoTo 0
End Function

Private Sub WriteSession(value As String)
    On Error Resume Next
    ActivePresentation.CustomDocumentProperties(SESSION_PROP).Value = value
    On Error GoTo 0
End Sub
