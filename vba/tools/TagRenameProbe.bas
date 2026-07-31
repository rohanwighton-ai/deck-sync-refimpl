Attribute VB_Name = "TagRenameProbe"
Option Explicit

' One-off empirical probe -- NOT part of the shipped add-in.
'
' Q8 of the Excel Control Layer confirmation asks for the rename migration plan
' and its rollback. Both depend on behaviour that has not been measured:
'
'   1. Does Tags.Add with an EXISTING tag name overwrite the value, or add a
'      second tag with the same name? The whole migration is "set role to the
'      new FieldID" -- if Add appends rather than replaces, every migrated
'      shape ends up carrying two role tags and the matcher's own
'      ambiguity guard would then refuse to inject into it. That turns a
'      rename into a deck-wide outage.
'
'   2. Can every tag on a shape be ENUMERATED (Count / Name(i) / Value(i))?
'      A migration tool has to find what is there before changing it, and the
'      D5 uniqueness check has to read identity tags across a whole deck.
'      Neither is possible if tags can only be read by asking for a name you
'      already know.
'
' Both are the kind of claim this project has been wrong about before when it
' asserted rather than probed -- see feedback_verify_office_automation.
Public Function TagRenameProbe() As String
    Dim r As String

    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim sld As Object
    Set sld = pres.Slides.Add(pres.Slides.count + 1, 12)

    Dim shp As Object
    Set shp = sld.Shapes.AddTextbox(1, 50, 50, 300, 50)
    shp.TextFrame.TextRange.Text = "probe"

    ' --- Q1: does Add overwrite an existing tag? --------------------------
    shp.Tags.Add "role", "About text"
    Dim countAfterFirst As Long
    countAfterFirst = shp.Tags.count

    shp.Tags.Add "role", "PROGRESS_BODY"
    Dim countAfterSecond As Long
    countAfterSecond = shp.Tags.count

    r = "=== Tags.Add on an existing name ===" & vbCrLf & _
        "  Tags.Count after first Add:  " & countAfterFirst & vbCrLf & _
        "  Tags.Count after second Add: " & countAfterSecond & vbCrLf & _
        "  value now:                   '" & shp.Tags("role") & "'" & vbCrLf & _
        "  OVERWRITES (not duplicates): " & (countAfterSecond = countAfterFirst) & vbCrLf & _
        "  new value took:              " & (shp.Tags("role") = "PROGRESS_BODY") & vbCrLf & vbCrLf

    ' --- Q2: can tags be enumerated? --------------------------------------
    shp.Tags.Add "probe_second_tag", "yes"
    Dim listing As String
    Dim i As Long
    On Error Resume Next
    For i = 1 To shp.Tags.count
        listing = listing & "    " & i & ": " & shp.Tags.Name(i) & " = '" & shp.Tags.Value(i) & "'" & vbCrLf
    Next i
    Dim enumErr As Long
    enumErr = Err.Number
    On Error GoTo 0

    r = r & "=== Enumeration ===" & vbCrLf & _
        "  Tags.Count: " & shp.Tags.count & vbCrLf & _
        "  Err during enumeration: " & enumErr & vbCrLf & _
        listing & _
        "  ENUMERABLE: " & (enumErr = 0 And listing <> "") & vbCrLf & vbCrLf

    ' --- Q3: is a rename reversible? --------------------------------------
    ' The rollback story is only as good as this.
    shp.Tags.Add "role", "About text"
    r = r & "=== Rollback (set it straight back) ===" & vbCrLf & _
        "  value after reverting: '" & shp.Tags("role") & "'" & vbCrLf & _
        "  Tags.Count:            " & shp.Tags.count & vbCrLf & _
        "  REVERSIBLE:            " & (shp.Tags("role") = "About text") & vbCrLf

    sld.Delete
    TagRenameProbe = r
End Function
