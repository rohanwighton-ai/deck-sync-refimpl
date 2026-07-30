Attribute VB_Name = "HiddenFixCheck"
Option Explicit

' One-off diagnostic -- NOT part of the shipped add-in, same posture as
' SyncRealDeck.bas / PreviewRealDeck.bas / AuditRealDeck.bas beside it.
'
' Proves S1's fix on a real deck: SlideDuplication.DuplicateAndTag copies the
' source slide via Slide.Duplicate, which carries SlideShowTransition.Hidden
' across as well as the slide-level tags. Because the master template is
' deliberately hidden, every record cloned from it arrived hidden -- a brand
' new project silently absent from the presented deck, reported by nothing.
'
' The test has to be a DELETE-AND-RECREATE rather than an inspection: the
' existing 3_P005 slide was created by the defective build, so its state says
' nothing about the fix. Only a slide created by the current code is evidence.
'
' Prints a full slide inventory before and after, because the assertion is a
' comparison between two specific slides -- the recreated record must be
' visible while the template it was cloned FROM stays hidden. A report showing
' only the new slide could not distinguish "the fix works" from "nothing is
' hidden any more", and the second would be a different bug.
Public Function HiddenFixCheck(deckPath As String, targetKey As String, Optional saveWhenDone As Boolean = False) As String
    Dim report As String

    Dim pres As Object
    ' Read-WRITE and WITH a window: RunSync.GatherInstances reads
    ' Application.ActivePresentation, so the deck must be the active one.
    Set pres = Application.Presentations.Open(deckPath, msoFalse, msoFalse, msoTrue)
    pres.Windows(1).Activate

    report = "Deck: " & deckPath & vbCrLf & vbCrLf & _
             "--- BEFORE ---" & vbCrLf & Inventory(pres)

    ' --- Delete the target record ----------------------------------------
    ' By instance_key, never by slide index: the index is exactly what
    ' resequencing moves around, so deleting "slide 5" would be deleting
    ' whatever happens to sit there this run.
    Dim i As Long
    Dim deleted As Boolean
    For i = pres.Slides.count To 1 Step -1
        Dim inst As SlideInstance
        inst = Resolve.ResolveSlideInstance(pres.Slides(i))
        If inst.HasInstanceKey And inst.InstanceKey = targetKey And Not inst.IsTemplate Then
            pres.Slides(i).Delete
            deleted = True
        End If
    Next i

    If Not deleted Then
        report = report & vbCrLf & "TARGET NOT FOUND: no non-template slide carries instance_key '" & targetKey & "'." & vbCrLf
        pres.Saved = msoTrue
        pres.Close
        HiddenFixCheck = report
        Exit Function
    End If

    report = report & vbCrLf & "Deleted the slide carrying '" & targetKey & "'." & vbCrLf & vbCrLf & _
             "--- AFTER DELETE ---" & vbCrLf & Inventory(pres)

    ' --- Re-sync, which recreates it from the template --------------------
    Dim workbookPath As String
    workbookPath = DeckRegistry.GetWorkbookPath(pres)

    Dim xl As Object, wb As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(workbookPath, 0, True)   ' read-only: sync reads the sheet

    Dim types() As String
    types = DeckRegistry.ListRegisteredTypes(pres)
    Dim lo As Long, hi As Long
    lo = LBound(types): hi = UBound(types)

    For i = lo To hi
        Dim templateSld As Object
        Dim wsName As String
        If DeckRegistry.LookupType(pres, types(i), templateSld, wsName) Then
            report = report & vbCrLf & RunSync.RunRoutineSync(wb.Worksheets(wsName), types(i), templateSld)
        End If
    Next i

    wb.Close False
    xl.Quit

    report = report & vbCrLf & "--- AFTER RE-SYNC ---" & vbCrLf & Inventory(pres)

    ' --- The assertion, stated rather than left to the reader -------------
    Dim recreatedHidden As String, templateHidden As String
    recreatedHidden = "NOT FOUND"
    templateHidden = "NOT FOUND"
    For i = 1 To pres.Slides.count
        Dim chk As SlideInstance
        chk = Resolve.ResolveSlideInstance(pres.Slides(i))
        If chk.IsTemplate Then
            templateHidden = CStr(pres.Slides(i).SlideShowTransition.Hidden <> 0)
        ElseIf chk.HasInstanceKey And chk.InstanceKey = targetKey Then
            recreatedHidden = CStr(pres.Slides(i).SlideShowTransition.Hidden <> 0)
        End If
    Next i

    report = report & vbCrLf & "VERDICT" & vbCrLf & _
        "  recreated '" & targetKey & "' hidden: " & recreatedHidden & "   (must be False)" & vbCrLf & _
        "  master template hidden:        " & templateHidden & "   (must be True)" & vbCrLf

    If saveWhenDone Then
        pres.Save
        report = report & vbCrLf & "SAVED to disk." & vbCrLf
    Else
        pres.Saved = msoTrue
        pres.Close
        report = report & vbCrLf & "Closed WITHOUT saving -- nothing written to disk." & vbCrLf
    End If

    HiddenFixCheck = report
End Function

' index / key / template? / hidden?, one line per slide.
Private Function Inventory(pres As Object) As String
    Dim s As String
    Dim i As Long
    For i = 1 To pres.Slides.count
        Dim inst As SlideInstance
        inst = Resolve.ResolveSlideInstance(pres.Slides(i))
        s = s & "  slide " & i & _
            "  key=" & IIf(inst.HasInstanceKey, inst.InstanceKey, "<none>") & _
            "  template=" & inst.IsTemplate & _
            "  hidden=" & (pres.Slides(i).SlideShowTransition.Hidden <> 0) & vbCrLf
    Next i
    Inventory = s
End Function
