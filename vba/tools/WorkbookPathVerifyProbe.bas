Attribute VB_Name = "WorkbookPathVerifyProbe"
Option Explicit

' THROWAWAY DIAGNOSTIC. Closes purpose-hound's one live finding: SetWorkbookPathVerified
' borrowed SetDeckPeriodVerified's retry/escalation design but never had its own named
' live proof. This gives it one, same discipline as item S/P: repeated writes, each
' confirmed from the SAVED FILE'S BYTES via WorkbookPathOnDisk, not the in-process cache,
' across genuine close/reopen cycles. Delete after use.
Public Function ProbeWorkbookPathVerified(deckPath As String) As String
    Dim report As String
    Dim pres As Object
    Dim i As Long
    ' Two rounds, matching item S's own "two genuinely separate sessions"
    ' proof standard -- a 49MB real deck makes each round genuinely slow
    ' (full SaveAs rewrite), so this stays at the minimum that still proves
    ' "survives a real close/reopen, repeatedly" rather than piling on
    ' rounds for their own sake.
    ' Must be REAL, existing files -- SetWorkbookPathVerified correctly
    ' refuses to pair with a path that doesn't exist on disk (found by this
    ' probe's first run: both non-existent placeholder paths were refused,
    ' which is the function working as intended, not a defect).
    Dim testPaths(1 To 2) As String
    testPaths(1) = "C:\Users\rohan\AppData\Local\Temp\wbpath-probe-target-1.xlsx"
    testPaths(2) = "C:\Users\rohan\AppData\Local\Temp\wbpath-probe-target-2.xlsx"

    For i = 1 To 2
        Set pres = Application.Presentations.Open(deckPath, , , msoFalse)

        Dim result As String
        result = DeckRegistry.SetWorkbookPathVerified(pres, testPaths(i), 3)

        pres.Close
        Set pres = Nothing

        ' Confirm from the SAVED FILE'S OWN BYTES, a fresh read, not the
        ' in-process object we just had open.
        Dim onDisk As String
        onDisk = DeckRegistry.WorkbookPathOnDisk(deckPath)

        Dim ok As Boolean
        ok = (StrComp(onDisk, testPaths(i), vbTextCompare) = 0)

        report = report & "Round " & i & ": wrote '" & testPaths(i) & "', " & _
            "SetWorkbookPathVerified returned '" & result & "', " & _
            "on-disk-after-close-reopen='" & onDisk & "', " & _
            IIf(ok, "MATCH", "*** MISMATCH ***") & vbCrLf
    Next i

    ProbeWorkbookPathVerified = report
End Function
