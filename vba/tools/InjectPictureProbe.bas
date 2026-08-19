Attribute VB_Name = "InjectPictureProbe"
Option Explicit

' THROWAWAY DIAGNOSTIC, 2026-08-19. Application.Run cannot call
' InjectPrimitive.InjectPictureField directly from an external COM client --
' it returns a UDT (InjectResult), which VBA refuses to expose across the
' Application.Run boundary ("Sub or function not defined", a known
' misleading error for this exact case). This Sub has no return value, so
' it crosses fine, and it logs the real result to a file instead. Delete
' after use.
Public Sub RunInjectPicture(sld As Object, identityTag As String, sourceId As String, locator As String, logPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim logf As Object
    Set logf = fso.OpenTextFile(logPath, 8, True) ' 8 = ForAppending

    Dim result As InjectResult
    result = InjectPrimitive.InjectPictureField(sld, identityTag, sourceId, locator, False)

    logf.WriteLine identityTag & ": Found=" & result.Found & " WouldChange=" & result.WouldChange & _
        " Written=" & result.Written & " Verified=" & result.Verified & _
        " CurrentValue='" & result.CurrentValue & "' Error='" & result.ErrorMessage & "'"
    logf.Close
End Sub
