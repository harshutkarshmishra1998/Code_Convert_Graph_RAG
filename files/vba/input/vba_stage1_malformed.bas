Option Explicit

Public Sub BrokenButRecoverable()
    Dim value As String
    value = "this string never closes
    Debug.Print value
End Sub
