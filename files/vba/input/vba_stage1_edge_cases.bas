Attribute VB_Name = "Stage1EdgeCases"
Option Explicit
Option Compare Binary
Option Base 0
#Const TRACE_MODE = True

' Strings with punctuation: "a:b" and escaped quote ""ok"".
Public Enum ModeBits
    ModeOff = 0
    ModeAll = 1 Or 2
End Enum

Public Type RecordValue
    TextValue As String * 12
    Amount As Currency
    Values() As Variant
End Type

Private m_value As Variant
Private Const LIMIT As Long = &H7FFF

Public Function ExerciseAll(ByRef target As Variant, Optional ByVal count As Long = 2) As Variant
    Dim matrix(0 To 2, 1 To 3) As Double
    Dim index As Long
    Dim record As RecordValue
    ReDim record.Values(1 To 3)
    ReDim Preserve record.Values(1 To 5)
    Erase matrix
    Set target = Nothing
    Let target = Empty
    target = Null
    target = IIf(count > 0, "x:y", "z")

    If count > 10 Then
        target = "large"
    ElseIf count = 10 Then
        target = "ten"
    Else
        target = "small": target = "still small"
    End If

    For index = 0 To count Step 2
        matrix(index Mod 3, 1) = index ^ 2
    Next index
    For Each target In Array("a", "b")
        Debug.Print target
    Next target
    Do While count > 0
        count = count - 1
    Loop
    Do
        count = count + 1
    Loop Until count = 1
    While count < 3
        count = count + 1
    Wend

    Select Case count
        Case 0, 1: target = "low"
        Case Is > 10: target = "high"
        Case Else: target = "other"
    End Select

    With record
        .Amount = CCur(1.25@)
        .TextValue = "x" & vbNullString
    End With

    100: On Error GoTo Handler
    If count = 2 Then GoSub LegacyPart
    GoTo Finish
LegacyPart:
    count = count + 1
    Return
Handler:
    Resume Next
Finish:
    On Error Resume Next
    On Error GoTo 0
    ExerciseAll = target
End Function

Public Sub OpaqueAndContinuation()
    Dim text As String
    text = "continued " & _
        "across physical lines" ' trailing comment
    CallByName Application, "Run", VbMethod, text
    DoEvents
End Sub
