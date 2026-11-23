Option Explicit

Public Function NormalizeSku(ByVal rawSku As Variant) As String
    Dim normalized As String

    On Error GoTo InvalidSku

    If IsNull(rawSku) Or IsEmpty(rawSku) Then
        Err.Raise vbObjectError + 1001, "NormalizeSku", "SKU is required."
    End If

    If VarType(rawSku) <> vbString Then
        Err.Raise vbObjectError + 1002, "NormalizeSku", "SKU must be text."
    End If

    normalized = UCase$(Trim$(CStr(rawSku)))
    If Len(normalized) < 4 Or Len(normalized) > 24 Then
        Err.Raise vbObjectError + 1003, "NormalizeSku", "SKU length is invalid."
    End If

    Select Case True
        Case normalized Like "*[!A-Z0-9-]*"
            Err.Raise vbObjectError + 1004, "NormalizeSku", "SKU contains unsupported characters."
        Case Else
            NormalizeSku = normalized
    End Select

    Exit Function

InvalidSku:
    Err.Raise Err.Number, "NormalizeSku", Err.Description
End Function
