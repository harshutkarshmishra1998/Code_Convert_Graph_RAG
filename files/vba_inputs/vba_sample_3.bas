Option Explicit

Public Sub ImportDailySales(ByVal tsvPath As String, ByVal reportDate As Date)
    Dim stream As Object
    Dim salesByRegion As Object
    Dim fields As Variant
    Dim lineText As String
    Dim regionCode As String
    Dim saleAmount As Currency
    Dim lastRow As Long
    Dim rowNumber As Long

    On Error GoTo HandleError

    If Len(Dir$(tsvPath)) = 0 Then
        Err.Raise vbObjectError + 2001, "ImportDailySales", "Input file does not exist."
    End If

    Set salesByRegion = CreateObject("Scripting.Dictionary")
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile tsvPath

    If Not stream.EOS Then
        lineText = stream.ReadText(-2)
    End If

    Do While Not stream.EOS
        lineText = stream.ReadText(-2)
        fields = Split(lineText, vbTab)

        If UBound(fields) >= 2 Then
            regionCode = UCase$(Trim$(CStr(fields(1))))
            If IsNumeric(fields(2)) And Len(regionCode) > 0 Then
                saleAmount = CCur(fields(2))
                If salesByRegion.Exists(regionCode) Then
                    salesByRegion(regionCode) = salesByRegion(regionCode) + saleAmount
                Else
                    salesByRegion.Add regionCode, saleAmount
                End If
            End If
        End If
    Loop

    stream.Close
    Set stream = Nothing

    With ThisWorkbook.Worksheets("Daily Sales")
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        For rowNumber = 2 To lastRow
            If IsDate(.Cells(rowNumber, "A").Value) Then
                If DateValue(.Cells(rowNumber, "A").Value) = DateValue(reportDate) Then
                    regionCode = UCase$(Trim$(CStr(.Cells(rowNumber, "B").Value)))
                    If salesByRegion.Exists(regionCode) Then
                        .Cells(rowNumber, "C").Value = salesByRegion(regionCode)
                    Else
                        .Cells(rowNumber, "C").Value = 0
                    End If
                End If
            End If
        Next rowNumber
    End With

    Exit Sub

HandleError:
    If Not stream Is Nothing Then
        If stream.State = 1 Then stream.Close
    End If
    Err.Raise Err.Number, "ImportDailySales", Err.Description
End Sub
