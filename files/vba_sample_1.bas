Option Explicit

Public Sub RefreshMonthlySummary()
    Dim sourceSheet As Worksheet
    Dim reportSheet As Worksheet
    Dim lastRow As Long
    Dim rowNumber As Long
    Dim reportRow As Long

    On Error GoTo HandleError

    Set sourceSheet = ThisWorkbook.Worksheets("Transactions")
    Set reportSheet = ThisWorkbook.Worksheets("Monthly Summary")
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).Row
    reportRow = 2

    For rowNumber = 2 To lastRow
        If Len(Trim$(CStr(sourceSheet.Cells(rowNumber, "A").Value))) > 0 Then
            If IsDate(sourceSheet.Cells(rowNumber, "B").Value) Then
                reportSheet.Cells(reportRow, "A").Value = sourceSheet.Cells(rowNumber, "A").Value
                reportSheet.Cells(reportRow, "B").Value = CDate(sourceSheet.Cells(rowNumber, "B").Value)
                reportSheet.Cells(reportRow, "C").Value = sourceSheet.Cells(rowNumber, "C").Value
                reportRow = reportRow + 1
            Else
                sourceSheet.Cells(rowNumber, "B").Interior.Color = vbYellow
            End If
        End If
    Next rowNumber

    reportSheet.Columns("A:C").AutoFit
    MsgBox "Summary refreshed with " & (reportRow - 2) & " records.", vbInformation
    Exit Sub

HandleError:
    MsgBox "Unable to refresh the monthly summary: " & Err.Description, vbExclamation
End Sub
