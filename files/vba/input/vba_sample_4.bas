Option Explicit

Public Sub ConsolidateDepartmentFiles(ByVal inputFolder As String, ByVal outputPath As String)
    Dim fileName As String
    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim outputBook As Workbook
    Dim fileSystem As Object
    Dim sourceRange As Range
    Dim nextRow As Long
    Dim lastRow As Long
    Dim processedFiles As Long
    Dim savedCalculation As XlCalculation
    Dim savedEvents As Boolean
    Dim savedScreenUpdating As Boolean

    On Error GoTo HandleError

    Err.Clear
    savedCalculation = Application.Calculation
    savedEvents = Application.EnableEvents
    savedScreenUpdating = Application.ScreenUpdating

    If LCase$(Right$(outputPath, 5)) <> ".xlsx" Then
        Err.Raise vbObjectError + 3002, "ConsolidateDepartmentFiles", "Output path must use the .xlsx extension."
    End If
    If StrComp(outputPath, ThisWorkbook.FullName, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 3003, "ConsolidateDepartmentFiles", "Output cannot replace the macro workbook."
    End If
    If Right$(inputFolder, 1) <> Application.PathSeparator Then
        inputFolder = inputFolder & Application.PathSeparator
    End If
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FolderExists(inputFolder) Then
        Err.Raise vbObjectError + 3001, "ConsolidateDepartmentFiles", "Input folder is unavailable."
    End If
    Set fileSystem = Nothing

    Set targetSheet = ThisWorkbook.Worksheets("Consolidated")
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    targetSheet.Cells.ClearContents
    nextRow = 1
    fileName = Dir$(inputFolder & "*.xlsx")

    Do While Len(fileName) > 0
        If StrComp(inputFolder & fileName, ThisWorkbook.FullName, vbTextCompare) <> 0 Then
            Set sourceBook = Workbooks.Open(Filename:=inputFolder & fileName, ReadOnly:=True)
            Set sourceSheet = sourceBook.Worksheets(1)
            lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).Row

            If lastRow >= 1 Then
                Set sourceRange = sourceSheet.Range("A1:F" & lastRow)
                If nextRow = 1 Then
                    targetSheet.Cells(nextRow, 1).Resize(sourceRange.Rows.Count, sourceRange.Columns.Count).Value = sourceRange.Value
                    nextRow = nextRow + sourceRange.Rows.Count
                ElseIf sourceRange.Rows.Count > 1 Then
                    targetSheet.Cells(nextRow, 1).Resize(sourceRange.Rows.Count - 1, sourceRange.Columns.Count).Value = _
                        sourceRange.Offset(1, 0).Resize(sourceRange.Rows.Count - 1, sourceRange.Columns.Count).Value
                    nextRow = nextRow + sourceRange.Rows.Count - 1
                End If
            End If

            sourceBook.Close SaveChanges:=False
            Set sourceBook = Nothing
            processedFiles = processedFiles + 1
        End If
        fileName = Dir$()
    Loop

    targetSheet.Columns("A:F").AutoFit
    Set outputBook = Workbooks.Add(xlWBATWorksheet)
    outputBook.Worksheets(1).Name = "Consolidated"
    If nextRow > 1 Then
        outputBook.Worksheets(1).Range("A1:F" & nextRow - 1).Value = _
            targetSheet.Range("A1:F" & nextRow - 1).Value
    End If
    outputBook.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbook
    outputBook.Close SaveChanges:=False
    Set outputBook = Nothing
    GoTo CleanUp

HandleError:
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    If Not outputBook Is Nothing Then outputBook.Close SaveChanges:=False
    MsgBox "Consolidation failed for '" & fileName & "': " & Err.Description, vbCritical

CleanUp:
    Application.Calculation = savedCalculation
    Application.EnableEvents = savedEvents
    Application.ScreenUpdating = savedScreenUpdating
    If Err.Number = 0 Then
        MsgBox "Consolidated " & processedFiles & " department files.", vbInformation
    End If
End Sub
