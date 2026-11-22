Option Explicit

Private Const LOCK_TIMEOUT_SECONDS As Long = 30

Public Sub ReconcileInventory(ByVal inventorySheet As Worksheet, ByVal movementSheet As Worksheet)
    Dim inventory As Object
    Dim movementTotals As Object
    Dim lastInventoryRow As Long
    Dim lastMovementRow As Long
    Dim rowNumber As Long
    Dim itemCode As String
    Dim quantity As Double
    Dim transactionDate As Date
    Dim previousEvents As Boolean
    Dim startedAt As Double
    Dim elapsedSeconds As Double
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo HandleError

    previousEvents = Application.EnableEvents
    Application.EnableEvents = False
    startedAt = Timer
    Set inventory = CreateObject("Scripting.Dictionary")
    Set movementTotals = CreateObject("Scripting.Dictionary")

    lastInventoryRow = inventorySheet.Cells(inventorySheet.Rows.Count, "A").End(xlUp).Row
    lastMovementRow = movementSheet.Cells(movementSheet.Rows.Count, "A").End(xlUp).Row

    For rowNumber = 2 To lastInventoryRow
        itemCode = UCase$(Trim$(CStr(inventorySheet.Cells(rowNumber, "A").Value2)))
        If Len(itemCode) > 0 Then
            If IsNumeric(inventorySheet.Cells(rowNumber, "B").Value2) Then
                If inventory.Exists(itemCode) Then
                    Err.Raise vbObjectError + 4001, "ReconcileInventory", "Duplicate inventory item: " & itemCode
                End If
                inventory.Add itemCode, CDbl(inventorySheet.Cells(rowNumber, "B").Value2)
            Else
                inventorySheet.Cells(rowNumber, "B").Interior.Color = vbRed
            End If
        End If
    Next rowNumber

    For rowNumber = 2 To lastMovementRow
        itemCode = UCase$(Trim$(CStr(movementSheet.Cells(rowNumber, "A").Value2)))
        If IsDate(movementSheet.Cells(rowNumber, "B").Value2) _
            And IsNumeric(movementSheet.Cells(rowNumber, "C").Value2) Then
            transactionDate = CDate(movementSheet.Cells(rowNumber, "B").Value2)
            quantity = CDbl(movementSheet.Cells(rowNumber, "C").Value2)

            If DateDiff("d", transactionDate, Date) <= 30 And transactionDate <= Date Then
                If movementTotals.Exists(itemCode) Then
                    movementTotals(itemCode) = movementTotals(itemCode) + quantity
                Else
                    movementTotals.Add itemCode, quantity
                End If
            End If
        End If
    Next rowNumber

    For rowNumber = 2 To lastInventoryRow
        itemCode = UCase$(Trim$(CStr(inventorySheet.Cells(rowNumber, "A").Value2)))
        If inventory.Exists(itemCode) Then
            quantity = inventory(itemCode)
            If movementTotals.Exists(itemCode) Then
                quantity = quantity + movementTotals(itemCode)
            End If

            Select Case True
                Case quantity < 0
                    inventorySheet.Cells(rowNumber, "C").Value = "INVESTIGATE"
                Case quantity = 0
                    inventorySheet.Cells(rowNumber, "C").Value = "OUT OF STOCK"
                Case quantity < 5
                    inventorySheet.Cells(rowNumber, "C").Value = "REORDER"
                Case Else
                    inventorySheet.Cells(rowNumber, "C").Value = "OK"
            End Select
            inventorySheet.Cells(rowNumber, "D").Value = quantity
        End If

        elapsedSeconds = Timer - startedAt
        If elapsedSeconds < 0 Then elapsedSeconds = elapsedSeconds + 86400
        If elapsedSeconds > LOCK_TIMEOUT_SECONDS Then
            Err.Raise vbObjectError + 4002, "ReconcileInventory", "Reconciliation exceeded its time limit."
        End If
    Next rowNumber

    inventorySheet.Range("A1:D" & lastInventoryRow).Sort _
        Key1:=inventorySheet.Range("A1"), Header:=xlYes
    GoTo CleanUp

HandleError:
    failureNumber = Err.Number
    failureDescription = Err.Description
    LogInventoryError failureNumber, failureDescription, itemCode
    MsgBox "Inventory reconciliation failed: " & failureDescription, vbExclamation

CleanUp:
    Application.EnableEvents = previousEvents
    Set movementTotals = Nothing
    Set inventory = Nothing
End Sub

Public Sub LogInventoryError(ByVal errorNumber As Long, ByVal description As String, ByVal itemCode As String)
    Dim logSheet As Worksheet
    Dim nextLogRow As Long

    On Error Resume Next
    Set logSheet = ThisWorkbook.Worksheets("Error Log")
    If logSheet Is Nothing Then Exit Sub

    nextLogRow = logSheet.Cells(logSheet.Rows.Count, "A").End(xlUp).Row + 1
    logSheet.Cells(nextLogRow, "A").Value = Now
    logSheet.Cells(nextLogRow, "B").Value = errorNumber
    logSheet.Cells(nextLogRow, "C").Value = description
    logSheet.Cells(nextLogRow, "D").Value = itemCode
End Sub
