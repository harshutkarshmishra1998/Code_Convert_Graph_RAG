Attribute VB_Name = "ComplexConversionStressTest"
Option Explicit
Option Compare Text
Option Base 1

' =============================================================================
' Module:      ComplexConversionStressTest
' Purpose:     Stress-test VBA-to-modern-language code converters with legacy
'              idioms, late/early binding, UDTs, error jumps, and recursion.
' Compatibility: Standard VBA (.bas) - no form/class dependencies required.
' =============================================================================

Public Enum PipelineModeFlags
    ModeNone = 0
    ModeNormalize = 1
    ModeDeduplicate = 2
    ModeAggregate = 4
    ModeAudit = 8
    ModeAll = ModeNormalize Or ModeDeduplicate Or ModeAggregate Or ModeAudit
End Enum

Public Type LedgerEntry
    EntryId As Long
    AccountCode As String
    Amount As Currency
    PostedDate As Date
    Tags() As String
    IsReconciled As Boolean
End Type

Public Type BatchSummary
    BatchId As String
    ModeMask As Long
    RecordCount As Long
    TotalAmount As Currency
    Entries() As LedgerEntry
End Type

Private Type InternalGraphNode
    NodeKey As String
    Weight As Double
    Visited As Boolean
    Children() As String
End Type

Private m_GlobalSequence As Long
Private m_AuditLog As Collection
Private m_NodeRegistry As Object ' Late-bound Scripting.Dictionary or fallback Collection

' -----------------------------------------------------------------------------
' Public Entry Point: Orchestrates multi-stage pipeline with deliberate
' VBA constructs that challenge AST/transpiler engines.
' -----------------------------------------------------------------------------
Public Function ExecuteStressPipeline( _
    ByVal rawPayload As String, _
    ByVal modeMask As Long, _
    Optional ByRef outReport As String, _
    Optional ByVal maxIterations As Long = 1000 _
) As Boolean
    On Error GoTo ErrHandler

    Dim lines() As String
    Dim batch As BatchSummary
    Dim idx As Long
    Dim processedCount As Long
    Dim tempAmount As Currency
    Dim statusMsg As String

    If Len(Trim$(rawPayload)) = 0 Then
        outReport = "EMPTY_PAYLOAD"
        ExecuteStressPipeline = False
        Exit Function
    End If

    InitInfrastructure
    batch.BatchId = "BATCH_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(NextSequence())
    batch.ModeMask = modeMask
    batch.RecordCount = 0
    batch.TotalAmount = 0@

    lines = Split(NormalizeLineEndings(rawPayload), vbLf)
    ReDim batch.Entries(1 To UBound(lines) - LBound(lines) + 1)

    For idx = LBound(lines) To UBound(lines)
        If idx - LBound(lines) + 1 > maxIterations Then
            LogAudit "WARN: Max iteration cap reached at row " & CStr(idx)
            Exit For
        End If

        If Len(Trim$(lines(idx))) > 0 Then
            If TryParseEntry(lines(idx), idx, batch.Entries(batch.RecordCount + 1)) Then
                batch.RecordCount = batch.RecordCount + 1

                If (modeMask And ModeNormalize) = ModeNormalize Then
                    NormalizeEntryInPlace batch.Entries(batch.RecordCount)
                End If

                tempAmount = batch.Entries(batch.RecordCount).Amount
                If (modeMask And ModeAggregate) = ModeAggregate Then
                    batch.TotalAmount = batch.TotalAmount + tempAmount
                End If

                processedCount = processedCount + 1
            Else
                LogAudit "SKIP: Malformed row " & CStr(idx) & " -> " & Left$(lines(idx), 60)
            End If
        End If

        If (idx Mod 250) = 0 Then
            DoEvents
        End If
    Next idx

    If batch.RecordCount > 0 Then
        ReDim Preserve batch.Entries(1 To batch.RecordCount)
    Else
        Erase batch.Entries
    End If

    If (modeMask And ModeDeduplicate) = ModeDeduplicate And batch.RecordCount > 1 Then
        DeduplicateBatchEntries batch
    End If

    If (modeMask And ModeAudit) = ModeAudit Then
        RunLegacyGoSubAudit batch, statusMsg
    Else
        statusMsg = "AUDIT_SKIPPED"
    End If

    outReport = BuildSummaryReport(batch, processedCount, statusMsg)
    ExecuteStressPipeline = True
    Exit Function

ErrHandler:
    outReport = "PIPELINE_ERROR [" & CStr(Err.Number) & "]: " & Err.Description & " at line/step context."
    LogAudit outReport
    ExecuteStressPipeline = False
End Function

' -----------------------------------------------------------------------------
' Initializes late-bound dictionary with fallback so converters see both
' CreateObject and dynamic error recovery.
' -----------------------------------------------------------------------------
Private Sub InitInfrastructure()
    On Error Resume Next
    Set m_NodeRegistry = CreateObject("Scripting.Dictionary")
    If Err.Number <> 0 Or m_NodeRegistry Is Nothing Then
        Err.Clear
        Set m_NodeRegistry = New Collection
    End If
    On Error GoTo 0

    If m_AuditLog Is Nothing Then
        Set m_AuditLog = New Collection
    End If
End Sub

Private Function NextSequence() As Long
    m_GlobalSequence = m_GlobalSequence + 1
    NextSequence = m_GlobalSequence
End Sub

Private Sub LogAudit(ByVal message As String)
    If m_AuditLog Is Nothing Then Set m_AuditLog = New Collection
    m_AuditLog.Add Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & message
End Sub

' -----------------------------------------------------------------------------
' String normalization & Variant array manipulation.
' -----------------------------------------------------------------------------
Public Function NormalizeLineEndings(ByVal text As String) As String
    Dim result As String
    result = Replace(text, vbCrLf, vbLf)
    result = Replace(result, vbCr, vbLf)
    NormalizeLineEndings = result
End Function

Private Function TryParseEntry( _
    ByVal rawLine As String, _
    ByVal fallbackId As Long, _
    ByRef target As LedgerEntry _
) As Boolean
    On Error GoTo ParseFail

    Dim parts As Variant
    Dim tagTokens As Variant
    Dim i As Long
    Dim rawAmt As Double

    parts = Split(rawLine, "|")
    If Not IsArray(parts) Then GoTo ParseFail
    If UBound(parts) - LBound(parts) + 1 < 4 Then GoTo ParseFail

    If IsNumeric(Trim$(parts(LBound(parts)))) Then
        target.EntryId = CLng(Trim$(parts(LBound(parts))))
    Else
        target.EntryId = fallbackId
    End If

    target.AccountCode = UCase$(Trim$(CStr(parts(LBound(parts) + 1))))
    If Len(target.AccountCode) = 0 Then GoTo ParseFail

    rawAmt = CDbl(Trim$(CStr(parts(LBound(parts) + 2))))
    target.Amount = CCur(rawAmt)

    If IsDate(Trim$(CStr(parts(LBound(parts) + 3)))) Then
        target.PostedDate = CDate(Trim$(CStr(parts(LBound(parts) + 3))))
    Else
        target.PostedDate = Date
    End If

    target.IsReconciled = False

    If UBound(parts) - LBound(parts) + 1 >= 5 And Len(Trim$(CStr(parts(LBound(parts) + 4)))) > 0 Then
        tagTokens = Split(CStr(parts(LBound(parts) + 4)), ",")
        ReDim target.Tags(1 To UBound(tagTokens) - LBound(tagTokens) + 1)
        For i = LBound(tagTokens) To UBound(tagTokens)
            target.Tags(i - LBound(tagTokens) + 1) = LCase$(Trim$(CStr(tagTokens(i))))
        Next i
    Else
        ReDim target.Tags(1 To 1)
        target.Tags(1) = "untagged"
    End If

    TryParseEntry = True
    Exit Function

ParseFail:
    TryParseEntry = False
End Function

' -----------------------------------------------------------------------------
' Mutates a UDT passed ByRef, including resizing/populating nested dynamic array.
' -----------------------------------------------------------------------------
Private Sub NormalizeEntryInPlace(ByRef entry As LedgerEntry)
    Dim i As Long
    Dim hasCoreTag As Boolean

    entry.AccountCode = Replace(entry.AccountCode, " ", "_")
    entry.AccountCode = Replace(entry.AccountCode, "-", "_")

    If entry.Amount < 0@ Then
        entry.IsReconciled = False
    ElseIf entry.Amount >= 10000@ Then
        entry.IsReconciled = False
    Else
        entry.IsReconciled = True
    End If

    hasCoreTag = False
    If (Not Not entry.Tags) <> 0 Then
        For i = LBound(entry.Tags) To UBound(entry.Tags)
            entry.Tags(i) = Trim$(entry.Tags(i))
            If entry.Tags(i) = "core" Then hasCoreTag = True
        Next i
    End If

    If Not hasCoreTag Then
        If (Not Not entry.Tags) = 0 Then
            ReDim entry.Tags(1 To 1)
            entry.Tags(1) = "core"
        Else
            ReDim Preserve entry.Tags(1 To UBound(entry.Tags) + 1)
            entry.Tags(UBound(entry.Tags)) = "core"
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' Deduplicates entries by EntryId + AccountCode using a composite key.
' -----------------------------------------------------------------------------
Private Sub DeduplicateBatchEntries(ByRef batch As BatchSummary)
    Dim seenKeys As Object
    Dim kept() As LedgerEntry
    Dim keptCount As Long
    Dim i As Long
    Dim key As String
    Dim newTotal As Currency

    On Error Resume Next
    Set seenKeys = CreateObject("Scripting.Dictionary")
    On Error GoTo 0

    If seenKeys Is Nothing Then Exit Sub

    ReDim kept(1 To batch.RecordCount)
    keptCount = 0
    newTotal = 0@

    For i = LBound(batch.Entries) To UBound(batch.Entries)
        key = CStr(batch.Entries(i).EntryId) & "::" & batch.Entries(i).AccountCode
        If Not seenKeys.Exists(key) Then
            seenKeys.Add key, i
            keptCount = keptCount + 1
            kept(keptCount) = batch.Entries(i)
            newTotal = newTotal + batch.Entries(i).Amount
        Else
            LogAudit "DEDUP: Dropped duplicate key " & key
        End If
    Next i

    If keptCount > 0 Then
        ReDim Preserve kept(1 To keptCount)
        batch.Entries = kept
    Else
        Erase batch.Entries
    End If

    batch.RecordCount = keptCount
    batch.TotalAmount = newTotal
End Sub

' -----------------------------------------------------------------------------
' Classic GoSub / Return subroutine inside a procedure. Most modern target
' languages lack native GoSub, making this a prime transpiler test case.
' -----------------------------------------------------------------------------
Private Sub RunLegacyGoSubAudit(ByRef batch As BatchSummary, ByRef summaryStatus As String)
    Dim i As Long
    Dim flaggedCount As Long
    Dim currentCheckId As Long
    Dim checkResult As String

    flaggedCount = 0

    For i = LBound(batch.Entries) To UBound(batch.Entries)
        currentCheckId = batch.Entries(i).EntryId
        GoSub ValidateSingleEntry
        If checkResult <> "OK" Then
            flaggedCount = flaggedCount + 1
            LogAudit "FLAG [" & CStr(currentCheckId) & "]: " & checkResult
        End If
    Next i

    summaryStatus = "AUDITED_" & CStr(batch.RecordCount) & "_FLAGGED_" & CStr(flaggedCount)
    Exit Sub

ValidateSingleEntry:
    If batch.Entries(i).Amount = 0@ Then
        checkResult = "ZERO_AMOUNT"
    ElseIf Year(batch.Entries(i).PostedDate) < 1990 Or Year(batch.Entries(i).PostedDate) > 2099 Then
        checkResult = "OUT_OF_RANGE_DATE"
    Else
        checkResult = "OK"
    End If
    Return
End Sub

' -----------------------------------------------------------------------------
' Recursive dependency/lineage graph evaluator over delimited edge strings.
' Format of edgesSpec: "A>B,B>C,A>D,C>E"
' -----------------------------------------------------------------------------
Public Function EvaluateDependencyDepth( _
    ByVal edgesSpec As String, _
    ByVal startNode As String, _
    Optional ByVal currentDepth As Long = 0, _
    Optional ByVal maxDepth As Long = 25 _
) As Long
    Dim pairs As Variant
    Dim edgeParts As Variant
    Dim i As Long
    Dim srcNode As String
    Dim dstNode As String
    Dim childDepth As Long
    Dim maxChildDepth As Long

    If currentDepth >= maxDepth Then
        EvaluateDependencyDepth = currentDepth
        Exit Function
    End If

    If Len(Trim$(edgesSpec)) = 0 Or Len(Trim$(startNode)) = 0 Then
        EvaluateDependencyDepth = currentDepth
        Exit Function
    End If

    pairs = Split(edgesSpec, ",")
    maxChildDepth = currentDepth

    For i = LBound(pairs) To UBound(pairs)
        edgeParts = Split(CStr(pairs(i)), ">")
        If UBound(edgeParts) - LBound(edgeParts) + 1 = 2 Then
            srcNode = Trim$(CStr(edgeParts(LBound(edgeParts))))
            dstNode = Trim$(CStr(edgeParts(LBound(edgeParts) + 1)))

            If StrComp(srcNode, Trim$(startNode), vbTextCompare) = 0 And Len(dstNode) > 0 Then
                If StrComp(dstNode, srcNode, vbTextCompare) <> 0 Then
                    childDepth = EvaluateDependencyDepth(edgesSpec, dstNode, currentDepth + 1, maxDepth)
                    If childDepth > maxChildDepth Then
                        maxChildDepth = childDepth
                    End If
                End If
            End If
        End If
    Next i

    EvaluateDependencyDepth = maxChildDepth
End Function

' -----------------------------------------------------------------------------
' Parameterized-style SQL builder with basic identifier/value escaping.
' -----------------------------------------------------------------------------
Public Function BuildUpsertSql( _
    ByVal tableName As String, _
    ByRef entry As LedgerEntry _
) As String
    Dim cleanTable As String
    Dim cleanAcct As String
    Dim cleanDate As String
    Dim recFlag As Long

    cleanTable = SanitizeSqlIdentifier(tableName)
    cleanAcct = Replace(entry.AccountCode, "'", "''")
    cleanDate = Format$(entry.PostedDate, "yyyy-mm-dd")
    recFlag = IIf(entry.IsReconciled, 1, 0)

    BuildUpsertSql = _
        "MERGE INTO " & cleanTable & " AS tgt " & _
        "USING (SELECT " & CStr(entry.EntryId) & " AS entry_id, '" & cleanAcct & "' AS account_code) AS src " & _
        "ON tgt.entry_id = src.entry_id AND tgt.account_code = src.account_code " & _
        "WHEN MATCHED THEN UPDATE SET amount = " & Replace(CStr(entry.Amount), ",", ".") & _
        ", posted_date = DATE '" & cleanDate & "', is_reconciled = " & CStr(recFlag) & " " & _
        "WHEN NOT MATCHED THEN INSERT (entry_id, account_code, amount, posted_date, is_reconciled) " & _
        "VALUES (" & CStr(entry.EntryId) & ", '" & cleanAcct & "', " & _
        Replace(CStr(entry.Amount), ",", ".") & ", DATE '" & cleanDate & "', " & CStr(recFlag) & ");"
End Function

Private Function SanitizeSqlIdentifier(ByVal ident As String) As String
    Dim i As Long
    Dim ch As String
    Dim outStr As String

    outStr = ""
    For i = 1 To Len(ident)
        ch = Mid$(ident, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or _
           (ch >= "0" And ch <= "9") Or ch = "_" Or ch = "." Then
            outStr = outStr & ch
        End If
    Next i

    If Len(outStr) = 0 Then outStr = "default_ledger_table"
    SanitizeSqlIdentifier = outStr
End Function

' -----------------------------------------------------------------------------
' Report & audit log formatter.
' -----------------------------------------------------------------------------
Private Function BuildSummaryReport( _
    ByRef batch As BatchSummary, _
    ByVal parsedCount As Long, _
    ByVal statusMsg As String _
) As String
    Dim sb As String
    Dim i As Long
    Dim logLimit As Long

    sb = "=== BATCH SUMMARY ===" & vbCrLf
    sb = sb & "Batch ID:      " & batch.BatchId & vbCrLf
    sb = sb & "Mode Mask:     " & CStr(batch.ModeMask) & vbCrLf
    sb = sb & "Parsed Rows:   " & CStr(parsedCount) & vbCrLf
    sb = sb & "Final Rows:    " & CStr(batch.RecordCount) & vbCrLf
    sb = sb & "Total Amount:  " & Format$(batch.TotalAmount, "#,##0.00") & vbCrLf
    sb = sb & "Audit Status:  " & statusMsg & vbCrLf

    If Not m_AuditLog Is Nothing And m_AuditLog.Count > 0 Then
        sb = sb & "--- RECENT AUDIT LOGS ---" & vbCrLf
        logLimit = m_AuditLog.Count
        If logLimit > 15 Then logLimit = 15
        For i = 1 To logLimit
            sb = sb & CStr(m_AuditLog(i)) & vbCrLf
        Next i
    End If

    BuildSummaryReport = sb
End Function

Public Function GetAuditEntryCount() As Long
    If m_AuditLog Is Nothing Then
        GetAuditEntryCount = 0
    Else
        GetAuditEntryCount = m_AuditLog.Count
    End If
End Function

Public Sub ResetModuleState()
    m_GlobalSequence = 0
    Set m_AuditLog = New Collection
    Set m_NodeRegistry = Nothing
End Sub