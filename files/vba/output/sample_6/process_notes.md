# VBA Process Notes: `vba_sample_6.bas`

## File details

- Notes format version: `2`
- Source path: `/Users/harshutkarshmishra/Downloads/VS Code/Code_Convert_Graph_RAG/files/vba/input/vba_sample_6.bas`
- SHA-256: `7289ca3eb6d951118004db2ff7c58f5f44b6b8ee6a818a3efa5490dea9542736`
- Decoded as: `utf-8`
- Physical lines: 489
- Procedures found: 15

## What this file contains

The sections below list source statements in their original order. Labels are mechanical descriptions based on statement shape; they do not claim to know business intent. Exact VBA text and line numbers are included so each note can be checked against the source.

## Module-level declarations and settings

- Lines 1–1 — **Module attribute**: `Attribute VB_Name = "ComplexConversionStressTest"`
- Lines 2–2 — **Module option**: `Option Explicit`
- Lines 3–3 — **Module option**: `Option Compare Text`
- Lines 4–4 — **Module option**: `Option Base 1`
- Lines 6–6 — **Comment**: `' =============================================================================`
- Lines 7–7 — **Comment**: `' Module:      ComplexConversionStressTest`
- Lines 8–8 — **Comment**: `' Purpose:     Stress-test VBA-to-modern-language code converters with legacy`
- Lines 9–9 — **Comment**: `'              idioms, late/early binding, UDTs, error jumps, and recursion.`
- Lines 10–10 — **Comment**: `' Compatibility: Standard VBA (.bas) - no form/class dependencies required.`
- Lines 11–11 — **Comment**: `' =============================================================================`
- Lines 13–13 — **Enum declaration**: `Public Enum PipelineModeFlags`
- Lines 14–14 — **Enum member**: `ModeNone = 0`
- Lines 15–15 — **Enum member**: `ModeNormalize = 1`
- Lines 16–16 — **Enum member**: `ModeDeduplicate = 2`
- Lines 17–17 — **Enum member**: `ModeAggregate = 4`
- Lines 18–18 — **Enum member**: `ModeAudit = 8`
- Lines 19–19 — **Enum member**: `ModeAll = ModeNormalize Or ModeDeduplicate Or ModeAggregate Or ModeAudit`
- Lines 20–20 — **End enum**: `End Enum`
- Lines 22–22 — **User-defined type declaration**: `Public Type LedgerEntry`
- Lines 23–23 — **User-defined type member**: `EntryId As Long`
- Lines 24–24 — **User-defined type member**: `AccountCode As String`
- Lines 25–25 — **User-defined type member**: `Amount As Currency`
- Lines 26–26 — **User-defined type member**: `PostedDate As Date`
- Lines 27–27 — **User-defined type member**: `Tags() As String`
- Lines 28–28 — **User-defined type member**: `IsReconciled As Boolean`
- Lines 29–29 — **End user-defined type**: `End Type`
- Lines 31–31 — **User-defined type declaration**: `Public Type BatchSummary`
- Lines 32–32 — **User-defined type member**: `BatchId As String`
- Lines 33–33 — **User-defined type member**: `ModeMask As Long`
- Lines 34–34 — **User-defined type member**: `RecordCount As Long`
- Lines 35–35 — **User-defined type member**: `TotalAmount As Currency`
- Lines 36–36 — **User-defined type member**: `Entries() As LedgerEntry`
- Lines 37–37 — **End user-defined type**: `End Type`
- Lines 39–39 — **User-defined type declaration**: `Private Type InternalGraphNode`
- Lines 40–40 — **User-defined type member**: `NodeKey As String`
- Lines 41–41 — **User-defined type member**: `Weight As Double`
- Lines 42–42 — **User-defined type member**: `Visited As Boolean`
- Lines 43–43 — **User-defined type member**: `Children() As String`
- Lines 44–44 — **End user-defined type**: `End Type`
- Lines 46–46 — **Declaration / storage**: `Private m_GlobalSequence As Long`
- Lines 47–47 — **Declaration / storage**: `Private m_AuditLog As Collection`
- Lines 48–48 — **Declaration / storage**: `Private m_NodeRegistry As Object ' Late-bound Scripting.Dictionary or fallback Collection`
- Lines 50–50 — **Comment**: `' -----------------------------------------------------------------------------`
- Lines 51–51 — **Comment**: `' Public Entry Point: Orchestrates multi-stage pipeline with deliberate`
- Lines 52–52 — **Comment**: `' VBA constructs that challenge AST/transpiler engines.`
- Lines 53–53 — **Comment**: `' -----------------------------------------------------------------------------`
- Lines 140–140 — **Comment**: `' -----------------------------------------------------------------------------`
- Lines 141–141 — **Comment**: `' Initializes late-bound dictionary with fallback so converters see both`
- Lines 142–142 — **Comment**: `' CreateObject and dynamic error recovery.`
- Lines 143–143 — **Comment**: `' -----------------------------------------------------------------------------`
- Lines 168–168 — **Comment**: `' -----------------------------------------------------------------------------`
- Lines 169–169 — **Comment**: `' String normalization & Variant array manipulation.`
- Lines 170–170 — **Comment**: `' -----------------------------------------------------------------------------`
- Lines 232–232 — **Comment**: `' -----------------------------------------------------------------------------`
- Lines 233–233 — **Comment**: `' Mutates a UDT passed ByRef, including resizing/populating nested dynamic array.`
- Lines 234–234 — **Comment**: `' -----------------------------------------------------------------------------`
- Lines 269–269 — **Comment**: `' -----------------------------------------------------------------------------`
- Lines 270–270 — **Comment**: `' Deduplicates entries by EntryId + AccountCode using a composite key.`
- Lines 271–271 — **Comment**: `' -----------------------------------------------------------------------------`
- Lines 313–313 — **Comment**: `' -----------------------------------------------------------------------------`
- Lines 314–314 — **Comment**: `' Classic GoSub / Return subroutine inside a procedure. Most modern target`
- Lines 315–315 — **Comment**: `' languages lack native GoSub, making this a prime transpiler test case.`
- Lines 316–316 — **Comment**: `' -----------------------------------------------------------------------------`
- Lines 348–348 — **Comment**: `' -----------------------------------------------------------------------------`
- Lines 349–349 — **Comment**: `' Recursive dependency/lineage graph evaluator over delimited edge strings.`
- Lines 350–350 — **Comment**: `' Format of edgesSpec: "A>B,B>C,A>D,C>E"`
- Lines 351–351 — **Comment**: `' -----------------------------------------------------------------------------`
- Lines 399–399 — **Comment**: `' -----------------------------------------------------------------------------`
- Lines 400–400 — **Comment**: `' Parameterized-style SQL builder with basic identifier/value escaping.`
- Lines 401–401 — **Comment**: `' -----------------------------------------------------------------------------`
- Lines 445–445 — **Comment**: `' -----------------------------------------------------------------------------`
- Lines 446–446 — **Comment**: `' Report & audit log formatter.`
- Lines 447–447 — **Comment**: `' -----------------------------------------------------------------------------`

## Function: `ExecuteStressPipeline` (lines 54–138)

**Declaration:** `Public Function ExecuteStressPipeline( _ ByVal rawPayload As String, _ ByVal modeMask As Long, _ Optional ByRef outReport As String, _ Optional ByVal maxIterations As Long = 1000 _ ) As Boolean`

### Ordered process notes

- **Lines 60–60 · Error handling:** Changes VBA error handling or refers to the active error; preserve this behavior during conversion.
  - VBA: `On Error GoTo ErrHandler`
- **Lines 62–62 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim lines() As String`
- **Lines 63–63 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim batch As BatchSummary`
- **Lines 64–64 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim idx As Long`
- **Lines 65–65 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim processedCount As Long`
- **Lines 66–66 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim tempAmount As Currency`
- **Lines 67–67 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim statusMsg As String`
- **Lines 69–69 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If Len(Trim$(rawPayload)) = 0 Then`
- **Lines 70–70 · Calculation / assignment:** Assign `outReport` using the exact expression `"EMPTY_PAYLOAD"`.
  - VBA: `outReport = "EMPTY_PAYLOAD"`
- **Lines 71–71 · Calculation / assignment:** Assign `ExecuteStressPipeline` using the exact expression `False`.
  - VBA: `ExecuteStressPipeline = False`
- **Lines 72–72 · Jump / procedure flow:** Changes procedure flow; preserve the target and execution order during conversion.
  - VBA: `Exit Function`
- **Lines 73–73 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 75–75 · Call / action:** Calls a procedure, property, or host action; identify its dependency and side effects.
  - VBA: `InitInfrastructure`
- **Lines 76–76 · Calculation / assignment:** Assign `batch.BatchId` using the exact expression `"BATCH_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(NextSequence())`.
  - VBA: `batch.BatchId = "BATCH_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(NextSequence())`
- **Lines 77–77 · Calculation / assignment:** Assign `batch.ModeMask` using the exact expression `modeMask`.
  - VBA: `batch.ModeMask = modeMask`
- **Lines 78–78 · Calculation / assignment:** Assign `batch.RecordCount` using the exact expression `0`.
  - VBA: `batch.RecordCount = 0`
- **Lines 79–79 · Calculation / assignment:** Assign `batch.TotalAmount` using the exact expression `0@`.
  - VBA: `batch.TotalAmount = 0@`
- **Lines 81–81 · Calculation / assignment:** Assign `lines` using the exact expression `Split(NormalizeLineEndings(rawPayload), vbLf)`.
  - VBA: `lines = Split(NormalizeLineEndings(rawPayload), vbLf)`
- **Lines 82–82 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `ReDim batch.Entries(1 To UBound(lines) - LBound(lines) + 1)`
- **Lines 84–84 · Loop control:** This line starts, controls, or ends repeated work; bounds and conditions are preserved below.
  - VBA: `For idx = LBound(lines) To UBound(lines)`
- **Lines 85–85 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If idx - LBound(lines) + 1 > maxIterations Then`
- **Lines 86–86 · Call / action:** Calls a procedure, property, or host action; identify its dependency and side effects.
  - VBA: `LogAudit "WARN: Max iteration cap reached at row " & CStr(idx)`
- **Lines 87–87 · Jump / procedure flow:** Changes procedure flow; preserve the target and execution order during conversion.
  - VBA: `Exit For`
- **Lines 88–88 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 90–90 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If Len(Trim$(lines(idx))) > 0 Then`
- **Lines 91–91 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If TryParseEntry(lines(idx), idx, batch.Entries(batch.RecordCount + 1)) Then`
- **Lines 92–92 · Calculation / assignment:** Assign `batch.RecordCount` using the exact expression `batch.RecordCount + 1`.
  - VBA: `batch.RecordCount = batch.RecordCount + 1`
- **Lines 94–94 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If (modeMask And ModeNormalize) = ModeNormalize Then`
- **Lines 95–95 · Call / action:** Calls a procedure, property, or host action; identify its dependency and side effects.
  - VBA: `NormalizeEntryInPlace batch.Entries(batch.RecordCount)`
- **Lines 96–96 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 98–98 · Calculation / assignment:** Assign `tempAmount` using the exact expression `batch.Entries(batch.RecordCount).Amount`.
  - VBA: `tempAmount = batch.Entries(batch.RecordCount).Amount`
- **Lines 99–99 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If (modeMask And ModeAggregate) = ModeAggregate Then`
- **Lines 100–100 · Calculation / assignment:** Assign `batch.TotalAmount` using the exact expression `batch.TotalAmount + tempAmount`.
  - VBA: `batch.TotalAmount = batch.TotalAmount + tempAmount`
- **Lines 101–101 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 103–103 · Calculation / assignment:** Assign `processedCount` using the exact expression `processedCount + 1`.
  - VBA: `processedCount = processedCount + 1`
- **Lines 104–104 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `Else`
- **Lines 105–105 · Call / action:** Calls a procedure, property, or host action; identify its dependency and side effects.
  - VBA: `LogAudit "SKIP: Malformed row " & CStr(idx) & " -> " & Left$(lines(idx), 60)`
- **Lines 106–106 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 107–107 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 109–109 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If (idx Mod 250) = 0 Then`
- **Lines 110–110 · Call / action:** Calls a procedure, property, or host action; identify its dependency and side effects.
  - VBA: `DoEvents`
- **Lines 111–111 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 112–112 · Loop control:** This line starts, controls, or ends repeated work; bounds and conditions are preserved below.
  - VBA: `Next idx`
- **Lines 114–114 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If batch.RecordCount > 0 Then`
- **Lines 115–115 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `ReDim Preserve batch.Entries(1 To batch.RecordCount)`
- **Lines 116–116 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `Else`
- **Lines 117–117 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Erase batch.Entries`
- **Lines 118–118 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 120–120 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If (modeMask And ModeDeduplicate) = ModeDeduplicate And batch.RecordCount > 1 Then`
- **Lines 121–121 · Call / action:** Calls a procedure, property, or host action; identify its dependency and side effects.
  - VBA: `DeduplicateBatchEntries batch`
- **Lines 122–122 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 124–124 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If (modeMask And ModeAudit) = ModeAudit Then`
- **Lines 125–125 · Call / action:** Calls a procedure, property, or host action; identify its dependency and side effects.
  - VBA: `RunLegacyGoSubAudit batch, statusMsg`
- **Lines 126–126 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `Else`
- **Lines 127–127 · Calculation / assignment:** Assign `statusMsg` using the exact expression `"AUDIT_SKIPPED"`.
  - VBA: `statusMsg = "AUDIT_SKIPPED"`
- **Lines 128–128 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 130–130 · Calculation / assignment:** Assign `outReport` using the exact expression `BuildSummaryReport(batch, processedCount, statusMsg)`.
  - VBA: `outReport = BuildSummaryReport(batch, processedCount, statusMsg)`
- **Lines 131–131 · Calculation / assignment:** Assign `ExecuteStressPipeline` using the exact expression `True`.
  - VBA: `ExecuteStressPipeline = True`
- **Lines 132–132 · Jump / procedure flow:** Changes procedure flow; preserve the target and execution order during conversion.
  - VBA: `Exit Function`
- **Lines 134–134 · Label:** Label.
  - VBA: `ErrHandler:`
- **Lines 135–135 · Calculation / assignment:** Assign `outReport` using the exact expression `"PIPELINE_ERROR [" & CStr(Err.Number) & "]: " & Err.Description & " at line/step context."`.
  - VBA: `outReport = "PIPELINE_ERROR [" & CStr(Err.Number) & "]: " & Err.Description & " at line/step context."`
- **Lines 136–136 · Call / action:** Calls a procedure, property, or host action; identify its dependency and side effects.
  - VBA: `LogAudit outReport`
- **Lines 137–137 · Calculation / assignment:** Assign `ExecuteStressPipeline` using the exact expression `False`.
  - VBA: `ExecuteStressPipeline = False`
- **Lines 138–138 · Procedure end:** Procedure end.
  - VBA: `End Function`

## Sub: `InitInfrastructure` (lines 144–156)

**Declaration:** `Private Sub InitInfrastructure()`

### Ordered process notes

- **Lines 145–145 · Error handling:** Changes VBA error handling or refers to the active error; preserve this behavior during conversion.
  - VBA: `On Error Resume Next`
- **Lines 146–146 · Calculation / assignment:** Assign `Set m_NodeRegistry` using the exact expression `CreateObject("Scripting.Dictionary")`.
  - VBA: `Set m_NodeRegistry = CreateObject("Scripting.Dictionary")`
- **Lines 147–147 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If Err.Number <> 0 Or m_NodeRegistry Is Nothing Then`
- **Lines 148–148 · Error handling:** Changes VBA error handling or refers to the active error; preserve this behavior during conversion.
  - VBA: `Err.Clear`
- **Lines 149–149 · Calculation / assignment:** Assign `Set m_NodeRegistry` using the exact expression `New Collection`.
  - VBA: `Set m_NodeRegistry = New Collection`
- **Lines 150–150 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 151–151 · Error handling:** Changes VBA error handling or refers to the active error; preserve this behavior during conversion.
  - VBA: `On Error GoTo 0`
- **Lines 153–153 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If m_AuditLog Is Nothing Then`
- **Lines 154–154 · Calculation / assignment:** Assign `Set m_AuditLog` using the exact expression `New Collection`.
  - VBA: `Set m_AuditLog = New Collection`
- **Lines 155–155 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 156–156 · Procedure end:** Procedure end.
  - VBA: `End Sub`

## Function: `NextSequence` (lines 158–161)

**Declaration:** `Private Function NextSequence() As Long`

### Ordered process notes

- **Lines 159–159 · Calculation / assignment:** Assign `m_GlobalSequence` using the exact expression `m_GlobalSequence + 1`.
  - VBA: `m_GlobalSequence = m_GlobalSequence + 1`
- **Lines 160–160 · Calculation / assignment:** Assign `NextSequence` using the exact expression `m_GlobalSequence`.
  - VBA: `NextSequence = m_GlobalSequence`
- **Lines 161–161 · Procedure end:** Procedure end.
  - VBA: `End Sub`

## Sub: `LogAudit` (lines 163–166)

**Declaration:** `Private Sub LogAudit(ByVal message As String)`

### Ordered process notes

- **Lines 164–164 · Decision with inline action (review):** This line has a decision and an inline action; both remain in the exact source statement and need review.
  - VBA: `If m_AuditLog Is Nothing Then Set m_AuditLog = New Collection`
- **Lines 165–165 · Call / action:** Calls a procedure, property, or host action; identify its dependency and side effects.
  - VBA: `m_AuditLog.Add Format$(Now, "yyyy-mm-dd hh:nn:ss") & " \| " & message`
- **Lines 166–166 · Procedure end:** Procedure end.
  - VBA: `End Sub`

## Function: `NormalizeLineEndings` (lines 171–176)

**Declaration:** `Public Function NormalizeLineEndings(ByVal text As String) As String`

### Ordered process notes

- **Lines 172–172 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim result As String`
- **Lines 173–173 · Calculation / assignment:** Assign `result` using the exact expression `Replace(text, vbCrLf, vbLf)`.
  - VBA: `result = Replace(text, vbCrLf, vbLf)`
- **Lines 174–174 · Calculation / assignment:** Assign `result` using the exact expression `Replace(result, vbCr, vbLf)`.
  - VBA: `result = Replace(result, vbCr, vbLf)`
- **Lines 175–175 · Calculation / assignment:** Assign `NormalizeLineEndings` using the exact expression `result`.
  - VBA: `NormalizeLineEndings = result`
- **Lines 176–176 · Procedure end:** Procedure end.
  - VBA: `End Function`

## Function: `TryParseEntry` (lines 178–230)

**Declaration:** `Private Function TryParseEntry( _ ByVal rawLine As String, _ ByVal fallbackId As Long, _ ByRef target As LedgerEntry _ ) As Boolean`

### Ordered process notes

- **Lines 183–183 · Error handling:** Changes VBA error handling or refers to the active error; preserve this behavior during conversion.
  - VBA: `On Error GoTo ParseFail`
- **Lines 185–185 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim parts As Variant`
- **Lines 186–186 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim tagTokens As Variant`
- **Lines 187–187 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim i As Long`
- **Lines 188–188 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim rawAmt As Double`
- **Lines 190–190 · Calculation / assignment:** Assign `parts` using the exact expression `Split(rawLine, "\|")`.
  - VBA: `parts = Split(rawLine, "\|")`
- **Lines 191–191 · Decision with inline action (review):** This line has a decision and an inline action; both remain in the exact source statement and need review.
  - VBA: `If Not IsArray(parts) Then GoTo ParseFail`
- **Lines 192–192 · Decision with inline action (review):** This line has a decision and an inline action; both remain in the exact source statement and need review.
  - VBA: `If UBound(parts) - LBound(parts) + 1 < 4 Then GoTo ParseFail`
- **Lines 194–194 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If IsNumeric(Trim$(parts(LBound(parts)))) Then`
- **Lines 195–195 · Calculation / assignment:** Assign `target.EntryId` using the exact expression `CLng(Trim$(parts(LBound(parts))))`.
  - VBA: `target.EntryId = CLng(Trim$(parts(LBound(parts))))`
- **Lines 196–196 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `Else`
- **Lines 197–197 · Calculation / assignment:** Assign `target.EntryId` using the exact expression `fallbackId`.
  - VBA: `target.EntryId = fallbackId`
- **Lines 198–198 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 200–200 · Calculation / assignment:** Assign `target.AccountCode` using the exact expression `UCase$(Trim$(CStr(parts(LBound(parts) + 1))))`.
  - VBA: `target.AccountCode = UCase$(Trim$(CStr(parts(LBound(parts) + 1))))`
- **Lines 201–201 · Decision with inline action (review):** This line has a decision and an inline action; both remain in the exact source statement and need review.
  - VBA: `If Len(target.AccountCode) = 0 Then GoTo ParseFail`
- **Lines 203–203 · Calculation / assignment:** Assign `rawAmt` using the exact expression `CDbl(Trim$(CStr(parts(LBound(parts) + 2))))`.
  - VBA: `rawAmt = CDbl(Trim$(CStr(parts(LBound(parts) + 2))))`
- **Lines 204–204 · Calculation / assignment:** Assign `target.Amount` using the exact expression `CCur(rawAmt)`.
  - VBA: `target.Amount = CCur(rawAmt)`
- **Lines 206–206 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If IsDate(Trim$(CStr(parts(LBound(parts) + 3)))) Then`
- **Lines 207–207 · Calculation / assignment:** Assign `target.PostedDate` using the exact expression `CDate(Trim$(CStr(parts(LBound(parts) + 3))))`.
  - VBA: `target.PostedDate = CDate(Trim$(CStr(parts(LBound(parts) + 3))))`
- **Lines 208–208 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `Else`
- **Lines 209–209 · Calculation / assignment:** Assign `target.PostedDate` using the exact expression `Date`.
  - VBA: `target.PostedDate = Date`
- **Lines 210–210 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 212–212 · Calculation / assignment:** Assign `target.IsReconciled` using the exact expression `False`.
  - VBA: `target.IsReconciled = False`
- **Lines 214–214 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If UBound(parts) - LBound(parts) + 1 >= 5 And Len(Trim$(CStr(parts(LBound(parts) + 4)))) > 0 Then`
- **Lines 215–215 · Calculation / assignment:** Assign `tagTokens` using the exact expression `Split(CStr(parts(LBound(parts) + 4)), ",")`.
  - VBA: `tagTokens = Split(CStr(parts(LBound(parts) + 4)), ",")`
- **Lines 216–216 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `ReDim target.Tags(1 To UBound(tagTokens) - LBound(tagTokens) + 1)`
- **Lines 217–217 · Loop control:** This line starts, controls, or ends repeated work; bounds and conditions are preserved below.
  - VBA: `For i = LBound(tagTokens) To UBound(tagTokens)`
- **Lines 218–218 · Calculation / assignment:** Assign `target.Tags(i - LBound(tagTokens) + 1)` using the exact expression `LCase$(Trim$(CStr(tagTokens(i))))`.
  - VBA: `target.Tags(i - LBound(tagTokens) + 1) = LCase$(Trim$(CStr(tagTokens(i))))`
- **Lines 219–219 · Loop control:** This line starts, controls, or ends repeated work; bounds and conditions are preserved below.
  - VBA: `Next i`
- **Lines 220–220 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `Else`
- **Lines 221–221 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `ReDim target.Tags(1 To 1)`
- **Lines 222–222 · Calculation / assignment:** Assign `target.Tags(1)` using the exact expression `"untagged"`.
  - VBA: `target.Tags(1) = "untagged"`
- **Lines 223–223 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 225–225 · Calculation / assignment:** Assign `TryParseEntry` using the exact expression `True`.
  - VBA: `TryParseEntry = True`
- **Lines 226–226 · Jump / procedure flow:** Changes procedure flow; preserve the target and execution order during conversion.
  - VBA: `Exit Function`
- **Lines 228–228 · Label:** Label.
  - VBA: `ParseFail:`
- **Lines 229–229 · Calculation / assignment:** Assign `TryParseEntry` using the exact expression `False`.
  - VBA: `TryParseEntry = False`
- **Lines 230–230 · Procedure end:** Procedure end.
  - VBA: `End Function`

## Sub: `NormalizeEntryInPlace` (lines 235–267)

**Declaration:** `Private Sub NormalizeEntryInPlace(ByRef entry As LedgerEntry)`

### Ordered process notes

- **Lines 236–236 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim i As Long`
- **Lines 237–237 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim hasCoreTag As Boolean`
- **Lines 239–239 · Calculation / assignment:** Assign `entry.AccountCode` using the exact expression `Replace(entry.AccountCode, " ", "_")`.
  - VBA: `entry.AccountCode = Replace(entry.AccountCode, " ", "_")`
- **Lines 240–240 · Calculation / assignment:** Assign `entry.AccountCode` using the exact expression `Replace(entry.AccountCode, "-", "_")`.
  - VBA: `entry.AccountCode = Replace(entry.AccountCode, "-", "_")`
- **Lines 242–242 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If entry.Amount < 0@ Then`
- **Lines 243–243 · Calculation / assignment:** Assign `entry.IsReconciled` using the exact expression `False`.
  - VBA: `entry.IsReconciled = False`
- **Lines 244–244 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `ElseIf entry.Amount >= 10000@ Then`
- **Lines 245–245 · Calculation / assignment:** Assign `entry.IsReconciled` using the exact expression `False`.
  - VBA: `entry.IsReconciled = False`
- **Lines 246–246 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `Else`
- **Lines 247–247 · Calculation / assignment:** Assign `entry.IsReconciled` using the exact expression `True`.
  - VBA: `entry.IsReconciled = True`
- **Lines 248–248 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 250–250 · Calculation / assignment:** Assign `hasCoreTag` using the exact expression `False`.
  - VBA: `hasCoreTag = False`
- **Lines 251–251 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If (Not Not entry.Tags) <> 0 Then`
- **Lines 252–252 · Loop control:** This line starts, controls, or ends repeated work; bounds and conditions are preserved below.
  - VBA: `For i = LBound(entry.Tags) To UBound(entry.Tags)`
- **Lines 253–253 · Calculation / assignment:** Assign `entry.Tags(i)` using the exact expression `Trim$(entry.Tags(i))`.
  - VBA: `entry.Tags(i) = Trim$(entry.Tags(i))`
- **Lines 254–254 · Decision with inline action (review):** This line has a decision and an inline action; both remain in the exact source statement and need review.
  - VBA: `If entry.Tags(i) = "core" Then hasCoreTag = True`
- **Lines 255–255 · Loop control:** This line starts, controls, or ends repeated work; bounds and conditions are preserved below.
  - VBA: `Next i`
- **Lines 256–256 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 258–258 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If Not hasCoreTag Then`
- **Lines 259–259 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If (Not Not entry.Tags) = 0 Then`
- **Lines 260–260 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `ReDim entry.Tags(1 To 1)`
- **Lines 261–261 · Calculation / assignment:** Assign `entry.Tags(1)` using the exact expression `"core"`.
  - VBA: `entry.Tags(1) = "core"`
- **Lines 262–262 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `Else`
- **Lines 263–263 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `ReDim Preserve entry.Tags(1 To UBound(entry.Tags) + 1)`
- **Lines 264–264 · Calculation / assignment:** Assign `entry.Tags(UBound(entry.Tags))` using the exact expression `"core"`.
  - VBA: `entry.Tags(UBound(entry.Tags)) = "core"`
- **Lines 265–265 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 266–266 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 267–267 · Procedure end:** Procedure end.
  - VBA: `End Sub`

## Sub: `DeduplicateBatchEntries` (lines 272–311)

**Declaration:** `Private Sub DeduplicateBatchEntries(ByRef batch As BatchSummary)`

### Ordered process notes

- **Lines 273–273 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim seenKeys As Object`
- **Lines 274–274 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim kept() As LedgerEntry`
- **Lines 275–275 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim keptCount As Long`
- **Lines 276–276 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim i As Long`
- **Lines 277–277 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim key As String`
- **Lines 278–278 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim newTotal As Currency`
- **Lines 280–280 · Error handling:** Changes VBA error handling or refers to the active error; preserve this behavior during conversion.
  - VBA: `On Error Resume Next`
- **Lines 281–281 · Calculation / assignment:** Assign `Set seenKeys` using the exact expression `CreateObject("Scripting.Dictionary")`.
  - VBA: `Set seenKeys = CreateObject("Scripting.Dictionary")`
- **Lines 282–282 · Error handling:** Changes VBA error handling or refers to the active error; preserve this behavior during conversion.
  - VBA: `On Error GoTo 0`
- **Lines 284–284 · Decision with inline action (review):** This line has a decision and an inline action; both remain in the exact source statement and need review.
  - VBA: `If seenKeys Is Nothing Then Exit Sub`
- **Lines 286–286 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `ReDim kept(1 To batch.RecordCount)`
- **Lines 287–287 · Calculation / assignment:** Assign `keptCount` using the exact expression `0`.
  - VBA: `keptCount = 0`
- **Lines 288–288 · Calculation / assignment:** Assign `newTotal` using the exact expression `0@`.
  - VBA: `newTotal = 0@`
- **Lines 290–290 · Loop control:** This line starts, controls, or ends repeated work; bounds and conditions are preserved below.
  - VBA: `For i = LBound(batch.Entries) To UBound(batch.Entries)`
- **Lines 291–291 · Calculation / assignment:** Assign `key` using the exact expression `CStr(batch.Entries(i).EntryId) & "::" & batch.Entries(i).AccountCode`.
  - VBA: `key = CStr(batch.Entries(i).EntryId) & "::" & batch.Entries(i).AccountCode`
- **Lines 292–292 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If Not seenKeys.Exists(key) Then`
- **Lines 293–293 · Call / action:** Calls a procedure, property, or host action; identify its dependency and side effects.
  - VBA: `seenKeys.Add key, i`
- **Lines 294–294 · Calculation / assignment:** Assign `keptCount` using the exact expression `keptCount + 1`.
  - VBA: `keptCount = keptCount + 1`
- **Lines 295–295 · Calculation / assignment:** Assign `kept(keptCount)` using the exact expression `batch.Entries(i)`.
  - VBA: `kept(keptCount) = batch.Entries(i)`
- **Lines 296–296 · Calculation / assignment:** Assign `newTotal` using the exact expression `newTotal + batch.Entries(i).Amount`.
  - VBA: `newTotal = newTotal + batch.Entries(i).Amount`
- **Lines 297–297 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `Else`
- **Lines 298–298 · Call / action:** Calls a procedure, property, or host action; identify its dependency and side effects.
  - VBA: `LogAudit "DEDUP: Dropped duplicate key " & key`
- **Lines 299–299 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 300–300 · Loop control:** This line starts, controls, or ends repeated work; bounds and conditions are preserved below.
  - VBA: `Next i`
- **Lines 302–302 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If keptCount > 0 Then`
- **Lines 303–303 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `ReDim Preserve kept(1 To keptCount)`
- **Lines 304–304 · Calculation / assignment:** Assign `batch.Entries` using the exact expression `kept`.
  - VBA: `batch.Entries = kept`
- **Lines 305–305 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `Else`
- **Lines 306–306 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Erase batch.Entries`
- **Lines 307–307 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 309–309 · Calculation / assignment:** Assign `batch.RecordCount` using the exact expression `keptCount`.
  - VBA: `batch.RecordCount = keptCount`
- **Lines 310–310 · Calculation / assignment:** Assign `batch.TotalAmount` using the exact expression `newTotal`.
  - VBA: `batch.TotalAmount = newTotal`
- **Lines 311–311 · Procedure end:** Procedure end.
  - VBA: `End Sub`

## Sub: `RunLegacyGoSubAudit` (lines 317–346)

**Declaration:** `Private Sub RunLegacyGoSubAudit(ByRef batch As BatchSummary, ByRef summaryStatus As String)`

### Ordered process notes

- **Lines 318–318 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim i As Long`
- **Lines 319–319 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim flaggedCount As Long`
- **Lines 320–320 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim currentCheckId As Long`
- **Lines 321–321 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim checkResult As String`
- **Lines 323–323 · Calculation / assignment:** Assign `flaggedCount` using the exact expression `0`.
  - VBA: `flaggedCount = 0`
- **Lines 325–325 · Loop control:** This line starts, controls, or ends repeated work; bounds and conditions are preserved below.
  - VBA: `For i = LBound(batch.Entries) To UBound(batch.Entries)`
- **Lines 326–326 · Calculation / assignment:** Assign `currentCheckId` using the exact expression `batch.Entries(i).EntryId`.
  - VBA: `currentCheckId = batch.Entries(i).EntryId`
- **Lines 327–327 · Jump / procedure flow:** Changes procedure flow; preserve the target and execution order during conversion.
  - VBA: `GoSub ValidateSingleEntry`
- **Lines 328–328 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If checkResult <> "OK" Then`
- **Lines 329–329 · Calculation / assignment:** Assign `flaggedCount` using the exact expression `flaggedCount + 1`.
  - VBA: `flaggedCount = flaggedCount + 1`
- **Lines 330–330 · Call / action:** Calls a procedure, property, or host action; identify its dependency and side effects.
  - VBA: `LogAudit "FLAG [" & CStr(currentCheckId) & "]: " & checkResult`
- **Lines 331–331 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 332–332 · Loop control:** This line starts, controls, or ends repeated work; bounds and conditions are preserved below.
  - VBA: `Next i`
- **Lines 334–334 · Calculation / assignment:** Assign `summaryStatus` using the exact expression `"AUDITED_" & CStr(batch.RecordCount) & "_FLAGGED_" & CStr(flaggedCount)`.
  - VBA: `summaryStatus = "AUDITED_" & CStr(batch.RecordCount) & "_FLAGGED_" & CStr(flaggedCount)`
- **Lines 335–335 · Jump / procedure flow:** Changes procedure flow; preserve the target and execution order during conversion.
  - VBA: `Exit Sub`
- **Lines 337–337 · Label:** Label.
  - VBA: `ValidateSingleEntry:`
- **Lines 338–338 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If batch.Entries(i).Amount = 0@ Then`
- **Lines 339–339 · Calculation / assignment:** Assign `checkResult` using the exact expression `"ZERO_AMOUNT"`.
  - VBA: `checkResult = "ZERO_AMOUNT"`
- **Lines 340–340 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `ElseIf Year(batch.Entries(i).PostedDate) < 1990 Or Year(batch.Entries(i).PostedDate) > 2099 Then`
- **Lines 341–341 · Calculation / assignment:** Assign `checkResult` using the exact expression `"OUT_OF_RANGE_DATE"`.
  - VBA: `checkResult = "OUT_OF_RANGE_DATE"`
- **Lines 342–342 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `Else`
- **Lines 343–343 · Calculation / assignment:** Assign `checkResult` using the exact expression `"OK"`.
  - VBA: `checkResult = "OK"`
- **Lines 344–344 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 345–345 · Jump / procedure flow:** Changes procedure flow; preserve the target and execution order during conversion.
  - VBA: `Return`
- **Lines 346–346 · Procedure end:** Procedure end.
  - VBA: `End Sub`

## Function: `EvaluateDependencyDepth` (lines 352–397)

**Declaration:** `Public Function EvaluateDependencyDepth( _ ByVal edgesSpec As String, _ ByVal startNode As String, _ Optional ByVal currentDepth As Long = 0, _ Optional ByVal maxDepth As Long = 25 _ ) As Long`

### Ordered process notes

- **Lines 358–358 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim pairs As Variant`
- **Lines 359–359 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim edgeParts As Variant`
- **Lines 360–360 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim i As Long`
- **Lines 361–361 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim srcNode As String`
- **Lines 362–362 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim dstNode As String`
- **Lines 363–363 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim childDepth As Long`
- **Lines 364–364 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim maxChildDepth As Long`
- **Lines 366–366 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If currentDepth >= maxDepth Then`
- **Lines 367–367 · Calculation / assignment:** Assign `EvaluateDependencyDepth` using the exact expression `currentDepth`.
  - VBA: `EvaluateDependencyDepth = currentDepth`
- **Lines 368–368 · Jump / procedure flow:** Changes procedure flow; preserve the target and execution order during conversion.
  - VBA: `Exit Function`
- **Lines 369–369 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 371–371 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If Len(Trim$(edgesSpec)) = 0 Or Len(Trim$(startNode)) = 0 Then`
- **Lines 372–372 · Calculation / assignment:** Assign `EvaluateDependencyDepth` using the exact expression `currentDepth`.
  - VBA: `EvaluateDependencyDepth = currentDepth`
- **Lines 373–373 · Jump / procedure flow:** Changes procedure flow; preserve the target and execution order during conversion.
  - VBA: `Exit Function`
- **Lines 374–374 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 376–376 · Calculation / assignment:** Assign `pairs` using the exact expression `Split(edgesSpec, ",")`.
  - VBA: `pairs = Split(edgesSpec, ",")`
- **Lines 377–377 · Calculation / assignment:** Assign `maxChildDepth` using the exact expression `currentDepth`.
  - VBA: `maxChildDepth = currentDepth`
- **Lines 379–379 · Loop control:** This line starts, controls, or ends repeated work; bounds and conditions are preserved below.
  - VBA: `For i = LBound(pairs) To UBound(pairs)`
- **Lines 380–380 · Calculation / assignment:** Assign `edgeParts` using the exact expression `Split(CStr(pairs(i)), ">")`.
  - VBA: `edgeParts = Split(CStr(pairs(i)), ">")`
- **Lines 381–381 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If UBound(edgeParts) - LBound(edgeParts) + 1 = 2 Then`
- **Lines 382–382 · Calculation / assignment:** Assign `srcNode` using the exact expression `Trim$(CStr(edgeParts(LBound(edgeParts))))`.
  - VBA: `srcNode = Trim$(CStr(edgeParts(LBound(edgeParts))))`
- **Lines 383–383 · Calculation / assignment:** Assign `dstNode` using the exact expression `Trim$(CStr(edgeParts(LBound(edgeParts) + 1)))`.
  - VBA: `dstNode = Trim$(CStr(edgeParts(LBound(edgeParts) + 1)))`
- **Lines 385–385 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If StrComp(srcNode, Trim$(startNode), vbTextCompare) = 0 And Len(dstNode) > 0 Then`
- **Lines 386–386 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If StrComp(dstNode, srcNode, vbTextCompare) <> 0 Then`
- **Lines 387–387 · Calculation / assignment:** Assign `childDepth` using the exact expression `EvaluateDependencyDepth(edgesSpec, dstNode, currentDepth + 1, maxDepth)`.
  - VBA: `childDepth = EvaluateDependencyDepth(edgesSpec, dstNode, currentDepth + 1, maxDepth)`
- **Lines 388–388 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If childDepth > maxChildDepth Then`
- **Lines 389–389 · Calculation / assignment:** Assign `maxChildDepth` using the exact expression `childDepth`.
  - VBA: `maxChildDepth = childDepth`
- **Lines 390–390 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 391–391 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 392–392 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 393–393 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 394–394 · Loop control:** This line starts, controls, or ends repeated work; bounds and conditions are preserved below.
  - VBA: `Next i`
- **Lines 396–396 · Calculation / assignment:** Assign `EvaluateDependencyDepth` using the exact expression `maxChildDepth`.
  - VBA: `EvaluateDependencyDepth = maxChildDepth`
- **Lines 397–397 · Procedure end:** Procedure end.
  - VBA: `End Function`

## Function: `BuildUpsertSql` (lines 402–425)

**Declaration:** `Public Function BuildUpsertSql( _ ByVal tableName As String, _ ByRef entry As LedgerEntry _ ) As String`

### Ordered process notes

- **Lines 406–406 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim cleanTable As String`
- **Lines 407–407 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim cleanAcct As String`
- **Lines 408–408 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim cleanDate As String`
- **Lines 409–409 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim recFlag As Long`
- **Lines 411–411 · Calculation / assignment:** Assign `cleanTable` using the exact expression `SanitizeSqlIdentifier(tableName)`.
  - VBA: `cleanTable = SanitizeSqlIdentifier(tableName)`
- **Lines 412–412 · Calculation / assignment:** Assign `cleanAcct` using the exact expression `Replace(entry.AccountCode, "'", "''")`.
  - VBA: `cleanAcct = Replace(entry.AccountCode, "'", "''")`
- **Lines 413–413 · Calculation / assignment:** Assign `cleanDate` using the exact expression `Format$(entry.PostedDate, "yyyy-mm-dd")`.
  - VBA: `cleanDate = Format$(entry.PostedDate, "yyyy-mm-dd")`
- **Lines 414–414 · Calculation / assignment:** Assign `recFlag` using the exact expression `IIf(entry.IsReconciled, 1, 0)`.
  - VBA: `recFlag = IIf(entry.IsReconciled, 1, 0)`
- **Lines 416–424 · Calculation / assignment:** Assign `BuildUpsertSql` using the exact expression `_ "MERGE INTO " & cleanTable & " AS tgt " & _ "USING (SELECT " & CStr(entry.EntryId) & " AS entry_id, '" & cleanAcct & "' AS account_code) AS src " & _ "ON tgt.entry_id = src.entry_id AND tgt.account_code = src.account_code " & _ "WHEN MATCHED THEN UPDATE SET amount = " & Replace(CStr(entry.Amount), ",", ".") & _ ", posted_date = DATE '" & cleanDate & "', is_reconciled = " & CStr(recFlag) & " " & _ "WHEN NOT MATCHED THEN INSERT (entry_id, account_code, amount, posted_date, is_reconciled) " & _ "VALUES (" & CStr(entry.EntryId) & ", '" & cleanAcct & "', " & _ Replace(CStr(entry.Amount), ",", ".") & ", DATE '" & cleanDate & "', " & CStr(recFlag) & ");"`.
  - VBA: `BuildUpsertSql = _ "MERGE INTO " & cleanTable & " AS tgt " & _ "USING (SELECT " & CStr(entry.EntryId) & " AS entry_id, '" & cleanAcct & "' AS account_code) AS src " & _ "ON tgt.entry_id = src.entry_id AND tgt.account_code = src.account_code " & _ "WHEN MATCHED THEN UPDATE SET amount = " & Replace(CStr(entry.Amount), ",", ".") & _ ", posted_date = DATE '" & cleanDate & "', is_reconciled = " & CStr(recFlag) & " " & _ "WHEN NOT MATCHED THEN INSERT (entry_id, account_code, amount, posted_date, is_reconciled) " & _ "VALUES (" & CStr(entry.EntryId) & ", '" & cleanAcct & "', " & _ Replace(CStr(entry.Amount), ",", ".") & ", DATE '" & cleanDate & "', " & CStr(recFlag) & ");"`
- **Lines 425–425 · Procedure end:** Procedure end.
  - VBA: `End Function`

## Function: `SanitizeSqlIdentifier` (lines 427–443)

**Declaration:** `Private Function SanitizeSqlIdentifier(ByVal ident As String) As String`

### Ordered process notes

- **Lines 428–428 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim i As Long`
- **Lines 429–429 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim ch As String`
- **Lines 430–430 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim outStr As String`
- **Lines 432–432 · Calculation / assignment:** Assign `outStr` using the exact expression `""`.
  - VBA: `outStr = ""`
- **Lines 433–433 · Loop control:** This line starts, controls, or ends repeated work; bounds and conditions are preserved below.
  - VBA: `For i = 1 To Len(ident)`
- **Lines 434–434 · Calculation / assignment:** Assign `ch` using the exact expression `Mid$(ident, i, 1)`.
  - VBA: `ch = Mid$(ident, i, 1)`
- **Lines 435–436 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or _ (ch >= "0" And ch <= "9") Or ch = "_" Or ch = "." Then`
- **Lines 437–437 · Calculation / assignment:** Assign `outStr` using the exact expression `outStr & ch`.
  - VBA: `outStr = outStr & ch`
- **Lines 438–438 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 439–439 · Loop control:** This line starts, controls, or ends repeated work; bounds and conditions are preserved below.
  - VBA: `Next i`
- **Lines 441–441 · Decision with inline action (review):** This line has a decision and an inline action; both remain in the exact source statement and need review.
  - VBA: `If Len(outStr) = 0 Then outStr = "default_ledger_table"`
- **Lines 442–442 · Calculation / assignment:** Assign `SanitizeSqlIdentifier` using the exact expression `outStr`.
  - VBA: `SanitizeSqlIdentifier = outStr`
- **Lines 443–443 · Procedure end:** Procedure end.
  - VBA: `End Function`

## Function: `BuildSummaryReport` (lines 448–475)

**Declaration:** `Private Function BuildSummaryReport( _ ByRef batch As BatchSummary, _ ByVal parsedCount As Long, _ ByVal statusMsg As String _ ) As String`

### Ordered process notes

- **Lines 453–453 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim sb As String`
- **Lines 454–454 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim i As Long`
- **Lines 455–455 · Declaration / storage:** Declares, sizes, or clears VBA storage; the exact declaration is preserved below.
  - VBA: `Dim logLimit As Long`
- **Lines 457–457 · Calculation / assignment:** Assign `sb` using the exact expression `"=== BATCH SUMMARY ===" & vbCrLf`.
  - VBA: `sb = "=== BATCH SUMMARY ===" & vbCrLf`
- **Lines 458–458 · Calculation / assignment:** Assign `sb` using the exact expression `sb & "Batch ID:      " & batch.BatchId & vbCrLf`.
  - VBA: `sb = sb & "Batch ID:      " & batch.BatchId & vbCrLf`
- **Lines 459–459 · Calculation / assignment:** Assign `sb` using the exact expression `sb & "Mode Mask:     " & CStr(batch.ModeMask) & vbCrLf`.
  - VBA: `sb = sb & "Mode Mask:     " & CStr(batch.ModeMask) & vbCrLf`
- **Lines 460–460 · Calculation / assignment:** Assign `sb` using the exact expression `sb & "Parsed Rows:   " & CStr(parsedCount) & vbCrLf`.
  - VBA: `sb = sb & "Parsed Rows:   " & CStr(parsedCount) & vbCrLf`
- **Lines 461–461 · Calculation / assignment:** Assign `sb` using the exact expression `sb & "Final Rows:    " & CStr(batch.RecordCount) & vbCrLf`.
  - VBA: `sb = sb & "Final Rows:    " & CStr(batch.RecordCount) & vbCrLf`
- **Lines 462–462 · Calculation / assignment:** Assign `sb` using the exact expression `sb & "Total Amount:  " & Format$(batch.TotalAmount, "#,##0.00") & vbCrLf`.
  - VBA: `sb = sb & "Total Amount:  " & Format$(batch.TotalAmount, "#,##0.00") & vbCrLf`
- **Lines 463–463 · Calculation / assignment:** Assign `sb` using the exact expression `sb & "Audit Status:  " & statusMsg & vbCrLf`.
  - VBA: `sb = sb & "Audit Status:  " & statusMsg & vbCrLf`
- **Lines 465–465 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If Not m_AuditLog Is Nothing And m_AuditLog.Count > 0 Then`
- **Lines 466–466 · Calculation / assignment:** Assign `sb` using the exact expression `sb & "--- RECENT AUDIT LOGS ---" & vbCrLf`.
  - VBA: `sb = sb & "--- RECENT AUDIT LOGS ---" & vbCrLf`
- **Lines 467–467 · Calculation / assignment:** Assign `logLimit` using the exact expression `m_AuditLog.Count`.
  - VBA: `logLimit = m_AuditLog.Count`
- **Lines 468–468 · Decision with inline action (review):** This line has a decision and an inline action; both remain in the exact source statement and need review.
  - VBA: `If logLimit > 15 Then logLimit = 15`
- **Lines 469–469 · Loop control:** This line starts, controls, or ends repeated work; bounds and conditions are preserved below.
  - VBA: `For i = 1 To logLimit`
- **Lines 470–470 · Calculation / assignment:** Assign `sb` using the exact expression `sb & CStr(m_AuditLog(i)) & vbCrLf`.
  - VBA: `sb = sb & CStr(m_AuditLog(i)) & vbCrLf`
- **Lines 471–471 · Loop control:** This line starts, controls, or ends repeated work; bounds and conditions are preserved below.
  - VBA: `Next i`
- **Lines 472–472 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 474–474 · Calculation / assignment:** Assign `BuildSummaryReport` using the exact expression `sb`.
  - VBA: `BuildSummaryReport = sb`
- **Lines 475–475 · Procedure end:** Procedure end.
  - VBA: `End Function`

## Function: `GetAuditEntryCount` (lines 477–483)

**Declaration:** `Public Function GetAuditEntryCount() As Long`

### Ordered process notes

- **Lines 478–478 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `If m_AuditLog Is Nothing Then`
- **Lines 479–479 · Calculation / assignment:** Assign `GetAuditEntryCount` using the exact expression `0`.
  - VBA: `GetAuditEntryCount = 0`
- **Lines 480–480 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `Else`
- **Lines 481–481 · Calculation / assignment:** Assign `GetAuditEntryCount` using the exact expression `m_AuditLog.Count`.
  - VBA: `GetAuditEntryCount = m_AuditLog.Count`
- **Lines 482–482 · Decision / branch:** This line selects or changes a decision branch; the condition text is preserved below.
  - VBA: `End If`
- **Lines 483–483 · Procedure end:** Procedure end.
  - VBA: `End Function`

## Sub: `ResetModuleState` (lines 485–489)

**Declaration:** `Public Sub ResetModuleState()`

### Ordered process notes

- **Lines 486–486 · Calculation / assignment:** Assign `m_GlobalSequence` using the exact expression `0`.
  - VBA: `m_GlobalSequence = 0`
- **Lines 487–487 · Calculation / assignment:** Assign `Set m_AuditLog` using the exact expression `New Collection`.
  - VBA: `Set m_AuditLog = New Collection`
- **Lines 488–488 · Calculation / assignment:** Assign `Set m_NodeRegistry` using the exact expression `Nothing`.
  - VBA: `Set m_NodeRegistry = Nothing`
- **Lines 489–489 · Procedure end:** Procedure end.
  - VBA: `End Sub`

## Calculation ledger

Every calculation-like assignment found by the deterministic source scan is listed here. The original statement and expression are kept verbatim for conversion checks.

- Line 14: `ModeNone` = `0`; source: `ModeNone = 0`
- Line 15: `ModeNormalize` = `1`; source: `ModeNormalize = 1`
- Line 16: `ModeDeduplicate` = `2`; source: `ModeDeduplicate = 2`
- Line 17: `ModeAggregate` = `4`; source: `ModeAggregate = 4`
- Line 18: `ModeAudit` = `8`; source: `ModeAudit = 8`
- Line 19: `ModeAll` = `ModeNormalize Or ModeDeduplicate Or ModeAggregate Or ModeAudit`; source: `ModeAll = ModeNormalize Or ModeDeduplicate Or ModeAggregate Or ModeAudit`
- Line 70: `outReport` = `"EMPTY_PAYLOAD"`; source: `outReport = "EMPTY_PAYLOAD"`
- Line 71: `ExecuteStressPipeline` = `False`; source: `ExecuteStressPipeline = False`
- Line 76: `batch.BatchId` = `"BATCH_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(NextSequence())`; source: `batch.BatchId = "BATCH_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(NextSequence())`
- Line 77: `batch.ModeMask` = `modeMask`; source: `batch.ModeMask = modeMask`
- Line 78: `batch.RecordCount` = `0`; source: `batch.RecordCount = 0`
- Line 79: `batch.TotalAmount` = `0@`; source: `batch.TotalAmount = 0@`
- Line 81: `lines` = `Split(NormalizeLineEndings(rawPayload), vbLf)`; source: `lines = Split(NormalizeLineEndings(rawPayload), vbLf)`
- Line 92: `batch.RecordCount` = `batch.RecordCount + 1`; source: `batch.RecordCount = batch.RecordCount + 1`
- Line 98: `tempAmount` = `batch.Entries(batch.RecordCount).Amount`; source: `tempAmount = batch.Entries(batch.RecordCount).Amount`
- Line 100: `batch.TotalAmount` = `batch.TotalAmount + tempAmount`; source: `batch.TotalAmount = batch.TotalAmount + tempAmount`
- Line 103: `processedCount` = `processedCount + 1`; source: `processedCount = processedCount + 1`
- Line 127: `statusMsg` = `"AUDIT_SKIPPED"`; source: `statusMsg = "AUDIT_SKIPPED"`
- Line 130: `outReport` = `BuildSummaryReport(batch, processedCount, statusMsg)`; source: `outReport = BuildSummaryReport(batch, processedCount, statusMsg)`
- Line 131: `ExecuteStressPipeline` = `True`; source: `ExecuteStressPipeline = True`
- Line 135: `outReport` = `"PIPELINE_ERROR [" & CStr(Err.Number) & "]: " & Err.Description & " at line/step context."`; source: `outReport = "PIPELINE_ERROR [" & CStr(Err.Number) & "]: " & Err.Description & " at line/step context."`
- Line 137: `ExecuteStressPipeline` = `False`; source: `ExecuteStressPipeline = False`
- Line 146: `Set m_NodeRegistry` = `CreateObject("Scripting.Dictionary")`; source: `Set m_NodeRegistry = CreateObject("Scripting.Dictionary")`
- Line 149: `Set m_NodeRegistry` = `New Collection`; source: `Set m_NodeRegistry = New Collection`
- Line 154: `Set m_AuditLog` = `New Collection`; source: `Set m_AuditLog = New Collection`
- Line 159: `m_GlobalSequence` = `m_GlobalSequence + 1`; source: `m_GlobalSequence = m_GlobalSequence + 1`
- Line 164: `Set m_AuditLog` = `New Collection`; source: `If m_AuditLog Is Nothing Then Set m_AuditLog = New Collection`
- Line 173: `result` = `Replace(text, vbCrLf, vbLf)`; source: `result = Replace(text, vbCrLf, vbLf)`
- Line 174: `result` = `Replace(result, vbCr, vbLf)`; source: `result = Replace(result, vbCr, vbLf)`
- Line 175: `NormalizeLineEndings` = `result`; source: `NormalizeLineEndings = result`
- Line 190: `parts` = `Split(rawLine, "\|")`; source: `parts = Split(rawLine, "\|")`
- Line 195: `target.EntryId` = `CLng(Trim$(parts(LBound(parts))))`; source: `target.EntryId = CLng(Trim$(parts(LBound(parts))))`
- Line 197: `target.EntryId` = `fallbackId`; source: `target.EntryId = fallbackId`
- Line 200: `target.AccountCode` = `UCase$(Trim$(CStr(parts(LBound(parts) + 1))))`; source: `target.AccountCode = UCase$(Trim$(CStr(parts(LBound(parts) + 1))))`
- Line 203: `rawAmt` = `CDbl(Trim$(CStr(parts(LBound(parts) + 2))))`; source: `rawAmt = CDbl(Trim$(CStr(parts(LBound(parts) + 2))))`
- Line 204: `target.Amount` = `CCur(rawAmt)`; source: `target.Amount = CCur(rawAmt)`
- Line 207: `target.PostedDate` = `CDate(Trim$(CStr(parts(LBound(parts) + 3))))`; source: `target.PostedDate = CDate(Trim$(CStr(parts(LBound(parts) + 3))))`
- Line 209: `target.PostedDate` = `Date`; source: `target.PostedDate = Date`
- Line 212: `target.IsReconciled` = `False`; source: `target.IsReconciled = False`
- Line 215: `tagTokens` = `Split(CStr(parts(LBound(parts) + 4)), ",")`; source: `tagTokens = Split(CStr(parts(LBound(parts) + 4)), ",")`
- Line 218: `target.Tags(i - LBound(tagTokens) + 1)` = `LCase$(Trim$(CStr(tagTokens(i))))`; source: `target.Tags(i - LBound(tagTokens) + 1) = LCase$(Trim$(CStr(tagTokens(i))))`
- Line 222: `target.Tags(1)` = `"untagged"`; source: `target.Tags(1) = "untagged"`
- Line 225: `TryParseEntry` = `True`; source: `TryParseEntry = True`
- Line 229: `TryParseEntry` = `False`; source: `TryParseEntry = False`
- Line 239: `entry.AccountCode` = `Replace(entry.AccountCode, " ", "_")`; source: `entry.AccountCode = Replace(entry.AccountCode, " ", "_")`
- Line 240: `entry.AccountCode` = `Replace(entry.AccountCode, "-", "_")`; source: `entry.AccountCode = Replace(entry.AccountCode, "-", "_")`
- Line 243: `entry.IsReconciled` = `False`; source: `entry.IsReconciled = False`
- Line 245: `entry.IsReconciled` = `False`; source: `entry.IsReconciled = False`
- Line 247: `entry.IsReconciled` = `True`; source: `entry.IsReconciled = True`
- Line 250: `hasCoreTag` = `False`; source: `hasCoreTag = False`
- Line 253: `entry.Tags(i)` = `Trim$(entry.Tags(i))`; source: `entry.Tags(i) = Trim$(entry.Tags(i))`
- Line 254: `hasCoreTag` = `True`; source: `If entry.Tags(i) = "core" Then hasCoreTag = True`
- Line 261: `entry.Tags(1)` = `"core"`; source: `entry.Tags(1) = "core"`
- Line 264: `entry.Tags(UBound(entry.Tags))` = `"core"`; source: `entry.Tags(UBound(entry.Tags)) = "core"`
- Line 281: `Set seenKeys` = `CreateObject("Scripting.Dictionary")`; source: `Set seenKeys = CreateObject("Scripting.Dictionary")`
- Line 287: `keptCount` = `0`; source: `keptCount = 0`
- Line 288: `newTotal` = `0@`; source: `newTotal = 0@`
- Line 291: `key` = `CStr(batch.Entries(i).EntryId) & "::" & batch.Entries(i).AccountCode`; source: `key = CStr(batch.Entries(i).EntryId) & "::" & batch.Entries(i).AccountCode`
- Line 294: `keptCount` = `keptCount + 1`; source: `keptCount = keptCount + 1`
- Line 295: `kept(keptCount)` = `batch.Entries(i)`; source: `kept(keptCount) = batch.Entries(i)`
- Line 296: `newTotal` = `newTotal + batch.Entries(i).Amount`; source: `newTotal = newTotal + batch.Entries(i).Amount`
- Line 304: `batch.Entries` = `kept`; source: `batch.Entries = kept`
- Line 309: `batch.RecordCount` = `keptCount`; source: `batch.RecordCount = keptCount`
- Line 310: `batch.TotalAmount` = `newTotal`; source: `batch.TotalAmount = newTotal`
- Line 323: `flaggedCount` = `0`; source: `flaggedCount = 0`
- Line 326: `currentCheckId` = `batch.Entries(i).EntryId`; source: `currentCheckId = batch.Entries(i).EntryId`
- Line 329: `flaggedCount` = `flaggedCount + 1`; source: `flaggedCount = flaggedCount + 1`
- Line 334: `summaryStatus` = `"AUDITED_" & CStr(batch.RecordCount) & "_FLAGGED_" & CStr(flaggedCount)`; source: `summaryStatus = "AUDITED_" & CStr(batch.RecordCount) & "_FLAGGED_" & CStr(flaggedCount)`
- Line 339: `checkResult` = `"ZERO_AMOUNT"`; source: `checkResult = "ZERO_AMOUNT"`
- Line 341: `checkResult` = `"OUT_OF_RANGE_DATE"`; source: `checkResult = "OUT_OF_RANGE_DATE"`
- Line 343: `checkResult` = `"OK"`; source: `checkResult = "OK"`
- Line 367: `EvaluateDependencyDepth` = `currentDepth`; source: `EvaluateDependencyDepth = currentDepth`
- Line 372: `EvaluateDependencyDepth` = `currentDepth`; source: `EvaluateDependencyDepth = currentDepth`
- Line 376: `pairs` = `Split(edgesSpec, ",")`; source: `pairs = Split(edgesSpec, ",")`
- Line 377: `maxChildDepth` = `currentDepth`; source: `maxChildDepth = currentDepth`
- Line 380: `edgeParts` = `Split(CStr(pairs(i)), ">")`; source: `edgeParts = Split(CStr(pairs(i)), ">")`
- Line 382: `srcNode` = `Trim$(CStr(edgeParts(LBound(edgeParts))))`; source: `srcNode = Trim$(CStr(edgeParts(LBound(edgeParts))))`
- Line 383: `dstNode` = `Trim$(CStr(edgeParts(LBound(edgeParts) + 1)))`; source: `dstNode = Trim$(CStr(edgeParts(LBound(edgeParts) + 1)))`
- Line 387: `childDepth` = `EvaluateDependencyDepth(edgesSpec, dstNode, currentDepth + 1, maxDepth)`; source: `childDepth = EvaluateDependencyDepth(edgesSpec, dstNode, currentDepth + 1, maxDepth)`
- Line 389: `maxChildDepth` = `childDepth`; source: `maxChildDepth = childDepth`
- Line 396: `EvaluateDependencyDepth` = `maxChildDepth`; source: `EvaluateDependencyDepth = maxChildDepth`
- Line 411: `cleanTable` = `SanitizeSqlIdentifier(tableName)`; source: `cleanTable = SanitizeSqlIdentifier(tableName)`
- Line 412: `cleanAcct` = `Replace(entry.AccountCode, "'", "''")`; source: `cleanAcct = Replace(entry.AccountCode, "'", "''")`
- Line 413: `cleanDate` = `Format$(entry.PostedDate, "yyyy-mm-dd")`; source: `cleanDate = Format$(entry.PostedDate, "yyyy-mm-dd")`
- Line 414: `recFlag` = `IIf(entry.IsReconciled, 1, 0)`; source: `recFlag = IIf(entry.IsReconciled, 1, 0)`
- Line 416: `BuildUpsertSql` = `_ "MERGE INTO " & cleanTable & " AS tgt " & _ "USING (SELECT " & CStr(entry.EntryId) & " AS entry_id, '" & cleanAcct & "' AS account_code) AS src " & _ "ON tgt.entry_id = src.entry_id AND tgt.account_code = src.account_code " & _ "WHEN MATCHED THEN UPDATE SET amount = " & Replace(CStr(entry.Amount), ",", ".") & _ ", posted_date = DATE '" & cleanDate & "', is_reconciled = " & CStr(recFlag) & " " & _ "WHEN NOT MATCHED THEN INSERT (entry_id, account_code, amount, posted_date, is_reconciled) " & _ "VALUES (" & CStr(entry.EntryId) & ", '" & cleanAcct & "', " & _ Replace(CStr(entry.Amount), ",", ".") & ", DATE '" & cleanDate & "', " & CStr(recFlag) & ");"`; source: `BuildUpsertSql = _ "MERGE INTO " & cleanTable & " AS tgt " & _ "USING (SELECT " & CStr(entry.EntryId) & " AS entry_id, '" & cleanAcct & "' AS account_code) AS src " & _ "ON tgt.entry_id = src.entry_id AND tgt.account_code = src.account_code " & _ "WHEN MATCHED THEN UPDATE SET amount = " & Replace(CStr(entry.Amount), ",", ".") & _ ", posted_date = DATE '" & cleanDate & "', is_reconciled = " & CStr(recFlag) & " " & _ "WHEN NOT MATCHED THEN INSERT (entry_id, account_code, amount, posted_date, is_reconciled) " & _ "VALUES (" & CStr(entry.EntryId) & ", '" & cleanAcct & "', " & _ Replace(CStr(entry.Amount), ",", ".") & ", DATE '" & cleanDate & "', " & CStr(recFlag) & ");"`
- Line 432: `outStr` = `""`; source: `outStr = ""`
- Line 434: `ch` = `Mid$(ident, i, 1)`; source: `ch = Mid$(ident, i, 1)`
- Line 437: `outStr` = `outStr & ch`; source: `outStr = outStr & ch`
- Line 441: `outStr` = `"default_ledger_table"`; source: `If Len(outStr) = 0 Then outStr = "default_ledger_table"`
- Line 442: `SanitizeSqlIdentifier` = `outStr`; source: `SanitizeSqlIdentifier = outStr`
- Line 457: `sb` = `"=== BATCH SUMMARY ===" & vbCrLf`; source: `sb = "=== BATCH SUMMARY ===" & vbCrLf`
- Line 458: `sb` = `sb & "Batch ID:      " & batch.BatchId & vbCrLf`; source: `sb = sb & "Batch ID:      " & batch.BatchId & vbCrLf`
- Line 459: `sb` = `sb & "Mode Mask:     " & CStr(batch.ModeMask) & vbCrLf`; source: `sb = sb & "Mode Mask:     " & CStr(batch.ModeMask) & vbCrLf`
- Line 460: `sb` = `sb & "Parsed Rows:   " & CStr(parsedCount) & vbCrLf`; source: `sb = sb & "Parsed Rows:   " & CStr(parsedCount) & vbCrLf`
- Line 461: `sb` = `sb & "Final Rows:    " & CStr(batch.RecordCount) & vbCrLf`; source: `sb = sb & "Final Rows:    " & CStr(batch.RecordCount) & vbCrLf`
- Line 462: `sb` = `sb & "Total Amount:  " & Format$(batch.TotalAmount, "#,##0.00") & vbCrLf`; source: `sb = sb & "Total Amount:  " & Format$(batch.TotalAmount, "#,##0.00") & vbCrLf`
- Line 463: `sb` = `sb & "Audit Status:  " & statusMsg & vbCrLf`; source: `sb = sb & "Audit Status:  " & statusMsg & vbCrLf`
- Line 466: `sb` = `sb & "--- RECENT AUDIT LOGS ---" & vbCrLf`; source: `sb = sb & "--- RECENT AUDIT LOGS ---" & vbCrLf`
- Line 467: `logLimit` = `m_AuditLog.Count`; source: `logLimit = m_AuditLog.Count`
- Line 468: `logLimit` = `15`; source: `If logLimit > 15 Then logLimit = 15`
- Line 470: `sb` = `sb & CStr(m_AuditLog(i)) & vbCrLf`; source: `sb = sb & CStr(m_AuditLog(i)) & vbCrLf`
- Line 474: `BuildSummaryReport` = `sb`; source: `BuildSummaryReport = sb`
- Line 479: `GetAuditEntryCount` = `0`; source: `GetAuditEntryCount = 0`
- Line 481: `GetAuditEntryCount` = `m_AuditLog.Count`; source: `GetAuditEntryCount = m_AuditLog.Count`
- Line 486: `m_GlobalSequence` = `0`; source: `m_GlobalSequence = 0`
- Line 487: `Set m_AuditLog` = `New Collection`; source: `Set m_AuditLog = New Collection`
- Line 488: `Set m_NodeRegistry` = `Nothing`; source: `Set m_NodeRegistry = Nothing`

## LLM-assisted process explanation

**LLM status:** Complete — 15/15 procedures; openai/gpt-oss-20b; structured output and citations checked

### LLM interpretation: `ExecuteStressPipeline`

**Confidence:** high

The ExecuteStressPipeline function processes a raw payload string into a batch summary, applying optional mode flags for normalization, aggregation, deduplication, and audit. It returns a Boolean indicating success and outputs a report string. The function validates input, initializes infrastructure, parses each line up to a maximum iteration cap, updates batch statistics, optionally normalizes and aggregates entries, performs deduplication and audit, then builds a summary report. Errors trigger an error handler that logs and reports the exception. (source lines: 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 127, 128, 129, 130, 131, 132, 133, 134, 135, 136, 137, 138)

#### Inputs and outputs

- Inputs: rawPayload (String), modeMask (Long), optional outReport (String), optional maxIterations (Long). Outputs: Boolean return value, outReport string. (source lines: 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 127, 128, 129, 130, 131, 132, 133, 134, 135, 136, 137, 138)

#### Calculations

- outReport = "EMPTY_PAYLOAD"; exact VBA: `outReport = "EMPTY_PAYLOAD"` (source lines: 70)
- ExecuteStressPipeline = False; exact VBA: `ExecuteStressPipeline = False` (source lines: 71)
- batch.BatchId = "BATCH_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(NextSequence()); exact VBA: `batch.BatchId = "BATCH_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(NextSequence())` (source lines: 76)
- batch.ModeMask = modeMask; exact VBA: `batch.ModeMask = modeMask` (source lines: 77)
- batch.RecordCount = 0; exact VBA: `batch.RecordCount = 0` (source lines: 78)
- batch.TotalAmount = 0@; exact VBA: `batch.TotalAmount = 0@` (source lines: 79)
- lines = Split(NormalizeLineEndings(rawPayload), vbLf); exact VBA: `lines = Split(NormalizeLineEndings(rawPayload), vbLf)` (source lines: 81)
- batch.RecordCount = batch.RecordCount + 1; exact VBA: `batch.RecordCount = batch.RecordCount + 1` (source lines: 92)
- tempAmount = batch.Entries(batch.RecordCount).Amount; exact VBA: `tempAmount = batch.Entries(batch.RecordCount).Amount` (source lines: 98)
- batch.TotalAmount = batch.TotalAmount + tempAmount; exact VBA: `batch.TotalAmount = batch.TotalAmount + tempAmount` (source lines: 100)
- processedCount = processedCount + 1; exact VBA: `processedCount = processedCount + 1` (source lines: 103)
- statusMsg = "AUDIT_SKIPPED"; exact VBA: `statusMsg = "AUDIT_SKIPPED"` (source lines: 127)
- outReport = BuildSummaryReport(batch, processedCount, statusMsg); exact VBA: `outReport = BuildSummaryReport(batch, processedCount, statusMsg)` (source lines: 130)
- ExecuteStressPipeline = True; exact VBA: `ExecuteStressPipeline = True` (source lines: 131)
- outReport = "PIPELINE_ERROR [" & CStr(Err.Number) & "]: " & Err.Description & " at line/step context."; exact VBA: `outReport = "PIPELINE_ERROR [" & CStr(Err.Number) & "]: " & Err.Description & " at line/step context."` (source lines: 135)
- ExecuteStressPipeline = False; exact VBA: `ExecuteStressPipeline = False` (source lines: 137)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `"EMPTY_PAYLOAD"` (source lines: 70)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `False` (source lines: 71)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `"BATCH_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(NextSequence())` (source lines: 76)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `modeMask` (source lines: 77)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `0` (source lines: 78)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `0@` (source lines: 79)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `Split(NormalizeLineEndings(rawPayload), vbLf)` (source lines: 81)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `batch.RecordCount + 1` (source lines: 92)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `batch.Entries(batch.RecordCount).Amount` (source lines: 98)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `batch.TotalAmount + tempAmount` (source lines: 100)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `processedCount + 1` (source lines: 103)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `"AUDIT_SKIPPED"` (source lines: 127)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `BuildSummaryReport(batch, processedCount, statusMsg)` (source lines: 130)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `True` (source lines: 131)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `"PIPELINE_ERROR [" & CStr(Err.Number) & "]: " & Err.Description & " at line/step context."` (source lines: 135)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `False` (source lines: 137)

#### Decisions

- If rawPayload is empty, set outReport to "EMPTY_PAYLOAD", return False, and exit. (source lines: 69, 70, 71, 72, 73)
- If idx - LBound(lines) + 1 > maxIterations, log warning and exit loop. (source lines: 84, 85, 86, 87, 88)
- If line is non-empty and TryParseEntry succeeds, increment RecordCount and optionally normalize and aggregate. (source lines: 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103)
- If (idx Mod 250) = 0, call DoEvents. (source lines: 109, 110, 111)
- If batch.RecordCount > 0, resize Entries array to actual count; else erase Entries. (source lines: 114, 115, 116, 117, 118)
- If ModeDeduplicate flag set and more than one record, call DeduplicateBatchEntries. (source lines: 120, 121, 122)
- If ModeAudit flag set, run legacy audit; else set statusMsg to "AUDIT_SKIPPED". (source lines: 124, 125, 126, 127, 128)

#### Loops

- For idx = LBound(lines) To UBound(lines) iterates over each line. (source lines: 84, 112)

#### Data effects

- batch.BatchId is set to a unique string based on current timestamp and sequence. (source lines: 76)
- batch.ModeMask is set to modeMask. (source lines: 77)
- batch.RecordCount initialized to 0. (source lines: 78)
- batch.TotalAmount initialized to 0@ (Currency). (source lines: 79)
- lines array is split from rawPayload. (source lines: 81)
- batch.Entries array is dimensioned to number of lines. (source lines: 82)
- batch.RecordCount incremented when a line parses successfully. (source lines: 92)
- tempAmount holds the Amount of the current entry. (source lines: 98)
- batch.TotalAmount incremented by tempAmount when aggregation mode is set. (source lines: 100)
- processedCount incremented for each successfully parsed entry. (source lines: 103)
- statusMsg set to "AUDIT_SKIPPED" if audit not performed. (source lines: 127)
- outReport built from BuildSummaryReport with batch, processedCount, statusMsg. (source lines: 130)
- ExecuteStressPipeline returns True on success, False on error or empty payload. (source lines: 131, 137)

#### Dependencies

- Depends on external functions: InitInfrastructure, NormalizeLineEndings, TryParseEntry, NormalizeEntryInPlace, DeduplicateBatchEntries, RunLegacyGoSubAudit, BuildSummaryReport, LogAudit, NextSequence. (source lines: 75, 81, 91, 94, 98, 120, 124, 130, 135)

#### Unclear behavior / questions

- The behavior of NextSequence() and its effect on BatchId uniqueness is not defined within this procedure. (source lines: 76)
- The exact format and content of the summary report generated by BuildSummaryReport is unknown. (source lines: 130)
- The error handling references Err.Number and Err.Description, but the context of the error is not specified. (source lines: 135)

### LLM interpretation: `InitInfrastructure`

**Confidence:** high

Initializes infrastructure objects, handling errors when creating a dictionary for node registry and ensuring an audit log collection exists. (source lines: 144, 145, 146, 147, 148, 149, 150, 151, 152, 153, 154, 155, 156)

#### Calculations

- Creates a scripting dictionary for node registry.; exact VBA: `Set m_NodeRegistry = CreateObject("Scripting.Dictionary")` (source lines: 146)
- Falls back to a collection if dictionary creation fails.; exact VBA: `Set m_NodeRegistry = New Collection` (source lines: 149)
- Initializes audit log collection if not already set.; exact VBA: `Set m_AuditLog = New Collection` (source lines: 154)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `CreateObject("Scripting.Dictionary")` (source lines: 146)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `New Collection` (source lines: 149)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `New Collection` (source lines: 154)

#### Decisions

- Checks for error after attempting to create dictionary or if the registry is Nothing.; exact VBA: `If Err.Number <> 0 Or m_NodeRegistry Is Nothing Then` (source lines: 147)

#### Data effects

- Clears error state after handling failure.; exact VBA: `Err.Clear` (source lines: 148)
- Resets error handling to default after registry setup.; exact VBA: `On Error GoTo 0` (source lines: 151)
- Ensures audit log collection exists. (source lines: 153, 154, 155)

### LLM interpretation: `NextSequence`

**Confidence:** high

The NextSequence function increments the global sequence counter and returns the new value. (source lines: 158, 159, 160, 161)

#### Calculations

- Increment global sequence counter by one.; exact VBA: `m_GlobalSequence = m_GlobalSequence + 1` (source lines: 159)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `m_GlobalSequence + 1` (source lines: 159)

### LLM interpretation: `LogAudit`

**Confidence:** high

The LogAudit procedure records a timestamped message into a global audit log collection. It initializes the collection if it is not already created and then appends a formatted string containing the current date/time and the supplied message. (source lines: 163, 164, 165)

#### Inputs and outputs

- Adds a formatted timestamped message to the audit log collection.; exact VBA: `m_AuditLog.Add Format$(Now, "yyyy-mm-dd hh:nn:ss") & " \| " & message` (source lines: 165)

#### Calculations

- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `New Collection` (source lines: 164)

#### Data effects

- Initializes the global audit log collection if it is currently Nothing.; exact VBA: `Set m_AuditLog = New Collection` (source lines: 164)

### LLM interpretation: `NormalizeLineEndings`

**Confidence:** high

Normalizes line endings in a string by converting CRLF and CR to LF. (source lines: 171, 172, 173, 174, 175, 176)

#### Calculations

- Assigns result to the string with CRLF replaced by LF.; exact VBA: `result = Replace(text, vbCrLf, vbLf)` (source lines: 173)
- Replaces any remaining CR characters in result with LF.; exact VBA: `result = Replace(result, vbCr, vbLf)` (source lines: 174)
- Returns the normalized string as the function result.; exact VBA: `NormalizeLineEndings = result` (source lines: 175)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `Replace(text, vbCrLf, vbLf)` (source lines: 173)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `Replace(result, vbCr, vbLf)` (source lines: 174)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `result` (source lines: 175)

### LLM interpretation: `TryParseEntry`

**Confidence:** high

Parses a delimited string into a LedgerEntry object, handling optional fields and defaults, returning success status. (source lines: 178, 179, 180, 181, 182, 183, 184, 185, 186, 187, 188, 189, 190, 191, 192, 193, 194, 195, 196, 197, 198, 199, 200, 201, 202, 203, 204, 205, 206, 207, 208, 209, 210, 211, 212, 213, 214, 215, 216, 217, 218, 219, 220, 221, 222, 223, 224, 225, 226, 227, 228, 229, 230)

#### Inputs and outputs

- Inputs: rawLine (String), fallbackId (Long), target (LedgerEntry by reference). Output: Boolean return value indicating success. (source lines: 178, 179, 180, 181, 182, 225, 229)

#### Calculations

- Splits rawLine into parts array.; exact VBA: `parts = Split(rawLine, "\|")` (source lines: 190)
- Assigns EntryId from first part if numeric, else uses fallbackId.; exact VBA: `target.EntryId = CLng(Trim$(parts(LBound(parts))))` (source lines: 195)
- Assigns EntryId from fallbackId when first part not numeric.; exact VBA: `target.EntryId = fallbackId` (source lines: 197)
- Sets AccountCode to uppercased trimmed string of second part.; exact VBA: `target.AccountCode = UCase$(Trim$(CStr(parts(LBound(parts) + 1))))` (source lines: 200)
- Parses third part as double and assigns to rawAmt.; exact VBA: `rawAmt = CDbl(Trim$(CStr(parts(LBound(parts) + 2))))` (source lines: 203)
- Converts rawAmt to currency and assigns to Amount.; exact VBA: `target.Amount = CCur(rawAmt)` (source lines: 204)
- Parses fourth part as date if valid, else uses current date.; exact VBA: `target.PostedDate = CDate(Trim$(CStr(parts(LBound(parts) + 3))))` (source lines: 207)
- Sets PostedDate to current date when fourth part not a date.; exact VBA: `target.PostedDate = Date` (source lines: 209)
- Initializes IsReconciled to False.; exact VBA: `target.IsReconciled = False` (source lines: 212)
- Splits fifth part into tagTokens if present.; exact VBA: `tagTokens = Split(CStr(parts(LBound(parts) + 4)), ",")` (source lines: 215)
- Assigns each tag token to target.Tags array in lower case.; exact VBA: `target.Tags(i - LBound(tagTokens) + 1) = LCase$(Trim$(CStr(tagTokens(i))))` (source lines: 218)
- If no tags provided, assigns single tag "untagged".; exact VBA: `target.Tags(1) = "untagged"` (source lines: 222)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `Split(rawLine, "\|")` (source lines: 190)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `CLng(Trim$(parts(LBound(parts))))` (source lines: 195)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `fallbackId` (source lines: 197)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `UCase$(Trim$(CStr(parts(LBound(parts) + 1))))` (source lines: 200)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `CDbl(Trim$(CStr(parts(LBound(parts) + 2))))` (source lines: 203)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `CCur(rawAmt)` (source lines: 204)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `CDate(Trim$(CStr(parts(LBound(parts) + 3))))` (source lines: 207)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `Date` (source lines: 209)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `False` (source lines: 212)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `Split(CStr(parts(LBound(parts) + 4)), ",")` (source lines: 215)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `LCase$(Trim$(CStr(tagTokens(i))))` (source lines: 218)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `"untagged"` (source lines: 222)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `True` (source lines: 225)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `False` (source lines: 229)

#### Decisions

- Fails if parts array not valid or has fewer than 4 elements. (source lines: 191, 192, 193)
- Fails if AccountCode length is zero. (source lines: 201)
- Fails if fourth part is not a valid date and no date provided? (handled by default). (source lines: 206, 208)

#### Loops

- Iterates over tagTokens to populate target.Tags. (source lines: 217, 218, 219)

#### Data effects

- Writes to target object's properties: EntryId, AccountCode, Amount, PostedDate, IsReconciled, Tags. (source lines: 195, 197, 200, 203, 204, 207, 209, 212, 215, 218, 222)

#### Dependencies

- Depends on VBA functions: Split, IsArray, UBound, LBound, IsNumeric, Trim$, CStr, CLng, CDbl, CCur, CDate, IsDate, Len, UCase$, LCase$, ReDim, Exit Function, On Error GoTo. (source lines: 190, 191, 192, 193, 194, 195, 196, 197, 198, 199, 200, 201, 202, 203, 204, 205, 206, 207, 208, 209, 210, 211, 212, 213, 214, 215, 216, 217, 218, 219, 220, 221, 222, 223, 224, 225, 226, 228, 229)

#### Unclear behavior / questions

- Behavior when parts array contains more than expected elements beyond the fifth is not explicitly handled; extra elements are ignored. (source lines: 214, 215, 216)

### LLM interpretation: `NormalizeEntryInPlace`

**Confidence:** high

Normalizes a LedgerEntry by sanitizing its AccountCode, setting its IsReconciled flag based on Amount, trimming and ensuring a "core" tag is present. (source lines: 235, 236, 237, 238, 239, 240, 241, 242, 243, 244, 245, 246, 247, 248, 249, 250, 251, 252, 253, 254, 255, 256, 257, 258, 259, 260, 261, 262, 263, 264, 265, 266, 267)

#### Calculations

- Replace spaces in AccountCode with underscores; exact VBA: `entry.AccountCode = Replace(entry.AccountCode, " ", "_")` (source lines: 239)
- Replace hyphens in AccountCode with underscores; exact VBA: `entry.AccountCode = Replace(entry.AccountCode, "-", "_")` (source lines: 240)
- Set IsReconciled to False when Amount is negative; exact VBA: `entry.IsReconciled = False` (source lines: 243)
- Set IsReconciled to False when Amount is >= 10000; exact VBA: `entry.IsReconciled = False` (source lines: 245)
- Set IsReconciled to True otherwise; exact VBA: `entry.IsReconciled = True` (source lines: 247)
- Initialize hasCoreTag to False; exact VBA: `hasCoreTag = False` (source lines: 250)
- Trim each tag in entry.Tags; exact VBA: `entry.Tags(i) = Trim$(entry.Tags(i))` (source lines: 253)
- Set hasCoreTag to True if a tag equals "core"; exact VBA: `If entry.Tags(i) = "core" Then hasCoreTag = True` (source lines: 254)
- Assign "core" to first tag when Tags array is empty; exact VBA: `entry.Tags(1) = "core"` (source lines: 261)
- Assign "core" to last tag when Tags array already has elements; exact VBA: `entry.Tags(UBound(entry.Tags)) = "core"` (source lines: 264)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `Replace(entry.AccountCode, " ", "_")` (source lines: 239)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `Replace(entry.AccountCode, "-", "_")` (source lines: 240)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `False` (source lines: 243)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `False` (source lines: 245)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `True` (source lines: 247)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `False` (source lines: 250)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `Trim$(entry.Tags(i))` (source lines: 253)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `True` (source lines: 254)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `"core"` (source lines: 261)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `"core"` (source lines: 264)

#### Decisions

- Branch on Amount to set IsReconciled (source lines: 242, 243, 244, 245, 246, 247, 248)

#### Loops

- Iterate over existing tags to trim and detect "core" (source lines: 251, 252, 253, 254, 255, 256)

#### Data effects

- Modifies entry.AccountCode, entry.IsReconciled, and entry.Tags in place (source lines: 239, 240, 243, 245, 247, 253, 254, 261, 264)

### LLM interpretation: `DeduplicateBatchEntries`

**Confidence:** high

The procedure removes duplicate ledger entries from a batch by tracking unique keys composed of EntryId and AccountCode. It builds a new array of unique entries, updates the batch’s record count and total amount, and logs any duplicates found. (source lines: 272, 281, 284, 286, 290, 291, 292, 293, 294, 295, 296, 298, 302, 304, 309, 310, 311)

#### Inputs and outputs

- Inputs: batch (ByRef) containing Entries array, RecordCount, TotalAmount. Outputs: batch modified in place with deduplicated Entries, updated RecordCount and TotalAmount. (source lines: 272, 309, 310, 311)

#### Calculations

- Creates a dictionary object.; exact VBA: `Set seenKeys = CreateObject("Scripting.Dictionary")` (source lines: 281)
- Initializes keptCount to zero.; exact VBA: `keptCount = 0` (source lines: 287)
- Initializes newTotal to zero currency.; exact VBA: `newTotal = 0@` (source lines: 288)
- Builds a unique key string from EntryId and AccountCode.; exact VBA: `key = CStr(batch.Entries(i).EntryId) & "::" & batch.Entries(i).AccountCode` (source lines: 291)
- Increments keptCount by one.; exact VBA: `keptCount = keptCount + 1` (source lines: 294)
- Stores the current entry into the kept array.; exact VBA: `kept(keptCount) = batch.Entries(i)` (source lines: 295)
- Adds the entry amount to newTotal.; exact VBA: `newTotal = newTotal + batch.Entries(i).Amount` (source lines: 296)
- Assigns the deduplicated array back to batch.Entries.; exact VBA: `batch.Entries = kept` (source lines: 304)
- Updates batch.RecordCount with the number of kept entries.; exact VBA: `batch.RecordCount = keptCount` (source lines: 309)
- Updates batch.TotalAmount with the summed amount of kept entries.; exact VBA: `batch.TotalAmount = newTotal` (source lines: 310)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `CreateObject("Scripting.Dictionary")` (source lines: 281)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `0` (source lines: 287)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `0@` (source lines: 288)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `CStr(batch.Entries(i).EntryId) & "::" & batch.Entries(i).AccountCode` (source lines: 291)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `keptCount + 1` (source lines: 294)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `batch.Entries(i)` (source lines: 295)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `newTotal + batch.Entries(i).Amount` (source lines: 296)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `kept` (source lines: 304)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `keptCount` (source lines: 309)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `newTotal` (source lines: 310)

#### Decisions

- Checks if a key already exists in seenKeys; if not, keeps the entry; otherwise logs a duplicate. (source lines: 292, 293, 298)

#### Loops

- Iterates over all entries in batch.Entries using a For loop from LBound to UBound. (source lines: 290, 300)

#### Dependencies

- Uses Scripting.Dictionary to store seen keys. (source lines: 281, 292, 293)

### LLM interpretation: `RunLegacyGoSubAudit`

**Confidence:** high

The procedure audits each entry in a batch, counts flagged entries, logs reasons, and returns a status string. It iterates over the batch entries, validates each via a GoSub, updates a counter and logs when a check fails, then composes a summary status. Lines 317‑346 are used. (source lines: 317, 318, 319, 320, 321, 322, 323, 325, 326, 327, 328, 329, 330, 331, 332, 333, 334, 335, 337, 338, 339, 340, 341, 342, 343, 344, 345, 346)

#### Inputs and outputs

- Inputs: batch (BatchSummary) and summaryStatus (String) passed ByRef. Output: summaryStatus is set to a status string. (source lines: 317, 334, 335)

#### Calculations

- flaggedCount = 0; exact VBA: `flaggedCount = 0` (source lines: 323)
- currentCheckId = batch.Entries(i).EntryId; exact VBA: `currentCheckId = batch.Entries(i).EntryId` (source lines: 326)
- flaggedCount = flaggedCount + 1; exact VBA: `flaggedCount = flaggedCount + 1` (source lines: 329)
- summaryStatus = "AUDITED_" & CStr(batch.RecordCount) & "_FLAGGED_" & CStr(flaggedCount); exact VBA: `summaryStatus = "AUDITED_" & CStr(batch.RecordCount) & "_FLAGGED_" & CStr(flaggedCount)` (source lines: 334)
- checkResult = "ZERO_AMOUNT"; exact VBA: `checkResult = "ZERO_AMOUNT"` (source lines: 339)
- checkResult = "OUT_OF_RANGE_DATE"; exact VBA: `checkResult = "OUT_OF_RANGE_DATE"` (source lines: 341)
- checkResult = "OK"; exact VBA: `checkResult = "OK"` (source lines: 343)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `0` (source lines: 323)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `batch.Entries(i).EntryId` (source lines: 326)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `flaggedCount + 1` (source lines: 329)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `"AUDITED_" & CStr(batch.RecordCount) & "_FLAGGED_" & CStr(flaggedCount)` (source lines: 334)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `"ZERO_AMOUNT"` (source lines: 339)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `"OUT_OF_RANGE_DATE"` (source lines: 341)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `"OK"` (source lines: 343)

#### Decisions

- If checkResult <> "OK" then flag and log. (source lines: 328, 329, 330, 331)

#### Loops

- Iterates over batch entries from LBound to UBound. (source lines: 325, 332)

#### Dependencies

- Uses batch.Entries(i).EntryId, Amount, PostedDate, and batch.RecordCount. (source lines: 326, 338, 340, 334)

### LLM interpretation: `EvaluateDependencyDepth`

**Confidence:** high

The function computes the maximum depth of dependencies starting from a node by recursively traversing edges defined in a comma‑separated string. It returns the depth as a Long, bounded by a maximum depth parameter. The procedure validates inputs, splits the edge list, iterates over each edge, and updates the maximum child depth when a deeper path is found. (source lines: 352, 353, 354, 355, 356, 357, 358, 359, 360, 361, 362, 363, 364, 365, 366, 367, 368, 369, 370, 371, 372, 373, 374, 375, 376, 377, 378, 379, 380, 381, 382, 383, 384, 385, 386, 387, 388, 389, 390, 391, 392, 393, 394, 395, 396, 397)

#### Inputs and outputs

- Inputs: edgesSpec (String), startNode (String), optional currentDepth (Long, default 0), optional maxDepth (Long, default 25). Output: Long return value. (source lines: 352, 353, 354, 355, 356, 357, 363, 364, 396, 397)

#### Calculations

- EvaluateDependencyDepth = currentDepth; exact VBA: `EvaluateDependencyDepth = currentDepth` (source lines: 367, 372)
- pairs = Split(edgesSpec, ","); exact VBA: `pairs = Split(edgesSpec, ",")` (source lines: 376)
- maxChildDepth = currentDepth; exact VBA: `maxChildDepth = currentDepth` (source lines: 377)
- edgeParts = Split(CStr(pairs(i)), ">"); exact VBA: `edgeParts = Split(CStr(pairs(i)), ">")` (source lines: 380)
- srcNode = Trim$(CStr(edgeParts(LBound(edgeParts)))); exact VBA: `srcNode = Trim$(CStr(edgeParts(LBound(edgeParts))))` (source lines: 382)
- dstNode = Trim$(CStr(edgeParts(LBound(edgeParts) + 1))); exact VBA: `dstNode = Trim$(CStr(edgeParts(LBound(edgeParts) + 1)))` (source lines: 383)
- childDepth = EvaluateDependencyDepth(edgesSpec, dstNode, currentDepth + 1, maxDepth); exact VBA: `childDepth = EvaluateDependencyDepth(edgesSpec, dstNode, currentDepth + 1, maxDepth)` (source lines: 387)
- maxChildDepth = childDepth; exact VBA: `maxChildDepth = childDepth` (source lines: 389)
- EvaluateDependencyDepth = maxChildDepth; exact VBA: `EvaluateDependencyDepth = maxChildDepth` (source lines: 396)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `currentDepth` (source lines: 367)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `currentDepth` (source lines: 372)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `Split(edgesSpec, ",")` (source lines: 376)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `currentDepth` (source lines: 377)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `Split(CStr(pairs(i)), ">")` (source lines: 380)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `Trim$(CStr(edgeParts(LBound(edgeParts))))` (source lines: 382)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `Trim$(CStr(edgeParts(LBound(edgeParts) + 1)))` (source lines: 383)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `EvaluateDependencyDepth(edgesSpec, dstNode, currentDepth + 1, maxDepth)` (source lines: 387)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `childDepth` (source lines: 389)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `maxChildDepth` (source lines: 396)

#### Decisions

- If currentDepth >= maxDepth then return currentDepth. (source lines: 366, 367, 368, 369)
- If edgesSpec or startNode are empty after trimming then return currentDepth. (source lines: 371, 372, 373, 374)

#### Loops

- For i = LBound(pairs) To UBound(pairs) iterates over each edge pair. (source lines: 379, 380, 381, 382, 383, 384, 385, 386, 387, 388, 389, 390, 391, 392, 393, 394, 395)

#### Data effects

- pairs is assigned the result of Split(edgesSpec, ",") and used as the loop collection. (source lines: 376)
- maxChildDepth is initialized to currentDepth and may be updated to a deeper child depth. (source lines: 377, 378, 389, 390)

#### Dependencies

- Recursive call to EvaluateDependencyDepth with dstNode and incremented currentDepth. (source lines: 387)

### LLM interpretation: `BuildUpsertSql`

**Confidence:** high

BuildUpsertSql constructs a MERGE SQL statement for inserting or updating a ledger entry. It sanitizes the table name, escapes the account code, formats the posted date, converts the reconciled flag to 0/1, and builds the MERGE statement using these values. (source lines: 402, 403, 404, 405, 406, 407, 408, 409, 410, 411, 412, 413, 414, 416, 417, 418, 419, 420, 421, 422, 423, 424, 425)

#### Inputs and outputs

- Inputs: tableName (String), entry (LedgerEntry). Output: BuildUpsertSql (String). (source lines: 402, 403, 404, 405, 425)

#### Calculations

- cleanAcct = Replace(entry.AccountCode, "'", "''"); exact VBA: `cleanAcct = Replace(entry.AccountCode, "'", "''")` (source lines: 412)
- cleanDate = Format$(entry.PostedDate, "yyyy-mm-dd"); exact VBA: `cleanDate = Format$(entry.PostedDate, "yyyy-mm-dd")` (source lines: 413)
- recFlag = IIf(entry.IsReconciled, 1, 0); exact VBA: `recFlag = IIf(entry.IsReconciled, 1, 0)` (source lines: 414)
- BuildUpsertSql = "MERGE INTO " & cleanTable & " AS tgt " & "USING (SELECT " & CStr(entry.EntryId) & " AS entry_id, '" & cleanAcct & "' AS account_code) AS src " & "ON tgt.entry_id = src.entry_id AND tgt.account_code = src.account_code " & "WHEN MATCHED THEN UPDATE SET amount = " & Replace(CStr(entry.Amount), ",", ".") & ", posted_date = DATE '" & cleanDate & "', is_reconciled = " & CStr(recFlag) & " " & "WHEN NOT MATCHED THEN INSERT (entry_id, account_code, amount, posted_date, is_reconciled) " & "VALUES (" & CStr(entry.EntryId) & ", '" & cleanAcct & "', " & Replace(CStr(entry.Amount), ",", ".") & ", DATE '" & cleanDate & "', " & CStr(recFlag) & ");"; exact VBA: `BuildUpsertSql = _ "MERGE INTO " & cleanTable & " AS tgt " & _ "USING (SELECT " & CStr(entry.EntryId) & " AS entry_id, '" & cleanAcct & "' AS account_code) AS src " & _ "ON tgt.entry_id = src.entry_id AND tgt.account_code = src.account_code " & _ "WHEN MATCHED THEN UPDATE SET amount = " & Replace(CStr(entry.Amount), ",", ".") & _ ", posted_date = DATE '" & cleanDate & "', is_reconciled = " & CStr(recFlag) & " " & _ "WHEN NOT MATCHED THEN INSERT (entry_id, account_code, amount, posted_date, is_reconciled) " & _ "VALUES (" & CStr(entry.EntryId) & ", '" & cleanAcct & "', " & _ Replace(CStr(entry.Amount), ",", ".") & ", DATE '" & cleanDate & "', " & CStr(recFlag) & ");"` (source lines: 416, 417, 418, 419, 420, 421, 422, 423, 424)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `SanitizeSqlIdentifier(tableName)` (source lines: 411)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `Replace(entry.AccountCode, "'", "''")` (source lines: 412)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `Format$(entry.PostedDate, "yyyy-mm-dd")` (source lines: 413)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `IIf(entry.IsReconciled, 1, 0)` (source lines: 414)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `_ "MERGE INTO " & cleanTable & " AS tgt " & _ "USING (SELECT " & CStr(entry.EntryId) & " AS entry_id, '" & cleanAcct & "' AS account_code) AS src " & _ "ON tgt.entry_id = src.entry_id AND tgt.account_code = src.account_code " & _ "WHEN MATCHED THEN UPDATE SET amount = " & Replace(CStr(entry.Amount), ",", ".") & _ ", posted_date = DATE '" & cleanDate & "', is_reconciled = " & CStr(recFlag) & " " & _ "WHEN NOT MATCHED THEN INSERT (entry_id, account_code, amount, posted_date, is_reconciled) " & _ "VALUES (" & CStr(entry.EntryId) & ", '" & cleanAcct & "', " & _ Replace(CStr(entry.Amount), ",", ".") & ", DATE '" & cleanDate & "', " & CStr(recFlag) & ");"` (source lines: 416)

#### Dependencies

- Depends on external functions: SanitizeSqlIdentifier, Replace, Format$, IIf, CStr. (source lines: 411, 412, 413, 414, 416, 420, 423)

#### Unclear behavior / questions

- The behavior of SanitizeSqlIdentifier is not shown; its effect on cleanTable is uncertain. (source lines: 411)

### LLM interpretation: `SanitizeSqlIdentifier`

**Confidence:** high

The function builds a sanitized SQL identifier by iterating over each character of the input string, appending only alphanumeric characters, underscores, or periods to an output string. If the result is empty, it defaults to "default_ledger_table" and returns the final string. (source lines: 427, 428, 429, 430, 431, 432, 433, 434, 435, 436, 437, 438, 439, 440, 441, 442, 443)

#### Inputs and outputs

- Input: ident (String). Output: function return value (String). (source lines: 427, 442)

#### Calculations

- Initializes outStr to an empty string.; exact VBA: `outStr = ""` (source lines: 432)
- Assigns ch to the current character of ident.; exact VBA: `ch = Mid$(ident, i, 1)` (source lines: 434)
- Appends allowed character to outStr.; exact VBA: `outStr = outStr & ch` (source lines: 437)
- If outStr is empty after processing, assigns default value.; exact VBA: `outStr = "default_ledger_table"` (source lines: 441)
- Returns the sanitized string.; exact VBA: `SanitizeSqlIdentifier = outStr` (source lines: 442)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `""` (source lines: 432)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `Mid$(ident, i, 1)` (source lines: 434)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `outStr & ch` (source lines: 437)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `"default_ledger_table"` (source lines: 441)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `outStr` (source lines: 442)

#### Decisions

- Conditional check to determine if a character is allowed (letters, digits, underscore, period). (source lines: 435, 436)

#### Loops

- Iterates over each character of ident using a For loop from 1 to Len(ident). (source lines: 433, 439)

### LLM interpretation: `BuildSummaryReport`

**Confidence:** high

BuildSummaryReport constructs a textual summary of a batch, including batch metadata and up to 15 recent audit log entries. It concatenates strings using the vbCrLf line separator and formats numeric values. The function returns the assembled string. (source lines: 448, 449, 450, 451, 452, 453, 454, 455, 456, 457, 458, 459, 460, 461, 462, 463, 464, 465, 466, 467, 468, 469, 470, 471, 472, 473, 474, 475)

#### Inputs and outputs

- The procedure receives a BatchSummary object, a parsed row count, and a status message as inputs and returns a string. (source lines: 448, 449, 450, 451, 452, 474)

#### Calculations

- sb = "=== BATCH SUMMARY ===" & vbCrLf; exact VBA: `sb = "=== BATCH SUMMARY ===" & vbCrLf` (source lines: 457)
- sb = sb & "Batch ID:      " & batch.BatchId & vbCrLf; exact VBA: `sb = sb & "Batch ID:      " & batch.BatchId & vbCrLf` (source lines: 458)
- sb = sb & "Mode Mask:     " & CStr(batch.ModeMask) & vbCrLf; exact VBA: `sb = sb & "Mode Mask:     " & CStr(batch.ModeMask) & vbCrLf` (source lines: 459)
- sb = sb & "Parsed Rows:   " & CStr(parsedCount) & vbCrLf; exact VBA: `sb = sb & "Parsed Rows:   " & CStr(parsedCount) & vbCrLf` (source lines: 460)
- sb = sb & "Final Rows:    " & CStr(batch.RecordCount) & vbCrLf; exact VBA: `sb = sb & "Final Rows:    " & CStr(batch.RecordCount) & vbCrLf` (source lines: 461)
- sb = sb & "Total Amount:  " & Format$(batch.TotalAmount, "#,##0.00") & vbCrLf; exact VBA: `sb = sb & "Total Amount:  " & Format$(batch.TotalAmount, "#,##0.00") & vbCrLf` (source lines: 462)
- sb = sb & "Audit Status:  " & statusMsg & vbCrLf; exact VBA: `sb = sb & "Audit Status:  " & statusMsg & vbCrLf` (source lines: 463)
- logLimit = m_AuditLog.Count; exact VBA: `logLimit = m_AuditLog.Count` (source lines: 467)
- If logLimit > 15 Then logLimit = 15; exact VBA: `If logLimit > 15 Then logLimit = 15` (source lines: 468)
- sb = sb & CStr(m_AuditLog(i)) & vbCrLf; exact VBA: `sb = sb & CStr(m_AuditLog(i)) & vbCrLf` (source lines: 470)
- BuildSummaryReport = sb; exact VBA: `BuildSummaryReport = sb` (source lines: 474)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `"=== BATCH SUMMARY ===" & vbCrLf` (source lines: 457)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `sb & "Batch ID:      " & batch.BatchId & vbCrLf` (source lines: 458)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `sb & "Mode Mask:     " & CStr(batch.ModeMask) & vbCrLf` (source lines: 459)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `sb & "Parsed Rows:   " & CStr(parsedCount) & vbCrLf` (source lines: 460)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `sb & "Final Rows:    " & CStr(batch.RecordCount) & vbCrLf` (source lines: 461)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `sb & "Total Amount:  " & Format$(batch.TotalAmount, "#,##0.00") & vbCrLf` (source lines: 462)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `sb & "Audit Status:  " & statusMsg & vbCrLf` (source lines: 463)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `sb & "--- RECENT AUDIT LOGS ---" & vbCrLf` (source lines: 466)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `m_AuditLog.Count` (source lines: 467)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `15` (source lines: 468)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `sb & CStr(m_AuditLog(i)) & vbCrLf` (source lines: 470)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `sb` (source lines: 474)

#### Decisions

- If m_AuditLog is not Nothing and contains entries, the audit log section is appended. (source lines: 465, 466)

#### Loops

- For i = 1 To logLimit; exact VBA: `For i = 1 To logLimit` (source lines: 469)

#### Dependencies

- The function depends on the global collection m_AuditLog to append audit log entries. (source lines: 465, 466, 467, 468, 469, 470, 471, 472)

### LLM interpretation: `GetAuditEntryCount`

**Confidence:** high

Returns the number of audit log entries. If the audit log object is not initialized, it returns 0; otherwise it returns the count of entries in the audit log. (source lines: 477, 478, 479, 480, 481, 482, 483)

#### Calculations

- Assigns 0 to the function result when the audit log is uninitialized.; exact VBA: `GetAuditEntryCount = 0` (source lines: 479)
- Assigns the count of the audit log to the function result when the audit log is initialized.; exact VBA: `GetAuditEntryCount = m_AuditLog.Count` (source lines: 481)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `0` (source lines: 479)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `m_AuditLog.Count` (source lines: 481)

### LLM interpretation: `ResetModuleState`

**Confidence:** high

Resets module state by initializing global sequence counter, creating a new audit log collection, and clearing the node registry reference. (source lines: 485, 486, 487, 488, 489)

#### Calculations

- Sets the global sequence counter to zero.; exact VBA: `m_GlobalSequence = 0` (source lines: 486)
- Creates a new Collection object for the audit log.; exact VBA: `Set m_AuditLog = New Collection` (source lines: 487)
- Clears the node registry reference.; exact VBA: `Set m_NodeRegistry = Nothing` (source lines: 488)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `0` (source lines: 486)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `New Collection` (source lines: 487)
- Deterministic source inventory entry; interpretation not supplied by the LLM.; exact VBA: `Nothing` (source lines: 488)

## Review checklist

- [ ] Confirm every calculation and its exact expression.
- [ ] Confirm each decision condition and both/all branch outcomes.
- [ ] Confirm loop start, end, step, and exit conditions.
- [ ] Identify data read or written, including workbook, file, database, and UI effects.
- [ ] Identify dependencies called by each procedure.
- [ ] Review all statements marked `(review)`.
- [ ] Mark each procedure verified, unclear, or not reviewed before Java conversion.

## Coverage counts

- Assignment-like statements: 99
- Unclassified statements to review: 8
- Procedures without a recognized end: 0

> These notes are an evidence index, not a claim that the program's business meaning has been fully understood. Verify every explanation against the cited VBA before converting it.
