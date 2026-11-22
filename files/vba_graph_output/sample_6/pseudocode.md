# Pseudocode: vba_sample_6



> Generated from the VBA structural graph. This is a language-neutral behavioral description for downstream intermediate-representation design, not executable code.



## Translation notes and known limitations

- Graph nodes: 331; graph edges: 520; procedures detected: 15.

- Parser diagnostics reported: 3.

- Treat every `[UNCERTAIN]` / `[OPAQUE]` marker as a review item before translation.

- Supporting granular segments supplied (procedure snapshots excluded): 237.

- The graph may omit comments, formatting, or syntax it could not model; consult the original VBA for unresolved constructs.



## Module overview and shared declarations

**Module name**: `ComplexConversionStressTest`  
**Module kind**: Standard VBA module (no form or class dependencies).  
**Compiler options**:  
- `Option Explicit` – all variables must be declared.  
- `Option Compare Text` – string comparisons are case‑insensitive.  
- `Option Base 1` – the lower bound of all dynamic arrays is 1.  

### Enumerated constants

```pseudocode
Enum PipelineModeFlags
    ModeNone        = 0
    ModeNormalize   = 1
    ModeDeduplicate = 2
    ModeAggregate   = 4
    ModeAudit       = 8
    ModeAll         = ModeNormalize OR ModeDeduplicate OR ModeAggregate OR ModeAudit
End Enum
```

These constants are used as bit‑flags to control the behaviour of the main pipeline routine.

### User‑defined types

```pseudocode
Type LedgerEntry
    EntryId      : Long
    AccountCode  : String
    Amount       : Currency
    PostedDate   : Date
    Tags         : Array of String   // dynamic array
    IsReconciled : Boolean
End Type
```

```pseudocode
Type BatchSummary
    BatchId      : String
    ModeMask     : Long
    RecordCount  : Long
    TotalAmount  : Currency
    Entries      : Array of LedgerEntry   // dynamic array
End Type
```

```pseudocode
Private Type InternalGraphNode
    NodeKey   : String
    Weight    : Double
    Visited   : Boolean
    Children  : Array of String   // dynamic array
End Type
```

### Module‑level state variables

```pseudocode
Private m_GlobalSequence : Long          // sequence counter for unique IDs
Private m_AuditLog       : Collection    // holds audit log strings
Private m_NodeRegistry   : Object        // late‑bound Scripting.Dictionary or fallback Collection
```

- `m_GlobalSequence` is incremented by `NextSequence()` to generate unique identifiers.  
- `m_AuditLog` collects audit messages; it is instantiated lazily in `InitInfrastructure`.  
- `m_NodeRegistry` is intended to store graph nodes; it is created as a `Scripting.Dictionary` if available, otherwise as a `Collection`.  

### Initialization procedures

```pseudocode
Private Sub InitInfrastructure()
    // Attempt to create a late‑bound dictionary
    On Error Resume Next
    m_NodeRegistry = CreateObject("Scripting.Dictionary")
    If Err.Number <> 0 Or m_NodeRegistry Is Nothing Then
        Err.Clear
        m_NodeRegistry = New Collection
    End If
    On Error GoTo 0

    // Ensure the audit log collection exists
    If m_AuditLog Is Nothing Then
        m_AuditLog = New Collection
    End If
End Sub
```

`InitInfrastructure` is called by the public entry point `ExecuteStressPipeline` before any processing occurs. It guarantees that `m_NodeRegistry` and `m_AuditLog` are ready for use.

```pseudocode
Public Sub ResetModuleState()
    m_GlobalSequence = 0
    m_AuditLog = New Collection
    m_NodeRegistry = Nothing
End Sub
```

`ResetModuleState` clears the module’s internal state, resetting the sequence counter, recreating the audit log collection, and discarding any previously stored graph nodes.

### Summary of module‑level declarations

| Category | Declaration | Purpose |
|----------|-------------|---------|
| Compiler options | `Option Explicit`, `Option Compare Text`, `Option Base 1` | Enforce variable declaration, case‑insensitive string comparison, array lower bound = 1 |
| Attribute | `Attribute VB_Name = "ComplexConversionStressTest"` | Sets the module name |
| Enum | `PipelineModeFlags` | Bit‑flag constants for pipeline configuration |
| Types | `LedgerEntry`, `BatchSummary`, `InternalGraphNode` | Structured data used throughout the module |
| Private state | `m_GlobalSequence`, `m_AuditLog`, `m_NodeRegistry` | Internal counters, audit log, and graph node registry |
| Initialization | `InitInfrastructure`, `ResetModuleState` | Prepare and reset module state |

No other module‑level initialization logic is present; all runtime setup occurs within `InitInfrastructure` and is invoked by the public entry point.



### Procedure: InitInfrastructure (part 1/1)

**Context**  
- Module: `module_c7488e4fcd800f9b` (file `vba_sample_6.bas`)  
- Procedure ID: `procedure_2954e225a14758ac`  
- Called by: `procedure_3f3c2e33f7565363` (line 75)  
- Visibility: `PRIVATE`  
- Kind: `SUB` (no return value)  
- Parameters: none  

**Purpose**  
Initialise two module‑level collections used by the application:  
- `m_NodeRegistry` – a dictionary of nodes (fallback to a generic collection if dictionary creation fails).  
- `m_AuditLog` – a collection of audit entries (initialised only if not already set).  

The procedure deliberately suppresses errors during dictionary creation, checks for failure, and restores normal error handling before proceeding to audit log initialisation.

---

#### Pseudocode

```
PROCEDURE InitInfrastructure
    // ------------------------------------------------------------------
    // 1. Suppress runtime errors while attempting to create a dictionary.
    // ------------------------------------------------------------------
    ON ERROR RESUME NEXT

    // 2. Attempt to create a Scripting.Dictionary object.
    SET m_NodeRegistry = CreateObject("Scripting.Dictionary")

    // ------------------------------------------------------------------
    // 3. If an error occurred (Err.Number <> 0) or the dictionary is
    //    still Nothing, fall back to a generic Collection.
    // ------------------------------------------------------------------
    IF (Err.Number <> 0) OR (m_NodeRegistry IS NOTHING) THEN
        // Clear the error state so subsequent code is not affected.
        CALL Err.Clear

        // Replace the failed dictionary with a new Collection instance.
        SET m_NodeRegistry = New Collection
    END IF

    // ------------------------------------------------------------------
    // 4. Restore default error handling behaviour.
    // ------------------------------------------------------------------
    ON ERROR GOTO 0

    // ------------------------------------------------------------------
    // 5. Initialise the audit log collection if it has not already been
    //    created elsewhere in the program.
    // ------------------------------------------------------------------
    IF m_AuditLog IS NOTHING THEN
        SET m_AuditLog = New Collection
    END IF

    // ------------------------------------------------------------------
    // 6. Procedure ends – no return value.
    // ------------------------------------------------------------------
END PROCEDURE
```

**Notes on the pseudocode**

- `ON ERROR RESUME NEXT` and `ON ERROR GOTO 0` are explicit error‑handling directives that control how the runtime responds to runtime errors.  
- `Err.Number` is the global error object’s numeric code; a non‑zero value indicates that the previous statement raised an error.  
- `CreateObject("Scripting.Dictionary")` attempts to instantiate a COM dictionary object.  
- `New Collection` creates a standard VBA collection object.  
- The `IS NOTHING` test checks whether a reference variable has been assigned an object instance.  
- The procedure performs no other side effects beyond the assignments to `m_NodeRegistry` and `m_AuditLog`.  
- All identifiers (`m_NodeRegistry`, `m_AuditLog`, `Err`) are preserved exactly as they appear in the source.

### Procedure: NextSequence (part 1/1)

**Purpose**  
`NextSequence` is a private function that generates a monotonically increasing sequence number.  
It increments a module‑level counter `m_GlobalSequence` and returns the new value.

**Inputs**  
None.

**Outputs**  
*Return value*: a `Long` integer representing the next sequence number.

**State Changes**  
* Reads the current value of the module‑level variable `m_GlobalSequence`.  
* Writes a new value to `m_GlobalSequence` (current value + 1).  
* Assigns that new value to the function name `NextSequence`, which serves as the return value.

**Side Effects**  
* The global counter `m_GlobalSequence` is permanently altered.  
* The function is referenced by other procedures (e.g., at line 76 of `vba_sample_6.bas`) and by itself via the implicit return assignment at line 160.

**Pseudocode**

```
FUNCTION NextSequence() RETURNS Long
    // Increment the global sequence counter
    m_GlobalSequence = m_GlobalSequence + 1

    // Return the updated sequence value
    NextSequence = m_GlobalSequence
END FUNCTION
```

*Notes*  
- `m_GlobalSequence` is a module‑level variable; its initial value is defined elsewhere in the module.  
- The assignment to `NextSequence` is the implicit return mechanism used by VBA functions.  
- No error handling or additional logic is present in this fragment.



### Procedure: NormalizeLineEndings (part 1/1)

**Purpose**  
Normalizes all line‑ending characters in a supplied text string so that the result contains only the line‑feed character (`vbLf`).  
The function is public, accepts a single `String` argument passed by value, and returns a `String`.

**Inputs**  
- `text` (String, ByVal): The original text that may contain any combination of carriage‑return (`vbCr`), line‑feed (`vbLf`), or carriage‑return‑line‑feed (`vbCrLf`) sequences.

**Outputs**  
- Returns a `String` that contains the same characters as `text` except that every line‑ending sequence has been converted to a single `vbLf`.

**Local State**  
- `result` (String): Temporary variable used to accumulate the transformed text.

**Side Effects**  
- None. The procedure performs only local string manipulation and returns a value.

**Dependencies**  
- Built‑in `Replace` function: `Replace(sourceString, findText, replaceText)` returns a new string with all occurrences of `findText` replaced by `replaceText`.  
- Constants:  
  - `vbCrLf` – carriage‑return + line‑feed (`CRLF`).  
  - `vbCr`   – carriage‑return (`CR`).  
  - `vbLf`   – line‑feed (`LF`).

**Procedure Flow**  

```
1. Declare local variable:
      result : String

2. First replacement pass:
      result = Replace(text, vbCrLf, vbLf)
   // All CRLF sequences are replaced with a single LF.

3. Second replacement pass:
      result = Replace(result, vbCr, vbLf)
   // Any remaining CR characters (not part of CRLF) are replaced with LF.

4. Return the normalized string:
      NormalizeLineEndings = result
```

**Notes**  
- The two passes are necessary because `Replace` operates on literal substrings; replacing `vbCrLf`





### Procedure: DeduplicateBatchEntries (part 1/1)

**Purpose**  
Remove duplicate ledger entries from a `BatchSummary` object.  
A duplicate is defined as an entry that shares the same `EntryId` and `AccountCode`.  
The procedure keeps the first occurrence of each unique key, discards subsequent duplicates, and updates the batch’s entry list, record count, and total amount.

**Signature**  
```
PRIVATE SUB DeduplicateBatchEntries(ByRef batch As BatchSummary)
```

**Inputs**  
- `batch` – a reference to a `BatchSummary` object that contains:
  - `Entries` – an array of `LedgerEntry` objects.
  - `RecordCount` – the number of entries currently in `Entries`.
  - `TotalAmount` – the sum of the `Amount` fields of all entries.

**Outputs / Side‑Effects**  
- `batch.Entries` is replaced with a new array that contains only the first occurrence of each unique key.
- `batch.RecordCount` is set to the number of unique entries retained.
- `batch.TotalAmount` is set to the sum of the `Amount` values of the retained entries.
- Duplicate entries are logged via `LogAudit`.
- If the dictionary creation fails, the procedure exits immediately without modifying `batch`.

**Local Variables**  
| Variable | Type | Purpose |
|----------|------|---------|
| `seenKeys` | `Object` (Scripting.Dictionary) | Tracks keys that have already been seen. |
| `kept()` | `LedgerEntry` array | Temporary array to hold unique entries. |
| `keptCount` | `Long` | Count of unique entries found. |
| `i` | `Long` | Loop index over `batch.Entries`. |
| `key` | `String` | Composite key `EntryId::AccountCode`. |
| `newTotal` | `Currency` | Running total of amounts for unique entries. |

**Pseudocode**

```pseudocode
// 1. Declare local variables
DECLARE seenKeys AS Object
DECLARE kept() AS LedgerEntry
DECLARE keptCount AS Long
DECLARE i AS Long
DECLARE key AS String
DECLARE newTotal AS Currency

// 2. Suppress errors while creating the dictionary
ON ERROR RESUME NEXT
SET seenKeys = CreateObject("Scripting.Dictionary")
ON ERROR GO TO 0

// 3. If dictionary creation failed, exit early
IF seenKeys IS NOTHING THEN
    EXIT SUB
END IF

// 4. Prepare storage for unique entries
REDIM kept(1 TO batch.RecordCount)   // allocate maximum possible size
keptCount = 0
newTotal = 0@                           // Currency literal

// 5. Iterate over all entries in the batch
FOR i = LBound(batch.Entries) TO UBound(batch.Entries)
    // 5a. Build a composite key for the current entry
    key = CStr(batch.Entries(i).EntryId) & "::" & batch.Entries(i).AccountCode

    // 5b. Check if this key has already been seen
    IF NOT seenKeys.Exists(key) THEN
        // 5b.i. Record the key and its index
        seenKeys.Add key, i

        // 5b.ii. Store the entry in the temporary array
        keptCount = keptCount + 1
        kept(keptCount) = batch.Entries(i)

        // 5b.iii. Add the entry amount to the running total
        newTotal = newTotal + batch.Entries(i).Amount
    ELSE
        // 5c. Duplicate found – log the event
        LogAudit("DEDUP: Dropped duplicate key " & key)
    END IF
NEXT i

// 6. Replace the original entries with the unique set
IF keptCount > 0 THEN
    // Shrink the array to the actual number of unique entries
    REDIM PRESERVE kept(1 TO keptCount)
    batch.Entries = kept
ELSE
    // No unique entries – clear the array
    ERASE batch.Entries
END IF

// 7. Update batch metadata
batch.RecordCount = keptCount
batch.TotalAmount = newTotal
```

**Notes & Observations**

- `On Error Resume Next` is used only around the dictionary creation; any failure results in `seenKeys` remaining `Nothing`, which triggers an early exit.
- The composite key uses a double colon (`::`) as a separator

### Procedure: RunLegacyGoSubAudit (part 1/1)

**Purpose**  
`RunLegacyGoSubAudit` performs a legacy audit of a batch of entries.  
For each entry it validates the amount and posted date, counts how many
entries fail validation, logs flagged entries, and finally returns a
summary status string.

**Visibility**  
`PRIVATE` – only callable from within the same module.

**Parameters**  
| Name          | Passing | Type          | Description |
|---------------|---------|---------------|-------------|
| `batch`       | ByRef   | `BatchSummary` | The batch to audit. The procedure may modify the batch object. |
| `summaryStatus` | ByRef | `String` | Output parameter that receives a status string after the audit. |

**Return value**  
None (SUB).

**Side‑effects**  
*

### Procedure: EvaluateDependencyDepth (part 1/1)

**Signature**  
```
FUNCTION EvaluateDependencyDepth(
    edgesSpec   : String,
    startNode   : String,
    currentDepth : Long = 0,
    maxDepth     : Long = 25
) RETURNS Long
```

**Purpose**  
Recursively determine the maximum depth of a directed dependency graph described by `edgesSpec`.  
`edgesSpec` is a comma‑separated list of edges in the form `"source>destination"`.  
The function starts from `startNode` and explores all reachable nodes, returning the greatest number of hops required to reach any descendant, bounded by `maxDepth`.  

**Local Variables**  
```
pairs          : Variant   // array of edge strings
edgeParts      : Variant   // array of source/destination strings for a single edge
i              : Long      // loop index
srcNode        : String    // source node of current edge
dstNode        : String    // destination node of current edge
childDepth     : Long      // depth returned from recursive call
maxChildDepth  : Long      // maximum depth found among all children
```

**Algorithm**

```
1. // Early exit if recursion depth limit reached
2. IF currentDepth >= maxDepth THEN
3.     RETURN currentDepth
4. END IF

5. // Early exit if input specification or start node is empty
6. IF Len(Trim$(edgesSpec)) = 0 OR Len(Trim$(startNode)) = 0 THEN
7.     RETURN currentDepth
8. END IF

9. // Split the edge specification into individual edges
10. pairs = Split(edgesSpec, ",")

11. // Initialise maximum child depth to the current depth
12. maxChildDepth = currentDepth

13. // Iterate over each edge
14. FOR i = LBound(pairs) TO UBound(pairs)
15.     // Split the edge into source and destination
16.     edgeParts = Split(CStr(pairs(i)), ">")

17.     // Ensure the edge has exactly two parts
18.     IF UBound(edgeParts) - LBound(edgeParts) + 1 = 2 THEN
19.         // Extract and trim source and destination nodes
20.         srcNode = Trim$(CStr(edgeParts(LBound(edgeParts))))
21.         dstNode = Trim$(CStr(edgeParts(LBound(edgeParts) + 1)))

22.         // Process only if the source matches the start node and destination is non‑empty
23.         IF StrComp(srcNode, Trim$(startNode), vbTextCompare) = 0 AND Len(dstNode) > 0 THEN
24.             // Avoid self‑loops
25.             IF StrComp(dstNode, srcNode, vbTextCompare) <> 0 THEN
26.                 // Recursive call to evaluate depth from the destination node
27.                 childDepth = EvaluateDependencyDepth(
28.                                 edgesSpec,
29.                                 dstNode,
30.                                 currentDepth + 1,
31.                                 maxDepth
32.                              )
33.                 // Update maximum depth if this child is deeper
34.                 IF childDepth > maxChildDepth THEN
35.                     maxChildDepth = childDepth
36.                 END IF
37.             END IF
38.         END IF
39.     END IF
40. NEXT i

41. // Return the deepest depth found among all children
42. RETURN maxChildDepth
```

**Notes & Observations**

- The function is **public** and returns a `Long`.  
- Parameters `currentDepth` and `maxDepth` are optional and default to `0` and `25` respectively.  
- All string manipulation uses built‑in VBA functions: `Trim$`, `CStr`, `Split`, `Len`, and `StrComp` with `vbTextCompare`.  
- Recursion is the core mechanism; each call explores one level deeper until the depth limit or an empty input is reached.  
- No external resources (files

### Procedure: BuildUpsertSql (part 1/1)

**Signature**  
```
FUNCTION BuildUpsertSql(tableName AS String, entry AS LedgerEntry) AS String
```
* `tableName` – passed ByVal, a string containing the target table name.  
* `entry` – passed ByRef, a `LedgerEntry` object whose fields are read but not modified.  
* Return value – a single SQL string that performs a MERGE (upsert) into the specified table.

**Local Variables**  
```
cleanTable  AS String   // sanitized table name
cleanAcct   AS String   // account code with single quotes escaped
cleanDate   AS String   // posted date formatted as yyyy-mm-dd
recFlag     AS Long     // 1 if entry.IsReconciled is true, otherwise 0
```

**Procedure Body**

1. **Sanitize the table name**  
   ```
   cleanTable = SanitizeSqlIdentifier(tableName)
   ```
   *Calls the external routine `SanitizeSqlIdentifier` to escape or quote the table name so it is safe for use in SQL.*

2. **Escape single quotes in the account code**  
   ```
   cleanAcct = Replace(entry.AccountCode, "'", "''")
   ```
   *Replaces each single‑quote character in `entry.AccountCode` with two single quotes, the SQL escape sequence.*

3. **Format the posted date**  
   ```
   cleanDate = Format$(entry.PostedDate, "yyyy-mm-dd")
   ```
   *Converts the `Date` value `entry.PostedDate` to a string in ISO‑8601 format.*

4. **Determine the reconciliation flag**  
   ```
   recFlag = IIf(entry.IsReconciled, 1, 0)
   ```
   *If `entry.IsReconciled` is true, `recFlag` becomes 1; otherwise it becomes 0.*

5. **Build the MERGE statement**  
   ```
   BuildUpsertSql =
       "MERGE INTO " & cleanTable & " AS tgt " &
       "USING (SELECT " & CStr(entry.EntryId) & " AS entry_id, '" & cleanAcct & "' AS account_code) AS src " &
       "ON tgt.entry_id = src

### Procedure: SanitizeSqlIdentifier (part 1/1)

**Purpose**  
Sanitizes a supplied SQL identifier string by retaining only characters that are valid in most SQL identifiers (letters, digits, underscore, and period).  
If the input contains no valid characters, the function returns the default string `"default_ledger_table"`.

**Signature**  
```
PRIVATE FUNCTION SanitizeSqlIdentifier(ident AS String) AS String
```

**Inputs**  
- `ident` – the original identifier string to be sanitized (passed ByVal).

**Outputs**  
- Returns a string containing only the allowed characters from `ident`.  
- If no allowed characters are found, returns `"default_ledger_table"`.

**Local Variables**  
- `i` (Long) – loop counter for iterating over each character of `ident`.  
- `ch` (String) – the current character extracted from `ident`.  
- `outStr` (String) – accumulator for the sanitized identifier; initially empty.

**Side‑Effects**  
- No external state is modified.  
- The function only reads the input `ident` and writes to its local variables.

**Pseudocode**

```
BEGIN FUNCTION SanitizeSqlIdentifier(ident)

    // 1. Declare local variables
    DECLARE i AS Long
    DECLARE ch AS String
    DECLARE outStr AS String

    // 2. Initialize the output string to empty
    outStr = ""

    // 3. Iterate over each character of the input identifier
    FOR i FROM 1 TO LEN(ident) DO
        // 3a. Extract the i‑th character (1‑based indexing)
        ch = MID$(ident, i, 1)

        // 3b. Check if the character is allowed
        IF (ch >= "A" AND ch <= "Z") OR
           (ch >= "a" AND ch <= "z") OR
           (ch >= "0" AND ch <= "9") OR
           (ch = "_") OR
           (ch = ".") THEN

            // 3c. Append the allowed character to the output
            outStr = outStr & ch
        END IF
    END FOR

    // 4. If no valid characters were found, use the default name
    IF LEN(outStr) = 0 THEN
        outStr = "default_ledger_table"
    END IF

    // 5. Return the sanitized identifier
    RETURN outStr

END FUNCTION
```

**Notes**

- `MID$` is a VBA string function that returns a substring; here it is used to obtain a single character.
- The comparison operators (`>=`, `<=`, `=`) perform lexicographic comparison of single‑character strings.
- The function is marked `PRIVATE`, meaning it is intended for internal use within its module only.
- No error handling is present; the function assumes `ident` is a valid string (including the empty string).

### Procedure: BuildSummaryReport (part 1/1)

**Signature**  
```
PRIVATE FUNCTION BuildSummaryReport(batch As BatchSummary, parsedCount As Long, statusMsg As String) As String
```
* `batch` – passed **ByRef**; a `BatchSummary` object whose fields are read but not modified.  
* `parsedCount` – passed **ByVal**; a `Long` indicating how many rows were parsed.  
* `statusMsg` – passed **ByVal**; a `String` containing the audit status message.  
* Returns a `String` that contains a formatted summary of the batch and, if available, recent audit log entries.

---

#### Local Variables
| Name | Type | Purpose |
|------|------|---------|
| `sb` | `String` | Accumulates the report text. |
| `i` | `Long` | Loop counter for audit log entries. |
| `logLimit` | `Long` | Upper bound for the audit log loop (capped at 15). |

---

#### Pseudocode Body

```
1.  // Initialise the report header
2.  sb = "=== BATCH SUMMARY ===" + vbCrLf

3.  // Append batch metadata
4.  sb = sb + "Batch ID:      " + batch.BatchId + vbCrLf
5.  sb = sb + "Mode Mask:     " + CStr(batch.ModeMask) + vbCrLf
6.  sb = sb + "Parsed Rows:   " + CStr(parsedCount) + vbCrLf
7.  sb = sb + "Final Rows:    " + CStr(batch.RecordCount) + vbCrLf
8.  sb = sb + "Total Amount:  " + Format$(batch.TotalAmount, "#,##0.00") + vbCrLf
9.  sb = sb + "Audit Status:  " + statusMsg + vbCrLf

10. // If an audit log exists and contains entries, append recent logs
11. IF (m_AuditLog IS NOT Nothing) AND (m_AuditLog.Count > 0) THEN
12.     sb = sb + "--- RECENT AUDIT LOGS ---" + vbCrLf
13.     logLimit = m_AuditLog.Count
14.     IF logLimit > 15 THEN
15.         logLimit = 15
16.     END IF
17.     FOR i = 1 TO logLimit
18.         sb = sb + CStr(m_AuditLog(i)) + vbCrLf
19.     NEXT i
20. END IF

21. // Return the accumulated string
22. BuildSummaryReport = sb
```

---

#### Notes on Behavior

* **String Concatenation** – Each `sb = sb + ...` statement appends a new line to the report.  
* **Formatting** – `Format$(batch.TotalAmount, "#,##0.00")` formats the total amount with two decimal places and thousands separators.  
* **Audit Log Handling** –  
  * `m_AuditLog` is a global collection (or array) of audit messages.  
  * The function checks that it is not `Nothing` and that it contains at least one entry.  
  * It limits the number of displayed entries to a maximum of 15 to avoid excessive output.  
  * Each log entry is converted to a string with `CStr` before concatenation.  
* **Return Value** – The function assigns the final string to its own name (`BuildSummaryReport`) and returns it.  
* **Side Effects** – The only observable side effect is the returned string; no external state is modified.  
* **Error Handling** – No explicit error handling is present; any runtime errors (e.g., accessing an out‑of‑range audit log index) will propagate to the caller.  

This pseudocode preserves the exact control flow, evaluation order, and data transformations of the original VBA fragment while remaining language‑neutral.

### Procedure: GetAuditEntryCount (part 1/1)

**Signature**  
```
FUNCTION GetAuditEntryCount() RETURNS Long
```

**Description**  
`GetAuditEntryCount` reports the number of audit entries stored in the module‑level collection `m_AuditLog`.  
If `m_AuditLog` has not been initialized (`Nothing`), the function returns `0`.  
Otherwise it returns the value of the `Count` property of `m_AuditLog`.

**State accessed**  
- Reads the global variable `m_AuditLog`.  
- Reads `m_AuditLog.Count` when `m_AuditLog` is not `Nothing`.

**State modified**  
- The function’s return value (`GetAuditEntryCount`) is set to the computed count.  
- No other global or local state is altered.

**Side effects**  
- None beyond the return value.

**Pseudocode**  
```
FUNCTION GetAuditEntryCount() RETURNS Long
    // Determine the number of audit entries
    IF m_AuditLog IS Nothing THEN
        // Audit log not initialized → zero entries
        GetAuditEntryCount = 0
    ELSE
        // Audit log exists → use its Count property
        GetAuditEntryCount = m_AuditLog.Count
    END IF
    // The function returns the value assigned to GetAuditEntryCount
END FUNCTION
```

**Notes**  
- The assignment to `GetAuditEntryCount` is the mechanism by which a VBA function returns a value.  
- The `Count` property is assumed to be a numeric value of type `Long`.  
- No error handling or additional logic is present in this fragment.

### Procedure: ResetModuleState (part 1/1)

**Context**  
- Procedure `ResetModuleState` is a public `SUB` defined in module `module_c7488e4fcd800f9b` of project `vba_sample_6`.  
- It has no parameters and returns no value.  
- The procedure is intended to reset module‑level state variables to their initial values.

**Inputs**  
- None.

**Outputs / Return Value**  
- None (procedure returns `void`/`SUB`).

**State Read / Written**  
| Variable | Type | Read? | Written? | Description |
|----------|------|-------|----------|-------------|
| `m_GlobalSequence` | numeric (Integer/Long) | No | Yes | Global sequence counter, reset to `0`. |
| `m_AuditLog` | Collection object | No | Yes | Audit log collection, replaced with a new empty collection. |
| `m_NodeRegistry` | Object reference | No | Yes | Registry of nodes, cleared by setting to `Nothing`. |

**Side Effects**  
- After execution, `m_GlobalSequence` equals `0`.  
- `m_AuditLog` refers to a brand‑new, empty `Collection`.  
- `m_NodeRegistry` is set to `Nothing` (null reference).  
- No external resources (files, databases, UI) are accessed.  
- No errors are anticipated; the procedure performs straightforward assignments.

**Pseudocode**

```
PROCEDURE ResetModuleState
    // Reset the global sequence counter to zero
    m_GlobalSequence = 0

    // Create a new, empty audit log collection
    SET m_AuditLog = NEW Collection

    // Clear the node registry reference
    SET m_NodeRegistry = NOTHING
END PROCEDURE
```

**Control Flow**  
1. Execute the first assignment (`m_GlobalSequence = 0`).  
2. Execute the second assignment (`SET m_AuditLog = NEW Collection`).  
3. Execute the third assignment (`SET m_NodeRegistry = NOTHING`).  
4. Procedure terminates normally.

**Dependencies**  
- Requires the `Collection` class to be available in the runtime environment.  
- No other procedures or modules are called.

**Notes**  
- The `SET` keyword indicates object assignment; it is distinct from simple value assignment.  
- The procedure is deterministic and side‑effect‑only on the three module‑level variables.  
- No error handling constructs (e.g., `On Error`) are present in the body.



## Parser diagnostics (verbatim from graph)



- **warning / CROSS_NESTED_BLOCK_CLOSE** — End If closed a non-top IF frame; inspect nested syntax. (line 223)

- **warning / CROSS_NESTED_BLOCK_CLOSE** — End If closed a non-top IF frame; inspect nested syntax. (line 256)

- **warning / CROSS_NESTED_BLOCK_CLOSE** — End If closed a non-top IF frame; inspect nested syntax. (line 472)
