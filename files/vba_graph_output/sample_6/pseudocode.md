# Pseudocode: vba_sample_6



> Generated from the VBA structural graph. This is a language-neutral behavioral description for downstream intermediate-representation design, not executable code.



## Translation notes and known limitations

- Graph nodes: 323; graph edges: 502; procedures detected: 15.

- Parser diagnostics reported: 0.

- Treat every `[UNCERTAIN]` / `[OPAQUE]` marker as a review item before translation.

- Supporting granular segments supplied (procedure snapshots excluded): 228.

- The graph may omit comments, formatting, or syntax it could not model; consult the original VBA for unresolved constructs.



## Module overview and shared declarations

The module **ComplexConversionStressTest** is a standard VBA module that contains a set of constants, enumerations, user‑defined types, and module‑level state variables.  
The module is configured with the following compiler options:

```
Option Explicit   ' All variables must be declared
Option Compare Text   ' String comparisons are case‑insensitive
Option Base 1   ' Default lower bound for arrays is 1
```

The module name is declared via the attribute:

```
Attribute VB_Name = "ComplexConversionStressTest"
```

### Enumerations

```
Enum PipelineModeFlags
    ModeNone        = 0
    ModeNormalize   = 1
    ModeDeduplicate = 2
    ModeAggregate   = 4
    ModeAudit       = 8
    ModeAll         = ModeNormalize Or ModeDeduplicate Or ModeAggregate Or ModeAudit
End Enum
```

`PipelineModeFlags` is a bit‑mask that controls which stages of the processing pipeline are executed.

### User‑defined types

```
Type LedgerEntry
    EntryId      As Long
    AccountCode  As String
    Amount       As Currency
    PostedDate   As Date
    Tags()       As String      ' dynamic array of tags
    IsReconciled As Boolean
End Type
```

```
Type BatchSummary
    BatchId      As String
    ModeMask     As Long
    RecordCount  As Long
    TotalAmount  As Currency
    Entries()    As LedgerEntry   ' dynamic array of LedgerEntry
End Type
```

```
Private Type InternalGraphNode
    NodeKey   As String
    Weight    As Double
    Visited   As Boolean
    Children() As String   ' dynamic array of child node keys
End Private Type
```

### Module‑level state variables

```
Private m_GlobalSequence As Long          ' monotonically increasing counter
Private m_AuditLog       As Collection     ' collection of audit log strings
Private m_NodeRegistry   As Object         ' late‑bound dictionary or fallback collection
```

### Initialization helpers

The module provides two procedures that initialise or reset the module‑level state.  
These procedures are **not** automatically executed; they must be called explicitly by client code.

```
Private Sub InitInfrastructure()
    ' Attempt to create a late‑bound Scripting.Dictionary.
    ' If that fails, fall back to a standard VBA Collection.
    On Error Resume Next
    Set m_NodeRegistry = CreateObject("Scripting.Dictionary")
    If Err.Number <> 0 Or m_NodeRegistry Is Nothing Then
        Err.Clear
        Set m_NodeRegistry = New Collection
    End If
    On Error GoTo 0

    ' Ensure the audit log collection exists.
    If m_AuditLog Is Nothing Then
        Set m_AuditLog = New Collection
    End If
End Sub
```

```
Public Sub ResetModuleState()
    m_GlobalSequence = 0
    Set m_AuditLog = New Collection
    Set m_NodeRegistry = Nothing
End Sub
```

These helpers establish the runtime environment required by the rest of the module (e.g., `ExecuteStressPipeline` relies on `m_NodeRegistry` and `m_AuditLog` being initialised). No other implicit initialisation occurs at module load time.

### Procedure: ExecuteStressPipeline (part 1/1)

[GENERATION FAILED: ValueError('LLM returned an empty procedure section')]

### Procedure: InitInfrastructure (part 1/1)
**Purpose**  
Initializes two internal collections used by the module:  
- `m_NodeRegistry` – a dictionary‑like object that holds node information.  
- `m_AuditLog` – a collection that records audit events.  

The procedure is private, takes no arguments, and returns nothing.  
It performs defensive programming by temporarily disabling error handling while
attempting to create the dictionary, then restores normal error handling.

---

#### Pseudocode

```text
PROCEDURE InitInfrastructure
    // ------------------------------------------------------------------
    // 1. Disable default error handling so that CreateObject failures
    //    do not raise a runtime error.
    // ------------------------------------------------------------------
    SET error_handling_mode TO "ResumeNext"

    // ------------------------------------------------------------------
    // 2. Attempt to create a Scripting.Dictionary object.
    //    The result is stored in the module‑level variable m_NodeRegistry.
    // ------------------------------------------------------------------
    SET m_NodeRegistry TO CreateObject("Scripting.Dictionary")

    // ------------------------------------------------------------------
    // 3. If the previous step failed (Err.Number <> 0) or the object
    //    is still Nothing, clear the error state and fall back to a
    //    plain Collection object.
    // ------------------------------------------------------------------
    IF (Err.Number <> 0) OR (m_NodeRegistry IS Nothing) THEN
        CALL Err.Clear
        SET m_NodeRegistry TO New Collection
    END IF

    // ------------------------------------------------------------------
    // 4. Restore normal error handling behaviour.
    // ------------------------------------------------------------------
    SET error_handling_mode TO "GoTo0"

    // ------------------------------------------------------------------
    // 5. Ensure the audit log collection exists.
    // ------------------------------------------------------------------
    IF m_AuditLog IS Nothing THEN
        SET m_AuditLog TO New Collection
    END IF
END PROCEDURE
```

#### Detailed Notes

| Step | Action | Effect | Notes |
|------|--------|--------|-------|
| 1 | `On Error Resume Next` | Temporarily suppresses runtime errors. | Allows `CreateObject` to fail silently. |
| 2 | `Set m_NodeRegistry = CreateObject("Scripting.Dictionary")` | Attempts to instantiate a dictionary. | If the scripting runtime is unavailable, `Err.Number` will be non‑zero. |
| 3 | `If Err.Number <> 0 Or m_NodeRegistry Is Nothing Then` | Detects failure of dictionary creation. | `Err.Clear` resets the error state. |
| 3a | `Set m_NodeRegistry = New Collection` | Fallback to a generic collection. | Guarantees `m_NodeRegistry` is usable. |
| 4 | `On Error GoTo 0` | Restores default error handling. | Subsequent code will raise errors normally. |
| 5 | `If m_AuditLog Is Nothing Then` | Checks audit log existence. | `m_AuditLog` is a module‑level collection. |
| 5a | `Set m_AuditLog = New Collection` | Creates audit log if missing. | No error handling needed because this is a simple constructor. |

**Side Effects**

- `m_NodeRegistry` is guaranteed to reference either a `Scripting.Dictionary` or a `Collection`.  
- `m_AuditLog` is guaranteed to reference a `Collection`.  
- The global error state (`Err`) is cleared if the dictionary creation fails.  
- The procedure does not return a value; it only mutates module‑level state.

**Dependencies**

- Requires the Windows Scripting Runtime (`Scripting.Dictionary`).  
- Relies on the global `Err` object for error detection.  
- Uses module‑level variables `m_NodeRegistry` and `m_AuditLog`.  

**Uncertainty**  
None. All behavior is directly observable from the supplied AST and source fragments.

### Procedure: NextSequence (part 1/1)
**Context**  
- Module: `module_c7488e4fcd800f9b` (file `vba_sample_6.bas`)  
- Procedure is declared `PRIVATE FUNCTION NextSequence() As Long`  
- No parameters.  
- Returns a `Long` value.  
- The procedure is referenced by `procedure_3f3c2e33f7565363` at line 76 and by itself at line 160 (self‑reference is a no‑op in this fragment).  

**Side effects**  
- Modifies the global variable `m_GlobalSequence`.  
- No other state changes or external interactions occur in this fragment.  

**Pseudocode**

```
PROCEDURE NextSequence() RETURNS Long
    // Increment the global sequence counter
    m_GlobalSequence = m_GlobalSequence + 1

    // Return the updated counter value
    RETURN m_GlobalSequence
END PROCEDURE
```

**Explanation**  
1. `m_GlobalSequence` is a module‑level variable that holds the current sequence number.  
2. The first statement increases this counter by one.  
3. The second statement assigns the new counter value to the function’s return value (`NextSequence`).  
4. The function then returns that value to the caller.  

No error handling, branching, or loops are present in this fragment. The procedure’s sole purpose is to generate a monotonically increasing sequence number.

### Procedure: LogAudit (part 1/1)
**Context**  
- Module: `vba_sample_6`  
- Procedure visibility: `PRIVATE`  
- Procedure kind: `SUB` (no return value)  
- Parameter: `message` (String, passed ByVal)  

**Global state**  
- `m_AuditLog`: a module‑level variable that may hold a `Collection` object.  

**Pseudocode**

```
PROCEDURE LogAudit(message AS STRING)

    // Ensure the audit log collection exists
    IF m_AuditLog IS NOTHING THEN
        SET m_AuditLog = NEW Collection
    END IF

    // Create a timestamped audit entry
    timestamp = FORMAT$(NOW, "yyyy-mm-dd hh:nn:ss")
    entry = timestamp & " | " & message

    // Append the entry to the audit log
    m_AuditLog.Add(entry)

END PROCEDURE
```

**Behavioral notes**

- The procedure does **not** return a value; it performs side effects only.  
- `m_AuditLog` is created lazily on the first call; subsequent calls reuse the same collection.  
- The timestamp uses the current system date and time (`NOW`) formatted as `yyyy-mm-dd hh:nn:ss`.  
- The audit entry is a single string composed of the timestamp, a separator `" | "`, and the supplied `message`.  
- The entry is added to the collection via its `Add` method.  
- No error handling or validation is present; any runtime errors (e.g., if `m_AuditLog` cannot be instantiated) will propagate to the caller.  
- The procedure may be invoked by other procedures within the same module (e.g., `procedure_3f3c2e33f7565363`, `procedure_3104c2440f9bdac3`, `procedure_389d8383c644215c`).

### Procedure: NormalizeLineEndings (part 1/1)
**Purpose**  
Normalizes all line‑ending sequences in a given text string to the single line‑feed character (`vbLf`).  
The function accepts a string by value and returns a new string with the normalized line endings.

**Inputs**  
- `text` : *String* – the original text whose line endings are to be normalized.  
  *Passing mode*: ByVal (the function receives a copy; the caller’s variable is not modified).

**Outputs**  
- Returns a *String* containing the same characters as `text` but with every CRLF (`vbCrLf`) and CR (`vbCr`) replaced by a single LF (`vbLf`).

**Local State**  
- `result` : *String* – temporary variable used to hold intermediate results of the replacement operations.

**Procedure Steps**

1. **Declare local variable**  
   ```
   result : String
   ```

2. **Replace CRLF with LF**  
   ```
   result = Replace(text, vbCrLf, vbLf)
   ```
   *Effect*: All occurrences of the carriage‑return + line‑feed sequence (`vbCrLf`) in `text` are replaced with a single line‑feed (`vbLf`). The result is stored in `result`.

3. **Replace remaining CR with LF**  
   ```
   result = Replace(result, vbCr, vbLf)
   ```
   *Effect*: Any remaining carriage‑return characters (`vbCr`) that were not part of a CRLF pair are replaced with a line‑feed. The updated string is again stored in `result`.

4. **Return the normalized string**  
   ```
   return result
   ```

**Side Effects**  
- No external resources are accessed (no file, database, or workbook operations).  
- No global state is modified.  
- The function is pure: given the same input, it will always return the same output.

**Dependencies**  
- Built‑in string function `Replace(source, find, replace)` – performs a global, case‑sensitive replacement of all occurrences of `find` in `source` with `replace`.  
- Constants `vbCrLf`, `vbCr`, `vbLf` – predefined string constants representing carriage‑return + line‑feed, carriage‑return, and line‑feed respectively.

**Control Flow**  
The procedure executes sequentially: declaration → first replacement → second replacement → return. There are no branches, loops, or error handling constructs.

### Procedure: TryParseEntry (part 1/1)
**Purpose**  
Attempts to parse a single line of ledger data (`rawLine`) into a `LedgerEntry` object (`target`).  
If parsing succeeds, the function returns `True`; otherwise it returns `False` and leaves `target` unchanged.

**Parameters**  
| Name        | Passing | Type   | Description |
|-------------|---------|--------|-------------|
| `rawLine`   | ByVal   | String | The raw text line to parse. |
| `fallbackId`| ByVal   | Long   | Value to use for `EntryId` if the first field is not numeric. |
| `target`    | ByRef   | LedgerEntry | Object that will receive the parsed data. |

**Return type**  
`Boolean` – `True` if parsing succeeded, `False` otherwise.

**Local variables**  
| Name      | Type    | Purpose |
|-----------|---------|---------|
| `parts`   | Variant | Array of fields obtained by splitting `rawLine`. |
| `tagTokens` | Variant | Array of tag strings when a tag field is present. |
| `i`       | Long    | Loop counter for tag processing. |
| `rawAmt`  | Double  | Numeric value of the amount field before conversion to currency. |

**Error handling**  
`On Error GoTo ParseFail` – any runtime error during parsing causes an immediate jump to the `ParseFail` label, where the function returns `False`.

---

#### Pseudocode

```pseudocode
Procedure TryParseEntry(rawLine As String, fallbackId As Long, ByRef target As LedgerEntry) As Boolean
    // Enable error forwarding to the ParseFail label
    OnErrorGoto ParseFail

    // 1. Split the raw line into fields using the pipe character
    parts = Split(rawLine, "|")

    // 2. Validate that the split produced an array with at least 4 elements
    If Not IsArray(parts) Then
        GoTo ParseFail
    End If

    If UBound(parts) - LBound(parts) + 1 < 4 Then
        GoTo ParseFail
    End If

    // 3. Parse the EntryId field (first element)
    If IsNumeric(Trim(parts[LBound(parts)])) Then
        target.EntryId = CLng(Trim(parts[LBound(parts)]))
    Else
        target.EntryId = fallbackId
    End If

    // 4. Parse the AccountCode field (second element)
    target.AccountCode = UCase(Trim(CStr(parts[LBound(parts) + 1])))
    If Len(target.AccountCode) = 0 Then
        GoTo ParseFail
    End If

    // 5. Parse the Amount field (third element)
    rawAmt = CDbl(Trim(CStr(parts[LBound(parts) + 2])))
    target.Amount = CCur(rawAmt)

    // 6. Parse the PostedDate field (fourth element)
    If IsDate(Trim(CStr(parts[LBound(parts) + 3]))) Then
        target.PostedDate = CDate(Trim(CStr(parts[LBound(parts) + 3])))
    Else
        target.PostedDate = Date
    End If

    // 7. Default reconciliation status
    target.IsReconciled = False

    // 8. Optional Tags field (fifth element, if present and non‑empty)
    If UBound(parts) - LBound(parts) + 1 >= 5 And Len(Trim(CStr(parts[LBound(parts) + 4]))) > 0 Then
        // Split the tags string on commas
        tagTokens = Split(CStr(parts[LBound(parts) + 4]), ",")

        // Allocate the Tags array in the target object
        ReDim target.Tags[1 .. UBound(tagTokens) - LBound(tagTokens) + 1]

        // Populate the Tags array with lower‑case, trimmed tokens
        For i = LBound(tagTokens) To UBound(tagTokens)
            target.Tags[i - LBound(tagTokens) + 1] = LCase(Trim(CStr(tagTokens[i])))
        Next i
    Else
        // No tags supplied – use a single default tag
        ReDim target.Tags[1 .. 1]
        target.Tags[1] = "untagged"
    End If

    // 9. Parsing succeeded
    TryParseEntry = True
    Exit

### Procedure: NormalizeEntryInPlace (part 1/1)
**Purpose**  
Normalizes a `LedgerEntry` object in place. The procedure modifies the passed‑by‑reference `entry` by sanitizing its `AccountCode`, determining its reconciliation status based on `Amount`, and ensuring that the `Tags` array contains a `"core"` tag. No value is returned.

**Signature**  
```
PRIVATE SUB NormalizeEntryInPlace(ByRef entry As LedgerEntry)
```

**Local Variables**  
- `i` : Long – loop counter for iterating over `entry.Tags`.  
- `hasCoreTag` : Boolean – flag indicating whether a `"core"` tag was found.

**Pseudocode**

```
BEGIN NormalizeEntryInPlace(entry)

    // 1. Declare local variables
    DECLARE i AS Long
    DECLARE hasCoreTag AS Boolean

    // 2. Replace spaces and hyphens in the account code with underscores
    entry.AccountCode = Replace(entry.AccountCode, " ", "_")
    entry.AccountCode = Replace(entry.AccountCode, "-", "_")

    // 3. Determine reconciliation status based on Amount
    IF entry.Amount < 0 THEN
        entry.IsReconciled = False
    ELSE IF entry.Amount >= 10000 THEN
        entry.IsReconciled = False
    ELSE
        entry.IsReconciled = True
    END IF

    // 4. Initialize core‑tag flag
    hasCoreTag = False

    // 5. If the Tags array is defined and non‑empty, trim each tag
    //    and check for the literal string "core".
    IF (Not Not entry.Tags) <> 0 THEN
        FOR i = LBound(entry.Tags) TO UBound(entry.Tags)
            entry.Tags(i) = Trim$(entry.Tags(i))
            IF entry.Tags(i) = "core" THEN
                hasCoreTag = True
            END IF
        NEXT i
    END IF

    // 6. If no core tag was found, add one.
    IF NOT hasCoreTag THEN
        IF (Not Not entry.Tags) = 0 THEN
            // Tags array is empty or undefined – create a single‑element array
            ReDim entry.Tags(1 TO 1)
            entry.Tags(1) = "core"
        ELSE
            // Preserve existing tags and append a new element
            ReDim Preserve entry.Tags(1 TO UBound(entry.Tags) + 1)
            entry.Tags(UBound(entry.Tags)) = "core"
        END IF
    END IF

END NormalizeEntryInPlace
```

**Notes & Observations**

- `Replace` and `Trim$` are standard VBA string functions; they are called exactly as shown.
- The expression `(Not Not entry.Tags)` is a common VBA idiom for testing whether an array is defined and non‑empty. In this pseudocode it is treated as a boolean that is `True` when the array contains at least one element.
- `ReDim` without `Preserve` creates a new array, discarding any existing contents. `ReDim Preserve` expands the array while keeping existing elements.
- The procedure has no return value; all changes are made directly to the `entry` object passed by reference.
- The procedure is marked `PRIVATE`, meaning it is intended for internal use within its module only.

### Procedure: DeduplicateBatchEntries (part 1/1)
**Signature**  
```
PRIVATE SUB DeduplicateBatchEntries(ByRef batch As BatchSummary)
```
*Input*: `batch` – a reference to a `BatchSummary` object that contains an array of `LedgerEntry` objects (`batch.Entries`), a record count (`batch.RecordCount`), and a total amount (`batch.TotalAmount`).  
*Output*: The procedure modifies `batch` in place; it does not return a value.

---

#### Pseudocode

```
1  // Declare local variables
2  DECLARE seenKeys AS Object          // will hold a Scripting.Dictionary
3  DECLARE kept() AS LedgerEntry       // dynamic array for unique entries
4  DECLARE keptCount AS Long
5  DECLARE i AS Long
6  DECLARE key AS String
7  DECLARE newTotal AS Currency

8  // Attempt to create a dictionary; ignore any error that occurs
9  ON ERROR RESUME NEXT
10 SET seenKeys = CreateObject("Scripting.Dictionary")
11 ON ERROR GOTO 0

12 // If dictionary creation failed, exit immediately
13 IF seenKeys IS NOTHING THEN
14     EXIT SUB
15 END IF

16 // Allocate the temporary array to the original record count
17 REDIM kept(1 TO batch.RecordCount)

18 // Initialise counters
19 keptCount = 0
20 newTotal = 0@

21 // Iterate over every entry in the batch
22 FOR i = LBound(batch.Entries) TO UBound(batch.Entries)
23     // Build a unique key from EntryId and AccountCode
24     key = CStr(batch.Entries(i).EntryId) & "::" & batch.Entries(i).AccountCode

25     // If this key has not been seen before, keep the entry
26     IF NOT seenKeys.Exists(key) THEN
27         seenKeys.Add key, i
28         keptCount = keptCount + 1
29         kept(keptCount) = batch.Entries(i)
30         newTotal = newTotal + batch.Entries(i).Amount
31     ELSE
32         // Duplicate found – log the event
33         Log

### Procedure: RunLegacyGoSubAudit (part 1/1)

[GENERATION FAILED: ValueError('LLM returned an empty procedure section')]

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
Recursively determine the maximum depth of a dependency chain that starts at `startNode`.  
`edgesSpec` is a comma‑separated list of directed edges written as `"source>destination"`.  
The function returns the longest number of hops that can be reached from `startNode` without exceeding `maxDepth`.  

**Local Variables**  
```
pairs          : Variant          // array of edge strings split from edgesSpec
edgeParts      : Variant          // array of two strings split from a single edge
i              : Long             // loop index
srcNode        : String           // source node of the current edge
dstNode        : String           // destination node of the current edge
childDepth     : Long             // depth returned by a recursive call
maxChildDepth  : Long             // maximum depth found among all children
```

**Algorithm**

```
1.  // Early exit if recursion depth limit reached
2.  IF currentDepth >= maxDepth THEN
3.      RETURN currentDepth
4.  END IF

5.  // Early exit if input is invalid or empty
6.  IF Len(Trim$(edgesSpec)) = 0 OR Len(Trim$(startNode)) = 0 THEN
7.      RETURN currentDepth
8.  END IF

9.  // Split the edge specification string into individual edges
10. pairs = Split(edgesSpec, ",")

11. // Initialise the maximum depth found so far to the current depth
12. maxChildDepth = currentDepth

13. // Iterate over each edge string
14. FOR i = LBound(pairs) TO UBound(pairs)
15.     // Split the edge into source and destination parts
16.     edgeParts = Split(CStr(pairs(i)), ">")

17.     // Ensure the split produced exactly two parts
18.     IF UBound(edgeParts) - LBound(edgeParts) + 1 = 2 THEN
19.         // Trim whitespace from both parts
20.         srcNode = Trim$(CStr(edgeParts(LBound(edgeParts))))
21.         dstNode = Trim$(CStr(edgeParts(LBound(edgeParts) + 1)))

22.         // Check that the current edge originates from the start node
23.         IF StrComp(srcNode, Trim$(startNode), vbTextCompare) = 0
24.            AND Len(dstNode) > 0 THEN

25.             //

### Procedure: BuildUpsertSql (part 1/1)
**Purpose**  
Constructs a SQL `MERGE` statement that will upsert a ledger entry into the specified table.  
The function sanitizes identifiers, escapes single quotes, formats dates, and converts numeric values to the proper string representation for SQL.

**Signature**  
```
FUNCTION BuildUpsertSql(tableName AS String, entry AS LedgerEntry) AS String
```
* `tableName` – the name of the target table (passed ByVal).  
* `entry` – a `LedgerEntry` object containing the data to be inserted/updated (passed ByRef).  
* Returns a single `String` containing the complete `MERGE` statement.

**Local Variables**  
```
cleanTable  AS String   // sanitized table name
cleanAcct   AS String   // account code with single quotes escaped
cleanDate   AS String   // posted date formatted as yyyy-mm-dd
recFlag     AS Long     // 1 if entry.IsReconciled is true, otherwise 0
```

**Execution Flow**

1. **Sanitize Table Identifier**  
   ```
   cleanTable = SanitizeSqlIdentifier(tableName)
   ```
   *Calls the external routine `SanitizeSqlIdentifier` to ensure the table name is safe for SQL.*

2. **Escape Account Code**  
   ```
   cleanAcct = Replace(entry.AccountCode, "'", "''")
   ```
   *Replaces

### Procedure: SanitizeSqlIdentifier (part 1/1)
**Signature**  
```
PRIVATE FUNCTION SanitizeSqlIdentifier(ByVal ident As String) As String
```

**Inputs**  
- `ident` – a string supplied by the caller, passed by value.

**Outputs**  
- Returns a string that contains only characters allowed in a SQL identifier:  
  - Upper‑case letters `A`–`Z`  
  - Lower‑case letters `a`–`z`  
  - Digits `0`–`9`  
  - Underscore `_`  
  - Period `.`  
  If no such characters are present in `ident`, the function returns the literal `"default_ledger_table"`.

**State Changes & Side Effects**  
- Local variables `i`, `ch`, and `outStr` are created and modified only within the function.  
- No external state (files, databases, workbook objects, etc.) is accessed or modified.  
- The function has no observable side effects beyond its return value.

**Pseudocode**

```
BEGIN FUNCTION SanitizeSqlIdentifier
    // Local variable declarations
    DECLARE i AS Long
    DECLARE ch AS String
    DECLARE outStr AS String

    // Initialize the output string to empty
    outStr = ""

    // Iterate over each character of the input string
    FOR i FROM 1 TO Len(ident) DO
        // Extract the i‑th character (1‑based indexing)
        ch = Mid$(ident, i, 1)

        // Check if the character is allowed
        IF (ch >= "A" AND ch <= "Z") OR
           (ch >= "a" AND ch <= "z") OR
           (ch >= "0" AND ch <= "9") OR
           ch = "_" OR
           ch = "." THEN
            // Append allowed character to the output
            outStr = outStr & ch
        END IF
    NEXT i

    // If no allowed characters were found, use the default name
    IF Len(outStr) = 0 THEN
        outStr = "default_ledger_table"
    END IF

    // Return the sanitized identifier
    SanitizeSqlIdentifier = outStr
END FUNCTION
```

**Explanation of Control Flow**

1. **Initialization** – `outStr` starts as an empty string.  
2. **Loop** – For each position `i` from 1 to the length of `ident`:
   - `ch` is set to the single character at that position.
   - The `IF` condition tests whether `ch` falls within the allowed ranges or equals `_` or `.`.
   - If the condition is true, `ch` is concatenated to `outStr`; otherwise it is discarded.
3. **Post‑loop Check** – After the loop, if `outStr` remains empty, it is replaced with the literal `"default_ledger_table"`.
4. **Return** – The function assigns `outStr` to its own name, which is the return value in VBA.

No error handling constructs (e.g., `On Error`) are present in this fragment. All operations are deterministic and purely functional with respect to the input string.

### Procedure: BuildSummaryReport (part 1/1)
**Purpose**  
Builds a textual summary of a batch processing operation and returns it as a single string.  
The summary includes batch identifiers, counts, totals, audit status, and optionally recent audit log entries.

**Signature**  
```
PRIVATE FUNCTION BuildSummaryReport(batch As BatchSummary, parsedCount As Long, statusMsg As String) As String
```

| Parameter | Passing | Type | Description |
|-----------|---------|------|-------------|
| `batch`   | ByRef   | `BatchSummary` | The batch object whose properties are displayed. |
| `parsedCount` | ByVal | `Long` | Number of rows parsed from the source. |
| `statusMsg` | ByVal | `String` | Audit status message to include in the summary. |

**Local Variables**  
- `sb` (String) – accumulator for the summary text.  
- `i` (Long) – loop counter for audit log entries.  
- `logLimit` (Long) – maximum number of audit log entries to include (capped at 15).

**External Dependencies**  
- `m_AuditLog` – module‑level collection or array of audit log strings.  
- `vbCrLf` – constant for a carriage‑return line‑feed.  
- `CStr` – function converting a value to its string representation.  
- `Format$` – function formatting a numeric value according to a format string.

**Control Flow & Side‑Effects**

1. **Initialize the summary header**  
   ```pseudocode
   sb = "=== BATCH SUMMARY ===" + vbCrLf
   ```

2. **Append batch details**  
   ```pseudocode
   sb = sb + "Batch ID:      " + CStr(batch.BatchId) + vbCrLf
   sb = sb + "Mode Mask:     " + CStr(batch.ModeMask) + vbCrLf
   sb = sb + "Parsed Rows:   " + CStr(parsedCount) + vbCrLf
   sb = sb + "Final Rows:    " + CStr(batch.RecordCount) + vbCrLf
   sb = sb + "Total Amount:  " + Format$(batch.TotalAmount, "#,##0.00") + vbCrLf
   sb = sb + "Audit Status:  " + statusMsg + vbCrLf
   ```

3. **Optional audit log section**  
   ```pseudocode
   IF (m_AuditLog IS NOT Nothing) AND (m_AuditLog.Count > 0) THEN
       sb = sb + "--- RECENT AUDIT LOGS ---" + vbCrLf
       logLimit = m_AuditLog.Count
       IF logLimit > 15 THEN
           logLimit = 15
       END IF
       FOR i = 1 TO logLimit
           sb = sb + CStr(m_AuditLog(i)) + vbCrLf
       NEXT i
   END IF
   ```

4. **Return the accumulated string**  
   ```pseudocode
   BuildSummaryReport = sb
   ```

**Observations**

- The function **does not modify** any of its input parameters (`batch`, `parsedCount`, `statusMsg`) or the global `m_AuditLog`.  
- All string concatenations are performed in the order shown, preserving the exact sequence of lines.  
- The audit log section is included only when `m_AuditLog` is defined and contains at least one entry; otherwise it is omitted.  
- The maximum number of audit log entries appended is capped at 15 to avoid excessive output.  
- No error handling or early exits are present; the function assumes all referenced properties and collections are valid.  
- The function’s sole observable side effect is the returned string; it does not write to files, modify the UI, or alter global state beyond reading `m_AuditLog`.

### Procedure: GetAuditEntryCount (part 1/1)
**Signature**  
```
FUNCTION GetAuditEntryCount() RETURNS Long
```

**Description**  
`GetAuditEntryCount` is a public function that reports the number of audit entries stored in the module‑level collection `m_AuditLog`.  
- If `m_AuditLog` has not been initialized (`Nothing`), the function returns `0`.  
- Otherwise it returns the value of `m_AuditLog.Count`, which represents the current number of audit entries.

The function performs only a read of the global variable `m_AuditLog`; it does not modify any state or produce side effects beyond its return value.

**Pseudocode**

```
FUNCTION GetAuditEntryCount() RETURNS Long
    // Evaluate whether the audit log collection has been created.
    IF m_AuditLog IS Nothing THEN
        // No audit log exists; report zero entries.
        GetAuditEntryCount = 0
    ELSE
        // Audit log exists; report its current count.
        GetAuditEntryCount = m_AuditLog.Count
    END IF
END FUNCTION
```

**Notes on Behavior**

- `m_AuditLog` is a module‑level variable; its type is assumed to be a collection or similar object that exposes a `Count` property.  
- The function assigns to its own name (`GetAuditEntryCount`) to set the return value, following VBA semantics.  
- No error handling or exception paths are present in this fragment.  
- The function has no parameters and returns a 64‑bit signed integer (`Long`).

### Procedure: ResetModuleState (part 1/1)
**Purpose**  
`ResetModuleState` is a public subroutine that reinitializes the module‑level state used by the application.  
It performs three assignments in sequence:

1. **Reset the global sequence counter** – sets the numeric variable `m_GlobalSequence` to `0`.  
2. **Create a fresh audit log** – assigns a new instance of the `Collection` object to the variable `m_AuditLog`.  
3. **Clear the node registry** – sets the object reference `m_NodeRegistry` to `Nothing`, effectively discarding any previously stored nodes.

No parameters are accepted and the procedure does not return a value.  
The routine has no error handling; any runtime errors that occur during the assignments will propagate to the caller.

---

#### Pseudocode

```
PROCEDURE ResetModuleState
    // 1. Reset the global sequence counter
    m_GlobalSequence ← 0

    // 2. Create a new audit log collection
    m_AuditLog ← NEW Collection

    // 3. Clear the node registry reference
    m_NodeRegistry ← Nothing
END PROCEDURE
```

**Notes**

- `m_GlobalSequence` is a module‑level numeric variable that tracks a global sequence number.  
- `m_AuditLog` is a module‑level variable of type `Collection`; a new empty collection is created each time the procedure runs.  
- `m_NodeRegistry` is a module‑level object reference that holds node objects; setting it to `Nothing` releases any existing reference.  
- The assignments are executed in the order shown; there is no branching or looping.  
- The procedure is intended to be called whenever the module needs to be reset to a clean state, such as at application start or after a major operation.
