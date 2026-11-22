# Semantic Analyzer — Universal IR Normalization

**Version:** `0.1.0`  
**Python:** 3.11+  
**Dependencies:** Python standard library only

This component normalizes the existing language-neutral Universal IR before a target backend (Java first) is allowed to generate code.

## Outputs

- `normalized_ir.json` — structured expression trees, resolved identifiers and UDT fields, call classifications, explicit array/index access, explicit `ReDim` bounds, `ByRef` call contracts, and explicit VBA default `For` step values.
- `semantic_analysis.json` — parser/validation results, unresolved symbols, graph/segment provenance checks, pseudocode coverage, frontend diagnostics, and target-runtime compatibility requirements.

## Run from the repository root

```bash
python -m vba_to_java.semantic_analyzer
```

The entry point processes sample 6. It reads `universal_ir.json`, `graph.json`,
and `segments.json` from `files/vba_graph_output/sample_6/` and writes
`normalized_ir.json` and `semantic_analysis.json` directly into that same
folder; it does not create a nested output directory. Exit code `2` means
conversion is blocked pending review; output artifacts are still written so the
blockers can be inspected.

## What the pass does

1. Parses source-backed opaque expression nodes with a deterministic VBA expression parser.
2. Resolves local, parameter, module, procedure, enum-member, and UDT-field references where declarations support resolution.
3. Distinguishes array indexing from function invocation and VBA Collection default-item access.
4. Classifies VBA built-ins, project procedure calls, and late-bound method calls.
5. Preserves parameter passing mode and checks whether `ByRef` call arguments are mutable lvalues.
6. Models assignment to a VBA Function's own name as an implicit return-value assignment.
7. Parses `ReDim` bounds and preserves `Preserve`; supplies VBA's default `For` step of `1` with an explicit language-rule inference marker.
8. Checks labels, `GoSub`/`Return`, and `On Error` handler references.
9. Verifies graph-node and granular-segment provenance. Full `MODULE` and `PROCEDURE` segment snapshots are excluded; graph module/procedure source snapshots are not read.
10. Reports compatibility requirements instead of silently translating source-language runtime behavior.

## Safety and accuracy policy

- No LLM is called in this version. No API key is required.
- No unsupported expression is replaced with a guessed AST. Parse failures remain explicit and block conversion.
- Pseudocode is supporting evidence and coverage metadata, not source-of-truth logic.
- Existing graph hierarchy repairs and `CROSS_NESTED_BLOCK_CLOSE` diagnostics are not auto-approved.
- A successful expression parse is not proof of equivalent VBA runtime behavior.
- A structurally valid normalized IR is not proof of control-flow correctness or behavioural equivalence.

## Current parser subset

Identifiers, literals (including VBA Currency suffix and date literals), member access, parenthesized calls, array/index expressions, nested calls, unary operators, VBA binary operators, parentheses, and statement-style calls with omitted parentheses.

Unsupported syntax is explicitly reported. Extend the parser and tests before using it on additional legacy constructs.

## Sample 6 status

The Sample 6 run is **blocked**, intentionally:

- 15 procedures retained.
- 926 expression roots normalized; zero unresolved/unsupported expression nodes and zero parser errors.
- 330 graph provenance references matched; no missing graph-node references.
- 237 granular segments available; 228 referenced by IR nodes.
- 24 frontend hierarchy repairs remain unreviewed.
- Three `CROSS_NESTED_BLOCK_CLOSE` warnings remain at source lines 223, 256, and 472.
- Four pseudocode procedure sections are missing: `ExecuteStressPipeline`, `LogAudit`, `TryParseEntry`, and `NormalizeEntryInPlace`.
- Five target-compatibility blockers are reported: VBA error handling, legacy GoSub/Return, late-bound Object members, COM `CreateObject`, and the double-`Not` dynamic-array initialization idiom.

The report also identifies review requirements for `ByRef`, `Option Compare Text`, `Option Base 1`, `Variant`, `Currency`, `Date`, dynamic arrays, `Collection`, `IIf` eager evaluation, `DoEvents`, object identity, and numeric bitwise operators.

Do not generate production Java from this sample until the blocking items have explicit, tested target-runtime decisions and the frontend control-flow warnings/repairs have been dispositioned.

The lower-level CLI remains available for explicitly selecting input files and
an output directory:

```bash
python -m vba_to_java.semantic_analyzer.cli \
  --ir files/vba_graph_output/sample_6/universal_ir.json \
  --graph files/vba_graph_output/sample_6/graph.json \
  --segments files/vba_graph_output/sample_6/segments.json \
  --pseudocode files/vba_graph_output/sample_6/pseudocode.md \
  --manifest files/vba_graph_output/sample_6/manifest.json \
  --out-dir files/vba_graph_output/sample_6
```

The package includes sanitized Sample 6 test fixtures. The fixture graph omits full module/procedure source snapshots; the fixture segments omit `PROCEDURE` snapshots.
