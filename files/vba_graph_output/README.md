# VBA graph JSON outputs

This folder contains one output directory per input module:

| Input | Output directory |
|---|---|
| [vba_sample_1.bas](../vba_sample_1.bas) | `sample_1/` |
| [vba_sample_2.bas](../vba_sample_2.bas) | `sample_2/` |
| [vba_sample_3.bas](../vba_sample_3.bas) | `sample_3/` |
| [vba_sample_4.bas](../vba_sample_4.bas) | `sample_4/` |
| [vba_sample_5.bas](../vba_sample_5.bas) | `sample_5/` |
| [vba_sample_6.bas](../vba_sample_6.bas) | `sample_6/` |

Each output directory contains `graph.json`, `segments.json`, `report.json`, and
`manifest.json`. These are analysis artifacts; the converter does not generate
pseudocode or Java code.

## What each JSON file is for

### `graph.json` - structure and relationships

This is the main structured representation of a VBA project. It includes:

- Project and module nodes, procedure nodes, and statement/control-flow nodes.
- Node IDs, parent IDs, source paths, line numbers, source text, and metadata.
- Edges such as `CONTAINS`, `NEXT`, branch relationships, and conservative
  `CALLS`/`REFERENCES` links.
- `symbols` for parameters and locally identified variables.
- `diagnostics` for parser warnings, plus the manifest in the graph.

Use it to understand procedure boundaries, nesting, ordering, branches, loops,
and possible procedure calls. Semantic links are best-effort hints, not a
complete VBA type or call analysis.

### `segments.json` - source chunks for downstream processing

This is a flat list of source-preserving chunks. Each segment has an ID, parent
ID, kind, name, source path, line range, and source text. Procedure segments
contain whole procedures; other segments represent statements and blocks.
`summary` and `annotations` are available for enrichment, but are empty by
default.

Use this file to feed procedures or smaller source units into a pseudocode
generator while retaining their graph IDs and line locations.

### `report.json` - conversion counts and quick quality check

This contains counts for modules, procedures, nodes, edges, segments,
diagnostics, and opaque statements. Use it to compare runs and decide whether
parser warnings need review. It is not needed to construct pseudocode.

### `manifest.json` - input provenance

This records each input file's path, detected encoding, byte size, and SHA-256
hash. Use it to trace generated graph data back to the exact input version. It
does not describe VBA behavior and is not needed to construct pseudocode.

## Which files to use for pseudocode

Use **both `graph.json` and `segments.json`** as the intermediate input:

1. Use `graph.json` for procedure structure, parent/child relationships,
   control-flow branches, statement ordering, symbols, and possible calls.
2. Use `segments.json` for the source text and line spans of each procedure and
   statement. Match segments to graph nodes using their IDs and parent IDs.
3. Use the original `.bas` file to verify any source represented as opaque or
   flagged by diagnostics.

`report.json` and `manifest.json` are useful for validation and provenance, but
are optional for pseudocode generation. Do not rely on summaries: they are
`null` unless added by a separate annotation step. Preserve unknown statements
instead of silently dropping them; this parser is conservative and does not
implement the full VBA language.

## Samples included

The figures below come from converting all six supplied samples. Node, edge,
and segment counts describe the generated graph; diagnostic counts indicate
warnings, not necessarily conversion failure.

| Sample | Procedures | Nodes | Edges | Segments | Diagnostics |
|---|---:|---:|---:|---:|---:|
| `sample_1` | 1 | 30 | 48 | 25 | 1 |
| `sample_2` | 1 | 23 | 32 | 16 | 0 |
| `sample_3` | 1 | 61 | 92 | 48 | 0 |
| `sample_4` | 1 | 83 | 140 | 71 | 1 |
| `sample_5` | 2 | 98 | 156 | 80 | 0 |
| `sample_6` | 15 | 331 | 520 | 252 | 3 |

Procedure names represented in the samples:

- `sample_1`: `RefreshMonthlySummary`
- `sample_2`: `NormalizeSku`
- `sample_3`: `ImportDailySales`
- `sample_4`: `ConsolidateDepartmentFiles`
- `sample_5`: `ReconcileInventory`, `LogInventoryError`
- `sample_6`: `ExecuteStressPipeline`, `InitInfrastructure`, `NextSequence`,
  `LogAudit`, `NormalizeLineEndings`, `TryParseEntry`,
  `NormalizeEntryInPlace`, `DeduplicateBatchEntries`, `RunLegacyGoSubAudit`,
  `EvaluateDependencyDepth`, `BuildUpsertSql`, `SanitizeSqlIdentifier`,
  `BuildSummaryReport`, `GetAuditEntryCount`, `ResetModuleState`

Samples 1 and 4 each have one `OPAQUE_STATEMENT` warning for an `.AutoFit`
statement. Sample 6 has three `CROSS_NESTED_BLOCK_CLOSE` warnings; review those
locations against the original source before treating its control-flow graph as
authoritative.

## Regenerating all six output directories

From the repository root, call the converter's per-input `main(input, output)`
function for each sample:

```python
from pathlib import Path

from vba_to_java.vba_graph.cli import main

inputs = Path("files")
outputs = inputs / "vba_graph_output"
n = [1, 2, 3, 4, 5, 6]

for i in n:
    main(
        inputs / f"vba_sample_{i}.bas",
        outputs / f"sample_{i}",
    )
```
