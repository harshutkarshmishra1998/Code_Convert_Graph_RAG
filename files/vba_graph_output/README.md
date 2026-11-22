# VBA analysis and LLM handoff outputs

Each sample directory contains flat, same-folder outputs for the matching
`.bas` module:

| Input | Output directory |
|---|---|
| [vba_sample_1.bas](../vba_sample_1.bas) | `sample_1/` |
| [vba_sample_2.bas](../vba_sample_2.bas) | `sample_2/` |
| [vba_sample_3.bas](../vba_sample_3.bas) | `sample_3/` |
| [vba_sample_4.bas](../vba_sample_4.bas) | `sample_4/` |
| [vba_sample_5.bas](../vba_sample_5.bas) | `sample_5/` |
| [vba_sample_6.bas](../vba_sample_6.bas) | `sample_6/` |

These are analysis and handoff artifacts, not translated application code.
The handoff generator does not make an external LLM request.

## Files and their purpose

### Frontend evidence

- **`graph.json`** — procedure and statement nodes, parent/child structure,
  source locations, symbols, control-flow relationships, and conservative
  call/reference links. Use it to inspect structure and ordering.
- **`segments.json`** — source-preserving procedure, block, and statement
  chunks. Match `id` and `parent_id` values to graph nodes; use the chunks as
  grounded input for pseudocode or code generation.
- **`report.json`** — graph counts and parser diagnostic/opaque-statement
  counts. It is a quick quality check, not a behavior specification.
- **`manifest.json`** — input path, encoding, byte size, and SHA-256. It
  verifies which exact source version produced the artifacts.

### Language-neutral conversion evidence

- **`universal_ir.json`** — structured procedures, declarations, expressions,
  statements, and control flow, with source provenance. It is the primary
  target-independent starting point for code generation.
- **`normalized_ir.json`** — Universal IR after deterministic expression,
  symbol, type, and control-flow normalization. Unsupported or unresolved
  constructs remain explicit; they are not guessed.
- **`semantic_analysis.json`** — readiness status, unresolved-expression and
  pseudocode-coverage summaries, compatibility requirements, provenance
  checks, and semantic issues. A `blocked` status means the target needs
  compatibility work or missing evidence; it does not mean the files failed
  to generate.

### Optional LLM-generated pseudocode

- **`pseudocode.md`** — generated module/procedure explanations, if available.
- **`pseudocode_run.json`** — generation attempt and coverage status, if
  available.

Pseudocode is optional and is never the sole source of truth. Check the
semantic report's missing-procedure list before relying on it.

### Intelligent-fixer audit

- **`intelligent_fixer_log.json`** — per-run issue detection and resolution
  status, including whether the LLM adapter was enabled and whether each issue
  was `addressed_and_verified` or remains unresolved.
- **`final_quality_report.json`** — the final structural, semantic, regression,
  compile, and behavioral validation summary.
- **`unresolved_issues.json`** and **`repair_history.json`** — current unresolved
  findings and append-only run history.

The checked-in reports were refreshed in offline, no-auto-apply mode. The
opaque-IR finding is marked `addressed_and_verified` only where normalized IR
and semantic validation prove zero unresolved nodes and parse errors. No
generated-file patches were applied; other findings remain explicitly
unresolved.

### Recommended LLM input

- **`llm_handoff.json`** — a compact project index with exact source, module
  declarations and procedure signatures, artifact hashes, semantic readiness,
  compatibility blockers, intelligent-fixer resolution statuses, and links to
  procedure context files.
- **`llm_handoff_<Procedure>.json`** — a per-procedure package with the exact
  procedure source, normalized IR, provenance, callee references, relevant
  semantic report, and pseudocode section when available. Keeping procedures
  separate avoids forcing a model to ingest the entire project at once.
- **`llm_handoff.md`** — a concise readiness summary and instructions for using
  the index and procedure files.

Give an LLM `llm_handoff.json`, the relevant procedure file(s), and the target
language/runtime. Include a callee's procedure file when translating a caller.
The original VBA source remains authoritative; the graph, segments, IR, and
semantic analysis provide structure and expose uncertainty. Ask the LLM to
preserve behavior, list assumptions, and call out unresolved requirements
rather than inventing APIs or silently dropping unsupported behavior.

## Which files are needed?

- **To write pseudocode:** use `graph.json`, `segments.json`, and the original
  `.bas` source. The graph provides control-flow structure; segments provide
  exact source text and line spans.
- **To generate code in another language:** use `llm_handoff.json` plus the
  relevant `llm_handoff_<Procedure>.json` file(s). They include the source,
  normalized IR, module declarations, procedure signatures, semantic
  constraints, provenance, and optional pseudocode. The original graph,
  segments, Universal IR, and normalized IR files remain available beside the
  handoff files for deeper inspection.
- **To validate freshness:** compare the source hash in `manifest.json` with
  the source. The handoff generator performs this check and refuses stale
  inputs.

## Current graph results

Counts are from the six checked-in samples after the parser fixes. The
diagnostic count is zero for each sample; opaque statements are also zero.

| Sample | Procedures | Nodes | Edges | Segments | Diagnostics |
|---|---:|---:|---:|---:|---:|
| `sample_1` | 1 | 29 | 46 | 24 | 0 |
| `sample_2` | 1 | 23 | 32 | 16 | 0 |
| `sample_3` | 1 | 59 | 88 | 46 | 0 |
| `sample_4` | 1 | 83 | 140 | 71 | 0 |
| `sample_5` | 2 | 95 | 150 | 77 | 0 |
| `sample_6` | 15 | 323 | 502 | 243 | 0 |

The semantic analyzer reports zero unresolved expressions for every sample
after normalization. Migration readiness is still `blocked` because some
samples use VBA-specific error handling, COM/late-bound objects, `GoSub`,
arrays, or other compatibility behaviors that must be deliberately implemented
in the target runtime. Sample 6 also has four procedures without detailed
pseudocode sections; the precise names are recorded in its semantic report and
handoff bundle.

## Regenerating outputs

The per-input `main(input_path, output)` functions are used internally, so no
terminal arguments are needed. Rebuild graph, Universal IR, and semantic
artifacts for samples 1–6 from the repository root:

```python
from pathlib import Path

from vba_to_java.vba_graph.cli import main as build_graph
from vba_to_java.universal_ir.cli import main as build_ir
from vba_to_java.semantic_analyzer.cli import main as analyze

root = Path.cwd()
n = [1, 2, 3, 4, 5, 6]

for i in n:
    source = root / "files" / f"vba_sample_{i}.bas"
    output = root / "files" / "vba_graph_output" / f"sample_{i}"
    build_graph(source, output)
    build_ir(source, output)
    analyze(source, output)
```

Generate or refresh the project index, procedure bundles, and Markdown summary
for every sample with:

```sh
python -m vba_to_java.llm_handoff
```
