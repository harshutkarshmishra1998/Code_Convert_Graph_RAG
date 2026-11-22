# VBA conversion outputs and Java runners

Run the complete local pipeline from the repository root:

```sh
python3 -m vba_to_java.main
```

It processes `vba_sample_1.bas` through `vba_sample_6.bas` sequentially and writes
the graph, intermediate reports, and `vba_sample_N.java` directly into
`files/vba_graph_output/sample_N/`. It does not create per-stage subfolders.
Java generation is deterministic and offline; pseudocode generation is recorded
as skipped because that stage requires an external LLM. A JDK 21+ and Maven are
required for the final Apache POI project compile.

## JSON artifacts

| File | Relevance |
| --- | --- |
| `manifest.json` | Source file name, encoding, and SHA-256 used to reject stale analysis artifacts. |
| `graph.json` | VBA control/data-flow graph nodes and edges. Useful for navigation and graph-based analysis. |
| `segments.json` | Ordered source segments with source locations and structural classifications. Useful for tracing intermediate output back to VBA. |
| `universal_ir.json` | Front-end intermediate representation produced from the graph. |
| `normalized_ir.json` | Canonical procedure, statement, and expression representation; primary structured input for pseudocode or code-conversion stages. |
| `semantic_analysis.json` | Normalized IR plus procedure/type/binding analysis and compatibility requirements. Use with `normalized_ir.json` to guide conversion and identify VBA-specific behavior that needs attention. |
| `report.json` | Graph-stage diagnostics and summary. |
| `unresolved_issues.json` | Remaining parsing or semantic items that need review. |
| `llm_handoff.json`, `llm_handoff_<Procedure>.json`, `llm_handoff.md` | Source-backed context bundles for optional, manually run LLM work. They do not mean an LLM was called by the local pipeline. |
| `intelligent_fixer_log.json`, `final_quality_report.json`, `repair_history.json` | Offline fixer findings, readiness assessment, and repair history. The pipeline does not auto-apply LLM changes. |
| `java_codegen_report.json` | Source freshness/procedure checks, Java compilation status, and known behavior limitations. |
| `pipeline_run.json` | Per-sample stage results and final compile status. |

For pseudocode, start with `normalized_ir.json` and `semantic_analysis.json`.
Include the matching `.bas` source and `segments.json` when exact source wording
or locations matter. `graph.json` is optional context; it is not a substitute
for the normalized procedure bodies.

## Generated Java files

| Sample | Java behavior | Runtime input |
| --- | --- | --- |
| 1 | Copies valid transaction rows into `Monthly Summary`, highlights invalid dates in yellow, and sizes the report columns. | Input workbook with `Transactions` and `Monthly Summary` sheets; optional output `.xlsx`. |
| 2 | Validates and normalizes one SKU. | One SKU string argument. |
| 3 | Reads UTF-8 TSV sales, aggregates by region, and updates matching `Daily Sales` rows for a report date. | Workbook, TSV path, ISO report date (`yyyy-MM-dd`), optional output `.xlsx`. |
| 4 | Consolidates the first worksheet's columns A:F from `.xlsx` files, retaining the first header only. | Input folder and optional output `.xlsx`; files are processed by name. |
| 5 | Reconciles inventory against movements from the preceding 30 days, logs handled errors when an `Error Log` sheet exists, and sorts inventory. | Workbook with `Inventory` and `Movements` sheets; optional output `.xlsx`. |
| 6 | Runs the ledger parsing, normalization, deduplication, aggregation, audit, dependency-depth, SQL-building, and summary-report routines. | UTF-8 payload text file, mode mask, optional maximum iteration count. |

The generated sources and the shared POI helper are compiled together by the
pipeline's Maven stage. After running the pipeline, run a workbook-based sample
with the resolved Maven dependencies on the classpath, for example:

```sh
mvn -q dependency:build-classpath -Dmdep.outputFile=target/dependencies.txt
java -cp "target/classes:$(cat target/dependencies.txt)" vba_sample_1 input.xlsx output.xlsx
```

The synthetic behavior smoke tests can be run with:

```sh
mvn -q test-compile
mvn -q dependency:build-classpath -Dmdep.outputFile=target/dependencies.txt
java -cp "target/classes:target/test-classes:$(cat target/dependencies.txt)" VbaTemplatesSmokeTest
```

Compilation is not proof of behavioral equivalence. Reports intentionally keep
`accuracy_status` as `NOT_VERIFIED`: the smoke tests use synthetic inputs, and
there has not been a differential run against a VBA/Excel host or the user's
real workbooks. See each sample's `java_codegen_report.json` for known limits.
