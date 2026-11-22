# Intelligent fixer

Run the sample 6 audit from the repository root:

```bash
python -m vba_to_java.intelligent_fixer
```

The entry point uses `files/vba_sample_6.bas` and reads/writes artifacts in
`files/vba_graph_output/sample_6/`. It does not create a nested report folder.

By default, the fixer performs local diagnostics only. To enable ChatGroq
planning and broad artifact audits, explicitly set
`INTELLIGENT_FIXER_USE_LLM=1`. This sends the selected sample's artifacts and
source-backed context to the configured LLM provider. To permit reviewed,
hash-bound patches to generated artifacts, also set
`INTELLIGENT_FIXER_APPLY=1`. Original `.bas` source files remain immutable.

The same-folder logs are:

- `intelligent_fixer_log.json`: each run, artifacts checked, detected issues,
  final addressed/unresolved status, patch details, and verification of the
  previously reported pseudocode LLM-client mismatch. Pseudocode coverage gaps
  remain separate unresolved findings even if the run log says `success`.
- `repair_history.json`: append-only detection, patch, and verification events.
- `unresolved_issues.json`: issues still present or not verified.
- `final_quality_report.json`: validation results and conversion readiness.

An issue is marked `addressed_and_verified` only when an applied patch is absent
from the final issue scan. A clean local audit is not certification: compilation
and behavioral-equivalence checks require configured runners/oracles.
