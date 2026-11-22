# VBA graph → human-readable pseudocode

Place this folder at `PROJECT_ROOT/vba_to_java/pseudocode/`.

## Expected input

```text
PROJECT_ROOT/
├── api_keys.py
├── files/
│   └── vba_output/
│       ├── sample_1/
│       │   ├── graph.json
│       │   ├── segments.json       # optional
│       │   ├── report.json         # optional
│       │   └── manifest.json       # optional
│       └── sample_2/
│           └── graph.json
└── vba_to_java/
    └── pseudocode/
        ├── generate_pseudocode.py
        └── README.md
```

The script reads the graph's `nodes`, `edges`, `symbols`, and `diagnostics`. It uses
node-level source fragments as evidence and does not use a complete source snapshot as
a substitute for graph structure. `segments.json`, `report.json`, and `manifest.json`
are optional; the current implementation intentionally derives the pseudocode from the
graph, since that is the structural input for later translation.

## Run

From `PROJECT_ROOT` with the same Python environment used by your project:

```bash
python vba_to_java/pseudocode/generate_pseudocode.py
```

One sample only:

```bash
python vba_to_java/pseudocode/generate_pseudocode.py --sample sample_6
```

Custom input path:

```bash
python vba_to_java/pseudocode/generate_pseudocode.py --input files/vba_output
```

## LLM client

The script loads exactly `api_keys._loaded["LLM"]`. It supports an OpenAI-compatible
client (`chat.completions.create` or `responses.create`), a LangChain-style object with
`invoke()`, or a callable wrapper. For OpenAI-compatible clients, set `LLM_MODEL` if the
client does not expose `model_name` or `model`:

```bash
export LLM_MODEL="your-enabled-model-name"
```

## Output

Each sample directory receives:

- `pseudocode.md` — human-readable, language-neutral pseudocode.
- `pseudocode_run.json` — counts, diagnostics, generation status, and failed sections.

The input root receives `pseudocode_batch_summary.json`.

Large procedures are split at graph subtree boundaries; no body text is silently
truncated. Failed LLM calls are recorded and retried. The output is a translation aid,
not proof of semantic equivalence. Review parser diagnostics and `[UNCERTAIN]` / `[OPAQUE]`
markers against the original VBA before generating an intermediate representation.
