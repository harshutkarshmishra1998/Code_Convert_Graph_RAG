# VBA to Java Conversion

This submodule uses the original `.bas` file and its `process_notes.md`. It refreshes notes when the note format, source hash, procedure list, deterministic calculation ledger, LLM calculation coverage, or LLM status is incomplete. It then asks the configured LLM to generate Java using both files as references.

Run from the repository root:

```sh
python -m vba_to_java_2.java_conversion.main
```

The current `n = [6]` setting converts the stress sample only. Results go to `files/vba/output/sample_6/`.

The converter makes up to five generation/repair attempts. It asks the LLM for Java source text directly, without wrapping the code in JSON. Java contains many quotes and escapes; removing the JSON wrapper avoids provider failures while encoding long source files. The prompt requires unresolved behavior and assumptions to be called out in `// UNRESOLVED:` and `// ASSUMPTION:` comments, initializes Java locals to VBA type defaults, and keeps `ByRef` representations consistent at callers and callees. The Java source must use one public top-level class, with helpers and Excel adapters nested inside it, so it can compile as a single `.java` file. Repair prompts include the previous Java source and exact audit failures. When `javac` identifies a negated string concatenation before `equalsIgnoreCase`, the converter applies a narrow parenthesization fix and recompiles. It writes a `.java` file only if the output declares the expected class, contains a Java method declaration for every VBA procedure, has no unresolved/assumption markers, and compiles with `javac`. It stops early when a repair repeats the same failures without progress. On failure it writes `conversion_report.md` and stops the batch.

For `openai/gpt-oss-20b`, generation allows up to 60,000 completion tokens (below Groq's documented 65,536-token model maximum) because complete source code is returned in one response. The prompt also asks for concise code to reduce truncation risk.

Excel operations must be represented by explicit Java adapter interfaces. Those interfaces describe what a caller must connect to a workbook implementation; they are not a working Excel integration by themselves. If an operation cannot be represented safely through the adapter, conversion remains blocked and the report explains the unresolved behavior.

The deterministic calculation ledger is provided to the LLM as source context, but this stage does not prove that each formula is translated exactly. Compilation and method-presence checks do not prove behavioral equivalence. The report says `NOT VERIFIED` until a reference test corpus executes both the original VBA and Java and compares results and side effects. Exact 100% behavior cannot be claimed from an LLM response or successful Java compilation alone. VBA-specific features such as `Currency`, `Variant`, `ByRef`, dynamic arrays, error handling, and `GoSub` need particular review in sample 6.
