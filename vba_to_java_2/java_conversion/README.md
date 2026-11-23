# VBA to Java Conversion

This submodule uses the original `.bas` file and its `process_notes.md`. It refreshes notes when the note format, source hash, procedure list, deterministic calculation ledger, LLM calculation coverage, or LLM status is incomplete. It then asks the configured LLM to generate Java using both files as references.

Run from the repository root:

```sh
python -m vba_to_java_2.java_conversion.main
```

The current `n = [6]` setting converts the stress sample only. Results go to `files/vba/output/sample_6/`.

The converter makes up to five generation/repair attempts. The response schema is intentionally small: it asks for Java source, unresolved items, and assumptions. This avoids large repeated calculation tables, which some LLM providers reject as invalid structured JSON. Repair prompts include the previous Java source and exact audit failures. It writes a `.java` file only if the output declares the expected class, contains a Java method declaration for every VBA procedure, reports no unresolved semantics or assumptions, and compiles with `javac`. On failure it writes `conversion_report.md` and stops the batch.

The deterministic calculation ledger is provided to the LLM as source context, but this stage does not prove that each formula is translated exactly. Compilation and method-presence checks do not prove behavioral equivalence. The report says `NOT VERIFIED` until a reference test corpus executes both the original VBA and Java and compares results and side effects. Exact 100% behavior cannot be claimed from an LLM response or successful Java compilation alone. VBA-specific features such as `Currency`, `Variant`, `ByRef`, dynamic arrays, error handling, and `GoSub` need particular review in sample 6.
