# VBA Process Notes

## What this does

This module reads each selected VBA file and creates a plain-language process note for later Java work. It keeps the original statement and source line next to the explanation, so a person can check calculations, decisions, loops, reads/writes, procedure calls, and error handling against the VBA.

The note has two parts:

1. **Source inventory:** generated from the VBA text. It keeps the exact statements and labels their apparent role.
2. **LLM explanation:** asks the configured Groq model to explain one procedure at a time using strict JSON Schema output. It summarizes the flow, calculations, inputs/outputs, branches, side effects, dependencies, and unanswered questions. Every LLM finding must cite source line numbers. The program checks that cited lines exist and fall within the relevant procedure.

The LLM can still misunderstand code. Its confidence and citations make the explanation easier to review; they do not prove it is correct. Keep uncertain behavior as a question and have a person verify the note before Java conversion. This stage does not generate Java.

## Run

From the project root:

```sh
python -m vba_to_java_2.process_notes.main
```

The small `main.py` runner selects samples `1` through `8` from `files/vba/input` and calls `main(input_location, output_location)` from `notes_generator.py`. Each note is written as `process_notes.md` in its corresponding folder under `files/vba/output`.

The terminal prints one progress sequence per input. Look for:

```text
[PROCESS] ...
[DETERMINISTIC] START — ...
[DETERMINISTIC] COMPLETE — ...
[LLM] START — ...
[LLM] PROCEDURE 1/1 — ...
[LLM] COMPLETE — ...
[OUTPUT] SAVED — ...
```

`[LLM] SKIPPED` means LLM analysis was deliberately disabled. `[LLM] FAILED` means the model request, schema validation, or citation check did not complete; the current note is saved with the failure recorded, then `[STOP]` ends the batch immediately. `[DETERMINISTIC] FAILED` or `[OUTPUT] FAILED` also stops the batch. The runner does not move to the next procedure or input after a failure.

To run the source inventory without contacting the LLM:

```sh
VBA_NOTES_DISABLE_LLM=1 python -m vba_to_java_2.process_notes.main
```

## LLM setup and privacy

The LLM connection uses the existing project `api_keys.py` configuration (`GROQ_API` and `LLM`). `LLM_MODEL` can override the configured model. Do not put credentials in source files or process notes.

When enabled, the procedure source from each VBA input is sent to the configured Groq model, one procedure per request. Use this only for source that you are allowed to send to that provider. If the configuration, network call, structured response, or citation checks fail, the note still contains its source inventory and clearly records each LLM failure.

## Output and review

Each `process_notes.md` contains file hash and encoding, module settings and UDT/enum members, procedure signatures, ordered source-linked statements, a deterministic calculation ledger, the LLM explanation (when available), a review checklist, and counts of statements that need attention. LLM findings are checked against every detected procedure calculation. If the LLM omits a calculation, the notes include the exact expression and source line from the deterministic ledger, explicitly marked as not interpreted by the LLM.

Review calculations and formulas exactly; check every branch outcome and loop boundary; identify workbook, file, UI, and other side effects; confirm procedure dependencies; and resolve every question. A generated note is a draft record of behavior, not a measured accuracy score or a guarantee of equivalence.
