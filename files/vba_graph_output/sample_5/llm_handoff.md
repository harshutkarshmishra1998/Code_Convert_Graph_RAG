# LLM code-generation handoff: vba_sample_5

This bundle is a source-grounded, target-language-neutral handoff. No external
LLM was called while creating it.

## Readiness

- Source provenance verified: `7599243eb5451741202a92e8f2d6254f463167ca9c90e129ee10df726fe244cb`
- Graph: 95 nodes, 150 edges,
  0 diagnostics
- Procedures: ReconcileInventory, LogInventoryError
- Semantic status: **blocked**
- Migration readiness: **blocked**
- Unresolved expressions: 0
- Pseudocode available: **no**
- Intelligent fixer LLM enabled: **False**
- Intelligent fixer issues: 1 addressed and verified,
  11 unresolved
- Missing pseudocode procedure sections:
  2
- Procedures missing detailed pseudocode:
- `ReconcileInventory`
- `LogInventoryError`

## Blocking compatibility requirements

- **VBA_ERROR_HANDLING — VBA On Error runtime semantics:** Implement a compatibility runtime for Resume Next, GoTo handler, GoTo 0, Err state and handler activation; this pass preserves directives but does not emulate execution.
- **LATE_BOUND_OBJECT_MEMBERS — Late-bound COM/Object members:** Define explicit Java interfaces/adapters for the actual COM object behavior. Object-typed references alone do not identify the target API.
- **COM_CREATEOBJECT — COM object activation:** Replace ProgID-based activation with an explicit Java implementation or integration adapter; preserve fallback behavior.

## How to use

Provide `llm_handoff.json` and one or more matching procedure files to the
code-generation model, specify a target language/runtime, and ask it to follow
the `generation_instructions`. Include procedure files listed as direct
callees when translating a caller. The original VBA source is authoritative;
graph, segments, Universal IR, normalized IR, and semantic analysis remain
available beside the handoff files and may explicitly preserve uncertainty.
Do not treat a blocked semantic status as a successful conversion.

## Procedure context files

- `llm_handoff_ReconcileInventory.json` — `ReconcileInventory`
- `llm_handoff_LogInventoryError.json` — `LogInventoryError`
