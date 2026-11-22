# LLM code-generation handoff: vba_sample_4

This bundle is a source-grounded, target-language-neutral handoff. No external
LLM was called while creating it.

## Readiness

- Source provenance verified: `3a6100a24086f35320010e5524d9db8e8566e5965564172b12befffd97bd6f5a`
- Graph: 83 nodes, 140 edges,
  0 diagnostics
- Procedures: ConsolidateDepartmentFiles
- Semantic status: **blocked**
- Migration readiness: **blocked**
- Unresolved expressions: 0
- Pseudocode available: **no**
- Intelligent fixer LLM enabled: **False**
- Intelligent fixer issues: 1 addressed and verified,
  9 unresolved
- Missing pseudocode procedure sections:
  1
- Procedures missing detailed pseudocode:
- `ConsolidateDepartmentFiles`

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

- `llm_handoff_ConsolidateDepartmentFiles.json` — `ConsolidateDepartmentFiles`
