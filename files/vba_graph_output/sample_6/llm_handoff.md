# LLM code-generation handoff: vba_sample_6

This bundle is a source-grounded, target-language-neutral handoff. No external
LLM was called while creating it.

## Readiness

- Source provenance verified: `7289ca3eb6d951118004db2ff7c58f5f44b6b8ee6a818a3efa5490dea9542736`
- Graph: 323 nodes, 502 edges,
  0 diagnostics
- Procedures: ExecuteStressPipeline, InitInfrastructure, NextSequence, LogAudit, NormalizeLineEndings, TryParseEntry, NormalizeEntryInPlace, DeduplicateBatchEntries, RunLegacyGoSubAudit, EvaluateDependencyDepth, BuildUpsertSql, SanitizeSqlIdentifier, BuildSummaryReport, GetAuditEntryCount, ResetModuleState
- Semantic status: **blocked**
- Migration readiness: **blocked**
- Unresolved expressions: 0
- Pseudocode available: **no**
- Intelligent fixer LLM enabled: **False**
- Intelligent fixer issues: 1 addressed and verified,
  37 unresolved
- Missing pseudocode procedure sections:
  15
- Procedures missing detailed pseudocode:
- `ExecuteStressPipeline`
- `InitInfrastructure`
- `NextSequence`
- `LogAudit`
- `NormalizeLineEndings`
- `TryParseEntry`
- `NormalizeEntryInPlace`
- `DeduplicateBatchEntries`
- `RunLegacyGoSubAudit`
- `EvaluateDependencyDepth`
- `BuildUpsertSql`
- `SanitizeSqlIdentifier`
- `BuildSummaryReport`
- `GetAuditEntryCount`
- `ResetModuleState`

## Blocking compatibility requirements

- **VBA_ERROR_HANDLING — VBA On Error runtime semantics:** Implement a compatibility runtime for Resume Next, GoTo handler, GoTo 0, Err state and handler activation; this pass preserves directives but does not emulate execution.
- **VBA_GOSUB_RETURN — Legacy GoSub/Return control flow:** Lower GoSub to an explicit continuation/return stack or a proven equivalent; never map Return to Java method return.
- **LATE_BOUND_OBJECT_MEMBERS — Late-bound COM/Object members:** Define explicit Java interfaces/adapters for the actual COM object behavior. Object-typed references alone do not identify the target API.
- **COM_CREATEOBJECT — COM object activation:** Replace ProgID-based activation with an explicit Java implementation or integration adapter; preserve fallback behavior.
- **VBA_ARRAY_INITIALIZATION_IDIOM — Double-Not dynamic-array initialization check:** Confirm the source runtime meaning and replace with an explicit initialized/uninitialized array state in Java. Do not translate Not Not as ordinary boolean negation.

## How to use

Provide `llm_handoff.json` and one or more matching procedure files to the
code-generation model, specify a target language/runtime, and ask it to follow
the `generation_instructions`. Include procedure files listed as direct
callees when translating a caller. The original VBA source is authoritative;
graph, segments, Universal IR, normalized IR, and semantic analysis remain
available beside the handoff files and may explicitly preserve uncertainty.
Do not treat a blocked semantic status as a successful conversion.

## Procedure context files

- `llm_handoff_ExecuteStressPipeline.json` — `ExecuteStressPipeline`
- `llm_handoff_InitInfrastructure.json` — `InitInfrastructure`
- `llm_handoff_NextSequence.json` — `NextSequence`
- `llm_handoff_LogAudit.json` — `LogAudit`
- `llm_handoff_NormalizeLineEndings.json` — `NormalizeLineEndings`
- `llm_handoff_TryParseEntry.json` — `TryParseEntry`
- `llm_handoff_NormalizeEntryInPlace.json` — `NormalizeEntryInPlace`
- `llm_handoff_DeduplicateBatchEntries.json` — `DeduplicateBatchEntries`
- `llm_handoff_RunLegacyGoSubAudit.json` — `RunLegacyGoSubAudit`
- `llm_handoff_EvaluateDependencyDepth.json` — `EvaluateDependencyDepth`
- `llm_handoff_BuildUpsertSql.json` — `BuildUpsertSql`
- `llm_handoff_SanitizeSqlIdentifier.json` — `SanitizeSqlIdentifier`
- `llm_handoff_BuildSummaryReport.json` — `BuildSummaryReport`
- `llm_handoff_GetAuditEntryCount.json` — `GetAuditEntryCount`
- `llm_handoff_ResetModuleState.json` — `ResetModuleState`
