"""Build source-grounded, target-language-neutral context bundles for code generation."""
from __future__ import annotations

import hashlib
import json
import re
import sys
from pathlib import Path
from typing import Any


ARTIFACT_FILES = {
    "graph": "graph.json",
    "segments": "segments.json",
    "graph_report": "report.json",
    "manifest": "manifest.json",
    "universal_ir": "universal_ir.json",
    "normalized_ir": "normalized_ir.json",
    "semantic_analysis": "semantic_analysis.json",
}
FIXER_FILES = {
    "intelligent_fixer_log": "intelligent_fixer_log.json",
    "final_quality_report": "final_quality_report.json",
    "unresolved_issues": "unresolved_issues.json",
}

GENERATION_INSTRUCTIONS = [
    "Generate code only after the target language and runtime are specified.",
    "Treat the original VBA source as the behavior authority; use graph, segments, IR, and semantic analysis as supporting evidence.",
    "Preserve control flow, errors, data conversions, evaluation order, side effects, and VBA-specific behavior.",
    "Do not silently omit unsupported or opaque statements, unresolved symbols, compatibility requirements, or missing pseudocode sections.",
    "Treat intelligent-fixer issues as resolved only when its final status is addressed_and_verified; preserve unresolved findings.",
    "Do not invent external APIs, workbook schemas, file formats, or runtime behavior; list assumptions and ask for missing requirements.",
    "Keep source-location and graph/segment provenance with each translated procedure where practical.",
    "Return a procedure-by-procedure implementation, a behavior/compatibility note, and a list of unresolved questions.",
]


def _read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _walk(value: Any):
    if isinstance(value, dict):
        yield value
        for child in value.values():
            yield from _walk(child)
    elif isinstance(value, list):
        for child in value:
            yield from _walk(child)


def _pseudocode_section(text: str | None, procedure_name: str) -> str | None:
    if not text:
        return None
    header = re.compile(
        rf"(?im)^[ \t]*#{{2,4}}[ \t]+Procedure:[ \t]*{re.escape(procedure_name)}(?:[ \t(]|$).*$"
    )
    match = header.search(text)
    if not match:
        return None
    next_heading = re.search(r"(?m)^[ \t]*#{1,3}[ \t]+", text[match.end():])
    end = match.end() + next_heading.start() if next_heading else len(text)
    section = text[match.start():end].strip()
    return section or None


def build_handoff(input_path: str | Path, output: str | Path) -> tuple[Path, Path]:
    source_path = Path(input_path).expanduser().resolve()
    sample_dir = Path(output).expanduser().resolve()
    if not source_path.is_file():
        raise FileNotFoundError(f"VBA source file does not exist: {source_path}")

    artifacts: dict[str, Any] = {}
    for key, filename in ARTIFACT_FILES.items():
        path = sample_dir / filename
        if not path.is_file():
            raise FileNotFoundError(f"Required handoff input does not exist: {path}")
        artifacts[key] = _read_json(path)

    project = artifacts["graph"].get("project")
    graph_name = project.get("name") if isinstance(project, dict) else project
    if graph_name and str(graph_name) != source_path.stem:
        raise ValueError(f"Input {source_path.name} does not match graph project {graph_name!r}")

    source_bytes = source_path.read_bytes()
    source_hash = hashlib.sha256(source_bytes).hexdigest()
    manifest_entries = artifacts["manifest"]
    if not isinstance(manifest_entries, list):
        raise ValueError("manifest.json must contain a source-file list")
    source_record = next(
        (entry for entry in manifest_entries
         if Path(str(entry.get("path", ""))).name == source_path.name),
        None,
    )
    if source_record is None:
        raise ValueError(f"manifest.json has no entry for {source_path.name}")
    if source_record.get("sha256") != source_hash:
        raise ValueError(
            f"Source hash mismatch for {source_path.name}; regenerate graph artifacts before bundling"
        )
    source_text = source_bytes.decode(source_record.get("encoding") or "utf-8")

    pseudocode_path = sample_dir / "pseudocode.md"
    pseudocode_run_path = sample_dir / "pseudocode_run.json"
    pseudocode = pseudocode_path.read_text(encoding="utf-8") if pseudocode_path.is_file() else None
    pseudocode_run = _read_json(pseudocode_run_path) if pseudocode_run_path.is_file() else None
    fixer_artifacts = {
        key: _read_json(sample_dir / filename)
        for key, filename in FIXER_FILES.items()
        if key != "unresolved_issues" and (sample_dir / filename).is_file()
    }
    fixer_run_list = fixer_artifacts.get("intelligent_fixer_log", {}).get("runs", [])
    latest_fixer_run = fixer_run_list[-1] if fixer_run_list else {}
    fixer_quality = fixer_artifacts.get("final_quality_report", {})
    fixer_issues = latest_fixer_run.get(
        "issues", fixer_quality.get("issue_resolution_log", [])
    )
    fixer_summary = {
        "run_available": bool(latest_fixer_run or fixer_quality),
        "llm_enabled": latest_fixer_run.get("llm_enabled"),
        "conversion_readiness": fixer_quality.get("conversion_readiness"),
        "issue_count": latest_fixer_run.get("issue_count", fixer_quality.get("issue_count")),
        "addressed_issue_count": latest_fixer_run.get(
            "addressed_issue_count", fixer_quality.get("addressed_issue_count")
        ),
        "unresolved_issue_count": latest_fixer_run.get(
            "unresolved_issue_count", fixer_quality.get("unresolved_issue_count")
        ),
        "issues": fixer_issues,
    }
    fixer_shared_summary = {
        key: value for key, value in fixer_summary.items() if key != "issues"
    }

    semantic = artifacts["semantic_analysis"]
    semantic_summary = semantic.get("summary", {})
    compatibility = semantic.get("compatibility_requirements", [])
    blocking_requirements = [item for item in compatibility if item.get("severity") == "blocking"]
    report_issues = semantic.get("issues", [])
    error_issues = [item for item in report_issues if item.get("severity") == "error"]
    normalized_ir = artifacts["normalized_ir"]
    pseudocode_gaps = semantic.get("pseudocode_coverage_gaps", [])
    segment_list = artifacts["segments"]
    if not isinstance(segment_list, list):
        raise ValueError("segments.json must contain a source-segment list")
    procedure_segments = {
        str(segment.get("name", "")).casefold(): segment
        for segment in segment_list
        if segment.get("kind") == "PROCEDURE"
    }
    procedure_reports = {
        str(procedure.get("procedure", "")).casefold(): procedure
        for module in semantic.get("modules", [])
        for procedure in module.get("procedures", [])
    }
    all_procedures = [
        (module, procedure)
        for module in normalized_ir.get("modules", [])
        for procedure in module.get("procedures", [])
    ]
    signatures = [
        {
            "module": module.get("name"),
            "name": procedure.get("name"),
            "kind": procedure.get("kind"),
            "visibility": procedure.get("visibility"),
            "parameters": procedure.get("parameters", []),
            "return_type": procedure.get("return_type"),
        }
        for module, procedure in all_procedures
    ]
    module_declarations = {
        str(module.get("name", "")): module.get("declarations", [])
        for module in normalized_ir.get("modules", [])
    }

    artifact_manifest = {}
    for key, filename in ARTIFACT_FILES.items():
        path = sample_dir / filename
        artifact_manifest[key] = {
            "file": filename,
            "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
            "size_bytes": path.stat().st_size,
        }
    for key, filename in FIXER_FILES.items():
        path = sample_dir / filename
        if path.is_file():
            artifact_manifest[key] = {
                "file": filename,
                "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
                "size_bytes": path.stat().st_size,
            }
    for key, path in (("pseudocode", pseudocode_path), ("pseudocode_run", pseudocode_run_path)):
        if path.is_file():
            artifact_manifest[key] = {
                "file": path.name,
                "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
                "size_bytes": path.stat().st_size,
            }

    procedure_files = []
    procedure_file_map = {}
    procedure_calls_map = {}
    for module, procedure in all_procedures:
        name = str(procedure.get("name", ""))
        key = name.casefold()
        segment = procedure_segments.get(key)
        if segment is None:
            raise ValueError(f"segments.json has no whole-procedure source segment for {name!r}")
        source_span = (procedure.get("provenance") or {}).get("source_location")
        if not segment.get("source"):
            raise ValueError(f"Procedure segment for {name!r} has no source text")
        direct_callees = sorted({
            str(target.get("name"))
            for node in _walk(procedure.get("body", []))
            if node.get("kind") == "invocation_expression"
            for target in [node.get("target_procedure")]
            if isinstance(target, dict) and target.get("name")
        })
        filename = f"llm_handoff_{re.sub(r'[^A-Za-z0-9_-]+', '_', name)}.json"
        procedure_file_map[key] = filename
        procedure_calls_map[key] = direct_callees
        procedure_files.append({
            "module": module.get("name"),
            "name": name,
            "file": filename,
            "kind": procedure.get("kind"),
            "source_location": source_span,
            "graph_node_ids": (procedure.get("provenance") or {}).get("graph_node_ids", []),
            "segment_ids": (procedure.get("provenance") or {}).get("segment_ids", []),
            "direct_callees": direct_callees,
        })

    shared_context = {
        "source": {
            "path": source_path.name,
            "encoding": source_record.get("encoding") or "utf-8",
            "sha256": source_hash,
            "manifest_entry": source_record,
        },
        "generation_instructions": GENERATION_INSTRUCTIONS,
        "signatures": signatures,
        "module_declarations": module_declarations,
        "semantic_status": semantic.get("status", "unknown"),
        "migration_readiness": semantic.get("migration_readiness", "unknown"),
        "semantic_summary": semantic_summary,
        "blocking_requirements": blocking_requirements,
        "semantic_errors": error_issues,
        "pseudocode_coverage_gaps": pseudocode_gaps,
        "intelligent_fixer": fixer_shared_summary,
        "artifact_files": artifact_manifest,
    }

    for module, procedure in all_procedures:
        name = str(procedure.get("name", ""))
        segment = procedure_segments[name.casefold()]
        source_span = (procedure.get("provenance") or {}).get("source_location")
        calls = [
            {
                "name": callee,
                "handoff_file": procedure_file_map.get(callee.casefold()),
            }
            for callee in procedure_calls_map[name.casefold()]
        ]
        procedure_bundle = {
            "schema_version": "1.0.0",
            "project": source_path.stem,
            "module": module.get("name"),
            "shared_context": shared_context,
            "procedure": procedure,
            "source": {
                "exact_procedure_text": segment["source"],
                "source_location": source_span or {
                    "file": segment.get("path"),
                    "line": segment.get("start_line"),
                    "end_line": segment.get("end_line"),
                },
                "segment_id": segment.get("id"),
            },
            "direct_callees": calls,
            "semantic_procedure_report": procedure_reports.get(name.casefold()),
            "pseudocode_section": _pseudocode_section(pseudocode, name),
            "fixer_issue_statuses": [
                item for item in fixer_issues
                if str((item.get("issue") or {}).get("procedure", "")).casefold() == name.casefold()
            ],
        }
        target = sample_dir / procedure_file_map[name.casefold()]
        target.write_text(
            json.dumps(procedure_bundle, ensure_ascii=False, separators=(",", ":")) + "\n",
            encoding="utf-8",
        )

    bundle = {
        "schema_version": "1.0.0",
        "project": source_path.stem,
        "source": {
            "path": source_path.name,
            "encoding": source_record.get("encoding") or "utf-8",
            "sha256": source_hash,
            "text": source_text,
        },
        "context_quality": {
            "source_provenance_verified": True,
            "all_required_artifacts_present": True,
            "llm_generation_performed": False,
            "semantic_status": semantic.get("status", "unknown"),
            "migration_readiness": semantic.get("migration_readiness", "unknown"),
            "blocking_compatibility_requirement_count": len(blocking_requirements),
            "semantic_error_count": len(error_issues),
            "unresolved_expression_count": semantic_summary.get("unresolved_expression_count"),
            "opaque_statement_count": (artifacts["graph_report"] or {}).get("opaque_count"),
            "pseudocode_available": pseudocode is not None,
            "missing_pseudocode_procedure_sections": semantic_summary.get("pseudocode_procedure_sections_missing"),
            "pseudocode_coverage_gaps": pseudocode_gaps,
            "fixer_run_available": fixer_summary["run_available"],
            "fixer_llm_enabled": fixer_summary["llm_enabled"],
            "fixer_addressed_issue_count": fixer_summary["addressed_issue_count"],
            "fixer_unresolved_issue_count": fixer_summary["unresolved_issue_count"],
        },
        "shared_context": shared_context,
        "procedure_files": procedure_files,
        "pseudocode_run": pseudocode_run,
        "intelligent_fixer": fixer_summary,
    }

    sample_dir.mkdir(parents=True, exist_ok=True)
    json_path = sample_dir / "llm_handoff.json"
    json_path.write_text(json.dumps(bundle, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

    graph_report = artifacts["graph_report"]
    procedure_names = ", ".join(item["name"] for item in procedure_files) or "None detected"
    procedure_file_list = "\n".join(
        f"- `{item['file']}` — `{item['name']}`" for item in procedure_files
    ) or "- No procedures detected."
    blockers = "\n".join(
        f"- **{item.get('code', 'UNKNOWN')} — {item.get('title', 'Compatibility requirement')}:** "
        f"{item.get('target_guidance', '')}"
        for item in blocking_requirements
    ) or "- None identified."
    missing_pseudocode = "\n".join(f"- `{name}`" for name in pseudocode_gaps) or "- None."
    markdown = f"""# LLM code-generation handoff: {source_path.stem}

This bundle is a source-grounded, target-language-neutral handoff. No external
LLM was called while creating it.

## Readiness

- Source provenance verified: `{source_hash}`
- Graph: {graph_report.get('node_count', 0)} nodes, {graph_report.get('edge_count', 0)} edges,
  {graph_report.get('diagnostic_count', 0)} diagnostics
- Procedures: {procedure_names}
- Semantic status: **{semantic.get('status', 'unknown')}**
- Migration readiness: **{semantic.get('migration_readiness', 'unknown')}**
- Unresolved expressions: {semantic_summary.get('unresolved_expression_count', 'unknown')}
- Pseudocode available: **{'yes' if pseudocode is not None else 'no'}**
- Intelligent fixer LLM enabled: **{fixer_summary['llm_enabled']}**
- Intelligent fixer issues: {fixer_summary['addressed_issue_count']} addressed and verified,
  {fixer_summary['unresolved_issue_count']} unresolved
- Missing pseudocode procedure sections:
  {semantic_summary.get('pseudocode_procedure_sections_missing', 'unknown')}
- Procedures missing detailed pseudocode:
{missing_pseudocode}

## Blocking compatibility requirements

{blockers}

## How to use

Provide `llm_handoff.json` and one or more matching procedure files to the
code-generation model, specify a target language/runtime, and ask it to follow
the `generation_instructions`. Include procedure files listed as direct
callees when translating a caller. The original VBA source is authoritative;
graph, segments, Universal IR, normalized IR, and semantic analysis remain
available beside the handoff files and may explicitly preserve uncertainty.
Do not treat a blocked semantic status as a successful conversion.

## Procedure context files

{procedure_file_list}
"""
    markdown_path = sample_dir / "llm_handoff.md"
    markdown_path.write_text(markdown, encoding="utf-8")
    return json_path, markdown_path


def main(input_path: str | Path, output: str | Path) -> int:
    try:
        json_path, markdown_path = build_handoff(input_path, output)
        print(json.dumps({"json": str(json_path), "markdown": str(markdown_path)}, indent=2))
        return 0
    except Exception as exc:
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    project_root = Path(__file__).resolve().parents[1]
    output_root = project_root / "files" / "vba_graph_output"
    n = [1, 2, 3, 4, 5, 6]
    status = 0

    for i in n:
        input_path = project_root / "files" / f"vba_sample_{i}.bas"
        output = output_root / f"sample_{i}"
        status = max(status, main(input_path, output))

    raise SystemExit(status)
