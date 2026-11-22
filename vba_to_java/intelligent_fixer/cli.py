"""Sample-folder entry point for the intelligent fixer."""
from __future__ import annotations

import os
from pathlib import Path

from vba_to_java.intelligent_fixer.core.fixer_orchestrator import IntelligentFixer
from vba_to_java.intelligent_fixer.core.project_loader import ProjectLoader
from vba_to_java.intelligent_fixer.llm_adapter import ChatGroqJSONAdapter


def _enabled(name: str) -> bool:
    return os.getenv(name, "").strip().lower() in {"1", "true", "yes"}


def main(input_path: str | Path, output: str | Path) -> int:
    source_path = Path(input_path).expanduser().resolve()
    sample_dir = Path(output).expanduser().resolve()
    if not source_path.is_file():
        raise FileNotFoundError(f"VBA source file not found: {source_path}")
    if not sample_dir.is_dir():
        raise FileNotFoundError(f"Sample output directory not found: {sample_dir}")
    if not (sample_dir / "graph.json").is_file():
        raise FileNotFoundError(f"Graph artifact not found: {sample_dir / 'graph.json'}")

    project = ProjectLoader(sample_dir).load()
    if not project.get("artifacts", {}).get("graph"):
        raise ValueError(f"No usable graph artifact found in {sample_dir}")

    use_llm = _enabled("INTELLIGENT_FIXER_USE_LLM")
    apply_changes = _enabled("INTELLIGENT_FIXER_APPLY")
    if apply_changes and not use_llm:
        raise ValueError(
            "Set INTELLIGENT_FIXER_USE_LLM=1 as well before enabling patch application."
        )

    llm = (
        ChatGroqJSONAdapter.from_project(Path(__file__).resolve().parents[2])
        if use_llm else None
    )
    fixer = IntelligentFixer(
        sample_dir,
        llm=llm,
        reviewer_llm=llm,
        allow_auto_apply=apply_changes,
        reports_dir=sample_dir,
    )
    result = fixer.run(
        apply=apply_changes,
        run_metadata={
            "source_file": source_path.name,
            "sample_folder": sample_dir.name,
            "outputs_written_to": str(sample_dir),
        },
    )
    report = result["quality_report"]
    print(
        f"{sample_dir.name}: {report['conversion_readiness']} "
        f"({len(result['unresolved'])} unresolved, "
        f"{report['addressed_issue_count']} addressed); "
        f"details: {sample_dir / 'intelligent_fixer_log.json'}"
    )
    return 2 if report["conversion_readiness"] != "VERIFIED" else 0
