"""Run the deterministic VBA-to-Java analysis and generation pipeline."""
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Callable

from vba_to_java.intelligent_fixer.core.fixer_orchestrator import IntelligentFixer
from vba_to_java.java_codegen import main as generate_java
from vba_to_java.llm_handoff import main as build_handoff
from vba_to_java.semantic_analyzer.cli import main as analyze
from vba_to_java.universal_ir.cli import main as build_ir
from vba_to_java.vba_graph.cli import main as build_graph


def _run_fixer_offline(input_path: Path, output: Path) -> dict:
    fixer = IntelligentFixer(
        output,
        llm=None,
        reviewer_llm=None,
        allow_auto_apply=False,
        reports_dir=output,
    )
    return fixer.run(
        apply=False,
        run_metadata={
            "source_file": input_path.name,
            "sample_folder": output.name,
            "outputs_written_to": str(output),
            "orchestrator": "vba_to_java.main",
        },
    )


def main(input_path: str | Path, output: str | Path) -> int:
    source_path = Path(input_path).expanduser().resolve()
    sample_dir = Path(output).expanduser().resolve()
    run_report: dict = {
        "source_file": source_path.name,
        "output_directory": str(sample_dir),
        "stages": [],
        "status": "running",
    }

    def stage(name: str, action: Callable[[], int], accepted_codes=(0,)) -> int:
        print(f"[{sample_dir.name}] {name}...")
        code = action()
        run_report["stages"].append({
            "name": name,
            "return_code": code,
            "status": "completed" if code in accepted_codes else "failed",
        })
        return code

    try:
        if not source_path.is_file():
            raise FileNotFoundError(f"VBA source file does not exist: {source_path}")

        code = stage("VBA graph", lambda: build_graph(source_path, sample_dir))
        if code != 0:
            raise RuntimeError(f"VBA graph stage failed with return code {code}")

        code = stage("Universal IR", lambda: build_ir(source_path, sample_dir))
        if code != 0:
            raise RuntimeError(f"Universal IR stage failed with return code {code}")

        code = stage("Semantic analysis", lambda: analyze(source_path, sample_dir), accepted_codes=(0, 2))
        semantic_report = json.loads((sample_dir / "semantic_analysis.json").read_text(encoding="utf-8"))
        run_report["stages"][-1]["analysis_status"] = semantic_report.get("status")
        if code not in (0, 2):
            raise RuntimeError(f"Semantic analysis failed with return code {code}")

        run_report["stages"].append({
            "name": "Pseudocode generation",
            "status": "skipped",
            "reason": "The pseudocode module sends source to an external LLM; this local pipeline does not make external requests.",
        })

        code = stage("Java generation and javac check", lambda: generate_java(source_path, sample_dir))
        if code != 0:
            raise RuntimeError(f"Java generation or compilation failed with return code {code}")

        print(f"[{sample_dir.name}] Intelligent-fixer audit (offline, no auto-apply)...")
        fixer_result = _run_fixer_offline(source_path, sample_dir)
        quality = fixer_result["quality_report"]
        run_report["stages"].append({
            "name": "Intelligent-fixer audit",
            "status": "completed",
            "llm_enabled": False,
            "auto_apply": False,
            "conversion_readiness": quality.get("conversion_readiness"),
            "unresolved_issue_count": len(fixer_result.get("unresolved", [])),
        })
        print(
            f"[{sample_dir.name}] fixer: {quality.get('conversion_readiness')} "
            f"({len(fixer_result.get('unresolved', []))} unresolved)"
        )

        code = stage("LLM handoff artifacts", lambda: build_handoff(source_path, sample_dir))
        if code != 0:
            raise RuntimeError(f"LLM handoff stage failed with return code {code}")

        run_report["status"] = "completed"
        run_report["java_file"] = f"{source_path.stem}.java"
        run_report["java_compile_status"] = json.loads(
            (sample_dir / "java_codegen_report.json").read_text(encoding="utf-8")
        ).get("compile_status")
        return_code = 0
    except Exception as exc:
        run_report["status"] = "failed"
        run_report["error"] = f"{type(exc).__name__}: {exc}"
        print(f"[{sample_dir.name}] ERROR: {run_report['error']}", file=sys.stderr)
        return_code = 1
    finally:
        sample_dir.mkdir(parents=True, exist_ok=True)
        report_path = sample_dir / "pipeline_run.json"
        report_path.write_text(json.dumps(run_report, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
        print(f"[{sample_dir.name}] pipeline report: {report_path}")
    return return_code


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
