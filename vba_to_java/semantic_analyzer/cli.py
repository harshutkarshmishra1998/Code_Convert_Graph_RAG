from __future__ import annotations
import argparse
import json
import sys
from pathlib import Path
from .normalize import load_json, run_files


def main(input_path: str | Path, output: str | Path) -> int:
    source_path = Path(input_path).expanduser().resolve()
    sample_dir = Path(output).expanduser().resolve()
    try:
        if not source_path.is_file():
            raise FileNotFoundError(f"VBA source file does not exist: {source_path}")

        ir_path = sample_dir / "universal_ir.json"
        graph_path = sample_dir / "graph.json"
        segments_path = sample_dir / "segments.json"
        for required_path in (ir_path, graph_path, segments_path):
            if not required_path.is_file():
                raise FileNotFoundError(f"Required input file does not exist: {required_path}")

        graph = load_json(graph_path)
        project = graph.get("project") if isinstance(graph, dict) else None
        graph_name = project.get("name") if isinstance(project, dict) else project
        if graph_name and str(graph_name) != source_path.stem:
            raise ValueError(
                f"Input {source_path.name} does not match graph project {graph_name!r}"
            )

        manifest_path = sample_dir / "manifest.json"
        pseudocode_path = sample_dir / "pseudocode.md"
        ir_out, report_out = run_files(
            ir_path=ir_path,
            graph_path=graph_path,
            segments_path=segments_path,
            pseudocode_path=pseudocode_path if pseudocode_path.is_file() else None,
            manifest_path=manifest_path if manifest_path.is_file() else None,
            output_dir=sample_dir,
        )
        report = json.loads(report_out.read_text(encoding="utf-8"))
        print(json.dumps({
            "normalized_ir": str(ir_out),
            "semantic_analysis": str(report_out),
            "status": report["status"],
            "summary": report["summary"],
        }, indent=2))
        return 0 if report["status"] != "blocked" else 2
    except Exception as exc:
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1


def cli() -> int:
    parser = argparse.ArgumentParser(
        prog="vba-semantic-analyzer",
        description="Normalize Universal IR using VBA graph and granular source evidence.",
    )
    parser.add_argument("--ir", required=True, help="Input universal_ir.json")
    parser.add_argument("--graph", help="Frontend graph.json")
    parser.add_argument("--segments", help="Frontend segments.json")
    parser.add_argument("--pseudocode", help="Supporting pseudocode.md")
    parser.add_argument("--manifest", help="Frontend manifest.json")
    parser.add_argument(
        "--out-dir",
        required=True,
        help="Output directory for normalized_ir.json and semantic_analysis.json",
    )
    args = parser.parse_args()
    ir_out, report_out = run_files(
        ir_path=args.ir,
        graph_path=args.graph,
        segments_path=args.segments,
        pseudocode_path=args.pseudocode,
        manifest_path=args.manifest,
        output_dir=args.out_dir,
    )
    report = json.loads(report_out.read_text(encoding="utf-8"))
    print(json.dumps({
        "normalized_ir": str(ir_out),
        "semantic_analysis": str(report_out),
        "status": report["status"],
        "summary": report["summary"],
    }, indent=2))
    return 0 if report["status"] != "blocked" else 2

if __name__=='__main__':
    raise SystemExit(cli())
