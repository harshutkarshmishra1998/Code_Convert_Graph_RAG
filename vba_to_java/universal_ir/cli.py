"""Command-line interface for Universal IR."""
from __future__ import annotations
import argparse
import json
import sys
from pathlib import Path
from .builder import build_ir
from .io import save_ir
from .validator import validate_program


def main(input_path: str | Path, output: str | Path) -> int:
    source_path = Path(input_path).expanduser().resolve()
    sample_dir = Path(output).expanduser().resolve()
    try:
        if not source_path.is_file():
            raise FileNotFoundError(f"VBA source file does not exist: {source_path}")

        graph_path = sample_dir / "graph.json"
        segments_path = sample_dir / "segments.json"
        if not graph_path.is_file():
            raise FileNotFoundError(f"Graph JSON not found: {graph_path}")
        if not segments_path.is_file():
            raise FileNotFoundError(f"Segments JSON not found: {segments_path}")

        graph = json.loads(graph_path.read_text(encoding="utf-8"))
        project = graph.get("project") if isinstance(graph, dict) else None
        graph_name = project.get("name") if isinstance(project, dict) else project
        if graph_name and str(graph_name) != source_path.stem:
            raise ValueError(
                f"Input {source_path.name} does not match graph project {graph_name!r}"
            )

        pseudocode_path = sample_dir / "pseudocode.md"
        ir = build_ir(
            graph_path,
            segments_path,
            pseudocode_path if pseudocode_path.is_file() else None,
        )
        issues = validate_program(ir)
        if issues:
            print(json.dumps(issues, indent=2))
            return 2

        output_path = save_ir(ir, sample_dir / "universal_ir.json")
        print(json.dumps({
            "output": str(output_path),
            "modules": len(ir["modules"]),
            "procedures": sum(len(module["procedures"]) for module in ir["modules"]),
            "diagnostics": len(ir["diagnostics"]),
            "validation_issues": [],
        }, indent=2))
        return 0
    except Exception as exc:
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1


def cli() -> int:
    parser = argparse.ArgumentParser(prog="universal-ir", description="Build and validate language-neutral IR")
    sub = parser.add_subparsers(dest="command", required=True)
    build = sub.add_parser("build", help="Build IR from VBA graph and granular segments")
    build.add_argument("--graph", required=True)
    build.add_argument("--segments")
    build.add_argument("--pseudocode")
    build.add_argument("--out", required=True)
    build.add_argument("--allow-invalid", action="store_true")
    validate = sub.add_parser("validate", help="Validate an existing IR JSON file")
    validate.add_argument("ir")
    args = parser.parse_args()
    if args.command == "build":
        ir = build_ir(args.graph, args.segments, args.pseudocode)
        issues = validate_program(ir)
        if issues and not args.allow_invalid:
            print(json.dumps(issues, indent=2))
            return 2
        out = save_ir(ir, args.out, validate=not args.allow_invalid)
        print(json.dumps({"output": str(out), "modules": len(ir["modules"]),
                          "procedures": sum(len(m["procedures"]) for m in ir["modules"]),
                          "diagnostics": len(ir["diagnostics"]), "validation_issues": issues}, indent=2))
        return 0
    data = json.loads(Path(args.ir).read_text(encoding="utf-8"))
    issues = validate_program(data)
    print(json.dumps({"valid": not issues, "issues": issues}, indent=2))
    return 0 if not issues else 2

if __name__ == "__main__":
    raise SystemExit(cli())
