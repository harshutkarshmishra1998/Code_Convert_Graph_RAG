import json
import sys
from pathlib import Path

from .engine import convert_path


def main(input_path, output, annotation_payload=None, strict=False):
    try:
        result = convert_path(input_path, annotation_payload)
        out = Path(output)
        out.mkdir(parents=True, exist_ok=True)
        for filename, key in (
            ("graph.json", "graph"),
            ("segments.json", "segments"),
            ("report.json", "report"),
        ):
            (out / filename).write_text(
                json.dumps(result[key], indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
        (out / "manifest.json").write_text(
            json.dumps(result["graph"]["manifest"], indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        print(json.dumps(result["report"], indent=2))
        return 2 if strict and result["report"]["diagnostic_count"] else 0
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
