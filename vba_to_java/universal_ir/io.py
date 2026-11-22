"""JSON serialization helpers for Universal IR."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .validator import validate_program

IR_VERSION = "1.0.0"


def new_program(name: str, source_language: str = "VBA") -> dict[str, Any]:
    return {
        "ir_version": IR_VERSION,
        "source_language": source_language,
        "name": name,
        "modules": [],
        "dependencies": [],
        "diagnostics": [],
        "metadata": {},
        "semantics": {},
    }


def load_ir(path: str | Path, validate: bool = True) -> dict[str, Any]:
    with Path(path).open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    if validate:
        issues = validate_program(data)
        if issues:
            summary = "\n".join(f"- {x['path']}: {x['message']} [{x['code']}]" for x in issues)
            raise ValueError(f"Invalid Universal IR:\n{summary}")
    return data


def save_ir(program: dict[str, Any], path: str | Path, validate: bool = True) -> Path:
    if validate:
        issues = validate_program(program)
        if issues:
            summary = "\n".join(f"- {x['path']}: {x['message']} [{x['code']}]" for x in issues)
            raise ValueError(f"Refusing to save invalid Universal IR:\n{summary}")
    destination = Path(path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.with_suffix(destination.suffix + ".tmp")
    with temporary.open("w", encoding="utf-8", newline="\n") as handle:
        json.dump(program, handle, indent=2, ensure_ascii=False)
        handle.write("\n")
    temporary.replace(destination)
    return destination
