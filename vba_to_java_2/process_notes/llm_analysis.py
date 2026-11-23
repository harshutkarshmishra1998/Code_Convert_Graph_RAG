"""Procedure-by-procedure LLM interpretation with source-citation checks."""

import importlib
import json
import os
import sys
from collections import Counter
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, Field


class Finding(BaseModel):
    category: Literal[
        "input_output", "calculation", "decision", "loop",
        "data_effect", "dependency", "uncertainty",
    ]
    description: str
    exact_vba: str = ""
    source_lines: list[int] = Field(min_length=1)


class ProcedureAnalysis(BaseModel):
    name: str
    summary: str
    summary_source_lines: list[int] = Field(min_length=1)
    confidence: Literal["high", "medium", "low"]
    findings: list[Finding]


_RESPONSE_FORMAT = {
    "type": "json_schema",
    "json_schema": {
        "name": "procedure_analysis",
        "strict": False,
        "schema": {
            "type": "object",
            "additionalProperties": False,
            "required": ["name", "summary", "summary_source_lines", "confidence", "findings"],
            "properties": {
                "name": {"type": "string"},
                "summary": {"type": "string"},
                "summary_source_lines": {"type": "array", "items": {"type": "integer"}},
                "confidence": {"type": "string", "enum": ["high", "medium", "low"]},
                "findings": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "additionalProperties": False,
                        "required": ["category", "description", "source_lines"],
                        "properties": {
                            "category": {"type": "string"},
                            "description": {"type": "string"},
                            "exact_vba": {"type": "string"},
                            "source_lines": {"type": "array", "items": {"type": "integer"}},
                        },
                    },
                },
            },
        },
    },
}


_CLIENT = None
_MODEL_NAME = None
_JSON_SCHEMA_CLIENT = None


def _client():
    global _CLIENT, _MODEL_NAME, _JSON_SCHEMA_CLIENT
    if _JSON_SCHEMA_CLIENT is not None:
        return _JSON_SCHEMA_CLIENT, _MODEL_NAME
    project_root = Path(__file__).resolve().parents[2]
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))
    try:
        api_keys = importlib.import_module("api_keys")
        from langchain_groq import ChatGroq

        _MODEL_NAME = os.getenv("LLM_MODEL") or api_keys._loaded["LLM"]
        _CLIENT = ChatGroq(
            model=_MODEL_NAME, temperature=0, max_tokens=4096,
            reasoning_format="hidden", reasoning_effort="low",
        )
        _JSON_SCHEMA_CLIENT = _CLIENT.bind(response_format=_RESPONSE_FORMAT)
        return _JSON_SCHEMA_CLIENT, str(_MODEL_NAME)
    except Exception as exc:
        raise RuntimeError("LLM unavailable; check project GROQ_API, LLM, and model configuration") from exc


def _as_dict(value):
    if isinstance(value, BaseModel):
        value = value.model_dump()
    elif hasattr(value, "content"):
        value = value.content
    for _ in range(3):
        if isinstance(value, dict):
            if "name" in value and "summary" in value:
                return value
            value = value.get("content", value.get("text", value))
            continue
        if isinstance(value, list):
            value = "\n".join(
                block.get("text", "") if isinstance(block, dict) else str(block)
                for block in value
            )
            continue
        break
    if isinstance(value, str):
        text = value.strip()
        if text.startswith("```"):
            lines = text.splitlines()
            text = "\n".join(lines[1:-1] if lines[-1].strip() == "```" else lines[1:])
        parsed = json.loads(text)
        if isinstance(parsed, dict):
            return parsed
        raise ValueError("LLM JSON response must be an object")
    raise ValueError("LLM response did not contain a JSON object")


def _normal_expression(value):
    output, quoted, i = [], False, 0
    while i < len(value):
        char = value[i]
        if char == '"':
            output.append(char)
            if quoted and i + 1 < len(value) and value[i + 1] == '"':
                output.append('"')
                i += 2
                continue
            quoted = not quoted
        elif quoted:
            output.append(char)
        elif not char.isspace():
            output.append(char.casefold())
        i += 1
    return "".join(output)


def _validate_procedure(result, procedure, line_count, required_calculations):
    result = _as_dict(result)
    for finding in result.get("findings", []):
        if finding.get("category") == "output":
            finding["category"] = "input_output"
    result = ProcedureAnalysis.model_validate(result).model_dump()
    if result.get("name", "").casefold() != procedure["name"].casefold():
        raise ValueError("LLM response procedure name did not match the requested procedure")
    low, high = procedure["start"], procedure["end"]

    def check_lines(values):
        if not isinstance(values, list) or not values:
            raise ValueError("LLM finding is missing source line citations")
        if any(not isinstance(line, int) or line < low or line > high or line > line_count for line in values):
            raise ValueError("LLM cited a line outside the requested procedure")

    check_lines(result.get("summary_source_lines"))
    actual_calculations = Counter()
    for finding in result.get("findings", []):
        check_lines(finding.get("source_lines"))
        if finding.get("category") == "calculation":
            for line in finding["source_lines"]:
                actual_calculations[(line, _normal_expression(finding.get("exact_vba", "")))] += 1
    expected_calculations = Counter(
        (item["source_line"], _normal_expression(item["expression"]))
        for item in required_calculations
    )
    missing = expected_calculations - actual_calculations
    if missing:
        for item in required_calculations:
            key = (item["source_line"], _normal_expression(item["expression"]))
            if not missing[key]:
                continue
            check_lines([item["source_line"]])
            result["findings"].append({
                "category": "calculation",
                "description": "Deterministic source inventory entry; interpretation not supplied by the LLM.",
                "exact_vba": item["expression"],
                "source_lines": [item["source_line"]],
            })
            missing[key] -= 1
    return result


def analyze_source(source, filename, procedures, progress=print, calculations=None):
    """Return structured explanations; isolate errors to individual procedures."""
    if os.getenv("VBA_NOTES_DISABLE_LLM") == "1":
        return "Skipped — VBA_NOTES_DISABLE_LLM=1", {"procedures": [], "errors": []}
    if not procedures:
        return "Skipped — no procedures found", {"procedures": [], "errors": []}
    try:
        client, model = _client()
    except Exception as exc:
        return f"Unavailable — {type(exc).__name__}: {exc}", {"procedures": [], "errors": []}

    source_lines = source.splitlines()
    calculations = calculations or []
    results, errors = [], []
    total = len(procedures)
    for index, procedure in enumerate(procedures, 1):
        progress(f"[LLM] PROCEDURE {index}/{total} — {procedure['name']}")
        excerpt = "\n".join(
            f"{number}: {source_lines[number - 1]}"
            for number in range(procedure["start"], min(procedure["end"], len(source_lines)) + 1)
        )
        prompt = (
        "Explain this one VBA procedure for later Java implementation. Do not write Java.\n"
            "Return exactly one JSON object with keys name, summary, summary_source_lines, "
            "confidence, and findings. Each finding must have category, description, and source_lines. "
            "Include exact_vba when applicable; otherwise omit it or use an empty string. "
            "category must be input_output, calculation, decision, loop, data_effect, dependency, or uncertainty.\n"
            "Use only the supplied lines. Preserve exact formulas and conditions in exact_vba. "
            "Cover inputs/outputs, calculations, branch outcomes, loop bounds, reads/writes, "
            "calls and error behavior when present. Put ambiguity in an uncertainty finding; "
            "do not guess. Every summary and finding must cite one or more supplied 1-based line "
            "numbers in source_lines. Use an empty findings list when no finding is supported.\n\n"
            "You must produce one calculation finding for each required_calculation entry below. "
            "Do not combine separate assignments into one finding. The finding's exact_vba must be that expression verbatim.\n"
            f"File: {filename}\nRequested procedure: {procedure['name']}\n"
            "Required calculation findings (include every item with exact_vba and its source_line):\n"
            f"{json.dumps([item for item in calculations if procedure['start'] <= item['source_line'] <= procedure['end']], ensure_ascii=False)}\n"
            f"Procedure source (line-numbered):\n{excerpt}"
        )
        try:
            response = client.invoke(prompt)
            required = [item for item in calculations if procedure["start"] <= item["source_line"] <= procedure["end"]]
            results.append(_validate_procedure(response, procedure, len(source_lines), required))
        except Exception as exc:
            message = f"{procedure['name']}: {type(exc).__name__}: {exc}"
            errors.append(message)
            progress(f"[LLM] PROCEDURE FAILED — {message}")
            return (
                f"Failed — {len(results)}/{total} procedures completed before {procedure['name']} failed; {model}",
                {"procedures": results, "errors": errors},
            )

    count = len(results)
    if count == total:
        status = f"Complete — {count}/{total} procedures; {model}; structured output and citations checked"
    else:
        status = f"Unavailable — 0/{total} procedures completed; {model}"
    return status, {"procedures": results, "errors": errors}
