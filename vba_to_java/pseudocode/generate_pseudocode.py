#!/usr/bin/env python3
"""Generate language-neutral, human-readable pseudocode from VBA graph outputs.

Expected layout (relative to PROJECT_ROOT):
    files/vba_graph_output/sample_1/graph.json
    files/vba_graph_output/sample_1/segments.json
    files/vba_graph_output/sample_1/report.json         # optional
    files/vba_graph_output/sample_1/manifest.json       # optional

Writes pseudocode.md and pseudocode_run.json into each sample directory.

The graph is the structural source of truth. Segment/source snapshots are not
used as a replacement for graph structure. Node-level source fragments are
included as evidence for the LLM to explain operations faithfully.
"""
from __future__ import annotations

import importlib
import json
import logging
import os
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

LOG = logging.getLogger("vba_pseudocode")

SYSTEM_PROMPT = """You are a senior software analyst documenting legacy VBA behavior.
Convert the supplied VBA graph/tree and node-level source fragments into
clear, precise, language-neutral PSEUDOCODE. The result will be used later as
an intermediate representation for translation into different programming
languages, so do not emit Java, Python, C#, VBA, or any other target-language
syntax.

Rules:
1. Explain what the code does, not merely restate each source line.
2. Preserve control flow, evaluation order, loop bounds, branch conditions,
   early exits, error handling, state changes, side effects, and call order.
3. Preserve identifiers, constants, formulas, field names, external calls,
   file/database/workbook interactions, and meaningful literal values exactly.
4. Describe every procedure's inputs, outputs/return value, state read/written,
   calls/dependencies, and observable side effects when supported by evidence.
5. For branches and loops, use explicit IF/ELSE, FOR EACH, WHILE, SELECT CASE,
   TRY/ON-ERROR, etc. in language-neutral prose/pseudocode.
6. Do not invent missing behavior, types, business rules, or dependencies.
   Mark uncertainty as [UNCERTAIN: ...] and unsupported/opaque code as
   [OPAQUE: preserve/inspect original construct ...].
7. Distinguish facts directly visible in the graph/source fragments from
   reasonable interpretation. Do not silently repair parser diagnostics.
8. Do not omit small statements, error paths, labels, GoTo/GoSub/Return,
   declarations, or unusual VBA constructs. Explain their control-flow effect.
9. Do not include raw VBA source listings. Use readable pseudocode statements.
10. Output Markdown only, with the requested headings. No preamble about being
    an AI and no target-language implementation.
"""


def json_compact(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"), default=str)


def load_json(path: Path, required: bool = False) -> Any:
    if not path.exists():
        if required:
            raise FileNotFoundError(path)
        return None
    with path.open("r", encoding="utf-8-sig") as handle:
        return json.load(handle)


def pick_graph_file(sample_dir: Path) -> Path | None:
    preferred = sample_dir / "graph.json"
    if preferred.is_file():
        return preferred
    candidates = sorted(p for p in sample_dir.glob("graph*.json") if p.is_file())
    return candidates[0] if len(candidates) == 1 else None


def node_source(node: dict[str, Any]) -> str:
    source = node.get("source")
    if isinstance(source, dict):
        value = source.get("text") or source.get("code")
        return value if isinstance(value, str) else ""
    return source if isinstance(source, str) else ""


def source_span(node: dict[str, Any]) -> dict[str, Any] | None:
    for key in ("span", "source_span", "location"):
        value = node.get(key)
        if isinstance(value, dict):
            return value
    source = node.get("source")
    if isinstance(source, dict):
        for key in ("span", "source_span", "location"):
            value = source.get(key)
            if isinstance(value, dict):
                return value
    return None


def build_tree(nodes: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], dict[str, dict[str, Any]]]:
    """Build parent/child tree while retaining original node fields."""
    by_id: dict[str, dict[str, Any]] = {}
    for raw in nodes:
        if not isinstance(raw, dict) or not raw.get("id"):
            continue
        item = {
            "id": raw.get("id"),
            "type": raw.get("type", "UNKNOWN"),
            "name": raw.get("name"),
            "parent_id": raw.get("parent_id"),
            "metadata": raw.get("metadata") or {},
            "source_text": node_source(raw),
            "span": source_span(raw),
            "children": [],
        }
        by_id[str(item["id"])] = item
    roots: list[dict[str, Any]] = []
    for item in by_id.values():
        parent_id = item.get("parent_id")
        parent = by_id.get(str(parent_id)) if parent_id is not None else None
        if parent is None or parent is item:
            roots.append(item)
        else:
            parent["children"].append(item)
    # Stable source order: source span line first, then original graph order.
    def line_of(item: dict[str, Any]) -> int:
        span = item.get("span") or {}
        for key in ("start_line", "line", "line_start"):
            if isinstance(span.get(key), int):
                return span[key]
        return 10**12
    for item in by_id.values():
        item["children"].sort(key=line_of)
    roots.sort(key=line_of)
    return roots, by_id


def node_payload(node: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": node.get("id"), "type": node.get("type"),
        "name": node.get("name"), "metadata": node.get("metadata") or {},
        "source_text": node.get("source_text") or "", "span": node.get("span"),
        "children": [node_payload(child) for child in node.get("children", [])],
    }


def flatten(node: dict[str, Any]) -> Iterable[dict[str, Any]]:
    yield node
    for child in node.get("children", []):
        yield from flatten(child)


def get_procedure_roots(nodes: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    roots, _ = build_tree(nodes)
    procedures: list[dict[str, Any]] = []
    module_level: list[dict[str, Any]] = []
    for root in roots:
        for item in flatten(root):
            if str(item.get("type", "")).upper() in {"PROCEDURE", "SUB", "FUNCTION", "PROPERTY"}:
                procedures.append(item)
        if str(root.get("type", "")).upper() not in {"PROCEDURE", "SUB", "FUNCTION", "PROPERTY"}:
            # Keep module/project trees; procedure subtrees are separately rendered.
            module_level.append(root)
    # A procedure may be nested under a MODULE root; avoid duplicate procedures
    # by returning unique IDs and use module-level nodes only for declarations.
    unique = {p["id"]: p for p in procedures}
    return list(unique.values()), module_level


def signature_from_node(proc: dict[str, Any]) -> dict[str, Any]:
    metadata = proc.get("metadata") or {}
    return {
        "name": proc.get("name"),
        "kind": metadata.get("kind") or metadata.get("procedure_type") or proc.get("type"),
        "visibility": metadata.get("visibility"),
        "parameters": metadata.get("parameters") or metadata.get("params"),
        "return_type": metadata.get("return_type"),
        "metadata": metadata,
    }


def chunk_children(proc: dict[str, Any], max_chars: int) -> list[list[dict[str, Any]]]:
    """Split large procedure bodies at child-subtree boundaries, never truncate."""
    children = proc.get("children", [])
    chunks: list[list[dict[str, Any]]] = []
    current: list[dict[str, Any]] = []
    size = 0
    for child in children:
        encoded = json_compact(node_payload(child))
        if current and size + len(encoded) > max_chars:
            chunks.append(current)
            current, size = [], 0
        if len(encoded) > max_chars:
            # Preserve the complete oversized subtree by recursively splitting its
            # direct children and carrying its block header/context with each part.
            descendants = child.get("children", [])
            if not descendants:
                chunks.append([child])
                continue
            header = {k: v for k, v in child.items() if k != "children"}
            sub_current: list[dict[str, Any]] = []
            sub_size = len(json_compact(header))
            for grandchild in descendants:
                gsize = len(json_compact(node_payload(grandchild)))
                if sub_current and sub_size + gsize > max_chars:
                    chunks.append([{**header, "children": sub_current}])
                    sub_current, sub_size = [], len(json_compact(header))
                sub_current.append(grandchild)
                sub_size += gsize
            if sub_current:
                chunks.append([{**header, "children": sub_current}])
            continue
        current.append(child)
        size += len(encoded)
    if current:
        chunks.append(current)
    return chunks or [[]]


@dataclass
class LLMAdapter:
    client: Any
    model: str
    temperature: float = 0.0
    max_tokens: int = 5000

    @classmethod
    def from_project(cls, project_root: Path) -> "LLMAdapter":
        if str(project_root) not in sys.path:
            sys.path.insert(0, str(project_root))
        try:
            api_keys = importlib.import_module("api_keys")
            from langchain_groq import ChatGroq

            model = api_keys._loaded["LLM"]
            client = ChatGroq(model=model, temperature=0)
        except Exception as exc:
            raise RuntimeError(
                "Could not initialize ChatGroq. Confirm the project environment has "
                "langchain-groq installed and api_keys loads GROQ_API and LLM."
            ) from exc
        return cls(client=client, model=str(model))

    @staticmethod
    def _extract_text(response: Any) -> str:
        if isinstance(response, str):
            return response.strip()
        # OpenAI Chat Completions response
        choices = getattr(response, "choices", None)
        if choices:
            message = getattr(choices[0], "message", None)
            content = getattr(message, "content", None)
            if isinstance(content, str):
                return content.strip()
        # OpenAI Responses API response
        output_text = getattr(response, "output_text", None)
        if isinstance(output_text, str):
            return output_text.strip()
        # LangChain AIMessage
        content = getattr(response, "content", None)
        if isinstance(content, str):
            return content.strip()
        if isinstance(content, list):
            pieces = []
            for part in content:
                if isinstance(part, dict) and isinstance(part.get("text"), str):
                    pieces.append(part["text"])
                elif isinstance(getattr(part, "text", None), str):
                    pieces.append(part.text)
            if pieces:
                return "\n".join(pieces).strip()
        if isinstance(response, dict):
            if isinstance(response.get("output_text"), str):
                return response["output_text"].strip()
            try:
                return str(response["choices"][0]["message"]["content"]).strip()
            except Exception:
                pass
        raise TypeError(f"Unsupported LLM response shape: {type(response).__name__}")

    def generate(self, user_prompt: str) -> str:
        client = self.client
        # OpenAI-compatible SDK client
        chat = getattr(client, "chat", None)
        completions = getattr(chat, "completions", None) if chat is not None else None
        create = getattr(completions, "create", None) if completions is not None else None
        if callable(create):
            response = create(
                model=self.model,
                messages=[{"role": "system", "content": SYSTEM_PROMPT}, {"role": "user", "content": user_prompt}],
                temperature=self.temperature,
                max_tokens=self.max_tokens,
            )
            return self._extract_text(response)
        # OpenAI Responses API client
        responses = getattr(client, "responses", None)
        response_create = getattr(responses, "create", None) if responses is not None else None
        if callable(response_create):
            response = response_create(
                model=self.model,
                instructions=SYSTEM_PROMPT,
                input=user_prompt,
                temperature=self.temperature,
                max_output_tokens=self.max_tokens,
            )
            return self._extract_text(response)
        # LangChain ChatModel or compatible wrapper
        invoke = getattr(client, "invoke", None)
        if callable(invoke):
            response = invoke([("system", SYSTEM_PROMPT), ("human", user_prompt)])
            return self._extract_text(response)
        # Generic callable wrapper
        if callable(client):
            response = client(system=SYSTEM_PROMPT, prompt=user_prompt, model=self.model, temperature=self.temperature)
            return self._extract_text(response)
        raise TypeError(
            "api_keys._loaded['LLM'] is not a supported client. Expected an OpenAI-compatible "
            "chat/responses client, a LangChain object with invoke(), or a callable wrapper."
        )


def procedure_prompt(project_name: str, proc: dict[str, Any], children: list[dict[str, Any]],
                     part_no: int, part_count: int, edges: list[dict[str, Any]],
                     symbols: list[dict[str, Any]], diagnostics: list[dict[str, Any]],
                     supporting_segments: list[dict[str, Any]], report: dict[str, Any] | None,
                     manifest: Any) -> str:
    proc_id = proc.get("id")
    relevant_edges = [e for e in edges if e.get("source") == proc_id or e.get("target") == proc_id]
    node_ids = {n.get("id") for child in children for n in flatten(child)}
    relevant_edges += [e for e in edges if e.get("source") in node_ids or e.get("target") in node_ids]
    relevant_symbols = [s for s in symbols if str(s.get("scope", "")).find(str(proc.get("name", ""))) >= 0]
    payload = {
        "project": project_name,
        "procedure": signature_from_node(proc),
        "procedure_id": proc_id,
        "part": {"number": part_no, "total": part_count},
        "body_tree": [node_payload(n) for n in children],
        "related_edges": relevant_edges,
        "possibly_relevant_symbols": relevant_symbols,
        "supporting_non_procedure_segments": supporting_segments,
        "report_summary": report or {},
        "input_manifest_metadata": manifest or [],
        "parser_diagnostics": diagnostics,
    }
    return (
        "Write the pseudocode section for this procedure. If this is one of several parts, "
        "document only the supplied body fragment while retaining its surrounding procedure "
        "context; do not imply that the fragment is the whole procedure. Use heading "
        f"'### Procedure: {proc.get('name') or proc_id} (part {part_no}/{part_count})'.\n\n"
        "Use this evidence payload (JSON):\n" + json.dumps(payload, ensure_ascii=False, indent=2, default=str)
    )


def module_prompt(project_name: str, module_nodes: list[dict[str, Any]],
                  symbols: list[dict[str, Any]], diagnostics: list[dict[str, Any]],
                  supporting_segments: list[dict[str, Any]], report: dict[str, Any] | None,
                  manifest: Any) -> str:
    # Only top-level module/project declarations and metadata, not procedure bodies.
    def declarations_only(node: dict[str, Any]) -> dict[str, Any] | None:
        typ = str(node.get("type", "")).upper()
        if typ in {"PROCEDURE", "SUB", "FUNCTION", "PROPERTY"}:
            return None
        copied = {k: v for k, v in node.items() if k != "children"}
        copied["children"] = [x for c in node.get("children", []) if (x := declarations_only(c)) is not None]
        return copied
    payload = {
        "project": project_name,
        "module_and_global_declarations": [x for n in module_nodes if (x := declarations_only(n)) is not None],
        "symbols": symbols,
        "supporting_non_procedure_segments": supporting_segments,
        "report_summary": report or {},
        "input_manifest_metadata": manifest or [],
        "parser_diagnostics": diagnostics,
    }
    return (
        "Document the module-level setup and declarations as language-neutral pseudocode. "
        "Include constants, types, enums, module state, configuration, and initialization "
        "only when present. Do not invent initialization behavior. Use heading '## Module "
        "overview and shared declarations'.\n\nEvidence payload (JSON):\n" +
        json.dumps(payload, ensure_ascii=False, indent=2, default=str)
    )


def safe_filename(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9_.-]+", "_", value or "procedure")
    return value[:100] or "procedure"


def _validate_procedure_section(
    content: str, procedure_name: str, part_number: int, part_count: int
) -> str:
    if not isinstance(content, str) or not content.strip():
        raise ValueError("LLM returned an empty procedure section")

    canonical_heading = (
        f"### Procedure: {procedure_name} (part {part_number}/{part_count})"
    )
    heading_pattern = re.compile(
        r"(?m)^[ \t]*#{2,4}\s+Procedure:\s*(.+?)\s*$", re.IGNORECASE
    )
    match = heading_pattern.search(content)
    if match:
        returned_name = re.sub(
            r"\s*\(?part\s+\d+\s*/\s*\d+\)?\s*$",
            "",
            match.group(1),
            flags=re.IGNORECASE,
        ).strip()
        if returned_name.casefold() != procedure_name.casefold():
            raise ValueError(
                f"LLM section heading names {returned_name!r}, expected {procedure_name!r}"
            )
        content = content[:match.start()] + canonical_heading + content[match.end():]
        body = content.replace(canonical_heading, "", 1).strip()
    else:
        body = content.strip()
        content = f"{canonical_heading}\n\n{body}"

    if not body:
        raise ValueError("LLM returned a procedure heading without pseudocode")
    return content.strip()


def process_sample(sample_dir: Path, llm: LLMAdapter, max_chars: int, retries: int) -> dict[str, Any]:
    graph_path = pick_graph_file(sample_dir)
    if graph_path is None:
        return {"sample": sample_dir.name, "status": "skipped", "reason": "No unique graph.json / graph*.json found"}
    graph = load_json(graph_path, required=True)
    if not isinstance(graph, dict) or not isinstance(graph.get("nodes"), list):
        raise ValueError(f"Invalid graph schema in {graph_path}: expected object with nodes[]")
    nodes = graph["nodes"]
    edges = graph.get("edges") or []
    symbols = graph.get("symbols") or []
    diagnostics = graph.get("diagnostics") or []
    segments = load_json(sample_dir / "segments.json", required=True)
    if not isinstance(segments, list) or any(not isinstance(seg, dict) for seg in segments):
        raise ValueError(f"Invalid segments file in {sample_dir}: expected an array of objects")
    report = load_json(sample_dir / "report.json") or {}
    manifest = load_json(sample_dir / "manifest.json") or graph.get("manifest") or []
    # Procedure snapshots are intentionally excluded. Granular segment fragments
    # are supplementary evidence only; graph hierarchy remains authoritative.
    non_procedure_segments = [
        {k: v for k, v in seg.items() if k in {"id", "parent_id", "kind", "name", "path", "start_line", "end_line", "source", "summary", "annotations"}}
        for seg in segments if isinstance(seg, dict) and str(seg.get("kind", "")).upper() != "PROCEDURE"
    ]
    procedures, module_nodes = get_procedure_roots(nodes)
    project = graph.get("project")
    project_name = (
        str(project.get("name"))
        if isinstance(project, dict) and project.get("name")
        else str(project or sample_dir.name)
    )
    sections: list[str] = [
        f"# Pseudocode: {project_name}",
        "",
        "> Generated from the VBA structural graph. This is a language-neutral behavioral "
        "description for downstream intermediate-representation design, not executable code.",
        "",
        "## Translation notes and known limitations",
        f"- Graph nodes: {len(nodes)}; graph edges: {len(edges)}; procedures detected: {len(procedures)}.",
        f"- Parser diagnostics reported: {len(diagnostics)}.",
        "- Treat every `[UNCERTAIN]` / `[OPAQUE]` marker as a review item before translation.",
        f"- Supporting granular segments supplied (procedure snapshots excluded): {len(non_procedure_segments)}.",
        "- The graph may omit comments, formatting, or syntax it could not model; consult the "
        "original VBA for unresolved constructs.",
        "",
    ]
    status: dict[str, Any] = {
        "sample": sample_dir.name, "graph_file": graph_path.name,
        "project": project_name, "node_count": len(nodes), "edge_count": len(edges),
        "procedure_count": len(procedures), "diagnostics": diagnostics,
        "generated_sections": 0, "generated_section_details": [],
        "failed_sections": [],
    }

    if module_nodes:
        module_node_ids: set[str] = set()
        def collect_module_ids(node: dict[str, Any]) -> None:
            if str(node.get("type", "")).upper() in {"PROCEDURE", "SUB", "FUNCTION", "PROPERTY"}:
                return
            module_node_ids.add(str(node.get("id")))
            for child in node.get("children", []):
                collect_module_ids(child)
        for module_root in module_nodes:
            collect_module_ids(module_root)
        module_segments = [seg for seg in non_procedure_segments
                           if str(seg.get("parent_id")) in module_node_ids]
        prompt = module_prompt(project_name, module_nodes, symbols, diagnostics,
                              module_segments, report, manifest)
        last_error = None
        for attempt in range(retries + 1):
            try:
                content = llm.generate(prompt)
                if not isinstance(content, str) or not content.strip():
                    raise ValueError("LLM returned an empty module overview")
                sections.append(content.strip())
                status["generated_sections"] += 1
                status["generated_section_details"].append({
                    "section": "module_overview",
                    "status": "generated",
                    "characters": len(content.strip()),
                })
                last_error = None
                break
            except Exception as exc:
                last_error = repr(exc)
                LOG.warning("%s: module overview attempt %s failed: %s", sample_dir.name, attempt + 1, exc)
                if attempt < retries:
                    time.sleep(min(2 ** attempt, 8))
        if last_error:
            status["failed_sections"].append({"section": "module_overview", "error": last_error})
            sections.append("## Module overview and shared declarations\n\n[GENERATION FAILED: " + last_error + "]")

    completed_procedure_ids = set()
    for proc_index, proc in enumerate(procedures, 1):
        procedure_name = proc.get("name") or proc.get("id")
        procedure_complete = True
        chunks = chunk_children(proc, max_chars=max_chars)
        for part_index, children in enumerate(chunks, 1):
            subtree_ids = {n.get("id") for child in children for n in flatten(child)}
            subtree_ids.add(proc.get("id"))
            supporting_segments = [seg for seg in non_procedure_segments
                                   if seg.get("parent_id") in subtree_ids]
            prompt = procedure_prompt(project_name, proc, children, part_index, len(chunks),
                                      edges, symbols, diagnostics, supporting_segments, report, manifest)
            last_error = None
            for attempt in range(retries + 1):
                try:
                    content = _validate_procedure_section(
                        llm.generate(prompt),
                        str(procedure_name),
                        part_index,
                        len(chunks),
                    )
                    sections.append(content)
                    status["generated_sections"] += 1
                    status["generated_section_details"].append({
                        "procedure": procedure_name,
                        "procedure_id": proc.get("id"),
                        "part": part_index,
                        "status": "generated",
                        "characters": len(content),
                    })
                    last_error = None
                    break
                except Exception as exc:
                    last_error = repr(exc)
                    LOG.warning("%s/%s part %s attempt %s failed: %s", sample_dir.name, proc.get("name"), part_index, attempt + 1, exc)
                    if attempt < retries:
                        time.sleep(min(2 ** attempt, 8))
            if last_error:
                procedure_complete = False
                record = {"procedure": procedure_name, "procedure_id": proc.get("id"), "part": part_index, "error": last_error}
                status["failed_sections"].append(record)
                sections.append(
                    f"### Procedure: {procedure_name} (part {part_index}/{len(chunks)})\n\n"
                    f"[GENERATION FAILED: {last_error}]"
                )
        if procedure_complete:
            completed_procedure_ids.add(proc.get("id"))

    coverage_gaps = [
        proc.get("name") or proc.get("id")
        for proc in procedures
        if proc.get("id") not in completed_procedure_ids
    ]
    status["pseudocode_coverage_gaps"] = coverage_gaps

    if diagnostics:
        sections.extend(["", "## Parser diagnostics (verbatim from graph)", ""])
        for diag in diagnostics:
            sections.append(f"- **{diag.get('severity', 'UNKNOWN')} / {diag.get('code', 'UNKNOWN')}**"
                            f" — {diag.get('message', '')}"
                            + (f" (line {diag.get('line')})" if diag.get("line") is not None else ""))
    output_path = sample_dir / "pseudocode.md"
    output_path.write_text("\n\n".join(s.strip() for s in sections if s is not None) + "\n", encoding="utf-8")
    status["output_file"] = output_path.name
    status["status"] = (
        "partial"
        if status["failed_sections"] or status["pseudocode_coverage_gaps"]
        else "success"
    )
    (sample_dir / "pseudocode_run.json").write_text(json.dumps(status, ensure_ascii=False, indent=2, default=str) + "\n", encoding="utf-8")
    return status


def main(input_path: Path, output: Path) -> int:
    project_root = Path(__file__).resolve().parents[2]
    source_path = Path(input_path).expanduser().resolve()
    sample_dir = Path(output).expanduser().resolve()
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    try:
        if not source_path.is_file():
            raise FileNotFoundError(f"VBA source file does not exist: {source_path}")
        graph_path = pick_graph_file(sample_dir)
        if graph_path is None:
            raise FileNotFoundError(f"No graph.json found in output directory: {sample_dir}")
        graph = load_json(graph_path, required=True)
        project = graph.get("project") if isinstance(graph, dict) else None
        graph_project = project.get("name") if isinstance(project, dict) else project
        if graph_project and str(graph_project) != source_path.stem:
            raise ValueError(
                f"Input {source_path.name} does not match graph project {graph_project!r}"
            )
        if not (sample_dir / "segments.json").is_file():
            raise FileNotFoundError(f"Required segments file does not exist: {sample_dir / 'segments.json'}")

        llm = LLMAdapter.from_project(project_root)
        result = process_sample(sample_dir, llm, max_chars=45000, retries=2)
        LOG.info(
            "%s: %s -> %s",
            sample_dir.name,
            result.get("status"),
            result.get("output_file", result.get("reason", "")),
        )
        return 0 if result.get("status") == "success" else 1
    except Exception:
        LOG.exception("Failed processing %s", source_path)
        return 1


if __name__ == "__main__":
    project_root = Path(__file__).resolve().parents[2]
    output_root = project_root / "files" / "vba_graph_output"
    n = [6]
    status = 0

    for i in n:
        input_path = project_root / "files" / f"vba_sample_{i}.bas"
        output = output_root / f"sample_{i}"
        status = max(status, main(input_path, output))

    raise SystemExit(status)
