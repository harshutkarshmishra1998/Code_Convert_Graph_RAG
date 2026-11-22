"""Build Universal IR from VBA graph + granular segments + optional pseudocode.

The builder deliberately does NOT consume full MODULE/PROCEDURE source snapshots.
It uses graph metadata and node-level source fragments, plus non-PROCEDURE segments.
Expressions remain explicitly source-backed/opaque unless the frontend parsed them.
"""
from __future__ import annotations

import json
import re
import hashlib
from collections import defaultdict
from pathlib import Path
from typing import Any

from .io import IR_VERSION

VBA_PRIMITIVES = {
    "boolean": "boolean", "byte": "uint8", "integer": "int16", "long": "int32",
    "longlong": "int64", "longptr": "intptr", "single": "float32", "double": "float64", "currency": "decimal",
    "decimal": "decimal", "date": "datetime", "string": "string", "object": "object",
    "variant": "variant", "void": "void"
}


def _type_spec(source_type: str | None) -> dict[str, Any]:
    if not source_type:
        return {"kind": "unknown", "name": "unresolved", "source_type": source_type}
    raw = source_type.strip()
    low = raw.lower()
    # Preserve array rank and dynamic/fixed array syntax rather than erasing it.
    array = raw.endswith("()") or "(" in raw and ")" in raw
    base = re.sub(r"\s*\([^)]*\)\s*$", "", raw).strip()
    base_low = base.lower()
    if array:
        return {
            "kind": "array", "element_type": _type_spec(base),
            "dimensions": [{"kind": "dynamic" if not raw[raw.find("(") + 1:raw.rfind(")")].strip() else "source_declared",
                            "bounds": raw[raw.find("(") + 1:raw.rfind(")")] if "(" in raw else "dynamic"}],
            "source_type": raw, "semantics": {"vba_array": True}
        }
    if base_low == "collection":
        return {"kind": "collection", "element_type": {"kind": "variant", "name": "Variant"},
                "source_type": raw, "semantics": {"vba_collection": True, "index_base": 1,
                "ordered": True, "keyed_items_supported": True}}
    if base_low in VBA_PRIMITIVES:
        mapped = VBA_PRIMITIVES[base_low]
        if base_low == "variant":
            return {"kind": "variant", "name": "Variant", "source_type": raw, "semantics": {"dynamic_type": True}}
        if base_low == "object":
            return {"kind": "object", "name": "Object", "source_type": raw, "semantics": {"late_bound_possible": True}}
        if mapped == "void":
            return {"kind": "void", "name": "void", "source_type": raw}
        semantics = {"vba_type": base}
        if base_low == "currency": semantics["fixed_scale"] = 4
        if base_low == "date": semantics["vba_date_representation"] = "OLE Automation date serial"
        if base_low == "boolean": semantics["vba_true_value"] = -1
        if base_low == "longptr": semantics["pointer_size"] = "platform_dependent"
        return {"kind": "primitive", "name": mapped, "source_type": raw, "semantics": semantics}
    return {"kind": "named", "name": base, "source_type": raw, "semantics": {"resolution": "unresolved_named_type"}}


def _simple_literal_expr(source: str | None) -> dict[str, Any]:
    if source is None:
        return _expr(None)
    value = str(source).strip()
    if re.fullmatch(r"[-+]?\d+", value):
        return {"kind": "literal", "value": int(value), "semantics": {"source_literal": value}}
    if re.fullmatch(r"[-+]?(?:\d+\.\d*|\.\d+)(?:[Ee][-+]?\d+)?", value):
        return {"kind": "literal", "value": float(value), "semantics": {"source_literal": value}}
    if value.startswith('"') and value.endswith('"'):
        return {"kind": "literal", "value": value[1:-1].replace('""', '"'), "semantics": {"source_literal": value}}
    if value.lower() in {"true", "false"}:
        return {"kind": "literal", "value": value.lower() == "true", "semantics": {"source_literal": value}}
    return _expr(value)


def _expr(source: str | None, *, reason: str = "Frontend supplied source text but no expression AST") -> dict[str, Any]:
    """Never fabricate an AST: retain exact expression source as an opaque node."""
    if source is None or not str(source).strip():
        return {"kind": "unsupported", "semantics": {"reason": "Expression source is missing", "source_expression": source}}
    return {"kind": "unsupported", "semantics": {
        "reason": reason, "source_expression": str(source), "parse_status": "not_parsed",
        "requires_semantic_resolution": True
    }}


def _source(node: dict[str, Any]) -> str:
    src = node.get("source") or {}
    return str(src.get("text") or "").strip()


def _location(node: dict[str, Any], source_path: str | None) -> dict[str, Any]:
    src = node.get("source") or {}
    loc: dict[str, Any] = {"line": src.get("start_line", 0)}
    if src.get("end_line") is not None:
        loc["end_line"] = src["end_line"]
    if source_path or src.get("path"):
        loc["file"] = source_path or src.get("path")
    return loc


def _provenance(node: dict[str, Any], segment_ids: list[str], source_path: str | None,
                segment_sources: list[dict[str, Any]] | None = None, include_node_source: bool = True) -> dict[str, Any]:
    result = {
        "source_language": "VBA", "source_location": _location(node, source_path),
        "graph_node_ids": [node["id"]], "segment_ids": segment_ids,
        "segment_sources": segment_sources or [], "inferred": False
    }
    if include_node_source:
        result["source_text"] = _source(node)
    return result


def _decl_from_line(line: str, node: dict[str, Any], seg_ids: list[str], source_path: str | None) -> dict[str, Any]:
    clean = line.strip()
    vis = "private" if clean.lower().startswith("private ") else "public" if clean.lower().startswith("public ") else "default"
    body = re.sub(r"^(public|private|global|dim|static)\s+", "", clean, flags=re.I).strip()
    if body.lower().startswith("const "):
        m = re.match(r"Const\s+(\w+)\s+As\s+([\w.]+)\s*=\s*(.*)", body, re.I)
        return {"id": node["id"], "kind": "constant", "name": m.group(1) if m else node.get("name", "unknown"),
                "visibility": vis, "type": _type_spec(m.group(2)) if m else {"kind": "unknown"},
                "value": _expr(m.group(3)) if m else _expr(None), "semantics": {"source_declaration": clean},
                "provenance": _provenance(node, seg_ids, source_path)}
    if body.lower().startswith("enum ") or body.lower().startswith("type "):
        enum = body.lower().startswith("enum ")
        name = re.match(r"(?:Enum|Type)\s+(\w+)", body, re.I)
        return {"id": node["id"], "kind": "enum" if enum else "record", "name": name.group(1) if name else node.get("name", "unknown"),
                "visibility": vis, "members": [], "semantics": {"source_declaration": clean, "definition_completeness": "header_only",
                "warning": "Frontend graph contains only the declaration header; member declarations were not supplied as separate nodes."},
                "provenance": _provenance(node, seg_ids, source_path)}
    m = re.match(r"(\w+)\s*(\([^)]*\))?\s+As\s+([\w.]+(?:\([^)]*\))?)", body, re.I)
    if m:
        name, arr, typ = m.groups()
        if arr:
            typ += arr
        return {"id": node["id"], "kind": "variable", "name": name, "visibility": vis,
                "type": _type_spec(typ), "modifiers": [], "semantics": {"source_declaration": clean},
                "provenance": _provenance(node, seg_ids, source_path)}
    return {"id": node["id"], "kind": "unsupported", "name": node.get("name", "unknown"),
            "visibility": vis, "semantics": {"source_declaration": clean, "reason": "Declaration syntax not recognized; preserved without interpretation."},
            "provenance": _provenance(node, seg_ids, source_path)}


def _inline_child_statement(source: str, parent_id: str, suffix: str) -> dict[str, Any]:
    text = source.strip()
    base = {"id": f"{parent_id}:{suffix}", "provenance": {"source_language": "VBA", "source_text": text,
            "graph_node_ids": [], "segment_ids": [], "inferred": True,
            "inference_reason": "Inline statement split from a single-line If node."}}
    m = re.match(r"(?:Set\s+)?([\w.]+(?:\([^)]*\))?)\s*=\s*(.+)$", text, re.I)
    if m:
        return {**base, "kind": "assignment", "target": {"kind": "identifier", "name": m.group(1)},
                "value": _expr(m.group(2)), "semantics": {"raw_source": text, "inline_if_body": True}}
    m = re.match(r"GoTo\s+(\w+)$", text, re.I)
    if m:
        return {**base, "kind": "goto", "label": m.group(1), "semantics": {"inline_if_body": True}}
    m = re.match(r"Exit\s+(Sub|Function|Property|For|Do)$", text, re.I)
    if m:
        return {**base, "kind": "exit", "exit_target": m.group(1).lower(), "semantics": {"inline_if_body": True}}
    return {**base, "kind": "call", "expression": _expr(text),
            "semantics": {"raw_source": text, "call_resolution": "unresolved", "inline_if_body": True}}


def _split_inline_else(text: str) -> tuple[str, str | None]:
    # Locate Else only outside quoted strings and balanced parentheses.
    depth = 0; in_string = False; i = 0
    while i < len(text):
        ch = text[i]
        if ch == '"':
            if in_string and i + 1 < len(text) and text[i + 1] == '"':
                i += 2; continue
            in_string = not in_string
        elif not in_string:
            if ch == "(": depth += 1
            elif ch == ")": depth = max(0, depth - 1)
            elif depth == 0 and text[i:i+6].lower() == " else ":
                return text[:i].strip(), text[i+6:].strip()
        i += 1
    return text.strip(), None


def _statement(node: dict[str, Any], children: list[dict[str, Any]], seg_ids: list[str], source_path: str | None) -> dict[str, Any] | None:
    typ = node.get("type", "")
    meta = node.get("metadata") or {}
    raw = _source(node)
    prov = _provenance(node, seg_ids, source_path, node.get("_segment_sources"))
    sid = node["id"]
    if typ in {"IF_BRANCH", "ELSE_BRANCH", "ELSEIF_BRANCH"}:
        return None  # consumed by IF_BLOCK
    if typ == "IF_BLOCK":
        branches = []
        ordered_branches = sorted(children, key=lambda b: b.get("_branch_order", 50))
        for branch in ordered_branches:
            btype = branch.get("type")
            if btype not in {"IF_BRANCH", "ELSEIF_BRANCH", "ELSE_BRANCH"}:
                continue
            bmeta = branch.get("metadata") or {}
            cond = (bmeta.get("condition") or {}).get("source")
            bchildren = []
            for child in branch.get("_children", []):
                stmt = _statement(child, child.get("_children", []), child.get("_segment_ids", []), source_path)
                if stmt is not None:
                    bchildren.append(stmt)
            branch_ir = {"body": bchildren, "branch_kind": btype.lower(),
                         "provenance": _provenance(branch, branch.get("_segment_ids", []), source_path,
                                                    branch.get("_segment_sources"))}
            if btype != "ELSE_BRANCH" and cond:
                branch_ir["condition"] = _expr(cond)
            branches.append(branch_ir)
        first_condition = (meta.get("condition") or {}).get("source")
        return {"id": sid, "kind": "if", "branches": branches,
                "condition": _expr(first_condition),
                "provenance": prov, "semantics": {"source_control_type": "IF_BLOCK", "branch_count": len(branches)}}
    if typ == "FOR_BLOCK":
        def ex(k):
            v = meta.get(k)
            return _expr(v.get("source") if isinstance(v, dict) else None)
        body = []
        terminator = None
        for child in children:
            child_raw = _source(child)
            if child.get("type") == "CALL_OR_COMMAND" and re.match(r"^Next\b", child_raw, re.I):
                terminator = {"source_text": child_raw, "graph_node_id": child["id"],
                              "provenance": _provenance(child, child.get("_segment_ids", []), source_path,
                                                         child.get("_segment_sources"))}
                continue
            stmt = _statement(child, child.get("_children", []), child.get("_segment_ids", []), source_path)
            if stmt is not None: body.append(stmt)
        return {"id": sid, "kind": "for", "variable": meta.get("variable", ""),
                "start": ex("start"), "end": ex("end"), "step": ex("step"), "body": body,
                "provenance": prov, "semantics": {"source_control_type": "FOR_BLOCK", "raw_header": raw,
                                                     "terminator": terminator}}
    # Preserve nested block-like constructs we cannot safely lower yet.
    if typ in {"WHILE_BLOCK", "DO_BLOCK", "SELECT_BLOCK", "WITH_BLOCK", "CASE_BLOCK"}:
        body = []
        for child in children:
            stmt = _statement(child, child.get("_children", []), child.get("_segment_ids", []), source_path)
            if stmt is not None: body.append(stmt)
        return {"id": sid, "kind": "unsupported", "body": body,
                "semantics": {"reason": f"Control construct {typ} requires a dedicated frontend mapping", "raw_source": raw}, "provenance": prov}
    if typ == "ERROR_CONTROL":
        match = re.search(r"On\s+Error\s+(GoTo\s+\w+|Resume\s+Next|GoTo\s+0)", raw, re.I)
        form = match.group(1) if match else "unparsed"
        mode = "resume_next" if form.lower() == "resume next" else "disable_handler" if form.lower() == "goto 0" else "goto_handler" if form.lower().startswith("goto") else "unknown"
        handler = form.split(None, 1)[1] if mode == "goto_handler" and len(form.split(None, 1)) > 1 else None
        on_error_stmt = {"id": sid, "kind": "on_error", "mode": mode, "provenance": prov,
                         "semantics": {"raw_source": raw, "requires_resolution": mode == "unknown", "vba_error_model": True}}
        if handler:
            on_error_stmt["handler"] = handler
        return on_error_stmt
    if typ == "JUMP":
        operation = str(meta.get("operation", "JUMP")).upper()
        return {"id": sid, "kind": "gosub" if operation == "GOSUB" else "goto",
                "label": meta.get("target") or node.get("name", ""), "provenance": prov,
                "semantics": {"source_operation": operation, "raw_source": raw}}
    if typ == "RETURN_CONTROL":
        return {"id": sid, "kind": "gosub_return", "provenance": prov,
                "semantics": {"legacy_gosub_return": True, "raw_source": raw,
                              "note": "VBA Return in a GoSub routine resumes after the GoSub call; it is not a function return."}}
    if typ == "EXIT":
        return {"id": sid, "kind": "exit", "exit_target": (meta.get("target") or node.get("name", "")).lower(),
                "provenance": prov}
    if typ == "ASSIGNMENT":
        target = meta.get("target") or node.get("name")
        expression = (meta.get("expression") or {}).get("source")
        # Function-name assignment is represented as a return-value assignment, not a plain local.
        return {"id": sid, "kind": "assignment", "target": {"kind": "identifier", "name": target},
                "value": _expr(expression), "provenance": prov,
                "semantics": {"assignment_kind": meta.get("assignment_kind", "unknown"), "raw_source": raw}}
    if typ == "CALL_OR_COMMAND":
        inline_if = re.match(r"^If\s+(.+?)\s+Then\s+(.+)$", raw, re.I)
        if inline_if:
            condition_source, tail = inline_if.groups()
            then_source, else_source = _split_inline_else(tail)
            branches = [{"branch_kind": "if_branch", "condition": _expr(condition_source),
                         "body": [_inline_child_statement(then_source, sid, "then")]}]
            if else_source:
                branches.append({"branch_kind": "else_branch", "condition": None,
                                 "body": [_inline_child_statement(else_source, sid, "else")]})
            return {"id": sid, "kind": "if", "condition": _expr(condition_source), "branches": branches,
                    "provenance": prov, "semantics": {"inline_vba_if": True, "raw_source": raw}}
        if re.match(r"^Next\b", raw, re.I):
            return {"id": sid, "kind": "unsupported", "provenance": prov,
                    "semantics": {"reason": "Loop terminator is detached from its FOR_BLOCK", "raw_source": raw}}
        if re.match(r"^Else\s*$", raw, re.I):
            return {"id": sid, "kind": "unsupported", "provenance": prov,
                    "semantics": {"reason": "ELSE token was not attached to its IF block by the frontend", "raw_source": raw}}
        label_match = re.match(r"^([A-Za-z_][A-Za-z0-9_]*)\s*:$", raw)
        if label_match:
            return {"id": sid, "kind": "label", "label": label_match.group(1), "provenance": prov,
                    "semantics": {"source_label": True}}
        expression = (meta.get("expression") or {}).get("source")
        return {"id": sid, "kind": "call", "expression": _expr(expression), "provenance": prov,
                "semantics": {"raw_source": raw, "call_resolution": "unresolved"}}
    if typ == "DECLARATION":
        redim = re.match(r"ReDim\s+(Preserve\s+)?(.+)$", raw, re.I)
        if redim:
            return {"id": sid, "kind": "redim", "target": {"kind": "identifier", "name": redim.group(2).split("(", 1)[0].strip()},
                    "provenance": prov, "semantics": {"raw_source": raw, "preserve": bool(redim.group(1)),
                    "bounds_source": redim.group(2)[redim.group(2).find("("):] if "(" in redim.group(2) else None}}
        erased = re.match(r"Erase\s+(.+)$", raw, re.I)
        if erased:
            return {"id": sid, "kind": "erase", "target": {"kind": "identifier", "name": erased.group(1).strip()},
                    "provenance": prov, "semantics": {"raw_source": raw}}
        nested_declaration = _decl_from_line(raw, node, seg_ids, source_path)
        nested_declaration["id"] = sid + ":decl"
        return {"id": sid, "kind": "variable_declaration", "declarations": [nested_declaration],
                "provenance": prov, "semantics": {"raw_source": raw}}
    if typ == "OPAQUE_MODULE_LINE":
        return {"id": sid, "kind": "unsupported", "provenance": prov,
                "semantics": {"reason": "Opaque module-level line", "raw_source": raw}}
    if typ in {"PROCEDURE", "PROJECT", "MODULE", "MODULE_DECLARATION", "IF_BRANCH", "ELSE_BRANCH", "ELSEIF_BRANCH"}:
        return None
    return {"id": sid, "kind": "unsupported", "provenance": prov,
            "semantics": {"reason": f"No IR mapping for graph node type {typ}", "raw_source": raw}}


def _extract_pseudocode_sections(text: str | None) -> tuple[str | None, dict[str, dict[str, Any]]]:
    if not text:
        return None, {}
    pattern = re.compile(r"^###\s+Procedure:\s*(.+?)(?:\s+\(part\s+\d+\s*/\s*\d+\))?\s*$", re.M)
    matches = list(pattern.finditer(text))
    if not matches:
        return text.strip(), {}
    overview = text[:matches[0].start()].strip()
    sections: dict[str, dict[str, Any]] = {}
    for index, match in enumerate(matches):
        heading = match.group(0).strip()
        name = match.group(1).strip()
        end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        section_text = text[match.start():end].strip()
        has_pseudocode = bool(re.search(r"^####\s+Pseudocode\b|^\*\*Pseudocode\*\*", section_text, re.M | re.I))
        sections[name.casefold()] = {"heading": heading, "text": section_text,
                                     "status": "pseudocode_available" if has_pseudocode else "summary_only",
                                     "contains_pseudocode_block": has_pseudocode,
                                     "authority": "generated_semantic_assistance_not_ground_truth"}
    return overview, sections


def build_ir(graph_path: str | Path, segments_path: str | Path | None = None,
             pseudocode_path: str | Path | None = None) -> dict[str, Any]:
    graph = json.loads(Path(graph_path).read_text(encoding="utf-8"))
    segments = json.loads(Path(segments_path).read_text(encoding="utf-8")) if segments_path else []
    if pseudocode_path and not Path(pseudocode_path).exists():
        raise FileNotFoundError(f"Pseudocode file not found: {pseudocode_path}")
    pseudocode = Path(pseudocode_path).read_text(encoding="utf-8") if pseudocode_path else None
    pseudocode_overview, pseudocode_sections = _extract_pseudocode_sections(pseudocode)
    nodes = graph.get("nodes", [])
    by_id = {n["id"]: dict(n) for n in nodes}
    # Repair one deterministic frontend nesting failure: children after a FOR's Next
    # terminator belong to the FOR's enclosing scope, not inside the completed loop.
    frontend_repairs: list[dict[str, Any]] = []
    for loop in [n for n in by_id.values() if n.get("type") == "FOR_BLOCK"]:
        loop_children = [n for n in by_id.values() if n.get("parent_id") == loop["id"]]
        next_nodes = [n for n in loop_children if n.get("type") == "CALL_OR_COMMAND"
                      and re.match(r"^Next\b", ((n.get("source") or {}).get("text") or "").strip(), re.I)]
        if not next_nodes:
            continue
        terminator = min(next_nodes, key=lambda n: (n.get("source") or {}).get("start_line", 10**9))
        next_line = (terminator.get("source") or {}).get("start_line", 0)
        old_parent = loop.get("parent_id")
        recovered_else_id = None
        for candidate in loop_children:
            line = (candidate.get("source") or {}).get("start_line", 0)
            if line <= next_line or candidate["id"] == terminator["id"]:
                continue
            original_parent = candidate.get("parent_id")
            raw = ((candidate.get("source") or {}).get("text") or "").strip()
            new_parent = recovered_else_id or old_parent
            reason = "Node appeared after the FOR loop's explicit Next terminator."
            if re.match(r"^Else\s*$", raw, re.I) and old_parent in by_id and by_id[old_parent].get("type") in {"IF_BRANCH", "ELSEIF_BRANCH"}:
                new_parent = by_id[old_parent].get("parent_id")
                candidate["type"] = "ELSE_BRANCH"
                candidate["name"] = "ELSE"
                candidate["metadata"] = {"recovered_from_misnested_else": True}
                recovered_else_id = candidate["id"]
                reason = "Recovered ELSE branch marker found after a completed FOR loop."
            candidate["parent_id"] = new_parent
            frontend_repairs.append({"node_id": candidate["id"], "source_line": line,
                                     "original_parent_id": original_parent, "repaired_parent_id": new_parent,
                                     "reason": reason})
    branch_priority = {"TRUE_BRANCH": 10, "ELSEIF_BRANCH": 20, "FALSE_BRANCH": 30}
    branch_order: dict[str, int] = {}
    for edge in graph.get("edges", []):
        if edge.get("relation") in branch_priority:
            branch_order[edge.get("target")] = branch_priority[edge["relation"]]
    for n in by_id.values():
        if n.get("type") in {"IF_BRANCH", "ELSEIF_BRANCH", "ELSE_BRANCH"}:
            n["_branch_order"] = branch_order.get(n["id"], 30 if n.get("type") == "ELSE_BRANCH" else 20)
    children: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for n in by_id.values():
        pid = n.get("parent_id")
        if pid in by_id:
            children[pid].append(n)
    # Source order is deterministic. The frontend's NEXT edges are retained separately.
    for pid, items in children.items():
        items.sort(key=lambda n: ((n.get("source") or {}).get("start_line", 10**9), n["id"]))
    segs_by_parent: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for seg in segments:
        if seg.get("kind") != "PROCEDURE":  # deliberately exclude full-procedure snapshots
            segs_by_parent[seg.get("parent_id", "")].append(seg)
    for nid, n in by_id.items():
        n["_children"] = children.get(nid, [])
        n["_segment_ids"] = [s["id"] for s in segs_by_parent.get(nid, [])]
        n["_segment_sources"] = [{"id": s.get("id"), "kind": s.get("kind"), "source": s.get("source", ""),
                                  "start_line": s.get("start_line"), "end_line": s.get("end_line")}
                                 for s in segs_by_parent.get(nid, [])]

    project_node = next((n for n in nodes if n.get("type") == "PROJECT"), {})
    module_nodes = [n for n in nodes if n.get("type") == "MODULE"]
    source_path = next((s.get("path") for s in graph.get("manifest", []) if s.get("path")), None)
    all_proc_nodes = [n for n in nodes if n.get("type") == "PROCEDURE"]
    program: dict[str, Any] = {
        "ir_version": IR_VERSION, "source_language": "VBA", "source_language_version": "unknown",
        "name": (graph.get("project") or {}).get("name") or project_node.get("name", "VBAProject"),
        "modules": [], "dependencies": [], "diagnostics": [], "metadata": {
            "frontend_schema_version": graph.get("schema_version"), "source_path": source_path,
            "source_manifest": graph.get("manifest", []), "pseudocode_provided": bool(pseudocode),
            "pseudocode_embedded": bool(pseudocode_sections),
            "pseudocode_artifact": ({"path": str(Path(pseudocode_path)),
                                     "sha256": hashlib.sha256(Path(pseudocode_path).read_bytes()).hexdigest(),
                                     "character_count": len(pseudocode),
                                     "procedure_sections_found": len(pseudocode_sections),
                                     "procedure_sections_missing": [p.get("name") for p in all_proc_nodes
                                                                     if p.get("name", "").casefold() not in pseudocode_sections]}
                                    if pseudocode is not None else None),
            "frontend_symbols": graph.get("symbols", []),
            "frontend_edges": graph.get("edges", []),
            "frontend_repairs": frontend_repairs,
            "construction_policy": "graph metadata + node-level fragments + non-PROCEDURE granular segments; full MODULE/PROCEDURE snapshots excluded"
        }, "semantics": {"migration_readiness": "partial", "target_language_independent": True}
    }

    proc_names: dict[str, str] = {}
    for p in all_proc_nodes:
        proc_names[p.get("name", "").lower()] = p["id"]

    # Keep graph edges as explicit relationship evidence; backends can resolve them later.
    for edge in graph.get("edges", []):
        if edge.get("relation") in {"CALLS", "REFERENCES"}:
            program["dependencies"].append({"source_id": edge.get("source"), "target_id": edge.get("target"),
                                            "relation": edge.get("relation"), "evidence": edge.get("evidence", {})})

    for mn in module_nodes:
        module_children = children.get(mn["id"], [])
        declarations = []
        procedures = []
        annotations = {"vba_options": (mn.get("metadata") or {}).get("options", []),
                       "vba_attributes": (mn.get("metadata") or {}).get("attributes", [])}
        if pseudocode_overview:
            annotations["pseudocode_module_overview"] = {
                "text": pseudocode_overview, "authority": "generated_semantic_assistance_not_ground_truth",
                "source_path": str(Path(pseudocode_path)) if pseudocode_path else None
            }
        open_type_decl = None
        for child in module_children:
            child_type = child.get("type")
            raw_module_line = _source(child)
            if child_type == "MODULE_DECLARATION":
                decl = _decl_from_line(raw_module_line, child, child.get("_segment_ids", []), source_path)
                declarations.append(decl)
                open_type_decl = decl if decl.get("kind") in {"enum", "record"} else None
            elif child_type == "OPAQUE_MODULE_LINE":
                if open_type_decl is not None and re.match(r"End\s+(Enum|Type)\b", raw_module_line, re.I):
                    open_type_decl.setdefault("semantics", {})["definition_completeness"] = "closed_with_members"
                    open_type_decl["semantics"].pop("warning", None)
                    open_type_decl["closing_provenance"] = _provenance(child, child.get("_segment_ids", []), source_path,
                                                                          child.get("_segment_sources"))
                    open_type_decl = None
                elif open_type_decl is not None:
                    enum_member = re.match(r"(\w+)\s*=\s*(.+)$", raw_module_line, re.I)
                    field_member = re.match(r"(\w+)\s*(\([^)]*\))?\s+As\s+([\w.]+)", raw_module_line, re.I)
                    if open_type_decl.get("kind") == "enum" and enum_member:
                        member = {"id": child["id"], "name": enum_member.group(1),
                                  "value": _simple_literal_expr(enum_member.group(2)),
                                  "provenance": _provenance(child, child.get("_segment_ids", []), source_path,
                                                             child.get("_segment_sources"))}
                        open_type_decl.setdefault("members", []).append(member)
                    elif open_type_decl.get("kind") == "record" and field_member:
                        field_type = field_member.group(3) + (field_member.group(2) or "")
                        member = {"id": child["id"], "name": field_member.group(1), "type": _type_spec(field_type),
                                  "provenance": _provenance(child, child.get("_segment_ids", []), source_path,
                                                             child.get("_segment_sources"))}
                        open_type_decl.setdefault("members", []).append(member)
                    else:
                        open_type_decl.setdefault("semantics", {}).setdefault("unparsed_member_lines", []).append(raw_module_line)
                else:
                    declarations.append({"id": child["id"], "kind": "unsupported", "name": child.get("name", "module_line"),
                                         "semantics": {"reason": "Opaque module-level line outside a recognized Enum/Type declaration", "raw_source": raw_module_line},
                                         "provenance": _provenance(child, child.get("_segment_ids", []), source_path,
                                                                    child.get("_segment_sources"))})
            elif child_type == "PROCEDURE":
                open_type_decl = None
                pm = child.get("metadata") or {}
                params = []
                for i, p in enumerate(pm.get("parameters", [])):
                    passing_token = str(p.get("passing", "")).lower()
                    passing = "value" if passing_token == "byval" else "reference" if passing_token == "byref" else "reference"
                    parameter = {"id": f"{child['id']}:param:{i}", "name": p.get("name", f"arg{i}"),
                                 "type": _type_spec(p.get("type")), "passing": passing,
                                 "optional": bool(p.get("optional")), "param_array": bool(p.get("param_array")),
                                 "provenance": {"source_language": "VBA", "source_text": p.get("source", ""),
                                                "source_location": _location(child, source_path), "graph_node_ids": [child["id"]],
                                                "inferred": passing_token not in {"byval", "byref"},
                                                "inference_reason": "VBA parameters default to ByRef when no passing modifier is specified."}}
                    if p.get("default") is not None:
                        parameter["default"] = _simple_literal_expr(p.get("default"))
                    params.append(parameter)
                local_vars = []
                for dn in _descendants(child, by_id, children):
                    if dn.get("type") == "DECLARATION":
                        raw = _source(dn)
                        if re.match(r"\s*(Dim|Static|Private|Public)\s+", raw, re.I) and not raw.lstrip().lower().startswith(("redim", "erase")):
                            d = _decl_from_line(raw, dn, dn.get("_segment_ids", []), source_path)
                            if d.get("kind") == "variable":
                                d["id"] = d["id"] + ":local"
                                d["scope"] = "local"
                                local_vars.append(d)
                body = []
                for ch in children.get(child["id"], []):
                    st = _statement(ch, ch.get("_children", []), ch.get("_segment_ids", []), source_path)
                    if st is not None: body.append(st)
                source_proc_kind = str(pm.get("kind", "")).upper()
                proc_kind = {"FUNCTION": "function", "SUB": "procedure", "PROPERTY_GET": "property_get",
                             "PROPERTY_LET": "property_let", "PROPERTY_SET": "property_set"}.get(source_proc_kind, "procedure")
                proc_annotations = {"source_kind": pm.get("kind"),
                                    "graph_metadata": {k:v for k,v in pm.items() if k not in {"parameters"}}}
                pseudo_section = pseudocode_sections.get(str(child.get("name", "")).casefold())
                if pseudo_section:
                    proc_annotations["pseudocode_context"] = {
                        **pseudo_section,
                        "source_path": str(Path(pseudocode_path)) if pseudocode_path else None,
                        "review_required": True
                    }
                else:
                    proc_annotations["pseudocode_context"] = {
                        "status": "missing", "source_path": str(Path(pseudocode_path)) if pseudocode_path else None,
                        "authority": "generated_semantic_assistance_not_ground_truth",
                        "review_required": True
                    }
                proc = {"id": child["id"], "name": child.get("name"),
                        "kind": proc_kind,
                        "visibility": str(pm.get("visibility", "default")).lower(), "parameters": params,
                        "return_type": _type_spec(pm.get("return_type")) if pm.get("return_type") else {"kind": "void", "name": "void"},
                        "local_variables": local_vars, "body": body,
                        "annotations": proc_annotations,
                        "provenance": _provenance(child, child.get("_segment_ids", []), source_path,
                                                   child.get("_segment_sources"), include_node_source=False),
                        "semantics": {"body_completeness": "graph_represented", "expression_completeness": "opaque_source_backed",
                                      "control_flow_integrity": "pending_diagnostic_assessment",
                                      "full_source_snapshot_used": False}}
                procedures.append(proc)
        option_sources = [o.get("source", "") for o in annotations["vba_options"]]
        base_values = [re.search(r"Option\s+Base\s+(\d+)", value, re.I) for value in option_sources]
        base_match = next((match for match in base_values if match), None)
        compare_values = [re.search(r"Option\s+Compare\s+(Text|Binary)", value, re.I) for value in option_sources]
        compare_match = next((match for match in compare_values if match), None)
        vba_module_kind = (mn.get("metadata") or {}).get("kind", "STANDARD_MODULE")
        vb_name = next((re.search(r'Attribute\s+VB_Name\s*=\s*"([^"]+)"', a.get("source", ""), re.I).group(1)
                        for a in annotations["vba_attributes"]
                        if re.search(r'Attribute\s+VB_Name\s*=\s*"([^"]+)"', a.get("source", ""), re.I)), None)
        module_semantics = {"vba_module_kind": vba_module_kind, "frontend_module_name": mn.get("name"),
                            "default_array_lower_bound": int(base_match.group(1)) if base_match else 0,
                            "string_comparison": compare_match.group(1).lower() if compare_match else "binary",
                            "full_module_snapshot_used": False}
        module_ir_kind = "interface" if "INTERFACE" in str(vba_module_kind).upper() else "class" if "STANDARD" not in str(vba_module_kind).upper() else "module"
        module = {"id": mn["id"], "name": vb_name or mn.get("name", "Module"), "kind": module_ir_kind,
                  "declarations": declarations, "procedures": procedures, "imports": [], "dependencies": [],
                  "annotations": annotations,
                  "provenance": {"source_language": "VBA", "source_location": _location(mn, source_path),
                                 "graph_node_ids": [mn["id"]], "inferred": False},
                  "semantics": module_semantics}
        program["modules"].append(module)

    # Preserve frontend diagnostics exactly; graph parser diagnostics are migration blockers until resolved.
    for d in graph.get("diagnostics", []):
        diagnostic_location = d.get("location") or {"file": d.get("path"), "line": d.get("line")}
        diagnostic_ir = {"code": d.get("code", "FRONTEND_DIAGNOSTIC"),
                         "severity": d.get("severity", "warning"), "message": d.get("message", str(d)),
                         "location": diagnostic_location,
                         "suggestion": "Resolve or explicitly disposition this diagnostic before claiming behavioural equivalence."}
        if d.get("node_id"):
            diagnostic_ir["node_id"] = d["node_id"]
        program["diagnostics"].append(diagnostic_ir)
    # Attach structural parser diagnostics to affected procedures so backends cannot overlook them.
    for module in program["modules"]:
        for proc in module["procedures"]:
            start = (proc.get("provenance", {}).get("source_location") or {}).get("line", 0)
            end = (proc.get("provenance", {}).get("source_location") or {}).get("end_line", start)
            affected = [d for d in program["diagnostics"] if d.get("code") == "CROSS_NESTED_BLOCK_CLOSE"
                        and start <= (d.get("location") or {}).get("line", -1) <= end]
            repairs = [r for r in frontend_repairs if start <= r.get("source_line", -1) <= end]
            if affected:
                proc["semantics"]["control_flow_integrity"] = "uncertain_due_to_frontend_diagnostic"
                proc["semantics"]["blocking_diagnostics"] = [d["code"] + "@" + str((d.get("location") or {}).get("line")) for d in affected]
            elif repairs:
                proc["semantics"]["control_flow_integrity"] = "repaired_needs_review"
            else:
                proc["semantics"]["control_flow_integrity"] = "no_frontend_block_diagnostic_reported"
            if repairs:
                proc["semantics"]["frontend_repairs"] = [r["node_id"] for r in repairs]
    if any(d.get("severity") == "warning" and d.get("code") == "CROSS_NESTED_BLOCK_CLOSE" for d in program["diagnostics"]):
        program["semantics"]["migration_readiness"] = "blocked_pending_control_flow_review"
    # Explicitly flag constructs for which the frontend currently provides only raw text/metadata.
    opaque_expr_count = sum(1 for m in program["modules"] for p in m["procedures"] for s in _walk_statements(p["body"])
                            for k in ("expression", "value", "condition", "start", "end", "step")
                            if isinstance(s.get(k), dict) and s[k].get("kind") == "unsupported")
    program["diagnostics"].append({"code": "IR_EXPRESSIONS_OPAQUE", "severity": "warning",
                                   "message": f"{opaque_expr_count} expression/control operands are preserved as source-backed opaque expressions, not parsed AST nodes.",
                                   "suggestion": "Add a real expression parser or validated LLM-assisted expression normalization before code generation."})
    if frontend_repairs:
        program["diagnostics"].append({"code": "IR_FRONTEND_HIERARCHY_REPAIRED", "severity": "warning",
                                       "message": f"Applied {len(frontend_repairs)} deterministic hierarchy repairs using explicit FOR/Next boundaries.",
                                       "suggestion": "Review every entry in metadata.frontend_repairs against the source/pseudocode before code generation."})
    if opaque_expr_count or frontend_repairs or any(d.get("code") == "CROSS_NESTED_BLOCK_CLOSE" for d in program["diagnostics"]):
        program["semantics"]["migration_readiness"] = "blocked_pending_semantic_and_control_flow_review"
    incomplete_types = [d["name"] for m in program["modules"] for d in m["declarations"]
                        if d.get("kind") in {"enum", "record"} and d.get("semantics", {}).get("definition_completeness") != "closed_with_members"]
    if incomplete_types:
        program["diagnostics"].append({"code": "IR_DECLARATION_MEMBERS_INCOMPLETE", "severity": "warning",
                                       "message": "Enum/record declarations remain incomplete: " + ", ".join(incomplete_types),
                                       "suggestion": "Extend the frontend declaration model or supply validated declaration fragments."})
    return program


def _descendants(node: dict[str, Any], by_id: dict[str, dict[str, Any]], children: dict[str, list[dict[str, Any]]]) -> list[dict[str, Any]]:
    out = []
    stack = list(children.get(node["id"], []))
    while stack:
        current = stack.pop(0)
        out.append(current)
        stack[0:0] = children.get(current["id"], [])
    return out


def _walk_statements(statements: list[dict[str, Any]]):
    for stmt in statements:
        yield stmt
        for key in ("body", "else_body", "statements"):
            child = stmt.get(key)
            if isinstance(child, list):
                yield from _walk_statements(child)
        for branch in stmt.get("branches", []):
            yield from _walk_statements(branch.get("body", []))
