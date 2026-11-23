"""Create readable, source-linked process notes for VBA inputs."""

import hashlib
import re
from pathlib import Path

from .llm_analysis import analyze_source


PROCEDURE = re.compile(
    r"^(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function|Property\s+(?:Get|Let|Set))\s+"
    r"([A-Za-z_]\w*)\b(.*)$", re.IGNORECASE
)
END_PROCEDURE = re.compile(r"^End\s+(Sub|Function|Property)\s*$", re.IGNORECASE)
NOTES_VERSION = "2"


def _decode(data):
    if data.startswith(b"\xef\xbb\xbf"):
        return data.decode("utf-8-sig"), "utf-8-sig"
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        return data.decode("utf-16"), "utf-16"
    for encoding in ("utf-8", "cp1252"):
        try:
            return data.decode(encoding), encoding
        except UnicodeDecodeError:
            pass
    return data.decode("latin-1"), "latin-1"


def _logical_statements(source):
    """Join explicit VBA line continuations, keeping original text and line range."""
    rows = []
    pending = []
    start = 1
    for line_number, line in enumerate(source.splitlines(), 1):
        if not pending:
            start = line_number
        pending.append(line)
        code, _ = _strip_comment(line)
        trimmed = code.rstrip()
        if trimmed.endswith("_") and len(trimmed) > 1 and trimmed[-2].isspace():
            continue
        rows.append({"start": start, "end": line_number, "raw": "\n".join(pending),
                     "text": " ".join(part.strip() for part in pending).strip()})
        pending = []
    if pending:
        rows.append({"start": start, "end": len(source.splitlines()), "raw": "\n".join(pending),
                     "text": " ".join(part.strip() for part in pending).strip()})
    return rows


def _strip_comment(line):
    quoted = False
    i = 0
    while i < len(line):
        if line[i] == '"':
            if quoted and i + 1 < len(line) and line[i + 1] == '"':
                i += 2
                continue
            quoted = not quoted
        elif line[i] == "'" and not quoted:
            return line[:i], line[i:]
        i += 1
    return line, ""


def _assignment(text):
    """Find an assignment operator outside strings and parenthesized expressions."""
    quoted = False
    depth = 0
    i = 0
    while i < len(text):
        char = text[i]
        if char == '"':
            if quoted and i + 1 < len(text) and text[i + 1] == '"':
                i += 2
                continue
            quoted = not quoted
        elif not quoted:
            if char == "(":
                depth += 1
            elif char == ")":
                depth = max(0, depth - 1)
            elif char == "=" and depth == 0:
                before = text[i - 1] if i else ""
                after = text[i + 1] if i + 1 < len(text) else ""
                if before not in "<>:" and after != "=":
                    return text[:i].strip(), text[i + 1:].strip()
        i += 1
    return None


def _kind(text):
    clean, comment = _strip_comment(text)
    value = clean.strip()
    low = value.casefold()
    if not value:
        return "Comment" if comment or text.strip().startswith("'") else "Blank line"
    if value.startswith("#"):
        return "Compiler directive"
    if value.casefold().startswith("attribute "):
        return "Module attribute"
    if value.casefold().startswith("option "):
        return "Module option"
    if re.match(r"^(Public|Private|Friend|Static)\s+(Sub|Function|Property)\b", value, re.I):
        return "Procedure declaration"
    if re.match(r"^(End\s+(Sub|Function|Property))\s*$", value, re.I):
        return "Procedure end"
    if re.match(r"^if\b.*\bthen\s+\S", value, re.I):
        return "Decision with inline action (review)"
    if re.match(r"^(if|elseif|else|select case|case\b|end if|end select)\b", value, re.I):
        return "Decision / branch"
    if re.match(r"^(for|next|do|loop|while|wend)\b", value, re.I):
        return "Loop control"
    if re.match(r"^(on error|resume\b|err\.)", value, re.I):
        return "Error handling"
    if re.match(r"^(?:[A-Za-z_]\w*|\d+)\s*:\s*\S+", value):
        return "Label with inline statement (review)"
    if re.match(r"^[A-Za-z_]\w*\s*:$|^\d+\s*:$", value):
        return "Label"
    if re.match(r"^(goto|gosub|return\b|exit\b)\b", value, re.I):
        return "Jump / procedure flow"
    if re.match(r"^(with\b|end with\b)", value, re.I):
        return "With block"
    if re.match(r"^(dim|static|const|redim|erase|public|private|global|type\b|enum\b|end type|end enum)\b", value, re.I):
        return "Declaration / storage"
    if re.match(r"^(set|let)\s+", value, re.I) or _assignment(value):
        return "Calculation / assignment"
    if re.match(r"^(call\s+)?[A-Za-z_.][\w.!]*\s*(\(|\s|$)", value, re.I):
        return "Call / action"
    return "Other VBA statement (review)"


def _escape(value):
    return value.replace("|", "\\|").replace("\n", "<br>").replace("`", "\\`")


def _procedure_sections(rows):
    sections = []
    current = None
    for row in rows:
        text = row["text"].strip()
        match = PROCEDURE.match(text)
        if match and current is None:
            current = {"name": match.group(2), "kind": match.group(1).upper(),
                       "header": text, "start": row["start"], "rows": [row]}
            continue
        if current is not None:
            current["rows"].append(row)
            if END_PROCEDURE.match(text):
                current["end"] = row["end"]
                sections.append(current)
                current = None
        elif not match:
            continue
    if current is not None:
        current["end"] = current["rows"][-1]["end"]
        current["unclosed"] = True
        sections.append(current)
    return sections


def _split_inline_statements(text):
    """Expose actions after a label, inline If, or VBA colon separator."""
    text = _strip_comment(text)[0].strip()
    conditional = re.match(r"^If\s+(.+?)\s+Then\s+(.+)$", text, re.I)
    branches = [text]
    if conditional:
        tail = conditional.group(2).strip()
        quoted = False
        split_at = None
        i = 0
        while i < len(tail):
            if tail[i] == '"':
                if quoted and i + 1 < len(tail) and tail[i + 1] == '"':
                    i += 2
                    continue
                quoted = not quoted
            if not quoted and tail[i:i + 5].casefold() == " else":
                split_at = i
                break
            i += 1
        branches = [tail[:split_at].strip(), tail[split_at + 6:].strip()] if split_at is not None else [tail]
    label = re.match(r"^(?:[A-Za-z_]\w*|\d+)\s*:(.*)$", text)
    if label and not conditional:
        branches = [label.group(1).strip()]
    result = []
    for branch in branches:
        buffer, quoted, depth, i = [], False, 0, 0
        while i < len(branch):
            char = branch[i]
            if char == '"':
                if quoted and i + 1 < len(branch) and branch[i + 1] == '"':
                    buffer.extend(['"', '"'])
                    i += 2
                    continue
                quoted = not quoted
            elif not quoted:
                if char == "(":
                    depth += 1
                elif char == ")":
                    depth = max(0, depth - 1)
                elif char == ":" and depth == 0 and not (i + 1 < len(branch) and branch[i + 1] == "="):
                    part = "".join(buffer).strip()
                    if part:
                        result.append(part)
                    buffer = []
                    i += 1
                    continue
            buffer.append(char)
            i += 1
        last = "".join(buffer).strip()
        if last:
            result.append(last)
    return result


def source_calculations(source, rows=None, procedures=None):
    """List source assignments, including inline actions and module constants."""
    rows = rows if rows is not None else _logical_statements(source)
    procedures = procedures if procedures is not None else _procedure_sections(rows)
    in_procedure = {row["start"] for proc in procedures for row in proc["rows"]}
    in_enum = in_type = False
    records = []
    for row in rows:
        text = row["text"].strip()
        if not text:
            continue
        inside = row["start"] in in_procedure
        if not inside:
            if re.match(r"^(?:Public|Private|Friend)?\s*Enum\b", text, re.I):
                in_enum = True
                continue
            if re.match(r"^End\s+Enum\b", text, re.I):
                in_enum = False
                continue
            if re.match(r"^(?:Public|Private|Friend)?\s*Type\b", text, re.I):
                in_type = True
                continue
            if re.match(r"^End\s+Type\b", text, re.I):
                in_type = False
                continue
            if in_type or (not in_enum and not re.match(r"^(?:Public|Private|Friend)?\s*Const\b", text, re.I)):
                continue
        for statement in _split_inline_statements(text):
            if re.match(r"^(if|elseif|else|for|next|do|loop|while|wend|select case|case\b|end if|end select|with\b|end with|on error|resume\b|goto\b|gosub\b|return\b|exit\b|dim\b|redim\b|erase\b|call\b|debug\.print\b)", statement, re.I):
                continue
            assignment = _assignment(statement)
            if assignment:
                records.append({"source_line": row["start"], "statement": statement,
                                "source_statement": text, "target": assignment[0],
                                "expression": assignment[1]})
    return records


def _render_file(path):
    data = path.read_bytes()
    source, encoding = _decode(data)
    rows = _logical_statements(source)
    procedures = _procedure_sections(rows)
    lines = [f"# VBA Process Notes: `{path.name}`", "",
             "## File details", "",
             f"- Notes format version: `{NOTES_VERSION}`",
             f"- Source path: `{path}`", f"- SHA-256: `{hashlib.sha256(data).hexdigest()}`",
             f"- Decoded as: `{encoding}`", f"- Physical lines: {len(source.splitlines())}",
             f"- Procedures found: {len(procedures)}", "",
             "## What this file contains", "",
             "The sections below list source statements in their original order. Labels are mechanical descriptions based on statement shape; they do not claim to know business intent. Exact VBA text and line numbers are included so each note can be checked against the source.", ""]
    module_rows = [row for row in rows if not any(p["start"] <= row["start"] <= p["end"] for p in procedures)]
    lines.extend(["## Module-level declarations and settings", ""])
    if module_rows:
        module_scope = None
        for row in module_rows:
            text = row["text"].strip()
            if not text:
                continue
            if re.match(r"^(Public|Private|Friend)?\s*Enum\b", text, re.I):
                module_scope, kind = "Enum member", "Enum declaration"
            elif re.match(r"^(Public|Private|Friend)?\s*Type\b", text, re.I):
                module_scope, kind = "User-defined type member", "User-defined type declaration"
            elif re.match(r"^End\s+Enum\b", text, re.I):
                kind, module_scope = "End enum", None
            elif re.match(r"^End\s+Type\b", text, re.I):
                kind, module_scope = "End user-defined type", None
            else:
                kind = module_scope or _kind(text)
            lines.append(f"- Lines {row['start']}–{row['end']} — **{kind}**: `{_escape(text)}`")
    else:
        lines.append("No module-level statements found.")
    lines.append("")

    assignments = 0
    uncertain = 0
    for procedure in procedures:
        lines.extend([f"## {procedure['kind'].title()}: `{procedure['name']}` (lines {procedure['start']}–{procedure['end']})", "",
                      f"**Declaration:** `{_escape(procedure['header'])}`", "",
                      "### Ordered process notes", ""])
        for row in procedure["rows"][1:]:
            text = row["text"].strip()
            if not text:
                continue
            kind = _kind(text)
            clean, _ = _strip_comment(text)
            assignment = _assignment(clean.strip())
            if kind == "Calculation / assignment":
                assignments += 1
                if assignment:
                    explanation = f"Assign `{_escape(assignment[0])}` using the exact expression `{_escape(assignment[1])}`."
                else:
                    explanation = "Assignment-like statement; inspect the exact VBA text below."
            elif kind == "Decision / branch":
                explanation = "This line selects or changes a decision branch; the condition text is preserved below."
            elif kind == "Decision with inline action (review)":
                explanation = "This line has a decision and an inline action; both remain in the exact source statement and need review."
                uncertain += 1
            elif kind == "Loop control":
                explanation = "This line starts, controls, or ends repeated work; bounds and conditions are preserved below."
            elif kind in {"Other VBA statement (review)", "Label with inline statement (review)",
                          "Decision with inline action (review)"}:
                explanation = "Not classified by this note generator; review this source statement."
                uncertain += 1
            elif kind == "Declaration / storage":
                explanation = "Declares, sizes, or clears VBA storage; the exact declaration is preserved below."
            elif kind == "Error handling":
                explanation = "Changes VBA error handling or refers to the active error; preserve this behavior during conversion."
            elif kind == "Jump / procedure flow":
                explanation = "Changes procedure flow; preserve the target and execution order during conversion."
            elif kind == "Call / action":
                explanation = "Calls a procedure, property, or host action; identify its dependency and side effects."
            elif kind == "Comment":
                explanation = "Source comment retained as context; verify its intent before relying on it."
            else:
                explanation = kind + "."
            lines.append(f"- **Lines {row['start']}–{row['end']} · {kind}:** {explanation}")
            lines.append(f"  - VBA: `{_escape(text)}`")
        if procedure.get("unclosed"):
            lines.append("- **Review required:** no matching `End Sub`, `End Function`, or `End Property` was found.")
        lines.append("")

    calculations = source_calculations(source, rows, procedures)
    lines.extend(["## Calculation ledger", "",
                  "Every calculation-like assignment found by the deterministic source scan is listed here. The original statement and expression are kept verbatim for conversion checks.", ""])
    for item in calculations:
        lines.append(f"- Line {item['source_line']}: `{_escape(item['target'])}` = `{_escape(item['expression'])}`; source: `{_escape(item['source_statement'])}`")
    lines.append("")

    print(f"[DETERMINISTIC] COMPLETE — {len(source.splitlines())} source lines, "
          f"{len(procedures)} procedures, {assignments} procedure assignments, "
          f"{len(calculations)} calculation ledger entries, {uncertain} review items", flush=True)
    print(f"[LLM] START — analyzing {path.name}", flush=True)
    llm_status, analysis = analyze_source(
        source, path.name, procedures,
        progress=lambda message: print(message, flush=True),
        calculations=calculations,
    )
    if llm_status.startswith("Complete"):
        print(f"[LLM] COMPLETE — {llm_status}", flush=True)
    elif llm_status.startswith("Skipped"):
        print(f"[LLM] SKIPPED — {llm_status}", flush=True)
    else:
        print(f"[LLM] FAILED — {llm_status}", flush=True)
    lines.extend(["## LLM-assisted process explanation", "",
                  f"**LLM status:** {llm_status}", ""])
    if analysis:
        for item in analysis.get("procedures", []):
            lines.extend([f"### LLM interpretation: `{item['name']}`", "",
                          f"**Confidence:** {item.get('confidence', 'not rated')}", "",
                          f"{item.get('summary', 'No summary returned.')} (source lines: "
                          f"{', '.join(str(number) for number in item.get('summary_source_lines', []))})", ""])
            categories = (("input_output", "Inputs and outputs"), ("calculation", "Calculations"),
                          ("decision", "Decisions"), ("loop", "Loops"),
                          ("data_effect", "Data effects"), ("dependency", "Dependencies"),
                          ("uncertainty", "Unclear behavior / questions"))
            for category, label in categories:
                entries = [finding for finding in item.get("findings", []) if finding["category"] == category]
                if not entries:
                    continue
                lines.extend([f"#### {label}", ""])
                for entry in entries:
                    source_lines = ", ".join(str(number) for number in entry["source_lines"])
                    description = entry["description"]
                    if entry.get("exact_vba"):
                        description += f"; exact VBA: `{_escape(entry['exact_vba'])}`"
                    lines.append(f"- {description} (source lines: {source_lines})")
                lines.append("")
        for error in analysis.get("errors", []):
            lines.append(f"- **LLM procedure failure:** {_escape(error)}")
        if analysis.get("errors"):
            lines.append("")
    else:
        lines.extend(["LLM interpretation was not available for this run. The source-linked inventory above is still available; do not treat it as a verified explanation.", ""])

    lines.extend(["## Review checklist", "",
                  "- [ ] Confirm every calculation and its exact expression.",
                  "- [ ] Confirm each decision condition and both/all branch outcomes.",
                  "- [ ] Confirm loop start, end, step, and exit conditions.",
                  "- [ ] Identify data read or written, including workbook, file, database, and UI effects.",
                  "- [ ] Identify dependencies called by each procedure.",
                  "- [ ] Review all statements marked `(review)`.",
                  "- [ ] Mark each procedure verified, unclear, or not reviewed before Java conversion.", "",
                  "## Coverage counts", "",
                  f"- Assignment-like statements: {assignments}",
                  f"- Unclassified statements to review: {uncertain}",
                  f"- Procedures without a recognized end: {sum(1 for item in procedures if item.get('unclosed'))}", "",
                  "> These notes are an evidence index, not a claim that the program's business meaning has been fully understood. Verify every explanation against the cited VBA before converting it.", ""])
    return "\n".join(lines), llm_status


def main(input_location, output_location):
    """Write one readable process note for each VBA file at input_location."""
    source = Path(input_location)
    output = Path(output_location)
    output.mkdir(parents=True, exist_ok=True)
    paths = [source] if source.is_file() else sorted(
        item for item in source.rglob("*") if item.is_file() and item.suffix.lower() in {".bas", ".cls", ".frm", ".vba", ".txt"}
    )
    for path in paths:
        print(f"[PROCESS] {path}", flush=True)
        try:
            print(f"[DETERMINISTIC] START — building source inventory", flush=True)
            note, llm_status = _render_file(path)
        except Exception as exc:
            print(f"[DETERMINISTIC] FAILED — {path}: {type(exc).__name__}: {exc}", flush=True)
            return False
        destination = output / "process_notes.md"
        try:
            destination.write_text(note, encoding="utf-8")
            print(f"[OUTPUT] SAVED — {destination}", flush=True)
        except Exception as exc:
            print(f"[OUTPUT] FAILED — {destination}: {type(exc).__name__}: {exc}", flush=True)
            return False
        if llm_status.startswith(("Failed", "Unavailable")):
            print(f"[STOP] Stopping the batch after the LLM failure for {path.name}.", flush=True)
            return False
    return True
