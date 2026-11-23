"""Generate Java from original VBA plus verified process notes."""

import hashlib
import importlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from ..process_notes import notes_generator


_CLIENT = None
_MODEL = None


def _get_client():
    global _CLIENT, _MODEL
    if _CLIENT is not None:
        return _CLIENT, _MODEL
    project_root = Path(__file__).resolve().parents[2]
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))
    try:
        api_keys = importlib.import_module("api_keys")
        from langchain_groq import ChatGroq

        _MODEL = os.getenv("LLM_MODEL") or api_keys._loaded["LLM"]
        # GPT-OSS 20B permits up to 65,536 completion tokens; leave headroom below that cap.
        client = ChatGroq(model=_MODEL, temperature=0, max_tokens=60000,
                          reasoning_format="hidden", reasoning_effort="low")
        _CLIENT = client
        return _CLIENT, str(_MODEL)
    except Exception as exc:
        raise RuntimeError("Java LLM unavailable; check GROQ_API, LLM, and model configuration") from exc


def _response_source(response):
    if hasattr(response, "content"):
        response = response.content
    if isinstance(response, list):
        response = "\n".join(
            item.get("text", "") if isinstance(item, dict) else str(item)
            for item in response
        )
    if not isinstance(response, str) or not response.strip():
        raise ValueError("LLM response did not contain Java source text")

    source = response.strip()
    fenced = re.search(r"```(?:java)?\s*([\s\S]*?)```", source, flags=re.IGNORECASE)
    if fenced:
        source = fenced.group(1).strip()
    else:
        start = re.search(r"(?m)^\s*(?:package\s+|import\s+|public\s+class\s+)", source)
        if start:
            source = source[start.start():].strip()
    return source


def _source_calculations(source):
    rows = notes_generator._logical_statements(source)
    procedures = notes_generator._procedure_sections(rows)
    return procedures, notes_generator.source_calculations(source, rows, procedures)


def _note_gaps(source_path, note_path, source, procedures, calculations):
    if not note_path.exists():
        return ["process_notes.md is missing"]
    note = note_path.read_text(encoding="utf-8")
    digest = hashlib.sha256(source_path.read_bytes()).hexdigest()
    if "- Notes format version: `2`" not in note:
        return ["process note uses an older format and needs regeneration"]
    if f"- SHA-256: `{digest}`" not in note:
        return ["process note hash does not match this .bas file"]
    if "**LLM status:** Complete —" not in note:
        return ["process note does not contain a complete LLM analysis"]
    gaps = []
    for procedure in procedures:
        if f"## {procedure['kind'].title()}: `{procedure['name']}`" not in note:
            gaps.append(f"procedure {procedure['name']} is missing from process notes")
        if f"### LLM interpretation: `{procedure['name']}`" not in note:
            gaps.append(f"LLM explanation for procedure {procedure['name']} is missing")
    for calculation in calculations:
        ledger = (f"- Line {calculation['source_line']}: `{notes_generator._escape(calculation['target'])}` = "
                  f"`{notes_generator._escape(calculation['expression'])}`;")
        if ledger not in note:
            gaps.append(f"calculation on line {calculation['source_line']} is missing from the deterministic ledger")
        belongs_to_procedure = any(
            procedure["start"] <= calculation["source_line"] <= procedure["end"]
            for procedure in procedures
        )
        llm_finding = f"exact VBA: `{notes_generator._escape(calculation['expression'])}`"
        if belongs_to_procedure and llm_finding not in note:
            gaps.append(f"LLM explanation is missing calculation on line {calculation['source_line']}")
    return gaps


def _java_client_prompt(source, process_note, class_name, procedures, calculations,
                        prior_source="", repair_errors=None):
    request = {
        "task": "Translate the supplied VBA module into one Java source file.",
        "requirements": [
            "Use both the original .bas and its process_notes.md. The .bas is the source of truth; notes are an aid.",
            "Do not omit procedures, branches, loops, labels, error paths, declarations, or side effects. Do not invent behavior.",
            "Preserve every listed VBA calculation exactly in meaning and operand order. The calculation ledger is reference material; do not return a repeated mapping table.",
            "If exact behavior cannot be represented safely, mark it with a // UNRESOLVED: comment; never hide it as an assumption.",
            "Represent Excel/VBA host objects through explicit Java adapter interfaces passed into the translated procedure. Do not fabricate workbook data or no-op required Excel operations. A host adapter that must be supplied at integration time is a documented dependency, not an unresolved translation, when the Java interface preserves every operation and its arguments/results.",
            "Keep each translated Java method name exactly equal to its VBA procedure name, including capitalization.",
            "Return one compilable .java file with exactly one public top-level type: the class named class_name. Put every adapter interface, helper class, and enum inside that class as a nested type; do not declare any other public top-level type.",
            "Keep Java concise: include only code needed to preserve the VBA operations, without long explanatory comments or duplicated documentation.",
            "Use Java standard library only unless the source clearly requires a dependency.",
            "Java's ! operator can negate only a boolean. For StrComp(left, right, vbTextCompare) <> 0, use !(left).equalsIgnoreCase(right); parenthesize concatenated string operands, such as !(left + next).equalsIgnoreCase(right).",
            "Keep calls consistent with the adapter methods actually declared in the Java source; add or correct adapter declarations rather than calling nonexistent methods.",
            "Initialize Java local variables to VBA's declared-type defaults when needed, including on error-handler paths: String to empty string, numbers to zero, Boolean to false, and object references to null.",
            "Preserve ByRef mutations with a compatible mutable representation at the parameter, every call site, and later consumers. Do not pass String where StringBuilder is expected; keep all declarations and uses type-consistent.",
            "The Java class must compile with the system javac. Do not claim runtime equivalence; report unresolved semantic differences.",
            "Return only the complete Java source text. Do not return JSON, markdown fences, or an explanation.",
            "If behavior truly cannot be represented, include a source comment beginning // UNRESOLVED: . If you must make an assumption, include // ASSUMPTION: . Do not silently omit uncertain behavior.",
        ],
        "class_name": class_name,
        "required_procedure_names": [procedure["name"] for procedure in procedures],
        "source_calculations_to_preserve": calculations,
        "original_vba": source,
        "process_notes_markdown": process_note,
    }
    if prior_source:
        request["previous_java_source_to_repair"] = prior_source
        request["repair_findings"] = repair_errors or []
        repair_instruction = (
            "Repair every listed issue. Keep all previously correct behavior and formulas, and "
            "return the complete corrected Java file. If the only unresolved item is an Excel/VBA "
            "host dependency, preserve its operations behind explicit Java adapter interfaces and "
            "document that integration requirement in the source; do not call it unresolved just "
            "because a host implementation must be provided. Do not use fake data or no-op methods. "
            "Keep translated procedure names exactly as spelled in VBA. Return only complete Java source text with no JSON or markdown fences."
        )
        repair_error_text = "\n".join(repair_errors or []).casefold()
        if "might not have been initialized" in repair_error_text:
            repair_instruction += (
                " Initialize the local reported by javac using the default value for its VBA type, "
                "so error-handler paths preserve VBA behavior."
            )
        if "cannot be converted to" in repair_error_text:
            repair_instruction += (
                " Resolve the reported type mismatch consistently across the caller, callee, and "
                "all uses. For ByRef strings, use one mutable representation and convert only where "
                "a String value is required."
            )
        request["repair_instruction"] = repair_instruction
    return json.dumps(request, ensure_ascii=False)


def _audit(result, class_name, procedures, calculations):
    java_source = _response_source(result)
    errors = []
    if not re.search(rf"\bclass\s+{re.escape(class_name)}\b", java_source):
        errors.append("Java source does not declare the required class")
    for procedure in procedures:
        # Match Java method declarations by name. Calls alone do not satisfy this check.
        declaration = rf"\b{re.escape(procedure['name'])}\s*\([^;{{}}]*\)\s*(?:throws\s+[\w.,\s]+)?\{{"
        if not re.search(declaration, java_source, flags=re.IGNORECASE):
            errors.append(f"Java procedure declaration is missing: {procedure['name']}")
    for marker in ("UNRESOLVED", "ASSUMPTION"):
        details = re.findall(rf"(?im)^\s*//\s*{marker}\s*:\s*(.+)$", java_source)
        errors.extend(f"{marker.lower()} semantics: {detail.strip()}" for detail in details)
    return java_source, errors


def _compile(java_source, class_name):
    javac = shutil.which("javac")
    if not javac:
        return "javac was not found on PATH"
    with tempfile.TemporaryDirectory(prefix="vba-java-check-") as temp:
        source = Path(temp) / f"{class_name}.java"
        source.write_text(java_source, encoding="utf-8")
        run = subprocess.run([javac, "-encoding", "UTF-8", "-d", temp, str(source)],
                             capture_output=True, text=True, timeout=60)
        if run.returncode:
            return (run.stderr or run.stdout).strip()[-12000:]
    return None


def _normalize_repair_error(error):
    error = re.sub(r"(?m)^.*?([^/\\]+\.java):\d+:", r"\1:<line>:", error)
    return error


def _repair_negated_string_concatenation(java_source, compile_error):
    diagnostic = re.search(
        r"\.java:(\d+): error: bad operand type String for unary operator '!'",
        compile_error,
    )
    if not diagnostic:
        return java_source
    lines = java_source.splitlines()
    line_index = int(diagnostic.group(1)) - 1
    if not 0 <= line_index < len(lines):
        return java_source
    pattern = re.compile(
        r"!\s*(?P<value>[A-Za-z_$][\w$]*(?:\s*\+\s*[A-Za-z_$][\w$]*)+)"
        r"\s*\.\s*equalsIgnoreCase\s*\(",
        re.DOTALL,
    )
    chunk_end = min(line_index + 4, len(lines))
    chunk = "\n".join(lines[line_index:chunk_end])
    repaired_line, count = pattern.subn(
        lambda match: f"!({match.group('value')}).equalsIgnoreCase(",
        chunk,
        count=1,
    )
    if not count:
        return java_source
    lines[line_index:chunk_end] = repaired_line.splitlines()
    return "\n".join(lines) + ("\n" if java_source.endswith("\n") else "")


def _write_report(path, heading, bullets):
    path.parent.mkdir(parents=True, exist_ok=True)
    body = [f"# Java Conversion Report: {heading}", "", *[f"- {item}" for item in bullets], ""]
    path.write_text("\n".join(body), encoding="utf-8")


def main(input_location, process_note_location, output_location):
    """Verify process notes, generate Java, repair compile/static failures, and report limits."""
    source_path, note_path, output = Path(input_location), Path(process_note_location), Path(output_location)
    report_path = output / "conversion_report.md"
    source, _encoding = notes_generator._decode(source_path.read_bytes())
    procedures, calculations = _source_calculations(source)

    for attempt in range(2):
        gaps = _note_gaps(source_path, note_path, source, procedures, calculations)
        if not gaps:
            break
        print(f"[NOTES] GAPS — {len(gaps)}; regenerating process note ({attempt + 1}/2)", flush=True)
        if attempt == 1 or not notes_generator.main(source_path, note_path.parent):
            _write_report(report_path, source_path.name, [*gaps, "Conversion stopped because process notes are incomplete."])
            print("[STOP] Process note gaps remain; Java generation was not started.", flush=True)
            return False
        note_path = note_path.parent / "process_notes.md"
    process_note = note_path.read_text(encoding="utf-8")
    note_gaps = _note_gaps(source_path, note_path, source, procedures, calculations)
    if note_gaps:
        _write_report(report_path, source_path.name, [*note_gaps, "Conversion stopped because process notes are incomplete."])
        print("[STOP] Process note gaps remain; Java generation was not started.", flush=True)
        return False

    class_name = re.sub(r"\W+", "_", source_path.stem)
    if not class_name or class_name[0].isdigit():
        class_name = "Vba_" + class_name
    class_name += "Java"
    try:
        client, model = _get_client()
    except Exception as exc:
        _write_report(report_path, source_path.name, [f"LLM setup failed: {type(exc).__name__}: {exc}"])
        print(f"[JAVA] FAILED — LLM setup: {exc}", flush=True)
        return False

    previous_source, repair_errors = "", []
    max_attempts = 5
    for attempt in range(1, max_attempts + 1):
        print(f"[JAVA] GENERATION {attempt}/{max_attempts} — {source_path.name} using {model}", flush=True)
        prompt = _java_client_prompt(source, process_note, class_name, procedures, calculations,
                                     previous_source, repair_errors)
        try:
            response = client.invoke(prompt)
            java_source, errors = _audit(response, class_name, procedures, calculations)
        except Exception as exc:
            _write_report(report_path, source_path.name, [f"LLM Java generation failed: {type(exc).__name__}: {exc}"])
            print(f"[JAVA] FAILED — {type(exc).__name__}: {exc}", flush=True)
            print("[STOP] Java conversion stopped; no later file will be processed.", flush=True)
            return False
        if not errors:
            try:
                compile_error = _compile(java_source, class_name)
            except Exception as exc:
                compile_error = f"javac validation failed: {type(exc).__name__}: {exc}"
            if compile_error:
                repaired_source = _repair_negated_string_concatenation(java_source, compile_error)
                if repaired_source != java_source:
                    java_source = repaired_source
                    try:
                        compile_error = _compile(java_source, class_name)
                    except Exception as exc:
                        compile_error = f"javac validation failed: {type(exc).__name__}: {exc}"
                if compile_error:
                    errors.append("javac compilation failed: " + compile_error)
        if not errors:
            try:
                output.mkdir(parents=True, exist_ok=True)
                java_path = output / f"{class_name}.java"
                java_path.write_text(java_source, encoding="utf-8")
                _write_report(report_path, source_path.name, [
                    f"Java file: {java_path.name}", "Java compilation: PASS",
                    f"Source calculation ledger entries supplied to the LLM: {len(calculations)}",
                    "Calculation-to-Java equivalence: NOT AUTOMATICALLY VERIFIED.",
                    "Behavioral equivalence: NOT VERIFIED; no execution comparison against VBA was run.",
                    "Do not claim 100% accuracy until the Java and original VBA produce matching results across a reference test corpus.",
                ])
            except Exception as exc:
                _write_report(report_path, source_path.name, [f"Writing Java output failed: {type(exc).__name__}: {exc}"])
                print(f"[JAVA] OUTPUT FAILED — {type(exc).__name__}: {exc}", flush=True)
                return False
            print(f"[JAVA] COMPILED — {java_path}", flush=True)
            print(f"[AUDIT] Source calculation ledger supplied — {len(calculations)} entries", flush=True)
            print("[AUDIT] Behavioral equivalence — NOT VERIFIED", flush=True)
            return True
        if repair_errors and sorted(map(_normalize_repair_error, errors)) == sorted(
            map(_normalize_repair_error, repair_errors)
        ):
            repair_errors = errors
            print("[STOP] Repair made no progress; no Java file was emitted.", flush=True)
            break
        previous_source = java_source
        repair_errors = errors
        preview = "; ".join(errors[:3])
        if len(errors) > 3:
            preview += f"; and {len(errors) - 3} more"
        print(f"[JAVA] REPAIR REQUIRED — {len(errors)} issue(s): {preview}", flush=True)

    _write_report(report_path, source_path.name, [*repair_errors, "No Java file was emitted because the checks did not pass."])
    if len(repair_errors) and attempt < max_attempts:
        return False
    print(f"[STOP] Repair limit ({max_attempts}) reached; no Java file was emitted.", flush=True)
    return False
