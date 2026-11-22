"""LLM repair planning for known diagnostics and broad output-quality audits.

The model is treated as an untrusted proposer. Every change is hash-bound,
preconditioned, independently reviewed, and validated after application.
"""
from __future__ import annotations
import json
from pathlib import Path


class RepairPlanner:
    def __init__(self, llm=None, max_audit_chars=40000, max_source_chars=48000):
        self.llm = llm
        self.max_audit_chars = max_audit_chars
        self.max_source_chars = max_source_chars

    @staticmethod
    def _decode_result(result):
        if isinstance(result, dict):
            return result
        if isinstance(result, str):
            try:
                value = json.loads(result)
                return value if isinstance(value, dict) else None
            except json.JSONDecodeError:
                return None
        return None

    def plan(self, issue, context, project):
        """Plan a narrow fix for a specific diagnostic."""
        if not self.llm:
            return None
        prompt = {
            "task": "Propose a minimal, reversible repair to a VBA conversion artifact.",
            "issue": issue,
            "evidence": context,
            "project_artifacts": {
                k: {"path": v.get("path"), "sha256": v.get("sha256")}
                for k, v in project["artifacts"].items() if isinstance(v, dict)
            },
            "existing_project_code_files": project.get("project_files", []),
            "required_output": {
                "issue_id": "string", "artifact": "artifact key", "relative_path": "existing inventoried path",
                "operation": "json_set | json_replace | json_add | json_remove | text_replace | file_replace",
                "pointer": "RFC6901 JSON pointer for JSON operations",
                "old_value": "exact expected existing value", "new_value": "replacement value",
                "old_text": "exact unique text", "new_text": "replacement text",
                "new_content": "complete replacement file content only for small files",
                "rationale": "string", "source_lines": [1, 2], "confidence": 0.0,
                "expected_sha256": "exact current hash from inventory"
            },
            "constraints": [
                "Only patch an existing inventoried non-source file.",
                "Never modify original VBA source files (.bas, .cls, .frm, .vba, .vbs).",
                "Do not create files, suppress diagnostics, or rewrite unrelated content.",
                "Use exact old-value or old-text preconditions.",
                "For a whole-file replacement, preserve all unrelated behavior and return the complete file.",
                "If evidence is insufficient, return no_safe_patch=true."
            ]
        }
        try:
            result = self._decode_result(self.llm(prompt))
            if not result or result.get("no_safe_patch"):
                return None
            result["status"] = "proposed"
            result["issue_id"] = issue.get("id")
            return result
        except Exception:
            return None

    def audit_outputs(self, project):
        """Inspect every readable generated artifact/implementation file, not only known diagnostics."""
        if not self.llm:
            return {"findings": [], "proposals": [], "skipped": [
                {"reason": "No LLM configured; broad output audit was not run."}]}

        sources = project.get("artifacts", {}).get("sources", [])
        source_context = []
        remaining = self.max_source_chars
        for source in sources:
            text = source.get("text", "")
            if not text or remaining <= 0:
                continue
            excerpt = text[:remaining]
            source_context.append({"path": source.get("path"), "sha256": source.get("sha256"),
                                   "text": excerpt, "truncated": len(excerpt) < len(text)})
            remaining -= len(excerpt)

        inventory = [{"path": x.get("path"), "sha256": x.get("sha256"),
                      "size_bytes": x.get("size_bytes"), "suffix": x.get("suffix")}
                     for x in project.get("output_files", [])]
        inventory_hashes = {x.get("path"): x.get("sha256") for x in project.get("output_files", [])}
        inventory_hashes.update({x.get("path"): x.get("sha256") for x in project.get("project_files", [])})
        producer_candidates = self._producer_candidates(project)
        findings, proposals, skipped = [], [], []
        for item in project.get("output_files", []):
            if not item.get("readable"):
                skipped.append({"path": item.get("path"), "reason": item.get("reason", "Not readable")})
                continue
            content = item.get("text", "")
            if not content.strip():
                continue
            chunks = self._chunks(content, self.max_audit_chars)
            for chunk_index, (start_line, end_line, chunk) in enumerate(chunks, 1):
                prompt = {
                    "task": "Perform a rigorous, source-grounded quality audit of this generated conversion output file. Find both known and previously unknown defects.",
                    "objective": [
                        "Check correctness, completeness, consistency, provenance, syntax, data integrity, and downstream usability.",
                        "Compare translated/generated behavior against original VBA where applicable.",
                        "Check consistency with related artifact inventory and source evidence.",
                        "Identify omissions, invented behavior, incorrect control flow, type/coercion errors, array bounds, error handling, ByRef, object semantics, dates, numeric precision, string comparison, and target-language syntax when relevant.",
                        "Do not report stylistic preferences as defects unless they impair correctness or maintainability."
                    ],
                    "source_of_truth": source_context,
                    "artifact_inventory": inventory,
                    "converter_implementation_candidates": producer_candidates,
                    "target_file": {"path": item["path"], "sha256": item["sha256"],
                                    "suffix": item.get("suffix"), "size_bytes": item.get("size_bytes"),
                                    "chunk_index": chunk_index, "chunk_count": len(chunks),
                                    "start_line": start_line, "end_line": end_line, "content": chunk},
                    "required_output": {
                        "findings": [{"code": "stable_issue_code", "severity": "critical|error|warning|info",
                                      "description": "specific defect", "evidence": "exact evidence from source/output",
                                      "source_reference": "source path and line(s) or unavailable",
                                      "impact": "what conversion or behavior is affected",
                                      "confidence": 0.0}],
                        "patches": [{"target_kind": "generated_artifact|converter_implementation",
                                     "relative_path": "exact target_file.path or candidate path",
                                     "expected_sha256": "exact hash from target_file or candidate",
                                     "operation": "text_replace | json_set | json_replace | json_add | json_remove | file_replace",
                                     "pointer": "RFC6901 JSON pointer where applicable",
                                     "old_value": "exact current JSON value", "new_value": "replacement JSON value",
                                     "old_text": "exact unique existing text", "new_text": "replacement text",
                                     "new_content": "complete corrected file only when entire file was provided",
                                     "rationale": "why this preserves source semantics",
                                     "source_lines": [1], "confidence": 0.0}],
                        "no_safe_patch_reason": "string or null"
                    },
                    "hard_rules": [
                        "Treat original VBA as source of truth, but distinguish source facts from inference.",
                        "Do not alter the original VBA source; only propose patches to the exact target file.",
                        "Never invent missing source behavior or silently remove unsupported constructs.",
                        "Never suppress, delete, downgrade, or rewrite a diagnostic merely to make quality status green.",
                        "A patch must be minimal, reversible, and grounded in exact content shown in this prompt.",
                        "For chunked files, only use exact text_replace patches wholly contained in this chunk; do not return a full-file replacement or JSON pointer unless the entire file is present.",
                        "If the defect is caused by the converter implementation, prefer a minimal patch to a supplied implementation candidate and set target_kind=converter_implementation. If the generated artifact alone is wrong and the producer is correct, set target_kind=generated_artifact.",
                        "Implementation-candidate excerpts may be truncated: only propose exact text_replace patches contained in the displayed excerpt; never rewrite a partial file.",
                        "For JSON, preserve schema and unrelated keys; use JSON pointer operations where safe.",
                        "If a defect cannot be fixed confidently from available evidence, report it without a patch.",
                        "Return only a JSON object matching required_output."
                    ]
                }
                try:
                    result = self._decode_result(self.llm(prompt))
                except Exception as exc:
                    skipped.append({"path": item["path"], "chunk": chunk_index,
                                    "reason": f"LLM audit failed: {type(exc).__name__}: {exc}"})
                    continue
                if not result:
                    skipped.append({"path": item["path"], "chunk": chunk_index,
                                    "reason": "LLM returned invalid structured output"})
                    continue
                for finding in result.get("findings", []) if isinstance(result.get("findings"), list) else []:
                    if isinstance(finding, dict):
                        finding.update({"artifact": "output_file", "path": item["path"],
                                        "chunk": chunk_index, "status": "detected"})
                        findings.append(finding)
                if item.get("read_only"):
                    if result.get("patches"):
                        skipped.append({"path": item["path"], "chunk": chunk_index,
                                        "reason": "File is evidence/read-only; findings retained but patches suppressed"})
                    continue
                for patch in result.get("patches", []) if isinstance(result.get("patches"), list) else []:
                    if not isinstance(patch, dict):
                        continue
                    patch_path = patch.get("relative_path") or item["path"]
                    if patch_path not in inventory_hashes:
                        skipped.append({"path": patch_path, "chunk": chunk_index,
                                        "reason": "LLM proposed a path outside the inventoried project"})
                        continue
                    patch.update({"artifact": "output_file", "relative_path": patch_path,
                                  "expected_sha256": patch.get("expected_sha256") or inventory_hashes[patch_path],
                                  "target_kind": patch.get("target_kind") or (
                                      "converter_implementation" if patch_path != item["path"] else "generated_artifact"),
                                  "status": "proposed"})
                    proposals.append(patch)
        return {"findings": findings, "proposals": proposals, "skipped": skipped,
                "files_scanned": sum(1 for x in project.get("output_files", []) if x.get("readable")),
                "files_skipped": len(skipped)}

    @staticmethod
    def _producer_candidates(project):
        """Return small, hash-bound excerpts from likely converter implementation files."""
        candidates = []
        for item in project.get("project_files", [])[:]:
            path = item.get("path", "")
            if not path or path.lower().endswith((".bas", ".cls", ".frm", ".vba", ".vbs")):
                continue
            # The project loader already excludes the fixer package, dependencies and secrets.
            try:
                text = (Path(project["root"]) / path).read_text(encoding="utf-8-sig")
            except (OSError, UnicodeDecodeError):
                continue
            name = Path(path).name.lower()
            keywords = ("parser", "builder", "converter", "translator", "normaliz", "semantic", "lower", "generator", "runtime")
            score = sum(1 for word in keywords if word in name)
            if score:
                candidates.append((score, {"path": path, "sha256": item.get("sha256"),
                                           "excerpt": text[:6000], "truncated": len(text) > 6000}))
        return [x[1] for x in sorted(candidates, key=lambda x: (-x[0], x[1]["path"]))[:4]]

    @staticmethod
    def _chunks(text, max_chars):
        """Line-aware chunks with line-number provenance; chunks do not overlap to keep patches unambiguous."""
        lines = text.splitlines(keepends=True)
        if not lines:
            return []
        chunks, current, size, start, line_no = [], [], 0, 1, 1
        for line in lines:
            if current and size + len(line) > max_chars:
                chunks.append((start, line_no - 1, "".join(current)))
                current, size, start = [], 0, line_no
            current.append(line)
            size += len(line)
            line_no += 1
        if current:
            chunks.append((start, line_no - 1, "".join(current)))
        return chunks
