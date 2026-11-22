"""Independent review of source-grounded patches; fail closed by default."""
class IndependentReviewer:
    def __init__(self, llm=None):
        self.llm = llm

    def review(self, proposal, project):
        if not self.llm:
            return {"approved": False, "reason": "Independent reviewer LLM is not configured."}
        path = proposal.get("relative_path")
        target = next((x for x in project.get("output_files", []) if x.get("path") == path), None)
        if target is None:
            target = next((x for x in project.get("project_files", []) if x.get("path") == path), None)
        source_context = []
        source_budget = 48000
        for source in project.get("artifacts", {}).get("sources", []):
            if source_budget <= 0:
                break
            source_text = source.get("text", "")[:source_budget]
            source_context.append({"path": source.get("path"), "sha256": source.get("sha256"),
                                   "text": source_text, "truncated": len(source_text) < len(source.get("text", ""))})
            source_budget -= len(source_text)
        if target and isinstance(target.get("text"), str):
            target_text = target["text"][:24000]
            target = {**target, "text": target_text, "truncated": len(target_text) < len(target.get("text", ""))}
        prompt = {
            "task": "Independently audit this proposed repair for correctness, safety and semantic fidelity. Do not trust the planner's rationale.",
            "proposal": proposal,
            "target_file": {k: target.get(k) for k in ("path", "sha256", "suffix", "size_bytes", "text")} if target else None,
            "original_vba_source": source_context,
            "artifact_inventory": [{"path": x.get("path"), "sha256": x.get("sha256")}
                                   for x in project.get("output_files", [])],
            "review_criteria": [
                "Check that the target path exists in the supplied inventory and is not original VBA source.",
                "Verify the old_text/old_value precondition against the supplied current content.",
                "Check the change against original VBA and related artifact evidence; identify unsupported assumptions.",
                "Reject diagnostic suppression, invented behavior, broad rewrites, unrelated edits and provenance damage.",
                "Reject whole-file replacement unless the complete current file is present and the proposed result is demonstrably complete.",
                "Consider syntax, schema, control flow, types, runtime semantics, data loss and downstream effects.",
                "Approve only if the patch is narrow, evidence-backed, and has a credible validation plan.",
                "If evidence is incomplete or ambiguous, reject."
            ],
            "output": {"approved": "boolean", "reason": "string", "risks": ["string"],
                       "required_tests": ["string"], "source_evidence_checked": "boolean"}
        }
        try:
            result = self.llm(prompt)
            if isinstance(result, dict) and result.get("approved") is True and result.get("source_evidence_checked") is True:
                return result
            return {"approved": False, "reason": "Reviewer did not explicitly approve with source evidence checked.",
                    "review_result": result}
        except Exception as exc:
            return {"approved": False, "reason": f"Reviewer failed: {type(exc).__name__}: {exc}"}
