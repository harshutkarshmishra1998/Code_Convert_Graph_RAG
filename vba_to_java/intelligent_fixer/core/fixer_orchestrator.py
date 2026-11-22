"""Iterative, source-grounded repair orchestration for arbitrary conversion outputs."""
from __future__ import annotations
import hashlib
from .project_loader import ProjectLoader
from .dependency_manager import DependencyManager
from .repair_state import RepairState, issue_id
from ..diagnostics.log_parser import LogParser
from ..diagnostics.issue_classifier import IssueClassifier
from ..diagnostics.root_cause_analyzer import RootCauseAnalyzer
from ..source_verification.vba_retriever import VBARetriever
from ..source_verification.procedure_context import ProcedureContext
from ..source_verification.dependency_context import DependencyContext
from ..source_verification.source_traceability import SourceTraceability
from ..agents.investigator import Investigator
from ..agents.repair_planner import RepairPlanner
from ..agents.repair_executor import RepairExecutor
from ..agents.independent_reviewer import IndependentReviewer
from ..validation.structural_validator import StructuralValidator
from ..validation.semantic_validator import SemanticValidator
from ..validation.regression_validator import RegressionValidator
from ..validation.compile_validator import CompileValidator
from ..validation.behavioral_validator import BehavioralValidator


class IntelligentFixer:
    def __init__(self, project_root, llm=None, reviewer_llm=None, rebuilders=None,
                 compile_runner=None, behavioral_oracle=None, allow_auto_apply=False,
                 max_audit_passes=3, reports_dir=None):
        self.loader = ProjectLoader(project_root)
        self.llm = llm
        self.rebuilders = rebuilders or {}
        self.allow_auto_apply = allow_auto_apply
        self.max_audit_passes = max(1, int(max_audit_passes))
        self.deps = DependencyManager()
        self.state = RepairState(
            reports_dir or self.loader.root / "intelligent_fixer" / "reports"
        )
        self.logs = LogParser()
        self.classifier = IssueClassifier()
        self.root_cause = RootCauseAnalyzer()
        self.retriever = VBARetriever()
        self.proc = ProcedureContext()
        self.dep_context = DependencyContext()
        self.trace = SourceTraceability()
        self.investigator = Investigator(llm)
        self.planner = RepairPlanner(llm)
        self.executor = RepairExecutor(self.loader.root)
        self.reviewer = IndependentReviewer(reviewer_llm)
        self.structural = StructuralValidator()
        self.semantic = SemanticValidator()
        self.regression = RegressionValidator()
        self.compile = CompileValidator(compile_runner)
        self.behavioral = BehavioralValidator(behavioral_oracle)

    def _discover_issues(self, project):
        artifacts = project["artifacts"]
        raw = []
        for kind, item in artifacts.items():
            if kind == "logs" and "text" in item:
                raw.extend(self.logs.parse(item["text"], item["path"]))
            if kind in {"graph", "ir", "semantic_analysis", "pseudocode_run", "report"} and "data" in item:
                raw.extend(self.logs.from_artifact(kind, item["data"], item["path"]))
        dedup = {}
        for item in raw:
            key = (item.get("code"), item.get("path"), item.get("line"), item.get("message"))
            if key not in dedup:
                dedup[key] = dict(item, reported_by=[item.get("artifact")])
            elif item.get("artifact") not in dedup[key]["reported_by"]:
                dedup[key]["reported_by"].append(item.get("artifact"))

        global_trace = self.trace.check({}, artifacts)
        issues = []
        for issue in dedup.values():
            issue = self.classifier.classify(issue)
            evidence = self.retriever.retrieve(issue, artifacts)
            context = {"source": evidence, "procedure": self.proc.build(evidence, artifacts),
                       "producer_code": self.loader.producer_code_context(issue),
                       "dependencies": self.dep_context.build(issue, artifacts),
                       "traceability": global_trace}
            cause = self.root_cause.analyze(issue, context, artifacts)
            investigation = self.investigator.investigate(issue, context, cause)
            issue.update({"id": issue_id(issue), "context": context, "root_cause": cause,
                          "investigation": investigation})
            issues.append(issue)

        structural = self.structural.validate(project)
        semantic = self.semantic.validate(project)
        for item in structural["issues"] + semantic["issues"]:
            item.setdefault("context", {"traceability": global_trace})
        combined = issues + structural["issues"] + semantic["issues"]
        compact = {}
        for item in combined:
            key = (item.get("code"), item.get("line"), item.get("procedure"), item.get("message"))
            if key not in compact:
                compact[key] = dict(item)
            else:
                compact[key].setdefault("reported_by", []).append(item.get("artifact") or "validator")
        all_issues = [{**item, "id": item.get("id") or issue_id(item)} for item in compact.values()]
        return all_issues, structural, semantic, global_trace

    @staticmethod
    def _inventory_hashes(project):
        allowed = {}
        for collection in (project.get("artifacts", {}).values(),
                           project.get("project_files", []), project.get("output_files", [])):
            for value in collection:
                if isinstance(value, dict) and value.get("path") and value.get("sha256"):
                    allowed[value["path"]] = value["sha256"]
        return allowed

    def _proposals(self, project, issues):
        proposals = []
        for issue in issues:
            proposal = self.planner.plan(issue, issue.get("context", {}), project)
            if proposal:
                proposals.append(proposal)
        # Broad LLM audit is not limited to the currently known diagnostic taxonomy.
        audit = self.planner.audit_outputs(project)
        for finding in audit.get("findings", []):
            finding.setdefault("message", finding.get("description", ""))
            finding.setdefault("id", issue_id(finding))
            finding.setdefault("context", {"traceability": self.trace.check({}, project["artifacts"])})
            issues.append(finding)
        for patch in audit.get("proposals", []):
            matching = next((x for x in audit.get("findings", [])
                             if x.get("path") == patch.get("relative_path")), None)
            linked_id = matching.get("id") if matching else issue_id(patch)
            patch.setdefault("issue_id", linked_id)
            if not any(x.get("id") == linked_id for x in issues):
                synthetic = {"id": linked_id, "code": "LLM_OUTPUT_REPAIR",
                             "path": patch.get("relative_path"),
                             "message": patch.get("rationale", "LLM-proposed output repair"),
                             "context": {"traceability": self.trace.check({}, project["artifacts"])}}
                issues.append(synthetic)
            proposals.append(patch)
        allowed = self._inventory_hashes(project)
        read_only_paths = {x.get("path") for x in project.get("output_files", []) if x.get("read_only")}
        valid = []
        for proposal in proposals:
            path = proposal.get("relative_path")
            if path not in allowed:
                proposal["status"] = "rejected_unrecognized_artifact_path"
                continue
            if proposal.get("expected_sha256") != allowed[path]:
                proposal["status"] = "rejected_hash_not_bound_to_inventory"
                continue
            if path.lower().endswith((".bas", ".cls", ".frm", ".vba", ".vbs")):
                proposal["status"] = "rejected_original_source_is_immutable"
                continue
            if path in read_only_paths or path.rsplit("/", 1)[-1].lower().startswith("manifest") or path.lower().endswith(".log"):
                proposal["status"] = "rejected_evidence_file_is_read_only"
                continue
            valid.append(proposal)
        # Avoid duplicate patches emitted by diagnostic planning and broad audits.
        unique = {}
        for p in valid:
            key = (p.get("relative_path"), p.get("operation"), p.get("pointer"),
                   p.get("old_text"), repr(p.get("old_value")))
            unique.setdefault(key, p)
        return list(unique.values()), audit

    def _apply_proposals(self, proposals, issues, project, apply):
        if not apply:
            return [], proposals
        if not self.allow_auto_apply:
            for proposal in proposals:
                proposal["status"] = "proposed_only_auto_apply_disabled"
            return [], proposals
        issue_by_id = {x.get("id"): x for x in issues}
        applied, rejected = [], []
        for proposal in proposals:
            issue = issue_by_id.get(proposal.get("issue_id"), {})
            trace = issue.get("context", {}).get("traceability", {})
            if not trace.get("safe_for_source_grounded_repair", False):
                proposal["status"] = "rejected_source_not_verified"
                rejected.append(proposal)
                self.state.append({"issue_id": proposal.get("issue_id"), "status": proposal["status"]})
                continue
            review = self.reviewer.review(proposal, project)
            if not review.get("approved", False):
                proposal["status"] = "rejected_by_reviewer"
                proposal["review"] = review
                rejected.append(proposal)
                self.state.append({"issue_id": proposal.get("issue_id"), "status": proposal["status"], "review": review})
                continue
            result = self.executor.apply(proposal)
            if result.get("applied"):
                result["artifact"] = proposal.get("artifact") or "output_file"
                result["target_kind"] = proposal.get("target_kind", "generated_artifact")
                result["issue_id"] = proposal.get("issue_id")
                result["operation"] = proposal.get("operation")
                result["rationale"] = proposal.get("rationale")
                applied.append(result)
                self.state.append({"issue_id": proposal.get("issue_id"), "status": "patch_applied", "result": result})
            else:
                proposal["status"] = "patch_failed"
                proposal["error"] = result.get("error")
                rejected.append(proposal)
        return applied, rejected

    def _rebuild(self, project, applied):
        changed = set()
        path_to_artifact = {v.get("path"): k for k, v in project["artifacts"].items()
                            if isinstance(v, dict) and v.get("path")}
        for patch in applied:
            artifact = path_to_artifact.get(patch.get("path"))
            if artifact:
                changed.add(artifact)
            elif patch.get("target_kind") == "converter_implementation":
                changed.add("source")
        plan = self.deps.plan(changed) if changed else {"changed": [], "invalidate_or_rebuild": []}
        results = []
        for name in plan["invalidate_or_rebuild"]:
            fn = self.rebuilders.get(name)
            if fn:
                try:
                    value = fn(self.loader.root)
                    results.append({"artifact": name, "result": value,
                                    "passed": not (isinstance(value, dict) and value.get("passed") is False)})
                except Exception as exc:
                    results.append({"artifact": name, "error": f"{type(exc).__name__}: {exc}", "passed": False})
            else:
                results.append({"artifact": name, "status": "not_configured", "passed": False})
        return plan, results

    def run(self, apply=False, max_passes=None, run_metadata=None):
        """Audit, repair, rebuild and repeat until stable or the pass limit is reached."""
        pass_limit = max(1, int(max_passes or self.max_audit_passes))
        pass_history, all_applied, last = [], [], None
        detected_by_id = {}
        self.state.append({"status": "run_started", "project_root": str(self.loader.root)})
        for pass_number in range(1, pass_limit + 1):
            before = self.loader.load()
            issues, structural, semantic, trace = self._discover_issues(before)
            proposals, audit = self._proposals(before, issues)
            pass_issues = {}
            for issue in issues:
                issue["id"] = issue.get("id") or issue_id(issue)
                pass_issues.setdefault(issue["id"], issue)
                detected_by_id.setdefault(issue["id"], issue)
            self.state.append_many([
                {
                    "issue_id": item["id"],
                    "status": "detected",
                    "pass": pass_number,
                    "issue": self._issue_summary(item),
                }
                for item in pass_issues.values()
            ])
            applied, rejected = self._apply_proposals(proposals, issues, before, apply)
            after = self.loader.load()
            rebuild_plan, rebuild_results = self._rebuild(before, applied)
            after = self.loader.load()
            regression = self.regression.validate(before, after, applied)
            final_struct = self.structural.validate(after)
            final_sem = self.semantic.validate(after)
            compile_result = self.compile.validate(self.loader.root)
            behavioral = self.behavioral.validate(before, after)
            pass_record = {"pass": pass_number, "issues_detected": len(issues),
                           "detected_issues": [
                               self._issue_summary(item) | {"id": item["id"]}
                               for item in pass_issues.values()
                           ],
                           "proposals": len(proposals), "applied": len(applied),
                           "rejected_or_failed": len(rejected), "audit": audit,
                           "rebuild_plan": rebuild_plan, "rebuild_results": rebuild_results,
                           "regression": regression, "structural": final_struct,
                           "semantic": final_sem, "compile": compile_result,
                           "behavioral": behavioral}
            pass_history.append(pass_record)
            all_applied.extend(applied)
            last = {"project": after, "issues": issues, "proposals": proposals,
                    "applied": applied, "rejected": rejected, "audit": audit,
                    "rebuild_plan": rebuild_plan, "rebuild_results": rebuild_results,
                    "regression": regression, "structural": final_struct,
                    "semantic": final_sem, "compile": compile_result,
                    "behavioral": behavioral}
            # If no patch was committed, another pass would only repeat the same proposals.
            if not applied:
                break

        # If the final allowed pass committed changes, perform a fresh read-only verification
        # so the report does not carry forward diagnostics from before those changes.
        if last and last.get("applied") and len(pass_history) >= pass_limit:
            verified_project = self.loader.load()
            verified_issues, verified_struct, verified_sem, _ = self._discover_issues(verified_project)
            verified_audit = self.planner.audit_outputs(verified_project)
            for finding in verified_audit.get("findings", []):
                finding.setdefault("message", finding.get("description", ""))
                verified_issues.append(finding)
            last.update({"project": verified_project, "issues": verified_issues,
                         "audit": verified_audit, "structural": verified_struct, "semantic": verified_sem})

        unresolved = []
        if last:
            unresolved.extend(last["issues"])
            unresolved.extend(last["audit"].get("findings", []))
            unresolved.extend(last["structural"].get("issues", []))
            unresolved.extend(last["semantic"].get("issues", []))
        compact = {}
        for item in unresolved:
            item["id"] = item.get("id") or issue_id(item)
            key = (item.get("code"), item.get("path"), item.get("line"),
                   item.get("procedure"), item.get("description") or item.get("message"))
            compact.setdefault(key, {**item, "status": "unresolved_or_unverified"})
        unresolved = list(compact.values())
        unresolved_by_id = {item["id"]: item for item in unresolved}
        final_scan_findings = [
            item for item in unresolved if item["id"] not in detected_by_id
        ]
        for item in final_scan_findings:
            detected_by_id[item["id"]] = item
        self.state.append_many([
            {
                "issue_id": item["id"],
                "status": "detected",
                "pass": "final_validation",
                "issue": self._issue_summary(item),
            }
            for item in final_scan_findings
        ])
        applied_by_id = {}
        for change in all_applied:
            if change.get("issue_id"):
                applied_by_id.setdefault(change["issue_id"], []).append({
                    key: change.get(key) for key in (
                        "path", "operation", "before_sha256", "after_sha256",
                        "backup", "rationale",
                    ) if change.get(key) is not None
                })
        issue_resolution_log = []
        all_issue_ids = dict(detected_by_id)
        for item in unresolved:
            all_issue_ids.setdefault(item["id"], item)
        for item_id, issue in all_issue_ids.items():
            remaining = item_id in unresolved_by_id
            changes = applied_by_id.get(item_id, [])
            if changes and not remaining:
                status = "addressed_and_verified"
            elif changes:
                status = "patch_applied_but_issue_remains"
            elif remaining:
                status = "detected_unresolved"
            else:
                status = "not_reproduced_on_final_scan"
            record = {
                "issue_id": item_id,
                "issue": self._issue_summary(issue),
                "detected": item_id in detected_by_id,
                "addressed": status == "addressed_and_verified",
                "status": status,
                "changes": changes,
                "verification": {
                    "present_in_final_issue_set": remaining,
                    "absent_after_final_scan": not remaining,
                },
            }
            issue_resolution_log.append(record)
            self.state.append({
                "issue_id": item_id,
                "status": status,
                "addressed": record["addressed"],
                "changes": changes,
                "verification": record["verification"],
            })
        rebuild_complete = all(x.get("passed") is True for p in pass_history for x in p.get("rebuild_results", []))
        verified = (not unresolved and last and last["structural"].get("errors", 0) == 0
                    and last["semantic"].get("blocking", 0) == 0
                    and last["regression"].get("passed") is True
                    and last["compile"].get("status") == "passed"
                    and last["behavioral"].get("status") == "verified"
                    and rebuild_complete)
        report = {"schema_version": "2.0", "project_root": str(self.loader.root),
                  "audit_passes": len(pass_history), "max_audit_passes": pass_limit,
                  "issue_count": len(unresolved), "applied_patch_count": len(all_applied),
                  "unresolved_issue_count": len(unresolved), "rebuild_complete": rebuild_complete,
                  "addressed_issue_count": sum(1 for item in issue_resolution_log if item["addressed"]),
                  "issue_resolution_log": issue_resolution_log,
                  "source_hashes": [{"path": x.get("path"), "sha256": x.get("sha256")}
                                    for x in last["project"].get("artifacts", {}).get("sources", [])] if last else [],
                  "passes": pass_history,
                  "validation": {"structural": last["structural"] if last else {},
                                 "semantic": last["semantic"] if last else {},
                                 "regression": last["regression"] if last else {},
                                 "compile": last["compile"] if last else {},
                                 "behavioral": last["behavioral"] if last else {}},
                  "conversion_readiness": "VERIFIED" if verified else "NOT_CERTIFIED",
                  "limitations": [
                      "The LLM can discover and propose fixes beyond predefined diagnostic categories, but cannot guarantee discovery of every defect.",
                      "Original VBA files are immutable; generated artifacts and inventoried implementation files may be patched.",
                      "Compilation and behavioral equivalence require trusted injected runners and a representative VBA reference oracle.",
                      "Files exceeding the audit size limit or unsupported/binary files are reported as skipped, not silently considered clean."
                  ]}
        self.state.write_unresolved(unresolved)
        self.state.write_quality(report)
        self.state.write_issue_run({
            "project_root": str(self.loader.root),
            "entry_point": run_metadata or {},
            "llm_enabled": self.llm is not None,
            "llm_adapter": type(self.llm).__name__ if self.llm is not None else None,
            "checked_artifacts": [
                {"kind": kind, "path": item.get("path"), "sha256": item.get("sha256")}
                for kind, item in (last or {}).get("project", {}).get("artifacts", {}).items()
                if isinstance(item, dict) and item.get("path")
            ],
            "issue_count": len(issue_resolution_log),
            "addressed_issue_count": report["addressed_issue_count"],
            "unresolved_issue_count": len(unresolved),
            "issues": issue_resolution_log,
            "applied_changes": applied_by_id,
            "pseudocode_generation_check": self._pseudocode_check(
                (last or {}).get("project", {}).get("artifacts", {}).get("pseudocode_run", {}),
                (last or {}).get("project", {}).get("artifacts", {}).get("semantic_analysis", {}),
            ),
        })
        self.state.append({
            "status": "run_completed",
            "issue_count": len(issue_resolution_log),
            "addressed_issue_count": report["addressed_issue_count"],
            "unresolved_issue_count": len(unresolved),
            "conversion_readiness": report["conversion_readiness"],
        })
        return {**(last or {}), "issues": last["issues"] if last else [],
                "applied": all_applied, "unresolved": unresolved,
                "pass_history": pass_history, "quality_report": report,
                "issue_resolution_log": issue_resolution_log}

    @staticmethod
    def _issue_summary(issue):
        keys = (
            "code", "severity", "priority", "subsystem", "message", "description",
            "path", "line", "procedure", "artifact", "reported_by",
        )
        return {key: issue[key] for key in keys if issue.get(key) is not None}

    @staticmethod
    def _pseudocode_check(artifact, semantic_artifact=None):
        data = artifact.get("data") or {}
        failed_sections = data.get("failed_sections") or []
        status = data.get("status")
        semantic_data = (semantic_artifact or {}).get("data") or {}
        summary = semantic_data.get("summary") or {}
        coverage_gaps = data.get("pseudocode_coverage_gaps")
        if not isinstance(coverage_gaps, list):
            coverage_gaps = semantic_data.get("pseudocode_coverage_gaps")
        coverage_gap_count = (
            len(coverage_gaps)
            if isinstance(coverage_gaps, list)
            else summary.get("pseudocode_procedure_sections_missing", 0)
        )
        client_check_passed = status == "success" and not failed_sections
        issue_detected = (
            (bool(data) and not client_check_passed) or bool(coverage_gap_count)
        )
        current_run_passed = bool(data) and client_check_passed and coverage_gap_count == 0
        return {
            "artifact": artifact.get("path"),
            "issue_detected": issue_detected,
            "current_run_passed": current_run_passed,
            "status": status or "not_available",
            "generated_sections": data.get("generated_sections", 0),
            "failed_sections": failed_sections,
            "coverage_gap_count": coverage_gap_count,
            "coverage_gaps": coverage_gaps if isinstance(coverage_gaps, list) else [],
            "known_previous_issue": {
                "issue_id": "PSEUDOCODE_LLM_CLIENT_UNSUPPORTED",
                "detected": True,
                "detection_source": "previous user-reported pseudocode execution",
                "description": "The configured LLM value was a model identifier, not an invokable chat client.",
                "addressed": client_check_passed,
                "status": "addressed_and_verified" if client_check_passed else "not_verified",
                "change": (
                    "The pseudocode LLM adapter now constructs a ChatGroq client "
                    "from the configured model."
                ),
                "changed_file": "vba_to_java/pseudocode/generate_pseudocode.py",
                "verification": {
                    "run_status": status or "not_available",
                    "generated_sections": data.get("generated_sections", 0),
                    "failed_sections": failed_sections,
                    "coverage_gap_count": coverage_gap_count,
                },
            },
        }
