"""Evidence-based root-cause hypotheses; never equates a diagnostic with its cause."""
class RootCauseAnalyzer:
    def analyze(self,issue,context,artifacts):
        code=issue.get("code","")
        evidence=context.get("source",{})
        present=bool(evidence.get("text"))
        hypotheses=[]
        if code=="CROSS_NESTED_BLOCK_CLOSE":
            hypotheses += ["Parser block-stack or branch-association defect",
                           "Source syntax or continuation handling not represented correctly"]
        elif code=="IR_EXPRESSIONS_OPAQUE":
            hypotheses += ["Expression grammar does not cover one or more VBA operators/constructs",
                           "Frontend emitted source-backed operands without expression AST"]
        elif code=="IR_FRONTEND_HIERARCHY_REPAIRED":
            hypotheses += ["Frontend required heuristic hierarchy repairs; verify parser's canonical structure"]
        elif issue.get("subsystem")=="runtime_semantics":
            hypotheses += ["Target backend lacks an explicit lowering/runtime implementation"]
        elif issue.get("subsystem")=="documentation_generation":
            hypotheses += ["Pseudocode generator omitted or failed a procedure section"]
        else: hypotheses.append("Insufficient diagnostic-specific evidence; inspect source and producing stage")
        return {"hypotheses":hypotheses,"source_available":present,
          "source_lines":evidence.get("line_range"),"confidence":"medium" if present else "low",
          "recommended_action":"inspect_producer_and_rebuild_downstream",
          "warning":"A root-cause hypothesis is not proof. Confirm against producer code and regenerated artifacts."}
