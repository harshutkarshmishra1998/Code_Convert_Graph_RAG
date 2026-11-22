"""Check semantic completeness without claiming target-language equivalence."""
class SemanticValidator:
    def validate(self,project):
        a=project["artifacts"]; issues=[]; blocking=0
        sem=a.get("semantic_analysis",{}).get("data",{})
        summary=sem.get("summary",{})
        for key in ("unresolved_expression_count","parse_error_count","issue_count"):
            value=summary.get(key,0)
            if value:
                issues.append({"code":f"SEMANTIC_{key.upper()}","severity":"blocking","message":f"{key}={value}"})
                blocking+=1
        for req in sem.get("compatibility_requirements",[]):
            if req.get("severity")=="blocking":
                issues.append({"code":req.get("code"),"severity":"blocking",
                  "message":req.get("title"),"target_guidance":req.get("target_guidance")})
                blocking+=1
        for name in sem.get("pseudocode_coverage_gaps",[]):
            issues.append({"code":"PSEUDOCODE_COVERAGE_GAP","severity":"warning","procedure":name,
                           "message":f"Detailed pseudocode section missing: {name}"})
        return {"blocking":blocking,"issues":issues,
          "summary":summary,"migration_readiness":sem.get("migration_readiness","unknown"),
          "note":"A zero semantic issue count does not prove behavior-preserving translation."}
