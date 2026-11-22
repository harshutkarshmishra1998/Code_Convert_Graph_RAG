"""Gather artifact and symbol dependencies relevant to an issue."""
class DependencyContext:
    def build(self,issue,artifacts):
        out={"artifact_hashes":{k:v.get("sha256") for k,v in artifacts.items() if isinstance(v,dict) and v.get("sha256")},
             "related_artifacts":[],"symbols":[]}
        code=issue.get("code","")
        mapping={"CROSS_NESTED_BLOCK_CLOSE":["graph","segments","ir","normalized_ir","semantic_analysis"],
                 "IR_EXPRESSIONS_OPAQUE":["ir","normalized_ir","semantic_analysis"],
                 "IR_FRONTEND_HIERARCHY_REPAIRED":["graph","segments","ir","normalized_ir"],
                 "VBA_ERROR_HANDLING":["source","ir","normalized_ir","semantic_analysis"],
                 "VBA_GOSUB_RETURN":["source","ir","normalized_ir","semantic_analysis"],
                 "LATE_BOUND_OBJECT_MEMBERS":["source","ir","normalized_ir","semantic_analysis"]}
        out["related_artifacts"]=mapping.get(code,list(artifacts))
        return out
