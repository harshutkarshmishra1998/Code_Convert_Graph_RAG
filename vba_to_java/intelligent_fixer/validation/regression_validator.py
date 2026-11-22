"""Detect accidental loss of source and structural content after a repair."""
class RegressionValidator:
    def validate(self,before,after,applied):
        ba=before["artifacts"]; aa=after["artifacts"]
        checks={}
        for kind in ("source","graph","segments","ir","normalized_ir"):
            b=ba.get(kind,{}); a=aa.get(kind,{})
            checks[kind+"_present"]=not bool(b) or bool(a)
            if kind=="source" and b.get("sha256"): checks["source_unchanged"]=b.get("sha256")==a.get("sha256")
        # Never count disappearance of a diagnostic as proof of repair.
        before_n=len(ba.get("graph",{}).get("data",{}).get("nodes",[]))
        after_n=len(aa.get("graph",{}).get("data",{}).get("nodes",[]))
        checks["graph_node_count_nonzero"]=after_n>0 if before_n else True
        return {"passed":all(v for v in checks.values()),"checks":checks,
          "applied_patch_count":len(applied),
          "warning":"Diagnostic disappearance is not a semantic regression test."}
