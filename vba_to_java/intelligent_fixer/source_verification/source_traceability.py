"""Check source provenance and report traceability gaps."""
class SourceTraceability:
    def check(self,issue,artifacts):
        sources=artifacts.get("sources") or ([artifacts["source"]] if artifacts.get("source") else [])
        manifest=artifacts.get("manifest",{}).get("data",[])
        expected={x.get("path"):x.get("sha256") for x in manifest if isinstance(x,dict)} if isinstance(manifest,list) else {}
        comparisons=[]
        for src in sources:
            exp=expected.get(src.get("path"))
            comparisons.append({"path":src.get("path"),"actual":src.get("sha256"),
                                 "expected":exp,"matches":(src.get("sha256")==exp) if exp else None})
        has_source=bool(sources)
        mismatch=any(x["matches"] is False for x in comparisons)
        missing_manifest=bool(expected) and any(x["path"] not in {s.get("path") for s in sources} for x in manifest if isinstance(x,dict))
        safe=has_source and not mismatch and not missing_manifest
        return {"source_present":has_source,"sources":comparisons,
          "hash_matches":not mismatch and not missing_manifest if expected else None,
          "safe_for_source_grounded_repair":safe,
          "warning":"Source location is supporting evidence; generated artifact offsets must be mapped, not assumed."}
