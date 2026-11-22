"""LLM-assisted investigation. LLM is injected by caller; no provider dependency."""
class Investigator:
    def __init__(self,llm=None): self.llm=llm
    def investigate(self,issue,context,cause):
        base={"summary":"No LLM configured; deterministic evidence package only.",
              "evidence_sufficient":bool(context.get("source",{}).get("available")),
              "root_cause":cause,"recommended_next_step":"inspect producer code and regenerate"}
        if not self.llm: return base
        prompt={"task":"Investigate a VBA conversion diagnostic. Treat original VBA as source of truth.",
          "issue":issue,"source_evidence":context.get("source"),
          "procedure_context":context.get("procedure"),"dependencies":context.get("dependencies"),
          "traceability":context.get("traceability"),"root_cause_hypotheses":cause,
          "rules":["Do not invent source facts.","Separate proven facts from hypotheses.",
                   "Do not propose a patch yet.","Return JSON object with facts, hypotheses, confidence, missing_evidence, producer_to_inspect."]}
        try:
            result=self.llm(prompt)
            return result if isinstance(result,dict) else {**base,"llm_raw":str(result)}
        except Exception as e: return {**base,"llm_error":f"{type(e).__name__}: {e}"}
