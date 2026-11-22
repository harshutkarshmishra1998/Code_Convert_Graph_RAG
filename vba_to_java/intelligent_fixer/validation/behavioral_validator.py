"""Behavioral equivalence requires an executable oracle; absence is reported honestly."""
class BehavioralValidator:
    def __init__(self,oracle=None): self.oracle=oracle
    def validate(self,before,after):
        if not self.oracle:
            return {"status":"not_verified","equivalent":None,
              "reason":"No paired VBA/reference and target-language test oracle supplied.",
              "required":["same input corpus","same external dependencies or controlled adapters",
                "compare return values, mutations, errors, ordering, dates, numeric rounding and side effects"]}
        try:
            result=self.oracle(before,after)
            return {"status":"verified" if result.get("equivalent") is True else "failed_or_inconclusive",**result}
        except Exception as e:
            return {"status":"error","equivalent":None,"reason":f"{type(e).__name__}: {e}"}
