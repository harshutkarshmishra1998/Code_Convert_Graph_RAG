"""Normalize severity and classify conversion issues by likely subsystem."""
class IssueClassifier:
    RULES={
      "CROSS_NESTED_BLOCK_CLOSE":("frontend_control_flow","high"),
      "IR_EXPRESSIONS_OPAQUE":("expression_frontend","high"),
      "IR_FRONTEND_HIERARCHY_REPAIRED":("frontend_control_flow","high"),
      "VBA_ERROR_HANDLING":("runtime_semantics","blocking"),
      "VBA_GOSUB_RETURN":("control_flow_lowering","blocking"),
      "LATE_BOUND_OBJECT_MEMBERS":("type_resolution","blocking"),
      "COM_CREATEOBJECT":("runtime_interop","blocking"),
      "VBA_ARRAY_INITIALIZATION_IDIOM":("array_semantics","blocking"),
      "PSEUDOCODE_SECTION_FAILED":("documentation_generation","high"),
    }
    def classify(self,issue):
        code=issue.get("code","UNCLASSIFIED")
        subsystem,priority=self.RULES.get(code,("unknown","normal"))
        severity=issue.get("severity","warning").lower()
        if severity=="review": severity="warning"
        if priority=="blocking": severity="blocking"
        return {**issue,"subsystem":subsystem,"priority":priority,"severity":severity,
                "auto_fix_policy":"source_evidence_and_independent_validation_required"}
