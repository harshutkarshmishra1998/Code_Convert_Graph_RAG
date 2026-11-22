import json
import tempfile
import unittest
from pathlib import Path

from vba_to_java.intelligent_fixer.core.fixer_orchestrator import (
    _mark_downstream_resolutions,
)
from vba_to_java.intelligent_fixer.core.repair_state import RepairState


class IntelligentFixerTests(unittest.TestCase):
    def test_opaque_ir_issue_is_resolved_only_with_clean_normalized_evidence(self):
        issues = [{"code": "IR_EXPRESSIONS_OPAQUE"}]
        artifacts = {
            "normalized_ir": {
                "data": {"modules": [{"procedures": [{"body": [{"kind": "literal"}]}]}]}
            },
            "semantic_analysis": {
                "data": {
                    "summary": {
                        "parse_error_count": 0,
                        "unresolved_expression_count": 0,
                    }
                }
            },
        }

        _mark_downstream_resolutions(issues, artifacts)
        self.assertEqual(
            issues[0]["downstream_resolution"]["method"],
            "semantic_normalization",
        )

        incomplete = [{"code": "IR_EXPRESSIONS_OPAQUE"}]
        artifacts["normalized_ir"]["data"]["modules"][0]["procedures"][0]["body"].append(
            {"kind": "unresolved_expression"}
        )
        _mark_downstream_resolutions(incomplete, artifacts)
        self.assertNotIn("downstream_resolution", incomplete[0])

    def test_corrupt_repair_history_is_not_silently_discarded(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "repair_history.json"
            path.write_text("{", encoding="utf-8")
            state = RepairState(directory)

            with self.assertRaises(json.JSONDecodeError):
                state.read(path, {"records": []})


if __name__ == "__main__":
    unittest.main()
