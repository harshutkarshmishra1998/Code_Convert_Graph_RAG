import unittest

from vba_to_java.semantic_analyzer.expression_parser import (
    parse_expression,
    parse_vba_call_statement,
)
from vba_to_java.semantic_analyzer.normalize import normalize_ir
from vba_to_java.vba_graph.parser import classify_statement


class VbaParsingTests(unittest.TestCase):
    def test_named_arguments_are_not_misclassified_as_assignments(self):
        examples = (
            "sourceBook.Close SaveChanges:=False",
            "outputBook.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbook",
            'inventorySheet.Range("A1:D" & lastRow).Sort Key1:=inventorySheet.Range("A1"), Header:=xlYes',
        )
        for source in examples:
            with self.subTest(source=source):
                node = classify_statement(source)
                self.assertEqual(node["type"], "CALL_OR_COMMAND")
                self.assertEqual(node["expression"]["source"], source)

        assignment = classify_statement("rowNumber = rowNumber + 1")
        self.assertEqual(assignment["type"], "ASSIGNMENT")
        self.assertEqual(assignment["target"], "rowNumber")

    def test_parenthesized_and_statement_calls_preserve_named_arguments(self):
        parenthesized = parse_expression(
            "Workbooks.Open(Filename:=inputFolder & fileName, ReadOnly:=True)"
        )
        self.assertEqual(
            [argument["name"] for argument in parenthesized["arguments"]],
            ["Filename", "ReadOnly"],
        )
        self.assertEqual(parenthesized["arguments"][0]["kind"], "named_argument")

        statement_call = parse_vba_call_statement(
            "sourceBook.Close SaveChanges:=False"
        )
        self.assertEqual(statement_call["arguments"][0]["name"], "SaveChanges")
        self.assertEqual(statement_call["arguments"][0]["value"]["value"], False)

    def test_multiline_call_and_implicit_with_member_parse(self):
        call = parse_vba_call_statement(
            'inventorySheet.Range("A1:D" & lastRow).Sort _\n'
            '    Key1:=inventorySheet.Range("A1"), Header:=xlYes'
        )
        self.assertEqual(
            [argument["name"] for argument in call["arguments"]],
            ["Key1", "Header"],
        )

        expression = parse_expression('.Cells(rowNumber, "C").Value')
        receiver = expression["object"]["callee"]["object"]
        self.assertEqual(receiver["name"], "__vba_with_receiver__")
        self.assertTrue(receiver["semantics"]["implicit_with_receiver"])

    def test_normalization_resolves_module_constants_and_with_context(self):
        unsupported = lambda source: {
            "kind": "unsupported",
            "semantics": {
                "reason": "Frontend supplied source text but no expression AST",
                "source_expression": source,
                "parse_status": "not_parsed",
                "requires_semantic_resolution": True,
            },
        }
        program = {
            "modules": [
                {
                    "name": "Sample",
                    "declarations": [
                        {
                            "kind": "constant",
                            "name": "LIMIT",
                            "type": {"kind": "primitive", "name": "int32"},
                            "value": unsupported("30"),
                        }
                    ],
                    "procedures": [
                        {
                            "name": "Update",
                            "kind": "procedure",
                            "parameters": [],
                            "local_variables": [
                                {"name": "sheet", "type": {"kind": "object", "name": "Worksheet"}},
                                {"name": "rowNumber", "type": {"kind": "primitive", "name": "int32"}},
                            ],
                            "body": [
                                {
                                    "kind": "with",
                                    "expression": unsupported("sheet"),
                                    "body": [
                                        {
                                            "kind": "assignment",
                                            "target": {"kind": "identifier", "name": "rowNumber"},
                                            "value": unsupported('.Cells(rowNumber, "C").Value'),
                                        }
                                    ],
                                }
                            ],
                        }
                    ],
                }
            ]
        }

        normalized, report = normalize_ir(program)
        declaration_value = normalized["modules"][0]["declarations"][0]["value"]
        with_value = normalized["modules"][0]["procedures"][0]["body"][0]["body"][0]["value"]

        self.assertEqual(declaration_value["kind"], "literal")
        self.assertEqual(declaration_value["value"], 30)
        self.assertEqual(with_value["kind"], "member_access")
        self.assertEqual(
            with_value["object"]["callee"]["object"]["binding"]["symbol_kind"],
            "implicit_with_receiver",
        )
        self.assertEqual(report["summary"]["parse_error_count"], 0)
        self.assertEqual(report["summary"]["unresolved_expression_count"], 0)


if __name__ == "__main__":
    unittest.main()
