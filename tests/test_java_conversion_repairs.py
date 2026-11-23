import json
import unittest

from vba_to_java_2.java_conversion.converter import (
    _java_client_prompt,
    _normalize_repair_error,
    _repair_negated_string_concatenation,
)


class JavaConversionRepairTests(unittest.TestCase):
    def test_parenthesizes_concatenation_before_negated_string_comparison(self):
        source = (
            "class Sample {\n"
            "    void check(String inputFolder, String f, String path) {\n"
            "        if (!inputFolder + f\n"
            "                .equalsIgnoreCase(path)) {}\n"
            "    }\n"
            "}\n"
        )
        compile_error = (
            "/tmp/vba-java-check-abc/Sample.java:3: error: "
            "bad operand type String for unary operator '!'\n"
        )

        repaired = _repair_negated_string_concatenation(source, compile_error)

        self.assertIn(
            "if (!(inputFolder + f).equalsIgnoreCase(path)) {}",
            repaired,
        )

    def test_leaves_unrelated_compile_errors_unchanged(self):
        source = "class Sample {}\n"
        compile_error = "error: cannot find symbol\n"

        self.assertEqual(
            _repair_negated_string_concatenation(source, compile_error),
            source,
        )

    def test_normalizes_temporary_paths_and_line_numbers(self):
        first = (
            "/tmp/vba-java-check-a/Sample.java:3: error: cannot find symbol"
        )
        second = (
            "/tmp/vba-java-check-b/Sample.java:7: error: cannot find symbol"
        )

        self.assertEqual(
            _normalize_repair_error(first),
            _normalize_repair_error(second),
        )

    def test_repair_prompt_addresses_vba_default_and_byref_type_errors(self):
        prompt = json.loads(_java_client_prompt(
            source="",
            process_note="",
            class_name="Sample",
            procedures=[],
            calculations=[],
            prior_source="class Sample {}",
            repair_errors=[
                "variable itemCode might not have been initialized",
                "String cannot be converted to StringBuilder",
            ],
        ))

        self.assertIn("VBA type", prompt["repair_instruction"])
        self.assertIn("caller, callee", prompt["repair_instruction"])


if __name__ == "__main__":
    unittest.main()
