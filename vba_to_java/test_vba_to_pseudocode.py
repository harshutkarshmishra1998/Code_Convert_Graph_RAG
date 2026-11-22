import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from vba_to_java.vba_to_pseudocode import convert_vba_to_pseudocode, main


class ConvertVbaToPseudocodeTests(unittest.TestCase):
    def test_converts_procedure_control_flow_and_comments(self):
        source = '''\
Public Sub SumValues()
    Dim total As Long
    total = 0
    For i = 1 To 3
        total = total + i
    Next i
    If total > 0 Then
        MsgBox "Total: " & total ' report result
    Else
        MsgBox "Empty"
    End If
End Sub
'''

        result = convert_vba_to_pseudocode(source)

        self.assertIn("PROCEDURE SumValues()", result)
        self.assertIn("DECLARE total: Long", result)
        self.assertIn("FOR i FROM 1 TO 3", result)
        self.assertIn("IF total > 0 THEN", result)
        self.assertIn('DISPLAY "Total: " & total', result)
        self.assertIn("END PROCEDURE", result)
        self.assertNotIn("report result", result)

    def test_preserves_apostrophes_and_colons_inside_strings(self):
        result = convert_vba_to_pseudocode(
            "Sub SayIt()\nMsgBox \"It's ready: now\"\nEnd Sub"
        )

        self.assertIn('DISPLAY "It\'s ready: now"', result)

    def test_handles_case_and_unknown_statements(self):
        result = convert_vba_to_pseudocode(
            "Sub Choose()\n"
            "Select Case choice\n"
            "Case 1\n"
            "Call UseOption\n"
            "Case Else\n"
            "RaiseCustomError\n"
            "End Select\n"
            "End Sub"
        )

        self.assertIn("SELECT choice", result)
        self.assertIn("CASE 1", result)
        self.assertIn("CALL UseOption", result)
        self.assertIn("CASE OTHERWISE", result)
        self.assertIn("EXECUTE RaiseCustomError", result)

    def test_handles_nested_arguments_inline_if_labels_and_rem_comments(self):
        source = '''\
Private Function FindValue(Optional ByVal key As String = "a,b") As Variant
    Dim result As Variant
    If InStr(1, key, "Then") > 0 Then result = "Found: " & key Else result = "Not found"
    GoTo Done
Handler: Rem Skip the rest of this comment: including colons
    Resume Next
Done:
    FindValue = result
End Function
'''

        result = convert_vba_to_pseudocode(source)

        self.assertIn('FUNCTION FindValue(key: String = "a,b") RETURNS Variant', result)
        self.assertIn("IF InStr(1, key, \"Then\") > 0 THEN", result)
        self.assertIn('SET result TO "Found: " & key', result)
        self.assertIn("ELSE", result)
        self.assertIn("LABEL Handler", result)
        self.assertIn("LABEL Done", result)
        self.assertIn("RETRY NEXT STATEMENT", result)
        self.assertNotIn("including colons", result)

    def test_reports_unmatched_and_unterminated_blocks(self):
        result = convert_vba_to_pseudocode(
            "Sub Incomplete()\nFor i = 1 To 2\nEnd If\nEnd Sub\n"
        )

        self.assertIn("END IF [unmatched block terminator]", result)
        self.assertIn("END FOR [implicit: closed by END PROCEDURE]", result)
        self.assertIn("END PROCEDURE", result)

    def test_does_not_mistake_named_arguments_or_comparisons_for_assignments(self):
        result = convert_vba_to_pseudocode(
            "Sub OpenWorkbook()\n"
            "Workbooks.Open Filename:=path, ReadOnly:=True\n"
            "Debug.Print item = expected\n"
            "sheet.Cells(rowNumber, 1).Value = total\n"
            "End Sub"
        )

        self.assertIn(
            "EXECUTE Workbooks.Open Filename:=path, ReadOnly:=True",
            result,
        )
        self.assertIn("DISPLAY item = expected", result)
        self.assertIn("SET sheet.Cells(rowNumber, 1).Value TO total", result)

    def test_progressive_vba_samples_convert_without_unbalanced_blocks(self):
        samples_dir = Path(__file__).resolve().parent.parent / "files"
        samples = sorted(samples_dir.glob("vba_sample_*.bas"))

        self.assertEqual(len(samples), 5)
        for sample in samples:
            with self.subTest(sample=sample.name):
                result = convert_vba_to_pseudocode(
                    sample.read_text(encoding="utf-8-sig")
                )
                self.assertTrue(result)
                self.assertNotIn("[implicit:", result)
                self.assertNotIn("[unmatched", result)

    def test_main_accepts_input_and_output_paths(self):
        with TemporaryDirectory() as directory:
            input_path = Path(directory) / "module.bas"
            output_path = Path(directory) / "result" / "module.txt"
            input_path.write_text('Sub RunIt()\nMsgBox "Ready"\nEnd Sub\n', encoding="utf-8")

            result_path = main(input_path, output_path)

            self.assertEqual(result_path, output_path)
            self.assertIn('DISPLAY "Ready"', output_path.read_text(encoding="utf-8"))

            default_input = Path(directory) / "default.bas"
            default_input.write_text("Sub RunIt()\nEnd Sub\n", encoding="utf-8")
            default_output = main(default_input)

            self.assertEqual(default_output.name, "default.pseudocode.txt")
            self.assertTrue(default_output.is_file())

    def test_main_rejects_overwriting_the_input(self):
        with TemporaryDirectory() as directory:
            input_path = Path(directory) / "module.bas"
            input_path.write_text("Sub RunIt()\nEnd Sub\n", encoding="utf-8")

            with self.assertRaises(ValueError):
                main(input_path, input_path)


if __name__ == "__main__":
    unittest.main()
