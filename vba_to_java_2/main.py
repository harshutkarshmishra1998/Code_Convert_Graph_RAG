"""Generate process notes and Java conversions for selected VBA samples."""

from pathlib import Path

from .java_conversion.converter import main as convert_to_java
from .process_notes.notes_generator import main as generate_process_notes


INPUT_NAMES = {
    1: "vba_sample_1.bas",
    2: "vba_sample_2.bas",
    3: "vba_sample_3.bas",
    4: "vba_sample_4.bas",
    5: "vba_sample_5.bas",
    6: "vba_sample_6.bas",
    7: "vba_stage1_edge_cases.bas",
    8: "vba_stage1_malformed.bas",
}


def main(sample_ids):
    project_root = Path(__file__).resolve().parents[1]
    input_root = project_root / "files" / "vba" / "input"
    output_root = project_root / "files" / "vba" / "output"
    for sample_id in sample_ids:
        source = input_root / INPUT_NAMES[sample_id]
        sample_output = output_root / f"sample_{sample_id}"
        if not generate_process_notes(source, sample_output):
            return False
        if not convert_to_java(source, sample_output / "process_notes.md", sample_output):
            return False
    return True


if __name__ == "__main__":
    main(sample_ids=(1,2,3,4,5,6,7,8))