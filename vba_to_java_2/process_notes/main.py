"""Select the VBA samples and run the process-notes module."""

from pathlib import Path

from .notes_generator import main


if __name__ == "__main__":
    project_root = Path(__file__).resolve().parents[2]
    input_root = project_root / "files" / "vba" / "input"
    output_root = project_root / "files" / "vba" / "output"
    # n = [1, 2, 3, 4, 5, 6, 7, 8]
    n = [6]
    input_names = {
        1: "vba_sample_1.bas", 2: "vba_sample_2.bas", 3: "vba_sample_3.bas",
        4: "vba_sample_4.bas", 5: "vba_sample_5.bas", 6: "vba_sample_6.bas",
        7: "vba_stage1_edge_cases.bas", 8: "vba_stage1_malformed.bas",
    }
    for i in n:
        if not main(input_root / input_names[i], output_root / f"sample_{i}"):
            break
