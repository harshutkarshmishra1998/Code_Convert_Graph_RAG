from pathlib import Path

from .cli import main

project_root = Path(__file__).resolve().parents[2]
output_root = project_root / "files" / "vba_graph_output"
n = [6]
status = 0

for i in n:
    input_path = project_root / "files" / f"vba_sample_{i}.bas"
    output = output_root / f"sample_{i}"
    status = max(status, main(input_path, output))

raise SystemExit(status)
