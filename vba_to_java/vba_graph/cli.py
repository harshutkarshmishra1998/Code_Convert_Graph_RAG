import argparse,json,sys
from pathlib import Path
from .engine import convert_path
from .schema import validate_graph

def main(argv=None):
    p=argparse.ArgumentParser(description="Hybrid VBA segmentation and hierarchical JSON graph")
    p.add_argument("input",help="Exported VBA file or directory")
    p.add_argument("-o","--output",default="vba_graph_output")
    p.add_argument("--strict",action="store_true",help="Exit non-zero when any diagnostic exists")
    p.add_argument("--llm-json",help="Optional JSON annotations from an LLM; cannot alter structure/source")
    a=p.parse_args(argv)
    try:
        payload=json.loads(Path(a.llm_json).read_text(encoding="utf-8")) if a.llm_json else None
        result=convert_path(a.input,payload)
        out=Path(a.output); out.mkdir(parents=True,exist_ok=True)
        for filename,key in (("graph.json","graph"),("segments.json","segments"),("report.json","report")):
            (out/filename).write_text(json.dumps(result[key],indent=2,ensure_ascii=False),encoding="utf-8")
        (out/"manifest.json").write_text(json.dumps(result["graph"]["manifest"],indent=2),encoding="utf-8")
        print(json.dumps(result["report"],indent=2))
        return 2 if a.strict and result["report"]["diagnostic_count"] else 0
    except Exception as exc:
        print(f"ERROR: {exc}",file=sys.stderr); return 1
