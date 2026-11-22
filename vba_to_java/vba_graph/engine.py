from pathlib import Path
import hashlib,re
from .source import read_project, EXTENSIONS
from .parser import parse_module
from .graph import build_graph
from .semantic import analyze
from .schema import validate_graph

def convert_sources(sources, project_name="VBA_Project", manifest=None, annotation_payload=None):
    modules=[]; diagnostics=[]
    for path,source in sources.items():
        ext=Path(path).suffix.lower()
        kind={".bas":"STANDARD_MODULE",".cls":"CLASS_MODULE",".frm":"USERFORM",".vba":"UNKNOWN",".txt":"UNKNOWN"}.get(ext,"UNKNOWN")
        name=re.sub(r"\W+","_",Path(path).stem) or "Module"
        mod,diags=parse_module(source,path,name,kind)
        mod["sha256"]=hashlib.sha256(source.encode("utf-8",errors="replace")).hexdigest()
        modules.append(mod); diagnostics.extend(diags)
    graph,segments=build_graph(project_name,modules,diagnostics,manifest or [])
    symbols,semantic_edges,semantic_diags=analyze(modules)
    graph["symbols"]=symbols
    # Resolve semantic procedure names to the stable structural node IDs.
    procedure_nodes = {n["name"].casefold(): [] for n in graph["nodes"] if n["type"] == "PROCEDURE"}
    for n in graph["nodes"]:
        if n["type"] == "PROCEDURE":
            procedure_nodes.setdefault(n["name"].casefold(), []).append(n["id"])
    qualified_to_id = {}
    for mod in modules:
        for proc in mod["procedures"]:
            qualified = f'{mod["name"]}.{proc["name"]}'
            match = next((n for n in graph["nodes"] if n["type"] == "PROCEDURE"
                          and n["name"] == proc["name"]
                          and n["source"]["path"] == mod["path"]), None)
            if match:
                qualified_to_id[qualified] = match["id"]
    for e in semantic_edges:
        src = qualified_to_id.get(e["source"])
        tgt = qualified_to_id.get(e["target"])
        if src and tgt:
            graph["edges"].append({"source":src,"target":tgt,"relation":e["relation"],
                                   "evidence":e.get("evidence",{})})
    graph["diagnostics"].extend(semantic_diags)
    if annotation_payload is not None:
        from .enrichment import apply_annotations
        graph["llm_enrichment"]=apply_annotations(graph,segments,annotation_payload)
    errors=validate_graph(graph)
    if errors: raise ValueError("Invalid graph: "+"; ".join(errors))
    report={"project":project_name,"module_count":len(modules),
            "procedure_count":sum(len(m["procedures"]) for m in modules),
            "node_count":len(graph["nodes"]),"edge_count":len(graph["edges"]),
            "segment_count":len(segments),"diagnostic_count":len(graph["diagnostics"]),
            "opaque_count":sum(1 for n in graph["nodes"] if n["type"]=="OPAQUE_STATEMENT")}
    return {"graph":graph,"segments":segments,"report":report}

def convert_path(path, annotation_payload=None):
    sources,manifest=read_project(path)
    result=convert_sources(sources,Path(path).stem,manifest,annotation_payload)
    return result
