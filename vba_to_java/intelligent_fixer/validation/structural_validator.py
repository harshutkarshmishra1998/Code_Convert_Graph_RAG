"""Deterministic structural and syntax checks over discovered conversion artifacts."""
import ast, json, xml.etree.ElementTree as ET
class StructuralValidator:
    def validate(self,project):
        a=project["artifacts"]; issues=[]; errors=0
        def add(code,msg,severity="error"):
            nonlocal errors
            issues.append({"code":code,"severity":severity,"message":msg})
            if severity=="error": errors+=1
        graph=a.get("graph",{}).get("data",{})
        ids=set(); segids=set()
        if graph:
            nodes=graph.get("nodes",[]); ids={n.get("id") for n in nodes if n.get("id")}
            if len(ids)!=len(nodes): add("DUPLICATE_GRAPH_NODE_ID","Graph node IDs are missing or duplicated")
            for e in graph.get("edges",[]):
                s=e.get("source") or e.get("from") or e.get("source_id")
                t=e.get("target") or e.get("to") or e.get("target_id")
                if s and s not in ids: add("DANGLING_GRAPH_EDGE_SOURCE",str(s))
                if t and t not in ids: add("DANGLING_GRAPH_EDGE_TARGET",str(t))
        segments=a.get("segments",{}).get("data",[])
        if isinstance(segments,list):
            segids={x.get("id") for x in segments if isinstance(x,dict)}
            if len(segids)!=len(segments): add("DUPLICATE_SEGMENT_ID","Segment IDs are missing or duplicated")
        ir=a.get("normalized_ir",{}).get("data") or a.get("ir",{}).get("data",{})
        if ir and not isinstance(ir.get("modules"),list): add("INVALID_IR_MODULES","IR modules must be an array")
        # Validate every explicit provenance reference; opaque/source-backed text remains valid evidence.
        missing_graph=set(); missing_segments=set()
        def walk(value):
            if isinstance(value,dict):
                prov=value.get("provenance")
                if isinstance(prov,dict):
                    missing_graph.update(set(prov.get("graph_node_ids",[]))-ids)
                    missing_segments.update(set(prov.get("segment_ids",[]))-segids)
                for child in value.values(): walk(child)
            elif isinstance(value,list):
                for child in value: walk(child)
        if ir: walk(ir)
        if missing_graph: add("IR_DANGLING_GRAPH_PROVENANCE",f"{len(missing_graph)} missing graph node references")
        if missing_segments: add("IR_DANGLING_SEGMENT_PROVENANCE",f"{len(missing_segments)} missing segment references")
        sources=a.get("sources") or ([a["source"]] if a.get("source") else [])
        source_by_path={x.get("path"):x.get("text","") for x in sources}
        segment_mismatches=[]
        for seg in segments if isinstance(segments,list) else []:
            source_text=source_by_path.get(seg.get("path"))
            start,end=seg.get("start_line"),seg.get("end_line")
            if source_text is None or not isinstance(start,int) or not isinstance(end,int): continue
            source_lines=source_text.splitlines()
            if 1<=start<=end<=len(source_lines):
                expected="\n".join(source_lines[start-1:end]).strip()
                actual=str(seg.get("source","")).strip()
                if expected!=actual: segment_mismatches.append(seg.get("id"))
        if segment_mismatches: add("SEGMENT_SOURCE_MISMATCH",f"{len(segment_mismatches)} segments differ from original source line ranges")
        source=a.get("source",{}); manifest=a.get("manifest",{}).get("data",[])
        if source and isinstance(manifest,list) and manifest:
            expected=manifest[0].get("sha256")
            if expected and source.get("sha256")!=expected: add("SOURCE_HASH_MISMATCH","Original source hash differs from manifest")
        syntax_checks=[]
        for item in project.get("output_files", []):
            if not item.get("readable") or not item.get("text"):
                continue
            path=item.get("path", ""); suffix=path.rsplit(".",1)[-1].lower() if "." in path else ""
            try:
                if suffix in {"json", "ipynb"}: json.loads(item["text"])
                elif suffix == "jsonl":
                    for line_no, line in enumerate(item["text"].splitlines(), 1):
                        if line.strip(): json.loads(line)
                elif suffix == "py": ast.parse(item["text"], filename=path)
                elif suffix == "xml": ET.fromstring(item["text"])
                elif suffix == "toml":
                    try:
                        import tomllib
                        tomllib.loads(item["text"])
                    except ImportError:
                        pass
                syntax_checks.append({"path":path,"status":"passed"})
            except Exception as exc:
                add("OUTPUT_SYNTAX_INVALID",f"{path}: {type(exc).__name__}: {exc}")
                syntax_checks.append({"path":path,"status":"failed","error":str(exc)})
        return {"errors":errors,"issues":issues,"checks":{"graph_nodes":len(graph.get("nodes",[])),
          "graph_edges":len(graph.get("edges",[])),"segments":len(segments) if isinstance(segments,list) else 0,
          "source_present":bool(source.get("text")),"output_syntax_checks":syntax_checks}}
