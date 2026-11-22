"""Optional LLM annotation adapter. Structural graph/source remain authoritative."""
import json

def apply_annotations(graph, segments, payload):
    if isinstance(payload, str):
        payload = json.loads(payload)
    annotations = payload.get("annotations", [])
    segment_index = {s["id"]:s for s in segments}
    node_ids = {n["id"] for n in graph["nodes"]}
    applied, rejected = 0, []
    for item in annotations:
        target = item.get("target_id")
        if target not in segment_index and target not in node_ids:
            rejected.append({"annotation":item,"reason":"unknown_target_id"})
            continue
        safe = {k:item[k] for k in ("summary","purpose","confidence","tags") if k in item}
        if target in segment_index:
            segment_index[target]["annotations"].update(safe)
            if "summary" in safe: segment_index[target]["summary"] = safe["summary"]
        else:
            node = next(n for n in graph["nodes"] if n["id"]==target)
            node["metadata"].setdefault("llm_annotations",{}).update(safe)
        applied += 1
    return {"applied":applied,"rejected":rejected}
