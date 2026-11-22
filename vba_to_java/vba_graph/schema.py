def validate_graph(graph):
    errors=[]
    for key in ("schema_version","language","project","nodes","edges","diagnostics","manifest"):
        if key not in graph: errors.append(f"Missing graph field: {key}")
    ids=[n.get("id") for n in graph.get("nodes",[])]
    if len(ids)!=len(set(ids)): errors.append("Duplicate node IDs")
    known=set(ids)
    for i,e in enumerate(graph.get("edges",[])):
        if e.get("source") not in known: errors.append(f"Edge {i} has unknown source")
        if e.get("target") not in known: errors.append(f"Edge {i} has unknown target")
        if not e.get("relation"): errors.append(f"Edge {i} missing relation")
    return errors
