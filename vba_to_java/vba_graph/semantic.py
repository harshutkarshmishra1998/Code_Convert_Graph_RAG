"""Best-effort semantic links over source-preserving structural nodes."""
import re

BUILTINS = {
    "if","then","else","elseif","for","each","in","to","step","next","do","loop",
    "while","wend","select","case","with","end","sub","function","property",
    "dim","as","set","let","call","public","private","friend","static","const",
    "byval","byref","optional","paramarray","true","false","nothing","null","empty",
    "and","or","not","xor","eqv","imp","mod","is","like","on","error","resume",
    "goto","gosub","exit","return","new","redim","preserve","erase","me"
}

def walk(items):
    for item in items:
        yield item
        for key in ("body","else_body"):
            if isinstance(item.get(key),list):
                yield from walk(item[key])
        for branch in item.get("branches",[]):
            yield from walk(branch.get("body",[]))
        for case in item.get("cases",[]):
            yield from walk(case.get("body",[]))

def analyze(modules):
    proc_lookup = {}
    for mod in modules:
        for proc in mod["procedures"]:
            proc_lookup.setdefault(proc["name"].casefold(),[]).append(
                f'{mod["name"]}.{proc["name"]}'
            )
    edges, symbols, diagnostics = [], [], []
    for mod in modules:
        module_vars = set()
        for decl in mod["declarations"]:
            src = decl.get("source","")
            m = re.match(r"^(?:Public|Private|Global|Dim|Const|Static)\s+(.+)$",src,re.I)
            if m:
                for part in m.group(1).split(","):
                    name = re.match(r"\s*([A-Za-z_]\w*)",part)
                    if name: module_vars.add(name.group(1).casefold())
        for proc in mod["procedures"]:
            proc_id = f'{mod["name"]}.{proc["name"]}'
            local_vars = {p["name"].casefold() for p in proc["parameters"] if "name" in p}
            for node in walk(proc["body"]):
                if node.get("type") == "DECLARATION":
                    m = re.match(r"^(?:Dim|Static|Const|ReDim)\s+(.+)$",node.get("source",""),re.I)
                    if m:
                        for part in m.group(1).split(","):
                            vm = re.match(r"\s*([A-Za-z_]\w*)",part)
                            if vm: local_vars.add(vm.group(1).casefold())
            for p in proc["parameters"]:
                symbols.append({"id":f"{proc_id}.param.{p.get('name','unknown')}",
                                "name":p.get("name"),"kind":"PARAMETER","scope":proc_id,
                                "data_type":p.get("type","Unknown"),"passing":p.get("passing")})
            for var in sorted(local_vars):
                symbols.append({"id":f"{proc_id}.var.{var}","name":var,"kind":"LOCAL_OR_PARAMETER",
                                "scope":proc_id})
            for node in walk(proc["body"]):
                src = node.get("source","")
                if node.get("type") in {"CALL_OR_COMMAND","HOST_CALL"}:
                    # Conservative call target extraction; excludes obvious keywords.
                    match = re.match(r"^(?:Call\s+)?([A-Za-z_]\w*)",src,re.I)
                    if match:
                        name = match.group(1)
                        for target in proc_lookup.get(name.casefold(),[]):
                            edges.append({"source":proc_id,"target":target,"relation":"CALLS",
                                          "evidence":{"path":mod["path"],"line":node.get("location",{}).get("start_line")}})
                # References are lexical evidence, not proof of a semantic read/write.
                identifiers = re.findall(r"\b[A-Za-z_]\w*\b",src)
                for name in identifiers:
                    key=name.casefold()
                    if key in BUILTINS or key in local_vars or key in module_vars:
                        continue
                    if key in proc_lookup and node.get("type") not in {"CALL_OR_COMMAND"}:
                        edges.append({"source":proc_id,"target":proc_lookup[key][0],"relation":"REFERENCES",
                                      "evidence":{"path":mod["path"],"line":node.get("location",{}).get("start_line")}})
    # Stable de-duplication
    unique, seen = [], set()
    for edge in edges:
        key=(edge["source"],edge["target"],edge["relation"])
        if key not in seen: seen.add(key); unique.append(edge)
    return symbols, unique, diagnostics
