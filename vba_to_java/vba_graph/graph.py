from .ids import stable_id
from .semantic import walk

def build_graph(project_name, modules, diagnostics, manifest, annotations=None):
    nodes, edges, segments = [], [], []
    def add_node(node):
        nodes.append(node)
    def edge(source,target,relation,evidence=None):
        edges.append({"source":source,"target":target,"relation":relation,"evidence":evidence or {}})

    project_id=stable_id("project",project_name,1,1,project_name)
    add_node({"id":project_id,"type":"PROJECT","name":project_name,"parent_id":None,
              "source":None,"metadata":{}})
    for mod in modules:
        path=mod["path"]
        mod_start=1
        mod_end=max(1,len(mod["raw_source"].splitlines()))
        mid=stable_id("module",path,mod_start,mod_end,mod["name"])
        add_node({"id":mid,"type":"MODULE","name":mod["name"],"parent_id":project_id,
                  "source":{"path":path,"start_line":1,"end_line":mod_end,"text":mod["raw_source"]},
                  "metadata":{"kind":mod["kind"],"options":mod["options"],"attributes":mod["attributes"]}})
        edge(project_id,mid,"CONTAINS")
        for i,decl in enumerate(mod["declarations"]):
            loc=decl.get("location",{})
            start=loc.get("start_line",1); end=loc.get("end_line",start)
            did=stable_id("decl",path,start,end,str(i))
            add_node({"id":did,"type":decl["type"],"name":f"module_declaration_{i+1}",
                      "parent_id":mid,"source":{"path":path,"start_line":start,"end_line":end,
                      "text":decl.get("source","")},"metadata":{}})
            edge(mid,did,"CONTAINS")
        for proc in mod["procedures"]:
            loc=proc["location"]; start=loc["start_line"]
            end=proc.get("end_location",{}).get("end_line",max(
                [n.get("location",{}).get("end_line",start) for n in walk(proc["body"])] or [start]))
            pid=stable_id("procedure",path,start,end,proc["name"])
            add_node({"id":pid,"type":"PROCEDURE","name":proc["name"],"parent_id":mid,
                      "source":{"path":path,"start_line":start,"end_line":end,
                                "text":_slice(mod["raw_source"],start,end)},
                      "metadata":{"kind":proc["kind"],"visibility":proc["visibility"],
                                  "parameters":proc["parameters"],"return_type":proc["return_type"]}})
            edge(mid,pid,"CONTAINS")
            seg_id=stable_id("segment",path,start,end,proc["name"])
            segments.append({"id":seg_id,"parent_id":pid,"kind":"PROCEDURE",
                             "name":proc["name"],"path":path,"start_line":start,"end_line":end,
                             "source":_slice(mod["raw_source"],start,end),
                             "summary":None,"annotations":{}})
            _add_body_nodes(proc["body"],pid,path,mod["raw_source"],nodes,edges,segments)

    graph={"schema_version":"1.0.0","language":"vba","project":{"name":project_name,"node_id":project_id},
           "nodes":nodes,"edges":edges,"symbols":[],"diagnostics":diagnostics,
           "manifest":manifest}
    return graph,segments

def _slice(source,start,end):
    lines=source.splitlines()
    return "\n".join(lines[max(0,start-1):max(0,end)])

def _add_body_nodes(items,parent,path,source,nodes,edges,segments):
    previous=None
    for index,item in enumerate(items):
        loc=item.get("location",{})
        start=loc.get("start_line",1)
        end=item.get("end_location",{}).get("end_line",loc.get("end_line",start))
        typ=item.get("type","OPAQUE")
        name=item.get("name") or item.get("target") or f"{typ.lower()}_{index+1}"
        nid=stable_id("node",path,start,end,f"{parent}:{typ}:{index}:{name}")
        raw=item.get("source","")
        node={"id":nid,"type":typ,"name":str(name),"parent_id":parent,
              "source":{"path":path,"start_line":start,"end_line":end,
                        "text":_slice(source,start,end)},
              "metadata":{k:v for k,v in item.items() if k not in {
                  "source","location","end_location","body","branches","else_body","cases"
              }}}
        nodes.append(node); edges.append({"source":parent,"target":nid,"relation":"CONTAINS","evidence":{}})
        if previous:
            edges.append({"source":previous,"target":nid,"relation":"NEXT","evidence":{"order":"source"}})
        previous=nid
        segments.append({"id":stable_id("segment",path,start,end,nid),"parent_id":nid,
                         "kind":typ,"name":str(name),"path":path,"start_line":start,"end_line":end,
                         "source":_slice(source,start,end),"summary":None,"annotations":{}})
        if typ=="IF_BLOCK":
            for bidx,branch in enumerate(item.get("branches",[])):
                branch_id=stable_id("branch",path,start,end,f"{nid}:{bidx}")
                branch_type=branch.get("branch","IF")
                nodes.append({"id":branch_id,"type":f"{branch_type}_BRANCH","name":branch_type,
                              "parent_id":nid,"source":{"path":path,"start_line":start,"end_line":end,
                              "text":""},"metadata":{"condition":branch.get("condition")}})
                edges.append({"source":nid,"target":branch_id,"relation":"TRUE_BRANCH" if bidx==0 else "ELSEIF_BRANCH","evidence":{}})
                _add_body_nodes(branch.get("body",[]),branch_id,path,source,nodes,edges,segments)
            if item.get("else_body") is not None:
                branch_id=stable_id("branch",path,start,end,f"{nid}:else")
                nodes.append({"id":branch_id,"type":"ELSE_BRANCH","name":"ELSE","parent_id":nid,
                              "source":{"path":path,"start_line":start,"end_line":end,"text":""},"metadata":{}})
                edges.append({"source":nid,"target":branch_id,"relation":"FALSE_BRANCH","evidence":{}})
                _add_body_nodes(item["else_body"],branch_id,path,source,nodes,edges,segments)
        elif typ=="SELECT_BLOCK":
            for cidx,case in enumerate(item.get("cases",[])):
                cid=stable_id("case",path,start,end,f"{nid}:{cidx}")
                nodes.append({"id":cid,"type":"CASE_BRANCH","name":case.get("condition_source",f"Case {cidx+1}"),
                              "parent_id":nid,"source":{"path":path,"start_line":start,"end_line":end,"text":""},
                              "metadata":{"condition_source":case.get("condition_source")}})
                edges.append({"source":nid,"target":cid,"relation":"CASE_BRANCH","evidence":{}})
                _add_body_nodes(case.get("body",[]),cid,path,source,nodes,edges,segments)
        else:
            child=item.get("body")
            if isinstance(child,list) and child:
                _add_body_nodes(child,nid,path,source,nodes,edges,segments)
