"""Extract structured diagnostics from JSON-like reports and plain text logs."""
import re
class LogParser:
    PATTERNS=[
      (re.compile(r"(?P<code>[A-Z][A-Z0-9_]{3,})"),"diagnostic"),
      (re.compile(r"(?:line|ln)\s*[:#]?\s*(?P<line>\d+)",re.I),"location"),
      (re.compile(r"(?P<path>[\w./\\-]+\.(?:bas|cls|frm|json|py|md))",re.I),"path"),
    ]
    def parse(self,text,path="<log>"):
        out=[]
        for n,line in enumerate(text.splitlines(),1):
            codes=re.findall(r"\b[A-Z][A-Z0-9_]{3,}\b",line)
            if not codes and not re.search(r"\b(error|warning|failed|unresolved|opaque|unsupported|exception)\b",line,re.I): continue
            m=re.search(r"\b(?:line|ln)\s*[:#]?\s*(\d+)",line,re.I)
            out.append({"code":codes[0] if codes else "UNCLASSIFIED_LOG_MESSAGE",
                "severity":"error" if re.search(r"error|failed|exception",line,re.I) else "warning",
                "message":line.strip(),"path":path,"log_line":n,
                "line":int(m.group(1)) if m else None,"artifact":"logs"})
        return out

    def from_artifact(self,kind,data,path):
        out=[]
        if kind=="graph":
            for d in data.get("diagnostics",[]):
                out.append({"code":d.get("code","GRAPH_DIAGNOSTIC"),"severity":d.get("severity","warning"),
                  "message":d.get("message",""),"path":d.get("path") or path,"line":d.get("line"),"artifact":"graph"})
        elif kind=="ir":
            for d in data.get("diagnostics",[]):
                loc=d.get("location",{})
                out.append({"code":d.get("code","IR_DIAGNOSTIC"),"severity":d.get("severity","warning"),
                  "message":d.get("message",""),"path":loc.get("file",path),"line":loc.get("line"),"artifact":"ir"})
        elif kind=="semantic_analysis":
            for d in data.get("upstream_diagnostics",[]):
                loc=d.get("location",{})
                out.append({"code":d.get("code","SEMANTIC_DIAGNOSTIC"),"severity":d.get("severity","warning"),
                  "message":d.get("message",""),"path":loc.get("file",path),"line":loc.get("line"),"artifact":"semantic_analysis"})
            for r in data.get("compatibility_requirements",[]):
                out.append({"code":r.get("code","COMPATIBILITY_REQUIREMENT"),"severity":r.get("severity","review"),
                  "message":r.get("title",""),"details":r,"path":path,"artifact":"semantic_analysis"})
        elif kind=="pseudocode_run":
            for d in data.get("diagnostics",[]):
                out.append({"code":d.get("code","PSEUDOCODE_DIAGNOSTIC"),"severity":d.get("severity","warning"),
                  "message":d.get("message",""),"path":d.get("path",path),"line":d.get("line"),"artifact":"pseudocode_run"})
            for name in data.get("failed_sections",[]):
                out.append({"code":"PSEUDOCODE_SECTION_FAILED","severity":"error","message":str(name),"path":path,"artifact":"pseudocode_run"})
        elif kind=="report" and data.get("diagnostic_count",0):
            out.append({"code":"REPORT_DIAGNOSTICS_PRESENT","severity":"warning",
                "message":f"{data.get('diagnostic_count')} diagnostics in graph report","path":path,"artifact":"report"})
        return out
