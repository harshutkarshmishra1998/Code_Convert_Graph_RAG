"""Build procedure context from graph node source ranges and IR procedure metadata."""
import re
class ProcedureContext:
    def build(self,evidence,artifacts):
        text=evidence.get("text","")
        names=[]
        for m in re.finditer(r"(?im)^\s*(?:(?:Public|Private|Friend|Static)\s+)*(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)",text):
            names.append(m.group(1))
        graph=artifacts.get("graph",{}).get("data",{})
        procs=[n for n in graph.get("nodes",[]) if n.get("type") in {"PROCEDURE","FUNCTION","SUB"}]
        selected=[]
        for n in procs:
            if not names or n.get("name","").lower() in {x.lower() for x in names}:
                selected.append({"id":n.get("id"),"name":n.get("name"),"source":n.get("source")})
        return {"procedure_names":names,"graph_procedures":selected[:20]}
