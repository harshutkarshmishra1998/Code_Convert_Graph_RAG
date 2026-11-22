"""Retrieve source excerpts by line or identifier from source files or graph snapshots."""
import re
class VBARetriever:
    def retrieve(self,issue,artifacts,radius=12):
        candidates=artifacts.get("sources") or ([artifacts["source"]] if artifacts.get("source") else [])
        requested=issue.get("path")
        selected=next((x for x in candidates if requested and (x.get("path")==requested or x.get("path","").endswith("/"+str(requested)))),None)
        path_was_matched=selected is not None
        selected=selected or artifacts.get("source",{})
        src=selected.get("text"); source_path=selected.get("path")
        if not src: return {"available":False,"reason":"Original source unavailable"}
        lines=src.splitlines()
        line=issue.get("line")
        if line and 1 <= int(line) <= len(lines):
            start=max(1,int(line)-radius); end=min(len(lines),int(line)+radius)
            return {"available":True,"path":source_path,"line_range":[start,end],
                    "text":"\n".join(f"{i}: {lines[i-1]}" for i in range(start,end+1)),
                    "sha256":selected.get("sha256"),"exact_line":int(line)}
        details=issue.get("details",{}) if isinstance(issue.get("details"),dict) else {}
        ev=details.get("evidence")
        evidence_items=ev if isinstance(ev,list) else [ev] if isinstance(ev,dict) else []
        evidence_items=[x for x in evidence_items if isinstance(x,dict)]
        proc_name=issue.get("procedure") or next((x.get("procedure") for x in evidence_items if x.get("procedure")),None)
        if proc_name and not path_was_matched:
            for candidate in candidates:
                candidate_lines=candidate.get("text","").splitlines()
                if any(re.search(rf"(?i)^\s*(?:(?:Public|Private|Friend|Static)\s+)*(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+{re.escape(str(proc_name))}\b",line_text) for line_text in candidate_lines):
                    selected=candidate; src=selected.get("text"); source_path=selected.get("path"); lines=src.splitlines()
                    break
        if proc_name:
            starts=[i for i,line_text in enumerate(lines,1) if re.search(
                rf"(?i)^\s*(?:(?:Public|Private|Friend|Static)\s+)*(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+{re.escape(str(proc_name))}\b",line_text)]
            if starts:
                start=starts[0]
                next_proc=next((i for i in range(start+1,len(lines)+1) if re.search(
                    r"(?i)^\s*(?:(?:Public|Private|Friend|Static)\s+)*(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+[A-Za-z_]\w*",lines[i-1])),len(lines)+1)
                end=min(next_proc-1,start+500)
                return {"available":True,"path":source_path,"line_range":[start,end],
                    "text":"\n".join(f"{j}: {lines[j-1]}" for j in range(start,end+1)),
                    "sha256":selected.get("sha256"),"procedure":proc_name}
        needle=issue.get("symbol") or issue.get("procedure") or issue.get("message","")
        # Prefer exact source phrases embedded in compatibility evidence.
        phrases=[x for x in ev if isinstance(x,str)] if isinstance(ev,list) else ([ev] if isinstance(ev,str) else [])
        for phrase in phrases:
            if phrase and str(phrase).strip().lower() in src.lower():
                i=next((n for n,line_text in enumerate(lines,1) if str(phrase).strip().lower() in line_text.lower()),1)
                start=max(1,i-radius); end=min(len(lines),i+radius)
                return {"available":True,"path":source_path,"line_range":[start,end],
                    "text":"\n".join(f"{j}: {lines[j-1]}" for j in range(start,end+1)),
                    "sha256":selected.get("sha256"),"matched_phrase":phrase}
        ident=re.search(r"\b[A-Za-z_]\w*\b",str(needle))
        if ident:
            hits=[i for i,s in enumerate(lines,1) if re.search(rf"\b{re.escape(ident.group())}\b",s,re.I)]
            if hits:
                i=hits[0]; start=max(1,i-radius); end=min(len(lines),i+radius)
                return {"available":True,"path":source_path,"line_range":[start,end],
                    "text":"\n".join(f"{j}: {lines[j-1]}" for j in range(start,end+1)),
                    "sha256":selected.get("sha256"),"matching_lines":hits[:30]}
        return {"available":True,"path":source_path,"line_range":None,"text":src[:12000],
                "sha256":selected.get("sha256"),"note":"No precise line supplied; bounded source prefix returned."}
