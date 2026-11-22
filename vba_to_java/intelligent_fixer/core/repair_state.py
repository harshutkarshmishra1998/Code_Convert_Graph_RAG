"""Persistent, append-only repair state with deterministic issue identities."""
from __future__ import annotations
import hashlib, json
from pathlib import Path
from datetime import datetime, timezone

def issue_id(issue):
    stable={k:issue.get(k) for k in ("code","path","line","message","artifact","pointer")}
    return hashlib.sha256(json.dumps(stable,sort_keys=True,default=str).encode()).hexdigest()[:16]

class RepairState:
    def __init__(self, reports_dir):
        self.dir=Path(reports_dir); self.dir.mkdir(parents=True,exist_ok=True)
        self.history=self.dir/"repair_history.json"
        self.unresolved=self.dir/"unresolved_issues.json"
        self.quality=self.dir/"final_quality_report.json"
        self.issue_log=self.dir/"intelligent_fixer_log.json"

    def read(self,path,default):
        try: return json.loads(path.read_text(encoding="utf-8"))
        except (FileNotFoundError,json.JSONDecodeError): return default

    def append(self, record):
        self.append_many([record])

    def append_many(self, records):
        if not records:
            return
        data=self.read(self.history,{"schema_version":"1.0","records":[]})
        now=datetime.now(timezone.utc).isoformat()
        data.setdefault("records",[]).extend(
            {**record,"timestamp":now} for record in records
        )
        self._atomic(self.history,data)

    def write_issue_run(self, run):
        data=self.read(self.issue_log,{"schema_version":"1.0","runs":[]})
        data.setdefault("runs",[]).append({
            **run,"timestamp":datetime.now(timezone.utc).isoformat()
        })
        self._atomic(self.issue_log,data)

    def write_unresolved(self, issues):
        self._atomic(self.unresolved,{"schema_version":"1.0","issues":issues})

    def write_quality(self, report):
        self._atomic(self.quality,report)

    @staticmethod
    def _atomic(path,data):
        tmp=path.with_suffix(path.suffix+".tmp")
        tmp.write_text(json.dumps(data,indent=2,ensure_ascii=False,default=str)+"\n",encoding="utf-8")
        tmp.replace(path)
