"""Discover and load VBA conversion artifacts without assuming one project layout."""
from __future__ import annotations
import json, hashlib
from pathlib import Path
from typing import Any

ALIASES = {
    "source": ("*.bas", "*.cls", "*.frm", "*.vba"),
    "graph": ("graph.json", "graph*.json"),
    "segments": ("segments.json", "segments*.json"),
    "ir": ("universal_ir.json", "universal_ir*.json", "ir.json"),
    "normalized_ir": ("normalized_ir.json", "normalized_ir*.json"),
    "semantic_analysis": ("semantic_analysis.json", "semantic_analysis*.json"),
    "manifest": ("manifest.json", "manifest*.json"),
    "pseudocode": ("pseudocode.md", "pseudocode*.md"),
    "pseudocode_run": ("pseudocode_run.json", "pseudocode_run*.json"),
    "report": ("report.json", "report*.json"),
    "logs": ("*.log", "*.txt"),
}
JSON_KINDS = {"graph", "segments", "ir", "normalized_ir", "semantic_analysis",
              "manifest", "pseudocode_run", "report"}
CODE_SUFFIXES={".py",".java",".js",".ts",".tsx",".jsx",".cs",".go",".rs",".kt",".scala",
               ".cpp",".c",".h",".hpp",".sql",".sh",".ps1",".yaml",".yml",".toml",".xml",".vb",
               ".swift",".dart",".rb",".php",".r",".lua",".pl",".hs",".ex",".exs",".erl",
               ".clj",".cljs",".fs",".fsx",".m",".mm",".proto",".graphql",".gql",".vue",".svelte",".bat"}
TEXT_SUFFIXES = CODE_SUFFIXES | {".json", ".jsonl", ".ipynb", ".md", ".mdx", ".txt", ".log", ".ini", ".cfg", ".properties", ".html", ".css", ".csv", ".tsv", ".rst"}
SOURCE_SUFFIXES = {".bas", ".cls", ".frm", ".vba", ".vbs"}
EXCLUDED_DIRS = {".git", ".venv", "venv", "node_modules", "__pycache__", ".idea", ".pytest_cache"}
SENSITIVE_NAMES = {".env", ".env.local", "id_rsa", "id_ed25519", "credentials", "secrets"}
FIXER_REPORT_NAMES = {
    "final_quality_report.json", "intelligent_fixer_log.json",
    "repair_history.json", "unresolved_issues.json",
}

def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()

class ProjectLoader:
    def __init__(self, project_root: str | Path):
        self.root = Path(project_root).expanduser().resolve()
        if not self.root.exists() or not self.root.is_dir():
            raise FileNotFoundError(f"Project root does not exist: {self.root}")

    def _find(self, kind: str) -> list[Path]:
        found=[]
        for pattern in ALIASES[kind]:
            found.extend(p for p in self.root.rglob(pattern)
                         if p.is_file() and not any(x in {".git", ".venv", "venv", "__pycache__"} for x in p.parts))
        # Prefer shallow paths, then newest names lexically; retain alternates for diagnostics.
        return sorted(set(found), key=lambda p: (len(p.relative_to(self.root).parts), str(p)))

    def load(self) -> dict[str, Any]:
        artifacts={}
        for kind in ALIASES:
            matches=self._find(kind)
            if not matches: continue
            if kind=="source":
                loaded=[]
                for source_path in matches:
                    raw_source=source_path.read_bytes()
                    try: source_text=raw_source.decode("utf-8-sig")
                    except UnicodeDecodeError: continue
                    loaded.append({"path":str(source_path.relative_to(self.root)),
                      "absolute_path":str(source_path),"sha256":sha256_bytes(raw_source),
                      "size_bytes":len(raw_source),"text":source_text})
                if loaded:
                    artifacts["sources"]=loaded
                    artifacts["source"]=loaded[0]
                continue
            p=matches[0]
            raw=p.read_bytes()
            item={"path":str(p.relative_to(self.root)), "absolute_path":str(p),
                  "sha256":sha256_bytes(raw), "size_bytes":len(raw), "alternates":[str(x.relative_to(self.root)) for x in matches[1:]]}
            if kind in JSON_KINDS:
                try: item["data"]=json.loads(raw.decode("utf-8-sig"))
                except Exception as e: item["load_error"]=f"{type(e).__name__}: {e}"
            elif kind in {"source","pseudocode","logs"}:
                try: item["text"]=raw.decode("utf-8-sig")
                except UnicodeDecodeError: item["load_error"]="File is not valid UTF-8; configure an explicit source encoding."
            artifacts[kind]=item
        # Graph exports commonly carry the complete source module; use only when hash matches manifest.
        if "source" not in artifacts and "graph" in artifacts and "manifest" in artifacts:
            graph=artifacts["graph"].get("data",{})
            manifest=artifacts["manifest"].get("data",[])
            expected={x.get("path"):x.get("sha256") for x in manifest if isinstance(x,dict)} if isinstance(manifest,list) else {}
            embedded=[]
            for node in graph.get("nodes",[]):
                if node.get("type")=="MODULE" and isinstance(node.get("source"),dict):
                    text=node["source"].get("text"); path=node["source"].get("path","<embedded-in-graph>")
                    if isinstance(text,str):
                        digest=sha256_bytes(text.encode("utf-8"))
                        if path in expected and digest!=expected[path]: continue
                        embedded.append({"path":path,"sha256":digest,"size_bytes":len(text.encode("utf-8")),
                          "text":text,"virtual":True,"origin":"graph MODULE.source.text"})
            if embedded:
                artifacts["sources"]=embedded
                artifacts["source"]=embedded[0]
        inventory=self._code_files()
        output_files=self._output_files()
        return {"root":str(self.root), "artifacts":artifacts,"project_files":inventory,
                "output_files":output_files,
                "missing_kinds":[k for k in ("source","graph","segments","ir","normalized_ir","semantic_analysis") if k not in artifacts]}

    def producer_code_context(self, issue, limit=5, max_chars=12000):
        code=str(issue.get("code","")).lower()
        subsystem=str(issue.get("subsystem","")).lower()
        keywords={
          "frontend_control_flow":["parser","frontend","graph","control_flow"],
          "expression_frontend":["expression","parser","ir","builder"],
          "documentation_generation":["pseudo","generator","documentation"],
          "runtime_semantics":["runtime","error","exception","backend","translator"],
          "control_flow_lowering":["control","lower","backend","translator"],
          "type_resolution":["semantic","type","binding","resolver"],
          "runtime_interop":["interop","com","runtime","backend"],
          "array_semantics":["array","runtime","backend","translator"],
        }.get(subsystem,["parser","builder","converter","translator"])
        candidates=[]
        for item in self._code_files():
            name=Path(item["path"]).name.lower()
            score=sum(2 for word in keywords if word in name)
            if score: candidates.append((score,item))
        out=[]
        for score,item in sorted(candidates,key=lambda x:(-x[0],x[1]["path"]))[:limit]:
            path=self.resolve(item["path"])
            try: text=path.read_text(encoding="utf-8")
            except Exception: continue
            out.append({"path":item["path"],"sha256":item["sha256"],
                        "excerpt":text[:max_chars],"truncated":len(text)>max_chars})
        return out

    def _iter_project_files(self):
        for fp in self.root.rglob("*"):
            if not fp.is_file():
                continue
            rel=fp.relative_to(self.root)
            if any(part in EXCLUDED_DIRS or part == "intelligent_fixer" for part in rel.parts):
                continue
            if fp.name.lower() in SENSITIVE_NAMES or fp.name.lower().startswith(".env"):
                continue
            yield fp

    def _code_files(self):
        out=[]
        for fp in self._iter_project_files():
            if fp.suffix.lower() not in CODE_SUFFIXES:
                continue
            try:
                raw=fp.read_bytes()
                out.append({"path":str(fp.relative_to(self.root)),"sha256":sha256_bytes(raw),
                            "size_bytes":len(raw),"suffix":fp.suffix.lower()})
            except OSError:
                pass
        return out

    def _output_files(self, max_file_bytes=8_000_000):
        """Inventory readable generated artifacts and implementation files; never include VBA source."""
        out=[]
        for fp in self._iter_project_files():
            suffix=fp.suffix.lower()
            if suffix not in TEXT_SUFFIXES or suffix in SOURCE_SUFFIXES:
                continue
            if fp.name.lower() in FIXER_REPORT_NAMES:
                continue
            try:
                raw=fp.read_bytes()
                if len(raw)>max_file_bytes:
                    out.append({"path":str(fp.relative_to(self.root)),"sha256":sha256_bytes(raw),
                                "size_bytes":len(raw),"suffix":suffix,"readable":False,
                                "reason":"File exceeds configured audit size limit"})
                    continue
                text=raw.decode("utf-8-sig")
                rel=str(fp.relative_to(self.root))
                read_only = suffix == ".log" or fp.name.lower().startswith("manifest")
                out.append({"path":rel,"sha256":sha256_bytes(raw),
                            "size_bytes":len(raw),"suffix":suffix,"readable":True,
                            "read_only":read_only,"text":text})
            except (OSError, UnicodeDecodeError):
                continue
        return sorted(out,key=lambda x:x["path"])

    def resolve(self, relative_path: str) -> Path:
        candidate=(self.root / relative_path).resolve()
        if candidate != self.root and self.root not in candidate.parents:
            raise ValueError(f"Path escapes project root: {relative_path}")
        return candidate
