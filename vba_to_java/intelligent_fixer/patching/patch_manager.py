"""Apply only narrow, preconditioned edits to existing non-source artifacts."""
from pathlib import Path
import hashlib
import ast, json, xml.etree.ElementTree as ET
from .file_editor import FileEditor
from .backup_manager import BackupManager
from .rollback_manager import RollbackManager

class PatchManager:
    FORBIDDEN={"*.bas","*.cls","*.frm","*.vba"}
    ALLOWED_SUFFIXES={".json",".md",".txt",".log",".py",".java",".js",".ts",".tsx",".jsx",
                      ".cs",".go",".rs",".kt",".scala",".cpp",".c",".h",".hpp",".sql",
                      ".sh",".ps1",".yaml",".yml",".toml",".xml",".html",".css",".ini",".cfg",".properties",".vb",".csv",".tsv",".ipynb",".rst"}
    def __init__(self,project_root):
        self.root=Path(project_root).resolve(); self.editor=FileEditor()
        self.backups=BackupManager(); self.rollbacker=RollbackManager()

    @staticmethod
    def _validate_candidate(path, data):
        text=data.decode("utf-8-sig")
        suffix=path.suffix.lower()
        if suffix in {".json", ".ipynb"}:
            json.loads(text)
        elif suffix == ".jsonl":
            for line in text.splitlines():
                if line.strip(): json.loads(line)
        elif suffix == ".py":
            ast.parse(text, filename=str(path))
        elif suffix == ".toml":
            try:
                import tomllib
                tomllib.loads(text)
            except ImportError:
                pass
        elif suffix == ".xml":
            ET.fromstring(text)

    def apply(self,proposal):
        rel=proposal.get("relative_path")
        if not rel: return {"applied":False,"error":"relative_path is required"}
        path=(self.root/rel).resolve()
        if self.root not in path.parents: return {"applied":False,"error":"Path escapes project root"}
        if not path.is_file(): return {"applied":False,"error":"Target must be an existing file"}
        if path.suffix.lower() not in self.ALLOWED_SUFFIXES: return {"applied":False,"error":"Unsupported artifact type"}
        if path.suffix.lower() in {".bas",".cls",".frm",".vba",".vbs"}: return {"applied":False,"error":"Original VBA source is immutable"}
        if path.suffix.lower()==".log" or path.name.lower().startswith("manifest"):
            return {"applied":False,"error":"Diagnostic logs and source manifests are read-only evidence"}
        old=path.read_bytes(); expected=proposal.get("expected_sha256")
        if not expected: return {"applied":False,"error":"expected_sha256 is mandatory"}
        if hashlib.sha256(old).hexdigest()!=expected:
            return {"applied":False,"error":"File changed since proposal was created"}
        op=proposal.get("operation")
        try:
            if op=="json_set": new=self.editor.json_set(old,proposal["pointer"],proposal.get("old_value"),proposal.get("new_value"))
            elif op=="json_replace": new=self.editor.json_replace(old,proposal["pointer"],proposal.get("old_value"),proposal.get("new_value"))
            elif op=="json_add": new=self.editor.json_add(old,proposal["pointer"],proposal.get("new_value"))
            elif op=="json_remove": new=self.editor.json_remove(old,proposal["pointer"],proposal.get("old_value"))
            elif op=="text_replace": new=self.editor.text_replace(old,proposal["old_text"],proposal["new_text"])
            elif op=="file_replace": new=self.editor.file_replace(old,proposal.get("new_content"),path.suffix)
            else: return {"applied":False,"error":f"Unsupported operation: {op}"}
            self._validate_candidate(path, new)
            backup=self.backups.backup(path)
            try:
                tmp=path.with_name(path.name+".intelligent_fixer.tmp")
                tmp.write_bytes(new); tmp.replace(path)
                return {"applied":True,"path":str(path.relative_to(self.root)),
                    "backup":str(backup.relative_to(self.root)),
                    "before_sha256":hashlib.sha256(old).hexdigest(),
                    "after_sha256":hashlib.sha256(new).hexdigest(),"operation":op}
            except Exception:
                self.rollbacker.rollback(path,backup); raise
        except Exception as e:
            return {"applied":False,"error":f"{type(e).__name__}: {e}"}
