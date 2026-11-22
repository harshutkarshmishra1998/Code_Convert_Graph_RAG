"""Compile gate with an explicitly injected, trusted build runner."""
from pathlib import Path
class CompileValidator:
    def __init__(self,runner=None): self.runner=runner
    def validate(self,project_root):
        root=Path(project_root)
        candidates=[]
        for pattern in ("pom.xml","build.gradle","build.gradle.kts","package.json","pyproject.toml","Cargo.toml","go.mod","*.sln","*.csproj"):
            candidates.extend(str(p.relative_to(root)) for p in root.rglob(pattern)
                              if p.is_file() and not any(x in {".git",".venv","node_modules","venv"} for x in p.parts))
        if not self.runner:
            return {"status":"not_run","build_manifests":sorted(set(candidates)),
              "reason":"No trusted compile runner injected; build scripts are never executed automatically."}
        try:
            result=self.runner(root)
            passed=isinstance(result,dict) and result.get("passed") is True
            return {"status":"passed" if passed else "failed_or_inconclusive",
                    "result":result,"build_manifests":sorted(set(candidates))}
        except Exception as e:
            return {"status":"error","reason":f"{type(e).__name__}: {e}"}
