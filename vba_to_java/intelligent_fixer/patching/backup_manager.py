"""Create byte-for-byte backups before every in-place patch."""
from pathlib import Path
import shutil, hashlib
class BackupManager:
    def backup(self,path):
        p=Path(path); backup=p.with_name(p.name+".intelligent_fixer.bak")
        if backup.exists():
            digest=hashlib.sha256(p.read_bytes()).hexdigest()
            existing=hashlib.sha256(backup.read_bytes()).hexdigest()
            if digest==existing: return backup
            # Preserve first backup; never overwrite it.
            backup=p.with_name(p.name+f".{digest[:8]}.intelligent_fixer.bak")
        shutil.copy2(p,backup)
        return backup
