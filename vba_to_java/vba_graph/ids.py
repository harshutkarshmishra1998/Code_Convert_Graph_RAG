import hashlib

def stable_id(prefix: str, path: str, start: int, end: int, name: str = "") -> str:
    raw = f"{path}\0{start}\0{end}\0{name}".encode("utf-8", errors="replace")
    return f"{prefix}_{hashlib.sha1(raw).hexdigest()[:16]}"
