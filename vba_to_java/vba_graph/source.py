from pathlib import Path
import hashlib

EXTENSIONS = {".bas", ".cls", ".frm", ".frx", ".vba", ".txt"}

def decode_bytes(data: bytes):
    if data.startswith(b"\xef\xbb\xbf"):
        return data.decode("utf-8-sig"), "utf-8-sig"
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        return data.decode("utf-16"), "utf-16"
    for encoding in ("utf-8", "cp1252", "latin-1"):
        try:
            return data.decode(encoding), encoding
        except UnicodeDecodeError:
            pass
    return data.decode("utf-8", errors="replace"), "utf-8-replace"

def read_project(path):
    root = Path(path)
    if not root.exists():
        raise FileNotFoundError(root)
    paths = [root] if root.is_file() else sorted(
        p for p in root.rglob("*") if p.is_file() and p.suffix.lower() in EXTENSIONS
    )
    sources, manifest = {}, []
    for p in paths:
        if p.suffix.lower() == ".frx":
            # Binary form resource: inventory/hash it; never decode as source text.
            data = p.read_bytes()
            rel = p.name if root.is_file() else str(p.relative_to(root))
            manifest.append({"path": rel, "kind": "binary_resource", "sha256": hashlib.sha256(data).hexdigest(),
                             "size_bytes": len(data)})
            continue
        data = p.read_bytes()
        text, encoding = decode_bytes(data)
        rel = p.name if root.is_file() else str(p.relative_to(root))
        sources[rel] = text
        manifest.append({"path": rel, "kind": "source", "encoding": encoding,
                         "sha256": hashlib.sha256(data).hexdigest(), "size_bytes": len(data)})
    return sources, manifest
