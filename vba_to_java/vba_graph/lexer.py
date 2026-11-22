"""VBA-aware physical/logical line handling; original lines remain intact."""
def split_comment(line):
    quoted = False
    i = 0
    while i < len(line):
        c = line[i]
        if c == '"':
            if quoted and i + 1 < len(line) and line[i + 1] == '"':
                i += 2
                continue
            quoted = not quoted
        if c == "'" and not quoted:
            return line[:i], line[i:]
        i += 1
    return line, ""

def logical_lines(source):
    physical = source.splitlines()
    result, parts, raw_parts = [], [], []
    start = 1
    for lineno, raw in enumerate(physical, 1):
        if not parts:
            start = lineno
        raw_parts.append(raw)
        code, _ = split_comment(raw)
        stripped = code.rstrip()
        if stripped.endswith("_") and len(stripped) > 1 and stripped[-2].isspace():
            parts.append(stripped[:-1].rstrip())
            continue
        parts.append(code.strip())
        result.append({
            "text": " ".join(p for p in parts if p).strip(),
            "start_line": start, "end_line": lineno,
            "raw": "\n".join(raw_parts)
        })
        parts, raw_parts = [], []
    if parts:
        result.append({"text": " ".join(parts).strip(), "start_line": start,
                       "end_line": len(physical), "raw": "\n".join(raw_parts)})
    return result

def split_colon_statements(text):
    """Split VBA statement separators outside string literals. Preserve label prefixes."""
    out, buf = [], []
    quoted, i = False, 0
    while i < len(text):
        c = text[i]
        if c == '"':
            if quoted and i + 1 < len(text) and text[i+1] == '"':
                buf.extend(['"', '"']); i += 2; continue
            quoted = not quoted
        if c == ":" and not quoted and not (i + 1 < len(text) and text[i+1] == "="):
            part = "".join(buf).strip()
            if part:
                out.append(part)
            buf = []
        else:
            buf.append(c)
        i += 1
    part = "".join(buf).strip()
    if part:
        out.append(part)
    return out
