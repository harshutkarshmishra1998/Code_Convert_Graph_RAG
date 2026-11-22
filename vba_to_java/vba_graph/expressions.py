"""Conservative expression metadata. Raw expression text is always retained."""
import re

TOKEN_RE = re.compile(
    r'"(?:[^"]|"")*"|#(?:[^#]|##)*#|<=|>=|<>|:=|'
    r'[A-Za-z_][A-Za-z0-9_]*|(?:&H[0-9A-F]+|&O[0-7]+|'
    r'\d+(?:\.\d*)?(?:[Ee][+-]?\d+)?)|[^\s]', re.I
)

def expression_info(source):
    tokens = TOKEN_RE.findall(source)
    identifiers = []
    for token in tokens:
        if token.startswith('"') or token.startswith("#") or token[:1].isdigit():
            continue
        if re.match(r"^[A-Za-z_]\w*$", token):
            if token.upper() not in {
                "AND","OR","NOT","XOR","EQV","IMP","MOD","IS","LIKE",
                "TRUE","FALSE","NOTHING","NULL","EMPTY","ME","NEW"
            }:
                identifiers.append(token)
    return {"source": source, "tokens": tokens, "identifiers": identifiers,
            "parse_status": "tokenized_not_fully_parsed"}
