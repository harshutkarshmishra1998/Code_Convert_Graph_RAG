"""Hierarchical, source-preserving VBA parser.

This is a conservative structural parser. Every source statement is represented, even when
its syntax is not understood. It does not claim to be a full VBA grammar implementation.
"""
import re
from .lexer import logical_lines, split_colon_statements
from .expressions import expression_info

PROC_RE = re.compile(
    r"^(?:(Public|Private|Friend|Static)\s+)?"
    r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+"
    r"([A-Za-z_][A-Za-z0-9_]*)\s*(?:\((.*)\))?(?:\s+As\s+(.+))?$", re.I
)

def split_args(src):
    if not src.strip():
        return []
    parts, buf, depth, quoted = [], [], 0, False
    i = 0
    while i < len(src):
        c = src[i]
        if c == '"':
            if quoted and i + 1 < len(src) and src[i+1] == '"':
                buf.extend(['"', '"']); i += 2; continue
            quoted = not quoted
        if not quoted:
            if c == "(": depth += 1
            elif c == ")": depth -= 1
            elif c == "," and depth == 0:
                parts.append("".join(buf).strip()); buf=[]; i+=1; continue
        buf.append(c); i += 1
    if buf:
        parts.append("".join(buf).strip())
    result = []
    for arg in parts:
        m = re.match(
            r"(?:(Optional|ParamArray)\s+)?(?:(ByVal|ByRef)\s+)?"
            r"([A-Za-z_]\w*)(?:\(\))?(?:\s+As\s+([^=]+?))?(?:\s*=\s*(.*))?$",
            arg, re.I
        )
        if m:
            result.append({"name":m.group(3),"passing":(m.group(2) or "ByRef").title(),
                           "optional":(m.group(1) or "").lower()=="optional",
                           "param_array":(m.group(1) or "").lower()=="paramarray",
                           "type":(m.group(4) or "Variant").strip(),"default":m.group(5),
                           "source":arg})
        else:
            result.append({"name":arg,"type":"Unknown","source":arg,"parse_status":"raw"})
    return result


def _assignment_parts(source):
    quoted = False
    depth = 0
    i = 0
    while i < len(source):
        char = source[i]
        if char == '"':
            if quoted and i + 1 < len(source) and source[i + 1] == '"':
                i += 2
                continue
            quoted = not quoted
        elif not quoted:
            if char == "(":
                depth += 1
            elif char == ")":
                depth = max(0, depth - 1)
            elif char == "=" and depth == 0:
                previous = source[i - 1] if i else ""
                following = source[i + 1] if i + 1 < len(source) else ""
                if previous not in "<>:" and following != "=":
                    return source[:i].strip(), source[i + 1:].strip()
        i += 1
    return None


def classify_statement(text):
    s = text.strip()
    low = s.lower()
    base = {"source":s}
    if not s:
        return None
    if re.match(r"^(rem\b|')", s, re.I):
        return dict(base, type="COMMENT")
    label = re.match(r"^([A-Za-z_]\w*|\d+):$", s)
    if label:
        return dict(base, type="LABEL", name=label.group(1))
    if re.match(r"^(on\s+error|resume\b)", s, re.I):
        return dict(base, type="ERROR_CONTROL")
    if re.match(r"^(goto|gosub)\b", s, re.I):
        m = re.match(r"^(GoTo|GoSub)\s+(.+)$", s, re.I)
        return dict(base, type="JUMP", operation=m.group(1).upper(), target=m.group(2).strip())
    if re.match(r"^return\b", s, re.I):
        return dict(base, type="RETURN_CONTROL")
    if re.match(r"^exit\s+", s, re.I):
        return dict(base, type="EXIT", target=s.split(None,1)[1])
    if re.match(r"^(dim|static|const|redim|erase|preserve)\b", s, re.I):
        return dict(base, type="DECLARATION")
    if re.match(r"^debug\.print\b", s, re.I):
        return dict(base, type="HOST_CALL", host="Debug.Print")
    if re.match(r"^(set|let)\s+", s, re.I):
        m = re.match(r"^(Set|Let)\s+(.+)$", s, re.I)
        parts = _assignment_parts(m.group(2)) if m else None
        if m and parts:
            return dict(base, type="ASSIGNMENT", assignment_kind=m.group(1).upper(),
                        target=parts[0], expression=expression_info(parts[1]))
    parts = _assignment_parts(s)
    if parts and not re.match(r"^(if|elseif)\b", s, re.I):
        return dict(base, type="ASSIGNMENT", assignment_kind="IMPLICIT",
                    target=parts[0], expression=expression_info(parts[1]))
    if re.match(r"^(?:call\s+)?(?:[A-Za-z_]\w*(?:[.!][A-Za-z_]\w*)*|\.[A-Za-z_]\w*)\s*(?:\(|\s+|$)", s, re.I):
        return dict(base, type="CALL_OR_COMMAND",
                    expression=expression_info(re.sub(r"^call\s+","",s,flags=re.I)))
    return dict(base, type="OPAQUE_STATEMENT", parse_status="unparsed")

def parse_module(source, path, module_name, module_kind):
    module = {"name":module_name,"kind":module_kind,"path":path,"options":[],"attributes":[],
              "declarations":[],"procedures":[],"raw_source":source}
    diagnostics = []
    current = None
    # stack frame: block node and currently active child list
    stack = []

    def active_body():
        if stack:
            return stack[-1]["body"]
        if current is not None:
            return current["body"]
        return module["declarations"]

    def add(node):
        active_body().append(node)

    for logical in logical_lines(source):
        text = logical["text"].strip()
        loc = {"start_line":logical["start_line"],"end_line":logical["end_line"]}
        if not text:
            continue
        if text.startswith("#"):
            add({"type":"COMPILER_DIRECTIVE","source":text,"location":loc})
            continue
        if re.match(r"^Attribute\b", text, re.I):
            module["attributes"].append({"source":text,"location":loc}); continue
        if re.match(r"^Option\b", text, re.I):
            module["options"].append({"source":text,"location":loc}); continue

        pm = PROC_RE.match(text)
        if pm and current is None:
            current = {"name":pm.group(3),"kind":re.sub(r"\s+"," ",pm.group(2)).upper(),
                       "visibility":(pm.group(1) or "Public").upper(),
                       "parameters":split_args(pm.group(4) or ""),
                       "return_type":pm.group(5) or ("Variant" if "function" in pm.group(2).lower() else None),
                       "body":[],"location":loc,"raw_header":text}
            module["procedures"].append(current)
            continue

        if current is None:
            module["declarations"].append({
                "type":"MODULE_DECLARATION" if re.match(
                    r"^(public|private|global|dim|const|declare|type|enum|implements|event)\b",text,re.I
                ) else "OPAQUE_MODULE_LINE",
                "source":text,"location":loc
            })
            continue

        if re.match(r"^End\s+(Sub|Function|Property)\s*$", text, re.I):
            if stack:
                diagnostics.append({
                    "severity": "warning",
                    "code": "UNCLOSED_BLOCK_AT_PROCEDURE_END",
                    "message": f"{len(stack)} block frame(s) remained open at procedure end.",
                    "path": path,
                    "line": loc["start_line"],
                })
                for frame in stack:
                    frame["node"].setdefault("end_location", loc)
            current["end_location"] = loc
            current = None
            stack = []
            continue

        for statement in split_colon_statements(text):
            # Block openers and branch transitions.
            m = re.match(r"^If\s+(.+?)\s+Then\s*$", statement, re.I)
            if m:
                cond = expression_info(m.group(1))
                node = {"type":"IF_BLOCK","condition":cond,
                        "branches":[{"branch":"IF","condition":cond,"body":[]}],
                        "else_body":None,"source":statement,"location":loc}
                add(node); stack.append({"type":"IF","node":node,"body":node["branches"][0]["body"]}); continue

            m = re.match(r"^ElseIf\s+(.+?)\s+Then\s*$", statement, re.I)
            if m and stack and stack[-1]["type"] == "IF":
                branch = {"branch":"ELSEIF","condition":expression_info(m.group(1)),"body":[]}
                stack[-1]["node"]["branches"].append(branch); stack[-1]["body"] = branch["body"]; continue

            if re.match(r"^Else\s*$", statement, re.I) and stack and stack[-1]["type"] == "IF":
                stack[-1]["node"]["else_body"] = []
                stack[-1]["body"] = stack[-1]["node"]["else_body"]; continue

            m = re.match(r"^For\s+Each\s+(.+?)\s+In\s+(.+)$", statement, re.I)
            if m:
                node = {"type":"FOR_EACH_BLOCK","variable":m.group(1),
                        "collection":expression_info(m.group(2)),"body":[],"source":statement,"location":loc}
                add(node); stack.append({"type":"LOOP","node":node,"body":node["body"]}); continue

            m = re.match(r"^For\s+(.+?)\s*=\s*(.+?)\s+To\s+(.+?)(?:\s+Step\s+(.+))?$", statement, re.I)
            if m:
                node = {"type":"FOR_BLOCK","variable":m.group(1),"start":expression_info(m.group(2)),
                        "end":expression_info(m.group(3)),
                        "step":expression_info(m.group(4)) if m.group(4) else None,
                        "body":[],"source":statement,"location":loc}
                add(node); stack.append({"type":"LOOP","node":node,"body":node["body"]}); continue

            m = re.match(r"^Next(?:\s+(.+))?$", statement, re.I)
            if m:
                variables = [item.strip() for item in (m.group(1) or "").split(",") if item.strip()]
                for variable in variables or [None]:
                    if (not stack or stack[-1]["type"] != "LOOP"
                            or stack[-1]["node"].get("type") not in {"FOR_BLOCK", "FOR_EACH_BLOCK"}):
                        add({"type":"UNMATCHED_BLOCK_TERMINATOR","source":statement,"location":loc})
                        diagnostics.append({"severity":"warning","code":"UNMATCHED_FOR_NEXT",
                            "message":"Next has no matching open For block.","path":path,
                            "line":loc["start_line"],"source":statement})
                        break
                    frame = stack[-1]
                    expected = re.sub(r"\s+As\s+.+$", "", str(frame["node"].get("variable", "")),
                                      flags=re.I).strip()
                    if variable and expected.casefold() != variable.casefold():
                        add({"type":"UNMATCHED_BLOCK_TERMINATOR","source":statement,"location":loc})
                        diagnostics.append({"severity":"warning","code":"FOR_NEXT_VARIABLE_MISMATCH",
                            "message":f"Next variable {variable!r} does not match open For variable {expected!r}.",
                            "path":path,"line":loc["start_line"],"source":statement})
                        break
                    stack.pop()
                    frame["node"]["end_location"] = loc
                    frame["node"]["terminator"] = {
                        "source": statement,
                        "location": loc,
                    }
                continue

            m = re.match(r"^Do(?:\s+(While|Until)\s+(.+))?$", statement, re.I)
            if m:
                node = {"type":"DO_LOOP_BLOCK","test_position":"PRE" if m.group(1) else "POST",
                        "test":m.group(1).upper() if m.group(1) else None,
                        "condition":expression_info(m.group(2)) if m.group(2) else None,
                        "body":[],"source":statement,"location":loc}
                add(node); stack.append({"type":"LOOP","node":node,"body":node["body"]}); continue

            m = re.match(r"^Loop(?:\s+(While|Until)\s+(.+))?$", statement, re.I)
            if m:
                if (stack and stack[-1]["type"] == "LOOP"
                        and stack[-1]["node"].get("type") == "DO_LOOP_BLOCK"):
                    frame = stack.pop()
                    frame["node"]["end_location"] = loc
                    frame["node"]["terminator"] = {"source": statement, "location": loc}
                    if frame["node"].get("test_position") == "POST":
                        frame["node"]["test"] = m.group(1).upper() if m.group(1) else None
                        frame["node"]["condition"] = expression_info(m.group(2)) if m.group(2) else None
                else:
                    add({"type":"UNMATCHED_BLOCK_TERMINATOR","source":statement,"location":loc})
                    diagnostics.append({"severity":"warning","code":"UNMATCHED_LOOP_END",
                        "message":"Loop has no matching open Do block.","path":path,
                        "line":loc["start_line"],"source":statement})
                continue

            m = re.match(r"^While\s+(.+)$", statement, re.I)
            if m:
                node = {"type":"WHILE_BLOCK","condition":expression_info(m.group(1)),
                        "body":[],"source":statement,"location":loc}
                add(node); stack.append({"type":"LOOP","node":node,"body":node["body"]}); continue

            if re.match(r"^Wend$", statement, re.I):
                if (stack and stack[-1]["type"] == "LOOP"
                        and stack[-1]["node"].get("type") == "WHILE_BLOCK"):
                    frame = stack.pop()
                    frame["node"]["end_location"] = loc
                    frame["node"]["terminator"] = {"source": statement, "location": loc}
                else:
                    add({"type":"UNMATCHED_BLOCK_TERMINATOR","source":statement,"location":loc})
                    diagnostics.append({"severity":"warning","code":"UNMATCHED_WEND",
                        "message":"Wend has no matching open While block.","path":path,
                        "line":loc["start_line"]})
                continue

            m = re.match(r"^Select\s+Case\s+(.+)$", statement, re.I)
            if m:
                node = {"type":"SELECT_BLOCK","expression":expression_info(m.group(1)),
                        "cases":[],"body":[],"source":statement,"location":loc}
                add(node); stack.append({"type":"SELECT","node":node,"body":node["body"]}); continue

            m = re.match(r"^Case\s+(.+)$", statement, re.I)
            if m and stack and stack[-1]["type"] == "SELECT":
                case = {"condition_source":m.group(1),"body":[]}
                stack[-1]["node"]["cases"].append(case); stack[-1]["body"] = case["body"]; continue

            m = re.match(r"^With\s+(.+)$", statement, re.I)
            if m:
                node = {"type":"WITH_BLOCK","expression":expression_info(m.group(1)),
                        "body":[],"source":statement,"location":loc}
                add(node); stack.append({"type":"WITH","node":node,"body":node["body"]}); continue

            if re.match(r"^End\s+With$", statement, re.I):
                if stack and stack[-1]["type"] == "WITH":
                    frame = stack.pop()
                    frame["node"]["end_location"] = loc
                    frame["node"]["terminator"] = {"source": statement, "location": loc}
                else:
                    add({"type":"UNMATCHED_BLOCK_TERMINATOR","source":statement,"location":loc})
                    diagnostics.append({"severity":"warning","code":"UNMATCHED_END_WITH",
                        "message":"End With has no matching open With block.","path":path,
                        "line":loc["start_line"]})
                continue

            if re.match(r"^End\s+Select$", statement, re.I):
                if stack and stack[-1]["type"] == "SELECT":
                    frame = stack.pop()
                    frame["node"]["end_location"] = loc
                    frame["node"]["terminator"] = {"source": statement, "location": loc}
                else:
                    add({"type":"UNMATCHED_BLOCK_TERMINATOR","source":statement,"location":loc})
                    diagnostics.append({"severity":"warning","code":"UNMATCHED_END_SELECT",
                        "message":"End Select has no matching open Select Case block.",
                        "path":path,"line":loc["start_line"]})
                continue

            if re.match(r"^End\s+If$", statement, re.I):
                idx = next((i for i in range(len(stack)-1,-1,-1) if stack[i]["type"]=="IF"), None)
                if idx is not None:
                    stack[idx]["node"]["end_location"] = loc
                    stack[idx]["node"]["terminator"] = {"source": statement, "location": loc}
                    if idx != len(stack)-1:
                        diagnostics.append({"severity":"warning","code":"CROSS_NESTED_BLOCK_CLOSE",
                            "message":"End If closed a non-top IF frame; inspect nested syntax.",
                            "path":path,"line":loc["start_line"]})
                    stack = stack[:idx]
                else:
                    add({"type":"UNMATCHED_BLOCK_TERMINATOR","source":statement,"location":loc})
                    diagnostics.append({"severity":"warning","code":"UNMATCHED_END_IF",
                        "message":"End If has no matching open block If.","path":path,
                        "line":loc["start_line"]})
                continue

            node = classify_statement(statement)
            if node:
                node["location"] = loc
                add(node)
                if node["type"] == "OPAQUE_STATEMENT":
                    diagnostics.append({"severity":"warning","code":"OPAQUE_STATEMENT",
                        "message":"Statement retained verbatim but not structurally interpreted.",
                        "path":path,"line":loc["start_line"],"source":statement})

    if current is not None:
        diagnostics.append({"severity":"warning","code":"UNCLOSED_PROCEDURE",
            "message":f"Procedure {current['name']} has no recognized terminator.",
            "path":path,"line":current["location"]["start_line"]})
    if stack:
        diagnostics.append({"severity":"warning","code":"UNCLOSED_BLOCK",
            "message":f"{len(stack)} block frame(s) remained open at end of procedure/file.",
            "path":path})
    return module, diagnostics
