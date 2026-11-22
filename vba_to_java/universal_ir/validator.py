"""Structural and cross-reference validation for Universal IR (stdlib only)."""
from __future__ import annotations

from typing import Any

SUPPORTED_TYPE_KINDS = {
    "primitive", "named", "array", "collection", "map", "nullable",
    "variant", "object", "pointer", "unknown", "void"
}
SUPPORTED_STATEMENT_KINDS = {
    "block", "variable_declaration", "assignment", "expression", "if",
    "switch", "for", "for_each", "while", "do_while", "do_until",
    "call", "return", "gosub_return", "break", "continue", "exit", "label", "goto",
    "gosub", "on_error", "error_handler", "raise_error", "with",
    "redim", "erase", "file_io", "unsupported", "comment"
}
SUPPORTED_EXPRESSION_KINDS = {
    "literal", "identifier", "member_access", "index", "call",
    "unary", "binary", "conditional", "new", "nothing", "missing",
    "array_literal", "type_cast", "lambda", "unsupported"
}


def validate_program(program: dict[str, Any]) -> list[dict[str, str]]:
    """Return all validation issues; does not raise for user-data errors."""
    issues: list[dict[str, str]] = []

    def issue(path: str, message: str, code: str = "IR_INVALID") -> None:
        issues.append({"path": path, "code": code, "message": message})

    if not isinstance(program, dict):
        return [{"path": "$", "code": "IR_INVALID_ROOT", "message": "Program must be a JSON object."}]

    for required in ("ir_version", "source_language", "name", "modules"):
        if required not in program:
            issue("$", f"Missing required field '{required}'.", "IR_MISSING_FIELD")
    if program.get("ir_version") != "1.0.0":
        issue("$.ir_version", "This validator supports IR version 1.0.0.", "IR_VERSION_UNSUPPORTED")
    if not isinstance(program.get("modules", []), list):
        issue("$.modules", "modules must be an array.")
        return issues

    ids: dict[str, str] = {}
    labels: set[str] = set()

    def register_id(obj: dict[str, Any], path: str) -> None:
        obj_id = obj.get("id")
        if obj_id is None:
            issue(path, "Node is missing a stable 'id'.", "IR_MISSING_ID")
        elif not isinstance(obj_id, str) or not obj_id.strip():
            issue(path + ".id", "id must be a non-empty string.", "IR_INVALID_ID")
        elif obj_id in ids:
            issue(path + ".id", f"Duplicate id '{obj_id}' (also used at {ids[obj_id]}).", "IR_DUPLICATE_ID")
        else:
            ids[obj_id] = path

    def check_type(t: Any, path: str) -> None:
        if not isinstance(t, dict):
            issue(path, "Type must be an object.", "IR_INVALID_TYPE")
            return
        kind = t.get("kind")
        if kind not in SUPPORTED_TYPE_KINDS:
            issue(path + ".kind", f"Unsupported type kind: {kind!r}.", "IR_UNKNOWN_TYPE_KIND")
        if kind in {"primitive", "named", "object"} and not t.get("name"):
            issue(path + ".name", f"Type kind '{kind}' requires a name.", "IR_MISSING_TYPE_NAME")
        if kind in {"array", "collection", "nullable", "pointer"}:
            if "element_type" not in t:
                issue(path + ".element_type", f"Type kind '{kind}' requires element_type.", "IR_MISSING_TYPE_MEMBER")
            else:
                check_type(t["element_type"], path + ".element_type")
        if kind == "map":
            for key in ("key_type", "value_type"):
                if key not in t:
                    issue(path + "." + key, "Map type requires key_type and value_type.", "IR_MISSING_TYPE_MEMBER")
                else:
                    check_type(t[key], path + "." + key)

    def check_expr(expr: Any, path: str) -> None:
        if not isinstance(expr, dict):
            issue(path, "Expression must be an object.", "IR_INVALID_EXPRESSION")
            return
        kind = expr.get("kind")
        if kind not in SUPPORTED_EXPRESSION_KINDS:
            issue(path + ".kind", f"Unsupported expression kind: {kind!r}.", "IR_UNKNOWN_EXPRESSION_KIND")
        if kind == "literal" and "value" not in expr:
            issue(path + ".value", "Literal expression requires value (null is allowed).", "IR_MISSING_EXPRESSION_VALUE")
        if kind == "identifier" and not expr.get("name"):
            issue(path + ".name", "Identifier expression requires name.", "IR_MISSING_EXPRESSION_NAME")
        if kind == "binary":
            for key in ("operator", "left", "right"):
                if key not in expr:
                    issue(path + "." + key, f"Binary expression requires {key}.", "IR_INCOMPLETE_EXPRESSION")
        if kind == "unary":
            for key in ("operator", "operand"):
                if key not in expr:
                    issue(path + "." + key, f"Unary expression requires {key}.", "IR_INCOMPLETE_EXPRESSION")
        if kind == "member_access":
            if "object" not in expr or not expr.get("member"):
                issue(path, "Member access requires object and member.", "IR_INCOMPLETE_EXPRESSION")
        if kind == "index":
            if "target" not in expr or not isinstance(expr.get("indices"), list):
                issue(path, "Index expression requires target and indices array.", "IR_INCOMPLETE_EXPRESSION")
        if "type" in expr:
            check_type(expr["type"], path + ".type")
        for key in ("operand", "left", "right", "target", "object", "condition", "when_true", "when_false"):
            if key in expr and isinstance(expr[key], dict):
                check_expr(expr[key], path + "." + key)
        for key in ("arguments", "elements", "indices"):
            if key in expr and isinstance(expr[key], list):
                for i, child in enumerate(expr[key]):
                    if key == "arguments" and isinstance(child, dict) and "value" in child:
                        check_expr(child["value"], f"{path}.{key}[{i}].value")
                    elif key != "arguments":
                        check_expr(child, f"{path}.{key}[{i}]")

    def check_stmt(stmt: Any, path: str) -> None:
        if not isinstance(stmt, dict):
            issue(path, "Statement must be an object.", "IR_INVALID_STATEMENT")
            return
        register_id(stmt, path)
        kind = stmt.get("kind")
        if kind not in SUPPORTED_STATEMENT_KINDS:
            issue(path + ".kind", f"Unsupported statement kind: {kind!r}.", "IR_UNKNOWN_STATEMENT_KIND")
        if kind == "if":
            if "condition" not in stmt:
                issue(path + ".condition", "If statement requires condition.", "IR_INCOMPLETE_CONTROL_FLOW")
            if "body" not in stmt and "branches" not in stmt:
                issue(path, "If statement requires either body/else_body or branches.", "IR_INCOMPLETE_CONTROL_FLOW")
            if "condition" in stmt:
                check_expr(stmt["condition"], path + ".condition")
            for key in ("body", "else_body"):
                if key in stmt:
                    if not isinstance(stmt[key], list):
                        issue(path + "." + key, "Branch body must be an array.")
                    else:
                        for i, child in enumerate(stmt[key]):
                            check_stmt(child, f"{path}.{key}[{i}]")
            for bi, branch in enumerate(stmt.get("branches", [])):
                if not isinstance(branch, dict):
                    issue(f"{path}.branches[{bi}]", "If branch must be an object.")
                    continue
                if branch.get("branch_kind") not in {"if_branch", "elseif_branch", "else_branch"}:
                    issue(f"{path}.branches[{bi}].branch_kind", "Unknown IF branch kind.", "IR_INVALID_BRANCH_KIND")
                if branch.get("condition") is not None:
                    check_expr(branch["condition"], f"{path}.branches[{bi}].condition")
                if not isinstance(branch.get("body", []), list):
                    issue(f"{path}.branches[{bi}].body", "IF branch body must be an array.")
                else:
                    for si, child in enumerate(branch.get("body", [])):
                        check_stmt(child, f"{path}.branches[{bi}].body[{si}]")
        if kind in {"for", "for_each"}:
            if not stmt.get("variable"):
                issue(path + ".variable", f"{kind} statement requires a variable.", "IR_INCOMPLETE_CONTROL_FLOW")
            for key in ("start", "end") if kind == "for" else ("iterable",):
                if key not in stmt:
                    issue(path + "." + key, f"{kind} statement requires {key}.", "IR_INCOMPLETE_CONTROL_FLOW")
        if kind == "while" and "condition" not in stmt:
            issue(path + ".condition", "While statement requires condition.", "IR_INCOMPLETE_CONTROL_FLOW")
        if kind in {"do_while", "do_until"} and "condition" not in stmt and not stmt.get("semantics", {}).get("unconditional"):
            issue(path + ".condition", "Conditional Do statement requires condition.", "IR_INCOMPLETE_CONTROL_FLOW")
        if kind == "switch":
            if "expression" not in stmt:
                issue(path + ".expression", "Switch statement requires an expression.", "IR_INCOMPLETE_CONTROL_FLOW")
            branches = stmt.get("branches", [])
            if not isinstance(branches, list):
                issue(path + ".branches", "Switch branches must be an array.", "IR_INVALID_BRANCHES")
            else:
                defaults = 0
                for bi, branch in enumerate(branches):
                    bpath = f"{path}.branches[{bi}]"
                    if not isinstance(branch, dict):
                        issue(bpath, "Switch branch must be an object.", "IR_INVALID_BRANCH")
                        continue
                    branch_kind = branch.get("branch_kind")
                    if branch_kind not in {"case", "case_else"}:
                        issue(bpath + ".branch_kind", "Switch branch must be case or case_else.", "IR_INVALID_BRANCH_KIND")
                    if branch_kind == "case_else":
                        defaults += 1
                    elif not isinstance(branch.get("condition"), dict):
                        issue(bpath + ".condition", "Case branch requires a condition.", "IR_INCOMPLETE_CONTROL_FLOW")
                    elif branch.get("condition") is not None:
                        check_expr(branch["condition"], bpath + ".condition")
                    body = branch.get("body")
                    if not isinstance(body, list):
                        issue(bpath + ".body", "Switch branch body must be an array.", "IR_INVALID_BRANCH_BODY")
                    else:
                        for si, child in enumerate(body):
                            check_stmt(child, f"{bpath}.body[{si}]")
                if defaults > 1:
                    issue(path + ".branches", "Switch cannot contain multiple Case Else branches.", "IR_DUPLICATE_DEFAULT_CASE")
        if kind in {"for", "for_each", "while", "do_while", "do_until", "switch", "block", "with", "error_handler"}:
            for key in ("body", "statements"):
                if key in stmt:
                    if not isinstance(stmt[key], list):
                        issue(path + "." + key, "Statement body must be an array.")
                    else:
                        for i, child in enumerate(stmt[key]):
                            check_stmt(child, f"{path}.{key}[{i}]")
        for key in ("expression", "target", "value", "condition", "iterable", "start", "end", "step", "error_code", "message"):
            if key in stmt and isinstance(stmt[key], dict):
                check_expr(stmt[key], path + "." + key)
        for key in ("initializer", "update"):
            if key in stmt and isinstance(stmt[key], list):
                for i, child in enumerate(stmt[key]):
                    check_stmt(child, f"{path}.{key}[{i}]")
        if kind in {"goto", "gosub"} and not stmt.get("label"):
            issue(path + ".label", f"{kind} requires a label.", "IR_MISSING_LABEL")
        if kind == "variable_declaration":
            for di, decl in enumerate(stmt.get("declarations", [])):
                if not isinstance(decl, dict):
                    issue(f"{path}.declarations[{di}]", "Nested declaration must be an object.")
                    continue
                register_id(decl, f"{path}.declarations[{di}]")
                if not decl.get("name"):
                    issue(f"{path}.declarations[{di}].name", "Nested declaration requires a name.")
                if "type" in decl:
                    check_type(decl["type"], f"{path}.declarations[{di}].type")
        if kind == "label":
            label = stmt.get("label")
            if not label:
                issue(path + ".label", "Label statement requires a label name.", "IR_MISSING_LABEL")
            elif label in labels:
                issue(path + ".label", f"Duplicate label '{label}'.", "IR_DUPLICATE_LABEL")
            else:
                labels.add(label)
        if kind == "unsupported" and not stmt.get("semantics", {}).get("reason"):
            issue(path + ".semantics.reason", "Unsupported statement must explain why it is not represented.", "IR_UNEXPLAINED_UNSUPPORTED")

    for mi, module in enumerate(program.get("modules", [])):
        mpath = f"$.modules[{mi}]"
        if not isinstance(module, dict):
            issue(mpath, "Module must be an object.")
            continue
        register_id(module, mpath)
        if not module.get("name"):
            issue(mpath + ".name", "Module requires a name.")
        for di, decl in enumerate(module.get("declarations", [])):
            dpath = f"{mpath}.declarations[{di}]"
            if not isinstance(decl, dict):
                issue(dpath, "Declaration must be an object.")
                continue
            register_id(decl, dpath)
            if not decl.get("kind") or not decl.get("name"):
                issue(dpath, "Declaration requires kind and name.")
            if "type" in decl:
                check_type(decl["type"], dpath + ".type")
            if "value" in decl:
                check_expr(decl["value"], dpath + ".value")
        for pi, proc in enumerate(module.get("procedures", [])):
            # Labels are scoped to a procedure, not globally to the whole project.
            labels.clear()
            ppath = f"{mpath}.procedures[{pi}]"
            if not isinstance(proc, dict):
                issue(ppath, "Procedure must be an object.")
                continue
            register_id(proc, ppath)
            if not proc.get("name") or not proc.get("kind"):
                issue(ppath, "Procedure requires name and kind.")
            if "return_type" in proc:
                check_type(proc["return_type"], ppath + ".return_type")
            for vi, var in enumerate(proc.get("local_variables", [])):
                vpath = f"{ppath}.local_variables[{vi}]"
                register_id(var, vpath)
                if not var.get("name"):
                    issue(vpath + ".name", "Variable requires a name.")
                if "type" in var:
                    check_type(var["type"], vpath + ".type")
            for ai, param in enumerate(proc.get("parameters", [])):
                apath = f"{ppath}.parameters[{ai}]"
                if param.get("id"):
                    register_id(param, apath)
                if not param.get("name"):
                    issue(apath + ".name", "Parameter requires a name.")
                if "type" in param:
                    check_type(param["type"], apath + ".type")
            if not isinstance(proc.get("body", []), list):
                issue(ppath + ".body", "Procedure body must be an array.")
            else:
                for si, stmt in enumerate(proc.get("body", [])):
                    check_stmt(stmt, f"{ppath}.body[{si}]")
                # Validate local jump/error-handler labels within the current procedure.
                labels_in_proc: set[str] = set()
                refs_in_proc: list[tuple[str, str]] = []
                def collect_labels(items: list[Any]) -> None:
                    for item in items:
                        if not isinstance(item, dict):
                            continue
                        if item.get("kind") == "label" and item.get("label"):
                            labels_in_proc.add(str(item["label"]).casefold())
                        if item.get("kind") in {"goto", "gosub"} and item.get("label"):
                            refs_in_proc.append((str(item["label"]), "jump"))
                        if item.get("kind") == "on_error" and item.get("handler"):
                            refs_in_proc.append((str(item["handler"]), "error handler"))
                        for key in ("body", "else_body", "statements"):
                            if isinstance(item.get(key), list): collect_labels(item[key])
                        for branch in item.get("branches", []):
                            if isinstance(branch, dict) and isinstance(branch.get("body"), list): collect_labels(branch["body"])
                collect_labels(proc.get("body", []))
                for target, ref_kind in refs_in_proc:
                    if target.casefold() not in labels_in_proc:
                        issue(ppath + ".body", f"Unresolved {ref_kind} label '{target}'.", "IR_UNRESOLVED_LABEL")

    # Validate simple intra-program references after all ids have been registered.
    for ri, ref in enumerate(program.get("references", [])):
        target = ref.get("target_id") if isinstance(ref, dict) else None
        if target and target not in ids:
            issue(f"$.references[{ri}].target_id", f"Reference target '{target}' does not exist.", "IR_DANGLING_REFERENCE")
    return issues
