"""Convert common VBA constructs into readable, structured pseudocode."""

from __future__ import annotations

import argparse
import re
from pathlib import Path


_PROCEDURE = re.compile(
    r"^(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function|Property\s+Get|"
    r"Property\s+Let|Property\s+Set)\s+([A-Za-z_]\w*)\s*(?:\((.*?)\))?"
    r"(?:\s+As\s+([\w.]+))?$",
    re.IGNORECASE,
)


def _strip_comment(line: str) -> str:
    """Remove apostrophe and Rem comments without changing string contents."""
    in_string = False
    result: list[str] = []
    index = 0

    while index < len(line):
        char = line[index]
        if char == '"':
            result.append(char)
            if in_string and index + 1 < len(line) and line[index + 1] == '"':
                result.append('"')
                index += 2
                continue
            in_string = not in_string
        elif char == "'" and not in_string:
            break
        elif not in_string and line[index : index + 3].casefold() == "rem":
            before = line[index - 1] if index else " "
            after_index = index + 3
            after = line[after_index] if after_index < len(line) else " "
            if (before.isspace() or before == ":") and after.isspace():
                break
            result.append(char)
        else:
            result.append(char)
        index += 1

    return "".join(result).strip()


def _logical_lines(source: str) -> list[str]:
    lines: list[str] = []
    continued = ""

    for physical_line in source.splitlines():
        line = physical_line.rstrip()
        if line.endswith(" _") and _strip_comment(line).endswith("_"):
            continued += line[:-2].strip() + " "
            continue
        lines.append((continued + line).strip())
        continued = ""

    if continued:
        lines.append(continued.strip())
    return lines


def _find_keyword(statement: str, keyword: str) -> int:
    """Find a VBA keyword outside strings and parenthesized expressions."""
    in_string = False
    depth = 0
    index = 0

    while index < len(statement):
        char = statement[index]
        if char == '"':
            if in_string and index + 1 < len(statement) and statement[index + 1] == '"':
                index += 2
                continue
            in_string = not in_string
        elif not in_string:
            if char == "(":
                depth += 1
            elif char == ")":
                depth = max(0, depth - 1)
            elif depth == 0 and statement[index : index + len(keyword)].casefold() == keyword.casefold():
                before = statement[index - 1] if index else " "
                after_index = index + len(keyword)
                after = statement[after_index] if after_index < len(statement) else " "
                if not (before.isalnum() or before == "_") and not (
                    after.isalnum() or after == "_"
                ):
                    return index
        index += 1
    return -1


def _split_statements(line: str) -> list[str]:
    """Split VBA's colon-separated statements, except colons inside strings."""
    label = re.match(r"^\s*((?:[A-Za-z_]\w*|\d+)\s*:)(.*)$", line)
    if label and label.group(1).strip(" :").casefold() not in {
        "else",
        "then",
        "case",
    }:
        rest = label.group(2).strip()
        return [label.group(1).strip(), *_split_statements(rest)] if rest else [
            label.group(1).strip()
        ]

    if re.match(r"^\s*If\b", line, re.IGNORECASE):
        then_index = _find_keyword(line, "Then")
        if then_index >= 0 and line[then_index + len("Then") :].strip():
            return [line.strip()]

    statements: list[str] = []
    current: list[str] = []
    in_string = False
    index = 0

    while index < len(line):
        char = line[index]
        if char == '"':
            current.append(char)
            if in_string and index + 1 < len(line) and line[index + 1] == '"':
                current.append('"')
                index += 2
                continue
            in_string = not in_string
        elif (
            char == ":"
            and not in_string
            and not (index + 1 < len(line) and line[index + 1] == "=")
        ):
            if current:
                statements.append("".join(current).strip())
            current = []
        else:
            current.append(char)
        index += 1

    if current:
        statements.append("".join(current).strip())
    return [statement for statement in statements if statement]


def _format_parameters(parameters: str | None) -> str:
    if parameters is None:
        return ""
    if not parameters.strip():
        return "()"
    formatted = []
    for parameter in _split_declarations(parameters):
        parameter = re.sub(
            r"\b(ByVal|ByRef|Optional|ParamArray)\b\s*",
            "",
            parameter.strip(),
            flags=re.IGNORECASE,
        )
        parameter = re.sub(r"\s+As\s+", ": ", parameter, flags=re.IGNORECASE)
        formatted.append(parameter)
    return f"({', '.join(formatted)})"


def _split_declarations(declarations: str) -> list[str]:
    parts: list[str] = []
    current: list[str] = []
    depth = 0
    in_string = False

    index = 0
    while index < len(declarations):
        char = declarations[index]
        if char == '"':
            if in_string and index + 1 < len(declarations) and declarations[index + 1] == '"':
                current.extend(('"', '"'))
                index += 2
                continue
            in_string = not in_string
        elif not in_string and char == "(":
            depth += 1
        elif not in_string and char == ")":
            depth = max(0, depth - 1)
        elif not in_string and depth == 0 and char == ",":
            parts.append("".join(current).strip())
            current = []
            index += 1
            continue
        current.append(char)
        index += 1

    if current:
        parts.append("".join(current).strip())
    return parts


def _assignment_parts(statement: str) -> tuple[str, str] | None:
    in_string = False
    depth = 0
    index = 0

    while index < len(statement):
        char = statement[index]
        if char == '"':
            if in_string and index + 1 < len(statement) and statement[index + 1] == '"':
                index += 2
                continue
            in_string = not in_string
        elif not in_string:
            if char == "(":
                depth += 1
            elif char == ")":
                depth = max(0, depth - 1)
            elif char == "=" and depth == 0:
                previous = statement[index - 1] if index else ""
                following = statement[index + 1] if index + 1 < len(statement) else ""
                if previous in ":<>" or following == "=":
                    index += 1
                    continue

                target = statement[:index].strip()
                target = re.sub(r"^(?:Let|Set)\s+", "", target, flags=re.IGNORECASE)
                target_depth = 0
                target_in_string = False
                for target_char in target:
                    if target_char == '"':
                        target_in_string = not target_in_string
                    elif not target_in_string and target_char == "(":
                        target_depth += 1
                    elif not target_in_string and target_char == ")":
                        target_depth = max(0, target_depth - 1)
                    elif not target_in_string and target_depth == 0 and target_char.isspace():
                        return None
                if not target or target.endswith(":"):
                    return None
                return target, statement[index + 1 :].strip()
        index += 1
    return None


def _simple_statement(statement: str) -> str:
    statement = statement.strip()
    if not statement:
        return ""

    if re.match(r"^(Option|Attribute)\b", statement, re.IGNORECASE):
        return ""
    if re.match(r"^Rem(?:\s|$)", statement, re.IGNORECASE):
        return ""

    constant = re.match(
        r"^(?:(?:Private|Public|Global)\s+)?Const\s+(.+)$",
        statement,
        re.IGNORECASE,
    )
    if constant:
        value = re.sub(r"\s*=\s*", " initialized to ", constant.group(1), count=1)
        return f"CONSTANT {value}"

    declaration = re.match(
        r"^(?:Dim|Private|Public|Global|Static)\s+(.+)$",
        statement,
        re.IGNORECASE,
    )
    if declaration:
        variables = []
        for item in _split_declarations(declaration.group(1)):
            item = re.sub(r"\s+As\s+", ": ", item, flags=re.IGNORECASE)
            item = re.sub(r"\s*=\s*", " initialized to ", item, count=1)
            variables.append(item)
        return "DECLARE " + ", ".join(variables)

    if re.match(r"^MsgBox\b", statement, re.IGNORECASE):
        return re.sub(r"^MsgBox\b", "DISPLAY", statement, flags=re.IGNORECASE)
    if re.match(r"^Debug\.Print\b", statement, re.IGNORECASE):
        return re.sub(r"^Debug\.Print\b", "DISPLAY", statement, flags=re.IGNORECASE)
    if re.match(r"^InputBox\b", statement, re.IGNORECASE):
        return re.sub(r"^InputBox\b", "GET INPUT", statement, flags=re.IGNORECASE)
    if re.match(r"^Call\s+", statement, re.IGNORECASE):
        return "CALL " + re.sub(r"^Call\s+", "", statement, flags=re.IGNORECASE)
    if re.match(r"^Exit\s+(Sub|Function|Property)\b", statement, re.IGNORECASE):
        return "RETURN"
    if re.match(r"^Exit\s+(For|Do)\b", statement, re.IGNORECASE):
        return "BREAK"
    if re.match(r"^On\s+Error\s+GoTo\s+0$", statement, re.IGNORECASE):
        return "DISABLE ERROR HANDLING"
    if re.match(r"^On\s+Error\s+Resume\s+Next$", statement, re.IGNORECASE):
        return "ON ERROR, CONTINUE WITH NEXT STATEMENT"

    error_handler = re.match(r"^On\s+Error\s+GoTo\s+(\w+)$", statement, re.IGNORECASE)
    if error_handler:
        return f"ON ERROR, GO TO {error_handler.group(1)}"

    if re.match(r"^GoTo\s+", statement, re.IGNORECASE):
        return re.sub(r"^GoTo\s+", "GO TO ", statement, flags=re.IGNORECASE)
    if re.match(r"^Resume\b", statement, re.IGNORECASE):
        resume_target = re.sub(r"^Resume\b", "", statement, flags=re.IGNORECASE).strip()
        return "RETRY NEXT STATEMENT" if resume_target.casefold() == "next" else (
            f"RETRY FROM {resume_target}" if resume_target else "RETRY"
        )
    if re.match(r"^On\s+Error\s+GoTo\s+-1$", statement, re.IGNORECASE):
        return "RESET ERROR STATE"

    resize = re.match(r"^ReDim\s+(Preserve\s+)?(.+)$", statement, re.IGNORECASE)
    if resize:
        preserve = " PRESERVING EXISTING VALUES" if resize.group(1) else ""
        return f"RESIZE{preserve} {resize.group(2)}"
    if re.match(r"^Erase\s+", statement, re.IGNORECASE):
        return re.sub(r"^Erase\s+", "CLEAR ", statement, flags=re.IGNORECASE)

    assignment = _assignment_parts(statement)
    if assignment:
        target, value = assignment
        value = re.sub(r"\bNothing\b", "NULL", value, flags=re.IGNORECASE)
        value = re.sub(r"\bTrue\b", "TRUE", value, flags=re.IGNORECASE)
        value = re.sub(r"\bFalse\b", "FALSE", value, flags=re.IGNORECASE)
        value = re.sub(r"\bNew\b", "CREATE", value, flags=re.IGNORECASE)
        if target.startswith("."):
            target = "CURRENT WITH-OBJECT" + target
        return f"SET {target} TO {value}"

    label = re.match(r"^((?:[A-Za-z_]\w*|\d+))\s*:$", statement)
    if label:
        return f"LABEL {label.group(1)}"
    if statement.startswith("."):
        return f"EXECUTE CURRENT WITH-OBJECT{statement}"

    return f"EXECUTE {statement}"


def _translate_inline_if(statement: str) -> tuple[str, str, str] | None:
    if_match = re.match(r"^If\s+", statement, re.IGNORECASE)
    if not if_match:
        return None

    then_index = _find_keyword(statement, "Then")
    if then_index < 0:
        return None
    condition = statement[if_match.end() : then_index].strip()
    body = statement[then_index + len("Then") :].strip()
    if not body:
        return None

    else_index = _find_keyword(body, "Else")
    if else_index < 0:
        return condition, body, ""
    return condition, body[:else_index].strip(), body[else_index + len("Else") :].strip()


def convert_vba_to_pseudocode(source: str) -> str:
    """Convert common VBA control flow and statements into pseudocode.

    Constructs outside the supported set are retained as ``EXECUTE`` steps
    rather than silently discarded.
    """
    output: list[str] = []
    indent = 0
    blocks: list[tuple[str, int]] = []
    lines = _logical_lines(source)

    def emit(text: str, level: int | None = None) -> None:
        if text:
            output.append("    " * (indent if level is None else level) + text)

    def open_block(kind: str, text: str) -> None:
        nonlocal indent
        emit(text)
        blocks.append((kind, indent))
        indent += 1

    def close_block(kind: str, text: str) -> None:
        nonlocal indent
        matching_index = next(
            (index for index in range(len(blocks) - 1, -1, -1) if blocks[index][0] == kind),
            None,
        )
        if matching_index is None:
            emit(f"EXECUTE {text} [unmatched block terminator]")
            return

        while len(blocks) - 1 > matching_index:
            unclosed_kind, base_indent = blocks.pop()
            indent = base_indent
            emit(
                f"{_block_end_text(unclosed_kind)} "
                f"[implicit: closed by {text}]"
            )
        _, base_indent = blocks.pop()
        indent = base_indent
        emit(text)

    for raw_line in lines:
        line = _strip_comment(raw_line)
        if not line:
            continue

        for statement in _split_statements(line):
            inline_if = _translate_inline_if(statement)
            if inline_if:
                condition_text, true_branch, false_branch = inline_if
                open_block("if", f"IF {condition_text} THEN")
                for action in _split_statements(true_branch):
                    translated = _simple_statement(action)
                    if translated:
                        emit(translated)
                if false_branch:
                    indent -= 1
                    emit("ELSE")
                    indent += 1
                    for action in _split_statements(false_branch):
                        translated = _simple_statement(action)
                        if translated:
                            emit(translated)
                close_block("if", "END IF")
                continue

            procedure = _PROCEDURE.match(statement)
            if procedure:
                kind = procedure.group(1).upper()
                name = procedure.group(2)
                parameters = _format_parameters(procedure.group(3))
                return_type = procedure.group(4)
                if kind == "FUNCTION" or kind.startswith("PROPERTY"):
                    if kind.startswith("PROPERTY"):
                        result = f"{kind} {name}{parameters}"
                    else:
                        result = f"FUNCTION {name}{parameters}"
                    if return_type and not kind.startswith("PROPERTY"):
                        result += f" RETURNS {return_type}"
                    open_block("procedure", result)
                else:
                    open_block("procedure", f"PROCEDURE {name}{parameters}")
                continue

            if re.match(r"^End\s+(Sub|Function|Property)\b", statement, re.IGNORECASE):
                close_block("procedure", "END PROCEDURE")
                continue
            if re.match(r"^End\s+If$", statement, re.IGNORECASE):
                close_block("if", "END IF")
                continue
            if re.match(r"^End\s+Select$", statement, re.IGNORECASE):
                close_block("select", "END SELECT")
                continue
            if re.match(r"^End\s+With$", statement, re.IGNORECASE):
                close_block("with", "END WITH")
                continue
            if re.match(r"^Wend$", statement, re.IGNORECASE):
                close_block("while", "END WHILE")
                continue

            branch = re.match(r"^ElseIf\s+(.+?)\s+Then$", statement, re.IGNORECASE)
            if branch:
                if blocks and blocks[-1][0] == "if":
                    _, base_indent = blocks[-1]
                    indent = base_indent
                    emit(f"ELSE IF {branch.group(1)} THEN")
                    indent += 1
                else:
                    emit("EXECUTE ElseIf [unmatched conditional branch]")
                continue
            if re.match(r"^Else$", statement, re.IGNORECASE):
                if blocks and blocks[-1][0] == "if":
                    _, base_indent = blocks[-1]
                    indent = base_indent
                    emit("ELSE")
                    indent += 1
                else:
                    emit("EXECUTE Else [unmatched conditional branch]")
                continue

            condition = re.match(r"^If\s+(.+?)\s+Then$", statement, re.IGNORECASE)
            if condition:
                open_block("if", f"IF {condition.group(1)} THEN")
                continue

            for_match = re.match(r"^For\s+Each\s+(\w+)\s+In\s+(.+)$", statement, re.IGNORECASE)
            if for_match:
                open_block("for", f"FOR EACH {for_match.group(1)} IN {for_match.group(2)}")
                continue
            for_match = re.match(
                r"^For\s+(\w+)\s*=\s*(.+?)\s+To\s+(.+?)(?:\s+Step\s+(.+))?$",
                statement,
                re.IGNORECASE,
            )
            if for_match:
                text = f"FOR {for_match.group(1)} FROM {for_match.group(2)} TO {for_match.group(3)}"
                if for_match.group(4):
                    text += f" STEP {for_match.group(4)}"
                open_block("for", text)
                continue
            next_match = re.match(r"^Next(?:\s+(.+))?$", statement, re.IGNORECASE)
            if next_match:
                variables = next_match.group(1)
                close_count = len(_split_declarations(variables)) if variables else 1
                for _ in range(close_count):
                    close_block("for", "END FOR")
                continue

            while_match = re.match(r"^While\s+(.+)$", statement, re.IGNORECASE)
            if while_match:
                open_block("while", f"WHILE {while_match.group(1)}")
                continue

            do_match = re.match(r"^Do(?:\s+(While|Until)\s+(.+))?$", statement, re.IGNORECASE)
            if do_match:
                mode, condition_text = do_match.groups()
                if mode and mode.lower() == "until":
                    open_block("do_until", f"WHILE NOT ({condition_text})")
                elif mode:
                    open_block("do_while", f"WHILE {condition_text}")
                else:
                    open_block("do", "REPEAT")
                continue

            loop_match = re.match(r"^Loop(?:\s+(While|Until)\s+(.+))?$", statement, re.IGNORECASE)
            if loop_match:
                mode, condition_text = loop_match.groups()
                if blocks and blocks[-1][0] in {"do", "do_while", "do_until"}:
                    if mode and mode.lower() == "while":
                        close_block(blocks[-1][0], f"UNTIL NOT ({condition_text})")
                    elif mode:
                        close_block(blocks[-1][0], f"UNTIL ({condition_text})")
                    elif blocks[-1][0] == "do":
                        close_block("do", "END REPEAT")
                    else:
                        close_block(blocks[-1][0], "END WHILE")
                else:
                    emit("EXECUTE Loop [unmatched loop terminator]")
                continue

            select_match = re.match(r"^Select\s+Case\s+(.+)$", statement, re.IGNORECASE)
            if select_match:
                open_block("select", f"SELECT {select_match.group(1)}")
                continue
            case_match = re.match(r"^Case\s+(.+)$", statement, re.IGNORECASE)
            if case_match and blocks and blocks[-1][0] == "select":
                _, base_indent = blocks[-1]
                indent = base_indent + 1
                case_value = (
                    "OTHERWISE"
                    if case_match.group(1).strip().casefold() == "else"
                    else case_match.group(1)
                )
                emit(f"CASE {case_value}")
                indent = base_indent + 2
                continue
            if re.match(r"^Case\b", statement, re.IGNORECASE):
                emit("EXECUTE Case [outside Select Case]")
                continue

            with_match = re.match(r"^With\s+(.+)$", statement, re.IGNORECASE)
            if with_match:
                open_block("with", f"WITH {with_match.group(1)}")
                continue

            translated = _simple_statement(statement)
            if translated:
                emit(translated)

    while blocks:
        kind, base_indent = blocks.pop()
        indent = base_indent
        emit(f"{_block_end_text(kind)} [implicit: source ended before block terminator]")

    return "\n".join(output)


def _block_end_text(kind: str) -> str:
    return {
        "if": "END IF",
        "for": "END FOR",
        "while": "END WHILE",
        "do_while": "END WHILE",
        "do_until": "END WHILE",
        "do": "END REPEAT",
        "select": "END SELECT",
        "with": "END WITH",
        "procedure": "END PROCEDURE",
    }[kind]


def main(input_path: str | Path, output_path: str | Path | None = None) -> Path:
    """Convert one VBA source file and write its pseudocode to disk."""
    input_path = Path(input_path)
    if not input_path.is_file():
        raise FileNotFoundError(f"VBA input file does not exist: {input_path}")

    if output_path is None:
        output_path = input_path.with_name(f"{input_path.stem}.pseudocode.txt")
    else:
        output_path = Path(output_path)
    if input_path.resolve() == output_path.resolve() or (
        output_path.exists() and input_path.samefile(output_path)
    ):
        raise ValueError("input and output paths must be different")

    pseudocode = convert_vba_to_pseudocode(input_path.read_text(encoding="utf-8-sig"))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(pseudocode + "\n", encoding="utf-8")
    return output_path


# def cli() -> None:
#     parser = argparse.ArgumentParser(
#         description="Convert a VBA source file into readable pseudocode."
#     )
#     parser.add_argument("input", type=Path, help="VBA source file (.bas, .cls, .frm, or text)")
#     parser.add_argument(
#         "-o",
#         "--output",
#         type=Path,
#         help="Output file (defaults to <input-name>.pseudocode.txt)",
#     )
#     args = parser.parse_args()

#     try:
#         output_path = main(args.input, args.output)
#     except (FileNotFoundError, ValueError) as error:
#         parser.error(str(error))
#     print(f"Pseudocode saved to: {output_path}")


if __name__ == "__main__":
    # cli()
    n = [1, 2, 3, 4, 5]
    for i in n:
        main(f"files/vba_sample_{i}.bas", f"files/vba_sample_{i}.pseudo.txt")
