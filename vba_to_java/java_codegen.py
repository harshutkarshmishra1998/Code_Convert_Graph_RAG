"""Deterministic Java emission from normalized VBA Universal IR."""
from __future__ import annotations

import hashlib
import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any


JAVA_KEYWORDS = {
    "abstract", "assert", "boolean", "break", "byte", "case", "catch", "char",
    "class", "const", "continue", "default", "do", "double", "else", "enum",
    "extends", "final", "finally", "float", "for", "goto", "if", "implements",
    "import", "instanceof", "int", "interface", "long", "native", "new",
    "package", "private", "protected", "public", "return", "short", "static",
    "strictfp", "super", "switch", "synchronized", "this", "throw", "throws",
    "transient", "try", "void", "volatile", "while", "true", "false", "null",
}


def _java_identifier(name: str) -> str:
    value = re.sub(r"[^A-Za-z0-9_$]", "_", name)
    if not value or not re.match(r"[A-Za-z_$]", value):
        value = f"_{value}"
    if value in JAVA_KEYWORDS:
        value = f"_{value}"
    return value


def _java_string(value: str) -> str:
    return json.dumps(value, ensure_ascii=True)


def _walk(value: Any):
    if isinstance(value, dict):
        yield value
        for child in value.values():
            yield from _walk(child)
    elif isinstance(value, list):
        for child in value:
            yield from _walk(child)


class JavaEmitter:
    def __init__(
        self,
        program: dict[str, Any],
        semantic_report: dict[str, Any] | None = None,
        class_name: str | None = None,
    ):
        self.program = program
        self.semantic_report = semantic_report or {}
        self.module = (program.get("modules") or [{}])[0]
        self.module_name = str(self.module.get("name") or program.get("name") or "VbaModule")
        self.class_name = _java_identifier(class_name or self.module_name)
        self.procedures = self.module.get("procedures", [])
        self.procedure_names = {
            str(procedure.get("name", "")).casefold(): _java_identifier(
                str(procedure.get("name", "procedure"))
            )
            for procedure in self.procedures
        }
        self.module_symbols = {
            str(declaration.get("name", "")).casefold(): _java_identifier(
                str(declaration.get("name", "declaration"))
            )
            for declaration in self.module.get("declarations", [])
            if declaration.get("name")
        }
        for declaration in self.module.get("declarations", []):
            if declaration.get("kind") == "enum":
                for member in declaration.get("members", []):
                    if member.get("name"):
                        self.module_symbols[str(member["name"]).casefold()] = _java_identifier(
                            str(member["name"])
                        )
        self.lines: list[str] = []
        self.notes: list[dict[str, Any]] = []
        self.local_names: dict[str, str] = {}
        self.with_receivers: list[str] = []
        self.return_name: str | None = None
        self._temp_index = 0

    def emit(self) -> tuple[str, list[dict[str, Any]]]:
        self.lines = [
            "import java.lang.reflect.Array;",
            "import java.lang.reflect.Method;",
            "import java.math.BigDecimal;",
            "import java.time.LocalDate;",
            "import java.time.LocalDateTime;",
            "import java.time.format.DateTimeFormatter;",
            "import java.util.ArrayList;",
            "import java.util.Arrays;",
            "import java.util.Collections;",
            "import java.util.HashMap;",
            "import java.util.List;",
            "import java.util.Map;",
            "import java.util.Objects;",
            "",
            f"public class {self.class_name} {{",
            "    // Generated from normalized VBA IR. Host/API adapters and VBA compatibility",
            "    // behavior listed in java_codegen_report.json still require implementation.",
        ]
        self._emit_module_declarations()
        for procedure in self.procedures:
            self._emit_procedure(procedure)
        self.lines.extend(_JAVA_RUNTIME.splitlines())
        self.lines.append("}")
        return "\n".join(self.lines) + "\n", self._deduplicate_notes()

    def _line(self, text: str, indent: int = 1) -> None:
        self.lines.append("    " * indent + text)

    def _add_note(self, code: str, message: str, statement: dict[str, Any] | None = None) -> None:
        provenance = (statement or {}).get("provenance") or {}
        self.notes.append({
            "code": code,
            "message": message,
            "procedure": self._current_procedure if hasattr(self, "_current_procedure") else None,
            "source_location": provenance.get("source_location"),
            "source_text": provenance.get("source_text"),
        })

    def _deduplicate_notes(self) -> list[dict[str, Any]]:
        unique = {}
        for note in self.notes:
            key = (note["code"], note["procedure"], str(note["source_location"]), note["message"])
            unique.setdefault(key, note)
        return list(unique.values())

    def _allocate_locals(self, procedure: dict[str, Any]) -> dict[str, str]:
        self.local_names = {}
        used = set()
        names = [
            str(item.get("name", ""))
            for item in procedure.get("parameters", []) + procedure.get("local_variables", [])
            if item.get("name")
        ]
        for statement in _walk(procedure.get("body", [])):
            if statement.get("kind") in {"for", "for_each"} and statement.get("variable"):
                names.append(str(statement["variable"]))
        if procedure.get("kind") == "function" and procedure.get("name"):
            names.append(str(procedure["name"]))
        for name in names:
            key = name.casefold()
            if key in self.local_names:
                continue
            base = _java_identifier(name)
            candidate = base
            suffix = 2
            while candidate.casefold() in used:
                candidate = f"{base}_{suffix}"
                suffix += 1
            used.add(candidate.casefold())
            self.local_names[key] = candidate
        return self.local_names

    def _local(self, name: str) -> str:
        return self.local_names.get(name.casefold(), _java_identifier(name))

    def _emit_module_declarations(self) -> None:
        for declaration in self.module.get("declarations", []):
            kind = declaration.get("kind")
            if kind == "constant" and declaration.get("name"):
                self._line(
                    f"private static final Object {self.module_symbols[str(declaration['name']).casefold()]} "
                    f"= {self.expr(declaration.get('value'))};"
                )
            elif kind == "variable" and declaration.get("name"):
                visibility = "public" if declaration.get("visibility") == "public" else "private"
                name = self.module_symbols[str(declaration["name"]).casefold()]
                self._line(f"{visibility} static Object {name} = null;")
            elif kind == "enum":
                for member in declaration.get("members", []):
                    if member.get("name"):
                        self._line(
                            f"private static final Object {self.module_symbols[str(member['name']).casefold()]} "
                            f"= {self.expr(member.get('value'))};"
                        )
            elif kind in {"record", "type_alias", "unsupported"}:
                self._line(f"// VBA declaration retained in IR: {declaration.get('name', kind)}")
                self._add_note(
                    "MODULE_TYPE_NOT_LOWERED",
                    f"Module declaration {declaration.get('name', kind)!r} needs a target-language type mapping.",
                )

    def _emit_procedure(self, procedure: dict[str, Any]) -> None:
        self._current_procedure = str(procedure.get("name", ""))
        self._allocate_locals(procedure)
        self.return_name = (
            self._local(self._current_procedure)
            if procedure.get("kind") == "function"
            else None
        )
        visibility = "public" if procedure.get("visibility") == "public" else "private"
        return_type = "Object" if procedure.get("kind") == "function" else "void"
        parameters = [
            f"Object {self._local(str(parameter.get('name', f'arg{i + 1}')))}"
            for i, parameter in enumerate(procedure.get("parameters", []))
        ]
        self._line(
            f"{visibility} {return_type} {self.procedure_names[self._current_procedure.casefold()]}"
            f"({', '.join(parameters)}) {{"
        )
        self._line("// ByRef/COM/runtime semantics are documented in the semantic report.", 2)
        parameter_keys = {
            str(parameter.get("name", "")).casefold()
            for parameter in procedure.get("parameters", [])
        }
        declared_keys = {
            str(variable.get("name", "")).casefold()
            for variable in procedure.get("local_variables", [])
            if variable.get("name")
        }
        loop_keys = {
            str(statement.get("variable", "")).casefold()
            for statement in _walk(procedure.get("body", []))
            if statement.get("kind") in {"for", "for_each"} and statement.get("variable")
        }
        for variable in procedure.get("local_variables", []):
            name = str(variable.get("name", ""))
            if name and name.casefold() not in parameter_keys and name.casefold() != self._current_procedure.casefold():
                self._line(f"Object {self._local(name)} = null;", 2)
        for key in sorted(loop_keys - parameter_keys - declared_keys):
            self._line(f"Object {self._local(key)} = null;", 2)
        if self.return_name:
            self._line(f"Object {self.return_name} = null;", 2)
        self._emit_statements(procedure.get("body", []), 2)
        if self.return_name:
            self._line(f"return {self.return_name};", 2)
        self._line("}", 1)
        self.lines.append("")

    def _emit_statements(self, statements: list[dict[str, Any]], indent: int) -> None:
        for statement in statements:
            self._emit_statement(statement, indent)

    def _emit_statement(self, statement: dict[str, Any], indent: int) -> None:
        kind = statement.get("kind")
        if kind == "variable_declaration":
            return
        if kind == "assignment":
            self._emit_assignment(statement, indent)
        elif kind == "if":
            self._emit_if(statement, indent)
        elif kind == "for":
            self._emit_for(statement, indent)
        elif kind == "for_each":
            self._emit_for_each(statement, indent)
        elif kind == "while":
            self._line(f"while (VbaRuntime.truthy({self.expr(statement.get('condition'))})) {{", indent)
            self._emit_statements(statement.get("body", []), indent + 1)
            self._line("}", indent)
        elif kind in {"do_while", "do_until"}:
            self._emit_do(statement, indent)
        elif kind == "switch":
            self._emit_switch(statement, indent)
        elif kind == "with":
            self._emit_with(statement, indent)
        elif kind == "call":
            self._line(f"{self.expr(statement.get('expression'))};", indent)
        elif kind == "exit":
            self._emit_exit(statement, indent)
        elif kind == "return":
            returned = self.expr(statement.get("expression")) if statement.get("expression") else "null"
            self._line(f"if (VbaRuntime.returnNow()) return {returned};", indent)
        elif kind == "on_error":
            mode = _java_string(str(statement.get("mode", "unknown")))
            handler = _java_string(str(statement.get("handler") or ""))
            self._line(f"VbaRuntime.onError({mode}, {handler});", indent)
            self._add_note("VBA_ERROR_HANDLING", "On Error runtime behavior needs a target implementation.", statement)
        elif kind in {"goto", "gosub", "gosub_return"}:
            label = str(statement.get("label") or kind)
            self._line(f"VbaRuntime.unsupportedControlFlow({_java_string(kind + ' ' + label)});", indent)
            self._add_note("VBA_LEGACY_JUMP", f"{kind} control flow is explicitly marked for implementation.", statement)
        elif kind == "label":
            self._line(f"// VBA label: {statement.get('label', '')}", indent)
        elif kind == "redim":
            self._emit_redim(statement, indent)
        elif kind == "erase":
            self._emit_erase(statement, indent)
        elif kind == "raise_error":
            self._line(f"VbaRuntime.raiseError({self.expr(statement.get('error_code'))}, {self.expr(statement.get('message'))});", indent)
        elif kind == "unsupported":
            self._line(f"VbaRuntime.untranslated({_java_string(str(statement.get('source') or statement.get('semantics') or ''))});", indent)
            self._add_note("UNSUPPORTED_STATEMENT", "A statement remains unsupported by the Java emitter.", statement)
        else:
            self._line(f"VbaRuntime.untranslated({_java_string(str(statement.get('kind', 'unknown')))});", indent)
            self._add_note("UNSUPPORTED_STATEMENT_KIND", f"Statement kind {kind!r} is not handled.", statement)

    def _emit_assignment(self, statement: dict[str, Any], indent: int) -> None:
        target = statement.get("target") or {}
        if target.get("kind") == "lvalue":
            target = target.get("expression") or {}
        value = self.expr(statement.get("value"))
        if target.get("kind") == "implicit_function_return_target":
            self._line(f"{self.return_name or 'result'} = {value};", indent)
        elif target.get("kind") == "identifier":
            name = str(target.get("name", ""))
            target_name = self.module_symbols.get(name.casefold(), self._local(name))
            self._line(f"{target_name} = {value};", indent)
        elif target.get("kind") == "member_access":
            receiver = self.expr(target.get("object"))
            self._line(f"VbaRuntime.set({receiver}, {_java_string(str(target.get('member', '')))}, {value});", indent)
        elif target.get("kind") == "index_access":
            indices = ", ".join(self.expr(index) for index in target.get("indices", []))
            self._line(
                f"VbaRuntime.setIndex({self.expr(target.get('object'))}, new Object[]{{{indices}}}, {value});",
                indent,
            )
        else:
            self._line(f"VbaRuntime.untranslated({_java_string(str(target))});", indent)
            self._add_note("UNSUPPORTED_ASSIGNMENT_TARGET", "Assignment target cannot be lowered safely.", statement)

    def _emit_if(self, statement: dict[str, Any], indent: int) -> None:
        branches = statement.get("branches", [])
        if not branches:
            condition = self.expr(statement.get("condition"))
            self._line(f"if (VbaRuntime.truthy({condition})) {{", indent)
            self._emit_statements(statement.get("body", []), indent + 1)
            if statement.get("else_body"):
                self._line("} else {", indent)
                self._emit_statements(statement["else_body"], indent + 1)
            self._line("}", indent)
            return
        emitted_branch = False
        for branch in branches:
            if branch.get("branch_kind") == "else_branch":
                self._line("} else {" if emitted_branch else "{", indent)
            else:
                keyword = "if" if not emitted_branch else "} else if"
                self._line(
                    f"{keyword} (VbaRuntime.truthy({self.expr(branch.get('condition'))})) {{",
                    indent,
                )
            self._emit_statements(branch.get("body", []), indent + 1)
            emitted_branch = True
        self._line("}", indent)

    def _emit_for(self, statement: dict[str, Any], indent: int) -> None:
        variable = self._local(str(statement.get("variable", "index")))
        step = self.expr(statement.get("step") or {"kind": "literal", "value": 1})
        self._line(f"{variable} = {self.expr(statement.get('start'))};", indent)
        self._line(
            f"for (; VbaRuntime.forCondition({variable}, {self.expr(statement.get('end'))}, {step}); "
            f"{variable} = VbaRuntime.binary(\"+\", {variable}, {step})) {{",
            indent,
        )
        self._emit_statements(statement.get("body", []), indent + 1)
        self._line("}", indent)

    def _emit_for_each(self, statement: dict[str, Any], indent: int) -> None:
        variable = self._local(str(statement.get("variable", "item")))
        temporary = self._temp("item")
        self._line(f"for (Object {temporary} : VbaRuntime.iterable({self.expr(statement.get('iterable'))})) {{", indent)
        self._line(f"{variable} = {temporary};", indent + 1)
        self._emit_statements(statement.get("body", []), indent + 1)
        self._line("}", indent)

    def _emit_do(self, statement: dict[str, Any], indent: int) -> None:
        condition = self.expr(statement.get("condition"))
        until = statement.get("kind") == "do_until"
        test = f"!VbaRuntime.truthy({condition})" if until else f"VbaRuntime.truthy({condition})"
        test_position = (statement.get("semantics") or {}).get("test_position")
        if not statement.get("condition"):
            self._line("while (true) {", indent)
            self._emit_statements(statement.get("body", []), indent + 1)
            self._line("}", indent)
        elif test_position == "post":
            self._line("do {", indent)
            self._emit_statements(statement.get("body", []), indent + 1)
            self._line(f"}} while ({test});", indent)
        else:
            self._line(f"while ({test}) {{", indent)
            self._emit_statements(statement.get("body", []), indent + 1)
            self._line("}", indent)

    def _emit_switch(self, statement: dict[str, Any], indent: int) -> None:
        value = self._temp("select")
        self._line(f"Object {value} = {self.expr(statement.get('expression'))};", indent)
        branches = statement.get("branches", [])
        branches = [branch for branch in branches if branch.get("branch_kind") != "case_else"] + [
            branch for branch in branches if branch.get("branch_kind") == "case_else"
        ]
        emitted_branch = False
        for branch in branches:
            if branch.get("branch_kind") == "case_else":
                self._line("} else {" if emitted_branch else "{", indent)
            else:
                selected = statement.get("expression") or {}
                if selected.get("kind") == "literal" and selected.get("value") is True:
                    predicate = f"VbaRuntime.truthy({self.expr(branch.get('condition'))})"
                else:
                    predicate = f"VbaRuntime.equal({value}, {self.expr(branch.get('condition'))})"
                keyword = "if" if not emitted_branch else "} else if"
                self._line(f"{keyword} ({predicate}) {{", indent)
            self._emit_statements(branch.get("body", []), indent + 1)
            emitted_branch = True
        self._line("}", indent)

    def _emit_with(self, statement: dict[str, Any], indent: int) -> None:
        receiver = self._temp("with")
        self._line(f"Object {receiver} = {self.expr(statement.get('expression'))};", indent)
        self.with_receivers.append(receiver)
        self._line("{", indent)
        self._emit_statements(statement.get("body", []), indent + 1)
        self._line("}", indent)
        self.with_receivers.pop()

    def _emit_exit(self, statement: dict[str, Any], indent: int) -> None:
        target = str(statement.get("exit_target", "")).casefold()
        if target == "function":
            self._line(f"if (VbaRuntime.returnNow()) return {self.return_name or 'null'};", indent)
        elif target in {"sub", "property"}:
            self._line("if (VbaRuntime.returnNow()) return;", indent)
        elif target in {"for", "do"}:
            self._line("if (VbaRuntime.returnNow()) break;", indent)
        else:
            self._line(f"VbaRuntime.unsupportedControlFlow({_java_string('Exit ' + target)});", indent)
            self._add_note("UNSUPPORTED_EXIT", f"Exit {target} needs review.", statement)

    def _emit_redim(self, statement: dict[str, Any], indent: int) -> None:
        target = statement.get("target") or {}
        expression = target.get("expression") if target.get("kind") == "lvalue" else target
        if isinstance(expression, dict) and expression.get("kind") == "identifier":
            name = str(expression.get("name", ""))
            dimensions = [self.expr(item.get("upper")) for item in statement.get("dimensions", [])]
            array = f"VbaRuntime.redim({self._local(name)}, new Object[]{{{', '.join(dimensions)}}})"
            self._line(f"{self._local(name)} = {array};", indent)
        else:
            self._line("VbaRuntime.unsupportedControlFlow(\"ReDim target\");", indent)
            self._add_note("UNSUPPORTED_REDIM_TARGET", "ReDim target is not a local array variable.", statement)

    def _emit_erase(self, statement: dict[str, Any], indent: int) -> None:
        target = statement.get("target") or {}
        expression = target.get("expression") if target.get("kind") == "lvalue" else target
        if isinstance(expression, dict) and expression.get("kind") == "identifier":
            self._line(f"{self._local(str(expression.get('name', '')))} = null;", indent)
        else:
            self._line("VbaRuntime.unsupportedControlFlow(\"Erase target\");", indent)
            self._add_note("UNSUPPORTED_ERASE_TARGET", "Erase target is not a local array variable.", statement)

    def _temp(self, stem: str) -> str:
        self._temp_index += 1
        return f"_vba_{stem}_{self._temp_index}"

    def expr(self, expression: Any) -> str:
        if not isinstance(expression, dict):
            return "null"
        kind = expression.get("kind")
        if kind == "literal":
            value = expression.get("value")
            if isinstance(value, str):
                if expression.get("literal_type") == "date":
                    return f"VbaRuntime.date({_java_string(value)})"
                return _java_string(value)
            if value is None:
                return "null"
            if isinstance(value, bool):
                return "Boolean.TRUE" if value else "Boolean.FALSE"
            return str(value)
        if kind == "identifier":
            name = str(expression.get("name", ""))
            if (expression.get("semantics") or {}).get("implicit_with_receiver"):
                return self.with_receivers[-1] if self.with_receivers else "VbaRuntime.untranslatedExpression(\"With receiver\")"
            lowered = name.casefold()
            if lowered in self.local_names:
                return self._local(name)
            if lowered in self.module_symbols:
                return self.module_symbols[lowered]
            if lowered in self.procedure_names:
                return self.procedure_names[lowered]
            if lowered == "err":
                return "VbaRuntime.errorState()"
            if lowered in {
                "thisworkbook", "application", "workbooks", "activeworkbook",
                "activesheet", "worksheet", "range", "cells",
            }:
                self._add_note(
                    "VBA_HOST_ADAPTER_REQUIRED",
                    f"VBA host object {name!r} requires an application-specific Java adapter.",
                )
                return f"VbaRuntime.host({_java_string(name)})"
            return f"VbaRuntime.name({_java_string(name)})"
        if kind == "self_reference":
            return "this"
        if kind == "member_access":
            obj = expression.get("object") or {}
            if obj.get("kind") == "identifier" and (obj.get("semantics") or {}).get("implicit_with_receiver"):
                receiver = self.with_receivers[-1] if self.with_receivers else "null"
            else:
                receiver = self.expr(obj)
            return f"VbaRuntime.get({receiver}, {_java_string(str(expression.get('member', '')))} )"
        if kind == "invocation_expression":
            arguments = ", ".join(self._argument(arg) for arg in expression.get("arguments", []))
            args = f"new Object[]{{{arguments}}}"
            callee = expression.get("callee") or {}
            target = expression.get("target_procedure") or {}
            if target.get("name"):
                method = self.procedure_names.get(str(target["name"]).casefold(), _java_identifier(str(target["name"])))
                return f"this.{method}({arguments})"
            if expression.get("invocation_kind") == "vba_builtin":
                builtin = str(expression.get("builtin_name") or (callee.get("name") if isinstance(callee, dict) else ""))
                if builtin.casefold().replace("$", "") == "createobject":
                    self._add_note(
                        "VBA_COM_ADAPTER_REQUIRED",
                        "CreateObject requires a Java adapter for the requested COM server.",
                    )
                return f"VbaRuntime.builtin({_java_string(builtin)}, {args})"
            if callee.get("kind") == "member_access":
                receiver = self.expr(callee.get("object"))
                member = str(callee.get("member", ""))
                return f"VbaRuntime.call({receiver}, {_java_string(member)}, {args})"
            if callee.get("kind") == "identifier":
                return f"VbaRuntime.call(null, {_java_string(str(callee.get('name', '')))}, {args})"
            return f"VbaRuntime.call({self.expr(callee)}, \"invoke\", {args})"
        if kind == "binary_expression":
            return (
                f"VbaRuntime.binary({_java_string(str(expression.get('operator', '')))}, "
                f"{self.expr(expression.get('left'))}, {self.expr(expression.get('right'))})"
            )
        if kind == "unary_expression":
            return f"VbaRuntime.unary({_java_string(str(expression.get('operator', '')))}, {self.expr(expression.get('operand'))})"
        if kind == "parenthesized_expression":
            return f"({self.expr(expression.get('expression'))})"
        if kind == "index_access":
            indices = ", ".join(self.expr(index) for index in expression.get("indices", []))
            return f"VbaRuntime.index({self.expr(expression.get('object'))}, new Object[]{{{indices}}})"
        if kind == "collection_item_access":
            indices = ", ".join(self.expr(index) for index in expression.get("key_or_index", []))
            return f"VbaRuntime.index({self.expr(expression.get('collection'))}, new Object[]{{{indices}}})"
        if kind == "named_argument":
            return self._argument(expression)
        if kind == "new_object_expression":
            self._add_note("LATE_BOUND_OBJECT_CREATION", "VBA object construction requires a target-specific adapter.")
            return f"VbaRuntime.newObject({_java_string(str(expression.get('type_name', 'Object')))})"
        if kind == "implicit_function_return_target":
            return self.return_name or "null"
        if kind in {"unresolved_expression", "unresolved_lvalue", "unresolved_call", "unsupported"}:
            self._add_note("UNRESOLVED_EXPRESSION", f"Expression kind {kind!r} remains unresolved.")
            raw = expression.get("source_expression") or expression.get("source") or kind
            return f"VbaRuntime.untranslatedExpression({_java_string(str(raw))})"
        if kind in {"nothing", "missing"}:
            return "null"
        self._add_note("UNSUPPORTED_EXPRESSION_KIND", f"Expression kind {kind!r} needs review.")
        return f"VbaRuntime.untranslatedExpression({_java_string(str(expression.get('source', kind)))})"

    def _argument(self, argument: Any) -> str:
        if not isinstance(argument, dict):
            return self.expr(argument)
        if argument.get("kind") == "named_argument":
            return (
                f"new VbaRuntime.NamedArg({_java_string(str(argument.get('name', '')))}, "
                f"{self.expr(argument.get('value'))})"
            )
        if argument.get("kind") == "missing":
            return "VbaRuntime.MISSING"
        return self.expr(argument)


_JAVA_RUNTIME = r'''
    private static final class VbaRuntime {
        static final Object MISSING = new Object();
        private static final ErrorState ERR = new ErrorState();

        static final class NamedArg {
            final String name;
            final Object value;
            NamedArg(String name, Object value) { this.name = name; this.value = value; }
        }

        static final class HostObject {
            final String name;
            HostObject(String name) { this.name = name; }
            @Override public String toString() { return name; }
        }

        static final class ErrorState {
            int number;
            String description = "";
            void clear() { number = 0; description = ""; }
            void raise(Object number, Object source, Object description) {
                this.number = (int) VbaRuntime.number(number);
                this.description = String.valueOf(description);
                throw new IllegalStateException(this.description);
            }
        }

        static Object errorState() { return ERR; }
        static Object host(String name) { return new HostObject(name); }
        static Object name(String name) {
            return switch (name.toLowerCase()) {
                case "true" -> Boolean.TRUE;
                case "false" -> Boolean.FALSE;
                case "nothing", "null", "empty" -> null;
                case "vbcrlf", "vbnewline" -> "\r\n";
                case "vbcr" -> "\r";
                case "vblf" -> "\n";
                case "vbtab" -> "\t";
                case "vbnullstring" -> "";
                case "vbtextcompare" -> 1;
                case "vbbinarycompare" -> 0;
                case "vbobjecterror" -> -2147221504;
                case "xlup" -> -4162;
                case "xlyes" -> 1;
                case "xlnone" -> -4142;
                case "xlopenxmlworkbook" -> 51;
                case "vbyellow" -> 6;
                case "vbinformation" -> 64;
                case "vbexclamation" -> 48;
                case "vbcritical" -> 16;
                case "now" -> LocalDateTime.now();
                case "timer" -> LocalDateTime.now().toLocalTime().toSecondOfDay();
                default -> new HostObject(name);
            };
        }

        static Object get(Object target, String member) {
            if (target == null) return null;
            if (target == ERR) {
                return switch (member.toLowerCase()) {
                    case "number" -> ERR.number;
                    case "description" -> ERR.description;
                    default -> null;
                };
            }
            if (target instanceof HostObject host) return new HostObject(host.name + "." + member);
            if (target instanceof Map<?, ?> map) return map.get(member);
            String suffix = Character.toUpperCase(member.charAt(0)) + member.substring(1);
            for (String candidate : new String[]{member, "get" + suffix, "is" + suffix}) {
                try {
                    Method method = target.getClass().getMethod(candidate);
                    return method.invoke(target);
                } catch (ReflectiveOperationException ignored) { }
            }
            try { return target.getClass().getField(member).get(target); }
            catch (ReflectiveOperationException ignored) {
                throw new IllegalStateException("Cannot read VBA member " + member + " from " + target.getClass().getName());
            }
        }

        static Object set(Object target, String member, Object value) {
            if (target == null || target instanceof HostObject)
                throw new UnsupportedOperationException("VBA host adapter required for member " + member);
            if (target instanceof Map<?, ?> raw) {
                @SuppressWarnings("unchecked") Map<Object, Object> map = (Map<Object, Object>) raw;
                map.put(member, value);
                return value;
            }
            String setter = "set" + Character.toUpperCase(member.charAt(0)) + member.substring(1);
            for (Method method : target.getClass().getMethods()) {
                if (method.getName().equalsIgnoreCase(setter) && method.getParameterCount() == 1) {
                    try { method.invoke(target, value); return value; }
                    catch (ReflectiveOperationException exception) { throw new IllegalStateException(exception); }
                }
            }
            throw new IllegalStateException("Cannot write VBA member " + member + " on " + target.getClass().getName());
        }

        static Object call(Object target, String member, Object[] arguments) {
            if (target == ERR && member.equalsIgnoreCase("raise")) {
                Object[] values = unwrap(arguments);
                ERR.raise(
                    values.length > 0 ? values[0] : 0,
                    values.length > 1 ? values[1] : "",
                    values.length > 2 ? values[2] : ""
                );
                return null;
            }
            if (target == ERR && member.equalsIgnoreCase("clear")) {
                ERR.clear();
                return null;
            }
            if (target == null || target instanceof HostObject)
                throw new UnsupportedOperationException("VBA host/API adapter required for " + member);
            Object[] values = unwrap(arguments);
            for (Method method : target.getClass().getMethods()) {
                if (!method.getName().equalsIgnoreCase(member) || method.getParameterCount() != values.length) continue;
                try { return method.invoke(target, values); }
                catch (ReflectiveOperationException exception) { throw new IllegalStateException(exception); }
            }
            throw new IllegalStateException("Cannot invoke VBA member " + member + " on " + target.getClass().getName());
        }

        static Object[] unwrap(Object[] arguments) {
            Object[] values = new Object[arguments.length];
            for (int i = 0; i < arguments.length; i++)
                values[i] = arguments[i] instanceof NamedArg named ? named.value : arguments[i];
            return values;
        }

        static Object builtin(String rawName, Object[] arguments) {
            Object[] args = unwrap(arguments);
            Object first = args.length == 0 ? null : args[0];
            String name = rawName.toLowerCase().replace("$", "");
            return switch (name) {
                case "cstr", "str" -> String.valueOf(first);
                case "len" -> String.valueOf(first == null ? "" : first).length();
                case "trim" -> String.valueOf(first == null ? "" : first).trim();
                case "ltrim" -> String.valueOf(first == null ? "" : first).replaceFirst("^\\s+", "");
                case "rtrim" -> String.valueOf(first == null ? "" : first).replaceFirst("\\s+$", "");
                case "ucase" -> String.valueOf(first == null ? "" : first).toUpperCase();
                case "lcase" -> String.valueOf(first == null ? "" : first).toLowerCase();
                case "isnumeric" -> isNumeric(first);
                case "isdate" -> canParseDate(first);
                case "isarray" -> first != null && (first.getClass().isArray() || first instanceof List<?>);
                case "isnull", "isempty" -> first == null;
                case "cbool" -> truthy(first);
                case "clng", "cint", "cbyte" -> (long) number(first);
                case "cdbl", "csng", "val", "abs" -> name.equals("abs") ? Math.abs(number(first)) : number(first);
                case "ccur", "cdec" -> new BigDecimal(String.valueOf(first));
                case "cdate", "datevalue" -> parseDate(first);
                case "split" -> Arrays.asList(String.valueOf(first).split(
                    args.length > 1 ? java.util.regex.Pattern.quote(String.valueOf(args[1])) : "\\s+", -1));
                case "join" -> String.join(args.length > 1 ? String.valueOf(args[1]) : " ",
                    iterable(args.length == 0 ? null : args[0]).stream().map(String::valueOf).toList());
                case "array" -> args;
                case "ubound" -> upperBound(first);
                case "lbound" -> 0;
                case "replace" -> String.valueOf(first).replace(
                    args.length > 1 ? String.valueOf(args[1]) : "",
                    args.length > 2 ? String.valueOf(args[2]) : "");
                case "left" -> String.valueOf(first).substring(0,
                    Math.min(String.valueOf(first).length(), (int) number(args.length > 1 ? args[1] : 0)));
                case "right" -> {
                    String text = String.valueOf(first);
                    int count = Math.min(text.length(), (int) number(args.length > 1 ? args[1] : 0));
                    yield text.substring(text.length() - count);
                }
                case "mid" -> {
                    String text = String.valueOf(first);
                    int start = Math.max(0, (int) number(args.length > 1 ? args[1] : 1) - 1);
                    int end = args.length > 2 ? Math.min(text.length(), start + (int) number(args[2])) : text.length();
                    yield start >= text.length() ? "" : text.substring(start, end);
                }
                case "iif" -> args.length > 1 && truthy(args[0]) ? args[1] : args.length > 2 ? args[2] : null;
                case "now" -> LocalDateTime.now();
                case "date" -> java.time.LocalDate.now().toString();
                case "year" -> parseDate(first).getYear();
                case "strcomp" -> {
                    String left = String.valueOf(first);
                    String right = args.length > 1 ? String.valueOf(args[1]) : "";
                    boolean textCompare = args.length > 2 && number(args[2]) == 1;
                    yield Integer.signum(textCompare
                        ? left.compareToIgnoreCase(right)
                        : left.compareTo(right));
                }
                case "format" -> format(first, args.length > 1 ? String.valueOf(args[1]) : "");
                case "doevents" -> null;
                case "msgbox" -> { System.out.println(first); yield 0; }
                case "createobject" -> new HostObject("COM:" + String.valueOf(first));
                default -> throw new UnsupportedOperationException("VBA builtin requires implementation: " + rawName);
            };
        }

        static Object binary(String operator, Object left, Object right) {
            return switch (operator.toLowerCase()) {
                case "&" -> String.valueOf(left == null ? "" : left) + String.valueOf(right == null ? "" : right);
                case "+" -> left instanceof String || right instanceof String
                    ? String.valueOf(left) + String.valueOf(right) : number(left) + number(right);
                case "-" -> number(left) - number(right);
                case "*" -> number(left) * number(right);
                case "/" -> number(left) / number(right);
                case "\\" -> (long) (number(left) / number(right));
                case "mod" -> number(left) % number(right);
                case "^" -> Math.pow(number(left), number(right));
                case "=" -> equal(left, right);
                case "<>" -> !equal(left, right);
                case "is" -> left == right;
                case "is_not" -> left != right;
                case "<" -> compare(left, right) < 0;
                case ">" -> compare(left, right) > 0;
                case "<=" -> compare(left, right) <= 0;
                case ">=" -> compare(left, right) >= 0;
                case "and" -> left instanceof Number && right instanceof Number
                    ? ((long) number(left) & (long) number(right)) : truthy(left) && truthy(right);
                case "or" -> left instanceof Number && right instanceof Number
                    ? ((long) number(left) | (long) number(right)) : truthy(left) || truthy(right);
                case "xor" -> left instanceof Number && right instanceof Number
                    ? ((long) number(left) ^ (long) number(right)) : truthy(left) ^ truthy(right);
                case "like" -> String.valueOf(left).matches(
                    java.util.regex.Pattern.quote(String.valueOf(right)).replace("*", "\\E.*\\Q").replace("?", "\\E.\\Q"));
                default -> throw new UnsupportedOperationException("VBA operator requires implementation: " + operator);
            };
        }

        static Object unary(String operator, Object value) {
            return switch (operator.toLowerCase()) {
                case "+" -> number(value);
                case "-" -> -number(value);
                case "not" -> value instanceof Number ? ~((long) number(value)) : !truthy(value);
                default -> throw new UnsupportedOperationException("VBA unary operator requires implementation: " + operator);
            };
        }

        static boolean truthy(Object value) {
            if (value == null || value == MISSING) return false;
            if (value instanceof Boolean bool) return bool;
            if (value instanceof Number number) return number.doubleValue() != 0;
            if (value instanceof String text) return !text.isEmpty();
            return true;
        }

        static boolean equal(Object left, Object right) {
            if (left instanceof Number && right instanceof Number)
                return Double.compare(number(left), number(right)) == 0;
            return Objects.equals(left, right);
        }

        static int compare(Object left, Object right) {
            if (left instanceof Number && right instanceof Number)
                return Double.compare(number(left), number(right));
            return String.valueOf(left).compareTo(String.valueOf(right));
        }

        static double number(Object value) {
            if (value instanceof Number number) return number.doubleValue();
            if (value instanceof Boolean bool) return bool ? -1 : 0;
            if (value == null || String.valueOf(value).isBlank()) return 0;
            try { return Double.parseDouble(String.valueOf(value)); }
            catch (NumberFormatException exception) { throw new IllegalArgumentException("Not numeric: " + value, exception); }
        }

        static boolean isNumeric(Object value) {
            try { number(value); return value != null; }
            catch (RuntimeException exception) { return false; }
        }

        static boolean forCondition(Object current, Object end, Object step) {
            return number(step) >= 0 ? number(current) <= number(end) : number(current) >= number(end);
        }

        static List<Object> iterable(Object value) {
            if (value instanceof List<?> list) return new ArrayList<>(list);
            if (value instanceof Iterable<?> iterable) {
                List<Object> result = new ArrayList<>();
                iterable.forEach(result::add);
                return result;
            }
            if (value != null && value.getClass().isArray()) {
                List<Object> result = new ArrayList<>();
                for (int i = 0; i < Array.getLength(value); i++) result.add(Array.get(value, i));
                return result;
            }
            return Collections.emptyList();
        }

        static Object index(Object value, Object[] indices) {
            int index = indices.length == 0 ? 0 : (int) number(indices[0]);
            if (value instanceof List<?> list) return list.get(index);
            if (value != null && value.getClass().isArray()) return Array.get(value, index);
            if (value instanceof Map<?, ?> map) return map.get(indices.length == 0 ? null : indices[0]);
            throw new IllegalStateException("Cannot index " + value);
        }

        static Object setIndex(Object value, Object[] indices, Object item) {
            int index = indices.length == 0 ? 0 : (int) number(indices[0]);
            if (value instanceof List<?> raw) {
                @SuppressWarnings("unchecked") List<Object> list = (List<Object>) raw;
                while (list.size() <= index) list.add(null);
                list.set(index, item);
                return item;
            }
            if (value != null && value.getClass().isArray()) { Array.set(value, index, item); return item; }
            if (value instanceof Map<?, ?> raw && indices.length > 0) {
                @SuppressWarnings("unchecked") Map<Object, Object> map = (Map<Object, Object>) raw;
                map.put(indices[0], item);
                return item;
            }
            throw new IllegalStateException("Cannot assign indexed value " + value);
        }

        static int upperBound(Object value) {
            if (value instanceof List<?> list) return list.size() - 1;
            if (value != null && value.getClass().isArray()) return Array.getLength(value) - 1;
            return -1;
        }

        static Object redim(Object value, Object[] bounds) {
            int length = bounds.length == 0 ? 0 : Math.max(0, (int) number(bounds[bounds.length - 1]) + 1);
            Object[] result = new Object[length];
            if (value != null && value.getClass().isArray())
                System.arraycopy(value, 0, result, 0, Math.min(Array.getLength(value), result.length));
            return result;
        }

        static Object date(String value) { return value; }
        static LocalDateTime parseDate(Object value) {
            if (value instanceof LocalDateTime dateTime) return dateTime;
            if (value instanceof LocalDate date) return date.atStartOfDay();
            String text = String.valueOf(value);
            for (DateTimeFormatter formatter : new DateTimeFormatter[]{
                    DateTimeFormatter.ISO_LOCAL_DATE_TIME,
                    DateTimeFormatter.ISO_LOCAL_DATE,
                    DateTimeFormatter.ofPattern("M/d/uuuu"),
                    DateTimeFormatter.ofPattern("M/d/uuuu H:mm:ss")}) {
                try {
                    try { return LocalDateTime.parse(text, formatter); }
                    catch (java.time.format.DateTimeParseException ignored) {
                        return LocalDate.parse(text, formatter).atStartOfDay();
                    }
                } catch (java.time.format.DateTimeParseException ignored) { }
            }
            throw new IllegalArgumentException("Unsupported VBA date value: " + value);
        }
        static boolean canParseDate(Object value) {
            if (value == null) return false;
            try { parseDate(value); return true; }
            catch (RuntimeException exception) { return false; }
        }
        static String format(Object value, String vbaPattern) {
            if (!(value instanceof LocalDateTime)) return String.valueOf(value);
            String pattern = vbaPattern.replace("nn", "__MINUTE__")
                .replace("mm", "MM").replace("__MINUTE__", "mm").replace("hh", "HH");
            try { return ((LocalDateTime) value).format(DateTimeFormatter.ofPattern(pattern)); }
            catch (IllegalArgumentException exception) { return String.valueOf(value); }
        }
        static Object newObject(String typeName) { return new HostObject(typeName); }
        static boolean returnNow() { return true; }
        static void onError(String mode, String handler) { }
        static void raiseError(Object code, Object message) {
            throw new IllegalStateException(String.valueOf(message));
        }
        static void unsupportedControlFlow(String description) {
            throw new UnsupportedOperationException("Translate VBA control flow: " + description);
        }
        static Object untranslatedExpression(String source) {
            throw new UnsupportedOperationException("Translate VBA expression: " + source);
        }
        static void untranslated(String source) {
            throw new UnsupportedOperationException("Translate VBA statement: " + source);
        }
    }
'''


def _compile(java_path: Path) -> tuple[bool | None, str]:
    javac = shutil.which("javac")
    if not javac:
        return None, "javac is not installed; source compilation was skipped."
    with tempfile.TemporaryDirectory(prefix="vba-java-check-") as directory:
        result = subprocess.run(
            [javac, "-d", directory, str(java_path)],
            capture_output=True,
            text=True,
            timeout=60,
            check=False,
        )
    output = "\n".join(part for part in (result.stdout.strip(), result.stderr.strip()) if part)
    return result.returncode == 0, output


def generate_java(input_path: str | Path, output: str | Path) -> tuple[Path, Path, dict[str, Any]]:
    source_path = Path(input_path).expanduser().resolve()
    sample_dir = Path(output).expanduser().resolve()
    if not source_path.is_file():
        raise FileNotFoundError(f"VBA source file does not exist: {source_path}")
    ir_path = sample_dir / "normalized_ir.json"
    if not ir_path.is_file():
        raise FileNotFoundError(f"Normalized IR does not exist: {ir_path}")
    program = json.loads(ir_path.read_text(encoding="utf-8"))
    modules = program.get("modules", [])
    if len(modules) != 1:
        raise ValueError(f"Java emitter expects one VBA module; found {len(modules)}")
    if program.get("name") and program["name"] != source_path.stem:
        raise ValueError(f"Input {source_path.name} does not match IR project {program['name']!r}")
    manifest_path = sample_dir / "manifest.json"
    if not manifest_path.is_file():
        raise FileNotFoundError(f"Source manifest does not exist: {manifest_path}")
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if not isinstance(manifest, list):
        raise ValueError("manifest.json must contain a source-file list")
    source_record = next(
        (entry for entry in manifest if Path(str(entry.get("path", ""))).name == source_path.name),
        None,
    )
    if source_record is None:
        raise ValueError(f"manifest.json has no entry for {source_path.name}")
    source_hash = hashlib.sha256(source_path.read_bytes()).hexdigest()
    if source_record.get("sha256") != source_hash:
        raise ValueError(f"Source hash mismatch for {source_path.name}; regenerate analysis artifacts first")
    semantic_path = sample_dir / "semantic_analysis.json"
    semantic = json.loads(semantic_path.read_text(encoding="utf-8")) if semantic_path.is_file() else {}
    emitter = JavaEmitter(program, semantic, class_name=source_path.stem)
    java_source, notes = emitter.emit()
    java_path = sample_dir / f"{_java_identifier(source_path.stem)}.java"
    java_path.write_text(java_source, encoding="utf-8")
    compiled, compiler_output = _compile(java_path)
    report = {
        "schema_version": "1.0",
        "source_file": source_path.name,
        "source_sha256": source_hash,
        "source_language": "VBA",
        "target_language": "Java",
        "java_file": java_path.name,
        "procedure_count": sum(len(module.get("procedures", [])) for module in program.get("modules", [])),
        "semantic_status": semantic.get("status", "not_available"),
        "compatibility_requirements": semantic.get("compatibility_requirements", []),
        "generated_with_runtime_requirements": bool(notes or semantic.get("status") == "blocked"),
        "compile_status": "passed" if compiled is True else "failed" if compiled is False else "skipped",
        "compiler_output": compiler_output,
        "limitations": notes,
        "note": "Compilation verifies Java syntax only, not behavioral equivalence or availability of VBA host adapters.",
    }
    report_path = sample_dir / "java_codegen_report.json"
    report_path.write_text(json.dumps(report, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    return java_path, report_path, report


def main(input_path: str | Path, output: str | Path) -> int:
    try:
        java_path, report_path, report = generate_java(input_path, output)
        print(json.dumps({
            "java_file": str(java_path),
            "report": str(report_path),
            "compile_status": report["compile_status"],
            "runtime_requirements": len(report["limitations"]),
        }, indent=2))
        return 0 if report["compile_status"] in {"passed", "skipped"} else 1
    except Exception as exc:
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    project_root = Path(__file__).resolve().parents[1]
    output_root = project_root / "files" / "vba_graph_output"
    n = [1, 2, 3, 4, 5, 6]
    status = 0

    for i in n:
        input_path = project_root / "files" / f"vba_sample_{i}.bas"
        output = output_root / f"sample_{i}"
        status = max(status, main(input_path, output))

    raise SystemExit(status)
