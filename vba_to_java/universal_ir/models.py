"""Typed, language-neutral data contracts for Universal IR.

The runtime representation is JSON-compatible. TypedDicts provide editor/type-checker
support without adding a third-party dependency.
"""
from __future__ import annotations

from typing import Any, Literal, NotRequired, TypedDict


class SourceLocation(TypedDict, total=False):
    file: str
    module: str
    procedure: str
    line: int
    column: int
    end_line: int
    end_column: int


class Provenance(TypedDict, total=False):
    source_language: str
    source_text: str
    source_location: SourceLocation
    graph_node_ids: list[str]
    segment_ids: list[str]
    segment_sources: list[dict[str, Any]]
    confidence: float
    inferred: bool
    inference_reason: str


class TypeSpec(TypedDict, total=False):
    kind: Literal[
        "primitive", "named", "array", "collection", "map", "nullable",
        "variant", "object", "pointer", "unknown", "void"
    ]
    name: str
    element_type: "TypeSpec"
    key_type: "TypeSpec"
    value_type: "TypeSpec"
    dimensions: list[dict[str, Any]]
    nullable: bool
    bounds: dict[str, Any]
    source_type: str
    semantics: dict[str, Any]


class Parameter(TypedDict, total=False):
    id: str
    name: str
    type: TypeSpec
    passing: Literal["value", "reference", "unspecified"]
    optional: bool
    default: "Expression"
    param_array: bool
    modifiers: list[str]
    provenance: Provenance


class Variable(TypedDict, total=False):
    id: str
    name: str
    type: TypeSpec
    scope: Literal["local", "module", "global", "parameter", "field"]
    storage: Literal["constant", "mutable", "static"]
    initial_value: "Expression"
    visibility: Literal["public", "private", "protected", "package", "default"]
    modifiers: list[str]
    provenance: Provenance


class Expression(TypedDict, total=False):
    kind: str
    value: Any
    name: str
    operator: str
    operand: "Expression"
    left: "Expression"
    right: "Expression"
    target: "Expression"
    object: "Expression"
    member: str
    arguments: list["Argument"]
    elements: list["Expression"]
    indices: list["Expression"]
    condition: "Expression"
    when_true: "Expression"
    when_false: "Expression"
    type: TypeSpec
    semantics: dict[str, Any]
    provenance: Provenance


class Argument(TypedDict, total=False):
    value: Expression
    passing: Literal["value", "reference", "omitted", "unspecified"]
    name: str


class Statement(TypedDict, total=False):
    kind: str
    id: str
    expression: Expression
    target: Expression
    exit_target: str
    value: Expression
    statements: list["Statement"]
    body: list["Statement"]
    else_body: list["Statement"]
    branches: list[dict[str, Any]]
    initializer: list["Statement"]
    condition: Expression
    update: list["Statement"]
    variable: str
    iterable: Expression
    start: Expression
    end: Expression
    step: Expression
    label: str
    handler: str
    mode: str
    error_code: Expression
    message: Expression
    dimensions: list[dict[str, Any]]
    declarations: list[Variable]
    arguments: list[Argument]
    provenance: Provenance
    semantics: dict[str, Any]


class Procedure(TypedDict, total=False):
    id: str
    name: str
    kind: Literal["function", "procedure", "constructor", "destructor", "property_get", "property_let", "property_set", "operator"]
    visibility: str
    modifiers: list[str]
    parameters: list[Parameter]
    return_type: TypeSpec
    local_variables: list[Variable]
    body: list[Statement]
    annotations: dict[str, Any]
    provenance: Provenance
    semantics: dict[str, Any]


class Declaration(TypedDict, total=False):
    id: str
    kind: Literal["variable", "constant", "enum", "record", "type_alias", "external", "class", "interface", "property", "unsupported"]
    name: str
    visibility: str
    type: TypeSpec
    members: list[dict[str, Any]]
    value: Expression
    modifiers: list[str]
    provenance: Provenance
    semantics: dict[str, Any]


class Module(TypedDict, total=False):
    id: str
    name: str
    kind: Literal["module", "class", "interface", "record", "package"]
    declarations: list[Declaration]
    procedures: list[Procedure]
    imports: list[dict[str, Any]]
    dependencies: list[dict[str, Any]]
    annotations: dict[str, Any]
    provenance: Provenance
    semantics: dict[str, Any]


class Diagnostic(TypedDict, total=False):
    code: str
    severity: Literal["error", "warning", "info"]
    message: str
    location: SourceLocation
    node_id: str
    suggestion: str


class Program(TypedDict, total=False):
    ir_version: str
    source_language: str
    source_language_version: str
    name: str
    modules: list[Module]
    dependencies: list[dict[str, Any]]
    diagnostics: list[Diagnostic]
    metadata: dict[str, Any]
    semantics: dict[str, Any]
