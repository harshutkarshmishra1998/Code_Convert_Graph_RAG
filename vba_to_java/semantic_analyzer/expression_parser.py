"""Deterministic parser for the VBA expression subset found in legacy projects.

Unsupported syntax is returned as an explicit unresolved_expression node; the
parser never substitutes guessed semantics.
"""
from __future__ import annotations
import re
from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class Token:
    kind: str
    value: str
    pos: int


_TOKEN_RE = re.compile(
    r'''(?P<SPACE>\s+)
    |(?P<STRING>"(?:""|[^"])*")
    |(?P<DATE>\#[^\#]*\#)
    |(?P<HEX>&[Hh][0-9A-Fa-f]+)
    |(?P<OCT>&[Oo][0-7]+)
    |(?P<NUMBER>(?:\d+\.\d*|\.\d+|\d+)(?:[Ee][+-]?\d+)?[!#@%&]?(?![A-Za-z0-9_]))
    |(?P<IDENT>[A-Za-z_][A-Za-z0-9_]*[$%&#!]?)
    |(?P<OP><=|>=|<>|:=|[+\-*/\\^=<>(),.&])
    |(?P<OTHER>.)''', re.X
)

# VBA precedence, low to high. Exponentiation is right-associative.
_PRECEDENCE = {
    'imp': 1, 'eqv': 2, 'xor': 3, 'or': 4, 'and': 5,
    '=': 6, '<>': 6, '<': 6, '>': 6, '<=': 6, '>=': 6, 'is': 6, 'like': 6,
    '&': 7, '+': 8, '-': 8, '*': 9, '/': 9, '\\': 9, 'mod': 9, '^': 10,
}


class ParseError(ValueError):
    pass


def tokenize(source: str) -> list[Token]:
    out: list[Token] = []
    for m in _TOKEN_RE.finditer(source):
        kind, value = m.lastgroup, m.group()
        if kind == 'SPACE':
            continue
        if kind == 'OTHER':
            raise ParseError(f"Unexpected character {value!r} at offset {m.start()}")
        out.append(Token(kind or '', value, m.start()))
    out.append(Token('EOF', '', len(source)))
    return out


class Parser:
    def __init__(self, source: str):
        self.source = source
        self.tokens = tokenize(source)
        self.i = 0

    @property
    def current(self) -> Token:
        return self.tokens[self.i]

    def take(self) -> Token:
        t = self.current
        self.i += 1
        return t

    def accept(self, value: str) -> bool:
        if self.current.value.lower() == value.lower():
            self.take(); return True
        return False

    def expect(self, value: str) -> Token:
        if not self.accept(value):
            raise ParseError(f"Expected {value!r} at offset {self.current.pos}")
        return self.tokens[self.i - 1]

    def parse(self) -> dict[str, Any]:
        node = self.expression(0)
        if self.current.kind != 'EOF':
            raise ParseError(f"Unexpected token {self.current.value!r} at offset {self.current.pos}")
        node.setdefault('source', self.source.strip())
        return node

    def expression(self, min_prec: int) -> dict[str, Any]:
        left = self.prefix()
        left = self.postfix(left)
        while self.current.kind in {'OP', 'IDENT'}:
            op = self.current.value.lower()
            if op not in _PRECEDENCE or _PRECEDENCE[op] < min_prec:
                break
            prec = _PRECEDENCE[op]
            self.take()
            # `Is Not` is one comparison operator in VBA.
            if op == 'is' and self.accept('not'):
                op = 'is_not'
            right = self.expression(prec if op == '^' else prec + 1)
            left = {'kind': 'binary_expression', 'operator': op, 'left': left, 'right': right}
        return left

    def prefix(self) -> dict[str, Any]:
        t = self.current
        low = t.value.lower()
        if t.kind == 'OP' and t.value in {'+', '-'}:
            self.take(); return {'kind': 'unary_expression', 'operator': t.value, 'operand': self.expression(11)}
        if t.kind == 'IDENT' and low in {'not'}:
            self.take(); return {'kind': 'unary_expression', 'operator': low, 'operand': self.expression(6)}
        if t.kind == 'OP' and t.value == '(':
            self.take(); node = self.expression(0); self.expect(')')
            return {'kind': 'parenthesized_expression', 'expression': node}
        if t.kind == 'STRING':
            self.take(); return {'kind': 'literal', 'literal_type': 'string', 'value': t.value[1:-1].replace('""', '"')}
        if t.kind == 'DATE':
            self.take(); return {'kind': 'literal', 'literal_type': 'date', 'value': t.value[1:-1], 'vba_date_literal': True}
        if t.kind in {'NUMBER', 'HEX', 'OCT'}:
            self.take(); return self.number_literal(t)
        if t.kind == 'IDENT':
            self.take()
            if low in {'true', 'false'}:
                return {'kind': 'literal', 'literal_type': 'boolean', 'value': low == 'true'}
            if low in {'nothing', 'null', 'empty'}:
                return {'kind': 'literal', 'literal_type': low, 'value': None}
            if low == 'new':
                if self.current.kind != 'IDENT':
                    raise ParseError("Expected class/type name after New")
                type_name = self.take().value
                return {'kind': 'new_object_expression', 'type_name': type_name}
            if low == 'me':
                return {'kind': 'self_reference', 'name': t.value}
            return {'kind': 'identifier', 'name': t.value}
        raise ParseError(f"Expected expression at offset {t.pos}; got {t.value!r}")

    def number_literal(self, t: Token) -> dict[str, Any]:
        raw = t.value
        if t.kind == 'HEX':
            return {'kind': 'literal', 'literal_type': 'integer', 'value': int(raw[2:], 16), 'source_literal': raw}
        if t.kind == 'OCT':
            return {'kind': 'literal', 'literal_type': 'integer', 'value': int(raw[2:], 8), 'source_literal': raw}
        suffix = raw[-1] if raw[-1:] in {'!', '#', '@', '%', '&'} else ''
        body = raw[:-1] if suffix else raw
        if suffix == '@':
            return {'kind': 'literal', 'literal_type': 'currency', 'value': body, 'source_literal': raw, 'scale': 4}
        if suffix == '#':
            typ = 'float64'
        elif suffix == '!':
            typ = 'float32'
        elif suffix == '%':
            typ = 'int16'
        elif suffix == '&':
            typ = 'int32'
        else:
            typ = 'float64' if any(c in body for c in '.Ee') else 'integer'
        value: Any = float(body) if any(c in body for c in '.Ee') else int(body)
        return {'kind': 'literal', 'literal_type': typ, 'value': value, 'source_literal': raw}

    def postfix(self, node: dict[str, Any]) -> dict[str, Any]:
        while True:
            if self.accept('.'):
                if self.current.kind != 'IDENT':
                    raise ParseError(f"Expected member name at offset {self.current.pos}")
                member = self.take().value
                node = {'kind': 'member_access', 'object': node, 'member': member}
                continue
            if self.accept('('):
                args: list[dict[str, Any]] = []
                if not self.accept(')'):
                    while True:
                        args.append(self.expression(0))
                        if self.accept(')'):
                            break
                        self.expect(',')
                node = {'kind': 'invocation_expression', 'callee': node, 'arguments': args, 'syntax': 'parenthesized'}
                continue
            break
        return node


def parse_expression(source: str) -> dict[str, Any]:
    return Parser(source).parse()


def _split_top_level_commas(source: str) -> list[str]:
    pieces=[]; start=0; depth=0; in_string=False; i=0
    while i < len(source):
        ch=source[i]
        if ch == '"':
            if in_string and i+1 < len(source) and source[i+1] == '"': i += 2; continue
            in_string = not in_string
        elif not in_string:
            if ch == '(': depth += 1
            elif ch == ')': depth -= 1
            elif ch == ',' and depth == 0:
                pieces.append(source[start:i].strip()); start=i+1
        i += 1
    pieces.append(source[start:].strip())
    return [p for p in pieces if p]


def parse_vba_call_statement(source: str) -> dict[str, Any]:
    """Parse VBA's statement-style invocation, e.g. `LogAudit message`.

    This is deliberately separate from expression calls because VBA permits
    omitted parentheses and whitespace-separated arguments in statements.
    """
    text = source.strip()
    if text.lower().startswith('call '):
        text = text[5:].strip()
    # Find the callee prefix, allowing a dotted method name.
    m = re.match(r'([A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z_][A-Za-z0-9_]*)*)\b(.*)$', text, re.S)
    if not m:
        raise ParseError('Could not identify VBA statement-call target')
    callee_text, rest = m.group(1), m.group(2).strip()
    callee = parse_expression(callee_text)
    if rest.startswith('(') and rest.endswith(')'):
        # Parenthesized argument list after a statement-style Call.
        expr = parse_expression(callee_text + rest)
        return expr
    args: list[dict[str, Any]] = []
    if rest:
        for piece in _split_top_level_commas(rest):
            args.append(parse_expression(piece))
    return {'kind': 'invocation_expression', 'callee': callee, 'arguments': args,
            'syntax': 'vba_statement', 'source': source.strip()}
