"""Structural checks for the normalized expression/control-flow subset."""
from __future__ import annotations
from typing import Any

FORBIDDEN = {'unsupported','unresolved_expression','unresolved_lvalue','unresolved_call'}


def validate_normalized_ir(program: dict[str, Any]) -> list[dict[str, Any]]:
    issues=[]
    def issue(path,code,message):
        issues.append({'severity':'error','code':code,'path':path,'message':message})
    def walk(x: Any, path: str):
        if isinstance(x,dict):
            kind=x.get('kind')
            if kind in FORBIDDEN:
                issue(path,'NORMALIZED_IR_UNRESOLVED_NODE',f"Node kind {kind!r} is not permitted in normalized IR.")
            if kind=='binary_expression':
                if x.get('operator') is None or not isinstance(x.get('left'),dict) or not isinstance(x.get('right'),dict):
                    issue(path,'INVALID_BINARY_EXPRESSION','Binary expression requires operator, left and right operands.')
            elif kind=='unary_expression':
                if x.get('operator') is None or not isinstance(x.get('operand'),dict):
                    issue(path,'INVALID_UNARY_EXPRESSION','Unary expression requires operator and operand.')
            elif kind=='member_access':
                if not isinstance(x.get('object'),dict) or not x.get('member'):
                    issue(path,'INVALID_MEMBER_ACCESS','Member access requires an object expression and member name.')
            elif kind=='invocation_expression':
                if not isinstance(x.get('callee'),dict) or not isinstance(x.get('arguments'),list):
                    issue(path,'INVALID_INVOCATION','Invocation requires callee and argument list.')
                if not x.get('invocation_kind'):
                    issue(path,'UNCLASSIFIED_INVOCATION','Invocation kind was not resolved.')
            elif kind=='index_access':
                if not isinstance(x.get('object'),dict) or not isinstance(x.get('indices'),list):
                    issue(path,'INVALID_INDEX_ACCESS','Array index access requires an object and index list.')
            elif kind=='collection_item_access':
                if not isinstance(x.get('collection'),dict) or not isinstance(x.get('key_or_index'),list):
                    issue(path,'INVALID_COLLECTION_ITEM_ACCESS','Collection item access requires collection and index/key.')
            elif kind=='identifier' and not x.get('name'):
                issue(path,'INVALID_IDENTIFIER','Identifier name is required.')
            elif kind=='for' and not isinstance(x.get('step'),dict):
                issue(path,'MISSING_FOR_STEP','For loop must have an explicit normalized step expression.')
            if kind=='named' and (x.get('semantics') or {}).get('resolution')=='unresolved_named_type':
                issue(path,'UNRESOLVED_NAMED_TYPE',f"Named type {x.get('name')!r} was not resolved.")
            for key,value in x.items(): walk(value,f'{path}.{key}')
        elif isinstance(x,list):
            for i,value in enumerate(x): walk(value,f'{path}[{i}]')
    walk(program,'$')
    return issues
