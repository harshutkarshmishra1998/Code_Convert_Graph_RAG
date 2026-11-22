"""Semantic normalization pass for Universal IR.

Consumes the Universal IR plus the exact frontend graph, granular segments,
pseudocode and manifest. Full MODULE/PROCEDURE snapshots are deliberately
excluded from the evidence index. This pass is deterministic and does not call
an LLM or invent expression trees: parse failures remain explicit blockers.
"""
from __future__ import annotations
import copy, hashlib, json, re
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

from .expression_parser import parse_expression, parse_vba_call_statement, ParseError
from .normalized_validator import validate_normalized_ir

BUILTINS = {
    'len','trim','trim$','ltrim','ltrim$','rtrim','rtrim$','replace','split','join','ubound','lbound',
    'cstr','clng','cint','cdbl','ccur','cdate','cbool','cbyte','csng','cdec','cvar','isarray',
    'isnumeric','isdate','ucase','ucase$','lcase','lcase$','left','left$','right','right$','mid','mid$',
    'format','format$','now','date','time','year','month','day','hour','minute','second','strcomp',
    'iif','abs','round','int','fix','val','asc','chr','chr$','space','space$','string','string$',
    'doevents','createobject','array','typename','isnull','isempty','isobject','err.clear'
}
INTRINSIC_CONSTANTS = {
    'vbcrlf','vbcr','vblf','vbtab','vbnullstring','vbtextcompare','vbinarycompare','vbobjecterror',
    'vbnewline','vbnullchar','vbnull','true','false','nothing','null','empty'
}
INTRINSIC_OBJECTS = {'err'}


def _norm(s: str) -> str:
    return s.casefold()


def _type_name(t: Any) -> str | None:
    if not isinstance(t, dict): return None
    if t.get('kind') == 'array': return 'array'
    if t.get('kind') == 'collection': return 'collection'
    if t.get('kind') == 'variant': return 'variant'
    return t.get('name') or t.get('source_type')


def _walk_statements(items):
    for s in items or []:
        if not isinstance(s, dict): continue
        yield s
        for b in s.get('branches', []) or []:
            yield from _walk_statements(b.get('body', []))
        yield from _walk_statements(s.get('body', []))


def _make_symbol_table(module: dict[str, Any]) -> dict[str, Any]:
    table: dict[str, Any] = {'module': {}, 'procedures': {}, 'records': {}, 'enums': {}, 'enum_members': {}, 'procedure_symbols': {}}
    for d in module.get('declarations', []):
        name = d.get('name', '')
        if not name: continue
        table['module'][_norm(name)] = {'name': name, 'kind': d.get('kind'), 'type': d.get('type'), 'id': d.get('id')}
        if d.get('kind') == 'record':
            table['records'][_norm(name)] = { _norm(f.get('name','')): f for f in d.get('members', []) if f.get('name') }
        if d.get('kind') == 'enum':
            table['enums'][_norm(name)] = { _norm(x.get('name','')): x for x in d.get('members', []) if x.get('name') }
            for member in d.get('members', []):
                if member.get('name'):
                    table['enum_members'][_norm(member['name'])] = {'name':member['name'],'kind':'enum_member','type':{'kind':'primitive','name':'int32'},'id':member.get('id'),'enum':name}
    for p in module.get('procedures', []):
        pn = p.get('name', '')
        if not pn: continue
        table['procedures'][_norm(pn)] = p
        symbols: dict[str, Any] = {}
        for param in p.get('parameters', []):
            symbols[_norm(param.get('name',''))] = {'name':param.get('name'), 'kind':'parameter', 'type':param.get('type'), 'passing':param.get('passing'), 'id':param.get('id')}
        for var in p.get('local_variables', []):
            if var.get('name'):
                # A parameter wins over its duplicated frontend local symbol.
                symbols.setdefault(_norm(var['name']), {'name':var['name'], 'kind':'local', 'type':var.get('type'), 'id':var.get('id')})
        table['procedure_symbols'][_norm(pn)] = symbols
    return table


def _symbol_for(name: str, proc_name: str, table: dict[str, Any]):
    key = _norm(name)
    local = table['procedure_symbols'].get(_norm(proc_name), {})
    if key in local: return local[key]
    if key in table['module']: return table['module'][key]
    if key in table['enum_members']: return table['enum_members'][key]
    if key in table['procedures']: return {'name':name, 'kind':'procedure', 'procedure':table['procedures'][key]}
    return None


def _record_for_type(type_obj: Any, table):
    if not isinstance(type_obj, dict): return None
    nm = type_obj.get('name') or type_obj.get('source_type')
    return table['records'].get(_norm(str(nm))) if nm else None


def _infer_expr_type(node: dict[str, Any], proc_name: str, table: dict[str, Any]) -> dict[str, Any] | None:
    k = node.get('kind')
    if k == 'literal':
        typ = node.get('literal_type')
        return {'kind':'primitive','name': {'string':'string','boolean':'boolean','integer':'int32','int16':'int16','int32':'int32','float32':'float32','float64':'float64','currency':'decimal','date':'datetime'}.get(typ, typ or 'unknown')}
    if k == 'identifier':
        sym = _symbol_for(node.get('name',''), proc_name, table)
        if sym:
            return sym.get('runtime_value_type') or sym.get('type')
        if _norm(node.get('name','')) in INTRINSIC_CONSTANTS:
            return {'kind':'intrinsic','name':node['name']}
        return None
    if k == 'self_reference': return {'kind':'object','name':'module_self'}
    if k == 'member_access':
        owner_t = _infer_expr_type(node.get('object',{}), proc_name, table)
        fields = _record_for_type(owner_t, table)
        field = fields.get(_norm(node.get('member',''))) if fields else None
        if field: return field.get('type')
        return None
    if k == 'index_access':
        base_t = _infer_expr_type(node.get('object',{}), proc_name, table)
        if isinstance(base_t, dict) and base_t.get('kind') == 'array': return base_t.get('element_type')
        return None
    if k == 'collection_item_access':
        return {'kind':'variant','name':'Variant','semantics':{'collection_default_item':True}}
    if k == 'invocation_expression':
        callee = node.get('callee', {})
        if callee.get('kind') == 'identifier':
            name = _norm(callee.get('name',''))
            proc = table['procedures'].get(name)
            if proc: return proc.get('return_type')
            if name in {'split','array','filter'}: return {'kind':'array','element_type':{'kind':'primitive','name':'string'},'semantics':{'runtime_array_from_builtin':name}}
            if name in {'len','ubound','lbound','year','month','day','hour','minute','second','clng','cint'}:
                return {'kind':'primitive','name':'int32'}
            if name in {'cstr','trim','trim$','ltrim','ltrim$','rtrim','rtrim$','ucase','ucase$','lcase','lcase$','left','left$','right','right$','mid','mid$','replace','format','format$','str'}:
                return {'kind':'primitive','name':'string'}
            if name in {'isarray','isnumeric','isdate','isnull','isempty','isobject'}:
                return {'kind':'primitive','name':'boolean'}
            if name in {'cdbl','csng','val','abs'}: return {'kind':'primitive','name':'float64'}
            if name == 'ccur': return {'kind':'primitive','name':'decimal','semantics':{'fixed_scale':4}}
            if name == 'cdate' or name == 'date': return {'kind':'primitive','name':'datetime'}
        return None
    if k == 'new_object_expression': return {'kind':'object','name':node.get('type_name')}
    if k == 'unary_expression':
        op = node.get('operator')
        if op == 'not':
            operand_type=_infer_expr_type(node.get('operand',{}),proc_name,table)
            return operand_type if _type_name(operand_type) in {'boolean','int16','int32','int64','uint8'} else None
        return _infer_expr_type(node.get('operand',{}), proc_name, table)
    if k == 'binary_expression':
        op = node.get('operator')
        if op in {'=','<>','<','>','<=','>=','is','is_not','like'}: return {'kind':'primitive','name':'boolean'}
        if op == '&': return {'kind':'primitive','name':'string'}
        lt = _infer_expr_type(node.get('left',{}), proc_name, table)
        rt = _infer_expr_type(node.get('right',{}), proc_name, table)
        if op in {'and','or','xor','eqv','imp'}:
            return lt if _type_name(lt) in {'int32','int16','int64','boolean'} else {'kind':'primitive','name':'boolean'}
        return lt or rt
    if k == 'parenthesized_expression': return _infer_expr_type(node.get('expression',{}), proc_name, table)
    return None


def _resolve_expr(node: dict[str, Any], proc_name: str, table: dict[str, Any], issues: list[dict[str, Any]], path: str) -> dict[str, Any]:
    node = copy.deepcopy(node)
    k = node.get('kind')
    if k == 'identifier':
        name = node.get('name','')
        sym = _symbol_for(name, proc_name, table)
        if sym:
            node['binding'] = {'symbol_kind':sym.get('kind'), 'symbol_name':sym.get('name'), 'symbol_id':sym.get('id')}
            if sym.get('type'): node['inferred_type'] = sym['type']
        elif _norm(name) in INTRINSIC_CONSTANTS:
            node['binding'] = {'symbol_kind':'intrinsic_constant', 'symbol_name':name}
        elif _norm(name) in INTRINSIC_OBJECTS:
            node['binding'] = {'symbol_kind':'intrinsic_object', 'symbol_name':name}
        elif _norm(name) in BUILTINS:
            node['binding'] = {'symbol_kind':'intrinsic_function', 'symbol_name':name}
        else:
            issues.append({'severity':'warning','code':'UNRESOLVED_IDENTIFIER','procedure':proc_name,'path':path,'symbol':name,
                           'message':'Identifier was not found in local, parameter, module, procedure, or known intrinsic symbols.'})
            node['binding'] = {'symbol_kind':'unresolved','symbol_name':name}
    elif k == 'member_access':
        node['object'] = _resolve_expr(node['object'], proc_name, table, issues, path+'.object')
        owner_t = _infer_expr_type(node['object'], proc_name, table)
        fields = _record_for_type(owner_t, table)
        member = fields.get(_norm(node.get('member',''))) if fields else None
        if member:
            node['binding'] = {'symbol_kind':'record_field','symbol_name':member.get('name'),'symbol_id':member.get('id')}
            if member.get('type'): node['inferred_type'] = member['type']
        elif _type_name(owner_t) == 'collection' and _norm(node.get('member','')) == 'count':
            node['binding']={'symbol_kind':'collection_property','symbol_name':node['member']}; node['inferred_type']={'kind':'primitive','name':'int32'}
        elif _type_name(owner_t) == 'collection' and _norm(node.get('member','')) in {'add','item','remove'}:
            node['binding']={'symbol_kind':'collection_method','symbol_name':node['member']}
        elif node.get('object',{}).get('kind') == 'identifier' and _norm(node['object'].get('name','')) == 'err' and _norm(node.get('member','')) in {'clear','number','description','source','helpcontext','helpfile'}:
            node['binding']={'symbol_kind':'vba_err_member','symbol_name':node['member']}
        elif _norm(node.get('member','')) == 'number' and node['object'].get('kind') == 'identifier' and _norm(node['object'].get('name','')) == 'err':
            node['binding']={'symbol_kind':'vba_err_property','symbol_name':'Number'}; node['inferred_type']={'kind':'primitive','name':'int32'}
        else:
            node['binding']={'symbol_kind':'late_bound_or_unresolved_member','owner_type':_type_name(owner_t),'member':node.get('member')}
            if owner_t is None:
                issues.append({'severity':'warning','code':'UNRESOLVED_MEMBER','procedure':proc_name,'path':path,'member':node.get('member'),
                               'message':'Member owner type could not be resolved; retain as late-bound member access.'})
    elif k == 'invocation_expression':
        node['callee'] = _resolve_expr(node['callee'], proc_name, table, issues, path+'.callee')
        node['arguments'] = [_resolve_expr(a, proc_name, table, issues, f'{path}.arguments[{i}]') for i,a in enumerate(node.get('arguments',[]))]
        callee = node.get('callee', {})
        if callee.get('kind') == 'member_access' and _type_name(callee.get('inferred_type')) == 'array':
            node = {'kind':'index_access','object':callee,'indices':node.get('arguments',[]),'syntax':'vba_array_index',
                    'source':node.get('source'),'normalization':{'method':'symbol_type_resolution','reason':'callee resolves to an array field'}}
        elif callee.get('kind') == 'identifier' and _symbol_for(callee.get('name',''),proc_name,table) and _type_name(_symbol_for(callee.get('name',''),proc_name,table).get('runtime_value_type') or _symbol_for(callee.get('name',''),proc_name,table).get('type')) == 'array':
            node = {'kind':'index_access','object':callee,'indices':node.get('arguments',[]),'syntax':'vba_array_index',
                    'source':node.get('source'),'normalization':{'method':'symbol_or_dataflow_type_resolution','reason':'variable is declared or inferred as an array'}}
        elif callee.get('kind') == 'identifier' and _symbol_for(callee.get('name',''),proc_name,table) and _type_name(_symbol_for(callee.get('name',''),proc_name,table).get('type')) == 'collection':
            node = {'kind':'collection_item_access','collection':callee,'key_or_index':node.get('arguments',[]),
                    'syntax':'vba_collection_default_item','source':node.get('source'),
                    'semantics':{'vba_default_member':'Item','collection_is_one_based':True}}
        elif callee.get('kind') == 'identifier':
            name = callee.get('name',''); key = _norm(name)
            if key in table['procedures']:
                node['invocation_kind']='procedure_call'
                target = table['procedures'][key]
                node['target_procedure']={'name':target.get('name'),'id':target.get('id')}
                if target.get('return_type'): node['inferred_type']=target['return_type']
                bindings=[]
                for ai,arg in enumerate(node.get('arguments',[])):
                    params=target.get('parameters',[])
                    if ai < len(params):
                        param=params[ai]
                        passing=param.get('passing','reference')
                        binding={'argument_index':ai,'parameter_name':param.get('name'),'passing':passing,
                                 'parameter_type':param.get('type')}
                        if passing == 'reference':
                            binding['requires_mutable_lvalue']=True
                            binding['argument_is_lvalue']=arg.get('kind') in {'identifier','member_access','index_access','collection_item_access'}
                        bindings.append(binding)
                node['argument_bindings']=bindings
            elif key in BUILTINS:
                node['invocation_kind']='vba_builtin'
                node['builtin_name']=name
                t = _infer_expr_type(node, proc_name, table)
                if t: node['inferred_type']=t
            else:
                node['invocation_kind']='unresolved_function_or_array_access'
                issues.append({'severity':'warning','code':'UNRESOLVED_INVOCATION','procedure':proc_name,'path':path,'callee':name,
                               'message':'Invocation could be a user function, external function, or array; no declaration resolved it.'})
        elif callee.get('kind') == 'member_access':
            node['invocation_kind']='method_call'
            node['method_name']=callee.get('member')
            node['owner_expression']=callee.get('object')
            if _norm(callee.get('member','')) in {'add','exists','remove','item'}:
                node['method_family']='collection_or_dictionary_method'
        else:
            node['invocation_kind']='computed_callable'
    elif k in {'binary_expression'}:
        node['left'] = _resolve_expr(node['left'], proc_name, table, issues, path+'.left')
        node['right'] = _resolve_expr(node['right'], proc_name, table, issues, path+'.right')
    elif k == 'unary_expression':
        node['operand'] = _resolve_expr(node['operand'], proc_name, table, issues, path+'.operand')
        if node.get('operator') == 'not':
            operand_type=_infer_expr_type(node['operand'],proc_name,table)
            node['semantics']={'vba_not_operator':'operand_type_dependent','operand_type':_type_name(operand_type)}
            if _type_name(operand_type) == 'array': node['semantics']['possible_array_initialization_idiom']=True
    elif k == 'parenthesized_expression':
        node['expression'] = _resolve_expr(node['expression'], proc_name, table, issues, path+'.expression')
    elif k == 'index_access':
        node['object'] = _resolve_expr(node['object'], proc_name, table, issues, path+'.object')
        node['indices'] = [_resolve_expr(a, proc_name, table, issues, f'{path}.indices[{i}]') for i,a in enumerate(node.get('indices',[]))]
    inferred = _infer_expr_type(node, proc_name, table)
    if inferred and 'inferred_type' not in node:
        node['inferred_type'] = inferred
    return node


def _parse_or_unresolved(source: str | None, *, proc: str, path: str, issues: list[dict[str, Any]]) -> dict[str, Any]:
    if source is None or not str(source).strip():
        return {'kind':'unresolved_expression','source_expression':source,'parse_status':'missing_source',
                'reason':'No expression text was supplied by the frontend.'}
    try:
        return parse_expression(str(source))
    except ParseError as exc:
        issues.append({'severity':'error','code':'EXPRESSION_PARSE_FAILED','procedure':proc,'path':path,
                       'source_expression':str(source),'message':str(exc)})
        return {'kind':'unresolved_expression','source_expression':str(source),'parse_status':'failed','reason':str(exc)}


def _normalize_expr_slot(slot: dict[str, Any], proc: str, path: str, table, issues):
    if slot.get('kind') == 'unsupported':
        source = (slot.get('semantics') or {}).get('source_expression')
        parsed = _parse_or_unresolved(source, proc=proc, path=path, issues=issues)
        if parsed.get('kind') == 'unresolved_expression':
            return {**parsed, 'original_ir_node':copy.deepcopy(slot), 'provenance':slot.get('provenance')}
        parsed['source'] = source
        parsed['normalization'] = {'method':'deterministic_vba_expression_parser','original_ir_kind':'unsupported'}
        return _resolve_expr(parsed, proc, table, issues, path)
    # Existing structured expression trees are still recursively resolved.
    return _resolve_expr(slot, proc, table, issues, path)


def _parse_target(target: dict[str, Any], proc: str, path: str, table, issues):
    if not isinstance(target, dict): return target
    if target.get('kind') == 'identifier':
        raw = target.get('name','')
        proc_def = table['procedures'].get(_norm(proc))
        if proc_def and _norm(raw) == _norm(proc) and proc_def.get('kind') == 'function':
            return {'kind':'implicit_function_return_target','procedure_name':proc,
                    'return_type':proc_def.get('return_type'),
                    'semantics':{'vba_function_name_assignment_is_return_value':True}}
        parsed = _parse_or_unresolved(raw, proc=proc, path=path, issues=issues)
        if parsed.get('kind') == 'unresolved_expression': return {'kind':'unresolved_lvalue','source':raw,'reason':parsed.get('reason')}
        # For assignment targets, a call-like postfix over a declared array is an index access.
        parsed = _resolve_expr(parsed, proc, table, issues, path)
        return {'kind':'lvalue','expression':parsed}
    return target


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


def _split_to_bounds(source: str) -> tuple[str | None, str]:
    # A VBA ReDim dimension may be `lower To upper`; ignore `To` in strings.
    m = re.match(r'^\s*(.*?)\s+To\s+(.*?)\s*$', source, re.I|re.S)
    return (m.group(1).strip(),m.group(2).strip()) if m else (None,source.strip())


def _collect_runtime_array_assignments(items, proc_name, table):
    """Infer Variant variables that receive arrays from known VBA array functions."""
    symbols = table['procedure_symbols'].get(_norm(proc_name), {})
    for stmt in _walk_statements(items):
        if stmt.get('kind') != 'assignment':
            continue
        target = stmt.get('target', {})
        value = stmt.get('value', {})
        target_name = target.get('name') if isinstance(target, dict) and target.get('kind') == 'identifier' else None
        if not target_name or not isinstance(value, dict) or value.get('kind') != 'unsupported':
            continue
        source = (value.get('semantics') or {}).get('source_expression') or ''
        try:
            parsed = parse_expression(source)
        except ParseError:
            continue
        if parsed.get('kind') == 'invocation_expression' and parsed.get('callee',{}).get('kind') == 'identifier' and _norm(parsed['callee'].get('name','')) in {'split','array','filter'}:
            sym = symbols.get(_norm(target_name))
            if sym:
                sym['runtime_value_type'] = {'kind':'array','element_type':{'kind':'primitive','name':'string'},
                                             'semantics':{'inferred_from_assignment':source,'inference_rule':'VBA Split/Array/Filter returns an array'}}


def _normalize_statement_inner(stmt: dict[str, Any], proc: str, path: str, table, issues):
    k = stmt.get('kind')
    if k == 'assignment':
        stmt['target'] = _parse_target(stmt.get('target',{}), proc, path+'.target', table, issues)
        stmt['value'] = _normalize_expr_slot(stmt.get('value',{}), proc, path+'.value', table, issues)
    elif k == 'if':
        stmt['condition'] = _normalize_expr_slot(stmt.get('condition',{}), proc, path+'.condition', table, issues)
        for bi, branch in enumerate(stmt.get('branches',[])):
            if branch.get('condition'):
                branch['condition'] = _normalize_expr_slot(branch['condition'], proc, f'{path}.branches[{bi}].condition', table, issues)
            branch['body'] = [_normalize_statement(s,proc,f'{path}.branches[{bi}].body[{i}]',table,issues) for i,s in enumerate(branch.get('body',[]))]
    elif k == 'for':
        for field in ('start','end','step'):
            if isinstance(stmt.get(field),dict) and stmt[field].get('kind') == 'unsupported':
                if field == 'step' and not (stmt[field].get('semantics') or {}).get('source_expression'):
                    stmt[field] = {'kind':'literal','literal_type':'integer','value':1,'source_literal':'1',
                                   'normalization':{'method':'vba_language_default','rule':'For loop default Step is 1','inferred':True}}
                else:
                    stmt[field] = _normalize_expr_slot(stmt[field],proc,path+'.'+field,table,issues)
        stmt['body'] = [_normalize_statement(s,proc,f'{path}.body[{i}]',table,issues) for i,s in enumerate(stmt.get('body',[]))]
    elif k == 'call':
        raw = (stmt.get('semantics') or {}).get('raw_source') or (stmt.get('provenance') or {}).get('source_text')
        try:
            expr = parse_vba_call_statement(raw or '')
            expr = _resolve_expr(expr,proc,table,issues,path+'.expression')
            stmt['expression'] = expr
            stmt['semantics'] = {**(stmt.get('semantics') or {}), 'call_resolution':expr.get('invocation_kind','unresolved'),
                                 'normalized_from_statement_syntax':True}
        except ParseError as exc:
            issues.append({'severity':'error','code':'CALL_PARSE_FAILED','procedure':proc,'path':path,'source':raw,'message':str(exc)})
            stmt['expression'] = {'kind':'unresolved_call','source':raw,'reason':str(exc)}
    elif k == 'redim':
        raw = (stmt.get('semantics') or {}).get('bounds_source','')
        dims=[]
        for dim in _split_top_level_commas(raw.strip().strip('()')):
            lower, upper = _split_to_bounds(dim)
            if lower is None:
                dims.append({'kind':'upper_bound_only','upper':_resolve_expr(_parse_or_unresolved(upper,proc=proc,path=path+'.dimension.upper',issues=issues),proc,table,issues,path+'.dimension.upper')})
            else:
                dims.append({'kind':'explicit_bounds','lower':_resolve_expr(_parse_or_unresolved(lower,proc=proc,path=path+'.dimension.lower',issues=issues),proc,table,issues,path+'.dimension.lower'),
                             'upper':_resolve_expr(_parse_or_unresolved(upper,proc=proc,path=path+'.dimension.upper',issues=issues),proc,table,issues,path+'.dimension.upper')})
        stmt['target'] = _parse_target(stmt.get('target',{}),proc,path+'.target',table,issues)
        stmt['dimensions'] = dims
        stmt['semantics'] = {**(stmt.get('semantics') or {}), 'bounds_parse_status':'parsed' if dims and all(d.get('upper',{}).get('kind')!='unresolved_expression' for d in dims) else 'partial'}
    elif k == 'erase':
        if isinstance(stmt.get('target'),dict): stmt['target']=_parse_target(stmt['target'],proc,path+'.target',table,issues)
    elif k == 'goto':
        pass
    elif k in {'gosub','gosub_return','label','exit','on_error'}:
        pass
    else:
        # Some nested statements can be opaque expressions directly.
        for key,val in list(stmt.items()):
            if isinstance(val,dict) and val.get('kind') == 'unsupported':
                stmt[key] = _normalize_expr_slot(val,proc,path+'.'+key,table,issues)
            elif isinstance(val,list):
                stmt[key] = [_normalize_statement(x,proc,f'{path}.{key}[{i}]',table,issues) if isinstance(x,dict) and 'kind' in x else x for i,x in enumerate(val)]
    return stmt


def _normalize_statement(s,proc,path,table,issues):
    return _normalize_statement_inner(copy.deepcopy(s),proc,path,table,issues)


def _segment_evidence(segments: dict[str, Any]) -> tuple[dict[str,dict], dict[str,int]]:
    items = segments if isinstance(segments,list) else segments.get('segments',[]) if isinstance(segments,dict) else []
    by_id={}; kinds=Counter()
    for seg in items:
        if not isinstance(seg,dict): continue
        kind=seg.get('kind','UNKNOWN'); kinds[kind]+=1
        # Explicitly exclude complete procedure/module snapshots from the evidence index.
        if kind in {'PROCEDURE','MODULE'}: continue
        if seg.get('id'): by_id[seg['id']]=seg
    return by_id,dict(kinds)


def _pseudocode_sections(text: str) -> set[str]:
    names=set()
    headers=list(re.finditer(r'^\s*#{2,4}\s+Procedure:\s*([^(`]+).*$',text,re.I|re.M))
    for match in headers:
        next_heading=re.search(r'(?m)^[ \t]*#{1,3}\s+',text[match.end():])
        next_start=match.end()+next_heading.start() if next_heading else len(text)
        body=text[match.end():next_start].strip()
        if not body or re.search(r'\[GENERATION FAILED:',body,re.I):
            continue
        names.add(_norm(match.group(1).strip()))
    return names


def _check_control_flow(proc: dict[str, Any], issues: list[dict[str, Any]]):
    name=proc.get('name','')
    labels={_norm(s.get('label','')) for s in _walk_statements(proc.get('body',[])) if s.get('kind')=='label'}
    gosub_returns=sum(1 for s in _walk_statements(proc.get('body',[])) if s.get('kind')=='gosub_return')
    gosubs=[s for s in _walk_statements(proc.get('body',[])) if s.get('kind')=='gosub']
    for s in _walk_statements(proc.get('body',[])):
        if s.get('kind') in {'goto','gosub'}:
            label=_norm(s.get('label',''))
            if label not in labels:
                issues.append({'severity':'error','code':'UNRESOLVED_LABEL','procedure':name,'label':s.get('label'),
                               'line':(s.get('provenance') or {}).get('source_location',{}).get('line'),
                               'message':'Jump target label is not represented in this procedure.'})
        if s.get('kind')=='on_error' and s.get('handler') and _norm(s['handler']) not in labels:
            issues.append({'severity':'error','code':'UNRESOLVED_ERROR_HANDLER','procedure':name,'label':s['handler'],
                           'message':'On Error handler label is not represented in this procedure.'})
    if gosubs and not gosub_returns:
        issues.append({'severity':'error','code':'GOSUB_WITHOUT_RETURN','procedure':name,'message':'GoSub exists but no GoSub Return node was found.'})


def _resolve_type_tree(type_obj: Any, table: dict[str, Any]) -> Any:
    if not isinstance(type_obj, dict): return type_obj
    t=copy.deepcopy(type_obj)
    if t.get('kind') == 'named' and _norm(t.get('name','')) in table['records']:
        t['kind']='record'; t['definition_name']=t['name']; t['semantics']={**(t.get('semantics') or {}),'resolution':'resolved_record','resolved':True}
    for key in ('element_type','underlying_type','value_type','key_type'):
        if isinstance(t.get(key),dict): t[key]=_resolve_type_tree(t[key],table)
    if isinstance(t.get('dimensions'),list):
        t['dimensions']=[copy.deepcopy(x) for x in t['dimensions']]
    return t


def _resolve_embedded_declaration_types(node: Any, table: dict[str, Any]) -> None:
    if isinstance(node,dict):
        if isinstance(node.get('type'),dict): node['type']=_resolve_type_tree(node['type'],table)
        for declaration in node.get('declarations',[]) if isinstance(node.get('declarations'),list) else []:
            if isinstance(declaration,dict) and isinstance(declaration.get('type'),dict):
                declaration['type']=_resolve_type_tree(declaration['type'],table)
        for value in node.values(): _resolve_embedded_declaration_types(value,table)
    elif isinstance(node,list):
        for value in node: _resolve_embedded_declaration_types(value,table)


def _count_nodes(root: Any, kinds: set[str]) -> int:
    count=0
    def walk(x):
        nonlocal count
        if isinstance(x,dict):
            if x.get('kind') in kinds: count+=1
            for v in x.values(): walk(v)
        elif isinstance(x,list):
            for v in x: walk(v)
    walk(root); return count


def _compatibility_requirements(program: dict[str, Any]) -> list[dict[str, Any]]:
    req=[]
    modules=program.get('modules',[])
    def add(code,title,guidance,evidence,severity='review'):
        req.append({'code':code,'title':title,'severity':severity,'target_guidance':guidance,'evidence':evidence})
    options=[]
    for m in modules:
        options.extend([x.get('source','') for x in (m.get('annotations') or {}).get('vba_options',[])])
    option_compare=[x for x in options if 'Option Compare Text' in x]
    option_base=[x for x in options if 'Option Base 1' in x]
    if option_compare:
        add('VBA_OPTION_COMPARE_TEXT','Case-insensitive string comparison','Preserve locale-aware VBA text-comparison semantics for affected string comparisons; do not assume Java String.equalsIgnoreCase is fully equivalent.',option_compare)
    if option_base:
        add('VBA_OPTION_BASE_1','One-based default array bounds','Use an array abstraction that preserves explicit lower bounds and VBA default lower-bound rules.',option_base)

    on_errors=[{'procedure':p.get('name'),'mode':s.get('mode'),'handler':s.get('handler')}
               for m in modules for p in m.get('procedures',[]) for s in _walk_statements(p.get('body',[])) if s.get('kind')=='on_error']
    if on_errors:
        add('VBA_ERROR_HANDLING','VBA On Error runtime semantics','Implement a compatibility runtime for Resume Next, GoTo handler, GoTo 0, Err state and handler activation; this pass preserves directives but does not emulate execution.',on_errors,'blocking')
    gosubs=[{'procedure':p.get('name'),'label':s.get('label')}
            for m in modules for p in m.get('procedures',[]) for s in _walk_statements(p.get('body',[])) if s.get('kind')=='gosub']
    if gosubs:
        add('VBA_GOSUB_RETURN','Legacy GoSub/Return control flow','Lower GoSub to an explicit continuation/return stack or a proven equivalent; never map Return to Java method return.',gosubs,'blocking')

    late=[]; create_objects=[]; byref=[]; iif_calls=[]; doevents=[]
    for m in modules:
        for p in m.get('procedures',[]):
            def scan(x):
                if isinstance(x,dict):
                    if (x.get('binding') or {}).get('symbol_kind')=='late_bound_or_unresolved_member':
                        late.append({'procedure':p.get('name'),'member':x.get('member'),'owner_type':(x.get('binding') or {}).get('owner_type')})
                    if x.get('kind')=='invocation_expression':
                        if x.get('builtin_name','').casefold()=='createobject':
                            create_objects.append({'procedure':p.get('name'),'expression':x.get('source')})
                        if x.get('builtin_name','').casefold()=='iif':
                            iif_calls.append({'procedure':p.get('name'),'source':x.get('source')})
                        if x.get('builtin_name','').casefold()=='doevents':
                            doevents.append({'procedure':p.get('name'),'source':x.get('source')})
                        for b in x.get('argument_bindings',[]):
                            if b.get('passing')=='reference':
                                byref.append({'caller':p.get('name'),'target':x.get('target_procedure',{}).get('name'),
                                              'parameter':b.get('parameter_name'),'argument_is_lvalue':b.get('argument_is_lvalue')})
                    for v in x.values(): scan(v)
                elif isinstance(x,list):
                    for v in x: scan(v)
            scan(p.get('body',[]))
    late=list({(x['procedure'],x['member'],x['owner_type']):x for x in late}.values())
    byref=list({(x['caller'],x['target'],x['parameter']):x for x in byref}.values())
    if late:
        add('LATE_BOUND_OBJECT_MEMBERS','Late-bound COM/Object members','Define explicit Java interfaces/adapters for the actual COM object behavior. Object-typed references alone do not identify the target API.',late,'blocking')
    if create_objects:
        add('COM_CREATEOBJECT','COM object activation','Replace ProgID-based activation with an explicit Java implementation or integration adapter; preserve fallback behavior.',create_objects,'blocking')
    if byref:
        add('VBA_BYREF','ByRef parameter mutation','Generate mutable holders or pass mutable record/array references so writes in the callee are visible to the caller.',byref)
    if iif_calls:
        add('VBA_IIF_EAGER_EVALUATION','IIf evaluates both result expressions','Do not translate IIf to Java’s lazy conditional operator without checking side effects and exceptions; VBA IIf evaluates both branches.',iif_calls)
    if doevents:
        add('VBA_DOEVENTS','DoEvents cooperative yield','Decide whether a Java event-loop/thread-yield equivalent is required; a no-op changes responsiveness semantics.',doevents)

    redims=[{'procedure':p.get('name'),'preserve':s.get('semantics',{}).get('preserve')}
            for m in modules for p in m.get('procedures',[]) for s in _walk_statements(p.get('body',[])) if s.get('kind')=='redim']
    if redims:
        add('VBA_DYNAMIC_ARRAYS','Dynamic arrays, ReDim, Preserve and Erase','Implement resizable arrays with preserved lower/upper bounds, element copying, and Erase semantics.',redims)

    all_types=[]
    for m in modules:
        all_types.extend([d.get('type') for d in m.get('declarations',[]) if d.get('type')])
        for d in m.get('declarations',[]): all_types.extend([f.get('type') for f in d.get('members',[]) if f.get('type')])
        for p in m.get('procedures',[]):
            all_types.extend([v.get('type') for v in p.get('parameters',[]) if v.get('type')])
            all_types.extend([v.get('type') for v in p.get('local_variables',[]) if v.get('type')])
    def types_walk(t):
        if not isinstance(t,dict): return
        yield t
        for key in ('element_type','underlying_type','value_type','key_type'):
            yield from types_walk(t.get(key))
    flat_types=[z for t in all_types for z in types_walk(t)]
    currencies=[x for x in flat_types if str(x.get('source_type','')).casefold()=='currency' or x.get('semantics',{}).get('fixed_scale')==4]
    if currencies:
        add('VBA_CURRENCY','VBA Currency fixed-point arithmetic','Use fixed-point decimal arithmetic with four decimal places and verify rounding/overflow rules.',{'type_occurrences':len(currencies)})
    dates=[x for x in flat_types if str(x.get('source_type','')).casefold()=='date' or x.get('semantics',{}).get('vba_date_representation')]
    if dates:
        add('VBA_DATE','VBA Date representation','Map VBA/OLE Automation dates carefully; verify serial-date conversions, timezone assumptions and date parsing/formatting.',{'type_occurrences':len(dates)})
    variants=[{'procedure':p.get('name'),'variable':v.get('name')}
              for m in modules for p in m.get('procedures',[]) for v in p.get('local_variables',[]) if (v.get('type') or {}).get('kind')=='variant']
    if variants:
        add('VBA_VARIANT','Variant dynamic values','Preserve runtime type changes, Empty/Null distinctions, arrays stored in Variant, and VBA conversion rules.',variants)
    collections=[{'module':m.get('name'),'variable':d.get('name')}
                 for m in modules for d in m.get('declarations',[]) if (d.get('type') or {}).get('kind')=='collection']
    if collections:
        add('VBA_COLLECTION','VBA Collection semantics','Preserve insertion order, one-based Item indexing, Count, Add and Remove behavior.',collections)

    array_not=[]; object_is=[]; numeric_logic=[]
    for m in modules:
        for p in m.get('procedures',[]):
            def scan_operators(x, source_context=None):
                if isinstance(x,dict):
                    current_source=x.get('source') or source_context
                    if x.get('kind')=='unary_expression' and x.get('operator')=='not' and (x.get('semantics') or {}).get('possible_array_initialization_idiom'):
                        array_not.append({'procedure':p.get('name'),'source':current_source,'operand_type':(x.get('semantics') or {}).get('operand_type')})
                    if x.get('kind')=='binary_expression' and x.get('operator') in {'is','is_not'}:
                        object_is.append({'procedure':p.get('name'),'operator':x.get('operator'),'source':current_source})
                    if x.get('kind')=='binary_expression' and x.get('operator') in {'and','or','xor','eqv','imp'} and _type_name(x.get('inferred_type')) in {'int16','int32','int64','uint8'}:
                        numeric_logic.append({'procedure':p.get('name'),'operator':x.get('operator'),'source':current_source})
                    for v in x.values(): scan_operators(v,current_source)
                elif isinstance(x,list):
                    for v in x: scan_operators(v,source_context)
            scan_operators(p.get('body',[]))
    array_not=list({(x['procedure'],x['source']):x for x in array_not}.values())
    object_is=list({(x['procedure'],x['operator'],x['source']):x for x in object_is}.values())
    numeric_logic=list({(x['procedure'],x['operator'],x['source']):x for x in numeric_logic}.values())
    if array_not:
        add('VBA_ARRAY_INITIALIZATION_IDIOM','Double-Not dynamic-array initialization check','Confirm the source runtime meaning and replace with an explicit initialized/uninitialized array state in Java. Do not translate Not Not as ordinary boolean negation.',array_not,'blocking')
    if object_is:
        add('VBA_OBJECT_IDENTITY','VBA Is object-reference comparison','Preserve object identity/null comparison semantics; distinguish Is from value equality.',object_is)
    if numeric_logic:
        add('VBA_BITWISE_LOGICAL_OPERATORS','Bitwise logical operators on numeric values','Preserve VBA bitwise behavior for numeric operands; do not map every And/Or expression to Java boolean operators.',numeric_logic)
    return req

def normalize_ir(ir: dict[str, Any], *, graph: dict[str, Any] | None=None, segments: dict[str, Any] | list | None=None,
                 pseudocode: str | None=None, manifest: dict[str, Any] | None=None) -> tuple[dict[str, Any],dict[str, Any]]:
    normalized=copy.deepcopy(ir)
    issues: list[dict[str, Any]]=[]
    module_reports=[]
    segment_map,segment_kinds=_segment_evidence(segments or {})
    pseudo_sections=_pseudocode_sections(pseudocode or '')
    used_segment_ids=set()
    for mi,module in enumerate(normalized.get('modules',[])):
        # First pass: resolve all declared types before resolving any call sites.
        table=_make_symbol_table(module)
        for declaration in module.get('declarations',[]):
            if isinstance(declaration.get('type'),dict): declaration['type']=_resolve_type_tree(declaration['type'],table)
            for field in declaration.get('members',[]):
                if isinstance(field.get('type'),dict): field['type']=_resolve_type_tree(field['type'],table)
            for member in declaration.get('members',[]):
                if isinstance(member.get('value'),dict):
                    member['value']=_normalize_expr_slot(member['value'], '<module-declaration>',
                                                         f'modules[{mi}].declarations.{declaration.get("name")}.{member.get("name")}.value',
                                                         table,issues)
        for proc in module.get('procedures',[]):
            for collection in (proc.get('parameters',[]),proc.get('local_variables',[])):
                for item in collection:
                    if isinstance(item.get('type'),dict): item['type']=_resolve_type_tree(item['type'],table)
                    if isinstance(item.get('default'),dict):
                        item['default']=_normalize_expr_slot(item['default'],proc.get('name','<unnamed>'),
                                                             f'{proc.get("name")}.parameter.{item.get("name")}.default',table,issues)
            if isinstance(proc.get('return_type'),dict): proc['return_type']=_resolve_type_tree(proc['return_type'],table)
        # Rebuild after type normalization so every cross-procedure call sees resolved parameter types.
        table=_make_symbol_table(module)
        _resolve_embedded_declaration_types(module.get('procedures',[]),table)
        proc_reports=[]
        for pi,proc in enumerate(module.get('procedures',[])):
            pname=proc.get('name','<unnamed>')
            _collect_runtime_array_assignments(proc.get('body',[]),pname,table)
            proc['body']=[_normalize_statement(s,pname,f'modules[{mi}].procedures[{pi}].body[{si}]',table,issues)
                          for si,s in enumerate(proc.get('body',[]))]
            _check_control_flow(proc,issues)
            # Collect provenance references and verify granular segment evidence.
            for s in _walk_statements(proc.get('body',[])):
                prov=s.get('provenance') or {}
                for sid in prov.get('segment_ids',[]):
                    used_segment_ids.add(sid)
                    if segments is not None and sid not in segment_map:
                        issues.append({'severity':'warning','code':'SEGMENT_EVIDENCE_MISSING','procedure':pname,'segment_id':sid,
                                       'message':'IR references a segment not present in the supplied granular-segment evidence set.'})
            expr_unresolved=0; expr_parsed=0; call_count=0
            for s in _walk_statements(proc.get('body',[])):
                def scan(x):
                    nonlocal expr_unresolved,expr_parsed,call_count
                    if isinstance(x,dict):
                        if x.get('kind') in {'unresolved_expression','unsupported'}: expr_unresolved+=1
                        if x.get('kind') in {'literal','identifier','binary_expression','unary_expression','member_access','index_access','invocation_expression','new_object_expression'} and x.get('source'):
                            expr_parsed+=1
                        if x.get('kind')=='invocation_expression': call_count+=1
                        for v in x.values(): scan(v)
                    elif isinstance(x,list):
                        for v in x: scan(v)
                scan(s)
            pseudo_status='detailed_section_present' if _norm(pname) in pseudo_sections else 'missing_detailed_section'
            proc['semantics']={**(proc.get('semantics') or {}),
                'semantic_normalization':{'status':'blocked' if expr_unresolved else 'parsed_with_review_items',
                                          'unresolved_expression_count':expr_unresolved,
                                          'parsed_expression_count':expr_parsed,
                                          'normalized_invocation_count':call_count},
                'pseudocode_coverage':pseudo_status}
            proc_reports.append({'procedure':pname,'unresolved_expression_count':expr_unresolved,
                                 'parsed_expression_count':expr_parsed,'invocation_count':call_count,
                                 'pseudocode_coverage':pseudo_status,
                                 'control_flow_integrity':(proc.get('semantics') or {}).get('control_flow_integrity','unknown')})
        module_reports.append({'module':module.get('name'),'procedures':proc_reports})
    # Graph/manifest checks are provenance checks, not permission to consume full snapshots.
    graph_summary={}
    if graph:
        graph_nodes=graph.get('nodes',[])
        graph_ids={x.get('id') for x in graph_nodes if x.get('id')}
        referenced_graph_ids=set()
        def collect_graph_refs(x):
            if isinstance(x,dict):
                for provenance_key in ('provenance','closing_provenance'):
                    prov=x.get(provenance_key)
                    if isinstance(prov,dict): referenced_graph_ids.update(prov.get('graph_node_ids',[]))
                for v in x.values(): collect_graph_refs(v)
            elif isinstance(x,list):
                for v in x: collect_graph_refs(v)
        collect_graph_refs(normalized.get('modules',[]))
        missing_graph_refs=sorted(referenced_graph_ids-graph_ids)
        graph_summary={'node_count':len(graph_nodes),'edge_count':len(graph.get('edges',[])),
                       'ir_graph_node_references':len(referenced_graph_ids),
                       'matched_graph_node_references':len(referenced_graph_ids & graph_ids),
                       'unreferenced_graph_node_count':len(graph_ids-referenced_graph_ids),
                       'missing_graph_node_references':missing_graph_refs,
                       'full_module_snapshots_excluded':True,'full_procedure_snapshots_excluded':True}
    manifest_summary={}
    if manifest:
        manifest_summary={'project':manifest.get('project') if isinstance(manifest,dict) else None,
                          'sample':manifest.get('sample') if isinstance(manifest,dict) else None,
                          'source_files':(manifest.get('files') or manifest.get('source_files') or manifest.get('sources')) if isinstance(manifest,dict) else manifest}
    structural_issues=validate_normalized_ir(normalized)
    issues.extend(structural_issues)
    # Retain upstream diagnostics; no frontend repair is silently accepted here.
    upstream=[copy.deepcopy(x) for x in normalized.get('diagnostics',[])]
    unresolved=sum(x.get('unresolved_expression_count',0) for m in module_reports for x in m['procedures']) + _count_nodes(normalized.get('modules',[]), {'unresolved_expression','unsupported'})
    upstream_blockers=[d for d in upstream if d.get('severity') in {'error','warning'} and d.get('code') in {
        'CROSS_NESTED_BLOCK_CLOSE','IR_FRONTEND_HIERARCHY_REPAIRED'}]
    superseded_opaque_diagnostics=[d for d in upstream if d.get('code')=='IR_EXPRESSIONS_OPAQUE'] if unresolved == 0 else []
    parse_errors=sum(1 for x in issues if x.get('severity')=='error' and x.get('code') in {'EXPRESSION_PARSE_FAILED','CALL_PARSE_FAILED'})
    hard_errors=sum(1 for x in issues if x.get('severity')=='error')
    compatibility_requirements=_compatibility_requirements(normalized)
    compatibility_blockers=[x for x in compatibility_requirements if x.get('severity')=='blocking']
    repair_ids=set()
    for m in normalized.get('modules',[]):
        for p in m.get('procedures',[]): repair_ids.update((p.get('semantics') or {}).get('frontend_repairs',[]))
    pseudocode_gaps=[p['procedure'] for m in module_reports for p in m['procedures'] if p['pseudocode_coverage']=='missing_detailed_section']
    readiness='blocked' if unresolved or parse_errors or hard_errors or upstream_blockers or compatibility_blockers else 'review_required'
    report={
        'analyzer_version':'0.1.0','input_ir_version':normalized.get('ir_version'),'project':normalized.get('name'),
        'status':readiness,'migration_readiness':readiness,
        'summary':{'module_count':len(module_reports),'procedure_count':sum(len(x['procedures']) for x in module_reports),
                   'parsed_expression_count':sum(p['parsed_expression_count'] for m in module_reports for p in m['procedures']) + _count_nodes(normalized.get('modules',[]), {'binary_expression','unary_expression'}),
                   'unresolved_expression_count':unresolved,'parse_error_count':parse_errors,'issue_count':len(issues),
                   'upstream_diagnostic_count':len(upstream),'blocking_upstream_diagnostic_count':len(upstream_blockers),
                   'blocking_compatibility_requirement_count':len(compatibility_blockers),
                   'frontend_hierarchy_repair_count':len(repair_ids),
                   'pseudocode_procedure_sections_missing':len(pseudocode_gaps),
                   'granular_segment_count':len(segment_map),'segment_kinds':segment_kinds,
                   'referenced_segment_count':len(used_segment_ids)},
        'graph_evidence':graph_summary,'manifest_evidence':manifest_summary,
        'compatibility_requirements':compatibility_requirements,
        'pseudocode_coverage_gaps':pseudocode_gaps,
        'upstream_diagnostics':upstream,'superseded_diagnostics':superseded_opaque_diagnostics,'issues':issues,'modules':module_reports,
        'policy':{'full_module_and_procedure_snapshots_used':False,
                  'llm_used':False,'opaque_expressions_guessed':False,
                  'parse_failures_preserved_as_unresolved':True,
                  'frontend_control_flow_repairs_auto_approved':False,
                  'pseudocode_is_supporting_evidence_not_ground_truth':True}
    }
    normalized['semantics']={**(normalized.get('semantics') or {}),
        'normalization':{'analyzer_version':'0.1.0','status':readiness,
                         'expression_parser':'deterministic_vba_subset','unresolved_expression_count':unresolved,
                         'source_evidence_policy':'graph metadata + granular segments; full snapshots excluded'},
        'migration_readiness':readiness}
    normalized['diagnostics']=upstream + [
        {'code':x.get('code'),'severity':x.get('severity'),'message':x.get('message'),
         'procedure':x.get('procedure'),'path':x.get('path'),'source_expression':x.get('source_expression')}
        for x in issues if x.get('severity') in {'error','warning'}
    ]
    return normalized,report


def load_json(path: str | Path) -> Any:
    return json.loads(Path(path).read_text(encoding='utf-8'))


def run_files(*, ir_path: str|Path, graph_path: str|Path|None=None, segments_path: str|Path|None=None,
              pseudocode_path: str|Path|None=None, manifest_path: str|Path|None=None,
              output_dir: str|Path) -> tuple[Path,Path]:
    ir=load_json(ir_path)
    graph=load_json(graph_path) if graph_path else None
    segments=load_json(segments_path) if segments_path else None
    pseudo=Path(pseudocode_path).read_text(encoding='utf-8') if pseudocode_path else None
    manifest=load_json(manifest_path) if manifest_path else None
    normalized,report=normalize_ir(ir,graph=graph,segments=segments,pseudocode=pseudo,manifest=manifest)
    out=Path(output_dir); out.mkdir(parents=True,exist_ok=True)
    ir_out=out/'normalized_ir.json'; report_out=out/'semantic_analysis.json'
    ir_out.write_text(json.dumps(normalized,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
    report_out.write_text(json.dumps(report,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
    return ir_out,report_out
