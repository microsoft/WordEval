"""Complete equation observations with explicit Word child-role semantics."""
from typing import Dict
from .compare import canonical_angle_bracket, normalized_equation_text

# WdOMathFunctionType: Frac has Num/Den; Delim/EqArray.E is OMathArgs.
# https://learn.microsoft.com/office/vba/api/word.wdomathfunctiontype
_CONTAINERS = {
    1: ('Acc', ('E',)), 2: ('Bar', ('E',)), 3: ('Box', ('E',)),
    4: ('BorderBox', ('E',)), 5: ('Delim', ()), 6: ('EqArray', ()),
    7: ('Frac', ('Num', 'Den')), 8: ('Func', ('FName', 'E')),
    9: ('GroupChar', ('E',)), 10: ('LimLow', ('E', 'Lim')),
    11: ('LimUpp', ('E', 'Lim')), 12: ('Mat', ()),
    13: ('Nary', ('E', 'Sub', 'Sup')), 14: ('Phantom', ('E',)),
    15: ('ScrPre', ('E', 'Sub', 'Sup')), 16: ('Rad', ('E', 'Deg')),
    17: ('ScrSub', ('E', 'Sub')), 18: ('ScrSubSup', ('E', 'Sub', 'Sup')),
    19: ('ScrSup', ('E', 'Sup')),
}
_PROPERTIES = {
    1: {'accent_char':'Char'}, 2: {'bar_top':'BarTop'},
    5: {'delim_begin_char':'BegChar', 'delim_end_char':'EndChar', 'delim_sep_char':'SepChar'},
    7: {'frac_type':'Type'},
    9: {'group_char':'Char', 'group_char_top':'CharTop'},
    13: {'nary_char':'Char', 'nary_sub_sup_lim':'SubSupLim', 'nary_hide_sub':'HideSub', 'nary_hide_sup':'HideSup'},
    16: {'rad_hide_degree':'HideDeg'},
}

def _observe(label, read):
    try:
        return read()
    except Exception as exc:
        raise RuntimeError(f'Cannot read equation {label}: {exc}') from exc

def _text(raw):
    return normalized_equation_text({'equation_text': raw or ''})

def extract_equation_info(docx_path: str, *, close_document, open_document) -> Dict:
    """Read every equation and its full text; a read failure is not an empty item."""
    doc = None
    try:
        doc = open_document(docx_path)
        count = _observe('count', lambda: doc.OMaths.Count)
        equations = []
        for index in range(1, count + 1):
            omath = _observe(f'{index}', lambda: doc.OMaths(index))
            value = _observe(f'{index} text', lambda: omath.Range.Text)
            equations.append({'index': index, 'text': value or ''})
        return {'equation_count': count, 'equations': equations}
    finally:
        close_document(doc)

def _argument(omath, depth):
    functions = _observe('argument functions', lambda: omath.Functions)
    count = _observe('argument function count', lambda: functions.Count)
    entries = [extract_omath_func_info(_observe(f'argument function {i}', lambda: functions(i)), i, depth)
               for i in range(1, count + 1)]
    # Concatenate adjacent text runs so harmless Word run splitting is ignored.
    expression = ''.join(item['semantic_expression'] for item in entries)
    if not entries:
        expression = _text(_observe('empty argument text', lambda: omath.Range.Text))
    return expression, entries

def extract_omath_func_info(func, index: int, depth: int = 0) -> Dict:
    """Read the active function's scalars and all named/collection children.

    Missing mandatory observations are indeterminate. A defensive depth bound
    raises instead of accepting two equally truncated observations.
    """
    if depth > 64:
        raise RuntimeError('Equation nesting exceeds the supported observation depth')
    kind = int(_observe('function type', lambda: func.Type))
    info = {'index': index, 'type': kind}
    if kind in (20, 21, 22):
        info['semantic_expression'] = _text(_observe('function text', lambda: func.Range.Text))
        return info
    if kind not in _CONTAINERS:
        raise RuntimeError(f'Unsupported equation function type: {kind}')
    name, roles = _CONTAINERS[kind]
    container = _observe(f'{name} object', lambda: getattr(func, name))
    for key, prop in _PROPERTIES.get(kind, {}).items():
        info[key] = _observe(f'{name}.{prop}', lambda: getattr(container, prop))
    arguments = {}
    for role in roles:
        child = _observe(f'{name}.{role}', lambda: getattr(container, role))
        expression, entries = _argument(child, depth + 1)
        arguments[role] = expression
        if entries:
            key = name.lower() if role == 'E' else f'{name.lower()}_{role.lower()}'
            info[f'{key}_nested'] = entries
    if kind in (5, 6):
        children = _observe(f'{name}.E', lambda: container.E)
        count = _observe(f'{name}.E.Count', lambda: children.Count)
        for i in range(1, count + 1):
            child = _observe(f'{name}.E({i})', lambda: children(i))
            arguments[str(i)], entries = _argument(child, depth + 1)
            info[f'{name.lower()}_{i}_nested'] = entries
    if kind == 12:
        info['mat_rows'] = _observe('matrix row count', lambda: container.Rows.Count)
        info['mat_cols'] = _observe('matrix column count', lambda: container.Cols.Count)
        for row in range(1, info['mat_rows'] + 1):
            for col in range(1, info['mat_cols'] + 1):
                child = _observe(f'matrix cell {row},{col}', lambda: container.Cell(row, col))
                arguments[f'{row},{col}'], entries = _argument(child, depth + 1)
                info[f'mat_{row}_{col}_nested'] = entries
    info['argument_expressions'] = arguments
    e = arguments.get('E', '')
    if kind == 7:
        expression = f"({arguments['Num']})/({arguments['Den']})"
    elif kind == 8:
        expression = f"func({arguments['FName']};{e})"
    elif kind in (10, 11):
        expression = f"{name}({e};{arguments['Lim']})"
    elif kind == 16:
        expression = f'sqrt({e})' if info['rad_hide_degree'] else f"root({arguments['Deg']};{e})"
    elif kind == 17:
        expression = f"{e}_({arguments['Sub']})"
    elif kind == 19:
        expression = f"{e}^({arguments['Sup']})"
    elif kind in (15, 18):
        scripts = f"_({arguments['Sub']})^({arguments['Sup']})"
        expression = scripts + e if kind == 15 else e + scripts
    elif kind == 5:
        begin = chr(canonical_angle_bracket(info['delim_begin_char']) or 0)
        end = chr(canonical_angle_bracket(info['delim_end_char']) or 0)
        expression = begin + '|'.join(arguments.values()) + end
    elif kind in (3, 4):
        expression = e
    else:
        expression = f'{name}(' + ';'.join(f'{k}={v}' for k, v in arguments.items()) + ')'
    info['semantic_expression'] = expression
    return info

def extract_equation_structure(docx_path: str, *, close_document, open_document) -> Dict:
    """Read all equations/functions without truncation or swallowed failures."""
    doc = None
    try:
        doc = open_document(docx_path)
        count = _observe('count', lambda: doc.OMaths.Count)
        result = {'equation_count': count, 'equation_text': '', 'function_count': 0,
                  'functions': [], 'equations': []}
        for index in range(1, count + 1):
            omath = _observe(f'{index}', lambda: doc.OMaths(index))
            functions = _observe(f'{index} functions', lambda: omath.Functions)
            function_count = _observe(f'{index} function count', lambda: functions.Count)
            entries = [extract_omath_func_info(_observe(f'{index} function {i}', lambda: functions(i)), i)
                       for i in range(1, function_count + 1)]
            result['equations'].append({
                'omath_type': _observe(f'{index} type', lambda: omath.Type),
                'equation_text': _observe(f'{index} text', lambda: omath.Range.Text) or '',
                'function_count': function_count, 'functions': entries,
                'semantic_expression': ''.join(item['semantic_expression'] for item in entries),
            })
        if result['equations']:
            first = result['equations'][0]
            for key in ('equation_text', 'function_count', 'functions', 'semantic_expression'):
                result[key] = first[key]
        return result
    finally:
        close_document(doc)
