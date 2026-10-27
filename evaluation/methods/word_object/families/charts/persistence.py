"""Editable chart backing and ordered cache/formula consistency."""
from .extract_saved import extract_insertion_signature
from .workbook import project_workbook
from ....base import VerificationResult


def _failures(state):
    failures = list(state.get('errors', []))
    charts = state.get('charts', [])
    if len(charts) != 1:
        return failures + ['expected one referenced chart']
    chart = charts[0]
    cells = chart.get('workbook_cells')
    if not cells:
        return failures + ['embedded editable workbook is missing or invalid']
    if not chart.get('series'):
        failures.append('chart has no data series')
    for index, series in enumerate(chart.get('series', []),1):
        for role in ('name','categories','values'):
            projected = project_workbook(cells, series.get('formulas',{}).get(role))
            cache = [series.get(role)] if role == 'name' else series.get(role)
            if not projected or projected != cache:
                failures.append(f'series {index} {role} cache differs from its workbook reference')
    return failures


def verify_switch_persistence(pred_path, gold_path):
    try:
        gold = extract_insertion_signature(gold_path)
        gold_failures = _failures(gold)
    except Exception as exc:
        return VerificationResult.error_result(f'Cannot read gold chart backing: {exc}')
    if gold_failures:
        result = VerificationResult.error_result('Gold chart backing is invalid')
        result.details.update(error_owner='gold_fixture',failures=gold_failures)
        return result
    try:
        pred = extract_insertion_signature(pred_path)
        failures = _failures(pred)
    except OSError as exc:
        return VerificationResult.error_result(f'Cannot read prediction chart backing: {exc}')
    except Exception as exc:
        pred = None
        failures = [f'prediction chart backing is invalid: {exc}']
    return VerificationResult(passed=not failures,
        message=('PASS: Chart workbook and cache are consistent' if not failures
                 else 'FAIL: Chart workbook or cache is invalid'),
        details={'verification_source':'saved_ooxml_chart_workbook_guardrail',
                 'failures':failures,'gold':gold,'pred':pred})


def verify_data_preservation(pred_path, gold_path):
    """Existing editable charts must retain data in both cache and workbook.

    Uses referenced parts and projected formula values, never workbook bytes or
    cell addresses as an equality oracle. Insertion has a separate contract for
    literal caches and workbook-free native representations.
    """
    result = verify_switch_persistence(pred_path, gold_path)
    if not result.passed:
        return result
    gold = result.details['gold']['charts'][0]['series']
    pred = result.details['pred']['charts'][0]['series']
    projection = lambda series: [
        {role: item.get(role) for role in ('name', 'categories', 'values')}
        for item in series
    ]
    if projection(gold) != projection(pred):
        return VerificationResult(
            passed=False, message='FAIL: Saved chart data changed',
            details={'verification_source': 'saved_ooxml_chart_workbook_guardrail',
                     'gold_data': projection(gold), 'pred_data': projection(pred)})
    return result
