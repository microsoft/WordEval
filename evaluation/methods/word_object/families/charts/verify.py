"""Chart effect and composition evaluation with an explicit COM reader."""

from pathlib import Path
from typing import Callable, Dict, List, Optional

from ....base import (
    VerificationResult,
    combine_com_primary_with_ooxml_guardrail,
    tag_verification_source,
)
from ...core import DocumentOpenError
from ...composition import (
    composition_command_ids,
    composition_steps,
    composition_step_key,
    composition_step_declaration_error,
    unhandled_composition_step,
)
from .commands import EFFECT_COMMANDS
from .compare import diff_command, diff_operation
from .contract import validate_chart_observations, diff_data
from .extract_saved import extract_properties
from .persistence import verify_switch_persistence, verify_data_preservation


def verify_effect(pred_path, gold_path, spec=None, *, read_com):
    result = _verify_effect(pred_path, gold_path, spec, read_com=read_com)
    if not result.passed:
        return result
    return combine_com_primary_with_ooxml_guardrail(
        result, verify_data_preservation(pred_path, gold_path))


def _verify_effect(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    read_com: Callable[[str], Dict],
) -> VerificationResult:
    """
    Verify chart properties match.

    Handles ChartStylesGallery, ChartColorsGallery, ChartLayoutGallery,
    ChartLegend, ChartDataLabel, ChartDataTable, ChartTrendline,
    ChartErrorBars, ChartTitle, ChartGridlinesGallery, ChartLines,
    ChartUpDownBars, ColorPickerChartLineStyles.
    """
    gold_chart = read_com(str(gold_path))
    pred_chart = read_com(str(pred_path))

    if "error" in gold_chart:
        return VerificationResult.error_result(f"Gold: {gold_chart['error']}")
    if "error" in pred_chart:
        if not pred_chart.get("absent"):
            return VerificationResult.error_result(f"Prediction: {pred_chart['error']}", observation_owner="pred")
        return VerificationResult(
            passed=False,
            message=f"FAIL: Predicted chart is unavailable ({pred_chart['error']})",
            details={
                "pred_error": pred_chart["error"],
                "idMso": spec.get("idMso", "") if spec else "",
            },
        )

    idMso = spec.get("idMso", "") if spec else ""
    observation_error = validate_chart_observations(gold_chart, pred_chart, [idMso], spec)
    if observation_error is not None:
        return observation_error
    differences = diff_command(idMso, gold_chart, pred_chart, spec)
    differences.extend(diff_data(gold_chart, pred_chart))

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Chart mismatch ({len(differences)} difference(s))",
            details={
                "differences": differences,
                "idMso": idMso,
                "gold_chart": gold_chart,
                "pred_chart": pred_chart,
            },
        )

    com_result = VerificationResult(
        passed=True, message="PASS: Chart properties match", details={"idMso": idMso}
    )
    if idMso == "ChartSwitchRowColumn":
        return combine_com_primary_with_ooxml_guardrail(
            com_result, verify_switch_persistence(pred_path, gold_path))
    if idMso == "ColorPickerChartLineStyles":
        gold_saved = extract_properties(gold_path)
        pred_saved = extract_properties(pred_path)
        if not gold_saved.get("series_order_valid"):
            message = "Gold saved chart series order is invalid or ambiguous"
            return VerificationResult(
                passed=False,
                message=f"ERROR: {message}",
                details={
                    "idMso": idMso,
                    "match_mode": "saved_ooxml_chart_line_guardrail",
                    "gold_series_order_issues": gold_saved.get("series_order_issues", []),
                    "error_owner": "gold_fixture_or_evaluator",
                },
                error=message,
            )
        gold_collateral = sorted(
            (
                series
                for series in gold_saved.get("series_line_fills", [])
                if series.get("index") != 1
            ),
            key=lambda series: (
                series.get("index") is None,
                series.get("index") or 0,
            ),
        )
        pred_collateral = sorted(
            (
                series
                for series in pred_saved.get("series_line_fills", [])
                if series.get("index") != 1
            ),
            key=lambda series: (
                series.get("index") is None,
                series.get("index") or 0,
            ),
        )
        guardrail_differences = []
        if not pred_saved.get("series_order_valid"):
            guardrail_differences.append(
                {
                    "property": "series_order",
                    "gold": gold_saved.get("series_order_values", []),
                    "pred": pred_saved.get("series_order_values", []),
                    "pred_issues": pred_saved.get("series_order_issues", []),
                }
            )
        checked_properties = ["non_target_series_line_fills"]
        if not pred_saved.get("series_order_valid"):
            checked_properties.append("series_order")
        if gold_collateral != pred_collateral:
            guardrail_differences.append(
                {
                    "property": "non_target_series_line_fills",
                    "gold": gold_collateral,
                    "pred": pred_collateral,
                }
            )
        return combine_com_primary_with_ooxml_guardrail(
            com_result,
            VerificationResult(
                passed=not guardrail_differences,
                message=(
                    "PASS: Non-target series line fills are preserved"
                    if not guardrail_differences
                    else "FAIL: Non-target series line fills changed"
                ),
                details={
                    "match_mode": "saved_ooxml_chart_line_guardrail",
                    "properties": checked_properties,
                    "differences": guardrail_differences,
                },
            ),
        )
    return com_result


def verify_composition(pred_path, gold_path, spec=None, *, read_com):
    result = _verify_composition(pred_path, gold_path, spec, read_com=read_com)
    if not result.passed:
        return result
    commands = composition_command_ids(spec)
    inserts = "ChartInsert" in commands or any(
        step.get("operation") == "insert_chart" for step in composition_steps(spec))
    if inserts:
        return result
    backing = verify_data_preservation(pred_path, gold_path)
    if result.details.get("verification_source") == "saved_ooxml_fallback":
        # A saved check cannot retrospectively claim COM primary evidence.
        result.details["chart_workbook_guardrail"] = backing.to_dict()
        return result if backing.passed else tag_verification_source(backing, "saved_ooxml_fallback")
    return combine_com_primary_with_ooxml_guardrail(result, backing)


def _verify_composition(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    read_com: Callable[[str], Dict],
) -> VerificationResult:
    """
    Verify chart composition by checking ONLY what the composition's steps
    explicitly test — aligned with atomic per-idMso scoping.

    For each step in spec.composition.steps:
      - if it has an idMso, use _chart_diffs_for_idmso (same logic as the
        atomic _verify_chart for that command);
      - if it has an operation name (insert_chart, change_chart_type, …),
        check the step's explicit params via _chart_diffs_for_operation.

    Falls back to composition atomic intents, then legacy commands, when no
    steps are present. This way the composition checks the
    same properties the equivalent atomic tests check — no over-checking
    of unrelated side-effects (see N4 in
    the composition review findings).
    """
    declaration_error = composition_step_declaration_error("chart", spec)
    if declaration_error is not None:
        return declaration_error

    composition = (spec or {}).get("composition", {}) or {}
    steps = composition.get("steps") or []
    commands = composition_command_ids(spec)
    supported_operations = {
        "insert_chart",
        "change_chart_type",
        "add_drop_lines",
        "add_hi_lo_lines",
        "add_up_down_bars",
        "add_chart_title",
        "format_chart_title",
    }
    supported_commands = EFFECT_COMMANDS | {"ChartInsert"}
    seen_steps: set[str] = set()
    normalized_steps = composition_steps(spec)
    for step_index, step in enumerate(normalized_steps, start=1):
        _, step_error = composition_step_key("chart", step_index, step, seen_steps)
        if step_error is not None:
            return step_error
        id_mso = step.get("idMso")
        operation = step.get("operation")
        if id_mso:
            if id_mso not in supported_commands:
                return unhandled_composition_step("chart", step_index, step)
        elif operation not in supported_operations:
            return unhandled_composition_step("chart", step_index, step)
        elif operation == "change_chart_type":
            params = step.get("params")
            if not isinstance(params, dict) or params.get("chart_type") is None:
                result = VerificationResult.error_result(
                    "chart composition change_chart_type step has no " "target chart_type"
                )
                result.details.update(
                    {
                        "route": "chart",
                        "step_index": step_index,
                        "step": step,
                        "reason": "malformed_composition_step_parameters",
                    }
                )
                return result

    if not steps:
        for step_index, step in enumerate(normalized_steps, start=1):
            if not step.get("idMso"):
                return unhandled_composition_step("chart", step_index, step)

    saved_chart_commands = {
        "ChartInsert",
        "ChartTitle",
        "ChartLegend",
        "ChartDataLabel",
        "ChartColorsGallery",
        "ChartTrendline",
    }
    verification_source = "word_com"
    com_error = None
    try:
        gold_chart = read_com(str(gold_path))
        pred_chart = read_com(str(pred_path))
    except DocumentOpenError:
        raise
    except Exception as exc:
        if not (commands and set(commands) <= saved_chart_commands):
            raise
        com_error = str(exc)
        gold_chart = extract_properties(gold_path)
        pred_chart = extract_properties(pred_path)
        verification_source = "saved_ooxml_fallback"

    if "error" in gold_chart:
        return VerificationResult.error_result(f"Gold: {gold_chart['error']}")
    if "error" in pred_chart:
        if not pred_chart.get("absent"):
            return VerificationResult.error_result(f"Prediction: {pred_chart['error']}", observation_owner="pred")
        return VerificationResult(
            passed=False,
            message=f"FAIL: Predicted chart is unavailable ({pred_chart['error']})",
            details={
                "pred_error": pred_chart["error"],
                "scoped_to": commands,
                "verification_source": verification_source,
            },
        )

    if verification_source == "word_com":
        observation_error = validate_chart_observations(gold_chart, pred_chart, commands, spec)
        if observation_error is not None:
            return tag_verification_source(observation_error, verification_source)
    differences: List[Dict] = []
    scope_summary: List[str] = []

    if steps:
        # Per-step scoping: each step contributes only its own checks.
        for step in steps:
            if not isinstance(step, dict):
                continue
            if step.get("idMso"):
                idmso = step["idMso"]
                differences.extend(diff_command(idmso, gold_chart, pred_chart, spec))
                scope_summary.append(f"idMso:{idmso}")
            elif "operation" in step:
                op = step["operation"]
                params = step.get("params") or {}
                differences.extend(diff_operation(op, params, gold_chart, pred_chart))
                scope_summary.append(f"op:{op}")
    elif commands:
        # AdditiveChartComposition format: list of idMso, no per-step params.
        # Take the union of per-idMso diffs.
        for idmso in commands:
            differences.extend(diff_command(idmso, gold_chart, pred_chart, spec))
            scope_summary.append(f"idMso:{idmso}")
    else:
        # No spec scoping info — fall back to chart-level structure check
        # (chart_type + series_count). Avoids over-checking but still
        # catches gross mismatches.
        if gold_chart.get("chart_type") != pred_chart.get("chart_type"):
            differences.append(
                {
                    "property": "chart_type",
                    "gold": gold_chart.get("chart_type"),
                    "pred": pred_chart.get("chart_type"),
                }
            )
        gold_series = gold_chart.get("series_info", []) or []
        pred_series = pred_chart.get("series_info", []) or []
        if len(gold_series) != len(pred_series):
            differences.append(
                {
                    "property": "series_count",
                    "gold": len(gold_series),
                    "pred": len(pred_series),
                }
            )
        scope_summary.append("fallback:chart_type+series_count")

    # A saved mismatch is sufficient failure evidence even if that projection
    # cannot establish the rest of the contract. A match still needs data.
    if verification_source == "word_com" or not differences:
        observation_error = validate_chart_observations(gold_chart, pred_chart, commands, spec)
        if observation_error is not None:
            return tag_verification_source(observation_error, verification_source)
        differences.extend(diff_data(gold_chart, pred_chart))

    # De-duplicate diffs by property+location signature (multiple steps
    # may legitimately check the same property; avoid double-reporting).
    seen = set()
    unique_diffs = []
    for d in differences:
        key = (d.get("property"), str(d.get("gold")), str(d.get("pred")))
        if key in seen:
            continue
        seen.add(key)
        unique_diffs.append(d)
    differences = unique_diffs

    if differences:
        return tag_verification_source(
            VerificationResult(
                passed=False,
                message=f"FAIL: Chart composition mismatch ({len(differences)} difference(s))",
                details={
                    "differences": differences,
                    "scoped_to": scope_summary,
                    "gold_chart": gold_chart,
                    "pred_chart": pred_chart,
                },
            ),
            verification_source,
            **({"word_com_error": com_error} if com_error else {}),
        )

    com_result = VerificationResult(
        passed=True,
        message="PASS: Chart composition properties match",
        details={
            "operation_count": spec.get("composition", {}).get("operation_count"),
        },
    )
    should_check_saved_guardrail = commands and (
        set(commands) <= saved_chart_commands or "ColorPickerChartLineStyles" in commands
    )
    if "ChartSwitchRowColumn" in commands:
        backing = verify_switch_persistence(pred_path, gold_path)
        if not backing.passed:
            return combine_com_primary_with_ooxml_guardrail(com_result, backing)
        com_result.details['chart_workbook_guardrail'] = backing.to_dict()
    if verification_source == "word_com" and should_check_saved_guardrail:
        gold_saved = extract_properties(gold_path)
        pred_saved = extract_properties(pred_path)
        guardrail_properties = set()
        guardrail_by_command = {
            "ChartInsert": {
                "chart_type",
                "series_count",
                "series_names",
                "series_values",
                "categories",
            },
            "ChartTitle": {
                # Word can persist the untouched default "Chart Title" as
                # an empty title text node while exposing the visible
                # placeholder through ChartTitle.Text after reopen.  COM
                # already checks the effective title text above, so the
                # saved guardrail should only require the persisted title
                # object and its layout mode.
                "has_title",
                "title_include_in_layout",
            },
            "ChartLegend": {"has_legend", "legend_position"},
            "ChartDataLabel": {"series_data_labels"},
            "ChartColorsGallery": {"chart_color"},
            "ChartTrendline": {"series_trendlines"},
        }
        if set(commands) <= saved_chart_commands:
            for command in commands:
                guardrail_properties.update(guardrail_by_command.get(command, set()))
        if "ColorPickerChartLineStyles" in commands:
            guardrail_properties.add("non_target_series_line_fills")
        guardrail_differences = [
            {
                "property": prop,
                "gold": gold_saved.get(prop),
                "pred": pred_saved.get(prop),
            }
            for prop in sorted(guardrail_properties)
            if prop != "non_target_series_line_fills"
            if gold_saved.get(prop) != pred_saved.get(prop)
        ]
        if "non_target_series_line_fills" in guardrail_properties:
            if not gold_saved.get("series_order_valid"):
                message = "Gold saved chart series order is invalid or ambiguous"
                return VerificationResult(
                    passed=False,
                    message=f"ERROR: {message}",
                    details={
                        "match_mode": "saved_ooxml_chart_guardrail",
                        "gold_series_order_issues": gold_saved.get("series_order_issues", []),
                        "error_owner": "gold_fixture_or_evaluator",
                    },
                    error=message,
                )
            gold_collateral = sorted(
                (
                    series
                    for series in gold_saved.get("series_line_fills", [])
                    if series.get("index") != 1
                ),
                key=lambda series: (
                    series.get("index") is None,
                    series.get("index") or 0,
                ),
            )
            pred_collateral = sorted(
                (
                    series
                    for series in pred_saved.get("series_line_fills", [])
                    if series.get("index") != 1
                ),
                key=lambda series: (
                    series.get("index") is None,
                    series.get("index") or 0,
                ),
            )
            if not pred_saved.get("series_order_valid"):
                guardrail_properties.add("series_order")
                guardrail_differences.append(
                    {
                        "property": "series_order",
                        "gold": gold_saved.get("series_order_values", []),
                        "pred": pred_saved.get("series_order_values", []),
                        "pred_issues": pred_saved.get("series_order_issues", []),
                    }
                )
            if gold_collateral != pred_collateral:
                guardrail_differences.append(
                    {
                        "property": "non_target_series_line_fills",
                        "gold": gold_collateral,
                        "pred": pred_collateral,
                    }
                )
        if not guardrail_properties:
            return tag_verification_source(com_result, verification_source)
        saved_guardrail = VerificationResult(
            passed=not guardrail_differences,
            message=(
                "PASS: Saved chart state matches"
                if not guardrail_differences
                else "FAIL: Saved chart state differs"
            ),
            details={
                "match_mode": "saved_ooxml_chart_guardrail",
                "properties": sorted(guardrail_properties),
                "differences": guardrail_differences,
            },
        )
        return combine_com_primary_with_ooxml_guardrail(com_result, saved_guardrail)
    return tag_verification_source(
        com_result,
        verification_source,
        **({"word_com_error": com_error} if com_error else {}),
    )
