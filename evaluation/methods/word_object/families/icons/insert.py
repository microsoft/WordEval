"""families / icons / insert with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.families.icons import (
    insert_saved as families_icons_insert_saved,
)
from pathlib import Path
from typing import Any, Dict, List, Optional
import math


def expected_svg_icon_contract(spec: Dict[str, Any]) -> Dict[str, Any]:
    variation = spec.get("variation", "")
    contracts = {
        "insert_blue_star_icon": {"shape": "star", "fill": "#4472C4"},
        "insert_orange_circle_icon": {"shape": "circle", "fill": "#ED7D31"},
    }
    return {
        **contracts.get(variation, {}),
        "width": 150.0,
        "height": 150.0,
    }


def svg_icon_com_contract_failures(state: Dict[str, Any], expected: Dict[str, Any]) -> List[str]:
    failures = []
    if not expected.get("shape") or not expected.get("fill"):
        failures.append("test contract lacks SVG shape or fill color")
        return failures
    if state.get("visual_object_count") != 1:
        failures.append(
            f"expected one body visual object, found {state.get('visual_object_count')}"
        )
    for dimension in ("width", "height"):
        actual = state.get(dimension)
        target = expected[dimension]
        if actual is None or abs(float(actual) - target) > 1.0:
            failures.append(f"{dimension} {actual} != {target}")
    return failures


def svg_icon_saved_contract_failures(state: Dict[str, Any], expected: Dict[str, Any]) -> List[str]:
    failures = list(state.get("errors") or [])
    if state.get("object_count") != 1:
        failures.append(f"expected one referenced SVG icon, found {state.get('object_count')}")
    if state.get("body_object_count") != 1:
        failures.append(f"expected one saved body object, found {state.get('body_object_count')}")
    icons = state.get("icons") or []
    if len(icons) != 1:
        return failures
    icon = icons[0]
    if not icon.get("has_svg_root"):
        failures.append("referenced image is not an SVG document")
    if icon.get("shape") != expected.get("shape"):
        failures.append(f"SVG shape {icon.get('shape')} != {expected.get('shape')}")
    if icon.get("fill_colors") != [expected.get("fill")]:
        failures.append(f"SVG effective fill {icon.get('fill_colors')} != {[expected.get('fill')]}")
    for dimension in ("width", "height"):
        actual = icon.get(dimension)
        target = expected.get(dimension)
        if actual is None or abs(float(actual) - target) > 1.0:
            failures.append(f"saved {dimension} {actual} != {target}")
    viewport = icon.get("effective_viewport")
    if (
        not isinstance(viewport, list)
        or len(viewport) != 4
        or not all(isinstance(value, (int, float)) for value in viewport)
        or not all(math.isfinite(float(value)) for value in viewport)
        or viewport[2] <= 0
        or viewport[3] <= 0
        or abs(float(viewport[2]) / float(viewport[3]) - 1.0) > 0.001
    ):
        failures.append("SVG effective viewport is missing, invalid, or non-square")
    return failures


def verify_icon_insert_saved(
    pred_path: Path,
    gold_path: Path,
    spec: Dict[str, Any],
    *,
    canonical_chart_value,
    error_with_details,
    saved_part_relationships,
    verify_saved_insertion_context,
) -> VerificationResult:
    expected = expected_svg_icon_contract(spec)
    try:
        gold = families_icons_insert_saved.extract_saved_svg_icon_signature(
            gold_path,
            canonical_chart_value=canonical_chart_value,
            saved_part_relationships=saved_part_relationships,
        )
    except Exception as exc:
        return error_with_details(f"Cannot extract gold saved SVG icon evidence: {exc}")
    gold_failures = svg_icon_saved_contract_failures(gold, expected)
    if gold_failures:
        return error_with_details(
            "Gold saved SVG icon evidence is incomplete or invalid",
            {"failures": gold_failures, "gold": gold, "expected": expected},
        )
    try:
        pred = families_icons_insert_saved.extract_saved_svg_icon_signature(
            pred_path,
            canonical_chart_value=canonical_chart_value,
            saved_part_relationships=saved_part_relationships,
        )
    except Exception as exc:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Predicted saved SVG icon package is unreadable ({exc})",
            details={"gold": gold, "expected": expected},
        )
    failures = svg_icon_saved_contract_failures(pred, expected)
    context_failure = verify_saved_insertion_context(pred_path, gold_path)
    if context_failure is not None:
        failures.append("surrounding document content/layout changed")
    return VerificationResult(
        passed=not failures,
        message=(
            "PASS: Referenced SVG identity, fill, geometry, size, and context are valid"
            if not failures
            else f"FAIL: Saved SVG icon insertion differs ({'; '.join(failures)})"
        ),
        details={
            "match_mode": "saved_ooxml_referenced_svg_payload",
            "failures": failures,
            "gold": gold,
            "pred": pred,
            "expected": expected,
            "context_failure": (context_failure.to_dict() if context_failure else None),
        },
    )


def verify_icon_insert(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    canonical_chart_value,
    com_primary_with_saved_guardrail,
    error_with_details,
    read_inserted_svg_icon_state,
    saved_part_relationships,
    verify_saved_insertion_context,
) -> VerificationResult:
    """Require one referenced SVG icon; inline and floating are equivalent."""
    spec = spec or {}
    expected = expected_svg_icon_contract(spec)

    def verify_com() -> VerificationResult:
        gold = read_inserted_svg_icon_state(str(gold_path))
        pred = read_inserted_svg_icon_state(str(pred_path))
        gold_failures = svg_icon_com_contract_failures(gold, expected)
        if gold_failures:
            return error_with_details(
                "Gold SVG icon display evidence contradicts the test contract",
                {"failures": gold_failures, "gold": gold, "expected": expected},
            )
        failures = svg_icon_com_contract_failures(pred, expected)
        return VerificationResult(
            passed=not failures,
            message=(
                "PASS: One displayed icon has the requested dimensions"
                if not failures
                else f"FAIL: Icon display state differs ({'; '.join(failures)})"
            ),
            details={
                "failures": failures,
                "gold": gold,
                "pred": pred,
                "expected": expected,
            },
        )

    return com_primary_with_saved_guardrail(
        verify_com,
        lambda: verify_icon_insert_saved(
            pred_path,
            gold_path,
            spec,
            canonical_chart_value=canonical_chart_value,
            error_with_details=error_with_details,
            saved_part_relationships=saved_part_relationships,
            verify_saved_insertion_context=verify_saved_insertion_context,
        ),
    )
