"""families / shapes / insert_bevel with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.policies import object_effect as policies_object_effect
from pathlib import Path
from typing import Any, Dict, Optional


def verify_insert_then_bevel_saved(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    extract_bevel_render_signatures,
    extract_saved_shape_properties,
    match_shapes_by_name,
) -> VerificationResult:
    """Verify the persisted state of the prerequisite bevel case.

    Word occasionally hangs while opening this particular generated shape.
    The requested state is fully persisted in DrawingML: shape kind and
    geometry live on the anchor/shape, while the bevel preset and the
    required 3-D render scene live in ``sp3d``/``scene3d``.  Keep this
    fallback scoped to ``insert_then_bevel`` so other object-effect cases
    continue to use the Word Object Model after a timeout.
    """
    try:
        gold_shapes = extract_saved_shape_properties(gold_path)
    except Exception as exc:
        return policies_object_effect.saved_evidence_error(
            f"Gold saved shape extraction failed: {exc}",
            evidence_owner="gold",
            exception=exc,
        )
    if not gold_shapes:
        return policies_object_effect.saved_evidence_error(
            "Gold saved shape evidence is empty",
            evidence_owner="gold",
        )
    try:
        pred_shapes = extract_saved_shape_properties(pred_path)
    except Exception as exc:
        return policies_object_effect.saved_evidence_error(
            f"Prediction saved shape extraction failed: {exc}",
            evidence_owner="prediction",
            exception=exc,
        )
    pred_shapes, gold_shapes = match_shapes_by_name(pred_shapes, gold_shapes)

    geometry_targets = []
    shape_count_matches = len(pred_shapes) == len(gold_shapes)
    property_map = {
        "auto_shape_type": "auto_shape_type",
        "actual_left": "actual_left",
        "actual_top": "actual_top",
        "width": "width",
        "height": "height",
    }

    def saved_value(shape: Dict[str, Any], prop: Optional[str]) -> Any:
        if prop == "actual_left":
            value = float(shape.get("left") or 0.0)
            # Word stores offsets relative to the selected frame.  This
            # case uses either Page or the first Column/Margin frame.
            if shape.get("relative_h_position") in {0, 2, 4}:
                value += float(shape.get("left_margin") or 0.0)
            return round(value, 4)
        if prop == "actual_top":
            value = float(shape.get("top") or 0.0)
            # The generated gold is paragraph-relative at the first body
            # paragraph; its rendered origin is the top margin.  Model
            # outputs may save the equivalent coordinate page-relative.
            if shape.get("relative_v_position") in {0, 2, 4}:
                value += float(shape.get("top_margin") or 0.0)
            return round(value, 4)
        return shape.get(prop) if prop is not None else None

    intents = (spec or {}).get("composition", {}).get("atomic_intents", [])
    insert_intent = next(
        (intent for intent in intents if intent.get("operation") == "insert_shape"),
        {},
    )
    for target in insert_intent.get("verification_targets", []) or []:
        prop = target.get("property")
        saved_prop = property_map.get(prop)
        gold_values = [saved_value(shape, saved_prop) for shape in gold_shapes]
        pred_values = [saved_value(shape, saved_prop) for shape in pred_shapes]
        tolerance = float(target.get("tolerance", 0.0) or 0.0)
        if saved_prop is None or not shape_count_matches:
            passed = False
        elif tolerance:
            passed = all(
                gold is not None
                and pred is not None
                and abs(float(pred) - float(gold)) <= tolerance
                for pred, gold in zip(pred_values, gold_values)
            )
        else:
            passed = pred_values == gold_values
        geometry_targets.append(
            {
                "target": target,
                "passed": passed,
                "gold": gold_values,
                "pred": pred_values,
            }
        )

    try:
        gold_bevels = extract_bevel_render_signatures(gold_path)
    except Exception as exc:
        return policies_object_effect.saved_evidence_error(
            f"Gold saved bevel extraction failed: {exc}",
            evidence_owner="gold",
            exception=exc,
        )
    if not gold_bevels:
        return policies_object_effect.saved_evidence_error(
            "Gold saved bevel evidence is empty",
            evidence_owner="gold",
        )
    try:
        pred_bevels = extract_bevel_render_signatures(pred_path)
    except Exception as exc:
        return policies_object_effect.saved_evidence_error(
            f"Prediction saved bevel extraction failed: {exc}",
            evidence_owner="prediction",
            exception=exc,
        )
    bevel_targets = []
    bevel_intent = next(
        (intent for intent in intents if intent.get("operation") == "shape_bevel"),
        {},
    )
    bevel_property_map = {
        "threed_bevel_top_type": "bevel_prst",
        "scene_present": "scene_present",
        "camera_present": "camera_present",
        "light_rig_present": "light_rig_present",
    }
    for target in bevel_intent.get("verification_targets", []) or []:
        prop = target.get("property")
        saved_prop = bevel_property_map.get(prop)
        gold_values = [item.get(saved_prop) for item in gold_bevels]
        pred_values = [item.get(saved_prop) for item in pred_bevels]
        bevel_targets.append(
            {
                "target": target,
                "passed": (
                    saved_prop is not None
                    and len(pred_bevels) == len(gold_bevels)
                    and pred_values == gold_values
                ),
                "gold": gold_values,
                "pred": pred_values,
            }
        )

    atomic_results = []
    for intent, target_results in (
        (insert_intent, geometry_targets),
        (bevel_intent, bevel_targets),
    ):
        atomic_results.append(
            {
                "intent_id": intent.get("id"),
                "operation": intent.get("operation"),
                "atomic_case": intent.get("atomic_case"),
                "passed": bool(target_results)
                and all(result.get("passed") for result in target_results),
                "target_results": target_results,
            }
        )
    flat_targets = [
        {
            **target_result,
            "intent_id": intent_result.get("intent_id"),
            "operation": intent_result.get("operation"),
            "atomic_case": intent_result.get("atomic_case"),
        }
        for intent_result in atomic_results
        for target_result in intent_result.get("target_results", [])
    ]
    passed = (
        shape_count_matches
        and bool(atomic_results)
        and all(result.get("passed") for result in atomic_results)
    )
    return VerificationResult(
        passed=passed,
        message=(
            "PASS: Saved inserted shape and bevel state match"
            if passed
            else "FAIL: Saved inserted shape or bevel state differs"
        ),
        details={
            "match_mode": "saved_ooxml_insert_then_bevel",
            "fallback_contract": "complete",
            "shape_count_match": shape_count_matches,
            "atomic_results": atomic_results,
            "target_results": flat_targets,
            "atomic_summary": {
                "passed_intents": sum(1 for result in atomic_results if result.get("passed")),
                "total_intents": len(atomic_results),
                "passed_targets": sum(1 for result in flat_targets if result.get("passed")),
                "total_targets": len(flat_targets),
            },
        },
    )
