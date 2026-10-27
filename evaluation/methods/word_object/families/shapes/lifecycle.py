"""families / shapes / lifecycle with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.core import DocumentOpenError
from evaluation.methods.word_object.families.shapes import (
    extract_saved as families_shapes_extract_saved,
)
from evaluation.methods.word_object.families.shapes import geometry as families_shapes_geometry
from evaluation.methods.word_object.families.shapes import layout as families_shapes_layout
from pathlib import Path
from typing import Dict, Optional
from ...observations import require_observations


def _lifecycle_observations(gold, pred, required):
    if not gold:
        return VerificationResult.error_result("Gold shape target is missing", observation_owner="gold")
    for owner, shapes in (("gold", gold), ("pred", pred)):
        for shape in shapes:
            fields = set(required)
            if "is_group" in fields and shape.get("is_group"):
                fields.add("group_item_count")
            error = require_observations({owner: shape}, fields, context="shape lifecycle")
            if error is not None:
                return error
    return None


def verify_grouping(
    pred_path: Path, gold_path: Path, *, read_shape_properties
) -> VerificationResult:
    """Verify shape grouping structure matches."""
    try:
        gold_shapes = read_shape_properties(str(gold_path))
    except DocumentOpenError:
        raise
    except Exception as com_error:
        try:
            defect = families_shapes_extract_saved.saved_group_root_defect(gold_path)
        except Exception:
            raise com_error
        if defect is None:
            raise
        error = "Gold grouped DrawingML uses invalid wpg:wGrp instead of " "wpg:wgp"
        return VerificationResult(
            passed=False,
            message=f"ERROR: {error}",
            details={
                "verification_source": ("word_com_failure+saved_ooxml_invalid_group"),
                "error_owner": "gold_fixture",
                "word_com_error": str(com_error),
                "saved_group_evidence": defect,
            },
            error=error,
        )

    try:
        pred_shapes = read_shape_properties(str(pred_path))
    except DocumentOpenError:
        raise
    except Exception as com_error:
        try:
            defect = families_shapes_extract_saved.saved_group_root_defect(pred_path)
        except Exception:
            raise com_error
        if defect is None:
            raise
        return VerificationResult(
            passed=False,
            message=(
                "FAIL: Predicted grouped DrawingML uses invalid wpg:wGrp " "instead of wpg:wgp"
            ),
            details={
                "verification_source": ("word_com_failure+saved_ooxml_invalid_group"),
                "reason": "invalid_output",
                "failure_owner": "prediction_saved_state",
                "detected_format": "schema_invalid_ooxml",
                "invalid_element": "wpg:wGrp",
                "expected_element": "wpg:wgp",
                "word_com_error": str(com_error),
                "saved_group_evidence": defect,
            },
        )

    error = _lifecycle_observations(gold_shapes, pred_shapes, {"is_group"})
    if error is not None:
        return error
    gold_groups = [s for s in gold_shapes if s.get("is_group")]
    pred_groups = [s for s in pred_shapes if s.get("is_group")]

    if not gold_groups:
        result = VerificationResult.error_result(
            "Gold grouping evidence contains no observable group"
        )
        result.details.update(
            {
                "error_owner": "gold_fixture",
                "gold_groups": 0,
                "gold_top_level_shapes": len(gold_shapes),
                "reason": "empty_gold_grouping_evidence",
            }
        )
        return result

    if len(gold_groups) != len(pred_groups):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Group count mismatch (gold={len(gold_groups)}, pred={len(pred_groups)})",
            details={"gold_groups": len(gold_groups), "pred_groups": len(pred_groups)},
        )

    if len(gold_shapes) != len(pred_shapes):
        return VerificationResult(
            passed=False,
            message=(
                "FAIL: Top-level shape count mismatch after grouping "
                f"(gold={len(gold_shapes)}, pred={len(pred_shapes)})"
            ),
            details={
                "gold_top_level_shapes": len(gold_shapes),
                "pred_top_level_shapes": len(pred_shapes),
            },
        )

    gold_group_item_counts = [group.get("group_item_count") for group in gold_groups]
    pred_group_item_counts = [group.get("group_item_count") for group in pred_groups]
    if any(count is None for count in gold_group_item_counts):
        result = VerificationResult.error_result("Gold grouping evidence lacks group-item counts")
        result.details.update(
            {
                "error_owner": "gold_fixture",
                "gold_group_item_counts": gold_group_item_counts,
                "reason": "unobservable_gold_group_item_count",
            }
        )
        return result
    if any(count is None for count in pred_group_item_counts):
        return VerificationResult(
            passed=False,
            message="FAIL: Prediction grouping lacks group-item counts",
            details={
                "gold_group_item_counts": gold_group_item_counts,
                "pred_group_item_counts": pred_group_item_counts,
                "reason": "unobservable_prediction_group_item_count",
            },
        )

    gold_group_item_counts = sorted(int(count) for count in gold_group_item_counts)
    pred_group_item_counts = sorted(int(count) for count in pred_group_item_counts)
    if gold_group_item_counts != pred_group_item_counts:
        return VerificationResult(
            passed=False,
            message=(
                "FAIL: Group item-count mismatch "
                f"(gold={gold_group_item_counts}, "
                f"pred={pred_group_item_counts})"
            ),
            details={
                "gold_group_item_counts": gold_group_item_counts,
                "pred_group_item_counts": pred_group_item_counts,
            },
        )

    gold_effective_shapes = sum(
        int(shape.get("group_item_count")) if shape.get("is_group") else 1 for shape in gold_shapes
    )
    pred_effective_shapes = sum(
        int(shape.get("group_item_count")) if shape.get("is_group") else 1 for shape in pred_shapes
    )
    if gold_effective_shapes != pred_effective_shapes:
        return VerificationResult(
            passed=False,
            message=(
                "FAIL: Effective shape count mismatch after grouping "
                f"(gold={gold_effective_shapes}, "
                f"pred={pred_effective_shapes})"
            ),
            details={
                "gold_effective_shapes": gold_effective_shapes,
                "pred_effective_shapes": pred_effective_shapes,
            },
        )

    return VerificationResult(
        passed=True,
        message=f"PASS: Group structure matches ({len(pred_groups)} groups)",
        details={
            "group_count": len(pred_groups),
            "top_level_shape_count": len(pred_shapes),
            "effective_shape_count": pred_effective_shapes,
            "group_item_counts": pred_group_item_counts,
        },
    )


def verify_ungrouping(
    pred_path: Path, gold_path: Path, *, read_shape_properties
) -> VerificationResult:
    """Verify shapes were ungrouped correctly (no groups remain)."""
    gold_shapes = read_shape_properties(str(gold_path))
    pred_shapes = read_shape_properties(str(pred_path))

    error = _lifecycle_observations(gold_shapes, pred_shapes, {"is_group"})
    if error is not None:
        return error
    gold_groups = [s for s in gold_shapes if s.get("is_group")]
    pred_groups = [s for s in pred_shapes if s.get("is_group")]

    if len(gold_groups) != len(pred_groups):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Group count mismatch after ungroup (gold={len(gold_groups)}, pred={len(pred_groups)})",
            details={"gold_groups": len(gold_groups), "pred_groups": len(pred_groups)},
        )

    # After ungrouping, there should be more individual shapes
    if len(gold_shapes) != len(pred_shapes):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Shape count mismatch (gold={len(gold_shapes)}, pred={len(pred_shapes)})",
            details={"gold_shapes": len(gold_shapes), "pred_shapes": len(pred_shapes)},
        )

    return VerificationResult(
        passed=True,
        message=f"PASS: Ungroup verified ({len(pred_shapes)} shapes, {len(pred_groups)} groups)",
        details={"shape_count": len(pred_shapes), "group_count": len(pred_groups)},
    )


def verify_z_order(
    pred_path: Path, gold_path: Path, *, read_shape_properties
) -> VerificationResult:
    """Verify shape z-order matches."""
    gold_shapes = read_shape_properties(str(gold_path))
    pred_shapes = read_shape_properties(str(pred_path))

    error = _lifecycle_observations(gold_shapes, pred_shapes, {"name", "z_order"})
    if error is not None:
        return error
    # Sort by z-order and get name sequence
    gold_order = [s["name"] for s in sorted(gold_shapes, key=lambda x: x["z_order"])]
    pred_order = [s["name"] for s in sorted(pred_shapes, key=lambda x: x["z_order"])]

    if gold_order == pred_order:
        return VerificationResult(
            passed=True, message="PASS: Shape z-order matches", details={"z_order": gold_order}
        )
    else:
        return VerificationResult(
            passed=False,
            message="FAIL: Shape z-order differs",
            details={"gold_order": gold_order, "pred_order": pred_order},
        )


def verify_rotation(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, read_shape_properties
) -> VerificationResult:
    """
    Verify shape rotation and flip state matches.

    Handles ObjectRotateGallery variants:
    - rotate_right_90, rotate_left_90 (rotation angle)
    - flip_horizontal, flip_vertical (flip state)
    """
    gold_shapes = read_shape_properties(str(gold_path))
    pred_shapes = read_shape_properties(str(pred_path))

    if len(gold_shapes) != len(pred_shapes):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Shape count mismatch (gold={len(gold_shapes)}, pred={len(pred_shapes)})",
            details={},
        )

    if not gold_shapes:
        return VerificationResult.error_result("Gold shape target is missing", observation_owner="gold")

    # Determine what to check from parameter or variation name
    parameter = spec.get("parameter", {}) if spec else {}
    param_type = parameter.get("type")
    direction = parameter.get("direction")

    # If not specified in parameter, infer from variation name
    if not param_type:
        variation = spec.get("variation", "") if spec else ""
        if "flip_horizontal" in variation:
            param_type = "flip"
            direction = "horizontal"
        elif "flip_vertical" in variation:
            param_type = "flip"
            direction = "vertical"
        else:
            param_type = "rotate"

    # Sort shapes by name for consistent comparison
    required = "rotation" if param_type == "rotate" else f"flip_{direction}"
    for owner, shapes in (("gold", gold_shapes), ("pred", pred_shapes)):
        for shape in shapes:
            error = require_observations({owner: shape}, [required], context="shape rotation/flip")
            if error is not None:
                return error
    gold_sorted = sorted(gold_shapes, key=lambda s: s.get("name", ""))
    pred_sorted = sorted(pred_shapes, key=lambda s: s.get("name", ""))

    differences = []
    tolerance = 1.0  # 1 degree tolerance for rotation

    for gold_shape, pred_shape in zip(gold_sorted, pred_sorted):
        shape_diffs = []

        if param_type == "rotate":
            # Check rotation angle
            gold_rot = gold_shape.get("rotation", 0)
            pred_rot = pred_shape.get("rotation", 0)
            # Circular comparison: gold 0.0 vs pred 359.6 is a 0.4 degree
            # difference, not 359.6 — plain modulo subtraction wrongly
            # fails at the wrap-around.
            angle_delta = abs((float(pred_rot) - float(gold_rot) + 180.0) % 360.0 - 180.0)
            if angle_delta > tolerance:
                shape_diffs.append(
                    {"property": "rotation", "gold": gold_rot % 360, "pred": pred_rot % 360}
                )

        elif param_type == "flip":
            # Check flip state (direction already set above)
            if direction == "horizontal":
                gold_flip = gold_shape.get("flip_horizontal", False)
                pred_flip = pred_shape.get("flip_horizontal", False)
                if gold_flip != pred_flip:
                    shape_diffs.append(
                        {"property": "flip_horizontal", "gold": gold_flip, "pred": pred_flip}
                    )
            else:  # vertical
                gold_flip = gold_shape.get("flip_vertical", False)
                pred_flip = pred_shape.get("flip_vertical", False)
                if gold_flip != pred_flip:
                    shape_diffs.append(
                        {"property": "flip_vertical", "gold": gold_flip, "pred": pred_flip}
                    )

        if shape_diffs:
            differences.append({"shape_name": gold_shape.get("name"), "differences": shape_diffs})

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: {len(differences)} shape(s) have rotation/flip differences",
            details={"differences": differences, "param_type": param_type},
        )

    return VerificationResult(
        passed=True,
        message=f"PASS: Shape rotation/flip matches",
        details={"param_type": param_type, "shape_count": len(gold_shapes)},
    )


def verify_insert_shape_operation(
    pred_path: Path,
    gold_path: Path,
    params: Optional[Dict] = None,
    spec: Optional[Dict] = None,
    *,
    read_shape_placement,
    read_shape_properties,
) -> VerificationResult:
    """Verify prerequisite shape insertion geometry and type."""
    params = params or {}
    try:
        pred_shapes, gold_shapes, extraction_source, extraction_note = (
            families_shapes_layout.extract_shape_pair(
                pred_path,
                gold_path,
                prefer_saved=False,
                read_shape_properties=read_shape_properties,
            )
        )
    except DocumentOpenError:
        raise
    except Exception as exc:
        return VerificationResult.error_result(
            f"Shape extraction error: {exc}",
            observation_owner="unknown", reason="shape_extraction_failed",
        )
    if not gold_shapes:
        return VerificationResult.error_result(
            "Gold insertion target contains no shapes", observation_owner="gold",
            reason="missing_gold_shape_target",
        )
    pred_shapes, gold_shapes = families_shapes_geometry.match_shapes_by_name(
        pred_shapes, gold_shapes
    )

    if len(gold_shapes) != len(pred_shapes):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Shape count mismatch (gold={len(gold_shapes)}, pred={len(pred_shapes)})",
            details={"gold_count": len(gold_shapes), "pred_count": len(pred_shapes)},
        )
    if not gold_shapes:
        return VerificationResult.error_result("Gold shape target is missing", observation_owner="gold")

    gold_sorted = sorted(gold_shapes, key=lambda s: s.get("name", ""))
    pred_sorted = sorted(pred_shapes, key=lambda s: s.get("name", ""))
    differences = []
    page_width = gold_sorted[0].get("page_width", 612)
    page_height = gold_sorted[0].get("page_height", 792)
    left_margin = gold_sorted[0].get("left_margin", 72)
    top_margin = gold_sorted[0].get("top_margin", 72)
    right_margin = gold_sorted[0].get("right_margin", 72)
    bottom_margin = gold_sorted[0].get("bottom_margin", 72)

    for idx, (gold_shape, pred_shape) in enumerate(zip(gold_sorted, pred_sorted), start=1):
        shape_diffs = []
        if params.get("type") is not None and gold_shape.get("auto_shape_type") != pred_shape.get(
            "auto_shape_type"
        ):
            shape_diffs.append(
                {
                    "property": "auto_shape_type",
                    "gold": gold_shape.get("auto_shape_type"),
                    "pred": pred_shape.get("auto_shape_type"),
                }
            )
        for prop, size_prop, page_size, margin_start, margin_end, rel_prop in [
            ("left", "width", page_width, left_margin, right_margin, "relative_h_position"),
            ("top", "height", page_height, top_margin, bottom_margin, "relative_v_position"),
        ]:
            gold_raw = gold_shape.get(prop)
            pred_raw = pred_shape.get(prop)
            if gold_raw is None or pred_raw is None:
                shape_diffs.append(
                    {
                        "property": prop,
                        "gold": gold_raw,
                        "pred": pred_raw,
                    }
                )
                continue
            gold_val = families_shapes_geometry.calculate_actual_position(
                gold_raw,
                gold_shape.get(size_prop, 0),
                gold_shape.get(rel_prop, 0),
                page_size,
                margin_start,
                margin_end,
            )
            pred_val = families_shapes_geometry.calculate_actual_position(
                pred_raw,
                pred_shape.get(size_prop, 0),
                pred_shape.get(rel_prop, 0),
                page_size,
                margin_start,
                margin_end,
            )
            if abs(gold_val - pred_val) > 2.0:
                shape_diffs.append(
                    {
                        "property": prop,
                        "gold_actual": round(gold_val, 2),
                        "pred_actual": round(pred_val, 2),
                        "gold_raw": gold_raw,
                        "pred_raw": pred_raw,
                    }
                )
        for prop in ("width", "height"):
            if params.get(prop) is None:
                continue
            gold_val = gold_shape.get(prop)
            pred_val = pred_shape.get(prop)
            if gold_val is None or pred_val is None or abs(gold_val - pred_val) > 2.0:
                shape_diffs.append(
                    {
                        "property": prop,
                        "gold": gold_val,
                        "pred": pred_val,
                    }
                )
        if shape_diffs:
            differences.append({"shape_index": idx, "differences": shape_diffs})

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Inserted shape geometry mismatch ({len(differences)} shape(s))",
            details={"differences": differences},
        )
    heading_text = str(((spec or {}).get("precondition") or {}).get("heading") or "").strip()
    if heading_text:
        shape_type = params.get("type")
        gold_placement = read_shape_placement(str(gold_path), heading_text, shape_type)
        pred_placement = read_shape_placement(str(pred_path), heading_text, shape_type)
        if not gold_placement.get("found") or not gold_placement.get("below_heading"):
            return VerificationResult.error_result(
                "Gold prerequisite shape is not visibly below its heading"
            )
        if not pred_placement.get("found") or not pred_placement.get("below_heading"):
            return VerificationResult(
                passed=False,
                message="FAIL: Inserted shape is not visibly below the prerequisite heading",
                details={
                    "match_mode": "word_com_shape_heading_relation",
                    "heading": heading_text,
                    "gold_placement": gold_placement,
                    "pred_placement": pred_placement,
                },
            )
    return VerificationResult(
        passed=True,
        message="PASS: Inserted shape geometry matches",
        details={
            "shape_count": len(gold_shapes),
            **({"heading_relation": "below"} if heading_text else {}),
        },
    )
