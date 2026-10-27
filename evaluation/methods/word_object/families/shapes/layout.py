"""families / shapes / layout with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.core import DocumentOpenError
from evaluation.methods.word_object.families.shapes import (
    extract_saved as families_shapes_extract_saved,
)
from evaluation.methods.word_object.families.shapes import geometry as families_shapes_geometry
from pathlib import Path
from ...observations import require_observations
from typing import Any, Dict, List, Optional


def extract_shape_pair(
    pred_path: Path, gold_path: Path, *, prefer_saved: bool = False, read_shape_properties
) -> tuple[List[Dict[str, Any]], List[Dict[str, Any]], str, Optional[str]]:
    """Extract both sides through one coordinate/property representation.

    Simple layout commands prefer saved OOXML because it fully captures
    their persisted geometry and avoids known Word-open hangs.  Other
    callers retain COM as primary and use OOXML only after a real error.
    """
    if prefer_saved:
        try:
            gold_shapes = families_shapes_extract_saved.extract_saved_shape_properties(gold_path)
            pred_shapes = families_shapes_extract_saved.extract_saved_shape_properties(pred_path)
            return pred_shapes, gold_shapes, "saved_ooxml", None
        except Exception as saved_error:
            gold_shapes = read_shape_properties(str(gold_path))
            pred_shapes = read_shape_properties(str(pred_path))
            return pred_shapes, gold_shapes, "word_com", str(saved_error)
    try:
        gold_shapes = read_shape_properties(str(gold_path))
        pred_shapes = read_shape_properties(str(pred_path))
        return pred_shapes, gold_shapes, "word_com", None
    except DocumentOpenError:
        raise
    except Exception as com_error:
        gold_shapes = families_shapes_extract_saved.extract_saved_shape_properties(gold_path)
        pred_shapes = families_shapes_extract_saved.extract_saved_shape_properties(pred_path)
        return pred_shapes, gold_shapes, "saved_ooxml_fallback", str(com_error)


def verify_alignment(
    pred_path: Path, gold_path: Path, alignment_type: str, *, read_shape_properties
) -> VerificationResult:
    """
    Verify shapes are aligned correctly.

    Args:
        alignment_type: 'left', 'right', 'top', 'bottom', 'center_h', 'center_v'
    """
    try:
        pred_shapes, gold_shapes, extraction_source, extraction_note = extract_shape_pair(
            pred_path, gold_path, prefer_saved=False, read_shape_properties=read_shape_properties
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
            "Gold alignment target contains no shapes", observation_owner="gold",
            reason="missing_gold_shape_target",
        )
    pred_shapes, gold_shapes = families_shapes_geometry.match_shapes_by_name(
        pred_shapes, gold_shapes
    )

    if len(gold_shapes) != len(pred_shapes):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Shape count mismatch (gold={len(gold_shapes)}, pred={len(pred_shapes)})",
            details={},
        )

    if not pred_shapes:
        return VerificationResult(passed=False, message="FAIL: No shapes found", details={})

    if alignment_type not in {
        "left",
        "right",
        "top",
        "bottom",
        "center_h",
        "center_v",
    }:
        return VerificationResult(
            passed=False, message=f"ERROR: Unknown alignment type: {alignment_type}",
            error=f"Unknown alignment type: {alignment_type}", details={}
        )

    try:
        pred_values = [
            families_shapes_geometry.normalized_alignment_value(shape, alignment_type)
            for shape in pred_shapes
        ]
        gold_values = [
            families_shapes_geometry.normalized_alignment_value(shape, alignment_type)
            for shape in gold_shapes
        ]
    except (KeyError, TypeError, ValueError) as exc:
        result = VerificationResult.error_result(f"Alignment geometry normalization failed: {exc}")
        result.details.update(
            {
                "alignment_type": alignment_type,
                "verification_source": extraction_source,
                "extraction_note": extraction_note,
            }
        )
        return result

    tolerance = 2.0

    # Check pred shapes are aligned (all same value within tolerance)
    pred_variance = max(pred_values) - min(pred_values)

    if pred_variance > tolerance:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Shapes not aligned ({alignment_type}), variance={pred_variance:.2f}pt",
            details={
                "alignment_type": alignment_type,
                "pred_values": pred_values,
                "pred_variance": pred_variance,
                "verification_source": (extraction_source),
                "extraction_note": extraction_note,
            },
        )

    pred_mean = sum(pred_values) / len(pred_values)

    preservation_failures = families_shapes_geometry.relative_layout_preservation_failures(
        pred_shapes, gold_shapes, tolerance
    )
    try:
        preservation_failures.extend(
            families_shapes_geometry.orthogonal_alignment_preservation_failures(
                pred_shapes,
                gold_shapes,
                alignment_type,
                tolerance,
            )
        )
    except (KeyError, TypeError, ValueError) as exc:
        result = VerificationResult.error_result(
            f"Alignment preservation normalization failed: {exc}"
        )
        result.details.update(
            {
                "alignment_type": alignment_type,
                "verification_source": extraction_source,
                "extraction_note": extraction_note,
            }
        )
        return result
    if preservation_failures:
        return VerificationResult(
            passed=False,
            message="FAIL: Alignment changed preserved shape state",
            details={
                "alignment_type": alignment_type,
                "failures": preservation_failures,
                "reason": "layout_preservation_mismatch",
                "verification_source": (extraction_source),
                "extraction_note": extraction_note,
            },
        )

    return VerificationResult(
        passed=True,
        message=f"PASS: Shapes aligned ({alignment_type}), variance={pred_variance:.2f}pt",
        details={
            "alignment_type": alignment_type,
            "alignment_position": round(pred_mean, 2),
            "relative_only_layout": True,
            "gold_values_observed_not_anchored": gold_values,
            "pred_variance": pred_variance,
            "verification_source": (extraction_source),
            "extraction_note": extraction_note,
        },
    )


def verify_distribution(
    pred_path: Path, gold_path: Path, direction: str, *, read_shape_properties
) -> VerificationResult:
    """
    Verify shapes are evenly distributed.

    Args:
        direction: 'horizontal' or 'vertical'
    """
    try:
        pred_shapes, gold_shapes, extraction_source, extraction_note = extract_shape_pair(
            pred_path, gold_path, prefer_saved=False, read_shape_properties=read_shape_properties
        )
    except DocumentOpenError:
        raise
    except Exception as exc:
        return VerificationResult.error_result(
            f"Shape extraction error: {exc}",
            observation_owner="unknown", reason="shape_extraction_failed",
        )
    if len(gold_shapes) < 3:
        return VerificationResult.error_result(
            "Gold distribution target contains fewer than three shapes",
            observation_owner="gold", reason="invalid_gold_shape_count",
        )
    pred_shapes, gold_shapes = families_shapes_geometry.match_shapes_by_name(
        pred_shapes, gold_shapes
    )

    if len(pred_shapes) != len(gold_shapes):
        return VerificationResult(
            passed=False,
            message=(
                "FAIL: Shape count mismatch for distribution "
                f"(gold={len(gold_shapes)}, pred={len(pred_shapes)})"
            ),
            details={
                "direction": direction,
                "gold_count": len(gold_shapes),
                "pred_count": len(pred_shapes),
                "verification_source": extraction_source,
                "extraction_note": extraction_note,
            },
        )

    if len(pred_shapes) < 3:
        return VerificationResult(
            passed=False, message="FAIL: Need at least 3 shapes for distribution check", details={}
        )

    # Get positions based on direction
    if direction == "horizontal":
        pred_positions = sorted([s["center_h"] for s in pred_shapes])
        gold_positions = sorted([s["center_h"] for s in gold_shapes])
    else:
        pred_positions = sorted([s["center_v"] for s in pred_shapes])
        gold_positions = sorted([s["center_v"] for s in gold_shapes])

    # Calculate gaps between consecutive shapes
    pred_gaps = [pred_positions[i + 1] - pred_positions[i] for i in range(len(pred_positions) - 1)]

    tolerance = 2.0

    # Check if gaps are equal (within tolerance)
    gap_variance = max(pred_gaps) - min(pred_gaps)

    if gap_variance > tolerance:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Shapes not evenly distributed ({direction}), gap variance={gap_variance:.2f}pt",
            details={
                "direction": direction,
                "pred_positions": pred_positions,
                "pred_gaps": pred_gaps,
                "gap_variance": gap_variance,
                "verification_source": (extraction_source),
                "extraction_note": extraction_note,
            },
        )

    preservation_failures = families_shapes_geometry.relative_layout_preservation_failures(
        pred_shapes, gold_shapes, tolerance
    )
    if preservation_failures:
        return VerificationResult(
            passed=False,
            message="FAIL: Distribution changed preserved shape state",
            details={
                "direction": direction,
                "failures": preservation_failures,
                "verification_source": (extraction_source),
                "extraction_note": extraction_note,
            },
        )

    return VerificationResult(
        passed=True,
        message=f"PASS: Shapes evenly distributed ({direction})",
        details={
            "direction": direction,
            "pred_gaps": pred_gaps,
            "gap_variance": gap_variance,
            "verification_source": (extraction_source),
            "extraction_note": extraction_note,
        },
    )


def verify_position(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    get_position_tolerance,
    read_shape_properties,
) -> VerificationResult:
    """
    Verify shape position by comparing actual rendered positions.

    Uses Word COM API to get shape properties and page dimensions,
    then calculates the actual rendered position accounting for
    preset constants and relative positioning anchors.

    This approach follows the principle from docs/design/why-word-object-verification.md:
    compare actual rendered positions, not XML attributes.
    """
    # Extract shape properties and page info via COM API
    gold_shapes = read_shape_properties(str(gold_path))
    pred_shapes = read_shape_properties(str(pred_path))

    if len(gold_shapes) != len(pred_shapes):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Shape count mismatch (gold={len(gold_shapes)}, pred={len(pred_shapes)})",
            details={},
        )

    if not gold_shapes:
        return VerificationResult(passed=False, message="FAIL: No shapes found", details={})

    required = {"left", "top", "width", "height", "relative_h_position", "relative_v_position"}
    for owner, shapes in (("gold", gold_shapes), ("pred", pred_shapes)):
        for shape in shapes:
            error = require_observations({owner: shape}, required, context="shape position")
            if error is not None:
                return error
    # Get page dimensions (use defaults if not available)
    # Standard Letter: 612x792pt, 1" margins = 72pt
    page_width = gold_shapes[0].get("page_width", 612)
    page_height = gold_shapes[0].get("page_height", 792)
    left_margin = gold_shapes[0].get("left_margin", 72)
    top_margin = gold_shapes[0].get("top_margin", 72)
    right_margin = gold_shapes[0].get("right_margin", 72)
    bottom_margin = gold_shapes[0].get("bottom_margin", 72)

    # Sort shapes by name for consistent comparison
    gold_sorted = sorted(gold_shapes, key=lambda s: s.get("name", ""))
    pred_sorted = sorted(pred_shapes, key=lambda s: s.get("name", ""))

    differences = []
    tolerance = get_position_tolerance()

    for gold_shape, pred_shape in zip(gold_sorted, pred_sorted):
        # Calculate actual rendered positions for gold
        gold_actual_left = families_shapes_geometry.calculate_actual_position(
            gold_shape.get("left", 0),
            gold_shape.get("width", 0),
            gold_shape.get("relative_h_position", 0),
            page_width,
            left_margin,
            right_margin,
        )
        gold_actual_top = families_shapes_geometry.calculate_actual_position(
            gold_shape.get("top", 0),
            gold_shape.get("height", 0),
            gold_shape.get("relative_v_position", 0),
            page_height,
            top_margin,
            bottom_margin,
        )

        # Calculate actual rendered positions for pred
        pred_actual_left = families_shapes_geometry.calculate_actual_position(
            pred_shape.get("left", 0),
            pred_shape.get("width", 0),
            pred_shape.get("relative_h_position", 0),
            page_width,
            left_margin,
            right_margin,
        )
        pred_actual_top = families_shapes_geometry.calculate_actual_position(
            pred_shape.get("top", 0),
            pred_shape.get("height", 0),
            pred_shape.get("relative_v_position", 0),
            page_height,
            top_margin,
            bottom_margin,
        )

        # Compare actual rendered positions
        shape_diffs = []

        if abs(gold_actual_left - pred_actual_left) > tolerance:
            shape_diffs.append(
                {
                    "property": "left",
                    "gold_actual": round(gold_actual_left, 1),
                    "pred_actual": round(pred_actual_left, 1),
                    "diff": round(abs(gold_actual_left - pred_actual_left), 1),
                }
            )

        if abs(gold_actual_top - pred_actual_top) > tolerance:
            shape_diffs.append(
                {
                    "property": "top",
                    "gold_actual": round(gold_actual_top, 1),
                    "pred_actual": round(pred_actual_top, 1),
                    "diff": round(abs(gold_actual_top - pred_actual_top), 1),
                }
            )

        if shape_diffs:
            differences.append({"shape_name": gold_shape.get("name"), "differences": shape_diffs})

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: {len(differences)} shape(s) have position differences",
            details={"differences": differences, "tolerance": tolerance},
        )

    return VerificationResult(
        passed=True,
        message="PASS: Shape positions match",
        details={"shape_count": len(gold_shapes), "tolerance": tolerance},
    )


def verify_wrap(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    read_shape_properties,
    uses_tight_wrap_polygon_projection,
) -> VerificationResult:
    """
    Verify text wrap type matches.

    Handles TextWrapGallery variants:
    - square (wrap_type=0)
    - tight (wrap_type=1)
    - top_bottom (wrap_type=4)
    - behind (wrap_type=5)
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
        return VerificationResult(passed=False, message="FAIL: No shapes found", details={})

    # Sort shapes by name for consistent comparison
    gold_sorted = sorted(gold_shapes, key=lambda s: s.get("name", ""))
    pred_sorted = sorted(pred_shapes, key=lambda s: s.get("name", ""))

    differences = []

    for gold_shape, pred_shape in zip(gold_sorted, pred_sorted):
        gold_wrap = gold_shape.get("wrap_type")
        pred_wrap = pred_shape.get("wrap_type")

        error = require_observations(
            {"gold": gold_shape, "pred": pred_shape},
            ["wrap_type", "left", "top", "width", "height", "relative_h_position", "relative_v_position"], context="text wrap",
        )
        if error is not None:
            return error

        if gold_wrap != pred_wrap:
            differences.append(
                {
                    "shape_name": gold_shape.get("name"),
                    "gold_wrap_type": gold_wrap,
                    "pred_wrap_type": pred_wrap,
                }
            )
        for key in ("name", "auto_shape_type"):
            if (
                gold_shape.get(key) is not None or pred_shape.get(key) is not None
            ) and gold_shape.get(key) != pred_shape.get(key):
                differences.append(
                    {
                        "shape_name": gold_shape.get("name"),
                        "property": key,
                        "gold": gold_shape.get(key),
                        "pred": pred_shape.get(key),
                    }
                )
        for key in ("width", "height"):
            gold_value = gold_shape.get(key)
            pred_value = pred_shape.get(key)
            if gold_value is None or pred_value is None or abs(gold_value - pred_value) > 2.0:
                differences.append(
                    {
                        "shape_name": gold_shape.get("name"),
                        "property": key,
                        "gold": gold_value,
                        "pred": pred_value,
                    }
                )
        for raw_key, size_key, page_key, start_margin, end_margin, rel_key in (
            ("left", "width", "page_width", "left_margin", "right_margin", "relative_h_position"),
            ("top", "height", "page_height", "top_margin", "bottom_margin", "relative_v_position"),
        ):
            gold_raw = gold_shape.get(raw_key)
            pred_raw = pred_shape.get(raw_key)
            if gold_raw is None or pred_raw is None:
                continue
            page_size = gold_shape.get(page_key, 612 if raw_key == "left" else 792)
            margin_start = gold_shape.get(start_margin, 72)
            margin_end = gold_shape.get(end_margin, 72)
            gold_actual = families_shapes_geometry.calculate_actual_position(
                gold_raw,
                gold_shape.get(size_key, 0),
                gold_shape.get(rel_key, 0),
                page_size,
                margin_start,
                margin_end,
            )
            pred_actual = families_shapes_geometry.calculate_actual_position(
                pred_raw,
                pred_shape.get(size_key, 0),
                pred_shape.get(rel_key, 0),
                page_size,
                margin_start,
                margin_end,
            )
            if abs(gold_actual - pred_actual) > 2.0:
                differences.append(
                    {
                        "shape_name": gold_shape.get("name"),
                        "property": raw_key,
                        "gold": round(gold_actual, 2),
                        "pred": round(pred_actual, 2),
                    }
                )

    if uses_tight_wrap_polygon_projection(spec):
        try:
            polygon_failures = families_shapes_geometry.wrap_polygon_failures(pred_path, gold_path)
        except Exception as exc:
            polygon_failures = [f"wrap polygon extraction error: {exc}"]
        if polygon_failures:
            differences.append(
                {
                    "property": "wrap_polygon",
                    "failures": polygon_failures,
                }
            )

    if differences:
        # Get expected wrap type from spec for better error message
        parameter = spec.get("parameter", {}) if spec else {}
        expected_wrap = parameter.get("wrap_type")

        return VerificationResult(
            passed=False,
            message=f"FAIL: {len(differences)} shape(s) have wrap type differences",
            details={"differences": differences, "expected_wrap_type": expected_wrap},
        )

    return VerificationResult(
        passed=True,
        message=f"PASS: Text wrap type matches",
        details={
            "shape_count": len(gold_shapes),
            "wrap_type": gold_shapes[0].get("wrap_type") if gold_shapes else None,
        },
    )
