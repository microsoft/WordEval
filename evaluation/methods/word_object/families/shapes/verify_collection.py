"""families / shapes / verify_collection with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from pathlib import Path
from typing import Dict, List, Optional


def verify_shapes_match(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    get_position_tolerance,
    read_document_structure,
    read_shape_properties,
) -> VerificationResult:
    """
    Verify all shape properties match between pred and gold.

    Used as fallback for openxml_comparison and when no specific
    verification method is specified.
    """
    gold_shapes = read_shape_properties(str(gold_path))
    pred_shapes = read_shape_properties(str(pred_path))

    if len(gold_shapes) != len(pred_shapes):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Shape count mismatch (gold={len(gold_shapes)}, pred={len(pred_shapes)})",
            details={"gold_count": len(gold_shapes), "pred_count": len(pred_shapes)},
        )

    if not gold_shapes:
        # No shapes - compare document structure instead
        return verify_document_structure(
            pred_path, gold_path, read_document_structure=read_document_structure
        )

    tolerance = get_position_tolerance()

    # Sort by name for consistent comparison
    gold_sorted = sorted(gold_shapes, key=lambda s: s.get("name", ""))
    pred_sorted = sorted(pred_shapes, key=lambda s: s.get("name", ""))

    differences = []

    for i, (gold_shape, pred_shape) in enumerate(zip(gold_sorted, pred_sorted)):
        shape_diffs = []

        # Compare position/size properties
        for key in ["left", "top", "width", "height", "rotation"]:
            gold_val = gold_shape.get(key)
            pred_val = pred_shape.get(key)
            if gold_val is not None and pred_val is not None:
                if abs(gold_val - pred_val) > tolerance:
                    shape_diffs.append(
                        {
                            "property": key,
                            "gold": gold_val,
                            "pred": pred_val,
                            "diff": round(abs(gold_val - pred_val), 2),
                        }
                    )

        # Compare boolean/enum properties
        for key in ["flip_horizontal", "flip_vertical", "wrap_type", "is_group"]:
            gold_val = gold_shape.get(key)
            pred_val = pred_shape.get(key)
            if gold_val is not None and pred_val is not None:
                if gold_val != pred_val:
                    shape_diffs.append({"property": key, "gold": gold_val, "pred": pred_val})

        if shape_diffs:
            differences.append(
                {"shape_index": i, "shape_name": gold_shape.get("name"), "differences": shape_diffs}
            )

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: {len(differences)} shape(s) have property differences",
            details={
                "differences": differences,
                "gold_shapes": gold_shapes,
                "pred_shapes": pred_shapes,
            },
        )

    return VerificationResult(
        passed=True,
        message=f"PASS: All {len(gold_shapes)} shape(s) match",
        details={"shape_count": len(gold_shapes)},
    )


def verify_document_structure(
    pred_path: Path, gold_path: Path, *, read_document_structure
) -> VerificationResult:
    """Fallback: verify by comparing document structure counts."""
    gold_struct = read_document_structure(str(gold_path))
    pred_struct = read_document_structure(str(pred_path))

    if "error" in gold_struct:
        return VerificationResult.error_result(f"Gold: {gold_struct['error']}")
    if "error" in pred_struct:
        return VerificationResult.error_result(f"Pred: {pred_struct['error']}")

    differences = []
    for key in gold_struct.keys():
        gold_val = gold_struct.get(key)
        pred_val = pred_struct.get(key)
        if gold_val != pred_val:
            differences.append({"property": key, "gold": gold_val, "pred": pred_val})

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: {len(differences)} structure difference(s)",
            details={"differences": differences, "gold": gold_struct, "pred": pred_struct},
        )

    return VerificationResult(
        passed=True, message="PASS: Document structure matches", details={"structure": gold_struct}
    )


def compare_shape_positions(
    gold_shapes: List[Dict], pred_shapes: List[Dict], tolerance: float
) -> Optional[VerificationResult]:
    """
    Compare individual shape positions.

    Returns None if all match, or a failure result if differences found.
    """
    gold_sorted = sorted(gold_shapes, key=lambda s: s.get("name", ""))
    pred_sorted = sorted(pred_shapes, key=lambda s: s.get("name", ""))

    shape_differences = []

    for i, (gold_shape, pred_shape) in enumerate(zip(gold_sorted, pred_sorted)):
        shape_diffs = []

        for key in ["left", "top", "width", "height"]:
            gold_val = gold_shape.get(key)
            pred_val = pred_shape.get(key)
            if gold_val is not None and pred_val is not None:
                if abs(gold_val - pred_val) > tolerance:
                    shape_diffs.append(
                        {
                            "property": key,
                            "gold": gold_val,
                            "pred": pred_val,
                            "diff": round(abs(gold_val - pred_val), 2),
                        }
                    )

        if shape_diffs:
            shape_differences.append(
                {"shape_index": i, "shape_name": gold_shape.get("name"), "differences": shape_diffs}
            )

    if shape_differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: {len(shape_differences)} shape(s) have position/size differences",
            details={"shape_differences": shape_differences},
        )

    return None  # All shapes match
