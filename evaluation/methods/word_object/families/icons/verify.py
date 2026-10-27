"""families / icons / verify with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from pathlib import Path
from typing import Any, Dict


def extract_icon_info(docx_path: Path, *, read_svg_icon_properties) -> Dict[str, Any]:
    return read_svg_icon_properties(str(docx_path))


def verify_insert_icon_operation(
    pred_path: Path, gold_path: Path, *, read_svg_icon_properties
) -> VerificationResult:
    """Verify prerequisite SVG icon insertion shape count, size, and fill color."""
    gold_info = extract_icon_info(gold_path, read_svg_icon_properties=read_svg_icon_properties)
    pred_info = extract_icon_info(pred_path, read_svg_icon_properties=read_svg_icon_properties)
    failures = []

    if not pred_info.get("has_icon"):
        return VerificationResult(
            passed=False,
            message="FAIL: No embedded SVG icon found in predicted document",
            details={"gold": gold_info, "pred": pred_info},
        )

    if pred_info.get("count") != gold_info.get("count"):
        failures.append(
            {
                "property": "count",
                "gold": gold_info.get("count"),
                "pred": pred_info.get("count"),
            }
        )
    for prop in ("width", "height"):
        gold_val = gold_info.get(prop)
        pred_val = pred_info.get(prop)
        if gold_val is None or pred_val is None or abs(gold_val - pred_val) > 1.0:
            failures.append({"property": prop, "gold": gold_val, "pred": pred_val})
    if gold_info.get("fill_color_rgb") is not None and pred_info.get("fill_color_rgb") is not None:
        if gold_info.get("fill_color_rgb") != pred_info.get("fill_color_rgb"):
            failures.append(
                {
                    "property": "fill_color_rgb",
                    "gold": gold_info.get("fill_color_rgb"),
                    "pred": pred_info.get("fill_color_rgb"),
                }
            )

    gold_svg_shapes = [s.get("shape") for s in gold_info.get("svg_signatures", [])]
    pred_svg_shapes = [s.get("shape") for s in pred_info.get("svg_signatures", [])]
    if gold_svg_shapes != pred_svg_shapes:
        failures.append(
            {
                "property": "svg_shape",
                "gold": gold_svg_shapes,
                "pred": pred_svg_shapes,
            }
        )

    if failures:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Inserted icon mismatch ({len(failures)} difference(s))",
            details={"differences": failures, "gold": gold_info, "pred": pred_info},
        )
    return VerificationResult(
        passed=True,
        message="PASS: Inserted icon matches including SVG shape",
        details={"gold": gold_info, "pred": pred_info},
    )
