"""families / shapes / effects with explicit evidence dependencies."""

from evaluation.methods.base import (
    VerificationResult,
    combine_com_primary_with_ooxml_guardrail,
    tag_verification_source,
)
from evaluation.methods.word_object.families.icons import saved_color as families_icons_saved_color
from evaluation.methods.word_object.families.shapes import (
    extract_saved as families_shapes_extract_saved,
)
from pathlib import Path
from ...observations import require_observations
from typing import Dict, Optional


def verify_transparency(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, read_shape_properties
) -> VerificationResult:
    """
    Verify shape/picture transparency matches between pred and gold.

    Handles PictureTransparencyGallery command.
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

    differences = []
    tolerance = 0.01  # 1% tolerance for transparency

    for gold_shape, pred_shape in zip(gold_shapes, pred_shapes):
        gold_transparency = gold_shape.get("fill_transparency", 0)
        pred_transparency = pred_shape.get("fill_transparency", 0)

        if abs(gold_transparency - pred_transparency) > tolerance:
            differences.append(
                {
                    "shape_name": gold_shape.get("name"),
                    "gold_transparency": gold_transparency,
                    "pred_transparency": pred_transparency,
                }
            )

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: {len(differences)} shape(s) have transparency differences",
            details={"differences": differences},
        )

    return VerificationResult(
        passed=True,
        message=f'PASS: Transparency matches ({gold_shapes[0].get("fill_transparency", 0):.0%})',
        details={
            "shape_count": len(gold_shapes),
            "transparency": gold_shapes[0].get("fill_transparency", 0),
        },
    )


def verify_shadow(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, read_shape_properties
) -> VerificationResult:
    """
    Verify shape shadow properties match between pred and gold.

    Handles ShadowNudge*, ShadowOnOrOff*, ShadowSemitransparent* commands.
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
    tolerance = 1.0  # 1 point tolerance for shadow offsets

    for gold_shape, pred_shape in zip(gold_sorted, pred_sorted):
        shape_diffs = []

        # Check shadow visibility
        gold_visible = gold_shape.get("shadow_visible")
        pred_visible = pred_shape.get("shadow_visible")
        if gold_visible != pred_visible:
            shape_diffs.append(
                {"property": "shadow_visible", "gold": gold_visible, "pred": pred_visible}
            )

        # Check shadow offsets (X and Y)
        for prop in ["shadow_offset_x", "shadow_offset_y"]:
            gold_val = gold_shape.get(prop)
            pred_val = pred_shape.get(prop)
            if gold_val is not None and pred_val is not None:
                if abs(gold_val - pred_val) > tolerance:
                    shape_diffs.append(
                        {
                            "property": prop,
                            "gold": gold_val,
                            "pred": pred_val,
                            "diff": round(abs(gold_val - pred_val), 2),
                        }
                    )

        if shape_diffs:
            differences.append({"shape_name": gold_shape.get("name"), "differences": shape_diffs})

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: {len(differences)} shape(s) have shadow differences",
            details={
                "differences": differences,
                "gold_shapes": [
                    {k: v for k, v in s.items() if "shadow" in k or k == "name"}
                    for s in gold_sorted
                ],
                "pred_shapes": [
                    {k: v for k, v in s.items() if "shadow" in k or k == "name"}
                    for s in pred_sorted
                ],
            },
        )

    return VerificationResult(
        passed=True,
        message="PASS: Shadow properties match",
        details={
            "shape_count": len(gold_shapes),
            "shadow_visible": gold_shapes[0].get("shadow_visible") if gold_shapes else None,
        },
    )


def verify_threed(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, read_shape_properties
) -> VerificationResult:
    """
    Verify 3D effect properties match between pred and gold.

    Handles _3DEffectsOnOffClassic, _3DTilt*, _3DExtrusion* commands.
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

    def _numeric_mismatch(gold_value, pred_value, tolerance=1.0):
        # 3D rotation/depth round-trip through DrawingML 1/60000-degree
        # units, so an XML-authored 9.999983 vs COM's 10.0 must not fail
        # on exact equality. Matches the 1.0-degree tolerance of the
        # _3DRotationGallery geometry branch. NOT for tristate props
        # (visible/perspective): |msoTrue(-1) - msoFalse(0)| == 1.0 would
        # compare as equal under this tolerance.
        if gold_value is None or pred_value is None:
            return gold_value != pred_value
        try:
            return abs(float(gold_value) - float(pred_value)) > tolerance
        except (TypeError, ValueError):
            return gold_value != pred_value

    for gold_shape, pred_shape in zip(gold_sorted, pred_sorted):
        # Check 3D visibility
        gold_vis = gold_shape.get("threed_visible")
        pred_vis = pred_shape.get("threed_visible")
        if gold_vis != pred_vis:
            return VerificationResult(
                passed=False,
                message=f"FAIL: 3D visibility mismatch (gold={gold_vis}, pred={pred_vis})",
                details={"shape_name": gold_shape.get("name")},
            )

        # Check rotation X
        gold_rx = gold_shape.get("threed_rotation_x")
        pred_rx = pred_shape.get("threed_rotation_x")
        if _numeric_mismatch(gold_rx, pred_rx):
            return VerificationResult(
                passed=False,
                message=f"FAIL: 3D RotationX mismatch (gold={gold_rx}, pred={pred_rx})",
                details={"shape_name": gold_shape.get("name")},
            )

        # Check rotation Y
        gold_ry = gold_shape.get("threed_rotation_y")
        pred_ry = pred_shape.get("threed_rotation_y")
        if _numeric_mismatch(gold_ry, pred_ry):
            return VerificationResult(
                passed=False,
                message=f"FAIL: 3D RotationY mismatch (gold={gold_ry}, pred={pred_ry})",
                details={"shape_name": gold_shape.get("name")},
            )

        # Check depth
        gold_depth = gold_shape.get("threed_depth")
        pred_depth = pred_shape.get("threed_depth")
        if _numeric_mismatch(gold_depth, pred_depth):
            return VerificationResult(
                passed=False,
                message=f"FAIL: 3D Depth mismatch (gold={gold_depth}, pred={pred_depth})",
                details={"shape_name": gold_shape.get("name")},
            )

        # Check perspective — MsoTriState (perspective vs parallel
        # projection), a categorical toggle: exact comparison only.
        gold_persp = gold_shape.get("threed_perspective")
        pred_persp = pred_shape.get("threed_perspective")
        if gold_persp != pred_persp:
            return VerificationResult(
                passed=False,
                message=f"FAIL: 3D Perspective mismatch (gold={gold_persp}, pred={pred_persp})",
                details={"shape_name": gold_shape.get("name")},
            )

    return VerificationResult(
        passed=True, message="PASS: 3D properties match", details={"shape_count": len(gold_shapes)}
    )


def verify_fill_outline_color(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    read_shape_properties,
    saved_evidence_error,
) -> VerificationResult:
    """
    Verify shape fill and outline color properties match.

    Handles ShapeFillColorPicker, ShapeOutlineColorPicker, OutlineColorPicker,
    GlowColorPicker, GraphicsColorPicker, GraphicsOutlineColorPicker.
    """
    idMso = spec.get("idMso", "") if spec else ""
    if idMso == "GraphicsOutlineColorPicker":
        color_rgb = (spec or {}).get("parameter", {}).get("color_rgb")
        if color_rgb is None:
            color_rgb = {"red": 0xFF0000, "blue": 0x0000FF}.get((spec or {}).get("variation"))
        if color_rgb is not None:
            color_hex = f"{int(color_rgb) & 0xFFFFFF:06X}"
            gold_count = families_icons_saved_color.saved_svg_stroke_color_count(
                gold_path, color_hex
            )
            pred_count = families_icons_saved_color.saved_svg_stroke_color_count(
                pred_path, color_hex
            )
            if not gold_count:
                # Without gold evidence 0 == 0 would pass every
                # prediction vacuously; the reference itself must carry
                # the instructed stroke color.
                return saved_evidence_error(
                    "Gold contains no SVG stroke with the instructed " f"outline color {color_hex}",
                    evidence_owner="gold",
                )
            if gold_count != pred_count:
                return VerificationResult(
                    passed=False,
                    message="FAIL: Saved SVG outline color signature differs",
                    details={
                        "match_mode": "saved_svg_stroke_color",
                        "color": color_hex,
                        "gold_stroke_count": gold_count,
                        "pred_stroke_count": pred_count,
                        "verification_source": "saved_ooxml_only",
                    },
                )
            return VerificationResult(
                passed=True,
                message="PASS: Saved SVG outline color signature matches",
                details={
                    "match_mode": "saved_svg_stroke_color",
                    "color": color_hex,
                    "stroke_count": gold_count,
                    "verification_source": "saved_ooxml_only",
                },
            )

    saved_guardrail = None
    if idMso == "GlowColorPicker":
        gold_glows = families_shapes_extract_saved.extract_saved_glow_signatures(gold_path)
        pred_glows = families_shapes_extract_saved.extract_saved_glow_signatures(pred_path)
        if gold_glows is not None and pred_glows is not None:
            if gold_glows != pred_glows:
                saved_guardrail = VerificationResult(
                    passed=False,
                    message="FAIL: Saved shape glow signature differs",
                    details={
                        "match_mode": "saved_ooxml_shape_glow",
                        "gold": gold_glows,
                        "pred": pred_glows,
                    },
                )
            else:
                saved_guardrail = VerificationResult(
                    passed=True,
                    message="PASS: Saved shape glow signature matches",
                    details={
                        "match_mode": "saved_ooxml_shape_glow",
                        "signature": gold_glows,
                    },
                )

    gold_shapes = read_shape_properties(str(gold_path))
    pred_shapes = read_shape_properties(str(pred_path))

    if len(gold_shapes) != len(pred_shapes):
        return tag_verification_source(
            VerificationResult(
                passed=False,
                message=f"FAIL: Shape count mismatch (gold={len(gold_shapes)}, pred={len(pred_shapes)})",
                details={},
            ),
            "word_com",
        )

    if not gold_shapes:
        return tag_verification_source(
            VerificationResult(passed=False, message="FAIL: No shapes found", details={}),
            "word_com",
        )

    gold_sorted = sorted(gold_shapes, key=lambda s: s.get("name", ""))
    pred_sorted = sorted(pred_shapes, key=lambda s: s.get("name", ""))

    differences = []
    idMso = spec.get("idMso", "") if spec else ""

    # Determine which properties to compare based on command
    props_to_check = []
    if "Fill" in idMso or idMso == "GraphicsColorPicker":
        props_to_check.extend([("fill_fore_color_rgb", "exact"), ("fill_type", "exact")])
    if "Outline" in idMso:
        props_to_check.extend([("line_fore_color_rgb", "exact"), ("line_visible", "exact")])
    if "Glow" in idMso:
        props_to_check.extend([("glow_color_rgb", "exact"), ("glow_radius", 1.0)])

    # Fallback: check all color-related props
    has_explicit_target = bool(props_to_check)
    if not props_to_check:
        props_to_check = [
            ("fill_fore_color_rgb", "exact"),
            ("line_fore_color_rgb", "exact"),
            ("glow_color_rgb", "exact"),
        ]

    for gold_shape, pred_shape in zip(gold_sorted, pred_sorted):
        shape_diffs = []
        if has_explicit_target:
            error = require_observations(
                {"gold": gold_shape, "pred": pred_shape},
                [prop for prop, _ in props_to_check], context=idMso,
            )
            if error is not None:
                return error
        for prop, tol in props_to_check:
            gold_val = gold_shape.get(prop)
            pred_val = pred_shape.get(prop)
            if gold_val is None and pred_val is None:
                continue
            if tol == "exact":
                if gold_val != pred_val:
                    shape_diffs.append(
                        {
                            "property": prop,
                            "gold": gold_val,
                            "pred": pred_val,
                        }
                    )
            else:
                if gold_val is not None and pred_val is not None:
                    if abs(gold_val - pred_val) > tol:
                        shape_diffs.append(
                            {
                                "property": prop,
                                "gold": gold_val,
                                "pred": pred_val,
                            }
                        )
        if shape_diffs:
            differences.append(
                {
                    "shape_name": gold_shape.get("name"),
                    "differences": shape_diffs,
                }
            )

    if differences:
        return tag_verification_source(
            VerificationResult(
                passed=False,
                message=f"FAIL: {len(differences)} shape(s) have fill/outline/color differences",
                details={"differences": differences, "idMso": idMso},
            ),
            "word_com",
        )

    com_result = VerificationResult(
        passed=True,
        message="PASS: Fill/outline/color properties match",
        details={"shape_count": len(gold_shapes), "idMso": idMso},
    )
    if saved_guardrail is not None:
        return combine_com_primary_with_ooxml_guardrail(com_result, saved_guardrail)
    return tag_verification_source(com_result, "word_com")
