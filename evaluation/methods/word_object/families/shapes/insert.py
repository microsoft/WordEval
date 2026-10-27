"""families / shapes / insert with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.families.shapes import (
    insert_geometry as families_shapes_insert_geometry,
)
from evaluation.methods.word_object.families.shapes import (
    insert_saved as families_shapes_insert_saved,
)
from pathlib import Path
from typing import Any, Dict, List, Optional
import re

EXPECTED_CANVAS_HEIGHT = 252.0
EXPECTED_CANVAS_WIDTH = 432.0
SIZE_TOLERANCE = 1.0
TEXTBOX_POSITION_TOLERANCE = 2.0


def verify_textbox_gallery(
    pred_path: Path,
    gold_path: Path,
    *,
    canonical_saved_element,
    com_primary_with_saved_guardrail,
    error_with_details,
    read_shape_properties,
    read_shape_text_styles,
    verify_saved_insertion_context,
) -> VerificationResult:
    return com_primary_with_saved_guardrail(
        lambda: verify_textbox_gallery_com(
            pred_path,
            gold_path,
            error_with_details=error_with_details,
            read_shape_properties=read_shape_properties,
            read_shape_text_styles=read_shape_text_styles,
        ),
        lambda: verify_textbox_gallery_saved(
            pred_path,
            gold_path,
            canonical_saved_element=canonical_saved_element,
            error_with_details=error_with_details,
            verify_saved_insertion_context=verify_saved_insertion_context,
        ),
    )


def verify_textbox_gallery_saved(
    pred_path: Path,
    gold_path: Path,
    *,
    canonical_saved_element,
    error_with_details,
    verify_saved_insertion_context,
) -> VerificationResult:
    def semantic(signatures: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        return [
            {
                "geometry": signature["geometry"],
                "text": signature["text"],
                "paragraph_styles": signature["paragraph_styles"],
                "orientation": signature["orientation"],
                "errors": signature["errors"],
            }
            for signature in signatures
        ]

    gold_saved = semantic(
        families_shapes_insert_saved.extract_saved_textboxes(
            gold_path, canonical_saved_element=canonical_saved_element
        )
    )
    pred_saved = semantic(
        families_shapes_insert_saved.extract_saved_textboxes(
            pred_path, canonical_saved_element=canonical_saved_element
        )
    )
    if not gold_saved:
        return error_with_details(
            "Gold TextBox gallery target evidence is missing",
            {"match_mode": "saved_ooxml_textbox_gallery", "gold": gold_saved},
        )
    if not pred_saved:
        return VerificationResult(
            passed=False,
            message="FAIL: Predicted TextBox gallery target is missing",
            details={
                "match_mode": "saved_ooxml_textbox_gallery",
                "gold": gold_saved,
                "pred": pred_saved,
            },
        )
    if gold_saved != pred_saved:
        return VerificationResult(
            passed=False,
            message="FAIL: Saved TextBox gallery signature differs",
            details={
                "match_mode": "saved_ooxml_textbox_gallery",
                "gold": gold_saved,
                "pred": pred_saved,
            },
        )
    context_result = verify_saved_insertion_context(pred_path, gold_path)
    if context_result is not None:
        return context_result
    return VerificationResult(
        passed=True,
        message="PASS: TextBox gallery style matches",
        details={
            "match_mode": "saved_ooxml_textbox_gallery",
            "shape_count": len(gold_saved),
        },
    )


def verify_textbox_gallery_com(
    pred_path: Path,
    gold_path: Path,
    *,
    error_with_details,
    read_shape_properties,
    read_shape_text_styles,
) -> VerificationResult:
    """Verify TextBoxInsertGallery through the Word Object Model."""
    gold_shapes = read_shape_properties(str(gold_path))
    pred_shapes = read_shape_properties(str(pred_path))

    if not gold_shapes:
        return error_with_details(
            "Gold TextBox gallery target evidence is missing",
            {"match_mode": "textbox_gallery_effective_appearance"},
        )

    if len(gold_shapes) != len(pred_shapes):
        return VerificationResult(
            passed=False,
            message=f"FAIL: TextBox shape count mismatch (pred={len(pred_shapes)}, gold={len(gold_shapes)})",
            details={
                "pred_shape_count": len(pred_shapes),
                "gold_shape_count": len(gold_shapes),
                "pred_shape_types": [s.get("type") for s in pred_shapes],
                "gold_shape_types": [s.get("type") for s in gold_shapes],
            },
        )

    pred_shape_types = sorted([s.get("type") for s in pred_shapes])
    gold_shape_types = sorted([s.get("type") for s in gold_shapes])
    if pred_shape_types != gold_shape_types:
        return VerificationResult(
            passed=False,
            message="FAIL: TextBox shape types differ",
            details={
                "pred_shape_types": pred_shape_types,
                "gold_shape_types": gold_shape_types,
            },
        )

    position_mismatches = []
    for index, (gold_shape, pred_shape) in enumerate(zip(gold_shapes, pred_shapes), start=1):
        for axis in ("left", "top"):
            gold_value = families_shapes_insert_geometry.effective_shape_coordinate(
                gold_shape, axis
            )
            pred_value = families_shapes_insert_geometry.effective_shape_coordinate(
                pred_shape, axis
            )
            if abs(gold_value - pred_value) > TEXTBOX_POSITION_TOLERANCE:
                position_mismatches.append(f"{index}.{axis}: pred={pred_value}, gold={gold_value}")
        for key in ("width", "height"):
            gold_value = float(gold_shape.get(key) or 0)
            pred_value = float(pred_shape.get(key) or 0)
            if abs(gold_value - pred_value) > TEXTBOX_POSITION_TOLERANCE:
                position_mismatches.append(f"{index}.{key}: pred={pred_value}, gold={gold_value}")
    if position_mismatches:
        return VerificationResult(
            passed=False,
            message="FAIL: TextBox gallery effective layout differs",
            details={
                "match_mode": "textbox_gallery_effective_layout",
                "mismatches": position_mismatches,
                "tolerance": TEXTBOX_POSITION_TOLERANCE,
            },
        )

    def _norm_ws(text: str) -> str:
        return re.sub(r"\s+", " ", text or "").strip()

    def _shape_text_sig_list(doc_path: Path) -> List[Dict]:
        sigs = []
        for sh in read_shape_text_styles(str(doc_path), include_group_items=True):
            sigs.append(
                {
                    "index": sh.get("index"),
                    "type": sh.get("type"),
                    "auto_shape_type": sh.get("auto_shape_type"),
                    "width": sh.get("width"),
                    "height": sh.get("height"),
                    "rotation": sh.get("rotation"),
                    "wrap_type": sh.get("wrap_type"),
                    "text_orientation": sh.get("text_orientation"),
                    "text_margins": [
                        sh.get("text_margin_left"),
                        sh.get("text_margin_top"),
                        sh.get("text_margin_right"),
                        sh.get("text_margin_bottom"),
                    ],
                    "fill_visible": sh.get("fill_visible"),
                    "fill_type": sh.get("fill_type"),
                    "fill_color": sh.get("fill_color"),
                    "fill_transparency": sh.get("fill_transparency"),
                    "line_visible": sh.get("line_visible"),
                    "line_style": sh.get("line_style"),
                    "line_dash_style": sh.get("line_dash_style"),
                    "line_weight": sh.get("line_weight"),
                    "line_color": sh.get("line_color"),
                    "line_transparency": sh.get("line_transparency"),
                    "paragraph_styles": [
                        (_s or "").strip() for _s in (sh.get("paragraph_styles") or [])
                    ],
                    "paragraph_texts": [_norm_ws(t) for t in (sh.get("paragraph_texts") or [])],
                    "paragraph_alignments": sh.get("paragraph_alignments") or [],
                    "paragraph_fonts": sh.get("paragraph_fonts") or [],
                }
            )
        return sorted(
            sigs,
            key=lambda s: (
                str(s.get("index")),
                str(s.get("type")),
                str(s.get("paragraph_styles", [])),
                str(s.get("paragraph_texts", [])),
            ),
        )

    gold_sigs = _shape_text_sig_list(gold_path)
    pred_sigs = _shape_text_sig_list(pred_path)
    if gold_sigs != pred_sigs:
        return VerificationResult(
            passed=False,
            message="FAIL: TextBox gallery effective appearance differs",
            details={
                "match_mode": "textbox_gallery_effective_appearance",
                "gold": gold_sigs,
                "pred": pred_sigs,
            },
        )

    return VerificationResult(
        passed=True,
        message="PASS: TextBox gallery style matches",
        details={
            "match_mode": "textbox_gallery_signature",
            "shape_count": len(gold_shapes),
            "shape_types": gold_shape_types,
        },
    )


def verify_textbox_insert(
    pred_path: Path,
    gold_path: Path,
    *,
    canonical_saved_element,
    com_primary_with_saved_guardrail,
    error_with_details,
    read_shape_properties,
    verify_saved_insertion_context,
) -> VerificationResult:
    return com_primary_with_saved_guardrail(
        lambda: verify_textbox_insert_com(
            pred_path,
            gold_path,
            error_with_details=error_with_details,
            read_shape_properties=read_shape_properties,
        ),
        lambda: verify_textbox_insert_saved(
            pred_path,
            gold_path,
            canonical_saved_element=canonical_saved_element,
            error_with_details=error_with_details,
            verify_saved_insertion_context=verify_saved_insertion_context,
        ),
    )


def verify_textbox_insert_saved(
    pred_path: Path,
    gold_path: Path,
    *,
    canonical_saved_element,
    error_with_details,
    verify_saved_insertion_context,
) -> VerificationResult:
    keys = ("orientation", "geometry", "text", "errors")
    gold_saved = [
        {key: signature[key] for key in keys}
        for signature in families_shapes_insert_saved.extract_saved_textboxes(
            gold_path, canonical_saved_element=canonical_saved_element
        )
    ]
    pred_saved = [
        {key: signature[key] for key in keys}
        for signature in families_shapes_insert_saved.extract_saved_textboxes(
            pred_path, canonical_saved_element=canonical_saved_element
        )
    ]
    if not gold_saved:
        return error_with_details(
            "Gold drawn TextBox target evidence is missing",
            {"match_mode": "saved_ooxml_textbox_insert", "gold": gold_saved},
        )
    if not pred_saved:
        return VerificationResult(
            passed=False,
            message="FAIL: Predicted drawn TextBox target is missing",
            details={
                "match_mode": "saved_ooxml_textbox_insert",
                "gold": gold_saved,
                "pred": pred_saved,
            },
        )
    if any((signature["text"] or "").strip() for signature in gold_saved):
        return error_with_details(
            "Gold drawn TextBox is not empty as required by the contract",
            {"match_mode": "saved_ooxml_textbox_insert", "gold": gold_saved},
        )
    if any((signature["text"] or "").strip() for signature in pred_saved):
        return VerificationResult(
            passed=False,
            message="FAIL: Drawn TextBox must remain empty",
            details={
                "match_mode": "saved_ooxml_textbox_insert",
                "gold": gold_saved,
                "pred": pred_saved,
            },
        )
    if gold_saved != pred_saved:
        return VerificationResult(
            passed=False,
            message="FAIL: Saved TextBox position/size signature differs",
            details={
                "match_mode": "saved_ooxml_textbox_insert",
                "gold": gold_saved,
                "pred": pred_saved,
            },
        )
    context_result = verify_saved_insertion_context(pred_path, gold_path)
    if context_result is not None:
        return context_result
    return VerificationResult(
        passed=True,
        message=f"PASS: TextBox insert matches ({len(gold_saved)} shape(s))",
        details={
            "match_mode": "saved_ooxml_textbox_insert",
            "signature": gold_saved,
        },
    )


def verify_textbox_insert_com(
    pred_path: Path, gold_path: Path, *, error_with_details, read_shape_properties
) -> VerificationResult:
    """
    Verify TextBoxInsert / TextBoxInsertVerticalWord by comparing:
    - shape count and type
    - text frame orientation (text direction)
    - position (left, top) and size (width, height)
    """
    gold_shapes = read_shape_properties(str(gold_path))
    pred_shapes = read_shape_properties(str(pred_path))

    if not gold_shapes:
        return error_with_details(
            "Gold drawn TextBox target evidence is missing",
            {"match_mode": "textbox_insert"},
        )

    if len(gold_shapes) != len(pred_shapes):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Shape count mismatch (pred={len(pred_shapes)}, gold={len(gold_shapes)})",
            details={
                "pred_shape_count": len(pred_shapes),
                "gold_shape_count": len(gold_shapes),
            },
        )

    for i, (g, p) in enumerate(zip(gold_shapes, pred_shapes)):
        if g.get("type") != p.get("type"):
            return VerificationResult(
                passed=False,
                message=f"FAIL: Shape type mismatch at index {i} (pred={p.get('type')}, gold={g.get('type')})",
                details={"pred": p.get("type"), "gold": g.get("type")},
            )

        g_orient = g.get("text_frame_orientation")
        p_orient = p.get("text_frame_orientation")
        if g_orient != p_orient:
            return VerificationResult(
                passed=False,
                message=f"FAIL: Text orientation mismatch (pred={p_orient}, gold={g_orient})",
                details={"pred_orientation": p_orient, "gold_orientation": g_orient},
            )

        if (g.get("text") or "").strip():
            return error_with_details(
                "Gold drawn TextBox is not empty as required by the contract",
                {"shape_index": i, "gold_text": g.get("text")},
            )
        if (p.get("text") or "").strip():
            return VerificationResult(
                passed=False,
                message="FAIL: Drawn TextBox must remain empty",
                details={
                    "shape_index": i,
                    "pred_text": p.get("text"),
                },
            )

        usability_failure = families_shapes_insert_geometry.shape_usable_page_failure(p)
        if usability_failure:
            return VerificationResult(
                passed=False,
                message=f"FAIL: Drawn TextBox {usability_failure}",
                details={"shape_index": i, "pred": p},
            )

        tol = TEXTBOX_POSITION_TOLERANCE
        mismatches = []

        for key in ("left", "top"):
            g_abs = families_shapes_insert_geometry.effective_shape_coordinate(g, key)
            p_abs = families_shapes_insert_geometry.effective_shape_coordinate(p, key)
            if abs(g_abs - p_abs) > tol:
                mismatches.append(f"{key}: pred={p_abs}, gold={g_abs}")
        for key in ("width", "height"):
            g_val = g.get(key, 0)
            p_val = p.get(key, 0)
            if abs(g_val - p_val) > tol:
                mismatches.append(f"{key}: pred={p_val}, gold={g_val}")

        if mismatches:
            return VerificationResult(
                passed=False,
                message=f"FAIL: TextBox position/size mismatch ({'; '.join(mismatches)})",
                details={
                    "mismatches": mismatches,
                    "tolerance": tol,
                },
            )

    return VerificationResult(
        passed=True,
        message="PASS: TextBox insert matches (orientation, position, size)",
        details={
            "match_mode": "textbox_insert",
            "shape_count": len(gold_shapes),
        },
    )


def verify_gallery_all_shapes_and_canvas(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    com_primary_with_saved_guardrail,
    error_with_details,
    read_shape_properties,
    verify_saved_insertion_context,
) -> VerificationResult:
    return com_primary_with_saved_guardrail(
        lambda: verify_gallery_all_shapes_and_canvas_com(
            pred_path,
            gold_path,
            spec,
            error_with_details=error_with_details,
            read_shape_properties=read_shape_properties,
        ),
        lambda: verify_gallery_all_shapes_and_canvas_saved(
            pred_path,
            gold_path,
            error_with_details=error_with_details,
            verify_saved_insertion_context=verify_saved_insertion_context,
        ),
    )


def verify_gallery_all_shapes_and_canvas_saved(
    pred_path: Path, gold_path: Path, *, error_with_details, verify_saved_insertion_context
) -> VerificationResult:
    gold_geometry = families_shapes_insert_saved.extract_saved_autoshape_geometries(gold_path)
    pred_geometry = families_shapes_insert_saved.extract_saved_autoshape_geometries(pred_path)
    if not gold_geometry:
        return error_with_details(
            "Gold AutoShape target evidence is missing",
            {
                "match_mode": "saved_ooxml_autoshape_geometry",
                "gold": gold_geometry,
            },
        )
    if not pred_geometry:
        return VerificationResult(
            passed=False,
            message="FAIL: Predicted AutoShape target is missing",
            details={
                "match_mode": "saved_ooxml_autoshape_geometry",
                "gold": gold_geometry,
                "pred": pred_geometry,
            },
        )
    if gold_geometry != pred_geometry:
        return VerificationResult(
            passed=False,
            message="FAIL: Saved AutoShape geometry differs",
            details={
                "match_mode": "saved_ooxml_autoshape_geometry",
                "gold": gold_geometry,
                "pred": pred_geometry,
            },
        )
    context_result = verify_saved_insertion_context(pred_path, gold_path)
    if context_result is not None:
        return context_result
    return VerificationResult(
        passed=True,
        message="PASS: GalleryAllShapesAndCanvas geometry matches",
        details={
            "match_mode": "saved_ooxml_autoshape_geometry",
            "geometries": gold_geometry,
        },
    )


def verify_gallery_all_shapes_and_canvas_com(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    error_with_details,
    read_shape_properties,
) -> VerificationResult:
    """Verify inserted shape identity through the Word Object Model."""
    gold_shapes = read_shape_properties(str(gold_path))
    pred_shapes = read_shape_properties(str(pred_path))

    if not gold_shapes:
        return error_with_details(
            "Gold AutoShape target evidence is missing",
            {"match_mode": "word_com_autoshape_semantics"},
        )

    if len(gold_shapes) != len(pred_shapes):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Shape count mismatch (pred={len(pred_shapes)}, gold={len(gold_shapes)})",
            details={
                "pred_shape_count": len(pred_shapes),
                "gold_shape_count": len(gold_shapes),
                "pred_shape_types": [s.get("type") for s in pred_shapes],
                "gold_shape_types": [s.get("type") for s in gold_shapes],
            },
        )

    pred_types = sorted(s.get("type") for s in pred_shapes)
    gold_types = sorted(s.get("type") for s in gold_shapes)
    if pred_types != gold_types:
        return VerificationResult(
            passed=False,
            message="FAIL: Shape type list mismatch for GalleryAllShapesAndCanvas",
            details={
                "pred_shape_types": pred_types,
                "gold_shape_types": gold_types,
            },
        )

    # Compare auto_shape_type to distinguish shapes within the same msoShapeType
    pred_auto = sorted(s.get("auto_shape_type", -1) for s in pred_shapes)
    gold_auto = sorted(s.get("auto_shape_type", -1) for s in gold_shapes)
    if pred_auto != gold_auto:
        return VerificationResult(
            passed=False,
            message="FAIL: AutoShapeType mismatch for GalleryAllShapesAndCanvas",
            details={
                "pred_auto_shape_types": pred_auto,
                "gold_auto_shape_types": gold_auto,
            },
        )

    expected_shape_type = None
    if spec:
        expected_shape_type = (spec.get("parameter") or {}).get("shape_type")

    if expected_shape_type is not None:
        expected_shape_type = int(expected_shape_type)
        gold_match = {
            "count": sum(
                1 for shape in gold_shapes if shape.get("auto_shape_type") == expected_shape_type
            )
        }
        pred_match = {
            "count": sum(
                1 for shape in pred_shapes if shape.get("auto_shape_type") == expected_shape_type
            )
        }
        if gold_match["count"] != 1:
            return error_with_details(
                "Gold does not contain exactly one requested AutoShape target",
                {
                    "expected_shape_type": expected_shape_type,
                    "gold_match": gold_match,
                },
            )
        if pred_match["count"] != gold_match["count"]:
            return VerificationResult(
                passed=False,
                message=(
                    "FAIL: Target shape_type count mismatch "
                    f"(type={expected_shape_type}, pred={pred_match.get('count')}, gold={gold_match.get('count')})"
                ),
                details={
                    "expected_shape_type": expected_shape_type,
                    "pred_match": pred_match,
                    "gold_match": gold_match,
                },
            )

    for index, shape in enumerate(pred_shapes, start=1):
        usability_failure = families_shapes_insert_geometry.shape_usable_page_failure(shape)
        if usability_failure:
            return VerificationResult(
                passed=False,
                message=f"FAIL: AutoShape {usability_failure}",
                details={
                    "match_mode": "word_com_autoshape_visibility",
                    "shape_index": index,
                    "pred": shape,
                },
            )
        if (shape.get("text") or "").strip():
            return VerificationResult(
                passed=False,
                message="FAIL: AutoShape contains unrequested visible text",
                details={
                    "match_mode": "word_com_autoshape_text",
                    "shape_index": index,
                    "pred_text": shape.get("text"),
                },
            )

    return VerificationResult(
        passed=True,
        message="PASS: GalleryAllShapesAndCanvas shape types match",
        details={
            "shape_count": len(pred_shapes),
            "shape_types": pred_types,
            "expected_shape_type": expected_shape_type,
        },
    )


def verify_drawing_canvas(
    pred_path: Path, gold_path: Path, *, read_find_shapes_by_type
) -> VerificationResult:
    """
    Verify ink drawing canvas exists with correct size.

    Checks:
    1. Shape with Type = msoCanvas (20) exists
    2. Canvas size is 100x200 points (as per instruction)
    """
    MSO_CANVAS = 20  # msoCanvas shape type

    pred_info = read_find_shapes_by_type(str(pred_path), MSO_CANVAS)
    gold_info = read_find_shapes_by_type(str(gold_path), MSO_CANVAS)

    # Check if canvas exists
    if not pred_info["found"]:
        return VerificationResult(
            passed=False,
            message="FAIL: No drawing canvas (msoCanvas) found in document",
            details={"pred_canvas_count": 0, "gold_canvas_count": gold_info["count"]},
        )

    # Check count matches gold
    if pred_info["count"] != gold_info["count"]:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Canvas count mismatch (pred={pred_info['count']}, gold={gold_info['count']})",
            details={
                "pred_count": pred_info["count"],
                "gold_count": gold_info["count"],
                "pred_shapes": pred_info["shapes"],
            },
        )

    # Check canvas size (100x200 points as per instruction)
    for canvas in pred_info["shapes"]:
        width = canvas.get("width", 0)
        height = canvas.get("height", 0)

        width_ok = abs(width - EXPECTED_CANVAS_WIDTH) <= SIZE_TOLERANCE
        height_ok = abs(height - EXPECTED_CANVAS_HEIGHT) <= SIZE_TOLERANCE

        if not width_ok or not height_ok:
            return VerificationResult(
                passed=False,
                message=f"FAIL: Canvas size mismatch (pred={width}x{height}, expected={EXPECTED_CANVAS_WIDTH}x{EXPECTED_CANVAS_HEIGHT})",
                details={
                    "pred_width": width,
                    "pred_height": height,
                    "expected_width": EXPECTED_CANVAS_WIDTH,
                    "expected_height": EXPECTED_CANVAS_HEIGHT,
                    "canvas": canvas,
                },
            )

    return VerificationResult(
        passed=True,
        message=f"PASS: Drawing canvas found with correct size ({pred_info['count']} canvas, {EXPECTED_CANVAS_WIDTH}x{EXPECTED_CANVAS_HEIGHT}pt)",
        details={
            "canvas_count": pred_info["count"],
            "expected_size": f"{EXPECTED_CANVAS_WIDTH}x{EXPECTED_CANVAS_HEIGHT}",
            "canvas_shapes": pred_info["shapes"],
        },
    )
