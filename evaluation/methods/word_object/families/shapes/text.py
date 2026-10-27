"""families / shapes / text with explicit evidence dependencies."""

from evaluation.methods.base import (
    VerificationResult,
    combine_com_primary_with_ooxml_guardrail,
    tag_verification_source,
)
from evaluation.methods.word_object.core import DocumentOpenError
from ...observations import require_observations
from evaluation.methods.word_object.families.shapes import (
    extract_saved as families_shapes_extract_saved,
)
from pathlib import Path
from typing import Dict, List, Optional


def textbox_shapes(docx_path: Path, *, read_shape_properties) -> List[Dict]:
    shapes = read_shape_properties(str(docx_path))
    return sorted(
        [shape for shape in shapes if shape.get("text") is not None],
        key=lambda s: s.get("name", ""),
    )


def verify_textbox_insert(
    pred_path: Path, gold_path: Path, *, read_shape_properties
) -> VerificationResult:
    """Verify inserted textbox count and text content."""
    gold_boxes = textbox_shapes(gold_path, read_shape_properties=read_shape_properties)
    pred_boxes = textbox_shapes(pred_path, read_shape_properties=read_shape_properties)

    if len(gold_boxes) != len(pred_boxes):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Textbox count mismatch (gold={len(gold_boxes)}, pred={len(pred_boxes)})",
            details={"gold_count": len(gold_boxes), "pred_count": len(pred_boxes)},
        )
    if not gold_boxes:
        return VerificationResult(passed=False, message="FAIL: No textbox found", details={})

    differences = []
    for idx, (gold_box, pred_box) in enumerate(zip(gold_boxes, pred_boxes), start=1):
        if (gold_box.get("text") or "").strip() != (pred_box.get("text") or "").strip():
            differences.append(
                {
                    "textbox_index": idx,
                    "property": "text",
                    "gold": (gold_box.get("text") or "").strip(),
                    "pred": (pred_box.get("text") or "").strip(),
                }
            )

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Textbox content mismatch ({len(differences)} difference(s))",
            details={"differences": differences},
        )
    return VerificationResult(
        passed=True,
        message="PASS: Textbox content matches",
        details={"textbox_count": len(gold_boxes)},
    )


def verify_textbox_text_fill(
    pred_path: Path, gold_path: Path, *, read_shape_properties
) -> VerificationResult:
    """Verify text color/fill inside textbox shapes."""
    gold_boxes = textbox_shapes(gold_path, read_shape_properties=read_shape_properties)
    pred_boxes = textbox_shapes(pred_path, read_shape_properties=read_shape_properties)

    if len(gold_boxes) != len(pred_boxes):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Textbox count mismatch (gold={len(gold_boxes)}, pred={len(pred_boxes)})",
            details={"gold_count": len(gold_boxes), "pred_count": len(pred_boxes)},
        )

    differences = []
    for idx, (gold_box, pred_box) in enumerate(zip(gold_boxes, pred_boxes), start=1):
        for prop in ("text_font_color", "text_fill_color", "text_fill_type"):
            gold_val = gold_box.get(prop)
            pred_val = pred_box.get(prop)
            if gold_val is None and pred_val is None:
                continue
            if gold_val != pred_val:
                differences.append(
                    {
                        "textbox_index": idx,
                        "property": prop,
                        "gold": gold_val,
                        "pred": pred_val,
                    }
                )

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Textbox text fill mismatch ({len(differences)} difference(s))",
            details={"differences": differences},
        )
    return VerificationResult(
        passed=True,
        message="PASS: Textbox text fill matches",
        details={"textbox_count": len(gold_boxes)},
    )


def verify_text_alignment(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    get_composition_steps,
    read_shape_properties,
) -> VerificationResult:
    """
    Verify text properties inside textbox shapes by comparing all
    text-related properties between gold and pred.

    Compares: text_alignment, text_frame_orientation, text_vertical_anchor.
    No command-specific logic — the gold document defines what's correct.
    """
    idMso = spec.get("idMso", "") if spec else ""
    saved_guardrail = None
    if idMso == "TextAlignGallery":
        gold_anchors = families_shapes_extract_saved.extract_saved_text_vertical_anchors(gold_path)
        pred_anchors = families_shapes_extract_saved.extract_saved_text_vertical_anchors(pred_path)
        if gold_anchors and pred_anchors is not None:
            if gold_anchors != pred_anchors:
                saved_guardrail = VerificationResult(
                    passed=False,
                    message="FAIL: Saved text vertical alignment differs",
                    details={
                        "match_mode": "saved_ooxml_text_vertical_anchor",
                        "gold": gold_anchors,
                        "pred": pred_anchors,
                    },
                )
            else:
                saved_guardrail = VerificationResult(
                    passed=True,
                    message="PASS: Saved text vertical alignment matches",
                    details={
                        "match_mode": "saved_ooxml_text_vertical_anchor",
                        "anchors": gold_anchors,
                    },
                )

    try:
        gold_shapes = read_shape_properties(str(gold_path))
        pred_shapes = read_shape_properties(str(pred_path))
    except DocumentOpenError:
        raise
    except Exception as com_error:
        if saved_guardrail is not None:
            return tag_verification_source(
                saved_guardrail,
                "saved_ooxml_fallback",
                word_com_error=str(com_error),
            )
        raise

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

    # Sort shapes by name for consistent comparison
    gold_sorted = sorted(gold_shapes, key=lambda s: s.get("name", ""))
    pred_sorted = sorted(pred_shapes, key=lambda s: s.get("name", ""))

    # TextAlignGallery directly targets the vertical anchor. Paragraph
    # alignment and text orientation are still visible preservation
    # invariants: changing either is collateral document damage, just as
    # formatting text outside the requested range would be. This is
    # intentionally different from locale/language metadata, which is
    # environment-dependent and non-semantic.
    text_props = [
        "text_alignment",
        "text_frame_orientation",
        "text_vertical_anchor",
    ]
    composition = (spec or {}).get("composition") or {}
    composition_steps = get_composition_steps(spec)
    inserts_new_textbox = any(
        step.get("operation") == "insert_textbox" for step in composition_steps
    )
    owns_vertical_alignment = any(
        step.get("idMso") == "TextAlignGallery" for step in composition_steps
    ) or any(
        str(intent.get("atomic_case") or "").casefold().startswith("textaligngallery/")
        for intent in composition.get("atomic_intents", [])
        if isinstance(intent, dict)
    )
    if inserts_new_textbox and not owns_vertical_alignment:
        # A newly created textbox has a Word/template-dependent vertical
        # anchor default. When the composition requests no vertical
        # alignment, centered and edge-aligned text are equivalent; the
        # surrounding verify() path still enforces visible preservation.
        text_props.remove("text_vertical_anchor")
    differences = []

    for gold_shape, pred_shape in zip(gold_sorted, pred_sorted):
        shape_diffs = []
        required = []
        if idMso == "TextAlignGallery":
            required = ["text_vertical_anchor"]
        elif idMso in {"TextDirection", "TextDirectionGallery"}:
            required = ["text_frame_orientation"]
        elif idMso in {"TextAlignLeft", "TextAlignCenter", "TextAlignRight", "TextAlignLetterJustify",
                       "TextAlignWordJustify", "TextAlignStretchJustify"}:
            required = ["text_alignment"]
        error = require_observations({"gold": gold_shape, "pred": pred_shape}, required, context=idMso)
        if error is not None:
            if saved_guardrail is not None:
                return tag_verification_source(saved_guardrail, "saved_ooxml_fallback",
                    word_com_error=error.error, unavailable_observations=error.details)
            return error
        for prop in text_props:
            gold_val = gold_shape.get(prop)
            pred_val = pred_shape.get(prop)
            # Skip if property is absent from both (not applicable to this shape)
            if gold_val is None and pred_val is None:
                continue
            if gold_val != pred_val:
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
        com_result = VerificationResult(
            passed=False,
            message=f"FAIL: {len(differences)} shape(s) have text property differences",
            details={"differences": differences},
        )
        if saved_guardrail is not None:
            com_result.details["saved_ooxml_guardrail"] = saved_guardrail.to_dict()
            return tag_verification_source(com_result, "word_com+saved_ooxml_guardrail")
        return tag_verification_source(com_result, "word_com")

    com_result = VerificationResult(
        passed=True,
        message="PASS: Text alignment/direction matches",
        details={"shape_count": len(gold_shapes)},
    )
    if saved_guardrail is not None:
        return combine_com_primary_with_ooxml_guardrail(com_result, saved_guardrail)
    return tag_verification_source(com_result, "word_com")


def verify_wordart(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, read_shape_properties
) -> VerificationResult:
    """
    Verify WordArt TextEffect properties match between gold and pred.

    Compares all WordArt-related properties generically:
    normalized_height, kerned_pairs, tracking, rotated_chars,
    text_frame_orientation, and shape rotation.

    No command-specific logic — the gold document defines what's correct.
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

    # All WordArt-related properties to compare
    # Exact-match properties (bool/enum/int)
    exact_props = [
        "wordart_normalized_height",
        "wordart_kerned_pairs",
        "wordart_rotated_chars",
        "text_frame_orientation",
    ]
    # Numeric properties with tolerance
    numeric_props = {
        "wordart_tracking": 0.1,  # tracking tolerance
        "rotation": 1.0,  # degree tolerance (catches shape rotation hacks)
    }

    differences = []

    for gold_shape, pred_shape in zip(gold_sorted, pred_sorted):
        shape_diffs = []

        # Compare exact-match properties
        for prop in exact_props:
            gold_val = gold_shape.get(prop)
            pred_val = pred_shape.get(prop)
            if gold_val is None and pred_val is None:
                continue
            if gold_val != pred_val:
                shape_diffs.append(
                    {
                        "property": prop,
                        "gold": gold_val,
                        "pred": pred_val,
                    }
                )

        # Compare numeric properties with tolerance
        for prop, tolerance in numeric_props.items():
            gold_val = gold_shape.get(prop)
            pred_val = pred_shape.get(prop)
            if gold_val is None and pred_val is None:
                continue
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
            elif gold_val != pred_val:
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
        return VerificationResult(
            passed=False,
            message=f"FAIL: {len(differences)} shape(s) have WordArt property differences",
            details={"differences": differences},
        )

    return VerificationResult(
        passed=True,
        message="PASS: WordArt properties match",
        details={"shape_count": len(gold_shapes)},
    )
