"""families / smartart / structure with explicit evidence dependencies."""

from evaluation.methods.base import (
    VerificationResult,
    combine_com_primary_with_ooxml_guardrail,
    tag_verification_source,
)
from evaluation.methods.word_object.families.smartart import refresh as families_smartart_refresh
from pathlib import Path
from typing import Dict, Optional
from ...observations import require_observations


def verify_smartart_structural(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    composition_steps,
    extract_refreshed_smartart_projection,
    extract_smartart_node_details,
    get_cleanup_word,
    read_smartart_drawing_info,
    read_smartart_node_scale_info,
    read_smartart_placement,
    read_smartart_reset_state,
) -> VerificationResult:
    """
    Verify SmartArt structural commands (add/remove nodes, reorder, resize, etc.).

    Uses extract_smartart_node_details() for node-level comparison.
    """
    idMso = spec.get("idMso", "") if spec else ""
    # Read the live SmartArt object model first.  The diagram drawing cache
    # is checked only after this semantic comparison has passed.
    gold_det = extract_smartart_node_details(gold_path)
    pred_det = extract_smartart_node_details(pred_path)

    if not gold_det.get("found"):
        return VerificationResult.error_result(
            "No SmartArt found in gold document", gold=gold_det,
            error_owner="gold_fixture_or_evaluator", verification_source="word_com")
    if not pred_det.get("found"):
        return VerificationResult(
            passed=False,
            message="FAIL: No SmartArt found in pred document",
            details={"pred": pred_det, "verification_source": "word_com"},
        )

    if idMso in {"SmartArtAddShapeAfter", "SmartArtAddShapeBefore", "SmartArtAddShapeAbove",
                 "SmartArtAddShapeBelow", "SmartArtAddAssistant"}:
        required = ["node_count", "all_node_count", "node_texts", "all_node_texts", "all_node_levels", "all_node_types"]
    elif idMso in {"SmartArtReorderUp", "SmartArtReorderDown"}:
        required = ["node_texts", "node_count"]
    elif idMso in {"SmartArtPromote", "SmartArtDemote"}:
        required = ["all_node_levels", "all_node_texts", "all_node_types", "node_count"]
    elif idMso in {"SmartArtLargerShape", "SmartArtSmallerShape"}:
        required = ["node_count", "node_texts", "node_levels", "node_sizes"]
    elif idMso == "SmartArtRightToLeft":
        required = ["reverse"]
    elif idMso == "SmartArtResetGraphic":
        required = ["layout_name", "node_count", "node_texts", "all_node_types"]
    else:
        required = ["layout_name", "node_count", "node_texts", "node_levels", "all_node_count",
                    "all_node_texts", "all_node_levels", "all_node_types", "reverse", "color_name", "quick_style_name"]
    error = require_observations({"gold": gold_det, "pred": pred_det}, required,
                                 context=f"SmartArt structure {idMso}")
    if error is not None:
        return error

    differences = []

    heading = (spec or {}).get("precondition", {}).get("heading")
    is_insert_composition = any(
        step.get("operation") == "insert_smartart" for step in composition_steps(spec)
    )
    if heading and is_insert_composition:
        gold_placement = read_smartart_placement(str(gold_path), heading)
        pred_placement = read_smartart_placement(str(pred_path), heading)
        if not gold_placement.get("found"):
            return VerificationResult.error_result("Gold: SmartArt placement could not be resolved")
        if not gold_placement.get("below_heading"):
            return VerificationResult.error_result(
                "Gold: SmartArt is not below its required heading"
            )
        if not pred_placement.get("found") or not pred_placement.get("below_heading"):
            differences.append(
                {
                    "property": "smartart_below_heading",
                    "gold": gold_placement,
                    "pred": pred_placement,
                }
            )
    saved_differences = []
    saved_guardrail_checked = False
    raw_cache_trigger = None

    if idMso in (
        "SmartArtAddShapeAfter",
        "SmartArtAddShapeBefore",
        "SmartArtAddShapeAbove",
        "SmartArtAddShapeBelow",
        "SmartArtAddAssistant",
    ):
        # Node count and text order should match (including all sub-nodes).
        for prop in [
            "node_count",
            "all_node_count",
            "node_texts",
            "all_node_texts",
            "all_node_levels",
            "all_node_types",
        ]:
            if gold_det.get(prop) != pred_det.get(prop):
                differences.append(
                    {
                        "property": prop,
                        "gold": gold_det.get(prop),
                        "pred": pred_det.get(prop),
                    }
                )
        gold_drawing = read_smartart_drawing_info(str(gold_path))
        pred_drawing = read_smartart_drawing_info(str(pred_path))
        saved_guardrail_checked = True
        if gold_drawing.get("node_order") != pred_drawing.get("node_order"):
            raw_cache_trigger = {
                "property": "drawing_node_order",
                "gold": gold_drawing.get("node_order"),
                "pred": pred_drawing.get("node_order"),
            }

    elif idMso in ("SmartArtReorderUp", "SmartArtReorderDown"):
        # COM Nodes can expose the changed data model even when the saved
        # drawing cache still renders the old order. Compare both views.
        for prop in ["node_texts", "node_count"]:
            if gold_det.get(prop) != pred_det.get(prop):
                differences.append(
                    {
                        "property": prop,
                        "gold": gold_det.get(prop),
                        "pred": pred_det.get(prop),
                    }
                )
        gold_drawing = read_smartart_drawing_info(str(gold_path))
        pred_drawing = read_smartart_drawing_info(str(pred_path))
        saved_guardrail_checked = True
        if gold_drawing.get("node_order") != pred_drawing.get("node_order"):
            raw_cache_trigger = {
                "property": "drawing_node_order",
                "gold": gold_drawing.get("node_order"),
                "pred": pred_drawing.get("node_order"),
            }

    elif idMso in ("SmartArtPromote", "SmartArtDemote"):
        # Node levels change. The saved drawing cache must also reflect it.
        for prop in ["all_node_levels", "all_node_texts", "all_node_types", "node_count"]:
            if gold_det.get(prop) != pred_det.get(prop):
                differences.append(
                    {
                        "property": prop,
                        "gold": gold_det.get(prop),
                        "pred": pred_det.get(prop),
                    }
                )
        gold_drawing = read_smartart_drawing_info(str(gold_path))
        pred_drawing = read_smartart_drawing_info(str(pred_path))
        saved_guardrail_checked = True
        if gold_drawing.get("node_order") != pred_drawing.get("node_order"):
            raw_cache_trigger = {
                "property": "drawing_node_order",
                "gold": gold_drawing.get("node_order"),
                "pred": pred_drawing.get("node_order"),
            }

    elif idMso in ("SmartArtLargerShape", "SmartArtSmallerShape"):
        # SmartArt auto-layout can choose different absolute dimensions
        # across Word builds. The command semantics are relative: the
        # target node (Charlie in this benchmark family) should become
        # larger/smaller than its siblings while preserving structure.
        # COM remains the live semantic projection, but the resize must
        # also persist as custom scale on the SmartArt data-model point.
        # drawing*.xml is only a render cache and can retain stale sizes.
        gold_sizes = gold_det.get("node_sizes", [])
        pred_sizes = pred_det.get("node_sizes", [])
        gold_texts = gold_det.get("node_texts", [])
        pred_texts = pred_det.get("node_texts", [])
        gold_scales = read_smartart_node_scale_info(str(gold_path))
        pred_scales = read_smartart_node_scale_info(str(pred_path))
        saved_guardrail_checked = True
        for prop in ["node_count", "node_texts", "node_levels"]:
            if gold_det.get(prop) != pred_det.get(prop):
                differences.append(
                    {
                        "property": prop,
                        "gold": gold_det.get(prop),
                        "pred": pred_det.get(prop),
                    }
                )
        if len(gold_sizes) != len(pred_sizes):
            differences.append(
                {
                    "property": "node_sizes_count",
                    "gold": len(gold_sizes),
                    "pred": len(pred_sizes),
                }
            )
        else:

            def _area(size):
                if not size:
                    return None
                width = size.get("width")
                height = size.get("height")
                if width is None or height is None:
                    return None
                return float(width) * float(height)

            def _target_index(texts):
                for i, text in enumerate(texts or []):
                    if str(text).strip().lower() == "charlie":
                        return i
                return None

            def _target_ratio(sizes, texts):
                idx = _target_index(texts)
                if idx is None or idx >= len(sizes):
                    return None
                target = _area(sizes[idx])
                siblings = [_area(size) for i, size in enumerate(sizes) if i != idx]
                siblings = [area for area in siblings if area and area > 0]
                if not target or not siblings:
                    return None
                return target / (sum(siblings) / len(siblings))

            def _saved_target_ratio(scale_info):
                nodes = scale_info.get("nodes") or []
                target = next(
                    (
                        node
                        for node in nodes
                        if str(node.get("text", "")).strip().lower() == "charlie"
                    ),
                    None,
                )
                siblings = [
                    node for node in nodes if str(node.get("text", "")).strip().lower() != "charlie"
                ]
                if target is None or not siblings:
                    return None

                def _scale_area(node):
                    scale_x = node.get("scale_x")
                    scale_y = node.get("scale_y")
                    if scale_x is None or scale_y is None:
                        return None
                    return float(scale_x) * float(scale_y)

                target_area = _scale_area(target)
                sibling_areas = [_scale_area(node) for node in siblings]
                sibling_areas = [area for area in sibling_areas if area]
                if not target_area or not sibling_areas:
                    return None
                return target_area / (sum(sibling_areas) / len(sibling_areas))

            gold_ratio = _target_ratio(gold_sizes, gold_texts)
            pred_ratio = _target_ratio(pred_sizes, pred_texts)
            gold_saved_ratio = _saved_target_ratio(gold_scales)
            pred_saved_ratio = _saved_target_ratio(pred_scales)
            if gold_saved_ratio is None:
                return VerificationResult(
                    passed=False,
                    message=("ERROR: Gold SmartArt custom-scale evidence " "could not be resolved"),
                    details={
                        "error_owner": "gold_fixture",
                        "verification_source": "saved_ooxml_guardrail",
                        "gold": gold_scales,
                    },
                    error=("Gold SmartArt custom-scale evidence could not " "be resolved"),
                )
            if pred_saved_ratio is None:
                saved_differences.append(
                    {
                        "property": "data_model_target_scale_ratio",
                        "gold": round(gold_saved_ratio, 3),
                        "pred": pred_saved_ratio,
                    }
                )
            if gold_ratio is None or pred_ratio is None:
                differences.append(
                    {
                        "property": "target_size_ratio",
                        "gold": gold_ratio,
                        "pred": pred_ratio,
                    }
                )
            elif idMso == "SmartArtLargerShape":
                threshold = max(1.10, gold_ratio * 0.75)
                saved_threshold = max(1.10, gold_saved_ratio * 0.75)
                if pred_ratio < threshold:
                    differences.append(
                        {
                            "property": "target_size_ratio",
                            "gold": round(gold_ratio, 3),
                            "pred": round(pred_ratio, 3),
                            "expected": f">= {round(threshold, 3)}",
                        }
                    )
                if pred_saved_ratio is not None and pred_saved_ratio < saved_threshold:
                    saved_differences.append(
                        {
                            "property": "data_model_target_scale_ratio",
                            "gold": round(gold_saved_ratio, 3),
                            "pred": round(pred_saved_ratio, 3),
                            "expected": f">= {round(saved_threshold, 3)}",
                        }
                    )
            else:
                threshold = min(0.90, gold_ratio * 1.25)
                saved_threshold = min(0.90, gold_saved_ratio * 1.25)
                if pred_ratio > threshold:
                    differences.append(
                        {
                            "property": "target_size_ratio",
                            "gold": round(gold_ratio, 3),
                            "pred": round(pred_ratio, 3),
                            "expected": f"<= {round(threshold, 3)}",
                        }
                    )
                if pred_saved_ratio is not None and pred_saved_ratio > saved_threshold:
                    saved_differences.append(
                        {
                            "property": "data_model_target_scale_ratio",
                            "gold": round(gold_saved_ratio, 3),
                            "pred": round(pred_saved_ratio, 3),
                            "expected": f"<= {round(saved_threshold, 3)}",
                        }
                    )

    elif idMso == "SmartArtRightToLeft":
        if gold_det.get("reverse") != pred_det.get("reverse"):
            differences.append(
                {
                    "property": "reverse",
                    "gold": gold_det.get("reverse"),
                    "pred": pred_det.get("reverse"),
                }
            )

    elif idMso == "SmartArtResetGraphic":
        for prop in [
            "layout_name",
            "node_count",
            "node_texts",
            "all_node_types",
        ]:
            if gold_det.get(prop) != pred_det.get(prop):
                differences.append(
                    {
                        "property": prop,
                        "gold": gold_det.get(prop),
                        "pred": pred_det.get(prop),
                    }
                )
        gold_reset = read_smartart_reset_state(str(gold_path))
        pred_reset = read_smartart_reset_state(str(pred_path))
        saved_guardrail_checked = True
        if not gold_reset.get("found") or not gold_reset.get("nodes"):
            return VerificationResult(
                passed=False,
                message="ERROR: Gold SmartArt Reset state is unavailable",
                details={
                    "gold": gold_reset,
                    "idMso": idMso,
                    "error_owner": "gold_fixture_or_evaluator",
                    "verification_source": "word_com",
                },
                error="Gold SmartArt Reset state is unavailable",
            )
        if not pred_reset.get("found") or not pred_reset.get("nodes"):
            return VerificationResult(
                passed=False,
                message="FAIL: Prediction SmartArt Reset state is unavailable",
                details={
                    "pred": pred_reset,
                    "idMso": idMso,
                    "verification_source": "word_com",
                },
            )

        error = require_observations({"gold": gold_reset, "pred": pred_reset},
                                     ["color_name", "quick_style_name"], context="SmartArt Reset galleries")
        if error is not None:
            return error
        for owner, reset in (("gold", gold_reset), ("pred", pred_reset)):
            for node_index, node in enumerate(reset["nodes"], 1):
                if not node.get("shape_found"):
                    continue  # Observed absence is handled below.
                for effect in ("fill", "text_fill", "line", "text_outline", "shadow"):
                    paint = node.get(effect) or {}
                    fields = ["visible"]
                    if paint.get("visible"):
                        if effect in ("fill", "text_fill"):
                            fields += ["rgb", "transparency"]
                        elif effect in ("line", "text_outline"):
                            fields += ["rgb", "weight", "dash_style"]
                    error = require_observations({owner: paint}, fields,
                        context=f"SmartArt Reset node {node_index} {effect}")
                    if error is not None:
                        return error

        for prop in ["color_name", "quick_style_name"]:
            if gold_reset.get(prop) != pred_reset.get(prop):
                differences.append(
                    {
                        "property": f"reset_{prop}",
                        "gold": gold_reset.get(prop),
                        "pred": pred_reset.get(prop),
                    }
                )

        def _visible(value):
            return value not in (None, False, 0)

        def _opaque(paint):
            transparency = paint.get("transparency")
            return (
                _visible(paint.get("visible"))
                and transparency is not None
                and 0.0 <= float(transparency) < 0.999
            )

        gold_nodes = gold_reset.get("nodes") or []
        pred_nodes = pred_reset.get("nodes") or []
        if len(gold_nodes) != len(pred_nodes):
            differences.append(
                {
                    "property": "reset_effective_node_count",
                    "gold": len(gold_nodes),
                    "pred": len(pred_nodes),
                }
            )
        for index, gold_node in enumerate(gold_nodes):
            if index >= len(pred_nodes):
                break
            pred_node = pred_nodes[index]
            label = gold_node.get("text")
            if not gold_node.get("shape_found"):
                return VerificationResult(
                    passed=False,
                    message=("ERROR: Gold SmartArt Reset node shape is " f"unavailable ({label})"),
                    details={
                        "gold": gold_reset,
                        "idMso": idMso,
                        "error_owner": "gold_fixture_or_evaluator",
                        "verification_source": "word_com",
                    },
                    error="Gold SmartArt Reset node shape is unavailable",
                )
            if not pred_node.get("shape_found"):
                differences.append(
                    {
                        "property": "reset_node_shape_present",
                        "node": label,
                        "gold": True,
                        "pred": False,
                    }
                )
                continue

            gold_fill = gold_node.get("fill") or {}
            pred_fill = pred_node.get("fill") or {}
            if not _opaque(gold_fill) or gold_fill.get("rgb") is None:
                return VerificationResult(
                    passed=False,
                    message=(
                        "ERROR: Gold SmartArt Reset fill evidence is " f"incomplete ({label})"
                    ),
                    details={
                        "gold": gold_reset,
                        "idMso": idMso,
                        "error_owner": "gold_fixture_or_evaluator",
                        "verification_source": "word_com",
                    },
                    error="Gold SmartArt Reset fill evidence is incomplete",
                )
            if not _opaque(pred_fill) or pred_fill.get("rgb") != gold_fill.get("rgb"):
                differences.append(
                    {
                        "property": "reset_effective_fill",
                        "node": label,
                        "gold": gold_fill,
                        "pred": pred_fill,
                    }
                )

            gold_text_fill = gold_node.get("text_fill") or {}
            pred_text_fill = pred_node.get("text_fill") or {}
            if not _opaque(gold_text_fill) or gold_text_fill.get("rgb") is None:
                return VerificationResult(
                    passed=False,
                    message=(
                        "ERROR: Gold SmartArt Reset glyph evidence is " f"incomplete ({label})"
                    ),
                    details={
                        "gold": gold_reset,
                        "idMso": idMso,
                        "error_owner": "gold_fixture_or_evaluator",
                        "verification_source": "word_com",
                    },
                    error="Gold SmartArt Reset glyph evidence is incomplete",
                )
            if not _opaque(pred_text_fill) or pred_text_fill.get("rgb") != gold_text_fill.get(
                "rgb"
            ):
                differences.append(
                    {
                        "property": "reset_effective_text_fill",
                        "node": label,
                        "gold": gold_text_fill,
                        "pred": pred_text_fill,
                    }
                )

            for effect_name in ["line", "text_outline"]:
                gold_effect = gold_node.get(effect_name) or {}
                pred_effect = pred_node.get(effect_name) or {}
                gold_effect_visible = _visible(gold_effect.get("visible"))
                pred_effect_visible = _visible(pred_effect.get("visible"))
                effect_matches = gold_effect_visible == pred_effect_visible
                if effect_matches and gold_effect_visible:
                    for field in ["rgb", "weight", "dash_style"]:
                        gold_value = gold_effect.get(field)
                        pred_value = pred_effect.get(field)
                        if field == "weight" and None not in (gold_value, pred_value):
                            if abs(float(gold_value) - float(pred_value)) > 0.01:
                                effect_matches = False
                        elif gold_value != pred_value:
                            effect_matches = False
                if not effect_matches:
                    differences.append(
                        {
                            "property": f"reset_residual_{effect_name}",
                            "node": label,
                            "gold": gold_effect,
                            "pred": pred_effect,
                        }
                    )

            gold_shadow = gold_node.get("shadow") or {}
            pred_shadow = pred_node.get("shadow") or {}
            if _visible(gold_shadow.get("visible")) != _visible(pred_shadow.get("visible")):
                differences.append(
                    {
                        "property": "reset_residual_shadow",
                        "node": label,
                        "gold": gold_shadow,
                        "pred": pred_shadow,
                    }
                )

    else:
        # Generic: compare all node detail properties
        for prop in [
            "layout_name",
            "node_count",
            "node_texts",
            "node_levels",
            "all_node_count",
            "all_node_texts",
            "all_node_levels",
            "all_node_types",
            "reverse",
            "color_name",
            "quick_style_name",
        ]:
            if gold_det.get(prop) != pred_det.get(prop):
                differences.append(
                    {
                        "property": prop,
                        "gold": gold_det.get(prop),
                        "pred": pred_det.get(prop),
                    }
                )

        is_insert_composition = any(
            step.get("operation") == "insert_smartart" for step in composition_steps(spec)
        )
        if is_insert_composition:
            gold_drawing = read_smartart_drawing_info(str(gold_path))
            pred_drawing = read_smartart_drawing_info(str(pred_path))
            saved_guardrail_checked = True
            # SmartArt responsively reflows the same layout when its
            # containing shape has a different aspect ratio (for example,
            # a Basic Block List may render as 2-2-1 or 3-2).  Absolute
            # cache coordinates and page reading order are therefore not
            # part of an insert instruction unless size/position is also
            # requested.  COM already verifies the semantic layout, node
            # hierarchy, style, and colors above; the saved cache only
            # guards against missing/stale visible node content here.
            gold_visible = sorted(str(text).strip() for text in gold_drawing.get("node_order", []))
            pred_visible = sorted(str(text).strip() for text in pred_drawing.get("node_order", []))
            if gold_visible != pred_visible:
                saved_differences.append(
                    {
                        "property": "drawing_node_texts",
                        "gold": gold_visible,
                        "pred": pred_visible,
                    }
                )

    if differences:
        return tag_verification_source(
            VerificationResult(
                passed=False,
                message=f"FAIL: SmartArt structural mismatch ({len(differences)} difference(s))",
                details={
                    "differences": differences,
                    "idMso": idMso,
                    "gold": gold_det,
                    "pred": pred_det,
                },
            ),
            "word_com",
        )

    if raw_cache_trigger is not None:
        refresh_result = families_smartart_refresh.verify_refreshed_smartart_projection(
            pred_path,
            gold_path,
            gold_det,
            pred_det,
            raw_cache_trigger,
            idMso,
            extract_refreshed_smartart_projection=extract_refreshed_smartart_projection,
            get_cleanup_word=get_cleanup_word,
        )
        if not refresh_result.passed:
            return refresh_result
        com_result = VerificationResult(
            passed=True,
            message="PASS: SmartArt structural properties match",
            details={"idMso": idMso},
        )
        return combine_com_primary_with_ooxml_guardrail(com_result, refresh_result)

    com_result = VerificationResult(
        passed=True, message="PASS: SmartArt structural properties match", details={"idMso": idMso}
    )
    if saved_guardrail_checked:
        saved_result = VerificationResult(
            passed=not saved_differences,
            message=(
                "PASS: Saved SmartArt state matches"
                if not saved_differences
                else "FAIL: Saved SmartArt state differs"
            ),
            details={
                "differences": saved_differences,
                "idMso": idMso,
                "match_mode": "saved_ooxml_smartart_state",
            },
        )
        return combine_com_primary_with_ooxml_guardrail(com_result, saved_result)
    return tag_verification_source(com_result, "word_com")
