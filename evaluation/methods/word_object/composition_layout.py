"""composition_layout with explicit evidence dependencies."""

from evaluation.methods.word_object.effect_commands import (
    ALIGNMENT_COLLECTION_METRICS,
    DISTRIBUTION_COLLECTION_METRICS,
)
from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object import composition_effects as composition_effects
from evaluation.methods.word_object.composition import composition_command_ids
from evaluation.methods.word_object.core import DocumentOpenError
from pathlib import Path
from typing import Dict, Optional


def uses_tight_wrap_polygon_projection(spec: Optional[Dict]) -> bool:
    """Return whether the case explicitly owns tight-wrap contour semantics."""
    variation = str((spec or {}).get("variation") or "").casefold()
    return variation == "apply_tight" or "tight" in variation


def relative_only_layout_composition(spec: Optional[Dict]) -> bool:
    """Return whether layout intent fixes relations but no absolute frame."""
    commands = set(composition_command_ids(spec))
    relative_commands = set(ALIGNMENT_COLLECTION_METRICS) | set(DISTRIBUTION_COLLECTION_METRICS)
    if not commands or not commands.issubset(relative_commands):
        return False
    explicit_position_keys = {
        "left",
        "top",
        "x",
        "y",
        "position",
        "placement",
        "relative_h_position",
        "relative_v_position",
    }
    for step in composition_effects.composition_steps(spec):
        id_mso = step.get("idMso")
        operation = (step.get("operation") or "").casefold()
        if id_mso and id_mso not in relative_commands:
            return False
        if "insert" in operation or "position" in operation:
            return False
        params = step.get("params") or {}
        if explicit_position_keys & set(params):
            return False
    return True


def verify_multi_shape_composition(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    calculate_actual_position,
    extract_shape_pair,
    match_shapes_by_name,
    wrap_polygon_failures,
) -> VerificationResult:
    """
    Verify layout-coupled compositions that need multiple property checks.

    Chains verification checks based on interaction_class:
    - align_then_distribute / distribute_then_align: position comparison
    - rotate_then_align: rotation + position comparison
    - group_lifecycle_zorder: z-order + group-absence check
    - group_lifecycle_rotate: rotation + group-absence check
    - wrap_then_position: wrap type + position comparison
    - crop_then_effect: crop + effect properties
    """
    interaction_class = spec.get("composition", {}).get("interaction_class", "")

    # Most layout-coupled cases just need full shape property comparison
    # against gold -- positions, rotation, z-order are all captured.
    try:
        pred_shapes, gold_shapes, extraction_source, extraction_note = extract_shape_pair(
            pred_path, gold_path, prefer_saved=False
        )
    except DocumentOpenError:
        raise
    except Exception as e:
        return VerificationResult.error_result(
            f"Multi-shape composition extraction error: {e}",
            observation_owner="unknown", reason="shape_extraction_failed",
            interaction_class=interaction_class,
        )

    if not gold_shapes:
        return VerificationResult.error_result(
            "Gold multi-shape composition target contains no shapes",
            observation_owner="gold", reason="missing_gold_shape_target",
        )
    if len(pred_shapes) != len(gold_shapes):
        return VerificationResult(
            passed=False,
            message=(
                f"FAIL: Shape count mismatch " f"(gold={len(gold_shapes)}, pred={len(pred_shapes)})"
            ),
            details={
                "interaction_class": interaction_class,
                "pred_count": len(pred_shapes),
                "gold_count": len(gold_shapes),
                "verification_source": (extraction_source),
                "extraction_note": extraction_note,
            },
        )

    pred_shapes, gold_shapes = match_shapes_by_name(pred_shapes, gold_shapes)

    failures = []

    tight_wrap_projection = (
        interaction_class == "wrap_then_position" and uses_tight_wrap_polygon_projection(spec)
    )
    if tight_wrap_projection:
        try:
            failures.extend(wrap_polygon_failures(pred_path, gold_path))
        except Exception as exc:
            failures.append(f"wrap polygon extraction error: {exc}")

    verification_source = extraction_source
    if tight_wrap_projection and extraction_source == "word_com":
        verification_source = "word_com+saved_ooxml_guardrail"

    # Check that no shapes remain grouped (group lifecycle cases)
    if "group_lifecycle" in interaction_class:
        for i, ps in enumerate(pred_shapes):
            if ps.get("is_group"):
                failures.append(f"shape {i+1} is still grouped (expected ungrouped)")

    # Compare properties with tolerances.
    # Use _calculate_actual_position to resolve Word's preset sentinel
    # constants (e.g. -999995 = wdShapeCenter) to actual page
    # coordinates.  This matches the approach in _verify_position
    # (used by the atomic PicturePositionGallery tests) and ensures a
    # prediction that computes the absolute position manually passes
    # when it matches a gold that stored the preset sentinel.
    position_tol = 2.0
    angle_tol = 1.0  # degrees

    # Alignment and distribution instructions constrain relationships,
    # not a unique absolute placement. Compare translation-invariant
    # residuals and skip gold-coordinate equality on the affected axes.
    commands = set(composition_command_ids(spec))
    relative_only_layout = relative_only_layout_composition(spec)
    targeted_position_axes = set()
    for command in commands:
        metric_info = ALIGNMENT_COLLECTION_METRICS.get(command)
        if metric_info:
            axis, metric = metric_info
            targeted_position_axes.add(axis)
            residual = composition_effects.custom_composition_collection_value(
                pred_shapes, metric, calculate_actual_position=calculate_actual_position
            )
            if residual > position_tol:
                failures.append(f"{command} semantic residual={residual:.2f}pt")
            if not relative_only_layout:
                gold_anchor = composition_effects.alignment_anchor(
                    gold_shapes, metric, calculate_actual_position=calculate_actual_position
                )
                pred_anchor = composition_effects.alignment_anchor(
                    pred_shapes, metric, calculate_actual_position=calculate_actual_position
                )
                if (
                    gold_anchor is None
                    or pred_anchor is None
                    or abs(pred_anchor - gold_anchor) > position_tol
                ):
                    failures.append(
                        f"{command} alignment anchor: " f"pred={pred_anchor}, gold={gold_anchor}"
                    )
        metric_info = DISTRIBUTION_COLLECTION_METRICS.get(command)
        if metric_info:
            axis, metric = metric_info
            targeted_position_axes.add(axis)
            if len(pred_shapes) < 3:
                failures.append(f"{command} requires at least 3 shapes")
            else:
                residual = composition_effects.custom_composition_collection_value(
                    pred_shapes, metric, calculate_actual_position=calculate_actual_position
                )
                if residual > position_tol:
                    failures.append(f"{command} semantic residual={residual:.2f}pt")
                if not relative_only_layout:
                    gold_endpoints = composition_effects.distribution_center_endpoints(
                        gold_shapes, axis, calculate_actual_position=calculate_actual_position
                    )
                    pred_endpoints = composition_effects.distribution_center_endpoints(
                        pred_shapes, axis, calculate_actual_position=calculate_actual_position
                    )
                    if (
                        len(gold_endpoints) != 2
                        or len(pred_endpoints) != 2
                        or any(
                            abs(pred - gold) > position_tol
                            for pred, gold in zip(pred_endpoints, gold_endpoints)
                        )
                    ):
                        failures.append(
                            f"{command} distribution endpoints: "
                            f"pred={pred_endpoints}, gold={gold_endpoints}"
                        )

    # Page dimensions (same for all shapes in a single-page doc)
    page_w = gold_shapes[0].get("page_width", 612)
    page_h = gold_shapes[0].get("page_height", 792)
    margin_l = gold_shapes[0].get("left_margin", 72)
    margin_t = gold_shapes[0].get("top_margin", 72)
    margin_r = gold_shapes[0].get("right_margin", 72)
    margin_b = gold_shapes[0].get("bottom_margin", 72)

    if interaction_class == "group_flip_ungroup_distribute":
        gold_endpoints = composition_effects.horizontal_center_endpoints(
            gold_shapes, calculate_actual_position=calculate_actual_position
        )
        pred_endpoints = composition_effects.horizontal_center_endpoints(
            pred_shapes, calculate_actual_position=calculate_actual_position
        )
        if (
            len(gold_endpoints) != 2
            or len(pred_endpoints) != 2
            or any(
                abs(pred - gold) > position_tol
                for pred, gold in zip(pred_endpoints, gold_endpoints)
            )
        ):
            failures.append(
                "horizontal distribution endpoints: "
                f"pred={pred_endpoints}, gold={gold_endpoints}"
            )

    for i, (ps, gs) in enumerate(zip(pred_shapes, gold_shapes)):
        shape_label = f"shape {i+1}"

        # Resolve left/top through _calculate_actual_position so
        # sentinels and absolute values are comparable.
        for raw_key, size_key, pg_size, m_start, m_end in [
            ("left", "width", page_w, margin_l, margin_r),
            ("top", "height", page_h, margin_t, margin_b),
        ]:
            axis = "horizontal" if raw_key == "left" else "vertical"
            if axis in targeted_position_axes:
                continue
            pv_raw = ps.get(raw_key)
            gv_raw = gs.get(raw_key)
            if pv_raw is None or gv_raw is None:
                continue
            pv = calculate_actual_position(
                pv_raw,
                ps.get(size_key, 0),
                ps.get("relative_h_position" if raw_key == "left" else "relative_v_position", 0),
                pg_size,
                m_start,
                m_end,
            )
            gv = calculate_actual_position(
                gv_raw,
                gs.get(size_key, 0),
                gs.get("relative_h_position" if raw_key == "left" else "relative_v_position", 0),
                pg_size,
                m_start,
                m_end,
            )
            if abs(pv - gv) > position_tol:
                failures.append(f"{shape_label} {raw_key}: " f"pred={pv:.1f}, gold={gv:.1f}")

        # Width and height (no sentinel resolution needed)
        for key in ("width", "height"):
            pv = ps.get(key)
            gv = gs.get(key)
            if pv is not None and gv is not None:
                if abs(pv - gv) > position_tol:
                    failures.append(f"{shape_label} {key}: " f"pred={pv:.1f}, gold={gv:.1f}")
        # Rotation
        pr = ps.get("rotation", 0)
        gr = gs.get("rotation", 0)
        angle_difference = abs((pr - gr + 180.0) % 360.0 - 180.0)
        if angle_difference > angle_tol:
            failures.append(f"{shape_label} rotation: pred={pr:.1f}, gold={gr:.1f}")
        # Z-order
        if ps.get("z_order") != gs.get("z_order"):
            failures.append(
                f"{shape_label} z_order: " f'pred={ps.get("z_order")}, ' f'gold={gs.get("z_order")}'
            )
        # Wrap type is semantic only when the composition requests wrap.
        # Position presets can change wrap as a Word side effect, but the
        # position-only atomic verifier deliberately ignores that default.
        if "wrap" in interaction_class and ps.get("wrap_type") != gs.get("wrap_type"):
            pw = ps.get("wrap_type")
            gw = gs.get("wrap_type")
            if pw is not None or gw is not None:
                failures.append(f"{shape_label} wrap_type: pred={pw}, gold={gw}")
        # Relative position frame (for wrap/position presets like
        # wdShapeCenter, wdShapeTop relative to margin)
        for key in ("relative_h_position", "relative_v_position"):
            axis = "horizontal" if key == "relative_h_position" else "vertical"
            if axis in targeted_position_axes:
                continue
            pv = ps.get(key)
            gv = gs.get(key)
            if pv is not None and gv is not None and pv != gv:
                failures.append(f"{shape_label} {key}: pred={pv}, gold={gv}")
        # Flip state (for group-flip-ungroup cases; position changes
        # catch flips of symmetric shapes, but asymmetric shapes need
        # explicit flip checks)
        for key in ("flip_horizontal", "flip_vertical"):
            pv = ps.get(key)
            gv = gs.get(key)
            if pv is not None and gv is not None and pv != gv:
                failures.append(f"{shape_label} {key}: pred={pv}, gold={gv}")

    if failures:
        return VerificationResult(
            passed=False,
            message=(f"FAIL: Multi-shape composition mismatch " f"({len(failures)} difference(s))"),
            details={
                "failures": failures,
                "interaction_class": interaction_class,
                "gold_shapes": gold_shapes,
                "pred_shapes": pred_shapes,
                "relative_only_layout": relative_only_layout,
                "verification_source": verification_source,
                "extraction_note": extraction_note,
            },
        )

    return VerificationResult(
        passed=True,
        message=(
            f"PASS: Multi-shape composition properties match " f"({len(pred_shapes)} shape(s))"
        ),
        details={
            "interaction_class": interaction_class,
            "shape_count": len(pred_shapes),
            "relative_only_layout": relative_only_layout,
            "verification_source": verification_source,
            **(
                {"fallback_contract": "complete"}
                if (
                    interaction_class == "wrap_then_position"
                    and extraction_source == "saved_ooxml_fallback"
                )
                else {}
            ),
            "extraction_note": extraction_note,
        },
    )
