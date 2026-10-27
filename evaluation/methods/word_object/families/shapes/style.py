"""families / shapes / style with explicit evidence dependencies."""

from evaluation.methods.base import (
    VerificationResult,
    combine_com_primary_with_ooxml_guardrail,
    tag_verification_source,
)
from evaluation.methods.word_object.core import DocumentOpenError
from evaluation.methods.word_object.observations import require_observations
from evaluation.methods.word_object.families.pictures import compare as families_pictures_compare
from evaluation.methods.word_object.families.shapes import (
    extract_saved as families_shapes_extract_saved,
)
from evaluation.methods.word_object.families.shapes import geometry as families_shapes_geometry
from evaluation.methods.word_object.families.shapes import layout as families_shapes_layout
from pathlib import Path
from typing import Any, Dict, List, Optional


def uses_named_shape_effect_projection(id_mso: str, spec: Optional[Dict]) -> bool:
    """Scope shape geometry projection to audited named presets only."""
    variation = str((spec or {}).get("variation") or "").casefold()
    approved = {
        "ObjectEffectShadowGallery": {"outer_center"},
        "ReflectionGallery": {"tight_touching"},
    }
    approved_variations = approved.get(id_mso, set())
    if variation in approved_variations:
        return True
    for intent in ((spec or {}).get("composition") or {}).get("atomic_intents") or []:
        if not isinstance(intent, dict):
            continue
        command, separator, intent_variation = str(intent.get("atomic_case") or "").partition("/")
        if (
            separator
            and command.casefold() == id_mso.casefold()
            and intent_variation.casefold() in approved_variations
        ):
            return True
    return False


def verify_shape_effects(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, read_shape_properties
) -> VerificationResult:
    """
    Verify shape effect properties match.

    Handles ObjectEffectShadowGallery, ObjectEffectGlowGallery,
    ObjectEffectSoftEdgesGallery, ObjectEffectPresetGallery,
    ReflectionGallery, GradientGallery, ShapeStylesGallery.
    """
    idMso = (spec or {}).get("idMso", "")
    if idMso == "ObjectEffectShadowGallery":
        try:
            gold_shapes = read_shape_properties(str(gold_path))
            pred_shapes = read_shape_properties(str(pred_path))
            pred_shapes, gold_shapes = families_shapes_geometry.match_shapes_by_name(
                pred_shapes, gold_shapes
            )
            com_differences = []
            if len(gold_shapes) != len(pred_shapes):
                com_differences.append(
                    {
                        "property": "shape_count",
                        "gold": len(gold_shapes),
                        "pred": len(pred_shapes),
                    }
                )
            if not com_differences and uses_named_shape_effect_projection(idMso, spec):
                return verify_named_shape_effect_state(
                    pred_path,
                    gold_path,
                    idMso,
                    pred_shapes,
                    gold_shapes,
                )
            for gold_shape, pred_shape in zip(gold_shapes, pred_shapes):
                if gold_shape.get("shadow_type") != pred_shape.get("shadow_type"):
                    com_differences.append(
                        {
                            "property": "shadow_type",
                            "shape_name": gold_shape.get("name"),
                            "gold": gold_shape.get("shadow_type"),
                            "pred": pred_shape.get("shadow_type"),
                        }
                    )
            if com_differences:
                return tag_verification_source(
                    VerificationResult(
                        passed=False,
                        message="FAIL: Word Object Model shadow properties differ",
                        details={"differences": com_differences},
                    ),
                    "word_com",
                )
            com_result = VerificationResult(
                passed=True,
                message="PASS: Word Object Model shadow properties match",
                details={"shape_count": len(gold_shapes)},
            )
        except DocumentOpenError:
            raise
        except Exception as com_error:
            gold_saved = families_shapes_extract_saved.extract_saved_shape_properties(gold_path)
            pred_saved = families_shapes_extract_saved.extract_saved_shape_properties(pred_path)
            gold_signatures = [shape.get("shadow_signature") for shape in gold_saved]
            pred_signatures = [shape.get("shadow_signature") for shape in pred_saved]
            fallback = VerificationResult(
                passed=gold_signatures == pred_signatures,
                message=(
                    "PASS: Saved shape shadow signature matches"
                    if gold_signatures == pred_signatures
                    else "FAIL: Saved shape shadow signature differs"
                ),
                details={
                    "match_mode": "saved_ooxml_shape_shadow",
                    "gold": gold_signatures,
                    "pred": pred_signatures,
                },
            )
            return tag_verification_source(
                fallback,
                "saved_ooxml_fallback",
                word_com_error=str(com_error),
            )

        gold_saved = families_shapes_extract_saved.extract_saved_shape_properties(gold_path)
        pred_saved = families_shapes_extract_saved.extract_saved_shape_properties(pred_path)
        gold_signatures = [shape.get("shadow_signature") for shape in gold_saved]
        pred_signatures = [shape.get("shadow_signature") for shape in pred_saved]
        saved_guardrail = VerificationResult(
            passed=gold_signatures == pred_signatures,
            message=(
                "PASS: Saved shape shadow signature matches"
                if gold_signatures == pred_signatures
                else "FAIL: Saved shape shadow signature differs"
            ),
            details={
                "match_mode": "saved_ooxml_shape_shadow",
                "gold": gold_signatures,
                "pred": pred_signatures,
            },
        )
        return combine_com_primary_with_ooxml_guardrail(com_result, saved_guardrail)

    if idMso == "GradientGallery":
        try:
            gold_shapes = read_shape_properties(str(gold_path))
        except DocumentOpenError:
            raise
        except Exception as exc:
            result = VerificationResult.error_result(
                "Gold GradientGallery evidence could not be extracted"
            )
            result.details.update(
                {
                    "error_owner": "gold_fixture",
                    "reason": "gold_gradient_extraction_failed",
                    "extraction_error": str(exc),
                }
            )
            return result

        if not gold_shapes:
            result = VerificationResult.error_result(
                "Gold gradient evidence contains no observable shape"
            )
            result.details.update(
                {
                    "error_owner": "gold_fixture",
                    "reason": "empty_gold_gradient_evidence",
                }
            )
            return result

        gold_evidence_issues = []
        for index, gold_shape in enumerate(gold_shapes, start=1):
            extraction_errors = [
                item
                for item in gold_shape.get("_extraction_errors", [])
                if item.get("property") == "fill_gradient_stops"
            ]
            missing = []
            stops = gold_shape.get("fill_gradient_stops")
            if not isinstance(stops, list) or len(stops) < 2:
                missing.append("fill_gradient_stops")
            if gold_shape.get("fill_gradient_has_visible_progression") is not True:
                missing.append("fill_gradient_has_visible_progression")
            if extraction_errors or missing:
                gold_evidence_issues.append(
                    {
                        "shape_index": index,
                        "shape_name": gold_shape.get("name"),
                        "missing_or_invalid": missing,
                        "extraction_errors": extraction_errors,
                    }
                )
        if gold_evidence_issues:
            result = VerificationResult.error_result(
                "Gold GradientGallery evidence is missing or unobservable"
            )
            result.details.update(
                {
                    "error_owner": "gold_fixture",
                    "reason": "invalid_gold_gradient_evidence",
                    "evidence_issues": gold_evidence_issues,
                }
            )
            return result

        try:
            pred_shapes = read_shape_properties(str(pred_path))
        except DocumentOpenError:
            raise
        except Exception as exc:
            return VerificationResult.error_result(
                "Predicted GradientGallery evidence could not be extracted",
                observation_owner="prediction",
                reason="prediction_gradient_extraction_failed",
                extraction_error=str(exc),
            )
    else:
        try:
            gold_shapes = read_shape_properties(str(gold_path))
            pred_shapes = read_shape_properties(str(pred_path))
        except DocumentOpenError:
            raise
        except Exception as com_error:
            if idMso != "ShapeStylesGallery":
                raise
            gold_saved = families_shapes_extract_saved.extract_saved_shape_style_signatures(
                gold_path
            )
            pred_saved = families_shapes_extract_saved.extract_saved_shape_style_signatures(
                pred_path
            )
            pred_saved, gold_saved = families_shapes_geometry.match_shapes_by_name(
                pred_saved, gold_saved
            )
            passed = bool(gold_saved) and gold_saved == pred_saved
            fallback = VerificationResult(
                passed=passed,
                message=(
                    "PASS: Saved composite shape style matches"
                    if passed
                    else "FAIL: Saved composite shape style differs"
                ),
                details={
                    "match_mode": "saved_ooxml_shape_style",
                    "fallback_contract": "complete",
                    "gold": gold_saved,
                    "pred": pred_saved,
                },
            )
            return tag_verification_source(
                fallback,
                "saved_ooxml_fallback",
                word_com_error=str(com_error),
            )

    if len(gold_shapes) != len(pred_shapes):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Shape count mismatch (gold={len(gold_shapes)}, pred={len(pred_shapes)})",
            details={},
        )

    if not gold_shapes:
        return VerificationResult(passed=False, message="FAIL: No shapes found", details={})

    if idMso == "ReflectionGallery" and uses_named_shape_effect_projection(idMso, spec):
        return verify_named_shape_effect_state(
            pred_path,
            gold_path,
            idMso,
            pred_shapes,
            gold_shapes,
        )

    gold_sorted = sorted(gold_shapes, key=lambda s: s.get("name", ""))
    pred_sorted = sorted(pred_shapes, key=lambda s: s.get("name", ""))

    # Compare only the state constrained by the requested command.  A
    # model may legitimately retain or choose different unrequested shape
    # defaults (for example fill/outline) while applying the right glow.
    # Composite gallery presets still use the broad fallback below.
    if idMso == "ObjectEffectGlowGallery":
        exact_props = []
        numeric_props = {"glow_radius": 1.0}
        color_props = ["glow_color_rgb"]
    elif idMso == "ObjectEffectSoftEdgesGallery":
        exact_props = []
        numeric_props = {"soft_edge_radius": 1.0}
        color_props = []
    elif idMso == "ReflectionGallery":
        exact_props = ["reflection_type"]
        numeric_props = {}
        color_props = []
    elif idMso == "GradientGallery":
        # The instruction constrains the Word gradient preset, not the
        # theme-resolved foreground RGB reported by Fill.ForeColor. The
        # latter can differ while the same gallery preset is retained.
        exact_props = [
            "fill_type",
            "fill_gradient_style",
            "fill_gradient_variant",
        ]
        numeric_props = {
            "fill_gradient_angle": 1.0,
            "fill_gradient_degree": 0.01,
        }
        color_props = []
    else:
        exact_props = [
            "shadow_type",
            "reflection_type",
            "fill_type",
            "auto_shape_type",
            "threed_bevel_top_type",
        ]
        numeric_props = {
            "glow_radius": 1.0,
            "soft_edge_radius": 1.0,
            "fill_gradient_angle": 1.0,
            "fill_transparency": 0.01,
            "threed_bevel_top_depth": 1.0,
        }
        color_props = [
            "fill_fore_color_rgb",
            "line_fore_color_rgb",
            "glow_color_rgb",
        ]

    differences = []

    for gold_shape, pred_shape in zip(gold_sorted, pred_sorted):
        shape_diffs = []
        # These named commands have explicit COM target properties. Broader
        # presets and the saved named-effect projections retain their own
        # active/inactive evidence contracts.
        if idMso in {"ObjectEffectGlowGallery", "ObjectEffectSoftEdgesGallery", "ReflectionGallery"}:
            error = require_observations(
                {"gold": gold_shape, "pred": pred_shape},
                [*exact_props, *numeric_props, *color_props], context=idMso,
            )
            if error is not None:
                return error

        if idMso == "GradientGallery":
            # Non-gradient fills establish a substantive defect even when
            # GradientStops is unavailable because it does not apply to them.
            # MsoFillType: 1 solid, 2 patterned, 4 textured, 5 background, 6 picture.
            if pred_shape.get("fill_type") in (1, 2, 4, 5, 6):
                return VerificationResult(
                    passed=False, message="FAIL: Predicted fill is not a gradient",
                    details={"reason": "wrong_fill_type", "property": "fill_type",
                             "pred": pred_shape["fill_type"], "expected": 3},
                )
            pred_extraction_errors = [
                item
                for item in pred_shape.get("_extraction_errors", [])
                if item.get("property") == "fill_gradient_stops"
            ]
            if pred_extraction_errors:
                return VerificationResult.error_result(
                    "Predicted gradient stops could not be extracted",
                    observation_owner="prediction",
                    reason="prediction_gradient_extraction_failed",
                    shape_name=pred_shape.get("name"),
                    extraction_errors=pred_extraction_errors,
                )
            pred_stops = pred_shape.get("fill_gradient_stops")
            if not isinstance(pred_stops, list) or len(pred_stops) < 2:
                shape_diffs.append(
                    {
                        "property": "fill_gradient_stops",
                        "gold": gold_shape.get("fill_gradient_stops"),
                        "pred": pred_stops,
                        "reason": "prediction gradient stop evidence is incomplete",
                    }
                )
            gold_progression = gold_shape.get("fill_gradient_has_visible_progression")
            pred_progression = pred_shape.get("fill_gradient_has_visible_progression")
            if gold_progression is not True:
                shape_diffs.append(
                    {
                        "property": "fill_gradient_has_visible_progression",
                        "gold": gold_progression,
                        "pred": pred_progression,
                        "reason": "reference gradient has no visible effective progression",
                    }
                )
            elif pred_progression is not True:
                shape_diffs.append(
                    {
                        "property": "fill_gradient_has_visible_progression",
                        "gold": True,
                        "pred": pred_progression,
                        "reason": "prediction is a gradient shell without visible progression",
                    }
                )

        # Exact match properties
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

        # Numeric properties with tolerance
        for prop, tol in numeric_props.items():
            gold_val = gold_shape.get(prop)
            pred_val = pred_shape.get(prop)
            if gold_val is None and pred_val is None:
                continue
            if gold_val is not None and pred_val is not None:
                if abs(gold_val - pred_val) > tol:
                    shape_diffs.append(
                        {
                            "property": prop,
                            "gold": gold_val,
                            "pred": pred_val,
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

        # Color properties (exact match)
        for prop in color_props:
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
            message=f"FAIL: {len(differences)} shape(s) have effect differences",
            details={"differences": differences},
        )

    return VerificationResult(
        passed=True,
        message="PASS: Shape effect properties match",
        details={"shape_count": len(gold_shapes)},
    )


def verify_shape_geometry(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, read_shape_properties
) -> VerificationResult:
    """
    Verify shape geometry properties match.

    Handles ShapeChangeShapeGallery, ArrowStyleGallery, OutlineDashesGallery,
    _3DRotationGallery, BevelShapeGallery.
    """
    idMso = spec.get("idMso", "") if spec else ""
    if idMso == "ShapeChangeShapeGallery":
        pred_shapes, gold_shapes, source, note = families_shapes_layout.extract_shape_pair(
            pred_path, gold_path, prefer_saved=False, read_shape_properties=read_shape_properties
        )
        if not gold_shapes:
            return VerificationResult(
                passed=False,
                message="FAIL: No shapes found in gold document",
                details={"match_mode": "saved_ooxml_autoshape_type"},
            )

        # None-safe sort: a document containing a non-autoshape alongside
        # the target yields auto_shape_type None, and sorted() over mixed
        # None/int raises TypeError.
        def _type_key(value):
            return (value is None, str(value))

        pred_types = sorted(
            (shape.get("auto_shape_type") for shape in pred_shapes),
            key=_type_key,
        )
        gold_types = sorted(
            (shape.get("auto_shape_type") for shape in gold_shapes),
            key=_type_key,
        )
        if pred_types != gold_types:
            return VerificationResult(
                passed=False,
                message="FAIL: Saved AutoShape type differs",
                details={
                    "match_mode": "saved_ooxml_autoshape_type",
                    "gold": gold_types,
                    "pred": pred_types,
                    "verification_source": source,
                    "extraction_note": note,
                },
            )
        return VerificationResult(
            passed=True,
            message="PASS: Saved AutoShape type matches",
            details={
                "match_mode": "saved_ooxml_autoshape_type",
                "signature": gold_types,
                "verification_source": source,
                "extraction_note": note,
            },
        )

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

    gold_sorted = sorted(gold_shapes, key=lambda s: s.get("name", ""))
    pred_sorted = sorted(pred_shapes, key=lambda s: s.get("name", ""))

    idMso = spec.get("idMso", "") if spec else ""

    # Determine which properties to compare based on command
    exact_props = []
    numeric_props = {}

    # ShapeChangeShapeGallery never reaches here: it returns from the
    # dedicated saved-autoshape branch above.
    if idMso == "ArrowStyleGallery":
        exact_props = ["line_begin_arrowhead_style", "line_end_arrowhead_style"]
    elif idMso == "OutlineDashesGallery":
        exact_props = ["line_dash_style"]
    elif idMso == "OutlineWeightGallery":
        numeric_props = {"line_weight": 0.5}
    elif idMso == "_3DRotationGallery":
        exact_props = ["threed_preset_camera"]
        numeric_props = {
            "threed_rotation_x": 1.0,
            "threed_rotation_y": 1.0,
        }
    elif idMso == "BevelShapeGallery":
        exact_props = ["threed_bevel_top_type"]
    else:
        # Fallback: compare all geometry props
        exact_props = [
            "auto_shape_type",
            "line_dash_style",
            "line_begin_arrowhead_style",
            "line_end_arrowhead_style",
            "threed_bevel_top_type",
            "threed_preset_camera",
        ]
        numeric_props = {"line_weight": 0.5}

    differences = []

    for gold_shape, pred_shape in zip(gold_sorted, pred_sorted):
        shape_diffs = []

        if idMso in {"ArrowStyleGallery", "OutlineDashesGallery", "OutlineWeightGallery",
                     "_3DRotationGallery", "BevelShapeGallery"}:
            error = require_observations(
                {"gold": gold_shape, "pred": pred_shape},
                [*exact_props, *numeric_props], context=idMso,
            )
            if error is not None:
                return error

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

        for prop, tol in numeric_props.items():
            gold_val = gold_shape.get(prop)
            pred_val = pred_shape.get(prop)
            if gold_val is None and pred_val is None:
                continue
            if gold_val is not None and pred_val is not None:
                if abs(gold_val - pred_val) > tol:
                    shape_diffs.append(
                        {
                            "property": prop,
                            "gold": gold_val,
                            "pred": pred_val,
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

    if idMso == "BevelShapeGallery":
        try:
            saved_failures = families_shapes_geometry.bevel_render_failures(pred_path, gold_path)
        except Exception as exc:
            saved_failures = [
                {
                    "property": "bevel_saved_structure",
                    "error": str(exc),
                }
            ]
        if saved_failures:
            differences.append(
                {
                    "shape_name": "saved DrawingML bevel",
                    "differences": saved_failures,
                }
            )

    if differences:
        return tag_verification_source(
            VerificationResult(
                passed=False,
                message=f"FAIL: {len(differences)} shape(s) have geometry differences",
                details={"differences": differences, "idMso": idMso},
            ),
            ("word_com+saved_ooxml_guardrail" if idMso == "BevelShapeGallery" else "word_com"),
        )

    return tag_verification_source(
        VerificationResult(
            passed=True,
            message="PASS: Shape geometry properties match",
            details={"shape_count": len(gold_shapes), "idMso": idMso},
        ),
        ("word_com+saved_ooxml_guardrail" if idMso == "BevelShapeGallery" else "word_com"),
    )


def verify_named_shape_effect_state(
    pred_path: Path,
    gold_path: Path,
    id_mso: str,
    pred_shapes: List[Dict[str, Any]],
    gold_shapes: List[Dict[str, Any]],
) -> VerificationResult:
    """Verify audited named shape effects by effective saved geometry."""
    family = "outerShdw" if id_mso == "ObjectEffectShadowGallery" else "reflection"
    com_property = "shadow_type" if family == "outerShdw" else "reflection_type"
    try:
        gold_effects = families_shapes_extract_saved.extract_saved_shape_named_effects(gold_path)
        pred_effects = families_shapes_extract_saved.extract_saved_shape_named_effects(pred_path)
    except Exception as exc:
        return VerificationResult.error_result(
            f"Named shape-effect saved evidence extraction failed: {exc}"
        )

    pred_shapes, gold_shapes = families_shapes_geometry.match_shapes_by_name(
        pred_shapes, gold_shapes
    )
    gold_names = [shape.get("name") for shape in gold_shapes]
    pred_names = [shape.get("name") for shape in pred_shapes]
    if (
        not gold_names
        or None in gold_names
        or None in pred_names
        or len(set(gold_names)) != len(gold_names)
        or gold_names != pred_names
    ):
        return VerificationResult(
            passed=False,
            message="FAIL: Named shape identities differ",
            details={
                "match_mode": "named_shape_effect_semantic_projection",
                "gold_names": gold_names,
                "pred_names": pred_names,
            },
        )

    def integer(attributes, name, default=None):
        value = attributes.get(name)
        if value is None:
            return default
        try:
            return int(value)
        except (TypeError, ValueError):
            return None

    failures = []
    evidence = []
    for gold_shape, pred_shape in zip(gold_shapes, pred_shapes):
        name = gold_shape["name"]
        gold_family = gold_effects.get(name, {}).get(family, [])
        pred_family = pred_effects.get(name, {}).get(family, [])
        if len(gold_family) != 1:
            return VerificationResult.error_result(
                f"Gold shape {name!r} must contain exactly one saved " f"{family} effect"
            )
        if len(pred_family) != 1:
            failures.append(
                f"pred shape {name!r} must contain exactly one saved " f"{family} effect"
            )
            evidence.append(
                {
                    "shape_name": name,
                    "gold_saved": gold_family,
                    "pred_saved": pred_family,
                }
            )
            continue

        gold_effect = gold_family[0]
        pred_effect = pred_family[0]
        gold_attributes = gold_effect["attributes"]
        pred_attributes = pred_effect["attributes"]
        shape_failures = []
        if family == "outerShdw":
            for owner, effect in (("gold", gold_effect), ("pred", pred_effect)):
                attributes = effect["attributes"]
                blur = integer(attributes, "blurRad")
                try:
                    alpha = int(effect.get("alpha"))
                except (TypeError, ValueError):
                    alpha = None
                black = (
                    effect.get("color_type") == "prstClr"
                    and (effect.get("color_value") or "").casefold() == "black"
                    or effect.get("color_type") == "srgbClr"
                    and (effect.get("color_value") or "").upper() == "000000"
                )
                if (
                    blur is None
                    or blur <= 0
                    or not families_pictures_compare.outer_shadow_has_centered_zero_offset(
                        attributes
                    )
                    or not black
                    or alpha is None
                    or not 0 < alpha <= 100000
                ):
                    shape_failures.append(
                        f"{owner} shadow is not centered, zero-distance, "
                        "blurred black geometry with alpha"
                    )
        else:
            required = ("blurRad", "stA", "endA", "endPos", "dir", "sy", "algn")
            for owner, attributes in (("gold", gold_attributes), ("pred", pred_attributes)):
                missing = [
                    property_name for property_name in required if property_name not in attributes
                ]
                if missing or integer(attributes, "dist", 0) != 0:
                    shape_failures.append(
                        f"{owner} reflection geometry is incomplete or " "nonzero-distance"
                    )
            if not shape_failures:
                for property_name in ("blurRad", "endA", "endPos", "dir", "sy"):
                    if integer(pred_attributes, property_name) != integer(
                        gold_attributes, property_name
                    ):
                        shape_failures.append(f"reflection {property_name} differs")
                if pred_attributes.get("algn") != gold_attributes.get("algn"):
                    shape_failures.append("reflection alignment differs")
                if abs(integer(pred_attributes, "stA") - integer(gold_attributes, "stA")) > 3000:
                    shape_failures.append("reflection start opacity differs")
                if pred_attributes.get("rotWithShape", "0") != gold_attributes.get(
                    "rotWithShape", "0"
                ):
                    shape_failures.append("reflection rotation semantics differ")

        gold_com = gold_shape.get(com_property)
        pred_com = pred_shape.get(com_property)
        if pred_com != gold_com and pred_com != -2:
            shape_failures.append(f"COM {com_property} pred={pred_com}, gold={gold_com}")
        failures.extend(f"{name}: {failure}" for failure in shape_failures)
        evidence.append(
            {
                "shape_name": name,
                "gold_saved": gold_effect,
                "pred_saved": pred_effect,
                "gold_com": gold_com,
                "pred_com": pred_com,
            }
        )

    result = VerificationResult(
        passed=not failures,
        message=(
            "PASS: Named shape effect semantic state matches"
            if not failures
            else "FAIL: Named shape effect semantic state differs"
        ),
        details={
            "match_mode": "named_shape_effect_semantic_projection",
            "family": family,
            "failures": failures,
            "shapes": evidence,
        },
    )
    return tag_verification_source(result, "word_com+saved_ooxml_guardrail")
