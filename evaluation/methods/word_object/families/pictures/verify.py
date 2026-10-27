"""families / pictures / verify with explicit evidence dependencies."""

from evaluation.methods.word_object.effect_commands import (
    POSITIVE_PICTURE_EFFECT_COMMANDS,
    POSITIVE_PICTURE_EFFECT_PARAM_TYPES,
)
from evaluation.methods.base import (
    VerificationResult,
    combine_com_primary_with_ooxml_guardrail,
    tag_verification_source,
)
from evaluation.methods.word_object.core import DocumentOpenError
from evaluation.methods.word_object.families.pictures import compare as families_pictures_compare
from evaluation.methods.word_object.families.pictures import (
    extract_saved as families_pictures_extract_saved,
)
from evaluation.methods.word_object.families.shapes import (
    extract_saved as families_shapes_extract_saved,
)
from evaluation.methods.word_object.families.shapes import geometry as families_shapes_geometry
from pathlib import Path
from ...observations import require_observations
from .camera_xml import CAMERA_PROPERTIES
from typing import Any, Dict, Optional


def verify_picture_smartart(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    get_position_tolerance,
    read_inline_shapes,
) -> VerificationResult:
    """Verify picture-to-SmartArt conversion on the first inline shape."""
    gold_info = read_inline_shapes(str(gold_path))
    pred_info = read_inline_shapes(str(pred_path))

    if gold_info["inline_shape_count"] != pred_info["inline_shape_count"]:
        return VerificationResult(
            passed=False,
            message=(
                f"FAIL: Inline shape count mismatch "
                f"(gold={gold_info['inline_shape_count']}, pred={pred_info['inline_shape_count']})"
            ),
            details={"gold": gold_info, "pred": pred_info},
        )

    if not gold_info["inline_shapes"] or not pred_info["inline_shapes"]:
        return VerificationResult(
            passed=False,
            message="FAIL: Missing inline shape after picture-to-SmartArt conversion",
            details={"gold": gold_info, "pred": pred_info},
        )

    gold_shape = gold_info["inline_shapes"][0]
    pred_shape = pred_info["inline_shapes"][0]

    differences = []
    for key in [
        "type",
        "has_smartart",
        "smartart_layout_name",
        "smartart_color_name",
        "smartart_quick_style_name",
        "smartart_node_count",
    ]:
        if gold_shape.get(key) != pred_shape.get(key):
            differences.append(
                {
                    "property": key,
                    "gold": gold_shape.get(key),
                    "pred": pred_shape.get(key),
                }
            )

    for key in ["width", "height"]:
        gold_val = gold_shape.get(key)
        pred_val = pred_shape.get(key)
        if gold_val is None or pred_val is None:
            continue
        if abs(gold_val - pred_val) > get_position_tolerance():
            differences.append(
                {
                    "property": key,
                    "gold": gold_val,
                    "pred": pred_val,
                }
            )

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: {len(differences)} picture-to-SmartArt difference(s)",
            details={
                "differences": differences,
                "gold_shape": gold_shape,
                "pred_shape": pred_shape,
            },
        )

    return VerificationResult(
        passed=True,
        message="PASS: Picture-to-SmartArt conversion matches gold",
        details={
            "gold_shape": gold_shape,
            "pred_shape": pred_shape,
        },
    )


def verify_picture_crop(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, read_picture_crop_properties
) -> VerificationResult:
    """
    Verify inline picture crop properties match between pred and gold.

    Handles PictureFillCrop and PictureFitCrop commands.
    These commands operate on inline pictures (not floating shapes).
    """
    gold_crop = read_picture_crop_properties(str(gold_path))
    pred_crop = read_picture_crop_properties(str(pred_path))

    # Check for extraction errors
    if "error" in gold_crop:
        return VerificationResult.error_result(f"Gold: {gold_crop['error']}")
    if "error" in pred_crop:
        return families_pictures_compare.picture_extraction_failure(
            pred_crop["error"],
            idMso=spec.get("idMso", "") if spec else "",
        )

    # Use larger tolerance for picture crop - different approaches can produce
    error = require_observations(
        {"gold": gold_crop, "pred": pred_crop},
        ("crop_left", "crop_right", "crop_top", "crop_bottom", "width", "height"),
        context="picture crop",
    )
    if error is not None:
        return error

    # slightly different crop values while achieving the same visual result
    tolerance = 10.0  # 10 points tolerance for crop values
    differences = []

    crop_props = ["crop_left", "crop_right", "crop_top", "crop_bottom"]
    for prop in crop_props:
        gold_val = gold_crop.get(prop, 0)
        pred_val = pred_crop.get(prop, 0)
        if abs(gold_val - pred_val) > tolerance:
            differences.append(
                {
                    "property": prop,
                    "gold": gold_val,
                    "pred": pred_val,
                    "diff": round(abs(gold_val - pred_val), 2),
                }
            )

    # Also check dimensions (width/height should match)
    for prop in ["width", "height"]:
        gold_val = gold_crop.get(prop, 0)
        pred_val = pred_crop.get(prop, 0)
        if abs(gold_val - pred_val) > tolerance:
            differences.append(
                {
                    "property": prop,
                    "gold": gold_val,
                    "pred": pred_val,
                    "diff": round(abs(gold_val - pred_val), 2),
                }
            )

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Picture crop mismatch ({len(differences)} difference(s))",
            details={"differences": differences, "gold_crop": gold_crop, "pred_crop": pred_crop},
        )

    return VerificationResult(
        passed=True,
        message="PASS: Picture crop properties match",
        details={"gold_crop": gold_crop, "pred_crop": pred_crop},
    )


def verify_picture_format(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, read_picture_format_properties
) -> VerificationResult:
    """Verify picture format properties (brightness, contrast, crop, dimensions)."""
    gold_fmt = read_picture_format_properties(str(gold_path))
    pred_fmt = read_picture_format_properties(str(pred_path))

    if "error" in gold_fmt:
        return VerificationResult.error_result(f"Gold: {gold_fmt['error']}")
    if "error" in pred_fmt:
        return families_pictures_compare.picture_extraction_failure(
            pred_fmt["error"],
            idMso=spec.get("idMso", "") if spec else "",
        )

    error = require_observations(
        {"gold": gold_fmt, "pred": pred_fmt},
        ("brightness", "contrast", "crop_left", "crop_right", "crop_top", "crop_bottom", "width", "height"),
        context="picture format",
    )
    if error is not None:
        return error
    differences = []
    # Brightness/contrast: small tolerance for floating point
    for prop in ["brightness", "contrast"]:
        gold_val = gold_fmt.get(prop, 0)
        pred_val = pred_fmt.get(prop, 0)
        if abs(gold_val - pred_val) > 0.01:
            differences.append({"property": prop, "gold": gold_val, "pred": pred_val})

    # Crop and dimensions: 10pt tolerance
    for prop in ["crop_left", "crop_right", "crop_top", "crop_bottom", "width", "height"]:
        gold_val = gold_fmt.get(prop, 0)
        pred_val = pred_fmt.get(prop, 0)
        if abs(gold_val - pred_val) > 10.0:
            differences.append({"property": prop, "gold": gold_val, "pred": pred_val})

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Picture format mismatch ({len(differences)} difference(s))",
            details={"differences": differences, "gold": gold_fmt, "pred": pred_fmt},
        )

    return VerificationResult(
        passed=True,
        message="PASS: Picture format properties match",
        details={"gold": gold_fmt, "pred": pred_fmt},
    )


def verify_picture_effects(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    read_picture_effect_properties,
    saved_evidence_error,
) -> VerificationResult:
    """
    Verify picture effect properties match.

    Handles PictureArtisticEffectsGallery, PictureEffects*Gallery,
    PictureRecolorGallery, PictureShapeGallery, PictureSharpenSoftenGallery,
    _3DBevelPictureTopGallery, PictureRotationGallery, etc.
    """
    idMso = spec.get("idMso", "") if spec else ""
    allow_floating = bool((spec or {}).get("_allow_floating_picture"))
    positive_command = idMso in POSITIVE_PICTURE_EFFECT_COMMANDS

    def extract(path: Path) -> Dict:
        if allow_floating:
            return read_picture_effect_properties(str(path), allow_floating=True)
        return read_picture_effect_properties(str(path))

    if positive_command:
        try:
            gold_props = extract(gold_path)
        except DocumentOpenError:
            raise
        except Exception as exc:
            result = VerificationResult.error_result(
                "Gold positive picture-effect evidence could not be extracted"
            )
            result.details.update(
                {
                    "error_owner": "gold_fixture",
                    "reason": "gold_picture_effect_extraction_failed",
                    "idMso": idMso,
                    "extraction_error": str(exc),
                }
            )
            return result

        if "error" in gold_props:
            result = VerificationResult.error_result(f"Gold: {gold_props['error']}")
            result.details.update(
                {
                    "error_owner": "gold_fixture",
                    "reason": "gold_picture_effect_extraction_failed",
                    "idMso": idMso,
                }
            )
            return result

        # Validate required gold evidence before touching the prediction so
        # a prediction-side failure cannot mask a defective fixture.
        if idMso == "PictureArtisticEffectsGallery":
            gold_evidence = families_pictures_compare.verify_picture_effects_param(
                gold_props,
                gold_props,
                idMso,
                compare_effect_type=True,
            )
        else:
            gold_evidence = families_pictures_compare.verify_picture_effects_param(
                gold_props,
                gold_props,
                idMso,
                effect_type=POSITIVE_PICTURE_EFFECT_PARAM_TYPES[idMso],
            )
        if not gold_evidence.passed:
            return gold_evidence

        try:
            pred_props = extract(pred_path)
        except DocumentOpenError:
            raise
        except Exception as exc:
            return VerificationResult.error_result(
                "Predicted positive picture-effect evidence could not be extracted",
                observation_owner="prediction",
                reason="prediction_picture_effect_extraction_failed",
                idMso=idMso,
                extraction_error=str(exc),
            )

        if "error" in pred_props:
            result = families_pictures_compare.picture_extraction_failure(
                pred_props["error"], idMso=idMso,
            )
            result.details["observation_owner"] = "prediction"
            if result.error:
                result.details["reason"] = "prediction_picture_effect_extraction_failed"
            return result
    else:
        # Camera recovery validates the reference first. Other routes retain
        # their existing order of extracting both sides before interpreting errors.
        gold_props = extract(gold_path)
        if idMso == "PictureRotationGallery":
            if "error" in gold_props:
                return VerificationResult.error_result(f"Gold: {gold_props['error']}")
            error = require_observations({"gold": gold_props}, CAMERA_PROPERTIES, context=idMso)
            if error is not None:
                return error
        pred_props = extract(pred_path)
        if "error" in gold_props:
            return VerificationResult.error_result(f"Gold: {gold_props['error']}")
        if "error" in pred_props:
            return families_pictures_compare.picture_extraction_failure(
                pred_props["error"],
                idMso=idMso,
            )

    if not allow_floating and families_pictures_compare.uses_named_picture_effect_projection(
        idMso, spec
    ):
        return verify_named_picture_effect_state(
            pred_path, gold_path, idMso, pred_props, gold_props
        )

    camera_fallback = {
        owner: props["_camera_xml_fallback"]
        for owner, props in (("gold", gold_props), ("pred", pred_props))
        if "_camera_xml_fallback" in props
    }

    # Determine which properties to compare based on command
    exact_props = []
    numeric_props = {}

    if idMso == "PictureEffectsShadowGallery":
        exact_props = ["shadow_type"]
    elif idMso == "PictureEffectsGlowGallery":
        numeric_props = {"glow_radius": 1.0}
        exact_props = ["glow_color_rgb"]
    elif idMso == "PictureEffectsSoftEdgesGallery":
        numeric_props = {"soft_edge_radius": 1.0}
    elif idMso == "PictureReflectionGallery":
        exact_props = ["reflection_type"]
    elif idMso == "PictureRotationGallery":
        exact_props = ["threed_preset_camera"]
        numeric_props = {"threed_rotation_x": 1.0, "threed_rotation_y": 1.0}
    elif idMso == "_3DBevelPictureTopGallery":
        exact_props = ["threed_bevel_top_type"]
        numeric_props = {"threed_bevel_top_depth": 1.0}
    elif idMso == "PictureShapeGallery":
        exact_props = ["auto_shape_type"]
    elif idMso == "PictureSharpenSoftenGallery":
        com_result = families_pictures_compare.verify_picture_effects_param(
            gold_props, pred_props, idMso, effect_type=25
        )
        if not com_result.passed:
            return tag_verification_source(com_result, "word_com")
        return combine_com_primary_with_ooxml_guardrail(com_result,
            verify_picture_effect_persistence(pred_path, gold_path, idMso,
                                              saved_evidence_error=saved_evidence_error))
    elif idMso in ("PictureRecolorGallery", "RecolorColorPicker"):
        exact_props = ["color_type"]
    elif idMso == "PictureColorTempertatureGallery":
        com_result = families_pictures_compare.verify_picture_effects_param(
            gold_props, pred_props, idMso, effect_type=7
        )
        if not com_result.passed:
            return tag_verification_source(com_result, "word_com")
        return combine_com_primary_with_ooxml_guardrail(com_result,
            verify_picture_effect_persistence(pred_path, gold_path, idMso,
                                              saved_evidence_error=saved_evidence_error))
    elif idMso == "PictureSaturationGallery":
        # Compare via PictureEffects (type=24) parameter value
        com_result = families_pictures_compare.verify_picture_effects_param(
            gold_props, pred_props, idMso, effect_type=24
        )
        if not com_result.passed:
            return tag_verification_source(com_result, "word_com")
        saved_result = verify_picture_saturation_persistence(
            pred_path, gold_path, saved_evidence_error=saved_evidence_error
        )
        return combine_com_primary_with_ooxml_guardrail(com_result, saved_result)
    elif idMso == "PictureArtisticEffectsGallery":
        com_result = families_pictures_compare.verify_picture_effects_param(
            gold_props, pred_props, idMso, compare_effect_type=True
        )
        if not com_result.passed:
            return tag_verification_source(com_result, "word_com")
        return combine_com_primary_with_ooxml_guardrail(com_result,
            verify_picture_effect_persistence(pred_path, gold_path, idMso,
                                              saved_evidence_error=saved_evidence_error))
    elif idMso == "OutlineColorPicker":
        exact_props = ["line_fore_color_rgb", "line_visible"]
    elif idMso == "PictureTransparencyGallery":
        numeric_props = {"fill_transparency": 0.01}
    else:
        # Fallback: compare all effect props (used for compositions)
        exact_props = [
            "shadow_type",
            "reflection_type",
            "auto_shape_type",
            "color_type",
            "threed_preset_camera",
            "threed_bevel_top_type",
            "line_fore_color_rgb",
            "line_visible",
        ]
        numeric_props = {
            "glow_radius": 1.0,
            "soft_edge_radius": 1.0,
            "brightness": 0.01,
            "contrast": 0.01,
            "threed_rotation_x": 1.0,
            "threed_rotation_y": 1.0,
            "threed_bevel_top_depth": 1.0,
            "fill_transparency": 0.01,
            # Crop values (for layout-coupled crop + effect compositions).
            # Matches _verify_picture_crop's 10pt tolerance.
            "crop_left": 10.0,
            "crop_right": 10.0,
            "crop_top": 10.0,
            "crop_bottom": 10.0,
        }

    # Compositions also compare these camera fields. Attribute recovery only
    # when this route actually consumes the recovered camera observations.
    if not set(CAMERA_PROPERTIES).intersection([*exact_props, *numeric_props]):
        camera_fallback = {}
    verification_source = "word_com+native_ooxml" if camera_fallback else "word_com"
    differences = []

    if idMso in {
        "PictureEffectsShadowGallery", "PictureEffectsGlowGallery", "PictureEffectsSoftEdgesGallery",
        "PictureReflectionGallery", "PictureRotationGallery", "_3DBevelPictureTopGallery",
        "PictureShapeGallery", "PictureRecolorGallery", "RecolorColorPicker",
        "OutlineColorPicker", "PictureTransparencyGallery",
    }:
        error = require_observations(
            {"gold": gold_props, "pred": pred_props},
            [*exact_props, *numeric_props], context=idMso,
        )
        if error is not None:
            return error

    for prop in exact_props:
        gold_val = gold_props.get(prop)
        pred_val = pred_props.get(prop)
        if gold_val is None and pred_val is None:
            continue
        if gold_val != pred_val:
            differences.append(
                {
                    "property": prop,
                    "gold": gold_val,
                    "pred": pred_val,
                }
            )

    for prop, tol in numeric_props.items():
        gold_val = gold_props.get(prop)
        pred_val = pred_props.get(prop)
        if gold_val is None and pred_val is None:
            continue
        if gold_val is not None and pred_val is not None:
            if abs(gold_val - pred_val) > tol:
                differences.append(
                    {
                        "property": prop,
                        "gold": gold_val,
                        "pred": pred_val,
                    }
                )
        elif gold_val != pred_val:
            differences.append(
                {
                    "property": prop,
                    "gold": gold_val,
                    "pred": pred_val,
                }
            )

    # For compositions (fallback path), also compare picture_effects array
    if not idMso:
        gold_effects = sorted(
            gold_props.get("picture_effects", []),
            key=lambda e: e.get("type", 0),
        )
        pred_effects = sorted(
            pred_props.get("picture_effects", []),
            key=lambda e: e.get("type", 0),
        )
        if len(gold_effects) != len(pred_effects):
            differences.append(
                {
                    "property": "picture_effects_count",
                    "gold": len(gold_effects),
                    "pred": len(pred_effects),
                }
            )
        else:
            for ge, pe in zip(gold_effects, pred_effects):
                if ge.get("type") != pe.get("type"):
                    differences.append(
                        {
                            "property": "picture_effect_type",
                            "gold": ge.get("type"),
                            "pred": pe.get("type"),
                        }
                    )
                elif ge.get("params") != pe.get("params"):
                    differences.append(
                        {
                            "property": f"picture_effect_{ge.get('type')}_params",
                            "gold": ge.get("params"),
                            "pred": pe.get("params"),
                        }
                    )

    if differences:
        return tag_verification_source(
            VerificationResult(
                passed=False,
                message=f"FAIL: Picture effect mismatch ({len(differences)} difference(s))",
                details={
                    "differences": differences,
                    "idMso": idMso,
                    "gold": gold_props,
                    "pred": pred_props,
                },
            ),
            verification_source,
        )

    com_result = VerificationResult(
        passed=True, message="PASS: Picture effect properties match", details={"idMso": idMso}
    )
    if camera_fallback:
        com_result.details["camera_fallback"] = camera_fallback
    if idMso == "_3DBevelPictureTopGallery":
        try:
            gold_bevels = families_shapes_extract_saved.extract_bevel_render_signatures(gold_path)
        except Exception as exc:
            return saved_evidence_error(
                f"Gold saved picture-bevel extraction failed: {exc}",
                evidence_owner="gold",
                exception=exc,
            )
        if not gold_bevels:
            return saved_evidence_error(
                "Gold saved picture-bevel evidence is empty",
                evidence_owner="gold",
            )
        required_scene_fields = (
            "scene_present",
            "camera_present",
            "light_rig_present",
            "camera_prst_present",
            "light_rig_type_present",
            "light_rig_direction_present",
        )
        incomplete_gold = [
            {
                "shape_index": index,
                "missing": [field for field in required_scene_fields if not signature.get(field)],
            }
            for index, signature in enumerate(gold_bevels, start=1)
            if any(not signature.get(field) for field in required_scene_fields)
        ]
        if incomplete_gold:
            result = saved_evidence_error(
                "Gold saved picture-bevel render scene is incomplete",
                evidence_owner="gold",
            )
            result.details["incomplete_gold"] = incomplete_gold
            return result
        try:
            saved_failures = families_shapes_geometry.bevel_render_failures(pred_path, gold_path)
        except Exception as exc:
            return saved_evidence_error(
                f"Prediction saved picture-bevel extraction failed: {exc}",
                evidence_owner="prediction",
                exception=exc,
            )
        saved_result = VerificationResult(
            passed=not saved_failures,
            message=(
                "PASS: Saved picture-bevel preset and render scene are complete"
                if not saved_failures
                else "FAIL: Saved picture-bevel preset or render scene is incomplete"
            ),
            details={
                "match_mode": "saved_ooxml_picture_bevel_effective_state",
                "differences": saved_failures,
                "gold": gold_bevels,
            },
        )
        return combine_com_primary_with_ooxml_guardrail(com_result, saved_result)
    if idMso == "PictureEffectsGlowGallery":
        gold_glows = families_pictures_extract_saved.extract_saved_picture_glows(gold_path)
        pred_glows = families_pictures_extract_saved.extract_saved_picture_glows(pred_path)
        if gold_glows is not None and pred_glows is not None:
            saved_result = VerificationResult(
                passed=not (gold_glows and pred_glows == []),
                message=(
                    "PASS: Saved picture glow is present"
                    if not (gold_glows and pred_glows == [])
                    else "FAIL: Saved picture glow is missing"
                ),
                details={
                    "gold_saved_glows": gold_glows,
                    "pred_saved_glows": pred_glows,
                    "match_mode": "saved_ooxml_picture_glow_presence",
                },
            )
            return combine_com_primary_with_ooxml_guardrail(com_result, saved_result)
    return tag_verification_source(com_result, verification_source)


def verify_picture_effect_persistence(pred_path, gold_path, command, *, saved_evidence_error):
    """COM parameters are primary; the submitted target must retain their source graph."""
    effect_name = {
        "PictureColorTempertatureGallery": "colorTemperature",
        "PictureSharpenSoftenGallery": "sharpenSoften",
        "PictureArtisticEffectsGallery": "artistic",
    }[command]
    def read(path):
        return families_pictures_extract_saved.extract_picture_effect_persistence(path, effect_name)
    try:
        gold = read(gold_path)
    except Exception as exc:
        return saved_evidence_error(f"Gold saved {effect_name} extraction failed: {exc}",
                                    evidence_owner="gold", exception=exc)
    if not gold or any(not item['graph_valid'] for item in gold):
        return saved_evidence_error(f"Gold {effect_name} persistence graph is missing or invalid",
                                    evidence_owner="gold")
    try:
        pred = read(pred_path)
    except Exception as exc:
        return saved_evidence_error(f"Prediction saved {effect_name} extraction failed: {exc}",
                                    evidence_owner="prediction", exception=exc)
    passed = ([item['effect'] for item in pred] == [item['effect'] for item in gold]
              and all(item['graph_valid'] for item in pred))
    return VerificationResult(passed=passed,
        message=("PASS: Saved picture effect persistence is valid" if passed
                 else "FAIL: Saved picture effect persistence graph is missing or invalid"),
        details={"match_mode":"saved_ooxml_picture_effect_hdphoto_persistence",
                 "gold":gold,"pred":pred,"idMso":command})


def verify_picture_saturation_persistence(
    pred_path: Path, gold_path: Path, *, saved_evidence_error
) -> VerificationResult:
    """Require the saved HD Photo graph that Word preserves on reopen."""
    try:
        gold = families_pictures_extract_saved.extract_picture_saturation_persistence(gold_path)
    except Exception as exc:
        return saved_evidence_error(
            f"Gold saved saturation extraction failed: {exc}",
            evidence_owner="gold",
            exception=exc,
        )
    if not gold or any(not item.get("graph_valid") for item in gold):
        return saved_evidence_error(
            "Gold saved saturation evidence lacks a valid internal "
            "HD Photo relationship and WDP part",
            evidence_owner="gold",
        )
    try:
        pred = families_pictures_extract_saved.extract_picture_saturation_persistence(pred_path)
    except Exception as exc:
        return saved_evidence_error(
            f"Prediction saved saturation extraction failed: {exc}",
            evidence_owner="prediction",
            exception=exc,
        )
    gold_values = [item.get("sat") for item in gold]
    pred_values = [item.get("sat") for item in pred]
    passed = (
        pred_values == gold_values
        and len(pred) == len(gold)
        and all(item.get("graph_valid") for item in pred)
    )
    return VerificationResult(
        passed=passed,
        message=(
            "PASS: Saved saturation persistence graph is valid"
            if passed
            else "FAIL: Saved saturation persistence graph is invalid"
        ),
        details={
            "match_mode": "saved_ooxml_saturation_hdphoto_persistence",
            "gold": gold,
            "pred": pred,
        },
    )


def verify_named_picture_effect_state(
    pred_path: Path,
    gold_path: Path,
    id_mso: str,
    pred_props: Dict[str, Any],
    gold_props: Dict[str, Any],
) -> VerificationResult:
    family = "outerShdw" if id_mso == "PictureEffectsShadowGallery" else "reflection"
    com_property = "shadow_type" if family == "outerShdw" else "reflection_type"
    try:
        gold_effects = families_pictures_extract_saved.extract_saved_picture_named_effects(
            gold_path
        )[family]
        pred_effects = families_pictures_extract_saved.extract_saved_picture_named_effects(
            pred_path
        )[family]
    except Exception as exc:
        return VerificationResult.error_result(
            f"Named picture-effect saved evidence extraction failed: {exc}"
        )
    if len(gold_effects) != 1:
        return VerificationResult.error_result(
            f"Gold must contain exactly one saved {family} effect"
        )
    if len(pred_effects) != 1:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Pred must contain exactly one saved {family} effect",
            details={
                "match_mode": "saved_ooxml_named_picture_effect",
                "family": family,
                "gold": gold_effects,
                "pred": pred_effects,
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

    gold_effect = gold_effects[0]
    pred_effect = pred_effects[0]
    failures = []
    pred_attributes = pred_effect["attributes"]
    gold_attributes = gold_effect["attributes"]
    if family == "outerShdw":
        for owner, effect in (("gold", gold_effect), ("pred", pred_effect)):
            attributes = effect["attributes"]
            blur = integer(attributes, "blurRad")
            alpha = integer({"alpha": effect.get("alpha")}, "alpha")
            black = (
                effect.get("color_type") == "prstClr"
                and (effect.get("color_value") or "").casefold() == "black"
                or effect.get("color_type") == "srgbClr"
                and (effect.get("color_value") or "").upper() == "000000"
            )
            if (
                blur is None
                or blur <= 0
                or not families_pictures_compare.outer_shadow_has_centered_zero_offset(attributes)
                or not black
                or alpha is None
                or not 0 < alpha <= 100000
            ):
                failures.append(
                    f"{owner} shadow is not centered, zero-distance, "
                    "blurred black geometry with alpha"
                )
    else:
        required = ("blurRad", "stA", "endA", "endPos", "dir", "sy", "algn")
        for owner, attributes in (("gold", gold_attributes), ("pred", pred_attributes)):
            missing = [name for name in required if name not in attributes]
            if missing or integer(attributes, "dist", 0) != 0:
                failures.append(f"{owner} reflection geometry is incomplete or nonzero-distance")
        if not failures:
            for name in ("blurRad", "endA", "endPos", "dir", "sy"):
                if integer(pred_attributes, name) != integer(gold_attributes, name):
                    failures.append(f"reflection {name} differs")
            if pred_attributes.get("algn") != gold_attributes.get("algn"):
                failures.append("reflection alignment differs")
            if abs(integer(pred_attributes, "stA") - integer(gold_attributes, "stA")) > 3000:
                failures.append("reflection start opacity differs")
            if pred_attributes.get("rotWithShape", "0") != gold_attributes.get("rotWithShape", "0"):
                failures.append("reflection rotation semantics differ")

    gold_com = gold_props.get(com_property)
    pred_com = pred_props.get(com_property)
    if pred_com != gold_com and pred_com != -2:
        failures.append(f"COM {com_property} pred={pred_com}, gold={gold_com}")
    if failures:
        return tag_verification_source(
            VerificationResult(
                passed=False,
                message="FAIL: Named picture effect semantic state differs",
                details={
                    "match_mode": "named_picture_effect_semantic_projection",
                    "family": family,
                    "failures": failures,
                    "gold_saved": gold_effect,
                    "pred_saved": pred_effect,
                    "gold_com": gold_com,
                    "pred_com": pred_com,
                },
            ),
            "word_com+saved_ooxml_guardrail",
        )
    return tag_verification_source(
        VerificationResult(
            passed=True,
            message="PASS: Named picture effect semantic state matches",
            details={
                "match_mode": "named_picture_effect_semantic_projection",
                "family": family,
                "gold_saved": gold_effect,
                "pred_saved": pred_effect,
                "gold_com": gold_com,
                "pred_com": pred_com,
            },
        ),
        "word_com+saved_ooxml_guardrail",
    )


def verify_alt_text(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, read_inline_shapes
) -> VerificationResult:
    """Verify AlternativeText and Title on inline shapes match between pred and gold."""
    gold_info = read_inline_shapes(str(gold_path))
    pred_info = read_inline_shapes(str(pred_path))

    if gold_info["inline_shape_count"] != pred_info["inline_shape_count"]:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Inline shape count mismatch (gold={gold_info['inline_shape_count']}, pred={pred_info['inline_shape_count']})",
            details={"gold": gold_info, "pred": pred_info},
        )

    differences = []
    for gold_shape, pred_shape in zip(gold_info["inline_shapes"], pred_info["inline_shapes"]):
        for prop in ["alternative_text", "title"]:
            gold_val = gold_shape.get(prop, "")
            pred_val = pred_shape.get(prop, "")
            if gold_val != pred_val:
                differences.append(
                    {
                        "index": gold_shape.get("index"),
                        "property": prop,
                        "gold": gold_val,
                        "pred": pred_val,
                    }
                )

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: {len(differences)} alt text property difference(s)",
            details={"differences": differences},
        )

    return VerificationResult(
        passed=True,
        message="PASS: Alt text properties match",
        details={
            "inline_shape_count": gold_info["inline_shape_count"],
            "gold_shapes": gold_info["inline_shapes"],
            "pred_shapes": pred_info["inline_shapes"],
        },
    )
