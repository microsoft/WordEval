"""families / smartart / text with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from pathlib import Path
from typing import Dict, Optional
from ...observations import require_observations


def verify_text_effects(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    read_smartart_camera_state,
    read_smartart_subnode_properties,
    verify_shapes_match,
) -> VerificationResult:
    """
    Verify text effect properties on shapes/SmartArt match.

    Handles TextEffectTransformGallery, TextEffects3DRotationGallery,
    TextFillColorPicker, BevelTextGallery, TextAlignGallery.

    Uses SmartArt sub-node properties when the gold exposes SmartArt.
    A spec-declared SmartArt target fails closed when that evidence is not
    observable.  Non-SmartArt targets retain the existing shape fallback.
    """
    idMso = spec.get("idMso", "") if spec else ""

    # For SmartArt text effects, use sub-node extraction for precise comparison
    gold_sub = read_smartart_subnode_properties(str(gold_path))
    camera_evidence = None
    precondition = (spec or {}).get("precondition") or {}
    declares_smartart = str(precondition.get("type") or "").casefold() == "smartart"
    required_smartart_property = (
        {
            "TextFillColorPicker": "text_fill_rgb",
            "TextEffectTransformGallery": "warp_format",
            "BevelTextGallery": "bevel_top_type",
        }.get(idMso)
        if declares_smartart
        else None
    )

    if not gold_sub.get("found"):
        if declares_smartart:
            result = VerificationResult.error_result(
                "Gold SmartArt text-effect target is not observable"
            )
            result.details.update(
                {
                    "idMso": idMso,
                    "error_owner": "gold_fixture",
                    "reason": "unobservable_gold_smartart_text_effect_target",
                    "gold": gold_sub,
                }
            )
            return result

        # The non-SmartArt WordArt/textbox property contract remains
        # intentionally deferred; preserve its existing shape fallback.
        return verify_shapes_match(pred_path, gold_path, spec)

    if required_smartart_property is not None and gold_sub.get(required_smartart_property) is None:
        result = VerificationResult.error_result(
            "Gold SmartArt text-effect evidence is missing required "
            f"property: {required_smartart_property}"
        )
        result.details.update(
            {
                "idMso": idMso,
                "error_owner": "gold_fixture",
                "reason": "missing_gold_smartart_text_effect_property",
                "missing_evidence": [required_smartart_property],
                "gold": gold_sub,
            }
        )
        return result

    pred_sub = read_smartart_subnode_properties(str(pred_path))

    if gold_sub.get("found"):
        if not pred_sub.get("found"):
            return VerificationResult(
                passed=False,
                message="FAIL: No SmartArt found in pred document",
                details={
                    "idMso": idMso,
                    "failure_owner": "prediction",
                    "reason": "unobservable_prediction_smartart_text_effect_target",
                    "gold": gold_sub,
                    "pred": pred_sub,
                },
            )

        differences = []

        if required_smartart_property is not None:
            error = require_observations(
                {"gold": gold_sub, "pred": pred_sub}, [required_smartart_property],
                context=f"SmartArt text effect {idMso}")
            if error is not None:
                error.details["missing_evidence"] = list(error.details.get("unavailable", {}))
                return error

        if idMso == "TextFillColorPicker":
            if gold_sub.get("text_fill_rgb") != pred_sub.get("text_fill_rgb"):
                differences.append(
                    {
                        "property": "text_fill_rgb",
                        "gold": gold_sub.get("text_fill_rgb"),
                        "pred": pred_sub.get("text_fill_rgb"),
                    }
                )
        elif idMso == "TextEffectTransformGallery":
            if gold_sub.get("warp_format") != pred_sub.get("warp_format"):
                differences.append(
                    {
                        "property": "warp_format",
                        "gold": gold_sub.get("warp_format"),
                        "pred": pred_sub.get("warp_format"),
                    }
                )
        elif idMso == "BevelTextGallery":
            if gold_sub.get("bevel_top_type") != pred_sub.get("bevel_top_type"):
                differences.append(
                    {
                        "property": "bevel_top_type",
                        "gold": gold_sub.get("bevel_top_type"),
                        "pred": pred_sub.get("bevel_top_type"),
                    }
                )
        elif idMso == "TextEffects3DRotationGallery":
            gold_camera_state = read_smartart_camera_state(str(gold_path))
            pred_camera_state = read_smartart_camera_state(str(pred_path))
            camera_evidence = {
                "gold": gold_camera_state,
                "pred": pred_camera_state,
            }
            if gold_camera_state.get("error"):
                return VerificationResult.error_result(
                    "Gold SmartArt camera-state extraction failed: " f"{gold_camera_state['error']}"
                )
            if pred_camera_state.get("error"):
                return VerificationResult(
                    passed=False,
                    message="FAIL: Prediction SmartArt camera-state extraction failed",
                    details={
                        "idMso": idMso,
                        "prediction_camera_state": pred_camera_state,
                    },
                )
            gold_data = gold_camera_state.get("data_presets") or []
            gold_drawing = gold_camera_state.get("drawing_presets") or []
            if not gold_data or not gold_drawing or gold_data != gold_drawing:
                result = VerificationResult.error_result(
                    "Gold SmartArt 3D rotation lacks agreeing data/drawing presets"
                )
                result.details["gold_camera_state"] = gold_camera_state
                return result

            # This gallery is a *text* effect.  Word writes the requested
            # camera under the text body, so a prediction that writes the
            # same preset under shape properties rotates the SmartArt shape
            # instead of its text.  Gold defines the required property
            # level, which keeps every legitimate Word serialization valid
            # without naming a preset, case, or system.
            gold_text_body = gold_camera_state.get("text_body_presets") or []
            if not gold_text_body:
                result = VerificationResult.error_result(
                    "Gold SmartArt 3D text rotation is not stored on a " "text-body property"
                )
                result.details["gold_camera_state"] = gold_camera_state
                return result
            pred_text_body = pred_camera_state.get("text_body_presets") or []
            if gold_text_body != pred_text_body:
                differences.append(
                    {
                        "property": "smartart_camera_text_body_presets",
                        "gold": gold_text_body,
                        "pred": pred_text_body,
                        "pred_shape_presets": (pred_camera_state.get("shape_presets") or []),
                        "reason": (
                            "requested text 3-D rotation is not applied to the "
                            "SmartArt text body"
                        ),
                    }
                )

            pred_data = pred_camera_state.get("data_presets") or []
            pred_drawing = pred_camera_state.get("drawing_presets") or []
            if gold_data != pred_data:
                differences.append(
                    {
                        "property": "smartart_camera_data_presets",
                        "gold": gold_data,
                        "pred": pred_data,
                    }
                )
            if gold_drawing != pred_drawing:
                differences.append(
                    {
                        "property": "smartart_camera_drawing_presets",
                        "gold": gold_drawing,
                        "pred": pred_drawing,
                    }
                )
            if pred_data != pred_drawing:
                differences.append(
                    {
                        "property": "smartart_camera_data_drawing_agreement",
                        "gold": True,
                        "pred": False,
                    }
                )
        else:
            # Fallback: compare all sub-node text properties
            for prop in ["text_fill_rgb", "warp_format", "bevel_top_type", "threed_preset_camera"]:
                if gold_sub.get(prop) != pred_sub.get(prop):
                    differences.append(
                        {
                            "property": prop,
                            "gold": gold_sub.get(prop),
                            "pred": pred_sub.get(prop),
                        }
                    )

        if differences:
            return VerificationResult(
                passed=False,
                message=f"FAIL: Text effect mismatch ({len(differences)} difference(s))",
                details={
                    "differences": differences,
                    "idMso": idMso,
                    "gold": gold_sub,
                    "pred": pred_sub,
                    **(
                        {"smartart_camera_state": camera_evidence}
                        if camera_evidence is not None
                        else {}
                    ),
                },
            )

        return VerificationResult(
            passed=True,
            message="PASS: Text effect properties match",
            details={
                "idMso": idMso,
                **(
                    {"smartart_camera_state": camera_evidence}
                    if camera_evidence is not None
                    else {}
                ),
            },
        )

    # ``gold_sub`` is known to be found above; this is defensive only.
    return verify_shapes_match(pred_path, gold_path, spec)
