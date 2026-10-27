"""families / pictures / compare with explicit evidence dependencies."""

from evaluation.methods.word_object.effect_commands import POSITIVE_PICTURE_EFFECT_COMMANDS
from evaluation.methods.base import VerificationResult
from typing import Any, Dict, List, Optional
from math import isfinite


def picture_extraction_failure(error: Any, *, idMso: str = "") -> VerificationResult:
    """Classify a predicted-picture extraction failure.

    A missing or wrong-kind predicted object is an observable benchmark
    failure, not evaluator infrastructure.  Other extraction failures can
    still mean Word could not make the requested observation and remain
    evaluator errors.
    """
    message = str(error)
    object_mismatch_prefixes = (
        "No inline shape",
        "No inline or floating picture",
        "Inline shape is not a picture",
    )
    if message.startswith(object_mismatch_prefixes):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Predicted picture is unavailable ({message})",
            details={"pred_error": message, "idMso": idMso},
        )
    return VerificationResult.error_result(f"Pred: {message}")


def uses_named_picture_effect_projection(id_mso: str, spec: Optional[Dict]) -> bool:
    """Scope semantic sentinel projection to audited atomic intents.

    Atomic cases declare the preset through ``variation``.  Composition
    cases retain the same contract in each intent's ``atomic_case`` while
    their top-level variation names the bundle.  Inspect both forms so a
    centered-shadow or touching-reflection composition receives the same
    saved-geometry projection without relaxing unrelated gallery presets.
    """
    variation = str((spec or {}).get("variation") or "").casefold()
    approved = {
        "PictureEffectsShadowGallery": {"outer_center"},
        "PictureReflectionGallery": {"full_touching", "tight_touching"},
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


def outer_shadow_has_centered_zero_offset(attributes: Dict[str, Any]) -> bool:
    """Return whether saved outer-shadow geometry is effectively centered.

    DrawingML alignment selects the origin for scale/skew transforms. If
    distance is zero and every transform keeps its identity default, there
    is no origin-dependent displacement regardless of the serialized
    alignment (including the schema-default bottom alignment). Any
    non-identity scale/skew still requires centered alignment.
    """
    try:
        distance = int(attributes.get("dist", 0))
        scale_x = int(attributes.get("sx", 100000))
        scale_y = int(attributes.get("sy", 100000))
        skew_x = int(attributes.get("kx", 0))
        skew_y = int(attributes.get("ky", 0))
    except (TypeError, ValueError):
        return False
    if distance != 0:
        return False
    alignment = attributes.get("algn")
    identity_transform = scale_x == 100000 and scale_y == 100000 and skew_x == 0 and skew_y == 0
    return alignment == "ctr" or identity_transform


def verify_picture_effects_param(
    gold_props: Dict,
    pred_props: Dict,
    idMso: str,
    effect_type: int = None,
    compare_effect_type: bool = False,
) -> VerificationResult:
    """
    Compare PictureEffects collection between gold and pred.

    For effect_type: find the effect with matching type and compare parameter values.
    For compare_effect_type: compare the effect type itself (e.g. artistic effect kind).
    """
    gold_effects = gold_props.get("picture_effects", []) or []
    pred_effects = pred_props.get("picture_effects", []) or []
    positive_commands = POSITIVE_PICTURE_EFFECT_COMMANDS

    def relevant_extraction_errors(props: Dict) -> List[Dict]:
        relevant = []
        for item in props.get("_extraction_errors", []) or []:
            prop = item.get("property")
            if prop == "picture_effects":
                relevant.append(item)
            elif (
                prop == "picture_effect_params"
                and not compare_effect_type
                and item.get("effect_type") == effect_type
            ):
                relevant.append(item)
        return relevant

    if idMso in positive_commands:
        gold_extraction_errors = relevant_extraction_errors(gold_props)
        if gold_extraction_errors:
            result = VerificationResult.error_result(
                "Gold picture-effect evidence could not be extracted"
            )
            result.details.update(
                {
                    "error_owner": "gold_fixture",
                    "reason": "gold_picture_effect_extraction_failed",
                    "idMso": idMso,
                    "extraction_errors": gold_extraction_errors,
                }
            )
            return result

        pred_extraction_errors = relevant_extraction_errors(pred_props)
        if pred_extraction_errors:
            return VerificationResult.error_result(
                "Predicted picture-effect evidence could not be extracted",
                observation_owner="prediction",
                reason="prediction_picture_effect_extraction_failed",
                idMso=idMso,
                extraction_errors=pred_extraction_errors,
            )

    if compare_effect_type:
        # Compare effect types (e.g. artisticGlowEdges=12 vs artisticPaintStrokes=18)
        for owner, effects in (("gold", gold_effects), ("prediction", pred_effects)):
            if any(effect.get("type") is None for effect in effects):
                return VerificationResult.error_result(
                    f"{owner} picture-effect type is unavailable",
                    observation_owner=owner, reason="unavailable_picture_effect_type",
                    idMso=idMso,
                )
        if idMso in positive_commands and not gold_effects:
            result = VerificationResult.error_result(
                "Gold positive picture-effect collection is empty"
            )
            result.details.update(
                {
                    "error_owner": "gold_fixture",
                    "reason": "empty_gold_picture_effect_collection",
                    "idMso": idMso,
                }
            )
            return result
        if idMso in positive_commands and not pred_effects:
            return VerificationResult(
                passed=False,
                message="FAIL: Predicted positive picture-effect collection is empty",
                details={
                    "failure_owner": "prediction",
                    "reason": "empty_prediction_picture_effect_collection",
                    "idMso": idMso,
                },
            )
        gold_types = sorted([e.get("type") for e in gold_effects])
        pred_types = sorted([e.get("type") for e in pred_effects])
        if gold_types != pred_types:
            return VerificationResult(
                passed=False,
                message=f"FAIL: Picture effect type mismatch",
                details={
                    "property": "picture_effect_types",
                    "gold": gold_types,
                    "pred": pred_types,
                    "idMso": idMso,
                },
            )
        return VerificationResult(
            passed=True, message="PASS: Picture effect types match", details={"idMso": idMso}
        )

    # Find effect with matching type and compare params
    gold_eff = next((e for e in gold_effects if e.get("type") == effect_type), None)
    pred_eff = next((e for e in pred_effects if e.get("type") == effect_type), None)

    if gold_eff is None and idMso in positive_commands:
        result = VerificationResult.error_result(
            f"Gold positive picture effect type={effect_type} is missing"
        )
        result.details.update(
            {
                "error_owner": "gold_fixture",
                "reason": "missing_gold_picture_effect",
                "effect_type": effect_type,
                "idMso": idMso,
            }
        )
        return result

    if gold_eff is None and pred_eff is None:
        return VerificationResult(
            passed=True,
            message=f"PASS: No picture effect type={effect_type} in either doc",
            details={"idMso": idMso},
        )

    if gold_eff is None or pred_eff is None:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Picture effect type={effect_type} present in only one doc",
            details={
                "gold_has_effect": gold_eff is not None,
                "pred_has_effect": pred_eff is not None,
                "idMso": idMso,
            },
        )

    gold_params = gold_eff.get("params", [])
    pred_params = pred_eff.get("params", [])
    if idMso in positive_commands and not gold_params:
        result = VerificationResult.error_result(
            f"Gold positive picture effect type={effect_type} has no parameters"
        )
        result.details.update(
            {
                "error_owner": "gold_fixture",
                "reason": "empty_gold_picture_effect_params",
                "effect_type": effect_type,
                "idMso": idMso,
            }
        )
        return result
    if idMso in positive_commands:
        for owner, params in (("gold", gold_params), ("prediction", pred_params)):
            if any(value is None or
                   (isinstance(value, (int, float)) and not isfinite(value))
                   for value in (params or [])):
                return VerificationResult.error_result(
                    f"{owner} picture-effect parameters are unavailable or non-finite",
                    observation_owner=owner, reason="invalid_picture_effect_params",
                    effect_type=effect_type, idMso=idMso,
                )
    if idMso in positive_commands and not pred_params:
        return VerificationResult(
            passed=False,
            message=(
                f"FAIL: Predicted positive picture effect type={effect_type} " "has no parameters"
            ),
            details={
                "failure_owner": "prediction",
                "reason": "empty_prediction_picture_effect_params",
                "effect_type": effect_type,
                "idMso": idMso,
            },
        )
    if gold_params != pred_params:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Picture effect params mismatch (type={effect_type})",
            details={
                "property": f"picture_effect_type_{effect_type}_params",
                "gold": gold_params,
                "pred": pred_params,
                "idMso": idMso,
            },
        )

    return VerificationResult(
        passed=True,
        message=f"PASS: Picture effect params match (type={effect_type})",
        details={"idMso": idMso},
    )
