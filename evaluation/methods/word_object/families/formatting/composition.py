"""families / formatting / composition with explicit evidence dependencies."""

from .text import _bullet_chars_equivalent
from .contract import NULLABLE_FIELDS
from ...observations import require_observations
from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.families.formatting import (
    compare as families_formatting_compare,
)
from pathlib import Path
from typing import Any, Dict, List, Optional


def verify_composition_atomic_targets(
    pred_path: Path, gold_path: Path, spec: Dict, *, read_all_paragraphs_formatting
) -> Optional[VerificationResult]:
    """Verify composition target scopes not covered by target-text comparison."""
    intents = spec.get("composition", {}).get("atomic_intents", [])
    if not intents:
        return None

    all_paragraph_targets = []
    for intent in intents:
        for target in intent.get("verification_targets", []) or []:
            if target.get("extractor") == "all_paragraphs_formatting":
                all_paragraph_targets.append(target)
    all_paragraph_targets = list(
        {
            (target.get("property"), target.get("comparison")): target
            for target in all_paragraph_targets
            if target.get("property")
        }.values()
    )

    if not all_paragraph_targets:
        return None

    gold_paras = read_all_paragraphs_formatting(str(gold_path))
    pred_paras = read_all_paragraphs_formatting(str(pred_path))

    if gold_paras is None or pred_paras is None:
        return VerificationResult(
            passed=False,
            message="ERROR: Could not extract all paragraph formatting for composition targets",
            error="Could not extract all paragraph formatting for composition targets",
            details={
                "gold_extracted": gold_paras is not None,
                "pred_extracted": pred_paras is not None,
            },
        )

    gold_paras = families_formatting_compare.normalize_dropcap_paragraphs(gold_paras)
    pred_paras = families_formatting_compare.normalize_dropcap_paragraphs(pred_paras)

    if len(gold_paras) != len(pred_paras):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Paragraph count mismatch (gold: {len(gold_paras)}, pred: {len(pred_paras)})",
            details={
                "gold_count": len(gold_paras),
                "pred_count": len(pred_paras),
                "source": "composition_atomic_targets",
            },
        )

    differences = []
    for i, (gold_para, pred_para) in enumerate(zip(gold_paras, pred_paras)):
        for target in all_paragraph_targets:
            prop = target.get("property")
            observation_error = require_observations(
                {"gold": gold_para, "pred": pred_para}, [prop],
                context=f"composition paragraph {i + 1}", nullable=NULLABLE_FIELDS)
            if observation_error is not None:
                return observation_error
            gold_val = gold_para.get(prop)
            pred_val = pred_para.get(prop)
            comparison = target.get(
                "comparison"
            ) or families_formatting_compare.composition_indent_comparison(prop, spec)
            if not families_formatting_compare.reference_comparison_matches(
                gold_val, pred_val, comparison
            ):
                differences.append(
                    {
                        "property": prop,
                        "comparison": comparison or "exact",
                        "gold": gold_val,
                        "pred": pred_val,
                        "location": f"paragraph {i + 1}",
                        "source": "composition_atomic_targets",
                    }
                )

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: {len(differences)} composition target difference(s) across paragraphs",
            details={
                "differences": differences[:10],
                "total_differences": len(differences),
                "paragraph_count": len(gold_paras),
                "properties": [target.get("property") for target in all_paragraph_targets],
            },
        )

    return None


def add_composition_atomic_details(
    details: Dict[str, Any],
    pred_path: Path,
    gold_path: Path,
    spec: Dict,
    target_text: str,
    *,
    read_all_paragraphs_formatting,
    read_document_text,
    read_dropcap_properties,
    read_text_formatting,
) -> None:
    """Attach intent-level target comparison details for composition specs."""
    atomic_results = build_composition_atomic_results(
        pred_path,
        gold_path,
        spec,
        target_text,
        read_all_paragraphs_formatting=read_all_paragraphs_formatting,
        read_document_text=read_document_text,
        read_dropcap_properties=read_dropcap_properties,
        read_text_formatting=read_text_formatting,
    )
    if not atomic_results:
        return

    target_results = []
    for intent_result in atomic_results:
        for target_result in intent_result.get("target_results", []):
            flat_target = dict(target_result)
            flat_target["intent_id"] = intent_result.get("intent_id")
            flat_target["operation"] = intent_result.get("operation")
            flat_target["atomic_case"] = intent_result.get("atomic_case")
            target_results.append(flat_target)

    passed_intents = sum(1 for result in atomic_results if result.get("passed"))
    details["atomic_results"] = atomic_results
    details["target_results"] = target_results
    details["atomic_summary"] = {
        "passed_intents": passed_intents,
        "total_intents": len(atomic_results),
        "passed_targets": sum(1 for result in target_results if result.get("passed")),
        "total_targets": len(target_results),
    }


def build_composition_atomic_results(
    pred_path: Path,
    gold_path: Path,
    spec: Dict,
    target_text: str,
    *,
    read_all_paragraphs_formatting,
    read_document_text,
    read_dropcap_properties,
    read_text_formatting,
) -> List[Dict[str, Any]]:
    intents = spec.get("composition", {}).get("atomic_intents", [])
    if not intents:
        return []

    cache: Dict[str, Any] = {}
    atomic_results = []
    for intent in intents:
        target_results = [
            evaluate_composition_target(
                target,
                pred_path,
                gold_path,
                target_text,
                spec,
                cache,
                operation=intent.get("operation"),
                atomic_case=intent.get("atomic_case"),
                read_all_paragraphs_formatting=read_all_paragraphs_formatting,
                read_document_text=read_document_text,
                read_dropcap_properties=read_dropcap_properties,
                read_text_formatting=read_text_formatting,
            )
            for target in intent.get("verification_targets", []) or []
        ]
        atomic_results.append(
            {
                "intent_id": intent.get("id"),
                "operation": intent.get("operation"),
                "atomic_case": intent.get("atomic_case"),
                "passed": bool(target_results)
                and all(result.get("passed") for result in target_results),
                "target_results": target_results,
            }
        )

    return atomic_results


def evaluate_composition_target(
    target: Dict[str, Any],
    pred_path: Path,
    gold_path: Path,
    target_text: str,
    spec: Dict,
    cache: Dict[str, Any],
    operation: Optional[str] = None,
    atomic_case: Optional[str] = None,
    *,
    read_all_paragraphs_formatting,
    read_document_text,
    read_dropcap_properties,
    read_text_formatting,
) -> Dict[str, Any]:
    extractor = target.get("extractor")
    prop = target.get("property")
    result: Dict[str, Any] = {
        "target": {
            "extractor": extractor,
            "property": prop,
        },
        "passed": False,
    }
    target_comparison = target.get(
        "comparison"
    ) or families_formatting_compare.composition_indent_comparison(prop, spec)
    if target_comparison:
        result["target"]["comparison"] = target_comparison

    if not extractor or not prop:
        result["error"] = "Missing extractor or property"
        return result

    if extractor in ("text_formatting", "paragraph_formatting"):
        cache_key = f"text_formatting:{target_text}"
        if cache_key not in cache:
            cache[cache_key] = {
                "gold": read_text_formatting(str(gold_path), target_text),
                "pred": read_text_formatting(str(pred_path), target_text),
            }
        extracted = cache[cache_key]
        gold_fmt = extracted["gold"]
        pred_fmt = extracted["pred"]
        if gold_fmt is None or pred_fmt is None:
            result.update(
                {
                    "error": "Target text not found",
                    "gold_extracted": gold_fmt is not None,
                    "pred_extracted": pred_fmt is not None,
                }
            )
            return result

        observation_error = require_observations(
            {"gold": gold_fmt, "pred": pred_fmt}, [prop],
            context="composition target", nullable=NULLABLE_FIELDS)
        if observation_error is not None:
            result.update(error=observation_error.error, observations=observation_error.details)
            return result
        gold_val = gold_fmt.get(prop)
        pred_val = pred_fmt.get(prop)
        shading_state = families_formatting_compare.character_shading_state(atomic_case)
        if (
            operation == "char_shading"
            and prop in {"shading", "shading_color"}
            and shading_state is not None
        ):
            pred_has_shading = families_formatting_compare.character_shading_present(pred_fmt)
            passed = pred_has_shading if shading_state == "on" else not pred_has_shading
        else:
            passed = families_formatting_compare.composition_target_values_match(
                prop,
                gold_val,
                pred_val,
                extractor,
                gold_fmt,
                pred_fmt,
                spec,
                comparison=target_comparison,
            )
        result.update(
            {
                "passed": passed,
                "gold": gold_val,
                "pred": pred_val,
            }
        )
        if not passed:
            result["differences"] = [
                {
                    "property": prop,
                    "gold": gold_val,
                    "pred": pred_val,
                    "location": "target_text",
                }
            ]
        return result

    if extractor == "all_paragraphs_formatting":
        if extractor not in cache:
            cache[extractor] = {
                "gold": read_all_paragraphs_formatting(str(gold_path)),
                "pred": read_all_paragraphs_formatting(str(pred_path)),
            }
        gold_paras = cache[extractor]["gold"]
        pred_paras = cache[extractor]["pred"]
        if gold_paras is None or pred_paras is None:
            result.update(
                {
                    "error": "Could not extract all paragraph formatting",
                    "gold_extracted": gold_paras is not None,
                    "pred_extracted": pred_paras is not None,
                }
            )
            return result
        gold_paras = families_formatting_compare.normalize_dropcap_paragraphs(gold_paras)
        pred_paras = families_formatting_compare.normalize_dropcap_paragraphs(pred_paras)
        if len(gold_paras) != len(pred_paras):
            result.update(
                {
                    "error": "Paragraph count mismatch",
                    "gold_count": len(gold_paras),
                    "pred_count": len(pred_paras),
                }
            )
            return result

        differences = []
        gold_values = []
        pred_values = []
        normalize_bullets = families_formatting_compare.composition_uses_bullet_normalization(spec)
        for i, (gold_para, pred_para) in enumerate(zip(gold_paras, pred_paras)):
            observation_error = require_observations(
                {"gold": gold_para, "pred": pred_para}, [prop],
                context=f"composition paragraph {i + 1}", nullable=NULLABLE_FIELDS)
            if observation_error is not None:
                result.update(error=observation_error.error, observations=observation_error.details)
                return result
            gold_val = gold_para.get(prop)
            pred_val = pred_para.get(prop)
            gold_values.append(gold_val)
            pred_values.append(pred_val)
            values_match = families_formatting_compare.reference_comparison_matches(
                gold_val, pred_val, target_comparison
            )
            if normalize_bullets and prop == "list_string":
                values_match = _bullet_chars_equivalent(
                    gold_val,
                    pred_val,
                    gold_para.get("list_font"),
                    pred_para.get("list_font"),
                )
            if not values_match:
                differences.append(
                    {
                        "property": prop,
                        "gold": gold_val,
                        "pred": pred_val,
                        "location": f"paragraph {i + 1}",
                    }
                )

        result.update(
            {
                "passed": not differences,
                "gold": gold_values,
                "pred": pred_values,
                "paragraph_count": len(gold_paras),
            }
        )
        if differences:
            result["differences"] = differences[:10]
            result["total_differences"] = len(differences)
        return result

    if extractor == "dropcap_properties":
        if extractor not in cache:
            cache[extractor] = {
                "gold": read_dropcap_properties(str(gold_path)),
                "pred": read_dropcap_properties(str(pred_path)),
            }
        gold_props = cache[extractor]["gold"]
        pred_props = cache[extractor]["pred"]
        observation_error = require_observations(
            {"gold": gold_props, "pred": pred_props}, [prop], context="composition drop cap target")
        if observation_error is not None:
            result.update(error=observation_error.error, observations=observation_error.details)
            return result
        gold_val = gold_props.get(prop)
        pred_val = pred_props.get(prop)
        passed = gold_val == pred_val
        result.update(
            {
                "passed": passed,
                "gold": gold_val,
                "pred": pred_val,
            }
        )
        if not passed:
            result["differences"] = [
                {
                    "property": prop,
                    "gold": gold_val,
                    "pred": pred_val,
                    "location": "document",
                }
            ]
        return result

    if extractor == "document_text":
        if extractor not in cache:
            cache[extractor] = {
                "gold": read_document_text(str(gold_path)),
                "pred": read_document_text(str(pred_path)),
            }
        gold_text = cache[extractor]["gold"]
        pred_text = cache[extractor]["pred"]
        if gold_text is None or pred_text is None:
            result.update(error="Composition document text could not be read")
            return result
        passed = gold_text == pred_text
        result.update(
            {
                "passed": passed,
                "gold": gold_text,
                "pred": pred_text,
            }
        )
        if not passed:
            result["differences"] = [
                {
                    "property": prop,
                    "gold": gold_text,
                    "pred": pred_text,
                    "location": "document_text",
                }
            ]
        return result

    result["error"] = f"Unsupported composition target extractor: {extractor}"
    return result
