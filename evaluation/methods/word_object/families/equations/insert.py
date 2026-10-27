"""families / equations / insert with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.families.equations import (
    insert_saved as families_equations_insert_saved,
)
from pathlib import Path
from typing import Any, Dict, Optional


def verify_equations(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    com_primary_with_saved_guardrail,
    error_with_details,
    extract_saved_insertion_context,
    normalize_word_default_layout_twips,
    read_equation_info,
) -> VerificationResult:
    """Verify effective equation text plus persistent structure and context."""
    return com_primary_with_saved_guardrail(
        lambda: verify_equations_com(
            pred_path,
            gold_path,
            error_with_details=error_with_details,
            read_equation_info=read_equation_info,
        ),
        lambda: verify_equations_saved(
            pred_path,
            gold_path,
            spec or {},
            error_with_details=error_with_details,
            extract_saved_insertion_context=extract_saved_insertion_context,
            normalize_word_default_layout_twips=normalize_word_default_layout_twips,
        ),
    )


def verify_equations_com(
    pred_path: Path, gold_path: Path, *, error_with_details, read_equation_info
) -> VerificationResult:
    """Use Word's effective OMath collection as the primary observation."""
    gold_info = read_equation_info(str(gold_path))
    pred_info = read_equation_info(str(pred_path))

    if not gold_info.get("equation_count"):
        return error_with_details(
            "Gold effective equation evidence is missing",
            {"match_mode": "word_com_equation_text", "gold": gold_info},
        )
    if gold_info["equation_count"] != pred_info["equation_count"]:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Equation count mismatch (pred={pred_info['equation_count']}, gold={gold_info['equation_count']})",
            details={
                "pred_count": pred_info["equation_count"],
                "gold_count": gold_info["equation_count"],
            },
        )

    gold_equations = gold_info.get("equations", [])
    pred_equations = pred_info.get("equations", [])
    if len(gold_equations) != gold_info["equation_count"]:
        return error_with_details(
            "Gold effective equation text could not be extracted",
            {"match_mode": "word_com_equation_text", "gold": gold_info},
        )
    if len(pred_equations) != pred_info["equation_count"]:
        return error_with_details(
            "Prediction effective equation text could not be completely extracted",
            {"match_mode": "word_com_equation_text", "pred": pred_info},
        )
    compare_count = len(gold_equations)
    text_mismatches = []

    for i in range(compare_count):
        gold_text = gold_equations[i].get("text", "")
        pred_text = pred_equations[i].get("text", "")
        gold_norm = families_equations_insert_saved.normalize_equation_text(gold_text)
        pred_norm = families_equations_insert_saved.normalize_equation_text(pred_text)
        if gold_norm != pred_norm:
            text_mismatches.append(
                {
                    "equation_index": i + 1,
                    "pred_text": pred_text,
                    "gold_text": gold_text,
                    "pred_normalized": pred_norm,
                    "gold_normalized": gold_norm,
                }
            )

    if text_mismatches:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Equation text mismatch in {len(text_mismatches)} equation(s) (whitespace ignored)",
            details={
                "pred_count": pred_info["equation_count"],
                "gold_count": gold_info["equation_count"],
                "mismatches": text_mismatches,
                "match_mode": "equation_text_ignore_whitespace",
            },
        )

    return VerificationResult(
        passed=True,
        message=(
            f"PASS: Equations match ({pred_info['equation_count']} equations; "
            "text compared with whitespace ignored)"
        ),
        details={
            **pred_info,
            "match_mode": "equation_text_ignore_whitespace",
        },
    )


def verify_equations_saved(
    pred_path: Path,
    gold_path: Path,
    spec: Dict[str, Any],
    *,
    error_with_details,
    extract_saved_insertion_context,
    normalize_word_default_layout_twips,
) -> VerificationResult:
    """Require persistent OMML semantics, target anchor, and preservation."""
    id_mso = spec.get("idMso", "")
    allow_equation_gallery_wrapper = id_mso == "InsertBuildingBlocksEquationsGallery"
    allow_native_equation_wrapper = id_mso == "EquationInsertNew"
    gold_state = families_equations_insert_saved.extract_saved_equation_state(
        gold_path,
        allow_equation_gallery_wrapper=allow_equation_gallery_wrapper,
        allow_native_equation_wrapper=allow_native_equation_wrapper,
        extract_saved_insertion_context=extract_saved_insertion_context,
        normalize_word_default_layout_twips=normalize_word_default_layout_twips,
    )
    pred_state = families_equations_insert_saved.extract_saved_equation_state(
        pred_path,
        allow_equation_gallery_wrapper=allow_equation_gallery_wrapper,
        allow_native_equation_wrapper=allow_native_equation_wrapper,
        extract_saved_insertion_context=extract_saved_insertion_context,
        normalize_word_default_layout_twips=normalize_word_default_layout_twips,
    )
    if not gold_state["equation_count"]:
        return error_with_details(
            "Gold saved equation evidence is missing",
            {"match_mode": "saved_ooxml_equation_semantics", "gold": gold_state},
        )
    if pred_state["equation_count"] != gold_state["equation_count"]:
        return VerificationResult(
            passed=False,
            message=(
                "FAIL: Saved equation count mismatch "
                f"(pred={pred_state['equation_count']}, "
                f"gold={gold_state['equation_count']})"
            ),
            details={
                "match_mode": "saved_ooxml_equation_semantics",
                "gold": gold_state,
                "pred": pred_state,
            },
        )

    gold_equations = gold_state["equations"]
    pred_equations = pred_state["equations"]
    mismatches = []
    for index, (gold_equation, pred_equation) in enumerate(
        zip(gold_equations, pred_equations), start=1
    ):
        compared_fields = ["expression", "superscript_count"]
        if id_mso == "EquationInsertNew":
            compared_fields.extend(
                [
                    "container",
                    "body_index",
                    "preceding_text",
                    "following_text",
                ]
            )
        for field in compared_fields:
            if pred_equation.get(field) != gold_equation.get(field):
                mismatches.append(
                    {
                        "equation_index": index,
                        "property": field,
                        "gold": gold_equation.get(field),
                        "pred": pred_equation.get(field),
                    }
                )
        if id_mso == "EquationInsertNew":
            for field in ("equation_control", "showing_placeholder"):
                if pred_equation.get(field) != gold_equation.get(field):
                    mismatches.append(
                        {
                            "equation_index": index,
                            "property": field,
                            "gold": gold_equation.get(field),
                            "pred": pred_equation.get(field),
                        }
                    )

    if mismatches:
        return VerificationResult(
            passed=False,
            message="FAIL: Saved equation semantics or insertion anchor differs",
            details={
                "match_mode": "saved_ooxml_equation_semantics",
                "mismatches": mismatches,
                "gold_equations": gold_equations,
                "pred_equations": pred_equations,
            },
        )

    if pred_state["context"] != gold_state["context"]:
        return VerificationResult(
            passed=False,
            message="FAIL: Equation insertion changed surrounding content/layout",
            details={
                "match_mode": "saved_ooxml_equation_context",
                "gold_context": gold_state["context"],
                "pred_context": pred_state["context"],
                "gold_transparent_gallery_wrapper_count": gold_state.get(
                    "transparent_gallery_wrapper_count", 0
                ),
                "pred_transparent_gallery_wrapper_count": pred_state.get(
                    "transparent_gallery_wrapper_count", 0
                ),
                "gold_transparent_native_equation_wrapper_count": gold_state.get(
                    "transparent_native_equation_wrapper_count", 0
                ),
                "pred_transparent_native_equation_wrapper_count": pred_state.get(
                    "transparent_native_equation_wrapper_count", 0
                ),
            },
        )

    return VerificationResult(
        passed=True,
        message="PASS: Saved equation semantics, anchor, and context match",
        details={
            "match_mode": "saved_ooxml_equation_semantics+context",
            "equations": pred_equations,
            "gold_transparent_gallery_wrapper_count": gold_state.get(
                "transparent_gallery_wrapper_count", 0
            ),
            "pred_transparent_gallery_wrapper_count": pred_state.get(
                "transparent_gallery_wrapper_count", 0
            ),
            "gold_transparent_native_equation_wrapper_count": gold_state.get(
                "transparent_native_equation_wrapper_count", 0
            ),
            "pred_transparent_native_equation_wrapper_count": pred_state.get(
                "transparent_native_equation_wrapper_count", 0
            ),
            "fallback_contract": "failure_only",
        },
    )
