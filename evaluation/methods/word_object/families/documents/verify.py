"""families / documents / verify with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.policies import (
    content_operations as policies_content_operations,
)
from evaluation.methods.word_object.preservation import context as preservation_context
from pathlib import Path
from typing import Dict, Optional
from ..formatting.characters import validate_character_observations


def verify_body_text_matches(
    pred_path: Path, gold_path: Path, *, read_document_text
) -> VerificationResult:
    """Verify body text matches between pred and gold.

    Used after note deletion to catch leftover reference marks (e.g., [1], [2])
    that remain as inline text when models remove footnotes/endnotes incorrectly.
    """
    gold_text = read_document_text(str(gold_path))
    pred_text = read_document_text(str(pred_path))

    if gold_text is None:
        return policies_content_operations.error_with_details(
            "Could not extract gold document body text",
            {"gold_extracted": False, "pred_extracted": pred_text is not None},
        )
    if pred_text is None:
        return policies_content_operations.error_with_details(
            "Could not extract predicted document body text",
            {"gold_extracted": True, "pred_extracted": False},
        )

    if gold_text != pred_text:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Body text differs (possible leftover reference marks)",
            details={
                "gold_text": gold_text[:500],
                "pred_text": pred_text[:500],
                "preservation_scope": "document_text",
            },
        )

    return VerificationResult(passed=True, message="PASS: Body text matches", details={})


def verify_text_content(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    read_all_paragraph_character_formatting,
    read_document_text,
) -> VerificationResult:
    """Verify document text content matches gold (used for symbol insertion, etc.)."""
    gold_text = read_document_text(str(gold_path)).strip()
    pred_text = read_document_text(str(pred_path)).strip()

    parameter = (spec or {}).get("parameter") or {}
    expected_symbol = parameter.get("char")
    expected_codepoint = parameter.get("code")
    if expected_codepoint is not None:
        try:
            expected_symbol = chr(int(expected_codepoint))
        except (TypeError, ValueError, OverflowError):
            pass
    if expected_symbol and f"TARGET{expected_symbol}" not in gold_text:
        return policies_content_operations.error_with_details(
            "Gold symbol target evidence is missing or at the wrong anchor",
            {
                "expected_symbol": expected_symbol,
                "gold_text": gold_text,
            },
        )
    if pred_text != gold_text:
        return VerificationResult(
            passed=False,
            message="FAIL: Document text content differs",
            details={
                "pred_text": pred_text[:200] + "..." if len(pred_text) > 200 else pred_text,
                "gold_text": gold_text[:200] + "..." if len(gold_text) > 200 else gold_text,
            },
        )

    gold_format = read_all_paragraph_character_formatting(str(gold_path))
    pred_format = read_all_paragraph_character_formatting(str(pred_path))
    error = validate_character_observations({"gold": gold_format, "pred": pred_format})
    if error is not None:
        return error
    if not gold_format:
        return policies_content_operations.error_with_details(
            "Gold effective character-format evidence is missing",
            {"match_mode": "word_com_symbol_effective_format"},
        )
    if pred_format != gold_format:
        return VerificationResult(
            passed=False,
            message="FAIL: Symbol insertion changed effective character formatting",
            details={
                "match_mode": "word_com_symbol_effective_format",
                "gold_format": gold_format,
                "pred_format": pred_format,
            },
        )

    context_result = preservation_context.verify_saved_insertion_context(
        pred_path,
        gold_path,
        ignore_body_visible_format=True,
    )
    if context_result is not None:
        context_result.message = (
            "FAIL: Symbol text matches but visible format/content layout differs"
        )
        return context_result

    return VerificationResult(
        passed=True,
        message="PASS: Symbol text, anchor, and visible formatting match",
        details={
            "text_length": len(gold_text),
            "expected_symbol": expected_symbol,
            "match_mode": "word_com_text+saved_visible_format",
        },
    )


def verify_structure(
    pred_path: Path, gold_path: Path, *, read_document_structure
) -> VerificationResult:
    """Verify overall document structure counts match."""
    gold_struct = read_document_structure(str(gold_path))
    pred_struct = read_document_structure(str(pred_path))

    if "error" in gold_struct:
        return VerificationResult.error_result(f"Gold: {gold_struct['error']}")
    if "error" in pred_struct:
        return VerificationResult.error_result(f"Pred: {pred_struct['error']}")

    differences = []
    for key in gold_struct.keys():
        gold_val = gold_struct.get(key)
        pred_val = pred_struct.get(key)
        if gold_val != pred_val:
            differences.append({"property": key, "gold": gold_val, "pred": pred_val})

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: {len(differences)} structure difference(s)",
            details={"differences": differences, "gold": gold_struct, "pred": pred_struct},
        )

    return VerificationResult(
        passed=True, message="PASS: Document structure matches", details={"structure": gold_struct}
    )
