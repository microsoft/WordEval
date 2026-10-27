"""policies / content_operations with explicit evidence dependencies."""

from evaluation.methods.base import (
    VerificationResult,
    combine_com_primary_with_ooxml_guardrail,
    tag_verification_source,
)
from evaluation.methods.word_object.core import DocumentOpenError
from typing import Any, Dict, Optional
import re


def error_with_details(
    error_message: str, details: Optional[Dict[str, Any]] = None
) -> VerificationResult:
    """Create an infrastructure/gold-evidence error without discarding context."""
    return VerificationResult(
        passed=False,
        message=f"ERROR: {error_message}",
        details=details or {},
        error=error_message,
    )


def saved_result_after_com_failure(saved_result: VerificationResult) -> VerificationResult:
    """Use incomplete saved evidence as failure-only after COM is unavailable."""
    if not saved_result.passed or saved_result.error:
        return saved_result
    error = (
        "Saved-state target match does not establish visible preservation " "after Word COM failure"
    )
    return VerificationResult(
        passed=False,
        message=f"ERROR: {error}",
        details={
            **(saved_result.details or {}),
            "fallback_policy": "saved_target_match_is_not_pass_evidence",
            "saved_target_result": saved_result.to_dict(),
        },
        error=error,
    )


def com_primary_with_saved_guardrail(
    com_verify,
    saved_verify,
    *,
    saved_pass_after_com_failure: bool = False,
    saved_pass_error_pattern: Optional[str] = None,
):
    """Run WOM first; use saved OOXML as guardrail or true fallback."""
    try:
        com_result = com_verify()
    except DocumentOpenError:
        raise
    except Exception as com_error:
        saved_result = saved_verify()
        error_matches = (
            saved_pass_error_pattern is None
            or re.search(saved_pass_error_pattern, str(com_error), re.IGNORECASE) is not None
        )
        if not saved_pass_after_com_failure or not error_matches:
            saved_result = saved_result_after_com_failure(saved_result)
        return tag_verification_source(
            saved_result,
            "saved_ooxml_fallback",
            word_com_error=str(com_error),
        )
    if not com_result.passed:
        return tag_verification_source(com_result, "word_com")
    return combine_com_primary_with_ooxml_guardrail(com_result, saved_verify())
