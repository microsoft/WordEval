"""policies / content_formatting with explicit evidence dependencies."""

from evaluation.methods.base import (
    VerificationResult,
    combine_com_primary_with_ooxml_guardrail,
    tag_verification_source,
)
from evaluation.methods.word_object.core import DocumentOpenError
from typing import Any


def saved_result_after_com_failure(
    saved_result: VerificationResult, com_error: Any
) -> VerificationResult:
    """Allow a saved mismatch to fail, but never let target-only evidence pass."""
    if saved_result.error or not saved_result.passed:
        return saved_result
    error_message = str(com_error)
    return VerificationResult(
        passed=False,
        message=f"ERROR: {error_message}",
        details={
            "saved_ooxml_evidence": saved_result.to_dict(),
            "missing_evidence": "word_com_visible_preservation",
        },
        error=error_message,
    )


def com_primary_with_saved_guardrail(com_verify, saved_verify):
    """Run WOM first; use saved OOXML as guardrail or true fallback."""
    try:
        com_result = com_verify()
    except DocumentOpenError:
        raise
    except Exception as com_error:
        saved_result = saved_result_after_com_failure(saved_verify(), com_error)
        return tag_verification_source(
            saved_result,
            "saved_ooxml_fallback",
            word_com_error=str(com_error),
        )
    if not com_result.passed:
        return tag_verification_source(com_result, "word_com")
    saved_result = saved_verify()
    return combine_com_primary_with_ooxml_guardrail(com_result, saved_result)
