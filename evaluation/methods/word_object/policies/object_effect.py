"""policies / object_effect with explicit evidence dependencies."""

from evaluation.methods.base import (
    VerificationResult,
    combine_com_primary_with_ooxml_guardrail,
    tag_verification_source,
)
from evaluation.methods.word_object.core import DocumentOpenError
from typing import Optional


def saved_evidence_error(
    message: str, *, evidence_owner: str, exception: Optional[Exception] = None
) -> VerificationResult:
    """Return a typed evaluator error for unavailable saved evidence."""
    detail = {
        "verification_source": "saved_ooxml_evidence_error",
        "evidence_owner": evidence_owner,
        "error_owner": "evaluator",
    }
    if exception is not None:
        detail["exception_type"] = type(exception).__name__
    return VerificationResult(
        passed=False,
        message=f"ERROR: {message}",
        details=detail,
        error=message,
    )


def com_primary_with_saved_guardrail(
    com_verify, saved_verify, *, saved_pass_after_com_failure: bool = False
):
    """Run WOM first; use saved OOXML as guardrail or true fallback."""
    try:
        com_result = com_verify()
    except DocumentOpenError:
        raise
    except Exception as com_error:
        saved_result = saved_verify()
        if saved_result.passed and not saved_result.error and not saved_pass_after_com_failure:
            error = (
                "Saved-state target match does not establish the complete "
                "preservation contract after Word COM failure"
            )
            saved_result = VerificationResult(
                passed=False,
                message=f"ERROR: {error}",
                details={
                    **(saved_result.details or {}),
                    "saved_target_result": saved_result.to_dict(),
                    "missing_evidence": "word_com_visible_preservation",
                },
                error=error,
            )
        return tag_verification_source(
            saved_result,
            "saved_ooxml_fallback",
            word_com_error=str(com_error),
        )
    if not com_result.passed:
        return tag_verification_source(com_result, "word_com")
    return combine_com_primary_with_ooxml_guardrail(com_result, saved_verify())
