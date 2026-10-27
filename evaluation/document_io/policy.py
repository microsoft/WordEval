"""Map document evidence to verdicts without depending on scoring methods."""
import re
from pathlib import Path
from typing import Any, Dict, Optional
from .package import validate_docx_output
from .errors import document_path_identity

def invalid_output_result(details: Dict[str, Any]) -> Dict[str, Any]:
    """Create a method-level failed result for an invalid model output file."""
    return {
        "status": "failed",
        "passed": False,
        "message": f"INVALID OUTPUT: {details.get('validation_error', 'pred.docx is not a valid OOXML .docx file.')}",
        "details": details,
        "error": None,
    }


def invalid_output_from_typed_word_open_failure(
    result: Dict[str, Any],
) -> Optional[Dict[str, Any]]:
    """Normalize prediction-owned Word open evidence without schema guessing."""

    details = result.get("details") or {}
    if details.get("open_failure_owner") != "prediction":
        return None
    evidence = details.get("document_open_error") or {}
    state = evidence.get("state")
    if state not in {"repair_required", "unopenable"}:
        return None
    repair_required = state == "repair_required"
    return invalid_output_result({
        "reason": "invalid_output",
        "detected_format": (
            "word_repair_required" if repair_required else "word_unopenable"
        ),
        "open_failure_owner": "prediction",
        "validation_error": (
            "pred.docx cannot be opened normally without repair."
            if repair_required
            else "pred.docx cannot be opened normally or with repair."
        ),
        "validation_context": "typed_document_open_error",
        "document_open_error": evidence,
        "word_open_error": result.get("error") or details.get("word_open_error"),
    })


def invalid_output_after_word_open_failure(
    path: Path,
    error: Any,
) -> Optional[Dict[str, Any]]:
    """Convert a proven unopenable schema-invalid prediction to model fail."""

    error_text = str(error or "")
    if "Document requires Word repair:" in error_text:
        reported_path = error_text.split(
            "Document requires Word repair:", 1
        )[1].strip()
        expected_path = document_path_identity(path)
        reported_path = document_path_identity(reported_path)
        if expected_path != reported_path:
            return None
        return invalid_output_result({
            "reason": "invalid_output",
            "detected_format": "word_repair_required",
            "open_failure_owner": "prediction",
            "validation_error": (
                "pred.docx cannot be opened normally; Microsoft Word only "
                "opens it after repairing the submitted file."
            ),
            "validation_context": (
                "after_word_normal_open_failure_repair_succeeded"
            ),
            "word_open_error": error_text,
        })
    match = re.match(r'^Failed to open (.+?) normally(?:\s|\()', error_text, re.DOTALL)
    if match is None:
        return None
    reported = document_path_identity(match.group(1))
    expected = document_path_identity(path)
    if reported != expected:
        return None
    try:
        details = validate_docx_output(path, include_word_open_failure_schema=True)
    except OSError:
        return None
    if details is None:
        return None
    details = dict(details)
    details["open_failure_owner"] = "prediction"
    details["validation_context"] = "after_word_normal_and_repair_open_failure"
    details["word_open_error"] = error_text
    return invalid_output_result(details)



def document_open_failure_result(error, pred_path, gold_path, doc_in_path=None):
    """Assign file ownership once, identically for every scoring adapter."""
    error_path = document_path_identity(error.docx_path)
    roles = [("prediction", pred_path), ("gold", gold_path), ("input", doc_in_path)]
    owner = next((role for role, path in roles if path is not None
                  and document_path_identity(path) == error_path), "environment_or_unknown")
    result = {
        "passed": False, "message": f"ERROR: {error}", "error": str(error),
        "details": {"open_failure_owner": owner, "document_open_error": error.to_details()},
    }
    return invalid_output_from_typed_word_open_failure(result) or result
