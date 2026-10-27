"""families / metadata / effects with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from pathlib import Path
from typing import Dict, Optional


def verify_docinfo(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, read_header_footer_content
) -> VerificationResult:
    """Verify DocInfo field/content control inserted in header."""
    gold_hf = read_header_footer_content(str(gold_path))
    pred_hf = read_header_footer_content(str(pred_path))

    idMso = spec.get("idMso", "") if spec else ""

    gold_headers = gold_hf.get("headers", [])
    pred_headers = pred_hf.get("headers", [])

    if not gold_headers:
        return VerificationResult(
            passed=False,
            message="FAIL: Gold document has no headers",
            details={"gold": gold_hf, "pred": pred_hf},
        )

    gold_h = gold_headers[0]
    pred_h = pred_headers[0] if pred_headers else {}

    if idMso == "DocInfoFileName":
        # Verify FILENAME field (Type=29) exists in header
        gold_field_types = gold_h.get("field_types", [])
        pred_field_types = pred_h.get("field_types", [])
        if 29 not in gold_field_types:
            return VerificationResult(
                passed=False,
                message="FAIL: Gold header missing FILENAME field",
                details={"gold_header": gold_h, "pred_header": pred_h},
            )
        if 29 not in pred_field_types:
            return VerificationResult(
                passed=False,
                message="FAIL: Pred header missing FILENAME field (Type=29)",
                details={"gold_header": gold_h, "pred_header": pred_h},
            )
        return VerificationResult(
            passed=True,
            message="PASS: FILENAME field present in header",
            details={"gold_header": gold_h, "pred_header": pred_h},
        )

    # DocInfoAuthor, DocInfoDocTitle - verify content control in header
    gold_cc = gold_h.get("content_control_values", {})
    pred_cc = pred_h.get("content_control_values", {})

    expected_key = "Author" if idMso == "DocInfoAuthor" else "Title"

    if expected_key not in gold_cc:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Gold header missing '{expected_key}' content control",
            details={"gold_cc": gold_cc, "pred_cc": pred_cc},
        )

    if expected_key not in pred_cc:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Pred header missing '{expected_key}' content control",
            details={"gold_cc": gold_cc, "pred_cc": pred_cc},
        )

    gold_val = gold_cc[expected_key]
    pred_val = pred_cc[expected_key]
    if gold_val != pred_val:
        return VerificationResult(
            passed=False,
            message=f"FAIL: '{expected_key}' content control value mismatch (gold='{gold_val}', pred='{pred_val}')",
            details={"gold_cc": gold_cc, "pred_cc": pred_cc},
        )

    return VerificationResult(
        passed=True,
        message=f"PASS: '{expected_key}' content control matches in header",
        details={"gold_cc": gold_cc, "pred_cc": pred_cc},
    )
