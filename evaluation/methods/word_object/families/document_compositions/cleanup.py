"""families / document_compositions / cleanup with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.families.comments import verify as families_comments_verify
from evaluation.methods.word_object.families.document_compositions import (
    extract_saved as families_document_compositions_extract_saved,
)
from evaluation.methods.word_object.families.headers import verify as families_headers_verify
from evaluation.methods.word_object.families.revisions import verify as families_revisions_verify
from evaluation.methods.word_object.policies import (
    content_operations as policies_content_operations,
)
from pathlib import Path
from typing import Dict, Optional


def verify_document_cleanup_composition(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    add_composition_atomic_details,
    doc_in_path,
    read_comment_info,
    read_document_text,
    read_header_footer_content,
    read_revision_info,
    summarize_sub_result,
) -> VerificationResult:
    return policies_content_operations.com_primary_with_saved_guardrail(
        lambda: verify_document_cleanup_com(
            pred_path,
            gold_path,
            spec,
            add_composition_atomic_details=add_composition_atomic_details,
            doc_in_path=doc_in_path,
            read_comment_info=read_comment_info,
            read_document_text=read_document_text,
            read_header_footer_content=read_header_footer_content,
            read_revision_info=read_revision_info,
            summarize_sub_result=summarize_sub_result,
        ),
        lambda: verify_document_cleanup_saved(pred_path, gold_path, spec),
        saved_pass_after_com_failure=False,
    )


def verify_document_cleanup_saved(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None
) -> VerificationResult:
    params = spec.get("verification", {}).get("primary", {}).get("params", {}) if spec else {}
    operations = params.get("operations", [])
    gold = families_document_compositions_extract_saved.extract_saved_cleanup_signature(gold_path)
    pred = families_document_compositions_extract_saved.extract_saved_cleanup_signature(pred_path)
    failures = []
    supported_operations = {
        "accept_changes",
        "reject_changes",
        "delete_comments",
        "delete_resolved_comments",
        "remove_header",
        "remove_watermark",
        "remove_footer",
    }
    unsupported_operations = sorted(set(operations) - supported_operations)
    revision_operation = next(
        (
            operation
            for operation in ("accept_changes", "reject_changes")
            if operation in operations
        ),
        None,
    )
    gold_errors = []
    if revision_operation:
        if gold["revision_count"]:
            gold_errors.append("gold retains tracked revisions")
        if gold["track_revisions"]:
            gold_errors.append("gold retains track revisions enabled")
        if not (gold["body_text"] or "").strip():
            gold_errors.append("gold final body text is empty")
    if "delete_comments" in operations and (gold["comment_count"] or gold["comment_marker_count"]):
        gold_errors.append("gold retains comments or comment anchors")
    if "remove_header" in operations and gold.get("header_payloads"):
        gold_errors.append("gold retains header payload")
    if "remove_watermark" in operations and gold["watermark_count"]:
        gold_errors.append("gold retains watermark")
    if "remove_footer" in operations and gold.get("footer_payloads"):
        gold_errors.append("gold retains footer payload")
    if gold_errors:
        return policies_content_operations.error_with_details(
            "Gold saved document cleanup evidence is missing or invalid",
            {
                "match_mode": "saved_ooxml_cleanup",
                "errors": gold_errors,
                "gold": gold,
            },
        )
    if revision_operation:
        if pred["revision_count"]:
            failures.append("tracked revisions remain")
        if pred["track_revisions"]:
            failures.append("track revisions remains enabled")
        if pred["body_text"] != gold["body_text"]:
            action = "accepted" if revision_operation == "accept_changes" else "rejected"
            failures.append(f"{action} body text differs from gold")
    if "delete_comments" in operations:
        if pred["comment_count"] or pred["comment_marker_count"]:
            failures.append("comments or comment anchors remain")
    if "delete_resolved_comments" in operations:
        if (
            pred["comment_count"] != gold["comment_count"]
            or pred["comment_marker_count"] != gold["comment_marker_count"]
        ):
            failures.append("resolved-comment cleanup differs from gold")
    if "remove_header" in operations and pred.get("header_payloads"):
        failures.append("header content remains")
    if "remove_watermark" in operations and pred["watermark_count"]:
        failures.append("watermark remains")
    if "remove_footer" in operations and pred["footer_payloads"]:
        failures.append("footer content remains")
    fallback_error = None
    if not failures and unsupported_operations:
        fallback_error = (
            "Saved-state cleanup fallback cannot verify requested operation(s): "
            + ", ".join(unsupported_operations)
        )
    return VerificationResult(
        passed=not failures and not unsupported_operations,
        message=(
            f"ERROR: {fallback_error}"
            if fallback_error
            else (
                f"PASS: Saved document cleanup ({len(operations)} operations)"
                if not failures
                else f"FAIL: Saved document cleanup ({len(failures)} failed check(s))"
            )
        ),
        details={
            "match_mode": "saved_ooxml_cleanup",
            "failures": failures,
            "unsupported_operations": unsupported_operations,
            "gold": gold,
            "pred": pred,
        },
        error=fallback_error,
    )


def verify_document_cleanup_com(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    add_composition_atomic_details,
    doc_in_path,
    read_comment_info,
    read_document_text,
    read_header_footer_content,
    read_revision_info,
    summarize_sub_result,
) -> VerificationResult:
    """Verify cleanup operations through the Word Object Model."""
    params = spec.get("verification", {}).get("primary", {}).get("params", {}) if spec else {}
    operations = params.get("operations", [])

    if gold_path.exists():
        gold_saved_probe = verify_document_cleanup_saved(gold_path, gold_path, spec)
        if gold_saved_probe.error:
            return gold_saved_probe

    sub_results = {}
    failures = []
    errors = []

    # Verify tracked changes resolved
    if "accept_changes" in operations or "reject_changes" in operations:
        rev_result = families_revisions_verify.verify_revisions(
            pred_path,
            gold_path,
            spec,
            doc_in_path=doc_in_path,
            read_document_text=read_document_text,
            read_revision_info=read_revision_info,
        )
        sub_results["revisions"] = summarize_sub_result(rev_result)
        if rev_result.error:
            errors.append(rev_result.error)
        elif not rev_result.passed:
            failures.append(rev_result.message)

    # Verify comments deleted
    if "delete_comments" in operations:
        cmt_result = families_comments_verify.verify_comments(
            pred_path, gold_path, spec, doc_in_path=doc_in_path, read_comment_info=read_comment_info
        )
        sub_results["comments"] = summarize_sub_result(cmt_result)
        if cmt_result.error:
            errors.append(cmt_result.error)
        elif not cmt_result.passed:
            failures.append(cmt_result.message)

    # Verify resolved comments deleted (but normal comments remain)
    if "delete_resolved_comments" in operations:
        rcmt_result = families_comments_verify.verify_resolved_comments_deleted(
            pred_path,
            gold_path,
            spec=spec,
            doc_in_path=doc_in_path,
            read_comment_info=read_comment_info,
        )
        sub_results["resolved_comments"] = summarize_sub_result(rcmt_result)
        if rcmt_result.error:
            errors.append(rcmt_result.error)
        elif not rcmt_result.passed:
            failures.append(rcmt_result.message)

    # Verify header/footer removed or watermark removed
    if any(op in operations for op in ["remove_header", "remove_footer", "remove_watermark"]):
        hf_result = families_headers_verify.verify_header_footer_composition(
            pred_path, gold_path, read_header_footer_content=read_header_footer_content
        )
        sub_results["header_footer"] = summarize_sub_result(hf_result)
        if hf_result.error:
            errors.append(hf_result.error)
        elif not hf_result.passed:
            failures.append(hf_result.message)

    details = {"sub_results": sub_results}
    atomic_failures, atomic_errors = add_composition_atomic_details(
        details, pred_path, gold_path, spec
    )
    failures.extend(atomic_failures)
    errors.extend(atomic_errors)

    if errors:
        error = "; ".join(errors)
        return VerificationResult(
            passed=False,
            message=(
                "ERROR: Document cleanup composition could not be fully "
                f"evaluated ({len(errors)} error(s))"
            ),
            details={**details, "errors": errors},
            error=error,
        )

    if failures:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Document cleanup composition ({len(failures)} failed check(s))",
            details=details,
        )

    return VerificationResult(
        passed=True,
        message=f"PASS: Document cleanup composition ({len(operations)} operations)",
        details=details,
    )
