"""families / comments / verify with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.composition import composition_command_ids
from evaluation.methods.word_object.families.comments import compare as families_comments_compare
from evaluation.methods.word_object.families.comments import (
    extract_saved as families_comments_extract_saved,
)
from evaluation.methods.word_object.policies import (
    content_operations as policies_content_operations,
)
from pathlib import Path
from typing import Any, Dict, Optional


def comment_operation_com(
    pred_path: Path, gold_path: Path, operation: str, *, doc_in_path, read_comment_info, **kwargs
) -> VerificationResult:
    pred_info = read_comment_info(str(pred_path))
    gold_info = read_comment_info(str(gold_path))
    input_info = read_comment_info(str(doc_in_path)) if doc_in_path else None
    return families_comments_compare.comment_operation_result(
        pred_info,
        gold_info,
        input_info,
        operation,
        match_mode="word_com_comment_threads",
        **kwargs,
    )


def verify_saved_comment_context(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, doc_in_path
) -> Optional[VerificationResult]:
    if len(composition_command_ids(spec)) > 1:
        # Other declared composition sub-verifiers own their body,
        # revision, header/footer, and object mutations.  Comment thread
        # semantics are still checked before this preservation guardrail.
        return None
    if doc_in_path is None:
        return policies_content_operations.error_with_details(
            "Input document is required for comment preservation verification"
        )
    input_context = families_comments_extract_saved.extract_saved_comment_context(doc_in_path)
    gold_context = families_comments_extract_saved.extract_saved_comment_context(gold_path)
    pred_context = families_comments_extract_saved.extract_saved_comment_context(pred_path)
    if gold_context != input_context:
        return policies_content_operations.error_with_details(
            "Gold changed content/layout outside the requested comment operation",
            {
                "match_mode": "saved_ooxml_comment_preservation",
                "input_context": input_context,
                "gold_context": gold_context,
            },
        )
    if pred_context != gold_context:
        return VerificationResult(
            passed=False,
            message="FAIL: Comment operation changed unrelated content/layout",
            details={
                "match_mode": "saved_ooxml_comment_preservation",
                "gold_context": gold_context,
                "pred_context": pred_context,
            },
        )
    return None


def comment_operation_saved(
    pred_path: Path,
    gold_path: Path,
    operation: str,
    *,
    spec: Optional[Dict] = None,
    doc_in_path,
    **kwargs,
) -> VerificationResult:
    pred_info = families_comments_extract_saved.extract_comment_info_ooxml(pred_path)
    gold_info = families_comments_extract_saved.extract_comment_info_ooxml(gold_path)
    input_info = (
        families_comments_extract_saved.extract_comment_info_ooxml(doc_in_path)
        if doc_in_path
        else None
    )
    if gold_info.get("errors"):
        return policies_content_operations.error_with_details(
            "Gold saved comment structure is invalid",
            {"errors": gold_info["errors"], "gold": gold_info},
        )
    if input_info is not None and input_info.get("errors"):
        return policies_content_operations.error_with_details(
            "Input saved comment structure is invalid",
            {"errors": input_info["errors"], "input": input_info},
        )
    if pred_info.get("errors"):
        return VerificationResult(
            passed=False,
            message="FAIL: Predicted saved comment structure is invalid",
            details={"errors": pred_info["errors"], "pred": pred_info},
        )
    result = families_comments_compare.comment_operation_result(
        pred_info,
        gold_info,
        input_info,
        operation,
        match_mode="saved_ooxml_comment_threads",
        **kwargs,
    )
    if not result.passed:
        return result
    if operation in {"insert_comment", "insert_reply", "delete_resolved"}:
        gold_orphans = list(gold_info.get("orphan_comment_residue") or [])
        pred_orphans = list(pred_info.get("orphan_comment_residue") or [])
        if pred_orphans != gold_orphans:
            # These operations must not receive credit by converting an
            # expected effective comment into unreachable package residue,
            # or by adding unrelated orphan definitions.  Target/all-comment
            # deletion may intentionally leave Word-cleanable residue and
            # is guarded by the effective thread delta above.
            return VerificationResult(
                passed=False,
                message=(
                    "FAIL: Predicted orphan comment residue differs from " f"gold for {operation}"
                ),
                details={
                    "match_mode": "saved_ooxml_comment_orphan_residue",
                    "operation": operation,
                    "gold_orphan_comment_residue": gold_orphans,
                    "pred_orphan_comment_residue": pred_orphans,
                },
            )
    context_result = verify_saved_comment_context(
        pred_path, gold_path, spec, doc_in_path=doc_in_path
    )
    return context_result or result


def verify_comments(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, doc_in_path, read_comment_info
) -> VerificationResult:
    """Verify the requested deletion by semantic thread delta."""
    operation = (
        "delete_target" if (spec or {}).get("idMso") == "ReviewDeleteComment" else "delete_all"
    )
    return policies_content_operations.com_primary_with_saved_guardrail(
        lambda: comment_operation_com(
            pred_path,
            gold_path,
            operation,
            doc_in_path=doc_in_path,
            read_comment_info=read_comment_info,
        ),
        lambda: comment_operation_saved(
            pred_path, gold_path, operation, spec=spec, doc_in_path=doc_in_path
        ),
    )


def verify_comments_saved(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, doc_in_path
) -> VerificationResult:
    operation = (
        "delete_target" if (spec or {}).get("idMso") == "ReviewDeleteComment" else "delete_all"
    )
    return comment_operation_saved(
        pred_path, gold_path, operation, spec=spec, doc_in_path=doc_in_path
    )


def verify_comment_text(
    pred_path: Path,
    gold_path: Path,
    expected_text: str,
    spec: Optional[Dict] = None,
    *,
    doc_in_path,
    read_comment_info,
) -> VerificationResult:
    return policies_content_operations.com_primary_with_saved_guardrail(
        lambda: comment_operation_com(
            pred_path,
            gold_path,
            "insert_comment",
            expected_text=expected_text,
            doc_in_path=doc_in_path,
            read_comment_info=read_comment_info,
        ),
        lambda: verify_comment_text_saved(
            pred_path, gold_path, expected_text, spec, doc_in_path=doc_in_path
        ),
    )


def verify_comment_text_saved(
    pred_path: Path,
    gold_path: Path,
    expected_text: str,
    spec: Optional[Dict] = None,
    *,
    doc_in_path,
) -> VerificationResult:
    return comment_operation_saved(
        pred_path,
        gold_path,
        "insert_comment",
        spec=spec,
        expected_text=expected_text,
        doc_in_path=doc_in_path,
    )


def infer_comment_expectation_from_gold(gold_path: Path, *, read_comment_info) -> Dict:
    """Infer missing comment expectations from gold metadata."""
    gold_info = read_comment_info(str(gold_path))
    replies = [c for c in gold_info["comments"] if c.get("is_reply")]
    if replies:
        reply = replies[0]
        return {
            "mode": "reply",
            "expected_text": reply.get("text", ""),
            "expected_author": reply.get("author") or None,
        }

    comments = [c for c in gold_info["comments"] if not c.get("is_reply")]
    if comments:
        comment = comments[0]
        return {
            "mode": "comment",
            "expected_text": comment.get("text", ""),
            "expected_author": comment.get("author") or None,
        }

    return {}


def infer_comment_expectation_from_saved_gold(gold_path: Path) -> Dict[str, Any]:
    gold_info = families_comments_extract_saved.extract_comment_info_ooxml(gold_path)
    replies = [comment for comment in gold_info["comments"] if comment.get("is_reply")]
    if replies:
        reply = replies[0]
        return {
            "mode": "reply",
            "expected_text": reply.get("text", ""),
            "expected_author": reply.get("author") or None,
        }
    comments = [comment for comment in gold_info["comments"] if not comment.get("is_reply")]
    if comments:
        return {
            "mode": "comment",
            "expected_text": comments[0].get("text", ""),
            "expected_author": comments[0].get("author") or None,
        }
    return {}


def verify_comment_reply(
    pred_path: Path,
    gold_path: Path,
    expected_text: str,
    expected_author: Optional[str] = None,
    spec: Optional[Dict] = None,
    *,
    doc_in_path,
    read_comment_info,
) -> VerificationResult:
    return policies_content_operations.com_primary_with_saved_guardrail(
        lambda: comment_operation_com(
            pred_path,
            gold_path,
            "insert_reply",
            expected_text=expected_text,
            expected_author=expected_author,
            doc_in_path=doc_in_path,
            read_comment_info=read_comment_info,
        ),
        lambda: verify_comment_reply_saved(
            pred_path, gold_path, expected_text, expected_author, spec, doc_in_path=doc_in_path
        ),
    )


def verify_comment_reply_saved(
    pred_path: Path,
    gold_path: Path,
    expected_text: str,
    expected_author: Optional[str] = None,
    spec: Optional[Dict] = None,
    *,
    doc_in_path,
) -> VerificationResult:
    return comment_operation_saved(
        pred_path,
        gold_path,
        "insert_reply",
        spec=spec,
        expected_text=expected_text,
        expected_author=expected_author,
        doc_in_path=doc_in_path,
    )


def verify_resolved_comments_deleted(
    pred_path: Path,
    gold_path: Path,
    expected_count: Optional[int] = None,
    spec: Optional[Dict] = None,
    *,
    doc_in_path,
    read_comment_info,
) -> VerificationResult:
    return policies_content_operations.com_primary_with_saved_guardrail(
        lambda: verify_resolved_comments_deleted_com(
            pred_path,
            gold_path,
            expected_count,
            doc_in_path=doc_in_path,
            read_comment_info=read_comment_info,
        ),
        lambda: verify_resolved_comments_deleted_saved(
            pred_path, gold_path, expected_count, spec, doc_in_path=doc_in_path
        ),
        saved_pass_after_com_failure=False,
    )


def verify_resolved_comments_deleted_com(
    pred_path: Path,
    gold_path: Path,
    expected_count: Optional[int] = None,
    *,
    doc_in_path,
    read_comment_info,
) -> VerificationResult:
    return comment_operation_com(
        pred_path,
        gold_path,
        "delete_resolved",
        expected_count=expected_count,
        doc_in_path=doc_in_path,
        read_comment_info=read_comment_info,
    )


def verify_resolved_comments_deleted_saved(
    pred_path: Path,
    gold_path: Path,
    expected_count: Optional[int] = None,
    spec: Optional[Dict] = None,
    *,
    doc_in_path,
) -> VerificationResult:
    return comment_operation_saved(
        pred_path,
        gold_path,
        "delete_resolved",
        spec=spec,
        expected_count=expected_count,
        doc_in_path=doc_in_path,
    )
