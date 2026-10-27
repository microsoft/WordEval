"""families / comments / compare with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.policies import (
    content_operations as policies_content_operations,
)
from typing import Any, Dict, List, Optional
import re


def normalized_comment_text(value: Optional[str]) -> str:
    return (value or "").replace("\r", "").replace("\n", "\n").strip()


def normalized_comment_scope(value: Optional[str]) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def comment_semantic_entries(info: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Return ID-agnostic comment/thread semantics from WOM or OOXML data."""
    comments = list(info.get("comments") or [])
    top_level = [comment for comment in comments if not comment.get("is_reply")]
    by_index = {
        comment.get("index"): comment for comment in top_level if comment.get("index") is not None
    }
    by_collection_index = {
        comment.get("collection_index"): comment
        for comment in top_level
        if comment.get("collection_index") is not None
    }
    by_range_start = {
        comment.get("range_start"): comment
        for comment in top_level
        if comment.get("range_start") is not None
    }
    by_para_id = {
        comment.get("para_id"): comment for comment in top_level if comment.get("para_id")
    }

    entries = []
    for comment in comments:
        is_reply = bool(comment.get("is_reply"))
        parent = None
        if is_reply:
            parent = by_para_id.get(comment.get("parent_para_id"))
            if parent is None:
                parent = by_range_start.get(comment.get("ancestor_range_start"))
            if parent is None:
                parent = by_index.get(comment.get("ancestor_index"))
            if parent is None:
                parent = by_collection_index.get(comment.get("ancestor_index"))

        scope = normalized_comment_scope(comment.get("scope_text"))
        if is_reply and not scope and parent is not None:
            scope = normalized_comment_scope(parent.get("scope_text"))
        entries.append(
            {
                "text": normalized_comment_text(comment.get("text")),
                "author": normalized_comment_text(comment.get("author")),
                "scope_text": scope,
                "is_reply": is_reply,
                "is_resolved": bool(comment.get("is_resolved")) if not is_reply else False,
                "parent_text": (
                    normalized_comment_text(parent.get("text")) if parent is not None else ""
                ),
                "parent_author": (
                    normalized_comment_text(parent.get("author")) if parent is not None else ""
                ),
                "parent_scope_text": (
                    normalized_comment_scope(parent.get("scope_text")) if parent is not None else ""
                ),
                "parent_missing": is_reply and parent is None,
            }
        )

    return sorted(
        entries,
        key=lambda entry: (
            entry["is_reply"],
            entry["parent_scope_text"],
            entry["parent_text"],
            entry["parent_author"],
            entry["scope_text"],
            entry["text"],
            entry["author"],
            entry["is_resolved"],
            entry["parent_missing"],
        ),
    )


def comment_parent_signature(entry: Dict[str, Any]):
    return (
        entry.get("text", ""),
        entry.get("author", ""),
        entry.get("scope_text", ""),
    )


def remove_comment_thread(
    entries: List[Dict[str, Any]], parent: Dict[str, Any]
) -> List[Dict[str, Any]]:
    parent_signature = comment_parent_signature(parent)
    removed_parent = False
    remaining = []
    for entry in entries:
        if not removed_parent and not entry["is_reply"] and entry == parent:
            removed_parent = True
            continue
        if (
            entry["is_reply"]
            and (
                entry.get("parent_text", ""),
                entry.get("parent_author", ""),
                entry.get("parent_scope_text", ""),
            )
            == parent_signature
        ):
            continue
        remaining.append(entry)
    return remaining


def find_single_added_comment(
    baseline: List[Dict[str, Any]], candidate: List[Dict[str, Any]], predicate
) -> Optional[Dict[str, Any]]:
    matches = []
    for index, entry in enumerate(candidate):
        if not predicate(entry):
            continue
        if candidate[:index] + candidate[index + 1 :] == baseline:
            matches.append(entry)
    return matches[0] if len(matches) == 1 else None


def comment_operation_result(
    pred_info: Dict[str, Any],
    gold_info: Dict[str, Any],
    input_info: Optional[Dict[str, Any]],
    operation: str,
    *,
    expected_text: str = "",
    expected_author: Optional[str] = None,
    expected_count: Optional[int] = None,
    match_mode: str,
) -> VerificationResult:
    if input_info is None:
        return policies_content_operations.error_with_details(
            "Input document is required for comment delta verification",
            {"operation": operation, "match_mode": match_mode},
        )

    input_entries = comment_semantic_entries(input_info)
    gold_entries = comment_semantic_entries(gold_info)
    pred_entries = comment_semantic_entries(pred_info)
    expected_entries: Optional[List[Dict[str, Any]]] = None
    gold_target = None
    pred_target = None

    if operation == "delete_all":
        if not input_entries:
            return policies_content_operations.error_with_details(
                "Gold/input contains no comment target for delete-all",
                {"input": input_entries, "gold": gold_entries},
            )
        expected_entries = []
    elif operation == "delete_target":
        targets = [
            entry
            for entry in input_entries
            if not entry["is_reply"] and entry["scope_text"] == "TARGET"
        ]
        if len(targets) != 1:
            return policies_content_operations.error_with_details(
                "Input must contain exactly one top-level TARGET comment thread",
                {"target_count": len(targets), "input": input_entries},
            )
        expected_entries = remove_comment_thread(input_entries, targets[0])
    elif operation == "delete_resolved":
        resolved = [
            entry for entry in input_entries if not entry["is_reply"] and entry["is_resolved"]
        ]
        if not resolved:
            return policies_content_operations.error_with_details(
                "Input contains no resolved comment thread",
                {"input": input_entries},
            )
        expected_entries = list(input_entries)
        for target in resolved:
            expected_entries = remove_comment_thread(expected_entries, target)
        if expected_count is not None and len(expected_entries) != int(expected_count):
            return policies_content_operations.error_with_details(
                "Configured resolved-comment expected_count disagrees with input delta",
                {
                    "configured_expected_count": expected_count,
                    "derived_expected_count": len(expected_entries),
                },
            )
    elif operation == "insert_comment":
        expected = normalized_comment_text(expected_text)
        predicate = lambda entry: (
            not entry["is_reply"] and entry["text"] == expected and entry["scope_text"] == "TARGET"
        )
        gold_target = find_single_added_comment(input_entries, gold_entries, predicate)
        if gold_target is None:
            return policies_content_operations.error_with_details(
                "Gold does not add exactly one expected top-level TARGET comment",
                {
                    "expected_text": expected,
                    "input": input_entries,
                    "gold": gold_entries,
                },
            )
        pred_target = find_single_added_comment(input_entries, pred_entries, predicate)
        if pred_target is None:
            return VerificationResult(
                passed=False,
                message="FAIL: Predicted document does not add exactly one expected TARGET comment",
                details={
                    "match_mode": match_mode,
                    "expected_text": expected,
                    "expected_scope": "TARGET",
                    "input": input_entries,
                    "gold": gold_entries,
                    "pred": pred_entries,
                },
            )
    elif operation == "insert_reply":
        expected = normalized_comment_text(expected_text)
        author = normalized_comment_text(expected_author)
        predicate = lambda entry: (
            entry["is_reply"]
            and not entry["parent_missing"]
            and entry["text"] == expected
            and entry["author"] == author
            and entry["parent_scope_text"] == "TARGET"
        )
        gold_target = find_single_added_comment(input_entries, gold_entries, predicate)
        if gold_target is None:
            return policies_content_operations.error_with_details(
                "Gold does not add exactly one expected reply under the TARGET thread",
                {
                    "expected_text": expected,
                    "expected_author": author,
                    "input": input_entries,
                    "gold": gold_entries,
                },
            )
        pred_target = find_single_added_comment(input_entries, pred_entries, predicate)
        if pred_target is None:
            return VerificationResult(
                passed=False,
                message="FAIL: Expected reply is missing, duplicated, or attached to the wrong parent",
                details={
                    "match_mode": match_mode,
                    "expected_text": expected,
                    "expected_author": author,
                    "expected_parent_scope": "TARGET",
                    "input": input_entries,
                    "gold": gold_entries,
                    "pred": pred_entries,
                },
            )
    else:
        return policies_content_operations.error_with_details(
            f"Unsupported comment operation: {operation}",
            {"match_mode": match_mode},
        )

    if expected_entries is not None:
        if gold_entries != expected_entries:
            return policies_content_operations.error_with_details(
                f"Gold comment delta is invalid for {operation}",
                {
                    "match_mode": match_mode,
                    "input": input_entries,
                    "expected": expected_entries,
                    "gold": gold_entries,
                },
            )
        if pred_entries != expected_entries:
            return VerificationResult(
                passed=False,
                message=f"FAIL: Predicted comment threads differ for {operation}",
                details={
                    "match_mode": match_mode,
                    "input": input_entries,
                    "expected": expected_entries,
                    "gold": gold_entries,
                    "pred": pred_entries,
                },
            )

    return VerificationResult(
        passed=True,
        message=f"PASS: Comment operation semantics match ({operation})",
        details={
            "match_mode": match_mode,
            "operation": operation,
            "input": input_entries,
            "gold": gold_entries,
            "pred": pred_entries,
            "gold_target": gold_target,
            "pred_target": pred_target,
            "input_orphan_comment_residue": list(
                (input_info or {}).get("orphan_comment_residue") or []
            ),
            "gold_orphan_comment_residue": list(gold_info.get("orphan_comment_residue") or []),
            "pred_orphan_comment_residue": list(pred_info.get("orphan_comment_residue") or []),
        },
    )


def resolved_comments_deleted_result(
    pred_info: Dict[str, Any],
    gold_info: Dict[str, Any],
    input_info: Optional[Dict[str, Any]],
    expected_count: Optional[int],
) -> VerificationResult:
    resolved_comments = [c for c in pred_info["comments"] if c["is_resolved"]]
    if resolved_comments:
        return VerificationResult(
            passed=False,
            message=f"FAIL: {len(resolved_comments)} resolved comment(s) still remain",
            details={
                "resolved_remaining": len(resolved_comments),
                "resolved_comments": resolved_comments,
                "total_comments": pred_info["comment_count"],
            },
        )

    # Determine expected count
    if expected_count is None and input_info:
        # Calculate expected: input_count - resolved_count
        expected_count = input_info["comment_count"] - input_info["resolved_count"]

    # Check count if expected_count is provided or calculated
    if expected_count is not None:
        if pred_info["comment_count"] != expected_count:
            return VerificationResult(
                passed=False,
                message=f"FAIL: Comment count mismatch (pred={pred_info['comment_count']}, expected={expected_count})",
                details={
                    "pred_count": pred_info["comment_count"],
                    "expected_count": expected_count,
                    "input_info": input_info,
                },
            )
    else:
        # Compare with gold
        if pred_info["comment_count"] != gold_info["comment_count"]:
            return VerificationResult(
                passed=False,
                message=f"FAIL: Comment count mismatch with gold (pred={pred_info['comment_count']}, gold={gold_info['comment_count']})",
                details={
                    "pred_count": pred_info["comment_count"],
                    "gold_count": gold_info["comment_count"],
                },
            )

    return VerificationResult(
        passed=True,
        message=f"PASS: All resolved comments deleted ({pred_info['comment_count']} unresolved remain)",
        details={
            "comment_count": pred_info["comment_count"],
            "resolved_count": 0,
            "input_info": input_info,
        },
    )
