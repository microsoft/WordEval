"""families / revisions / verify with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.composition import composition_command_ids
from evaluation.methods.word_object.families.revisions import (
    extract_saved as families_revisions_extract_saved,
)
from evaluation.methods.word_object.policies import (
    content_operations as policies_content_operations,
)
from evaluation.methods.word_object.preservation import context as preservation_context
from pathlib import Path
from typing import Any, Dict, Optional


def normalized_revision_body_text(value: Optional[str]) -> str:
    return (value or "").replace("\r", "").replace("\x07", "").strip()


def verify_revisions_saved(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, doc_in_path
) -> VerificationResult:
    if doc_in_path is None:
        return policies_content_operations.error_with_details(
            "Input document is required for accept/reject verification"
        )
    input_signature = families_revisions_extract_saved.extract_saved_revision_signature(doc_in_path)
    gold_signature = families_revisions_extract_saved.extract_saved_revision_signature(gold_path)
    pred_signature = families_revisions_extract_saved.extract_saved_revision_signature(pred_path)
    commands = set(composition_command_ids(spec))
    operation = (
        "reject"
        if (
            "RejectAllChanges" in (spec or {}).get("idMso", "")
            or any("RejectAllChanges" in command for command in commands)
        )
        else "accept"
    )
    if input_signature["revision_count"] < 1:
        return policies_content_operations.error_with_details(
            "Input has no saved revisions to process",
            {"input": input_signature},
        )
    if gold_signature["revision_count"] != 0 or gold_signature["track_revisions"]:
        return policies_content_operations.error_with_details(
            "Gold retains revisions or tracking after processing",
            {"gold": gold_signature},
        )
    if pred_signature["revision_count"] != 0:
        return VerificationResult(
            passed=False,
            message="FAIL: Predicted saved document retains live revisions",
            details={
                "match_mode": "saved_ooxml_revision_final_state",
                "pred": pred_signature,
            },
        )
    if pred_signature["track_revisions"]:
        return VerificationResult(
            passed=False,
            message="FAIL: Predicted saved document still tracks revisions",
            details={
                "match_mode": "saved_ooxml_revision_final_state",
                "pred": pred_signature,
            },
        )
    expected_text = input_signature["rejected_text" if operation == "reject" else "accepted_text"]
    gold_text = gold_signature["accepted_text"]
    pred_text = pred_signature["accepted_text"]
    if normalized_revision_body_text(gold_text) != normalized_revision_body_text(expected_text):
        return policies_content_operations.error_with_details(
            f"Gold does not represent the {operation}ed revision view",
            {
                "operation": operation,
                "expected_text": expected_text,
                "gold_text": gold_text,
            },
        )
    if normalized_revision_body_text(pred_text) != normalized_revision_body_text(gold_text):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Predicted {operation}ed body differs from gold",
            details={
                "match_mode": "saved_ooxml_revision_final_state",
                "operation": operation,
                "gold_text": gold_text,
                "pred_text": pred_text,
            },
        )
    gold_context = preservation_context.extract_saved_semantic_context(gold_path)
    pred_context = preservation_context.extract_saved_semantic_context(pred_path)
    if pred_context != gold_context:
        return VerificationResult(
            passed=False,
            message=f"FAIL: {operation.title()} changed formatting/structure outside the final state",
            details={
                "match_mode": "saved_ooxml_revision_preservation",
                "operation": operation,
                "gold_context": gold_context,
                "pred_context": pred_context,
            },
        )
    return VerificationResult(
        passed=True,
        message=f"PASS: Saved {operation} final state and preservation match",
        details={
            "match_mode": "saved_ooxml_revision_final_state",
            "operation": operation,
            "signature": gold_signature,
        },
    )


def verify_revisions_com(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    read_document_text,
    read_revision_info,
) -> VerificationResult:
    pred_info = read_revision_info(str(pred_path))
    gold_info = read_revision_info(str(gold_path))
    if gold_info["revision_count"] != 0 or gold_info["track_revisions"]:
        return policies_content_operations.error_with_details(
            "Gold retains effective revisions or tracking", {"gold": gold_info}
        )
    if pred_info["revision_count"] != 0:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Revisions still exist (count={pred_info['revision_count']})",
            details={"pred": pred_info},
        )
    if pred_info["track_revisions"]:
        return VerificationResult(
            passed=False,
            message="FAIL: TrackRevisions is still enabled",
            details={"pred": pred_info},
        )
    pred_text = read_document_text(str(pred_path))
    gold_text = read_document_text(str(gold_path))
    if normalized_revision_body_text(pred_text) != normalized_revision_body_text(gold_text):
        return VerificationResult(
            passed=False,
            message="FAIL: Final document text differs from command-specific gold",
            details={
                "pred_text": pred_text,
                "gold_text": gold_text,
            },
        )
    return VerificationResult(
        passed=True,
        message="PASS: Revisions processed, tracking stopped, final text matches",
        details={
            "match_mode": "word_com_revision_final_state",
            "pred": pred_info,
            "gold": gold_info,
        },
    )


def verify_track_changes_saved(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, doc_in_path
) -> VerificationResult:
    if doc_in_path is None:
        return policies_content_operations.error_with_details(
            "Input document is required for Track Changes verification"
        )
    expected_text = " TRACKED_EDIT"
    input_signature = families_revisions_extract_saved.extract_saved_revision_signature(doc_in_path)
    gold_signature = families_revisions_extract_saved.extract_saved_revision_signature(gold_path)
    pred_signature = families_revisions_extract_saved.extract_saved_revision_signature(pred_path)

    def target_valid(signature: Dict[str, Any]) -> bool:
        return (
            signature.get("track_revisions") is True
            and signature.get("revision_count") == 1
            and signature.get("revision_type_counts") == {"ins": 1}
            and len(signature.get("revisions", [])) == 1
            and signature["revisions"][0].get("type") == "ins"
            and signature["revisions"][0].get("text") == expected_text
            and signature["revisions"][0].get("at_body_end") is True
        )

    if input_signature["revision_count"] != 0:
        return policies_content_operations.error_with_details(
            "Track Changes input already contains revisions",
            {"input": input_signature},
        )
    if not target_valid(gold_signature):
        return policies_content_operations.error_with_details(
            "Gold lacks the exact tracked insertion at the body end",
            {"gold": gold_signature, "expected_text": expected_text},
        )
    expected_final = input_signature["accepted_text"].rstrip() + expected_text
    if normalized_revision_body_text(
        gold_signature["accepted_text"]
    ) != normalized_revision_body_text(expected_final):
        return policies_content_operations.error_with_details(
            "Gold tracked insertion does not preserve the input body",
            {
                "input_text": input_signature["accepted_text"],
                "gold_text": gold_signature["accepted_text"],
            },
        )
    if not target_valid(pred_signature):
        return VerificationResult(
            passed=False,
            message="FAIL: Predicted saved state lacks the exact tracked insertion",
            details={
                "match_mode": "saved_ooxml_tracked_insertion",
                "expected_text": expected_text,
                "gold": gold_signature,
                "pred": pred_signature,
            },
        )
    if normalized_revision_body_text(
        pred_signature["accepted_text"]
    ) != normalized_revision_body_text(gold_signature["accepted_text"]):
        return VerificationResult(
            passed=False,
            message="FAIL: Predicted tracked final text differs",
            details={
                "gold_text": gold_signature["accepted_text"],
                "pred_text": pred_signature["accepted_text"],
            },
        )
    context_kwargs = {"strip_tracked_insertions": True}
    input_context = preservation_context.extract_saved_semantic_context(
        doc_in_path, **context_kwargs
    )
    gold_context = preservation_context.extract_saved_semantic_context(gold_path, **context_kwargs)
    pred_context = preservation_context.extract_saved_semantic_context(pred_path, **context_kwargs)
    if input_context != gold_context:
        return policies_content_operations.error_with_details(
            "Gold changes state outside the tracked insertion",
            {"input_context": input_context, "gold_context": gold_context},
        )
    if pred_context != gold_context:
        return VerificationResult(
            passed=False,
            message="FAIL: Tracked insertion changed unrelated content/layout",
            details={
                "match_mode": "saved_ooxml_tracked_insertion_preservation",
                "gold_context": gold_context,
                "pred_context": pred_context,
            },
        )
    return VerificationResult(
        passed=True,
        message="PASS: Saved tracked insertion and preservation match",
        details={
            "match_mode": "saved_ooxml_tracked_insertion",
            "signature": gold_signature,
        },
    )


def verify_track_changes_com(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    read_document_text,
    read_revision_info,
) -> VerificationResult:
    expected_text = " TRACKED_EDIT"
    pred_info = read_revision_info(str(pred_path))
    gold_info = read_revision_info(str(gold_path))

    def target_valid(info: Dict[str, Any]) -> bool:
        revisions = info.get("revisions", [])
        return (
            info.get("track_revisions") is True
            and info.get("revision_count") == 1
            and len(revisions) == 1
            and revisions[0].get("type") == 1
            and revisions[0].get("text") == expected_text
            and revisions[0].get("story_type") == 1
            and revisions[0].get("end", -1) >= info.get("body_end", 0) - 1
        )

    if not target_valid(gold_info):
        return policies_content_operations.error_with_details(
            "Gold lacks the exact effective tracked insertion",
            {"gold": gold_info},
        )
    if not target_valid(pred_info):
        return VerificationResult(
            passed=False,
            message="FAIL: Effective tracked insertion type/text/position differs",
            details={
                "match_mode": "word_com_tracked_insertion",
                "expected_text": expected_text,
                "gold": gold_info,
                "pred": pred_info,
            },
        )
    pred_text = read_document_text(str(pred_path))
    gold_text = read_document_text(str(gold_path))
    if normalized_revision_body_text(pred_text) != normalized_revision_body_text(gold_text):
        return VerificationResult(
            passed=False,
            message="FAIL: Effective tracked final text differs",
            details={"gold_text": gold_text, "pred_text": pred_text},
        )
    return VerificationResult(
        passed=True,
        message="PASS: Tracking enabled with the exact body-end insertion",
        details={
            "match_mode": "word_com_tracked_insertion",
            "pred": pred_info,
            "gold": gold_info,
        },
    )


def verify_revisions(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    doc_in_path,
    read_document_text,
    read_revision_info,
) -> VerificationResult:
    """Verify command-specific Accept/Reject final state and preservation."""
    return policies_content_operations.com_primary_with_saved_guardrail(
        lambda: verify_revisions_com(
            pred_path,
            gold_path,
            spec,
            read_document_text=read_document_text,
            read_revision_info=read_revision_info,
        ),
        lambda: verify_revisions_saved(pred_path, gold_path, spec, doc_in_path=doc_in_path),
        saved_pass_after_com_failure=False,
    )


def verify_track_changes_enabled(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    doc_in_path,
    read_document_text,
    read_revision_info,
) -> VerificationResult:
    """Verify the exact tracked body-end insertion and preservation."""
    return policies_content_operations.com_primary_with_saved_guardrail(
        lambda: verify_track_changes_com(
            pred_path,
            gold_path,
            spec,
            read_document_text=read_document_text,
            read_revision_info=read_revision_info,
        ),
        lambda: verify_track_changes_saved(pred_path, gold_path, spec, doc_in_path=doc_in_path),
        saved_pass_after_com_failure=False,
    )
