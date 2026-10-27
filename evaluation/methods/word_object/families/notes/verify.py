"""families / notes / verify with explicit evidence dependencies."""

from collections import Counter
from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.families.notes import compare as families_notes_compare
from evaluation.methods.word_object.families.notes import (
    verify_saved as families_notes_verify_saved,
)
from evaluation.methods.word_object.policies import (
    content_operations as policies_content_operations,
)
from pathlib import Path
from typing import Dict, Optional


def verify_note_com_operation(
    pred_path: Path,
    gold_path: Path,
    kind: str,
    operation: str,
    spec: Optional[Dict] = None,
    *,
    read_document_text,
    read_footnote_endnote_content,
) -> VerificationResult:
    gold = read_footnote_endnote_content(str(gold_path))
    pred = read_footnote_endnote_content(str(pred_path))
    key = f"{kind}s"
    count_key = f"{kind}_count"
    if len(gold[key]) != gold[count_key]:
        return policies_content_operations.error_with_details(
            f"Gold {kind} extraction is incomplete", {"gold": gold}
        )
    if len(pred[key]) != pred[count_key]:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Predicted {kind} extraction is incomplete",
            details={"pred": pred},
        )

    if operation == "number_format":
        style_key = f"{kind}_number_style"
        if gold.get(style_key) != 4:
            return policies_content_operations.error_with_details(
                f"Gold {kind} number style is not lowercase letters",
                {"gold_style": gold.get(style_key)},
            )
        if pred.get(style_key) != 4:
            return VerificationResult(
                passed=False,
                message=f"FAIL: Predicted {kind} number style is not lowercase letters",
                details={
                    "pred_style": pred.get(style_key),
                    "gold_style": gold.get(style_key),
                },
            )
    elif operation == "numbering_rule":
        rule_key = f"{kind}_numbering_rule"
        if gold.get(rule_key) != 1:
            return policies_content_operations.error_with_details(
                f"Gold {kind} numbering rule is not restart each section",
                {"gold_rule": gold.get(rule_key)},
            )
        if pred.get(rule_key) != 1:
            return VerificationResult(
                passed=False,
                message=f"FAIL: Predicted {kind} numbering does not restart each section",
                details={
                    "pred_rule": pred.get(rule_key),
                    "gold_rule": gold.get(rule_key),
                },
            )
        gold_sections = Counter(row.get("section_number") for row in gold[key])
        pred_sections = Counter(row.get("section_number") for row in pred[key])
        if pred_sections != gold_sections:
            return VerificationResult(
                passed=False,
                message=f"FAIL: Predicted {kind} section numbering groups differ",
                details={
                    "gold_sections": dict(gold_sections),
                    "pred_sections": dict(pred_sections),
                },
            )

    mismatch = families_notes_compare.compare_com_note_entries(
        pred_path,
        gold_path,
        pred[key],
        gold[key],
        kind,
        spec,
        read_document_text=read_document_text,
    )
    if mismatch is not None:
        return mismatch
    opposite = "endnote" if kind == "footnote" else "footnote"
    opposite_mismatch = families_notes_compare.compare_com_note_entries(
        pred_path,
        gold_path,
        pred[f"{opposite}s"],
        gold[f"{opposite}s"],
        opposite,
        spec,
        read_document_text=read_document_text,
    )
    if opposite_mismatch is not None:
        return opposite_mismatch
    return VerificationResult(
        passed=True,
        message=f"PASS: Effective {kind} semantics match ({len(pred[key])})",
        details={
            "match_mode": "word_com_note_semantics",
            "kind": kind,
            "operation": operation,
            "entries": pred[key],
        },
    )


def verify_footnotes(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    doc_in_path,
    read_document_text,
    read_footnote_endnote_content,
) -> VerificationResult:
    """Verify footnotes through WOM, with saved package semantics as guardrail."""
    return policies_content_operations.com_primary_with_saved_guardrail(
        lambda: verify_footnotes_com(
            pred_path,
            gold_path,
            spec,
            read_document_text=read_document_text,
            read_footnote_endnote_content=read_footnote_endnote_content,
        ),
        lambda: families_notes_verify_saved.verify_footnotes_saved(
            pred_path, gold_path, spec, doc_in_path=doc_in_path
        ),
        saved_pass_after_com_failure=False,
    )


def verify_footnotes_com(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    read_document_text,
    read_footnote_endnote_content,
) -> VerificationResult:
    return verify_note_com_operation(
        pred_path,
        gold_path,
        "footnote",
        "insert",
        spec,
        read_document_text=read_document_text,
        read_footnote_endnote_content=read_footnote_endnote_content,
    )


def verify_endnotes(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    doc_in_path,
    read_document_text,
    read_footnote_endnote_content,
) -> VerificationResult:
    """Verify effective and persisted endnote semantics."""
    return policies_content_operations.com_primary_with_saved_guardrail(
        lambda: verify_note_com_operation(
            pred_path,
            gold_path,
            "endnote",
            "insert",
            spec,
            read_document_text=read_document_text,
            read_footnote_endnote_content=read_footnote_endnote_content,
        ),
        lambda: families_notes_verify_saved.verify_note_saved_operation(
            pred_path, gold_path, "endnote", "insert", spec, doc_in_path=doc_in_path
        ),
        saved_pass_after_com_failure=False,
    )


def verify_footnote_number_format(
    pred_path: Path,
    gold_path: Path,
    *,
    doc_in_path,
    read_document_text,
    read_footnote_endnote_content,
) -> VerificationResult:
    """Verify lowercase-letter footnote numbering and preservation."""
    return policies_content_operations.com_primary_with_saved_guardrail(
        lambda: verify_note_com_operation(
            pred_path,
            gold_path,
            "footnote",
            "number_format",
            read_document_text=read_document_text,
            read_footnote_endnote_content=read_footnote_endnote_content,
        ),
        lambda: families_notes_verify_saved.verify_note_saved_operation(
            pred_path, gold_path, "footnote", "number_format", doc_in_path=doc_in_path
        ),
        saved_pass_after_com_failure=False,
    )


def verify_endnote_number_format(
    pred_path: Path,
    gold_path: Path,
    *,
    doc_in_path,
    read_document_text,
    read_footnote_endnote_content,
) -> VerificationResult:
    """Verify lowercase-letter endnote numbering and preservation."""
    return policies_content_operations.com_primary_with_saved_guardrail(
        lambda: verify_note_com_operation(
            pred_path,
            gold_path,
            "endnote",
            "number_format",
            read_document_text=read_document_text,
            read_footnote_endnote_content=read_footnote_endnote_content,
        ),
        lambda: families_notes_verify_saved.verify_note_saved_operation(
            pred_path, gold_path, "endnote", "number_format", doc_in_path=doc_in_path
        ),
        saved_pass_after_com_failure=False,
    )


def verify_footnote_numbering_rule(
    pred_path: Path,
    gold_path: Path,
    *,
    doc_in_path,
    read_document_text,
    read_footnote_endnote_content,
) -> VerificationResult:
    """Verify effective restart-each-section semantics and preservation."""
    return policies_content_operations.com_primary_with_saved_guardrail(
        lambda: verify_note_com_operation(
            pred_path,
            gold_path,
            "footnote",
            "numbering_rule",
            read_document_text=read_document_text,
            read_footnote_endnote_content=read_footnote_endnote_content,
        ),
        lambda: families_notes_verify_saved.verify_note_saved_operation(
            pred_path, gold_path, "footnote", "numbering_rule", doc_in_path=doc_in_path
        ),
        saved_pass_after_com_failure=False,
    )


def verify_footnote_custom_symbol(
    pred_path: Path,
    gold_path: Path,
    *,
    doc_in_path,
    read_document_text,
    read_footnote_endnote_content,
) -> VerificationResult:
    """Verify the real club-symbol footnote and its semantic anchor."""
    return policies_content_operations.com_primary_with_saved_guardrail(
        lambda: verify_note_com_operation(
            pred_path,
            gold_path,
            "footnote",
            "custom_symbol",
            read_document_text=read_document_text,
            read_footnote_endnote_content=read_footnote_endnote_content,
        ),
        lambda: families_notes_verify_saved.verify_note_saved_operation(
            pred_path, gold_path, "footnote", "custom_symbol", doc_in_path=doc_in_path
        ),
        saved_pass_after_com_failure=True,
        saved_pass_error_pattern=r"Object has been deleted",
    )


def verify_endnote_custom_symbol(
    pred_path: Path,
    gold_path: Path,
    *,
    doc_in_path,
    read_document_text,
    read_footnote_endnote_content,
) -> VerificationResult:
    """Verify the real club-symbol endnote and its semantic anchor."""
    return policies_content_operations.com_primary_with_saved_guardrail(
        lambda: verify_note_com_operation(
            pred_path,
            gold_path,
            "endnote",
            "custom_symbol",
            read_document_text=read_document_text,
            read_footnote_endnote_content=read_footnote_endnote_content,
        ),
        lambda: families_notes_verify_saved.verify_note_saved_operation(
            pred_path, gold_path, "endnote", "custom_symbol", doc_in_path=doc_in_path
        ),
        saved_pass_after_com_failure=False,
    )


def verify_footnote_endnote_composition(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    add_composition_atomic_details,
    doc_in_path,
    read_document_text,
    read_footnote_endnote_content,
    summarize_sub_result,
) -> VerificationResult:
    """Verify both footnotes AND endnotes match between gold and pred."""
    sub_results = {}
    failures = []
    errors = []

    fn_result = verify_footnotes(
        pred_path,
        gold_path,
        spec,
        doc_in_path=doc_in_path,
        read_document_text=read_document_text,
        read_footnote_endnote_content=read_footnote_endnote_content,
    )
    sub_results["footnotes"] = summarize_sub_result(fn_result)
    if fn_result.error:
        errors.append(fn_result.error)
    elif not fn_result.passed:
        failures.append(fn_result.message)

    en_result = verify_endnotes(
        pred_path,
        gold_path,
        spec,
        doc_in_path=doc_in_path,
        read_document_text=read_document_text,
        read_footnote_endnote_content=read_footnote_endnote_content,
    )
    sub_results["endnotes"] = summarize_sub_result(en_result)
    if en_result.error:
        errors.append(en_result.error)
    elif not en_result.passed:
        failures.append(en_result.message)

    details = {"sub_results": sub_results}
    add_composition_atomic_details(details, pred_path, gold_path, spec)

    if errors:
        error = "; ".join(errors)
        return VerificationResult(
            passed=False,
            message=(
                "ERROR: Footnote/endnote composition could not be fully "
                f"evaluated ({len(errors)} error(s))"
            ),
            details={**details, "errors": errors},
            error=error,
        )

    if failures:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Footnote/endnote composition mismatch ({len(failures)} check(s))",
            details=details,
        )

    return VerificationResult(
        passed=True,
        message="PASS: Both footnotes and endnotes match",
        details=details,
    )
