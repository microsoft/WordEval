"""families / notes / verify_saved with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.composition import composition_command_ids
from evaluation.methods.word_object.families.notes import compare as families_notes_compare
from evaluation.methods.word_object.families.notes import (
    extract_saved as families_notes_extract_saved,
)
from evaluation.methods.word_object.policies import (
    content_operations as policies_content_operations,
)
from evaluation.methods.word_object.preservation import context as preservation_context
from pathlib import Path
from typing import Any, Dict, Optional


def verify_saved_note_context(
    pred_path: Path, gold_path: Path, *, strip_note_properties: Optional[str] = None, doc_in_path
) -> Optional[VerificationResult]:
    if doc_in_path is None:
        return policies_content_operations.error_with_details(
            "Input document is required for note preservation verification"
        )
    kwargs = {
        "strip_note_references": True,
        "strip_note_properties": strip_note_properties,
    }
    input_context = preservation_context.extract_saved_semantic_context(doc_in_path, **kwargs)
    gold_context = preservation_context.extract_saved_semantic_context(gold_path, **kwargs)
    pred_context = preservation_context.extract_saved_semantic_context(pred_path, **kwargs)
    if input_context != gold_context:
        return policies_content_operations.error_with_details(
            "Gold changed state outside the requested note operation",
            {
                "match_mode": "saved_ooxml_note_preservation",
                "input_context": input_context,
                "gold_context": gold_context,
            },
        )
    if pred_context != gold_context:
        return VerificationResult(
            passed=False,
            message="FAIL: Note operation changed unrelated content/layout",
            details={
                "match_mode": "saved_ooxml_note_preservation",
                "gold_context": gold_context,
                "pred_context": pred_context,
            },
        )
    return None


def verify_note_saved_operation(
    pred_path: Path,
    gold_path: Path,
    kind: str,
    operation: str,
    spec: Optional[Dict] = None,
    *,
    doc_in_path,
) -> VerificationResult:
    if doc_in_path is None:
        return policies_content_operations.error_with_details(
            "Input document is required for note semantic verification"
        )
    opposite = "endnote" if kind == "footnote" else "footnote"
    paths = {
        "input": doc_in_path,
        "gold": gold_path,
        "pred": pred_path,
    }
    signatures = {
        owner: {
            kind: families_notes_extract_saved.extract_saved_note_signature(path, kind),
            opposite: families_notes_extract_saved.extract_saved_note_signature(path, opposite),
        }
        for owner, path in paths.items()
    }
    for owner in ("input", "gold", "pred"):
        for note_kind in (kind, opposite):
            errors = families_notes_compare.saved_note_structure_errors(
                signatures[owner][note_kind], note_kind
            )
            if not errors:
                continue
            details = {
                "match_mode": "saved_ooxml_note_objects",
                "owner": owner,
                "kind": note_kind,
                "errors": errors,
                "signature": signatures[owner][note_kind],
            }
            if owner in {"input", "gold"}:
                return policies_content_operations.error_with_details(
                    f"{owner.title()} saved {note_kind} evidence is invalid",
                    details,
                )
            return VerificationResult(
                passed=False,
                message=f"FAIL: Predicted saved {note_kind} evidence is invalid",
                details=details,
            )

    input_entries = families_notes_compare.saved_note_entries(signatures["input"][kind])
    gold_entries = families_notes_compare.saved_note_entries(signatures["gold"][kind])
    pred_entries = families_notes_compare.saved_note_entries(signatures["pred"][kind])
    expected_text, expected_mark = families_notes_compare.note_expected_values(
        kind, operation, spec
    )
    target_property = None
    commands = set(composition_command_ids(spec))
    opposite_requested = (
        opposite == "footnote"
        and "FootnoteInsert" in commands
        or opposite == "endnote"
        and "EndnoteInsertWord" in commands
    )

    if operation in {"insert", "custom_symbol"}:
        candidates = [
            index
            for index, entry in enumerate(gold_entries)
            if (expected_text is None or entry["text"] == expected_text)
            and (
                operation != "custom_symbol"
                or entry["reference_kind"] == "custom"
                and entry["custom_mark"] == expected_mark
            )
            and gold_entries[:index] + gold_entries[index + 1 :] == input_entries
        ]
        if len(candidates) != 1:
            return policies_content_operations.error_with_details(
                "Gold does not add exactly one requested note object",
                {
                    "match_mode": "saved_ooxml_note_delta",
                    "kind": kind,
                    "operation": operation,
                    "expected_text": expected_text,
                    "expected_mark": expected_mark,
                    "input": input_entries,
                    "gold": gold_entries,
                    "candidate_indexes": candidates,
                },
            )
        target_index = candidates[0]
        target = gold_entries[target_index]
        execution = (spec or {}).get("execution", {})
        precondition = (spec or {}).get("precondition", {})
        if operation == "custom_symbol":
            anchor_valid = (
                target_index == 1
                and len(gold_entries) >= 2
                and target["paragraph_index"] == gold_entries[0]["paragraph_index"]
                and target["character_offset"] == gold_entries[0]["character_offset"]
            )
        elif execution.get("insert_para_index"):
            anchor_valid = target["paragraph_index"] == execution["insert_para_index"] and target[
                "character_offset"
            ] == len(target["paragraph_text"])
        else:
            anchor_hint = precondition.get("cursor_after", "TARGET")
            anchor_valid = (
                target["paragraph_text"][: target["character_offset"]]
                .rstrip()
                .endswith(anchor_hint)
            )
        if not anchor_valid:
            return policies_content_operations.error_with_details(
                "Gold note is not at the instruction-owned anchor/order",
                {
                    "match_mode": "saved_ooxml_note_delta",
                    "kind": kind,
                    "operation": operation,
                    "target_index": target_index,
                    "target": target,
                },
            )
        if pred_entries != gold_entries:
            return VerificationResult(
                passed=False,
                message=f"FAIL: Predicted {kind} objects/anchors differ",
                details={
                    "match_mode": "saved_ooxml_note_delta",
                    "kind": kind,
                    "operation": operation,
                    "gold": gold_entries,
                    "pred": pred_entries,
                },
            )
    elif operation == "number_format":
        target_property = "numFmt"
        expected = "lowerLetter"
        gold_values = families_notes_compare.effective_saved_note_property(
            signatures["gold"][kind], target_property
        )
        pred_values = families_notes_compare.effective_saved_note_property(
            signatures["pred"][kind], target_property
        )
        if not gold_values or any(value != expected for value in gold_values):
            return policies_content_operations.error_with_details(
                "Gold saved note number format is missing or incorrect",
                {"gold_values": gold_values, "expected": expected},
            )
        if not pred_values or any(value != expected for value in pred_values):
            return VerificationResult(
                passed=False,
                message=f"FAIL: Predicted {kind} number format is not lowercase letters",
                details={
                    "match_mode": "saved_ooxml_note_number_format",
                    "expected": expected,
                    "gold_values": gold_values,
                    "pred_values": pred_values,
                },
            )
    elif operation == "numbering_rule":
        target_property = "numRestart"
        expected = "eachSect"
        gold_values = families_notes_compare.effective_saved_note_property(
            signatures["gold"][kind], target_property
        )
        pred_values = families_notes_compare.effective_saved_note_property(
            signatures["pred"][kind], target_property
        )
        if not gold_values or any(value != expected for value in gold_values):
            return policies_content_operations.error_with_details(
                "Gold saved note restart rule is missing or incorrect",
                {"gold_values": gold_values, "expected": expected},
            )
        if not pred_values or any(value != expected for value in pred_values):
            return VerificationResult(
                passed=False,
                message=f"FAIL: Predicted {kind} numbering does not restart each section",
                details={
                    "match_mode": "saved_ooxml_note_numbering_rule",
                    "expected": expected,
                    "gold_values": gold_values,
                    "pred_values": pred_values,
                },
            )
    else:
        raise ValueError(f"Unsupported note operation: {operation}")

    if operation in {"number_format", "numbering_rule"}:
        if input_entries != gold_entries:
            return policies_content_operations.error_with_details(
                "Gold changed note objects during a setting-only operation",
                {"input": input_entries, "gold": gold_entries},
            )
        if pred_entries != gold_entries:
            return VerificationResult(
                passed=False,
                message=f"FAIL: Predicted {kind} objects changed",
                details={"gold": gold_entries, "pred": pred_entries},
            )

    for note_kind in (kind, opposite):
        if note_kind == opposite and opposite_requested:
            # The opposite note operation is independently verified by the
            # composition's other sub-verifier.
            continue
        gold_note_entries = families_notes_compare.saved_note_entries(signatures["gold"][note_kind])
        input_note_entries = families_notes_compare.saved_note_entries(
            signatures["input"][note_kind]
        )
        pred_note_entries = families_notes_compare.saved_note_entries(signatures["pred"][note_kind])
        if note_kind == opposite and input_note_entries != gold_note_entries:
            return policies_content_operations.error_with_details(
                f"Gold changed the opposite {note_kind} collection",
                {"input": input_note_entries, "gold": gold_note_entries},
            )
        if note_kind == opposite and pred_note_entries != gold_note_entries:
            return VerificationResult(
                passed=False,
                message=f"FAIL: Opposite {note_kind} collection changed",
                details={
                    "gold": gold_note_entries,
                    "pred": pred_note_entries,
                },
            )
        ignored = target_property if note_kind == kind else None
        input_properties = families_notes_compare.saved_note_property_state(
            signatures["input"][note_kind],
            ignore=ignored,
            kind=note_kind,
        )
        gold_properties = families_notes_compare.saved_note_property_state(
            signatures["gold"][note_kind],
            ignore=ignored,
            kind=note_kind,
        )
        pred_properties = families_notes_compare.saved_note_property_state(
            signatures["pred"][note_kind],
            ignore=ignored,
            kind=note_kind,
        )
        if input_properties != gold_properties:
            return policies_content_operations.error_with_details(
                f"Gold changed unrelated {note_kind} settings",
                {"input": input_properties, "gold": gold_properties},
            )
        if pred_properties != gold_properties:
            return VerificationResult(
                passed=False,
                message=f"FAIL: Unrelated {note_kind} settings changed",
                details={"gold": gold_properties, "pred": pred_properties},
            )

    if families_notes_compare.should_verify_note_body_text(spec):
        context_result = verify_saved_note_context(
            pred_path,
            gold_path,
            strip_note_properties=(kind if target_property else None),
            doc_in_path=doc_in_path,
        )
        if context_result is not None:
            return context_result
    return VerificationResult(
        passed=True,
        message=f"PASS: Saved {kind} semantics and preservation match",
        details={
            "match_mode": "saved_ooxml_note_semantics",
            "kind": kind,
            "operation": operation,
            "entries": gold_entries,
            **(
                {
                    "fallback_contract": "complete",
                    "requires_render_control": True,
                }
                if operation == "custom_symbol"
                else {}
            ),
        },
    )


def verify_footnotes_saved(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, doc_in_path
) -> VerificationResult:
    if doc_in_path is None:
        # Direct route-unit calls predate input-aware delta verification.
        # Retain their narrow gold/pred saved guardrail while production
        # verification (which always supplies doc_in) uses the full delta.
        gold = families_notes_extract_saved.extract_saved_note_signature(gold_path, "footnote")
        pred = families_notes_extract_saved.extract_saved_note_signature(pred_path, "footnote")
        gold_errors = families_notes_compare.saved_note_structure_errors(gold, "footnote")
        pred_errors = families_notes_compare.saved_note_structure_errors(pred, "footnote")
        if gold_errors:
            return policies_content_operations.error_with_details(
                "Gold saved footnote evidence is missing or invalid",
                {"gold": gold, "errors": gold_errors},
            )
        if pred_errors:
            return VerificationResult(
                passed=False,
                message="FAIL: Predicted saved footnote evidence is invalid",
                details={
                    "gold": gold,
                    "pred": pred,
                    "errors": pred_errors,
                },
            )
        strict_body = families_notes_compare.should_verify_note_body_text(spec)

        def scoped(signature: Dict[str, Any]) -> Dict[str, Any]:
            result = {
                "entries": [
                    {
                        key: reference.get(key)
                        for key in (
                            "paragraph_index",
                            "character_offset",
                            "note_text",
                            "valid_id",
                            "custom_mark_follows",
                            "custom_mark",
                        )
                    }
                    for reference in signature.get("references", [])
                ],
                "orphan_notes": signature.get("orphan_notes", []),
                "relationship_count": signature.get("footnotes_relationship_count", 0),
            }
            if strict_body:
                result["body_paragraphs"] = signature.get("body_paragraphs", [])
            return result

        gold_scoped = scoped(gold)
        pred_scoped = scoped(pred)
        if gold_scoped != pred_scoped:
            return VerificationResult(
                passed=False,
                message="FAIL: Saved footnote semantic signature differs",
                details={
                    "scope": "full_document" if strict_body else "footnote_only",
                    "gold": gold_scoped,
                    "pred": pred_scoped,
                },
            )
        return VerificationResult(
            passed=True,
            message="PASS: Saved footnote semantic signature matches",
            details={
                "scope": "full_document" if strict_body else "footnote_only",
                "signature": gold_scoped,
            },
        )
    return verify_note_saved_operation(
        pred_path, gold_path, "footnote", "insert", spec, doc_in_path=doc_in_path
    )
