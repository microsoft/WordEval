"""families / notes / compare with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.composition import composition_command_ids
from evaluation.methods.word_object.families.notes import (
    extract_saved as families_notes_extract_saved,
)
from evaluation.methods.word_object.policies import (
    content_operations as policies_content_operations,
)
from pathlib import Path
from typing import Any, Dict, List, Optional
import re


def should_verify_note_body_text(spec: Optional[Dict]) -> bool:
    """Only require exact body text when notes are the whole body mutation."""
    commands = composition_command_ids(spec)
    if not commands:
        return True

    note_commands = {"FootnoteInsert", "EndnoteInsertWord"}
    return set(commands).issubset(note_commands)


def normalize_note_anchor_text(text: Optional[str]) -> str:
    """Normalize body-text snippets used for semantic note anchor checks."""
    without_note_refs = (text or "").replace("\x02", "")
    return re.sub(r"\s+", " ", without_note_refs).strip().lower()


def extract_note_anchor_signature(
    body_text: str, position: int, anchor_hint: Optional[str] = None
) -> Dict[str, Optional[str]]:
    """Extract lightweight semantic context around a note reference."""
    pos = max(0, min(position, len(body_text)))
    left_text = body_text[:pos]
    right_text = body_text[pos:]

    prev_match = re.search(r"([\w]+)\W*$", left_text, flags=re.UNICODE)
    next_match = re.search(r"^\W*([\w]+)", right_text, flags=re.UNICODE)

    normalized_left = normalize_note_anchor_text(left_text)
    normalized_right = normalize_note_anchor_text(right_text)
    normalized_hint = normalize_note_anchor_text(anchor_hint)

    return {
        "prev_word": prev_match.group(1).lower() if prev_match else None,
        "next_word": next_match.group(1).lower() if next_match else None,
        "left_suffix": normalized_left[-80:],
        "right_prefix": normalized_right[:80],
        "matches_hint": bool(normalized_hint and normalized_left.endswith(normalized_hint)),
    }


def note_reference_positions_match(
    pred_path: Path,
    gold_path: Path,
    pred_pos: Optional[int],
    gold_pos: Optional[int],
    spec: Optional[Dict] = None,
    *,
    read_document_text,
) -> Dict[str, Optional[object]]:
    """
    Compare note reference positions with a semantic anchor fallback.

    Exact COM positions are still preferred, but Word can drift by one
    character at the same logical insertion point after note insertion.
    """
    if pred_pos is None or gold_pos is None:
        return {"matched": True, "mode": "position_unavailable"}
    if pred_pos == gold_pos:
        return {"matched": True, "mode": "exact_position"}

    precondition = spec.get("precondition", {}) if spec else {}
    anchor_hint = precondition.get("cursor_after")

    pred_text = read_document_text(str(pred_path))
    gold_text = read_document_text(str(gold_path))
    pred_anchor = extract_note_anchor_signature(pred_text, pred_pos, anchor_hint)
    gold_anchor = extract_note_anchor_signature(gold_text, gold_pos, anchor_hint)

    same_neighbor_words = (
        pred_anchor.get("prev_word")
        and pred_anchor.get("prev_word") == gold_anchor.get("prev_word")
        and pred_anchor.get("next_word") == gold_anchor.get("next_word")
    )
    if pred_anchor.get("matches_hint") and gold_anchor.get("matches_hint") and same_neighbor_words:
        return {
            "matched": True,
            "mode": "semantic_anchor",
            "pred_anchor": pred_anchor,
            "gold_anchor": gold_anchor,
        }

    same_anchor_hint = (
        anchor_hint
        and pred_anchor.get("matches_hint")
        and gold_anchor.get("matches_hint")
        and pred_anchor.get("prev_word")
        and pred_anchor.get("prev_word") == gold_anchor.get("prev_word")
    )
    if same_anchor_hint:
        return {
            "matched": True,
            "mode": "semantic_anchor_prev_word",
            "pred_anchor": pred_anchor,
            "gold_anchor": gold_anchor,
        }

    if abs(pred_pos - gold_pos) <= 1 and same_neighbor_words:
        return {
            "matched": True,
            "mode": "adjacent_boundary_tolerance",
            "pred_anchor": pred_anchor,
            "gold_anchor": gold_anchor,
        }

    return {
        "matched": False,
        "mode": "position_mismatch",
        "pred_anchor": pred_anchor,
        "gold_anchor": gold_anchor,
    }


def saved_note_anchor_projections_match(
    pred_path: Path, gold_path: Path, kind: str
) -> Dict[str, Any]:
    """Compare the complete saved anchor projection for one note kind."""
    try:
        gold = families_notes_extract_saved.extract_saved_note_signature(gold_path, kind)
    except Exception as exc:
        return {
            "matched": False,
            "mode": "saved_anchor_unavailable",
            "gold_error": str(exc),
        }
    try:
        pred = families_notes_extract_saved.extract_saved_note_signature(pred_path, kind)
    except Exception as exc:
        return {
            "matched": False,
            "mode": "saved_anchor_unavailable",
            "pred_error": str(exc),
        }

    gold_errors = saved_note_structure_errors(gold, kind)
    pred_errors = saved_note_structure_errors(pred, kind)

    def projection(signature: Dict[str, Any]) -> List[Dict[str, Any]]:
        return [
            {
                "paragraph_index": row.get("paragraph_index"),
                "character_offset": row.get("character_offset"),
                "paragraph_text": row.get("paragraph_text") or "",
                "note_text": row.get("note_text"),
                "reference_kind": ("custom" if row.get("custom_mark_follows") else "auto"),
                "custom_mark": row.get("custom_mark") or "",
                "section_index": row.get("section_index"),
                "valid_id": bool(row.get("valid_id")),
            }
            for row in signature.get("references", [])
        ]

    gold_projection = projection(gold)
    pred_projection = projection(pred)
    return {
        "matched": (not gold_errors and not pred_errors and pred_projection == gold_projection),
        "mode": "saved_ooxml_note_anchor_projection",
        "gold": gold_projection,
        "pred": pred_projection,
        "gold_errors": gold_errors,
        "pred_errors": pred_errors,
    }


def saved_note_entries(signature: Dict[str, Any]) -> List[Dict[str, Any]]:
    return [
        {
            "text": reference.get("note_text"),
            "reference_kind": ("custom" if reference.get("custom_mark_follows") else "auto"),
            "custom_mark": reference.get("custom_mark") or "",
            "paragraph_index": reference.get("paragraph_index"),
            "character_offset": reference.get("character_offset"),
            "paragraph_text": reference.get("paragraph_text") or "",
            "section_index": reference.get("section_index"),
        }
        for reference in signature.get("references", [])
    ]


def saved_note_property_state(
    signature: Dict[str, Any], *, ignore: Optional[str] = None, kind: Optional[str] = None
) -> Dict[str, Any]:
    def filtered(properties: Dict[str, Any]) -> Dict[str, Any]:
        result = {key: value for key, value in (properties or {}).items() if key != ignore}
        if kind == "endnote" and result.get("pos") == "docEnd":
            # ``docEnd`` is Word's effective endnote-position default.
            # Explicit and omitted serializations therefore describe the
            # same setting.
            result.pop("pos")
        return result

    return {
        "settings": filtered(signature.get("settings_properties", {})),
        "sections": [
            {
                "section_index": row.get("section_index"),
                "properties": filtered(row.get("properties", {})),
            }
            for row in signature.get("section_properties", [])
        ],
    }


def effective_saved_note_property(
    signature: Dict[str, Any], property_name: str
) -> List[Optional[str]]:
    inherited = signature.get("settings_properties", {}).get(property_name)
    sections = signature.get("section_properties", [])
    if not sections:
        return [inherited]
    return [row.get("properties", {}).get(property_name, inherited) for row in sections]


def saved_note_structure_errors(signature: Dict[str, Any], kind: str) -> List[str]:
    references = signature.get("references", [])
    errors = []
    if references and not signature.get(f"{kind}s_part_present"):
        errors.append(f"missing {kind}s part")
    if references and signature.get(f"{kind}s_relationship_count") != 1:
        errors.append(f"{kind}s relationship count is not one")
    if any(not row.get("valid_id") for row in references):
        errors.append(f"invalid {kind} reference ID")
    if signature.get("orphan_notes"):
        errors.append(f"orphan {kind} definition")
    if signature.get("duplicate_note_ids"):
        errors.append(f"duplicate {kind} definition ID")
    if signature.get("duplicate_reference_ids"):
        errors.append(f"duplicate {kind} reference ID")
    return errors


def note_expected_values(
    kind: str, operation: str, spec: Optional[Dict]
) -> tuple[Optional[str], Optional[str]]:
    execution = (spec or {}).get("execution", {})
    if operation == "custom_symbol":
        return (
            execution.get("custom_text") or f"Custom symbol {kind}.",
            execution.get("custom_mark") or "♣",
        )
    if operation == "insert" and composition_command_ids(spec):
        # Composition instructions can request note text that differs from
        # the atomic command's default without copying that prose into the
        # generated spec.  The unique input->gold note delta below remains
        # the source of truth, and pred must still match gold exactly.  This
        # branch precedes the execution.insert_text branch so that carrying
        # an atomic command's execution into a composition spec never pins
        # the expected note text to the atomic default.
        return None, None
    if execution.get("insert_text"):
        return execution["insert_text"], None
    return "See reference.", None


def compare_com_note_entries(
    pred_path: Path,
    gold_path: Path,
    pred_entries: List[Dict[str, Any]],
    gold_entries: List[Dict[str, Any]],
    kind: str,
    spec: Optional[Dict],
    *,
    read_document_text,
) -> Optional[VerificationResult]:
    if len(pred_entries) != len(gold_entries):
        return VerificationResult(
            passed=False,
            message=f"FAIL: {kind.title()} count mismatch",
            details={
                "pred_count": len(pred_entries),
                "gold_count": len(gold_entries),
            },
        )
    for index, (pred_entry, gold_entry) in enumerate(zip(pred_entries, gold_entries), start=1):
        if pred_entry.get("text") != gold_entry.get("text"):
            return VerificationResult(
                passed=False,
                message=f"FAIL: {kind.title()} {index} text differs",
                details={
                    "pred_text": pred_entry.get("text"),
                    "gold_text": gold_entry.get("text"),
                },
            )
        pred_reference = pred_entry.get("reference") or ""
        gold_reference = gold_entry.get("reference") or ""
        pred_mark = "auto" if pred_reference == "\x02" else pred_reference
        gold_mark = "auto" if gold_reference == "\x02" else gold_reference
        if pred_mark != gold_mark:
            return VerificationResult(
                passed=False,
                message=f"FAIL: {kind.title()} {index} reference mark differs",
                details={
                    "pred_reference": pred_mark,
                    "gold_reference": gold_mark,
                },
            )
        if pred_entry.get("section_number") != gold_entry.get("section_number"):
            return VerificationResult(
                passed=False,
                message=f"FAIL: {kind.title()} {index} section differs",
                details={
                    "pred_section": pred_entry.get("section_number"),
                    "gold_section": gold_entry.get("section_number"),
                },
            )
        position_match = note_reference_positions_match(
            pred_path,
            gold_path,
            pred_entry.get("reference_start"),
            gold_entry.get("reference_start"),
            spec,
            read_document_text=read_document_text,
        )
        absolute_offsets_differ = (
            pred_entry.get("reference_start") is not None
            and gold_entry.get("reference_start") is not None
            and pred_entry.get("reference_start") != gold_entry.get("reference_start")
        )
        saved_anchor_match = None
        if absolute_offsets_differ:
            saved_anchor_match = saved_note_anchor_projections_match(pred_path, gold_path, kind)
            if saved_anchor_match.get("gold_error") or saved_anchor_match.get("gold_errors"):
                return policies_content_operations.error_with_details(
                    f"Gold saved {kind} anchor evidence is invalid",
                    {"saved_anchor_projection": saved_anchor_match},
                )
        if (not position_match["matched"] and not absolute_offsets_differ) or (
            absolute_offsets_differ and not saved_anchor_match.get("matched")
        ):
            return VerificationResult(
                passed=False,
                message=f"FAIL: {kind.title()} {index} anchor differs",
                details={
                    "match_mode": position_match["mode"],
                    "pred_anchor": position_match.get("pred_anchor"),
                    "gold_anchor": position_match.get("gold_anchor"),
                    "saved_anchor_projection": saved_anchor_match,
                },
            )
    return None
