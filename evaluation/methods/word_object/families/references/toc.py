"""families / references / toc with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.families.references import (
    extract_saved as families_references_extract_saved,
)
from evaluation.methods.word_object.policies import (
    content_operations as policies_content_operations,
)
from pathlib import Path
from typing import Any, Dict, List, Optional
import re
from .fields import require_reference_sections


def verify_saved_toc_text_level(
    pred_path: Path, gold_path: Path, target_text: str, spec: Optional[Dict]
) -> Optional[VerificationResult]:
    gold = families_references_extract_saved.extract_saved_toc_add_text_state(
        gold_path, target_text
    )
    pred = families_references_extract_saved.extract_saved_toc_add_text_state(
        pred_path, target_text
    )
    if gold is None or pred is None:
        return None

    if not gold["source_paragraphs"]:
        return policies_content_operations.error_with_details(
            f'Target paragraph "{target_text}" not found in gold',
            {"target_text": target_text, "gold": gold},
        )
    if not pred["source_paragraphs"]:
        return VerificationResult(
            passed=False,
            message=f'FAIL: Target paragraph "{target_text}" not found in pred',
            details={"target_text": target_text, "gold": gold, "pred": pred},
        )

    gold_source = gold["source_paragraphs"][0]
    pred_source = pred["source_paragraphs"][0]
    differences = []
    for key in ("paragraph_style", "outline_level", "font_size"):
        if gold_source.get(key) != pred_source.get(key):
            differences.append(
                {
                    "property": key,
                    "gold": gold_source.get(key),
                    "pred": pred_source.get(key),
                }
            )
    if differences:
        return VerificationResult(
            passed=False,
            message="FAIL: TOC text level assignment differs on target paragraph",
            details={
                "target_text": target_text,
                "differences": differences,
                "gold": gold,
                "pred": pred,
                "match_mode": "saved_ooxml_toc_cache",
            },
        )

    if pred.get("source_context") != gold.get("source_context"):
        return VerificationResult(
            passed=False,
            message="FAIL: TOC Add Text changed surrounding source content or heading semantics",
            details={
                "target_text": target_text,
                "gold_source_context": gold.get("source_context"),
                "pred_source_context": pred.get("source_context"),
                "match_mode": "saved_ooxml_toc_source_preservation",
            },
        )

    if gold["toc_count"] == 0 or not gold["target_in_toc"]:
        return policies_content_operations.error_with_details(
            "Gold saved TOC does not include target text",
            {
                "target_text": target_text,
                "gold": gold,
                "match_mode": "saved_ooxml_toc_cache",
            },
        )
    if pred["toc_count"] != gold["toc_count"]:
        return VerificationResult(
            passed=False,
            message=(
                "FAIL: TOC count differs after Add Text "
                f"(pred={pred['toc_count']}, gold={gold['toc_count']})"
            ),
            details={
                "target_text": target_text,
                "gold": gold,
                "pred": pred,
                "match_mode": "saved_ooxml_toc_cache",
            },
        )
    if not pred["target_in_toc"]:
        return VerificationResult(
            passed=False,
            message="FAIL: Saved TOC does not include target text",
            details={
                "target_text": target_text,
                "gold": gold,
                "pred": pred,
                "match_mode": "saved_ooxml_toc_cache",
            },
        )

    return VerificationResult(
        passed=True,
        message="PASS: TOC text level matches and saved TOC includes target text",
        details={
            "target_text": target_text,
            "expected_level": (spec or {}).get("parameter", {}).get("level"),
            "values": pred_source,
            "toc_count": pred["toc_count"],
            "match_mode": "saved_ooxml_toc_cache",
        },
    )


def verify_toc_text_level(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    read_paragraph_outline_info_by_text,
    read_toc_content,
) -> VerificationResult:
    """Use WOM as primary and require the persisted TOC cache as guardrail."""

    def saved_guardrail():
        result = verify_saved_toc_text_level(
            pred_path,
            gold_path,
            (
                (spec or {}).get("precondition", {}).get("select_text")
                or (spec or {}).get("precondition", {}).get("full_text")
                or "Alpha TARGET Omega."
            ),
            spec,
        )
        return result or VerificationResult.error_result("Saved TOC state could not be extracted")

    return policies_content_operations.com_primary_with_saved_guardrail(
        lambda: verify_toc_text_level_com(
            pred_path,
            gold_path,
            spec,
            read_paragraph_outline_info_by_text=read_paragraph_outline_info_by_text,
            read_toc_content=read_toc_content,
        ),
        saved_guardrail,
        saved_pass_after_com_failure=False,
    )


def verify_toc_text_level_com(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    read_paragraph_outline_info_by_text,
    read_toc_content,
) -> VerificationResult:
    """
    Verify TOC Add Text gallery by comparing target paragraph TOC level metadata.

    This command often changes the paragraph's heading/outline level rather than
    the TOC object's text/count itself, so `_verify_toc()` is too weak.
    """
    precondition = spec.get("precondition", {}) if spec else {}
    target_text = (
        precondition.get("select_text") or precondition.get("full_text") or "Alpha TARGET Omega."
    )

    gold_info = read_paragraph_outline_info_by_text(str(gold_path), target_text)
    pred_info = read_paragraph_outline_info_by_text(str(pred_path), target_text)

    if gold_info is None:
        return policies_content_operations.error_with_details(
            f'Target paragraph "{target_text}" not found in gold',
            {"target_text": target_text},
        )

    if pred_info is None:
        return VerificationResult(
            passed=False,
            message=f'FAIL: Target paragraph "{target_text}" not found in pred',
            details={"target_text": target_text},
        )

    error = require_reference_sections({"gold": gold_info, "pred": pred_info},
                                       {"paragraph_style", "outline_level", "font_size"})
    if error is not None:
        return error
    differences = []
    for key in ("paragraph_style", "outline_level", "font_size"):
        if gold_info.get(key) != pred_info.get(key):
            differences.append(
                {
                    "property": key,
                    "gold": gold_info.get(key),
                    "pred": pred_info.get(key),
                }
            )

    if differences:
        return VerificationResult(
            passed=False,
            message="FAIL: TOC text level assignment differs on target paragraph",
            details={
                "target_text": target_text,
                "expected_level": (spec or {}).get("parameter", {}).get("level") if spec else None,
                "differences": differences,
                "gold": gold_info,
                "pred": pred_info,
            },
        )

    gold_toc = read_toc_content(str(gold_path), update_fields=True)
    pred_toc = read_toc_content(str(pred_path), update_fields=True)
    gold_toc_text = "\n".join(t.get("text") or "" for t in gold_toc.get("tocs", []))
    pred_toc_text = "\n".join(t.get("text") or "" for t in pred_toc.get("tocs", []))

    if gold_toc.get("toc_count", 0) == 0:
        return policies_content_operations.error_with_details(
            "Gold document has no TableOfContents object for TOC Add Text integration check",
            {
                "target_text": target_text,
                "paragraph_values": gold_info,
                "gold_toc": gold_toc,
            },
        )

    if target_text not in gold_toc_text:
        return policies_content_operations.error_with_details(
            "Gold TOC does not include target text after field update",
            {
                "target_text": target_text,
                "paragraph_values": gold_info,
                "gold_toc": gold_toc,
            },
        )

    if pred_toc.get("toc_count", 0) != gold_toc.get("toc_count", 0):
        return VerificationResult(
            passed=False,
            message=(
                "FAIL: TOC count differs after Add Text "
                f"(pred={pred_toc.get('toc_count', 0)}, gold={gold_toc.get('toc_count', 0)})"
            ),
            details={
                "target_text": target_text,
                "gold_toc": gold_toc,
                "pred_toc": pred_toc,
            },
        )

    if target_text not in pred_toc_text:
        return VerificationResult(
            passed=False,
            message="FAIL: Updated TOC does not include target text",
            details={
                "target_text": target_text,
                "expected_level": (spec or {}).get("parameter", {}).get("level") if spec else None,
                "paragraph_values": pred_info,
                "gold_toc": gold_toc,
                "pred_toc": pred_toc,
            },
        )

    return VerificationResult(
        passed=True,
        message="PASS: TOC text level assignment matches and updated TOC includes target text",
        details={
            "target_text": target_text,
            "expected_level": (spec or {}).get("parameter", {}).get("level") if spec else None,
            "values": gold_info,
            "toc_count": pred_toc.get("toc_count", 0),
        },
    )


def verify_toc(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, read_dynamic_reference_state
) -> VerificationResult:
    """Verify automatic insertion or removal using submitted and refreshed state."""
    gold = read_dynamic_reference_state(str(gold_path))
    pred = read_dynamic_reference_state(str(pred_path))
    error = require_reference_sections({"gold": gold, "pred": pred}, ('tocs', 'fields', 'source_headings', 'source_paragraphs', 'section_count', 'headers', 'footers', 'tables'))
    if error is not None:
        return error
    id_mso = (spec or {}).get("idMso", "")

    def normalized_source(state: Dict[str, Any]) -> Dict[str, Any]:
        return {
            "source_headings": state.get("source_headings", []),
            "source_paragraphs": state.get("source_paragraphs", []),
            "section_count": state.get("section_count"),
            "headers": state.get("headers", []),
            "footers": state.get("footers", []),
            "tables": [
                {
                    "row_count": table.get("row_count"),
                    "column_count": table.get("column_count"),
                    "cell_contents": table.get("cell_contents"),
                }
                for table in state.get("tables", [])
            ],
        }

    if id_mso == "TableOfContentsRemove":
        gold_tocs = gold.get("tocs", [])
        if gold_tocs:
            return policies_content_operations.error_with_details(
                "Gold still contains a TableOfContents object after remove",
                {"gold": gold},
            )
        pred_toc_fields = [
            field
            for field in pred.get("fields", [])
            if re.search(r"^\s*TOC\b", field.get("code") or "", re.I)
        ]
        if pred.get("tocs") or pred_toc_fields:
            return VerificationResult(
                passed=False,
                message="FAIL: TOC remove left a TOC object or field",
                details={"pred": pred, "toc_fields": pred_toc_fields},
            )
        if normalized_source(pred) != normalized_source(gold):
            return VerificationResult(
                passed=False,
                message="FAIL: TOC remove changed source content or layout",
                details={
                    "gold_source": normalized_source(gold),
                    "pred_source": normalized_source(pred),
                },
            )
        return VerificationResult(
            passed=True,
            message="PASS: TOC fully removed and source document preserved",
            details={
                "match_mode": "toc_remove_object_absence_and_preservation",
                "source_headings": pred.get("source_headings", []),
            },
        )

    gold_tocs = gold.get("tocs", [])
    pred_tocs = pred.get("tocs", [])
    if len(gold_tocs) != 1 or not gold_tocs[0].get("entries"):
        return policies_content_operations.error_with_details(
            "Gold lacks one non-empty automatic TOC object",
            {"gold": gold},
        )
    if len(pred_tocs) != 1:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Expected one TOC object, found {len(pred_tocs)}",
            details={"gold": gold, "pred": pred},
        )

    def entry_signature(toc: Dict[str, Any]) -> List[Dict[str, Any]]:
        return [
            {
                "text": " ".join((entry.get("text") or "").split()),
                "level": entry.get("level"),
            }
            for entry in toc.get("entries", [])
            if " ".join((entry.get("text") or "").split())
        ]

    def required_switches(code: str) -> Dict[str, bool]:
        normalized = " ".join((code or "").upper().split())
        return {
            "heading_1_3": bool(re.search(r'\\O\s+"?1-3"?', normalized)),
            "hyperlinks": bool(re.search(r"\\H(?:\s|$)", normalized)),
            "hide_web_page_numbers": bool(re.search(r"\\Z(?:\s|$)", normalized)),
            "outline_levels": bool(re.search(r"\\U(?:\s|$)", normalized)),
        }

    gold_toc = gold_tocs[0]
    pred_toc = pred_tocs[0]
    gold_entries = entry_signature(gold_toc)
    pred_entries = entry_signature(pred_toc)
    if pred_entries != gold_entries:
        return VerificationResult(
            passed=False,
            message="FAIL: Submitted TOC cache is empty, stale, or has wrong levels",
            details={
                "gold_entries": gold_entries,
                "pred_entries": pred_entries,
                "submitted_pred": pred_toc,
            },
        )
    if (
        " ".join((pred_toc.get("heading_text") or "").split())
        != " ".join((gold_toc.get("heading_text") or "").split())
        or (pred_toc.get("heading_style") or "").casefold()
        != (gold_toc.get("heading_style") or "").casefold()
    ):
        return VerificationResult(
            passed=False,
            message="FAIL: Automatic TOC heading text or style differs",
            details={"gold_toc": gold_toc, "pred_toc": pred_toc},
        )
    gold_switches = required_switches(gold_toc.get("code") or "")
    pred_switches = required_switches(pred_toc.get("code") or "")
    required = [key for key, present in gold_switches.items() if present]
    missing = [key for key in required if not pred_switches.get(key)]
    if missing:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Automatic TOC is missing required switch semantics: {', '.join(missing)}",
            details={
                "gold_code": gold_toc.get("code"),
                "pred_code": pred_toc.get("code"),
                "missing_switches": missing,
            },
        )
    if normalized_source(pred) != normalized_source(gold):
        return VerificationResult(
            passed=False,
            message="FAIL: Automatic TOC insertion changed source content or layout",
            details={
                "gold_source": normalized_source(gold),
                "pred_source": normalized_source(pred),
            },
        )

    refreshed = read_dynamic_reference_state(str(pred_path), update_fields=True)
    error = require_reference_sections({"refreshed": refreshed}, ('tocs', 'fields', 'source_headings', 'source_paragraphs', 'section_count', 'headers', 'footers', 'tables'))
    if error is not None:
        return error
    refreshed_tocs = refreshed.get("tocs", [])
    refreshed_entries = entry_signature(refreshed_tocs[0]) if len(refreshed_tocs) == 1 else []
    if len(refreshed_tocs) != 1 or refreshed_entries != gold_entries:
        return VerificationResult(
            passed=False,
            message="FAIL: Refreshed TOC does not contain all Heading 1-3 entries at the correct levels",
            details={
                "gold_entries": gold_entries,
                "submitted_pred_entries": pred_entries,
                "refreshed_pred_entries": refreshed_entries,
                "refreshed_pred": refreshed,
            },
        )
    return VerificationResult(
        passed=True,
        message="PASS: Automatic TOC submitted cache and refreshed semantics match",
        details={
            "match_mode": "toc_object_submitted_cache_refreshed_entries",
            "heading": pred_toc.get("heading_text"),
            "entries": pred_entries,
            "submitted_page_references": [
                entry.get("page_reference") for entry in pred_toc.get("entries", [])
            ],
            "refreshed_page_references": [
                entry.get("page_reference") for entry in refreshed_tocs[0].get("entries", [])
            ],
        },
    )
