"""families / pages / verify with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.families.pages import (
    extract_saved as families_pages_extract_saved,
)
from evaluation.methods.word_object.policies import (
    content_operations as policies_content_operations,
)
from pathlib import Path
from typing import Any, Dict, List, Optional
import re


def verify_sections(
    pred_path: Path, gold_path: Path, *, read_section_properties
) -> VerificationResult:
    """Verify section properties match."""
    gold_sections = read_section_properties(str(gold_path))
    pred_sections = read_section_properties(str(pred_path))

    if not gold_sections or not pred_sections:
        return policies_content_operations.error_with_details(
            'Section observation is empty; every Word document must have a section',
            {'gold':gold_sections,'pred':pred_sections})

    if len(gold_sections) != len(pred_sections):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Section count mismatch (pred={len(pred_sections)}, gold={len(gold_sections)})",
            details={"pred_count": len(pred_sections), "gold_count": len(gold_sections)},
        )

    for i, (pred_s, gold_s) in enumerate(zip(pred_sections, gold_sections)):
        if pred_s != gold_s:
            return VerificationResult(
                passed=False,
                message=f"FAIL: Section {i+1} properties differ",
                details={"section_index": i + 1, "pred_section": pred_s, "gold_section": gold_s},
            )

    return VerificationResult(
        passed=True,
        message=f"PASS: Sections match ({len(pred_sections)} sections)",
        details={"count": len(pred_sections), "sections": pred_sections},
    )


def verify_pages(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    get_optional_reader,
    read_page_info,
) -> VerificationResult:
    """Verify native page/blank-page semantics at the requested anchor."""
    gold_info = read_page_info(str(gold_path))
    pred_info = read_page_info(str(pred_path))

    id_mso = (spec or {}).get("idMso", "")
    expected_breaks = {
        "PageBreakInsertWord": 1,
        "BlankPageInsert": 2,
    }.get(id_mso)
    gold_breaks = gold_info.get("manual_page_breaks")
    if gold_info.get("page_count") is None or gold_breaks is None:
        return policies_content_operations.error_with_details(
            "Gold page-layout evidence could not be extracted",
            {"gold": gold_info, "idMso": id_mso},
        )
    if expected_breaks is not None and gold_breaks != expected_breaks:
        return policies_content_operations.error_with_details(
            "Gold does not contain the native page-layout target required " f"by {id_mso}",
            {
                "idMso": id_mso,
                "expected_manual_page_breaks": expected_breaks,
                "gold": gold_info,
            },
        )
    if expected_breaks is None and gold_breaks < 1:
        return policies_content_operations.error_with_details(
            "Gold contains no manual page-break target",
            {"gold": gold_info, "idMso": id_mso},
        )

    differences = {}

    if gold_info["page_count"] != pred_info["page_count"]:
        differences["page_count"] = {
            "pred": pred_info["page_count"],
            "gold": gold_info["page_count"],
        }

    if gold_info.get("manual_page_breaks") != pred_info.get("manual_page_breaks"):
        differences["manual_page_breaks"] = {
            "pred": pred_info.get("manual_page_breaks"),
            "gold": gold_info.get("manual_page_breaks"),
        }

    # A manual break can be stored inside the existing paragraph or between
    # new paragraph marks.  Those encodings are Word-visible equivalents,
    # so compare visible content and semantic offsets rather than requiring
    # the gold paragraph skeleton byte-for-byte.
    def visible_text(value: Any) -> Optional[str]:
        if value is None:
            return None
        return "".join(character for character in value if character not in {"\r", "\x0c", "\x07"})

    gold_visible_text = visible_text(gold_info.get("content_text"))
    pred_visible_text = visible_text(pred_info.get("content_text"))
    precondition = (spec or {}).get("precondition") or {}
    expected_visible_text = precondition.get("full_text")
    if expected_visible_text is not None:
        if gold_visible_text != expected_visible_text:
            return policies_content_operations.error_with_details(
                "Gold visible content does not satisfy the page-layout contract",
                {
                    "idMso": id_mso,
                    "expected_visible_text": expected_visible_text,
                    "gold_visible_text": gold_visible_text,
                },
            )
    else:
        expected_visible_text = gold_visible_text
    if pred_visible_text != expected_visible_text:
        differences["visible_content"] = {
            "pred": pred_visible_text,
            "expected": expected_visible_text,
        }

    def semantic_break_positions(info: Dict[str, Any]) -> List[Dict[str, Any]]:
        positions = []
        for position in info.get("page_break_positions") or []:
            preceding = position.get("visible_preceding_text")
            if preceding is None:
                preceding = visible_text(position.get("preceding_text"))
            positions.append(
                {
                    "visible_offset": position.get("visible_offset"),
                    "visible_preceding_text": preceding,
                }
            )
        return positions

    gold_positions = semantic_break_positions(gold_info)
    pred_positions = semantic_break_positions(pred_info)
    cursor_after = precondition.get("cursor_after")
    expected_anchor_text = None
    if expected_visible_text and cursor_after:
        anchor_end = expected_visible_text.find(cursor_after)
        if anchor_end >= 0:
            anchor_end += len(cursor_after)
            expected_anchor_text = expected_visible_text[:anchor_end]

    def positions_match_contract(positions: List[Dict[str, Any]]) -> bool:
        if len(positions) != gold_breaks:
            return False
        if expected_anchor_text is None:
            return positions == gold_positions
        expected_offset = len(expected_anchor_text)
        for position in positions:
            preceding = position.get("visible_preceding_text")
            if position.get("visible_offset") != expected_offset:
                return False
            if preceding is not None and not preceding.endswith(expected_anchor_text):
                return False
        return True

    if not positions_match_contract(gold_positions):
        return policies_content_operations.error_with_details(
            "Gold manual breaks do not satisfy the requested anchor contract",
            {
                "idMso": id_mso,
                "expected_anchor_text": expected_anchor_text,
                "gold_page_break_positions": gold_positions,
            },
        )
    if not positions_match_contract(pred_positions):
        differences["page_break_positions"] = {
            "pred": pred_positions,
            "expected_anchor_text": expected_anchor_text,
            "gold": gold_positions,
        }

    preservation = {}
    for label, extractor_name in (
        ("paragraph_formatting", "extract_all_paragraphs_formatting"),
        ("character_formatting", "extract_all_paragraph_character_formatting"),
        ("section_properties", "extract_section_properties"),
    ):
        extractor = get_optional_reader(extractor_name)
        if extractor is None:
            continue
        gold_value = extractor(str(gold_path))
        pred_value = extractor(str(pred_path))
        if label == "character_formatting":
            from ..formatting.characters import validate_character_observations
            error = validate_character_observations({"gold": gold_value, "pred": pred_value})
            if error is not None:
                return error
        preservation[label] = {"gold": gold_value, "pred": pred_value}
        if gold_value is None:
            return policies_content_operations.error_with_details(
                f"Gold {label} evidence could not be extracted",
                {"idMso": id_mso, **preservation},
            )
        if label == "section_properties":
            gold_effective = gold_value
            pred_effective = pred_value
        elif label == "character_formatting":
            from ..formatting.characters import visible_character_stream
            gold_effective = visible_character_stream(gold_value)
            pred_effective = visible_character_stream(pred_value)
        else:

            def effective_visible_formatting(value: Any) -> Any:
                if not isinstance(value, list):
                    return value
                effective = []
                for item in value:
                    if not isinstance(item, dict):
                        signature = item
                    else:
                        if not visible_text(item.get("text")):
                            continue
                        signature = {
                            key: item_value
                            for key, item_value in item.items()
                            if key not in {"index", "text"}
                        }
                    # Splitting a paragraph around a break can duplicate one
                    # effective format signature. Collapse only adjacent
                    # duplicates while preserving meaningful order.
                    if not effective or effective[-1] != signature:
                        effective.append(signature)
                return effective

            gold_effective = effective_visible_formatting(gold_value)
            pred_effective = effective_visible_formatting(pred_value)
        if pred_effective != gold_effective:
            differences[label] = {
                "pred": pred_effective,
                "gold": gold_effective,
            }

    if differences:
        return VerificationResult(
            passed=False,
            message="FAIL: Page-layout command semantics differ",
            details={
                "pred_pages": pred_info["page_count"],
                "gold_pages": gold_info["page_count"],
                "pred_manual_page_breaks": pred_info.get("manual_page_breaks"),
                "gold_manual_page_breaks": gold_info.get("manual_page_breaks"),
                "differences": differences,
                "idMso": id_mso,
            },
        )

    return VerificationResult(
        passed=True,
        message=(
            f"PASS: Page layout matches "
            f"({pred_info['page_count']} pages, {pred_info.get('manual_page_breaks', 0)} manual page breaks)"
        ),
        details={
            **pred_info,
            "idMso": id_mso,
            "match_mode": "word_page_layout_semantic_anchor_and_preservation",
            "preservation_checks": sorted(preservation),
        },
    )


def verify_cover_page(
    pred_path: Path, gold_path: Path, *, read_content_controls, read_document_text, read_page_info
) -> VerificationResult:
    """
    Verify cover-page building block style by comparing stable Word-level signals.

    Intentionally avoids shape/table count heuristics; compares:
    - rendered page count reported by Word
    - content control signature (count + type/title/tag)
    - saved paragraph-role and drawing-layout signature
    - document text with whitespace normalized
    """
    gold_page_info = read_page_info(str(gold_path))
    pred_page_info = read_page_info(str(pred_path))
    gold_page_count = gold_page_info.get("page_count")
    pred_page_count = pred_page_info.get("page_count")
    if not isinstance(gold_page_count, int) or gold_page_count < 1:
        return VerificationResult.error_result("Gold cover-page page count is unavailable")
    if not isinstance(pred_page_count, int) or pred_page_count < 1:
        return VerificationResult.error_result("Prediction cover-page page count is unavailable")
    if gold_page_count != pred_page_count:
        return VerificationResult(
            passed=False,
            message="FAIL: Cover page rendered page count differs",
            details={
                "match_mode": "cover_page_signature",
                "gold_page_count": gold_page_count,
                "pred_page_count": pred_page_count,
            },
        )

    gold_cc = read_content_controls(str(gold_path))
    pred_cc = read_content_controls(str(pred_path))

    def _norm_ws(text: str) -> str:
        return re.sub(r"\s+", " ", text or "").strip()

    def _cc_signature(info: Dict) -> Dict:
        controls = []
        for c in info.get("controls") or []:
            controls.append(
                {
                    "type": c.get("type"),
                    "title": _norm_ws(c.get("title", "")),
                    "tag": _norm_ws(c.get("tag", "")),
                }
            )
        controls = sorted(
            controls, key=lambda x: (str(x.get("type")), x.get("title", ""), x.get("tag", ""))
        )
        return {
            "content_control_count": info.get("content_control_count", 0),
            "controls": controls,
        }

    gold_sig = _cc_signature(gold_cc)
    pred_sig = _cc_signature(pred_cc)
    if gold_sig != pred_sig:
        return VerificationResult(
            passed=False,
            message="FAIL: Cover page content-control signature differs",
            details={
                "match_mode": "cover_page_signature",
                "gold": gold_sig,
                "pred": pred_sig,
            },
        )

    gold_layout = families_pages_extract_saved.cover_page_ooxml_layout_signature(gold_path)
    pred_layout = families_pages_extract_saved.cover_page_ooxml_layout_signature(pred_path)
    if gold_layout != pred_layout:
        return VerificationResult(
            passed=False,
            message="FAIL: Cover page saved layout signature differs",
            details={
                "match_mode": "cover_page_signature",
                "gold_layout": gold_layout,
                "pred_layout": pred_layout,
            },
        )

    gold_text = _norm_ws(read_document_text(str(gold_path)))
    pred_text = _norm_ws(read_document_text(str(pred_path)))
    if gold_text != pred_text:
        return VerificationResult(
            passed=False,
            message="FAIL: Cover page document text differs (whitespace normalized)",
            details={
                "match_mode": "cover_page_signature",
                "text_compare": "whitespace_normalized",
                "gold_text": gold_text[:300] + "..." if len(gold_text) > 300 else gold_text,
                "pred_text": pred_text[:300] + "..." if len(pred_text) > 300 else pred_text,
            },
        )

    # Compare content control fill values
    def _cc_value_map(cc_info):
        result = {}
        for cc in cc_info.get("controls", []):
            key = cc.get("title") or cc.get("tag") or ""
            if key:
                result[key] = (cc.get("text") or "").strip()
        return result

    gold_values = _cc_value_map(gold_cc)
    pred_values = _cc_value_map(pred_cc)

    from ..content_controls.values import filled_value_mismatches
    cc_mismatches = filled_value_mismatches(gold_cc.get("controls", []), pred_cc.get("controls", []))

    if cc_mismatches:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Cover page content control fill values differ ({len(cc_mismatches)} field(s))",
            details={
                "match_mode": "cover_page_signature",
                "mismatches": cc_mismatches,
                "gold_values": gold_values,
                "pred_values": pred_values,
            },
        )

    return VerificationResult(
        passed=True,
        message="PASS: Cover page style matches",
        details={
            "match_mode": "cover_page_signature",
            "content_control_count": gold_sig["content_control_count"],
            "body_paragraph_role_count": len(gold_layout["paragraph_roles"]),
            "drawing_anchor_count": len(gold_layout["anchors"]),
            "normalized_text_length": len(gold_text),
            "page_count": pred_page_count,
        },
    )
