"""families / content_controls / verify with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.families.references import (
    extract_saved as families_references_extract_saved,
)
from evaluation.methods.word_object.policies import (
    content_operations as policies_content_operations,
)
from pathlib import Path
from typing import Any, Dict, Optional
import re
from ..references.fields import require_reference_sections


def verify_content_controls(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    read_content_controls,
    read_dynamic_reference_state,
) -> VerificationResult:
    """Verify content control count and fill values match."""
    parameter_name = str(((spec or {}).get("parameter") or {}).get("name", ""))
    variation = str((spec or {}).get("variation", ""))
    primary = ((spec or {}).get("verification") or {}).get("primary") or {}
    manual_toc = ("manual" in parameter_name.casefold() or "manual" in variation.casefold()
                  or primary.get("method") == "content_control_verification")
    if (spec or {}).get("idMso") == "TableOfContentsGallery" and manual_toc:
        gold = read_dynamic_reference_state(str(gold_path))
        pred = read_dynamic_reference_state(str(pred_path))
        error = require_reference_sections({"gold": gold, "pred": pred}, ('content_controls', 'paragraphs', 'tocs'))
        if error is not None:
            return error

        def manual_signature(state: Dict[str, Any]) -> Dict[str, Any]:
            controls = state.get("content_controls", [])
            outers = [
                control
                for control in controls
                if (
                    "table of contents" in (control.get("building_block_gallery") or "").casefold()
                    or control.get("type") == 14
                )
            ]
            if len(outers) == 1:
                outer = outers[0]
                inner = [
                    control
                    for control in controls
                    if control is not outer
                    and (
                        outer.get("start") is None
                        or control.get("start") is None
                        or (
                            outer.get("start") <= control.get("start")
                            and control.get("end") <= outer.get("end")
                        )
                    )
                ]
            else:
                outer = None
                inner = controls if len(controls) == 6 else []
            if outer and outer.get("start") is not None:
                manual_paragraphs = [
                    paragraph
                    for paragraph in state.get("paragraphs", [])
                    if paragraph.get("start") is not None
                    and outer.get("start") <= paragraph.get("start") < outer.get("end")
                ]
            else:
                manual_paragraphs = state.get("paragraphs", [])
            heading_styles = [
                paragraph.get("style")
                for paragraph in manual_paragraphs
                if "toc heading" in (paragraph.get("style") or "").casefold()
            ]
            entry_levels = []
            for paragraph in manual_paragraphs:
                match = re.search(r"TOC\s*(\d+)", paragraph.get("style") or "", re.I)
                if match:
                    entry_levels.append(int(match.group(1)))
            return {
                "automatic_toc_count": len(state.get("tocs", [])),
                "outer_count": len(outers),
                "inner_count": len(inner),
                "inner_types": sorted(control.get("type") for control in inner),
                "placeholder_count": sum(
                    1 for control in inner if control.get("showing_placeholder")
                ),
                "heading_count": len(heading_styles),
                "entry_levels": entry_levels,
            }

        gold_signature = manual_signature(gold)
        pred_signature = manual_signature(pred)
        gold_saved = families_references_extract_saved.extract_saved_manual_toc_state(gold_path)
        pred_saved = families_references_extract_saved.extract_saved_manual_toc_state(pred_path)
        gold_com_complete = (
            gold_signature["automatic_toc_count"] == 0
            and gold_signature["inner_count"] == 6
            and gold_signature["placeholder_count"] == 6
            and gold_signature["heading_count"] == 1
            and len(gold_signature["entry_levels"]) == 6
        )
        gold_saved_complete = (
            gold_saved is not None
            and gold_saved["outer_count"] == 1
            and gold_saved["inner_count"] == 6
            and gold_saved["placeholder_count"] == 6
            and gold_saved["heading_count"] == 1
            and len(gold_saved["entry_levels"]) == 6
            and gold_saved["toc_field_count"] == 0
        )
        # Unit fakes do not have a package path; their explicit COM state
        # includes the outer type-14 control instead.
        fake_complete = (
            gold_saved is None and gold_signature["outer_count"] == 1 and gold_com_complete
        )
        if not gold_com_complete or not (gold_saved_complete or fake_complete):
            return policies_content_operations.error_with_details(
                "Gold lacks a complete editable manual-TOC building-block structure",
                {
                    "gold_signature": gold_signature,
                    "gold_saved": gold_saved,
                    "gold": gold,
                },
            )
        pred_com_signature = {
            key: value for key, value in pred_signature.items() if key != "outer_count"
        }
        gold_com_signature = {
            key: value for key, value in gold_signature.items() if key != "outer_count"
        }
        if pred_com_signature != gold_com_signature:
            return VerificationResult(
                passed=False,
                message="FAIL: Manual TOC controls, placeholders, heading, or levels differ",
                details={
                    "gold_signature": gold_signature,
                    "pred_signature": pred_signature,
                    "pred": pred,
                },
            )
        if gold_saved_complete:
            if pred_saved is None or pred_saved != gold_saved:
                return VerificationResult(
                    passed=False,
                    message="FAIL: Submitted manual-TOC building-block OOXML differs",
                    details={
                        "gold_saved": gold_saved,
                        "pred_saved": pred_saved,
                    },
                )
        elif pred_signature.get("outer_count") != 1:
            return VerificationResult(
                passed=False,
                message="FAIL: Manual TOC lacks its outer gallery control",
                details={"pred_signature": pred_signature},
            )
        if pred.get("source_headings", []) != gold.get("source_headings", []):
            return VerificationResult(
                passed=False,
                message="FAIL: Manual TOC insertion changed source headings",
                details={
                    "gold_source_headings": gold.get("source_headings", []),
                    "pred_source_headings": pred.get("source_headings", []),
                },
            )
        if gold_saved_complete:
            from ...preservation.context import extract_saved_semantic_context
            gold_context = extract_saved_semantic_context(gold_path, strip_manual_toc=True)
            pred_context = extract_saved_semantic_context(pred_path, strip_manual_toc=True)
            if gold_context != pred_context:
                return VerificationResult(
                    passed=False, message="FAIL: Manual TOC insertion changed surrounding document content/layout",
                    details={"preservation_scope": "outside_manual_toc", "gold": gold_context, "pred": pred_context},
                )
        return VerificationResult(
            passed=True,
            message="PASS: Editable manual-TOC building-block structure matches",
            details={
                "match_mode": "manual_toc_outer_gallery_inner_placeholders_levels",
                "signature": pred_signature,
                "saved_signature": pred_saved,
            },
        )

    gold_info = read_content_controls(str(gold_path))
    pred_info = read_content_controls(str(pred_path))

    if gold_info["content_control_count"] != pred_info["content_control_count"]:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Content control count mismatch (pred={pred_info['content_control_count']}, gold={gold_info['content_control_count']})",
            details={"pred": pred_info, "gold": gold_info},
        )

    # Compare fill values: build {title_or_tag: text} for each
    def _cc_value_map(controls):
        result = {}
        for cc in controls:
            key = cc.get("title") or cc.get("tag") or ""
            if key:
                result[key] = (cc.get("text") or "").strip()
        return result

    gold_values = _cc_value_map(gold_info.get("controls", []))
    pred_values = _cc_value_map(pred_info.get("controls", []))

    from .values import filled_value_mismatches
    mismatches = filled_value_mismatches(gold_info.get("controls", []), pred_info.get("controls", []))

    if mismatches:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Content control fill value mismatch ({len(mismatches)} field(s))",
            details={
                "mismatches": mismatches,
                "gold_values": gold_values,
                "pred_values": pred_values,
            },
        )

    return VerificationResult(
        passed=True,
        message=f"PASS: Content controls match ({pred_info['content_control_count']} controls, {len(gold_values)} filled)",
        details={
            "content_control_count": pred_info["content_control_count"],
            "filled_values": pred_values,
        },
    )
