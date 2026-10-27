"""families / references / verify with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.policies import (
    content_operations as policies_content_operations,
)
from pathlib import Path
from typing import Any, Dict, List, Optional
import re
from .fields import require_reference_sections


def verify_index(
    pred_path: Path, gold_path: Path, *, read_dynamic_reference_state
) -> VerificationResult:
    """Verify a real INDEX object against submitted and refreshed XE state."""
    gold = read_dynamic_reference_state(str(gold_path))
    pred = read_dynamic_reference_state(str(pred_path))
    error = require_reference_sections({"gold": gold, "pred": pred}, ('indexes', 'xe_entries', 'source_headings', 'section_count', 'headers', 'footers'))
    if error is not None:
        return error

    def normalized_entries(index: Dict[str, Any]) -> List[str]:
        values = []
        for line in index.get("entries", []):
            value = " ".join((line or "").split())
            value = re.sub(
                r"(?:[,\t]\s*)?\d+(?:\s*[-\u2013,]\s*\d+)*\s*$",
                "",
                value,
            ).strip(" ,\t")
            if value:
                values.append(value)
        return values

    gold_indexes = gold.get("indexes", [])
    gold_xe = sorted(gold.get("xe_entries", []))
    if len(gold_indexes) != 1 or not gold_xe:
        return policies_content_operations.error_with_details(
            "Gold lacks one INDEX object or XE source entries",
            {"gold": gold},
        )
    gold_entries = normalized_entries(gold_indexes[0])
    if not gold_entries or not any("cherry" in value.casefold() for value in gold_entries):
        return policies_content_operations.error_with_details(
            "Gold submitted index does not contain the post-index Cherry entry",
            {"gold": gold, "gold_entries": gold_entries},
        )
    pred_indexes = pred.get("indexes", [])
    if len(pred_indexes) != 1:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Expected one real INDEX object, found {len(pred_indexes)}",
            details={"gold": gold, "pred": pred},
        )
    pred_entries = normalized_entries(pred_indexes[0])
    if pred_entries != gold_entries:
        return VerificationResult(
            passed=False,
            message="FAIL: Submitted index cache is stale, incomplete, or wrong",
            details={
                "gold_entries": gold_entries,
                "pred_entries": pred_entries,
                "submitted_pred": pred_indexes[0],
            },
        )
    pred_xe = sorted(pred.get("xe_entries", []))
    if pred_xe != gold_xe:
        return VerificationResult(
            passed=False,
            message="FAIL: Index XE source entries were changed or removed",
            details={"gold_xe_entries": gold_xe, "pred_xe_entries": pred_xe},
        )
    if not re.search(r"^\s*INDEX\b", pred_indexes[0].get("code") or "", re.I):
        return VerificationResult(
            passed=False,
            message="FAIL: Index text is not backed by a real INDEX field",
            details={"pred_index": pred_indexes[0]},
        )
    refreshed = read_dynamic_reference_state(str(pred_path), update_fields=True)
    error = require_reference_sections({"refreshed": refreshed}, ('indexes', 'xe_entries', 'source_headings', 'section_count', 'headers', 'footers'))
    if error is not None:
        return error
    refreshed_indexes = refreshed.get("indexes", [])
    refreshed_entries = (
        normalized_entries(refreshed_indexes[0]) if len(refreshed_indexes) == 1 else []
    )
    if len(refreshed_indexes) != 1 or refreshed_entries != gold_entries:
        return VerificationResult(
            passed=False,
            message="FAIL: Refreshed INDEX entries differ from the XE-backed gold semantics",
            details={
                "gold_entries": gold_entries,
                "submitted_pred_entries": pred_entries,
                "refreshed_pred_entries": refreshed_entries,
            },
        )
    for key in ("source_headings", "section_count", "headers", "footers"):
        if pred.get(key) != gold.get(key):
            return VerificationResult(
                passed=False,
                message=f"FAIL: Index update changed preserved {key}",
                details={"gold": gold.get(key), "pred": pred.get(key)},
            )
    return VerificationResult(
        passed=True,
        message="PASS: INDEX object, submitted cache, refreshed entries, and XE sources match",
        details={
            "match_mode": "index_object_submitted_refreshed_xe_sources",
            "entries": pred_entries,
            "xe_entries": pred_xe,
        },
    )


def verify_citation(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, read_dynamic_reference_state
) -> VerificationResult:
    """Verify one anchored CITATION field, its source, and visible result."""
    gold = read_dynamic_reference_state(str(gold_path))
    pred = read_dynamic_reference_state(str(pred_path))
    error = require_reference_sections({"gold": gold, "pred": pred}, ('citation_fields', 'bibliography_sources'))
    if error is not None:
        return error
    gold_fields = gold.get("citation_fields", [])
    pred_fields = pred.get("citation_fields", [])
    expected_source_tag = (spec or {}).get("source_tag") or (
        (spec or {}).get("parameter") or {}
    ).get("source_tag")
    expected_source_tag = expected_source_tag or "Smith2020"

    def citation_tag(field: Dict[str, Any]) -> str:
        match = re.search(
            r"^\s*CITATION\s+([^\s\\]+)",
            field.get("code") or "",
            re.I,
        )
        return match.group(1) if match else ""

    def source_by_tag(state: Dict[str, Any]) -> List[Dict[str, Any]]:
        return [
            source
            for source in state.get("bibliography_sources", [])
            if (source.get("tag") or "").casefold() == expected_source_tag.casefold()
        ]

    gold_sources = source_by_tag(gold)
    if len(gold_fields) != 1 or len(gold_sources) != 1:
        return policies_content_operations.error_with_details(
            "Gold must contain one CITATION field and one matching bibliography source",
            {
                "gold_fields": gold_fields,
                "gold_sources": gold.get("bibliography_sources", []),
                "expected_source_tag": expected_source_tag,
            },
        )
    if len(pred_fields) != 1:
        return VerificationResult(
            passed=False,
            message=(
                "FAIL: Pred must contain exactly one CITATION field "
                f"(found {len(pred_fields)}); plain text and duplicates are not accepted"
            ),
            details={"gold_fields": gold_fields, "pred_fields": pred_fields},
        )
    pred_field = pred_fields[0]
    if citation_tag(pred_field).casefold() != expected_source_tag.casefold():
        return VerificationResult(
            passed=False,
            message=f"FAIL: CITATION field does not reference {expected_source_tag}",
            details={"pred_field": pred_field},
        )
    pred_sources = source_by_tag(pred)
    if len(pred_sources) != 1:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Document lacks one real source tagged {expected_source_tag}",
            details={"pred_sources": pred.get("bibliography_sources", [])},
        )
    gold_tags = sorted(source.get("tag") or "" for source in gold.get("bibliography_sources", []))
    pred_tags = sorted(source.get("tag") or "" for source in pred.get("bibliography_sources", []))
    if pred_tags != gold_tags:
        return VerificationResult(
            passed=False,
            message="FAIL: Citation insertion changed bibliography sources",
            details={"gold_source_tags": gold_tags, "pred_source_tags": pred_tags},
        )

    # The visible in-text citation is rendered by Word from the source, so
    # its shape follows the installation's citation style and language
    # rather than anything the solver does: the same field reads
    # "(SmithJohn, 2020)" on a Korean install and "(Smith, 2020)" on an
    # English one. Comparing that string would fail a correct submission
    # for the environment it ran in. Which source was cited is already
    # settled by the field-code tag comparison above, which reads the code
    # and is therefore locale independent, so only require that the field
    # resolved to something visible.
    pred_result = (pred_field.get("result") or "").strip()
    if not pred_result:
        return VerificationResult(
            passed=False,
            message="FAIL: CITATION field did not resolve to a visible citation",
            details={"pred_result": pred_field.get("result")},
        )
    gold_field = gold_fields[0]
    if pred_field.get("story_type") not in (None, 1):
        return VerificationResult(
            passed=False,
            message="FAIL: Citation is not in the main document story",
            details={"pred_field": pred_field},
        )
    semantic_anchor_available = all(
        value is not None
        for value in (
            gold_field.get("paragraph_index"),
            pred_field.get("paragraph_index"),
            gold_field.get("paragraph_code_offset"),
            pred_field.get("paragraph_code_offset"),
        )
    )
    if semantic_anchor_available:
        beginning_offsets = {0, 1, 2}
        anchored = pred_field["paragraph_index"] == gold_field["paragraph_index"] and (
            pred_field["paragraph_code_offset"] == gold_field["paragraph_code_offset"]
            or gold_field["paragraph_index"] == 1
            and gold_field["paragraph_code_offset"] in beginning_offsets
            and pred_field["paragraph_code_offset"] in beginning_offsets
        )
        gold_anchor = {
            "paragraph_index": gold_field["paragraph_index"],
            "paragraph_code_offset": gold_field["paragraph_code_offset"],
        }
        pred_anchor = {
            "paragraph_index": pred_field["paragraph_index"],
            "paragraph_code_offset": pred_field["paragraph_code_offset"],
        }
        if anchored and pred_field["paragraph_code_offset"] != gold_field["paragraph_code_offset"]:
            pred_anchor["normalization"] = "word_beginning_field_boundary_offset"
    else:
        gold_position = gold_field.get("result_start", gold_field.get("start"))
        pred_position = pred_field.get("result_start", pred_field.get("start"))
        anchored = not (
            gold_position is not None
            and pred_position is not None
            and abs(pred_position - gold_position) > 1
        )
        gold_anchor = {"legacy_position": gold_position}
        pred_anchor = {"legacy_position": pred_position}
    if not anchored:
        return VerificationResult(
            passed=False,
            message="FAIL: Citation is not anchored at the beginning of the document body",
            details={"gold_anchor": gold_anchor, "pred_anchor": pred_anchor},
        )
    refreshed = read_dynamic_reference_state(str(pred_path), update_fields=True)
    error = require_reference_sections({"refreshed": refreshed}, ('citation_fields', 'bibliography_sources'))
    if error is not None:
        return error
    refreshed_fields = refreshed.get("citation_fields", [])
    refreshed_sources = [
        source
        for source in refreshed.get("bibliography_sources", [])
        if (source.get("tag") or "").casefold() == expected_source_tag.casefold()
    ]
    last_name = (pred_sources[0].get("last") or "").casefold()
    refreshed_result = refreshed_fields[0].get("result") or "" if len(refreshed_fields) == 1 else ""
    if (
        len(refreshed_fields) != 1
        or len(refreshed_sources) != 1
        or citation_tag(refreshed_fields[0]).casefold() != expected_source_tag.casefold()
        or (last_name and last_name not in refreshed_result.casefold())
    ):
        return VerificationResult(
            passed=False,
            message="FAIL: Refreshed citation no longer resolves through the requested source",
            details={
                "refreshed_fields": refreshed_fields,
                "refreshed_sources": refreshed_sources,
            },
        )
    return VerificationResult(
        passed=True,
        message="PASS: Unique anchored CITATION field, source, submitted result, and refresh verified",
        details={
            "match_mode": "citation_field_source_anchor_submitted_refreshed",
            "source_tag": expected_source_tag,
            "submitted_result": pred_field.get("result"),
            "refreshed_result": refreshed_result,
            "anchor": pred_anchor,
        },
    )


def verify_caption(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, read_dynamic_reference_state
) -> VerificationResult:
    """Verify a unique SEQ-backed caption and its object-relative placement."""
    gold = read_dynamic_reference_state(str(gold_path))
    pred = read_dynamic_reference_state(str(pred_path))
    error = require_reference_sections({"gold": gold, "pred": pred}, ('caption_fields', 'tables', 'inline_shapes', 'body_text', 'section_count', 'headers', 'footers'))
    if error is not None:
        return error
    params = (spec or {}).get("verification", {}).get("primary", {}).get("params", {})
    label = params.get("label", "")
    title = params.get("title", "")
    expected_caption = f"{label} 1{title}"

    def matching_fields(state: Dict[str, Any]) -> List[Dict[str, Any]]:
        return [
            field
            for field in state.get("caption_fields", [])
            if not label or (field.get("label") or "").casefold() == label.casefold()
        ]

    gold_fields = matching_fields(gold)
    pred_fields = matching_fields(pred)
    if len(gold_fields) != 1:
        return policies_content_operations.error_with_details(
            f"Gold lacks exactly one {label} SEQ caption field",
            {"gold": gold, "gold_fields": gold_fields},
        )
    if len(pred_fields) != 1:
        return VerificationResult(
            passed=False,
            message=(
                f"FAIL: Pred must contain exactly one {label} SEQ field "
                f"(found {len(pred_fields)})"
            ),
            details={"gold_fields": gold_fields, "pred_fields": pred_fields},
        )
    gold_field = gold_fields[0]
    pred_field = pred_fields[0]
    if (pred_field.get("result") or "").strip() != (gold_field.get("result") or "").strip():
        return VerificationResult(
            passed=False,
            message="FAIL: Caption sequence number differs",
            details={"gold_field": gold_field, "pred_field": pred_field},
        )
    paragraph_text = " ".join((pred_field.get("paragraph_text") or "").split())
    if paragraph_text and " ".join(expected_caption.split()) not in paragraph_text:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Caption text '{expected_caption}' is incomplete or wrong",
            details={"pred_field": pred_field},
        )
    execution_position = (spec or {}).get("execution", {}).get("position")
    requested_position = params.get("position")
    if execution_position == 0 or str(requested_position).casefold() == "above":
        relation = "above"
    elif execution_position == 1 or str(requested_position).casefold() == "below":
        relation = "below"
    else:
        relation = "below" if label.casefold() == "figure" else "above"

    caption_start = pred_field.get("paragraph_start")
    if caption_start is None:
        caption_start = pred_field.get("start") or pred_field.get("result_start")
    if label.casefold() == "figure":
        gold_objects = gold.get("inline_shapes", [])
        pred_objects = pred.get("inline_shapes", [])
    else:
        gold_objects = gold.get("tables", [])
        pred_objects = pred.get("tables", [])
    if not gold_objects:
        return policies_content_operations.error_with_details(
            f"Gold lacks the source {label.lower()} object",
            {"gold": gold},
        )
    if len(pred_objects) != len(gold_objects):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Source {label.lower()} object count changed",
            details={"gold_objects": gold_objects, "pred_objects": pred_objects},
        )
    target = pred_objects[0]
    if relation == "above":
        correct_side = (
            caption_start is not None
            and target.get("start") is not None
            and caption_start < target.get("start")
        )
    else:
        correct_side = (
            caption_start is not None
            and target.get("end") is not None
            and caption_start >= target.get("end")
        )
    if not correct_side:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Caption is not {relation} the source {label.lower()}",
            details={
                "relation": relation,
                "caption_start": caption_start,
                "source_object": target,
            },
        )

    def object_signature(objects: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        keys = (
            ("type", "width", "height")
            if label.casefold() == "figure"
            else ("row_count", "column_count", "cell_contents")
        )
        return [{key: obj.get(key) for key in keys} for obj in objects]

    if object_signature(pred_objects) != object_signature(gold_objects):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Caption insertion changed the source {label.lower()}",
            details={
                "gold_objects": object_signature(gold_objects),
                "pred_objects": object_signature(pred_objects),
            },
        )
    if " ".join((pred.get("body_text") or "").split()) != " ".join(
        (gold.get("body_text") or "").split()
    ):
        return VerificationResult(
            passed=False,
            message="FAIL: Caption insertion changed unrelated document text",
            details={
                "gold_text": (gold.get("body_text") or "")[:500],
                "pred_text": (pred.get("body_text") or "")[:500],
            },
        )
    for key in ("section_count", "headers", "footers"):
        if pred.get(key) != gold.get(key):
            return VerificationResult(
                passed=False,
                message=f"FAIL: Caption insertion changed preserved {key}",
                details={"gold": gold.get(key), "pred": pred.get(key)},
            )
    refreshed = read_dynamic_reference_state(str(pred_path), update_fields=True)
    error = require_reference_sections({"refreshed": refreshed}, ('caption_fields', 'tables', 'inline_shapes', 'body_text', 'section_count', 'headers', 'footers'))
    if error is not None:
        return error
    refreshed_fields = matching_fields(refreshed)
    if len(refreshed_fields) != 1:
        return VerificationResult(
            passed=False,
            message="FAIL: Caption field count changed after refresh",
            details={"refreshed_fields": refreshed_fields},
        )
    refreshed_field = refreshed_fields[0]
    refreshed_result = (refreshed_field.get("result") or "").strip()
    refreshed_paragraph = " ".join((refreshed_field.get("paragraph_text") or "").split())
    if refreshed_result != (gold_field.get("result") or "").strip() or (
        refreshed_paragraph and " ".join(expected_caption.split()) not in refreshed_paragraph
    ):
        return VerificationResult(
            passed=False,
            message="FAIL: Refreshed caption sequence or title is wrong",
            details={"refreshed_field": refreshed_field},
        )
    return VerificationResult(
        passed=True,
        message=f"PASS: Real {label} SEQ caption is correctly placed {relation}",
        details={
            "match_mode": "caption_seq_relative_placement_preservation_refreshed",
            "label": label,
            "title": title,
            "relation": relation,
            "caption_field": pred_field,
        },
    )
