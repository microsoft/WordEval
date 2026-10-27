"""families / formatting / paragraphs with explicit evidence dependencies."""

from .text import _bullet_chars_equivalent
from .contract import required_fields, NULLABLE_FIELDS
from ...observations import require_observations
from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.composition import (
    composition_command_ids,
    composition_operation_ids,
)
from evaluation.methods.word_object.families.formatting import (
    compare as families_formatting_compare,
)
from evaluation.methods.word_object.families.formatting import (
    composition as families_formatting_composition,
)
from evaluation.methods.word_object.families.formatting import (
    preservation as families_formatting_preservation,
)
from pathlib import Path
from typing import Dict, Optional


def verify_dropcap(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    read_document_text,
    read_dropcap_properties,
) -> VerificationResult:
    """Verify drop cap properties on the first paragraph match gold.

    DropCapInsertGallery applies a paragraph-level formatting attribute
    (Paragraphs(1).DropCap.Position / .LinesToDrop), so it lives in the
    content_formatting verifier alongside the rest of paragraph formatting.
    """
    content_result = families_formatting_preservation.check_content_integrity(
        pred_path, gold_path, read_document_text=read_document_text
    )
    if content_result is not None:
        return content_result

    gold_props = read_dropcap_properties(str(gold_path))
    pred_props = read_dropcap_properties(str(pred_path))

    requested = {"position"}
    primary_params = (spec or {}).get("verification", {}).get("primary", {}).get("params", {})
    parameter = (spec or {}).get("parameter", {}) or {}
    for prop in ("lines_to_drop", "distance_from_text"):
        if prop in primary_params or prop in parameter:
            requested.add(prop)
    for intent in (spec or {}).get("composition", {}).get("atomic_intents", []):
        for target in intent.get("verification_targets", []):
            prop = target.get("property")
            if prop in {"lines_to_drop", "distance_from_text"}:
                requested.add(prop)

    observation_error = require_observations(
        {"gold": gold_props, "pred": pred_props}, requested, context="drop cap")
    if observation_error is not None:
        return observation_error
    differences = []
    for prop in (
        "position",
        "paragraph_index",
        "paragraph_start",
        "paragraph_text",
        "target_character",
    ):
        if gold_props.get(prop) != pred_props.get(prop):
            differences.append(
                {
                    "property": prop,
                    "gold": gold_props.get(prop),
                    "pred": pred_props.get(prop),
                }
            )
    if "lines_to_drop" in requested and (
        gold_props.get("lines_to_drop") != pred_props.get("lines_to_drop")
    ):
        differences.append(
            {
                "property": "lines_to_drop",
                "gold": gold_props.get("lines_to_drop"),
                "pred": pred_props.get("lines_to_drop"),
            }
        )
    if "distance_from_text" in requested:
        gold_distance = gold_props.get("distance_from_text")
        pred_distance = pred_props.get("distance_from_text")
        try:
            distance_matches = abs(float(gold_distance) - float(pred_distance)) <= 0.05
        except (TypeError, ValueError):
            distance_matches = gold_distance == pred_distance
        if not distance_matches:
            differences.append(
                {
                    "property": "distance_from_text",
                    "gold": gold_distance,
                    "pred": pred_distance,
                }
            )

    return VerificationResult(
        passed=not differences,
        message=(
            "PASS: Drop cap kind, anchor, target character, and requested " "properties match"
            if not differences
            else f"FAIL: Drop cap semantics differ ({len(differences)} check(s))"
        ),
        details={
            "pred": pred_props,
            "gold": gold_props,
            "requested_properties": sorted(requested),
            "differences": differences,
        },
    )


def verify_all_paragraphs(
    pred_path: Path,
    gold_path: Path,
    spec: Dict,
    *,
    read_all_paragraph_character_formatting,
    read_all_paragraphs_formatting,
    read_document_text,
    read_dropcap_properties,
    read_text_formatting,
) -> VerificationResult:
    """
    Verify formatting across all paragraphs matches between pred and gold.

    Used when select_text is "all" - e.g., LineSpacingGallery, BulletsGalleryWord.
    """
    # Content integrity check first
    content_result = families_formatting_preservation.check_content_integrity(
        pred_path, gold_path, read_document_text=read_document_text
    )
    if content_result is not None:
        families_formatting_composition.add_composition_atomic_details(
            content_result.details,
            pred_path,
            gold_path,
            spec,
            "all",
            read_all_paragraphs_formatting=read_all_paragraphs_formatting,
            read_document_text=read_document_text,
            read_dropcap_properties=read_dropcap_properties,
            read_text_formatting=read_text_formatting,
        )
        return content_result

    # Extract all paragraphs from both documents
    gold_paras = read_all_paragraphs_formatting(str(gold_path))
    pred_paras = read_all_paragraphs_formatting(str(pred_path))

    for owner, paragraphs in (("gold", gold_paras), ("pred", pred_paras)):
        if paragraphs is None:
            error = f"Could not extract paragraphs from {owner} document"
            return VerificationResult(
                passed=False,
                message=f"ERROR: {error}",
                details={"select_text": "all", "observation_owner": owner},
                error=error,
            )

    # Check paragraph counts match
    if len(gold_paras) != len(pred_paras):
        return VerificationResult(
            passed=False,
            message=f"FAIL: Paragraph count mismatch (gold: {len(gold_paras)}, pred: {len(pred_paras)})",
            details={
                "gold_count": len(gold_paras),
                "pred_count": len(pred_paras),
                "select_text": "all",
            },
        )

    # Determine which properties to compare based on command type
    idMso = spec.get("idMso", "")

    is_composition = bool(spec.get("composition"))

    # List gallery commands: focus on semantic list properties, not indent details
    # The instruction specifies the list style (e.g., "Roman numerals with period")
    # but doesn't specify exact indent/alignment which are Word template defaults
    LIST_GALLERY_COMMANDS = {
        "BulletsGalleryWord",
        "NumberingGalleryWord",
        "MultilevelListGallery",
    }
    FULL_PARAGRAPH_PROPERTIES = [
        "alignment",
        "left_indent",
        "right_indent",
        "first_line_indent",
        "line_spacing",
        "line_spacing_rule",
        "space_before",
        "space_after",
        "paragraph_borders",
        "dropcap_position",
        "dropcap_lines_to_drop",
        "list_type",
        "list_level",
        "list_string",
    ]
    PARAGRAPH_SPACING_TARGETS = {
        "ParagraphSpacingBefore": {"space_before"},
        "ParagraphSpacingAfter": {"space_after"},
    }
    target_properties = None

    if is_composition:
        # Composition specs should verify the union of the semantic
        # properties their subcommands explicitly test. Do not compare the
        # full paragraph bundle by default: Word list galleries set
        # template-specific indents/list_type as side effects that atomic
        # list verifiers intentionally ignore.
        composition_cmds = composition_command_ids(spec)
        composition_steps = spec.get("composition", {}).get("steps", [])
        step_cmds = [
            step.get("idMso")
            for step in composition_steps
            if isinstance(step, dict) and step.get("idMso")
        ]
        step_ops = composition_operation_ids(spec)
        all_cmds = set(composition_cmds) | set(step_cmds)
        all_ops = set(step_ops)
        para_properties = []

        if (
            "BulletsGalleryWord" in all_cmds
            or "NumberingGalleryWord" in all_cmds
            or "apply_bullets" in all_ops
            or "apply_numbering" in all_ops
        ):
            para_properties.append("list_string")
        if (
            "MultilevelListGallery" in all_cmds
            or "set_list_level" in all_ops
            or "set_list_level_range" in all_ops
        ):
            para_properties.extend(["list_string", "list_level"])
        if "LineSpacingGallery" in all_cmds:
            para_properties.extend(["line_spacing", "line_spacing_rule"])
        if any(
            cmd in all_cmds
            for cmd in [
                "ParagraphAlignmentGallery",
                "AlignLeft",
                "AlignCenter",
                "AlignRight",
                "AlignJustify",
            ]
        ):
            para_properties.append("alignment")
        if any(
            cmd in all_cmds
            for cmd in [
                "IndentIncreaseWord",
                "IndentDecreaseWord",
                "ParagraphIndentGallery",
                "ListIndentIncrease",
                "ListIndentDecrease",
            ]
        ):
            para_properties.extend(["left_indent", "first_line_indent"])
        if any(
            cmd in all_cmds
            for cmd in [
                "ParagraphSpacingBeforeGallery",
                "ParagraphSpacingAfterGallery",
                "ParagraphSpacingBefore",
                "ParagraphSpacingAfter",
            ]
        ):
            para_properties.extend(["space_before", "space_after"])
        if "BordersSelectionGallery" in all_cmds:
            para_properties.append("paragraph_borders")
        if "DropCapInsertGallery" in all_cmds:
            para_properties.extend(["dropcap_position", "dropcap_lines_to_drop"])

        # Unknown composition command: fall back to the historical full
        # comparison rather than silently under-checking.
        if not para_properties:
            para_properties = FULL_PARAGRAPH_PROPERTIES
        else:
            para_properties = list(dict.fromkeys(para_properties))
    elif idMso in LIST_GALLERY_COMMANDS:
        # For list galleries, only verify the semantic list properties
        # NOT checking:
        # - list_type: Internal Word categorization (3=SimpleNumbering, 4=OutlineNumbering)
        #   Both can produce identical visual output, so this is Word trivia, not semantics
        # - Indent/alignment: Word-specific template defaults not mentioned in instruction
        if idMso == "MultilevelListGallery":
            # For multilevel lists, also verify correct nesting level
            para_properties = ["list_string", "list_level"]
        else:
            # For single-level lists (NumberingGalleryWord, BulletsGalleryWord)
            # Only the visual format matters
            para_properties = ["list_string"]
    elif idMso in PARAGRAPH_SPACING_TARGETS:
        # Spacing commands explicitly target only SpaceBefore/SpaceAfter,
        # but the generator's API action should not alter the rest of the
        # paragraph environment. Keep the invariant check, while labelling
        # target vs invariant failures for downstream analysis.
        target_properties = PARAGRAPH_SPACING_TARGETS[idMso]
        para_properties = list(target_properties) + [
            prop for prop in FULL_PARAGRAPH_PROPERTIES if prop not in target_properties
        ]
    else:
        # Full property comparison for other commands
        para_properties = FULL_PARAGRAPH_PROPERTIES

    # For BulletsGalleryWord (or compositions that include it), normalize
    # bullet characters before comparison.
    # Word's gallery uses legacy PUA chars; models may use Unicode equivalents
    composition_cmds = composition_command_ids(spec) if is_composition else []
    composition_ops = composition_operation_ids(spec) if is_composition else []
    normalize_bullets = (
        idMso == "BulletsGalleryWord"
        or "BulletsGalleryWord" in composition_cmds
        or "apply_bullets" in composition_ops
    )

    # Compare formatting for each paragraph
    differences = []

    for i, (gold_para, pred_para) in enumerate(zip(gold_paras, pred_paras)):
        required = required_fields(spec, "all_paragraphs_formatting")
        if idMso in LIST_GALLERY_COMMANDS:
            required.add("list_string")
        if idMso == "MultilevelListGallery":
            required.add("list_level")
        if idMso == "LineSpacingGallery":
            required.add("line_spacing_rule")
        if is_composition:
            required.update(para_properties)
        observation_error = require_observations(
            {"gold": gold_para, "pred": pred_para}, required,
            context=f"paragraph {i + 1}", nullable=NULLABLE_FIELDS)
        if observation_error is not None:
            return observation_error
        for prop in para_properties:
            gold_val = gold_para.get(prop)
            pred_val = pred_para.get(prop)
            comparison = (
                families_formatting_compare.composition_indent_comparison(prop, spec)
                if is_composition
                else None
            )
            values_match = families_formatting_compare.reference_comparison_matches(
                gold_val, pred_val, comparison
            )
            if normalize_bullets and prop == "list_string":
                values_match = _bullet_chars_equivalent(
                    gold_val,
                    pred_val,
                    gold_para.get("list_font"),
                    pred_para.get("list_font"),
                )
            if not values_match:
                differences.append(
                    {
                        "property": prop,
                        "comparison": comparison or "exact",
                        "role": (
                            "target"
                            if target_properties and prop in target_properties
                            else "invariant"
                        ),
                        "gold": gold_val,
                        "pred": pred_val,
                        "location": f"paragraph {i + 1}",
                    }
                )

    if differences:
        target_differences = [diff for diff in differences if diff.get("role") == "target"]
        invariant_differences = [diff for diff in differences if diff.get("role") == "invariant"]
        details = {
            "differences": differences[:10],  # Limit to first 10
            "total_differences": len(differences),
            "select_text": "all",
            "paragraph_count": len(gold_paras),
        }
        if target_properties:
            details.update(
                {
                    "target_properties": sorted(target_properties),
                    "target_differences": target_differences[:10],
                    "invariant_differences": invariant_differences[:10],
                    "target_difference_count": len(target_differences),
                    "invariant_difference_count": len(invariant_differences),
                }
            )
        families_formatting_composition.add_composition_atomic_details(
            details,
            pred_path,
            gold_path,
            spec,
            "all",
            read_all_paragraphs_formatting=read_all_paragraphs_formatting,
            read_document_text=read_document_text,
            read_dropcap_properties=read_dropcap_properties,
            read_text_formatting=read_text_formatting,
        )
        if target_properties and not target_differences:
            message = (
                f"FAIL: {len(invariant_differences)} invariant formatting "
                "difference(s) across paragraphs"
            )
        elif target_properties:
            message = (
                f"FAIL: {len(target_differences)} target and "
                f"{len(invariant_differences)} invariant formatting "
                "difference(s) across paragraphs"
            )
        else:
            message = f"FAIL: {len(differences)} formatting difference(s) " "across paragraphs"
        return VerificationResult(passed=False, message=message, details=details)

    # Paragraph/list commands do not target character formatting. Preserve
    # visible run-level state so a prediction cannot pass after also making
    # every paragraph bold, recoloring text, adding effects, and so on.
    preservation_result = families_formatting_preservation.check_character_formatting_preservation(
        pred_path,
        gold_path,
        spec,
        "all",
        read_all_paragraph_character_formatting=read_all_paragraph_character_formatting,
        read_all_paragraphs_formatting=read_all_paragraphs_formatting,
        read_document_text=read_document_text,
        read_dropcap_properties=read_dropcap_properties,
        read_text_formatting=read_text_formatting,
    )
    if preservation_result is not None:
        return preservation_result

    details = {"select_text": "all", "paragraph_count": len(gold_paras)}
    families_formatting_composition.add_composition_atomic_details(
        details,
        pred_path,
        gold_path,
        spec,
        "all",
        read_all_paragraphs_formatting=read_all_paragraphs_formatting,
        read_document_text=read_document_text,
        read_dropcap_properties=read_dropcap_properties,
        read_text_formatting=read_text_formatting,
    )
    return VerificationResult(
        passed=True,
        message=f"PASS: All {len(gold_paras)} paragraphs have matching formatting",
        details=details,
    )


def verify_specific_paragraph(
    pred_path: Path,
    gold_path: Path,
    paragraph_index: int,
    spec: Dict,
    *,
    read_all_paragraph_character_formatting,
    read_all_paragraphs_formatting,
    read_document_text,
    read_dropcap_properties,
    read_paragraph_formatting,
    read_text_formatting,
) -> VerificationResult:
    """
    Verify formatting of a specific paragraph matches between pred and gold.

    Used when select_text is "paragraph_N" - e.g., ListLevelGallery.
    """
    # Content integrity check first
    content_result = families_formatting_preservation.check_content_integrity(
        pred_path, gold_path, read_document_text=read_document_text
    )
    if content_result is not None:
        families_formatting_composition.add_composition_atomic_details(
            content_result.details,
            pred_path,
            gold_path,
            spec,
            f"paragraph_{paragraph_index}",
            read_all_paragraphs_formatting=read_all_paragraphs_formatting,
            read_document_text=read_document_text,
            read_dropcap_properties=read_dropcap_properties,
            read_text_formatting=read_text_formatting,
        )
        return content_result

    # Extract specific paragraph from both documents
    gold_para = read_paragraph_formatting(str(gold_path), paragraph_index)
    pred_para = read_paragraph_formatting(str(pred_path), paragraph_index)

    if gold_para is None:
        return VerificationResult(
            passed=False,
            message=f"ERROR: Paragraph {paragraph_index} not found in gold document",
            error=f"Paragraph {paragraph_index} not found in gold document",
            details={"paragraph_index": paragraph_index},
        )

    if pred_para is None:
        return VerificationResult(
            passed=False,
            message=f"FAIL: Paragraph {paragraph_index} not found in pred document",
            details={"paragraph_index": paragraph_index},
        )

    # Compare formatting properties
    differences = []
    idMso = spec.get("idMso", "")
    if idMso == "ListLevelGallery":
        para_properties = ["list_level", "list_string"]
    else:
        para_properties = [
            "alignment",
            "left_indent",
            "right_indent",
            "first_line_indent",
            "line_spacing",
            "line_spacing_rule",
            "space_before",
            "space_after",
            "paragraph_borders",
            "dropcap_position",
            "dropcap_lines_to_drop",
            "list_type",
            "list_level",
            "list_string",
        ]

    required = required_fields(spec, "paragraph_formatting")
    if idMso == "ListLevelGallery":
        required.update(("list_level", "list_string"))
    observation_error = require_observations(
        {"gold": gold_para, "pred": pred_para}, required, context="selected paragraph",
        nullable=NULLABLE_FIELDS)
    if observation_error is not None:
        return observation_error

    for prop in para_properties:
        gold_val = gold_para.get(prop)
        pred_val = pred_para.get(prop)
        if gold_val != pred_val:
            differences.append(
                {
                    "property": prop,
                    "gold": gold_val,
                    "pred": pred_val,
                    "location": f"paragraph {paragraph_index}",
                }
            )

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: {len(differences)} formatting difference(s) on paragraph {paragraph_index}",
            details={
                "differences": differences,
                "paragraph_index": paragraph_index,
                "gold": gold_para,
                "pred": pred_para,
            },
        )

    preservation_result = families_formatting_preservation.check_character_formatting_preservation(
        pred_path,
        gold_path,
        spec,
        f"paragraph_{paragraph_index}",
        read_all_paragraph_character_formatting=read_all_paragraph_character_formatting,
        read_all_paragraphs_formatting=read_all_paragraphs_formatting,
        read_document_text=read_document_text,
        read_dropcap_properties=read_dropcap_properties,
        read_text_formatting=read_text_formatting,
    )
    if preservation_result is not None:
        return preservation_result

    return VerificationResult(
        passed=True,
        message=f"PASS: Paragraph {paragraph_index} formatting matches",
        details={"paragraph_index": paragraph_index, "values": gold_para},
    )
