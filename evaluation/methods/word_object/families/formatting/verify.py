"""families / formatting / verify with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from .contract import validate_formatting_observations
from ...observations import require_observations
from evaluation.methods.word_object.families.formatting import (
    compare as families_formatting_compare,
)
from evaluation.methods.word_object.families.formatting import (
    composition as families_formatting_composition,
)
from evaluation.methods.word_object.families.formatting import (
    preservation as families_formatting_preservation,
)
from evaluation.methods.word_object.families.formatting import (
    verify_saved as families_formatting_verify_saved,
)
from evaluation.methods.word_object.families.formatting import (
    visibility as families_formatting_visibility,
)
from evaluation.methods.word_object.policies import (
    content_formatting as policies_content_formatting,
)
from pathlib import Path
from typing import Dict, Optional


def verify_property_based(
    pred_path: Path,
    gold_path: Path,
    spec: Dict,
    *,
    read_all_paragraph_character_formatting,
    read_all_paragraphs_formatting,
    read_document_text,
    read_dropcap_properties,
    read_target_glyph_state,
    read_text_formatting,
) -> VerificationResult:
    """
    Verify formatting properties on a literal target text run.

    Caller contract (from `verify()`):
    - `spec` is non-None.
    - `select_text` is a literal text or the default 'Target'; special
      values 'all' and 'paragraph_N' are routed by `verify()` to
      `_verify_all_paragraphs` / `_verify_specific_paragraph` and never
      reach this method.
    - `idMso` is NOT one of the property_based-family special cases
      (LigatureGalleryWord, QuickStylesGallery); those are also routed
      by `verify()` to dedicated handlers.

    Even when `verification.primary.property_name` names a single
    property, this method compares every field returned by
    `extract_text_formatting` (with a small set of rendered-side-effect
    exclusions). The extra fields act as invariants: anything the
    instruction did not ask to change is expected to stay the same on
    the target text.
    """
    # Get target text from precondition
    precondition = spec.get("precondition", {})
    target_text = precondition.get("select_text", "Target")
    idMso = spec.get("idMso", "")

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
            target_text,
            read_all_paragraphs_formatting=read_all_paragraphs_formatting,
            read_document_text=read_document_text,
            read_dropcap_properties=read_dropcap_properties,
            read_text_formatting=read_text_formatting,
        )
        return content_result

    # Extract formatting from both documents (standard text search)
    gold_fmt = read_text_formatting(str(gold_path), target_text)
    pred_fmt = read_text_formatting(str(pred_path), target_text)

    if gold_fmt is None:
        return VerificationResult(
            passed=False,
            message=f'ERROR: Target text "{target_text}" not found in gold',
            error=f'Target text "{target_text}" not found in gold',
            details={"target_text": target_text},
        )

    if pred_fmt is None:
        return VerificationResult(
            passed=False,
            message=f'FAIL: Target text "{target_text}" not found in pred',
            details={"target_text": target_text},
        )

    # Compare formatting properties for target text
    differences = []
    # Properties that are internal rendered-position data, not formatting properties
    _rendered_keys = {"rendered_first_h", "rendered_last_h"}

    # For gallery commands that set multiple properties as a bundle,
    # only verify the property the instruction actually asks about.
    _OUTLINE_GALLERY_COMMANDS = {
        "TextOutlineDashesGallery",  # instruction asks about dash style
        "TextOutlineColorPickerAlternate",  # instruction asks about color
    }
    if idMso in _OUTLINE_GALLERY_COMMANDS:
        # These commands set weight as a side effect (gallery default).
        # Ignore exact weight differences when pred has a visible outline
        # (weight > 0), but still flag zero-weight outlines which are
        # visually invisible regardless of other properties.
        _rendered_keys = _rendered_keys | {"text_outline_weight"}
    if idMso == "TextOutlineDashesGallery":
        # The primary signal is text_outline_dash_style (if COM can read it).
        # Color is a side effect of the gallery preset, so skip it.
        _rendered_keys = _rendered_keys | {"text_outline_color"}
    if idMso == "TextFillGradientGallery":
        # Legacy Font.Color and Fill.ForeColor are incidental for a
        # gradient. The task-relevant WOM signals are fill type and angle;
        # persisted stop semantics are checked by the saved guardrail.
        _rendered_keys = _rendered_keys | {"font_color", "text_fill_color"}

    composition_intents = spec.get("composition", {}).get("atomic_intents", [])
    composition_operations = {intent.get("operation") for intent in composition_intents}
    if "drop_cap" in composition_operations:
        # Word derives the displayed initial's point size from the chosen
        # drop-cap kind and line count.  It is not an independently
        # requested character-formatting value.
        _rendered_keys = _rendered_keys | {"font_size"}
    if "fill_gradient" in composition_operations:
        # As with the atomic gradient command, legacy Font.Color and the
        # gradient object's ForeColor are incidental materialization
        # details. The composition contract names direction, not endpoint
        # colors.
        _rendered_keys = _rendered_keys | {"font_color", "text_fill_color"}
    character_shading_states = {
        families_formatting_compare.character_shading_state(intent.get("atomic_case"))
        for intent in composition_intents
        if intent.get("operation") == "char_shading"
    } - {None}
    if character_shading_states:
        # CharacterShading/on does not name an exact pattern or color, and
        # CharacterShading/off may be represented by several COM sentinel
        # values. Compare the requested on/off state below instead of the
        # reference's incidental raw values.
        _rendered_keys = _rendered_keys | {"shading", "shading_color"}

    observation_error = validate_formatting_observations(
        gold_fmt, pred_fmt, spec, ignored=_rendered_keys)
    if observation_error is not None:
        return observation_error

    for prop in gold_fmt.keys():
        if prop in _rendered_keys:
            continue
        gold_val = gold_fmt.get(prop)
        pred_val = pred_fmt.get(prop)
        comparison = (
            families_formatting_compare.composition_target_comparison(
                prop, spec, extractor="text_formatting"
            )
            if spec.get("composition")
            else None
        )
        if not families_formatting_compare.reference_comparison_matches(
            gold_val, pred_val, comparison
        ):
            if (
                prop == "text_outline_dash_style"
                and families_formatting_compare.text_outline_dash_styles_match(
                    gold_val, pred_val, spec
                )
            ):
                continue
            # For alignment: if the difference is left(0) vs right(2),
            # it may be a bidi encoding difference. Use rendered character
            # positions as the ground truth instead of the abstract property.
            # See the alignment bidi false-negative analysis for details.
            if (
                prop == "alignment"
                and families_formatting_compare.is_alignment_visually_equivalent(gold_fmt, pred_fmt)
            ):
                continue
            differences.append(
                {"property": prop, "gold": gold_val, "pred": pred_val, "location": "target_text"}
            )

    # For outline gallery commands where weight was skipped: still fail if
    # pred has zero weight (outline is visually invisible).
    # Note: COM API sometimes misreports zero-weight outlines, so this
    # is best-effort.
    if idMso in _OUTLINE_GALLERY_COMMANDS and not differences:
        pred_weight = pred_fmt.get("text_outline_weight")
        gold_weight = gold_fmt.get("text_outline_weight")
        if gold_weight and gold_weight > 0 and (pred_weight is None or pred_weight == 0):
            differences.append(
                {
                    "property": "text_outline_weight",
                    "gold": gold_weight,
                    "pred": pred_weight,
                    "location": "target_text (zero-weight = invisible outline)",
                }
            )

    pred_has_character_shading = families_formatting_compare.character_shading_present(pred_fmt)
    if "on" in character_shading_states and not pred_has_character_shading:
        differences.append(
            {
                "property": "character_shading_present",
                "gold": True,
                "pred": False,
                "location": "target_text",
            }
        )
    if "off" in character_shading_states and pred_has_character_shading:
        differences.append(
            {
                "property": "character_shading_present",
                "gold": False,
                "pred": True,
                "location": "target_text",
            }
        )

    if differences:
        details = {
            "differences": differences,
            "target_text": target_text,
            "gold": gold_fmt,
            "pred": pred_fmt,
        }
        families_formatting_composition.add_composition_atomic_details(
            details,
            pred_path,
            gold_path,
            spec,
            target_text,
            read_all_paragraphs_formatting=read_all_paragraphs_formatting,
            read_document_text=read_document_text,
            read_dropcap_properties=read_dropcap_properties,
            read_text_formatting=read_text_formatting,
        )
        return VerificationResult(
            passed=False,
            message=f"FAIL: {len(differences)} formatting difference(s) on target text",
            details=details,
        )

    composition_target_result = families_formatting_composition.verify_composition_atomic_targets(
        pred_path, gold_path, spec, read_all_paragraphs_formatting=read_all_paragraphs_formatting
    )
    if composition_target_result is not None:
        families_formatting_composition.add_composition_atomic_details(
            composition_target_result.details,
            pred_path,
            gold_path,
            spec,
            target_text,
            read_all_paragraphs_formatting=read_all_paragraphs_formatting,
            read_document_text=read_document_text,
            read_dropcap_properties=read_dropcap_properties,
            read_text_formatting=read_text_formatting,
        )
        return composition_target_result

    visibility_result = families_formatting_visibility.verify_target_glyph_visibility(
        pred_path, gold_path, spec, target_text, read_target_glyph_state=read_target_glyph_state
    )
    if visibility_result is not None and not visibility_result.passed:
        families_formatting_composition.add_composition_atomic_details(
            visibility_result.details,
            pred_path,
            gold_path,
            spec,
            target_text,
            read_all_paragraphs_formatting=read_all_paragraphs_formatting,
            read_document_text=read_document_text,
            read_dropcap_properties=read_dropcap_properties,
            read_text_formatting=read_text_formatting,
        )
        return visibility_result

    # Check non-target text (control text) wasn't modified
    control_texts = ["Alpha", "Omega"]
    control_result = families_formatting_preservation.check_control_text(
        pred_path, gold_path, control_texts, read_text_formatting=read_text_formatting
    )
    if control_result is not None:
        families_formatting_composition.add_composition_atomic_details(
            control_result.details,
            pred_path,
            gold_path,
            spec,
            target_text,
            read_all_paragraphs_formatting=read_all_paragraphs_formatting,
            read_document_text=read_document_text,
            read_dropcap_properties=read_dropcap_properties,
            read_text_formatting=read_text_formatting,
        )
        return control_result

    control_visibility_result = families_formatting_visibility.verify_control_glyph_visibility(
        pred_path, gold_path, spec, control_texts, read_target_glyph_state=read_target_glyph_state
    )
    if control_visibility_result is not None and not control_visibility_result.passed:
        families_formatting_composition.add_composition_atomic_details(
            control_visibility_result.details,
            pred_path,
            gold_path,
            spec,
            target_text,
            read_all_paragraphs_formatting=read_all_paragraphs_formatting,
            read_document_text=read_document_text,
            read_dropcap_properties=read_dropcap_properties,
            read_text_formatting=read_text_formatting,
        )
        return control_visibility_result

    if families_formatting_compare.composition_preserves_character_formatting(spec):
        preservation_result = (
            families_formatting_preservation.check_character_formatting_preservation(
                pred_path,
                gold_path,
                spec,
                target_text,
                read_all_paragraph_character_formatting=read_all_paragraph_character_formatting,
                read_all_paragraphs_formatting=read_all_paragraphs_formatting,
                read_document_text=read_document_text,
                read_dropcap_properties=read_dropcap_properties,
                read_text_formatting=read_text_formatting,
            )
        )
        if preservation_result is not None:
            return preservation_result

    details = {
        "target_text": target_text,
        "values": gold_fmt,
        "control_texts_checked": control_texts,
    }
    if visibility_result is not None:
        details["target_visibility"] = visibility_result.details
    if control_visibility_result is not None:
        details["control_visibility"] = control_visibility_result.details
    families_formatting_composition.add_composition_atomic_details(
        details,
        pred_path,
        gold_path,
        spec,
        target_text,
        read_all_paragraphs_formatting=read_all_paragraphs_formatting,
        read_document_text=read_document_text,
        read_dropcap_properties=read_dropcap_properties,
        read_text_formatting=read_text_formatting,
    )
    return VerificationResult(
        passed=True,
        message=f'PASS: Formatting matches for "{target_text}" and non-target text unchanged',
        details=details,
    )


def verify_ligatures_com(
    pred_path: Path, gold_path: Path, *, read_all_paragraph_font_properties, read_document_text
) -> VerificationResult:
    """Compare paragraph-wide OpenType ligature state through Word COM."""
    content_result = families_formatting_preservation.check_content_integrity(
        pred_path, gold_path, read_document_text=read_document_text
    )
    if content_result is not None:
        return content_result

    gold_fonts = read_all_paragraph_font_properties(str(gold_path))
    pred_fonts = read_all_paragraph_font_properties(str(pred_path))
    if not gold_fonts:
        return VerificationResult.error_result("Gold ligature projection has no paragraphs", observation_owner="gold")
    for owner, paragraphs in (("gold", gold_fonts), ("pred", pred_fonts)):
        if paragraphs is None:
            return VerificationResult.error_result("Ligature projection is unavailable", observation_owner=owner)
        for index, paragraph in enumerate(paragraphs, 1):
            error = require_observations({owner: paragraph}, ["ligatures"],
                                         context=f"paragraph {index} ligatures")
            if error is not None:
                return error
    gold_values = [row.get("ligatures") for row in gold_fonts]
    pred_values = [row.get("ligatures") for row in pred_fonts]
    if gold_values != pred_values:
        return VerificationResult(
            passed=False,
            message="FAIL: Word Object Model ligature settings differ",
            details={
                "differences": [
                    {
                        "property": "ligatures",
                        "gold": gold_values,
                        "pred": pred_values,
                    }
                ],
                "gold": gold_fonts,
                "pred": pred_fonts,
            },
        )
    return VerificationResult(
        passed=True,
        message="PASS: Word Object Model ligature settings match",
        details={
            "paragraph_count": len(gold_fonts),
            "ligatures": gold_values,
        },
    )


def verify_quick_styles(
    pred_path: Path,
    gold_path: Path,
    spec: Dict,
    *,
    read_all_paragraph_font_properties,
    read_body_paragraph_styles,
) -> VerificationResult:
    """Use Word styles as primary and saved direct formatting as guardrail."""
    return policies_content_formatting.com_primary_with_saved_guardrail(
        lambda: verify_quick_styles_com(
            pred_path,
            gold_path,
            read_all_paragraph_font_properties=read_all_paragraph_font_properties,
            read_body_paragraph_styles=read_body_paragraph_styles,
        ),
        lambda: families_formatting_verify_saved.verify_quick_styles_saved(
            pred_path, gold_path, spec
        ),
    )


def verify_quick_styles_com(
    pred_path: Path,
    gold_path: Path,
    *,
    read_all_paragraph_font_properties,
    read_body_paragraph_styles,
) -> VerificationResult:
    gold_styles = read_body_paragraph_styles(str(gold_path))
    pred_styles = read_body_paragraph_styles(str(pred_path))
    for owner, styles in (("gold", gold_styles), ("pred", pred_styles)):
        if styles is None or (owner == "gold" and not styles):
            return VerificationResult.error_result("Paragraph style projection is unavailable", observation_owner=owner)
        if any(not isinstance(style, str) or not style for style in styles):
            return VerificationResult.error_result("Required paragraph style name is unavailable", observation_owner=owner)
    gold_paragraphs = read_all_paragraph_font_properties(str(gold_path))
    pred_paragraphs = read_all_paragraph_font_properties(str(pred_path))
    gold_texts = [row.get("text") for row in gold_paragraphs]
    pred_texts = [row.get("text") for row in pred_paragraphs]
    differences = []
    if gold_styles != pred_styles:
        differences.append(
            {
                "property": "paragraph_styles",
                "gold": gold_styles,
                "pred": pred_styles,
            }
        )
    if gold_texts != pred_texts:
        differences.append(
            {
                "property": "paragraph_texts",
                "gold": gold_texts,
                "pred": pred_texts,
            }
        )
    return VerificationResult(
        passed=not differences,
        message=(
            "PASS: Word Object Model paragraph styles match"
            if not differences
            else "FAIL: Word Object Model paragraph styles differ"
        ),
        details={
            "differences": differences,
            "styles": gold_styles,
            "paragraph_count": len(gold_styles),
        },
    )


def verify_formatting_match(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, read_text_formatting
) -> VerificationResult:
    """
    Compare every formatting property on a literal target text run.

    Caller contract (from `verify()`):
    - `select_text` is a literal text or the default 'Target'. Special
      values 'all' and 'paragraph_N' are routed by `verify()` to the
      paragraph-aware handlers and never reach this method.
    - `spec` may be None (the no-spec early return in `verify()` forwards
      here for a baseline 'Target'-anchored comparison).
    """
    # Get target text
    target_text = "Target"
    if spec:
        precondition = spec.get("precondition", {})
        target_text = precondition.get("select_text", "Target")

    # Extract and compare (standard text search)
    gold_fmt = read_text_formatting(str(gold_path), target_text)
    pred_fmt = read_text_formatting(str(pred_path), target_text)

    if gold_fmt is None:
        return VerificationResult(
            passed=False,
            message=f'ERROR: Target text "{target_text}" not found in gold',
            error=f'Target text "{target_text}" not found in gold',
            details={},
        )

    if pred_fmt is None:
        return VerificationResult(
            passed=False, message=f'FAIL: Target text "{target_text}" not found in pred', details={}
        )

    observation_error = validate_formatting_observations(gold_fmt, pred_fmt, spec or {})
    if observation_error is not None:
        return observation_error

    # Compare all formatting properties for target text
    differences = []
    _rendered_keys = {"rendered_first_h", "rendered_last_h"}
    for prop in gold_fmt.keys():
        if prop in _rendered_keys:
            continue
        gold_val = gold_fmt.get(prop)
        pred_val = pred_fmt.get(prop)
        if gold_val != pred_val:
            if (
                prop == "alignment"
                and families_formatting_compare.is_alignment_visually_equivalent(gold_fmt, pred_fmt)
            ):
                continue
            differences.append(
                {"property": prop, "gold": gold_val, "pred": pred_val, "location": "target_text"}
            )

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: {len(differences)} formatting difference(s) on target text",
            details={
                "differences": differences,
                "target_text": target_text,
                "gold": gold_fmt,
                "pred": pred_fmt,
            },
        )

    # Check non-target text (control text) wasn't modified
    control_texts = ["Alpha", "Omega"]
    control_result = families_formatting_preservation.check_control_text(
        pred_path, gold_path, control_texts, read_text_formatting=read_text_formatting
    )
    if control_result is not None:
        return control_result

    return VerificationResult(
        passed=True,
        message=f'PASS: Formatting matches for "{target_text}" and non-target text unchanged',
        details={
            "target_text": target_text,
            "values": gold_fmt,
            "control_texts_checked": control_texts,
        },
    )
