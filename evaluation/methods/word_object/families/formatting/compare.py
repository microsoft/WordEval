"""families / formatting / compare with explicit evidence dependencies."""

from evaluation.methods.word_object.composition import (
    composition_command_ids,
    composition_operation_ids,
)
from typing import Any, Dict, List, Optional


def normalize_dropcap_paragraphs(paragraphs: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Remove Word's synthetic one-letter drop-cap paragraph.

    A real dropped cap is exposed by Word COM as a separate first
    paragraph containing only the dropped character, followed by the rest
    of the logical first paragraph.  All-paragraph intents such as spacing
    should compare logical body paragraphs and must not fail solely because
    one side attempted a non-rendering OOXML lookalike instead.
    """
    if len(paragraphs) < 2:
        return paragraphs
    first = paragraphs[0]
    position = first.get("dropcap_position")
    text = str(first.get("text") or "").strip()
    if position not in {None, 0} and len(text) <= 2:
        return paragraphs[1:]
    return paragraphs


def character_shading_present(formatting: Dict[str, Any]) -> bool:
    texture = formatting.get("shading")
    background = formatting.get("shading_color")
    texture_present = texture not in {None, 0, 9999999, -9999999}
    background_present = background not in {
        None,
        -16777216,  # wdColorAutomatic
        9999999,
        -9999999,
    }
    return texture_present or background_present


def character_shading_state(atomic_case: Optional[str]) -> Optional[str]:
    """Return the requested CharacterShading state encoded by an intent."""
    if not atomic_case:
        return None
    if atomic_case.endswith("/on"):
        return "on"
    if atomic_case.endswith("/off"):
        return "off"
    return None


def composition_target_values_match(
    prop: str,
    gold_val: Any,
    pred_val: Any,
    extractor: str,
    gold_fmt: Dict[str, Any],
    pred_fmt: Dict[str, Any],
    spec: Dict,
    comparison: Optional[str] = None,
) -> bool:
    comparison = comparison or composition_indent_comparison(prop, spec)
    if reference_comparison_matches(gold_val, pred_val, comparison):
        return True
    if prop == "text_outline_dash_style":
        return text_outline_dash_styles_match(gold_val, pred_val, spec)
    if prop == "alignment" and extractor in ("text_formatting", "paragraph_formatting"):
        return is_alignment_visually_equivalent(gold_fmt, pred_fmt)
    return False


def text_outline_dash_styles_match(gold_val: Any, pred_val: Any, spec: Dict) -> bool:
    """Accept Word's two round-cap dotted text-outline encodings.

    `msoLineRoundDot` is 3 and persists as `dot`. Word also produces
    `msoLineSysDot` (11 / `sysDot`) with a round line cap for text-outline
    dotted results that are visually and semantically valid for the current
    round-dot instruction. This equivalence is scoped to the explicit
    Round_Dot atomic intent; square dots and every dash style remain exact.
    Visibility is verified independently.
    """
    if gold_val == pred_val:
        return True
    if not round_dot_text_outline_requested(spec):
        return False
    return gold_val in {3, 11} and pred_val in {3, 11}


def round_dot_text_outline_requested(spec: Dict) -> bool:
    parameter = spec.get("parameter", {}) or {}
    if spec.get("idMso") == "TextOutlineDashesGallery" and (
        parameter.get("label") == "Round_Dot" or parameter.get("style") == 3
        or spec.get("variation") == "apply_Round_Dot"
    ):
        return True
    return any(
        intent.get("atomic_case") == "TextOutlineDashesGallery/apply_Round_Dot"
        for intent in spec.get("composition", {}).get("atomic_intents", [])
    )


def reference_comparison_matches(gold_val: Any, pred_val: Any, comparison: Optional[str]) -> bool:
    """Compare directional commands without requiring one exact Word step size."""
    if gold_val == pred_val:
        return True
    if comparison not in {"at_least_reference", "at_most_reference"}:
        return False
    try:
        gold_number = float(gold_val)
        pred_number = float(pred_val)
    except (TypeError, ValueError):
        return False
    if comparison == "at_least_reference":
        return pred_number >= gold_number
    return pred_number <= gold_number


def composition_indent_comparison(prop: str, spec: Dict) -> Optional[str]:
    """Infer comparison mode for historical specs without target metadata."""
    if prop != "left_indent":
        return None
    commands = set(composition_command_ids(spec))
    increases = {"IndentIncreaseWord", "ListIndentIncrease"} & commands
    decreases = {"IndentDecreaseWord", "ListIndentDecrease"} & commands
    if increases and not decreases:
        return "at_least_reference"
    if decreases and not increases:
        return "at_most_reference"
    return None


def composition_target_comparison(
    prop: str, spec: Dict, extractor: Optional[str] = None
) -> Optional[str]:
    """Return a case-declared comparison before historical inference.

    The exact/directional distinction belongs to the instruction, not only
    to the relative Word command used to generate the reference. Explicit
    ``comparison: exact`` therefore prevents an instruction such as
    "increase to 1.27 cm" from accepting a larger value.
    """
    for intent in spec.get("composition", {}).get("atomic_intents", []):
        for target in intent.get("verification_targets", []):
            if target.get("property") != prop:
                continue
            if extractor and target.get("extractor") != extractor:
                continue
            if "comparison" in target:
                return target.get("comparison")
    return composition_indent_comparison(prop, spec)


def composition_preserves_character_formatting(spec: Dict) -> bool:
    """Whether a composition declares only paragraph/list operations."""
    intents = spec.get("composition", {}).get("atomic_intents", [])
    operations = {intent.get("operation") for intent in intents if intent.get("operation")}
    paragraph_operations = {
        "alignment",
        "indent",
        "line_spacing",
        "para_border",
        "spacing_before",
        "spacing_after",
        "drop_cap",
        "list_type",
        "list_spacing",
        "list_alignment",
        "list_indent",
    }
    return bool(operations) and operations <= paragraph_operations


def composition_uses_bullet_normalization(spec: Dict) -> bool:
    idMso = spec.get("idMso", "")
    composition_cmds = composition_command_ids(spec)
    composition_ops = composition_operation_ids(spec)
    return (
        idMso == "BulletsGalleryWord"
        or "BulletsGalleryWord" in composition_cmds
        or "apply_bullets" in composition_ops
    )


def is_alignment_visually_equivalent(gold_fmt: Dict, pred_fmt: Dict) -> bool:
    """
    Check if two alignment values are visually equivalent despite different
    Paragraph.Alignment property values.

    This handles the bidi encoding case where a paragraph with bidi + right
    alignment (Alignment=2) renders identically to a non-bidi left-aligned
    paragraph (Alignment=0). The COM Alignment property reflects the abstract
    property, not the rendered result.

    We compare the rendered first-character and last-character horizontal
    positions (measured via Selection.Information(wdHorizontalPositionRelativeToPage)).
    If both match within tolerance, the alignment is visually equivalent.

    Only applies to left(0) vs right(2) discrepancies (the bidi flip case).
    Other discrepancies (e.g., left vs justify, left vs center) are not
    affected by bidi and should use direct property comparison.
    """
    ALIGNMENT_LEFT = 0
    ALIGNMENT_RIGHT = 2
    TOLERANCE_PT = 3.0  # points

    gold_align = gold_fmt.get("alignment")
    pred_align = pred_fmt.get("alignment")

    # Only handle left <-> right discrepancies (bidi flip)
    if {gold_align, pred_align} != {ALIGNMENT_LEFT, ALIGNMENT_RIGHT}:
        return False

    # Need rendered positions from both
    gold_first = gold_fmt.get("rendered_first_h")
    gold_last = gold_fmt.get("rendered_last_h")
    pred_first = pred_fmt.get("rendered_first_h")
    pred_last = pred_fmt.get("rendered_last_h")

    if any(v is None for v in [gold_first, gold_last, pred_first, pred_last]):
        return False  # Can't determine — fall back to property comparison

    # Compare both left edge (first char) and right edge (last char)
    first_match = abs(gold_first - pred_first) <= TOLERANCE_PT
    last_match = abs(gold_last - pred_last) <= TOLERANCE_PT

    return first_match and last_match
