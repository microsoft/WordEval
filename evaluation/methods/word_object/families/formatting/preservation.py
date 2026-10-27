"""families / formatting / preservation with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.families.formatting import (
    compare as families_formatting_compare,
)
from evaluation.methods.word_object.families.formatting import (
    composition as families_formatting_composition,
)
from pathlib import Path
from typing import Dict, List, Optional
from .contract import validate_formatting_observations
from .characters import validate_character_observations, visible_character_stream


def check_dedicated_character_preservation(
    pred_path, gold_path, *, read_characters, doc_in_path=None, drop_cap=False,
):
    """Compare existing character bundles, allowing only the drop-cap initial size.

    Word splits the dropped initial into a separate paragraph. Compare aligned
    visible characters rather than paragraph boundaries or saved run boundaries.
    The gold is also checked against input so collateral gold edits are not rules.
    """
    paths = {"gold": gold_path, "pred": pred_path}
    if doc_in_path is not None:
        paths = {"input": doc_in_path, **paths}
    states = {owner: read_characters(str(path)) for owner, path in paths.items()}
    error = validate_character_observations(states)
    if error is not None:
        return error

    def projection(paragraphs):
        characters = []
        for run in visible_character_stream(paragraphs):
            for char in run["text"]:
                properties = dict(run["format"])
                if drop_cap and not characters:
                    properties.pop("font_size", None)
                characters.append((char, properties))
        return characters

    baseline_owner = "input" if doc_in_path is not None else "gold"
    baseline = projection(states[baseline_owner])
    for owner in ("gold", "pred"):
        current = projection(states[owner])
        if current == baseline:
            continue
        details = {"preservation_scope": "non_target_character_formatting",
                   "baseline": baseline_owner, "observation_owner": owner,
                   "allowed_effect": "initial_font_size" if drop_cap else "ligatures_only"}
        if owner == "gold":
            return VerificationResult.error_result("Gold changed unrelated character formatting", **details)
        return VerificationResult(passed=False, message="FAIL: Unrelated character formatting changed", details=details)
    return None


def check_character_formatting_preservation(
    pred_path: Path,
    gold_path: Path,
    spec: Dict,
    select_text: str,
    *,
    read_all_paragraph_character_formatting,
    read_all_paragraphs_formatting,
    read_document_text,
    read_dropcap_properties,
    read_text_formatting,
) -> Optional[VerificationResult]:
    """Reject collateral visible character changes in paragraph commands."""
    gold_characters = read_all_paragraph_character_formatting(str(gold_path))
    pred_characters = read_all_paragraph_character_formatting(str(pred_path))
    operations = {
        intent.get("operation") for intent in spec.get("composition", {}).get("atomic_intents", [])
    }
    error = validate_character_observations(
        {"gold": gold_characters, "pred": pred_characters},
        ignored={"font_size"} if "drop_cap" in operations else (),
    )
    if error is not None:
        return error
    if "drop_cap" in operations:
        gold_characters = [
            {key: value for key, value in item.items() if key != "font_size"}
            for item in gold_characters
        ]
        pred_characters = [
            {key: value for key, value in item.items() if key != "font_size"}
            for item in pred_characters
        ]
    if gold_characters == pred_characters:
        return None

    character_differences = []
    for index, (gold_char, pred_char) in enumerate(zip(gold_characters, pred_characters), 1):
        if gold_char == pred_char:
            continue
        character_differences.append(
            {
                "location": f"paragraph {index}",
                "gold": gold_char,
                "pred": pred_char,
            }
        )
    if len(gold_characters) != len(pred_characters):
        character_differences.append(
            {
                "property": "paragraph_count",
                "gold": len(gold_characters),
                "pred": len(pred_characters),
            }
        )
    details = {
        "differences": character_differences[:10],
        "total_differences": len(character_differences),
        "select_text": select_text,
        "preservation_scope": "character_formatting",
    }
    families_formatting_composition.add_composition_atomic_details(
        details,
        pred_path,
        gold_path,
        spec,
        select_text,
        read_all_paragraphs_formatting=read_all_paragraphs_formatting,
        read_document_text=read_document_text,
        read_dropcap_properties=read_dropcap_properties,
        read_text_formatting=read_text_formatting,
    )
    return VerificationResult(
        passed=False,
        message=("FAIL: Visible character formatting changed during a " "paragraph/list command"),
        details=details,
    )


def check_content_integrity(
    pred_path: Path, gold_path: Path, *, read_document_text
) -> Optional[VerificationResult]:
    """
    Check that document text content matches.

    Returns None if content matches, or a failure result if content differs.
    """
    pred_text = read_document_text(str(pred_path))
    gold_text = read_document_text(str(gold_path))

    if pred_text != gold_text:
        # Identify what changed
        added = []
        removed = []

        for phrase in ["Alpha", "Omega", "Target"]:
            pred_count = pred_text.count(phrase)
            gold_count = gold_text.count(phrase)
            if pred_count > gold_count:
                added.append(f"'{phrase}' appears {pred_count} times (expected {gold_count})")
            elif pred_count < gold_count:
                removed.append(f"'{phrase}' appears {pred_count} times (expected {gold_count})")

        return VerificationResult(
            passed=False,
            message="FAIL: Document text content differs from gold",
            details={
                "pred_text": pred_text[:200] + ("..." if len(pred_text) > 200 else ""),
                "gold_text": gold_text[:200] + ("..." if len(gold_text) > 200 else ""),
                "pred_length": len(pred_text),
                "gold_length": len(gold_text),
                "added": added if added else None,
                "removed": removed if removed else None,
            },
        )

    return None  # Content matches


def check_control_text(
    pred_path: Path, gold_path: Path, control_texts: List[str] = None, *, read_text_formatting
) -> Optional[VerificationResult]:
    """
    Check that non-target text (control text) formatting is unchanged.

    Returns None if control text is unchanged, or a failure result if modified.
    """
    if control_texts is None:
        control_texts = ["Alpha", "Omega"]
    control_differences = []
    _rendered_keys = {"rendered_first_h", "rendered_last_h"}

    for control_text in control_texts:
        gold_fmt = read_text_formatting(str(gold_path), control_text)
        pred_fmt = read_text_formatting(str(pred_path), control_text)

        if gold_fmt is None:
            continue
        if pred_fmt is None:
            return VerificationResult(
                passed=False, message=f'FAIL: Control text "{control_text}" is missing',
                details={"control_text": control_text})
        observation_error = validate_formatting_observations(gold_fmt, pred_fmt, {})
        if observation_error is not None:
            return observation_error

        for prop in gold_fmt.keys():
            if prop in _rendered_keys:
                continue
            gold_val = gold_fmt.get(prop)
            pred_val = pred_fmt.get(prop)
            if gold_val != pred_val:
                # Apply same bidi-aware alignment check for control text
                if (
                    prop == "alignment"
                    and families_formatting_compare.is_alignment_visually_equivalent(
                        gold_fmt, pred_fmt
                    )
                ):
                    continue
                control_differences.append(
                    {
                        "property": prop,
                        "gold": gold_val,
                        "pred": pred_val,
                        "location": f'control_text "{control_text}"',
                    }
                )

    if control_differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: {len(control_differences)} formatting difference(s) on non-target text",
            details={
                "differences": control_differences,
                "explanation": "Formatting was applied to text outside the target.",
            },
        )

    return None  # Control text unchanged
