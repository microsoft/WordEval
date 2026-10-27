"""composition_visual with explicit evidence dependencies."""

from evaluation.methods.word_object.effect_commands import (
    FILL_OUTLINE_COLOR_COMMANDS,
    PICTURE_CROP_COMMANDS,
    PICTURE_EFFECT_COMMANDS,
    PICTURE_FORMAT_COMMANDS,
    SHAPE_EFFECT_COMMANDS,
    SHAPE_GEOMETRY_COMMANDS,
    SMARTART_SUBNODE_COMMANDS,
    TEXT_ALIGN_COMMANDS,
    TEXT_EFFECT_COMMANDS,
)
from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object import composition_effects as composition_effects
from pathlib import Path
from typing import Dict, Optional


def verify_shape_composition(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    composition_step_declaration_error,
    composition_step_key,
    unhandled_composition_step,
    verify_fill_outline_color,
    verify_insert_shape_operation,
    verify_shape_effects,
    verify_shape_geometry,
    verify_text_effects
) -> VerificationResult:
    sub_results = {}
    declaration_error = composition_step_declaration_error("shape", spec)
    if declaration_error is not None:
        return declaration_error
    steps = composition_effects.composition_steps(spec)
    if not steps:
        return verify_shape_effects(pred_path, gold_path, spec)

    seen_steps: set[str] = set()
    for step_index, step in enumerate(steps, start=1):
        step_key, step_error = composition_step_key("shape", step_index, step, seen_steps)
        if step_error is not None:
            return step_error
        op = step.get("operation")
        idMso = step.get("idMso")
        if op == "insert_shape":
            result = verify_insert_shape_operation(pred_path, gold_path, step.get("params"), spec)
        elif idMso in FILL_OUTLINE_COLOR_COMMANDS:
            result = verify_fill_outline_color(
                pred_path, gold_path, composition_effects.scoped_spec(spec, idMso)
            )
        elif idMso in SHAPE_EFFECT_COMMANDS:
            result = verify_shape_effects(
                pred_path, gold_path, composition_effects.scoped_spec(spec, idMso)
            )
        elif idMso in SHAPE_GEOMETRY_COMMANDS or idMso in SMARTART_SUBNODE_COMMANDS:
            result = verify_shape_geometry(
                pred_path, gold_path, composition_effects.scoped_spec(spec, idMso)
            )
        elif idMso in TEXT_EFFECT_COMMANDS:
            result = verify_text_effects(
                pred_path, gold_path, composition_effects.scoped_spec(spec, idMso)
            )
        else:
            return unhandled_composition_step("shape", step_index, step)
        if not result.passed:
            return result
        sub_results[step_key] = result.details

    return composition_effects.merge_sub_results(
        "PASS: Shape composition properties match", sub_results
    )


def verify_picture_composition(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    composition_step_declaration_error,
    composition_step_key,
    unhandled_composition_step,
    verify_insert_icon_operation,
    verify_picture_crop,
    verify_picture_effects,
    verify_picture_format
) -> VerificationResult:
    sub_results = {}
    declaration_error = composition_step_declaration_error("picture", spec)
    if declaration_error is not None:
        return declaration_error
    steps = composition_effects.composition_steps(spec)
    if not steps:
        last_idMso = _get_last_step_idMso(spec or {})
        if last_idMso in PICTURE_CROP_COMMANDS:
            return verify_picture_crop(pred_path, gold_path, spec)
        return verify_picture_effects(pred_path, gold_path, spec)

    inserted_icon = any(step.get("operation") == "insert_icon" for step in steps)
    seen_steps: set[str] = set()
    for step_index, step in enumerate(steps, start=1):
        step_key, step_error = composition_step_key("picture", step_index, step, seen_steps)
        if step_error is not None:
            return step_error
        op = step.get("operation")
        idMso = step.get("idMso")
        if op == "insert_icon":
            result = verify_insert_icon_operation(pred_path, gold_path)
        elif idMso in PICTURE_CROP_COMMANDS:
            result = verify_picture_crop(
                pred_path, gold_path, composition_effects.scoped_spec(spec, idMso)
            )
        elif idMso in PICTURE_EFFECT_COMMANDS:
            scoped_spec = composition_effects.scoped_spec(spec, idMso)
            if inserted_icon:
                scoped_spec["_allow_floating_picture"] = True
            result = verify_picture_effects(pred_path, gold_path, scoped_spec)
        elif idMso in PICTURE_FORMAT_COMMANDS:
            result = verify_picture_format(
                pred_path, gold_path, composition_effects.scoped_spec(spec, idMso)
            )
        else:
            return unhandled_composition_step("picture", step_index, step)
        if not result.passed:
            return result
        sub_results[step_key] = result.details

    return composition_effects.merge_sub_results(
        "PASS: Picture composition properties match", sub_results
    )


def verify_textbox_composition(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    composition_step_declaration_error,
    composition_step_key,
    unhandled_composition_step,
    verify_fill_outline_color,
    verify_text_alignment,
    verify_textbox_insert,
    verify_textbox_text_fill
) -> VerificationResult:
    sub_results = {}
    declaration_error = composition_step_declaration_error("textbox", spec)
    if declaration_error is not None:
        return declaration_error
    steps = composition_effects.composition_steps(spec)
    if not steps:
        return verify_text_alignment(pred_path, gold_path, spec)

    seen_steps: set[str] = set()
    for step_index, step in enumerate(steps, start=1):
        step_key, step_error = composition_step_key("textbox", step_index, step, seen_steps)
        if step_error is not None:
            return step_error
        op = step.get("operation")
        idMso = step.get("idMso")
        if op == "insert_textbox":
            result = verify_textbox_insert(pred_path, gold_path)
        elif idMso in TEXT_ALIGN_COMMANDS:
            result = verify_text_alignment(
                pred_path, gold_path, composition_effects.scoped_spec(spec, idMso)
            )
        elif idMso == "TextFillColorPicker":
            result = verify_textbox_text_fill(pred_path, gold_path)
        elif idMso in FILL_OUTLINE_COLOR_COMMANDS:
            result = verify_fill_outline_color(
                pred_path, gold_path, composition_effects.scoped_spec(spec, idMso)
            )
        else:
            return unhandled_composition_step("textbox", step_index, step)
        if not result.passed:
            return result
        sub_results[step_key] = result.details

    return composition_effects.merge_sub_results(
        "PASS: Textbox composition properties match", sub_results
    )


def _get_last_step_idMso(spec):
    """Extract the idMso from the last step in a composition spec."""
    steps = spec.get("composition", {}).get("steps", [])
    if steps:
        return steps[-1].get("idMso", "")
    return ""
