"""preservation / targets with explicit evidence dependencies."""

from evaluation.methods.word_object.effect_commands import (
    ALIGNMENT_COMMANDS,
    GROUP_COMMANDS,
    UNGROUP_COMMANDS,
    Z_ORDER_COMMANDS,
    ROTATION_COMMANDS,
    POSITION_COMMANDS,
    WRAP_COMMANDS,
    PICTURE_CROP_COMMANDS,
    PICTURE_FORMAT_COMMANDS,
    CHART_COMMANDS_SET,
    DISTRIBUTION_COMMANDS,
    EQUATION_COMMANDS_SET,
    FILL_OUTLINE_COLOR_COMMANDS,
    PICTURE_EFFECT_COMMANDS,
    PICTURE_NS,
    SHAPE_EFFECT_COMMANDS,
    SHAPE_GEOMETRY_COMMANDS,
    SMARTART_COMMANDS_SET,
    SMARTART_STRUCTURAL_COMMANDS,
    SMARTART_SUBNODE_COMMANDS,
    SMARTART_TEXT_EFFECT_COMMANDS,
    TABLE_COMMANDS,
    TEXT_ALIGN_COMMANDS,
    WORDPROCESSING_GROUP_NS,
    WORDPROCESSING_SHAPE_NS,
)
from evaluation.methods.word_object import composition_effects as composition_effects
from pathlib import Path
from typing import Any, Dict, Optional
import xml.etree.ElementTree as ET
import zipfile


def visual_object_guardrail_kind(spec: Optional[Dict]) -> Optional[str]:
    """Return the Session 2 semantic target family, if applicable."""
    if not spec:
        return None
    if spec.get("composition"):
        target = spec.get("composition", {}).get("target_object_type", "")
        if not target:
            target = spec.get("precondition", {}).get("type", "")
        if target in {"shape", "multi_shape", "graphic", "textbox"}:
            return "shape"
        if target == "picture":
            return "picture"
        if target in {
            "smartart",
            "smartart_list",
            "smartart_hierarchy",
            "smartart_orgchart",
        }:
            return "smartart"
        return None

    idMso = spec.get("idMso", "")
    precondition = spec.get("precondition", {}).get("type", "")
    if idMso in (GROUP_COMMANDS | UNGROUP_COMMANDS | Z_ORDER_COMMANDS |
                 ROTATION_COMMANDS | POSITION_COMMANDS | WRAP_COMMANDS):
        return "picture" if precondition == "picture" else "shape"
    if idMso in PICTURE_CROP_COMMANDS | PICTURE_FORMAT_COMMANDS:
        return "picture"
    if idMso in ALIGNMENT_COMMANDS or idMso in DISTRIBUTION_COMMANDS:
        return "shape"
    if idMso in TEXT_ALIGN_COMMANDS:
        return "shape"
    if idMso in SHAPE_EFFECT_COMMANDS or idMso in SHAPE_GEOMETRY_COMMANDS:
        return "shape"
    if idMso in FILL_OUTLINE_COLOR_COMMANDS:
        if idMso.startswith("Graphics") or precondition in {
            "picture",
            "icon",
            "graphic",
        }:
            return "picture"
        return "shape"
    if idMso in PICTURE_EFFECT_COMMANDS:
        return "picture"
    if idMso in SMARTART_COMMANDS_SET or idMso in SMARTART_STRUCTURAL_COMMANDS:
        return "smartart"
    if idMso in SMARTART_SUBNODE_COMMANDS and precondition == "smartart":
        return "smartart"
    if idMso in SMARTART_TEXT_EFFECT_COMMANDS and precondition == "smartart":
        return "smartart"
    return None


def saved_visual_target_present(docx_path: Path, kind: str) -> bool:
    """Return whether a saved package contains the required target kind."""
    with zipfile.ZipFile(docx_path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
        names = set(package.namelist())
    if kind == "shape":
        tags = {
            f"{{{WORDPROCESSING_SHAPE_NS}}}wsp",
            f"{{{WORDPROCESSING_GROUP_NS}}}wgp",
            "{urn:schemas-microsoft-com:vml}shape",
            "{urn:schemas-microsoft-com:vml}group",
        }
        return any(node.tag in tags for node in root.iter())
    if kind == "picture":
        picture_tag = f"{{{PICTURE_NS}}}pic"
        return any(node.tag == picture_tag for node in root.iter())
    if kind == "smartart":
        diagram_rel = "{http://schemas.openxmlformats.org/drawingml/2006/diagram}relIds"
        return any(node.tag == diagram_rel for node in root.iter()) and any(
            name.startswith("word/diagrams/data") and name.endswith(".xml") for name in names
        )
    raise ValueError(f"Unknown visual-object guardrail kind: {kind}")


def structured_object_guardrail_kind(spec: Optional[Dict]) -> Optional[str]:
    """Return the Session 3 structured target family, if applicable."""
    if not spec:
        return None
    if spec.get("composition"):
        target = spec.get("composition", {}).get("target_object_type", "")
        if not target:
            target = spec.get("precondition", {}).get("type", "")
        return target if target in {"chart", "equation", "table"} else None
    idMso = spec.get("idMso", "")
    if idMso in CHART_COMMANDS_SET:
        return "chart"
    if idMso in EQUATION_COMMANDS_SET:
        return "equation"
    if idMso in TABLE_COMMANDS:
        # SmartArt sub-node commands (e.g. OutlineWeightGallery) are
        # dual-listed in TABLE_COMMANDS.  In a SmartArt context the dispatch
        # in _verify_requested_semantics routes them to SmartArt
        # verification, so the structured table preflight must not demand a
        # table target in gold.  Mirror that dispatch precedence here.
        if (
            idMso in SMARTART_SUBNODE_COMMANDS
            and spec.get("precondition", {}).get("type", "") == "smartart"
        ):
            return None
        return "table"
    return None


def smartart_connector_fallback_contract_is_complete(
    spec: Optional[Dict], details: Dict[str, Any]
) -> bool:
    """Authorize the narrowly scoped SmartArt connector saved fallback.

    A SmartArt composition is complete only when every declared atomic step
    produced a distinct sub-result, the sole saved-only evidence is the
    connector-arrow property, and the shared visible-preservation guardrail
    passed.  Any unknown/skipped/duplicate step or any broader saved-state
    dependency remains failure-only under the evaluator's default gate.
    """
    composition = (spec or {}).get("composition") or {}
    target_type = composition.get("target_object_type") or (
        (spec or {}).get("precondition", {}).get("type")
    )
    if target_type not in {
        "smartart",
        "smartart_list",
        "smartart_hierarchy",
        "smartart_orgchart",
    }:
        return False
    if details.get("verification_source") != ("saved_ooxml_only_wom_unavailable"):
        return False

    preservation = details.get("visible_preservation") or {}
    if preservation.get("match_mode") != "saved_ooxml_visual_object_preservation" or not (
        preservation.get("projection_comparison") or {}
    ).get("equivalent"):
        return False

    steps = composition_effects.composition_steps(spec)
    step_keys = [step.get("operation") or step.get("idMso") for step in steps]
    if not step_keys or any(not key for key in step_keys) or len(step_keys) != len(set(step_keys)):
        return False
    declared_count = composition.get("operation_count")
    if declared_count is not None and declared_count != len(step_keys):
        return False

    sub_results = details.get("sub_results")
    if not isinstance(sub_results, dict) or set(sub_results) != set(step_keys):
        return False
    saved_backed = [
        sub_result
        for sub_result in sub_results.values()
        if isinstance(sub_result, dict)
        and str(sub_result.get("verification_source", "")).startswith("saved_ooxml")
    ]
    return bool(saved_backed) and all(
        sub_result.get("verification_source") == "saved_ooxml_only_wom_unavailable"
        and sub_result.get("idMso") == "ArrowStyleGallery"
        and sub_result.get("property") == "connector_arrows"
        for sub_result in saved_backed
    )
