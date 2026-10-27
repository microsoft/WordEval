"""families / smartart / verify with explicit evidence dependencies."""

from evaluation.methods.word_object.effect_commands import SMARTART_SUBNODE_COMMANDS
from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.composition import composition_command_ids
from evaluation.methods.word_object.families.smartart import (
    extract_saved as families_smartart_extract_saved,
)
from evaluation.methods.word_object.families.smartart import refresh as families_smartart_refresh
from .gallery import verify_gallery_definitions
from pathlib import Path
from typing import Dict, Optional
from ...observations import require_observations


def verify_smartart(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict] = None,
    *,
    extract_refreshed_smartart_projection,
    get_cleanup_word,
    read_smartart_drawing_info,
    read_smartart_properties,
    read_smartart_subnode_properties,
) -> VerificationResult:
    """
    Verify SmartArt layout/style/color matches.

    Uses extract_smartart_properties() which checks both inline and floating shapes.
    """
    idMso = spec.get("idMso", "") if spec else ""
    connector_operations = {
        intent.get("operation")
        for intent in (spec or {}).get("composition", {}).get("atomic_intents", [])
    }
    # Only the arrow operation targets the connector lines. Outline weight
    # and dash follow their linked atomics onto the Target node shape, so
    # comparing prstGeom="line" shapes for them would read the wrong shape.
    connector_property = {
        "ArrowStyleGallery": ("connector_arrow", "connector_arrows"),
    }.get(idMso)
    if connector_property and connector_property[0] in connector_operations:
        operation, prop = connector_property
        gold_connectors = families_smartart_extract_saved.extract_saved_smartart_connector_styles(
            gold_path
        )
        pred_connectors = families_smartart_extract_saved.extract_saved_smartart_connector_styles(
            pred_path
        )
        if gold_connectors.get("connector_count", 0) and gold_connectors.get(
            prop
        ) != pred_connectors.get(prop):
            return VerificationResult(
                passed=False,
                message="FAIL: Saved SmartArt connector formatting differs",
                details={
                    "differences": [
                        {
                            "property": prop,
                            "gold": gold_connectors.get(prop),
                            "pred": pred_connectors.get(prop),
                        }
                    ],
                    "operation": operation,
                    "idMso": idMso,
                    "verification_source": "saved_ooxml_only_wom_unavailable",
                },
            )
        if gold_connectors.get("connector_count", 0):
            return VerificationResult(
                passed=True,
                message="PASS: Saved SmartArt connector formatting matches",
                details={
                    "property": prop,
                    "idMso": idMso,
                    "normalized_gold": gold_connectors.get(prop),
                    "normalized_pred": pred_connectors.get(prop),
                    "scope_note": (
                        "saved drawing-cache connectors in document order; "
                        "target binding is not established by this helper"
                    ),
                    "verification_source": "saved_ooxml_only_wom_unavailable",
                },
            )

    # Word Object Model is the primary semantic source for SmartArt data,
    # gallery identity, and node properties.  The saved drawing cache below
    # is a persistence/visible-state guardrail, not a replacement for COM.
    gold_sa = read_smartart_properties(str(gold_path))
    pred_sa = read_smartart_properties(str(pred_path))

    if not gold_sa.get("found"):
        result = VerificationResult.error_result('No SmartArt found in gold document')
        result.details.update(error_owner='gold_fixture_or_evaluator',gold=gold_sa)
        return result

    if not pred_sa.get("found"):
        return VerificationResult(
            passed=False,
            message="FAIL: No SmartArt found in pred document",
            details={"pred": pred_sa},
        )

    differences = []

    # Composition case: check if any commands are sub-node commands
    if spec and spec.get("composition"):
        comp_commands = composition_command_ids(spec)
        has_subnode = any(c in SMARTART_SUBNODE_COMMANDS for c in comp_commands)
        if has_subnode:
            gold_sub = read_smartart_subnode_properties(str(gold_path))
            pred_sub = read_smartart_subnode_properties(str(pred_path))
            if not gold_sub.get("found"):
                return VerificationResult.error_result("Gold: SmartArt sub-node not found")
            if not pred_sub.get("found"):
                return VerificationResult(
                    passed=False,
                    message="FAIL: Predicted SmartArt sub-node not found",
                    details={"pred": pred_sub},
                )

            # Composition contracts own only the properties explicitly
            # named by their atomic-intent verification targets.  This is
            # especially important for named text bevels: Angle requests
            # bevel identity, while an unrelated custom depth remains
            # outside the declared contract unless depth is itself named.
            supported_exact = {
                "connector_begin_arrow",
                "connector_end_arrow",
                "line_dash_style",
                "text_fill_rgb",
                "warp_format",
                "bevel_top_type",
                "threed_preset_camera",
            }
            default_numeric_tolerances = {
                "line_weight": 0.5,
                "threed_rotation_x": 1.0,
                "threed_rotation_y": 1.0,
                "bevel_top_depth": 1.0,
            }
            requested_targets = []
            for intent in spec.get("composition", {}).get("atomic_intents", []) or []:
                for target in intent.get("verification_targets", []) or []:
                    if target.get("extractor") == ("smartart_subnode_properties"):
                        requested_targets.append(target)

            if requested_targets:
                unknown = sorted(
                    {
                        target.get("property")
                        for target in requested_targets
                        if target.get("property") not in supported_exact
                        and target.get("property") not in (default_numeric_tolerances)
                    },
                    key=str,
                )
                if unknown:
                    return VerificationResult.error_result(
                        "Unsupported SmartArt sub-node composition " f"target(s): {unknown}"
                    )
                exact_targets = sorted(
                    {
                        target.get("property")
                        for target in requested_targets
                        if target.get("property") in supported_exact
                    }
                )
                numeric_targets: Dict[str, float] = {}
                for target in requested_targets:
                    prop = target.get("property")
                    if prop not in default_numeric_tolerances:
                        continue
                    tolerance = target.get("tolerance", default_numeric_tolerances[prop])
                    try:
                        numeric_targets[prop] = float(tolerance)
                    except (TypeError, ValueError):
                        return VerificationResult.error_result(
                            "Invalid SmartArt sub-node tolerance for " f"{prop}: {tolerance!r}"
                        )
            else:
                # Legacy composition specs have no property projection.
                # Preserve the established broad comparison for them so a
                # missing schema cannot turn into a silent pass.
                exact_targets = sorted(supported_exact)
                numeric_targets = dict(default_numeric_tolerances)

            required = [*exact_targets, *numeric_targets]
            error = require_observations(
                {"gold": gold_sub, "pred": pred_sub}, required,
                context="SmartArt sub-node composition")
            if error is not None:
                error.details["missing_evidence"] = list(error.details.get("unavailable", {}))
                if error.details.get("observation_owner") == "gold":
                    error.details["error_owner"] = "gold_fixture_or_evaluator"
                return error

            for prop in exact_targets:
                gv = gold_sub.get(prop)
                pv = pred_sub.get(prop)
                if not requested_targets and gv is None and pv is None:
                    continue
                if gv != pv:
                    differences.append(
                        {
                            "property": prop,
                            "gold": gv,
                            "pred": pv,
                        }
                    )
            for prop, tol in numeric_targets.items():
                gv = gold_sub.get(prop)
                pv = pred_sub.get(prop)
                if not requested_targets and gv is None and pv is None:
                    continue
                if gv is not None and pv is not None:
                    if abs(gv - pv) > tol:
                        differences.append(
                            {
                                "property": prop,
                                "gold": gv,
                                "pred": pv,
                            }
                        )
                elif gv != pv:
                    differences.append(
                        {
                            "property": prop,
                            "gold": gv,
                            "pred": pv,
                        }
                    )
            if differences:
                return VerificationResult(
                    passed=False,
                    message=f"FAIL: SmartArt composition mismatch ({len(differences)} difference(s))",
                    details={"differences": differences, "gold": gold_sub, "pred": pred_sub},
                )
            return VerificationResult(
                passed=True,
                message="PASS: SmartArt composition properties match",
                details={
                    "operation_count": spec["composition"].get("operation_count"),
                    "requested_subnode_properties": [*exact_targets, *numeric_targets],
                },
            )

    # SmartArt sub-node commands: use sub-node extraction for precise comparison
    if idMso in SMARTART_SUBNODE_COMMANDS:
        gold_sub = read_smartart_subnode_properties(str(gold_path))
        pred_sub = read_smartart_subnode_properties(str(pred_path))

        if not gold_sub.get("found"):
            return VerificationResult.error_result("Gold: SmartArt sub-node not found")
        if not pred_sub.get("found"):
            return VerificationResult(
                passed=False,
                message="FAIL: Predicted SmartArt sub-node not found",
                details={"pred": pred_sub},
            )

        required = {
            "ArrowStyleGallery": ("connector_begin_arrow", "connector_end_arrow"),
            "OutlineDashesGallery": ("line_dash_style",),
            "OutlineWeightGallery": ("line_weight",),
        }.get(idMso, ())
        error = require_observations(
            {"gold": gold_sub, "pred": pred_sub}, required, context=f"SmartArt {idMso}")
        if error is not None:
            return error

        if idMso == "ArrowStyleGallery":
            for prop in ["connector_begin_arrow", "connector_end_arrow"]:
                if gold_sub.get(prop) != pred_sub.get(prop):
                    differences.append(
                        {
                            "property": prop,
                            "gold": gold_sub.get(prop),
                            "pred": pred_sub.get(prop),
                        }
                    )
        elif idMso == "OutlineDashesGallery":
            if gold_sub.get("line_dash_style") != pred_sub.get("line_dash_style"):
                differences.append(
                    {
                        "property": "line_dash_style",
                        "gold": gold_sub.get("line_dash_style"),
                        "pred": pred_sub.get("line_dash_style"),
                    }
                )
        elif idMso == "OutlineWeightGallery":
            gold_w = gold_sub.get("line_weight")
            pred_w = pred_sub.get("line_weight")
            if gold_w is not None and pred_w is not None:
                if abs(gold_w - pred_w) > 0.5:
                    differences.append(
                        {
                            "property": "line_weight",
                            "gold": gold_w,
                            "pred": pred_w,
                        }
                    )
            elif gold_w != pred_w:
                differences.append(
                    {
                        "property": "line_weight",
                        "gold": gold_w,
                        "pred": pred_w,
                    }
                )

        if differences:
            return VerificationResult(
                passed=False,
                message=f"FAIL: SmartArt sub-node mismatch ({len(differences)} difference(s))",
                details={
                    "differences": differences,
                    "idMso": idMso,
                    "gold": gold_sub,
                    "pred": pred_sub,
                },
            )
        return VerificationResult(
            passed=True,
            message="PASS: SmartArt sub-node properties match",
            details={"idMso": idMso},
        )

    # Determine which SmartArt properties to compare based on command
    if idMso == "SmartArtLayoutGallery":
        props = ["layout_name"]
    elif idMso == "SmartArtStylesGallery":
        props = ["quick_style_name"]
    elif idMso == "SmartArtChangeColorsGallery":
        props = ["color_name"]
    else:
        props = ["layout_name", "color_name", "quick_style_name"]

    error = require_observations(
        {"gold": gold_sa, "pred": pred_sa}, props, context=f"SmartArt {idMso} gallery")
    if error is not None:
        return error

    for prop in props:
        gold_val = gold_sa.get(prop, "")
        pred_val = pred_sa.get(prop, "")
        if gold_val != pred_val:
            differences.append(
                {
                    "property": prop,
                    "gold": gold_val,
                    "pred": pred_val,
                }
            )

    # Named gallery identity is the primary semantic contract.  Never use
    # a refreshed cache to excuse a Word-COM identity mismatch.
    if differences:
        return VerificationResult(
            passed=False,
            message=(
                "FAIL: SmartArt Word gallery identity differs "
                f"({len(differences)} difference(s))"
            ),
            details={
                "differences": differences,
                "idMso": idMso,
                "gold": gold_sa,
                "pred": pred_sa,
                "verification_source": "word_com",
            },
        )

    if idMso == "SmartArtLayoutGallery":
        gold_drawing = read_smartart_drawing_info(str(gold_path))
        pred_drawing = read_smartart_drawing_info(str(pred_path))
        if gold_drawing.get("found") and not pred_drawing.get("found"):
            differences.append(
                {
                    "property": "smartart_drawing_present",
                    "gold": True,
                    "pred": False,
                }
            )
        elif (
            gold_drawing.get("found")
            and gold_drawing.get("node_order")
            and pred_drawing.get("node_order")
            and sorted(str(text).strip() for text in gold_drawing.get("node_order", []))
            != sorted(str(text).strip() for text in pred_drawing.get("node_order", []))
        ):
            # An empty compact-cache extraction can still correspond to a
            # correctly rendered SmartArt whose text is stored through
            # diagram-data references.  Treat it as inconclusive; compare
            # visible node membership only when both saved caches expose it.
            differences.append(
                {
                    "property": "smartart_drawing_node_texts",
                    "gold": sorted(
                        str(text).strip() for text in gold_drawing.get("node_order", [])
                    ),
                    "pred": sorted(
                        str(text).strip() for text in pred_drawing.get("node_order", [])
                    ),
                }
            )
    elif idMso in {"SmartArtStylesGallery", "SmartArtChangeColorsGallery"}:
        definition_result = verify_gallery_definitions(pred_path, gold_path, idMso)
        if not definition_result.passed:
            return definition_result
        gold_drawing = read_smartart_drawing_info(str(gold_path))
        pred_drawing = read_smartart_drawing_info(str(pred_path))
        if gold_drawing.get("found") and not pred_drawing.get("found"):
            differences.append(
                {
                    "property": "smartart_drawing_present",
                    "gold": True,
                    "pred": False,
                }
            )
        elif gold_drawing.get("found") and gold_drawing.get("node_styles") != pred_drawing.get(
            "node_styles"
        ):
            raw_cache_trigger = {
                "property": "smartart_drawing_node_styles",
                "gold": gold_drawing.get("node_styles"),
                "pred": pred_drawing.get("node_styles"),
            }
            return families_smartart_refresh.verify_refreshed_smartart_gallery_projection(
                pred_path,
                gold_path,
                gold_sa,
                pred_sa,
                raw_cache_trigger,
                idMso,
                extract_refreshed_smartart_projection=extract_refreshed_smartart_projection,
                get_cleanup_word=get_cleanup_word,
            )

    if differences:
        return VerificationResult(
            passed=False,
            message=f"FAIL: SmartArt mismatch ({len(differences)} difference(s))",
            details={
                "differences": differences,
                "idMso": idMso,
                "gold": gold_sa,
                "pred": pred_sa,
            },
        )

    return VerificationResult(
        passed=True,
        message="PASS: SmartArt properties match",
        details={
            "idMso": idMso,
            "verification_source": (
                "word_com+saved_ooxml_guardrail"
                if idMso
                in {
                    "SmartArtLayoutGallery",
                    "SmartArtStylesGallery",
                    "SmartArtChangeColorsGallery",
                }
                else "word_com"
            ),
        },
    )
