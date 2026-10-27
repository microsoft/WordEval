"""composition_effects with explicit evidence dependencies."""

from evaluation.methods.word_object.effect_commands import SHAPE_COLLECTION_PROPERTIES
from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.composition import composition_steps as declared_steps
from evaluation.methods.word_object.core import DocumentOpenError
from evaluation.methods.word_object.families.equations.compare import (
    equation_contract_equivalence, normalize_equation_contract,
)
from pathlib import Path
from typing import Any, Dict, List, Optional
import math


def composition_steps(spec: Optional[Dict]) -> List[Dict]:
    """Return normalized composition steps from either spec format."""
    return declared_steps(spec)


def scoped_spec(spec: Optional[Dict], idMso: str) -> Dict:
    """Create a single-command spec so atomic verifiers keep their scope."""
    scoped = dict(spec or {})
    scoped["idMso"] = idMso
    scoped.setdefault("precondition", {})
    return scoped


def merge_sub_results(message: str, sub_results: Dict) -> VerificationResult:
    if not sub_results:
        result = VerificationResult.error_result(
            "Composition produced no handled verification steps"
        )
        result.details["reason"] = "empty_composition_sub_results"
        return result
    sources = {
        result.get("verification_source")
        for result in sub_results.values()
        if isinstance(result, dict) and result.get("verification_source")
    }
    source = None
    if "word_com+saved_ooxml_guardrail" in sources:
        source = "word_com+saved_ooxml_guardrail"
    elif "saved_ooxml_only_wom_unavailable" in sources:
        source = "saved_ooxml_only_wom_unavailable"
    elif "saved_ooxml_fallback" in sources:
        source = "saved_ooxml_fallback"
    elif "word_com" in sources:
        source = "word_com"
    details = {"sub_results": sub_results}
    if source:
        details["verification_source"] = source
    return VerificationResult(
        passed=True,
        message=message,
        details=details,
    )


def add_composition_atomic_details(
    details: Dict,
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict],
    *,
    calculate_actual_position,
    extract_bevel_render_signatures,
    extract_icon_info,
    extract_saved_chart_properties,
    extract_saved_shape_properties,
    extract_saved_smartart_connector_styles,
    extract_smartart_node_details,
    extract_table_autofit_flags,
    match_shapes_by_name,
    normalized_equation_text,
    get_position_tolerance,
    read_chart_properties,
    read_equation_structure,
    read_picture_crop_properties,
    read_picture_effect_properties,
    read_shape_properties,
    read_smartart_placement,
    read_smartart_properties,
    read_smartart_subnode_properties,
    read_table_properties,
) -> None:
    atomic_results = build_composition_atomic_results(
        pred_path,
        gold_path,
        spec,
        calculate_actual_position=calculate_actual_position,
        extract_bevel_render_signatures=extract_bevel_render_signatures,
        extract_icon_info=extract_icon_info,
        extract_saved_chart_properties=extract_saved_chart_properties,
        extract_saved_shape_properties=extract_saved_shape_properties,
        extract_saved_smartart_connector_styles=extract_saved_smartart_connector_styles,
        extract_smartart_node_details=extract_smartart_node_details,
        extract_table_autofit_flags=extract_table_autofit_flags,
        match_shapes_by_name=match_shapes_by_name,
        normalized_equation_text=normalized_equation_text,
        get_position_tolerance=get_position_tolerance,
        read_chart_properties=read_chart_properties,
        read_equation_structure=read_equation_structure,
        read_picture_crop_properties=read_picture_crop_properties,
        read_picture_effect_properties=read_picture_effect_properties,
        read_shape_properties=read_shape_properties,
        read_smartart_placement=read_smartart_placement,
        read_smartart_properties=read_smartart_properties,
        read_smartart_subnode_properties=read_smartart_subnode_properties,
        read_table_properties=read_table_properties,
    )
    if not atomic_results:
        return

    flat_targets = []
    for result in atomic_results:
        for target_result in result.get("target_results", []):
            flat = {
                "intent_id": result["intent_id"],
                "operation": result.get("operation"),
                "atomic_case": result.get("atomic_case"),
            }
            flat.update(target_result)
            flat_targets.append(flat)

    details["atomic_results"] = atomic_results
    details["target_results"] = flat_targets
    details["atomic_summary"] = {
        "passed_intents": sum(1 for r in atomic_results if r.get("passed")),
        "total_intents": len(atomic_results),
        "passed_targets": sum(1 for r in flat_targets if r.get("passed")),
        "total_targets": len(flat_targets),
    }


def build_composition_atomic_results(
    pred_path: Path,
    gold_path: Path,
    spec: Optional[Dict],
    *,
    calculate_actual_position,
    extract_bevel_render_signatures,
    extract_icon_info,
    extract_saved_chart_properties,
    extract_saved_shape_properties,
    extract_saved_smartart_connector_styles,
    extract_smartart_node_details,
    extract_table_autofit_flags,
    match_shapes_by_name,
    normalized_equation_text,
    get_position_tolerance,
    read_chart_properties,
    read_equation_structure,
    read_picture_crop_properties,
    read_picture_effect_properties,
    read_shape_properties,
    read_smartart_placement,
    read_smartart_properties,
    read_smartart_subnode_properties,
    read_table_properties,
) -> List[Dict[str, Any]]:
    intents = (spec or {}).get("composition", {}).get("atomic_intents", [])
    if not intents:
        return []

    cache: Dict[str, Any] = {
        "__allow_floating_picture__": any(
            intent.get("operation") == "insert_icon" for intent in intents
        ),
        "__smartart_heading__": (spec or {}).get("precondition", {}).get("heading"),
        "__equation_contract_equivalence__": equation_contract_equivalence(spec or {}),
    }
    results = []
    interaction_class = (spec or {}).get("composition", {}).get("interaction_class", "")
    for intent in intents:
        target_results = []
        for target in intent.get("verification_targets", []) or []:
            target_results.append(
                evaluate_composition_target(
                    target,
                    pred_path,
                    gold_path,
                    cache,
                    calculate_actual_position=calculate_actual_position,
                    extract_bevel_render_signatures=extract_bevel_render_signatures,
                    extract_icon_info=extract_icon_info,
                    extract_saved_chart_properties=extract_saved_chart_properties,
                    extract_saved_shape_properties=extract_saved_shape_properties,
                    extract_saved_smartart_connector_styles=extract_saved_smartart_connector_styles,
                    extract_smartart_node_details=extract_smartart_node_details,
                    extract_table_autofit_flags=extract_table_autofit_flags,
                    match_shapes_by_name=match_shapes_by_name,
                    normalized_equation_text=normalized_equation_text,
                    read_chart_properties=read_chart_properties,
                    read_equation_structure=read_equation_structure,
                    read_picture_crop_properties=read_picture_crop_properties,
                    read_picture_effect_properties=read_picture_effect_properties,
                    read_shape_properties=read_shape_properties,
                    read_smartart_placement=read_smartart_placement,
                    read_smartart_properties=read_smartart_properties,
                    read_smartart_subnode_properties=read_smartart_subnode_properties,
                    read_table_properties=read_table_properties,
                )
            )
        if (
            interaction_class == "group_flip_ungroup_distribute"
            and intent.get("operation") == "distribute_command"
        ):
            try:
                gold_shapes = extract_composition_target_data(
                    "multi_shape_properties",
                    gold_path,
                    cache,
                    "gold",
                    extract_bevel_render_signatures=extract_bevel_render_signatures,
                    extract_icon_info=extract_icon_info,
                    extract_saved_chart_properties=extract_saved_chart_properties,
                    extract_saved_shape_properties=extract_saved_shape_properties,
                    extract_saved_smartart_connector_styles=extract_saved_smartart_connector_styles,
                    extract_smartart_node_details=extract_smartart_node_details,
                    extract_table_autofit_flags=extract_table_autofit_flags,
                    read_chart_properties=read_chart_properties,
                    read_equation_structure=read_equation_structure,
                    read_picture_crop_properties=read_picture_crop_properties,
                    read_picture_effect_properties=read_picture_effect_properties,
                    read_shape_properties=read_shape_properties,
                    read_smartart_placement=read_smartart_placement,
                    read_smartart_properties=read_smartart_properties,
                    read_smartart_subnode_properties=read_smartart_subnode_properties,
                    read_table_properties=read_table_properties,
                )
                pred_shapes = extract_composition_target_data(
                    "multi_shape_properties",
                    pred_path,
                    cache,
                    "pred",
                    extract_bevel_render_signatures=extract_bevel_render_signatures,
                    extract_icon_info=extract_icon_info,
                    extract_saved_chart_properties=extract_saved_chart_properties,
                    extract_saved_shape_properties=extract_saved_shape_properties,
                    extract_saved_smartart_connector_styles=extract_saved_smartart_connector_styles,
                    extract_smartart_node_details=extract_smartart_node_details,
                    extract_table_autofit_flags=extract_table_autofit_flags,
                    read_chart_properties=read_chart_properties,
                    read_equation_structure=read_equation_structure,
                    read_picture_crop_properties=read_picture_crop_properties,
                    read_picture_effect_properties=read_picture_effect_properties,
                    read_shape_properties=read_shape_properties,
                    read_smartart_placement=read_smartart_placement,
                    read_smartart_properties=read_smartart_properties,
                    read_smartart_subnode_properties=read_smartart_subnode_properties,
                    read_table_properties=read_table_properties,
                )
                gold_endpoints = horizontal_center_endpoints(
                    gold_shapes, calculate_actual_position=calculate_actual_position
                )
                pred_endpoints = horizontal_center_endpoints(
                    pred_shapes, calculate_actual_position=calculate_actual_position
                )
                tolerance = get_position_tolerance()
                endpoint_match = (
                    len(gold_endpoints) == 2
                    and len(pred_endpoints) == 2
                    and all(
                        abs(pred - gold) <= tolerance
                        for pred, gold in zip(pred_endpoints, gold_endpoints)
                    )
                )
                target_results.append(
                    {
                        "target": {
                            "extractor": "multi_shape_properties",
                            "property": "horizontal_center_endpoints",
                            "tolerance": tolerance,
                            "scope": interaction_class,
                        },
                        "gold": gold_endpoints,
                        "pred": pred_endpoints,
                        "passed": endpoint_match,
                    }
                )
            except DocumentOpenError:
                raise
            except Exception as exc:
                target_results.append(
                    {
                        "target": {
                            "extractor": "multi_shape_properties",
                            "property": "horizontal_center_endpoints",
                            "scope": interaction_class,
                        },
                        "passed": False,
                        "error": str(exc),
                    }
                )
        results.append(
            {
                "intent_id": intent.get("id"),
                "operation": intent.get("operation"),
                "atomic_case": intent.get("atomic_case"),
                "passed": bool(target_results)
                and all(target.get("passed") for target in target_results),
                "target_results": target_results,
            }
        )
    return results


def evaluate_composition_target(
    target: Dict[str, Any],
    pred_path: Path,
    gold_path: Path,
    cache: Dict[str, Any],
    *,
    calculate_actual_position,
    extract_bevel_render_signatures,
    extract_icon_info,
    extract_saved_chart_properties,
    extract_saved_shape_properties,
    extract_saved_smartart_connector_styles,
    extract_smartart_node_details,
    extract_table_autofit_flags,
    match_shapes_by_name,
    normalized_equation_text,
    read_chart_properties,
    read_equation_structure,
    read_picture_crop_properties,
    read_picture_effect_properties,
    read_shape_properties,
    read_smartart_placement,
    read_smartart_properties,
    read_smartart_subnode_properties,
    read_table_properties,
) -> Dict[str, Any]:
    extractor = target.get("extractor")
    prop = target.get("property")
    result = {"target": target}

    try:
        gold_data = extract_composition_target_data(
            extractor,
            gold_path,
            cache,
            "gold",
            extract_bevel_render_signatures=extract_bevel_render_signatures,
            extract_icon_info=extract_icon_info,
            extract_saved_chart_properties=extract_saved_chart_properties,
            extract_saved_shape_properties=extract_saved_shape_properties,
            extract_saved_smartart_connector_styles=extract_saved_smartart_connector_styles,
            extract_smartart_node_details=extract_smartart_node_details,
            extract_table_autofit_flags=extract_table_autofit_flags,
            read_chart_properties=read_chart_properties,
            read_equation_structure=read_equation_structure,
            read_picture_crop_properties=read_picture_crop_properties,
            read_picture_effect_properties=read_picture_effect_properties,
            read_shape_properties=read_shape_properties,
            read_smartart_placement=read_smartart_placement,
            read_smartart_properties=read_smartart_properties,
            read_smartart_subnode_properties=read_smartart_subnode_properties,
            read_table_properties=read_table_properties,
        )
        pred_data = extract_composition_target_data(
            extractor,
            pred_path,
            cache,
            "pred",
            extract_bevel_render_signatures=extract_bevel_render_signatures,
            extract_icon_info=extract_icon_info,
            extract_saved_chart_properties=extract_saved_chart_properties,
            extract_saved_shape_properties=extract_saved_shape_properties,
            extract_saved_smartart_connector_styles=extract_saved_smartart_connector_styles,
            extract_smartart_node_details=extract_smartart_node_details,
            extract_table_autofit_flags=extract_table_autofit_flags,
            read_chart_properties=read_chart_properties,
            read_equation_structure=read_equation_structure,
            read_picture_crop_properties=read_picture_crop_properties,
            read_picture_effect_properties=read_picture_effect_properties,
            read_shape_properties=read_shape_properties,
            read_smartart_placement=read_smartart_placement,
            read_smartart_properties=read_smartart_properties,
            read_smartart_subnode_properties=read_smartart_subnode_properties,
            read_table_properties=read_table_properties,
        )
        if (
            extractor == "multi_shape_properties"
            and isinstance(gold_data, list)
            and isinstance(pred_data, list)
        ):
            # shape_index in an atomic intent refers to the stable named
            # object from the generator, not its incidental XML/collection
            # order after Word saves the document.
            pred_data, gold_data = match_shapes_by_name(pred_data, gold_data)
        equivalence = cache.get("__equation_contract_equivalence__")
        if extractor == "equation_structure" and equivalence:
            result.update(comparison_equivalence=equivalence,
                          raw_gold=composition_target_value(
                              gold_data, prop, target, calculate_actual_position=calculate_actual_position),
                          raw_pred=composition_target_value(
                              pred_data, prop, target, calculate_actual_position=calculate_actual_position))
            gold_data = normalize_equation_contract(gold_data, equivalence)
            pred_data = normalize_equation_contract(pred_data, equivalence)
        gold_value = composition_target_value(
            gold_data, prop, target, calculate_actual_position=calculate_actual_position
        )
        pred_value = composition_target_value(
            pred_data, prop, target, calculate_actual_position=calculate_actual_position
        )
        if extractor == "equation_structure" and prop in {
            "equation_text",
            "equations",
        }:
            # Word inserts run/paragraph separators and incidental spaces
            # into OMath.Range.Text.  Whole-equation verification already
            # treats those characters as representation noise; atomic
            # composition targets must use the same text contract while
            # retaining exact function/type/parameter structure.
            gold_value = normalize_composition_equation_value(
                gold_value, normalized_equation_text=normalized_equation_text
            )
            pred_value = normalize_composition_equation_value(
                pred_value, normalized_equation_text=normalized_equation_text
            )
        result.update(
            {
                "gold": gold_value,
                "pred": pred_value,
                "passed": composition_values_match(
                    gold_value,
                    pred_value,
                    target.get("tolerance"),
                ),
            }
        )
    except DocumentOpenError:
        raise
    except Exception as exc:
        result.update(
            {
                "passed": False,
                "error": str(exc),
            }
        )
    return result


def extract_composition_target_data(
    extractor: str,
    path: Path,
    cache: Dict[str, Any],
    side: str,
    *,
    extract_bevel_render_signatures,
    extract_icon_info,
    extract_saved_chart_properties,
    extract_saved_shape_properties,
    extract_saved_smartart_connector_styles,
    extract_smartart_node_details,
    extract_table_autofit_flags,
    read_chart_properties,
    read_equation_structure,
    read_picture_crop_properties,
    read_picture_effect_properties,
    read_shape_properties,
    read_smartart_placement,
    read_smartart_properties,
    read_smartart_subnode_properties,
    read_table_properties,
) -> Any:
    key = f"{side}:{extractor}"
    if key in cache:
        return cache[key]

    if extractor == "multi_shape_properties":
        try:
            data = extract_saved_shape_properties(path)
        except Exception:
            data = read_shape_properties(str(path))
    elif extractor == "shape_properties":
        data = read_shape_properties(str(path))
    elif extractor == "picture_effect_properties":
        if cache.get("__allow_floating_picture__", False):
            data = read_picture_effect_properties(str(path), allow_floating=True)
        else:
            data = read_picture_effect_properties(str(path))
    elif extractor == "picture_crop_properties":
        data = read_picture_crop_properties(str(path))
    elif extractor == "chart_properties":
        data = read_chart_properties(str(path))
    elif extractor == "saved_chart_properties":
        data = extract_saved_chart_properties(path)
    elif extractor == "smartart_properties":
        data = read_smartart_properties(str(path))
    elif extractor == "smartart_node_details":
        data = extract_smartart_node_details(path)
    elif extractor == "smartart_placement":
        heading = cache.get("__smartart_heading__")
        if not heading:
            raise ValueError("SmartArt placement target requires a heading")
        data = read_smartart_placement(str(path), heading)
    elif extractor == "smartart_subnode_properties":
        data = read_smartart_subnode_properties(str(path))
    elif extractor == "smartart_connector_styles":
        data = extract_saved_smartart_connector_styles(path)
    elif extractor == "table_properties":
        data = read_table_properties(str(path))
    elif extractor == "table_autofit_flags":
        data = extract_table_autofit_flags(path)
    elif extractor == "equation_structure":
        data = read_equation_structure(str(path))
    elif extractor == "icon_info":
        data = extract_icon_info(path)
    elif extractor == "bevel_render_signatures":
        data = extract_bevel_render_signatures(path)
    else:
        raise ValueError(f"Unsupported object-effect target extractor: {extractor}")

    cache[key] = data
    return data


def composition_target_value(
    data: Any, prop: str, target: Optional[Dict[str, Any]] = None, *, calculate_actual_position
) -> Any:
    if isinstance(data, dict):
        if "error" in data:
            raise ValueError(data["error"])
        if prop in getattr(data, "unavailable", {}):
            raise ValueError(f"Unavailable {prop}: {data.unavailable[prop]}")
        if prop in getattr(data, "mixed", ()):
            raise ValueError(f"Required composition property is mixed: {prop}")
        if prop == "picture_effect_params":
            return picture_effect_params(data, (target or {}).get("effect_type"))
        if prop == "picture_effect_present":
            return picture_effect_present(data, (target or {}).get("effect_type"))
        if prop == "outer_border_line_styles":
            borders = data.get("borders") or {}
            return {
                name: (borders.get(name) or {}).get("line_style")
                for name in ("top", "left", "bottom", "right")
            }
        custom = custom_composition_target_value(
            data, prop, calculate_actual_position=calculate_actual_position
        )
        if custom is not None:
            return custom
        return data.get(prop)
    if isinstance(data, list):
        if prop == "table_count":
            return len(data)
        if prop in SHAPE_COLLECTION_PROPERTIES:
            return custom_composition_collection_value(
                data, prop, calculate_actual_position=calculate_actual_position
            )
        shape_index = (target or {}).get("shape_index")
        item_index = shape_index
        if item_index is None:
            item_index = (target or {}).get("table_index")
        if isinstance(item_index, int):
            zero_based_index = item_index - 1
            if zero_based_index < 0 or zero_based_index >= len(data):
                raise ValueError(
                    f"item index {item_index} out of range " f"for {len(data)} item(s)"
                )
            item = data[zero_based_index]
            return composition_target_value(
                item, prop, target, calculate_actual_position=calculate_actual_position
            )
        return [
            (
                composition_target_value(
                    item, prop, target, calculate_actual_position=calculate_actual_position
                )
                if isinstance(item, dict)
                else None
            )
            for item in data
        ]
    return None


def custom_composition_collection_value(
    data: List[Dict[str, Any]], prop: str, *, calculate_actual_position
) -> float:
    """Return a translation-invariant alignment/distribution residual."""
    boxes = composition_visual_boxes(data, calculate_actual_position=calculate_actual_position)

    value_key = {
        "left_edge_spread": "left",
        "right_edge_spread": "right",
        "horizontal_center_spread": "center_h",
        "top_edge_spread": "top",
        "bottom_edge_spread": "bottom",
        "vertical_center_spread": "center_v",
    }.get(prop)
    if value_key:
        values = [box[value_key] for box in boxes]
        return max(values) - min(values) if values else 0.0

    center_key = "center_h" if prop == "horizontal_center_gap_spread" else "center_v"
    centers = sorted(box[center_key] for box in boxes)
    gaps = [centers[index + 1] - centers[index] for index in range(len(centers) - 1)]
    return max(gaps) - min(gaps) if gaps else 0.0


def composition_visual_boxes(
    data: List[Dict[str, Any]], *, calculate_actual_position
) -> List[Dict[str, float]]:
    """Return rendered edge and center coordinates for saved shapes."""
    boxes = []
    for item in data:
        left = custom_composition_target_value(
            item, "actual_left", calculate_actual_position=calculate_actual_position
        )
        top = custom_composition_target_value(
            item, "actual_top", calculate_actual_position=calculate_actual_position
        )
        width = float(item.get("width") or 0)
        height = float(item.get("height") or 0)
        rotation = math.radians(float(item.get("rotation") or 0) % 360.0)
        visual_width = abs(width * math.cos(rotation)) + abs(height * math.sin(rotation))
        visual_height = abs(width * math.sin(rotation)) + abs(height * math.cos(rotation))
        visual_left = left + (width - visual_width) / 2.0
        visual_top = top + (height - visual_height) / 2.0
        boxes.append(
            {
                "left": visual_left,
                "right": visual_left + visual_width,
                "center_h": visual_left + visual_width / 2,
                "top": visual_top,
                "bottom": visual_top + visual_height,
                "center_v": visual_top + visual_height / 2,
            }
        )
    return boxes


def alignment_anchor(
    shapes: List[Dict[str, Any]], metric: str, *, calculate_actual_position
) -> Optional[float]:
    """Return the aligned edge/center position for a shape collection."""
    value_key = {
        "left_edge_spread": "left",
        "right_edge_spread": "right",
        "horizontal_center_spread": "center_h",
        "top_edge_spread": "top",
        "bottom_edge_spread": "bottom",
        "vertical_center_spread": "center_v",
    }.get(metric)
    if value_key is None:
        return None
    values = [
        box[value_key]
        for box in composition_visual_boxes(
            shapes, calculate_actual_position=calculate_actual_position
        )
    ]
    return sum(values) / len(values) if values else None


def distribution_center_endpoints(
    shapes: List[Dict[str, Any]], axis: str, *, calculate_actual_position
) -> List[float]:
    """Return the selected objects' endpoint centers on one axis."""
    center_key = "center_h" if axis == "horizontal" else "center_v"
    centers = sorted(
        box[center_key]
        for box in composition_visual_boxes(
            shapes, calculate_actual_position=calculate_actual_position
        )
    )
    if not centers:
        return []
    return [round(centers[0], 4), round(centers[-1], 4)]


def horizontal_center_endpoints(
    shapes: List[Dict[str, Any]], *, calculate_actual_position
) -> List[float]:
    return distribution_center_endpoints(
        shapes, "horizontal", calculate_actual_position=calculate_actual_position
    )


def picture_effect_params(data: Dict[str, Any], effect_type: Optional[int]) -> Any:
    for effect in data.get("picture_effects", []) or []:
        if effect.get("type") == effect_type:
            return effect.get("params", [])
    return None


def picture_effect_present(data: Dict[str, Any], effect_type: Optional[int]) -> bool:
    return any(
        effect.get("type") == effect_type for effect in data.get("picture_effects", []) or []
    )


def custom_composition_target_value(
    data: Dict[str, Any], prop: str, *, calculate_actual_position
) -> Any:
    if prop == "actual_left":
        return calculate_actual_position(
            data.get("left", 0),
            data.get("width", 0),
            data.get("relative_h_position", 0),
            data.get("page_width", 612),
            data.get("left_margin", 72),
            data.get("right_margin", 72),
        )
    if prop == "actual_top":
        return calculate_actual_position(
            data.get("top", 0),
            data.get("height", 0),
            data.get("relative_v_position", 0),
            data.get("page_height", 792),
            data.get("top_margin", 72),
            data.get("bottom_margin", 72),
        )
    if prop == "series_data_labels":
        return [
            {
                "index": series.get("index"),
                "has_data_labels": series.get("has_data_labels"),
                "data_label_position": series.get("data_label_position"),
            }
            for series in data.get("series_info", []) or []
        ]
    if prop == "series_trendlines":
        return [
            {
                "index": series.get("index"),
                "trendline_count": series.get("trendline_count"),
                "trendline_type": series.get("trendline_type"),
            }
            for series in data.get("series_info", []) or []
        ]
    if prop == "series_line_colors":
        series_info = data.get("series_info", []) or []
        series_one = next(
            (series for series in series_info if series.get("index") == 1),
            series_info[0] if series_info else None,
        )
        return (
            [
                {
                    "index": series_one.get("index"),
                    "line_color_rgb": series_one.get("line_color_rgb"),
                }
                # ColorPickerChartLineStyles targets Series 1.  Other series
                # retain unrequested theme/inherited outlines and must not be
                # reference-matched as collateral state.
            ]
            if series_one is not None
            else []
        )
    if prop == "value_axis_major_gridlines":
        for axis in data.get("axes", []) or []:
            if axis.get("axis_type") == 2:
                return axis.get("has_major_gridlines")
        return None
    if prop == "value_axis_minor_gridlines":
        for axis in data.get("axes", []) or []:
            if axis.get("axis_type") == 2:
                return axis.get("has_minor_gridlines")
        return None
    return None


def composition_values_match(
    gold_value: Any, pred_value: Any, tolerance: Optional[float] = None
) -> bool:
    if tolerance is not None:
        if isinstance(gold_value, list) and isinstance(pred_value, list):
            if len(gold_value) != len(pred_value):
                return False
            return all(
                composition_values_match(gv, pv, tolerance)
                for gv, pv in zip(gold_value, pred_value)
            )
        if isinstance(gold_value, (int, float)) and isinstance(pred_value, (int, float)):
            return abs(float(gold_value) - float(pred_value)) <= tolerance
    return gold_value == pred_value


def normalize_composition_equation_value(value: Any, *, normalized_equation_text) -> Any:
    if isinstance(value, list):
        return [
            normalize_composition_equation_value(
                item, normalized_equation_text=normalized_equation_text
            )
            for item in value
        ]
    if isinstance(value, dict):
        return {
            key: (
                normalized_equation_text({"equation_text": item})
                if key == "equation_text"
                else normalize_composition_equation_value(
                    item, normalized_equation_text=normalized_equation_text
                )
            )
            for key, item in value.items()
        }
    if isinstance(value, str):
        return normalized_equation_text({"equation_text": value})
    return value
