"""families / smartart / saved_projection with explicit evidence dependencies."""

from typing import Dict, List, Optional, Any
import os
import re
import xml.etree.ElementTree as ET
import zipfile


def extract_smartart_camera_state(docx_path: str) -> Dict[str, Any]:
    """Extract 3-D camera presets separately from SmartArt data and cache.

    A ``drawing*.xml`` camera can be a stale rendering-cache shell.  Named
    text 3-D rotations become Word-effective only when the semantic
    ``data*.xml`` state and the drawing state carry the same non-default
    preset.  Return both projections so the route verifier can require
    agreement without comparing unrelated diagram markup.

    The same ``a:scene3d`` element means two different things depending on
    the property it is written under.  Beneath a text body (``a:bodyPr``)
    it rotates the text; beneath shape properties (``spPr``) it rotates the
    SmartArt shape itself.  Project both hosts separately so a text-effect
    route can require the effect at the property level Word actually uses
    for that command, instead of accepting any camera anywhere in the
    diagram.
    """

    def unique(values: List[str]) -> List[str]:
        result: List[str] = []
        for value in values:
            if value and value not in result:
                result.append(value)
        non_default = [value for value in result if value != "orthographicFront"]
        return sorted(non_default or result)

    def local(tag: str) -> str:
        return tag.rsplit("}", 1)[-1]

    result: Dict[str, Any] = {
        "data_presets": [],
        "drawing_presets": [],
        "parts": [],
    }
    try:
        with zipfile.ZipFile(docx_path) as zf:
            diagram_parts = sorted(
                name
                for name in zf.namelist()
                if name.startswith("word/diagrams/")
                and name.endswith(".xml")
                and os.path.basename(name).startswith(("data", "drawing"))
            )

            by_kind: Dict[str, List[str]] = {
                "data": [],
                "drawing": [],
            }
            by_host: Dict[str, List[str]] = {
                "text_body": [],
                "shape": [],
                "other": [],
            }
            for name in diagram_parts:
                basename = os.path.basename(name)
                kind = "data" if basename.startswith("data") else "drawing"
                root = ET.fromstring(zf.read(name))
                parents = {child: parent for parent in root.iter() for child in parent}
                part_by_host: Dict[str, List[str]] = {
                    "text_body": [],
                    "shape": [],
                    "other": [],
                }
                presets: List[str] = []
                for scene in root.iter():
                    if local(scene.tag) != "scene3d":
                        continue
                    cameras = unique(
                        [
                            node.get("prst", "")
                            for node in scene.iter()
                            if local(node.tag) == "camera"
                        ]
                    )
                    if not cameras:
                        continue
                    presets.extend(cameras)
                    host_tag = local(parents[scene].tag) if scene in parents else ""
                    host = (
                        "text_body"
                        if host_tag == "bodyPr"
                        else "shape" if host_tag == "spPr" else "other"
                    )
                    part_by_host[host].extend(cameras)
                    by_host[host].extend(cameras)
                presets = unique(presets)
                by_kind[kind].extend(presets)
                result["parts"].append(
                    {
                        "part": name,
                        "kind": kind,
                        "presets": presets,
                        "text_body_presets": unique(part_by_host["text_body"]),
                        "shape_presets": unique(part_by_host["shape"]),
                        "other_host_presets": unique(part_by_host["other"]),
                    }
                )

            result["data_presets"] = unique(by_kind["data"])
            result["drawing_presets"] = unique(by_kind["drawing"])
            result["text_body_presets"] = unique(by_host["text_body"])
            result["shape_presets"] = unique(by_host["shape"])
            result["other_host_presets"] = unique(by_host["other"])
            result["all_presets"] = unique(by_kind["data"] + by_kind["drawing"])
            result["data_drawing_agree"] = bool(
                result["data_presets"] and result["data_presets"] == result["drawing_presets"]
            )
            return result
    except (OSError, zipfile.BadZipFile, ET.ParseError) as exc:
        result["error"] = str(exc)
        result["all_presets"] = []
        result["text_body_presets"] = []
        result["shape_presets"] = []
        result["other_host_presets"] = []
        result["data_drawing_agree"] = False
        return result


def extract_smartart_camera_presets(docx_path: str) -> List[str]:
    """
    Extract SmartArt 3D camera presets from diagram OOXML.

    Word COM can report default ThreeD values for SmartArt text even when the
    applied camera preset is present in the underlying diagram XML. The
    SmartArt diagram parts are the source of truth for these gallery presets.
    """
    state = extract_smartart_camera_state(docx_path)
    return list(state.get("all_presets") or [])


def extract_smartart_drawing_info(docx_path: str) -> Dict:
    """Extract the rendered SmartArt drawing cache from OOXML.

    Word's ``SmartArt.Nodes`` can reflect a changed data model even when
    the saved ``drawing*.xml`` still contains the old visual arrangement.
    That inconsistency is user-visible: Word renders the stale drawing
    until it successfully regenerates the cache. Return only stable,
    semantic information needed by the verifier: drawing-part presence
    and node text in page reading order.
    """
    result: Dict[str, Any] = {
        "found": False,
        "part_count": 0,
        "parts": [],
        "node_order": [],
        "node_layout": [],
    }

    def local_name(tag: str) -> str:
        return tag.rsplit("}", 1)[-1]

    def first_color(container: Any) -> Optional[str]:
        if container is None:
            return None
        for node in container.iter():
            name = local_name(node.tag)
            if name in {"schemeClr", "srgbClr", "prstClr", "sysClr"}:
                value = node.attrib.get("val") or node.attrib.get("lastClr")
                if value:
                    return f"{name}:{value}"
        return None

    try:
        with zipfile.ZipFile(docx_path) as zf:
            names = sorted(
                name
                for name in zf.namelist()
                if name.startswith("word/diagrams/drawing") and name.endswith(".xml")
            )
            result["found"] = bool(names)
            result["part_count"] = len(names)
            combined_order: List[str] = []
            for name in names:
                root = ET.fromstring(zf.read(name))
                nodes = []
                for shape in root.iter():
                    if local_name(shape.tag) != "sp":
                        continue
                    texts = [
                        (node.text or "").strip()
                        for node in shape.iter()
                        if local_name(node.tag) == "t" and (node.text or "").strip()
                    ]
                    if not texts:
                        continue
                    shape_properties = next(
                        (child for child in shape if local_name(child.tag) == "spPr"),
                        None,
                    )
                    text_body = next(
                        (child for child in shape if local_name(child.tag) == "txBody"),
                        None,
                    )
                    x = y = width = height = None
                    for transform in shape.iter():
                        if local_name(transform.tag) != "xfrm":
                            continue
                        for child in transform:
                            if local_name(child.tag) == "off":
                                x = int(child.attrib.get("x", "0"))
                                y = int(child.attrib.get("y", "0"))
                            elif local_name(child.tag) == "ext":
                                width = int(child.attrib.get("cx", "0"))
                                height = int(child.attrib.get("cy", "0"))
                        break
                    orientation = None
                    if text_body is not None:
                        body_properties = next(
                            (child for child in text_body if local_name(child.tag) == "bodyPr"),
                            None,
                        )
                        if body_properties is not None:
                            orientation = body_properties.attrib.get("vert")
                    nodes.append(
                        {
                            "text": " ".join(texts),
                            "x": x,
                            "y": y,
                            "width": width,
                            "height": height,
                            "orientation": orientation,
                            "fill_color": first_color(shape_properties),
                            "text_color": first_color(text_body),
                        }
                    )
                ordered = sorted(
                    nodes,
                    key=lambda item: (
                        item["y"] if item["y"] is not None else float("inf"),
                        item["x"] if item["x"] is not None else float("inf"),
                        item["text"],
                    ),
                )
                node_order = [item["text"] for item in ordered]
                node_styles = [
                    {
                        "text": item["text"],
                        "fill_color": item["fill_color"],
                        "text_color": item["text_color"],
                    }
                    for item in ordered
                ]
                node_layout = [
                    {
                        "text": item["text"],
                        "x": item["x"],
                        "y": item["y"],
                        "width": item["width"],
                        "height": item["height"],
                        "orientation": item["orientation"],
                    }
                    for item in ordered
                ]
                combined_order.extend(node_order)
                result["parts"].append(
                    {
                        "part": name,
                        "node_count": len(nodes),
                        "node_order": node_order,
                        "node_styles": node_styles,
                        "node_layout": node_layout,
                    }
                )
            result["node_order"] = combined_order
            result["node_styles"] = [
                style for part in result["parts"] for style in part["node_styles"]
            ]
            result["node_layout"] = [
                layout for part in result["parts"] for layout in part["node_layout"]
            ]
            return result
    except (OSError, zipfile.BadZipFile, ET.ParseError) as exc:
        result["error"] = str(exc)
        return result


def extract_smartart_gallery_identities(docx_path: str) -> Dict[str, Any]:
    """Extract the semantic gallery IDs persisted in SmartArt data parts.

    Word appends an instance suffix such as ``#1`` or ``#2`` when it
    refreshes a diagram.  That suffix identifies the local gallery-part
    instance, not a different named style/color/layout.  Preserve the raw
    values for diagnostics while also returning suffix-free semantic IDs
    for comparison after a Word SaveAs refresh.
    """

    diagram = "http://schemas.openxmlformats.org/drawingml/2006/diagram"

    def semantic_id(value: Optional[str]) -> str:
        return re.sub(r"#\d+$", "", str(value or "").strip())

    result: Dict[str, Any] = {
        "found": False,
        "identities": [],
    }
    try:
        with zipfile.ZipFile(docx_path) as package:
            names = sorted(
                name
                for name in package.namelist()
                if name.startswith("word/diagrams/data") and name.endswith(".xml")
            )
            for name in names:
                root = ET.fromstring(package.read(name))
                document_points = [
                    point for point in root.iter(f"{{{diagram}}}pt") if point.get("type") == "doc"
                ]
                for point in document_points:
                    property_sets = [child for child in point if child.tag == f"{{{diagram}}}prSet"]
                    for properties in property_sets:
                        raw_layout = properties.get("loTypeId", "")
                        raw_style = properties.get("qsTypeId", "")
                        raw_color = properties.get("csTypeId", "")
                        if not any((raw_layout, raw_style, raw_color)):
                            continue
                        result["identities"].append(
                            {
                                "part": name,
                                "layout_type_id": semantic_id(raw_layout),
                                "quick_style_type_id": semantic_id(raw_style),
                                "color_type_id": semantic_id(raw_color),
                                "raw_layout_type_id": raw_layout,
                                "raw_quick_style_type_id": raw_style,
                                "raw_color_type_id": raw_color,
                            }
                        )
        result["found"] = bool(result["identities"])
        return result
    except (OSError, zipfile.BadZipFile, ET.ParseError) as exc:
        result["error"] = str(exc)
        return result


def extract_smartart_node_scale_info(docx_path: str) -> Dict:
    """Extract persisted per-node SmartArt custom scales from data XML.

    SmartArt larger/smaller operations are stored on presentation points
    as ``custScaleX``/``custScaleY``.  A drawing part is only a render
    cache and may retain dimensions that disagree with this data model, so
    it is not authoritative for whether the resize operation persisted.
    Missing custom-scale attributes have the schema-default value 100000.
    """
    dgm = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
    a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    ns = {"dgm": dgm, "a": a}
    result: Dict[str, Any] = {
        "found": False,
        "part_count": 0,
        "parts": [],
        "nodes": [],
    }
    try:
        with zipfile.ZipFile(docx_path) as zf:
            names = sorted(
                name
                for name in zf.namelist()
                if name.startswith("word/diagrams/data") and name.endswith(".xml")
            )
            result["found"] = bool(names)
            result["part_count"] = len(names)
            combined_nodes: List[Dict[str, Any]] = []
            for name in names:
                root = ET.fromstring(zf.read(name))
                text_by_model_id: Dict[str, str] = {}
                points = root.findall(".//dgm:pt", ns)
                for point in points:
                    model_id = point.attrib.get("modelId")
                    texts = [
                        (node.text or "").strip()
                        for node in point.findall(".//a:t", ns)
                        if (node.text or "").strip()
                    ]
                    if model_id and texts:
                        text_by_model_id[model_id] = " ".join(texts)

                nodes = []
                for point in points:
                    properties = point.find("dgm:prSet", ns)
                    if properties is None:
                        continue
                    associated_id = properties.attrib.get("presAssocID")
                    if (
                        properties.attrib.get("presName") != "node"
                        or not associated_id
                        or associated_id not in text_by_model_id
                    ):
                        continue
                    scale_x = int(properties.attrib.get("custScaleX", "100000"))
                    scale_y = int(properties.attrib.get("custScaleY", "100000"))
                    style_index = properties.attrib.get("presStyleIdx")
                    nodes.append(
                        {
                            "text": text_by_model_id[associated_id],
                            "scale_x": scale_x,
                            "scale_y": scale_y,
                            "explicit_custom_scale": (
                                "custScaleX" in properties.attrib
                                or "custScaleY" in properties.attrib
                            ),
                            "presentation_style_index": (
                                int(style_index) if style_index is not None else None
                            ),
                        }
                    )
                nodes.sort(
                    key=lambda row: (
                        (
                            row["presentation_style_index"]
                            if row["presentation_style_index"] is not None
                            else float("inf")
                        ),
                        row["text"],
                    )
                )
                result["parts"].append({"part": name, "nodes": nodes})
                combined_nodes.extend(nodes)
            result["nodes"] = combined_nodes
            return result
    except (OSError, ValueError, zipfile.BadZipFile, ET.ParseError) as exc:
        result["error"] = str(exc)
        return result
