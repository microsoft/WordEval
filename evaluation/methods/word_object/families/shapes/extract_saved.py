"""families / shapes / extract_saved with explicit evidence dependencies."""

from evaluation.methods.word_object.effect_commands import (
    DRAWINGML_NS,
    EMU_PER_POINT,
    PICTURE_NS,
    ROTATION_UNITS_PER_DEGREE,
    WORDPROCESSINGML_NS,
    WORDPROCESSING_DRAWING_NS,
    WORDPROCESSING_GROUP_NS,
    WORDPROCESSING_SHAPE_NS,
)
from pathlib import Path
from typing import Any, Dict, List, Optional
import xml.etree.ElementTree as ET
import zipfile


def extract_saved_shape_properties(docx_path: Path) -> List[Dict[str, Any]]:
    """Read the core floating-shape layout directly from DrawingML.

    This is intentionally a fallback for otherwise valid packages that
    Word cannot open or repair.  It covers the saved properties needed by
    alignment, distribution, rotation, flip, and group-lifecycle cases;
    it is not a replacement for COM where rendered Word state is required.
    """
    wp = f"{{{WORDPROCESSING_DRAWING_NS}}}"
    a = f"{{{DRAWINGML_NS}}}"
    wps = f"{{{WORDPROCESSING_SHAPE_NS}}}"
    wpg = f"{{{WORDPROCESSING_GROUP_NS}}}"
    w = f"{{{WORDPROCESSINGML_NS}}}"

    with zipfile.ZipFile(docx_path) as package:
        root = ET.fromstring(package.read("word/document.xml"))

    page_width = 612.0
    page_height = 792.0
    left_margin = right_margin = top_margin = bottom_margin = 72.0
    section = root.find(f".//{w}sectPr")
    if section is not None:
        page_size = section.find(f"{w}pgSz")
        page_margin = section.find(f"{w}pgMar")
        if page_size is not None:
            page_width = float(page_size.get(f"{w}w", 12240)) / 20.0
            page_height = float(page_size.get(f"{w}h", 15840)) / 20.0
        if page_margin is not None:
            left_margin = float(page_margin.get(f"{w}left", 1440)) / 20.0
            right_margin = float(page_margin.get(f"{w}right", 1440)) / 20.0
            top_margin = float(page_margin.get(f"{w}top", 1440)) / 20.0
            bottom_margin = float(page_margin.get(f"{w}bottom", 1440)) / 20.0

    horizontal_frames = {
        "margin": 0,
        "page": 1,
        "column": 2,
        "character": 3,
        "leftMargin": 4,
        "rightMargin": 5,
        "insideMargin": 6,
        "outsideMargin": 7,
    }
    vertical_frames = {
        "margin": 0,
        "page": 1,
        "paragraph": 2,
        "line": 3,
        "topMargin": 4,
        "bottomMargin": 5,
        "insideMargin": 6,
        "outsideMargin": 7,
    }

    def position(
        container: ET.Element,
        axis: str,
        frames: Dict[str, int],
    ) -> tuple[float, Optional[int]]:
        node = container.find(f"{wp}position{axis}")
        if node is None:
            if container.tag == f"{wp}inline":
                return 0.0, None
            raise ValueError(f"shape has no saved position{axis} in {docx_path}")
        offset = node.find(f"{wp}posOffset")
        if offset is None or offset.text is None:
            align = node.find(f"{wp}align")
            raise ValueError(
                f"unsupported saved alignment position "
                f"{align.text if align is not None else None!r} in {docx_path}"
            )
        return (
            float(offset.text) / EMU_PER_POINT,
            frames.get(node.get("relativeFrom", "")),
        )

    shapes: List[Dict[str, Any]] = []
    containers = [node for node in root.iter() if node.tag in {f"{wp}anchor", f"{wp}inline"}]
    for index, container in enumerate(containers, start=1):
        doc_pr = container.find(f"{wp}docPr")
        extent = container.find(f"{wp}extent")
        if extent is None:
            raise ValueError(f"shape has no saved extent in {docx_path}")
        width = float(extent.get("cx", 0)) / EMU_PER_POINT
        height = float(extent.get("cy", 0)) / EMU_PER_POINT
        left, relative_h = position(container, "H", horizontal_frames)
        top, relative_v = position(container, "V", vertical_frames)

        xfrm = container.find(f".//{wps}wsp/{wps}spPr/{a}xfrm")
        if xfrm is None:
            xfrm = container.find(f".//{a}xfrm")
        rotation = 0.0
        flip_h = False
        flip_v = False
        if xfrm is not None:
            rotation = (float(xfrm.get("rot", 0)) / ROTATION_UNITS_PER_DEGREE) % 360.0
            flip_h = xfrm.get("flipH") in {"1", "true"}
            flip_v = xfrm.get("flipV") in {"1", "true"}
        geometry = container.find(f".//{a}prstGeom")
        shadow = container.find(f".//{a}outerShdw")

        def saved_signature(element: Optional[ET.Element]) -> Any:
            if element is None:
                return None
            attributes = dict(element.attrib)
            tag = element.tag
            if element.tag == f"{a}outerShdw" and attributes.get("rotWithShape") in {"0", "false"}:
                attributes.pop("rotWithShape")
            if element.tag == f"{a}prstClr" and attributes.get("val", "").lower() == "black":
                # Word may serialize the same black shadow color as either
                # preset black or literal sRGB 000000.  Preserve alpha and
                # other descendants while canonicalizing that equivalent
                # color representation.
                tag = f"{a}srgbClr"
                attributes["val"] = "000000"
            elif element.tag == f"{a}srgbClr" and attributes.get("val"):
                attributes["val"] = attributes["val"].upper()
            return (
                tag,
                tuple(sorted(attributes.items())),
                tuple(saved_signature(child) for child in element),
            )

        shadow_signature = saved_signature(shadow)

        relative_height = container.get("relativeHeight")
        try:
            z_order = int(relative_height) if relative_height is not None else index
        except ValueError:
            z_order = index
        is_group = container.find(f".//{wpg}wgp") is not None
        shapes.append(
            {
                "name": doc_pr.get("name") if doc_pr is not None else f"Shape {index}",
                "left": round(left, 4),
                "top": round(top, 4),
                "width": round(width, 4),
                "height": round(height, 4),
                "center_h": round(left + width / 2, 4),
                "center_v": round(top + height / 2, 4),
                "z_order": z_order,
                "rotation": round(rotation, 4),
                "auto_shape_type": (geometry.get("prst") if geometry is not None else None),
                "shadow_type": shadow_signature,
                "shadow_signature": shadow_signature,
                "flip_horizontal": flip_h,
                "flip_vertical": flip_v,
                "is_group": is_group,
                "relative_h_position": relative_h,
                "relative_v_position": relative_v,
                "page_width": page_width,
                "page_height": page_height,
                "left_margin": left_margin,
                "right_margin": right_margin,
                "top_margin": top_margin,
                "bottom_margin": bottom_margin,
                "_source": "saved_ooxml",
            }
        )
    return shapes


def extract_saved_shape_style_signatures(docx_path: Path) -> List[Dict[str, Any]]:
    """Read composite shape-style state without using the Word OM.

    Shape style gallery presets are persisted through theme references in
    ``wps:style`` plus optional direct fill, line, effect, and 3-D
    overrides in ``wps:spPr``.  Geometry and placement are deliberately
    excluded because ShapeStylesGallery does not constrain them.
    """
    wp = f"{{{WORDPROCESSING_DRAWING_NS}}}"
    a = f"{{{DRAWINGML_NS}}}"
    wps = f"{{{WORDPROCESSING_SHAPE_NS}}}"

    with zipfile.ZipFile(docx_path) as package:
        root = ET.fromstring(package.read("word/document.xml"))

    def signature(element: Optional[ET.Element]) -> Any:
        if element is None:
            return None
        attributes = dict(element.attrib)
        if element.tag == f"{a}srgbClr" and attributes.get("val"):
            attributes["val"] = attributes["val"].upper()
        text = (element.text or "").strip() or None
        return (
            element.tag,
            tuple(sorted(attributes.items())),
            text,
            tuple(signature(child) for child in element),
        )

    names_by_shape: Dict[int, str] = {}
    for container in root.iter():
        if container.tag not in {f"{wp}anchor", f"{wp}inline"}:
            continue
        doc_pr = container.find(f"{wp}docPr")
        name = doc_pr.get("name") if doc_pr is not None else None
        if not name:
            continue
        for shape in container.findall(f".//{wps}wsp"):
            names_by_shape[id(shape)] = name

    direct_style_tags = {
        f"{a}noFill",
        f"{a}solidFill",
        f"{a}gradFill",
        f"{a}blipFill",
        f"{a}pattFill",
        f"{a}grpFill",
        f"{a}ln",
        f"{a}effectLst",
        f"{a}effectDag",
        f"{a}scene3d",
        f"{a}sp3d",
    }
    shapes = []
    for index, shape in enumerate(root.findall(f".//{wps}wsp"), start=1):
        shape_properties = shape.find(f"{wps}spPr")
        direct_overrides = []
        if shape_properties is not None:
            direct_overrides = [
                signature(child) for child in shape_properties if child.tag in direct_style_tags
            ]
        style = shape.find(f"{wps}style")
        shapes.append(
            {
                "name": names_by_shape.get(id(shape), f"Shape {index}"),
                "style": signature(style),
                "direct_overrides": tuple(direct_overrides),
            }
        )
    return shapes


def saved_group_root_defect(docx_path: Path) -> Optional[Dict[str, Any]]:
    """Return evidence for the known invalid ``wpg:wGrp`` root.

    This is deliberately a narrow failure diagnostic, not an OOXML
    replacement for Word's Object Model.  It is consulted only after COM
    cannot enumerate shapes.  A successful COM observation therefore
    remains authoritative for Word-tolerated or normalized markup.
    """
    a = f"{{{DRAWINGML_NS}}}"
    invalid_root = f"{{{WORDPROCESSING_GROUP_NS}}}wGrp"
    expected_root = f"{{{WORDPROCESSING_GROUP_NS}}}wgp"

    with zipfile.ZipFile(docx_path) as package:
        root = ET.fromstring(package.read("word/document.xml"))

    defects = []
    group_payload_index = 0
    for graphic_data in root.iter(f"{a}graphicData"):
        if graphic_data.get("uri") != WORDPROCESSING_GROUP_NS:
            continue
        group_payload_index += 1
        root_tags = [child.tag for child in graphic_data]
        if invalid_root in root_tags:
            defects.append(
                {
                    "group_payload_index": group_payload_index,
                    "root_tags": root_tags,
                }
            )

    if not defects:
        return None
    return {
        "document_part": "word/document.xml",
        "detected_element": invalid_root,
        "expected_element": expected_root,
        "defects": defects,
    }


def extract_wrap_polygon_signatures(docx_path: Path) -> List[Dict[str, Any]]:
    """Read Tight/Through wrap paths from the saved DrawingML.

    Word exposes the requested wrap *kind* through COM even when a manually
    authored polygon is degenerate. Retain the saved coordinates for a
    scale-invariant contour projection; Word legitimately serializes the
    same meaningful contour in more than one coordinate extent.
    """
    wp = f"{{{WORDPROCESSING_DRAWING_NS}}}"
    with zipfile.ZipFile(docx_path) as package:
        document_xml = package.read("word/document.xml")
    root = ET.fromstring(document_xml)
    signatures: List[Dict[str, Any]] = []
    for anchor in root.iter(f"{wp}anchor"):
        wrap = next(
            (child for child in anchor if child.tag in {f"{wp}wrapTight", f"{wp}wrapThrough"}),
            None,
        )
        if wrap is None:
            continue
        polygon = wrap.find(f"{wp}wrapPolygon")
        points = []
        if polygon is not None:
            for point in polygon:
                if point.tag not in {f"{wp}start", f"{wp}lineTo"}:
                    continue
                try:
                    points.append((int(point.attrib["x"]), int(point.attrib["y"])))
                except (KeyError, ValueError):
                    points.append((None, None))
        signatures.append(
            {
                "wrap_kind": wrap.tag.rsplit("}", 1)[-1],
                "points": points,
            }
        )
    return signatures


def extract_bevel_render_signatures(docx_path: Path) -> List[Dict[str, Any]]:
    """Read saved bevel and the 3D scene needed to render it visibly."""
    with zipfile.ZipFile(docx_path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    a = f"{{{DRAWINGML_NS}}}"
    container_tags = {
        f"{{{WORDPROCESSING_SHAPE_NS}}}wsp",
        f"{{{PICTURE_NS}}}pic",
    }
    signatures = []
    for container in root.iter():
        if container.tag not in container_tags:
            continue
        bevel = container.find(f".//{a}bevelT")
        if bevel is None:
            continue
        scene = container.find(f".//{a}scene3d")
        camera = scene.find(f"{a}camera") if scene is not None else None
        light_rig = scene.find(f"{a}lightRig") if scene is not None else None
        camera_prst = camera.get("prst", "") if camera is not None else ""
        light_rig_type = light_rig.get("rig", "") if light_rig is not None else ""
        light_rig_direction = light_rig.get("dir", "") if light_rig is not None else ""
        signatures.append(
            {
                "bevel_prst": bevel.get("prst", ""),
                "scene_present": scene is not None,
                "camera_present": camera is not None,
                "light_rig_present": light_rig is not None,
                "camera_prst_present": bool(camera_prst),
                "light_rig_type_present": bool(light_rig_type),
                "light_rig_direction_present": bool(light_rig_direction),
                "camera_prst": camera_prst,
                "light_rig_type": light_rig_type,
                "light_rig_direction": light_rig_direction,
            }
        )
    return signatures


def extract_saved_glow_signatures(docx_path: Path) -> Optional[List[Dict[str, Any]]]:
    """Return all saved DrawingML glow radius/color signatures."""
    a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    try:
        with zipfile.ZipFile(docx_path) as package:
            root = ET.fromstring(package.read("word/document.xml"))
    except (FileNotFoundError, KeyError, OSError, zipfile.BadZipFile, ET.ParseError):
        return None
    signatures = []
    for glow in root.iter(f"{{{a}}}glow"):
        radius = glow.get("rad")
        color = next(iter(glow), None)
        signatures.append(
            {
                "radius_pt": round(float(radius) / 12700, 2) if radius else None,
                "color_type": color.tag.rsplit("}", 1)[-1] if color is not None else None,
                "color_value": color.get("val") if color is not None else None,
            }
        )
    return signatures


def extract_saved_text_vertical_anchors(docx_path: Path) -> Optional[List[str]]:
    wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    try:
        with zipfile.ZipFile(docx_path) as package:
            root = ET.fromstring(package.read("word/document.xml"))
    except (FileNotFoundError, KeyError, OSError, zipfile.BadZipFile, ET.ParseError):
        return None
    return [node.get("anchor", "t") for node in root.iter(f"{{{wps}}}bodyPr")]


def extract_saved_shape_named_effects(
    docx_path: Path,
) -> Dict[str, Dict[str, List[Dict[str, Any]]]]:
    """Project named floating-shape shadow/reflection geometry."""
    a = f"{{{DRAWINGML_NS}}}"
    wp = f"{{{WORDPROCESSING_DRAWING_NS}}}"
    wps = f"{{{WORDPROCESSING_SHAPE_NS}}}"
    with zipfile.ZipFile(docx_path) as package:
        root = ET.fromstring(package.read("word/document.xml"))

    result: Dict[str, Dict[str, List[Dict[str, Any]]]] = {}
    containers = [node for node in root.iter() if node.tag in {f"{wp}anchor", f"{wp}inline"}]
    for index, container in enumerate(containers, start=1):
        shape = container.find(f".//{wps}wsp")
        if shape is None:
            continue
        doc_pr = container.find(f"{wp}docPr")
        name = doc_pr.get("name") if doc_pr is not None and doc_pr.get("name") else f"Shape {index}"
        families: Dict[str, List[Dict[str, Any]]] = {
            "outerShdw": [],
            "reflection": [],
        }
        for family in families:
            for effect in shape.iter(f"{a}{family}"):
                color = next(
                    (
                        child
                        for child in effect
                        if child.tag
                        in {
                            f"{a}prstClr",
                            f"{a}srgbClr",
                            f"{a}schemeClr",
                        }
                    ),
                    None,
                )
                alpha = color.find(f"{a}alpha") if color is not None else None
                families[family].append(
                    {
                        "family": family,
                        "attributes": dict(effect.attrib),
                        "color_type": (color.tag.rsplit("}", 1)[-1] if color is not None else None),
                        "color_value": (color.get("val") if color is not None else None),
                        "alpha": (alpha.get("val") if alpha is not None else None),
                    }
                )
        result[name] = families
    return result
