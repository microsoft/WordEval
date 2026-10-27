"""families / pages / extract_saved with explicit evidence dependencies."""

from evaluation.methods.word_object.ooxml import canonical as ooxml_canonical
from pathlib import Path
from typing import Any, Dict
import re
import xml.etree.ElementTree as ET
import zipfile


def cover_page_ooxml_layout_signature(docx_path: Path) -> Dict[str, Any]:
    """Extract stable saved-layout signals without resolving style names.

    Cover-page predictions can preserve the same direct formatting and
    ``w:pStyle`` role token while omitting an otherwise unused style
    definition. Word then reports a different ``Style.NameLocal`` even
    though the rendered building block is unchanged. Compare the saved
    paragraph roles and drawing geometry instead, ignoring empty trailing
    paragraphs that have no visible layout effect.
    """
    namespaces = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
        "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        "v": "urn:schemas-microsoft-com:vml",
    }
    with zipfile.ZipFile(docx_path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
        picture_signatures = ooxml_canonical.saved_picture_signatures(root, package)

    def q(prefix: str, local: str) -> str:
        return f"{{{namespaces[prefix]}}}{local}"

    paragraph_roles = []
    for paragraph in root.findall(".//w:body//w:p", namespaces):
        text = "".join(node.text or "" for node in paragraph.findall(".//w:t", namespaces))
        has_visible_object = (
            paragraph.find(".//w:drawing", namespaces) is not None
            or paragraph.find(".//w:pict", namespaces) is not None
        )
        if not text.strip() and not has_visible_object:
            continue
        style = paragraph.find("./w:pPr/w:pStyle", namespaces)
        paragraph_roles.append(style.get(q("w", "val"), "") if style is not None else "")

    def position_signature(anchor: ET.Element, name: str):
        position = anchor.find(f"wp:{name}", namespaces)
        if position is None:
            return None
        child = next(iter(position), None)
        return (
            position.get("relativeFrom", ""),
            child.tag.rsplit("}", 1)[-1] if child is not None else "",
            (child.text or "").strip() if child is not None else "",
        )

    anchors = []
    for anchor in root.findall(".//wp:anchor", namespaces):
        extent = anchor.find("wp:extent", namespaces)
        geometry = anchor.find(".//a:prstGeom", namespaces)
        wrap = next(
            (
                child.tag.rsplit("}", 1)[-1]
                for child in anchor
                if child.tag.rsplit("}", 1)[-1].startswith("wrap")
            ),
            "",
        )
        anchors.append(
            (
                int(extent.get("cx", "0")) if extent is not None else 0,
                int(extent.get("cy", "0")) if extent is not None else 0,
                position_signature(anchor, "positionH"),
                position_signature(anchor, "positionV"),
                geometry.get("prst", "") if geometry is not None else "",
                wrap,
            )
        )

    paragraph_formatting = ooxml_canonical.saved_body_visible_format_signature(root)
    drawing_namespaces = {
        namespaces["a"],
        namespaces["pic"],
        namespaces["wps"],
    }
    drawing_visuals = []
    for node in root.iter():
        namespace = node.tag[1:].split("}", 1)[0] if node.tag.startswith("{") else ""
        local = node.tag.rsplit("}", 1)[-1]
        if namespace in drawing_namespaces and local in {
            "spPr",
            "grpSpPr",
            "style",
        }:
            drawing_visuals.append(
                ooxml_canonical.canonical_visible_element(
                    node,
                    skipped_attribute_names={
                        "id",
                        "spid",
                        "anchorId",
                        "editId",
                        "objectId",
                    },
                )
            )

    vml_visuals = []
    vml_child_names = {
        "stroke",
        "path",
        "shadow",
        "textpath",
        "formulas",
        "handles",
        "wrap",
        "textbox",
    }
    vml_nodes = [
        node
        for node in root.iter()
        if node.tag
        in {
            f"{{{namespaces['v']}}}shape",
            f"{{{namespaces['v']}}}group",
        }
    ]
    z_values = []
    for node in vml_nodes:
        match = re.search(r"(?:^|;)z-index:(-?\d+)", node.get("style", ""))
        if match:
            z_values.append(int(match.group(1)))
    z_ranks = {value: rank for rank, value in enumerate(sorted(set(z_values)))}

    def normalize_vml_value(name: str, value: str) -> str:
        local = name.rsplit("}", 1)[-1]
        if local in {"fillcolor", "strokecolor", "color", "color2"}:
            theme_match = re.search(r"(\[\d+\])\s*$", value)
            if theme_match:
                return theme_match.group(1)
        if local == "style":
            return re.sub(
                r"(^|;)z-index:(-?\d+)",
                lambda match: (f"{match.group(1)}z-index:" f"{z_ranks[int(match.group(2))]}"),
                value,
            )
        return value

    for node in vml_nodes:
        attributes = tuple(
            sorted(
                (name, normalize_vml_value(name, value))
                for name, value in node.attrib.items()
                if name.rsplit("}", 1)[-1]
                not in {
                    "id",
                    "spid",
                    "anchorId",
                    "editId",
                    "objectId",
                    "gfxdata",
                }
            )
        )
        children = []
        for child in node:
            local = child.tag.rsplit("}", 1)[-1]
            if local not in vml_child_names:
                continue
            if local == "textbox":
                children.append(
                    (
                        child.tag,
                        tuple(
                            sorted(
                                (name, value)
                                for name, value in child.attrib.items()
                                if name.rsplit("}", 1)[-1]
                                not in {
                                    "id",
                                    "spid",
                                    "anchorId",
                                    "editId",
                                    "objectId",
                                    "gfxdata",
                                }
                            )
                        ),
                    )
                )
                continue
            children.append(
                ooxml_canonical.canonical_visible_element(
                    child,
                    skipped_attribute_names={
                        "id",
                        "spid",
                        "anchorId",
                        "editId",
                        "objectId",
                        "gfxdata",
                    },
                )
            )
        vml_visuals.append((node.tag, attributes, tuple(children)))

    section = root.find(".//w:body/w:sectPr", namespaces)
    page_size = section.find("w:pgSz", namespaces) if section is not None else None
    page_margin = section.find("w:pgMar", namespaces) if section is not None else None
    return {
        "paragraph_roles": paragraph_roles,
        "anchors": sorted(anchors, key=repr),
        "vml_shape_count": len(root.findall(".//v:shape", namespaces)),
        "vml_group_count": len(root.findall(".//v:group", namespaces)),
        "paragraph_formatting": paragraph_formatting,
        "drawing_visuals": drawing_visuals,
        "vml_visuals": vml_visuals,
        "pictures": picture_signatures,
        "page_size": dict(sorted(page_size.attrib.items())) if page_size is not None else {},
        "page_margin": dict(sorted(page_margin.attrib.items())) if page_margin is not None else {},
    }
