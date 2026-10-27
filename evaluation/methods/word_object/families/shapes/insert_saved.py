"""families / shapes / insert_saved with explicit evidence dependencies."""

from pathlib import Path
from typing import Any, Dict, List, Optional
import io
import math
import re
import xml.etree.ElementTree as ET
import zipfile


def extract_saved_textboxes(docx_path: Path, *, canonical_saved_element) -> List[Dict[str, Any]]:
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    ns = {"w": w, "wp": wp, "wps": wps, "a": a}
    with zipfile.ZipFile(docx_path) as package:
        root = ET.fromstring(package.read("word/document.xml"))

    margin_left = margin_top = 72.0
    section = root.find(".//w:sectPr", ns)
    if section is not None:
        margin = section.find("w:pgMar", ns)
        if margin is not None:
            margin_left = float(margin.get(f"{{{w}}}left", 1440)) / 20.0
            margin_top = float(margin.get(f"{{{w}}}top", 1440)) / 20.0

    def position(container: ET.Element, axis: str) -> tuple[Optional[float], list[str]]:
        errors = []
        node = container.find(f"wp:position{axis}", ns)
        if node is None:
            return None, [f"missing position{axis}"]
        relative = node.get("relativeFrom")
        offset = node.find("wp:posOffset", ns)
        if relative is None:
            errors.append(f"position{axis} missing unqualified relativeFrom")
        if offset is None or offset.text is None:
            align = node.find("wp:align", ns)
            if align is not None and align.text:
                return None, errors
            errors.append(f"position{axis} missing posOffset")
            return None, errors
        raw = float(offset.text) / 12700.0
        if axis == "H" and relative in {"column", "margin"}:
            raw += margin_left
        if axis == "V" and relative in {"paragraph", "margin"}:
            raw += margin_top
        return round(raw, 4), errors

    signatures = []
    for container in root.iter():
        if container.tag not in {f"{{{wp}}}anchor", f"{{{wp}}}inline"}:
            continue
        shape = container.find(".//wps:wsp", ns)
        textbox = container.find(".//w:txbxContent", ns)
        if shape is None or textbox is None:
            continue
        errors = []
        left, horizontal_errors = position(container, "H")
        top, vertical_errors = position(container, "V")
        errors.extend(horizontal_errors + vertical_errors)
        extent = container.find("wp:extent", ns)
        if extent is None:
            errors.append("missing extent")
            width = height = None
        else:
            width = round(float(extent.get("cx", 0)) / 12700.0, 4)
            height = round(float(extent.get("cy", 0)) / 12700.0, 4)
        body_pr = shape.find("wps:bodyPr", ns)
        style = shape.find("wps:style", ns)
        geometry = shape.find("wps:spPr/a:prstGeom", ns)
        paragraph_styles = []
        for paragraph in textbox.findall(".//w:p", ns):
            pstyle = paragraph.find("w:pPr/w:pStyle", ns)
            paragraph_styles.append(pstyle.get(f"{{{w}}}val") if pstyle is not None else "")
        signatures.append(
            {
                "left": left,
                "top": top,
                "width": width,
                "height": height,
                "orientation": body_pr.get("vert", "horz") if body_pr is not None else "horz",
                "body_properties": (
                    {
                        key: body_pr.get(key)
                        for key in (
                            "vert",
                            "wrap",
                            "lIns",
                            "tIns",
                            "rIns",
                            "bIns",
                            "numCol",
                            "anchor",
                            "anchorCtr",
                        )
                        if body_pr.get(key) is not None
                    }
                    if body_pr is not None
                    else {}
                ),
                "geometry": geometry.get("prst") if geometry is not None else None,
                "text": " ".join(
                    "".join(node.text or "" for node in textbox.findall(".//w:t", ns)).split()
                ),
                "paragraph_styles": paragraph_styles,
                "style": canonical_saved_element(style),
                "errors": errors,
            }
        )
    return signatures


def extract_saved_autoshape_geometries(docx_path: Path) -> List[str]:
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    wpg = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
    a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    mc = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    v = "urn:schemas-microsoft-com:vml"
    o = "urn:schemas-microsoft-com:office:office"
    ns = {"wp": wp, "wps": wps, "a": a}
    with zipfile.ZipFile(docx_path) as package:
        document_xml = package.read("word/document.xml")
    root = ET.fromstring(document_xml)
    namespace_map = {
        prefix: uri
        for _, (prefix, uri) in ET.iterparse(io.BytesIO(document_xml), events=("start-ns",))
    }
    supported_choice_namespaces = {w, wp, wps, wpg, a, v}
    geometries = []

    def positive_vml_dimension(value: str) -> bool:
        match = re.fullmatch(
            r"\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+))\s*" r"(?:pt|pc|in|cm|mm|px|%|em|ex)?\s*",
            value,
            flags=re.IGNORECASE,
        )
        if not match:
            return False
        numeric = float(match.group(1))
        return math.isfinite(numeric) and numeric > 0

    def positive_vml_opacity(value: Optional[str]) -> bool:
        if value is None:
            return True
        match = re.fullmatch(
            r"\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+))\s*(%|f)?\s*",
            value,
            flags=re.IGNORECASE,
        )
        if not match:
            return False
        numeric = float(match.group(1))
        return math.isfinite(numeric) and numeric > 0

    def visible_vml_rectangle(node: ET.Element) -> bool:
        style = {}
        for declaration in node.get("style", "").split(";"):
            if ":" not in declaration:
                continue
            name, value = declaration.split(":", 1)
            style[name.strip().casefold()] = value.strip().casefold()
        if not (
            positive_vml_dimension(style.get("width", ""))
            and positive_vml_dimension(style.get("height", ""))
        ):
            return False
        if style.get("display") == "none":
            return False
        if style.get("visibility") in {"hidden", "collapse"}:
            return False
        if style.get("mso-hide") in {"all", "screen"}:
            return False
        opacity = style.get("opacity")
        if not positive_vml_opacity(opacity):
            return False

        def enabled(attribute: Optional[str]) -> bool:
            return (attribute or "").strip().casefold() not in {
                "0",
                "false",
                "f",
                "off",
                "no",
            }

        fill = node.find(f"{{{v}}}fill")
        stroke = node.find(f"{{{v}}}stroke")
        fill_setting = fill.get("on") if fill is not None else None
        stroke_setting = stroke.get("on") if stroke is not None else None
        fill_opacity = fill.get("opacity") if fill is not None else None
        fill_opacity2 = (
            fill.get("opacity2") or fill.get(f"{{{o}}}opacity2") if fill is not None else None
        )
        fill_type = fill.get("type", "").strip().casefold() if fill is not None else ""
        has_visible_fill_alpha = positive_vml_opacity(fill_opacity)
        if fill_opacity2 is not None and fill_type in {"gradient", "gradientradial", "pattern"}:
            has_visible_fill_alpha = has_visible_fill_alpha or positive_vml_opacity(fill_opacity2)
        has_fill = enabled(fill_setting if fill_setting is not None else node.get("filled"))
        has_fill = (
            has_fill
            and style.get("fill") != "none"
            and positive_vml_opacity(style.get("fill-opacity"))
            and has_visible_fill_alpha
        )
        has_stroke = enabled(stroke_setting if stroke_setting is not None else node.get("stroked"))
        stroke_weight = (
            (stroke.get("weight") if stroke is not None else None)
            or node.get("strokeweight")
            or style.get("strokeweight")
            or style.get("stroke-width")
        )
        has_stroke = (
            has_stroke
            and style.get("stroke") != "none"
            and positive_vml_opacity(style.get("stroke-opacity"))
            and positive_vml_opacity(stroke.get("opacity") if stroke is not None else None)
            and (stroke_weight is None or positive_vml_dimension(stroke_weight))
        )
        return has_fill or has_stroke

    def choice_is_supported(choice: ET.Element) -> bool:
        requires = (choice.get("Requires") or choice.get(f"{{{mc}}}Requires") or "").split()
        return bool(requires) and all(
            namespace_map.get(prefix) in supported_choice_namespaces for prefix in requires
        )

    def visit(
        node: ET.Element,
        *,
        in_body: bool = False,
        in_run: bool = False,
        in_pict: bool = False,
    ) -> None:
        if node.tag == f"{{{mc}}}AlternateContent":
            # AlternateContent represents one logical object.  Prefer the
            # modern Choice branch and inspect Fallback only when no
            # Choice exists, so a DrawingML/VML compatibility pair is not
            # counted twice.
            choice = next(
                (
                    child
                    for child in node
                    if child.tag == f"{{{mc}}}Choice" and choice_is_supported(child)
                ),
                None,
            )
            fallback = next(
                (child for child in node if child.tag == f"{{{mc}}}Fallback"),
                None,
            )
            selected = choice if choice is not None else fallback
            if selected is not None:
                for child in selected:
                    visit(
                        child,
                        in_body=in_body,
                        in_run=in_run,
                        in_pict=in_pict,
                    )
            return

        in_body = in_body or node.tag == f"{{{w}}}body"
        in_run = in_run or node.tag == f"{{{w}}}r"
        in_pict = in_pict or node.tag == f"{{{w}}}pict"
        if node.tag in {f"{{{wp}}}anchor", f"{{{wp}}}inline"}:
            geometry = node.find(".//wps:wsp/wps:spPr/a:prstGeom", ns)
            if geometry is not None:
                geometries.append(geometry.get("prst"))
        elif (
            node.tag == f"{{{v}}}rect"
            and in_body
            and in_run
            and in_pict
            and visible_vml_rectangle(node)
        ):
            # Legacy VML rectangles are Word autoshapes too.  Do not
            # generalize this to v:shape or v:oval: their geometry cannot
            # be inferred as a rectangle from the element alone.  The
            # rectangle must also be a visible, non-degenerate Word object;
            # orphan cache markup is not saved-shape evidence.
            geometries.append("rect")

        for child in node:
            visit(
                child,
                in_body=in_body,
                in_run=in_run,
                in_pict=in_pict,
            )

    visit(root)
    return sorted(geometries)
