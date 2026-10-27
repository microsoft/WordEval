"""families / formatting / extract_saved with explicit evidence dependencies."""

from .text import (
    _adjacent_text_run_groups,
    _matching_adjacent_run_signatures,
    _xml_element_signature,
)
from pathlib import Path
from typing import Any, Dict, List, Optional
import xml.etree.ElementTree as ET
import zipfile


def extract_saved_ligature_signatures(docx_path: Path) -> List[Dict[str, Any]]:
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    w14 = "http://schemas.microsoft.com/office/word/2010/wordml"
    ns = {"w": w, "w14": w14}
    with zipfile.ZipFile(docx_path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    signatures = []
    for paragraph in root.findall(".//w:body/w:p", ns):
        values = []
        paragraph_ligature = paragraph.find("w:pPr/w:rPr/w14:ligatures", ns)
        paragraph_value = (
            paragraph_ligature.get(f"{{{w14}}}val") if paragraph_ligature is not None else None
        )
        for run in paragraph.findall("w:r", ns):
            if not "".join(node.text or "" for node in run.findall(".//w:t", ns)):
                continue
            ligature = run.find("w:rPr/w14:ligatures", ns)
            values.append(
                ligature.get(f"{{{w14}}}val") if ligature is not None else paragraph_value
            )
        signatures.append(
            {
                "text": "".join(node.text or "" for node in paragraph.findall(".//w:t", ns)),
                "ligatures": values,
            }
        )
    return signatures


def extract_saved_target_run(docx_path: Path, target_text: str) -> Dict[str, Any]:
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    w14 = "http://schemas.microsoft.com/office/word/2010/wordml"
    ns = {"w": w, "w14": w14}
    with zipfile.ZipFile(docx_path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    full_text = "".join(node.text or "" for node in root.findall(".//w:body//w:t", ns))

    def run_signature(run: ET.Element) -> Dict[str, Any]:
        text = "".join(node.text or "" for node in run.findall(".//w:t", ns))
        rpr = run.find("w:rPr", ns)

        def child(local: str, namespace: str = w) -> Optional[ET.Element]:
            return rpr.find(f"{{{namespace}}}{local}") if rpr is not None else None

        def on_off(local: str) -> bool:
            node = child(local)
            if node is None:
                return False
            return node.get(f"{{{w}}}val", "1").lower() not in {"0", "false", "off"}

        fonts = child("rFonts")
        color = child("color")
        size = child("sz")
        underline = child("u")
        border = child("bdr")
        shading = child("shd")
        outline = child("textOutline", w14)
        reflection = child("reflection", w14)
        return {
            "text": text,
            "_merge_run_properties": _xml_element_signature(rpr),
            "bold": on_off("b"),
            "italic": on_off("i"),
            "underline": underline.get(f"{{{w}}}val") if underline is not None else None,
            "border": (border.get(f"{{{w}}}val", "single") if border is not None else None),
            "shading": (tuple(sorted(shading.attrib.items())) if shading is not None else None),
            "font_name": (
                fonts.get(f"{{{w}}}ascii") or fonts.get(f"{{{w}}}hAnsi")
                if fonts is not None
                else None
            ),
            "font_size": (float(size.get(f"{{{w}}}val")) / 2.0 if size is not None else None),
            "font_color": color.get(f"{{{w}}}val") if color is not None else None,
            "outline_visible": outline is not None,
            "outline_weight": (
                float(outline.get(f"{{{w14}}}w", 0)) / 12700.0 if outline is not None else None
            ),
            "reflection_present": reflection is not None,
            "reflection_signature": (
                tuple(sorted(reflection.attrib.items())) if reflection is not None else None
            ),
        }

    runs = []
    for adjacent in _adjacent_text_run_groups(root, w):
        signatures = [run_signature(run) for run in adjacent]
        runs.extend(
            _matching_adjacent_run_signatures(
                signatures,
                target_text,
                strip_text=False,
            )
        )
    return {"full_text": full_text, "target_runs": runs}


def extract_saved_target_text_effect(
    docx_path: Path, target_text: str, idMso: str
) -> Dict[str, Any]:
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    w14 = "http://schemas.microsoft.com/office/word/2010/wordml"
    ns = {"w": w, "w14": w14}
    with zipfile.ZipFile(docx_path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    signatures = []

    def gradient_color_signature(color: Optional[ET.Element]) -> Any:
        if color is None:
            return None
        local_name = color.tag.rsplit("}", 1)[-1]
        value = (color.get(f"{{{w14}}}val") or "").upper()
        transforms = tuple(
            (
                child.tag.rsplit("}", 1)[-1],
                tuple(sorted(child.attrib.items())),
            )
            for child in color
        )
        if not transforms:
            if local_name == "srgbClr" and value in {"EE0000", "FF0000"}:
                return "named_red"
            if local_name == "srgbClr" and value == "000000":
                return "named_black"
            if local_name == "schemeClr" and value.casefold() in {"tx1", "dk1"}:
                return "named_black"
        return (local_name, value, transforms)

    def run_signature(run: ET.Element) -> Dict[str, Any]:
        text = "".join(node.text or "" for node in run.findall(".//w:t", ns))
        rpr = run.find("w:rPr", ns)
        if idMso == "TextHighlightColorPicker":
            node = rpr.find("w:highlight", ns) if rpr is not None else None
            signature = node.get(f"{{{w}}}val") if node is not None else None
        elif idMso == "TextOutlineDashesGallery":
            outline = rpr.find("w14:textOutline", ns) if rpr is not None else None
            dash = outline.find("w14:prstDash", ns) if outline is not None else None
            width = outline.get(f"{{{w14}}}w") if outline is not None else None
            try:
                positive_width = int(width or "0") > 0
            except ValueError:
                positive_width = False
            signature = {
                "outline_present": outline is not None,
                "outline_visible": (
                    outline is not None
                    and positive_width
                    and outline.find("w14:noFill", ns) is None
                ),
                "dash": dash.get(f"{{{w14}}}val") if dash is not None else None,
            }
        else:
            text_fill = rpr.find("w14:textFill", ns) if rpr is not None else None
            gradient = text_fill.find("w14:gradFill", ns) if text_fill is not None else None
            stops = []
            if gradient is not None:
                for stop in gradient.findall("w14:gsLst/w14:gs", ns):
                    color = next(iter(stop), None)
                    stops.append(
                        {
                            "position": stop.get(f"{{{w14}}}pos"),
                            "color": gradient_color_signature(color),
                        }
                    )
            linear = gradient.find("w14:lin", ns) if gradient is not None else None
            signature = {
                "gradient_present": gradient is not None,
                "stops": stops,
                "angle": linear.get(f"{{{w14}}}ang") if linear is not None else None,
            }
        return {
            "text": text,
            "effect": signature,
            "_merge_run_properties": _xml_element_signature(rpr),
        }

    for adjacent in _adjacent_text_run_groups(root, w):
        run_signatures = [run_signature(run) for run in adjacent]
        signatures.extend(
            _matching_adjacent_run_signatures(
                run_signatures,
                target_text,
                strip_text=True,
            )
        )
    return {"target_runs": signatures}


def extract_saved_multilevel_list(docx_path: Path) -> Dict[str, Any]:
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ns = {"w": w}
    with zipfile.ZipFile(docx_path) as package:
        document = ET.fromstring(package.read("word/document.xml"))
        numbering = (
            ET.fromstring(package.read("word/numbering.xml"))
            if "word/numbering.xml" in package.namelist()
            else None
        )

        styles = None
        if "word/styles.xml" in package.namelist():
            styles = ET.fromstring(package.read("word/styles.xml"))

    def num_pr_values(num_pr: Optional[ET.Element]) -> Dict[str, str]:
        values = {}
        if num_pr is None:
            return values
        for child_name, key in (("numId", "num_id"), ("ilvl", "level")):
            child = num_pr.find(f"w:{child_name}", ns)
            if child is not None:
                value = child.get(f"{{{w}}}val")
                if value is not None:
                    values[key] = value
        return values

    style_definitions: Dict[str, Dict[str, Any]] = {}
    if styles is not None:
        for style in styles.findall("w:style", ns):
            style_id = style.get(f"{{{w}}}styleId")
            if not style_id:
                continue
            based_on = style.find("w:basedOn", ns)
            style_definitions[style_id] = {
                "based_on": (based_on.get(f"{{{w}}}val") if based_on is not None else None),
                "numbering": num_pr_values(style.find("w:pPr/w:numPr", ns)),
            }

    resolved_styles: Dict[str, Dict[str, str]] = {}

    def resolve_style_numbering(
        style_id: Optional[str],
        active: Optional[set[str]] = None,
    ) -> Dict[str, str]:
        if not style_id:
            return {}
        if style_id in resolved_styles:
            return dict(resolved_styles[style_id])
        active = set(active or ())
        if style_id in active:
            return {}
        active.add(style_id)
        definition = style_definitions.get(style_id, {})
        values = resolve_style_numbering(definition.get("based_on"), active)
        values.update(definition.get("numbering") or {})
        resolved_styles[style_id] = values
        return dict(values)

    num_to_abstract = {}
    for num in numbering.findall("w:num", ns) if numbering is not None else ():
        num_id = num.get(f"{{{w}}}numId")
        abstract = num.find("w:abstractNumId", ns)
        num_to_abstract[num_id] = abstract.get(f"{{{w}}}val") if abstract is not None else None

    paragraphs = []
    used_num_ids = set()
    for paragraph in document.findall(".//w:body/w:p", ns):
        paragraph_properties = paragraph.find("w:pPr", ns)
        style = (
            paragraph_properties.find("w:pStyle", ns) if paragraph_properties is not None else None
        )
        style_id = style.get(f"{{{w}}}val") if style is not None else None
        numbering_values = resolve_style_numbering(style_id)
        numbering_values.update(
            num_pr_values(
                paragraph_properties.find("w:numPr", ns)
                if paragraph_properties is not None
                else None
            )
        )
        num_id_value = numbering_values.get("num_id")
        level_value = numbering_values.get("level")
        if num_id_value is not None and level_value is None:
            # Word treats an omitted ilvl on a numbered paragraph/style as
            # the first list level.  Heading-linked outline lists commonly
            # persist level 0 this way.
            level_value = "0"
        if num_id_value is not None:
            used_num_ids.add(num_id_value)
        paragraphs.append(
            {
                "text": "".join(node.text or "" for node in paragraph.findall(".//w:t", ns)),
                "level": level_value,
                "num_id": num_id_value,
            }
        )

    levels = {}
    invalid = [] if numbering is not None else ["missing word/numbering.xml"]
    used_abstract = {num_to_abstract.get(num_id) for num_id in used_num_ids}
    used_levels_by_abstract: Dict[Optional[str], set[str]] = {}
    for paragraph in paragraphs:
        abstract_id = num_to_abstract.get(paragraph["num_id"])
        level_value = paragraph["level"]
        if abstract_id is not None and level_value is not None:
            used_levels_by_abstract.setdefault(abstract_id, set()).add(level_value)
    for abstract in (numbering.findall("w:abstractNum", ns) if numbering is not None else ()):
        abstract_id = abstract.get(f"{{{w}}}abstractNumId")
        if abstract_id not in used_abstract:
            continue
        used_levels = used_levels_by_abstract.get(abstract_id, set())
        for level in abstract.findall("w:lvl", ns):
            ilvl = level.get(f"{{{w}}}ilvl")
            if ilvl is None:
                invalid.append("w:lvl missing w:ilvl")
                continue
            if ilvl not in used_levels:
                # A reference may define the full nine-level gallery while
                # a correct prediction persists only the levels actually
                # used by document paragraphs.  Unused definitions are not
                # observable requested output.
                continue
            num_fmt = level.find("w:numFmt", ns)
            level_text = level.find("w:lvlText", ns)
            levels[ilvl] = {
                "format": num_fmt.get(f"{{{w}}}val") if num_fmt is not None else None,
                "text": level_text.get(f"{{{w}}}val") if level_text is not None else None,
            }
    normalized_paragraphs = [{"text": row["text"], "level": row["level"]} for row in paragraphs]
    return {
        "paragraphs": normalized_paragraphs,
        "levels": levels,
        "invalid": invalid,
    }


def extract_quick_style_signatures(docx_path: Path) -> List[Dict[str, Any]]:
    """Read paragraph style roles and direct formatting from document.xml."""
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    namespaces = {"w": w}
    with zipfile.ZipFile(docx_path) as package:
        root = ET.fromstring(package.read("word/document.xml"))

    def canonical(element: Optional[ET.Element]) -> Any:
        if element is None:
            return None
        attributes = tuple(
            sorted(
                (name, value)
                for name, value in element.attrib.items()
                if not name.rsplit("}", 1)[-1].startswith("rsid")
            )
        )
        return (
            element.tag,
            attributes,
            (element.text or "").strip(),
            tuple(canonical(child) for child in element),
        )

    signatures = []
    for paragraph in root.findall(".//w:body/w:p", namespaces):
        paragraph_properties = paragraph.find("w:pPr", namespaces)
        style = (
            paragraph_properties.find("w:pStyle", namespaces)
            if paragraph_properties is not None
            else None
        )
        paragraph_direct = []
        if paragraph_properties is not None:
            paragraph_direct = [
                canonical(child)
                for child in paragraph_properties
                if child.tag
                not in {
                    f"{{{w}}}pStyle",
                    # pPr/rPr formats the paragraph mark, not the visible
                    # text run.  Treating it as a visible direct override
                    # falsely rejects style-correct documents.
                    f"{{{w}}}rPr",
                }
            ]
        run_direct = []
        for run in paragraph.findall("w:r", namespaces):
            run_text = "".join(node.text or "" for node in run.findall(".//w:t", namespaces))
            if not run_text:
                continue
            run_properties = canonical(run.find("w:rPr", namespaces))
            if (
                run_properties is not None
                and not run_properties[1]
                and not run_properties[2]
                and not run_properties[3]
            ):
                # <w:rPr/> is equivalent to no direct run formatting.
                run_properties = None
            run_direct.append(run_properties)
        signatures.append(
            {
                "text": "".join(
                    node.text or "" for node in paragraph.findall(".//w:t", namespaces)
                ),
                "style_id": (style.get(f"{{{w}}}val", "") if style is not None else ""),
                "paragraph_direct_formatting": paragraph_direct,
                "run_direct_formatting": run_direct,
            }
        )
    return signatures
