"""families / references / extract_saved with explicit evidence dependencies."""

from pathlib import Path
from typing import Any, Dict, Optional
import re
import xml.etree.ElementTree as ET
import zipfile


def extract_saved_toc_add_text_state(docx_path: Path, target_text: str) -> Optional[Dict[str, Any]]:
    """Read the saved TOC field/cache and source paragraph without Word COM."""
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ns = {"w": w}
    try:
        with zipfile.ZipFile(docx_path) as package:
            root = ET.fromstring(package.read("word/document.xml"))
            styles_root = (
                ET.fromstring(package.read("word/styles.xml"))
                if "word/styles.xml" in package.namelist()
                else None
            )
    except (FileNotFoundError, KeyError, OSError, zipfile.BadZipFile, ET.ParseError):
        return None

    style_outlines: Dict[str, int] = {}
    style_bases: Dict[str, str] = {}
    if styles_root is not None:
        for style_definition in styles_root.findall("w:style", ns):
            if style_definition.get(f"{{{w}}}type") != "paragraph":
                continue
            style_id = style_definition.get(f"{{{w}}}styleId")
            if not style_id:
                continue
            based_on = style_definition.find("w:basedOn", ns)
            if based_on is not None and based_on.get(f"{{{w}}}val"):
                style_bases[style_id] = based_on.get(f"{{{w}}}val")
            outline = style_definition.find("w:pPr/w:outlineLvl", ns)
            if outline is not None:
                try:
                    style_outlines[style_id] = int(outline.get(f"{{{w}}}val"))
                except (TypeError, ValueError):
                    pass

    def effective_style_outline(style_id: Optional[str]) -> Optional[int]:
        current = style_id
        visited = set()
        while current and current not in visited:
            visited.add(current)
            if current in style_outlines:
                return style_outlines[current]
            current = style_bases.get(current)
        return None

    def effective_paragraph_outline(
        paragraph: ET.Element,
        style_id: Optional[str],
    ) -> Optional[int]:
        direct = paragraph.find("./w:pPr/w:outlineLvl", ns)
        if direct is not None:
            try:
                return int(direct.get(f"{{{w}}}val"))
            except (TypeError, ValueError):
                return None
        return effective_style_outline(style_id)

    toc_field_count = 0
    for instruction in root.findall(".//w:instrText", ns):
        field_code = "".join(instruction.itertext()).strip()
        if re.match(r"^TOC(?:\s|\\)", field_code, flags=re.IGNORECASE):
            toc_field_count += 1

    toc_entries = []
    source_paragraphs = []
    source_context = []
    for paragraph in root.findall(".//w:body//w:p", ns):
        text = "".join(node.text or "" for node in paragraph.findall(".//w:t", ns)).strip()
        style_node = paragraph.find("./w:pPr/w:pStyle", ns)
        style = style_node.get(f"{{{w}}}val") if style_node is not None else None
        if style and re.fullmatch(r"TOC\d+", style, flags=re.IGNORECASE):
            toc_entries.append(text)
            continue
        if style and re.fullmatch(r"TOC(?:Heading|\s+Heading)", style, flags=re.IGNORECASE):
            continue
        if text:
            source_context.append(
                {
                    "text": " ".join(text.split()),
                    "paragraph_style": style,
                    "outline_level": effective_paragraph_outline(paragraph, style),
                }
            )
        if text != target_text:
            continue
        size_node = paragraph.find(".//w:rPr/w:sz", ns)
        size = size_node.get(f"{{{w}}}val") if size_node is not None else None
        source_paragraphs.append(
            {
                "paragraph_style": style,
                "outline_level": effective_paragraph_outline(paragraph, style),
                "font_size": float(size) / 2 if size is not None else None,
            }
        )

    return {
        "toc_count": toc_field_count,
        "toc_entries": toc_entries,
        "target_in_toc": any(target_text in entry for entry in toc_entries),
        "source_paragraphs": source_paragraphs,
        "source_context": source_context,
    }


def extract_saved_manual_toc_state(path: Path) -> Optional[Dict[str, Any]]:
    """Read manual-TOC outer/inner SDT identity from submitted OOXML."""
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ns = {"w": w}
    try:
        with zipfile.ZipFile(path) as package:
            root = ET.fromstring(package.read("word/document.xml"))
    except (OSError, KeyError, zipfile.BadZipFile, ET.ParseError):
        return None

    def attribute(node: Optional[ET.Element], name: str, default=""):
        if node is None:
            return default
        return node.get(f"{{{w}}}{name}", node.get(name, default))

    controls = root.findall(".//w:sdt", ns)
    outers = []
    for control in controls:
        properties = control.find("w:sdtPr", ns)
        doc_part = properties.find("w:docPartObj", ns) if properties is not None else None
        gallery = attribute(
            doc_part.find("w:docPartGallery", ns) if doc_part is not None else None,
            "val",
        )
        if "table of contents" in gallery.casefold():
            outers.append(control)
    if len(outers) == 1:
        outer = outers[0]
        inner_controls = [
            control for control in outer.findall(".//w:sdt", ns) if control is not outer
        ]
        manual_paragraphs = outer.findall(".//w:p", ns)
    else:
        inner_controls = []
        manual_paragraphs = []
    placeholder_count = 0
    temporary_count = 0
    for control in inner_controls:
        properties = control.find("w:sdtPr", ns)
        if properties is not None:
            placeholder_count += int(properties.find("w:showingPlcHdr", ns) is not None)
            temporary_count += int(properties.find("w:temporary", ns) is not None)
    heading_count = 0
    entry_levels = []
    for paragraph in manual_paragraphs:
        style_node = paragraph.find("./w:pPr/w:pStyle", ns)
        style = attribute(style_node, "val")
        if style.casefold() in {"tocheading", "toc heading"}:
            heading_count += 1
        match = re.fullmatch(r"TOC\s*(\d+)", style, re.I)
        if match:
            entry_levels.append(int(match.group(1)))
    toc_fields = [
        " ".join((node.text or "").split())
        for node in root.findall(".//w:instrText", ns)
        if re.search(r"^\s*TOC\b", node.text or "", re.I)
    ]
    return {
        "outer_count": len(outers),
        "inner_count": len(inner_controls),
        "placeholder_count": placeholder_count,
        "temporary_count": temporary_count,
        "heading_count": heading_count,
        "entry_levels": entry_levels,
        "toc_field_count": len(toc_fields),
    }
