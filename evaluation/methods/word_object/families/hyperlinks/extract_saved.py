"""families / hyperlinks / extract_saved with explicit evidence dependencies."""

from evaluation.methods.word_object.ooxml import canonical as ooxml_canonical
from pathlib import Path
from typing import Any, Dict, Optional
import xml.etree.ElementTree as ET
import zipfile


def saved_hyperlink_guardrail_signature(docx_path: Path) -> Dict[str, Any]:
    """Extract visible formatting and picture-preservation signals."""
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ns = {"w": w}
    with zipfile.ZipFile(docx_path) as package:
        names = set(package.namelist())
        root = ET.fromstring(package.read("word/document.xml"))
        styles = (
            ET.fromstring(package.read("word/styles.xml")) if "word/styles.xml" in names else None
        )
        pictures = ooxml_canonical.saved_picture_signatures(root, package)

    style_by_id = {}
    if styles is not None:
        style_by_id = {
            style.get(f"{{{w}}}styleId"): style
            for style in styles.findall("w:style", ns)
            if style.get(f"{{{w}}}styleId")
        }

    def style_properties(style_id: str, seen: Optional[set[str]] = None) -> Dict[str, Any]:
        if not style_id or style_id not in style_by_id:
            return {
                "builtin_hyperlink": style_id.lower() == "hyperlink",
                "color": None,
                "underline": None,
            }
        seen = set(seen or set())
        if style_id in seen:
            return {"builtin_hyperlink": False, "color": None, "underline": None}
        seen.add(style_id)
        style = style_by_id[style_id]
        based_on = style.find("w:basedOn", ns)
        base_id = based_on.get(f"{{{w}}}val", "") if based_on is not None else ""
        result = style_properties(base_id, seen)
        name = style.find("w:name", ns)
        style_name = name.get(f"{{{w}}}val", "") if name is not None else ""
        result["builtin_hyperlink"] = bool(
            result["builtin_hyperlink"]
            or style_id.lower() == "hyperlink"
            or style_name.lower() == "hyperlink"
        )
        rpr = style.find("w:rPr", ns)
        if rpr is not None:
            style_color = rpr.find("w:color", ns)
            style_underline = rpr.find("w:u", ns)
            if style_color is not None:
                result["color"] = style_color
            if style_underline is not None:
                result["underline"] = style_underline
        return result

    def standard_color(color: Optional[ET.Element]) -> bool:
        if color is None:
            return False
        theme = (color.get(f"{{{w}}}themeColor") or "").lower()
        value = (color.get(f"{{{w}}}val") or "").lower()
        return theme == "hyperlink" or value in {"0000ff", "0563c1"}

    def visible_underline(underline: Optional[ET.Element]) -> bool:
        if underline is None:
            return False
        return (underline.get(f"{{{w}}}val") or "single").lower() not in {
            "none",
            "0",
            "false",
            "off",
        }

    def run_uses_standard_hyperlink_style(run: ET.Element) -> bool:
        rpr = run.find("w:rPr", ns)
        if rpr is None:
            return False
        rstyle = rpr.find("w:rStyle", ns)
        style_id = rstyle.get(f"{{{w}}}val", "") if rstyle is not None else ""
        inherited = style_properties(style_id)
        direct_color = rpr.find("w:color", ns)
        direct_underline = rpr.find("w:u", ns)

        if direct_color is not None and not standard_color(direct_color):
            return False
        if direct_underline is not None and not visible_underline(direct_underline):
            return False
        if inherited["builtin_hyperlink"]:
            return True
        effective_color = direct_color if direct_color is not None else inherited["color"]
        effective_underline = (
            direct_underline if direct_underline is not None else inherited["underline"]
        )
        return standard_color(effective_color) and visible_underline(effective_underline)

    standard_range_styles = []
    for hyperlink in root.findall(".//w:body//w:hyperlink", ns):
        visible_runs = [
            run
            for run in hyperlink.findall(".//w:r", ns)
            if "".join(node.text or "" for node in run.findall(".//w:t", ns))
        ]
        standard_range_styles.append(
            bool(visible_runs)
            and all(run_uses_standard_hyperlink_style(run) for run in visible_runs)
        )

    hyperlink_extra = ooxml_canonical.saved_body_visible_format_signature(
        root, hyperlink_extra_only=True
    )
    return {
        "outside_hyperlink_formatting": ooxml_canonical.saved_body_visible_format_signature(
            root, exclude_hyperlink_runs=True
        ),
        "hyperlink_extra_formatting": [
            paragraph["runs"] for paragraph in hyperlink_extra if paragraph["runs"]
        ],
        "standard_range_styles": standard_range_styles,
        "pictures": pictures,
    }
