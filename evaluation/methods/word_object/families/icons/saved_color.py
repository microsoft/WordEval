"""families / icons / saved_color with explicit evidence dependencies."""

from pathlib import Path
from typing import Dict, Optional
import re
import xml.etree.ElementTree as ET
import zipfile


def normalize_svg_color(
    value: Optional[str], *, current_color: Optional[str] = None
) -> Optional[str]:
    """Normalize the SVG color spellings used by benchmark predictions."""
    if value is None:
        return None
    raw = value.strip().casefold()
    if raw == "currentcolor":
        return current_color
    if raw in {"none", "transparent"}:
        return "NONE"
    named = {
        "black": "000000",
        "blue": "0000FF",
        "red": "FF0000",
        "white": "FFFFFF",
    }
    if raw in named:
        return named[raw]
    if re.fullmatch(r"#[0-9a-f]{3}", raw):
        return "".join(character * 2 for character in raw[1:]).upper()
    if re.fullmatch(r"#[0-9a-f]{6}", raw):
        return raw[1:].upper()
    rgb_match = re.fullmatch(
        r"rgb\(\s*([0-9]+)\s*[, ]\s*([0-9]+)\s*[, ]\s*([0-9]+)\s*\)",
        raw,
    )
    if rgb_match:
        channels = [min(255, int(channel)) for channel in rgb_match.groups()]
        return "".join(f"{channel:02X}" for channel in channels)
    return raw.upper()


def saved_svg_stroke_color_count(docx_path: Path, color_hex: str) -> int:
    """Count drawable SVG elements whose effective stroke has ``color_hex``.

    SVG presentation attributes inherit. Word writes the reference color on
    every path, while a semantically equivalent prediction may set one
    ``stroke`` on the root ``svg`` or a parent ``g``. Counting literal
    attribute occurrences therefore rejects valid documents. Resolve the
    small presentation subset used by the benchmark instead.
    """
    drawable_tags = {
        "circle",
        "ellipse",
        "line",
        "path",
        "polygon",
        "polyline",
        "rect",
        "text",
        "use",
    }
    target_color = color_hex.upper()

    def style_properties(node: ET.Element) -> Dict[str, str]:
        properties: Dict[str, str] = {}
        for declaration in node.get("style", "").split(";"):
            name, separator, value = declaration.partition(":")
            if separator:
                properties[name.strip().casefold()] = value.strip()
        return properties

    count = 0

    def visit(
        node: ET.Element,
        inherited_stroke: str = "NONE",
        inherited_color: str = "000000",
    ) -> None:
        nonlocal count
        properties = style_properties(node)

        raw_color = properties.get("color", node.get("color"))
        if raw_color is None or raw_color.strip().casefold() == "inherit":
            effective_color = inherited_color
        else:
            effective_color = (
                normalize_svg_color(
                    raw_color,
                    current_color=inherited_color,
                )
                or inherited_color
            )

        raw_stroke = properties.get("stroke", node.get("stroke"))
        if raw_stroke is None or raw_stroke.strip().casefold() == "inherit":
            effective_stroke = inherited_stroke
        else:
            effective_stroke = (
                normalize_svg_color(
                    raw_stroke,
                    current_color=effective_color,
                )
                or inherited_stroke
            )

        local_name = node.tag.rsplit("}", 1)[-1].casefold()
        if local_name in drawable_tags and effective_stroke == target_color:
            count += 1
        for child in node:
            visit(child, effective_stroke, effective_color)

    with zipfile.ZipFile(docx_path) as package:
        for name in package.namelist():
            if not name.lower().endswith(".svg"):
                continue
            visit(ET.fromstring(package.read(name)))
    return count
