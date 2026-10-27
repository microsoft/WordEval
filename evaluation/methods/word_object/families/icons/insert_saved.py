"""families / icons / insert_saved with explicit evidence dependencies."""

from pathlib import Path
from typing import Any, Dict, List, Optional
import hashlib
import math
import re
import xml.etree.ElementTree as ET
import zipfile


def normalize_svg_color(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    # Accept the same input spellings as the object_effect normalizer
    # (named colors, none/transparent, space-separated rgb()) so an
    # identical prediction cannot pass one behavior's routes and fail the
    # other's. This file's call sites expect the '#RRGGBB' convention.
    text = value.strip().upper()
    named = {
        "BLACK": "#000000",
        "BLUE": "#0000FF",
        "RED": "#FF0000",
        "WHITE": "#FFFFFF",
    }
    if text in named:
        return named[text]
    if text in {"NONE", "TRANSPARENT"}:
        return "NONE"
    if re.fullmatch(r"#[0-9A-F]{6}", text):
        return text
    short = re.fullmatch(r"#([0-9A-F])([0-9A-F])([0-9A-F])", text)
    if short:
        return "#" + "".join(group * 2 for group in short.groups())
    rgb = re.fullmatch(r"RGB\(\s*(\d+)\s*[, ]\s*(\d+)\s*[, ]\s*(\d+)\s*\)", text)
    if rgb:
        values = [min(255, int(group)) for group in rgb.groups()]
        return "#" + "".join(f"{item:02X}" for item in values)
    return text


def parse_svg_absolute_length(value: Optional[str]) -> Optional[float]:
    """Return an absolute SVG length in CSS px, or None if unsupported."""
    if value is None:
        return None
    match = re.fullmatch(
        r"\s*([+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)" r"\s*(px|pt|pc|in|cm|mm|q)?\s*",
        value,
        flags=re.IGNORECASE,
    )
    if not match:
        return None
    numeric = float(match.group(1))
    if not math.isfinite(numeric) or numeric <= 0:
        return None
    unit = (match.group(2) or "px").lower()
    factors = {
        "px": 1.0,
        "pt": 96.0 / 72.0,
        "pc": 16.0,
        "in": 96.0,
        "cm": 96.0 / 2.54,
        "mm": 96.0 / 25.4,
        "q": 96.0 / 101.6,
    }
    return numeric * factors[unit]


def extract_saved_svg_payload(payload: bytes, *, canonical_chart_value) -> Dict[str, Any]:
    text = payload.decode("utf-8", errors="replace")
    root = ET.fromstring(payload)
    local = lambda tag: tag.rsplit("}", 1)[-1].lower()
    drawables = {"path", "polygon", "circle", "rect", "ellipse"}
    fills: List[str] = []

    def visit(node: ET.Element, inherited_fill: Optional[str] = None) -> None:
        style = {}
        for item in (node.get("style") or "").split(";"):
            if ":" in item:
                key, value = item.split(":", 1)
                style[key.strip().lower()] = value.strip()
        fill = node.get("fill") or style.get("fill") or inherited_fill
        if local(node.tag) in drawables and fill and fill.lower() != "none":
            normalized = normalize_svg_color(fill)
            if normalized:
                fills.append(normalized)
        for child in node:
            visit(child, fill)

    visit(root)
    shape = "unknown"
    if any(local(node.tag) == "circle" for node in root.iter()):
        shape = "circle"
    elif any(local(node.tag) == "polygon" for node in root.iter()):
        polygon = next(node for node in root.iter() if local(node.tag) == "polygon")
        if len(re.findall(r"-?\d+(?:\.\d+)?", polygon.get("points", ""))) >= 16:
            shape = "star"
    else:
        path = next((node for node in root.iter() if local(node.tag) == "path"), None)
        if path is not None:
            data = path.get("d", "")
            if (
                not re.search(r"[acqs]", data, flags=re.IGNORECASE)
                and len(re.findall(r"-?\d+(?:\.\d+)?", data)) >= 16
            ):
                shape = "star"

    view_box_present = "viewBox" in root.attrib or "viewbox" in root.attrib
    view_box = root.get("viewBox") if "viewBox" in root.attrib else root.get("viewbox")
    view_box_values = (
        [canonical_chart_value(item) for item in re.split(r"[\s,]+", view_box)]
        if view_box
        else None
    )
    svg_width = root.get("width")
    svg_height = root.get("height")
    absolute_width = parse_svg_absolute_length(svg_width)
    absolute_height = parse_svg_absolute_length(svg_height)
    effective_viewport = view_box_values
    viewport_source = "viewBox" if view_box_present else None
    if not view_box_present and absolute_width is not None and absolute_height is not None:
        effective_viewport = [0, 0, absolute_width, absolute_height]
        viewport_source = "root_absolute_dimensions"
    return {
        "shape": shape,
        "fill_colors": sorted(set(fills)),
        "view_box": view_box_values,
        "svg_width": svg_width,
        "svg_height": svg_height,
        "absolute_svg_width_px": absolute_width,
        "absolute_svg_height_px": absolute_height,
        "effective_viewport": effective_viewport,
        "viewport_source": viewport_source,
        "payload_sha256": hashlib.sha256(payload).hexdigest(),
        "has_svg_root": local(root.tag) == "svg",
        "text_length": len(text),
    }


def extract_saved_svg_icon_signature(
    docx_path: Path, *, canonical_chart_value, saved_part_relationships
) -> Dict[str, Any]:
    """Extract only SVGs referenced by visible Word drawing objects."""
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    asvg = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
    errors: List[str] = []
    records = []
    with zipfile.ZipFile(docx_path) as package:
        names = set(package.namelist())
        document = ET.fromstring(package.read("word/document.xml"))
        relationships = saved_part_relationships(package, "word/document.xml")
        containers = [
            node for node in document.iter() if node.tag in {f"{{{wp}}}inline", f"{{{wp}}}anchor"}
        ]
        referenced_parts = []
        for container in containers:
            candidates = []
            for reference in container.findall(f".//{{{asvg}}}svgBlip"):
                candidates.append(reference)
            for reference in container.findall(f".//{{{a}}}blip"):
                relationship_id = reference.get(f"{{{r}}}embed")
                relationship = relationships.get(relationship_id or "")
                target = (relationship or {}).get("target", "")
                part = (relationship or {}).get("part") or ""
                # A normal PNG/JPEG fallback is not an SVG reference.  A
                # direct a:blip is considered only when its relationship
                # actually targets an SVG (including invalid/external SVG
                # relationships, which must be reported below).
                if target.lower().endswith(".svg") or part.lower().endswith(".svg"):
                    candidates.append(reference)

            seen_relationships = set()
            for reference in candidates:
                relationship_id = reference.get(f"{{{r}}}embed")
                if relationship_id in seen_relationships:
                    continue
                seen_relationships.add(relationship_id)
                relationship = relationships.get(relationship_id or "")
                if (
                    not relationship
                    or relationship.get("external")
                    or not relationship.get("type", "").endswith("/image")
                ):
                    errors.append("SVG reference has no valid image relationship")
                    continue
                part = relationship.get("part")
                if not part or part not in names or not part.lower().endswith(".svg"):
                    errors.append("referenced SVG part is missing or not SVG")
                    continue
                referenced_parts.append(part)
                extent = container.find(f"{{{wp}}}extent")
                width = height = None
                if extent is None:
                    errors.append("SVG drawing has no displayed extent")
                else:
                    width = round(float(extent.get("cx", "0")) / 12700.0, 2)
                    height = round(float(extent.get("cy", "0")) / 12700.0, 2)
                records.append(
                    {
                        "part": part,
                        "object_kind": (
                            "inline" if container.tag == f"{{{wp}}}inline" else "floating"
                        ),
                        "width": width,
                        "height": height,
                        **extract_saved_svg_payload(
                            package.read(part), canonical_chart_value=canonical_chart_value
                        ),
                    }
                )
        svg_media = sorted(
            name
            for name in names
            if name.startswith("word/media/") and name.lower().endswith(".svg")
        )
        if sorted(referenced_parts) != svg_media:
            errors.append("package contains unreferenced or multiply referenced SVG media")
        return {
            "object_count": len(records),
            "body_object_count": sum(
                1
                for node in document.iter()
                if node.tag in {f"{{{w}}}drawing", f"{{{w}}}pict", f"{{{w}}}object"}
            ),
            "icons": records,
            "svg_media_parts": svg_media,
            "errors": errors,
        }
