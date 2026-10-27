"""families / pictures / extract_saved with explicit evidence dependencies."""

from evaluation.methods.word_object.effect_commands import DRAWINGML_NS, PICTURE_NS
from pathlib import Path
from typing import Any, Dict, List, Optional
import posixpath
import xml.etree.ElementTree as ET
import zipfile


def _extract_effect_persistence(docx_path: Path, effect_name: str, *, picture_only=False) -> List[Dict[str, Any]]:
    """Read selected effects together with their durable HD Photo graph."""
    a14 = "{http://schemas.microsoft.com/office/drawing/2010/main}"
    r = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    rel_ns = "{http://schemas.openxmlformats.org/package/2006/relationships}"
    with zipfile.ZipFile(docx_path) as package:
        names = set(package.namelist())
        root = ET.fromstring(package.read("word/document.xml"))
        relationships: Dict[str, Dict[str, str]] = {}
        relationship_part = "word/_rels/document.xml.rels"
        if relationship_part in names:
            relationship_root = ET.fromstring(package.read(relationship_part))
            relationships = {
                node.get("Id", ""): dict(node.attrib)
                for node in relationship_root.iter(f"{rel_ns}Relationship")
                if node.get("Id")
            }

    signatures: List[Dict[str, Any]] = []
    if picture_only:
        picture = next(root.iter(f"{{{PICTURE_NS}}}pic"), None)
        layers = picture.iter(f"{a14}imgLayer") if picture is not None else ()
    else:
        layers = root.iter(f"{a14}imgLayer")
    for layer in layers:
        relationship_id = layer.get(f"{r}embed", "")
        relationship = relationships.get(relationship_id)
        target = relationship.get("Target", "") if relationship else ""
        if target.startswith("/"):
            target_part = target.lstrip("/")
        else:
            target_part = posixpath.normpath(posixpath.join("word", target)) if target else ""
        relationship_type = relationship.get("Type", "") if relationship else ""
        target_mode = relationship.get("TargetMode", "") if relationship else ""
        graph_failures = []
        if not relationship_id:
            graph_failures.append("imgLayer has no embedded relationship")
        if relationship is None:
            graph_failures.append("imgLayer relationship is missing")
        if target_mode and target_mode.casefold() != "internal":
            graph_failures.append("imgLayer relationship is not internal")
        if not relationship_type.casefold().endswith("/hdphoto"):
            graph_failures.append("imgLayer relationship is not HD Photo")
        if not target_part.casefold().endswith(".wdp"):
            graph_failures.append("imgLayer target is not a WDP part")
        if target_part not in names:
            graph_failures.append("imgLayer WDP target part is missing")
        for effect in layer.iter():
            local_name = effect.tag.removeprefix(a14)
            if not effect.tag.startswith(a14) or not (
                local_name == effect_name or effect_name == "artistic" and local_name.startswith("artistic")
            ):
                continue
            raw_sat = effect.get("sat")
            try:
                sat: Any = int(raw_sat) if raw_sat is not None else None
            except ValueError:
                sat = raw_sat
            signatures.append(
                {
                    "sat": sat,
                    "effect": local_name,
                    "relationship_id": relationship_id,
                    "relationship_type": relationship_type,
                    "target_mode": target_mode,
                    "target_part": target_part,
                    "graph_valid": not graph_failures,
                    "graph_failures": graph_failures,
                }
            )
    return signatures


def extract_picture_saturation_persistence(docx_path: Path) -> List[Dict[str, Any]]:
    return _extract_effect_persistence(docx_path, "saturation")


def extract_picture_effect_persistence(docx_path: Path, effect_name: str) -> List[Dict[str, Any]]:
    """Require the effect on the target picture, ignoring unrelated orphan nodes."""
    return _extract_effect_persistence(docx_path, effect_name, picture_only=True)


def extract_saved_picture_glows(docx_path: Path) -> Optional[List[Dict[str, Any]]]:
    """Return glow elements saved on DrawingML pictures, before Word repairs them."""
    a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    pic = "http://schemas.openxmlformats.org/drawingml/2006/picture"
    ns = {"a": a, "pic": pic}
    try:
        with zipfile.ZipFile(docx_path) as package:
            root = ET.fromstring(package.read("word/document.xml"))
    except (FileNotFoundError, KeyError, OSError, zipfile.BadZipFile, ET.ParseError):
        return None

    glows = []
    for picture in root.findall(".//pic:pic", ns):
        for glow in picture.findall(".//a:glow", ns):
            radius = glow.get("rad")
            color = next(iter(glow), None)
            glows.append(
                {
                    "radius_pt": (round(float(radius) / 12700, 2) if radius is not None else None),
                    "color_type": (color.tag.rsplit("}", 1)[-1] if color is not None else None),
                    "color_value": (color.get("val") if color is not None else None),
                }
            )
    return glows


def extract_saved_picture_named_effects(docx_path: Path) -> Dict[str, List[Dict[str, Any]]]:
    """Project named picture shadow/reflection geometry from DrawingML."""
    a = f"{{{DRAWINGML_NS}}}"
    pic = f"{{{PICTURE_NS}}}"
    result: Dict[str, List[Dict[str, Any]]] = {
        "outerShdw": [],
        "reflection": [],
    }
    with zipfile.ZipFile(docx_path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    for picture in root.iter(f"{pic}pic"):
        for family in result:
            for effect in picture.iter(f"{a}{family}"):
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
                result[family].append(
                    {
                        "family": family,
                        "attributes": dict(effect.attrib),
                        "color_type": (color.tag.rsplit("}", 1)[-1] if color is not None else None),
                        "color_value": (color.get("val") if color is not None else None),
                        "alpha": (alpha.get("val") if alpha is not None else None),
                    }
                )
    return result
