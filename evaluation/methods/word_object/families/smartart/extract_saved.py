"""families / smartart / extract_saved with explicit evidence dependencies."""

from evaluation.methods.word_object.effect_commands import DRAWINGML_NS, EMU_PER_POINT
from pathlib import Path
from typing import Any, Dict, List
import xml.etree.ElementTree as ET
import zipfile


def extract_saved_smartart_connector_styles(docx_path: Path) -> Dict[str, Any]:
    """Read visible connector formatting from SmartArt drawing caches.

    SmartArt.Nodes exposes the data model and a node's own outline.  The
    connector lines are separate shapes in drawingN.xml, so node-outline
    properties can falsely pass connector weight/dash tasks.  Preserve all
    dash children and normalized arrow state so malformed lookalikes (for
    example tailEnd val="triangle" instead of type="triangle") fail.
    """
    diagram = "http://schemas.microsoft.com/office/drawing/2008/diagram"
    drawing = DRAWINGML_NS

    def line_end_signature(
        elements: List[ET.Element],
    ) -> tuple[tuple[str, str], ...]:
        # Office resolves omitted DrawingML line-end state to type=none,
        # width=medium, length=medium. Normalize explicit/default forms,
        # while retaining duplicates, unexpected attributes, text, and
        # child markup so malformed lookalikes still fail.
        element = elements[0] if elements else None
        attributes = dict(element.attrib) if element is not None else {}
        resolved = {
            "type": attributes.pop("type", "none"),
            "w": attributes.pop("w", "med"),
            "len": attributes.pop("len", "med"),
        }
        resolved.update({f"unexpected:attribute:{key}": value for key, value in attributes.items()})
        if len(elements) > 1:
            resolved["unexpected:element_count"] = str(len(elements))
            for index, duplicate in enumerate(elements[1:], start=2):
                resolved[f"unexpected:element_{index}"] = ET.tostring(duplicate, encoding="unicode")
        if element is not None:
            if (element.text or "").strip():
                resolved["unexpected:text"] = (element.text or "").strip()
            for index, child in enumerate(element, start=1):
                resolved[f"unexpected:child_{index}"] = ET.tostring(child, encoding="unicode")
        return tuple(sorted(resolved.items()))

    connectors = []
    with zipfile.ZipFile(docx_path) as package:
        drawing_parts = sorted(
            name
            for name in package.namelist()
            if name.startswith("word/diagrams/drawing") and name.endswith(".xml")
        )
        for part in drawing_parts:
            root = ET.fromstring(package.read(part))
            for shape in root.iter(f"{{{diagram}}}sp"):
                geometry = next(
                    shape.iter(f"{{{drawing}}}prstGeom"),
                    None,
                )
                if geometry is None or geometry.get("prst") != "line":
                    continue
                line = next(
                    shape.iter(f"{{{drawing}}}ln"),
                    None,
                )
                if line is None:
                    connectors.append(
                        {
                            "weight": None,
                            "dashes": (),
                            "head": line_end_signature([]),
                            "tail": line_end_signature([]),
                        }
                    )
                    continue
                dashes = tuple(
                    child.get("val") for child in line if child.tag == f"{{{drawing}}}prstDash"
                )
                head = [child for child in line if child.tag == f"{{{drawing}}}headEnd"]
                tail = [child for child in line if child.tag == f"{{{drawing}}}tailEnd"]

                connectors.append(
                    {
                        "weight": (
                            round(float(line.get("w")) / EMU_PER_POINT, 3)
                            if line.get("w") is not None
                            else None
                        ),
                        "dashes": dashes,
                        "head": line_end_signature(head),
                        "tail": line_end_signature(tail),
                    }
                )
    return {
        "connector_count": len(connectors),
        "connector_weights": [item["weight"] for item in connectors],
        "connector_dashes": [item["dashes"] for item in connectors],
        "connector_arrows": [{"head": item["head"], "tail": item["tail"]} for item in connectors],
    }
