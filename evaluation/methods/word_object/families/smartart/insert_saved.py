"""families / smartart / insert_saved with explicit evidence dependencies."""

from pathlib import Path
from typing import Any, Dict, List
import xml.etree.ElementTree as ET
import zipfile


def extract_saved_smartart_signature(
    docx_path: Path, *, saved_part_relationships
) -> Dict[str, Any]:
    """Extract SmartArt layout, ordered node graph, and drawing cache."""
    dgm = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
    a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ns = {"dgm": dgm, "a": a}
    errors: List[str] = []
    diagrams = []
    with zipfile.ZipFile(docx_path) as package:
        names = set(package.namelist())
        document = ET.fromstring(package.read("word/document.xml"))
        relationships = saved_part_relationships(package, "word/document.xml")
        references = document.findall(f".//{{{dgm}}}relIds")
        drawing_relationships = [
            value
            for value in relationships.values()
            if value.get("type", "").endswith("/diagramDrawing")
        ]
        for reference in references:
            parts = {}
            for key, suffix in (
                ("data", "/diagramData"),
                ("layout", "/diagramLayout"),
                ("quick_style", "/diagramQuickStyle"),
                ("colors", "/diagramColors"),
            ):
                relationship_id = reference.get(
                    f'{{{r}}}{dict(data="dm", layout="lo", quick_style="qs", colors="cs")[key]}'
                )
                relationship = relationships.get(relationship_id or "")
                if (
                    not relationship
                    or relationship.get("external")
                    or not relationship.get("type", "").endswith(suffix)
                    or relationship.get("part") not in names
                ):
                    errors.append(f"SmartArt {key} relationship/part is invalid")
                    parts[key] = None
                else:
                    parts[key] = relationship.get("part")
            if len(drawing_relationships) != 1:
                errors.append("SmartArt must have exactly one drawing-cache relationship")
                drawing_part = None
            else:
                drawing_part = drawing_relationships[0].get("part")
                if drawing_part not in names:
                    errors.append("SmartArt drawing-cache part is missing")
                    drawing_part = None

            if not parts.get("data"):
                continue
            data_root = ET.fromstring(package.read(parts["data"]))
            point_rows = []
            id_to_text = {}
            root_ids = set()
            layout_id = None
            for point in data_root.findall(".//dgm:ptLst/dgm:pt", ns):
                point_id = point.get("modelId", "")
                point_type = point.get("type")
                if point_type == "doc":
                    root_ids.add(point_id)
                    properties = point.find("dgm:prSet", ns)
                    if properties is not None:
                        layout_id = properties.get("loTypeId") or layout_id
                    continue
                if point_type:
                    continue
                text = "".join(node.text or "" for node in point.findall(".//a:t", ns)).strip()
                if not text:
                    continue
                id_to_text[point_id] = text
                point_rows.append((point_id, text))

            edges = []
            for connection in data_root.findall(".//dgm:cxnLst/dgm:cxn", ns):
                if connection.get("type"):
                    continue
                source_id = connection.get("srcId", "")
                destination_id = connection.get("destId", "")
                if source_id not in root_ids | set(id_to_text):
                    continue
                if destination_id not in id_to_text:
                    continue
                edges.append(
                    {
                        "parent": ("__root__" if source_id in root_ids else id_to_text[source_id]),
                        "child": id_to_text[destination_id],
                        "order": int(connection.get("srcOrd", "0")),
                    }
                )

            depths = {"__root__": 0}
            unresolved = list(edges)
            for _ in range(len(edges) + 1):
                changed = False
                for edge in unresolved:
                    if edge["parent"] in depths and edge["child"] not in depths:
                        depths[edge["child"]] = depths[edge["parent"]] + 1
                        changed = True
                if not changed:
                    break
            node_order = [text for _, text in point_rows]
            node_levels = [depths.get(text) for text in node_order]
            drawing_texts = []
            if drawing_part:
                drawing_root = ET.fromstring(package.read(drawing_part))
                drawing_texts = [
                    (node.text or "").strip()
                    for node in drawing_root.findall(".//a:t", ns)
                    if (node.text or "").strip()
                ]
                if drawing_texts != node_order:
                    errors.append("SmartArt drawing cache is stale or reordered")
            diagrams.append(
                {
                    "layout_id": layout_id,
                    "node_order": node_order,
                    "node_levels": node_levels,
                    "edges": edges,
                    "drawing_texts": drawing_texts,
                    "parts": {**parts, "drawing": drawing_part},
                }
            )
        return {
            "object_count": len(references),
            "body_object_count": sum(
                1
                for node in document.iter()
                if node.tag in {f"{{{w}}}drawing", f"{{{w}}}pict", f"{{{w}}}object"}
            ),
            "diagrams": diagrams,
            "errors": errors,
        }
