"""families / tables / insert_saved with explicit evidence dependencies."""

from pathlib import Path
from typing import Any, Dict
import xml.etree.ElementTree as ET
import zipfile


def extract_saved_table_insert_signature(docx_path: Path) -> Dict[str, Any]:
    """Extract inserted tables plus visible top-level preservation controls."""
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ns = {"w": w}
    with zipfile.ZipFile(docx_path) as package:
        document = ET.fromstring(package.read("word/document.xml"))

    signatures = []
    for table in document.findall(".//w:tbl", ns):
        rows = []
        for row in table.findall("w:tr", ns):
            cells = []
            for cell in row.findall("w:tc", ns):
                grid_span = cell.find("w:tcPr/w:gridSpan", ns)
                try:
                    span = int(grid_span.get(f"{{{w}}}val", "1"))
                except (AttributeError, TypeError, ValueError):
                    span = 1
                text = "".join(node.text or "" for node in cell.findall(".//w:t", ns)).strip()
                cells.append({"span": span, "text": text})
            rows.append(cells)
        signatures.append(
            {
                "row_count": len(rows),
                "column_counts": [sum(cell["span"] for cell in row) for row in rows],
                "cell_text": [[cell["text"] for cell in row] for row in rows],
            }
        )
    body_paragraph_text = [
        "".join(node.text or "" for node in paragraph.findall(".//w:t", ns))
        for paragraph in document.findall("./w:body/w:p", ns)
    ]
    return {
        "tables": signatures,
        "body_paragraph_text": body_paragraph_text,
        "drawing_count": len(document.findall(".//w:body//w:drawing", ns)),
        "legacy_drawing_count": len(document.findall(".//w:body//w:pict", ns)),
    }
