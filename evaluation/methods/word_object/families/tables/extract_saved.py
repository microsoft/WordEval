"""families / tables / extract_saved with explicit evidence dependencies."""

from evaluation.methods.word_object.effect_commands import WORDPROCESSINGML_NS
from pathlib import Path
from typing import Any, Dict, List
import xml.etree.ElementTree as ET
import zipfile


def extract_saved_table_style_signatures(docx_path: Path) -> List[Dict[str, Any]]:
    w = WORDPROCESSINGML_NS
    ns = {"w": w}
    with zipfile.ZipFile(docx_path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    signatures = []
    for table in root.findall(".//w:tbl", ns):
        properties = table.find("w:tblPr", ns)
        style = properties.find("w:tblStyle", ns) if properties is not None else None
        signatures.append(
            {
                "style": style.get(f"{{{w}}}val") if style is not None else None,
            }
        )
    return signatures


def extract_saved_table_composition_signatures(docx_path: Path) -> List[Dict[str, Any]]:
    """Read table structure, preferred widths, and cell alignment."""
    w = WORDPROCESSINGML_NS
    ns = {"w": w}
    with zipfile.ZipFile(docx_path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    signatures = []
    for table in root.findall(".//w:tbl", ns):
        rows = table.findall("w:tr", ns)
        cell_rows = [row.findall("w:tc", ns) for row in rows]
        column_count = max((len(row) for row in cell_rows), default=0)
        preferred_columns = []
        for column in range(column_count):
            values = []
            for cells in cell_rows:
                if column >= len(cells):
                    continue
                width = cells[column].find("w:tcPr/w:tcW", ns)
                if width is None:
                    values.append((None, None))
                    continue
                raw = width.get(f"{{{w}}}w")
                try:
                    numeric = float(raw) if raw is not None else None
                except ValueError:
                    numeric = None
                values.append((width.get(f"{{{w}}}type"), numeric))
            preferred_columns.append(values)

        cell_alignments = []
        cell_contents = []
        for cells in cell_rows:
            alignment_row = []
            content_row = []
            for cell in cells:
                vertical = cell.find("w:tcPr/w:vAlign", ns)
                horizontal = cell.find("w:p/w:pPr/w:jc", ns)
                alignment_row.append(
                    {
                        "vertical": (
                            vertical.get(f"{{{w}}}val") if vertical is not None else "top"
                        ),
                        "horizontal": (
                            horizontal.get(f"{{{w}}}val") if horizontal is not None else "left"
                        ),
                    }
                )
                content_row.append("".join(node.text or "" for node in cell.findall(".//w:t", ns)))
            cell_alignments.append(alignment_row)
            cell_contents.append(content_row)
        signatures.append(
            {
                "row_count": len(rows),
                "column_count": column_count,
                "cell_contents": cell_contents,
                "preferred_columns": preferred_columns,
                "cell_alignments": cell_alignments,
            }
        )
    return signatures


def extract_table_autofit_flags(path: Path) -> List[Dict[str, Any]]:
    """Read semantic AutoFit mode from saved table properties.

    Word COM can expose different AllowAutoFit flags after a subsequent
    Distribute Columns command even when both files still persist the same
    requested 100%-of-window table.  The saved tblW/tblLayout pair captures
    the requested mode without those interaction side effects.
    """
    w = WORDPROCESSINGML_NS
    ns = {"w": w}
    with zipfile.ZipFile(path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    flags = []
    for table in root.findall(".//w:tbl", ns):
        properties = table.find("w:tblPr", ns)
        width = properties.find("w:tblW", ns) if properties is not None else None
        layout = properties.find("w:tblLayout", ns) if properties is not None else None
        width_type = width.get(f"{{{w}}}type") if width is not None else None
        width_value = width.get(f"{{{w}}}w") if width is not None else None
        layout_type = layout.get(f"{{{w}}}type") if layout is not None else None
        try:
            numeric_width = float(width_value) if width_value is not None else None
        except ValueError:
            numeric_width = None

        if width_type == "pct":
            preferred_width_type = 2  # wdPreferredWidthPercent
            preferred_width = round(numeric_width / 50.0, 2) if numeric_width is not None else None
            allow_autofit = layout_type != "fixed"
        elif layout_type == "fixed":
            preferred_width_type = 3  # wdPreferredWidthPoints
            preferred_width = (
                round(numeric_width / 20.0, 2)
                if width_type == "dxa" and numeric_width is not None
                else None
            )
            allow_autofit = False
        else:
            preferred_width_type = 1  # wdPreferredWidthAuto
            preferred_width = 0.0
            allow_autofit = True
        flags.append(
            {
                "allow_autofit": allow_autofit,
                "preferred_width_type": preferred_width_type,
                "preferred_width": preferred_width,
                "saved_width_type": width_type,
                "saved_layout_type": layout_type or "autofit_default",
            }
        )
    return flags
