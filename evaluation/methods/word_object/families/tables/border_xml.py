"""Resolve explicit no-border states that Word exposes as wdUndefined.

The fallback reads Word's in-memory serialization of the exact native table.
It only proves absence when cell, row and style overrides cannot add borders;
it does not implement a second general-purpose border conflict resolver.
"""
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
NAMES = {"horizontal": "insideH", "vertical": "insideV"}
ABSENT = {"nil", "none"}


def explicit_empty_table_edges(xml):
    root = ET.fromstring(xml)
    tables = list(root.iter(W + "tbl"))
    if len(tables) != 1:
        return set()
    table = tables[0]
    borders = table.find(f"{W}tblPr/{W}tblBorders")
    if borders is None:
        return set()

    def adds_border(container):
        return any(edge.get(W + "val") not in ABSENT for edge in container)

    # Direct cell borders and row exceptions can override table-level nil.
    for path in (f"{W}tr/{W}tc/{W}tcPr/{W}tcBorders",
                 f"{W}tr/{W}tblPrEx/{W}tblBorders"):
        if any(adds_border(node) for node in table.findall(path)):
            return set()
    # Conditional cell borders may come from the applied/default table style.
    styles = {s.get(W + "styleId"): s for s in root.iter(W + "style")
              if s.get(W + "type") == "table"}
    declared = table.find(f"{W}tblPr/{W}tblStyle")
    if declared is not None:
        style_id = declared.get(W + "val")
    else:
        default = next(root.iter(W + "defaultTableStyle"), None)
        defaults = [key for key, value in styles.items() if value.get(W + "default") in {"1", "true", "on"}]
        if default is None and len(defaults) > 1:
            return set()
        style_id = default.get(W + "val") if default is not None else next(iter(defaults), None)
    seen = set()
    while style_id:
        if style_id in seen or style_id not in styles:
            return set()
        seen.add(style_id)
        style = styles[style_id]
        if any(adds_border(node) for node in style.iter(W + "tcBorders")):
            return set()
        parent = style.find(W + "basedOn")
        style_id = parent.get(W + "val") if parent is not None else None
    return {edge.tag.removeprefix(W) for edge in borders if edge.get(W + "val") in ABSENT}


def border_style(obj, kind, name, *, scope):
    value = obj.Borders(-kind).LineStyle
    if value not in (9999999, -9999999):
        return value
    if scope == "table":
        table, edge = obj, NAMES.get(name, name)
    else:
        rng = obj.Range
        if rng.Tables.Count != 1 or rng.Cells.Count != 1:
            return value
        table, cell = rng.Tables(1), rng.Cells(1)
        # Merged/irregular grids need physical border evidence, not an inferred
        # mapping from a serialized one-cell fragment to the original table.
        rows, columns = table.Rows.Count, table.Columns.Count
        if table.Range.Cells.Count != rows * columns:
            return value
        row, column = cell.RowIndex, cell.ColumnIndex
        edge = {"top": "top" if row == 1 else "insideH",
                "bottom": "bottom" if row == rows else "insideH",
                "left": "left" if column == 1 else "insideV",
                "right": "right" if column == columns else "insideV"}[name]
        if scope == "paragraph":
            paragraph_xml = ET.fromstring(rng.WordOpenXML)
            # Paragraph borders are separate from enclosing cell borders.
            body = paragraph_xml.find(f".//{W}body")
            if body is None:
                return value
            if any(child.get(W + "val") not in ABSENT
                   for node in body.iter(W + "pBdr") for child in node):
                return value
            styles = {s.get(W + "styleId"): s for s in paragraph_xml.iter(W + "style")
                      if s.get(W + "type") == "paragraph"}
            for paragraph in body.iter(W + "p"):
                declared = paragraph.find(f"{W}pPr/{W}pStyle")
                default = next((key for key, style in styles.items()
                                if style.get(W + "default") in {"1", "true", "on"}), None)
                style_id = declared.get(W + "val") if declared is not None else default
                seen = set()
                while style_id:
                    if style_id in seen or style_id not in styles:
                        return value
                    seen.add(style_id)
                    style = styles[style_id]
                    if any(child.get(W + "val") not in ABSENT
                           for node in style.iter(W + "pBdr") for child in node):
                        return value
                    parent = style.find(W + "basedOn")
                    style_id = parent.get(W + "val") if parent is not None else None
    return 0 if edge in explicit_empty_table_edges(table.Range.WordOpenXML) else value
