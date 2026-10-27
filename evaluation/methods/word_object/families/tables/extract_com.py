"""Complete physical-cell observations, retaining property-specific failures."""

from typing import Dict, List
from ...observations import ObservedProperties
from .borders import CELL_EDGES, TABLE_EDGES, read_borders


def extract_table_properties(docx_path: str, *, close_document, open_document, include_cell_borders=False) -> List[Dict]:
    doc = None
    try:
        doc = open_document(docx_path)
        tables = []
        for index in range(1, doc.Tables.Count + 1):
            table = doc.Tables(index)
            state = ObservedProperties(index=index)
            state.read("row_count", lambda: table.Rows.Count)
            state.read("column_count", lambda: table.Columns.Count)
            cells = {}
            cell_error = None
            try:
                physical = table.Range.Cells
                for cell_index in range(1, physical.Count + 1):
                    cell = physical(cell_index)
                    coordinate = (cell.RowIndex, cell.ColumnIndex)
                    if coordinate in cells:
                        raise ValueError(f"Ambiguous physical cell coordinate: {coordinate}")
                    cells[coordinate] = cell
            except Exception as exc:
                cell_error = exc

            def matrix(getter):
                if cell_error is not None:
                    raise ValueError(f"Physical cell enumeration failed: {cell_error}")
                if state.unavailable:
                    # Counts define the grid; unrelated fields never affect it.
                    for field in ("row_count", "column_count"):
                        if field in state.unavailable:
                            raise ValueError(state.unavailable[field])
                return [[getter(cells[(row, col)]) if (row, col) in cells else None
                         for col in range(1, state["column_count"] + 1)]
                        for row in range(1, state["row_count"] + 1)]

            def text(cell):
                value = cell.Range.Text
                return value.removesuffix("\r\x07").removesuffix("\x07").strip()

            for name, getter in {
                "cell_contents": text,
                "cell_alignments": lambda c: c.VerticalAlignment,
                "cell_text_orientations": lambda c: c.Range.Orientation,
                "cell_paragraph_alignments": lambda c: c.Range.Paragraphs(1).Alignment,
            }.items():
                state.read(name, lambda getter=getter: matrix(getter))
            state.read("column_widths", lambda: [round(table.Columns(i).Width, 2)
                       for i in range(1, table.Columns.Count + 1)])
            state.read("row_heights", lambda: [round(table.Rows(i).Height, 2)
                       for i in range(1, table.Rows.Count + 1)])

            def cell_geometry():
                if cell_error is not None:
                    raise ValueError(f"Physical cell enumeration failed: {cell_error}")
                result = []
                for (row, column), cell in sorted(cells.items()):
                    item = ObservedProperties(row=row, column=column)
                    item.read("width", lambda: cell.Width, lambda value: round(value, 2))
                    item.read("height_rule", lambda: cell.HeightRule)
                    # Word explicitly leaves Height undefined for automatic
                    # rows. Preserve the observed rule instead of requiring it.
                    item.read("height", lambda: None if item["height_rule"] == 0 else cell.Height,
                              lambda value: round(value, 2))
                    result.append(item)
                return result

            state.read("cell_geometry", cell_geometry)

            def cell_borders():
                if cell_error is not None:
                    raise ValueError(f"Physical cell enumeration failed: {cell_error}")
                return [ObservedProperties(row=row, column=column, borders=read_borders(cell, CELL_EDGES))
                        for (row, column), cell in sorted(cells.items())]

            if include_cell_borders:
                state.read("cell_borders", cell_borders)
            state.read("borders", lambda: read_borders(table, TABLE_EDGES))
            state.read("style", lambda: table.Style.NameLocal if hasattr(table.Style, "NameLocal") else str(table.Style))
            tables.append(state)
        return tables
    finally:
        close_document(doc)
