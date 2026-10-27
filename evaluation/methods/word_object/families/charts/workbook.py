"""Chart workbook decoding and bounded local-range projection."""

import io
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET
from typing import Any, Dict, List, Optional
from ...ooxml.package import part_relationships


def canonical_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        numeric = float(value)
    else:
        text = str(value).strip()
        try:
            numeric = float(text)
        except (TypeError, ValueError):
            return text
    if numeric.is_integer():
        return int(numeric)
    return round(numeric, 10)


def extract_workbook_cells(
    payload: bytes,
) -> Dict[str, Dict[str, Any]]:
    """Extract persisted cell values keyed by workbook sheet and A1 ref."""
    spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    office_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    package_rel = "http://schemas.openxmlformats.org/package/2006/relationships"
    ns = {"x": spreadsheet, "r": office_rel}
    with zipfile.ZipFile(io.BytesIO(payload)) as workbook:
        names = set(workbook.namelist())
        shared_strings: List[str] = []
        if "xl/sharedStrings.xml" in names:
            shared_root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
            for item in shared_root.findall("x:si", ns):
                shared_strings.append(
                    "".join(node.text or "" for node in item.findall(".//x:t", ns))
                )
        if "xl/workbook.xml" not in names or "xl/_rels/workbook.xml.rels" not in names:
            raise ValueError("embedded workbook has no workbook graph")
        workbook_root = ET.fromstring(workbook.read("xl/workbook.xml"))
        rels_root = ET.fromstring(workbook.read("xl/_rels/workbook.xml.rels"))
        relationship_parts = {
            relationship.get("Id", ""): posixpath.normpath(
                posixpath.join("xl", relationship.get("Target", ""))
            ).lstrip("/")
            for relationship in rels_root.findall(f"{{{package_rel}}}Relationship")
            if relationship.get("TargetMode") != "External"
        }
        sheets: Dict[str, str] = {}
        for sheet in workbook_root.findall("x:sheets/x:sheet", ns):
            name = sheet.get("name", "")
            relationship_id = sheet.get(f"{{{office_rel}}}id", "")
            part = relationship_parts.get(relationship_id)
            if name and part:
                sheets[name] = part
        if not sheets:
            raise ValueError("embedded workbook has no worksheet")
        result: Dict[str, Dict[str, Any]] = {}
        for sheet_name, sheet_part in sheets.items():
            if sheet_part not in names:
                raise ValueError(f"embedded workbook worksheet is missing: {sheet_name}")
            root = ET.fromstring(workbook.read(sheet_part))
            values: Dict[str, Any] = {}
            for cell in root.findall(".//x:sheetData/x:row/x:c", ns):
                reference = (cell.get("r") or "").upper()
                if not re.fullmatch(r"[A-Z]{1,3}[1-9]\d*", reference):
                    raise ValueError(f"worksheet {sheet_name!r} has an invalid cell reference")
                cell_type = cell.get("t")
                if cell_type == "inlineStr":
                    value: Any = "".join(node.text or "" for node in cell.findall(".//x:t", ns))
                else:
                    value_node = cell.find("x:v", ns)
                    value = value_node.text if value_node is not None else None
                    if cell_type == "s" and value is not None:
                        value = shared_strings[int(value)]
                    elif cell_type not in {"str", "e"}:
                        value = canonical_value(value)
                values[reference] = value
            _validate_sheet_tables(workbook, sheet_part, root, values)
            result[sheet_name] = values
        return result


def _validate_sheet_tables(workbook, sheet_part, sheet, values):
    """Check only referenced Excel tables; formatted blank cells do not set their extent."""
    x = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    table_parts = sheet.find(x+'tableParts')
    if table_parts is None:
        return
    references = table_parts.findall(x+'tablePart')
    if int(table_parts.get('count', str(len(references)))) != len(references):
        raise ValueError('worksheet tableParts count is inconsistent')
    relationships = part_relationships(workbook, sheet_part)
    for reference in references:
        relationship = relationships.get(reference.get(r+'id', ''))
        if (not relationship or relationship['external']
                or not relationship['type'].endswith('/table')
                or relationship['part'] not in workbook.namelist()):
            raise ValueError('worksheet table relationship is invalid or missing')
        table = ET.fromstring(workbook.read(relationship['part']))
        match = re.fullmatch(r'\$?([A-Z]{1,3})\$?([1-9]\d*):\$?([A-Z]{1,3})\$?([1-9]\d*)', table.get('ref',''))
        if not match:
            raise ValueError('worksheet table range is invalid')
        left, top, right, bottom = match.groups()
        start, end = column_number(left), column_number(right)
        if end < start or int(bottom) < int(top):
            raise ValueError('worksheet table range is reversed')
        columns = table.find(x+'tableColumns')
        if columns is None:
            raise ValueError('worksheet table has no column definitions')
        items = columns.findall(x+'tableColumn')
        if len(items) != end-start+1 or int(columns.get('count','-1')) != len(items):
            raise ValueError('worksheet table column count does not match its range')
        names = [item.get('name') for item in items]
        if any(name is None for name in names) or len({name.casefold() for name in names}) != len(names):
            raise ValueError('worksheet table column names are missing or duplicated')
        if table.get('headerRowCount','1') != '0':
            for offset, name in enumerate(names):
                cell = f'{column_name(start+offset)}{top}'
                if canonical_value(values.get(cell)) != canonical_value(name):
                    raise ValueError(f'worksheet table header does not match cell {cell}')


def column_number(column: str) -> int:
    number = 0
    for character in column:
        number = number * 26 + ord(character) - ord("A") + 1
    return number


def column_name(number: int) -> str:
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def parse_range_formula(formula: Optional[str]) -> Optional[Dict[str, Any]]:
    """Parse one local rectangular A1 chart reference.

    External workbooks, unions, formulas, and reversed ranges are rejected;
    they cannot define the bounded editable-data contract used here.
    """
    if formula is None:
        return None
    text = str(formula).strip()
    if text.startswith("="):
        text = text[1:].strip()
    if any(token in text for token in ("[", "]", ",")):
        return None
    match = re.fullmatch(
        r"(?:'((?:[^']|'')+)'|([^!']+))!"
        r"\$?([A-Za-z]{1,3})\$?([1-9]\d*)"
        r"(?::\$?([A-Za-z]{1,3})\$?([1-9]\d*))?",
        text,
    )
    if not match:
        return None
    quoted_sheet, plain_sheet, start_column, start_row, end_column, end_row = match.groups()
    sheet = quoted_sheet.replace("''", "'") if quoted_sheet is not None else plain_sheet.strip()
    start_column_number = column_number(start_column.upper())
    end_column_number = column_number((end_column or start_column).upper())
    start_row_number = int(start_row)
    end_row_number = int(end_row or start_row)
    if end_column_number < start_column_number or end_row_number < start_row_number:
        return None
    cell_count = (end_column_number - start_column_number + 1) * (
        end_row_number - start_row_number + 1
    )
    if cell_count > 10000:
        return None
    cells = [
        f"{column_name(column)}{row}"
        for row in range(start_row_number, end_row_number + 1)
        for column in range(start_column_number, end_column_number + 1)
    ]
    return {"sheet": sheet, "cells": cells, "formula": str(formula)}


def project_workbook(
    workbook_cells: Dict[str, Dict[str, Any]],
    formula: Optional[str],
) -> Optional[List[Any]]:
    parsed = parse_range_formula(formula)
    if parsed is None:
        return None
    sheet = workbook_cells.get(parsed["sheet"])
    if sheet is None:
        return None
    return [canonical_value(sheet.get(reference)) for reference in parsed["cells"]]


def extract_workbook_grid(payload: bytes) -> List[List[Any]]:
    """Compatibility view of the first sheet; chart checks use formulas."""
    sheets = extract_workbook_cells(payload)
    cells = next(iter(sheets.values()))
    coordinates = [re.fullmatch(r"([A-Z]{1,3})([1-9]\d*)", reference) for reference in cells]
    if not coordinates:
        return []
    max_column = max(column_number(match.group(1)) for match in coordinates if match)
    max_row = max(int(match.group(2)) for match in coordinates if match)
    return [
        [cells.get(f"{column_name(column)}{row}") for column in range(1, max_column + 1)]
        for row in range(1, max_row + 1)
    ]
