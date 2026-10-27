"""Saved chart projections for effect and insertion contracts.

The two projections intentionally preserve their existing scope: effect checks
read chart properties, while insertion checks follow document relationships and
validate caches and embedded workbook backing.
"""

import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Dict, List, Optional

from ...ooxml.package import part_relationships
from .workbook import canonical_value, extract_workbook_cells, extract_workbook_grid

DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def extract_properties(docx_path: Path) -> Dict[str, Any]:
    """Read the chart properties used by basic title/legend/label/color cases."""
    c = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    a = DRAWINGML_NS
    cs = "http://schemas.microsoft.com/office/drawing/2012/chartStyle"
    ns = {"c": c, "a": a, "cs": cs}
    with zipfile.ZipFile(docx_path) as package:
        chart_parts = sorted(
            name
            for name in package.namelist()
            if name.startswith("word/charts/chart") and name.endswith(".xml")
        )
        if not chart_parts:
            raise ValueError(f"no saved chart part in {docx_path}")
        root = ET.fromstring(package.read(chart_parts[0]))
        color_parts = sorted(
            name
            for name in package.namelist()
            if name.startswith("word/charts/colors") and name.endswith(".xml")
        )
        color_signature = None
        if color_parts:
            colors = ET.fromstring(package.read(color_parts[0]))
            color_signature = (
                colors.get("id"),
                colors.get("meth"),
                tuple(node.get("val") for node in colors.findall("a:schemeClr", ns)),
                tuple(
                    tuple(sorted(node.attrib.items()))
                    for node in colors.findall("cs:variation", ns)
                ),
            )

    title = root.find("c:chart/c:title", ns)
    title_text = (
        "".join(node.text or "" for node in title.findall(".//a:t", ns))
        if title is not None
        else ""
    )
    title_overlay = title.find("c:overlay", ns) if title is not None else None
    title_is_overlay = title_overlay is not None and title_overlay.get("val", "1").lower() not in {
        "0",
        "false",
        "off",
    }
    legend = root.find("c:chart/c:legend", ns)
    legend_position = legend.find("c:legendPos", ns) if legend is not None else None

    chart_tags = {
        f"{{{c}}}barChart",
        f"{{{c}}}lineChart",
        f"{{{c}}}pieChart",
        f"{{{c}}}areaChart",
        f"{{{c}}}scatterChart",
        f"{{{c}}}radarChart",
    }
    series_info = []
    series_order_issues = []
    series_order_values = []
    chart_type = None
    categories = None

    def cached_values(container: Optional[ET.Element]) -> List[str]:
        if container is None:
            return []
        return [
            (node.text or "").strip()
            for node in container.findall(".//c:pt/c:v", ns)
            if (node.text or "").strip()
        ]

    def series_name(series: ET.Element) -> Optional[str]:
        title = series.find("c:tx", ns)
        if title is None:
            return None
        values = cached_values(title)
        if values:
            return values[0]
        direct = title.find("c:v", ns)
        return (
            (direct.text or "").strip()
            if direct is not None and (direct.text or "").strip()
            else None
        )

    def canonical_element(element: ET.Element) -> tuple:
        """Return a stable saved-state signature for a DrawingML node."""
        return (
            element.tag.rsplit("}", 1)[-1],
            tuple(sorted(element.attrib.items())),
            (element.text or "").strip(),
            tuple(canonical_element(child) for child in element),
        )

    def series_line_fill_signature(series: ET.Element) -> tuple:
        line = series.find("c:spPr/a:ln", ns)
        if line is None:
            return ("inherited",)
        fill_tags = {
            f"{{{a}}}noFill",
            f"{{{a}}}solidFill",
            f"{{{a}}}gradFill",
            f"{{{a}}}pattFill",
        }
        fills = [canonical_element(child) for child in line if child.tag in fill_tags]
        return tuple(fills) if fills else ("inherited",)

    for chart in root.findall("c:chart/c:plotArea/*", ns):
        if chart.tag not in chart_tags:
            continue
        chart_type = chart_type or chart.tag.rsplit("}", 1)[-1]
        chart_labels = chart.find("c:dLbls", ns)
        for series in chart.findall("c:ser", ns):
            xml_position = len(series_info) + 1
            order_nodes = series.findall("c:order", ns)
            order_value = None
            order_raw = None
            if len(order_nodes) != 1:
                series_order_issues.append(
                    {
                        "xml_position": xml_position,
                        "reason": (
                            "missing_order" if not order_nodes else "duplicate_order_elements"
                        ),
                        "element_count": len(order_nodes),
                    }
                )
            else:
                order_raw = order_nodes[0].get("val")
                try:
                    order_value = int(order_raw)
                    if order_value < 0:
                        raise ValueError("negative chart series order")
                except (TypeError, ValueError):
                    series_order_issues.append(
                        {
                            "xml_position": xml_position,
                            "reason": "invalid_order_value",
                            "raw_value": order_raw,
                        }
                    )
                    order_value = None
            if order_value is not None:
                series_order_values.append(order_value)
            series_index = order_value + 1 if order_value is not None else None
            labels = series.find("c:dLbls", ns)
            if labels is None:
                labels = chart_labels
            show_value = labels.find("c:showVal", ns) if labels is not None else None
            position = labels.find("c:dLblPos", ns) if labels is not None else None
            has_labels = labels is not None and (
                show_value is None
                or show_value.get("val", "1").lower() not in {"0", "false", "off"}
            )
            series_categories = cached_values(series.find("c:cat", ns))
            if categories is None and series_categories:
                categories = series_categories
            numeric_values = cached_values(series.find("c:val", ns))
            trendlines = series.findall("c:trendline", ns)
            trendline_types = []
            for trendline in trendlines:
                trendline_type = trendline.find("c:trendlineType", ns)
                # OOXML defaults an omitted trendline type to linear.
                trendline_types.append(
                    trendline_type.get("val", "linear") if trendline_type is not None else "linear"
                )
            series_info.append(
                {
                    "index": series_index,
                    "saved_order": order_value,
                    "saved_order_raw": order_raw,
                    "xml_position": xml_position,
                    "name": series_name(series),
                    "categories": series_categories,
                    "values": [
                        float(value) if re.fullmatch(r"-?\d+(?:\.\d+)?", value) else value
                        for value in numeric_values
                    ],
                    "has_data_labels": has_labels,
                    "data_label_position": (position.get("val") if position is not None else None),
                    "trendline_count": len(trendlines),
                    "trendline_type": (trendline_types[0] if trendline_types else None),
                    "line_fill_signature": series_line_fill_signature(series),
                }
            )

    expected_order_values = list(range(len(series_info)))
    if (
        len(series_order_values) == len(series_info)
        and sorted(series_order_values) != expected_order_values
    ):
        series_order_issues.append(
            {
                "reason": "non_unique_or_non_contiguous_order_values",
                "observed": sorted(series_order_values),
                "expected": expected_order_values,
            }
        )
    series_info.sort(
        key=lambda series: (
            series.get("index") is None,
            series.get("index") or series.get("xml_position") or 0,
        )
    )
    style = root.find("c:style", ns)
    return {
        "chart_type": chart_type,
        "series_count": len(series_info),
        "series_order_valid": not series_order_issues,
        "series_order_values": sorted(series_order_values),
        "series_order_issues": series_order_issues,
        "series_info": series_info,
        "series_names": [series.get("name") for series in series_info],
        "series_values": [series.get("values") for series in series_info],
        "series_line_fills": [
            {
                "index": series.get("index"),
                "line_fill_signature": series.get("line_fill_signature"),
            }
            for series in series_info
        ],
        "categories": categories,
        "has_title": title is not None,
        "title_text": title_text,
        "title_include_in_layout": not title_is_overlay,
        "has_legend": legend is not None,
        "legend_position": (legend_position.get("val") if legend_position is not None else None),
        "chart_color": color_signature,
        "chart_style": style.get("val") if style is not None else None,
    }


def extract_insertion_signature(docx_path: Path) -> Dict[str, Any]:
    """Extract referenced chart data, cache, and workbook backing."""
    c = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    ns = {"c": c, "r": r, "w": w, "wp": wp}
    errors: List[str] = []
    charts = []
    with zipfile.ZipFile(docx_path) as package:
        names = set(package.namelist())
        document = ET.fromstring(package.read("word/document.xml"))
        document_relationships = part_relationships(package, "word/document.xml")
        references = document.findall(".//c:chart", ns)
        for reference in references:
            relationship_id = reference.get(f"{{{r}}}id")
            relationship = document_relationships.get(relationship_id or "")
            if not relationship or relationship.get("external"):
                errors.append("chart reference has no internal relationship")
                continue
            if not relationship.get("type", "").endswith("/chart"):
                errors.append("chart reference relationship has wrong type")
                continue
            chart_part = relationship.get("part")
            if not chart_part or chart_part not in names:
                errors.append("referenced chart part is missing")
                continue
            chart_root = ET.fromstring(package.read(chart_part))
            chart_relationships = part_relationships(package, chart_part)

            plot_kinds = []
            if chart_root.find(".//c:barChart", ns) is not None:
                bar = chart_root.find(".//c:barChart", ns)
                direction = bar.find("c:barDir", ns)
                grouping = bar.find("c:grouping", ns)
                if (
                    direction is not None
                    and direction.get("val") == "col"
                    and grouping is not None
                    and grouping.get("val") == "clustered"
                ):
                    plot_kinds.append("clustered_column")
                else:
                    plot_kinds.append("other_bar")
            if chart_root.find(".//c:lineChart", ns) is not None:
                plot_kinds.append("line")
            if chart_root.find(".//c:pieChart", ns) is not None:
                plot_kinds.append("pie")

            series = []
            for series_index, item in enumerate(
                chart_root.findall(".//c:plotArea/*/c:ser", ns), start=1
            ):

                def cache_points(
                    parent: Optional[ET.Element],
                    role: str,
                ) -> List[ET.Element]:
                    if parent is None:
                        return []
                    nodes = parent.findall(".//c:pt", ns)
                    if not nodes:
                        return []
                    try:
                        indexed = [(int(node.get("idx", "")), node) for node in nodes]
                    except (TypeError, ValueError):
                        errors.append(
                            f"series {series_index} {role} cache has an " "invalid point index"
                        )
                        return nodes
                    indexed.sort(key=lambda row: row[0])
                    indexes = [row[0] for row in indexed]
                    if indexes != list(range(len(indexed))):
                        errors.append(
                            f"series {series_index} {role} cache point "
                            "indexes are incomplete or duplicated"
                        )
                    point_count = parent.find(".//c:ptCount", ns)
                    try:
                        declared_count = int(point_count.get("val", ""))
                    except (AttributeError, TypeError, ValueError):
                        declared_count = None
                    if declared_count != len(indexed):
                        errors.append(
                            f"series {series_index} {role} cache point "
                            "count is missing or inconsistent"
                        )
                    return [row[1] for row in indexed]

                name_node = item.find("c:tx/c:v", ns)
                if name_node is None:
                    name_cache = item.find("c:tx/c:strRef/c:strCache", ns)
                    name_nodes = cache_points(name_cache, "name")
                    name_node = name_nodes[0].find("c:v", ns) if name_nodes else None

                def formula(path: str) -> Optional[str]:
                    node = item.find(path, ns)
                    if node is None or not (node.text or "").strip():
                        return None
                    return (node.text or "").strip()

                def points(parent_name: str) -> Optional[List[Any]]:
                    parent = item.find(f"c:{parent_name}", ns)
                    if parent is None:
                        return None
                    nodes = cache_points(parent, parent_name)
                    if not nodes:
                        return None
                    return [
                        canonical_value(
                            node.find("c:v", ns).text if node.find("c:v", ns) is not None else None
                        )
                        for node in nodes
                    ]

                series.append(
                    {
                        "name": ((name_node.text or "").strip() if name_node is not None else None),
                        "categories": points("cat"),
                        "values": points("val"),
                        "formulas": {
                            "name": formula("c:tx/c:strRef/c:f"),
                            "categories": (
                                formula("c:cat/c:strRef/c:f") or formula("c:cat/c:numRef/c:f")
                            ),
                            "values": formula("c:val/c:numRef/c:f"),
                        },
                    }
                )

            external_data = chart_root.find("c:externalData", ns)
            workbook_part = None
            workbook_grid = None
            workbook_cells = None
            if external_data is not None:
                workbook_id = external_data.get(f"{{{r}}}id")
                workbook_relationship = chart_relationships.get(workbook_id or "")
                if (
                    not workbook_relationship
                    or workbook_relationship.get("external")
                    or not workbook_relationship.get("type", "").endswith("/package")
                ):
                    errors.append("chart workbook relationship is invalid")
                else:
                    workbook_part = workbook_relationship.get("part")
                    if not workbook_part or workbook_part not in names:
                        errors.append("embedded chart workbook is missing")
                    else:
                        try:
                            workbook_payload = package.read(workbook_part)
                            workbook_cells = extract_workbook_cells(workbook_payload)
                            workbook_grid = extract_workbook_grid(workbook_payload)
                        except Exception as exc:
                            errors.append(f"embedded chart workbook is invalid: {exc}")
            charts.append(
                {
                    "kind": plot_kinds[0] if len(plot_kinds) == 1 else None,
                    "plot_kinds": plot_kinds,
                    "series": series,
                    "workbook_part": workbook_part,
                    "workbook_grid": workbook_grid,
                    "workbook_cells": workbook_cells,
                }
            )
        extents = [
            {
                "width_emu": int(node.get("cx", "0")),
                "height_emu": int(node.get("cy", "0")),
            }
            for node in document.findall(".//wp:extent", ns)
        ]
        page_size = document.find(".//w:sectPr/w:pgSz", ns)
        page_margins = document.find(".//w:sectPr/w:pgMar", ns)
        page_width_emu = (
            int(page_size.get(f"{{{w}}}w", "0")) * 635 if page_size is not None else None
        )
        page_height_emu = (
            int(page_size.get(f"{{{w}}}h", "0")) * 635 if page_size is not None else None
        )
        printable_width_emu = page_width_emu
        printable_height_emu = page_height_emu
        if page_margins is not None and page_width_emu and page_height_emu:
            printable_width_emu -= (
                int(page_margins.get(f"{{{w}}}left", "0"))
                + int(page_margins.get(f"{{{w}}}right", "0"))
            ) * 635
            printable_height_emu -= (
                int(page_margins.get(f"{{{w}}}top", "0"))
                + int(page_margins.get(f"{{{w}}}bottom", "0"))
            ) * 635
        return {
            "object_count": len(references),
            "body_object_count": sum(
                1
                for node in document.iter()
                if node.tag in {f"{{{w}}}drawing", f"{{{w}}}pict", f"{{{w}}}object"}
            ),
            "charts": charts,
            "errors": errors,
            "extents": extents,
            "page_width_emu": page_width_emu,
            "page_height_emu": page_height_emu,
            "printable_width_emu": printable_width_emu,
            "printable_height_emu": printable_height_emu,
        }
