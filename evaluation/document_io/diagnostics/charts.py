"""Structural diagnostics extracted without changing rule order."""
import io
import xml.etree.ElementTree as ET
from pathlib import PurePosixPath
from .common import (w_ns, r_ns, m_ns, c_ns, a_ns, cs_ns, wp_ns, wps_ns, a14_ns, mc_ns, invalid_schema, invalid_child_order)

def diagnose(xml_roots, include_word_open_failure_schema=False):
    for part_name, root in xml_roots.items():
        if not part_name.startswith("word/charts/"):
            continue
        if include_word_open_failure_schema:
            color_choice_tags = {
                f"{{{a_ns}}}{local_name}"
                for local_name in (
                    "scrgbClr", "srgbClr", "hslClr", "sysClr",
                    "schemeClr", "prstClr",
                )
            }
            for color_style in root.iter(f"{{{cs_ns}}}colorStyle"):
                direct_colors = [
                    child for child in color_style
                    if child.tag in color_choice_tags
                ]
                if not direct_colors:
                    return invalid_schema(
                        part_name,
                        "cs:colorStyle",
                        (
                            "pred.docx has an invalid chart color style: "
                            "cs:colorStyle must contain at least one "
                            "direct DrawingML color before any "
                            "cs:variation elements."
                        ),
                    )
                for variation in color_style.findall(
                    f"{{{cs_ns}}}variation"
                ):
                    invalid_color = next(
                        (
                            child for child in variation
                            if child.tag in color_choice_tags
                        ),
                        None,
                    )
                    if invalid_color is not None:
                        local_name = invalid_color.tag.rsplit("}", 1)[-1]
                        return invalid_schema(
                            part_name,
                            f"cs:variation/a:{local_name}",
                            (
                                "pred.docx places a DrawingML color "
                                f"a:{local_name} inside cs:variation; "
                                "chart color variations may contain "
                                "only color transforms."
                            ),
                        )
            required_chart_parents = {
                f"{{{c_ns}}}legend": (
                    f"{{{c_ns}}}chart", "c:legend", "c:chart",
                ),
                f"{{{c_ns}}}autoTitleDeleted": (
                    f"{{{c_ns}}}chart", "c:autoTitleDeleted", "c:chart",
                ),
                f"{{{c_ns}}}txBehindChart": (
                    f"{{{c_ns}}}chart", "c:txBehindChart", "c:chart",
                ),
                f"{{{c_ns}}}printSettings": (
                    f"{{{c_ns}}}chartSpace", "c:printSettings",
                    "c:chartSpace",
                ),
            }
            for parent in root.iter():
                for child in parent:
                    parent_rule = required_chart_parents.get(child.tag)
                    if parent_rule is None:
                        continue
                    expected_parent, child_label, parent_label = parent_rule
                    if parent.tag == expected_parent:
                        continue
                    actual_parent = parent.tag.rsplit("}", 1)[-1]
                    return invalid_schema(
                        part_name,
                        child_label,
                        (
                            f"pred.docx places {child_label} directly "
                            f"inside c:{actual_parent}; it must be a "
                            f"direct child of {parent_label}."
                        ),
                    )
        for chart in (
            root.iter(f"{{{c_ns}}}chart")
            if include_word_open_failure_schema else ()
        ):
            order_error = invalid_child_order(
                part_name,
                chart,
                "c:chart",
                (
                    ("title",), ("autoTitleDeleted",), ("pivotFmts",),
                    ("view3D",), ("floor",), ("sideWall",),
                    ("backWall",), ("plotArea",), ("legend",),
                    ("plotVisOnly",), ("dispBlanksAs",),
                    ("showDLblsOverMax",), ("extLst",),
                ),
                c_ns,
            )
            if order_error:
                return order_error
        for bar_chart in (
            root.iter(f"{{{c_ns}}}barChart")
            if include_word_open_failure_schema else ()
        ):
            for series in bar_chart.findall(f"{{{c_ns}}}ser"):
                order_error = invalid_child_order(
                    part_name,
                    series,
                    "c:barChart/c:ser",
                    (
                        ("idx",), ("order",), ("tx",), ("spPr",),
                        ("invertIfNegative",), ("pictureOptions",),
                        ("dPt",), ("dLbls",), ("trendline",),
                        ("errBars",), ("cat",), ("val",),
                        ("shape",), ("extLst",),
                    ),
                    c_ns,
                )
                if order_error:
                    return order_error
        for trendline in (
            root.iter(f"{{{c_ns}}}trendline")
            if include_word_open_failure_schema else ()
        ):
            order_error = invalid_child_order(
                part_name,
                trendline,
                "c:trendline",
                (
                    ("name",), ("spPr",), ("trendlineType",),
                    ("order",), ("period",), ("forward",),
                    ("backward",), ("intercept",), ("dispRSqr",),
                    ("dispEq",), ("trendlineLbl",), ("extLst",),
                ),
                c_ns,
            )
            if order_error:
                return order_error
        for value_axis in (
            root.iter(f"{{{c_ns}}}valAx")
            if include_word_open_failure_schema else ()
        ):
            order_error = invalid_child_order(
                part_name,
                value_axis,
                "c:valAx",
                (
                    ("axId",), ("scaling",), ("delete",), ("axPos",),
                    ("majorGridlines",), ("minorGridlines",), ("title",),
                    ("numFmt",), ("majorTickMark",), ("minorTickMark",),
                    ("tickLblPos",), ("spPr",), ("txPr",), ("crossAx",),
                    ("crosses", "crossesAt"), ("crossBetween",),
                    ("majorUnit",), ("minorUnit",), ("dispUnits",),
                    ("extLst",),
                ),
                c_ns,
            )
            if order_error:
                return order_error
        for line_chart in root.iter(f"{{{c_ns}}}lineChart"):
            if include_word_open_failure_schema:
                order_error = invalid_child_order(
                    part_name,
                    line_chart,
                    "c:lineChart",
                    (
                        ("grouping",), ("varyColors",), ("ser",),
                        ("dLbls",), ("dropLines",), ("hiLowLines",),
                        ("upDownBars",), ("marker",), ("smooth",),
                        ("axId",), ("extLst",),
                    ),
                    c_ns,
                )
                if order_error:
                    return order_error
                for series in line_chart.findall(f"{{{c_ns}}}ser"):
                    order_error = invalid_child_order(
                        part_name,
                        series,
                        "c:lineChart/c:ser",
                        (
                            ("idx",), ("order",), ("tx",), ("spPr",),
                            ("marker",), ("dPt",), ("dLbls",),
                            ("trendline",), ("errBars",), ("cat",),
                            ("val",), ("smooth",), ("extLst",),
                        ),
                        c_ns,
                    )
                    if order_error:
                        return order_error
            # Word can open a line chart with a leftover barDir and
            # expose the requested line-chart semantics normally.
            # Diagnose this schema violation only after an attributed
            # Word open failure; it is not proof of an unusable file.
            if (
                include_word_open_failure_schema
                and line_chart.find(f"{{{c_ns}}}barDir") is not None
            ):
                return invalid_schema(
                    part_name,
                    "c:lineChart/c:barDir",
                    (
                        "pred.docx places bar-chart-only c:barDir inside "
                        "c:lineChart."
                    ),
                )
            marker = line_chart.find(f"{{{c_ns}}}marker")
            if marker is not None and len(marker):
                return invalid_schema(
                    part_name,
                    "c:lineChart/c:marker",
                    (
                        "pred.docx gives the leaf c:lineChart/c:marker "
                        "element child content."
                    ),
                )

    return None
