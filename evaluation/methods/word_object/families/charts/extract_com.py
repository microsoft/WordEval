"""Chart COM readers using the caller-owned document lifecycle.

Documents are opened, read, and closed in the original order. This module does
not own a Word process or retain live COM objects between calls.
"""

import os
import tempfile
from typing import Any, Callable, Dict
from ...observations import ObservedProperties


def read_primary_axes(chart):
    """Read existing primary category/value axes without dropping failed reads.

    Chart.Axes() is the collection API, as documented at
    https://learn.microsoft.com/en-us/office/vba/api/word.chart.axes
    """
    count = int(chart.Axes().Count)
    # Word's nested Axes dispatch can expose Count/Item but no IEnumVARIANT
    # through pywin32. Resolve the documented type/group slots, then require
    # their number to agree with Count: a failed existing axis cannot disappear.
    members = {}
    for group in (1, 2):
        for kind in (1, 2, 3):
            if len(members) == count:
                break
            try:
                axis = chart.Axes(kind, group)
                identity = (axis.AxisGroup, axis.Type)
            except Exception:
                continue  # A slot can be absent; Count proves completeness below.
            if identity[0] not in (1, 2) or identity[1] not in (1, 2, 3):
                raise ValueError("Chart axis identity is unavailable")
            members[identity] = axis
    if len(members) != count:
        raise ValueError("Chart axis enumeration is incomplete")
    axes = []
    for (group, kind), axis in members.items():
        if group != 1 or kind not in (1, 2):
            continue
        state = ObservedProperties(axis_type=kind)
        state.read("has_major_gridlines", lambda: axis.HasMajorGridlines, bool)
        state.read("has_minor_gridlines", lambda: axis.HasMinorGridlines, bool)
        axes.append(state)
    return sorted(axes, key=lambda axis: axis["axis_type"])


def extract_properties(
    docx_path: str,
    *,
    open_document: Callable,
    close_document: Callable,
) -> Dict:
    """
    Extract chart properties from the first chart in document.

    Returns dict with chart properties:
    - chart_style, chart_color
    - has_legend, legend_position, legend_include_in_layout
    - has_title, title_text, title_include_in_layout
    - has_data_table, data_table_show_legend_key
    - series_info: per-series data labels, trendlines, error bars, line color
    - axes: axis type, gridlines
    - has_drop_lines, has_hi_lo_lines, has_up_down_bars
    """
    doc = None
    try:
        doc = open_document(docx_path)

        matches = []
        for collection_name in ("InlineShapes", "Shapes"):
            collection = getattr(doc, collection_name)
            for index in range(1, collection.Count + 1):
                candidate = collection(index)
                if candidate.HasChart:
                    matches.append((candidate, "inline" if collection_name == "InlineShapes" else "floating"))
        if not matches:
            return {"error": "No chart shape found", "absent": True}
        shape, source = matches[0]
        chart = shape.Chart
        result = {"source": source, "chart_count": len(matches)}

        # Chart style and color
        try:
            result["chart_style"] = chart.ChartStyle
        except:
            result["chart_style"] = None
        try:
            result["chart_color"] = chart.ChartColor
        except:
            result["chart_color"] = None

        # Chart type
        try:
            result["chart_type"] = chart.ChartType
        except:
            result["chart_type"] = None

        # Legend
        try:
            result["has_legend"] = bool(chart.HasLegend)
            if result["has_legend"]:
                result["legend_position"] = chart.Legend.Position
                try:
                    result["legend_include_in_layout"] = bool(chart.Legend.IncludeInLayout)
                except:
                    result["legend_include_in_layout"] = None
            else:
                result["legend_position"] = None
                result["legend_include_in_layout"] = None
        except:
            result["has_legend"] = None
            result["legend_position"] = None
            result["legend_include_in_layout"] = None

        # Title
        try:
            result["has_title"] = bool(chart.HasTitle)
            if result["has_title"]:
                result["title_text"] = chart.ChartTitle.Text
                try:
                    result["title_include_in_layout"] = bool(chart.ChartTitle.IncludeInLayout)
                except:
                    result["title_include_in_layout"] = None
                try:
                    result["title_font_bold"] = bool(
                        chart.ChartTitle.Format.TextFrame2.TextRange.Font.Bold
                    )
                except:
                    try:
                        result["title_font_bold"] = bool(chart.ChartTitle.Font.Bold)
                    except:
                        result["title_font_bold"] = None
            else:
                result["title_text"] = None
                result["title_include_in_layout"] = None
                result["title_font_bold"] = None
        except:
            result["has_title"] = None
            result["title_text"] = None
            result["title_include_in_layout"] = None
            result["title_font_bold"] = None

        # Data table
        try:
            result["has_data_table"] = bool(chart.HasDataTable)
            if result["has_data_table"]:
                try:
                    result["data_table_show_legend_key"] = bool(chart.DataTable.ShowLegendKey)
                except:
                    result["data_table_show_legend_key"] = None
            else:
                result["data_table_show_legend_key"] = None
        except:
            result["has_data_table"] = None
            result["data_table_show_legend_key"] = None

        # Series info (data labels, trendlines, error bars, line color)
        series_info = []
        categories = None

        def com_sequence(value):
            """Normalize scalar/SafeArray chart data returned by COM."""
            if value is None:
                return None
            if isinstance(value, (tuple, list)):
                return [item.item() if hasattr(item, "item") else item for item in value]
            return [value.item() if hasattr(value, "item") else value]

        for si in range(1, chart.SeriesCollection().Count + 1):
            series = chart.SeriesCollection(si)
            s_info = {"index": si}

            # Persisted chart data is available through the Word/
            # embedded-chart object model without activating or
            # mutating the workbook.  Keep these WOM values as the
            # semantic primary source for chart insertion tasks.
            try:
                name = series.Name
                s_info["name"] = None if name is None else str(name)
            except:
                s_info["name"] = None
            try:
                s_info["values"] = com_sequence(series.Values)
            except:
                s_info["values"] = None
            try:
                series_categories = com_sequence(series.XValues)
            except:
                series_categories = None
            s_info["categories"] = series_categories
            if categories is None and series_categories:
                categories = series_categories

            # Data labels
            try:
                s_info["has_data_labels"] = bool(series.HasDataLabels)
                if s_info["has_data_labels"]:
                    try:
                        s_info["data_label_position"] = series.DataLabels().Position
                    except:
                        s_info["data_label_position"] = None
                else:
                    s_info["data_label_position"] = None
            except:
                s_info["has_data_labels"] = None
                s_info["data_label_position"] = None

            # Trendlines
            try:
                s_info["trendline_count"] = series.Trendlines().Count
                if s_info["trendline_count"] > 0:
                    s_info["trendline_type"] = series.Trendlines(1).Type
                else:
                    s_info["trendline_type"] = None
            except:
                s_info["trendline_count"] = None
                s_info["trendline_type"] = None

            # Error bars
            # Note: ErrorBars.Type is NOT accessible via COM after creation.
            # The type (percentage, standard_error, etc.) is set via Series.ErrorBar()
            # but cannot be read back. Only has_error_bars is reliable.
            try:
                s_info["has_error_bars"] = bool(series.HasErrorBars)
            except:
                s_info["has_error_bars"] = None

            # Line color (for ColorPickerChartLineStyles)
            try:
                s_info["line_color_rgb"] = series.Format.Line.ForeColor.RGB
            except:
                s_info["line_color_rgb"] = None

            # Fill color (for ChartResetToMatchStyle detection)
            try:
                s_info["fill_color_rgb"] = series.Format.Fill.ForeColor.RGB
            except:
                s_info["fill_color_rgb"] = None

            series_info.append(s_info)
        result["series_info"] = series_info
        result["series_count"] = len(series_info)
        result["series_names"] = [series.get("name") for series in series_info]
        result["series_values"] = [series.get("values") for series in series_info]
        result["categories"] = categories

        # A failed read must not silently remove one axis from both projections.
        # Enumerating the collection distinguishes real absence (e.g. a pie)
        # from an unavailable axis or gridline property.
        try:
            result["axes"] = read_primary_axes(chart)
        except Exception as exc:
            result["axes"] = None
            result["axes_unavailable"] = str(exc)

        # Chart group features (drop lines, hi-lo lines, up/down bars)
        try:
            cg = chart.ChartGroups(1)
            result["has_drop_lines"] = bool(cg.HasDropLines)
        except:
            result["has_drop_lines"] = None
        try:
            cg = chart.ChartGroups(1)
            result["has_hi_lo_lines"] = bool(cg.HasHiLoLines)
        except:
            result["has_hi_lo_lines"] = None
        try:
            cg = chart.ChartGroups(1)
            result["has_up_down_bars"] = bool(cg.HasUpDownBars)
        except:
            result["has_up_down_bars"] = None

        return result

    finally:
        close_document(doc)


def extract_inserted_state(
    docx_path: str,
    *,
    open_document: Callable,
    close_document: Callable,
    strict_sequence: Callable,
    png_nonwhite_ratio: Callable,
) -> Dict[str, Any]:
    """Strict WOM state for the atomic chart-insertion contract."""
    doc = None
    try:
        doc = open_document(docx_path)
        inline_count = int(doc.InlineShapes.Count)
        floating_count = int(doc.Shapes.Count)
        matches = []
        for collection_name, object_kind in (
            ("InlineShapes", "inline"),
            ("Shapes", "floating"),
        ):
            collection = getattr(doc, collection_name)
            for index in range(1, int(collection.Count) + 1):
                candidate = collection(index)
                if bool(candidate.HasChart):
                    matches.append((candidate, object_kind))

        result: Dict[str, Any] = {
            "chart_count": len(matches),
            "visual_object_count": inline_count + floating_count,
        }
        if len(matches) != 1:
            return result

        shape, source = matches[0]
        width = float(shape.Width)
        height = float(shape.Height)
        page_setup = doc.PageSetup
        printable_width = float(
            page_setup.PageWidth - page_setup.LeftMargin - page_setup.RightMargin
        )
        printable_height = float(
            page_setup.PageHeight - page_setup.TopMargin - page_setup.BottomMargin
        )
        chart = shape.Chart
        with tempfile.TemporaryDirectory(prefix="word-chart-render-") as temp_dir:
            chart_render_path = os.path.join(temp_dir, "chart.png")
            if not bool(chart.Export(chart_render_path, "PNG")):
                raise ValueError("Word could not export the reopened chart")
            render = png_nonwhite_ratio(chart_render_path)
        collection = chart.SeriesCollection()
        series = []
        for index in range(1, int(collection.Count) + 1):
            item = collection(index)
            series.append(
                {
                    "name": str(item.Name),
                    "values": strict_sequence(item.Values),
                    "categories": strict_sequence(item.XValues),
                }
            )
        result.update(
            {
                "source": source,
                "chart_type": int(chart.ChartType),
                "series": series,
                "width_points": round(width, 2),
                "height_points": round(height, 2),
                "printable_width_points": round(printable_width, 2),
                "printable_height_points": round(printable_height, 2),
                "visible_on_page": (
                    width > 0
                    and height > 0
                    and width <= printable_width
                    and height <= printable_height
                ),
                "render_visible": render["visible"],
                "render_width_pixels": render["width"],
                "render_height_pixels": render["height"],
                "render_nonwhite_pixels": render["nonwhite_pixels"],
                "render_nonwhite_ratio": render["nonwhite_ratio"],
            }
        )
        return result
    finally:
        close_document(doc)
