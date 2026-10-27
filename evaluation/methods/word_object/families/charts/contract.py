"""Chart data preservation and required observations shared by effect routes."""
from ...observations import require_observations
from ....base import VerificationResult
from .workbook import canonical_value


COMMAND_FIELDS = {
    "ChartStylesGallery": ("chart_style",),
    "ChartColorsGallery": ("chart_color",),
    "ChartLayoutGallery": ("has_title", "has_legend", "has_data_table", "axes"),
    "ChartLegend": ("has_legend",),
    "ChartTitle": ("has_title",),
    "ChartDataTable": ("has_data_table",),
    "ChartUpDownBars": ("has_up_down_bars",),
    "ChartSwitchRowColumn": ("chart_type", "chart_color"),
    "ChartResetToMatchStyle": ("chart_color",),
    "ChartGridlinesGallery": ("axes",),
    "ChartLines": ("has_drop_lines", "has_hi_lo_lines"),
}

SERIES_FIELDS = {
    "ChartDataLabel": ("has_data_labels",),
    "ChartLayoutGallery": ("has_data_labels",),
    "ChartTrendline": ("trendline_count",),
    "ChartErrorBars": ("has_error_bars",),
    "ColorPickerChartLineStyles": ("line_color_rgb",),
}


def validate_chart_observations(gold, pred, commands, spec=None):
    """None means unreadable; an observed empty collection is handled separately."""
    fields = {"series_info", "chart_type"}
    for command in commands:
        fields.update(COMMAND_FIELDS.get(command, ()))
    operation_fields = {
        "add_drop_lines": ("has_drop_lines",),
        "add_hi_lo_lines": ("has_hi_lo_lines",),
        "add_up_down_bars": ("has_up_down_bars",),
        "add_chart_title": ("has_title",),
        "format_chart_title": ("has_title", "title_text"),
    }
    title_operation = False
    for step in (spec or {}).get("composition", {}).get("steps", []) or []:
        operation = step.get("operation")
        fields.update(operation_fields.get(operation, ()))
        if operation in {"add_chart_title", "format_chart_title"}:
            title_operation = True
        if operation == "format_chart_title" and (step.get("params") or {}).get("bold") is not None:
            fields.add("title_font_bold")
    for owner, state in (("gold", gold), ("pred", pred)):
        if "axes" in fields and state.get("axes_unavailable"):
            return VerificationResult.error_result("Chart axes could not be read", observation_owner=owner,
                                                   unavailable={"axes": state["axes_unavailable"]})
        error = require_observations({owner: state}, fields, context="chart")
        if error is not None:
            return error
        if "axes" in fields:
            if owner == "gold" and not state["axes"]:
                return VerificationResult.error_result("Gold chart has no required axes", observation_owner=owner)
            for axis in state["axes"]:
                error = require_observations({owner: axis},
                    ("axis_type", "has_major_gridlines", "has_minor_gridlines"), context="chart axis")
                if error is not None:
                    return error
        conditional = []
        if ("ChartTitle" in commands or "ChartLayoutGallery" in commands or title_operation) and state.get("has_title"):
            conditional += ["title_text", "title_include_in_layout"]
        if ("ChartLegend" in commands or "ChartLayoutGallery" in commands) and state.get("has_legend"):
            conditional += ["legend_position"]
        if "ChartDataTable" in commands and state.get("has_data_table"):
            conditional += ["data_table_show_legend_key"]
        error = require_observations({owner: state}, conditional, context="active chart element")
        if error is not None:
            return error
        series = state["series_info"]
        if not isinstance(series, list):
            return VerificationResult.error_result("Unreadable chart series collection", observation_owner=owner)
        if owner == "gold" and not series:
            return VerificationResult.error_result("Gold chart has no data series", observation_owner=owner)
        for index, item in enumerate(series, 1):
            required = {"name", "values", "categories"}
            for command in commands:
                if command != "ColorPickerChartLineStyles" or index == 1:
                    required.update(SERIES_FIELDS.get(command, ()))
            if "has_data_labels" in required and item.get("has_data_labels"):
                required.add("data_label_position")
            if "trendline_count" in required and item.get("trendline_count"):
                required.add("trendline_type")
            error = require_observations({owner: item}, required, context=f"chart series {index}")
            if error is not None:
                return error
    return None


def diff_data(gold, pred):
    """Compare final series semantics, preserving order and numeric zero.

    No supported effect command edits data. SwitchRowColumn and compositions
    compare the final reference projection, so intentional transposition is kept.
    """
    differences = []
    gold_series, pred_series = gold["series_info"], pred["series_info"]
    for field in ("chart_type", "chart_count"):
        if field in gold and gold.get(field) != pred.get(field):
            differences.append({"property": field, "gold": gold.get(field), "pred": pred.get(field)})
    if len(gold_series) != len(pred_series):
        differences.append({"property": "series_count", "gold": len(gold_series), "pred": len(pred_series)})
    for index, (g, p) in enumerate(zip(gold_series, pred_series), 1):
        for field in ("name", "values", "categories"):
            gv, pv = g.get(field), p.get(field)
            if field == "values":
                gv = [canonical_value(v) for v in gv]
                pv = [canonical_value(v) for v in pv]
            if gv != pv:
                differences.append({"property": f"series_{index}_{field}", "gold": gv, "pred": pv})
    return differences
