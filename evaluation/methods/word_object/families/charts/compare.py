"""Pure chart comparisons, scoped to the requested command or operation."""

from typing import Any, Dict, List, Optional

from .workbook import canonical_value, column_name, parse_range_formula, project_workbook


def diff_command(
    idMso: str,
    gold_chart: Dict,
    pred_chart: Dict,
    spec: Optional[Dict] = None,
) -> List[Dict]:
    """Per-idMso chart property comparison (shared by atomic and composition).

    Returns the list of property diffs that the given atomic command
    legitimately tests. Composition tests reuse this so they check the SAME
    properties as the equivalent atomic test — see N4 in
    the composition review findings.
    """
    differences: List[Dict] = []

    def _series_count_differs(gold_series, pred_series) -> bool:
        # zip() silently truncates: a prediction that dropped or added a
        # series must fail on the count, never skip the unmatched tail.
        if len(gold_series) != len(pred_series):
            differences.append(
                {
                    "property": "series_count",
                    "gold": len(gold_series),
                    "pred": len(pred_series),
                }
            )
            return True
        return False

    # Command-specific property checks
    if idMso == "ChartStylesGallery":
        if gold_chart.get("chart_style") != pred_chart.get("chart_style"):
            differences.append(
                {
                    "property": "chart_style",
                    "gold": gold_chart.get("chart_style"),
                    "pred": pred_chart.get("chart_style"),
                }
            )

    elif idMso == "ChartColorsGallery":
        if gold_chart.get("chart_color") != pred_chart.get("chart_color"):
            differences.append(
                {
                    "property": "chart_color",
                    "gold": gold_chart.get("chart_color"),
                    "pred": pred_chart.get("chart_color"),
                }
            )

    elif idMso == "ChartLayoutGallery":
        # Layout affects title, legend, data labels, gridlines — compare all
        for prop in ["has_title", "has_legend", "has_data_table", "legend_position"]:
            if gold_chart.get(prop) != pred_chart.get(prop):
                differences.append(
                    {
                        "property": prop,
                        "gold": gold_chart.get(prop),
                        "pred": pred_chart.get(prop),
                    }
                )
        # Compare axes gridlines (same zip-truncation hazard as series)
        gold_axes = gold_chart.get("axes", [])
        pred_axes = pred_chart.get("axes", [])
        if len(gold_axes) != len(pred_axes):
            differences.append(
                {
                    "property": "axes_count",
                    "gold": len(gold_axes),
                    "pred": len(pred_axes),
                }
            )
        for ga, pa in zip(gold_axes, pred_axes):
            for prop in ["has_major_gridlines", "has_minor_gridlines"]:
                if ga.get(prop) != pa.get(prop):
                    differences.append(
                        {
                            "property": f"axis_{ga.get('axis_type')}_{prop}",
                            "gold": ga.get(prop),
                            "pred": pa.get(prop),
                        }
                    )
        # Compare series data labels
        gold_series = gold_chart.get("series_info", [])
        pred_series = pred_chart.get("series_info", [])
        _series_count_differs(gold_series, pred_series)
        for gs, ps in zip(gold_series, pred_series):
            if gs.get("has_data_labels") != ps.get("has_data_labels"):
                differences.append(
                    {
                        "property": f"series_{gs.get('index')}_has_data_labels",
                        "gold": gs.get("has_data_labels"),
                        "pred": ps.get("has_data_labels"),
                    }
                )

    elif idMso == "ChartLegend":
        for prop in ["has_legend", "legend_position"]:
            if gold_chart.get(prop) != pred_chart.get(prop):
                differences.append(
                    {
                        "property": prop,
                        "gold": gold_chart.get(prop),
                        "pred": pred_chart.get(prop),
                    }
                )

    elif idMso == "ChartTitle":
        props = ["has_title", "title_text", "title_include_in_layout"]
        for prop in props:
            if gold_chart.get(prop) != pred_chart.get(prop):
                differences.append(
                    {
                        "property": prop,
                        "gold": gold_chart.get(prop),
                        "pred": pred_chart.get(prop),
                    }
                )

    elif idMso == "ChartDataLabel":
        gold_series = gold_chart.get("series_info", [])
        pred_series = pred_chart.get("series_info", [])
        _series_count_differs(gold_series, pred_series)
        for gs, ps in zip(gold_series, pred_series):
            if gs.get("has_data_labels") != ps.get("has_data_labels"):
                differences.append(
                    {
                        "property": f"series_{gs.get('index')}_has_data_labels",
                        "gold": gs.get("has_data_labels"),
                        "pred": ps.get("has_data_labels"),
                    }
                )
            if gs.get("data_label_position") != ps.get("data_label_position"):
                differences.append(
                    {
                        "property": f"series_{gs.get('index')}_data_label_position",
                        "gold": gs.get("data_label_position"),
                        "pred": ps.get("data_label_position"),
                    }
                )

    elif idMso == "ChartDataTable":
        for prop in ["has_data_table", "data_table_show_legend_key"]:
            if gold_chart.get(prop) != pred_chart.get(prop):
                differences.append(
                    {
                        "property": prop,
                        "gold": gold_chart.get(prop),
                        "pred": pred_chart.get(prop),
                    }
                )

    elif idMso == "ChartTrendline":
        gold_series = gold_chart.get("series_info", [])
        pred_series = pred_chart.get("series_info", [])
        _series_count_differs(gold_series, pred_series)
        for gs, ps in zip(gold_series, pred_series):
            if gs.get("trendline_count") != ps.get("trendline_count"):
                differences.append(
                    {
                        "property": f"series_{gs.get('index')}_trendline_count",
                        "gold": gs.get("trendline_count"),
                        "pred": ps.get("trendline_count"),
                    }
                )
            if gs.get("trendline_type") != ps.get("trendline_type"):
                differences.append(
                    {
                        "property": f"series_{gs.get('index')}_trendline_type",
                        "gold": gs.get("trendline_type"),
                        "pred": ps.get("trendline_type"),
                    }
                )

    elif idMso == "ChartErrorBars":
        # Note: error_bar_type is NOT readable via COM (set at creation, not queryable).
        # Only has_error_bars can be verified. openxml_comparison handles full verification.
        gold_series = gold_chart.get("series_info", [])
        pred_series = pred_chart.get("series_info", [])
        _series_count_differs(gold_series, pred_series)
        for gs, ps in zip(gold_series, pred_series):
            if gs.get("has_error_bars") != ps.get("has_error_bars"):
                differences.append(
                    {
                        "property": f"series_{gs.get('index')}_has_error_bars",
                        "gold": gs.get("has_error_bars"),
                        "pred": ps.get("has_error_bars"),
                    }
                )

    elif idMso == "ChartGridlinesGallery":
        gold_axes = gold_chart.get("axes", [])
        pred_axes = pred_chart.get("axes", [])
        parameter = dict((spec or {}).get("parameter") or {})
        if not parameter:
            preset = {"major_horizontal": (2, "major"), "minor_horizontal": (2, "minor"),
                      "major_vertical": (1, "major"), "minor_vertical": (1, "minor")}.get((spec or {}).get("variation"))
            if preset is not None:
                parameter = {"axis_type": preset[0], "gridline_type": preset[1]}
        if not parameter:
            for step in (spec or {}).get("composition", {}).get("steps", []) or []:
                if step.get("idMso") == "ChartGridlinesGallery":
                    parameter = dict(step.get("params") or {})
                    break

        axis_type = parameter.get("axis_type")
        axis_name = parameter.get("axis")
        gridline_type = parameter.get("gridline_type") or parameter.get("gridline")
        if axis_type is None and axis_name is not None:
            axis_type = {"category": 1, "value": 2}.get(str(axis_name).lower())

        if axis_type is None or gridline_type is None:
            for intent in (spec or {}).get("composition", {}).get("atomic_intents", []) or []:
                target_properties = {
                    target.get("property")
                    for target in intent.get("verification_targets", []) or []
                }
                if "value_axis_major_gridlines" in target_properties:
                    axis_type, gridline_type = 2, "major"
                    break
                if "value_axis_minor_gridlines" in target_properties:
                    axis_type, gridline_type = 2, "minor"
                    break

        requested_prop = {
            "major": "has_major_gridlines",
            "minor": "has_minor_gridlines",
        }.get(str(gridline_type).lower())
        if axis_type is not None and requested_prop is not None:
            gold_axis = next(
                (axis for axis in gold_axes if axis.get("axis_type") == axis_type),
                None,
            )
            pred_axis = next(
                (axis for axis in pred_axes if axis.get("axis_type") == axis_type),
                None,
            )
            gold_value = gold_axis.get(requested_prop) if gold_axis else None
            pred_value = pred_axis.get(requested_prop) if pred_axis else None
            if gold_value != pred_value:
                differences.append(
                    {
                        "property": f"axis_{axis_type}_{requested_prop}",
                        "gold": gold_value,
                        "pred": pred_value,
                    }
                )
        else:
            # Unknown legacy spec: retain conservative reference matching.
            if len(gold_axes) != len(pred_axes):
                differences.append(
                    {
                        "property": "axes_count",
                        "gold": len(gold_axes),
                        "pred": len(pred_axes),
                    }
                )
            for ga, pa in zip(gold_axes, pred_axes):
                for prop in ["has_major_gridlines", "has_minor_gridlines"]:
                    if ga.get(prop) != pa.get(prop):
                        differences.append(
                            {
                                "property": f"axis_{ga.get('axis_type')}_{prop}",
                                "gold": ga.get(prop),
                                "pred": pa.get(prop),
                            }
                        )

    elif idMso == "ChartLines":
        for prop in ["has_drop_lines", "has_hi_lo_lines"]:
            if gold_chart.get(prop) != pred_chart.get(prop):
                differences.append(
                    {
                        "property": prop,
                        "gold": gold_chart.get(prop),
                        "pred": pred_chart.get(prop),
                    }
                )

    elif idMso == "ChartUpDownBars":
        if gold_chart.get("has_up_down_bars") != pred_chart.get("has_up_down_bars"):
            differences.append(
                {
                    "property": "has_up_down_bars",
                    "gold": gold_chart.get("has_up_down_bars"),
                    "pred": pred_chart.get("has_up_down_bars"),
                }
            )

    elif idMso == "ColorPickerChartLineStyles":
        gold_series = gold_chart.get("series_info", [])
        pred_series = pred_chart.get("series_info", [])
        gs = next(
            (series for series in gold_series if series.get("index") == 1),
            gold_series[0] if gold_series else None,
        )
        ps = next(
            (series for series in pred_series if series.get("index") == 1),
            pred_series[0] if pred_series else {},
        )
        if gs is None:
            # No gold Series 1 means zero checks would run and the case
            # would pass vacuously; record the missing evidence instead.
            differences.append(
                {
                    "property": "series_1_line_color_rgb",
                    "gold": "missing_gold_series",
                    "pred": ps.get("line_color_rgb") if isinstance(ps, dict) else None,
                }
            )
        elif gs.get("line_color_rgb") != ps.get("line_color_rgb"):
            differences.append(
                {
                    "property": f"series_{gs.get('index')}_line_color_rgb",
                    "gold": gs.get("line_color_rgb"),
                    "pred": ps.get("line_color_rgb"),
                }
            )

    elif idMso == "ChartSwitchRowColumn":
        # Count alone does not establish the transposition or preserve data.
        for prop in ["series_count", "chart_type", "chart_color", "series_names",
                     "series_values", "categories"]:
            if gold_chart.get(prop) != pred_chart.get(prop):
                differences.append(
                    {
                        "property": prop,
                        "gold": gold_chart.get(prop),
                        "pred": pred_chart.get(prop),
                    }
                )

    elif idMso == "ChartResetToMatchStyle":
        # Reset clears direct formatting on series (fill colors revert to theme)
        # Chart.Style is not the operation target and Word can report a
        # different style id for the same inherited series appearance.
        for prop in ["chart_color", "series_count"]:
            if gold_chart.get(prop) != pred_chart.get(prop):
                differences.append(
                    {
                        "property": prop,
                        "gold": gold_chart.get(prop),
                        "pred": pred_chart.get(prop),
                    }
                )
        # Compare series fill colors (main signal for ClearToMatchStyle)
        gold_series = gold_chart.get("series_info", [])
        pred_series = pred_chart.get("series_info", [])
        for gs, ps in zip(gold_series, pred_series):
            # A successfully reset series may have no direct fill at all,
            # inheriting the chart style. Gold saved by Word can instead
            # materialize the same theme fill as an explicit RGB value.
            pred_fill = ps.get("fill_color_rgb")
            if pred_fill is not None and gs.get("fill_color_rgb") != pred_fill:
                differences.append(
                    {
                        "property": f"series_{gs.get('index')}_fill_color_rgb",
                        "gold": gs.get("fill_color_rgb"),
                        "pred": pred_fill,
                    }
                )

    else:
        # Generic: compare all chart props
        for prop in [
            "chart_style",
            "chart_color",
            "has_legend",
            "has_title",
            "has_data_table",
            "has_drop_lines",
            "has_hi_lo_lines",
            "has_up_down_bars",
        ]:
            if gold_chart.get(prop) != pred_chart.get(prop):
                differences.append(
                    {
                        "property": prop,
                        "gold": gold_chart.get(prop),
                        "pred": pred_chart.get(prop),
                    }
                )

    return differences


def diff_operation(op_name: str, op_params: Dict, gold_chart: Dict, pred_chart: Dict) -> List[Dict]:
    """Per-operation chart property comparison for composition steps.

    Op_params guides WHICH properties to check; the actual comparison is
    always pred vs gold so we only flag genuine pred deviations from the
    gold the generator produced.
    """
    differences: List[Dict] = []

    if op_name == "insert_chart":
        # The step specifies chart_type + categories + series (name+values).
        # Check chart_type and series_count vs gold for any insert_chart
        # operation; check per-series name/values only if the step explicitly
        # specifies them AND gold has a non-None value for that field.
        if pred_chart.get("chart_type") != gold_chart.get("chart_type"):
            differences.append(
                {
                    "property": "chart_type",
                    "gold": gold_chart.get("chart_type"),
                    "pred": pred_chart.get("chart_type"),
                }
            )
        gold_series = gold_chart.get("series_info", []) or []
        pred_series = pred_chart.get("series_info", []) or []
        if len(pred_series) != len(gold_series):
            differences.append(
                {
                    "property": "series_count",
                    "gold": len(gold_series),
                    "pred": len(pred_series),
                }
            )
        if op_params.get("series") is not None:
            for i, (gs, ps) in enumerate(zip(gold_series, pred_series)):
                # Series name: only compare if gold actually carries one.
                g_name = gs.get("name")
                if g_name is not None and ps.get("name") != g_name:
                    differences.append(
                        {
                            "property": f"series_{i+1}_name",
                            "gold": g_name,
                            "pred": ps.get("name"),
                        }
                    )
                # Series values: try multiple field names; only compare if
                # both gold and pred expose them.
                g_vals = gs.get("values") or gs.get("data_values")
                p_vals = ps.get("values") or ps.get("data_values")
                if g_vals is not None and p_vals is not None:
                    gv = [round(v, 2) if isinstance(v, float) else v for v in g_vals]
                    pv = [round(v, 2) if isinstance(v, float) else v for v in p_vals]
                    if pv != gv:
                        differences.append(
                            {
                                "property": f"series_{i+1}_values",
                                "gold": gv,
                                "pred": pv,
                            }
                        )
        # Categories: compare only if both gold and pred expose them.
        if op_params.get("categories") is not None:
            gold_cats = gold_chart.get("categories")
            pred_cats = pred_chart.get("categories")
            if gold_cats is not None and pred_cats is not None:
                if [str(c) for c in pred_cats] != [str(c) for c in gold_cats]:
                    differences.append(
                        {
                            "property": "categories",
                            "gold": gold_cats,
                            "pred": pred_cats,
                        }
                    )

    elif op_name == "change_chart_type":
        if op_params.get("chart_type") is not None:
            if pred_chart.get("chart_type") != gold_chart.get("chart_type"):
                differences.append(
                    {
                        "property": "chart_type",
                        "gold": gold_chart.get("chart_type"),
                        "pred": pred_chart.get("chart_type"),
                    }
                )

    elif op_name == "add_drop_lines":
        if pred_chart.get("has_drop_lines") != gold_chart.get("has_drop_lines"):
            differences.append(
                {
                    "property": "has_drop_lines",
                    "gold": gold_chart.get("has_drop_lines"),
                    "pred": pred_chart.get("has_drop_lines"),
                }
            )

    elif op_name == "add_hi_lo_lines":
        if pred_chart.get("has_hi_lo_lines") != gold_chart.get("has_hi_lo_lines"):
            differences.append(
                {
                    "property": "has_hi_lo_lines",
                    "gold": gold_chart.get("has_hi_lo_lines"),
                    "pred": pred_chart.get("has_hi_lo_lines"),
                }
            )

    elif op_name == "add_up_down_bars":
        if pred_chart.get("has_up_down_bars") != gold_chart.get("has_up_down_bars"):
            differences.append(
                {
                    "property": "has_up_down_bars",
                    "gold": gold_chart.get("has_up_down_bars"),
                    "pred": pred_chart.get("has_up_down_bars"),
                }
            )

    elif op_name == "add_chart_title":
        for prop in ["has_title", "title_include_in_layout"]:
            if pred_chart.get(prop) != gold_chart.get(prop):
                differences.append(
                    {
                        "property": prop,
                        "gold": gold_chart.get(prop),
                        "pred": pred_chart.get(prop),
                    }
                )

    elif op_name == "format_chart_title":
        props = ["title_text"]
        if op_params.get("bold") is not None:
            props.append("title_font_bold")
        for prop in props:
            if pred_chart.get(prop) != gold_chart.get(prop):
                differences.append(
                    {
                        "property": prop,
                        "gold": gold_chart.get(prop),
                        "pred": pred_chart.get(prop),
                    }
                )

    # Unknown operation: fall through with no diffs (relies on idMso steps)

    return differences


def expected_insert_contract(spec: Dict[str, Any]) -> Dict[str, Any]:
    params = spec.get("verification", {}).get("primary", {}).get("params", {})
    chart_type = params.get("chart_type")
    kind = {51: "clustered_column", 4: "line", 5: "pie"}.get(chart_type)
    categories = [canonical_value(item) for item in params.get("categories", [])]
    series = []
    for item in params.get("series", []):
        series.append(
            {
                "name": str(item.get("name", "")).strip(),
                "values": [canonical_value(value) for value in item.get("values", [])],
                "categories": categories,
            }
        )
    return {
        "chart_type": chart_type,
        "kind": kind,
        "series": series,
    }


def com_insert_failures(
    state: Dict[str, Any],
    expected: Dict[str, Any],
) -> List[str]:
    failures = []
    if expected.get("chart_type") is None or not expected.get("series"):
        failures.append("test contract lacks chart type or series data")
        return failures
    if state.get("chart_count") != 1:
        failures.append(f"expected one chart, found {state.get('chart_count')}")
    if state.get("visual_object_count") != 1:
        failures.append(
            f"expected one body visual object, found {state.get('visual_object_count')}"
        )
    if state.get("chart_type") != expected.get("chart_type"):
        failures.append(f"chart type {state.get('chart_type')} != {expected.get('chart_type')}")
    if state.get("visible_on_page", True) is not True:
        failures.append("chart has empty or page-clipped reopened geometry")
    if state.get("render_visible", True) is not True:
        failures.append("reopened chart export has no visible chart content")
    actual_series = []
    for item in state.get("series") or []:
        actual_series.append(
            {
                "name": str(item.get("name", "")).strip(),
                "values": (
                    None
                    if item.get("values") is None
                    else [canonical_value(value) for value in item["values"]]
                ),
                "categories": (
                    None
                    if item.get("categories") is None
                    else [canonical_value(value) for value in item["categories"]]
                ),
            }
        )
    if actual_series != expected.get("series"):
        failures.append("ordered series names/categories/values differ")
    return failures


def saved_insert_failures(
    state: Dict[str, Any],
    expected: Dict[str, Any],
) -> List[str]:
    failures = list(state.get("errors") or [])
    if state.get("object_count") != 1:
        failures.append(f"expected one referenced chart, found {state.get('object_count')}")
    if state.get("body_object_count") != 1:
        failures.append(f"expected one saved body object, found {state.get('body_object_count')}")
    charts = state.get("charts") or []
    if len(charts) != 1:
        return failures
    chart = charts[0]
    if chart.get("kind") != expected.get("kind"):
        failures.append(f"saved chart kind {chart.get('kind')} != {expected.get('kind')}")
    cached_series = [
        {
            "name": item.get("name"),
            "categories": item.get("categories"),
            "values": item.get("values"),
        }
        for item in chart.get("series") or []
    ]
    if cached_series != expected.get("series"):
        failures.append("saved ordered chart caches differ from requested data")
    expected_series = expected.get("series") or []
    actual_series = chart.get("series") or []
    all_formula_values = [
        value for item in actual_series for value in (item.get("formulas") or {}).values()
    ]
    has_formulas = any(value is not None for value in all_formula_values)
    formulas_complete = bool(actual_series) and all(
        all(
            (item.get("formulas") or {}).get(role) is not None
            for role in ("name", "categories", "values")
        )
        for item in actual_series
    )
    if has_formulas and not formulas_complete:
        failures.append("saved chart formulas are incomplete")
    workbook_cells = chart.get("workbook_cells")
    if workbook_cells:
        if formulas_complete:
            for index, (actual, requested) in enumerate(zip(actual_series, expected_series)):
                formulas = actual.get("formulas") or {}
                projections = {
                    role: project_workbook(workbook_cells, formulas.get(role))
                    for role in ("name", "categories", "values")
                }
                requested_projection = {
                    "name": [requested.get("name")],
                    "categories": requested.get("categories"),
                    "values": requested.get("values"),
                }
                for role, projected in projections.items():
                    if projected is None:
                        failures.append(
                            f"series {index + 1} {role} formula is not a "
                            "parseable local workbook range"
                        )
                    elif projected != requested_projection[role]:
                        failures.append(
                            f"series {index + 1} referenced {role} workbook "
                            "values differ from requested data"
                        )
        else:
            # The generated golds use literal chart caches and a canonical
            # embedded workbook without c:f nodes.  Preserve gold-as-pred
            # while checking the same effective cells.  Predictions that
            # do carry formulas always take the stricter projection path.
            sheet = next(iter(workbook_cells.values()))
            expected_cells: Dict[str, Any] = {"A1": ""}
            for series_index, requested in enumerate(expected_series, 2):
                expected_cells[f"{column_name(series_index)}1"] = requested.get("name")
            categories = expected_series[0].get("categories") if expected_series else []
            for row_index, category in enumerate(categories, 2):
                expected_cells[f"A{row_index}"] = category
                for series_index, requested in enumerate(expected_series, 2):
                    expected_cells[f"{column_name(series_index)}{row_index}"] = requested.get(
                        "values", []
                    )[row_index - 2]
            for reference, requested_value in expected_cells.items():
                actual_value = canonical_value(sheet.get(reference))
                if reference == "A1" and actual_value is None:
                    actual_value = ""
                if actual_value != requested_value:
                    failures.append("literal-cache workbook cells differ from requested data")
                    break
    elif formulas_complete:
        # Formula-backed caches without an embedded workbook are accepted
        # only when every role points at a bounded local A1 range whose
        # shape matches the requested/cache projection.  This prevents a
        # correct visible cache from masking a nonsensical backing range.
        # COM remains the primary proof that Word exposes the data.
        for index, (actual, requested) in enumerate(zip(actual_series, expected_series), start=1):
            expected_lengths = {
                "name": 1,
                "categories": len(requested.get("categories") or []),
                "values": len(requested.get("values") or []),
            }
            parsed_formulas: Dict[str, Dict[str, Any]] = {}
            for role, formula in (actual.get("formulas") or {}).items():
                parsed = parse_range_formula(formula)
                if parsed is None:
                    failures.append(
                        f"series {index} {role} formula is not a " "parseable local workbook range"
                    )
                elif len(parsed["cells"]) != expected_lengths[role]:
                    failures.append(
                        f"series {index} {role} formula range contains "
                        f"{len(parsed['cells'])} cell(s), expected "
                        f"{expected_lengths[role]}"
                    )
                else:
                    parsed_formulas[role] = parsed
            formula_sheets = {parsed["sheet"].casefold() for parsed in parsed_formulas.values()}
            if len(parsed_formulas) == 3 and len(formula_sheets) != 1:
                failures.append(f"series {index} formula ranges use different sheets")
    extents = state.get("extents") or []
    if len(extents) != 1:
        failures.append(f"expected one saved chart extent, found {len(extents)}")
    else:
        extent = extents[0]
        width = extent.get("width_emu") or 0
        height = extent.get("height_emu") or 0
        if width <= 0 or height <= 0:
            failures.append("saved chart has an empty display extent")
        printable_width = state.get("printable_width_emu")
        printable_height = state.get("printable_height_emu")
        if printable_width and width > printable_width:
            failures.append("saved chart width exceeds the printable page width")
        if printable_height and height > printable_height:
            failures.append("saved chart height exceeds the printable page height")
    return failures
