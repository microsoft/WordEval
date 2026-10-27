"""
Additive composition generator for object_effect behavior.

Compositions grouped by object type:
  - AdditiveShapeComposition (4 variants): fill, outline, effects, bevel, shape change
  - AdditiveChartComposition (5 variants): bar chart + line chart elements
  - AdditivePictureComposition (7 variants): effects, adjustments, crop
  - AdditiveSmartArtComposition (3 variants): style/colors + connectors + text effects
  - AdditiveTableComposition (3 variants): style/borders + structure/layout
  - AdditiveLayoutComposition (6 variants): align+distribute + rotate+distribute

Usage:
    python -m generator.compositions.additive.object_effect
    python -m generator.compositions.additive.object_effect --command AdditiveShapeComposition
"""

import json
import shutil
import time
from datetime import datetime
from pathlib import Path
from typing import Dict

from generator.config import BENCHMARK_OUTPUT_DIR
from generator.compositions.atomic_intents import build_object_effect_atomic_intents
from generator.compositions.interaction.shared import write_instructions_json
from generator.execution.setup import ensure_word
from generator.test_sets.spec_utils import write_spec_if_changed

TEST_SET = 'additive_composition_x_object_effect'

# Keep composition setters on the same Office standard-color values used by
# the linked contextual-parameterized atomic cases.
OFFICE_STANDARD_COLOR_BGR = {
    'red': 238,          # RGB(238, 0, 0)
    'blue': 12611584,    # RGB(0, 112, 192)
    'green': 5287936,    # RGB(0, 176, 80)
}


# ============================================================================
# Shared helper utilities
# ============================================================================

def _apply_visible_table_borders(table):
    """Apply visible single-line borders so table structure is reviewable."""
    try:
        for border_id in [1, 2, 3, 4, 5, 6]:
            table.Borders(border_id).LineStyle = 1
    except Exception:
        pass


# ============================================================================
# Shape effect slot operations
# ============================================================================

def _apply_shape_slot(shape, slot_id, value):
    """Apply a shape effect slot value to a Shape object."""
    if slot_id == 'fill_color':  # fill color
        shape.Fill.Solid()
        shape.Fill.ForeColor.RGB = OFFICE_STANDARD_COLOR_BGR[value]

    elif slot_id == 'outline_color':  # outline color
        shape.Line.Visible = True
        shape.Line.ForeColor.RGB = OFFICE_STANDARD_COLOR_BGR[value]

    elif slot_id == 'shadow':  # shadow
        shadow_map = {'outer_center': 25, 'outer_diagonal_bottom_right': 21}
        shape.Shadow.Type = shadow_map[value]

    elif slot_id == 'glow':  # glow
        if value == 'glow_11pt_accent_4':
            shape.Glow.Radius = 11.0
            # Start from the atomic preset's resolved RGB.  After save, the
            # same atomic-generator helper restores the gallery's schemeClr
            # transforms, which the Word Object Model cannot emit directly.
            shape.Glow.Color.RGB = 49407
        else:
            radius, color_bgr = {
                'glow_18pt_accent_6': (18.0, 4697456),
            }[value]
            shape.Glow.Radius = radius
            shape.Glow.Color.RGB = color_bgr

    elif slot_id == 'soft_edges':  # soft edges
        radius_map = {'5pt': 5.0, '10pt': 10.0, '25pt': 25.0}
        shape.SoftEdge.Radius = radius_map[value]

    elif slot_id == 'reflection':  # reflection
        type_map = {'tight_touching': 1, 'half_touching': 2}
        shape.Reflection.Type = type_map[value]


# ============================================================================
# Chart slot operations
# ============================================================================

def _apply_chart_slot(wd, chart, slot_id, value):
    """Apply a chart slot value to a Chart object."""
    if slot_id == 'title':  # chart title
        chart.HasTitle = True
        chart.ChartTitle.Text = 'Chart Title'
        if value == 'centered_overlay':
            chart.ChartTitle.IncludeInLayout = False
        else:  # above_chart
            chart.ChartTitle.IncludeInLayout = True

    elif slot_id == 'legend':  # legend
        position_map = {'left': -4131, 'right': -4152, 'top': -4160}
        chart.HasLegend = True
        chart.Legend.Position = position_map[value]

    elif slot_id == 'data_labels':  # data labels
        position_map = {'center': -4108, 'outside_end': 2}
        for i in range(1, chart.SeriesCollection().Count + 1):
            chart.SeriesCollection(i).HasDataLabels = True
            chart.SeriesCollection(i).DataLabels().Position = position_map[value]

    elif slot_id == 'color_scheme':  # chart colors
        color_map = {'monochromatic_accent5': 18, 'monochromatic_palette_7': 20}
        chart.ChartColor = color_map[value]

    elif slot_id == 'gridlines':  # gridlines
        # Show major gridlines on value axis
        chart.Axes(2).HasMajorGridlines = True

    elif slot_id == 'trendline':  # trendline
        s = chart.SeriesCollection(1)
        while s.Trendlines().Count > 0:
            s.Trendlines(1).Delete()
        s.Trendlines().Add(Type=-4132)  # xlLinear

    elif slot_id == 'data_table':  # data table
        chart.HasDataTable = True
        if value == 'no_legend_key':
            chart.DataTable.ShowLegendKey = False
        else:
            chart.DataTable.ShowLegendKey = True

    elif slot_id == 'series_color':  # series line color
        color_map = {'blue': 16711680, 'red': 255}
        s = chart.SeriesCollection(1)
        s.Format.Line.Visible = True
        s.Format.Line.ForeColor.RGB = color_map[value]


# ============================================================================
# Chart (line chart) slot operations
# ============================================================================

def _apply_line_chart_slot(wd, chart, slot_id, value):
    """Apply a line chart slot value to a Chart object."""
    if slot_id == 'chart_lines':  # chart lines (drop or hilo)
        cg = chart.ChartGroups(1)
        if value == 'drop_lines':
            cg.HasDropLines = True
        elif value == 'hilo_lines':
            cg.HasHiLoLines = True

    elif slot_id == 'up_down_bars':  # up/down bars
        chart.ChartGroups(1).HasUpDownBars = True

    else:
        # Reuse bar chart slot for shared commands (title, legend, labels, gridlines)
        _apply_chart_slot(wd, chart, slot_id, value)


# ============================================================================
# Picture extended slot operations
# ============================================================================

def _apply_picture_extended_slot(wd, obj, slot_id, value):
    """Apply a picture extended slot value to an InlineShape object."""
    if slot_id == 'saturation':  # color saturation
        sat_map = {'0pct': 0.0, '33pct': 0.33}
        shape = obj.ConvertToShape()
        try:
            effect = shape.Fill.PictureEffects.Insert(24)  # ColorSaturation
            effect.EffectParameters.Item(1).Value = sat_map[value]
        finally:
            obj = shape.ConvertToInlineShape()

    elif slot_id == 'temperature':  # color temperature
        temp_map = {'4700K': 4700, '11200K': 11200}
        shape = obj.ConvertToShape()
        try:
            effect = shape.Fill.PictureEffects.Insert(7)  # ColorTemperature
            effect.EffectParameters.Item(1).Value = temp_map[value]
        finally:
            obj = shape.ConvertToInlineShape()

    elif slot_id == 'sharpen_soften':  # sharpen/soften
        amount_map = {'sharpen_50': 0.5, 'soften_50': -0.5}
        shape = obj.ConvertToShape()
        try:
            effect = shape.Fill.PictureEffects.Insert(25)  # SharpenSoften
            effect.EffectParameters.Item(1).Value = amount_map[value]
        finally:
            obj = shape.ConvertToInlineShape()

    elif slot_id == 'artistic_effect':  # artistic effect
        effect_map = {'glow_edges': 12, 'paint_strokes': 18}
        shape = obj.ConvertToShape()
        try:
            shape.Fill.PictureEffects.Insert(effect_map[value])
        finally:
            obj = shape.ConvertToInlineShape()

    elif slot_id == 'rotation_3d':  # 3D rotation
        # Keep composition presets aligned with the atomic
        # PictureRotationGallery variants.
        preset_map = {'oblique_bottom_right': 47, 'isometric_top_up': 20}
        shape = obj.ConvertToShape()
        try:
            shape.ThreeD.SetPresetCamera(preset_map[value])
        finally:
            obj = shape.ConvertToInlineShape()

    elif slot_id == 'pic_outline':  # outline color
        shape = obj.ConvertToShape()
        try:
            shape.Line.Visible = True
            shape.Line.ForeColor.RGB = OFFICE_STANDARD_COLOR_BGR[value]
        finally:
            obj = shape.ConvertToInlineShape()

    elif slot_id == 'ext_shadow':  # shadow (reused from picture effects)
        shadow_map = {'outer_center': 25, 'outer_diagonal_bottom_right': 21}
        obj.Shadow.Visible = True
        obj.Shadow.Type = shadow_map[value]

    elif slot_id == 'ext_glow':  # glow (reused)
        glow_map = {
            'glow_11pt_accent_4': (11.0, 8),
        }
        radius, theme_color = glow_map[value]
        obj.Glow.Radius = radius
        obj.Glow.Color.ObjectThemeColor = theme_color
        obj.Glow.Transparency = 0.0

    elif slot_id == 'ext_soft_edges':  # soft edges (reused)
        radius_map = {'5pt': 5.0, '25pt': 25.0}
        obj.SoftEdge.Radius = radius_map[value]

    elif slot_id == 'ext_reflection':  # reflection (reused)
        type_map = {'tight_touching': 1}
        obj.Reflection.Type = type_map[value]

    return obj  # Return obj in case ConvertToShape/ConvertToInlineShape changed it


# ============================================================================
# SmartArt combined slot operations
# ============================================================================

def _target_connector(target_node):
    """Return the connector owned by the selected SmartArt node.

    This deliberately follows the same node-scoped selector as the linked
    ``ArrowStyleGallery`` atomic generator.  Geometric proximity in the saved
    drawing cache is not a stable ownership signal across SmartArt layouts or
    Word versions.
    """
    for i in range(1, target_node.Shapes.Count + 1):
        try:
            shape = target_node.Shapes(i)
            if shape.Type == 9:
                return shape
        except Exception:
            pass
    return None


def _apply_smartart_combined_slot(wd, smartart_shape, slot_id, value):
    """Apply a SmartArt combined slot (connector or text effect) value."""
    smartart = smartart_shape.SmartArt

    # Find target node
    target_node = None
    for i in range(1, smartart.Nodes.Count + 1):
        try:
            if smartart.Nodes(i).TextFrame2.TextRange.Text.strip() == 'Target':
                target_node = smartart.Nodes(i)
                break
        except Exception:
            pass
    if target_node is None:
        target_node = smartart.Nodes(1)

    if slot_id == 'node_outline_weight':  # outline weight of the target node
        weight_map = {'3pt': 3.0, '6pt': 6.0}
        # Mirrors the OutlineWeightGallery atomic, which targets the node shape.
        node_shape = target_node.Shapes(1)
        node_shape.Line.Visible = True
        node_shape.Line.Weight = weight_map[value]

    elif slot_id == 'node_outline_dash':  # outline dash style of the target node
        dash_map = {'dash': 4}  # msoLineDash = 4
        node_shape = target_node.Shapes(1)
        node_shape.Line.Visible = True
        node_shape.Line.DashStyle = dash_map[value]

    elif slot_id == 'connector_arrow':  # connector arrowhead style
        arrow_map = {
            'none_to_triangle': (1, 2),  # begin=none, end=triangle
            'triangle_to_none': (2, 1),
        }
        begin_arrow, end_arrow = arrow_map[value]
        connector = _target_connector(target_node)
        if connector is None:
            raise RuntimeError("Target SmartArt node has no connector shape")
        connector.Line.BeginArrowheadStyle = begin_arrow
        connector.Line.EndArrowheadStyle = end_arrow

    elif slot_id == 'text_fill':  # text fill color
        color_map = {'blue': 16711680, 'green': 65280}
        tf2 = target_node.Shapes(1).TextFrame2
        tf2.TextRange.Font.Fill.ForeColor.RGB = color_map[value]

    elif slot_id == 'text_transform':  # text transform (warp format)
        warp_map = {'fade_right': 28}
        tf2 = target_node.Shapes(1).TextFrame2
        tf2.WarpFormat = warp_map[value]

    elif slot_id == 'text_3d':  # text 3D rotation
        preset_map = {'isometric_left_down': 25}
        tf2 = target_node.Shapes(1).TextFrame2
        tf2.ThreeD.SetPresetCamera(preset_map[value])

    elif slot_id == 'text_bevel':  # text bevel
        bevel_map = {'angle': 6, 'circle': 3}
        tf2 = target_node.Shapes(1).TextFrame2
        td = tf2.ThreeD
        td.BevelTopType = bevel_map[value]
        td.BevelTopDepth = 4.0
        td.BevelTopInset = 4.0


# ============================================================================
# Shape extended slot operations
# ============================================================================

def _apply_shape_extended_slot(shape, slot_id, value):
    """Apply a shape extended slot value (bevel, shape change, or reuse base slots)."""
    if slot_id == 'bevel':  # bevel
        bevel_map = {'angle': 6, 'cross': 5}
        shape.ThreeD.BevelTopType = bevel_map[value]
        shape.ThreeD.BevelTopDepth = 6.0
        shape.ThreeD.BevelTopInset = 6.0

    elif slot_id == 'shape_change':  # shape change (auto shape type)
        shape_map = {'ellipse': 9, 'callout_line': 109, 'lightning_bolt': 22}
        shape.AutoShapeType = shape_map[value]

    else:
        # Reuse base shape slots (S1-S6)
        _apply_shape_slot(shape, slot_id, value)


# ============================================================================
# Table structure slot operations
# ============================================================================

def _apply_table_structure_slot(wd, table, slot_id, value):
    """Apply a table structure slot value."""
    if slot_id == 'autofit':  # auto-fit
        fit_map = {'contents': 1, 'window': 2}  # wdAutoFitContent=1, wdAutoFitWindow=2
        table.AutoFitBehavior(fit_map[value])

    elif slot_id == 'distribute_rows':  # distribute rows
        table.Rows.DistributeHeight()

    elif slot_id == 'distribute_cols':  # distribute columns
        table.Columns.DistributeWidth()

    elif slot_id == 'cell_align':  # cell align middle center
        for row in table.Rows:
            for cell in row.Cells:
                try:
                    cell.VerticalAlignment = 1  # wdCellAlignVerticalCenter
                    cell.Range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
                except Exception:
                    pass


# ============================================================================
# Variant definitions
# ============================================================================

SHAPE_VARIANTS = {
    # --- Base effects (fill, outline, shadow, glow, soft edges, reflection) ---
    'variant_fill_outline_effects': {
        'description': 'Shape fill + outline + effects (5 operations)',
        'instruction': "Set the shape's fill color to red (RGB: 238, 0, 0) and outline to blue (RGB: 0, 112, 192), add an outer center shadow, add an 11-point glow (Accent 4), and add a tight touching reflection.",
        'slots': {
            'fill_color': 'red', 'outline_color': 'blue', 'shadow': 'outer_center',
            'glow': 'glow_11pt_accent_4', 'reflection': 'tight_touching',
        },
    },
    'variant_fill_shadow_softedges': {
        'description': 'Shape fill + outline + shadow + soft edges (4 operations)',
        'instruction': "Set the shape's fill color to blue (RGB: 0, 112, 192) and outline to green (RGB: 0, 176, 80), add an outer center shadow, and add 25-point soft edges.",
        'slots': {
            'fill_color': 'blue', 'outline_color': 'green',
            'shadow': 'outer_center', 'soft_edges': '25pt',
        },
    },
    # --- Extended (shape change, bevel + base effects) ---
    'variant_ellipse_bevel_effects': {
        'description': 'Shape change + bevel + fill + outline + effects (7 operations)',
        'instruction': "Change the shape to an ellipse with 'Angle' bevel, set fill to red (RGB: 238, 0, 0) and outline to blue (RGB: 0, 112, 192), add an outer center shadow, an 11-point glow (Accent 4), and a tight touching reflection.",
        'slots': {
            'shape_change': 'ellipse', 'bevel': 'angle',
            'fill_color': 'red', 'outline_color': 'blue',
            'shadow': 'outer_center', 'glow': 'glow_11pt_accent_4',
            'reflection': 'tight_touching',
        },
    },
    'variant_lightning_bolt_bevel_softedges': {
        'description': 'Shape change + bevel + fill + outline + soft edges (5 operations)',
        'instruction': "Change the shape to a lightning bolt with 'Cross' bevel, set fill to blue (RGB: 0, 112, 192) and outline to green (RGB: 0, 176, 80), and add 25-point soft edges.",
        'slots': {
            'shape_change': 'lightning_bolt', 'bevel': 'cross',
            'fill_color': 'blue', 'outline_color': 'green',
            'soft_edges': '25pt',
        },
    },
}

# Command names for shape slots
SHAPE_COMMAND_MAP = {
    'fill_color': 'ShapeFillColorPicker', 'outline_color': 'ShapeOutlineColorPicker',
    'shadow': 'ObjectEffectShadowGallery', 'glow': 'ObjectEffectGlowGallery',
    'soft_edges': 'ObjectEffectSoftEdgesGallery', 'reflection': 'ReflectionGallery',
    'bevel': 'BevelShapeGallery', 'shape_change': 'ShapeChangeShapeGallery',
}

CHART_VARIANTS = {
    # --- Bar chart basic (doc_in_chart) ---
    'variant_bar_title_legend_colors': {
        'description': 'Chart title + legend + data labels + colors (4 operations)',
        'instruction': 'Add a title "Chart Title" above the bar chart, move the legend to the right side, add data labels at the center of each bar, and change the chart to a purple monochromatic color scheme.',
        'slots': {
            'title': 'above_chart', 'legend': 'right',
            'data_labels': 'center', 'color_scheme': 'monochromatic_accent5',
        },
    },
    'variant_bar_gridlines_trendline': {
        'description': 'Chart layout with gridlines + trendline (5 operations)',
        'instruction': 'Add a centered overlay title "Chart Title" on the bar chart, move the legend to the left, add data labels above each bar, show major gridlines on the value axis, and add a linear trendline to the first series.',
        'slots': {
            'title': 'centered_overlay', 'legend': 'left',
            'data_labels': 'outside_end', 'gridlines': 'major_value_gridlines',
            'trendline': 'linear_trendline',
        },
    },
    # --- Bar chart extended (doc_in_chart) ---
    'variant_bar_datatable_seriescolor': {
        'description': 'Chart with data table + series line color (6 operations)',
        'instruction': "Add a title \"Chart Title\" above the bar chart, move the legend to the left, add data labels above each bar, show major value gridlines, add a data table without legend key, and change the outline color of the first data series to blue (RGB: 0, 0, 255).",
        'slots': {
            'title': 'above_chart', 'legend': 'left',
            'data_labels': 'outside_end', 'gridlines': 'major_value_gridlines',
            'data_table': 'no_legend_key', 'series_color': 'blue',
        },
    },
    # --- Line chart (doc_in_chart_line_clean) ---
    # These two use a composition-only input with the default chart title
    # placeholder and the major value gridlines removed, so that adding them is
    # a real edit. doc_in_chart_line is left alone because three atomic cases
    # under object_effect_x_contextual_parameterized share it and are not
    # degenerate on it.
    'variant_line_droplines_updown': {
        'description': 'Drop lines + up/down bars + title + labels + legend (5 operations)',
        'instruction': "Add drop lines and up/down bars to the line chart, add a title \"Chart Title\" above the chart, show data labels at the center, and move the legend to the right.",
        'doc_in_name': 'doc_in_chart_line_clean',
        'type': 'line',
        'slots': {
            'chart_lines': 'drop_lines', 'up_down_bars': 'show',
            'title': 'above_chart', 'data_labels': 'center',
            'legend': 'right',
        },
    },
    'variant_line_hilo_gridlines': {
        'description': 'Hi-lo lines + title + legend + gridlines (4 operations)',
        'instruction': "Add high-low lines to the line chart, add a centered overlay title \"Chart Title\", move the legend to the left, and show major value gridlines.",
        'doc_in_name': 'doc_in_chart_line_clean',
        'type': 'line',
        'slots': {
            'chart_lines': 'hilo_lines',
            'title': 'centered_overlay', 'legend': 'left',
            'gridlines': 'major_value_gridlines',
        },
    },
}

CHART_COMMAND_MAP = {
    'title': 'ChartTitle', 'legend': 'ChartLegend', 'data_labels': 'ChartDataLabel',
    'color_scheme': 'ChartColorsGallery', 'gridlines': 'ChartGridlinesGallery',
    'trendline': 'ChartTrendline', 'data_table': 'ChartDataTable',
    'series_color': 'ColorPickerChartLineStyles',
    'chart_lines': 'ChartLines', 'up_down_bars': 'ChartUpDownBars',
}


# ============================================================================
# Picture effect slot operations
# ============================================================================

def _apply_picture_slot(obj, slot_id, value):
    """Apply a picture effect slot value to an InlineShape object."""
    if slot_id == 'pic_shadow':  # shadow
        shadow_map = {'outer_center': 25, 'outer_diagonal_bottom_right': 21}
        obj.Shadow.Visible = True
        obj.Shadow.Type = shadow_map[value]

    elif slot_id == 'pic_glow':  # glow
        glow_map = {
            'glow_11pt_accent_4': (11.0, 8),
            'glow_8pt_accent_6': (8.0, 6),
        }
        radius, theme_color = glow_map[value]
        obj.Glow.Radius = radius
        obj.Glow.Color.ObjectThemeColor = theme_color
        obj.Glow.Transparency = 0.0

    elif slot_id == 'pic_soft_edges':  # soft edges
        radius_map = {'5pt': 5.0, '10pt': 10.0, '25pt': 25.0}
        obj.SoftEdge.Radius = radius_map[value]

    elif slot_id == 'pic_reflection':  # reflection
        type_map = {'tight_touching': 1, 'half_touching': 2}
        obj.Reflection.Type = type_map[value]

    elif slot_id == 'pic_recolor':  # recolor
        color_map = {'grayscale': 2, 'black_white': 3}
        obj.PictureFormat.ColorType = color_map[value]


PICTURE_VARIANTS = {
    # --- Base effects (shadow, glow, soft edges, reflection, recolor) ---
    'variant_shadow_glow_reflection': {
        'description': 'Picture shadow + glow + reflection (3 operations)',
        'instruction': "Add an outer center shadow to the picture, add an 11-point glow effect (Accent 4), and add a tight touching reflection below it.",
        'slots': {
            'pic_shadow': 'outer_center',
            'pic_glow': 'glow_11pt_accent_4',
            'pic_reflection': 'tight_touching',
        },
    },
    'variant_shadow_softedges_recolor': {
        'description': 'Picture shadow + soft edges + recolor (3 operations)',
        'instruction': "Add an outer center shadow to the picture, add 25-point soft edges, and recolor it to grayscale.",
        'slots': {
            'pic_shadow': 'outer_center',
            'pic_soft_edges': '25pt',
            'pic_recolor': 'grayscale',
        },
    },
    # --- Extended (adjustments, artistic effects, 3D, outline + effects) ---
    'variant_adjustments_artistic_3d': {
        'type': 'extended',
        'description': 'Pure adjustments (5 operations)',
        'instruction': "Set the picture's color saturation to 0% and color temperature to 4700K, sharpen by 50%, apply 'Glow Edges' artistic effect, and set 'Oblique Bottom Right' 3D rotation.",
        'slots': {
            'saturation': '0pct', 'temperature': '4700K', 'sharpen_soften': 'sharpen_50',
            'artistic_effect': 'glow_edges', 'rotation_3d': 'oblique_bottom_right',
        },
    },
    'variant_adjustments_outline_effects': {
        'type': 'extended',
        'description': 'Mixed adjustments + effects (6 operations)',
        'instruction': "Set the picture's color saturation to 33% with 'Isometric Top Up' 3D rotation, add a red outline (RGB: 238, 0, 0), an outer center shadow, an 11-point glow (Accent 4), and a tight touching reflection.",
        'slots': {
            'saturation': '33pct', 'rotation_3d': 'isometric_top_up',
            'pic_outline': 'red', 'ext_shadow': 'outer_center',
            'ext_glow': 'glow_11pt_accent_4', 'ext_reflection': 'tight_touching',
        },
    },
    'variant_temperature_artistic_shadow': {
        'type': 'extended',
        'description': 'Different mix (5 operations)',
        'instruction': "Set the picture's color temperature to 11200K, soften by 50%, apply 'Paint Strokes' artistic effect, add an outer center shadow, and add 25-point soft edges.",
        'slots': {
            'temperature': '11200K', 'sharpen_soften': 'soften_50',
            'artistic_effect': 'paint_strokes', 'ext_shadow': 'outer_center',
            'ext_soft_edges': '25pt',
        },
    },
    # --- Crop + effect (landscape picture) ---
    'variant_crop_16x9_shadow': {
        'type': 'crop',
        'description': 'Crop to 16:9 + add outer center shadow',
        'instruction': (
            "Crop the picture to a 16:9 aspect ratio (centered crop "
            "removing equal amounts from the top and bottom), "
            "and add an outer center shadow to the picture."
        ),
        'ratio': 16 / 9,
        'effect': 'shadow',
        'shadow_type': 25,
    },
    'variant_crop_square_reflection': {
        'type': 'crop',
        'description': 'Crop to 1:1 square + add tight touching reflection',
        'instruction': (
            "Crop the picture to a 1:1 square aspect ratio (centered crop "
            "removing equal amounts from the left and right), "
            "and add a 'Tight Reflection, Touching' reflection effect "
            "to the picture."
        ),
        'ratio': 1.0,
        'effect': 'reflection',
        'reflection_type': 1,
    },
}

PICTURE_COMMAND_MAP = {
    'pic_shadow': 'PictureEffectsShadowGallery', 'pic_glow': 'PictureEffectsGlowGallery',
    'pic_soft_edges': 'PictureEffectsSoftEdgesGallery', 'pic_reflection': 'PictureReflectionGallery',
    'pic_recolor': 'PictureRecolorGallery',
}

PICTURE_EXTENDED_COMMAND_MAP = {
    'saturation': 'PictureSaturationGallery',
    'temperature': 'PictureColorTempertatureGallery',
    'sharpen_soften': 'PictureSharpenSoftenGallery',
    'artistic_effect': 'PictureArtisticEffectsGallery',
    'rotation_3d': 'PictureRotationGallery',
    'pic_outline': 'OutlineColorPicker',
    'ext_shadow': 'PictureEffectsShadowGallery',
    'ext_glow': 'PictureEffectsGlowGallery',
    'ext_soft_edges': 'PictureEffectsSoftEdgesGallery',
    'ext_reflection': 'PictureReflectionGallery',
}


# ============================================================================
# SmartArt slot operations
# ============================================================================

def _apply_smartart_slot(wd, smartart, slot_id, value):
    """Apply a SmartArt slot value."""
    if slot_id == 'sa_style':  # style
        styles = wd.Application.SmartArtQuickStyles
        smartart.QuickStyle = styles(int(value))

    elif slot_id == 'sa_colors':  # colors
        colors = wd.Application.SmartArtColors
        smartart.Color = colors(int(value))


SMARTART_VARIANTS = {
    # --- Base (style + colors) ---
    'variant_style_colors': {
        'description': 'SmartArt style + colors (2 operations)',
        # Aligned with atomic SmartArtStylesGallery (5='Intense Effect') and
        # SmartArtChangeColorsGallery (5='Colorful Range - Accent Colors 2 to 3').
        # Previously the instruction said "3-D polished" which neither matches
        # index 5 ('Intense Effect' is in the "Best Match for Document" group,
        # not the 3-D group) nor "colorful accent scheme" which is index 1, not 5.
        # Corrected during the composition review with newer models (finding N3).
        'instruction': "Apply the 'Intense Effect' quick style to the SmartArt and change its colors to the 'Colorful Range - Accent Colors 2 to 3' scheme.",
        'slots': {
            'sa_style': 5,   # SmartArtQuickStyles(5) = "Intense Effect"
            'sa_colors': 5,  # SmartArtColors(5) = "Colorful Range - Accent Colors 2 to 3"
        },
    },
    # --- Combined (connector + text effects) ---
    'variant_connector_text_full': {
        'type': 'combined',
        'description': 'Full bundle: connector + text effects (7 operations)',
        'instruction': "Set the outline of 'Target' in the SmartArt to a 3-point dashed line, add a triangle arrowhead at the right end of the line next to 'Target', and apply blue text fill (RGB: 0, 0, 255) with 'Fade Right' transform, 'Isometric Left Down' 3D rotation, and 'Angle' bevel to the 'Target' shape text.",
        'slots': {
            'node_outline_weight': '3pt', 'node_outline_dash': 'dash', 'connector_arrow': 'none_to_triangle',
            'text_fill': 'blue', 'text_transform': 'fade_right',
            'text_3d': 'isometric_left_down', 'text_bevel': 'angle',
        },
    },
    'variant_connector_text_basic': {
        'type': 'combined',
        'description': 'Connector-focused with one text effect (5 operations)',
        'instruction': "Set the outline of 'Target' in the SmartArt to a 6-point dashed line, add a triangle arrowhead at the right end of the line next to 'Target', and apply green text fill (RGB: 0, 255, 0) with 'Circle' bevel to the 'Target' shape text.",
        'slots': {
            'node_outline_weight': '6pt', 'node_outline_dash': 'dash', 'connector_arrow': 'none_to_triangle',
            'text_fill': 'green', 'text_bevel': 'circle',
        },
    },
}

SMARTART_COMMAND_MAP = {
    'sa_style': 'SmartArtStylesGallery',
    'sa_colors': 'SmartArtChangeColorsGallery',
}

SMARTART_COMBINED_COMMAND_MAP = {
    'node_outline_weight': 'OutlineWeightGallery',
    'node_outline_dash': 'OutlineDashesGallery',
    'connector_arrow': 'ArrowStyleGallery',
    'text_fill': 'TextFillColorPicker',
    'text_transform': 'TextEffectTransformGallery',
    'text_3d': 'TextEffects3DRotationGallery',
    'text_bevel': 'BevelTextGallery',
}


# ============================================================================
# Table slot operations
# ============================================================================

def _apply_table_slot(obj, slot_id, value):
    """Apply a table slot value to a Table object."""
    if slot_id == 'table_style':  # table style
        obj.Style = value

    elif slot_id == 'text_direction':  # text direction
        for row in obj.Rows:
            for cell in row.Cells:
                try:
                    cell.Range.Orientation = int(value)
                except Exception:
                    pass

    elif slot_id == 'border_style':  # outside border style
        for border_type in range(1, 5):
            try:
                obj.Borders(border_type).LineStyle = int(value)
            except Exception:
                pass


TABLE_VARIANTS = {
    # --- Base (style, text direction, border) ---
    'variant_style_textdir_border': {
        'description': 'Table style + text direction + border (3 operations)',
        'instruction': "Apply the 'Grid Table 4 - Accent 1' style to the table (blue accent with alternating row shading), rotate all cell text to vertical bottom-to-top direction, and change the outside border to double lines.",
        'slots': {
            'table_style': 'Grid Table 4 - Accent 1',
            'text_direction': 2,   # wdTextOrientationUpward (bottom to top)
            # wdLineStyleDouble = 7 (matches "double lines" in instruction).
            # Atomic BorderStylesGallery enumeration also pairs (7, 'double',
            # 'double line'); previously this was 2 (= wdLineStyleDot, "dotted")
            # which contradicted the instruction; corrected during the
            # composition review with newer models (finding N2).
            # Only outside borders are changed here, matching the atomic
            # BorderStylesGallery contract.
            'border_style': 7,
        },
    },
    # --- Structure (auto-fit, distribute, cell alignment) ---
    'variant_autofit_rows_align': {
        'type': 'structure',
        'description': 'Auto-fit contents + distribute rows + cell align (3 operations)',
        'instruction': "Auto-fit the table to its contents, distribute all rows to equal height, and set all cells to middle-center alignment.",
        'doc_in_name': 'doc_in_table_structure_contents',
        'slots': {
            'autofit': 'contents', 'distribute_rows': 'apply',
            'cell_align': 'on',
        },
    },
    'variant_autofit_cols_align': {
        'type': 'structure',
        'description': 'Auto-fit window + distribute columns + cell align (3 operations)',
        'instruction': "Auto-fit the table to the window width, distribute all columns equally, and set all cells to middle-center alignment.",
        'doc_in_name': 'doc_in_table_structure_window',
        'slots': {
            'autofit': 'window', 'distribute_cols': 'apply',
            'cell_align': 'on',
        },
    },
}

TABLE_COMMAND_MAP = {
    'table_style': 'TableStylesGalleryWord',
    'text_direction': 'TextDirection',
    'border_style': 'BorderStylesGallery',
}

TABLE_STRUCTURE_COMMAND_MAP = {
    'autofit_contents': 'TableAutoFitContents',
    'autofit_window': 'TableAutoFitWindow',
    'distribute_rows': 'TableRowsDistribute',
    'distribute_cols': 'TableColumnsDistribute',
    'cell_align': 'TableCellAlignMiddleCenter',
}




# ============================================================================
# Multi-shape layout variant definitions
#
# Align+distribute: orthogonal-axis pairs (align X + distribute Y or vice versa).
# Rotate+distribute: rotation-agnostic distribute (Word uses un-rotated coords).
# ============================================================================

LAYOUT_VARIANTS = {
    # --- Align + distribute (orthogonal axes) ---
    'variant_align_left_distribute_v': {
        'type': 'align',
        'description': 'Align lefts (X) + distribute vertically (Y) — orthogonal axes',
        'instruction': (
            "Align all three rectangles by their left edges, "
            "and distribute all three rectangles evenly vertically with "
            "equal spacing between them."
        ),
        'align_idMso': 'ObjectsAlignLeftSmart',
        'distribute_idMso': 'AlignDistributeVertically',
    },
    'variant_align_top_distribute_h': {
        'type': 'align',
        'description': 'Align tops (Y) + distribute horizontally (X) — orthogonal axes',
        'instruction': (
            "Align all three rectangles by their top edges, "
            "and distribute all three rectangles evenly horizontally with "
            "equal spacing between them."
        ),
        'align_idMso': 'ObjectsAlignTopSmart',
        'distribute_idMso': 'AlignDistributeHorizontally',
    },
    'variant_align_center_distribute_v': {
        'type': 'align',
        'description': 'Align horizontal centers (X) + distribute vertically (Y)',
        'instruction': (
            "Align all three rectangles so their horizontal centers "
            "(left-right midpoints) are at the same horizontal position, "
            "and distribute all three rectangles evenly vertically with "
            "equal spacing between them."
        ),
        'align_idMso': 'ObjectsAlignCenterHorizontalSmart',
        'distribute_idMso': 'AlignDistributeVertically',
    },
    'variant_align_middle_distribute_h': {
        'type': 'align',
        'description': 'Align vertical middles (Y) + distribute horizontally (X)',
        'instruction': (
            "Align all three rectangles so their vertical centers "
            "(top-bottom midpoints) are at the same vertical position, "
            "and distribute all three rectangles evenly horizontally with "
            "equal spacing between them."
        ),
        'align_idMso': 'ObjectsAlignMiddleVerticalSmart',
        'distribute_idMso': 'AlignDistributeHorizontally',
    },
    # --- Rotate + distribute (rotation-agnostic) ---
    'variant_rotate_cw_distribute_h': {
        'type': 'rotate',
        'description': 'Rotate middle shape 90° clockwise + distribute horizontally',
        'instruction': (
            "Rotate the middle rectangle 90° clockwise, "
            "and distribute all three rectangles evenly horizontally with "
            "equal spacing between them."
        ),
        'rotate_angle': 90,
        'distribute_type': 0,  # msoDistributeHorizontally
    },
    'variant_rotate_ccw_distribute_v': {
        'type': 'rotate',
        'description': 'Rotate middle shape 90° counter-clockwise + distribute vertically',
        'instruction': (
            "Rotate the middle rectangle 90° counter-clockwise, "
            "and distribute all three rectangles evenly vertically with "
            "equal spacing between them."
        ),
        'rotate_angle': -90,
        'distribute_type': 1,  # msoDistributeVertically
    },
}


# ============================================================================
# Instruction registry (command name → variants dict)
# ============================================================================

_COMMAND_VARIANTS = [
    ('AdditiveShapeComposition', SHAPE_VARIANTS),
    ('AdditiveChartComposition', CHART_VARIANTS),
    ('AdditivePictureComposition', PICTURE_VARIANTS),
    ('AdditiveSmartArtComposition', SMARTART_VARIANTS),
    ('AdditiveTableComposition', TABLE_VARIANTS),
    ('AdditiveLayoutComposition', LAYOUT_VARIANTS),
]


# Parameter fragments that must be visible to the solver whenever a
# composition links to the corresponding exact atomic case.  This is kept
# deliberately small and covers only slots whose display name is not enough
# to recover the atomic target value.
_ATOMIC_INSTRUCTION_FRAGMENTS = {
    ('title', 'above_chart'): ('"chart title"',),
    ('title', 'centered_overlay'): ('"chart title"',),
    ('fill_color', 'red'): ('red (rgb: 238, 0, 0)',),
    ('fill_color', 'blue'): ('blue (rgb: 0, 112, 192)',),
    ('outline_color', 'blue'): ('blue (rgb: 0, 112, 192)',),
    ('outline_color', 'green'): ('green (rgb: 0, 176, 80)',),
    ('glow', 'glow_11pt_accent_4'): ('accent 4',),
    ('pic_glow', 'glow_11pt_accent_4'): ('accent 4',),
    ('ext_glow', 'glow_11pt_accent_4'): ('accent 4',),
    ('pic_outline', 'red'): ('red outline (rgb: 238, 0, 0)',),
    ('series_color', 'blue'): (
        'outline color of the first data series', 'blue (rgb: 0, 0, 255)',
    ),
}


def _validate_atomic_instruction_alignment(variant_name, config):
    """Reject a composition prompt that hides an exact linked-atomic value."""
    instruction = str(config.get('instruction') or '').casefold()
    missing = []
    for slot_id, value in (config.get('slots') or {}).items():
        for fragment in _ATOMIC_INSTRUCTION_FRAGMENTS.get((slot_id, value), ()):
            if fragment not in instruction:
                missing.append(fragment)
    if missing:
        raise ValueError(
            f'{variant_name} instruction omits linked-atomic parameter(s): '
            + ', '.join(repr(fragment) for fragment in missing)
        )


def _write_all_instructions():
    """Write instructions JSON from variant defs (single source of truth)."""
    commands = {}
    for cmd_name, variants in _COMMAND_VARIANTS:
        commands[cmd_name] = {
            vname: vdef['instruction'] for vname, vdef in variants.items()
        }
    write_instructions_json(TEST_SET, commands)


# ============================================================================
# Generator class
# ============================================================================

class AdditiveObjectEffectGenerator:
    """Generator for object_effect x additive_composition test cases."""

    def __init__(self, output_dir: Path = None, hidden: bool = False):
        self.output_dir = output_dir or (BENCHMARK_OUTPUT_DIR / TEST_SET)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.inputs_dir = self.output_dir / 'inputs'
        self.inputs_dir.mkdir(parents=True, exist_ok=True)
        self.hidden = hidden

    def _get_shared_input(self, input_name, wd=None):
        """Get path to shared input file.

        Searches existing test-set input directories, then falls back to
        creating the input via helpers (wd required for those fallbacks).
        """
        dest = self.inputs_dir / f'{input_name}.docx'
        if dest.exists():
            return dest

        search_dirs = [
            BENCHMARK_OUTPUT_DIR / 'object_effect_x_contextual_parameterized' / 'inputs',
            BENCHMARK_OUTPUT_DIR / 'object_effect_x_core_parameterized' / 'inputs',
            BENCHMARK_OUTPUT_DIR / 'object_effect_x_core_actionable' / 'inputs',
            BENCHMARK_OUTPUT_DIR / 'interaction_composition_x_layout_coupled' / 'inputs',
        ]
        for src_dir in search_dirs:
            src = src_dir / f'{input_name}.docx'
            if src.exists():
                shutil.copy2(src, dest)
                print(f"  Copied input: {input_name}.docx")
                return dest

        # Fallback: create from helpers if wd provided.
        if wd is not None:
            from generator.compositions.interaction.shared import (
                create_landscape_picture,
                create_spread_tall_shapes_doc,
                create_tall_shapes_doc,
            )
            if input_name == 'doc_in_three_shapes_tall_spread':
                create_spread_tall_shapes_doc(wd, dest)
                print(f"  Created: {input_name}.docx")
                return dest
            if input_name == 'doc_in_three_shapes_tall':
                create_tall_shapes_doc(wd, dest)
                print(f"  Created: {input_name}.docx")
                return dest
            if input_name == 'doc_in_landscape_picture':
                doc = wd.Documents.Add()
                time.sleep(0.3)
                create_landscape_picture(doc)
                doc.SaveAs(str(dest.absolute()), FileFormat=16)
                doc.Close(False)
                time.sleep(0.3)
                print(f"  Created: {input_name}.docx")
                return dest
            if input_name == 'doc_in_table_structure_contents':
                self._create_table_structure_input(wd, dest, mode='contents')
                print(f"  Created: {input_name}.docx")
                return dest
            if input_name == 'doc_in_table_structure_window':
                self._create_table_structure_input(wd, dest, mode='window')
                print(f"  Created: {input_name}.docx")
                return dest

        raise FileNotFoundError(
            f"Shared input '{input_name}.docx' not found in any benchmark inputs directory"
        )

    def _create_table_structure_input(self, wd, dest: Path, mode: str):
        """Create a table input with visible preconditions for structure compositions."""
        doc = wd.Documents.Add()
        time.sleep(0.3)

        try:
            table = doc.Tables.Add(
                Range=wd.Selection.Range,
                NumRows=3,
                NumColumns=3,
            )

            cell_text = [
                ["ID", "Workstream", "Reviewer notes"],
                ["A1", "Q2 budget", "Approve once final numbers are confirmed."],
                ["C12", "Regional operations planning", "Needs a second pass on supplier timing and staffing assumptions."],
            ]
            for row_idx, row_values in enumerate(cell_text, start=1):
                for col_idx, text in enumerate(row_values, start=1):
                    cell = table.Cell(row_idx, col_idx)
                    cell.Range.Text = text
                    try:
                        cell.VerticalAlignment = 0  # wdCellAlignVerticalTop
                        cell.Range.ParagraphFormat.Alignment = 0  # wdAlignParagraphLeft
                    except Exception:
                        pass

            try:
                table.AutoFitBehavior(0)  # wdAutoFitFixed
            except Exception:
                pass

            try:
                if mode == 'contents':
                    for col in table.Columns:
                        col.Width = 144
                    row_heights = [36, 72, 108]
                    for row_idx, height in enumerate(row_heights, start=1):
                        table.Rows(row_idx).Height = height
                        table.Rows(row_idx).HeightRule = 1  # wdRowHeightExactly
                    table.AutoFitBehavior(2)  # wdAutoFitWindow
                elif mode == 'window':
                    uneven_widths = [60, 150, 260]
                    for col_idx, width in enumerate(uneven_widths, start=1):
                        table.Columns(col_idx).Width = width
                    for row in table.Rows:
                        row.Height = 72
                        row.HeightRule = 1  # wdRowHeightExactly
                else:
                    raise ValueError(f"Unknown table structure input mode: {mode}")
            except Exception:
                pass

            _apply_visible_table_borders(table)
            time.sleep(0.3)
            wd.Selection.HomeKey(Unit=6)
            doc.SaveAs(str(dest.absolute()), FileFormat=16)
            time.sleep(0.1)
        finally:
            doc.Close(False)
            time.sleep(0.3)

    def _get_shape(self, doc):
        """Get the first shape from the document."""
        if doc.Shapes.Count == 0:
            raise RuntimeError("No shapes found in document")
        return doc.Shapes(1)

    def _get_chart(self, wd, doc):
        """Get the first chart from the document."""
        for i in range(1, doc.InlineShapes.Count + 1):
            if doc.InlineShapes(i).HasChart:
                return doc.InlineShapes(i).Chart
        for i in range(1, doc.Shapes.Count + 1):
            if doc.Shapes(i).HasChart:
                return doc.Shapes(i).Chart
        raise RuntimeError("No chart found in document")

    def _generate_shape_variant(self, wd, case_name, variant_name, config):
        """Generate a single shape composition variant."""
        variant_dir = self.output_dir / case_name / variant_name
        variant_dir.mkdir(parents=True, exist_ok=True)
        gold_path = variant_dir / 'gold.docx'
        spec_path = variant_dir / 'spec.json'

        slots = config['slots']
        shared_input = self._get_shared_input('doc_in_shape')

        doc = wd.Documents.Open(str(shared_input.absolute()))
        try:
            shape = self._get_shape(doc)
            shape.Select()
            time.sleep(0.3)

            for slot_id, value in slots.items():
                _apply_shape_extended_slot(shape, slot_id, value)

            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
        finally:
            doc.Close(False)
        time.sleep(0.3)

        if slots.get('glow') == 'glow_11pt_accent_4':
            from generator.test_sets.object_effect_contextual_parameterized import (
                _fix_glow_scheme_color,
            )
            _fix_glow_scheme_color(gold_path, 0xFFC000, 'accent4')

        atomic_intents = build_object_effect_atomic_intents(slots)

        spec = {
            'behavior': 'object_effect',
            'scope': 'additive_composition',
            'variation': variant_name,
            'generated_at': datetime.now().isoformat(),
            'composition': {
                'type': 'additive',
                'atomic_intents': atomic_intents,
                'operation_count': len(slots),
            },
            'precondition': {
                'type': 'shape',
                'select_object': 'shape_1',
            },
            'verification': {
                'primary': {
                    'method': 'shape_property_comparison',
                    'description': 'Word COM shape property comparison with gold',
                },
                'invariants': [
                    {
                        'method': 'structure_preserved',
                        'description': 'Document structure preserved',
                    },
                ],
            },
            'files': {
                'doc_in': '../../inputs/doc_in_shape.docx',
                'gold': 'gold.docx',
            },
        }
        write_spec_if_changed(spec_path, spec)

        print(f"  [{case_name}/{variant_name}] Generated ({len(slots)} operations)")
        return True

    def _generate_chart_variant(self, wd, case_name, variant_name, config):
        """Generate a single chart composition variant (bar, extended, or line)."""
        variant_dir = self.output_dir / case_name / variant_name
        variant_dir.mkdir(parents=True, exist_ok=True)
        gold_path = variant_dir / 'gold.docx'
        spec_path = variant_dir / 'spec.json'

        slots = config['slots']
        doc_in_name = config.get('doc_in_name', 'doc_in_chart')
        shared_input = self._get_shared_input(doc_in_name)

        doc = wd.Documents.Open(str(shared_input.absolute()))
        try:
            chart = self._get_chart(wd, doc)

            for slot_id, value in slots.items():
                _apply_line_chart_slot(wd, chart, slot_id, value)

            # Deactivate chart before saving
            try:
                wd.Selection.HomeKey(Unit=6)
            except Exception:
                pass
            time.sleep(0.5)

            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
        finally:
            doc.Close(False)
        time.sleep(0.3)

        atomic_intents = build_object_effect_atomic_intents(slots)
        select_object = (
            'chart_line_1' if doc_in_name.startswith('doc_in_chart_line')
            else 'chart_1'
        )

        spec = {
            'behavior': 'object_effect',
            'scope': 'additive_composition',
            'variation': variant_name,
            'generated_at': datetime.now().isoformat(),
            'composition': {
                'type': 'additive',
                'atomic_intents': atomic_intents,
                'operation_count': len(slots),
            },
            'precondition': {
                'type': 'chart',
                'select_object': select_object,
            },
            'verification': {
                'primary': {
                    'method': 'shape_property_comparison',
                    'description': 'Word COM chart property comparison with gold',
                },
                'invariants': [
                    {
                        'method': 'structure_preserved',
                        'description': 'Document structure preserved',
                    },
                ],
            },
            'files': {
                'doc_in': f'../../inputs/{doc_in_name}.docx',
                'gold': 'gold.docx',
            },
        }
        write_spec_if_changed(spec_path, spec)

        print(f"  [{case_name}/{variant_name}] Generated ({len(slots)} operations)")
        return True

    def _generate_picture_variant(self, wd, case_name, variant_name, config):
        """Generate a single picture composition variant."""
        variant_dir = self.output_dir / case_name / variant_name
        variant_dir.mkdir(parents=True, exist_ok=True)
        gold_path = variant_dir / 'gold.docx'
        spec_path = variant_dir / 'spec.json'

        slots = config['slots']
        shared_input = self._get_shared_input('doc_in_picture')

        doc = wd.Documents.Open(str(shared_input.absolute()))
        try:
            pic = doc.InlineShapes(1)
            for slot_id, value in slots.items():
                _apply_picture_slot(pic, slot_id, value)
            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
        finally:
            doc.Close(False)
        time.sleep(0.3)

        if slots.get('pic_glow') == 'glow_11pt_accent_4':
            from generator.test_sets.object_effect_contextual_parameterized import (
                _fix_glow_scheme_color,
            )
            _fix_glow_scheme_color(gold_path, 0xFFC000, 'accent4')

        atomic_intents = build_object_effect_atomic_intents(slots)
        spec = {
            'behavior': 'object_effect',
            'scope': 'additive_composition',
            'variation': variant_name,
            'generated_at': datetime.now().isoformat(),
            'composition': {
                'type': 'additive',
                'atomic_intents': atomic_intents,
                'operation_count': len(slots),
            },
            'precondition': {
                'type': 'picture',
                'select_object': 'picture_1',
            },
            'verification': {
                'primary': {
                    'method': 'shape_property_comparison',
                    'description': 'Word COM picture property comparison with gold',
                },
                'invariants': [
                    {'method': 'structure_preserved', 'description': 'Document structure preserved'},
                ],
            },
            'files': {
                'doc_in': '../../inputs/doc_in_picture.docx',
                'gold': 'gold.docx',
            },
        }
        write_spec_if_changed(spec_path, spec)
        print(f"  [{case_name}/{variant_name}] Generated ({len(slots)} operations)")
        return True

    def _generate_smartart_variant(self, wd, case_name, variant_name, config):
        """Generate a single SmartArt composition variant."""
        variant_dir = self.output_dir / case_name / variant_name
        variant_dir.mkdir(parents=True, exist_ok=True)
        gold_path = variant_dir / 'gold.docx'
        spec_path = variant_dir / 'spec.json'

        slots = config['slots']
        shared_input = self._get_shared_input('doc_in_smartart')

        doc = wd.Documents.Open(str(shared_input.absolute()))
        try:
            # Find the SmartArt object
            smartart = None
            for i in range(1, doc.Shapes.Count + 1):
                if doc.Shapes(i).HasSmartArt:
                    smartart = doc.Shapes(i).SmartArt
                    doc.Shapes(i).Select()
                    time.sleep(0.3)
                    break
            if smartart is None:
                raise RuntimeError("No SmartArt found in document")

            for slot_id, value in slots.items():
                _apply_smartart_slot(wd, smartart, slot_id, value)

            try:
                wd.Selection.HomeKey(Unit=6)
            except Exception:
                pass
            time.sleep(0.5)

            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
        finally:
            doc.Close(False)
        time.sleep(0.3)

        atomic_intents = build_object_effect_atomic_intents(slots)
        spec = {
            'behavior': 'object_effect',
            'scope': 'additive_composition',
            'variation': variant_name,
            'generated_at': datetime.now().isoformat(),
            'composition': {
                'type': 'additive',
                'atomic_intents': atomic_intents,
                'operation_count': len(slots),
            },
            'precondition': {
                'type': 'smartart',
                'select_object': 'smartart_1',
            },
            'verification': {
                'primary': {
                    'method': 'shape_property_comparison',
                    'description': 'Word COM SmartArt property comparison with gold',
                },
                'invariants': [
                    {'method': 'structure_preserved', 'description': 'Document structure preserved'},
                ],
            },
            'files': {
                'doc_in': '../../inputs/doc_in_smartart.docx',
                'gold': 'gold.docx',
            },
        }
        write_spec_if_changed(spec_path, spec)
        print(f"  [{case_name}/{variant_name}] Generated ({len(slots)} operations)")
        return True

    def _generate_table_variant(self, wd, case_name, variant_name, config):
        """Generate a single table composition variant."""
        variant_dir = self.output_dir / case_name / variant_name
        variant_dir.mkdir(parents=True, exist_ok=True)
        gold_path = variant_dir / 'gold.docx'
        spec_path = variant_dir / 'spec.json'

        slots = config['slots']
        doc_in_name = config.get('doc_in_name', 'doc_in_table')
        shared_input = self._get_shared_input(doc_in_name, wd=wd)

        doc = wd.Documents.Open(str(shared_input.absolute()))
        try:
            table = doc.Tables(1)
            for slot_id, value in slots.items():
                _apply_table_slot(table, slot_id, value)
            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
        finally:
            doc.Close(False)
        time.sleep(0.3)

        atomic_intents = build_object_effect_atomic_intents(slots)
        spec = {
            'behavior': 'object_effect',
            'scope': 'additive_composition',
            'variation': variant_name,
            'generated_at': datetime.now().isoformat(),
            'composition': {
                'type': 'additive',
                'atomic_intents': atomic_intents,
                'operation_count': len(slots),
            },
            'precondition': {
                'type': 'table',
                'select_object': 'table_1',
            },
            'verification': {
                'primary': {
                    'method': 'shape_property_comparison',
                    'description': 'Word COM table property comparison with gold',
                },
                'invariants': [
                    {'method': 'structure_preserved', 'description': 'Document structure preserved'},
                ],
            },
            'files': {
                'doc_in': f'../../inputs/{doc_in_name}.docx',
                'gold': 'gold.docx',
            },
        }
        write_spec_if_changed(spec_path, spec)
        print(f"  [{case_name}/{variant_name}] Generated ({len(slots)} operations)")
        return True

    # ==== New composition generators (additive extension plan) ====

    def _generate_picture_extended_variant(self, wd, case_name, variant_name, config):
        """Generate a single picture extended composition variant."""
        variant_dir = self.output_dir / case_name / variant_name
        variant_dir.mkdir(parents=True, exist_ok=True)
        gold_path = variant_dir / 'gold.docx'
        spec_path = variant_dir / 'spec.json'

        slots = config['slots']
        shared_input = self._get_shared_input('doc_in_picture')

        doc = wd.Documents.Open(str(shared_input.absolute()))
        try:
            pic = doc.InlineShapes(1)
            for slot_id, value in slots.items():
                pic = _apply_picture_extended_slot(wd, pic, slot_id, value) or pic
            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
        finally:
            doc.Close(False)
        time.sleep(0.3)

        if slots.get('ext_glow') == 'glow_11pt_accent_4':
            from generator.test_sets.object_effect_contextual_parameterized import (
                _fix_glow_scheme_color,
            )
            _fix_glow_scheme_color(gold_path, 0xFFC000, 'accent4')

        atomic_intents = build_object_effect_atomic_intents(slots)
        spec = {
            'behavior': 'object_effect',
            'scope': 'additive_composition',
            'variation': variant_name,
            'generated_at': datetime.now().isoformat(),
            'composition': {
                'type': 'additive',
                'atomic_intents': atomic_intents,
                'operation_count': len(slots),
            },
            'precondition': {
                'type': 'picture',
                'select_object': 'picture_1',
            },
            'verification': {
                'primary': {
                    'method': 'shape_property_comparison',
                    'description': 'Word COM picture property comparison with gold',
                },
                'invariants': [
                    {'method': 'structure_preserved', 'description': 'Document structure preserved'},
                ],
            },
            'files': {
                'doc_in': '../../inputs/doc_in_picture.docx',
                'gold': 'gold.docx',
            },
        }
        write_spec_if_changed(spec_path, spec)
        print(f"  [{case_name}/{variant_name}] Generated ({len(slots)} operations)")
        return True

    def _generate_smartart_combined_variant(self, wd, case_name, variant_name, config):
        """Generate a single SmartArt combined composition variant."""
        variant_dir = self.output_dir / case_name / variant_name
        variant_dir.mkdir(parents=True, exist_ok=True)
        gold_path = variant_dir / 'gold.docx'
        spec_path = variant_dir / 'spec.json'

        slots = config['slots']
        shared_input = self._get_shared_input('doc_in_smartart')

        doc = wd.Documents.Open(str(shared_input.absolute()))
        try:
            smartart_shape = None
            for i in range(1, doc.Shapes.Count + 1):
                if doc.Shapes(i).HasSmartArt:
                    smartart_shape = doc.Shapes(i)
                    smartart_shape.Select()
                    time.sleep(0.3)
                    break
            if smartart_shape is None:
                raise RuntimeError("No SmartArt found in document")

            for slot_id, value in slots.items():
                _apply_smartart_combined_slot(wd, smartart_shape, slot_id, value)

            try:
                wd.Selection.HomeKey(Unit=6)
            except Exception:
                pass
            time.sleep(0.5)

            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
        finally:
            doc.Close(False)
        time.sleep(0.3)

        atomic_intents = build_object_effect_atomic_intents(slots)
        spec = {
            'behavior': 'object_effect',
            'scope': 'additive_composition',
            'variation': variant_name,
            'generated_at': datetime.now().isoformat(),
            'composition': {
                'type': 'additive',
                'atomic_intents': atomic_intents,
                'operation_count': len(slots),
            },
            'precondition': {
                'type': 'smartart',
                'select_object': 'smartart_1',
            },
            'verification': {
                'primary': {
                    'method': 'shape_property_comparison',
                    'description': 'Word COM SmartArt sub-element property comparison with gold',
                },
                'invariants': [
                    {'method': 'structure_preserved', 'description': 'Document structure preserved'},
                ],
            },
            'files': {
                'doc_in': '../../inputs/doc_in_smartart.docx',
                'gold': 'gold.docx',
            },
        }
        write_spec_if_changed(spec_path, spec)
        print(f"  [{case_name}/{variant_name}] Generated ({len(slots)} operations)")
        return True

    def _generate_table_structure_variant(self, wd, case_name, variant_name, config):
        """Generate a single table structure composition variant."""
        variant_dir = self.output_dir / case_name / variant_name
        variant_dir.mkdir(parents=True, exist_ok=True)
        gold_path = variant_dir / 'gold.docx'
        spec_path = variant_dir / 'spec.json'

        slots = config['slots']
        doc_in_name = config.get('doc_in_name', 'doc_in_table')
        shared_input = self._get_shared_input(doc_in_name, wd=wd)

        doc = wd.Documents.Open(str(shared_input.absolute()))
        try:
            table = doc.Tables(1)
            for slot_id, value in slots.items():
                _apply_table_structure_slot(wd, table, slot_id, value)
            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
        finally:
            doc.Close(False)
        time.sleep(0.3)

        atomic_intents = build_object_effect_atomic_intents(slots)

        spec = {
            'behavior': 'object_effect',
            'scope': 'additive_composition',
            'variation': variant_name,
            'generated_at': datetime.now().isoformat(),
            'composition': {
                'type': 'additive',
                'atomic_intents': atomic_intents,
                'operation_count': len(slots),
            },
            'precondition': {
                'type': 'table',
                'select_object': 'table_1',
            },
            'verification': {
                'primary': {
                    'method': 'shape_property_comparison',
                    'description': 'Word COM table property comparison with gold',
                },
                'invariants': [
                    {'method': 'structure_preserved', 'description': 'Document structure preserved'},
                ],
            },
            'files': {
                'doc_in': f'../../inputs/{doc_in_name}.docx',
                'gold': 'gold.docx',
            },
        }
        write_spec_if_changed(spec_path, spec)
        print(f"  [{case_name}/{variant_name}] Generated ({len(slots)} operations)")
        return True

    # ==== Multi-shape / picture layout-independent additive generators ====

    def _select_all_shapes(self, doc):
        """Multi-select all floating shapes in the document."""
        if doc.Shapes.Count == 0:
            return
        doc.Shapes(1).Select()
        for i in range(2, doc.Shapes.Count + 1):
            doc.Shapes(i).Select(False)
        time.sleep(0.2)

    def _generate_align_distribute_variant(self, wd, case_name, variant_name, config):
        """Generate an additive align+distribute composition on three shapes.

        Both operations are executed via ExecuteMso on the multi-selection,
        matching what a user would do from the ribbon. Order-independent
        because the two operations target orthogonal axes.
        """
        variant_dir = self.output_dir / case_name / variant_name
        variant_dir.mkdir(parents=True, exist_ok=True)
        gold_path = variant_dir / 'gold.docx'
        spec_path = variant_dir / 'spec.json'

        shared_input = self._get_shared_input('doc_in_three_shapes_uneven', wd=wd)

        doc = wd.Documents.Open(str(shared_input.absolute()))
        try:
            # Step A: align
            self._select_all_shapes(doc)
            wd.CommandBars.ExecuteMso(config['align_idMso'])
            time.sleep(0.5)
            # Step B: distribute
            self._select_all_shapes(doc)
            wd.CommandBars.ExecuteMso(config['distribute_idMso'])
            time.sleep(0.5)
            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
        finally:
            doc.Close(False)
        time.sleep(0.3)

        intent_slots = {
            'align_command': config['align_idMso'],
            'distribute_command': config['distribute_idMso'],
        }
        atomic_intents = build_object_effect_atomic_intents(intent_slots)
        spec = {
            'behavior': 'object_effect',
            'scope': 'additive_composition',
            'variation': variant_name,
            'generated_at': datetime.now().isoformat(),
            'composition': {
                'type': 'additive',
                'atomic_intents': atomic_intents,
                'operation_count': 2,
                'target_object_type': 'multi_shape',
            },
            'precondition': {
                'type': 'three_shapes_uneven',
                'description': (
                    'Document with three 100x100 rectangles at uneven '
                    'positions so both align and distribute produce visible '
                    'changes'
                ),
            },
            'verification': {
                'primary': {
                    'method': 'shape_property_comparison',
                    'description': (
                        'Multi-shape position comparison with gold '
                        '(positions, rotation, wrap)'
                    ),
                },
                'invariants': [
                    {
                        'method': 'structure_preserved',
                        'description': 'Document structure preserved',
                    },
                ],
            },
            'files': {
                'doc_in': '../../inputs/doc_in_three_shapes_uneven.docx',
                'gold': 'gold.docx',
            },
        }
        write_spec_if_changed(spec_path, spec)
        print(f"  [{case_name}/{variant_name}] Generated (2 operations)")
        return True

    def _generate_rotate_distribute_variant(self, wd, case_name, variant_name, config):
        """Generate an additive rotate+distribute composition on three shapes.

        Uses ShapeRange.Distribute (not ExecuteMso) so the behavior matches
        what the layout_coupled empirical test established: distribute is
        rotation-agnostic (operates on un-rotated .Left/.Top), so the two
        operations don't interact.
        """
        variant_dir = self.output_dir / case_name / variant_name
        variant_dir.mkdir(parents=True, exist_ok=True)
        gold_path = variant_dir / 'gold.docx'
        spec_path = variant_dir / 'spec.json'

        shared_input = self._get_shared_input(
            'doc_in_three_shapes_tall_spread', wd=wd)

        doc = wd.Documents.Open(str(shared_input.absolute()))
        try:
            # Step A: rotate middle shape
            doc.Shapes(2).IncrementRotation(config['rotate_angle'])
            time.sleep(0.3)
            # Step B: distribute
            shape_range = doc.Shapes.Range(
                list(range(1, doc.Shapes.Count + 1))
            )
            shape_range.Distribute(config['distribute_type'], False)
            time.sleep(0.5)
            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
        finally:
            doc.Close(False)
        time.sleep(0.3)

        distribute_cmd = (
            'AlignDistributeHorizontally'
            if config['distribute_type'] == 0
            else 'AlignDistributeVertically'
        )
        intent_slots = {
            'rotate_angle': config['rotate_angle'],
            'distribute_command': distribute_cmd,
        }
        atomic_intents = build_object_effect_atomic_intents(intent_slots)
        spec = {
            'behavior': 'object_effect',
            'scope': 'additive_composition',
            'variation': variant_name,
            'generated_at': datetime.now().isoformat(),
            'composition': {
                'type': 'additive',
                'atomic_intents': atomic_intents,
                'operation_count': 2,
                'target_object_type': 'multi_shape',
            },
            'precondition': {
                'type': 'three_shapes_tall_spread',
                'description': (
                    'Document with three 60x120 tall rectangles unevenly '
                    'spread at (50,50), (170,200), (400,550)'
                ),
            },
            'verification': {
                'primary': {
                    'method': 'shape_property_comparison',
                    'description': (
                        'Multi-shape position + rotation comparison with gold'
                    ),
                },
                'invariants': [
                    {
                        'method': 'structure_preserved',
                        'description': 'Document structure preserved',
                    },
                ],
            },
            'files': {
                'doc_in': '../../inputs/doc_in_three_shapes_tall_spread.docx',
                'gold': 'gold.docx',
            },
        }
        write_spec_if_changed(spec_path, spec)
        print(f"  [{case_name}/{variant_name}] Generated (2 operations)")
        return True

    def _generate_crop_effects_variant(self, wd, case_name, variant_name, config):
        """Generate an additive picture crop + effect composition.

        Crop changes PictureFormat.CropLeft/Right/Top/Bottom; shadow and
        reflection set Shadow.Type / Reflection.Type. These touch disjoint
        property slots, and shadow/reflection automatically follow the
        cropped outline, so the composition is order-independent.
        """
        variant_dir = self.output_dir / case_name / variant_name
        variant_dir.mkdir(parents=True, exist_ok=True)
        gold_path = variant_dir / 'gold.docx'
        spec_path = variant_dir / 'spec.json'

        shared_input = self._get_shared_input('doc_in_landscape_picture', wd=wd)

        doc = wd.Documents.Open(str(shared_input.absolute()))
        try:
            pic = doc.InlineShapes(1)

            # Step A: centered crop to target aspect ratio
            target_ratio = config['ratio']
            pic_w = pic.Width
            pic_h = pic.Height
            current_ratio = pic_w / pic_h
            if current_ratio > target_ratio:
                new_w = pic_h * target_ratio
                crop = (pic_w - new_w) / 2
                pic.PictureFormat.CropLeft = crop
                pic.PictureFormat.CropRight = crop
            else:
                new_h = pic_w / target_ratio
                crop = (pic_h - new_h) / 2
                pic.PictureFormat.CropTop = crop
                pic.PictureFormat.CropBottom = crop
            time.sleep(0.3)

            # Step B: visual effect
            if config['effect'] == 'shadow':
                pic.Shadow.Visible = True
                pic.Shadow.Type = config['shadow_type']
            elif config['effect'] == 'reflection':
                pic.Reflection.Type = config['reflection_type']
            time.sleep(0.3)

            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
        finally:
            doc.Close(False)
        time.sleep(0.3)

        effect_cmd = (
            'PictureEffectsShadowGallery'
            if config['effect'] == 'shadow'
            else 'PictureReflectionGallery'
        )
        # Crop is performed via PictureFormat.CropLeft etc. — there is no
        # single ribbon idMso for "crop to 16:9"; the ribbon uses
        # PictureCropGallery or the crop handles.
        ratio_key = '16_9' if abs(config['ratio'] - (16 / 9)) < 0.01 else '1_1'
        if config['effect'] == 'shadow':
            effect_slot = ('pic_shadow', 'outer_center')
        else:
            effect_slot = ('pic_reflection', 'tight_touching')
        intent_slots = {
            'crop_aspect_ratio': ratio_key,
            effect_slot[0]: effect_slot[1],
        }
        atomic_intents = build_object_effect_atomic_intents(intent_slots)
        spec = {
            'behavior': 'object_effect',
            'scope': 'additive_composition',
            'variation': variant_name,
            'generated_at': datetime.now().isoformat(),
            'composition': {
                'type': 'additive',
                'atomic_intents': atomic_intents,
                'operation_count': 2,
                'target_object_type': 'picture',
            },
            'precondition': {
                'type': 'landscape_picture',
                'description': (
                    'Document with a 640x480 landscape picture (4:3) with '
                    'colored quadrant markers for crop verification'
                ),
            },
            'verification': {
                'primary': {
                    'method': 'shape_property_comparison',
                    'description': (
                        'Picture effect and crop property comparison with gold'
                    ),
                },
                'invariants': [
                    {
                        'method': 'structure_preserved',
                        'description': 'Document structure preserved',
                    },
                ],
            },
            'files': {
                'doc_in': '../../inputs/doc_in_landscape_picture.docx',
                'gold': 'gold.docx',
            },
        }
        write_spec_if_changed(spec_path, spec)
        print(f"  [{case_name}/{variant_name}] Generated (2 operations)")
        return True

    def _generate_variant(self, wd, case_name, variant_name, config, results):
        """Generate a single variant, dispatching to the correct generator by type."""
        key = f'{case_name}/{variant_name}'
        try:
            vtype = config.get('type', 'base')

            if case_name == 'AdditiveShapeComposition':
                ok = self._generate_shape_variant(wd, case_name, variant_name, config)

            elif case_name == 'AdditiveChartComposition':
                ok = self._generate_chart_variant(wd, case_name, variant_name, config)

            elif case_name == 'AdditivePictureComposition':
                if vtype == 'extended':
                    ok = self._generate_picture_extended_variant(wd, case_name, variant_name, config)
                elif vtype == 'crop':
                    ok = self._generate_crop_effects_variant(wd, case_name, variant_name, config)
                else:
                    ok = self._generate_picture_variant(wd, case_name, variant_name, config)

            elif case_name == 'AdditiveSmartArtComposition':
                if vtype == 'combined':
                    ok = self._generate_smartart_combined_variant(wd, case_name, variant_name, config)
                else:
                    ok = self._generate_smartart_variant(wd, case_name, variant_name, config)

            elif case_name == 'AdditiveTableComposition':
                if vtype == 'structure':
                    ok = self._generate_table_structure_variant(wd, case_name, variant_name, config)
                else:
                    ok = self._generate_table_variant(wd, case_name, variant_name, config)

            elif case_name == 'AdditiveLayoutComposition':
                if vtype == 'rotate':
                    ok = self._generate_rotate_distribute_variant(wd, case_name, variant_name, config)
                else:
                    ok = self._generate_align_distribute_variant(wd, case_name, variant_name, config)

            else:
                raise ValueError(f"Unknown composition: {case_name}")

            results[key] = ok
        except Exception as e:
            print(f"  [ERROR] {variant_name}: {e}")
            import traceback
            traceback.print_exc()
            results[key] = False

    def generate_all(self, command: str = None, variant_names=None):
        """Generate all object effect composition test cases."""
        requested_variants = set(variant_names or [])
        generated_variants = set()
        wd = ensure_word(visible=not self.hidden)
        try:
            results = {}

            for case_name, variants in _COMMAND_VARIANTS:
                if command is not None and command != case_name:
                    continue
                print(f"\n=== {case_name} ===")
                for variant_name, config in variants.items():
                    if requested_variants and variant_name not in requested_variants:
                        continue
                    _validate_atomic_instruction_alignment(variant_name, config)
                    self._generate_variant(wd, case_name, variant_name, config, results)
                    generated_variants.add(variant_name)

            unknown_variants = requested_variants - generated_variants
            if unknown_variants:
                raise ValueError(
                    'Requested variant(s) not found under the selected command: '
                    + ', '.join(sorted(unknown_variants))
                )

            # Write instructions JSON (source of truth is in variant defs)
            _write_all_instructions()

            total = len(results)
            passed = sum(1 for v in results.values() if v)
            print(f"\n=== Summary: {passed}/{total} test cases generated ===")

            # A filtered run is not a test-set summary.  Preserve the last
            # complete summary rather than replacing it with a partial count.
            if command is None and not requested_variants:
                summary_path = self.output_dir / 'generation_summary.json'
                summary = {
                    'test_set': TEST_SET,
                    'generated_at': datetime.now().isoformat(),
                    'total': total,
                    'passed': passed,
                    'results': {
                        k: 'success' if v else 'failed'
                        for k, v in results.items()
                    },
                }
                with open(summary_path, 'w', encoding='utf-8') as f:
                    json.dump(summary, f, indent=2, ensure_ascii=False)

            return results
        finally:
            pass


def main():
    import argparse
    parser = argparse.ArgumentParser(
        description='Generate additive object effect composition test cases'
    )
    parser.add_argument('--command', '-c', help='Generate specific case only')
    parser.add_argument(
        '--variant', '-v', action='append',
        help='Generate only this variant (repeat for multiple variants)',
    )
    parser.add_argument('--output', '-o', help='Output directory')
    parser.add_argument('--hidden', action='store_true', help='Run Word hidden')
    args = parser.parse_args()

    output_dir = Path(args.output) if args.output else None
    gen = AdditiveObjectEffectGenerator(output_dir=output_dir, hidden=args.hidden)
    gen.generate_all(command=args.command, variant_names=args.variant)


if __name__ == '__main__':
    main()
