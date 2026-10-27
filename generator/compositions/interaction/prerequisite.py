"""Prerequisite chain generator for interaction-sensitive compositions.

Each case: Step A inserts/creates an object, Step B applies an operation to it.

Usage:
    python -m generator.compositions.run --test-set interaction_composition_x_prerequisite_chain
"""

import json
import time
import zipfile
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

from generator.config import BENCHMARK_OUTPUT_DIR
from generator.compositions.atomic_intents import (
    build_prerequisite_chain_atomic_intents,
)
from generator.execution.setup import ensure_word
from generator.test_sets.spec_utils import write_spec_if_changed

from .shared import (
    IMG_PATH,
    build_spec,
    close_doc,
    create_chart,
    create_equation,
    create_icon,
    create_picture,
    create_shape,
    create_smartart,
    create_stub_doc,
    create_table,
    create_textbox,
    create_empty_doc,
    get_smartart_layout,
    get_shared_input,
    open_doc,
    save_as_gold,
    write_instructions_json,
    write_spec,
)

TEST_SET = 'interaction_composition_x_prerequisite_chain'

SMARTART_LAYOUT_LABELS = [
    'Boss', 'Team1', 'Member1', 'Member2', 'Team2', 'Member3',
]
SMARTART_LAYOUT_LEVELS = [1, 2, 3, 3, 2, 3]
SMARTART_INITIAL_LAYOUT = 'Hierarchy'
SMARTART_FINAL_LAYOUT = 'Continuous Picture List'


def _validate_smartart_layout_snapshot(snapshot):
    """Return explicit contract failures for the approved IMP-17 fixture."""
    failures = []
    if not snapshot.get('found'):
        failures.append('native SmartArt object is missing')
    if snapshot.get('layout_name') != SMARTART_FINAL_LAYOUT:
        failures.append('final layout is not Continuous Picture List')
    if snapshot.get('all_node_count') != len(SMARTART_LAYOUT_LABELS):
        failures.append('SmartArt must contain exactly six data nodes')
    if snapshot.get('all_node_texts') != SMARTART_LAYOUT_LABELS:
        failures.append('SmartArt node text/order differs from the approved fixture')
    if snapshot.get('all_node_levels') != SMARTART_LAYOUT_LEVELS:
        failures.append('SmartArt hierarchy differs from the approved fixture')
    if not snapshot.get('below_heading'):
        failures.append("SmartArt is not visibly below 'Team Structure:'")
    return failures


def _smartart_package_snapshot(docx_path):
    """Inspect the saved package for an editable, non-flattened SmartArt."""
    a_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    with zipfile.ZipFile(docx_path) as package:
        names = set(package.namelist())
        data_parts = sorted(
            name for name in names
            if name.startswith('word/diagrams/data') and name.endswith('.xml')
        )
        layout_parts = sorted(
            name for name in names
            if name.startswith('word/diagrams/layout') and name.endswith('.xml')
        )
        drawing_parts = sorted(
            name for name in names
            if name.startswith('word/diagrams/drawing') and name.endswith('.xml')
        )
        document_xml = package.read('word/document.xml')
        labels = []
        for part in data_parts:
            root = ET.fromstring(package.read(part))
            labels.extend(
                (node.text or '').strip()
                for node in root.findall(f'.//{{{a_ns}}}t')
                if (node.text or '').strip()
            )
    return {
        'native_smartart_reference': b'relIds' in document_xml,
        'data_part_count': len(data_parts),
        'layout_part_count': len(layout_parts),
        'drawing_part_count': len(drawing_parts),
        'data_labels': labels,
    }


def _validate_smartart_package_snapshot(snapshot):
    failures = []
    if not snapshot.get('native_smartart_reference'):
        failures.append('document has no native SmartArt relationship reference')
    for key in ('data_part_count', 'layout_part_count', 'drawing_part_count'):
        if snapshot.get(key) != 1:
            failures.append(f'{key} must equal one')
    if snapshot.get('data_labels') != SMARTART_LAYOUT_LABELS:
        failures.append('saved SmartArt data labels differ from the approved fixture')
    return failures

# ─── Per-group contextual stub headings ──────────────────────────────────
#
# Each prerequisite group inserts into a doc_in that contains a single
# Heading 2 paragraph describing the *purpose* of the object being
# inserted.  This anchors the insertion in a semantic document context
# (e.g., a "Quarterly Sales:" heading makes a chart insertion feel like
# real document authoring rather than a sterile empty-canvas drop).
#
# The heading text is visible in doc_in, so the instruction can
# reference it by name and the agent can see a document-level intent.
STUB_HEADINGS = {
    'PrerequisiteTable':    'Project Status:',
    'PrerequisiteChart':    'Quarterly Sales:',
    'PrerequisiteIcon':     'Company Logo:',
    'PrerequisiteShape':    'Diagram:',
    'PrerequisiteSmartArt': 'Team Structure:',
    'PrerequisiteEquation': 'Formula:',
    'PrerequisiteTextBox':  'Note:',
}


# ═════════════════════════════════════════════════════════════════════════
# Case definitions — grouped by command/directory name
# ═════════════════════════════════════════════════════════════════════════

CASE_GROUPS = [
    {
        'group': 'PrerequisiteTable',
        'target_object_type': 'table',
        'behavior': 'object_effect',
        'verification_method': 'composition_dispatch',
        'composition_type': 'prerequisite',
        'variants': {
            'variant_style': {
                'steps': [
                    {'operation': 'insert_table', 'params': {'rows': 2, 'cols': 2}},
                    {'idMso': 'TableStylesGalleryWord',
                     'params': {'style_name': 'Grid Table 4 - Accent 1'}},
                ],
                'interaction_class': 'insert_then_style',
                'verification_description':
                    'Table existence + style name, borders, shading via _verify_table',
                'instruction':
                    "Below the 'Project Status:' heading, insert a 2x2 table with cell contents 'R1C1' through 'R2C2', then apply the 'Grid Table 4 - Accent 1' style to it (blue accent with alternating row shading).",
            },
            'variant_cell_align': {
                'steps': [
                    {'operation': 'insert_table', 'params': {'rows': 2, 'cols': 2}},
                    {'idMso': 'TableCellAlignMiddleCenter',
                     'params': {'scope': 'all_cells'}},
                ],
                'interaction_class': 'insert_then_align',
                'verification_description':
                    'Table existence + cell_alignments = middle-center via _verify_table',
                'instruction':
                    "Below the 'Project Status:' heading, insert a 2x2 table with cell contents 'R1C1' through 'R2C2', then set the cell alignment of all cells to middle-center.",
            },
            'variant_border': {
                'steps': [
                    {'operation': 'insert_table', 'params': {'rows': 2, 'cols': 2}},
                    {'idMso': 'BorderStylesGallery',
                     'params': {'line_style': 'double'}},
                ],
                'interaction_class': 'insert_then_border',
                'verification_description':
                    'Table existence + border line_style = double via _verify_table',
                'instruction':
                    "Below the 'Project Status:' heading, insert a 2x2 table with cell contents 'R1C1' through 'R2C2', then apply a double-line border style to the outer border of the table. Only change the outside border, not the inner cell borders.",
            },
            'variant_text_align': {
                'steps': [
                    {'operation': 'insert_table', 'params': {'rows': 2, 'cols': 2}},
                    {'idMso': 'AlignCenter',
                     'params': {'alignment': 'center'}},
                ],
                'interaction_class': 'insert_then_text_align',
                'verification_description':
                    'Table existence + paragraph alignment=center in cell R2C2 '
                    'via _verify_table',
                'instruction':
                    "Below the 'Project Status:' heading, insert a 2x2 table with cell contents 'R1C1' through 'R2C2', then center-align the text in cell R2C2 of the table.",
            },
        },
    },
    {
        'group': 'PrerequisiteChart',
        'target_object_type': 'chart',
        'behavior': 'object_effect',
        'verification_method': 'composition_dispatch',
        'composition_type': 'prerequisite',
        'variants': {
            'variant_title': {
                'steps': [
                    {'operation': 'insert_chart',
                     'params': {'chart_type': 51,
                                'categories': ['Q1', 'Q2'],
                                'series': [{'name': 'Revenue',
                                            'values': [100, 150]}]}},
                    {'idMso': 'ChartTitle', 'params': {'position': 'above_chart'}},
                ],
                'interaction_class': 'insert_then_title',
                'verification_description':
                    'Chart existence + HasTitle=True, title position via '
                    '_verify_chart_composition',
                'instruction':
                    "Below the 'Quarterly Sales:' heading, insert a clustered "
                    "column chart with categories 'Q1' and 'Q2', and one "
                    "series 'Revenue' with values 100 and 150, then add a "
                    "title \"Chart Title\" above the chart.",
            },
            'variant_legend': {
                'steps': [
                    {'operation': 'insert_chart',
                     'params': {'chart_type': 5,
                                'categories': ['A', 'B', 'C'],
                                'series': [{'name': 'Sales',
                                            'values': [40, 35, 25]}]}},
                    {'idMso': 'ChartLegend',
                     'params': {'position': 'right'}},  # xlLegendPositionRight=-4152
                ],
                'interaction_class': 'insert_then_legend',
                'verification_description':
                    'Chart existence + HasLegend=True, Legend.Position via '
                    '_verify_chart_composition',
                'instruction':
                    "Below the 'Quarterly Sales:' heading, insert a pie chart "
                    "with categories 'A', 'B', and 'C', and one series "
                    "'Sales' with values 40, 35, and 25, then move the "
                    "chart legend to the right side of the chart.",
            },
            'variant_trendline': {
                'steps': [
                    {'operation': 'insert_chart',
                     'params': {'chart_type': 4,
                                'categories': ['Jan', 'Feb', 'Mar'],
                                'series': [{'name': 'Temperature',
                                            'values': [10, 15, 20]}]}},
                    {'idMso': 'ChartTrendline',
                     'params': {'type': 'linear'}},  # xlLinear=-4132
                ],
                'interaction_class': 'insert_then_trendline',
                'verification_description':
                    'Chart existence + Trendlines.Count > 0 via '
                    '_verify_chart_composition',
                'instruction':
                    "Below the 'Quarterly Sales:' heading, insert a line chart "
                    "with categories 'Jan', 'Feb', and 'Mar', and one series "
                    "'Temperature' with values 10, 15, and 20, then add a "
                    "linear trendline to 'Temperature' in the chart.",
            },
            'variant_data_labels': {
                'steps': [
                    {'operation': 'insert_chart',
                     'params': {'chart_type': 51,
                                'categories': ['Q1', 'Q2'],
                                'series': [{'name': 'Revenue',
                                            'values': [100, 150]}]}},
                    {'idMso': 'ChartDataLabel',
                     'params': {'position': 'outside_end'}},  # xlLabelPositionOutsideEnd=2
                ],
                'interaction_class': 'insert_then_data_labels',
                'verification_description':
                    'Chart existence + has_data_labels=True, position=outside_end '
                    'via _verify_chart_composition',
                'instruction':
                    "Below the 'Quarterly Sales:' heading, insert a clustered "
                    "column chart with categories 'Q1' and 'Q2', "
                    "and one series 'Revenue' with values 100 and 150, "
                    "then add data labels at the 'Outside End' position.",
            },
            'variant_gridlines': {
                'steps': [
                    {'operation': 'insert_chart',
                     'params': {'chart_type': 4,
                                'categories': ['Jan', 'Feb', 'Mar'],
                                'series': [{'name': 'Temperature',
                                            'values': [10, 15, 20]}]}},
                    {'idMso': 'ChartGridlinesGallery',
                     'params': {'axis': 'value', 'gridline': 'minor'}},
                ],
                'interaction_class': 'insert_then_gridlines',
                'verification_description':
                    'Chart existence + value axis has_minor_gridlines=True via '
                    '_verify_chart_composition',
                'instruction':
                    "Below the 'Quarterly Sales:' heading, insert a line chart "
                    "with categories 'Jan', 'Feb', and 'Mar', and one series "
                    "'Temperature' with values 10, 15, and 20, then turn on "
                    "the primary horizontal minor gridlines of the chart.",
            },
        },
    },
    {
        'group': 'PrerequisiteIcon',
        'target_object_type': 'picture',
        'behavior': 'object_effect',
        'verification_method': 'composition_dispatch',
        'composition_type': 'prerequisite',
        'variants': {
            'variant_shadow': {
                'steps': [
                    {'operation': 'insert_icon',
                     'params': {'shape': 'star'}},
                    {'idMso': 'PictureEffectsShadowGallery',
                     'params': {'shadow_type': 25}},  # outer center
                ],
                'interaction_class': 'insert_then_shadow',
                'verification_description':
                    'Icon existence + shadow_type=25 via _verify_picture_effects',
                'instruction':
                    "Below the 'Company Logo:' heading, insert a 150 x 150 pt "
                    "blue star-shaped SVG icon (RGB: 68, 114, 196), then add "
                    "an outer center shadow effect.",
            },
            'variant_reflection': {
                'steps': [
                    {'operation': 'insert_icon',
                     'params': {'shape': 'circle'}},
                    {'idMso': 'PictureReflectionGallery',
                     'params': {'reflection_type': 5}},  # Half Reflection, 4 pt offset
                ],
                'interaction_class': 'insert_then_reflection',
                'verification_description':
                    'Icon existence + reflection_type via '
                    '_verify_picture_effects',
                'instruction':
                    "Below the 'Company Logo:' heading, insert a 150 x 150 pt "
                    "orange circle-shaped SVG icon (RGB: 237, 125, 49), then "
                    "apply the 'Half Reflection with a 4 pt offset' effect to it.",
            },
            'variant_glow': {
                'steps': [
                    {'operation': 'insert_icon',
                     'params': {'shape': 'star'}},
                    {'idMso': 'PictureEffectsGlowGallery',
                     'params': {'glow_radius': 11.0,
                                'theme_color': 'accent4'}},
                ],
                'interaction_class': 'insert_then_glow',
                'verification_description':
                    'Icon existence + glow_radius, glow_color_rgb via '
                    '_verify_picture_effects',
                'instruction':
                    "Below the 'Company Logo:' heading, insert a 150 x 150 pt "
                    "blue star-shaped SVG icon (RGB: 68, 114, 196), then "
                    "apply an 11 pt glow effect with the Accent 4 theme color.",
            },
        },
    },
    {
        'group': 'PrerequisiteShape',
        'target_object_type': 'shape',
        'behavior': 'object_effect',
        'verification_method': 'composition_dispatch',
        'composition_type': 'prerequisite',
        'variants': {
            'variant_fill': {
                'steps': [
                    {'operation': 'insert_shape',
                     'params': {'type': 1, 'left': 100, 'top': 100,
                                'width': 100, 'height': 80}},
                    {'idMso': 'ShapeFillColorPicker',
                     'params': {'color_bgr': 255}},  # RGB(255, 0, 0)
                ],
                'interaction_class': 'insert_then_fill',
                'verification_description':
                    'Shape existence + fill_fore_color_rgb via _verify_shape_effects',
                'instruction':
                    "Below the 'Diagram:' heading, insert a rectangle shape (100 by 80 points) with its top-left corner 100 points from the left edge of the page and 100 points from the top edge of the page, then set its fill color to Red (RGB: 255, 0, 0).",
            },
            'variant_shadow': {
                'steps': [
                    {'operation': 'insert_shape',
                     'params': {'type': 1, 'left': 100, 'top': 100,
                                'width': 100, 'height': 80}},
                    {'idMso': 'ObjectEffectShadowGallery',
                     'params': {'shadow_type': 25}},  # outer center
                ],
                'interaction_class': 'insert_then_shadow',
                'verification_description':
                    'Shape existence + shadow_type via _verify_shape_effects',
                'instruction':
                    "Below the 'Diagram:' heading, insert a rectangle shape (100 by 80 points) with its top-left corner 100 points from the left edge of the page and 100 points from the top edge of the page, then add an outer center shadow effect.",
            },
            'variant_bevel': {
                'steps': [
                    {'operation': 'insert_shape',
                     'params': {'type': 1, 'left': 100, 'top': 100,
                                'width': 100, 'height': 80}},
                    {'idMso': 'BevelShapeGallery',
                     'params': {'bevel_type': 6, 'bevel_depth': 6,
                                'bevel_inset': 6}},  # Angle
                ],
                'interaction_class': 'insert_then_bevel',
                'verification_description':
                    'Shape existence + threed_bevel_top_type via '
                    '_verify_shape_geometry',
                'instruction':
                    "Below the 'Diagram:' heading, insert a rectangle shape (100 by 80 points) with its top-left corner 100 points from the left edge of the page and 100 points from the top edge of the page, then apply the 'Angle' bevel effect.",
            },
            'variant_glow': {
                'steps': [
                    {'operation': 'insert_shape',
                     'params': {'type': 1, 'left': 100, 'top': 100,
                                'width': 100, 'height': 80}},
                    {'idMso': 'ObjectEffectGlowGallery',
                     'params': {'glow_radius': 11.0,
                                'glow_color_rgb': 49407}},
                ],
                'interaction_class': 'insert_then_glow',
                'verification_description':
                    'Shape existence + glow_radius, glow_color_rgb via '
                    '_verify_shape_effects',
                'instruction':
                    "Below the 'Diagram:' heading, insert a rectangle shape (100 by 80 points) with its top-left corner 100 points from the left edge of the page and 100 points from the top edge of the page, then apply an 11 pt glow effect with the Accent 4 theme color to it.",
            },
        },
    },
    {
        'group': 'PrerequisiteSmartArt',
        'target_object_type': 'smartart',
        'behavior': 'object_effect',
        'verification_method': 'composition_dispatch',
        'composition_type': 'prerequisite',
        'variants': {
            'variant_style': {
                'steps': [
                    {'operation': 'insert_smartart',
                     'params': {'layout_index': 1,
                                'labels': ['Alpha', 'Beta', 'Charlie',
                                           'Delta', 'Echo']}},
                    {'idMso': 'SmartArtStylesGallery',
                     'params': {'style_index': 5}},  # Intense Effect
                ],
                'interaction_class': 'insert_then_style',
                'verification_description':
                    'SmartArt existence + style properties via _verify_smartart',
                'instruction':
                    "Below the 'Team Structure:' heading, insert a Basic Block List SmartArt graphic with nodes 'Alpha', 'Beta', 'Charlie', 'Delta', and 'Echo', then apply the 'Intense Effect' SmartArt style.",
            },
            'variant_color': {
                'steps': [
                    {'operation': 'insert_smartart',
                     'params': {'layout_index': 1,
                                'labels': ['Alpha', 'Beta', 'Charlie',
                                           'Delta', 'Echo']}},
                    {'idMso': 'SmartArtChangeColorsGallery',
                     'params': {'color_index': 5}},  # Colorful Range 2-3
                ],
                'interaction_class': 'insert_then_color',
                'verification_description':
                    'SmartArt existence + color properties via _verify_smartart',
                'instruction':
                    "Below the 'Team Structure:' heading, insert a Basic Block List SmartArt graphic with nodes 'Alpha', 'Beta', 'Charlie', 'Delta', and 'Echo', then change its colors to Colorful Range (Accent 2-3).",
            },
            'variant_layout': {
                'steps': [
                    {'operation': 'insert_smartart',
                     'params': {'layout_index': 102,
                                'layout_name': SMARTART_INITIAL_LAYOUT,
                                'labels': ['Boss', 'Team1', 'Member1',
                                           'Member2', 'Team2', 'Member3']}},
                    {'idMso': 'SmartArtLayoutGallery',
                     'params': {'layout_index': 20,
                                'layout_name': SMARTART_FINAL_LAYOUT}},
                ],
                'interaction_class': 'insert_then_layout_change',
                'verification_description':
                    'SmartArt existence + layout change via '
                    '_verify_smartart_structural',
                'instruction':
                    "Below the 'Team Structure:' heading, insert a Hierarchy SmartArt graphic with 'Boss' at the top, 'Team1' and 'Team2' below, 'Member1' and 'Member2' under 'Team1', and 'Member3' under 'Team2', then change its layout to 'Continuous Picture List'.",
            },
        },
    },
    {
        'group': 'PrerequisiteEquation',
        'target_object_type': 'equation',
        'behavior': 'object_effect',
        'verification_method': 'composition_dispatch',
        'composition_type': 'prerequisite',
        'variants': {
            'variant_professional': {
                'steps': [
                    {'operation': 'insert_equation',
                     'params': {'content': 'x^2+y=z'}},
                    {'idMso': 'EquationProfessionalAll', 'params': {}},
                ],
                'interaction_class': 'insert_then_format',
                'verification_description':
                    'Equation existence + BuildUp=True via _verify_equation',
                'instruction':
                    "Below the 'Formula:' heading, insert a new equation containing 'x^2+y=z', then convert it to Professional (display) format.",
            },
            'variant_fraction': {
                'steps': [
                    {'operation': 'insert_equation',
                     'params': {'content': 'y='}},
                    {'idMso': 'EquationFractionGallery',
                     'params': {'frac_type': 0}},  # Stacked fraction
                ],
                'interaction_class': 'insert_then_structure',
                'verification_description':
                    'Equation existence + function_types containing fraction via '
                    '_verify_equation',
                'instruction':
                    "Below the 'Formula:' heading, insert a new equation containing 'y=', then insert a stacked fraction 'a' over 'b' after it.",
            },
            'variant_radical': {
                'steps': [
                    {'operation': 'insert_equation',
                     'params': {'content': 'x='}},
                    {'idMso': 'EquationRadicalGallery',
                     'params': {}},
                ],
                'interaction_class': 'insert_then_structure',
                'verification_description':
                    'Equation existence + function_types containing radical via '
                    '_verify_equation',
                'instruction':
                    "Below the 'Formula:' heading, insert a new equation containing 'x=', then add a square root of 'a' after the '=' sign.",
            },
            'variant_script': {
                'steps': [
                    {'operation': 'insert_equation',
                     'params': {'content': 'y='}},
                    {'idMso': 'EquationScriptGallery',
                     'params': {}},
                ],
                'interaction_class': 'insert_then_structure',
                'verification_description':
                    'Equation existence + function_types containing superscript '
                    'via _verify_equation',
                'instruction':
                    "Below the 'Formula:' heading, insert a new equation containing 'y=', then add 'x' with superscript '2' (x²) after the '=' sign.",
            },
        },
    },
    {
        'group': 'PrerequisiteTextBox',
        'target_object_type': 'textbox',
        'behavior': 'object_effect',
        'verification_method': 'composition_dispatch',
        'composition_type': 'prerequisite',
        'variants': {
            'variant_direction': {
                'steps': [
                    {'operation': 'insert_textbox',
                     'params': {'text': 'Sample Text', 'left': 100, 'top': 100,
                                'width': 200, 'height': 80}},
                    {'idMso': 'TextDirection',
                     'params': {'orientation': 3}},  # msoTextOrientationDownward
                ],
                'interaction_class': 'insert_then_direction',
                'verification_description':
                    'TextBox existence + text_frame_orientation=downward '
                    '(rotated 90° clockwise) via _verify_text_alignment',
                'instruction':
                    "Below the 'Note:' heading, draw a text box by dragging from (100, 100) to (300, 180), type 'Sample Text' in it, then rotate all text in the text box 90° clockwise (so the text reads from top to bottom with characters facing right).",
            },
            'variant_outline_color': {
                'steps': [
                    {'operation': 'insert_textbox',
                     'params': {'text': 'Sample Text', 'left': 100, 'top': 100,
                                'width': 200, 'height': 80}},
                    {'idMso': 'ShapeOutlineColorPicker',
                     'params': {'color_bgr': 12611584}},  # Blue (RGB 0,112,192)
                ],
                'interaction_class': 'insert_then_outline_color',
                'verification_description':
                    'TextBox existence + line_fore_color_rgb via '
                    '_verify_shape_effects',
                'instruction':
                    "Below the 'Note:' heading, draw a text box by dragging from (100, 100) to (300, 180), type 'Sample Text' in it, then set the text box's outline color to Blue (RGB: 0, 112, 192).",
            },
        },
    },
]


# ═════════════════════════════════════════════════════════════════════════
# Generator class
# ═════════════════════════════════════════════════════════════════════════

class PrerequisiteGenerator:
    """Generator for prerequisite chain test cases."""

    def __init__(self, output_dir: Path = None, hidden: bool = False):
        self.output_dir = output_dir or (BENCHMARK_OUTPUT_DIR / TEST_SET)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.inputs_dir = self.output_dir / 'inputs'
        self.inputs_dir.mkdir(parents=True, exist_ok=True)
        self.hidden = hidden

    def generate_all(self, command: str = None, variant_names=None):
        """Generate all prerequisite test cases."""
        requested_variants = set(variant_names or [])
        generated_variants = set()
        wd = ensure_word(visible=not self.hidden)
        results = {}

        for group_def in CASE_GROUPS:
            group_name = group_def['group']
            if command and command != group_name:
                continue
            for variant_name, variant_def in group_def['variants'].items():
                if requested_variants and variant_name not in requested_variants:
                    continue
                case_key = f'{group_name}/{variant_name}'
                # Merge group-level and variant-level fields
                case_def = self._merge_case_def(group_def, variant_def, variant_name)
                generated_variants.add(variant_name)
                try:
                    ok = self._generate_case(wd, group_name, variant_name, case_def)
                    results[case_key] = ok
                    status = 'OK' if ok else 'FAILED'
                    print(f"  [{case_key}] {status}")
                except Exception as e:
                    results[case_key] = False
                    print(f"  [{case_key}] ERROR: {e}")

        unknown_variants = requested_variants - generated_variants
        if unknown_variants:
            raise ValueError(
                'Requested variant(s) not found under the selected command: '
                + ', '.join(sorted(unknown_variants))
            )

        # Write instructions JSON (source of truth is in case defs)
        commands = {}
        for group_def in CASE_GROUPS:
            group_name = group_def['group']
            variants = {}
            for variant_name, variant_def in group_def['variants'].items():
                variants[variant_name] = variant_def['instruction']
            commands[group_name] = variants
        write_instructions_json(TEST_SET, commands)

        # Write generation summary
        total = len(results)
        passed = sum(1 for v in results.values() if v)
        summary = {
            'test_set': TEST_SET,
            'generated_at': datetime.now().isoformat(),
            'total': total,
            'passed': passed,
            'results': {k: 'success' if v else 'failed'
                        for k, v in results.items()},
        }
        if command is None and not requested_variants:
            summary_path = self.output_dir / 'generation_summary.json'
            with open(summary_path, 'w', encoding='utf-8') as f:
                json.dump(summary, f, indent=2, ensure_ascii=False)

        print(f"\nSummary: {passed}/{total} cases generated successfully")
        return results

    # ─────────────────────────────────────────────────────────────────
    # Merge helper
    # ─────────────────────────────────────────────────────────────────

    def _merge_case_def(self, group_def, variant_def, variant_name):
        """Merge group-level defaults with variant-level overrides."""
        case_def = {
            'group': group_def['group'],
            'variant': variant_name,
            'composition_type': group_def.get('composition_type', 'prerequisite'),
            'target_object_type': variant_def.get('target_object_type') or group_def.get('target_object_type'),
            'behavior': variant_def.get('behavior') or group_def.get('behavior'),
            'verification_method': variant_def.get('verification_method') or group_def.get('verification_method'),
            'steps': variant_def['steps'],
            'interaction_class': variant_def['interaction_class'],
            'verification_description': variant_def['verification_description'],
            'instruction': variant_def['instruction'],
        }
        # Optional fields from variant or group
        for key in ('auxiliary', 'doc_in_name'):
            val = variant_def.get(key) or group_def.get(key)
            if val:
                case_def[key] = val
        return case_def

    # ─────────────────────────────────────────────────────────────────
    # Per-case generation
    # ─────────────────────────────────────────────────────────────────

    def _generate_case(self, wd, group_name, variant_name, case_def):
        """Generate a single test case variant."""
        variant_dir = self.output_dir / group_name / variant_name
        variant_dir.mkdir(parents=True, exist_ok=True)
        gold_path = variant_dir / 'gold.docx'
        spec_path = variant_dir / 'spec.json'

        # Determine doc_in
        doc_in_ref, doc_in_path = self._resolve_doc_in(wd, case_def,
                                                        variant_dir)

        # Generate gold via COM
        self._generate_gold(wd, case_def, doc_in_path, gold_path)

        # Ensure img.png is in inputs/ for picture cases
        if case_def.get('auxiliary'):
            self._ensure_auxiliary(case_def)

        # Build and write spec
        atomic_intents = build_prerequisite_chain_atomic_intents(
            case_def['steps'])
        spec = build_spec(
            behavior=case_def['behavior'],
            composition_type=case_def['composition_type'],
            interaction_class=case_def['interaction_class'],
            target_object_type=case_def['target_object_type'],
            steps=case_def['steps'],
            precondition=self._build_precondition(case_def),
            verification_method=case_def['verification_method'],
            verification_description=case_def['verification_description'],
            doc_in_ref=doc_in_ref,
            variant_name=variant_name,
            auxiliary=case_def.get('auxiliary'),
            atomic_intents=atomic_intents,
        )
        write_spec(spec_path, spec)
        return True

    # ─────────────────────────────────────────────────────────────────
    # Doc_in resolution
    # ─────────────────────────────────────────────────────────────────

    def _resolve_doc_in(self, wd, case_def, variant_dir):
        """Return (relative_ref, absolute_path) for the doc_in.

        Default behavior: each prerequisite group gets a contextual stub
        document (a Heading 2 line naming the purpose of the insertion)
        so the agent has a document-level anchor rather than an empty
        canvas.  See STUB_HEADINGS.
        """
        # Variants that reference a shared input (explicit override)
        doc_in_name = case_def.get('doc_in_name')
        if doc_in_name:
            path = get_shared_input(self.inputs_dir, doc_in_name)
            return f'../../inputs/{doc_in_name}.docx', path

        # Default: per-group contextual stub
        group = case_def['group']
        heading = STUB_HEADINGS.get(group)
        if heading:
            stub_name = f'doc_in_{group.lower()}_stub.docx'
            stub_path = self.inputs_dir / stub_name
            if not stub_path.exists():
                create_stub_doc(wd, stub_path, heading)
                print(f'  Created: {stub_name}')
            return f'../../inputs/{stub_name}', stub_path

        # Fallback: empty document (should not hit for prerequisite cases)
        empty_path = self.inputs_dir / 'doc_in_empty.docx'
        if not empty_path.exists():
            create_empty_doc(wd, empty_path)
            print('  Created: doc_in_empty.docx')
        return '../../inputs/doc_in_empty.docx', empty_path

    # ─────────────────────────────────────────────────────────────────
    # Gold generation — dispatches by object type
    # ─────────────────────────────────────────────────────────────────

    def _generate_gold(self, wd, case_def, doc_in_path, gold_path):
        """Open doc_in, execute all steps via COM, save as gold."""
        target = case_def['target_object_type']

        if target == 'table':
            self._gold_table_prerequisite(wd, case_def, doc_in_path,
                                          gold_path)
        elif target == 'chart':
            self._gold_chart_prerequisite(wd, case_def, doc_in_path,
                                          gold_path)
        elif target == 'picture':
            self._gold_picture_prerequisite(wd, case_def, doc_in_path,
                                            gold_path)
        elif target == 'shape':
            self._gold_shape_prerequisite(wd, case_def, doc_in_path,
                                          gold_path)
        elif target == 'smartart':
            self._gold_smartart_prerequisite(wd, case_def, doc_in_path,
                                             gold_path)
        elif target == 'equation':
            self._gold_equation_prerequisite(wd, case_def, doc_in_path,
                                             gold_path)
        elif target == 'textbox':
            self._gold_textbox_prerequisite(wd, case_def, doc_in_path,
                                            gold_path)
        else:
            raise ValueError(f'Unknown target_object_type: {target}')

    # ── Table prerequisite ───────────────────────────────────────────

    def _gold_table_prerequisite(self, wd, case_def, doc_in_path, gold_path):
        doc = open_doc(wd, doc_in_path)
        try:
            # Step A: Insert table
            params_a = case_def['steps'][0]['params']
            table = create_table(doc, params_a['rows'], params_a['cols'])

            # Step B: Apply operation
            step_b = case_def['steps'][1]
            idMso_b = step_b['idMso']
            params_b = step_b['params']

            if idMso_b == 'TableStylesGalleryWord':
                table.Style = params_b['style_name']
                time.sleep(0.3)

            elif idMso_b == 'TableCellAlignMiddleCenter':
                for row in table.Rows:
                    for cell in row.Cells:
                        try:
                            cell.VerticalAlignment = 1  # wdCellAlignVerticalCenter
                            cell.Range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
                        except Exception:
                            pass
                time.sleep(0.2)

            elif idMso_b == 'BorderStylesGallery':
                # Apply double-line borders to the outer border only.
                for border_id in [-4, -3, -2, -1]:
                    # -4=Right, -3=Bottom, -2=Left, -1=Top
                    try:
                        table.Borders(border_id).LineStyle = 7
                    except Exception:
                        pass
                time.sleep(0.2)

            elif idMso_b == 'BorderColorPicker':
                # Apply color only to outer borders (top/bottom/left/right)
                color = params_b['color_bgr']
                for border_id in [-4, -3, -2, -1]:
                    try:
                        border = table.Borders(border_id)
                        border.LineStyle = 1  # wdLineStyleSingle (ensure visible)
                        border.Color = color
                    except Exception:
                        pass
                time.sleep(0.2)

            elif idMso_b == 'AlignCenter':
                # Center-align the text in cell R2C2 (middle cell)
                try:
                    cell = table.Cell(2, 2)
                    cell.Range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
                except Exception:
                    pass
                time.sleep(0.2)

            save_as_gold(doc, gold_path)
        finally:
            close_doc(doc)

    # ── Chart prerequisite ───────────────────────────────────────────

    def _gold_chart_prerequisite(self, wd, case_def, doc_in_path, gold_path):
        doc = open_doc(wd, doc_in_path)
        try:
            # Step A: Insert chart with explicit data
            params_a = case_def['steps'][0]['params']
            chart = create_chart(
                doc, params_a['chart_type'],
                categories=params_a.get('categories'),
                series=params_a.get('series'),
            )

            # Step B: Apply operation
            step_b = case_def['steps'][1]
            idMso_b = step_b['idMso']
            params_b = step_b['params']

            if idMso_b == 'ChartTitle':
                chart.HasTitle = True
                chart.ChartTitle.Text = 'Chart Title'
                chart.ChartTitle.IncludeInLayout = True  # above chart
                time.sleep(0.3)

            elif idMso_b == 'ChartLegend':
                chart.HasLegend = True
                chart.Legend.Position = -4152  # xlLegendPositionRight
                time.sleep(0.3)

            elif idMso_b == 'ChartTrendline':
                series = chart.SeriesCollection(1)
                series.Trendlines().Add(Type=-4132)  # xlLinear
                time.sleep(0.3)

            elif idMso_b == 'ChartDataLabel':
                # Add data labels to all series at the 'Outside End' position
                # xlLabelPositionOutsideEnd = 2
                for i in range(1, chart.SeriesCollection().Count + 1):
                    try:
                        series = chart.SeriesCollection(i)
                        series.HasDataLabels = True
                        series.DataLabels().Position = 2
                    except Exception:
                        pass
                time.sleep(0.3)

            elif idMso_b == 'ChartGridlinesGallery':
                # Turn on minor gridlines on the value axis (primary vertical)
                # xlValue = 2, xlPrimary = 1
                try:
                    value_axis = chart.Axes(2, 1)
                    value_axis.HasMinorGridlines = True
                except Exception:
                    pass
                time.sleep(0.3)

            save_as_gold(doc, gold_path)
        finally:
            close_doc(doc)

    # ── Picture prerequisite ─────────────────────────────────────────

    def _gold_picture_prerequisite(self, wd, case_def, doc_in_path,
                                   gold_path):
        doc = open_doc(wd, doc_in_path)
        try:
            # Step A: Insert picture or icon
            step_a = case_def['steps'][0]
            if step_a['operation'] == 'insert_icon':
                pic = create_icon(doc, shape=step_a.get('params', {}).get('shape', 'star'))
            else:
                pic = create_picture(doc)

            # Step B: Apply operation
            step_b = case_def['steps'][1]
            idMso_b = step_b['idMso']
            params_b = step_b['params']

            if idMso_b == 'PictureCropAspectRatio1To1':
                # Crop to 1:1 square
                pic_width = pic.Width
                pic_height = pic.Height
                current_ratio = pic_width / pic_height
                target_ratio = 1.0
                if current_ratio > target_ratio:
                    new_width = pic_height * target_ratio
                    crop_amount = (pic_width - new_width) / 2
                    pic.PictureFormat.CropLeft = crop_amount
                    pic.PictureFormat.CropRight = crop_amount
                else:
                    new_height = pic_width / target_ratio
                    crop_amount = (pic_height - new_height) / 2
                    pic.PictureFormat.CropTop = crop_amount
                    pic.PictureFormat.CropBottom = crop_amount
                time.sleep(0.3)

            elif idMso_b == 'PictureRecolorGallery':
                pic.PictureFormat.ColorType = params_b['color_type']
                time.sleep(0.3)

            elif idMso_b == 'PictureEffectsShadowGallery':
                pic.Shadow.Visible = True
                pic.Shadow.Type = params_b['shadow_type']
                time.sleep(0.3)

            elif idMso_b == 'PictureReflectionGallery':
                # Reflection requires ConvertToShape for the effect API
                shape = pic.ConvertToShape()
                try:
                    shape.Reflection.Type = params_b['reflection_type']
                finally:
                    shape.ConvertToInlineShape()
                time.sleep(0.3)

            elif idMso_b == 'PictureEffectsGlowGallery':
                pic.Glow.Radius = params_b['glow_radius']
                pic.Glow.Color.ObjectThemeColor = 8  # Accent 4
                pic.Glow.Transparency = 0.0
                time.sleep(0.3)

            save_as_gold(doc, gold_path)
        finally:
            close_doc(doc)

        if idMso_b == 'PictureEffectsGlowGallery':
            from generator.test_sets.object_effect_contextual_parameterized import (
                _fix_glow_scheme_color,
            )
            _fix_glow_scheme_color(gold_path, 0xFFC000, 'accent4')

    # ── Shape prerequisite ───────────────────────────────────────────

    def _gold_shape_prerequisite(self, wd, case_def, doc_in_path, gold_path):
        doc = open_doc(wd, doc_in_path)
        try:
            # Step A: Insert shape
            params_a = case_def['steps'][0]['params']
            shape = create_shape(doc, params_a['type'], params_a['left'],
                                 params_a['top'], params_a['width'],
                                 params_a['height'])

            # Step B: Apply operation
            step_b = case_def['steps'][1]
            idMso_b = step_b['idMso']
            params_b = step_b['params']

            if idMso_b == 'ShapeFillColorPicker':
                shape.Fill.Solid()
                shape.Fill.ForeColor.RGB = params_b['color_bgr']
                time.sleep(0.2)

            elif idMso_b == 'ObjectEffectShadowGallery':
                shape.Shadow.Type = params_b['shadow_type']
                time.sleep(0.2)

            elif idMso_b == 'BevelShapeGallery':
                shape.ThreeD.BevelTopType = params_b['bevel_type']
                shape.ThreeD.BevelTopDepth = params_b.get('bevel_depth', 6)
                shape.ThreeD.BevelTopInset = params_b.get('bevel_inset', 6)
                time.sleep(0.2)

            elif idMso_b == 'ObjectEffectGlowGallery':
                shape.Glow.Radius = params_b['glow_radius']
                shape.Glow.Color.RGB = params_b['glow_color_rgb']
                time.sleep(0.2)

            elif idMso_b == 'OutlineDashesGallery':
                shape.Line.DashStyle = params_b['dash_style']
                time.sleep(0.2)

            save_as_gold(doc, gold_path)
        finally:
            close_doc(doc)

        if idMso_b == 'ObjectEffectGlowGallery':
            from generator.test_sets.object_effect_contextual_parameterized import (
                _fix_glow_scheme_color,
            )
            _fix_glow_scheme_color(gold_path, 0xFFC000, 'accent4')

    # ── SmartArt prerequisite ────────────────────────────────────────

    def _gold_smartart_prerequisite(self, wd, case_def, doc_in_path,
                                    gold_path):
        doc = open_doc(wd, doc_in_path)
        try:
            # Step A: Insert SmartArt
            params_a = case_def['steps'][0]['params']
            sa_shape = create_smartart(wd, doc, params_a['layout_index'],
                                       params_a.get('labels'),
                                       layout_name=params_a.get('layout_name'))
            sa = sa_shape.SmartArt

            # Step B: Apply operation
            step_b = case_def['steps'][1]
            idMso_b = step_b['idMso']
            params_b = step_b['params']

            if idMso_b == 'SmartArtStylesGallery':
                styles = wd.Application.SmartArtQuickStyles
                sa.QuickStyle = styles(params_b['style_index'])
                time.sleep(0.5)

            elif idMso_b == 'SmartArtChangeColorsGallery':
                colors = wd.Application.SmartArtColors
                sa.Color = colors(params_b['color_index'])
                time.sleep(0.5)

            elif idMso_b == 'SmartArtLayoutGallery':
                sa.Layout = get_smartart_layout(
                    wd,
                    layout_index=params_b.get('layout_index'),
                    layout_name=params_b.get('layout_name'),
                )
                time.sleep(1.0)

            save_as_gold(doc, gold_path)
        finally:
            close_doc(doc)

        if idMso_b == 'SmartArtLayoutGallery':
            self._assert_smartart_layout_postconditions(wd, gold_path)

    @staticmethod
    def _smartart_layout_snapshot(doc):
        shape = next(
            (
                doc.Shapes(index)
                for index in range(1, doc.Shapes.Count + 1)
                if doc.Shapes(index).HasSmartArt
            ),
            None,
        )
        if shape is None:
            return {'found': False}

        heading = doc.Content.Duplicate
        heading.Find.ClearFormatting()
        heading.Find.Text = 'Team Structure:'
        heading_found = bool(heading.Find.Execute())
        below_heading = False
        if heading_found:
            heading_top = float(heading.Information(6))
            relative_vertical = int(shape.RelativeVerticalPosition)
            raw_top = float(shape.Top)
            anchor_top = float(shape.Anchor.Information(6))
            if relative_vertical == 1:
                object_top = raw_top
            elif relative_vertical == 0:
                object_top = float(doc.PageSetup.TopMargin) + raw_top
            else:
                object_top = anchor_top + raw_top
            below_heading = object_top > heading_top + 1.0

        smartart = shape.SmartArt
        nodes = smartart.AllNodes
        return {
            'found': True,
            'layout_name': str(smartart.Layout.Name or '').strip(),
            'all_node_count': int(nodes.Count),
            'all_node_texts': [
                str(nodes(index).TextFrame2.TextRange.Text or '').strip()
                for index in range(1, nodes.Count + 1)
            ],
            'all_node_levels': [
                int(nodes(index).Level)
                for index in range(1, nodes.Count + 1)
            ],
            'below_heading': below_heading,
        }

    @classmethod
    def _assert_smartart_layout_postconditions(cls, wd, gold_path):
        """Reopen the saved reference and reject semantic/package regressions."""
        doc = open_doc(wd, gold_path)
        try:
            snapshot = cls._smartart_layout_snapshot(doc)
        finally:
            close_doc(doc)
        failures = _validate_smartart_layout_snapshot(snapshot)
        package_snapshot = _smartart_package_snapshot(gold_path)
        failures.extend(_validate_smartart_package_snapshot(package_snapshot))
        if failures:
            raise ValueError('IMP-17 SmartArt postcondition failed: ' + '; '.join(failures))

    # ── Equation prerequisite ────────────────────────────────────────

    def _gold_equation_prerequisite(self, wd, case_def, doc_in_path,
                                    gold_path):
        doc = open_doc(wd, doc_in_path)
        try:
            # Step A: Insert equation and type content
            params_a = case_def['steps'][0]['params']
            create_equation(wd, doc, params_a['content'])

            # Step B: Apply operation
            step_b = case_def['steps'][1]
            idMso_b = step_b['idMso']
            params_b = step_b['params']

            if idMso_b == 'EquationProfessionalAll':
                # Build up the equation to professional format
                omath = doc.OMaths(1)
                omath.BuildUp()
                time.sleep(0.5)

            elif idMso_b == 'EquationFractionGallery':
                # Navigate to end of equation, insert fraction
                omath = doc.OMaths(1)
                rng = omath.Range
                rng.Collapse(0)  # wdCollapseEnd
                func = omath.Functions.Add(rng, 7)  # omMathFuncFrac
                func.Frac.Type = params_b.get('frac_type', 0)
                func.Frac.Num.Range.Text = 'a'
                func.Frac.Den.Range.Text = 'b'
                time.sleep(0.5)

            elif idMso_b == 'EquationRadicalGallery':
                # Insert a square root with 'a' inside, after the '=' sign
                omath = doc.OMaths(1)
                rng = omath.Range
                rng.Collapse(0)  # wdCollapseEnd
                func = omath.Functions.Add(rng, 16)  # wdOMathFunctionRad
                func.Rad.HideDeg = True  # hide degree → square root
                func.Rad.E.Range.Text = 'a'
                time.sleep(0.5)

            elif idMso_b == 'EquationScriptGallery':
                # Insert a superscript with base 'x' and exponent '2'
                omath = doc.OMaths(1)
                rng = omath.Range
                rng.Collapse(0)  # wdCollapseEnd
                func = omath.Functions.Add(rng, 19)  # omMathFuncScrSup
                func.ScrSup.E.Range.Text = 'x'
                func.ScrSup.Sup.Range.Text = '2'
                time.sleep(0.5)

            save_as_gold(doc, gold_path)
        finally:
            close_doc(doc)

    # ── TextBox prerequisite ─────────────────────────────────────────

    def _gold_textbox_prerequisite(self, wd, case_def, doc_in_path,
                                   gold_path):
        doc = open_doc(wd, doc_in_path)
        try:
            # Step A: Insert text box
            params_a = case_def['steps'][0]['params']
            tb = create_textbox(
                doc,
                params_a.get('text', 'Sample Text'),
                params_a.get('left', 100),
                params_a.get('top', 100),
                params_a.get('width', 200),
                params_a.get('height', 80),
            )

            # Step B: Apply operation
            step_b = case_def['steps'][1]
            idMso_b = step_b['idMso']
            params_b = step_b['params']

            if idMso_b == 'TextDirection':
                tb.TextFrame.Orientation = params_b['orientation']
                time.sleep(0.2)

            elif idMso_b == 'TextFillColorPicker':
                # Set the font color of the text inside the text box
                tb.TextFrame.TextRange.Font.Color = params_b['color_bgr']
                time.sleep(0.2)

            elif idMso_b == 'ShapeOutlineColorPicker':
                tb.Line.ForeColor.RGB = params_b['color_bgr']
                time.sleep(0.2)

            save_as_gold(doc, gold_path)
        finally:
            close_doc(doc)

    # ─────────────────────────────────────────────────────────────────
    # Helpers
    # ─────────────────────────────────────────────────────────────────

    def _build_precondition(self, case_def):
        """Build the precondition dict for the spec."""
        group = case_def.get('group', '')
        heading = STUB_HEADINGS.get(group)
        if heading:
            return {
                'type': 'stub',
                'heading': heading,
                'description': (
                    f"Document with a single Heading 2 paragraph '{heading}' "
                    'providing contextual anchor for the insertion'
                ),
            }
        return {
            'type': 'empty',
            'description': 'Empty document -- agent must create the object',
        }

    def _ensure_auxiliary(self, case_def):
        """Copy auxiliary assets (e.g., img.png) into inputs/."""
        img_dest = self.inputs_dir / 'img.png'
        if not img_dest.exists():
            import shutil
            shutil.copy2(IMG_PATH, img_dest)
            print('  Copied auxiliary: img.png')
