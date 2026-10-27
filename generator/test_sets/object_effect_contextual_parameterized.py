"""
Test case generator for object_effect × contextual_parameterized.

Generates test cases for Word commands that apply effects to contextual objects
(tables, pictures, shapes, charts, SmartArt, equations, etc.) with parameter selection:
- Gallery controls (styles, colors, effects)
- Parameterized menus

Parameters are loaded from variant definition file:
- object_effect_x_contextual_parameterized_variants.json

Object types supported:
- Table: TableStylesGalleryWord, BorderStylesGallery, etc.
- Picture: PictureStylesGallery, PictureEffects*, PictureRecolor*, etc.
- Shape: ShapeStylesGallery, ShapeFillColorPicker, OutlineColorPicker, etc.
- Chart: ChartStylesGallery, ChartLayoutGallery, ChartColorsGallery, etc.
- SmartArt: SmartArtStylesGallery, SmartArtLayoutGallery, etc.
- Equation: EquationIntegralGallery, EquationFractionGallery, etc.
- Text Box: TextAlignGallery, TextDirectionGallery, etc.
"""

import json
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional
from win32com.client import constants as c

from ..config import (
    COMMAND_SETS_DIR, BENCHMARK_OUTPUT_DIR,
)
from ..execution.setup import ensure_word, finalize_word
from ..execution.execution import DirectExecute
from generator.test_sets.spec_utils import write_spec_if_changed
from generator.docx_package import write_docx_package

# Path to variant files
CONTEXTUAL_PARAMETERIZED_VARIANTS_PATH = Path(__file__).parent.parent.parent / 'instructions' / 'object_effect_x_contextual_parameterized_variants.json'
EQUATION_GALLERY_VARIANTS_PATH = Path(__file__).parent.parent.parent / 'instructions' / 'equation_gallery_variants.json'
# Path to hand-written instruction overrides (deterministic, high-quality)
INSTRUCTIONS_OVERRIDE_PATH = Path(__file__).parent.parent.parent / 'instructions' / 'object_effect_x_contextual_parameterized.json'

# =============================================================================
# PRIORITIZED COMMANDS (PM List Scope)
# =============================================================================
# These are the only commands we're focusing on for this test set.
# All other commands are deprioritized as they are outside the PM list scope.

PRIORITIZED_COMMANDS = {
    # Table styling
    'TableStylesGalleryWord',       # Apply a table style
    # Table borders
    'OutlineWeightGallery',         # Border weight (also used for shapes, but PM focus is table)
    'BorderStylesGallery',          # Border styles
    # Text box text
    'TextDirectionGallery',         # Text box text direction
    # Picture
    'PictureTransparencyGallery',   # Image transparency
}

# =============================================================================
# COMMAND CATEGORIZATION BY OBJECT TYPE
# =============================================================================

# Table commands - require a table
TABLE_COMMANDS = {
    'TableStylesGalleryWord', 'BorderStylesGallery',
    # 'OutlineWeightGallery',  # Moved to SMARTART_COMMANDS
}

# Text box commands - require a floating textbox shape with text
TEXTBOX_COMMANDS = {
    'TextAlignGallery',
    'TextDirectionGallery',
}

# Picture commands - require a picture (shape placeholder)
PICTURE_COMMANDS = {
    'PictureSaturationGallery',
    'PictureTransparencyGallery',
    'PictureRecolorGallery',
    'PictureArtisticEffectsGallery',
    'PictureEffectsShadowGallery',
    'PictureEffectsGlowGallery',
    'PictureEffectsSoftEdgesGallery',
    'PictureReflectionGallery',
    'PictureRotationGallery',
    '_3DBevelPictureTopGallery',
    'PictureShapeGallery',
    'PictureEffectsPresetGallery',
    'PictureStylesGallery',
    'PicturesConvertToSmartArt',
    'PictureColorTempertatureGallery',
    'PictureSharpenSoftenGallery',
    'OutlineColorPicker',  # Picture Border color
    'RecolorColorPicker',  # Picture recolor effect
}

# Equation commands - require an equation
EQUATION_COMMANDS = {
    'EquationFractionGallery', 'EquationScriptGallery', 'EquationRadicalGallery',
    'EquationIntegralGallery', 'EquationLargeOperatorGallery', 'EquationDelimiterGallery',
    'EquationFunctionGallery', 'EquationAccentGallery', 'EquationLimitGallery',
    'EquationOperatorGallery', 'EquationMatrixGallery', 'EquationSymbolsInsertGallery',
}

# SmartArt commands - require a SmartArt
SMARTART_COMMANDS = {
    'SmartArtLayoutGallery', 'SmartArtStylesGallery', 'SmartArtChangeColorsGallery',
    'OutlineWeightGallery',  # SmartArt Tools outline weight
    'OutlineDashesGallery',  # SmartArt Tools dashes
    'ArrowStyleGallery',     # SmartArt Tools arrows
    'GallerySketchyShapes',  # SmartArt Tools sketched style
}

# Chart commands - require a chart (default: clustered column)
CHART_COMMANDS = {
    'ChartStylesGallery', 'ChartColorsGallery', 'ChartLayoutGallery',
    'ChartLegend', 'ChartDataLabel', 'ChartDataTable',
    'ChartTrendline', 'ChartErrorBars',
    'ChartAxesGallery', 'ColorPickerChartLineStyles', 'ChartTitle',
    'ChartGridlinesGallery',
}

# Chart commands that require a line chart (drop lines, hi-lo lines, up/down bars)
CHART_LINE_COMMANDS = {
    'ChartLines', 'ChartUpDownBars',
}

# Note: CHART_3D_COMMANDS removed — all reclassified to legacy
# Backed up in _deprecated_chart_legacy_code.py

# Graphic/SVG commands - require an SVG graphic (msoGraphic, Type=28)
GRAPHIC_COMMANDS = {
    'GraphicsColorPicker', 'GraphicsOutlineColorPicker',
}

# Shape commands - require a shape
SHAPE_COMMANDS = {
    'ShapeStylesGallery', 'ShapeFillColorPicker', 'ShapeOutlineColorPicker',
    'GradientGallery', 'ShapeFillTextureGallery', 'ShapeChangeShapeGallery',
    'ObjectEffectPresetGallery', 'ObjectEffectShadowGallery',
    'ObjectEffectGlowGallery', 'ObjectEffectSoftEdgesGallery',
    'ReflectionGallery', 'BevelShapeGallery',
    'GlowColorPicker', '_3DRotationGallery',
}

# SmartArt WordArtStyles commands - require a SmartArt (text effects on SmartArt nodes)
SMARTART_TEXT_COMMANDS = {
    'TextFillColorPicker', 'TextEffectTransformGallery', 'BevelTextGallery',
    'TextEffects3DRotationGallery',
}

# =============================================================================
# COMMANDS TO SKIP (Deprioritized - outside PM list scope)
# =============================================================================

SKIP_COMMANDS = {
    # =========================================================================
    # TECHNICAL LIMITATIONS — cannot generate via COM API
    # =========================================================================
    'GallerySketchyShapes': 'com_not_supported — sketchy outline is Office 365 feature, Line.Style only sets double-line patterns',
    'ShapeFillTextureGallery': 'com_no_xml_serialization — PresetTextured() returns OK but does not serialize to OOXML on save',
    'ShapeStylesOtherThemeFillsGallery': 'com_no_xml_serialization — ObjectThemeColor returns OK but does not serialize to OOXML on save',
    'LineThickness': 'requires_ink_object_not_shape',
    'TextStylesGallery': 'com_textframe2_not_serialized — composite WordArt style requires multiple XML elements, COM TextFrame2 API does not serialize',
    'TextFillTextureGallery': 'com_no_xml_serialization — Font.Fill.PresetTextured() does not serialize to OOXML via COM',
    'PictureEffectsPresetGallery': 'preset_names_not_descriptive — Preset 4/10 etc. are opaque numbered presets with no meaningful label for instruction',
    'GraphicsStylesGallery': 'no_preset_mapping — ShapeStyle not supported on msoGraphic; composite style presets not mapped to actual Word gallery items',
    'Gallery3DModelViewPresets': 'requires_3d_model_file — Add3DModel needs a valid .glb/.obj file, no built-in 3D model available',
    '_3DModelSceneGallery': 'requires_3d_model_file — same as Gallery3DModelViewPresets',
    'InkColorPicker': 'requires_ink_object — ink strokes require pen/stylus input, cannot be created via COM',
    'PensGallery': 'requires_ink_object — same as InkColorPicker',
    'PictureStylesGallery': 'com_not_supported — composite style (border+shadow+3D+reflection); COM ShapeStyle not available on InlineShapes, requires hardcoded OOXML injection',

    # =========================================================================
    # NO COM READ API — can generate gold but word_object cannot verify
    # =========================================================================
    'PictureShapeGallery': 'no_com_read — InlineShape.AutoShapeType always returns -2 (unknown) for crop shapes',
    'PictureTransparencyGallery': 'no_com_read — InlineShape fill_transparency always returns 1.0; picture transparency not readable via COM',
    'RecolorColorPicker': 'no_com_read — PictureFormat.ColorType returns same value for all recolor presets',
    'ChartErrorBars': 'no_com_read — error_bar_type is write-only via COM; only has_error_bars is readable, cannot distinguish percentage vs standard_error',

    # =========================================================================
    # RECLASSIFY — belongs in a different test set or is a container
    # =========================================================================
    'ChartAxesGallery': 'container_gallery — parent menu, not directly executable',
    'ChartAxisTitlesGallery': 'container_gallery — parent menu, not directly executable',
    # 'ChartTitle': removed from skip — has parameterized variants (above chart, centered overlay)
    'PicturesConvertToSmartArt': 'com_crashes_word — converts picture to SmartArt, destabilizes COM connection for subsequent commands',
    'ChartStylesGallery': 'no_tooltip_preset — Word UI has no tooltip, applies opaque preset style index with no describable parameters',

    # =========================================================================
    # SPINNER CONTROLS — reclassified from generic_control;
    # not yet implemented (need COM property setting, not ExecuteMso)
    # =========================================================================
    'ShapeHeight': 'spinner_control — set Shape.Height via COM',
    'ShapeWidth': 'spinner_control — set Shape.Width via COM',
    'TableColumnWidth': 'spinner_control — set Column.Width via COM',
    'TableRowHeight': 'spinner_control — set Row.Height via COM',
    'HeaderFooterPositionHeaderFromTop': 'spinner_control — set HeaderDistance via COM',
    'HeaderFooterPositionFooterFromBottom': 'spinner_control — set FooterDistance via COM',
}

# =============================================================================
# SPECIAL HANDLING COMMANDS (none for now - all deprioritized moved to SKIP)
# =============================================================================

SPECIAL_HANDLING_COMMANDS = {}

# =============================================================================
# PRECONDITION TYPE MAPPING
# =============================================================================

def get_precondition_type(idMso: str) -> str:
    """
    Determine the precondition type (object type) for a command.

    Returns:
        One of: 'table', 'picture', 'equation', 'chart', 'smartart',
                'shape'
    """
    if idMso == 'TextDirectionGallery':
        return 'textbox_text_direction'
    elif idMso in TEXTBOX_COMMANDS:
        return 'textbox'
    elif idMso in TABLE_COMMANDS:
        return 'table'
    elif idMso in PICTURE_COMMANDS:
        return 'picture'
    elif idMso in EQUATION_COMMANDS:
        return 'equation'
    elif idMso in CHART_COMMANDS:
        return 'chart'
    elif idMso in CHART_LINE_COMMANDS:
        return 'chart_line'
    elif idMso in SMARTART_COMMANDS or idMso in SMARTART_TEXT_COMMANDS:
        return 'smartart'
    elif idMso in GRAPHIC_COMMANDS:
        return 'graphic'
    elif idMso in SHAPE_COMMANDS:
        return 'shape'
    else:
        if idMso in PRIORITIZED_COMMANDS:
            return 'textbox'
        return 'table'


def _load_instruction_overrides() -> Dict[str, Dict[str, str]]:
    """Load hand-written instruction overrides from the instructions JSON file.

    Returns:
        Dict mapping command -> {variant_key -> instruction_text}
    """
    if not INSTRUCTIONS_OVERRIDE_PATH.exists():
        return {}
    with open(INSTRUCTIONS_OVERRIDE_PATH, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data.get('commands', {})


def load_contextual_parameterized_variants() -> Dict:
    """Load variant definitions from contextual parameterized gallery variant file.

    If hand-written instruction overrides exist in the instructions JSON,
    they take priority over auto-generated instructions from the variants file.
    """
    result = {}

    if CONTEXTUAL_PARAMETERIZED_VARIANTS_PATH.exists():
        with open(CONTEXTUAL_PARAMETERIZED_VARIANTS_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
            result.update(data.get('commands', {}))
    else:
        print(f"[WARNING] Contextual parameterized variants file not found: {CONTEXTUAL_PARAMETERIZED_VARIANTS_PATH}")

    if EQUATION_GALLERY_VARIANTS_PATH.exists():
        with open(EQUATION_GALLERY_VARIANTS_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
            result.update(data.get('commands', {}))
    else:
        print(f"[WARNING] Equation gallery variants file not found: {EQUATION_GALLERY_VARIANTS_PATH}")

    # Apply hand-written instruction overrides
    overrides = _load_instruction_overrides()
    override_count = 0
    for cmd_name, variant_instructions in overrides.items():
        if cmd_name not in result:
            continue
        variants = result[cmd_name].get('variants', [])
        for variant in variants:
            variant_key = f"{variant['label']}"
            if variant_key in variant_instructions:
                variant['instruction'] = variant_instructions[variant_key]
                override_count += 1
    if override_count > 0:
        print(f"  [Instructions] Applied {override_count} hand-written instruction overrides")

    return result


def extract_parameters(wd, idMso: str) -> List[Dict]:
    """
    Extract all parameters for a command from variant definitions.

    Returns:
        List of parameter dictionaries, each with 'label' and command-specific fields.
    """
    variants_dict = load_contextual_parameterized_variants()

    if idMso not in variants_dict:
        print(f"    [WARNING] Command '{idMso}' not found in variant definitions")
        return []

    variants = variants_dict[idMso].get('variants', [])
    # Filter to only variants marked for generation
    variants = [v for v in variants if v.get('generate', True)]

    if not variants:
        print(f"    [WARNING] No variants to generate for '{idMso}'")
        return []

    print(f"    [Variants] {idMso}: {len(variants)} variants")
    parameters = []
    for variant in variants:
        param = {
            'label': variant['label'],
            'instruction': variant.get('instruction'),
        }
        # Add all variant-specific fields dynamically
        for key in ['style_index', 'style_name', 'color', 'color_rgb', 'color_bgr', 'color_hex',
                    'theme_color',
                    'effect_index', 'weight', 'dash_style', 'line_style', 'direction', 'preset',
                    'transform', 'recolor_index', 'saturation', 'temperature', 'transparency',
                    'transparency_pct', 'softness', 'sharpness', 'reflection_index', 'glow_radius',
                    'bevel_type', 'rotation_x', 'rotation_y', 'rotation_z', 'const_name',
                    # Equation-specific fields
                    'omath_function', 'frac_type', 'accent_char', 'nary_char',
                    'numerator', 'denominator', 'preset_params',
                    'hide_sub', 'hide_sup', 'sub_sup_lim', 'hide_degree', 'degree_text', 'degree',
                    'begin_char', 'end_char', 'sep_char', 'num_args',
                    'group_char', 'position_top', 'func_name',
                    'display_func_name', 'argument', 'is_inverse',
                    'rows', 'cols', 'matrix_type', 'matrix_inside', 'has_dots', 'has_brace',
                    'wrapper_begin_char', 'wrapper_end_char',
                    'text_content', 'operator_text', 'preset_group',
                    'strikethrough', 'num_rows', 'size',
                    'section', 'name',
                    # Picture effect fields
                    'shadow_type', 'shadow_style', 'shadow_size',
                    'offset_x', 'offset_y', 'blur',
                    'glow_color_rgb', 'glow_theme_color', 'scheme_color', 'soft_edge_radius',
                    'reflection_type', 'preset_camera', 'camera_preset', 'ooxml_camera_preset', 'ooxml_effect_extent',
                    'auto_shape_type', 'geometry', 'bevel_top_preset', 'bevel_top_width',
                    'bevel_top_inset', 'preset_index',
                    'begin_arrowhead_style', 'end_arrowhead_style',
                     'gradient_style', 'gradient_variant', 'gradient_preset', 'texture_index',
                     'ocg_style', 'ocg_variant', 'ocg_degree', 'ocg_angle',
                     'ocg_stop_brightness', 'requires_visible_progression',
                     'gradient_name',
                    # Saturation, crop, color temperature, sharpen/soften fields
                    'saturation_pct', 'crop_left', 'crop_top', 'crop_right', 'crop_bottom',
                     'color_temp', 'amount',
                     # Text effect fields
                    'alignment_wml', 'orientation', 'vertical_anchor',
                    'transform_preset', 'bevel_preset',
                     'warp_format',
                     # Picture recolor/artistic fields
                     'color_type', 'effect_type', 'gallery_item_name', 'duotone_hex',
                     # SmartArt fields
                     'layout_index', 'color_index',
                    # Chart fields
                    'chart_style', 'chart_color', 'chart_layout',
                    'show_legend', 'legend_position', 'show_data_labels', 'label_position', 'show_data_table', 'show_legend_key',
                    'trend_type', 'show_error_bars', 'error_bar_type', 'error_amount', 'line_type',
                    'title_position', 'show', 'axis_type', 'gridline_type',
                    # Classic 3D fields
                    'threed_format', 'lighting_direction',
                    'preset_material', 'extrusion_depth',
                    # WordArt fields
                    'preset_text_effect', 'preset_shape']:
            if key in variant:
                param[key] = variant[key]
        parameters.append(param)

    return parameters


def _apply_equation_structure(wd, omath, param: Dict) -> bool:
    """Build an equation structure into the OMath object and fill default content."""
    func_type = param.get('omath_function')
    preset = param.get('preset_group')

    rng = omath.Range
    rng.Collapse(0)  # wdCollapseEnd

    # Text-only variants (symbols, differentials, operators, dots)
    if func_type is None:
        text = param.get('text_content') or param.get('operator_text')
        if preset == 'basic_operator':
            text = {
                'colon_equals': 'a:=b',
                'def_equals': 'a≝b',
            }.get(param.get('label', ''), text)
        if text:
            rng.Text = text
        return True

    # --- Fraction (type 7) ---
    if func_type == 7:
        func = omath.Functions.Add(rng, 7)
        frac_type = param.get('frac_type', 0)
        func.Frac.Type = frac_type
        if preset == 'common_fraction':
            _fill_fraction_preset(func, param.get('label', ''))
        else:
            func.Frac.Num.Range.Text = 'a'
            func.Frac.Den.Range.Text = 'b'
        return True

    # --- Accent (type 1) ---
    if func_type == 1:
        func = omath.Functions.Add(rng, 1)
        accent_char = param.get('accent_char')
        if accent_char:
            func.Acc.Char = accent_char
        if preset == 'common_accent':
            _fill_accent_preset(omath, func, param.get('label', ''))
        else:
            func.Acc.E.Range.Text = 'a'
        return True

    # --- Bar (type 2) ---
    if func_type == 2:
        func = omath.Functions.Add(rng, 2)
        position_top = param.get('position_top', True)
        func.Bar.BarTop = position_top
        if preset == 'common_accent' and param.get('label') == 'x_oplus_y':
            func.Bar.E.Range.Text = 'x\u2295y'
        else:
            func.Bar.E.Range.Text = 'a'
        return True

    # --- Box (type 3) ---
    if func_type == 3:
        func = omath.Functions.Add(rng, 3)
        if param.get('strikethrough'):
            func.Box.StrikeH = True
        func.Box.E.Range.Text = 'a+b'
        return True

    # --- BorderBox (type 4) ---
    if func_type == 4:
        func = omath.Functions.Add(rng, 4)
        label = param.get('label', '')
        if label == 'boxed_example':
            # Build real professional superscript functions inside the border
            # box. Calling OMath.BuildUp() after assigning linear text to an
            # existing BorderBox leaves the carets as a literal text run in
            # current Word builds.
            argument = func.BorderBox.E
            argument.Range.Text = 'a=b+c'
            # Use the Word Characters collection instead of numeric Range
            # offsets: professional math letters are exposed as astral Unicode
            # characters and occupy two Word range positions apiece.
            for character_index in (5, 3, 1):
                base = argument.Range.Characters(character_index).Duplicate
                script = argument.Functions.Add(base, 19)  # wdOMathFunctionScrSup
                script.ScrSup.Sup.Range.Text = '2'
        else:
            func.BorderBox.E.Range.Text = 'a+b'
        return True

    # --- Delim (type 5) ---
    if func_type == 5:
        func = omath.Functions.Add(rng, 5)
        beg = param.get('begin_char')
        end = param.get('end_char')
        if beg is not None:
            func.Delim.BegChar = beg
        if end is not None:
            func.Delim.EndChar = end
        sep = param.get('sep_char')
        if sep is not None:
            func.Delim.SepChar = sep
        num_args = param.get('num_args', 1)
        for _ in range(num_args - 1):
            func.Delim.E.Add()
        if beg == 0x0000:
            func.Delim.NoLeftChar = True
        if end == 0x0000:
            func.Delim.NoRightChar = True
        # Fill content
        labels = ['a', 'b', 'c', 'd']
        for i in range(min(num_args, len(labels))):
            try:
                func.Delim.E(i + 1).Range.Text = labels[i]
            except Exception:
                pass
        return True

    # --- EqArray (type 6) ---
    if func_type == 6:
        func = omath.Functions.Add(rng, 6)
        num_rows = param.get('num_rows', param.get('rows', 2))
        for _ in range(num_rows - 1):
            func.EqArray.E.Add()
        labels = ['a', 'b', 'c']
        for i in range(min(num_rows, len(labels))):
            try:
                func.EqArray.E(i + 1).Range.Text = labels[i]
            except Exception:
                pass
        return True

    # --- Func (type 8) ---
    if func_type == 8:
        func = omath.Functions.Add(rng, 8)
        func_name = param.get('func_name')
        if func_name:
            func.Func.FName.Range.Text = func_name
        func.Func.E.Range.Text = param.get('argument', 'x')
        return True

    # --- GroupChar (type 9) ---
    if func_type == 9:
        func = omath.Functions.Add(rng, 9)
        char = param.get('group_char')
        if char:
            func.GroupChar.Char = char
        align_top = param.get('position_top')
        if align_top is not None:
            func.GroupChar.AlignTop = align_top
        func.GroupChar.E.Range.Text = 'n'
        return True

    # --- LimLow (type 10) ---
    if func_type == 10:
        func = omath.Functions.Add(rng, 10)
        gc_char = param.get('group_char')
        if gc_char:
            # Grouping Character Below: LimLow wrapping GroupChar
            gc_func = func.LimLow.E.Functions.Add(func.LimLow.E.Range, 9)
            gc_func.GroupChar.Char = gc_char
            gc_func.GroupChar.E.Range.Text = 'a+b'
            func.LimLow.Lim.Range.Text = 'x'
        else:
            label = param.get('label', '')
            func_name = param.get('func_name')
            if func_name:
                func.LimLow.E.Range.Text = func_name
            if label == 'logarithm_with_empty_base':
                func.LimLow.Lim.Range.Text = '2'
            elif label == 'limit':
                func.LimLow.Lim.Range.Text = 'x→0'
            elif label in ('minimum', 'maximum'):
                func.LimLow.Lim.Range.Text = '0≤x≤1'
        return True

    # --- LimUpp (type 11) ---
    if func_type == 11:
        func = omath.Functions.Add(rng, 11)
        gc_char = param.get('group_char')
        if gc_char:
            # Grouping Character Above: LimUpp wrapping GroupChar
            gc_func = func.LimUpp.E.Functions.Add(func.LimUpp.E.Range, 9)
            gc_func.GroupChar.Char = gc_char
            gc_func.GroupChar.AlignTop = True
            gc_func.GroupChar.E.Range.Text = 'a+b'
            func.LimUpp.Lim.Range.Text = 'x'
        else:
            func.LimUpp.E.Range.Text = 'x'
            func.LimUpp.Lim.Range.Text = 'n'
        return True

    # --- Mat (type 12) ---
    if func_type == 12:
        func = omath.Functions.Add(rng, 12)
        size = param.get('size', 2)
        rows = param.get('rows', size)
        cols = param.get('cols', size)
        mat = func.Mat
        for _ in range(cols - 1):
            mat.Cols.Add()
        for _ in range(rows - 1):
            mat.Rows.Add()
        if preset == 'common_matrix':
            _fill_matrix_preset(mat, param)
        else:
            # Fill with sequential values
            vals = iter('abcdefghijklmnopqrstuvwxyz')
            for r in range(1, rows + 1):
                for cc in range(1, cols + 1):
                    try:
                        mat.Cell(r, cc).Range.Text = next(vals)
                    except (StopIteration, Exception):
                        pass
        return True

    # --- Nary (type 13) — integrals, large operators ---
    if func_type == 13:
        func = omath.Functions.Add(rng, 13)
        char = param.get('nary_char')
        if char:
            func.Nary.Char = char
        hide_sub = param.get('hide_sub')
        hide_sup = param.get('hide_sup')
        if hide_sub is not None:
            func.Nary.HideSub = hide_sub
        if hide_sup is not None:
            func.Nary.HideSup = hide_sup
        if hide_sub is not None and (not hide_sub or not hide_sup):
            func.Nary.SubSupLim = param.get('sub_sup_lim', False)
        if preset == 'common_large_operator' and param.get('label') == 'summation_from_i_equals_zero_to_n':
            func.Nary.E.Range.Text = 'x'
            func.Nary.Sub.Range.Text = 'i=0'
            func.Nary.Sup.Range.Text = 'n'
        else:
            # Fill content
            func.Nary.E.Range.Text = 'x'
            if not hide_sub:
                func.Nary.Sub.Range.Text = '0'
            if not hide_sup:
                func.Nary.Sup.Range.Text = '1'
        return True

    # --- ScrPre (type 15) ---
    if func_type == 15:
        func = omath.Functions.Add(rng, 15)
        if preset == 'common_script':
            _fill_script_preset(omath, func, param.get('label', ''))
        else:
            func.ScrPre.E.Range.Text = 'x'
            func.ScrPre.Sub.Range.Text = '1'
            func.ScrPre.Sup.Range.Text = 'n'
        return True

    # --- Rad (type 16) ---
    if func_type == 16:
        func = omath.Functions.Add(rng, 16)
        hide_deg = param.get('hide_degree', True)
        func.Rad.HideDeg = hide_deg
        degree = param.get('degree_text', param.get('degree'))
        if degree and not hide_deg:
            func.Rad.Deg.Range.Text = str(degree)
        if preset == 'common_radical':
            _fill_radical_preset(omath, func, param.get('label', ''))
        else:
            func.Rad.E.Range.Text = 'a'
        return True

    # --- ScrSub (type 17) ---
    if func_type == 17:
        func = omath.Functions.Add(rng, 17)
        if preset == 'common_script':
            _fill_script_preset(omath, func, param.get('label', ''))
        else:
            func.ScrSub.E.Range.Text = 'x'
            func.ScrSub.Sub.Range.Text = 'i'
        return True

    # --- ScrSubSup (type 18) ---
    if func_type == 18:
        func = omath.Functions.Add(rng, 18)
        if preset == 'common_script':
            _fill_script_preset(omath, func, param.get('label', ''))
        else:
            func.ScrSubSup.E.Range.Text = 'x'
            func.ScrSubSup.Sub.Range.Text = 'i'
            func.ScrSubSup.Sup.Range.Text = '2'
        return True

    # --- ScrSup (type 19) ---
    if func_type == 19:
        func = omath.Functions.Add(rng, 19)
        if preset == 'common_script':
            _fill_script_preset(omath, func, param.get('label', ''))
        else:
            func.ScrSup.E.Range.Text = 'x'
            func.ScrSup.Sup.Range.Text = '2'
        return True

    print(f"    [WARN] Unknown omath_function type: {func_type}")
    return False


# =============================================================================
# Preset content fillers
# =============================================================================

def _fill_fraction_preset(func, label: str):
    """Fill fraction preset content."""
    presets = {
        'dy_over_dx': ('dy', 'dx'),
    }
    num, den = presets.get(label, ('a', 'b'))
    func.Frac.Num.Range.Text = num
    func.Frac.Den.Range.Text = den


def _fill_accent_preset(omath, func, label: str):
    """Fill accent preset content."""
    if label == 'vector_a':
        func.Acc.E.Range.Text = 'A'
    elif label == 'ABC_overbar':
        func.Acc.E.Range.Text = 'ABC'
    else:
        func.Acc.E.Range.Text = 'a'


def _fill_script_preset(omath, func, label: str):
    """Fill script preset content."""
    if label == 'n_1_Y':
        func.ScrPre.E.Range.Text = 'Y'
        func.ScrPre.Sub.Range.Text = '1'
        func.ScrPre.Sup.Range.Text = 'n'


def _fill_radical_preset(omath, func, label: str):
    """Fill radical preset content."""
    if label == 'pythagorean_magnitude':
        func.Rad.E.Range.Text = 'a\u00b2+b\u00b2'
    else:
        func.Rad.E.Range.Text = 'a'


def _fill_matrix_preset(mat, param: Dict):
    """Fill matrix preset content."""
    label = param.get('label', '')
    size = param.get('size', 2)
    rows = param.get('rows', size)
    cols = param.get('cols', size)
    if label == 'identity_2x2_with_blank_off_diagonal_cells':
        for r in range(1, rows + 1):
            for cc in range(1, cols + 1):
                try:
                    if r == cc:
                        mat.Cell(r, cc).Range.Text = '1'
                except Exception:
                    pass
    elif 'identity' in label:
        for r in range(1, rows + 1):
            for cc in range(1, cols + 1):
                try:
                    mat.Cell(r, cc).Range.Text = '1' if r == cc else '0'
                except Exception:
                    pass
    elif label in {'sparse_3x3_parentheses', 'sparse_3x3_brackets'}:
        for r in range(1, rows + 1):
            for cc in range(1, cols + 1):
                try:
                    if r == 1 and cc == 1:
                        mat.Cell(r, cc).Range.Text = 'a11'
                    elif r == 1 and cc == 2:
                        mat.Cell(r, cc).Range.Text = '⋯'
                    elif r == 1 and cc == 3:
                        mat.Cell(r, cc).Range.Text = 'a1n'
                    elif r == 2 and cc == 1:
                        mat.Cell(r, cc).Range.Text = '⋮'
                    elif r == 2 and cc == 2:
                        mat.Cell(r, cc).Range.Text = '⋱'
                    elif r == 2 and cc == 3:
                        mat.Cell(r, cc).Range.Text = '⋮'
                    elif r == 3 and cc == 1:
                        mat.Cell(r, cc).Range.Text = 'an1'
                    elif r == 3 and cc == 2:
                        mat.Cell(r, cc).Range.Text = '⋯'
                    elif r == 3 and cc == 3:
                        mat.Cell(r, cc).Range.Text = 'ann'
                except Exception:
                    pass
    elif 'sparse' in label:
        for r in range(1, rows + 1):
            for cc in range(1, cols + 1):
                try:
                    if r == cc:
                        mat.Cell(r, cc).Range.Text = f'a{r}'
                    else:
                        mat.Cell(r, cc).Range.Text = '0'
                except Exception:
                    pass


def _fix_glow_scheme_color(gold_path: Path, color_rgb: int, scheme_color: str):
    """Post-process gold.docx to replace srgbClr with schemeClr for glow effect.

    Word UI applies glow gallery presets using schemeClr with satMod+alpha transforms,
    but COM API's Glow.Color.RGB only produces srgbClr. This patches the XML to match.
    """
    from zipfile import ZipFile, ZIP_DEFLATED
    from lxml import etree

    rgb_hex = f'{color_rgb:06X}'
    drawingml = 'http://schemas.openxmlformats.org/drawingml/2006/main'

    with ZipFile(gold_path, 'r') as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    root = etree.fromstring(files['word/document.xml'])
    patched = False
    for glow in root.iter(f'{{{drawingml}}}glow'):
        color = next(iter(glow), None)
        if color is None:
            continue
        local_name = etree.QName(color).localname
        value = color.get('val', '')
        if not (
            (local_name == 'srgbClr' and value.upper() == rgb_hex)
            or (local_name == 'schemeClr' and value == scheme_color)
        ):
            continue
        replacement = etree.Element(f'{{{drawingml}}}schemeClr')
        replacement.set('val', scheme_color)
        saturation = etree.SubElement(replacement, f'{{{drawingml}}}satMod')
        saturation.set('val', '175000')
        alpha = etree.SubElement(replacement, f'{{{drawingml}}}alpha')
        alpha.set('val', '40000')
        glow.replace(color, replacement)
        patched = True

    if not patched:
        return

    files['word/document.xml'] = etree.tostring(
        root, xml_declaration=True, encoding='UTF-8', standalone='yes'
    )

    write_docx_package(files, gold_path)


def _fix_groupchar_above(gold_path: Path):
    """Post-process gold.docx to add pos=top and vertJc=bot to groupChrPr for ⏞.

    Word COM's GroupChar.AlignTop = True doesn't emit the pos/vertJc attributes
    that Word's gallery UI produces. This patches the XML to match.
    """
    from zipfile import ZipFile, ZIP_DEFLATED
    from lxml import etree

    m_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

    with ZipFile(gold_path, 'r') as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    doc_xml = etree.fromstring(files['word/document.xml'])

    for gcpr in doc_xml.iter(f'{{{m_ns}}}groupChrPr'):
        chr_elem = gcpr.find(f'{{{m_ns}}}chr')
        if chr_elem is not None and chr_elem.get(f'{{{m_ns}}}val') == '\u23DE':
            # Add pos="top" if missing
            pos_elem = gcpr.find(f'{{{m_ns}}}pos')
            if pos_elem is None:
                pos_elem = etree.SubElement(gcpr, f'{{{m_ns}}}pos')
                # Insert after chr element
                gcpr.remove(pos_elem)
                chr_idx = list(gcpr).index(chr_elem)
                gcpr.insert(chr_idx + 1, pos_elem)
            pos_elem.set(f'{{{m_ns}}}val', 'top')

            # Add vertJc="bot" if missing
            vertjc_elem = gcpr.find(f'{{{m_ns}}}vertJc')
            if vertjc_elem is None:
                vertjc_elem = etree.SubElement(gcpr, f'{{{m_ns}}}vertJc')
                # Insert after pos element
                gcpr.remove(vertjc_elem)
                pos_idx = list(gcpr).index(pos_elem)
                gcpr.insert(pos_idx + 1, vertjc_elem)
            vertjc_elem.set(f'{{{m_ns}}}val', 'bot')

    files['word/document.xml'] = etree.tostring(
        doc_xml, xml_declaration=True, encoding='UTF-8', standalone='yes'
    )

    write_docx_package(files, gold_path)


def _fix_groupchar_position(docx_path: Path, position_top: bool):
    """Match Word UI structure for operator arrows."""
    from zipfile import ZipFile, ZIP_DEFLATED
    from lxml import etree

    m_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

    with ZipFile(docx_path, 'r') as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    doc_xml = etree.fromstring(files['word/document.xml'])

    target = None
    parent = None
    for group_chr in doc_xml.iter(f'{{{m_ns}}}groupChr'):
        if group_chr.getparent() is not None and group_chr.getparent().tag == f'{{{m_ns}}}oMath':
            target = group_chr
            parent = group_chr.getparent()
            break

    if target is None or parent is None:
        return

    gcpr = target.find(f'{{{m_ns}}}groupChrPr')
    if gcpr is None:
        return
    chr_elem = gcpr.find(f'{{{m_ns}}}chr')
    if chr_elem is None or chr_elem.get(f'{{{m_ns}}}val') not in {'→', '←', '⇒', '⇐', '↔'}:
        return

    ctrl_pr = gcpr.find(f'{{{m_ns}}}ctrlPr')
    gcpr[:] = []
    gcpr.append(chr_elem)
    if position_top:
        vertjc_elem = etree.Element(f'{{{m_ns}}}vertJc')
        vertjc_elem.set(f'{{{m_ns}}}val', 'bot')
        gcpr.append(vertjc_elem)
    else:
        pos_elem = etree.Element(f'{{{m_ns}}}pos')
        pos_elem.set(f'{{{m_ns}}}val', 'top')
        gcpr.append(pos_elem)
    if ctrl_pr is not None:
        gcpr.append(ctrl_pr)

    box = etree.Element(f'{{{m_ns}}}box')
    box_pr = etree.SubElement(box, f'{{{m_ns}}}boxPr')
    op_emu = etree.SubElement(box_pr, f'{{{m_ns}}}opEmu')
    op_emu.set(f'{{{m_ns}}}val', '1')
    etree.SubElement(box_pr, f'{{{m_ns}}}ctrlPr')
    box_e = etree.SubElement(box, f'{{{m_ns}}}e')
    parent.replace(target, box)
    box_e.append(target)

    files['word/document.xml'] = etree.tostring(
        doc_xml, xml_declaration=True, encoding='UTF-8', standalone='yes'
    )

    write_docx_package(files, docx_path)


def _wrap_eqarray_with_delimiters(docx_path: Path, begin_char: str, end_char: str):
    """Wrap the first top-level eqArr in a delimiter object."""
    from zipfile import ZipFile, ZIP_DEFLATED
    from lxml import etree

    m_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

    with ZipFile(docx_path, 'r') as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    doc_xml = etree.fromstring(files['word/document.xml'])

    target = None
    parent = None
    for elem in doc_xml.iter(f'{{{m_ns}}}eqArr'):
        if elem.getparent() is not None and elem.getparent().tag == f'{{{m_ns}}}oMath':
            target = elem
            parent = elem.getparent()
            break

    if target is None or parent is None:
        return

    d = etree.Element(f'{{{m_ns}}}d')
    dpr = etree.SubElement(d, f'{{{m_ns}}}dPr')
    beg = etree.SubElement(dpr, f'{{{m_ns}}}begChr')
    beg.set(f'{{{m_ns}}}val', begin_char)
    end = etree.SubElement(dpr, f'{{{m_ns}}}endChr')
    end.set(f'{{{m_ns}}}val', end_char)

    eqarr_pr = target.find(f'{{{m_ns}}}eqArrPr')
    if eqarr_pr is not None:
        ctrl_pr = eqarr_pr.find(f'{{{m_ns}}}ctrlPr')
        if ctrl_pr is not None:
            dpr.append(etree.fromstring(etree.tostring(ctrl_pr)))

    e = etree.SubElement(d, f'{{{m_ns}}}e')
    parent.replace(target, d)
    e.append(target)

    files['word/document.xml'] = etree.tostring(
        doc_xml, xml_declaration=True, encoding='UTF-8', standalone='yes'
    )

    write_docx_package(files, docx_path)


def _wrap_fraction_as_small(docx_path: Path):
    """Wrap the first top-level fraction in a small box argument."""
    from zipfile import ZipFile, ZIP_DEFLATED
    from lxml import etree

    m_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

    with ZipFile(docx_path, 'r') as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    doc_xml = etree.fromstring(files['word/document.xml'])

    target = None
    parent = None
    for elem in doc_xml.iter(f'{{{m_ns}}}f'):
        if elem.getparent() is not None and elem.getparent().tag == f'{{{m_ns}}}oMath':
            target = elem
            parent = elem.getparent()
            break

    if target is None or parent is None:
        return

    box = etree.Element(f'{{{m_ns}}}box')
    box_pr = etree.SubElement(box, f'{{{m_ns}}}boxPr')

    frac_pr = target.find(f'{{{m_ns}}}fPr')
    if frac_pr is not None:
        ctrl_pr = frac_pr.find(f'{{{m_ns}}}ctrlPr')
        if ctrl_pr is not None:
            box_pr.append(etree.fromstring(etree.tostring(ctrl_pr)))

    e = etree.SubElement(box, f'{{{m_ns}}}e')
    arg_pr = etree.SubElement(e, f'{{{m_ns}}}argPr')
    arg_sz = etree.SubElement(arg_pr, f'{{{m_ns}}}argSz')
    arg_sz.set(f'{{{m_ns}}}val', '-1')

    parent.replace(target, box)
    e.append(target)

    files['word/document.xml'] = etree.tostring(
        doc_xml, xml_declaration=True, encoding='UTF-8', standalone='yes'
    )

    write_docx_package(files, docx_path)


def _convert_function_name_to_inverse(docx_path: Path):
    """Convert the first top-level function name to a superscript -1 form."""
    from zipfile import ZipFile, ZIP_DEFLATED
    from lxml import etree

    m_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

    with ZipFile(docx_path, 'r') as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    doc_xml = etree.fromstring(files['word/document.xml'])

    func = None
    for elem in doc_xml.iter(f'{{{m_ns}}}func'):
        if elem.getparent() is not None and elem.getparent().tag == f'{{{m_ns}}}oMath':
            func = elem
            break

    if func is None:
        return

    f_name = func.find(f'{{{m_ns}}}fName')
    if f_name is None:
        return

    base_text = ''.join(t.text or '' for t in f_name.iter(f'{{{m_ns}}}t'))
    if not base_text:
        return

    for child in list(f_name):
        f_name.remove(child)

    s_sup = etree.SubElement(f_name, f'{{{m_ns}}}sSup')
    s_sup_pr = etree.SubElement(s_sup, f'{{{m_ns}}}sSupPr')
    etree.SubElement(s_sup_pr, f'{{{m_ns}}}ctrlPr')

    e = etree.SubElement(s_sup, f'{{{m_ns}}}e')
    base_r = etree.SubElement(e, f'{{{m_ns}}}r')
    base_t = etree.SubElement(base_r, f'{{{m_ns}}}t')
    base_t.text = base_text

    sup = etree.SubElement(s_sup, f'{{{m_ns}}}sup')
    sup_r = etree.SubElement(sup, f'{{{m_ns}}}r')
    sup_t = etree.SubElement(sup_r, f'{{{m_ns}}}t')
    sup_t.text = '-1'

    files['word/document.xml'] = etree.tostring(
        doc_xml, xml_declaration=True, encoding='UTF-8', standalone='yes'
    )

    write_docx_package(files, docx_path)


def _wrap_limlow_with_argument(docx_path: Path, argument_text: str):
    """Wrap the first top-level LimLow in a function object with a right-side argument."""
    from zipfile import ZipFile, ZIP_DEFLATED
    from lxml import etree

    m_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

    with ZipFile(docx_path, 'r') as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    doc_xml = etree.fromstring(files['word/document.xml'])

    target = None
    parent = None
    for elem in doc_xml.iter(f'{{{m_ns}}}limLow'):
        if elem.getparent() is not None and elem.getparent().tag == f'{{{m_ns}}}oMath':
            target = elem
            parent = elem.getparent()
            break

    if target is None or parent is None:
        return

    func = etree.Element(f'{{{m_ns}}}func')
    func_pr = etree.SubElement(func, f'{{{m_ns}}}funcPr')

    lim_pr = target.find(f'{{{m_ns}}}limLowPr')
    if lim_pr is not None:
        ctrl_pr = lim_pr.find(f'{{{m_ns}}}ctrlPr')
        if ctrl_pr is not None:
            func_pr.append(etree.fromstring(etree.tostring(ctrl_pr)))

    f_name = etree.SubElement(func, f'{{{m_ns}}}fName')
    parent.replace(target, func)
    f_name.append(target)

    e = etree.SubElement(func, f'{{{m_ns}}}e')
    r = etree.SubElement(e, f'{{{m_ns}}}r')
    t = etree.SubElement(r, f'{{{m_ns}}}t')
    t.text = argument_text

    files['word/document.xml'] = etree.tostring(
        doc_xml, xml_declaration=True, encoding='UTF-8', standalone='yes'
    )

    write_docx_package(files, docx_path)


def _convert_log_with_base_to_subscript(docx_path: Path, base_text: str, argument_text: str):
    """Convert the first top-level LimLow(log, base) into Func(fName=sSub(log, base), e=arg)."""
    from zipfile import ZipFile, ZIP_DEFLATED
    from lxml import etree

    m_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

    with ZipFile(docx_path, 'r') as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    doc_xml = etree.fromstring(files['word/document.xml'])

    target = None
    parent = None
    for elem in doc_xml.iter(f'{{{m_ns}}}limLow'):
        if elem.getparent() is not None and elem.getparent().tag == f'{{{m_ns}}}oMath':
            target = elem
            parent = elem.getparent()
            break

    if target is None or parent is None:
        return

    func = etree.Element(f'{{{m_ns}}}func')
    func_pr = etree.SubElement(func, f'{{{m_ns}}}funcPr')

    lim_pr = target.find(f'{{{m_ns}}}limLowPr')
    if lim_pr is not None:
        ctrl_pr = lim_pr.find(f'{{{m_ns}}}ctrlPr')
        if ctrl_pr is not None:
            func_pr.append(etree.fromstring(etree.tostring(ctrl_pr)))

    f_name = etree.SubElement(func, f'{{{m_ns}}}fName')
    s_sub = etree.SubElement(f_name, f'{{{m_ns}}}sSub')
    s_sub_pr = etree.SubElement(s_sub, f'{{{m_ns}}}sSubPr')
    etree.SubElement(s_sub_pr, f'{{{m_ns}}}ctrlPr')

    e_base = etree.SubElement(s_sub, f'{{{m_ns}}}e')
    e_base_r = etree.SubElement(e_base, f'{{{m_ns}}}r')
    e_base_t = etree.SubElement(e_base_r, f'{{{m_ns}}}t')
    e_base_t.text = 'log'

    sub = etree.SubElement(s_sub, f'{{{m_ns}}}sub')
    sub_r = etree.SubElement(sub, f'{{{m_ns}}}r')
    sub_t = etree.SubElement(sub_r, f'{{{m_ns}}}t')
    sub_t.text = base_text

    e = etree.SubElement(func, f'{{{m_ns}}}e')
    e_r = etree.SubElement(e, f'{{{m_ns}}}r')
    e_t = etree.SubElement(e_r, f'{{{m_ns}}}t')
    e_t.text = argument_text

    parent.replace(target, func)

    files['word/document.xml'] = etree.tostring(
        doc_xml, xml_declaration=True, encoding='UTF-8', standalone='yes'
    )

    write_docx_package(files, docx_path)


def _wrap_matrix_with_delimiters(docx_path: Path, begin_char: str, end_char: str):
    """Wrap the first top-level matrix in a delimiter object."""
    from zipfile import ZipFile, ZIP_DEFLATED
    from lxml import etree

    m_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

    with ZipFile(docx_path, 'r') as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    doc_xml = etree.fromstring(files['word/document.xml'])

    target = None
    parent = None
    for elem in doc_xml.iter(f'{{{m_ns}}}m'):
        if elem.getparent() is not None and elem.getparent().tag == f'{{{m_ns}}}oMath':
            target = elem
            parent = elem.getparent()
            break

    if target is None or parent is None:
        return

    d = etree.Element(f'{{{m_ns}}}d')
    dpr = etree.SubElement(d, f'{{{m_ns}}}dPr')
    beg = etree.SubElement(dpr, f'{{{m_ns}}}begChr')
    beg.set(f'{{{m_ns}}}val', begin_char)
    end = etree.SubElement(dpr, f'{{{m_ns}}}endChr')
    end.set(f'{{{m_ns}}}val', end_char)

    m_pr = target.find(f'{{{m_ns}}}mPr')
    if m_pr is not None:
        ctrl_pr = m_pr.find(f'{{{m_ns}}}ctrlPr')
        if ctrl_pr is not None:
            dpr.append(etree.fromstring(etree.tostring(ctrl_pr)))

    e = etree.SubElement(d, f'{{{m_ns}}}e')
    parent.replace(target, d)
    e.append(target)

    files['word/document.xml'] = etree.tostring(
        doc_xml, xml_declaration=True, encoding='UTF-8', standalone='yes'
    )

    write_docx_package(files, docx_path)


def _apply_picture_transparency_effect(docx_path: Path, transparency: float):
    """Patch the first picture blip with a transparency effect."""
    from zipfile import ZipFile, ZIP_DEFLATED
    from lxml import etree

    a_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'

    with ZipFile(docx_path, 'r') as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    doc_xml = etree.fromstring(files['word/document.xml'])
    blip = doc_xml.find(f'.//{{{a_ns}}}blip')
    if blip is None:
        return

    # Remove existing transparency effects.
    for child in list(blip):
        if child.tag in {
            f'{{{a_ns}}}alphaMod',
            f'{{{a_ns}}}alphaModFix',
        }:
            blip.remove(child)

    if transparency > 0:
        opacity = max(0, min(100000, int(round((1.0 - transparency) * 100000))))
        effect = etree.Element(f'{{{a_ns}}}alphaModFix')
        effect.set('amt', str(opacity))
        ext_lst = blip.find(f'{{{a_ns}}}extLst')
        if ext_lst is None:
            blip.append(effect)
        else:
            ext_idx = list(blip).index(ext_lst)
            blip.insert(ext_idx, effect)

    files['word/document.xml'] = etree.tostring(
        doc_xml, xml_declaration=True, encoding='UTF-8', standalone='yes'
    )

    write_docx_package(files, docx_path)


def _apply_recolor_duotone(docx_path: Path, color_hex: str):
    """Patch the first picture blip with a duotone recolor effect.

    Injects <a:duotone> into the blip element, matching Word's
    RecolorColorPicker behavior (tint=45000, satMod=400000).
    """
    from zipfile import ZipFile, ZIP_DEFLATED
    from lxml import etree

    a_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'

    with ZipFile(docx_path, 'r') as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    doc_xml = etree.fromstring(files['word/document.xml'])
    blip = doc_xml.find(f'.//{{{a_ns}}}blip')
    if blip is None:
        return

    # Remove existing duotone
    for child in list(blip):
        if child.tag == f'{{{a_ns}}}duotone':
            blip.remove(child)

    # Build duotone element
    duotone = etree.Element(f'{{{a_ns}}}duotone')
    prst_clr = etree.SubElement(duotone, f'{{{a_ns}}}prstClr')
    prst_clr.set('val', 'black')
    srgb_clr = etree.SubElement(duotone, f'{{{a_ns}}}srgbClr')
    srgb_clr.set('val', color_hex)
    tint = etree.SubElement(srgb_clr, f'{{{a_ns}}}tint')
    tint.set('val', '45000')
    sat_mod = etree.SubElement(srgb_clr, f'{{{a_ns}}}satMod')
    sat_mod.set('val', '400000')

    # Insert before extLst
    ext_lst = blip.find(f'{{{a_ns}}}extLst')
    if ext_lst is None:
        blip.append(duotone)
    else:
        ext_idx = list(blip).index(ext_lst)
        blip.insert(ext_idx, duotone)

    files['word/document.xml'] = etree.tostring(
        doc_xml, xml_declaration=True, encoding='UTF-8', standalone='yes'
    )

    write_docx_package(files, docx_path)


def _apply_picture_style(docx_path: Path, style_label: str):
    """Patch the first picture with a baked picture style payload.

    Picture styles are composite effects (border + shadow + 3D + reflection)
    that Word applies to InlineShapes. COM API doesn't expose ShapeStyle on
    InlineShapes, so we inject pre-computed OOXML directly.
    """
    from zipfile import ZipFile, ZIP_DEFLATED
    from lxml import etree

    # Pre-baked style payloads extracted from Word-saved reference documents.
    # Each style is a combination of line, shadow, 3D, and reflection effects.
    style_map = {
        'simple_frame_white': {
            # Simple Frame, White — thin white border
            'effect_extent': {'l': '19050', 't': '19050', 'r': '19050', 'b': '19050'},
            'line': {'w': '38100', 'color': 'FFFFFF'},
        },
        'simple_frame_black': {
            # Simple Frame, Black — thin black border
            'effect_extent': {'l': '19050', 't': '19050', 'r': '19050', 'b': '19050'},
            'line': {'w': '38100', 'color': '000000'},
        },
        'metal_frame': {
            # Metal Frame
            'effect_extent': {'l': '76200', 't': '76200', 'r': '76200', 'b': '76200'},
            'line': {'w': '88900', 'color': 'FFFFFF'},
            'camera_prst': 'orthographicFront',
            'light_rig': 'threePt',
            'light_dir': 't',
            'sp3d': {'prstMaterial': 'metal'},
            'bevelT': {'w': '88900', 'h': '88900', 'prst': 'cross'},
        },
        'center_shadow_rectangle': {
            # Center Shadow Rectangle — simple drop shadow
            'effect_extent': {'l': '0', 't': '0', 'r': '0', 'b': '0'},
            'line': {'noFill': True},
            'outer_shadow': {
                'blurRad': '57150',
                'dist': '19050',
                'dir': '5400000',
                'algn': 'ctr',
                'color': '000000',
                'alpha': '63000',
            },
        },
        'beveled_matte_white': {
            # Beveled Matte, White
            'effect_extent': {'l': '38100', 't': '38100', 'r': '48260', 'b': '53975'},
            'line': {'w': '95250', 'color': 'FFFFFF'},
            'camera_prst': 'orthographicFront',
            'light_rig': 'threePt',
            'light_dir': 't',
            'sp3d': {},
            'bevelT': {'w': '50800', 'prst': 'softRound'},
        },
    }

    payload = style_map.get(style_label)
    if payload is None:
        return

    a_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    pic_ns = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
    wp_ns = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'

    with ZipFile(docx_path, 'r') as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    doc_xml = etree.fromstring(files['word/document.xml'])

    # Update effectExtent on inline shape
    inline = doc_xml.find(f'.//{{{wp_ns}}}inline')
    if inline is not None and payload.get('effect_extent'):
        effect = inline.find(f'{{{wp_ns}}}effectExtent')
        if effect is None:
            extent = inline.find(f'{{{wp_ns}}}extent')
            effect = etree.Element(f'{{{wp_ns}}}effectExtent')
            if extent is not None:
                extent_idx = list(inline).index(extent)
                inline.insert(extent_idx + 1, effect)
            else:
                inline.insert(0, effect)
        for key, value in payload['effect_extent'].items():
            effect.set(key, value)

    sp_pr = doc_xml.find(f'.//{{{pic_ns}}}spPr')
    if sp_pr is None:
        return

    # Line / border
    line = sp_pr.find(f'{{{a_ns}}}ln')
    if line is not None:
        sp_pr.remove(line)
    if payload.get('line') is not None:
        line = etree.SubElement(sp_pr, f'{{{a_ns}}}ln')
        line_payload = payload['line']
        if line_payload.get('w'):
            line.set('w', line_payload['w'])
        if line_payload.get('noFill'):
            etree.SubElement(line, f'{{{a_ns}}}noFill')
        elif line_payload.get('color'):
            solid = etree.SubElement(line, f'{{{a_ns}}}solidFill')
            srgb = etree.SubElement(solid, f'{{{a_ns}}}srgbClr')
            srgb.set('val', line_payload['color'])

    # Shadow effect
    effect_lst = sp_pr.find(f'{{{a_ns}}}effectLst')
    if effect_lst is not None:
        sp_pr.remove(effect_lst)
    if payload.get('outer_shadow'):
        effect_lst = etree.SubElement(sp_pr, f'{{{a_ns}}}effectLst')
        shadow = etree.SubElement(effect_lst, f'{{{a_ns}}}outerShdw')
        for key, value in payload['outer_shadow'].items():
            if key in ('color', 'alpha'):
                continue
            shadow.set(key, value)
        srgb = etree.SubElement(shadow, f'{{{a_ns}}}srgbClr')
        srgb.set('val', payload['outer_shadow'].get('color', '000000'))
        if payload['outer_shadow'].get('alpha'):
            alpha = etree.SubElement(srgb, f'{{{a_ns}}}alpha')
            alpha.set('val', payload['outer_shadow']['alpha'])

    # 3D scene
    if payload.get('camera_prst'):
        scene3d = sp_pr.find(f'{{{a_ns}}}scene3d')
        if scene3d is None:
            scene3d = etree.SubElement(sp_pr, f'{{{a_ns}}}scene3d')
        else:
            for child in list(scene3d):
                scene3d.remove(child)
        camera = etree.SubElement(scene3d, f'{{{a_ns}}}camera')
        camera.set('prst', payload['camera_prst'])
        light_rig = etree.SubElement(scene3d, f'{{{a_ns}}}lightRig')
        light_rig.set('rig', payload.get('light_rig', 'threePt'))
        light_rig.set('dir', payload.get('light_dir', 't'))

    # 3D shape properties (bevel, material)
    if payload.get('sp3d') is not None or payload.get('bevelT') is not None:
        sp3d = sp_pr.find(f'{{{a_ns}}}sp3d')
        if sp3d is None:
            sp3d = etree.SubElement(sp_pr, f'{{{a_ns}}}sp3d')
        else:
            for key in list(sp3d.attrib.keys()):
                del sp3d.attrib[key]
            for child in list(sp3d):
                sp3d.remove(child)
        for key, value in payload.get('sp3d', {}).items():
            sp3d.set(key, value)
        if payload.get('bevelT'):
            bevel_t = etree.SubElement(sp3d, f'{{{a_ns}}}bevelT')
            for key, value in payload['bevelT'].items():
                bevel_t.set(key, value)

    files['word/document.xml'] = etree.tostring(
        doc_xml, xml_declaration=True, encoding='UTF-8', standalone='yes'
    )

    write_docx_package(files, docx_path)



def _docx_patch(docx_path: Path, transform_fn):
    """Generic OOXML patch helper: read, transform, write back."""
    from zipfile import ZipFile, ZIP_DEFLATED
    from lxml import etree

    with ZipFile(docx_path, 'r') as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    doc_xml = etree.fromstring(files['word/document.xml'])
    transform_fn(doc_xml)

    files['word/document.xml'] = etree.tostring(
        doc_xml, xml_declaration=True, encoding='UTF-8', standalone='yes'
    )
    write_docx_package(files, docx_path)


def _apply_text_fill_color(docx_path: Path, color_hex: str):
    """Set text color via w:color under w:rPr in shape's txbxContent."""
    from lxml import etree
    w_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    wps_ns = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'

    def transform(doc_xml):
        # Find w:rPr elements inside wps:txbx > w:txbxContent
        for txbx in doc_xml.findall(f'.//{{{wps_ns}}}txbx'):
            for rPr in txbx.findall(f'.//{{{w_ns}}}rPr'):
                # Remove existing w:color
                for old in rPr.findall(f'{{{w_ns}}}color'):
                    rPr.remove(old)
                # Add w:color
                color = etree.SubElement(rPr, f'{{{w_ns}}}color')
                color.set(f'{{{w_ns}}}val', color_hex)

    _docx_patch(docx_path, transform)


def _apply_text_align(docx_path: Path, alignment: str):
    """Set text alignment via w:jc under w:pPr in shape's txbxContent."""
    from lxml import etree
    w_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    wps_ns = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'

    def transform(doc_xml):
        for txbx in doc_xml.findall(f'.//{{{wps_ns}}}txbx'):
            for pPr in txbx.findall(f'.//{{{w_ns}}}pPr'):
                # Remove existing w:jc
                for old in pPr.findall(f'{{{w_ns}}}jc'):
                    pPr.remove(old)
                # Add w:jc
                jc = etree.SubElement(pPr, f'{{{w_ns}}}jc')
                jc.set(f'{{{w_ns}}}val', alignment)

    _docx_patch(docx_path, transform)


def _apply_text_transform(docx_path: Path, preset: str):
    """Apply text transform (warp) via a:prstTxWarp under wps:bodyPr."""
    from lxml import etree
    a_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    wps_ns = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'

    def transform(doc_xml):
        for bodyPr in doc_xml.findall(f'.//{{{wps_ns}}}bodyPr'):
            # Remove existing prstTxWarp
            for old in bodyPr.findall(f'{{{a_ns}}}prstTxWarp'):
                bodyPr.remove(old)
            warp = etree.SubElement(bodyPr, f'{{{a_ns}}}prstTxWarp')
            warp.set('prst', preset)
            etree.SubElement(warp, f'{{{a_ns}}}avLst')

    _docx_patch(docx_path, transform)


def _apply_text_bevel(docx_path: Path, bevel_preset: str):
    """Apply text bevel via a:sp3d under wps:bodyPr."""
    from lxml import etree
    a_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    wps_ns = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'

    def transform(doc_xml):
        for bodyPr in doc_xml.findall(f'.//{{{wps_ns}}}bodyPr'):
            # Remove existing sp3d
            for old in bodyPr.findall(f'{{{a_ns}}}sp3d'):
                bodyPr.remove(old)
            sp3d = etree.SubElement(bodyPr, f'{{{a_ns}}}sp3d')
            bevelT = etree.SubElement(sp3d, f'{{{a_ns}}}bevelT')
            bevelT.set('prst', bevel_preset)
            bevelT.set('w', '50800')
            bevelT.set('h', '50800')

    _docx_patch(docx_path, transform)


def _apply_picture_shape_geometry(docx_path: Path, geometry: str):
    """Patch the first picture geometry to match Crop to Shape output."""
    from zipfile import ZipFile, ZIP_DEFLATED
    from lxml import etree

    a_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    pic_ns = 'http://schemas.openxmlformats.org/drawingml/2006/picture'

    with ZipFile(docx_path, 'r') as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    doc_xml = etree.fromstring(files['word/document.xml'])
    sp_pr = doc_xml.find(f'.//{{{pic_ns}}}spPr')
    if sp_pr is None:
        return

    prst_geom = sp_pr.find(f'{{{a_ns}}}prstGeom')
    if prst_geom is None:
        prst_geom = etree.SubElement(sp_pr, f'{{{a_ns}}}prstGeom')
        etree.SubElement(prst_geom, f'{{{a_ns}}}avLst')
    prst_geom.set('prst', geometry)
    if prst_geom.find(f'{{{a_ns}}}avLst') is None:
        etree.SubElement(prst_geom, f'{{{a_ns}}}avLst')

    files['word/document.xml'] = etree.tostring(
        doc_xml, xml_declaration=True, encoding='UTF-8', standalone='yes'
    )

    write_docx_package(files, docx_path)


def _apply_picture_rotation_scene(docx_path: Path, camera_preset: str, effect_extent: Dict[str, str]):
    """Patch the first inline picture with scene3d camera/lightRig and effectExtent."""
    from zipfile import ZipFile, ZIP_DEFLATED
    from lxml import etree

    a_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    pic_ns = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
    wp_ns = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'

    with ZipFile(docx_path, 'r') as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    doc_xml = etree.fromstring(files['word/document.xml'])

    inline = doc_xml.find(f'.//{{{wp_ns}}}inline')
    if inline is not None:
        effect = inline.find(f'{{{wp_ns}}}effectExtent')
        if effect is None:
            extent = inline.find(f'{{{wp_ns}}}extent')
            effect = etree.Element(f'{{{wp_ns}}}effectExtent')
            effect.set('l', '0')
            effect.set('t', '0')
            effect.set('r', '0')
            effect.set('b', '0')
            if extent is not None:
                extent_idx = list(inline).index(extent)
                inline.insert(extent_idx + 1, effect)
            else:
                inline.insert(0, effect)
        for key, value in effect_extent.items():
            effect.set(key, value)

    sp_pr = doc_xml.find(f'.//{{{pic_ns}}}spPr')
    if sp_pr is None:
        return

    scene3d = sp_pr.find(f'{{{a_ns}}}scene3d')
    if scene3d is None:
        scene3d = etree.SubElement(sp_pr, f'{{{a_ns}}}scene3d')
    else:
        for child in list(scene3d):
            scene3d.remove(child)

    camera = etree.SubElement(scene3d, f'{{{a_ns}}}camera')
    camera.set('prst', camera_preset)
    light = etree.SubElement(scene3d, f'{{{a_ns}}}lightRig')
    light.set('rig', 'threePt')
    light.set('dir', 't')

    files['word/document.xml'] = etree.tostring(
        doc_xml, xml_declaration=True, encoding='UTF-8', standalone='yes'
    )

    write_docx_package(files, docx_path)


def _apply_picture_bevel(docx_path: Path, bevel_preset: str, width_emu: str = '50800'):
    """Patch the first picture with a top bevel effect matching Word's saved OOXML."""
    from zipfile import ZipFile, ZIP_DEFLATED
    from lxml import etree

    a_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    pic_ns = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
    wp_ns = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'

    with ZipFile(docx_path, 'r') as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    doc_xml = etree.fromstring(files['word/document.xml'])

    inline = doc_xml.find(f'.//{{{wp_ns}}}inline')
    if inline is not None:
        effect = inline.find(f'{{{wp_ns}}}effectExtent')
        if effect is None:
            extent = inline.find(f'{{{wp_ns}}}extent')
            effect = etree.Element(f'{{{wp_ns}}}effectExtent')
            if extent is not None:
                extent_idx = list(inline).index(extent)
                inline.insert(extent_idx + 1, effect)
            else:
                inline.insert(0, effect)
        effect.set('l', '38100')
        effect.set('t', '38100')
        effect.set('r', '48260')
        effect.set('b', '53975')

    sp_pr = doc_xml.find(f'.//{{{pic_ns}}}spPr')
    if sp_pr is None:
        return

    scene3d = sp_pr.find(f'{{{a_ns}}}scene3d')
    if scene3d is None:
        scene3d = etree.SubElement(sp_pr, f'{{{a_ns}}}scene3d')
    else:
        for child in list(scene3d):
            scene3d.remove(child)

    camera = etree.SubElement(scene3d, f'{{{a_ns}}}camera')
    camera.set('prst', 'orthographicFront')

    light_rig = etree.SubElement(scene3d, f'{{{a_ns}}}lightRig')
    light_rig.set('rig', 'threePt')
    light_rig.set('dir', 't')

    sp3d = sp_pr.find(f'{{{a_ns}}}sp3d')
    if sp3d is None:
        sp3d = etree.SubElement(sp_pr, f'{{{a_ns}}}sp3d')
    else:
        for child in list(sp3d):
            sp3d.remove(child)

    bevel_t = etree.SubElement(sp3d, f'{{{a_ns}}}bevelT')
    bevel_t.set('w', str(width_emu))
    bevel_t.set('prst', bevel_preset)

    files['word/document.xml'] = etree.tostring(
        doc_xml, xml_declaration=True, encoding='UTF-8', standalone='yes'
    )

    write_docx_package(files, docx_path)


def _apply_picture_effects_preset(docx_path: Path, preset_index: int):
    """Patch the first picture with baked preset payload values extracted from verified samples."""
    from zipfile import ZipFile, ZIP_DEFLATED
    from lxml import etree

    preset_map = {
        4: {
            'effect_extent': {'l': '57150', 't': '57150', 'r': '353060', 'b': '339725'},
            'line': {'noFill': True},
            'outer_shadow': {
                'blurRad': '190500',
                'dist': '228600',
                'dir': '2700000',
                'algn': 'ctr',
                'color': '000000',
                'alpha': '30000',
            },
            'camera_prst': 'orthographicFront',
            'camera_rot': {'lat': '0', 'lon': '0', 'rev': '0'},
            'light_rig': 'glow',
            'light_dir': 't',
            'light_rot': {'lat': '0', 'lon': '0', 'rev': '4800000'},
            'sp3d': {'prstMaterial': 'matte'},
            'bevelT': {'w': '127000', 'h': '63500'},
        },
        10: {
            'effect_extent': {'l': '552450', 't': '0', 'r': '543560', 'b': '282575'},
            'line': {'w': '34925', 'color': 'FFFFFF'},
            'outer_shadow': {
                'blurRad': '317500',
                'dir': '2700000',
                'algn': 'ctr',
                'color': '000000',
                'alpha': '43000',
            },
            'camera_prst': 'perspectiveFront',
            'camera_extra': {'fov': '2700000'},
            'camera_rot': {'lat': '19086000', 'lon': '19067999', 'rev': '3108000'},
            'light_rig': 'threePt',
            'light_dir': 't',
            'light_rot': {'lat': '0', 'lon': '0', 'rev': '0'},
            'sp3d': {'extrusionH': '38100', 'prstMaterial': 'clear'},
            'bevelT': {'w': '260350', 'h': '50800', 'prst': 'softRound'},
            'bevelB': {'prst': 'softRound'},
        },
    }

    payload = preset_map.get(int(preset_index))
    if payload is None:
        return

    a_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    pic_ns = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
    wp_ns = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'

    with ZipFile(docx_path, 'r') as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    doc_xml = etree.fromstring(files['word/document.xml'])

    inline = doc_xml.find(f'.//{{{wp_ns}}}inline')
    if inline is not None:
        effect = inline.find(f'{{{wp_ns}}}effectExtent')
        if effect is None:
            extent = inline.find(f'{{{wp_ns}}}extent')
            effect = etree.Element(f'{{{wp_ns}}}effectExtent')
            if extent is not None:
                extent_idx = list(inline).index(extent)
                inline.insert(extent_idx + 1, effect)
            else:
                inline.insert(0, effect)
        for key, value in payload['effect_extent'].items():
            effect.set(key, value)

    sp_pr = doc_xml.find(f'.//{{{pic_ns}}}spPr')
    if sp_pr is None:
        return

    line = sp_pr.find(f'{{{a_ns}}}ln')
    if line is not None:
        sp_pr.remove(line)
    if payload.get('line') is not None:
        line = etree.SubElement(sp_pr, f'{{{a_ns}}}ln')
        line_payload = payload['line']
        if line_payload.get('w'):
            line.set('w', line_payload['w'])
        if line_payload.get('noFill'):
            etree.SubElement(line, f'{{{a_ns}}}noFill')
        elif line_payload.get('color'):
            solid = etree.SubElement(line, f'{{{a_ns}}}solidFill')
            srgb = etree.SubElement(solid, f'{{{a_ns}}}srgbClr')
            srgb.set('val', line_payload['color'])

    effect_lst = sp_pr.find(f'{{{a_ns}}}effectLst')
    if effect_lst is not None:
        sp_pr.remove(effect_lst)
    if payload.get('outer_shadow'):
        effect_lst = etree.SubElement(sp_pr, f'{{{a_ns}}}effectLst')
        shadow = etree.SubElement(effect_lst, f'{{{a_ns}}}outerShdw')
        for key, value in payload['outer_shadow'].items():
            if key == 'color':
                continue
            if key == 'alpha':
                continue
            shadow.set(key, value)
        srgb = etree.SubElement(shadow, f'{{{a_ns}}}srgbClr')
        srgb.set('val', payload['outer_shadow'].get('color', '000000'))
        if payload['outer_shadow'].get('alpha'):
            alpha = etree.SubElement(srgb, f'{{{a_ns}}}alpha')
            alpha.set('val', payload['outer_shadow']['alpha'])

    scene3d = sp_pr.find(f'{{{a_ns}}}scene3d')
    if scene3d is None:
        scene3d = etree.SubElement(sp_pr, f'{{{a_ns}}}scene3d')
    else:
        for child in list(scene3d):
            scene3d.remove(child)

    camera = etree.SubElement(scene3d, f'{{{a_ns}}}camera')
    camera.set('prst', payload['camera_prst'])
    for key, value in payload.get('camera_extra', {}).items():
        camera.set(key, value)
    if payload.get('camera_rot'):
        rot = etree.SubElement(camera, f'{{{a_ns}}}rot')
        for key, value in payload['camera_rot'].items():
            rot.set(key, value)

    light_rig = etree.SubElement(scene3d, f'{{{a_ns}}}lightRig')
    light_rig.set('rig', payload['light_rig'])
    light_rig.set('dir', payload['light_dir'])
    if payload.get('light_rot'):
        rot = etree.SubElement(light_rig, f'{{{a_ns}}}rot')
        for key, value in payload['light_rot'].items():
            rot.set(key, value)

    sp3d = sp_pr.find(f'{{{a_ns}}}sp3d')
    if sp3d is None:
        sp3d = etree.SubElement(sp_pr, f'{{{a_ns}}}sp3d')
    else:
        for key in list(sp3d.attrib.keys()):
            del sp3d.attrib[key]
        for child in list(sp3d):
            sp3d.remove(child)
    for key, value in payload['sp3d'].items():
        sp3d.set(key, value)
    if payload.get('bevelT'):
        bevel_t = etree.SubElement(sp3d, f'{{{a_ns}}}bevelT')
        for key, value in payload['bevelT'].items():
            bevel_t.set(key, value)
    if payload.get('bevelB'):
        bevel_b = etree.SubElement(sp3d, f'{{{a_ns}}}bevelB')
        for key, value in payload['bevelB'].items():
            bevel_b.set(key, value)

    files['word/document.xml'] = etree.tostring(
        doc_xml, xml_declaration=True, encoding='UTF-8', standalone='yes'
    )

    write_docx_package(files, docx_path)


def _apply_picture_layout_gallery_item(wd, obj, gallery_item_name: str, timeout_s: float = 12.0) -> bool:
    """Apply a Picture Layout gallery item through Word's ribbon UI."""
    try:
        import ctypes
        import win32api
        import win32con
        import win32gui
        from pywinauto.application import Application
    except ImportError as e:
        print(f"    [WARN] PicturesConvertToSmartArt requires pywinauto/pywin32: {e}")
        return False

    popup_classes = {
        'Net UI Tool Window',
        'NetUIHWND',
        'MsoCommandBarPopup',
        'NUIDialog',
    }

    def _enum_popup_hwnds() -> List[int]:
        found: List[int] = []

        def _add(hwnd: int):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return
                if win32gui.GetClassName(hwnd) in popup_classes and hwnd not in found:
                    found.append(hwnd)
            except Exception:
                return

        def _child_cb(hwnd, _):
            _add(hwnd)

        def _top_cb(hwnd, _):
            _add(hwnd)
            win32gui.EnumChildWindows(hwnd, _child_cb, None)

        win32gui.EnumWindows(_top_cb, None)
        return found

    def _wait_for_new_popup(before: List[int], popup_timeout_s: float = 5.0) -> Optional[int]:
        before_set = set(before)
        deadline = time.monotonic() + popup_timeout_s
        while time.monotonic() < deadline:
            time.sleep(0.05)
            after = _enum_popup_hwnds()
            new = [hwnd for hwnd in after if hwnd not in before_set]
            if new:
                return new[-1]
        return None

    def _expand_control(ctrl) -> bool:
        try:
            ctrl.expand()
            return True
        except Exception:
            pass

        try:
            wrapper = ctrl.wrapper_object()
            children = wrapper.children()
            button_children = [child for child in children if child.friendly_class_name() == 'Button']
            if button_children:
                button_children[-1].click_input()
                return True
        except Exception:
            pass

        try:
            rect = ctrl.rectangle()
            ctrl.click_input(coords=(rect.width() - 4, rect.height() - 4))
            return True
        except Exception:
            return False

    def _close_popup():
        try:
            win32api.keybd_event(0x1B, 0, 0, 0)
            win32api.keybd_event(0x1B, 0, win32con.KEYEVENTF_KEYUP, 0)
        except Exception:
            pass

    def _find_gallery_item(dlg, popup_bounds, item_text: str):
        item_text_l = item_text.casefold()
        best = None
        best_area = -1
        for candidate in dlg.descendants(control_type='ListItem'):
            try:
                text = (candidate.window_text() or '').strip()
                if text.casefold() != item_text_l:
                    continue
                rect = candidate.rectangle()
                if (
                    rect.left < popup_bounds[0] or rect.right > popup_bounds[2]
                    or rect.top < popup_bounds[1] or rect.bottom > popup_bounds[3]
                ):
                    continue
                area = max(0, rect.width() * rect.height())
                if area > best_area:
                    best = candidate
                    best_area = area
            except Exception:
                continue
        return best

    try:
        try:
            obj.Select()
        except Exception:
            try:
                obj.Range.Select()
            except Exception:
                return False

        word_hwnd = int(wd.ActiveWindow.Hwnd)
        try:
            user32 = ctypes.windll.user32
            user32.ShowWindow(word_hwnd, 5)
            user32.SetForegroundWindow(word_hwnd)
            wd.Activate()
        except Exception:
            pass
        time.sleep(0.6)

        app = Application(backend='uia').connect(handle=word_hwnd)
        dlg = app.window(handle=word_hwnd)

        picture_format_tab = dlg.child_window(title='Picture Format', control_type='TabItem')
        picture_format_tab.wait('exists enabled', timeout=5.0)
        picture_format_tab.wrapper_object().select()
        time.sleep(0.8)

        gallery_ctrl = dlg.child_window(auto_id='PicturesConvertToSmartArt')
        gallery_ctrl.wait('exists enabled', timeout=5.0)

        before = _enum_popup_hwnds()
        if not _expand_control(gallery_ctrl):
            return False

        popup_hwnd = _wait_for_new_popup(before)
        if popup_hwnd is None:
            _close_popup()
            return False

        time.sleep(0.8)
        popup_bounds = win32gui.GetWindowRect(popup_hwnd)
        gallery_item = _find_gallery_item(dlg, popup_bounds, gallery_item_name)
        if gallery_item is None:
            visible_items = []
            for candidate in dlg.descendants(control_type='ListItem'):
                try:
                    text = (candidate.window_text() or '').strip()
                    rect = candidate.rectangle()
                    if (
                        text
                        and popup_bounds[0] <= rect.left
                        and rect.right <= popup_bounds[2]
                        and popup_bounds[1] <= rect.top
                        and rect.bottom <= popup_bounds[3]
                    ):
                        visible_items.append(text)
                except Exception:
                    continue
            print(f"    [WARN] Gallery item '{gallery_item_name}' not found. Visible items: {visible_items}")
            _close_popup()
            return False

        rect = gallery_item.rectangle()
        click_x = (rect.left + rect.right) // 2
        click_y = (rect.top + rect.bottom) // 2
        win32api.SetCursorPos((click_x, click_y))
        time.sleep(0.15)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

        deadline = time.monotonic() + timeout_s
        while time.monotonic() < deadline:
            time.sleep(0.25)
            try:
                doc = wd.ActiveDocument
                if doc.InlineShapes.Count == 0:
                    continue
                inline_shape = doc.InlineShapes(1)
                if not bool(getattr(inline_shape, 'HasSmartArt', False)):
                    continue
                layout_name = (inline_shape.SmartArt.Layout.Name or '').strip()
                if layout_name.casefold() == gallery_item_name.casefold():
                    return True
            except Exception:
                continue
        return False
    finally:
        time.sleep(0.1)


def _ole_rgb_channels(value: int) -> tuple[int, int, int]:
    """Return the three color channels from an Office RGB/OLE color value."""
    color = int(value) & 0xFFFFFF
    return color & 0xFF, (color >> 8) & 0xFF, (color >> 16) & 0xFF


def _gradient_stop_state(fill) -> List[Dict]:
    """Read effective, theme-resolved gradient stops from a FillFormat."""
    rows = []
    stops = fill.GradientStops
    for index in range(1, stops.Count + 1):
        stop = stops.Item(index)
        rows.append({
            'index': index,
            'rgb': int(stop.Color.RGB),
            'position': round(float(stop.Position), 4),
            'transparency': round(float(stop.Transparency), 4),
        })
    return rows


def _has_visible_gradient_progression(stops: List[Dict]) -> bool:
    """Return whether one fill has a real, visible stop-to-stop progression.

    This is deliberately property-scoped: colors are compared only with other
    stops in the same fill.  It neither requires exact RGB values nor a fixed
    stop count across Office themes.
    """
    usable = [
        row for row in stops
        if row.get('rgb') is not None and row.get('position') is not None
    ]
    if len(usable) < 2:
        return False
    if max(row['position'] for row in usable) - min(row['position'] for row in usable) < 0.5:
        return False

    for left_index, left in enumerate(usable):
        left_channels = _ole_rgb_channels(left['rgb'])
        for right in usable[left_index + 1:]:
            right_channels = _ole_rgb_channels(right['rgb'])
            deltas = [
                abs(a - b) for a, b in zip(left_channels, right_channels)
            ]
            if max(deltas) >= 12 and sum(deltas) >= 24:
                return True
    return False


def _validate_gradient_postconditions(fill, param: Dict) -> Dict:
    """Enforce the named GradientGallery preset and visible progression."""
    expected = {
        'type': 3,  # msoFillGradient
        'style': int(param['ocg_style']),
        'variant': int(param.get('ocg_variant', 1)),
        'degree': float(param.get('ocg_degree', 0.5)),
        'angle': float(param.get('ocg_angle', 0.0)),
    }
    actual = {
        'type': int(fill.Type),
        'style': int(fill.GradientStyle),
        'variant': int(fill.GradientVariant),
        'degree': float(fill.GradientDegree),
        'angle': float(fill.GradientAngle),
        'stops': _gradient_stop_state(fill),
    }
    actual['visible_progression'] = _has_visible_gradient_progression(
        actual['stops']
    )

    errors = []
    for key in ('type', 'style', 'variant'):
        if actual[key] != expected[key]:
            errors.append(f'{key}={actual[key]!r}, expected {expected[key]!r}')
    if abs(actual['degree'] - expected['degree']) > 0.01:
        errors.append(
            f"degree={actual['degree']!r}, expected {expected['degree']!r}"
        )
    if abs(actual['angle'] - expected['angle']) > 1.0:
        errors.append(
            f"angle={actual['angle']!r}, expected {expected['angle']!r}"
        )
    if param.get('requires_visible_progression') and not actual['visible_progression']:
        errors.append('gradient stops have no visible effective progression')

    if errors:
        name = param.get('gradient_name') or param.get('label') or 'gradient'
        raise ValueError(f'{name} postcondition failed: ' + '; '.join(errors))
    return actual


def _assert_saved_gradient_postconditions(wd, gold_path: Path, param: Dict) -> Dict:
    """Normal-reopen a saved gradient reference and validate effective state."""
    reopened = None
    try:
        reopened = wd.Documents.Open(
            str(gold_path.absolute()), ReadOnly=True, AddToRecentFiles=False
        )
        if reopened.Shapes.Count != 1:
            raise ValueError(
                f'gradient reference must contain one shape, found {reopened.Shapes.Count}'
            )
        return _validate_gradient_postconditions(reopened.Shapes(1).Fill, param)
    finally:
        if reopened is not None:
            reopened.Close(False)


def apply_parameter(wd, idMso: str, param: Dict, obj=None) -> bool:
    """
    Apply parameter to selected object.
    
    Args:
        wd: Word application
        idMso: Command identifier
        param: Parameter dictionary
        obj: Object to apply parameter to (table or picture)
    
    Returns:
        True if successful, False otherwise
    """
    try:
        # =====================================================================
        # TABLE COMMANDS (PM List Scope)
        # =====================================================================
        
        # Table style - Apply table style from gallery
        if idMso == 'TableStylesGalleryWord':
            if obj is None:
                return False
            style_name = param.get('style_name')
            if style_name:
                obj.Style = style_name
            return True
        
        # Outline weight - Set line weight on SmartArt "Target" node shape
        elif idMso == 'OutlineWeightGallery':
            if obj is None:
                return False
            target_node = None
            for i in range(1, obj.Nodes.Count + 1):
                if obj.Nodes(i).TextFrame2.TextRange.Text.strip() == 'Target':
                    target_node = obj.Nodes(i)
                    break
            if target_node is None:
                target_node = obj.Nodes(1)
            weight = param.get('weight')
            if weight is not None:
                s = target_node.Shapes(1)  # Main shape
                s.Select()
                time.sleep(0.3)
                s.Line.Visible = True
                s.Line.Weight = float(weight)
            return True
        
        # Border styles - Apply border style
        elif idMso == 'BorderStylesGallery':
            if obj is None:
                return False
            line_style = param.get('line_style')
            if line_style is not None:
                # Apply border line style to all borders
                for border_type in range(1, 5):
                    try:
                        obj.Borders(border_type).LineStyle = line_style
                    except:
                        pass
            return True
        
        # Text direction - Set text direction in a textbox
        elif idMso == 'TextDirectionGallery':
            if obj is None:
                return False
            orientation = param.get('orientation', param.get('direction'))
            if orientation is not None:
                try:
                    obj.TextFrame.Orientation = int(orientation)
                except:
                    try:
                        obj.TextFrame.TextRange.Orientation = int(orientation)
                    except:
                        pass
            return True

        # Align Text - Set vertical alignment inside a textbox
        elif idMso == 'TextAlignGallery':
            if obj is None:
                return False
            vertical_anchor = param.get('vertical_anchor')
            if vertical_anchor is not None:
                # MsoVerticalAnchor: 1=top, 3=middle, 4=bottom
                try:
                    obj.TextFrame.VerticalAnchor = int(vertical_anchor)
                except:
                    pass
            return True

        # =====================================================================
        # PICTURE COMMANDS (PM List Scope)
        # =====================================================================

        elif idMso == 'PicturesConvertToSmartArt':
            if obj is None:
                return False
            gallery_item_name = param.get('gallery_item_name')
            if not gallery_item_name:
                return False
            return _apply_picture_layout_gallery_item(wd, obj, gallery_item_name)

        # Picture saturation - via Fill.PictureEffects (type 24)
        elif idMso == 'PictureSaturationGallery':
            sat_pct = param.get('saturation_pct')
            if sat_pct is not None:
                shape = obj.ConvertToShape()
                effect = shape.Fill.PictureEffects.Insert(24)
                effect.EffectParameters.Item(1).Value = sat_pct / 100.0
                shape.ConvertToInlineShape()
            return True

        # Picture color temperature - via Fill.PictureEffects (type 7)
        elif idMso == 'PictureColorTempertatureGallery':
            color_temp = param.get('color_temp')
            if color_temp is not None:
                shape = obj.ConvertToShape()
                effect = shape.Fill.PictureEffects.Insert(7)
                effect.EffectParameters.Item(1).Value = color_temp
                shape.ConvertToInlineShape()
            return True

        # Picture sharpen/soften - via Fill.PictureEffects (type 25)
        elif idMso == 'PictureSharpenSoftenGallery':
            amount = param.get('amount')
            if amount is not None:
                shape = obj.ConvertToShape()
                effect = shape.Fill.PictureEffects.Insert(25)
                # amount is in EMU (-100000 to 100000), COM expects -1.0 to 1.0
                effect.EffectParameters.Item(1).Value = amount / 100000.0
                shape.ConvertToInlineShape()
            return True

        # Picture transparency - Set image transparency
        elif idMso == 'PictureTransparencyGallery':
            if obj is None:
                return False
            return True

        # Picture recolor - COM PictureFormat.ColorType (2=grayscale, 3=B&W, 4=washout)
        elif idMso == 'PictureRecolorGallery':
            if obj is None:
                return False
            color_type = param.get('color_type')
            if color_type is not None:
                obj.PictureFormat.ColorType = int(color_type)
            return True

        elif idMso == 'PictureArtisticEffectsGallery':
            if obj is None:
                return False
            effect_type = param.get('effect_type')
            if effect_type is not None:
                shape = obj.ConvertToShape()
                shape.Fill.PictureEffects.Insert(int(effect_type))
                shape.ConvertToInlineShape()
            return True

        # Picture styles - OOXML post-processing via _apply_picture_style
        elif idMso == 'PictureStylesGallery':
            if obj is None:
                return False
            # Actual style is applied via post-processing after save
            return True

        # Picture shadow - Apply via Shadow.Type preset (21-43)
        elif idMso == 'PictureEffectsShadowGallery':
            if obj is None:
                return False
            shadow_type = param.get('shadow_type')
            if shadow_type is not None:
                obj.Shadow.Visible = True
                obj.Shadow.Type = int(shadow_type)
            return True

        # Picture glow - Apply glow effect
        elif idMso == 'PictureEffectsGlowGallery':
            if obj is None:
                return False
            glow = obj.Glow
            radius = param.get('glow_radius')
            if radius is not None:
                glow.Radius = radius
            theme_color = param.get('glow_theme_color')
            if theme_color is not None:
                glow.Color.ObjectThemeColor = theme_color
            else:
                color_rgb = param.get('glow_color_rgb')
                if color_rgb is not None:
                    glow.Color.RGB = color_rgb
            glow.Transparency = 0.0
            return True

        # Picture soft edges - Apply soft edges effect
        elif idMso == 'PictureEffectsSoftEdgesGallery':
            if obj is None:
                return False
            radius = param.get('soft_edge_radius')
            if radius is not None:
                obj.SoftEdge.Radius = radius
            return True

        # Picture reflection - Apply reflection effect
        elif idMso == 'PictureReflectionGallery':
            if obj is None:
                return False
            reflection_type = param.get('reflection_type')
            if reflection_type is not None:
                obj.Reflection.Type = reflection_type
            return True

        # Picture 3D rotation - Apply 3D rotation via ThreeD.SetPresetCamera
        elif idMso == 'PictureRotationGallery':
            preset_camera = param.get('preset_camera')
            if preset_camera is not None:
                shape = obj.ConvertToShape()
                shape.ThreeD.SetPresetCamera(int(preset_camera))
                shape.ConvertToInlineShape()
            return True

        # Picture bevel - Apply top bevel via ThreeD.BevelTopType
        elif idMso == '_3DBevelPictureTopGallery':
            bevel_preset = param.get('bevel_top_preset')
            if bevel_preset is not None:
                bevel_map = {
                    'softRound': 7,
                    'convex': 8,
                    'angle': 6,
                    'artDeco': 13,
                    'circle': 3,
                    'slope': 4,
                    'cross': 5,
                    'coolSlant': 9,
                    'divot': 10,
                    'riblet': 11,
                    'hardEdge': 12,
                }
                bevel_type = bevel_map.get(bevel_preset, 3)
                width_emu = param.get('bevel_top_width', '50800')
                depth_pt = int(width_emu) / 12700.0  # EMU to points
                inset_emu = param.get('bevel_top_inset', width_emu)
                inset_pt = int(inset_emu) / 12700.0
                shape = obj.ConvertToShape()
                shape.ThreeD.BevelTopType = bevel_type
                shape.ThreeD.BevelTopDepth = depth_pt
                shape.ThreeD.BevelTopInset = inset_pt
                shape.ConvertToInlineShape()
            return True

        # Picture shape - Change picture shape via AutoShapeType
        elif idMso == 'PictureShapeGallery':
            auto_shape_type = param.get('auto_shape_type')
            if auto_shape_type is not None:
                shape = obj.ConvertToShape()
                shape.AutoShapeType = int(auto_shape_type)
                shape.ConvertToInlineShape()
            return True

        # Picture effects preset - Apply combined preset effect
        elif idMso == 'PictureEffectsPresetGallery':
            if obj is None:
                return False
            return True

        # =====================================================================
        # SMARTART TEXT COMMANDS (SmartArtToolsFormat > WordArtStyles)
        # =====================================================================

        elif idMso in SMARTART_TEXT_COMMANDS:
            if obj is None:
                return False
            # Find "Target" node in SmartArt
            target_node = None
            for i in range(1, obj.Nodes.Count + 1):
                if obj.Nodes(i).TextFrame2.TextRange.Text.strip() == 'Target':
                    target_node = obj.Nodes(i)
                    break
            if target_node is None:
                target_node = obj.Nodes(1)

            tf2 = target_node.TextFrame2

            if idMso == 'TextFillColorPicker':
                color_hex = param.get('color_hex')
                if color_hex:
                    r = int(color_hex[0:2], 16)
                    g = int(color_hex[2:4], 16)
                    b = int(color_hex[4:6], 16)
                    bgr = (b << 16) | (g << 8) | r
                    tf2.TextRange.Font.Fill.ForeColor.RGB = bgr

            elif idMso == 'TextEffectTransformGallery':
                warp_format = param.get('warp_format')
                if warp_format is not None:
                    tf2.WarpFormat = int(warp_format)

            elif idMso == 'BevelTextGallery':
                bevel_type = param.get('bevel_type')
                if bevel_type is not None:
                    td = tf2.ThreeD
                    td.BevelTopType = int(bevel_type)
                    td.BevelTopDepth = 4.0
                    td.BevelTopInset = 4.0

            elif idMso == 'TextEffects3DRotationGallery':
                preset_camera = param.get('preset_camera')
                if preset_camera is not None:
                    tf2.ThreeD.SetPresetCamera(int(preset_camera))

            return True

        # =====================================================================
        # SMARTART COMMANDS (obj is SmartArt object)
        # =====================================================================

        elif idMso == 'SmartArtLayoutGallery':
            if obj is None: return False
            layout_index = param.get('layout_index')
            if layout_index is not None:
                layouts = wd.Application.SmartArtLayouts
                obj.Layout = layouts(int(layout_index))
            return True

        elif idMso == 'SmartArtStylesGallery':
            if obj is None: return False
            style_index = param.get('style_index')
            if style_index is not None:
                styles = wd.Application.SmartArtQuickStyles
                obj.QuickStyle = styles(int(style_index))
            return True

        elif idMso == 'SmartArtChangeColorsGallery':
            if obj is None: return False
            color_index = param.get('color_index')
            if color_index is not None:
                colors = wd.Application.SmartArtColors
                obj.Color = colors(int(color_index))
            return True

        # =====================================================================
        # CHART COMMANDS (obj is Chart object)
        # =====================================================================

        elif idMso == 'ChartStylesGallery':
            if obj is None: return False
            style = param.get('chart_style')
            if style is not None:
                obj.ChartStyle = int(style)
            return True

        elif idMso == 'ChartColorsGallery':
            if obj is None: return False
            color = param.get('chart_color')
            if color is not None:
                obj.ChartColor = int(color)
            return True

        elif idMso == 'ChartLayoutGallery':
            if obj is None: return False
            layout = param.get('chart_layout')
            if layout is not None:
                obj.ApplyLayout(int(layout))
            return True

        elif idMso == 'ChartLegend':
            if obj is None: return False
            legend_position = param.get('legend_position')
            if legend_position is not None:
                obj.HasLegend = True
                obj.Legend.Position = int(legend_position)
            else:
                show = param.get('show_legend')
                if show is not None:
                    obj.HasLegend = bool(show)
            return True

        elif idMso == 'ChartDataLabel':
            if obj is None: return False
            position = param.get('label_position')
            if position is not None:
                for i in range(1, obj.SeriesCollection().Count + 1):
                    obj.SeriesCollection(i).HasDataLabels = True
                    obj.SeriesCollection(i).DataLabels().Position = int(position)
            else:
                show = param.get('show_data_labels', True)
                for i in range(1, obj.SeriesCollection().Count + 1):
                    obj.SeriesCollection(i).HasDataLabels = bool(show)
            return True

        elif idMso == 'ChartDataTable':
            if obj is None: return False
            show_legend_key = param.get('show_legend_key')
            if show_legend_key is not None:
                obj.HasDataTable = True
                obj.DataTable.ShowLegendKey = bool(show_legend_key)
            else:
                show = param.get('show_data_table', True)
                obj.HasDataTable = bool(show)
            return True

        elif idMso == 'ChartTrendline':
            if obj is None: return False
            trend_type = param.get('trend_type', -4132)  # xlLinear
            s = obj.SeriesCollection(1)
            # Clear existing trendlines
            while s.Trendlines().Count > 0:
                s.Trendlines(1).Delete()
            s.Trendlines().Add(Type=int(trend_type))
            return True

        elif idMso == 'ChartErrorBars':
            if obj is None: return False
            error_bar_type = param.get('error_bar_type')
            s = obj.SeriesCollection(1)
            if error_bar_type is not None:
                amount = param.get('error_amount', 5 if error_bar_type == 2 else 1)
                s.ErrorBar(1, 1, int(error_bar_type), amount)
                time.sleep(0.5)
            else:
                s.HasErrorBars = False
            return True

        elif idMso == 'ChartLines':
            if obj is None: return False
            line_type = param.get('line_type', 'drop')
            cg = obj.ChartGroups(1)
            if line_type == 'hilo':
                cg.HasHiLoLines = True
            elif line_type == 'drop':
                cg.HasDropLines = True
            return True

        elif idMso == 'ChartUpDownBars':
            if obj is None: return False
            obj.ChartGroups(1).HasUpDownBars = True
            return True

        elif idMso == 'ChartAxesGallery':
            if obj is None: return False
            action = param.get('axis_action')
            if action == 'show_all_titles':
                for ax_type in [1, 2]:  # xlCategory, xlValue
                    ax = obj.Axes(ax_type)
                    ax.HasTitle = True
                    ax.AxisTitle.Text = ['Category', 'Values'][ax_type - 1]
            elif action == 'hide_horizontal':
                obj.Axes(1).Delete()
            return True

        elif idMso == 'ColorPickerChartLineStyles':
            if obj is None: return False
            color_bgr = param.get('color_bgr')
            if color_bgr is not None:
                s = obj.SeriesCollection(1)
                s.Format.Line.Visible = True
                s.Format.Line.ForeColor.RGB = int(color_bgr)
            return True

        elif idMso == 'ChartTitle':
            if obj is None: return False
            position = param.get('title_position', 'above')
            obj.HasTitle = True
            obj.ChartTitle.Text = 'Chart Title'
            if position == 'overlay':
                obj.ChartTitle.IncludeInLayout = False  # Centered overlay
            else:
                obj.ChartTitle.IncludeInLayout = True   # Above chart
            return True

        elif idMso == 'ChartGridlinesGallery':
            if obj is None: return False
            axis_type = param.get('axis_type', 1)  # 1=xlCategory, 2=xlValue
            gridline_type = param.get('gridline_type', 'major')  # major or minor
            ax = obj.Axes(int(axis_type))
            if gridline_type == 'major':
                ax.HasMajorGridlines = True
            elif gridline_type == 'minor':
                ax.HasMinorGridlines = True
            return True

        elif idMso == 'RecolorColorPicker':
            # Duotone recolor — COM cannot apply this; handled via OOXML post-processing
            return True

        # =====================================================================
        # EQUATION COMMANDS
        # =====================================================================

        elif idMso in EQUATION_COMMANDS:
            if obj is None:
                return False
            # obj is the OMath object
            return _apply_equation_structure(wd, obj, param)

        # =====================================================================
        # GRAPHIC/SVG COMMANDS
        # =====================================================================

        elif idMso == 'GraphicsColorPicker':
            if obj is None: return False
            color_bgr = param.get('color_bgr')
            if color_bgr is not None:
                obj.Fill.Solid()
                obj.Fill.ForeColor.RGB = int(color_bgr)
            return True

        elif idMso == 'GraphicsOutlineColorPicker':
            if obj is None: return False
            color_bgr = param.get('color_bgr')
            if color_bgr is not None:
                obj.Line.Visible = True
                obj.Line.Weight = 2.0
                obj.Line.ForeColor.RGB = int(color_bgr)
            return True


        # =====================================================================
        # SHAPE COMMANDS
        # =====================================================================

        elif idMso == 'ShapeStylesGallery':
            if obj is None:
                return False
            style_index = param.get('style_index')
            if style_index is not None:
                obj.ShapeStyle = int(style_index)
            return True

        elif idMso == 'ShapeFillColorPicker':
            if obj is None:
                return False
            color_bgr = param.get('color_bgr')
            if color_bgr is not None:
                obj.Fill.Solid()
                obj.Fill.ForeColor.RGB = int(color_bgr)
            return True

        elif idMso in ('ShapeOutlineColorPicker', 'OutlineColorPicker'):
            if obj is None:
                return False
            color_bgr = param.get('color_bgr')
            if color_bgr is not None:
                obj.Line.Visible = True
                obj.Line.ForeColor.RGB = int(color_bgr)
            return True

        elif idMso == 'OutlineDashesGallery':
            if obj is None:
                return False
            # Apply to SmartArt "Target" node shape
            target_node = None
            for i in range(1, obj.Nodes.Count + 1):
                if obj.Nodes(i).TextFrame2.TextRange.Text.strip() == 'Target':
                    target_node = obj.Nodes(i)
                    break
            if target_node is None:
                target_node = obj.Nodes(1)
            dash_style = param.get('dash_style')
            if dash_style is not None:
                s = target_node.Shapes(1)  # Main shape
                s.Select()
                time.sleep(0.3)
                s.Line.Visible = True
                s.Line.DashStyle = int(dash_style)
            return True

        elif idMso == 'ArrowStyleGallery':
            if obj is None:
                return False
            # Find "Target" node's connector shape (Type=9), select then apply
            target_node = None
            for i in range(1, obj.Nodes.Count + 1):
                if obj.Nodes(i).TextFrame2.TextRange.Text.strip() == 'Target':
                    target_node = obj.Nodes(i)
                    break
            if target_node is None:
                target_node = obj.Nodes(1)
            # Shape 3 is the connector (Type=9) in Tab List
            for j in range(1, target_node.Shapes.Count + 1):
                s = target_node.Shapes(j)
                if s.Type == 9:  # msoGroup / connector
                    s.Select()
                    time.sleep(0.3)
                    begin_style = param.get('begin_arrowhead_style', 0)
                    end_style = param.get('end_arrowhead_style', 0)
                    s.Line.Visible = True
                    s.Line.BeginArrowheadStyle = int(begin_style)
                    s.Line.EndArrowheadStyle = int(end_style)
                    break
            return True

        elif idMso == 'GradientGallery':
            if obj is None:
                return False
            ocg_style = param.get('ocg_style')
            ocg_variant = param.get('ocg_variant', 1)
            ocg_degree = param.get('ocg_degree', 0.5)
            if ocg_style is not None:
                obj.Fill.OneColorGradient(int(ocg_style), int(ocg_variant), float(ocg_degree))
                stop_brightness = param.get('ocg_stop_brightness')
                if stop_brightness:
                    stops = obj.Fill.GradientStops
                    if stops.Count < 2:
                        raise ValueError('OneColorGradient did not create two gradient stops')
                    stops.Item(1).Color.Brightness = float(stop_brightness[0])
                    stops.Item(stops.Count).Color.Brightness = float(stop_brightness[-1])
                _validate_gradient_postconditions(obj.Fill, param)
            return True

        elif idMso == 'ShapeFillTextureGallery':
            if obj is None:
                return False
            texture_index = param.get('texture_index')
            if texture_index is not None:
                obj.Fill.PresetTextured(int(texture_index))
            return True

        elif idMso == 'ShapeChangeShapeGallery':
            if obj is None:
                return False
            auto_shape_type = param.get('auto_shape_type')
            if auto_shape_type is not None:
                obj.AutoShapeType = int(auto_shape_type)
            return True

        elif idMso == 'ObjectEffectShadowGallery':
            if obj is None:
                return False
            shadow_type = param.get('shadow_type')
            if shadow_type is not None:
                obj.Shadow.Type = int(shadow_type)
            return True

        elif idMso == 'ObjectEffectGlowGallery':
            if obj is None:
                return False
            radius = param.get('glow_radius')
            color_bgr = param.get('color_bgr', 0x4472C4)
            if radius is not None:
                obj.Glow.Radius = float(radius)
                obj.Glow.Color.RGB = int(color_bgr)
            return True

        elif idMso == 'ObjectEffectSoftEdgesGallery':
            if obj is None:
                return False
            radius = param.get('soft_edge_radius')
            if radius is not None:
                obj.SoftEdge.Radius = float(radius)
            return True

        elif idMso == 'ReflectionGallery':
            if obj is None:
                return False
            reflection_type = param.get('reflection_type')
            if reflection_type is not None:
                obj.Reflection.Type = int(reflection_type)
            return True

        elif idMso == '_3DRotationGallery':
            if obj is None:
                return False
            camera_preset = param.get('preset_camera') or param.get('camera_preset')
            if camera_preset is not None:
                obj.ThreeD.SetPresetCamera(int(camera_preset))
            return True

        elif idMso == 'BevelShapeGallery':
            if obj is None:
                return False
            bevel_type = param.get('bevel_type')
            if bevel_type is not None:
                obj.ThreeD.BevelTopType = int(bevel_type)
                obj.ThreeD.BevelTopDepth = float(param.get('bevel_depth', 6))
                obj.ThreeD.BevelTopInset = float(param.get('bevel_inset', 6))
            return True

        elif idMso == 'GlowColorPicker':
            if obj is None:
                return False
            color_bgr = param.get('color_bgr')
            radius = param.get('glow_radius', 8)
            if color_bgr is not None:
                obj.Glow.Radius = float(radius)
                obj.Glow.Color.RGB = int(color_bgr)
            return True

        elif idMso == 'ObjectEffectPresetGallery':
            if obj is None:
                return False
            # Combined effects: shadow + reflection + soft edges + glow
            shadow_type = param.get('shadow_type')
            if shadow_type is not None:
                obj.Shadow.Visible = True
                obj.Shadow.Type = int(shadow_type)
            reflection_type = param.get('reflection_type')
            if reflection_type is not None:
                obj.Reflection.Type = int(reflection_type)
            soft_edge_radius = param.get('soft_edge_radius')
            if soft_edge_radius is not None:
                obj.SoftEdge.Radius = float(soft_edge_radius)
            glow_radius = param.get('glow_radius')
            if glow_radius is not None:
                obj.Glow.Radius = float(glow_radius)
            return True

        elif idMso == 'GallerySketchyShapes':
            if obj is None:
                return False
            # Sketchy line effects use Line.Style — COM sets the base,
            # but the actual "sketchy" rendering is an Office 365 feature.
            # For now, apply the line dash style as a fallback.
            line_style = param.get('line_style', 1)
            obj.Line.Visible = True
            obj.Line.Style = int(line_style)
            return True

        elif idMso == 'ShapeStylesOtherThemeFillsGallery':
            if obj is None:
                return False
            theme_color = param.get('theme_color')
            if theme_color is not None:
                obj.Fill.Solid()
                obj.Fill.ForeColor.ObjectThemeColor = int(theme_color)
            return True

        else:
            print(f"    [SKIP] No apply method for {idMso}")
            return False

    except Exception as e:
        print(f"    [ERROR] Failed to apply parameter: {e}")
        import traceback
        traceback.print_exc()
        return False


class ObjectEffectContextualParameterizedGenerator:
    """Generator for object_effect × contextual_parameterized test cases."""

    def __init__(self, output_dir: Path = None):
        self.execution = DirectExecute()
        self.output_dir = output_dir or (BENCHMARK_OUTPUT_DIR / 'object_effect_x_contextual_parameterized')
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.inputs_dir = self.output_dir / 'inputs'
        self.inputs_dir.mkdir(parents=True, exist_ok=True)
        self.input_paths = {}  # Cache for generated input files

    def generate_input_for_command(self, idMso: str) -> Path:
        """
        Generate input document with appropriate precondition for a command.

        Args:
            idMso: Command identifier

        Returns:
            Path to the generated input document
        """
        precondition_type = get_precondition_type(idMso)
        input_filename = f'doc_in_{precondition_type}.docx'
        input_path = self.inputs_dir / input_filename

        # Return cached path if already generated
        if precondition_type in self.input_paths:
            return self.input_paths[precondition_type]

        if input_path.exists():
            print(f"  Reusing existing input: {input_filename}")
            self.input_paths[precondition_type] = input_path
            return input_path

        print(f"  Generating input: {input_filename}")

        wd = ensure_word()
        doc = None
        try:
            doc = wd.Documents.Add()
            time.sleep(1.0)

            if precondition_type in ('table', 'table_wide'):
                # Create 3x3 table for table commands
                table = doc.Tables.Add(doc.Range(), 3, 3)
                # Fill cells with identifiers
                for i in range(1, 4):
                    for j in range(1, 4):
                        table.Cell(i, j).Range.Text = f"R{i}C{j}"
                if precondition_type == 'table':
                    table.AutoFitBehavior(1)  # wdAutoFitContent
                # table_wide: keep default page-width distribution
                time.sleep(0.2)

            elif precondition_type == 'chart':
                # Create a clustered column chart with secondary axis
                doc.InlineShapes.AddChart2(-1, 51)  # xlColumnClustered
                time.sleep(3)
                chart = doc.InlineShapes(1).Chart
                chart.SeriesCollection(2).AxisGroup = 2  # Move series 2 to secondary
                # Turn off chart title (default has one) so ChartTitle variants show a diff
                chart.HasTitle = False
                # Turn off all gridlines (default has Primary Major Horizontal ON)
                for ax_type in [1, 2]:  # xlCategory, xlValue
                    ax = chart.Axes(ax_type)
                    ax.HasMajorGridlines = False
                    ax.HasMinorGridlines = False
                time.sleep(0.5)

            elif precondition_type == 'chart_line':
                # Create a line chart (needed for drop lines, hi-lo lines, up/down bars)
                doc.InlineShapes.AddChart2(-1, 4)  # xlLine
                time.sleep(3)

            elif precondition_type == 'smartart':
                # Create a basic SmartArt with "Target" text in first node
                layouts = wd.Application.SmartArtLayouts
                sa_shape = doc.Shapes.AddSmartArt(layouts(9), 72, 72, 400, 300)  # Tab List
                time.sleep(2)
                sa = sa_shape.SmartArt
                if sa.Nodes.Count > 0:
                    sa.Nodes(1).TextFrame2.TextRange.Text = "Target"
                    for i in range(2, sa.Nodes.Count + 1):
                        sa.Nodes(i).TextFrame2.TextRange.Text = f"Item {i}"
                time.sleep(0.5)

            elif precondition_type == 'picture':
                # Insert an actual image file as inline picture
                img_path = str(Path(__file__).parent.parent / 'assets' / 'img.png')
                doc.InlineShapes.AddPicture(img_path)
                time.sleep(0.5)

            elif precondition_type == 'equation':
                # Create an equation with 'y='
                doc.Range(0, 0).Select()
                wd.CommandBars.ExecuteMso("EquationInsertNew")
                time.sleep(1.0)
                sel = wd.Selection
                sel.TypeText("y=")
                time.sleep(0.3)

            elif precondition_type == 'graphic':
                # Insert SVG graphic
                svg_path = str(Path(__file__).parent.parent / 'assets' / 'img.svg')
                doc.Shapes.AddPicture(svg_path, False, True, 72, 72, 200, 200)
                time.sleep(0.5)

            elif precondition_type == 'textbox':
                shape = doc.Shapes.AddTextbox(
                    Orientation=1,  # msoTextOrientationHorizontal
                    Left=100,
                    Top=100,
                    Width=260,
                    Height=140,
                )
                shape.TextFrame.TextRange.Text = "Target text\nSecond line\nThird line"
                shape.TextFrame.MarginTop = 8
                shape.TextFrame.MarginBottom = 8
                shape.TextFrame.MarginLeft = 8
                shape.TextFrame.MarginRight = 8
                shape.Line.Visible = True
                shape.Line.Weight = 1.25
                time.sleep(0.2)

            elif precondition_type == 'textbox_text_direction':
                shape = doc.Shapes.AddTextbox(
                    Orientation=1,  # msoTextOrientationHorizontal
                    Left=100,
                    Top=100,
                    Width=260,
                    Height=140,
                )
                # Mixed Latin/CJK text makes vertical writing mode visibly
                # distinct from 90-degree text rotation in Word.
                shape.TextFrame.TextRange.Text = "Alpha 123\n测试 Text\nBeta 456"
                shape.TextFrame.MarginTop = 8
                shape.TextFrame.MarginBottom = 8
                shape.TextFrame.MarginLeft = 8
                shape.TextFrame.MarginRight = 8
                shape.Line.Visible = True
                shape.Line.Weight = 1.25
                time.sleep(0.2)

            elif precondition_type == 'shape':
                # Create a standard shape for shape effect commands
                shape = doc.Shapes.AddShape(2, 100, 100, 150, 150)  # Parallelogram
                shape.Fill.Solid()
                shape.Fill.ForeColor.RGB = 10498160  # Purple (#7030A0)
                shape.Line.Visible = True
                shape.Line.ForeColor.RGB = 4210752  # Dark gray (#404040)
                shape.Line.Weight = 1.5
                time.sleep(0.2)


            elif precondition_type == 'line':
                # Create a line shape for arrow style commands
                line = doc.Shapes.AddLine(100, 150, 300, 150)
                line.Line.Weight = 2.0
                time.sleep(0.2)

            else:
                # Default: create a simple shape
                shape = doc.Shapes.AddShape(1, 100, 100, 100, 100)  # Rectangle
                time.sleep(0.2)

            doc.SaveAs(str(input_path.absolute()), FileFormat=16)
            time.sleep(1.0)
            for _retry in range(3):
                try:
                    doc.Close(False)
                    break
                except Exception:
                    time.sleep(2.0)
            time.sleep(0.5)

            self.input_paths[precondition_type] = input_path
            print(f"    Created: {input_filename}")
            return input_path

        except Exception as e:
            print(f"    Error creating input: {e}")
            if doc:
                for _retry in range(3):
                    try:
                        doc.Close(False)
                        break
                    except Exception:
                        time.sleep(2.0)
            raise

    def load_commands(self) -> List[Dict]:
        """Load commands classified as object_effect × contextual_parameterized."""
        commands_file = COMMAND_SETS_DIR / 'commands_by_policy_id.json'

        if not commands_file.exists():
            print(f"Warning: {commands_file} not found.")
            return []

        with open(commands_file, 'r', encoding='utf-8') as f:
            all_commands = json.load(f)

        return [
            cmd for cmd in all_commands
            if cmd.get('behavior') == 'object_effect'
            and cmd.get('scope') == 'contextual_parameterized'
        ]

    def generate_test_case(self, wd, idMso: str, param: Dict, input_path: Path) -> Dict:
        """Generate a single test case."""
        result = {
            'idMso': idMso,
            'variation': f"{param['label']}",
            'parameter': param['label'],
            'success': False,
            'message': '',
            'paths': {}
        }

        try:
            # Check if this variant needs a local doc_in with precondition
            needs_local_doc_in = False
            local_doc_in_path = None
            
            # TextDirectionGallery/horizontal needs vertical text as precondition
            if idMso == 'TextDirectionGallery' and param.get('label') == 'horizontal':
                needs_local_doc_in = True
                test_dir = self.output_dir / idMso / f"{param['label']}"
                test_dir.mkdir(parents=True, exist_ok=True)
                local_doc_in_path = test_dir / 'doc_in.docx'
                
                # Create doc_in with vertical text direction in the textbox.
                doc = wd.Documents.Open(str(input_path.absolute()))
                time.sleep(0.5)
                if doc.Shapes.Count > 0:
                    try:
                        doc.Shapes(1).TextFrame.Orientation = 4  # msoTextOrientationVerticalFarEast
                    except:
                        pass
                            
                doc.SaveAs(str(local_doc_in_path.absolute()), FileFormat=16)
                doc.Close(False)
                time.sleep(0.2)
                input_path = local_doc_in_path

            # TextAlignGallery/top needs non-top text as precondition so the
            # case is not a no-op.
            elif idMso == 'TextAlignGallery' and param.get('label') == 'top':
                needs_local_doc_in = True
                test_dir = self.output_dir / idMso / f"{param['label']}"
                test_dir.mkdir(parents=True, exist_ok=True)
                local_doc_in_path = test_dir / 'doc_in.docx'

                doc = wd.Documents.Open(str(input_path.absolute()))
                time.sleep(0.5)
                if doc.Shapes.Count > 0:
                    try:
                        doc.Shapes(1).TextFrame.VerticalAnchor = 4  # msoAnchorBottom
                    except:
                        pass

                doc.SaveAs(str(local_doc_in_path.absolute()), FileFormat=16)
                doc.Close(False)
                time.sleep(0.2)
                input_path = local_doc_in_path
            
            # PictureTransparencyGallery/0_percent needs some transparency as precondition
            elif idMso == 'PictureTransparencyGallery' and param.get('label') == '0_percent':
                needs_local_doc_in = True
                test_dir = self.output_dir / idMso / f"{param['label']}"
                test_dir.mkdir(parents=True, exist_ok=True)
                local_doc_in_path = test_dir / 'doc_in.docx'
                
                # Create doc_in with 50% transparency as precondition
                doc = wd.Documents.Open(str(input_path.absolute()))
                time.sleep(0.5)
                doc.SaveAs(str(local_doc_in_path.absolute()), FileFormat=16)
                doc.Close(False)
                time.sleep(0.2)
                _apply_picture_transparency_effect(local_doc_in_path, 0.5)
                input_path = local_doc_in_path

            # OutlineDashesGallery/solid needs a non-solid outline as precondition.
            elif idMso == 'OutlineDashesGallery' and param.get('label') == 'solid':
                needs_local_doc_in = True
                test_dir = self.output_dir / idMso / f"{param['label']}"
                test_dir.mkdir(parents=True, exist_ok=True)
                local_doc_in_path = test_dir / 'doc_in.docx'

                doc = wd.Documents.Open(str(input_path.absolute()))
                time.sleep(0.5)
                if doc.Shapes.Count > 0:
                    shape = doc.Shapes(1)
                    shape.Line.Visible = True
                    shape.Line.Weight = 2.25
                    shape.Line.DashStyle = 4  # msoLineDash
                doc.SaveAs(str(local_doc_in_path.absolute()), FileFormat=16)
                doc.Close(False)
                time.sleep(0.2)
                input_path = local_doc_in_path

            # BevelShapeGallery/none needs an existing bevel so the task is not a no-op.
            elif idMso == 'BevelShapeGallery' and param.get('label') == 'none':
                needs_local_doc_in = True
                test_dir = self.output_dir / idMso / f"{param['label']}"
                test_dir.mkdir(parents=True, exist_ok=True)
                local_doc_in_path = test_dir / 'doc_in.docx'

                doc = wd.Documents.Open(str(input_path.absolute()))
                time.sleep(0.5)
                if doc.Shapes.Count > 0:
                    shape = doc.Shapes(1)
                    shape.ThreeD.BevelTopType = 2  # Circle
                    shape.ThreeD.BevelTopDepth = 6.0
                    shape.ThreeD.BevelTopInset = 6.0
                doc.SaveAs(str(local_doc_in_path.absolute()), FileFormat=16)
                doc.Close(False)
                time.sleep(0.2)
                input_path = local_doc_in_path

            # Chart toggle preconditions:
            # "off" variants need the feature ON in doc_in so the operation is not a no-op.

            # ChartErrorBars: "none" needs error bars ON
            elif idMso == 'ChartErrorBars' and param.get('error_bar_type') is None:
                needs_local_doc_in = True
                test_dir = self.output_dir / idMso / f"{param['label']}"
                test_dir.mkdir(parents=True, exist_ok=True)
                local_doc_in_path = test_dir / 'doc_in.docx'

                doc = wd.Documents.Open(str(input_path.absolute()))
                time.sleep(0.5)
                if doc.InlineShapes.Count > 0 and doc.InlineShapes(1).HasChart:
                    doc.InlineShapes(1).Chart.SeriesCollection(1).HasErrorBars = True
                doc.SaveAs(str(local_doc_in_path.absolute()), FileFormat=16)
                doc.Close(False)
                time.sleep(0.2)
                input_path = local_doc_in_path


            doc = wd.Documents.Open(str(input_path.absolute()))
            time.sleep(0.5)

            precondition_type = get_precondition_type(idMso)
            obj = None

            # Get the appropriate object based on precondition type
            if precondition_type in ('table', 'table_wide'):
                if doc.Tables.Count > 0:
                    obj = doc.Tables(1)
            elif precondition_type == 'equation':
                if doc.OMaths.Count > 0:
                    obj = doc.OMaths(1)
            elif precondition_type in ('chart', 'chart_line'):
                if doc.InlineShapes.Count > 0 and doc.InlineShapes(1).HasChart:
                    obj = doc.InlineShapes(1).Chart
            elif precondition_type == 'smartart':
                for i in range(1, doc.Shapes.Count + 1):
                    if doc.Shapes(i).HasSmartArt:
                        obj = doc.Shapes(i).SmartArt
                        break
            elif precondition_type == 'picture':
                if doc.InlineShapes.Count > 0:
                    obj = doc.InlineShapes(1)
                elif doc.Shapes.Count > 0:
                    obj = doc.Shapes(1)
            elif precondition_type == 'graphic':
                if doc.Shapes.Count > 0:
                    obj = doc.Shapes(1)
            elif precondition_type in ('shape', 'line', 'textbox', 'textbox_text_direction'):
                if doc.Shapes.Count > 0:
                    obj = doc.Shapes(1)
            
            if obj is None:
                result['message'] = f'No {precondition_type} found in document'
                doc.Close(False)
                return result

            # Apply parameter
            if not apply_parameter(wd, idMso, param, obj=obj):
                result['message'] = f'Failed to apply parameter'
                doc.Close(False)
                return result

            time.sleep(0.2)

            # Save gold output
            test_dir = self.output_dir / idMso / f"{param['label']}"
            test_dir.mkdir(parents=True, exist_ok=True)

            gold_path = test_dir / 'gold.docx'
            spec_path = test_dir / 'spec.json'

            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
            doc.Close(False)
            time.sleep(0.2)

            if idMso == 'ObjectEffectGlowGallery' and param.get('scheme_color'):
                _fix_glow_scheme_color(gold_path, param['color_rgb'], param['scheme_color'])
            elif idMso == 'PictureEffectsGlowGallery':
                picture_glow_theme = {
                    8: (0xFFC000, 'accent4'),
                    10: (0x70AD47, 'accent6'),
                }.get(param.get('glow_theme_color'))
                if picture_glow_theme is not None:
                    _fix_glow_scheme_color(gold_path, *picture_glow_theme)

            if idMso == 'PictureTransparencyGallery':
                _apply_picture_transparency_effect(gold_path, float(param.get('transparency', 0.0)))

            # PictureRecolorGallery: now handled by COM (ColorType), no post-processing needed

            if idMso == 'RecolorColorPicker':
                duotone_hex = param.get('duotone_hex')
                if duotone_hex:
                    _apply_recolor_duotone(gold_path, duotone_hex)

            if idMso == 'PictureEffectsPresetGallery':
                preset_index = param.get('preset_index')
                if preset_index is not None:
                    _apply_picture_effects_preset(gold_path, int(preset_index))

            # TextAlignGallery, TextFillColorPicker: now handled by COM, no post-processing

            if idMso == 'TextEffectTransformGallery':
                transform_preset = param.get('transform_preset')
                if transform_preset:
                    _apply_text_transform(gold_path, transform_preset)

            if idMso == 'BevelTextGallery':
                bevel_preset = param.get('bevel_preset')
                if bevel_preset:
                    _apply_text_bevel(gold_path, bevel_preset)

            if idMso == 'PictureStylesGallery':
                style_label = param.get('label')
                if style_label:
                    _apply_picture_style(gold_path, style_label)


            # Post-process: some Cases/Stacks UI items are delimiter-wrapped EqArray structures.
            if param.get('omath_function') == 6 and (
                param.get('wrapper_begin_char') is not None or param.get('wrapper_end_char') is not None
            ):
                _wrap_eqarray_with_delimiters(
                    gold_path,
                    param.get('wrapper_begin_char') or '',
                    param.get('wrapper_end_char') or '',
                )

            if param.get('label') == 'small_fraction':
                _wrap_fraction_as_small(gold_path)

            if param.get('omath_function') == 8 and param.get('is_inverse'):
                _convert_function_name_to_inverse(gold_path)

            if param.get('label') in {'limit', 'minimum', 'maximum'}:
                _wrap_limlow_with_argument(gold_path, 'x')

            if param.get('label') == 'logarithm_with_empty_base':
                _convert_log_with_base_to_subscript(gold_path, '2', 'x')

            if param.get('omath_function') == 12 and (
                param.get('wrapper_begin_char') is not None or param.get('wrapper_end_char') is not None
            ):
                _wrap_matrix_with_delimiters(
                    gold_path,
                    param.get('wrapper_begin_char') or '',
                    param.get('wrapper_end_char') or '',
                )

            # Post-process: fix groupChrPr for ⏞ (overbrace / grouping_char_above)
            # COM doesn't reliably emit pos="top" + vertJc="bot" for ⏞
            if param.get('group_char') == 9182 and param.get('omath_function') in (9, 11):
                _fix_groupchar_above(gold_path)

            if param.get('label') in {'right_arrow_above', 'right_arrow_below'}:
                _fix_groupchar_position(gold_path, bool(param.get('position_top')))

            if idMso == 'GradientGallery':
                _assert_saved_gradient_postconditions(wd, gold_path, param)

            # Calculate relative path to input file
            if needs_local_doc_in:
                doc_in_ref = 'doc_in.docx'  # Local doc_in in same folder
            else:
                input_filename = input_path.name
                doc_in_ref = f'../../inputs/{input_filename}'

            # Generate spec
            spec = {
                'idMso': idMso,
                'behavior': 'object_effect',
                'scope': 'contextual_parameterized',
                'variation': f"{param['label']}",
                'generated_at': datetime.now().isoformat(),
                'parameter': param,
                'precondition': {
                    'type': precondition_type,
                    'select_object': f'{precondition_type}_1',
                },
                'verification': {
                    'primary': {
                        'method': 'shape_property_comparison',
                        'description': 'Word COM shape/object property comparison',
                    },
                    'invariants': [
                        {
                            'method': 'structure_preserved',
                            'description': 'Document structure preserved (object properties change)',
                        }
                    ],
                },
                'files': {
                    'doc_in': doc_in_ref,
                    'gold': 'gold.docx',
                }
            }

            write_spec_if_changed(spec_path, spec)

            result['success'] = True
            result['paths'] = {
                'doc_in': str(input_path),
                'gold': str(gold_path),
                'spec': str(spec_path)
            }

        except Exception as e:
            result['message'] = f'Error: {str(e)}'
            try:
                doc.Close(False)
            except:
                pass

        return result

    def generate_all(self, limit: int = None, command_filter: str = None,
                     variant_filter: str = None) -> Dict:
        """
        Generate test cases for all commands or a specific command.

        Args:
            limit: Limit number of commands to process
            command_filter: If specified, only process this command (idMso)
            variant_filter: If specified, only generate the matching variant label
        """
        commands = self.load_commands()

        # Filter out skipped commands
        skipped_commands = [cmd for cmd in commands if cmd['idMso'] in SKIP_COMMANDS]
        special_commands = [cmd for cmd in commands if cmd['idMso'] in SPECIAL_HANDLING_COMMANDS]
        commands = [cmd for cmd in commands if cmd['idMso'] not in SKIP_COMMANDS and cmd['idMso'] not in SPECIAL_HANDLING_COMMANDS]

        # Filter by specific command if requested
        if command_filter:
            commands = [cmd for cmd in commands if cmd['idMso'] == command_filter]
            if not commands:
                # Check if it's in skipped or special
                if any(cmd['idMso'] == command_filter for cmd in skipped_commands):
                    print(f"[WARNING] Command '{command_filter}' is in SKIP_COMMANDS")
                elif any(cmd['idMso'] == command_filter for cmd in special_commands):
                    print(f"[WARNING] Command '{command_filter}' is in SPECIAL_HANDLING_COMMANDS")
                else:
                    print(f"[ERROR] Command '{command_filter}' not found in command set")
                return None

        if limit:
            commands = commands[:limit]

        print(f"\n=== Generating test cases ===")
        if command_filter:
            print(f"Command filter: {command_filter}")
        print(f"Total commands loaded: {len(commands) + len(skipped_commands) + len(special_commands)}")
        print(f"Skipped commands: {len(skipped_commands)}")
        print(f"Special handling commands: {len(special_commands)}")
        print(f"Processing commands: {len(commands)}")
        print(f"Output: {self.output_dir}\n")

        summary = {
            'test_set': 'object_effect_x_contextual_parameterized',
            'generated_at': datetime.now().isoformat(),
            'total_commands': len(commands) + len(skipped_commands) + len(special_commands),
            'skipped': len(skipped_commands),
            'skipped_commands': {cmd['idMso']: SKIP_COMMANDS.get(cmd['idMso'], 'unknown') for cmd in skipped_commands},
            'special_handling': len(special_commands),
            'special_handling_commands': {cmd['idMso']: SPECIAL_HANDLING_COMMANDS.get(cmd['idMso'], 'unknown') for cmd in special_commands},
            'successful': 0,
            'failed': 0,
            'results': []
        }

        wd = ensure_word()
        generation_start = time.time()

        for cmd in commands:
            idMso = cmd['idMso']
            print(f"\n{idMso}...")

            # Reconnect Word to avoid "Object is not connected to server"
            try:
                wd.Documents.Count  # test connection
            except Exception:
                wd = ensure_word()
                time.sleep(1.0)

            # Generate input document for this command
            input_path = self.generate_input_for_command(idMso)

            # Extract parameters
            params = extract_parameters(wd, idMso)

            if variant_filter:
                params = [
                    param for param in params
                    if param.get('label') == variant_filter
                ]

            if not params:
                detail = (
                    f" for variant '{variant_filter}'"
                    if variant_filter else ''
                )
                print(f"  [SKIP] No parameters extracted{detail}")
                summary['failed'] += 1
                summary['results'].append({
                    'idMso': idMso,
                    'success': False,
                    'message': 'No parameters extracted'
                })
                continue

            # Generate test case for each parameter variant
            for param in params:
                result = self.generate_test_case(wd, idMso, param, input_path)

                summary['results'].append(result)
                if result['success']:
                    summary['successful'] += 1
                    print(f"  {result['variation']}: OK")
                else:
                    summary['failed'] += 1
                    print(f"  {result['variation']}: FAIL - {result['message']}")

                time.sleep(1.0)

        # Save summary
        summary_path = self.output_dir / 'generation_summary.json'
        with open(summary_path, 'w', encoding='utf-8') as f:
            json.dump(summary, f, indent=2, ensure_ascii=False, sort_keys=True)

        elapsed = time.time() - generation_start
        minutes, seconds = divmod(elapsed, 60)

        print(f"\n=== Generation Complete ===")
        print(f"Successful: {summary['successful']}")
        print(f"Failed: {summary['failed']}")
        print(f"Skipped: {summary['skipped']}")
        print(f"Special handling: {summary['special_handling']}")
        print(f"Time: {int(minutes)}m {seconds:.1f}s")

        return summary


def main():
    """CLI entry point."""
    import argparse

    parser = argparse.ArgumentParser(
        description='Generate test cases for object_effect × contextual_parameterized'
    )
    parser.add_argument('--limit', '-l', type=int, help='Limit number of commands')
    parser.add_argument('--output', '-o', help='Output directory')
    parser.add_argument('--command', '-c', help='Generate only for specific command (idMso)')
    parser.add_argument('--variant', '-v', help='Generate only the exact variant label')

    args = parser.parse_args()

    output_dir = Path(args.output) if args.output else None
    generator = ObjectEffectContextualParameterizedGenerator(output_dir=output_dir)
    generator.generate_all(
        limit=args.limit,
        command_filter=args.command,
        variant_filter=args.variant,
    )


if __name__ == '__main__':
    main()
