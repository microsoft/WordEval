"""Helpers for composition atomic-intent metadata.

These helpers keep variant-level atomic mappings next to the composition
generators without relying on the lossy `composition.commands` field.
"""

from __future__ import annotations

from typing import Any, Dict, Iterable, List, Optional, Tuple


Target = Dict[str, Any]
Intent = Dict[str, Any]


CONTENT_FORMATTING_ATOMIC_CASES: Dict[Tuple[str, Any], str] = {
    ('bold', 'on'): 'Bold/on',
    ('bold', 'off'): 'Bold/off',
    ('italic', 'on'): 'Italic/on',
    ('italic', 'off'): 'Italic/off',
    ('underline_style', 'Single'): 'UnderlineGallery/apply_Single',
    ('underline_style', 'Double'): 'UnderlineGallery/apply_Double',
    ('underline_style', 'Dash'): 'UnderlineGallery/apply_Dash',
    ('underline_style', 'Dotted'): 'UnderlineGallery/apply_Dotted',
    ('underline_color', 'Red'): 'UnderlineColorPicker/apply_Red',
    ('underline_color', 'Yellow'): 'UnderlineColorPicker/apply_Yellow',
    ('strikethrough', 'on'): 'Strikethrough/on',
    ('strikethrough', 'off'): 'Strikethrough/off',
    ('sub_super', 'subscript'): 'Subscript/on',
    ('sub_super', 'superscript'): 'Superscript/on',
    ('sub_super', 'off'): 'Superscript/off',
    ('char_border', 'on'): 'CharacterBorder/on',
    ('char_border', 'off'): 'CharacterBorder/off',
    ('char_shading', 'on'): 'CharacterShading/on',
    ('char_shading', 'off'): 'CharacterShading/off',
    ('font_name', 'Arial'): 'Font/apply_Arial',
    ('font_name', 'Comic_Sans'): 'Font/apply_Comic_Sans',
    ('font_size', '8pt'): 'FontSize/apply_8pt',
    ('font_size', '18pt'): 'FontSize/apply_18pt',
    ('font_size', '32pt'): 'FontSize/apply_32pt',
    ('font_size', '40pt'): 'FontSize/apply_40pt',
    ('font_size', 'increase'): 'FontSizeIncreaseWord/apply',
    ('font_size', 'decrease'): 'FontSizeDecreaseWord/apply',
    ('font_color', 'Red'): 'FontColorPicker/apply_Red',
    ('font_color', 'Yellow'): 'FontColorPicker/apply_Yellow',
    ('highlight', 'Blue'): 'TextHighlightColorPicker/apply_Blue',
    ('highlight', 'Yellow'): 'TextHighlightColorPicker/apply_Yellow',
    ('text_case', 'UPPERCASE'): 'ChangeCaseGallery/apply_UPPERCASE',
    ('text_case', 'lowercase'): 'ChangeCaseGallery/apply_lowercase',
    ('text_case', 'tOGGLE'): 'ChangeCaseGallery/apply_tOGGLE_cASE',
    ('outline_color', 'Blue'): 'TextOutlineColorPickerAlternate/apply_Blue',
    ('outline_color', 'Yellow'): 'TextOutlineColorPickerAlternate/apply_Yellow',
    ('outline_weight', '1pt'): 'TextOutlineWeightGallery/apply_1pt',
    ('outline_weight', '2pt'): 'TextOutlineWeightGallery/apply_2pt',
    ('outline_weight', '3pt'): 'TextOutlineWeightGallery/apply_3pt',
    ('outline_dash', 'Dash'): 'TextOutlineDashesGallery/apply_Dash',
    ('outline_dash', 'Round_Dot'): 'TextOutlineDashesGallery/apply_Round_Dot',
    ('reflection', 'type_1'): 'TextReflectionGallery/apply_type_1',
    ('reflection', 'type_3'): 'TextReflectionGallery/apply_type_3',
    ('reflection', 'type_5'): 'TextReflectionGallery/apply_type_5',
    ('reflection', 'type_7'): 'TextReflectionGallery/apply_type_7',
    ('fill_gradient', 'horizontal'): 'TextFillGradientGallery/apply_horizontal',
    ('fill_gradient', 'vertical'): 'TextFillGradientGallery/apply_vertical',
    ('fill_gradient', 'diag_L'): 'TextFillGradientGallery/apply_diagonal_down_left',
    ('fill_gradient', 'diag_R'): 'TextFillGradientGallery/apply_diagonal_down_right',
    ('alignment', 'center'): 'AlignCenter/on',
    ('alignment', 'right'): 'AlignRight/on',
    ('alignment', 'justify'): 'AlignJustify/on',
    ('indent', 'increase'): 'IndentIncreaseWord/apply',
    ('indent', 'decrease'): 'IndentDecreaseWord/apply',
    ('line_spacing', '1.15'): 'LineSpacingGallery/apply_1_15',
    ('line_spacing', '1.5'): 'LineSpacingGallery/apply_1_5',
    ('line_spacing', 'Double'): 'LineSpacingGallery/apply_Double',
    ('line_spacing', 'Triple'): 'LineSpacingGallery/apply_Triple',
    ('para_border', 'bottom'): 'BordersSelectionGallery/apply_bottom_border',
    ('para_border', 'top'): 'BordersSelectionGallery/apply_top_border',
    ('spacing_before', '0pt'): 'ParagraphSpacingBefore/set_before_0pt',
    ('spacing_before', '24pt'): 'ParagraphSpacingBefore/set_before_24pt',
    ('spacing_after', '0pt'): 'ParagraphSpacingAfter/set_after_0pt',
    ('spacing_after', '24pt'): 'ParagraphSpacingAfter/set_after_24pt',
    ('drop_cap', 'dropped'): 'DropCapInsertGallery/dropped',
    ('drop_cap', 'in_margin'): 'DropCapInsertGallery/in_margin',
    ('list_spacing', '1.5'): 'LineSpacingGallery/apply_1_5',
    ('list_spacing', 'Double'): 'LineSpacingGallery/apply_Double',
    ('list_spacing', 'Triple'): 'LineSpacingGallery/apply_Triple',
    ('list_alignment', 'center'): 'AlignCenter/on',
    ('list_alignment', 'right'): 'AlignRight/on',
    ('list_indent', 'increase'): 'IndentIncreaseWord/apply',
}

LIST_TYPE_ATOMIC_CASES: Dict[Tuple[str, int], str] = {
    ('BulletsGalleryWord', 1): 'BulletsGalleryWord/apply_solid_circle',
    ('BulletsGalleryWord', 7): 'BulletsGalleryWord/apply_checkmark',
    ('NumberingGalleryWord', 1): 'NumberingGalleryWord/apply_arabic_period',
}

OBJECT_EFFECT_ATOMIC_CASES: Dict[Tuple[str, Any], Optional[str]] = {
    ('fill_color', 'red'): 'ShapeFillColorPicker/red',
    ('fill_color', 'blue'): 'ShapeFillColorPicker/blue',
    ('outline_color', 'red'): 'ShapeOutlineColorPicker/red',
    ('outline_color', 'blue'): 'ShapeOutlineColorPicker/blue',
    ('outline_color', 'green'): 'ShapeOutlineColorPicker/green',
    ('shadow', 'outer_center'): 'ObjectEffectShadowGallery/outer_center',
    ('shadow', 'outer_diagonal_bottom_right'): (
        'ObjectEffectShadowGallery/outer_diagonal_bottom_right'
    ),
    ('glow', 'glow_11pt_accent_4'): 'ObjectEffectGlowGallery/glow_11pt_accent_4',
    ('glow', 'glow_18pt_accent_6'): 'ObjectEffectGlowGallery/glow_18pt_accent_6',
    ('soft_edges', '25pt'): 'ObjectEffectSoftEdgesGallery/soft_edge_25pt',
    ('reflection', 'tight_touching'): 'ReflectionGallery/tight_touching',
    ('reflection', 'half_touching'): 'ReflectionGallery/half_touching',
    ('bevel', 'angle'): 'BevelShapeGallery/angle',
    ('bevel', 'cross'): 'BevelShapeGallery/cross',
    ('shape_change', 'ellipse'): 'ShapeChangeShapeGallery/ellipse',
    ('shape_change', 'callout_line'): 'ShapeChangeShapeGallery/callout_line',
    ('shape_change', 'lightning_bolt'): 'ShapeChangeShapeGallery/lightning_bolt',
    ('title', 'above_chart'): 'ChartTitle/above_chart',
    ('title', 'centered_overlay'): 'ChartTitle/centered_overlay',
    ('legend', 'left'): 'ChartLegend/left',
    ('legend', 'right'): 'ChartLegend/right',
    ('legend', 'top'): 'ChartLegend/top',
    ('data_labels', 'center'): 'ChartDataLabel/center',
    ('data_labels', 'outside_end'): 'ChartDataLabel/outside_end',
    ('color_scheme', 'monochromatic_accent5'): (
        'ChartColorsGallery/monochromatic_accent5'
    ),
    ('color_scheme', 'monochromatic_palette_7'): (
        'ChartColorsGallery/monochromatic_palette_7'
    ),
    ('gridlines', 'major_value_gridlines'): 'ChartGridlinesGallery/major_horizontal',
    ('gridlines', 'minor_horizontal'): 'ChartGridlinesGallery/minor_horizontal',
    ('trendline', 'linear_trendline'): 'ChartTrendline/linear',
    ('data_table', 'no_legend_key'): 'ChartDataTable/no_legend_key',
    ('series_color', 'blue'): 'ColorPickerChartLineStyles/blue',
    ('series_color', 'red'): 'ColorPickerChartLineStyles/red',
    ('chart_lines', 'drop_lines'): 'ChartLines/drop_lines',
    ('chart_lines', 'hilo_lines'): 'ChartLines/hilo_lines',
    ('up_down_bars', 'show'): 'ChartUpDownBars/show_up_down_bars',
    ('pic_shadow', 'outer_center'): 'PictureEffectsShadowGallery/outer_center',
    ('pic_shadow', 'outer_diagonal_bottom_right'): (
        'PictureEffectsShadowGallery/outer_diagonal_bottom_right'
    ),
    ('pic_glow', 'glow_11pt_accent_4'): 'PictureEffectsGlowGallery/11pt_accent_4',
    ('pic_soft_edges', '25pt'): 'PictureEffectsSoftEdgesGallery/25pt',
    ('pic_reflection', 'tight_touching'): 'PictureReflectionGallery/tight_touching',
    ('pic_reflection', 'half_touching'): 'PictureReflectionGallery/half_touching',
    ('pic_reflection', 'half_4pt_offset'): 'PictureReflectionGallery/half_4pt_offset',
    ('pic_recolor', 'grayscale'): 'PictureRecolorGallery/grayscale',
    ('pic_recolor', 'black_white'): 'PictureRecolorGallery/black_and_white_50',
    ('saturation', '0pct'): 'PictureSaturationGallery/0_percent',
    ('saturation', '33pct'): 'PictureSaturationGallery/33_percent',
    ('temperature', '4700K'): 'PictureColorTempertatureGallery/4700K',
    ('temperature', '11200K'): 'PictureColorTempertatureGallery/11200K',
    ('sharpen_soften', 'sharpen_50'): 'PictureSharpenSoftenGallery/sharpen_50',
    ('sharpen_soften', 'soften_50'): 'PictureSharpenSoftenGallery/soften_50',
    ('artistic_effect', 'glow_edges'): 'PictureArtisticEffectsGallery/glow_edges',
    ('artistic_effect', 'paint_strokes'): 'PictureArtisticEffectsGallery/paint_strokes',
    ('rotation_3d', 'oblique_bottom_right'): (
        'PictureRotationGallery/oblique_bottom_right'
    ),
    ('rotation_3d', 'isometric_top_up'): 'PictureRotationGallery/isometric_top_up',
    ('pic_outline', 'red'): 'OutlineColorPicker/red',
    ('pic_outline', 'blue'): 'OutlineColorPicker/blue',
    ('ext_shadow', 'outer_center'): 'PictureEffectsShadowGallery/outer_center',
    ('ext_shadow', 'outer_diagonal_bottom_right'): (
        'PictureEffectsShadowGallery/outer_diagonal_bottom_right'
    ),
    ('ext_glow', 'glow_11pt_accent_4'): 'PictureEffectsGlowGallery/11pt_accent_4',
    ('ext_soft_edges', '25pt'): 'PictureEffectsSoftEdgesGallery/25pt',
    ('ext_reflection', 'tight_touching'): 'PictureReflectionGallery/tight_touching',
    ('crop_aspect_ratio', '16_9'): 'PictureCropAspectRatio16To9/apply',
    ('crop_aspect_ratio', '1_1'): 'PictureCropAspectRatio1To1/apply',
    ('sa_style', 5): 'SmartArtStylesGallery/intense_effect',
    ('sa_colors', 5): 'SmartArtChangeColorsGallery/colorful_range_2_3',
    ('sa_layout', 20): 'SmartArtLayoutGallery/continuous_picture_list',
    ('node_outline_weight', '3pt'): 'OutlineWeightGallery/3_0pt',
    ('node_outline_weight', '6pt'): 'OutlineWeightGallery/6_0pt',
    ('node_outline_dash', 'dash'): 'OutlineDashesGallery/dash',
    ('connector_arrow', 'none_to_triangle'): 'ArrowStyleGallery/none_to_triangle',
    ('text_fill', 'blue'): 'TextFillColorPicker/blue',
    ('text_fill', 'green'): 'TextFillColorPicker/green',
    ('text_fill', 'red'): 'TextFillColorPicker/red',
    ('text_transform', 'fade_right'): 'TextEffectTransformGallery/fade_right',
    ('text_3d', 'isometric_left_down'): (
        'TextEffects3DRotationGallery/isometric_left_down'
    ),
    ('text_bevel', 'angle'): 'BevelTextGallery/angle',
    ('text_bevel', 'circle'): 'BevelTextGallery/circle',
    ('table_style', 'Grid Table 4 - Accent 1'): (
        'TableStylesGalleryWord/Grid_Table_4___Accent_1'
    ),
    ('text_direction', 2): 'TextDirectionGallery/rotate_90_upward',
    ('border_style', 7): 'BorderStylesGallery/double',
    ('autofit', 'contents'): 'TableAutoFitContents/apply',
    ('autofit', 'window'): 'TableAutoFitWindow/apply',
    ('distribute_rows', 'apply'): 'TableRowsDistribute/apply',
    ('distribute_cols', 'apply'): 'TableColumnsDistribute/apply',
    ('cell_align', 'on'): 'TableCellAlignMiddleCenter/on',
    ('align_command', 'ObjectsAlignLeftSmart'): 'ObjectsAlignLeftSmart/apply',
    ('align_command', 'ObjectsAlignRightSmart'): 'ObjectsAlignRightSmart/apply',
    ('align_command', 'ObjectsAlignTopSmart'): 'ObjectsAlignTopSmart/apply',
    ('align_command', 'ObjectsAlignBottomSmart'): 'ObjectsAlignBottomSmart/apply',
    ('align_command', 'ObjectsAlignCenterHorizontalSmart'): (
        'ObjectsAlignCenterHorizontalSmart/apply'
    ),
    ('align_command', 'ObjectsAlignMiddleVerticalSmart'): (
        'ObjectsAlignMiddleVerticalSmart/apply'
    ),
    ('distribute_command', 'AlignDistributeHorizontally'): (
        'AlignDistributeHorizontally/apply'
    ),
    ('distribute_command', 'AlignDistributeVertically'): (
        'AlignDistributeVertically/apply'
    ),
    ('group_command', 'ObjectsGroup'): 'ObjectsGroup/apply',
    ('ungroup_command', 'ObjectsUngroup'): 'ObjectsUngroup/apply',
    ('rotate_angle', 90): 'ObjectRotateGallery/apply_rotate_right_90',
    ('rotate_angle', -90): 'ObjectRotateGallery/apply_rotate_left_90',
    ('flip_direction', 'horizontal'): 'ObjectRotateGallery/apply_flip_horizontal',
    ('flip_direction', 'vertical'): 'ObjectRotateGallery/apply_flip_vertical',
    ('wrap_type', 0): 'TextWrapGallery/apply_square',
    ('wrap_type', 1): 'TextWrapGallery/apply_tight',
    ('position_preset', (-999995, -999999)): 'PicturePositionGallery/apply_top_center',
    ('position_preset', (-999998, -999999)): 'PicturePositionGallery/apply_top_left',
    ('position_preset', (-999995, -999995)): 'PicturePositionGallery/apply_middle_center',
    ('position_preset', (-999996, -999997)): 'PicturePositionGallery/apply_bottom_right',
}

PREREQUISITE_STEP_ATOMIC_CASES: Dict[Tuple[str, Any], Optional[str]] = {
    ('insert_table', (2, 2)): 'TableInsertGallery/2x2',
    ('insert_table', (3, 5)): 'TableInsertGallery/3x5',
    ('insert_table', (5, 3)): 'TableInsertGallery/5x3',
    ('insert_table', (5, 5)): 'TableInsertGallery/5x5',
    ('insert_icon', 'star'): 'IconInsertFromFile/insert_blue_star_icon',
    ('insert_icon', 'circle'): 'IconInsertFromFile/insert_orange_circle_icon',
    ('insert_shape', 1): 'GalleryAllShapesAndCanvas/rectangle',
    ('insert_shape', 33): 'GalleryAllShapesAndCanvas/right_arrow',
    ('insert_shape', 89): 'GalleryAllShapesAndCanvas/explosion_1',
    ('insert_shape', 93): 'GalleryAllShapesAndCanvas/star_8point',
    ('insert_smartart', 1): 'SmartArtInsert/insert_basic_list',
    ('insert_smartart', 102): 'SmartArtInsert/insert_hierarchy',
    ('insert_textbox', 'draw'): 'TextBoxInsert/draw',
    ('insert_equation', 'new'): 'EquationInsertNew/apply',
    ('chart_title', 'above_chart'): 'ChartTitle/above_chart',
    ('chart_legend', 'right'): 'ChartLegend/right',
    ('chart_trendline', 'linear'): 'ChartTrendline/linear',
    ('chart_data_labels', 'outside_end'): 'ChartDataLabel/outside_end',
    ('chart_gridlines', ('value', 'minor')): 'ChartGridlinesGallery/minor_horizontal',
    ('picture_shadow', 25): 'PictureEffectsShadowGallery/outer_center',
    ('picture_reflection', 5): 'PictureReflectionGallery/half_4pt_offset',
    ('picture_glow', '11pt_accent_4'): 'PictureEffectsGlowGallery/11pt_accent_4',
    ('shape_fill', 'red'): 'ShapeFillColorPicker/red',
    ('shape_shadow', 25): 'ObjectEffectShadowGallery/outer_center',
    ('shape_bevel', 6): 'BevelShapeGallery/angle',
    ('shape_glow', '11pt_accent_4'): 'ObjectEffectGlowGallery/glow_11pt_accent_4',
    ('table_style', 'Grid Table 4 - Accent 1'): (
        'TableStylesGalleryWord/Grid_Table_4___Accent_1'
    ),
    ('table_cell_align', 'middle_center'): 'TableCellAlignMiddleCenter/on',
    ('table_border_style', 7): 'BorderStylesGallery/double',
    ('table_border_style', 'double'): 'BorderStylesGallery/double',
    ('table_text_align', 'center'): 'AlignCenter/on',
    ('smartart_style', 5): 'SmartArtStylesGallery/intense_effect',
    ('smartart_colors', 5): 'SmartArtChangeColorsGallery/colorful_range_2_3',
    ('smartart_layout', 20): 'SmartArtLayoutGallery/continuous_picture_list',
    ('equation_professional', 'apply'): 'EquationProfessionalAll/on',
    ('equation_fraction', 0): 'EquationFractionGallery/stacked',
    ('equation_radical', 'square_root'): 'EquationRadicalGallery/square_root',
    ('equation_script', 'superscript'): 'EquationScriptGallery/superscript',
    ('textbox_direction', 3): 'TextDirection/apply',
    ('textbox_outline_color', 12611584): 'ShapeOutlineColorPicker/blue',
}

CONTENT_OPERATIONS_ATOMIC_CASES: Dict[Tuple[str, Any], Optional[str]] = {
    ('footnote', 'default'): 'FootnoteInsert/apply',
    ('endnote', 'default'): 'EndnoteInsertWord/apply',
    ('header_gallery', 'viewmaster'): 'HeaderInsertGallery/viewmaster',
    ('footer_gallery', 'ion_light'): 'FooterInsertGallery/ion_light',
    ('page_number_header', 'plain_number_1'): (
        'PageNumbersInHeaderInsertGallery/plain_number_1'
    ),
    ('page_number_footer', 'plain_number_1'): (
        'PageNambersInFooterInsertGallery/plain_number_1'
    ),
    ('citation', 'Smith2020'): 'CitationInsert/smith2020',
    ('bibliography', 'insert'): 'BibliographyInsert/apply',
    ('accept_changes', 'apply'): 'AcceptAllChangesInDocAndStopTracking/apply',
    ('reject_changes', 'apply'): 'RejectAllChangesInDocAndStopTracking/apply',
    ('delete_comments', 'apply'): 'ReviewDeleteAllCommentsInDocument/apply',
    ('delete_resolved_comments', 'apply'): 'DeleteAllResolvedComments/apply',
    ('remove_watermark', 'apply'): 'WatermarkRemove/apply',
    ('remove_header', 'apply'): 'HeaderFooterRemoveHeaderWord/apply',
    ('remove_footer', 'apply'): 'HeaderFooterRemoveFooterWord/apply',
}

SCOPE_CHANGING_ATOMIC_CASES: Dict[Tuple[str, Any], Optional[str]] = {
    ('insert_row_above', 'apply'): 'TableRowsInsertAboveWord/apply',
    ('insert_row', 'bottom'): 'TableRowsInsertBelowWord/apply',
    ('insert_row', 'below_first'): 'TableRowsInsertBelowWord/apply',
    ('insert_column', 'end'): 'TableColumnsInsertRight/apply',
    ('insert_column', 'right_of_first'): 'TableColumnsInsertRight/apply',
    ('merge_cells', 'apply'): 'MergeCells/apply',
    ('split_cells', 'apply'): 'SplitCells/apply',
    ('split_table', 'apply'): 'TableSplitTable/apply',
    ('style_table', 'Grid Table 4 - Accent 1'): (
        'TableStylesGalleryWord/Grid_Table_4___Accent_1'
    ),
    ('distribute_rows', 'apply'): 'TableRowsDistribute/apply',
    ('distribute_columns', 'apply'): 'TableColumnsDistribute/apply',
    ('align_cell', (1, 1)): 'TableCellAlignMiddleCenter/on',
    ('align_row', (3, 2)): 'TableCellAlignBottomRight/on',
    ('align_paragraph', 1): 'AlignCenter/on',
    ('align_paragraph', 2): 'AlignRight/on',
    ('fill_row', 'typing'): None,
    ('fill_cell', 'typing'): None,
    ('shade_row', 49407): None,
    ('shade_column', 49407): None,
    ('add_smartart_node', 2): 'SmartArtAddShapeAfter/apply',
    ('add_smartart_child', 5): 'SmartArtAddShapeBelow/apply',
    ('promote_node', 'apply'): 'SmartArtPromote/apply',
    ('demote_node', 'apply'): 'SmartArtDemote/apply',
    ('reorder_node_up', 'apply'): 'SmartArtReorderUp/apply',
    ('reorder_node_down', 'apply'): 'SmartArtReorderDown/apply',
    ('change_chart_type', 4): None,
    ('add_drop_lines', 'apply'): 'ChartLines/drop_lines',
    ('add_chart_title', 'above'): 'ChartTitle/above_chart',
    ('format_chart_title', 'typing_and_bold'): None,
    ('insert_fraction', 0): 'EquationFractionGallery/stacked',
    ('insert_integral_in_numerator', 8747): (
        'EquationIntegralGallery/integral_with_stacked_limits'
    ),
    ('insert_delimiter', (91, 93)): 'EquationDelimiterGallery/square_brackets',
    ('insert_matrix_in_delimiter', (2, 2)): 'EquationMatrixGallery/empty_2x2',
    ('apply_bullets', (1, 1)): 'BulletsGalleryWord/apply_solid_circle',
    ('apply_numbering', (2, 1)): 'NumberingGalleryWord/apply_arabic_period',
    ('set_list_level', 2): 'ListLevelGallery/apply_level_2',
    ('set_list_level_range', 2): 'ListLevelGallery/apply_level_2',
}


def _target(extractor: str, prop: str, **extra: Any) -> Target:
    target: Target = {'extractor': extractor, 'property': prop}
    target.update(extra)
    return target


def _targets_for_slot(slot: str, value: Any) -> List[Target]:
    text = 'text_formatting'
    para = 'paragraph_formatting'
    all_para = 'all_paragraphs_formatting'

    if slot in ('bold', 'italic', 'strikethrough', 'font_size', 'font_name', 'font_color'):
        prop = {
            'bold': 'bold',
            'italic': 'italic',
            'strikethrough': 'strikethrough',
            'font_size': 'font_size',
            'font_name': 'font_name',
            'font_color': 'font_color',
        }[slot]
        return [_target(text, prop)]
    if slot == 'underline_style':
        return [_target(text, 'underline')]
    if slot == 'underline_color':
        return [_target(text, 'underline_color')]
    if slot == 'sub_super':
        if value == 'subscript':
            return [_target(text, 'subscript'), _target(text, 'superscript')]
        if value in ('superscript', 'off'):
            return [_target(text, 'superscript'), _target(text, 'subscript')]
    if slot == 'char_border':
        return [_target(text, 'border')]
    if slot == 'char_shading':
        return [_target(text, 'shading'), _target(text, 'shading_color')]
    if slot == 'highlight':
        return [_target(text, 'highlight_color')]
    if slot == 'text_case':
        return [_target('document_text', 'selected_text')]
    if slot == 'outline_color':
        return [_target(text, 'text_outline_visible'), _target(text, 'text_outline_color')]
    if slot == 'outline_weight':
        return [_target(text, 'text_outline_visible'), _target(text, 'text_outline_weight')]
    if slot == 'outline_dash':
        return [_target(text, 'text_outline_visible'), _target(text, 'text_outline_dash_style')]
    if slot == 'reflection':
        return [_target(text, 'text_reflection_type')]
    if slot == 'fill_gradient':
        return [_target(text, 'text_fill_type'), _target(text, 'text_fill_gradient_angle')]
    if slot == 'alignment':
        return [_target(text, 'alignment')]
    if slot == 'indent':
        comparison = {
            'increase': 'at_least_reference',
            'decrease': 'at_most_reference',
        }.get(value)
        return [
            _target(text, 'left_indent', comparison=comparison),
            _target(text, 'first_line_indent'),
        ]
    if slot == 'line_spacing':
        return [_target(text, 'line_spacing')]
    if slot == 'para_border':
        return [_target(para, 'paragraph_borders')]
    if slot == 'spacing_before':
        return [_target(all_para, 'space_before')]
    if slot == 'spacing_after':
        return [_target(all_para, 'space_after')]
    if slot == 'drop_cap':
        return [_target('dropcap_properties', 'position')]
    if slot == 'list_type':
        return [_target(all_para, 'list_string')]
    if slot == 'list_spacing':
        return [_target(all_para, 'line_spacing'), _target(all_para, 'line_spacing_rule')]
    if slot == 'list_alignment':
        return [_target(all_para, 'alignment')]
    if slot == 'list_indent':
        comparison = {
            'increase': 'at_least_reference',
            'decrease': 'at_most_reference',
        }.get(value)
        return [
            _target(all_para, 'left_indent', comparison=comparison),
            _target(all_para, 'first_line_indent'),
        ]
    return []


def _targets_for_object_effect_slot(slot: str, value: Any) -> List[Target]:
    shape = 'shape_properties'
    picture = 'picture_effect_properties'
    picture_crop = 'picture_crop_properties'
    chart = 'chart_properties'
    smartart = 'smartart_properties'
    smartart_subnode = 'smartart_subnode_properties'
    table = 'table_properties'
    table_autofit = 'table_autofit_flags'
    multi_shape = 'multi_shape_properties'

    if slot == 'fill_color':
        return [_target(shape, 'fill_fore_color_rgb'), _target(shape, 'fill_type')]
    if slot == 'outline_color':
        return [_target(shape, 'line_fore_color_rgb'), _target(shape, 'line_visible')]
    if slot in ('shadow',):
        return [_target(shape, 'shadow_type')]
    if slot == 'glow':
        return [_target(shape, 'glow_radius'), _target(shape, 'glow_color_rgb')]
    if slot == 'soft_edges':
        return [_target(shape, 'soft_edge_radius')]
    if slot == 'reflection':
        return [_target(shape, 'reflection_type')]
    if slot == 'bevel':
        return [
            _target(shape, 'threed_bevel_top_type'),
            _target('bevel_render_signatures', 'scene_present'),
            _target('bevel_render_signatures', 'camera_present'),
            _target('bevel_render_signatures', 'light_rig_present'),
        ]
    if slot == 'shape_change':
        return [_target(shape, 'auto_shape_type')]
    if slot == 'title':
        return [
            _target(chart, 'has_title'),
            _target(chart, 'title_text'),
            _target(chart, 'title_include_in_layout'),
        ]
    if slot == 'legend':
        return [_target(chart, 'has_legend'), _target(chart, 'legend_position')]
    if slot == 'data_labels':
        return [_target(chart, 'series_data_labels')]
    if slot == 'trendline':
        return [_target(chart, 'series_trendlines')]
    if slot == 'series_color':
        return [_target(chart, 'series_line_colors')]
    if slot == 'color_scheme':
        return [_target(chart, 'chart_color')]
    if slot == 'gridlines':
        return [_target(chart, 'value_axis_major_gridlines')]
    if slot == 'data_table':
        return [
            _target(chart, 'has_data_table'),
            _target(chart, 'data_table_show_legend_key'),
        ]
    if slot == 'chart_lines':
        prop = 'has_drop_lines' if value == 'drop_lines' else 'has_hi_lo_lines'
        return [_target(chart, prop)]
    if slot == 'up_down_bars':
        return [_target(chart, 'has_up_down_bars')]
    if slot in ('pic_shadow', 'ext_shadow'):
        return [_target(picture, 'shadow_type')]
    if slot in ('pic_glow', 'ext_glow'):
        return [_target(picture, 'glow_radius'), _target(picture, 'glow_color_rgb')]
    if slot in ('pic_soft_edges', 'ext_soft_edges'):
        return [_target(picture, 'soft_edge_radius')]
    if slot in ('pic_reflection', 'ext_reflection'):
        return [_target(picture, 'reflection_type')]
    if slot == 'pic_recolor':
        return [_target(picture, 'color_type')]
    if slot == 'saturation':
        return [_target(picture, 'picture_effect_params', effect_type=24)]
    if slot == 'temperature':
        return [_target(picture, 'picture_effect_params', effect_type=7)]
    if slot == 'sharpen_soften':
        return [_target(picture, 'picture_effect_params', effect_type=25)]
    if slot == 'artistic_effect':
        effect_type = {'glow_edges': 12, 'paint_strokes': 18}.get(value)
        return [_target(picture, 'picture_effect_present', effect_type=effect_type)]
    if slot == 'rotation_3d':
        return [
            _target(picture, 'threed_preset_camera'),
            _target(picture, 'threed_rotation_x'),
            _target(picture, 'threed_rotation_y'),
        ]
    if slot == 'pic_outline':
        return [_target(picture, 'line_fore_color_rgb'), _target(picture, 'line_visible')]
    if slot == 'crop_aspect_ratio':
        return [
            _target(picture_crop, 'crop_left', tolerance=10.0),
            _target(picture_crop, 'crop_right', tolerance=10.0),
            _target(picture_crop, 'crop_top', tolerance=10.0),
            _target(picture_crop, 'crop_bottom', tolerance=10.0),
            _target(picture_crop, 'width', tolerance=10.0),
            _target(picture_crop, 'height', tolerance=10.0),
        ]
    if slot == 'sa_style':
        return [_target(smartart, 'quick_style_name')]
    if slot == 'sa_colors':
        return [_target(smartart, 'color_name')]
    if slot == 'node_outline_weight':
        # The linked OutlineWeightGallery atomic sets the outline of the Target
        # node shape, which is what smartart_subnode_properties reads.
        return [_target(smartart_subnode, 'line_weight', tolerance=0.5)]
    if slot == 'node_outline_dash':
        return [_target(smartart_subnode, 'line_dash_style')]
    if slot == 'connector_arrow':
        # ArrowStyleGallery applies the begin/end styles to the connector owned
        # by the Target node.  Reuse that target-bound atomic observation rather
        # than comparing every connector line in the SmartArt drawing cache.
        return [
            _target(smartart_subnode, 'connector_begin_arrow'),
            _target(smartart_subnode, 'connector_end_arrow'),
        ]
    if slot == 'text_fill':
        return [_target(smartart_subnode, 'text_fill_rgb')]
    if slot == 'text_transform':
        return [_target(smartart_subnode, 'warp_format')]
    if slot == 'text_3d':
        return [_target(smartart_subnode, 'threed_preset_camera')]
    if slot == 'text_bevel':
        return [_target(smartart_subnode, 'bevel_top_type')]
    if slot == 'table_style':
        return [_target(table, 'style')]
    if slot == 'text_direction':
        return [_target(table, 'cell_text_orientations')]
    if slot == 'border_style':
        return [_target(table, 'outer_border_line_styles')]
    if slot == 'autofit':
        return [
            _target(table_autofit, 'allow_autofit'),
            _target(table_autofit, 'preferred_width_type'),
            _target(table_autofit, 'preferred_width', tolerance=2.0),
        ]
    if slot == 'distribute_rows':
        return [_target(table, 'row_heights', tolerance=2.0)]
    if slot == 'distribute_cols':
        return [_target(table, 'column_widths', tolerance=2.0)]
    if slot == 'cell_align':
        return [
            _target(table, 'cell_alignments'),
            _target(table, 'cell_paragraph_alignments'),
        ]
    if slot == 'align_command':
        metric = {
            'ObjectsAlignLeftSmart': 'left_edge_spread',
            'ObjectsAlignCenterHorizontalSmart': 'horizontal_center_spread',
            'ObjectsAlignRightSmart': 'right_edge_spread',
            'ObjectsAlignTopSmart': 'top_edge_spread',
            'ObjectsAlignMiddleVerticalSmart': 'vertical_center_spread',
            'ObjectsAlignBottomSmart': 'bottom_edge_spread',
        }.get(value)
        return [_target(multi_shape, metric, tolerance=2.0)] if metric else []
    if slot == 'distribute_command':
        metric = {
            'AlignDistributeHorizontally': 'horizontal_center_gap_spread',
            'AlignDistributeVertically': 'vertical_center_gap_spread',
        }.get(value)
        return [_target(multi_shape, metric, tolerance=2.0)] if metric else []
    if slot == 'rotate_angle':
        return [
            _target(multi_shape, 'rotation', tolerance=1.0),
            _target(multi_shape, 'width', tolerance=2.0),
            _target(multi_shape, 'height', tolerance=2.0),
        ]
    if slot == 'group_command':
        return [
            _target(multi_shape, 'left', tolerance=2.0),
            _target(multi_shape, 'top', tolerance=2.0),
        ]
    if slot == 'ungroup_command':
        return [_target(multi_shape, 'is_group')]
    if slot == 'flip_direction':
        prop = 'flip_horizontal' if value == 'horizontal' else 'flip_vertical'
        return [
            _target(multi_shape, prop),
            _target(multi_shape, 'left', tolerance=2.0),
            _target(multi_shape, 'top', tolerance=2.0),
        ]
    if slot == 'wrap_type':
        return [_target(multi_shape, 'wrap_type')]
    if slot == 'position_preset':
        return [
            _target(multi_shape, 'actual_left', tolerance=2.0),
            _target(multi_shape, 'actual_top', tolerance=2.0),
        ]
    if slot == 'resize_shape':
        return [
            _target(multi_shape, 'width', tolerance=2.0),
            _target(multi_shape, 'height', tolerance=2.0),
        ]
    if slot == 'move_shape':
        return [
            _target(multi_shape, 'actual_left', tolerance=2.0),
            _target(multi_shape, 'actual_top', tolerance=2.0),
        ]
    return []


def _targets_for_content_operation(operation: str, value: Any) -> List[Target]:
    verifier = 'content_operation_verifier'
    hf = 'header_footer_content'

    if operation == 'footnote':
        return [_target(verifier, 'footnote_verification')]
    if operation == 'endnote':
        return [_target(verifier, 'endnote_verification')]
    if operation == 'header_gallery':
        return [
            _target(
                verifier,
                'header_footer_style_match',
                idMso='HeaderInsertGallery',
                part_type='header',
            )
        ]
    if operation == 'footer_gallery':
        return [
            _target(
                verifier,
                'header_footer_style_match',
                idMso='FooterInsertGallery',
                part_type='footer',
            )
        ]
    if operation == 'page_number_header':
        return [
            _target(
                verifier,
                'header_footer_style_match',
                idMso='PageNumbersInHeaderInsertGallery',
                part_type='header',
            )
        ]
    if operation == 'page_number_footer':
        return [
            _target(
                verifier,
                'header_footer_style_match',
                idMso='PageNambersInFooterInsertGallery',
                part_type='footer',
            )
        ]
    if operation == 'citation':
        return [_target(verifier, 'citation_verification', source_tag=value)]
    if operation == 'bibliography':
        return [_target(verifier, 'bibliography_verification', idMso='BibliographyInsert')]
    if operation in ('accept_changes', 'reject_changes'):
        return [
            _target('revision_info', 'revision_count'),
            _target('revision_info', 'track_revisions'),
            _target('document_text', 'text'),
        ]
    if operation == 'delete_comments':
        return [_target('comment_info', 'comment_count')]
    if operation == 'delete_resolved_comments':
        return [
            _target('comment_info', 'resolved_count', expected=0),
            _target('comment_info', 'comment_count'),
        ]
    if operation == 'remove_watermark':
        return [_target(hf, 'shape_count', part='header')]
    if operation == 'remove_header':
        return [
            _target(hf, 'text', part='header'),
            _target(hf, 'field_count', part='header'),
            _target(hf, 'shape_count', part='header'),
            _target(hf, 'inline_shape_count', part='header'),
        ]
    if operation == 'remove_footer':
        return [
            _target(hf, 'text', part='footer'),
            _target(hf, 'field_count', part='footer'),
            _target(hf, 'shape_count', part='footer'),
            _target(hf, 'inline_shape_count', part='footer'),
        ]
    return []


def _chart_insert_case(params: Dict[str, Any]) -> Optional[str]:
    categories = tuple(params.get('categories') or [])
    series = params.get('series') or []
    series_key = tuple(
        (item.get('name'), tuple(item.get('values') or []))
        for item in series
    )
    key = (params.get('chart_type'), categories, series_key)
    chart_cases = {
        (
            51,
            ('Q1', 'Q2'),
            (('Revenue', (100, 150)),),
        ): 'ChartInsert/insert_column_chart',
        (
            4,
            ('Jan', 'Feb', 'Mar'),
            (('Temperature', (10, 15, 20)),),
        ): 'ChartInsert/insert_line_chart',
        (
            5,
            ('A', 'B', 'C'),
            (('Sales', (40, 35, 25)),),
        ): 'ChartInsert/insert_pie_chart',
    }
    return chart_cases.get(key)


def _prerequisite_operation_key(step: Dict[str, Any]) -> Tuple[str, Any]:
    op = step.get('operation')
    params = step.get('params') or {}
    idMso = step.get('idMso')

    if op == 'insert_table':
        return ('insert_table', (params.get('rows'), params.get('cols')))
    if op == 'insert_chart':
        return ('insert_chart', params)
    if op == 'insert_icon':
        return ('insert_icon', params.get('shape', 'star'))
    if op == 'insert_shape':
        return ('insert_shape', params.get('type'))
    if op == 'insert_smartart':
        return ('insert_smartart', params.get('layout_index'))
    if op == 'insert_equation':
        return ('insert_equation', 'new')
    if op == 'insert_textbox':
        return ('insert_textbox', 'draw')

    if idMso == 'TableStylesGalleryWord':
        return ('table_style', params.get('style_name'))
    if idMso == 'TableCellAlignMiddleCenter':
        return ('table_cell_align', 'middle_center')
    if idMso == 'BorderStylesGallery':
        return ('table_border_style', params.get('line_style'))
    if idMso == 'BorderColorPicker':
        return ('table_border_color', params.get('color_bgr'))
    if idMso == 'AlignCenter':
        return ('table_text_align', params.get('alignment'))
    if idMso == 'ChartTitle':
        return ('chart_title', params.get('position'))
    if idMso == 'ChartLegend':
        return ('chart_legend', params.get('position'))
    if idMso == 'ChartTrendline':
        return ('chart_trendline', params.get('type'))
    if idMso == 'ChartDataLabel':
        return ('chart_data_labels', params.get('position'))
    if idMso == 'ChartGridlinesGallery':
        return ('chart_gridlines', (params.get('axis'), params.get('gridline')))
    if idMso == 'PictureEffectsShadowGallery':
        return ('picture_shadow', params.get('shadow_type'))
    if idMso == 'PictureReflectionGallery':
        return ('picture_reflection', params.get('reflection_type'))
    if idMso == 'PictureEffectsGlowGallery':
        return ('picture_glow', '11pt_accent_4')
    if idMso == 'ShapeFillColorPicker':
        color = 'red' if params.get('color_bgr') == 255 else params.get('color_bgr')
        return ('shape_fill', color)
    if idMso == 'ObjectEffectShadowGallery':
        return ('shape_shadow', params.get('shadow_type'))
    if idMso == 'BevelShapeGallery':
        return ('shape_bevel', params.get('bevel_type'))
    if idMso == 'ObjectEffectGlowGallery':
        return ('shape_glow', '11pt_accent_4')
    if idMso == 'OutlineDashesGallery':
        return ('shape_outline_dash', params.get('dash_style'))
    if idMso == 'SmartArtStylesGallery':
        return ('smartart_style', params.get('style_index'))
    if idMso == 'SmartArtChangeColorsGallery':
        return ('smartart_colors', params.get('color_index'))
    if idMso == 'SmartArtLayoutGallery':
        return ('smartart_layout', params.get('layout_index'))
    if idMso == 'EquationProfessionalAll':
        return ('equation_professional', 'apply')
    if idMso == 'EquationFractionGallery':
        return ('equation_fraction', params.get('frac_type', 0))
    if idMso == 'EquationRadicalGallery':
        return ('equation_radical', 'square_root')
    if idMso == 'EquationScriptGallery':
        return ('equation_script', 'superscript')
    if idMso == 'TextDirection':
        return ('textbox_direction', params.get('orientation'))
    if idMso == 'TextFillColorPicker':
        return ('textbox_text_fill', params.get('color_bgr'))
    if idMso == 'ShapeOutlineColorPicker':
        return ('textbox_outline_color', params.get('color_bgr'))
    return (op or idMso or 'unknown', None)


def atomic_case_for_prerequisite_step(step: Dict[str, Any]) -> Optional[str]:
    operation, value = _prerequisite_operation_key(step)
    if operation == 'insert_chart' and isinstance(value, dict):
        return _chart_insert_case(value)
    return PREREQUISITE_STEP_ATOMIC_CASES.get((operation, value))


def _targets_for_prerequisite_step(step: Dict[str, Any]) -> List[Target]:
    operation, value = _prerequisite_operation_key(step)

    if operation == 'insert_table':
        return [
            _target('table_properties', 'row_count'),
            _target('table_properties', 'column_count'),
        ]
    if operation == 'insert_chart':
        return [
            _target('saved_chart_properties', 'chart_type'),
            _target('saved_chart_properties', 'series_count'),
            _target('saved_chart_properties', 'series_names'),
            _target('saved_chart_properties', 'series_values'),
            _target('saved_chart_properties', 'categories'),
        ]
    if operation == 'insert_icon':
        return [
            _target('icon_info', 'has_icon'),
            _target('icon_info', 'svg_shapes'),
            _target('icon_info', 'width', tolerance=1.0),
            _target('icon_info', 'height', tolerance=1.0),
        ]
    if operation == 'insert_shape':
        return [
            _target('shape_properties', 'auto_shape_type'),
            _target('shape_properties', 'actual_left', tolerance=2.0),
            _target('shape_properties', 'actual_top', tolerance=2.0),
            _target('shape_properties', 'width', tolerance=2.0),
            _target('shape_properties', 'height', tolerance=2.0),
        ]
    if operation == 'insert_smartart':
        return [
            _target('smartart_properties', 'found'),
            _target('smartart_properties', 'layout_name'),
            _target('smartart_node_details', 'all_node_texts'),
            _target('smartart_node_details', 'all_node_levels'),
            _target('smartart_placement', 'below_heading'),
        ]
    if operation == 'insert_equation':
        return [_target('equation_structure', 'equation_count')]
    if operation == 'insert_textbox':
        return [
            _target('multi_shape_properties', 'actual_left', tolerance=2.0),
            _target('multi_shape_properties', 'actual_top', tolerance=2.0),
            _target('multi_shape_properties', 'width', tolerance=2.0),
            _target('multi_shape_properties', 'height', tolerance=2.0),
        ]

    if operation == 'table_style':
        return _targets_for_object_effect_slot('table_style', value)
    if operation == 'table_cell_align':
        return _targets_for_object_effect_slot('cell_align', 'on')
    if operation == 'table_border_style':
        return _targets_for_object_effect_slot('border_style', value)
    if operation == 'table_border_color':
        return [_target('table_properties', 'borders')]
    if operation == 'table_text_align':
        return [_target('table_properties', 'cell_paragraph_alignments')]
    if operation == 'chart_title':
        return [
            _target('saved_chart_properties', 'has_title'),
            _target('saved_chart_properties', 'title_text'),
            _target('saved_chart_properties', 'title_include_in_layout'),
        ]
    if operation == 'chart_legend':
        return [
            _target('saved_chart_properties', 'has_legend'),
            _target('saved_chart_properties', 'legend_position'),
        ]
    if operation == 'chart_trendline':
        return [_target('saved_chart_properties', 'series_trendlines')]
    if operation == 'chart_data_labels':
        return [_target('saved_chart_properties', 'series_data_labels')]
    if operation == 'chart_gridlines':
        if value == ('value', 'minor'):
            return [_target('chart_properties', 'value_axis_minor_gridlines')]
        return _targets_for_object_effect_slot('gridlines', value)
    if operation == 'picture_shadow':
        return _targets_for_object_effect_slot('pic_shadow', 'outer_center')
    if operation == 'picture_reflection':
        return _targets_for_object_effect_slot('pic_reflection', 'half_4pt_offset')
    if operation == 'picture_glow':
        return _targets_for_object_effect_slot('pic_glow', 'glow_11pt_accent_4')
    if operation == 'shape_fill':
        return _targets_for_object_effect_slot('fill_color', 'red')
    if operation == 'shape_shadow':
        return _targets_for_object_effect_slot('shadow', 'outer_center')
    if operation == 'shape_bevel':
        return _targets_for_object_effect_slot('bevel', 'angle')
    if operation == 'shape_glow':
        return _targets_for_object_effect_slot('glow', 'glow_11pt_accent_4')
    if operation == 'shape_outline_dash':
        return [_target('shape_properties', 'line_dash_style')]
    if operation == 'smartart_style':
        return _targets_for_object_effect_slot('sa_style', value)
    if operation == 'smartart_colors':
        return _targets_for_object_effect_slot('sa_colors', value)
    if operation == 'smartart_layout':
        return [_target('smartart_properties', 'layout_name')]
    if operation.startswith('equation_'):
        return [_target('equation_structure', 'equations')]
    if operation == 'textbox_direction':
        return [_target('shape_properties', 'text_frame_orientation')]
    if operation == 'textbox_text_fill':
        return [
            _target('shape_properties', 'text_font_color'),
            _target('shape_properties', 'text_fill_color'),
            _target('shape_properties', 'text_fill_type'),
        ]
    if operation == 'textbox_outline_color':
        return _targets_for_object_effect_slot('outline_color', 'blue')
    return []


def atomic_case_for_content_formatting_slot(slot: str, value: Any) -> Optional[str]:
    if slot == 'list_type' and isinstance(value, tuple):
        return LIST_TYPE_ATOMIC_CASES.get(value)
    return CONTENT_FORMATTING_ATOMIC_CASES.get((slot, value))


def build_content_formatting_atomic_intents(
    slots: Dict[str, Any],
    target_comparisons: Optional[Dict[str, Dict[str, str]]] = None,
) -> List[Intent]:
    """Build formatting intents with optional instruction-level comparisons.

    A relative Word command can implement either a directional instruction
    ("increase indentation") or an exact one ("increase to 1.27 cm"). The
    variant supplies an override for the latter so the spec preserves that
    semantic distinction.
    """
    intents: List[Intent] = []
    target_comparisons = target_comparisons or {}
    for index, (slot, value) in enumerate(slots.items(), start=1):
        targets = _targets_for_slot(slot, value)
        slot_comparisons = target_comparisons.get(slot, {})
        for target in targets:
            comparison = slot_comparisons.get(target.get('property'))
            if comparison is not None:
                target['comparison'] = comparison
        intents.append({
            'id': f'intent_{index:02d}',
            'operation': slot,
            'atomic_case': atomic_case_for_content_formatting_slot(slot, value),
            'verification_targets': targets,
        })
    return intents


def atomic_case_for_object_effect_slot(slot: str, value: Any) -> Optional[str]:
    return OBJECT_EFFECT_ATOMIC_CASES.get((slot, value))


def build_object_effect_atomic_intents(slots: Dict[str, Any]) -> List[Intent]:
    intents: List[Intent] = []
    for index, (slot, value) in enumerate(slots.items(), start=1):
        intents.append({
            'id': f'intent_{index:02d}',
            'operation': slot,
            'atomic_case': atomic_case_for_object_effect_slot(slot, value),
            'verification_targets': _targets_for_object_effect_slot(slot, value),
        })
    return intents


def atomic_case_for_content_operation(operation: str, value: Any) -> Optional[str]:
    return CONTENT_OPERATIONS_ATOMIC_CASES.get((operation, value))


def build_content_operations_atomic_intents(
    operations: Iterable[Tuple[str, Any]],
) -> List[Intent]:
    intents: List[Intent] = []
    for index, (operation, value) in enumerate(operations, start=1):
        intents.append({
            'id': f'intent_{index:02d}',
            'operation': operation,
            'atomic_case': atomic_case_for_content_operation(operation, value),
            'verification_targets': _targets_for_content_operation(operation, value),
        })
    return intents


def _scope_changing_operation_key(step: Dict[str, Any]) -> Tuple[str, Any]:
    op = step.get('operation')
    params = step.get('params') or {}

    if op == 'insert_row_above':
        return ('insert_row_above', 'apply')
    if op == 'insert_row':
        return ('insert_row', params.get('position'))
    if op == 'insert_column':
        return ('insert_column', params.get('position'))
    if op in ('merge_cells', 'split_cells', 'split_table'):
        return (op, 'apply')
    if op == 'style_table':
        return ('style_table', params.get('style_name'))
    if op in ('distribute_rows', 'distribute_columns'):
        return (op, 'apply')
    if op in ('fill_row', 'fill_cell'):
        return (op, 'typing')
    if op in ('shade_row', 'shade_column'):
        return (op, params.get('color_rgb'))
    if op == 'align_cell':
        return ('align_cell', (params.get('vertical'), params.get('horizontal')))
    if op == 'align_row':
        return ('align_row', (params.get('vertical'), params.get('horizontal')))
    if op == 'align_paragraph':
        return ('align_paragraph', params.get('alignment'))
    if op in ('add_smartart_node', 'add_smartart_child'):
        return (op, params.get('position'))
    if op in (
        'promote_node',
        'demote_node',
        'reorder_node_up',
        'reorder_node_down',
        'add_drop_lines',
    ):
        return (op, 'apply')
    if op == 'change_chart_type':
        return ('change_chart_type', params.get('chart_type'))
    if op == 'add_chart_title':
        return ('add_chart_title', params.get('position'))
    if op == 'format_chart_title':
        return ('format_chart_title', 'typing_and_bold')
    if op == 'insert_fraction':
        return ('insert_fraction', params.get('frac_type'))
    if op == 'insert_integral_in_numerator':
        return ('insert_integral_in_numerator', params.get('char'))
    if op == 'insert_delimiter':
        return ('insert_delimiter', (params.get('beg_char'), params.get('end_char')))
    if op == 'insert_matrix_in_delimiter':
        return ('insert_matrix_in_delimiter', (params.get('rows'), params.get('cols')))
    if op in ('apply_bullets', 'apply_numbering'):
        return (op, (params.get('gallery'), params.get('template')))
    if op in ('set_list_level', 'set_list_level_range'):
        return (op, params.get('level'))
    return (op or 'unknown', None)


def atomic_case_for_scope_changing_step(step: Dict[str, Any]) -> Optional[str]:
    return SCOPE_CHANGING_ATOMIC_CASES.get(_scope_changing_operation_key(step))


def _targets_for_scope_changing_step(step: Dict[str, Any]) -> List[Target]:
    operation, _value = _scope_changing_operation_key(step)
    params = step.get('params') or {}
    table = 'table_properties'

    if operation in ('insert_row_above', 'insert_row'):
        return [_target(table, 'row_count')]
    if operation == 'insert_column':
        return [_target(table, 'column_count')]
    if operation == 'merge_cells':
        return [
            _target(table, 'row_count'),
            _target(table, 'column_count'),
            _target(table, 'cell_contents'),
        ]
    if operation == 'split_cells':
        return [
            _target(table, 'row_count'),
            _target(table, 'column_count'),
            _target(table, 'cell_contents'),
        ]
    if operation == 'split_table':
        return [_target(table, 'table_count'), _target(table, 'row_count')]
    if operation == 'style_table':
        table_index = params.get('table_index')
        if table_index:
            return [_target(table, 'style', table_index=table_index)]
        return [_target(table, 'style')]
    if operation == 'distribute_rows':
        table_index = params.get('table_index')
        if table_index:
            return [_target(table, 'row_heights', table_index=table_index)]
        return [_target(table, 'row_heights')]
    if operation == 'distribute_columns':
        table_index = params.get('table_index')
        if table_index:
            return [_target(table, 'column_widths', table_index=table_index)]
        return [_target(table, 'column_widths')]
    if operation in ('fill_row', 'fill_cell'):
        return [_target(table, 'cell_contents')]
    if operation in ('shade_row', 'shade_column'):
        return [_target(table, 'cell_shading_colors')]
    if operation in ('align_cell', 'align_row'):
        return [
            _target(table, 'cell_alignments'),
            _target(table, 'cell_paragraph_alignments'),
        ]
    if operation == 'align_paragraph':
        return [_target(table, 'cell_paragraph_alignments')]

    if operation in (
        'add_smartart_node',
        'add_smartart_child',
        'promote_node',
        'demote_node',
        'reorder_node_up',
        'reorder_node_down',
    ):
        return [
            _target('smartart_node_details', 'all_node_count'),
            _target('smartart_node_details', 'all_node_texts'),
            _target('smartart_node_details', 'all_node_levels'),
            _target('smartart_node_details', 'node_texts'),
            _target('smartart_node_details', 'node_levels'),
        ]

    if operation == 'change_chart_type':
        return [_target('chart_properties', 'chart_type')]
    if operation == 'add_drop_lines':
        return [_target('chart_properties', 'has_drop_lines')]
    if operation == 'add_chart_title':
        return [
            _target('chart_properties', 'has_title'),
            _target('chart_properties', 'title_include_in_layout'),
        ]
    if operation == 'format_chart_title':
        targets = [_target('chart_properties', 'title_text')]
        if params.get('bold') is not None:
            targets.append(_target('chart_properties', 'title_font_bold'))
        return targets
    if operation.startswith('insert_') and operation not in (
        'insert_row',
        'insert_row_above',
        'insert_column',
    ):
        return [
            _target('equation_structure', 'function_count'),
            _target('equation_structure', 'functions'),
            _target('equation_structure', 'equation_text'),
        ]

    if operation in ('apply_bullets', 'apply_numbering'):
        return [_target('all_paragraphs_formatting', 'list_string')]
    if operation in ('set_list_level', 'set_list_level_range'):
        return [_target('all_paragraphs_formatting', 'list_level')]
    return []


def build_scope_changing_atomic_intents(
    steps: Iterable[Dict[str, Any]],
) -> List[Intent]:
    intents: List[Intent] = []
    for index, step in enumerate(steps, start=1):
        operation, _value = _scope_changing_operation_key(step)
        intents.append({
            'id': f'intent_{index:02d}',
            'operation': operation,
            'atomic_case': atomic_case_for_scope_changing_step(step),
            'verification_targets': _targets_for_scope_changing_step(step),
        })
    return intents


def build_prerequisite_chain_atomic_intents(
    steps: Iterable[Dict[str, Any]],
) -> List[Intent]:
    intents: List[Intent] = []
    for index, step in enumerate(steps, start=1):
        operation, _value = _prerequisite_operation_key(step)
        intents.append({
            'id': f'intent_{index:02d}',
            'operation': operation,
            'atomic_case': atomic_case_for_prerequisite_step(step),
            'verification_targets': _targets_for_prerequisite_step(step),
        })
    return intents


def _layout_align_command(align_type: Any) -> Optional[str]:
    return {
        0: 'ObjectsAlignLeftSmart',
        1: 'ObjectsAlignCenterHorizontalSmart',
        2: 'ObjectsAlignRightSmart',
        3: 'ObjectsAlignTopSmart',
        4: 'ObjectsAlignMiddleVerticalSmart',
        5: 'ObjectsAlignBottomSmart',
    }.get(align_type)


def _layout_distribute_command(distribute_type: Any) -> Optional[str]:
    return {
        0: 'AlignDistributeHorizontally',
        1: 'AlignDistributeVertically',
    }.get(distribute_type)


def _layout_operation_key(step: Dict[str, Any]) -> Tuple[str, Any]:
    op = step.get('operation')
    params = step.get('params') or {}

    if op in ('rotate_shape', 'rotate_group'):
        return ('rotate_angle', params.get('angle'))
    if op == 'flip_group':
        direction = 'horizontal' if params.get('direction') == 0 else 'vertical'
        return ('flip_direction', direction)
    if op == 'align_shapes':
        return ('align_command', _layout_align_command(params.get('align_type')))
    if op == 'distribute_shapes':
        return (
            'distribute_command',
            _layout_distribute_command(params.get('distribute_type')),
        )
    if op == 'group_all':
        return ('group_command', 'ObjectsGroup')
    if op == 'ungroup':
        return ('ungroup_command', 'ObjectsUngroup')
    if op == 'set_wrap':
        return ('wrap_type', params.get('wrap_type'))
    if op == 'set_position':
        return (
            'position_preset',
            (params.get('h_position'), params.get('v_position')),
        )
    if op == 'resize_shape':
        return ('resize_shape', (params.get('width'), params.get('height')))
    if op == 'move_shape':
        return ('move_shape', (params.get('left'), params.get('top')))
    return (op or 'unknown', None)


def build_layout_coupled_atomic_intents(
    steps: Iterable[Dict[str, Any]],
) -> List[Intent]:
    intents: List[Intent] = []
    for index, step in enumerate(steps, start=1):
        operation, value = _layout_operation_key(step)
        targets = _targets_for_object_effect_slot(operation, value)
        shape_index = (step.get('params') or {}).get('shape_index')
        if shape_index and operation == 'position_preset':
            targets = [dict(target, shape_index=shape_index) for target in targets]
        intents.append({
            'id': f'intent_{index:02d}',
            'operation': operation,
            'atomic_case': atomic_case_for_object_effect_slot(operation, value),
            'verification_targets': targets,
        })
    return intents


def referenced_atomic_cases(intents: Iterable[Intent]) -> List[str]:
    return [
        intent['atomic_case']
        for intent in intents
        if intent.get('atomic_case') is not None
    ]
