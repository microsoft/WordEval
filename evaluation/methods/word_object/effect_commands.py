"""Object-effect command identities and comparison constants."""

from .families.charts.commands import EFFECT_COMMANDS as CHART_EFFECT_COMMANDS

WORDPROCESSING_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

WORDPROCESSING_SHAPE_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"

WORDPROCESSING_GROUP_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"

WORDPROCESSINGML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

MARKUP_COMPATIBILITY_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

OFFICE_MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

PICTURE_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"

EMU_PER_POINT = 12700.0

ROTATION_UNITS_PER_DEGREE = 60000.0

WRAP_POLYGON_COORD_MAX = 21600

WRAP_POLYGON_POINT_TOLERANCE = 500

ALIGNMENT_COMMANDS = {
    "ObjectsAlignLeftSmart": "left",
    "ObjectsAlignRightSmart": "right",
    "ObjectsAlignTopSmart": "top",
    "ObjectsAlignBottomSmart": "bottom",
    "ObjectsAlignCenterHorizontalSmart": "center_h",
    "ObjectsAlignMiddleVerticalSmart": "center_v",
}

DISTRIBUTION_COMMANDS = {
    "AlignDistributeHorizontally": "horizontal",
    "AlignDistributeVertically": "vertical",
}

GROUP_COMMANDS = {"ObjectsGroup"}

UNGROUP_COMMANDS = {"ObjectsUngroup"}

Z_ORDER_COMMANDS = {
    "ObjectBringToFront",
    "ObjectSendToBack",
    "ObjectBringForward",
    "ObjectSendBackward",
}

ROTATION_COMMANDS = {"ObjectRotateGallery"}

POSITION_COMMANDS = {"PicturePositionGallery"}

WRAP_COMMANDS = {"TextWrapGallery"}

TABLE_COMMANDS = {
    "TableRowsDelete",
    "TableColumnsDelete",
    "TableRowsInsertAboveWord",
    "TableRowsInsertBelowWord",
    "TableColumnsInsertLeft",
    "TableColumnsInsertRight",
    "MergeCells",
    "SplitCells",
    "TableSplitTable",
    "TableDelete",
    "TableAutoFitContents",
    "TableAutoFitWindow",
    "TableAutoFitFixedColumnWidth",
    "TableCellAlignTopLeft",
    "TableCellAlignTopCenter",
    "TableCellAlignTopRight",
    "TableCellAlignMiddleLeft",
    "TableCellAlignMiddleCenter",
    "TableCellAlignMiddleRight",
    "TableCellAlignBottomLeft",
    "TableCellAlignBottomCenter",
    "TableCellAlignBottomRight",
    "TableColumnsDistribute",
    "TableRowsDistribute",
    "TableStyleClear",
    "TableRepeatHeaderRows",
    "TableStylesGalleryWord",
    "TableDrawBorderPenWeight",
    "OutlineWeightGallery",
    "BorderStylesGallery",
    "BorderColorPicker",
    "TableDrawBorderPenStyle",
}

DESTRUCTIVE_STRUCTURED_COMMANDS = {"TableDelete"}

TRANSPARENCY_COMMANDS = {"PictureTransparencyGallery"}

PICTURE_CROP_COMMANDS = {
    "PictureFillCrop",
    "PictureFitCrop",
    "PictureCropAspectRatio1To1",
    "PictureCropAspectRatio2To3",
    "PictureCropAspectRatio3To2",
    "PictureCropAspectRatio3To4",
    "PictureCropAspectRatio4To3",
    "PictureCropAspectRatio3To5",
    "PictureCropAspectRatio5To3",
    "PictureCropAspectRatio4To5",
    "PictureCropAspectRatio5To4",
    "PictureCropAspectRatio16To9",
    "PictureCropAspectRatio16To10",
}

PICTURE_FORMAT_COMMANDS = {"PictureReset", "PictureResetAndSize"}

ALT_TEXT_COMMANDS = {"AlternativeText"}

DOCINFO_COMMANDS = {"DocInfoFileName", "DocInfoAuthor", "DocInfoDocTitle"}

EQUATION_CONVERSION_COMMANDS = {
    "EquationNormalText",
    "EquationProfessionalAll",
    "EquationProfessionalOne",
    "EquationLinearAll",
    "EquationLinearOne",
}

PICTURE_SMARTART_COMMANDS = {"PicturesConvertToSmartArt"}

SHADOW_COMMANDS = {
    "ShadowOnOrOffClassic",
    "ShadowSemitransparentClassic",
    "ShadowNudgeUpClassic",
    "ShadowNudgeDownClassic",
    "ShadowNudgeLeftClassic",
    "ShadowNudgeRightClassic",
}

TEXT_ALIGN_COMMANDS = {
    "TextAlignLeft",
    "TextAlignCenter",
    "TextAlignRight",
    "TextAlignLetterJustify",
    "TextAlignWordJustify",
    "TextAlignStretchJustify",
    "TextDirection",
    "TextDirectionGallery",
    "TextAlignGallery",
}

ALIGNMENT_COLLECTION_METRICS = {
    "ObjectsAlignLeftSmart": ("horizontal", "left_edge_spread"),
    "ObjectsAlignCenterHorizontalSmart": ("horizontal", "horizontal_center_spread"),
    "ObjectsAlignRightSmart": ("horizontal", "right_edge_spread"),
    "ObjectsAlignTopSmart": ("vertical", "top_edge_spread"),
    "ObjectsAlignMiddleVerticalSmart": ("vertical", "vertical_center_spread"),
    "ObjectsAlignBottomSmart": ("vertical", "bottom_edge_spread"),
}

DISTRIBUTION_COLLECTION_METRICS = {
    "AlignDistributeHorizontally": ("horizontal", "horizontal_center_gap_spread"),
    "AlignDistributeVertically": ("vertical", "vertical_center_gap_spread"),
}

SHAPE_COLLECTION_PROPERTIES = {
    metric
    for _, metric in list(ALIGNMENT_COLLECTION_METRICS.values())
    + list(DISTRIBUTION_COLLECTION_METRICS.values())
}

WORDART_COMMANDS = {
    "WordArtEvenTextHeightClassic",
    "WordArtSpacingKernCharacterPairs",
    "WordArtSpacingLoose",
    "WordArtSpacingNormal",
    "WordArtSpacingTight",
    "WordArtSpacingVeryLoose",
    "WordArtSpacingVeryTight",
    "WordArtVerticalText",
}

THREED_COMMANDS = {
    "_3DEffectsOnOffClassic",
    "_3DTiltDownClassic",
    "_3DTiltUpClassic",
    "_3DTiltLeftClassic",
    "_3DTiltRightClassic",
    "_3DExtrusionPerspectiveClassic",
    "_3DExtrusionParallelClassic",
}

FILL_OUTLINE_COLOR_COMMANDS = {
    "ShapeFillColorPicker",
    "ShapeOutlineColorPicker",
    "GlowColorPicker",
    "GraphicsColorPicker",
    "GraphicsOutlineColorPicker",
}

SHAPE_EFFECT_COMMANDS = {
    "ObjectEffectShadowGallery",
    "ObjectEffectGlowGallery",
    "ObjectEffectSoftEdgesGallery",
    "ObjectEffectPresetGallery",
    "ReflectionGallery",
    "GradientGallery",
    "ShapeStylesGallery",
}

SHAPE_GEOMETRY_COMMANDS = {"ShapeChangeShapeGallery", "_3DRotationGallery", "BevelShapeGallery"}

SMARTART_SUBNODE_COMMANDS = {"ArrowStyleGallery", "OutlineDashesGallery", "OutlineWeightGallery"}

PICTURE_EFFECT_COMMANDS = {
    "PictureArtisticEffectsGallery",
    "PictureColorTempertatureGallery",
    "PictureEffectsGlowGallery",
    "PictureEffectsShadowGallery",
    "PictureEffectsSoftEdgesGallery",
    "PictureRecolorGallery",
    "PictureReflectionGallery",
    "PictureRotationGallery",
    "PictureSaturationGallery",
    "PictureShapeGallery",
    "PictureSharpenSoftenGallery",
    "_3DBevelPictureTopGallery",
    "RecolorColorPicker",
    "OutlineColorPicker",
}

POSITIVE_PICTURE_EFFECT_COMMANDS = frozenset(
    {
        "PictureArtisticEffectsGallery",
        "PictureColorTempertatureGallery",
        "PictureSaturationGallery",
        "PictureSharpenSoftenGallery",
    }
)

POSITIVE_PICTURE_EFFECT_PARAM_TYPES = {
    "PictureColorTempertatureGallery": 7,
    "PictureSaturationGallery": 24,
    "PictureSharpenSoftenGallery": 25,
}

CHART_COMMANDS_SET = CHART_EFFECT_COMMANDS

EQUATION_COMMANDS_SET = {
    "EquationAccentGallery",
    "EquationDelimiterGallery",
    "EquationFractionGallery",
    "EquationFunctionGallery",
    "EquationIntegralGallery",
    "EquationLargeOperatorGallery",
    "EquationLimitGallery",
    "EquationMatrixGallery",
    "EquationOperatorGallery",
    "EquationRadicalGallery",
    "EquationScriptGallery",
    "EquationSymbolsInsertGallery",
}

SMARTART_COMMANDS_SET = {
    "SmartArtLayoutGallery",
    "SmartArtStylesGallery",
    "SmartArtChangeColorsGallery",
}

SMARTART_STRUCTURAL_COMMANDS = {
    "SmartArtAddShapeAfter",
    "SmartArtAddShapeBefore",
    "SmartArtAddShapeAbove",
    "SmartArtAddShapeBelow",
    "SmartArtAddAssistant",
    "SmartArtPromote",
    "SmartArtDemote",
    "SmartArtReorderUp",
    "SmartArtReorderDown",
    "SmartArtLargerShape",
    "SmartArtSmallerShape",
    "SmartArtRightToLeft",
    "SmartArtResetGraphic",
}

SMARTART_OPERATION_COMMANDS = {
    "add_smartart_node": "SmartArtAddShapeAfter",
    "add_smartart_child": "SmartArtAddShapeBelow",
    "promote_node": "SmartArtPromote",
    "demote_node": "SmartArtDemote",
    "reorder_node_up": "SmartArtReorderUp",
    "reorder_node_down": "SmartArtReorderDown",
}

SMARTART_TEXT_EFFECT_COMMANDS = {
    "TextEffectTransformGallery",
    "TextEffects3DRotationGallery",
    "TextFillColorPicker",
    "BevelTextGallery",
}

TEXT_EFFECT_COMMANDS = SMARTART_TEXT_EFFECT_COMMANDS | {"TextAlignGallery"}

PRESET_POSITION_THRESHOLD = -900000.0

PRESET_VALUE_MAP = {
    -999998.0: "start",
    -999999.0: "start",
    -999995.0: "center",
    -999996.0: "end",
    -999997.0: "end",
}
