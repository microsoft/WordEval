"""Required target observations for literal-text formatting routes.

These are evidence requirements, not expected values. Values still come from the
case contract/reference and existing scoped equivalences. Inactive effect state
is nullable; a failed read of the same state is not.
"""
from ...observations import ObservedProperties, require_observations
from ....base import VerificationResult


COMMAND_FIELDS = {
    "Bold": ("bold",), "Italic": ("italic",),
    "Strikethrough": ("strikethrough",), "Subscript": ("subscript",),
    "Superscript": ("superscript",),
    "Font": ("font_name",), "FontColorPicker": ("font_color",),
    "FontSize": ("font_size",), "FontSizeIncreaseWord": ("font_size",),
    "FontSizeDecreaseWord": ("font_size",),
    "CharacterBorder": ("border",), "BordersSelectionGallery": ("paragraph_borders",),
    "CharacterShading": ("shading", "shading_color"),
    "ShadingColorPicker": ("shading_color",),
    "TextHighlightColorPicker": ("highlight_color",),
    "UnderlineGallery": ("underline",),
    "UnderlineColorPicker": ("underline", "underline_color"),
    "TextReflectionGallery": ("text_reflection_type",),
    "TextFillGradientGallery": ("text_fill_type", "text_fill_gradient_angle"),
    "TextOutlineWeightGallery": ("text_outline_visible", "text_outline_weight"),
    "TextOutlineColorPickerAlternate": ("text_outline_visible", "text_outline_color"),
    "TextOutlineDashesGallery": ("text_outline_visible", "text_outline_dash_style"),
    "AlignLeft": ("alignment",), "AlignRight": ("alignment",),
    "AlignCenter": ("alignment",), "AlignJustify": ("alignment",),
    "IndentIncreaseWord": ("left_indent",), "IndentDecreaseWord": ("left_indent",),
    "LineSpacingGallery": ("line_spacing",),
    "ParagraphSpacingBefore": ("space_before",),
    "ParagraphSpacingAfter": ("space_after",),
    "ClearFormatting": ("bold", "italic", "underline", "font_name", "font_size", "font_color"),
}

NULLABLE_FIELDS = {
    "shading", "shading_color", "highlight_color", "underline_color",
    "text_outline_color", "text_outline_weight", "text_outline_dash_style",
    "text_fill_gradient_angle", "text_fill_gradient_kind", "text_glow_radius",
}


def required_fields(spec, extractor="text_formatting"):
    fields = set(COMMAND_FIELDS.get(spec.get("idMso"), ()))
    primary = spec.get("verification", {}).get("primary", {})
    name = primary.get("property_name")
    if name and name != "clear_formatting":
        fields.add(name)
    for intent in spec.get("composition", {}).get("atomic_intents", []):
        for target in intent.get("verification_targets", []) or []:
            accepted = {extractor}
            if extractor == "text_formatting":
                accepted.add("paragraph_formatting")
            if target.get("extractor") in accepted and target.get("property"):
                fields.add(target["property"])
    return fields


def validate_formatting_observations(gold, pred, spec, *, ignored=()):
    for owner, state in (("gold", gold), ("pred", pred)):
        if not state:
            return VerificationResult.error_result(
                f"Empty {owner} formatting projection", observation_owner=owner)
    fields = required_fields(spec)
    # Full projections own preservation as well as the requested property.
    # Legacy narrow readers validate only their declared target observations.
    if isinstance(gold, ObservedProperties):
        fields.update(gold)
    elif "hidden" in gold:
        fields.add("hidden")
    fields -= {"rendered_first_h", "rendered_last_h", *ignored}
    for state in (gold, pred):
        if state.get("text_outline_visible"):
            if any(key.startswith("text_outline_") for key in fields):
                fields.add("text_outline_weight")
        if state.get("text_fill_type") == 3 and "text_fill_type" in fields:
            fields.add("text_fill_gradient_angle")
    for owner, state in (("gold", gold), ("pred", pred)):
        nullable = set(NULLABLE_FIELDS)
        # Inactive child values are absent by definition; the observed parent
        # mismatch is a semantic failure, not an extraction error.
        if state.get("text_outline_visible"):
            nullable -= {"text_outline_color", "text_outline_weight", "text_outline_dash_style"}
        if state.get("text_fill_type") == 3:
            nullable.discard("text_fill_gradient_kind")
            if state.get("text_fill_gradient_kind") != "path":
                nullable.discard("text_fill_gradient_angle")
        # Gold is validated first. A uniform reference and positively observed
        # mixed prediction disagree, including an observed absent gold effect.
        # Mixed gold itself needs a finer-grained projection and cannot pass.
        mixed = set(getattr(state, "mixed", ())) & fields
        if owner == "pred" and mixed:
            return VerificationResult(
                passed=False, message="FAIL: Target formatting is nonuniform",
                details={"differences": [{"property": name, "gold": gold[name], "pred": "mixed"}
                                         for name in sorted(mixed)]})
        error = require_observations({owner: state}, fields,
                                     context="target text formatting", nullable=nullable)
        if error is not None:
            return error
    return None
