"""families / formatting / extract_com with explicit evidence dependencies."""

from evaluation.methods.word_object.families.formatting import (
    paragraph_com as families_formatting_paragraph_com,
)
from typing import Dict, List, Optional, Any
from ...observations import ObservedProperties


def explicit_path_gradient(rng):
    """Prove that every selected text run has a native non-linear gradient.

    GradientAngle is inapplicable to a path gradient. Only Word's serialization
    of the exact selected text can disambiguate it from a failed linear read;
    unused styles and text outside the selected range cannot supply evidence.
    """
    from xml.etree import ElementTree as ET
    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    w14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
    root = ET.fromstring(rng.WordOpenXML)
    body = root.find(f".//{w}body")
    if body is None:
        return False
    runs = [run for run in body.iter(w + "r") if run.find(w + "t") is not None]
    if not runs or "".join(t.text or "" for run in runs for t in run.findall(w + "t")) != rng.Text:
        return False
    for run in runs:
        fill = run.find(f"{w}rPr/{w14}textFill/{w14}gradFill")
        if fill is None or fill.find(w14 + "path") is None or fill.find(w14 + "lin") is not None:
            return False
    return True


def hidden_text_ranges(doc, rng):
    """Observe hidden character intervals, independent of XML run boundaries.

    Split only mixed ranges. Uniform paragraphs take one property read; a
    partially hidden word cannot disappear inside a paragraph-wide mixed value.
    Offsets are relative to the supplied range, not absolute document positions.
    """
    if rng.End <= rng.Start:
        return []
    from .characters import character_property, text_width

    def visibility(value):
        if value not in (0, -1, 1):
            raise ValueError("Character visibility could not be observed")
        return bool(value)

    observed = character_property(doc, rng, lambda r: r.Font.Hidden, normalize=visibility)
    if isinstance(observed, list):
        return [(start, end) for start, end, hidden in observed if hidden]
    return [(0, text_width(rng.Text))] if observed else []


def extract_text_formatting(
    docx_path: str, target_text: str, *, close_document, get_word, open_document
) -> Optional[Dict]:
    """Read effective target formatting; None means the target is absent.

    Failed properties retain explicit unavailable diagnostics. Inactive properties
    may be observed None, while numeric zero (black, no effect, zero spacing) is a
    real value. Rendered positions are optional evidence for bidi equivalence.
    """
    doc = None
    try:
        doc = open_document(docx_path)
        rng = doc.Content
        rng.Find.ClearFormatting()
        rng.Find.Execute(FindText=target_text)
        if not rng.Find.Found:
            return None
        font, para = rng.Font, rng.ParagraphFormat
        result = ObservedProperties()
        result.read("body_hidden_ranges", lambda: hidden_text_ranges(doc, doc.Content))
        for name, attr in {
            "bold": "Bold", "italic": "Italic", "strikethrough": "StrikeThrough",
            "subscript": "Subscript", "superscript": "Superscript", "hidden": "Hidden",
        }.items():
            result.read(name, lambda attr=attr: getattr(font, attr), bool)
        for name, attr in {
            "underline": "Underline", "font_size": "Size", "font_name": "Name",
            "font_color": "Color",
        }.items():
            result.read(name, lambda attr=attr: getattr(font, attr))
        result.read("underline_color", lambda: font.UnderlineColor,
                    lambda value: None if value == -16777216 or not result["underline"] else value)

        def character_borders():
            edges = {}
            for name, kind in (("top", 1), ("left", 2), ("bottom", 3), ("right", 4)):
                border = rng.Borders(kind)
                line_style = border.LineStyle
                if line_style == 9999999:
                    raise ValueError(f"mixed {name} border")
                edges[name] = {
                    "line_style": line_style or None,
                    "line_width": border.LineWidth if line_style else None,
                    "color": border.Color if line_style else None,
                }
            return edges

        result.read("border", character_borders)
        result.read("shading", lambda: rng.Shading.Texture, lambda value: value or None)
        # Preserve the established clear/white normalization; zero is RGB black.
        result.read("shading_color", lambda: rng.Shading.BackgroundPatternColor,
                    lambda value: None if value in (-16777216, 16777215) else value)
        result.read("highlight_color", lambda: rng.HighlightColorIndex,
                    lambda value: value or None)
        result.read("text_outline_visible", lambda: font.Line.Visible, bool)
        for name, getter, normalize in (
            ("text_outline_color", lambda: font.Line.ForeColor.RGB, lambda v: v),
            ("text_outline_weight", lambda: font.Line.Weight, lambda v: round(v, 2)),
            ("text_outline_dash_style", lambda: font.Line.DashStyle, lambda v: v),
        ):
            if result["text_outline_visible"]:
                result.read(name, getter, normalize)
            else:
                result[name] = None
        result.read("text_fill_type", lambda: font.Fill.Type)
        result["text_fill_color"] = None
        result["text_fill_gradient_angle"] = None
        result["text_fill_gradient_kind"] = None
        if result["text_fill_type"] in (1, 3):
            result.read("text_fill_color", lambda: font.Fill.ForeColor.RGB)
        if result["text_fill_type"] == 3:
            result.read("text_fill_gradient_angle", lambda: font.Fill.GradientAngle,
                        lambda value: round(value, 1))
            if "text_fill_gradient_angle" not in result.unavailable:
                result["text_fill_gradient_kind"] = "linear"
            else:
                result.read("text_fill_gradient_kind", lambda: "path" if explicit_path_gradient(rng) else None)
                if result["text_fill_gradient_kind"] == "path":
                    result.read("text_fill_gradient_angle", lambda: None)
        result.read("text_reflection_type", lambda: font.Reflection.Type)
        result.read("text_glow_radius", lambda: font.Glow.Radius,
                    lambda value: round(value, 2) if value else None)
        result.read("dropcap_position", lambda: rng.Paragraphs(1).DropCap.Position)
        result.read("dropcap_lines_to_drop", lambda: rng.Paragraphs(1).DropCap.LinesToDrop)
        result.read("paragraph_borders", lambda:
            families_formatting_paragraph_com.extract_paragraph_borders(rng.Paragraphs(1)))
        for name, attr in {
            "alignment": "Alignment", "left_indent": "LeftIndent",
            "right_indent": "RightIndent", "first_line_indent": "FirstLineIndent",
            "line_spacing": "LineSpacing", "space_before": "SpaceBefore",
            "space_after": "SpaceAfter",
        }.items():
            result.read(name, lambda attr=attr: getattr(para, attr), lambda value: round(value, 2))

        # Selection can invalidate COM proxies, so read all formatting first.
        result["rendered_first_h"] = result["rendered_last_h"] = None
        try:
            word = get_word()
            para_range = rng.Paragraphs(1).Range
            doc.Range(para_range.Start, para_range.Start + 1).Select()
            result["rendered_first_h"] = round(word.Selection.Information(5), 1)
            doc.Range(para_range.End - 2, para_range.End - 1).Select()
            result["rendered_last_h"] = round(word.Selection.Information(5), 1)
        except Exception:
            pass
        return result
    finally:
        close_document(doc)


def extract_target_glyph_state(
    docx_path: str, target_text: str, *, close_document, open_document
) -> Dict[str, Any]:
    """Extract target-scoped Word glyph-paint visibility evidence.

    This is intentionally separate from :meth:`extract_text_formatting`.
    The latter is compared field-by-field by many content-formatting
    routes, while this projection is needed only when a text-effect
    composition can preserve every requested property yet make the glyph
    paint fully transparent.  Keeping the method additive avoids changing
    the comparison contract of existing callers.

    The projection does not judge color contrast.  It reports only Word's
    effective hidden flag plus fill/outline visibility, transparency, and
    outline weight after a normal reopen.  Callers must treat unavailable
    properties as indeterminate rather than as visible or invisible.
    """

    def normalized_bool(value: Any) -> Optional[bool]:
        if value in (None, 9999999, -2147483648):
            return None
        try:
            return bool(int(value))
        except (TypeError, ValueError, OverflowError):
            return None

    def normalized_transparency(value: Any) -> Optional[float]:
        try:
            numeric = float(value)
        except (TypeError, ValueError, OverflowError):
            return None
        if not 0.0 <= numeric <= 1.0:
            return None
        return round(numeric, 4)

    doc = None
    try:
        doc = open_document(docx_path)
        rng = doc.Content.Duplicate
        rng.Find.ClearFormatting()
        rng.Find.Execute(FindText=target_text)
        if not rng.Find.Found:
            return {
                "found": False,
                "target_text": target_text,
            }

        font = rng.Font
        result: Dict[str, Any] = {
            "found": True,
            "target_text": rng.Text,
            "hidden": normalized_bool(getattr(font, "Hidden", None)),
            "font_color": (font.Color if font.Color != 9999999 else None),
            "fill_visible": None,
            "fill_type": None,
            "fill_transparency": None,
            "outline_visible": None,
            "outline_transparency": None,
            "outline_weight": None,
        }

        try:
            fill = font.Fill
            result["fill_visible"] = normalized_bool(fill.Visible)
            try:
                fill_type = int(fill.Type)
                if fill_type != 9999999:
                    result["fill_type"] = fill_type
            except (TypeError, ValueError, OverflowError):
                pass
            result["fill_transparency"] = normalized_transparency(fill.Transparency)
        except Exception:
            pass

        try:
            line = font.Line
            result["outline_visible"] = normalized_bool(line.Visible)
            result["outline_transparency"] = normalized_transparency(line.Transparency)
            try:
                weight = float(line.Weight)
                if 0.0 <= weight < 1000.0:
                    result["outline_weight"] = round(weight, 4)
            except (TypeError, ValueError, OverflowError):
                pass
        except Exception:
            pass

        return result
    finally:
        close_document(doc)


def extract_all_paragraph_font_properties(
    docx_path: str, *, close_document, open_document
) -> List[Dict]:
    """Extract paragraph-level font properties not covered by paragraph formatting."""
    doc = None
    try:
        doc = open_document(docx_path)
        paragraphs = []
        for i in range(1, doc.Paragraphs.Count + 1):
            para_info = {"index": i}
            try:
                para = doc.Paragraphs(i)
                para_text = para.Range.Text
                if para_text.endswith("\r"):
                    para_text = para_text[:-1]
                para_info["text"] = para_text[:50] + ("..." if len(para_text) > 50 else "")
            except:
                para_info["text"] = ""
            try:
                ligatures = doc.Paragraphs(i).Range.Font.Ligatures
                para_info["ligatures"] = ligatures if ligatures != 9999999 else None
            except:
                para_info["ligatures"] = None
            paragraphs.append(para_info)
        return paragraphs
    finally:
        close_document(doc)


def extract_all_paragraph_character_formatting(
    docx_path: str, *, close_document, open_document
) -> List[Dict[str, Any]]:
    """Extract visible character-formatting invariants per body paragraph.

    Paragraph/list commands operate on paragraph properties and list markers;
    they should not silently change the formatting of the paragraph's text.
    Exclude the paragraph/cell end markers before reading ``Range.Font`` so
    paragraph-mark formatting does not create mixed-value noise.
    """
    doc = None
    try:
        doc = open_document(docx_path)
        paragraphs: List[Dict[str, Any]] = []

        from .characters import character_property, character_projection

        def outline_value(rng, attribute):
            line = rng.Font.Line
            if line.Visible in (-2, 9999999, -9999999):
                return 9999999
            if line.Visible == 0:
                return None
            return line.ForeColor.RGB if attribute == "color" else getattr(line, attribute)

        def fill_color(rng):
            fill = rng.Font.Fill
            if fill.Type in (-2, 9999999, -9999999):
                return 9999999
            return fill.ForeColor.RGB if fill.Type in {1, 3, 4} else None

        getters = {
            "bold": lambda r: r.Font.Bold,
            "italic": lambda r: r.Font.Italic,
            "underline": lambda r: r.Font.Underline,
            "strikethrough": lambda r: r.Font.StrikeThrough,
            "subscript": lambda r: r.Font.Subscript,
            "superscript": lambda r: r.Font.Superscript,
            "hidden": lambda r: r.Font.Hidden,
            "font_size": lambda r: r.Font.Size,
            "font_name": lambda r: r.Font.Name,
            "font_color": lambda r: r.Font.Color,
            "underline_color": lambda r: r.Font.UnderlineColor,
            "highlight_color": lambda r: r.HighlightColorIndex,
            "shading_texture": lambda r: r.Shading.Texture,
            "shading_color": lambda r: r.Shading.BackgroundPatternColor,
            "text_outline_visible": lambda r: r.Font.Line.Visible,
            "text_outline_color": lambda r: outline_value(r, "color"),
            "text_outline_weight": lambda r: outline_value(r, "Weight"),
            "text_outline_dash_style": lambda r: outline_value(r, "DashStyle"),
            "text_fill_type": lambda r: r.Font.Fill.Type,
            "text_fill_color": fill_color,
            "text_reflection_type": lambda r: r.Font.Reflection.Type,
            "text_glow_radius": lambda r: r.Font.Glow.Radius,
        }
        from .characters import INACTIVE_CHARACTER_FIELDS
        from .paragraph_ranges import is_row_end, paragraph_text_range
        for index in range(1, doc.Paragraphs.Count + 1):
            para = doc.Paragraphs(index)
            if is_row_end(para):
                continue
            rng = paragraph_text_range(para)
            signature = ObservedProperties(text=rng.Text)
            # Empty paragraphs have no visible characters to preserve.
            if rng.End > rng.Start:
                projection = character_projection(rng)
                for name, getter in getters.items():
                    rounded = name in {"font_size", "text_outline_weight", "text_glow_radius"}
                    signature.read(name, lambda getter=getter, name=name, rounded=rounded:
                        character_property(doc, rng, getter,
                            normalize=(lambda v: round(float(v), 2)) if rounded else (lambda v: v),
                            nullable=name in INACTIVE_CHARACTER_FIELDS,
                            empty_is_mixed=name == "font_name",
                            office_mixed=name.startswith("text_"), projection=projection))
                signature.read("hidden_ranges", lambda: hidden_text_ranges(doc, rng))
            paragraphs.append(signature)

        return paragraphs
    finally:
        close_document(doc)


def extract_body_paragraph_styles(docx_path: str, *, close_document, open_document) -> List[str]:
    """Extract Style.NameLocal (or string fallback) for body paragraphs."""
    doc = None
    try:
        doc = open_document(docx_path)
        styles = []
        for i in range(1, doc.Paragraphs.Count + 1):
            try:
                para = doc.Paragraphs(i)
                style = para.Style
                if hasattr(style, "NameLocal"):
                    styles.append(style.NameLocal if style else "")
                else:
                    styles.append(str(style) if style else "")
            except:
                styles.append("")
        return styles
    finally:
        close_document(doc)


def extract_list_properties(
    docx_path: str, target_text: str = None, *, close_document, open_document
) -> Optional[Dict]:
    """Extract list formatting properties."""
    doc = None
    try:
        doc = open_document(docx_path)

        if target_text:
            rng = doc.Content
            rng.Find.Execute(FindText=target_text)
            if not rng.Find.Found:
                return None
        else:
            rng = doc.Content

        list_fmt = rng.ListFormat

        return {
            "list_type": list_fmt.ListType,
            "list_level": list_fmt.ListLevelNumber,
            "list_string": list_fmt.ListString,
        }

    finally:
        close_document(doc)
