"""Complete, applicable native observations for margin page-number shapes."""
import re

from ...observations import ObservedProperties, require_observations
from ..formatting.characters import character_property
from ..pages.field_text import page_field_template
from ..references.fields import required_style_name, INSTRUCTION_ONLY_FIELD_TYPES


FONT_FIELDS = {"name", "size", "bold", "italic", "color", "color_index", "spacing"}
TEXT_SHAPES = {1, 5, 17}


def office_enum(value):
    if value == -2:
        raise ValueError("Required Office enum is mixed")
    return value


def extract_margin_shape_styles(doc, *, include_group_items=False):
    """Retain failed items without losing independent native observations."""
    def font_signature(rng):
        state = ObservedProperties()
        for name, attribute in {"name": "Name", "size": "Size", "bold": "Bold",
                                "italic": "Italic", "color": "Color", "color_index": "ColorIndex",
                                "spacing": "Spacing"}.items():
            def normalize(value, name=name):
                if name in {"bold", "italic"}:
                    return bool(value)
                if name == "color" and value in {-16777216, 0}:
                    return 0
                return round(float(value), 2) if name in {"size", "spacing"} else value
            state.read(name, lambda attribute=attribute, name=name: character_property(
                doc, rng, lambda r: getattr(r.Font, attribute), normalize=normalize,
                empty_is_mixed=name == "name"))
        return state

    def shape_state(shape, index, story):
        state = ObservedProperties(index=index, story=story)
        scalars = {"type": "Type", "relative_horizontal_position": "RelativeHorizontalPosition",
                   "relative_vertical_position": "RelativeVerticalPosition", "left": "Left", "top": "Top",
                   "width": "Width", "height": "Height", "rotation": "Rotation"}
        for key, attribute in scalars.items():
            state.read(key, lambda attribute=attribute: getattr(shape, attribute),
                       (lambda v: round(float(v), 2)) if key in {"left", "top", "width", "height", "rotation"} else office_enum)
        if state["type"] == 1:
            state.read("auto_shape_type", lambda: shape.AutoShapeType, office_enum)
        state.read("wrap_type", lambda: shape.WrapFormat.Type, office_enum)
        for prefix, parent, attributes in (
            ("fill", "Fill", {"type": "Type", "transparency": "Transparency"}),
            ("line", "Line", {"style": "Style", "dash_style": "DashStyle", "weight": "Weight", "transparency": "Transparency"}),
        ):
            if state["type"] == 6:
                continue  # Children own group appearance; parent aggregates may be mixed.
            state.read(prefix + "_visible", lambda parent=parent: getattr(shape, parent).Visible, bool)
            # Inactive effects have no required color/style. Their visibility
            # remains a required observation, so a failed getter is not absence.
            if state[prefix + "_visible"] != 0:
                for key, attribute in attributes.items():
                    state.read(prefix + "_" + key, lambda parent=parent, attribute=attribute: getattr(getattr(shape, parent), attribute),
                               office_enum if key in {"type", "style", "dash_style"} else (lambda v: v))
                state.read(prefix + "_color", lambda parent=parent: getattr(shape, parent).ForeColor.RGB)
                state.read(prefix + "_theme_color", lambda parent=parent: getattr(shape, parent).ForeColor.ObjectThemeColor, office_enum)

        state.update({key: [] for key in ("paragraph_styles", "paragraph_texts", "paragraph_templates",
                                         "paragraph_alignments", "paragraph_fonts", "field_codes", "field_keywords", "field_fonts")})
        state.update(page_field_count=0, num_pages_field_count=0)
        if state["type"] in TEXT_SHAPES:
            state.read("text_orientation", lambda: shape.TextFrame.Orientation, office_enum)
            for side in ("Left", "Right", "Top", "Bottom"):
                state.read("text_margin_" + side.lower(), lambda side=side: getattr(shape.TextFrame, "Margin" + side),
                           lambda v: round(float(v), 2))
            # HasText and collection failures cannot silently remove paragraphs
            # or PAGE fields from an otherwise matching shape.
            state.read("has_text", lambda: shape.TextFrame.HasText, bool)
            if state["has_text"]:
                tr = shape.TextFrame.TextRange
                for j in range(1, tr.Paragraphs.Count + 1):
                    para = tr.Paragraphs(j)
                    rng = para.Range
                    fields = [rng.Fields(i) for i in range(1, rng.Fields.Count + 1)]
                    state["paragraph_styles"].append(required_style_name(para.Style))
                    state["paragraph_texts"].append(rng.Text.rstrip("\r").strip())
                    state["paragraph_templates"].append(page_field_template(doc, rng, fields).strip())
                    state["paragraph_alignments"].append(para.Format.Alignment)
                    state["paragraph_fonts"].append(font_signature(rng))
                for j in range(1, tr.Fields.Count + 1):
                    field = tr.Fields(j)
                    code = field.Code.Text.strip()
                    match = re.match(r"\s*([A-Z]+)\b", code.upper())
                    keyword = match.group(1) if match else ""
                    state["field_codes"].append(code)
                    state["field_keywords"].append(keyword)
                    if field.Type not in INSTRUCTION_ONLY_FIELD_TYPES:
                        state["field_fonts"].append(font_signature(field.Result))
                    if keyword == "PAGE":
                        state["page_field_count"] += 1
                    elif keyword == "NUMPAGES":
                        state["num_pages_field_count"] += 1
        return state

    result = []
    def append(shape, index, story):
        result.append(shape_state(shape, index, story))
        if include_group_items and shape.Type == 6:
            for j in range(1, shape.GroupItems.Count + 1):
                append_item(shape.GroupItems, j, f"{index}.{j}", story)

    def append_item(collection, native_index, index, story):
        try:
            shape = collection(native_index)
        except Exception as exc:
            state = ObservedProperties(index=index, story=story)
            state.record_unavailable("collection_item", exc)
            result.append(state)
            return
        append(shape, index, story)

    for i in range(1, doc.Shapes.Count + 1):
        append_item(doc.Shapes, i, str(i), "body")
    for section_index in range(1, doc.Sections.Count + 1):
        section = doc.Sections(section_index)
        for collection_name, story in (("Headers", "header"), ("Footers", "footer")):
            collection = getattr(section, collection_name)
            for part_type in (1, 2, 3):
                part = collection(part_type)
                if not part.Exists:
                    continue
                shapes = part.Shapes
                for i in range(1, shapes.Count + 1):
                    append_item(shapes, i, f"{story}:{section_index}:{part_type}:{i}", story)
    return result


def validate_margin_shape_styles(states):
    for owner, shapes in states.items():
        for shape in shapes:
            if not isinstance(shape, ObservedProperties):
                continue  # Legacy focused fixtures; native reader always tags observations.
            error = require_observations({owner: shape}, set(shape) - {"index", "story"},
                                         context="margin page-number shape")
            if error is not None:
                return error
            for alignment in shape.get("paragraph_alignments", []):
                error = require_observations({owner: {"alignment": alignment}}, {"alignment"},
                                             context="margin page-number paragraph")
                if error is not None:
                    return error
            for key in ("paragraph_fonts", "field_fonts"):
                for font in shape[key]:
                    error = require_observations({owner: font}, FONT_FIELDS, context="margin page-number font")
                    if error is not None:
                        return error
    return None
