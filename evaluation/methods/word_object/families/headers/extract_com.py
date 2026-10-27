"""Read complete header/footer collections with independent observation errors."""

from typing import Dict
from ...observations import ObservedProperties
from ..pages.field_text import page_field_template, page_field_character_property
from ..formatting.paragraph_ranges import is_row_end, paragraph_text_range
from ..tables.border_xml import border_style


def extract_header_footer_content(docx_path: str, *, close_document, open_document) -> Dict:
    doc = None
    try:
        doc = open_document(docx_path)
        result = {"headers": [], "footers": []}

        def paragraph_signature(para):
            rng = paragraph_text_range(para)
            state = ObservedProperties(text=rng.Text.removesuffix("\r").strip())
            state.read("style", lambda: para.Style.NameLocal if hasattr(para.Style, "NameLocal") else str(para.Style))
            for name, attr in (("alignment", "Alignment"), ("space_before", "SpaceBefore"), ("space_after", "SpaceAfter")):
                state.read(name, lambda attr=attr: getattr(para.Format, attr))
            for name, attr in (("font_name", "Name"), ("font_size", "Size"), ("font_color", "Color"), ("bold", "Bold"), ("italic", "Italic")):
                state.read(name, lambda attr=attr: page_field_character_property(doc, rng,
                    [rng.Fields(i) for i in range(1, rng.Fields.Count + 1)],
                    lambda r: getattr(r.Font, attr), empty_is_mixed=attr == "Name"))
            for kind, name in enumerate(("top", "left", "bottom", "right"), 1):
                state.read(f"{name}_border_line_style", lambda kind=kind, name=name:
                           border_style(para, kind, name, scope="paragraph"))
                for field, attr in (("line_width", "LineWidth"), ("color", "Color")):
                    state.read(f"{name}_border_{field}", lambda kind=kind, name=name, attr=attr:
                        None if state[f"{name}_border_line_style"] == 0 else getattr(para.Borders(-kind), attr))
            return state

        def part_info(part, kind):
            rng = part.Range
            info = ObservedProperties(type=kind)
            info.read("text", lambda: rng.Text.strip())
            info.read("paragraph_signatures", lambda: [paragraph_signature(rng.Paragraphs(i))
                       for i in range(1, rng.Paragraphs.Count + 1) if not is_row_end(rng.Paragraphs(i))])
            info.read("paragraph_count", lambda: len(info["paragraph_signatures"]))
            info.read("paragraph_styles", lambda: [p["style"] for p in info["paragraph_signatures"]])
            info.read("text_template", lambda: page_field_template(doc, rng, [rng.Fields(i) for i in range(1, rng.Fields.Count + 1)]).strip())
            info.read("field_count", lambda: rng.Fields.Count)
            info.read("field_types", lambda: [rng.Fields(i).Type for i in range(1, rng.Fields.Count + 1)])
            info.read("field_codes", lambda: [rng.Fields(i).Code.Text.strip() for i in range(1, rng.Fields.Count + 1)])
            # Field.Type identifies PAGE (33) and NUMPAGES (26); unrelated
            # fields containing the letters PAGE do not become page counters.
            info.read("page_field_count", lambda: info["field_types"].count(33))
            info.read("num_pages_field_count", lambda: info["field_types"].count(26))
            info.read("inline_shape_count", lambda: rng.InlineShapes.Count)
            info.read("shape_count", lambda: part.Shapes.Count)

            def controls():
                values = {}
                def collect(r):
                    for i in range(1, r.ContentControls.Count + 1):
                        cc = r.ContentControls(i)
                        key = cc.Title.strip() or cc.Tag.strip() or f"_pos_{len(values) + 1}"
                        base, occurrence = key, 1
                        while key in values:
                            occurrence += 1
                            key = f"{base}#{occurrence}"
                        values[key] = cc.Range.Text.strip()
                def shape_controls(shape):
                    if shape.Type == 6:  # msoGroup
                        for i in range(1, shape.GroupItems.Count + 1):
                            shape_controls(shape.GroupItems(i))
                    elif shape.Type in {1, 5, 17}:  # AutoShape, freeform, textbox
                        if shape.TextFrame.HasText:
                            collect(shape.TextFrame.TextRange)
                collect(rng)
                for i in range(1, part.Shapes.Count + 1):
                    shape_controls(part.Shapes(i))
                return values
            info.read("content_control_values", controls)
            return info

        for section_index, section in enumerate(doc.Sections, 1):
            for name, collection in (("headers", section.Headers), ("footers", section.Footers)):
                for kind in (1, 2, 3):
                    part = collection(kind)
                    # Failure to discover a story is not observed absence.
                    if part.Exists:
                        info = part_info(part, kind)
                        info["section"] = section_index
                        result[name].append(info)
        return result
    finally:
        close_document(doc)
