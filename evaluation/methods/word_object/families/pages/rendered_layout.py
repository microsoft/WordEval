"""families / pages / rendered_layout with explicit evidence dependencies."""

from evaluation.methods.word_object.families.formatting import (
    paragraph_com as families_formatting_paragraph_com,
)
from typing import Dict, Any
import re
from .field_text import page_field_keyword, page_field_template
from ..formatting.characters import character_property


def extract_body_page_number_layout(
    docx_path: str, *, close_document, open_document
) -> Dict[str, Any]:
    """Extract visible body layout for current-position page-number blocks.

    The gallery inserts real PAGE/NUMPAGES fields into body paragraphs.  A
    rendered text comparison alone cannot distinguish a literal number from
    a dynamic field, while field counts alone miss the surrounding preset
    text and effective formatting.  This extractor keeps those two views
    together without depending on building-block names or package IDs.
    """
    doc = None

    def normalized(value: Any) -> Any:
        if value in (None, 9999999, -9999999):
            return None
        return value

    def rounded(value: Any) -> Any:
        value = normalized(value)
        if value is None:
            return None
        try:
            return round(float(value), 2)
        except (TypeError, ValueError):
            return value

    def font_signature(rng) -> Dict[str, Any]:
        def read(attribute, normalize=lambda value: value):
            return character_property(doc, rng, lambda r: getattr(r.Font, attribute),
                                      normalize=normalize, empty_is_mixed=attribute == "Name")
        return {
            "name": read("Name"),
            "size": read("Size", lambda value: round(float(value), 2)),
            "bold": read("Bold", bool),
            "italic": read("Italic", bool),
            "color": read("Color", lambda value: 0 if value in (-16777216, 0) else value),
            "spacing": read("Spacing", lambda value: round(float(value), 2)),
        }

    try:
        doc = open_document(docx_path)
        try:
            doc.Repaginate()
            doc.Fields.Update()
        except Exception:
            pass

        body_fields = []
        for index in range(1, doc.Fields.Count + 1):
            field = doc.Fields(index)
            story_type = int(field.Code.StoryType)
            if story_type != 1:  # wdMainTextStory
                continue
            keyword = page_field_keyword(field)
            if keyword not in {"PAGE", "NUMPAGES"}:
                continue
            body_fields.append(field)

        paragraphs = []
        fields = []
        for paragraph_index in range(1, doc.Paragraphs.Count + 1):
            paragraph = doc.Paragraphs(paragraph_index)
            paragraph_fields = [
                field
                for field in body_fields
                if field.Result.Start >= paragraph.Range.Start
                and field.Result.End <= paragraph.Range.End
            ]
            paragraph_fields.sort(key=lambda field: field.Result.Start)
            if not paragraph_fields:
                continue

            template = page_field_template(doc, paragraph.Range, paragraph_fields).rstrip("\r\x07")
            units = []
            cursor = paragraph.Range.Start
            for field in paragraph_fields:
                keyword = page_field_keyword(field)
                # Field begin/end characters bracket the literal segments.
                # Exclude those controls so the range font is the effective
                # formatting of visible preset text only.
                literal_end = max(cursor, int(field.Code.Start) - 1)
                if literal_end > cursor:
                    literal_range = doc.Range(cursor, literal_end)
                    literal_text = (
                        (literal_range.Text or "")
                        .replace("\x13", "")
                        .replace("\x14", "")
                        .replace("\x15", "")
                    )
                    if literal_text:
                        units.append(
                            {
                                "kind": "text",
                                "text": literal_text,
                                "font": font_signature(literal_range),
                            }
                        )

                field_entry = {
                    "keyword": keyword,
                    "font": font_signature(field.Result),
                }
                units.append({"kind": "field", **field_entry})
                fields.append(
                    {
                        "paragraph_index": paragraph_index,
                        **field_entry,
                    }
                )
                cursor = int(field.Result.End) + 1

            literal_end = paragraph.Range.End
            while literal_end > cursor and doc.Range(literal_end - 1, literal_end).Text in {
                "\r",
                "\x07",
            }:
                literal_end -= 1
            if literal_end > cursor:
                literal_range = doc.Range(cursor, literal_end)
                literal_text = (
                    (literal_range.Text or "")
                    .replace("\x13", "")
                    .replace("\x14", "")
                    .replace("\x15", "")
                )
                if literal_text:
                    units.append(
                        {
                            "kind": "text",
                            "text": literal_text,
                            "font": font_signature(literal_range),
                        }
                    )

            fmt = paragraph.Format
            paragraphs.append(
                {
                    "paragraph_index": paragraph_index,
                    "template": template,
                    "units": units,
                    "alignment": normalized(fmt.Alignment),
                    "left_indent": rounded(fmt.LeftIndent),
                    "right_indent": rounded(fmt.RightIndent),
                    "first_line_indent": rounded(fmt.FirstLineIndent),
                    "space_before": rounded(fmt.SpaceBefore),
                    "space_after": rounded(fmt.SpaceAfter),
                    "line_spacing": rounded(fmt.LineSpacing),
                    "line_spacing_rule": normalized(fmt.LineSpacingRule),
                    "paragraph_borders": families_formatting_paragraph_com.extract_paragraph_borders(
                        paragraph
                    ),
                }
            )

        return {
            "document_text": doc.Content.Text or "",
            "document_template": page_field_template(doc, doc.Content, body_fields),
            "page_count": doc.ComputeStatistics(2),
            "paragraphs": paragraphs,
            "fields": fields,
        }
    finally:
        close_document(doc)
