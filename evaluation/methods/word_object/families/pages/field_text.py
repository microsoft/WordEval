"""Normalize real page-field spans without erasing literal numbers."""


def page_field_keyword(field):
    return {33: "PAGE", 26: "NUMPAGES"}.get(field.Type, "")


def page_field_segments(rng, fields):
    def literal(start, end):
        segment = rng.Duplicate
        segment.Start, segment.End = start, end
        return segment

    spans = sorted(((int(field.Code.Start) - 1, int(field.Result.End) + 1,
                     page_field_keyword(field), field) for field in fields if page_field_keyword(field)),
                   key=lambda span: span[:3])
    cursor = rng.Start
    parts = []
    for start, end, keyword, field in spans:
        if start < cursor or end < start or end > rng.End:
            raise ValueError("Page field spans are overlapping or outside the observed range")
        if start > cursor:
            parts.append(("text", literal(cursor, start), ""))
        parts.append(("field", field.Result, keyword))
        cursor = end
    if cursor < rng.End:
        parts.append(("text", literal(cursor, rng.End), ""))
    return parts


def page_field_template(doc, rng, fields):
    return "".join("{" + keyword + "}" if kind == "field" else segment.Text
                   for kind, segment, keyword in page_field_segments(rng, fields))


def page_field_character_property(doc, rng, fields, getter, **kwargs):
    """Compare field-result formatting independently of cached counter width."""
    from ..formatting.characters import character_property

    if not any(page_field_keyword(field) for field in fields):
        return character_property(doc, rng, getter, **kwargs)
    return [dict(kind=kind, field=keyword,
                 value=character_property(doc, segment, getter, **kwargs))
            for kind, segment, keyword in page_field_segments(rng, fields)
            if kind == "field" or segment.Text]
