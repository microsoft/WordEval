"""families / references / extract_com with explicit evidence dependencies."""

from evaluation.methods.word_object.families.references import fields as families_references_fields
from ...ooxml.text import _xml_text_by_local_name
from typing import Dict, List, Any
import re
import xml.etree.ElementTree as ET
from ...observations import ObservedProperties, require_observations
from ..formatting.characters import character_property


def extract_dynamic_reference_state(
    docx_path: str, update_fields: bool = False, *, close_document, open_document
) -> Dict[str, Any]:
    """Extract TOC, caption, index, citation, and manual-TOC semantics.

    A false ``update_fields`` read represents the submitted effective state.
    A true read opens a fresh document instance, refreshes only in memory,
    and never saves over the submitted file.
    """
    doc = None
    try:
        doc = open_document(docx_path)
        refresh_operations = (
            families_references_fields.update_all_word_fields(doc) if update_fields else []
        )
        fields = families_references_fields.extract_fields_from_document(doc)

        state = ObservedProperties()

        def observe(name, reader):
            value = state.read(name, reader)
            return [] if value is None else value

        def range_bounds(rng):
            return rng.Start, rng.End

        def paragraph_info(paragraph, index: int) -> Dict[str, Any]:
            rng = paragraph.Range
            start, end = range_bounds(rng)
            text = (rng.Text or "").rstrip("\r\x07")
            style = families_references_fields.required_style_name(paragraph.Style)
            info = ObservedProperties({
                "index": index,
                "start": start,
                "end": end,
                "text": text,
                "style": style,
            })
            info.read("outline_level", lambda: paragraph.OutlineLevel)
            info.read("alignment", lambda: paragraph.Alignment)
            info.read("font_size", lambda: character_property(
                doc, rng, lambda r: r.Font.Size, normalize=lambda v: round(float(v), 2)))
            return info

        paragraphs = observe("paragraphs", lambda: [
            paragraph_info(doc.Paragraphs(index), index)
            for index in range(1, doc.Paragraphs.Count + 1)
        ])

        toc_ranges = []
        toc_heading_starts = set()
        def collect_tocs():
            tocs = []
            for index in range(1, doc.TablesOfContents.Count + 1):
                toc = doc.TablesOfContents(index)
                rng = toc.Range
                start, end = range_bounds(rng)
                toc_ranges.append((start, end))
                code = families_references_fields.reference_code(fields, start, end, "TOC")
                heading_text = ""
                heading_style = None
                heading_start = None
                if start > 0:
                    before = doc.Range(0, start)
                    heading = before.Paragraphs(before.Paragraphs.Count)
                    heading_text = (heading.Range.Text or "").strip()
                    heading_style = families_references_fields.required_style_name(heading.Style)
                    heading_start = heading.Range.Start
                    toc_heading_starts.add(heading_start)
                entries = []
                for paragraph_index in range(1, rng.Paragraphs.Count + 1):
                    paragraph = rng.Paragraphs(paragraph_index)
                    raw = (paragraph.Range.Text or "").rstrip("\r\x07")
                    if not raw.strip():
                        continue
                    parts = raw.split("\t")
                    page_reference = ""
                    entry_text = raw
                    if len(parts) > 1 and re.search(r"\d", parts[-1]):
                        page_reference = parts[-1].strip()
                        entry_text = "\t".join(parts[:-1]).strip()
                    style = families_references_fields.required_style_name(paragraph.Style) or ""
                    match = re.search(r"TOC\s*(\d+)", style, re.I)
                    entries.append(
                        {
                            "text": entry_text,
                            "page_reference": page_reference,
                            "style": style,
                            "level": int(match.group(1)) if match else None,
                        }
                    )
                tocs.append(
                    {
                        "index": index,
                        "start": start,
                        "end": end,
                        "code": code,
                        "heading_text": heading_text,
                        "heading_style": heading_style,
                        "heading_start": heading_start,
                        "include_page_numbers": bool(toc.IncludePageNumbers),
                        "right_align_page_numbers": bool(toc.RightAlignPageNumbers),
                        "entries": entries,
                    }
                )
            return tocs

        tocs = observe("tocs", collect_tocs)

        for paragraph in paragraphs:
            start = paragraph.get("start")
            paragraph["in_toc"] = any(
                range_start is not None
                and range_end is not None
                and start is not None
                and range_start <= start < range_end
                for range_start, range_end in toc_ranges
            )
            paragraph["is_toc_heading"] = start in toc_heading_starts

        def collect_content_controls():
            content_controls = []
            controls = doc.ContentControls
            for index in range(1, controls.Count + 1):
                control = controls(index)
                rng = control.Range
                start, end = range_bounds(rng)
                placeholder = ""
                building_block_gallery = ""
                building_block_category = ""
                try:
                    placeholder = control.PlaceholderText.BuildingBlock.Name or ""
                except Exception:
                    pass
                try:
                    building_block_gallery = str(control.BuildingBlockType)
                except Exception:
                    pass
                try:
                    building_block_category = str(control.BuildingBlockCategory)
                except Exception:
                    pass
                try:
                    temporary = bool(control.Temporary)
                except Exception:
                    temporary = None  # Diagnostic metadata; no route consumes this flag.
                content_controls.append(
                    {
                        "index": index,
                        "type": families_references_fields.word_scalar(control.Type),
                        "title": control.Title or "",
                        "tag": control.Tag or "",
                        "text": (rng.Text or "").strip(),
                        "start": start,
                        "end": end,
                        "showing_placeholder": bool(
                            control.ShowingPlaceholderText
                        ),
                        "temporary": temporary,
                        "placeholder": placeholder,
                        "building_block_gallery": building_block_gallery,
                        "building_block_category": building_block_category,
                    }
                )
            return content_controls

        content_controls = observe("content_controls", collect_content_controls)

        def collect_indexes():
            indexes = []
            for index in range(1, doc.Indexes.Count + 1):
                idx = doc.Indexes(index)
                rng = idx.Range
                start, end = range_bounds(rng)
                code = families_references_fields.reference_code(fields, start, end, "INDEX")
                entries = []
                for line in (rng.Text or "").replace("\x07", "").split("\r"):
                    line = " ".join(line.split())
                    if line:
                        entries.append(line)
                indexes.append(
                    {
                        "index": index,
                        "start": start,
                        "end": end,
                        "code": code,
                        "text": (rng.Text or "").strip(),
                        "entries": entries,
                    }
                )
            return indexes

        indexes = observe("indexes", collect_indexes)

        def source_entry(code: str) -> str:
            quoted = re.search(r'^\s*XE\s+"([^"]+)"', code, re.I)
            if quoted:
                return quoted.group(1)
            plain = re.search(r"^\s*XE\s+([^\\]+)", code, re.I)
            return plain.group(1).strip() if plain else ""

        xe_entries = sorted(
            value for value in (source_entry(field.get("code") or "") for field in fields) if value
        )
        citation_fields = []
        for field in fields:
            if not (
                field.get("type") == 96
                or re.search(r"^\s*CITATION\b", field.get("code") or "", re.I)
            ):
                continue
            paragraph = None
            position = field.get("code_start")
            if position is None:
                position = field.get("result_start")
            if position is not None:
                for candidate in paragraphs:
                    if (
                        candidate.get("start") is not None
                        and candidate.get("end") is not None
                        and candidate["start"] <= position <= candidate["end"]
                    ):
                        paragraph = candidate
                        break
            citation_fields.append(
                {
                    **field,
                    "paragraph_index": (paragraph.get("index") if paragraph else None),
                    "paragraph_start": (paragraph.get("start") if paragraph else None),
                    "paragraph_end": (paragraph.get("end") if paragraph else None),
                    "paragraph_text": (paragraph.get("text") if paragraph else ""),
                    "paragraph_code_offset": (
                        position - paragraph.get("start")
                        if paragraph and position is not None
                        else None
                    ),
                }
            )

        caption_fields = []
        for field in fields:
            code = field.get("code") or ""
            match = re.search(r"^\s*SEQ\s+([^\s\\]+)", code, re.I)
            if field.get("type") != 12 and not match:
                continue
            paragraph = None
            position = field.get("result_start")
            if position is not None:
                for candidate in paragraphs:
                    if (
                        candidate.get("start") is not None
                        and candidate.get("end") is not None
                        and candidate["start"] <= position <= candidate["end"]
                    ):
                        paragraph = candidate
                        break
            caption_fields.append(
                {
                    **field,
                    "label": match.group(1) if match else "",
                    "start": field.get("result_start"),
                    "paragraph_start": paragraph.get("start") if paragraph else None,
                    "paragraph_end": paragraph.get("end") if paragraph else None,
                    "paragraph_text": paragraph.get("text") if paragraph else "",
                    "paragraph_style": paragraph.get("style") if paragraph else None,
                }
            )

        def collect_bibliography_sources():
            bibliography_sources = []
            sources = doc.Bibliography.Sources
            for index in range(1, sources.Count + 1):
                source = sources(index)
                xml = str(source.XML or "")
                ET.fromstring(xml)

                def source_value(name: str) -> str:
                    match = re.search(
                        rf"<(?:\w+:)?{name}>(.*?)</(?:\w+:)?{name}>",
                        xml,
                        re.I | re.S,
                    )
                    return " ".join(match.group(1).split()) if match else ""

                bibliography_sources.append(
                    {
                        "tag": source_value("Tag"),
                        "source_type": source_value("SourceType"),
                        "title": source_value("Title"),
                        "year": source_value("Year"),
                        "last": source_value("Last"),
                        "first": source_value("First"),
                    }
                )
            return bibliography_sources

        bibliography_sources = observe("bibliography_sources", collect_bibliography_sources)

        def collect_tables():
            tables = []
            for index in range(1, doc.Tables.Count + 1):
                table = doc.Tables(index)
                start, end = range_bounds(table.Range)
                cells = []
                for row_index in range(1, table.Rows.Count + 1):
                    row = []
                    for column_index in range(1, table.Columns.Count + 1):
                        text = table.Cell(row_index, column_index).Range.Text
                        row.append(text.rstrip("\r\x07"))
                    cells.append(row)
                tables.append(
                    {
                        "index": index,
                        "start": start,
                        "end": end,
                        "row_count": table.Rows.Count,
                        "column_count": table.Columns.Count,
                        "cell_contents": cells,
                    }
                )
            return tables

        tables = observe("tables", collect_tables)

        def collect_inline_shapes():
            inline_shapes = []
            for index in range(1, doc.InlineShapes.Count + 1):
                shape = doc.InlineShapes(index)
                start, end = range_bounds(shape.Range)
                inline_shapes.append(
                    {
                        "index": index,
                        "type": families_references_fields.word_scalar(shape.Type),
                        "start": start,
                        "end": end,
                        "width": round(float(shape.Width), 2),
                        "height": round(float(shape.Height), 2),
                    }
                )
            return inline_shapes

        inline_shapes = observe("inline_shapes", collect_inline_shapes)

        def story_texts(kind: str) -> List[Dict[str, Any]]:
            values = []
            for section_index in range(1, doc.Sections.Count + 1):
                collection = getattr(doc.Sections(section_index), kind)
                for part_type in (1, 2, 3):
                    part = collection(part_type)
                    if part.Exists:
                        values.append(
                            {
                                "section_index": section_index,
                                "type": part_type,
                                "text": (part.Range.Text or "").strip(),
                                "field_codes": [
                                    (part.Range.Fields(i).Code.Text or "").strip()
                                    for i in range(1, part.Range.Fields.Count + 1)
                                ],
                            }
                        )
            return values


        source_headings = []
        source_paragraphs = []
        for paragraph in paragraphs:
            if paragraph["in_toc"] or paragraph["is_toc_heading"]:
                continue
            text = " ".join((paragraph.get("text") or "").split())
            if not text:
                continue
            error = require_observations(
                {f"paragraph {paragraph['index']}": paragraph}, {"outline_level", "font_size"},
                context="reference source paragraph")
            if error is not None:
                state.unavailable["source_paragraphs"] = str(error.details)
            signature = {
                "text": text,
                "style": paragraph.get("style"),
                "outline_level": paragraph.get("outline_level"),
                "font_size": paragraph.get("font_size"),
            }
            source_paragraphs.append(signature)
            style = paragraph.get("style") or ""
            heading_match = re.search(r"Heading\s*(\d+)", style, re.I)
            if heading_match:
                source_headings.append(
                    {
                        "text": text,
                        "level": int(heading_match.group(1)),
                    }
                )

        state.update({
            "refreshed": update_fields,
            "refresh_operations": refresh_operations,
            "section_count": doc.Sections.Count,
            "body_text": doc.Content.Text or "",
            "paragraphs": paragraphs,
            "source_paragraphs": source_paragraphs,
            "source_headings": source_headings,
            "fields": fields,
            "tocs": tocs,
            "content_controls": content_controls,
            "indexes": indexes,
            "xe_entries": xe_entries,
            "citation_fields": citation_fields,
            "caption_fields": caption_fields,
            "bibliography_sources": bibliography_sources,
            "tables": tables,
            "inline_shapes": inline_shapes,
            "shape_count": doc.Shapes.Count,
            "headers": observe("headers", lambda: story_texts("Headers")),
            "footers": observe("footers", lambda: story_texts("Footers")),
        })
        if "paragraphs" in state.unavailable:
            for key in ("source_paragraphs", "source_headings", "citation_fields", "caption_fields"):
                state.unavailable[key] = state.unavailable["paragraphs"]
        if "tocs" in state.unavailable:
            for key in ("source_paragraphs", "source_headings"):
                state.unavailable[key] = state.unavailable["tocs"]
        return state
    finally:
        close_document(doc)


def extract_bibliography_info(docx_path: str, *, close_document, open_document) -> Dict:
    """Extract bibliography metadata and BIBLIOGRAPHY fields."""
    doc = None
    try:
        doc = open_document(docx_path)
        info = ObservedProperties()
        info.read("refresh", lambda: families_references_fields.update_fields_of_type(doc.Fields, 97) or True)
        info.read("document_text", lambda: doc.Content.Text)
        info.read("bibliography_style", lambda: doc.Bibliography.BibliographyStyle)
        info.read("source_count", lambda: doc.Bibliography.Sources.Count)

        def sources():
            values = []
            collection = doc.Bibliography.Sources
            for i in range(1, collection.Count + 1):
                source = getattr(collection, "Item", collection)(i)
                source_xml = str(source.XML)
                # Validate XML before interpreting missing elements as absence.
                ET.fromstring(source_xml)
                item = {"title": _xml_text_by_local_name(source_xml, "Title")}
                for attribute in ("Tag", "Cited"):
                    try:
                        item[attribute.lower()] = getattr(source, attribute)
                    except Exception:
                        pass  # Diagnostic only; source identity uses XML title.
                values.append(item)
            return values

        info.read("sources", sources)
        info["bibliography_fields"] = [
            field
            for field in families_references_fields.extract_fields_from_document(doc)
            if field.get("type") == 97
            or re.search(r"\bBIBLIOGRAPHY\b", field.get("code") or "", flags=re.IGNORECASE)
        ]
        return info
    finally:
        close_document(doc)
