"""families / metadata / extract_com with explicit evidence dependencies."""

from evaluation.methods.word_object.families.references import fields as families_references_fields
from typing import Dict, Any


def extract_docinfo_state(
    docx_path: str, update_fields: bool = False, *, close_document, open_document
) -> Dict[str, Any]:
    """Extract DocInfo controls and fields from every header story.

    Unlike ``extract_header_footer_content``, this keeps each content
    control instance and exposes its XML-mapping state.  That distinction
    is necessary to reject visible but unbound lookalikes while ignoring
    non-semantic SDT ids and package serialization.
    """
    doc = None
    try:
        doc = open_document(docx_path)
        refresh_operations = (
            families_references_fields.update_all_word_fields(doc) if update_fields else []
        )

        def field_info(field) -> Dict[str, Any]:
            code = ""
            result = ""
            start = end = story_type = None
            try:
                code = (field.Code.Text or "").strip()
                story_type = field.Code.StoryType
            except Exception:
                pass
            try:
                result = (field.Result.Text or "").strip()
                start = field.Result.Start
                end = field.Result.End
            except Exception:
                pass
            return {
                "type": families_references_fields.word_scalar(getattr(field, "Type", None)),
                "code": code,
                "result": result,
                "start": start,
                "end": end,
                "story_type": story_type,
            }

        def control_info(control) -> Dict[str, Any]:
            mapping = None
            try:
                mapping = control.XMLMapping
            except Exception:
                pass
            is_mapped = False
            xpath = ""
            store_item_id = ""
            source_exists = False
            if mapping is not None:
                try:
                    is_mapped = bool(mapping.IsMapped)
                except Exception:
                    pass
                try:
                    xpath = mapping.XPath or ""
                except Exception:
                    pass
                try:
                    store_item_id = mapping.StoreItemID or ""
                except Exception:
                    pass
                try:
                    source_exists = mapping.CustomXMLPart is not None
                except Exception:
                    source_exists = False
            text = ""
            start = end = story_type = None
            try:
                text = (control.Range.Text or "").strip()
                start = control.Range.Start
                end = control.Range.End
                story_type = control.Range.StoryType
            except Exception:
                pass
            return {
                "type": families_references_fields.word_scalar(getattr(control, "Type", None)),
                "title": getattr(control, "Title", "") or "",
                "tag": getattr(control, "Tag", "") or "",
                "text": text,
                "start": start,
                "end": end,
                "story_type": story_type,
                "is_mapped": is_mapped,
                "xpath": xpath,
                "store_item_id": store_item_id,
                "source_exists": source_exists,
            }

        def part_info(part, section_index: int, part_type: int) -> Dict[str, Any]:
            rng = part.Range
            fields = []
            controls = []
            try:
                fields = [field_info(rng.Fields(index)) for index in range(1, rng.Fields.Count + 1)]
            except Exception:
                pass
            try:
                controls = [
                    control_info(rng.ContentControls(index))
                    for index in range(1, rng.ContentControls.Count + 1)
                ]
            except Exception:
                pass
            return {
                "section_index": section_index,
                "type": part_type,
                "exists": bool(getattr(part, "Exists", True)),
                "link_to_previous": bool(getattr(part, "LinkToPrevious", False)),
                "text": (rng.Text or "").strip(),
                "fields": fields,
                "controls": controls,
            }

        headers = []
        footers = []
        for section_index in range(1, doc.Sections.Count + 1):
            section = doc.Sections(section_index)
            for part_type in (1, 2, 3):
                try:
                    part = section.Headers(part_type)
                    if part.Exists:
                        headers.append(part_info(part, section_index, part_type))
                except Exception:
                    pass
                try:
                    part = section.Footers(part_type)
                    if part.Exists:
                        footers.append(part_info(part, section_index, part_type))
                except Exception:
                    pass
        return {
            "section_count": doc.Sections.Count,
            "body_text": doc.Content.Text or "",
            "headers": headers,
            "footers": footers,
            "refreshed": update_fields,
            "refresh_operations": refresh_operations,
        }
    finally:
        close_document(doc)
