"""families / references / fields with explicit evidence dependencies."""

from typing import Dict, List, Optional
import re
from ...observations import ObservedProperties, require_observations


# XE, TC, RD, TA and PRIVATE carry instructions without a result range.
# Word raises "property is not available" for Result on these native types.
INSTRUCTION_ONLY_FIELD_TYPES = frozenset({4, 9, 11, 74, 77})


def extract_citation_fields(docx_path: str, *, close_document, open_document) -> List[Dict]:
    """Extract CITATION fields (Type=96) from document.

    Returns:
        List of dicts with rendered text and field-code metadata for each
        citation field.
    """
    doc = None
    try:
        doc = open_document(docx_path)
        fields = []
        for i in range(1, doc.Fields.Count + 1):
            f = doc.Fields(i)
            if f.Type == 96:  # wdFieldCitation
                code_text = ""
                try:
                    code_text = (f.Code.Text or "").strip()
                except Exception:
                    pass
                fields.append(
                    {
                        "result": (f.Result.Text or "").strip(),
                        "code": code_text,
                    }
                )
        return fields
    finally:
        close_document(doc)


def extract_fields_from_document(doc) -> List[Dict]:
    fields: List[Dict] = []
    seen = set()
    ordinal = 0

    def _append_field(
        field,
        story_type: Optional[int] = None,
        story_identity: Optional[int] = None,
    ) -> None:
        nonlocal ordinal
        field_type = field.Type
        code_range = field.Code
        code_text = code_range.Text.strip()
        code_start, code_end = code_range.Start, code_range.End
        result_applicable = field_type not in INSTRUCTION_ONLY_FIELD_TYPES
        if result_applicable:
            result_range = field.Result
            result_text = result_range.Text.strip()
            result_start, result_end = result_range.Start, result_range.End
        else:
            result_text, result_start, result_end = "", None, None
        if story_type is None:
            story_type = code_range.StoryType

        position_key = (code_start, code_end, result_start, result_end)
        if any(value is not None for value in position_key):
            key = (story_identity, story_type, field_type, *position_key)
            if key in seen:
                return
            seen.add(key)
        ordinal += 1
        fields.append(
            {
                "index": ordinal,
                "type": field_type,
                "code": code_text,
                "result": result_text,
                "result_applicable": result_applicable,
                "story_type": story_type,
                "code_start": code_start,
                "code_end": code_end,
                "result_start": result_start,
                "result_end": result_end,
            }
        )

    story_found = False
    for rng in iter_word_story_ranges(doc):
        story_found = True
        story_type = rng.StoryType
        rng_fields = rng.Fields
        for j in range(1, rng_fields.Count + 1):
            _append_field(rng_fields(j), story_type, word_range_identity(rng))
    if not story_found:
        for i in range(1, doc.Fields.Count + 1):
            _append_field(doc.Fields(i))

    return fields


def extract_document_fields(
    docx_path: str, update_fields: bool = False, *, close_document, open_document
) -> List[Dict]:
    """Extract Word fields from all document stories.

    Returns each field's type, code text, rendered result text, and story
    type. This lets verifiers reject plain text that visually resembles a
    Word field-backed feature such as captions, page numbers, or
    bibliographies.
    """
    doc = None
    try:
        doc = open_document(docx_path)
        if update_fields:
            update_all_word_fields(doc)
        return extract_fields_from_document(doc)
    finally:
        close_document(doc)


def word_style_name(style_obj) -> Optional[str]:
    """Return a stable local style name from a Word COM style value."""
    try:
        if hasattr(style_obj, "NameLocal"):
            return style_obj.NameLocal
        return str(style_obj)
    except Exception:
        return None


def word_scalar(value):
    """Normalize Word's mixed/undefined sentinels to ``None``."""
    if value in (None, 9999999, -9999999):
        return None
    return value


def required_style_name(style):
    name = style if isinstance(style, str) else style.NameLocal
    if not isinstance(name, str) or not name:
        raise ValueError("Required Word style name is unavailable")
    return name


def reference_code(fields, start, end, keyword):
    """Bind to this object's result range, never an unrelated first field."""
    matches = [field["code"] for field in fields
               if re.match(rf"^\s*{keyword}\b", field["code"], re.I)
               and field.get("story_type") == 1
               and field.get("result_start") is not None
               and start <= field["result_start"] <= end
               and field.get("result_end") is not None
               and field["result_end"] <= end]
    if len(matches) != 1:
        raise ValueError(f"Cannot bind one {keyword} field to range {start}:{end}")
    return matches[0]


def require_reference_sections(states, sections):
    """Check only the sections consumed by this route, including derived faults."""
    for owner, state in states.items():
        if isinstance(state, ObservedProperties):
            error = require_observations({owner: state}, sections, context="reference state")
            if error is not None:
                return error
    return None


def update_field_collection(fields):
    """A failed refresh is not evidence that the submitted cache is current."""
    result = fields.Update()
    if result != 0:
        raise ValueError(f"Word field refresh did not succeed (result={result!r})")


def update_fields_of_type(fields, field_type):
    """Refresh only the native fields consumed by a scoped verifier."""
    for index in range(1, fields.Count + 1):
        field = fields(index)
        if field.Type == field_type:
            result = field.Update()
            if result not in (True, -1):
                raise ValueError(f"Word field {index} refresh did not succeed (result={result!r})")


def update_all_word_fields(doc) -> List[str]:
    """Refresh fields in-memory without saving the document."""
    operations: List[str] = []
    story_found = False
    for story in iter_word_story_ranges(doc):
        story_found = True
        update_field_collection(story.Fields)
        operations.append(f"story:{story.StoryType}")
    if not story_found:
        update_field_collection(doc.Fields)
        operations.append("document_fields")
    for index in range(1, doc.TablesOfContents.Count + 1):
        doc.TablesOfContents(index).Update()
        operations.append(f"toc:{index}")
    for index in range(1, doc.Indexes.Count + 1):
        doc.Indexes(index).Update()
        operations.append(f"index:{index}")
    return operations


def word_range_identity(rng) -> int:
    """Return a stable-enough identity for cycle-safe story traversal."""
    try:
        return id(rng._oleobj_)
    except Exception:
        return id(rng)


def iter_word_story_ranges(doc):
    """Yield every Word story through explicit ``WdStoryType`` keys.

    ``StoryRanges`` is keyed by the Word story constants; its collection
    position is not a story type. Header/footer stories can also form a
    linked ``NextStoryRange`` chain, so traversal is cycle-safe.
    """
    story_ranges = getattr(doc, "StoryRanges", None)
    if story_ranges is None:
        return

    def absent_story(error):
        # Word 5941: the requested member of the collection does not exist.
        # Other COM failures cannot establish an absent optional story.
        codes = [getattr(error, "hresult", None)]
        info = getattr(error, "excepinfo", None)
        if info:
            codes.append(info[-1])
        return isinstance(error, (KeyError, IndexError)) or any(
            isinstance(code, int) and code & 0xffff == 5941 for code in codes)

    for story_type in range(1, 18):
        try:
            story = story_ranges(story_type) if callable(story_ranges) else story_ranges.Item(story_type)
        except Exception as exc:
            if absent_story(exc):
                continue
            raise
        # Retain wrappers as well as their identities: Python may otherwise
        # reuse an id after a previous section's COM wrapper is released.
        seen_ranges = {}
        hops = 0
        while story is not None:
            identity = word_range_identity(story)
            if identity in seen_ranges:
                raise ValueError("Word story chain contains a cycle")
            if hops >= 256:
                raise ValueError("Word story chain exceeds the complete traversal limit")
            seen_ranges[identity] = story
            hops += 1
            yield story
            story = story.NextStoryRange
