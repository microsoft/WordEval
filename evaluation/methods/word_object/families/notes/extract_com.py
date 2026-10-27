"""families / notes / extract_com with explicit evidence dependencies."""

from typing import Dict, List


def extract_footnote_endnote_content(docx_path: str, *, close_document, open_document) -> Dict:
    """Extract complete effective footnote/endnote state.

    Counts, collection settings, note text, and reference anchors are all
    mandatory evidence.  In particular, a COM collection/property failure
    must propagate to the verifier; returning an empty collection would make
    two failed extractions look semantically equal.
    """
    doc = None
    try:
        doc = open_document(docx_path)

        def extract_collection(collection) -> tuple[int, List[Dict]]:
            count = int(collection.Count)
            entries = []
            for i in range(1, count + 1):
                note = collection(i)
                note_range = note.Range
                reference = note.Reference
                paragraph = reference.Paragraphs(1).Range
                entries.append(
                    {
                        "index": int(note.Index),
                        "text": (note_range.Text or "").strip(),
                        "reference": (reference.Text or "").strip(),
                        "reference_start": int(reference.Start),
                        "reference_end": int(reference.End),
                        # wdActiveEndSectionNumber = 2.  This is the effective
                        # section used by restart-each-section numbering.
                        "section_number": int(reference.Information(2)),
                        "paragraph_start": int(paragraph.Start),
                        "paragraph_end": int(paragraph.End),
                        "paragraph_text": paragraph.Text or "",
                    }
                )
            return count, entries

        footnotes = doc.Footnotes
        endnotes = doc.Endnotes
        footnote_count, footnote_entries = extract_collection(footnotes)
        endnote_count, endnote_entries = extract_collection(endnotes)

        return {
            "footnotes": footnote_entries,
            "endnotes": endnote_entries,
            "footnote_count": footnote_count,
            "endnote_count": endnote_count,
            "footnote_number_style": int(footnotes.NumberStyle),
            "endnote_number_style": int(endnotes.NumberStyle),
            "footnote_numbering_rule": int(footnotes.NumberingRule),
            "endnote_numbering_rule": int(endnotes.NumberingRule),
            "section_count": int(doc.Sections.Count),
            "body_start": int(doc.Content.Start),
            "body_end": int(doc.Content.End),
        }

    finally:
        close_document(doc)
