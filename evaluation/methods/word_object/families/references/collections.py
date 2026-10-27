"""families / references / collections with explicit evidence dependencies."""

from typing import Dict
from .fields import update_field_collection


def extract_index_info(docx_path: str, *, close_document, open_document) -> Dict:
    """Extract index information from document."""
    doc = None
    try:
        doc = open_document(docx_path)

        index_count = doc.Indexes.Count

        indexes = []
        for i in range(1, index_count + 1):
            idx = doc.Indexes(i)
            indexes.append({"index": i, "text": idx.Range.Text})

        return {"index_count": index_count, "indexes": indexes}

    finally:
        close_document(doc)


def extract_toc_content(
    docx_path: str, update_fields: bool = False, *, close_document, open_document
) -> Dict:
    """Extract Table of Contents information."""
    doc = None
    try:
        doc = open_document(docx_path)

        if update_fields:
            update_field_collection(doc.Fields)

        toc_count = doc.TablesOfContents.Count

        tocs = []
        for i in range(1, toc_count + 1):
            toc = doc.TablesOfContents(i)
            # Extract heading text from paragraph just before the TOC range
            heading_text = ""
            start = toc.Range.Start
            if start > 0:
                before_rng = doc.Range(0, start)
                last_para = before_rng.Paragraphs(before_rng.Paragraphs.Count)
                heading_text = last_para.Range.Text.strip()
            tocs.append(
                {
                    "index": i,
                    "text": toc.Range.Text.strip(),
                    "heading_text": heading_text,
                    "include_page_numbers": toc.IncludePageNumbers,
                }
            )

        return {"toc_count": toc_count, "tocs": tocs}

    finally:
        close_document(doc)
