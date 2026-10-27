"""families / pages / extract_com with explicit evidence dependencies."""

from typing import Dict, List


def extract_section_properties(docx_path: str, *, close_document, open_document) -> List[Dict]:
    """Extract section properties (page breaks, orientation, etc.)"""
    doc = None
    try:
        doc = open_document(docx_path)
        sections = []

        for i in range(1, doc.Sections.Count + 1):
            section = doc.Sections(i)
            page_setup = section.PageSetup

            sections.append(
                {
                    "index": i,
                    "orientation": page_setup.Orientation,
                    "page_width": round(page_setup.PageWidth, 2),
                    "page_height": round(page_setup.PageHeight, 2),
                    "section_start": page_setup.SectionStart,
                    "different_first_page_header_footer": page_setup.DifferentFirstPageHeaderFooter,
                }
            )

        return sections

    finally:
        close_document(doc)


def extract_page_info(docx_path: str, *, close_document, open_document) -> Dict:
    """Extract page information from document."""
    doc = None
    try:
        doc = open_document(docx_path)

        doc.Repaginate()
        page_count = doc.ComputeStatistics(2)  # wdStatisticPages = 2

        # Count page breaks
        page_breaks = 0
        page_break_positions = []
        rng = doc.Content
        rng.Find.ClearFormatting()
        while rng.Find.Execute(FindText="^m"):
            page_breaks += 1
            try:
                prefix = doc.Range(doc.Content.Start, rng.Start).Text or ""
                visible_prefix = "".join(
                    character for character in prefix if character not in {"\r", "\x0c", "\x07"}
                )
                page_break_positions.append(
                    {
                        "visible_offset": len(visible_prefix),
                        "preceding_text": prefix[-40:],
                        "visible_preceding_text": visible_prefix[-40:],
                    }
                )
            except Exception:
                pass
            rng.Collapse(0)

        return {
            "page_count": page_count,
            "manual_page_breaks": page_breaks,
            "page_break_positions": page_break_positions,
            "content_text": doc.Content.Text or "",
        }

    finally:
        close_document(doc)
