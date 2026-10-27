"""families / documents / extract_com with explicit evidence dependencies."""

from typing import Dict


def extract_document_structure(docx_path: str, *, close_document, open_document) -> Dict:
    """Extract document structure counts."""
    doc = None
    try:
        doc = open_document(docx_path)

        # Missing COM observations are not zero objects. A partial structure
        # must surface as an evaluator error instead of matching another zero.
        collections = {
            "footnote_count": "Footnotes",
            "endnote_count": "Endnotes",
            "page_count": None,
            "shape_count": "Shapes",
            "inline_shape_count": "InlineShapes",
            "content_control_count": "ContentControls",
            "equation_count": "OMaths",
            "toc_count": "TablesOfContents",
            "index_count": "Indexes",
            "section_count": "Sections",
            "comment_count": "Comments",
            "table_count": "Tables",
        }
        structure = {}
        for key, collection in collections.items():
            try:
                if collection is None:
                    doc.Repaginate()
                    structure[key] = doc.ComputeStatistics(2)
                else:
                    structure[key] = getattr(doc, collection).Count
            except Exception as exc:
                raise RuntimeError(f"Cannot read document {key}: {exc}") from exc

        return structure

    except Exception as e:
        return {"error": str(e)}

    finally:
        close_document(doc)


def extract_document_text(docx_path: str, *, close_document, open_document) -> str:
    """
    Extract all text content from document.

    Args:
        docx_path: Path to document

    Returns:
        String containing all document text
    """
    doc = None
    try:
        doc = open_document(docx_path)
        try:
            doc.Fields.Update()
        except Exception:
            pass
        return doc.Content.Text
    finally:
        close_document(doc)
