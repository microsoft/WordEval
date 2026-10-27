"""families / revisions / extract_com with explicit evidence dependencies."""

from typing import Dict


def extract_revision_info(docx_path: str, *, close_document, open_document) -> Dict:
    """
    Extract revision (track changes) information from document.

    Returns:
        Dict with:
        - revision_count: Number of tracked changes
        - track_revisions: Whether tracking is enabled
        - revisions: List of revision details
    """
    doc = None
    try:
        doc = open_document(docx_path)

        revision_count = int(doc.Revisions.Count)
        revisions = []
        metadata_errors = []
        for i in range(1, revision_count + 1):
            revision = doc.Revisions(i)
            revision_range = revision.Range
            entry = {
                "index": i,
                "type": int(revision.Type),  # wdRevisionInsert = 1
                "text": revision_range.Text or "",
                "start": int(revision_range.Start),
                "end": int(revision_range.End),
                "story_type": int(revision_range.StoryType),
                "paragraph_text": revision_range.Paragraphs(1).Range.Text or "",
            }
            for property_name, key in (
                ("Author", "author"),
                ("Date", "date"),
            ):
                try:
                    value = getattr(revision, property_name)
                    entry[key] = str(value) if value is not None else ""
                except Exception as error:
                    metadata_errors.append(
                        {
                            "revision_index": i,
                            "property": property_name,
                            "error": str(error),
                        }
                    )
            revisions.append(entry)

        return {
            "revision_count": revision_count,
            "track_revisions": bool(doc.TrackRevisions),
            "revisions": revisions,
            "body_start": int(doc.Content.Start),
            "body_end": int(doc.Content.End),
            "metadata_errors": metadata_errors,
        }

    finally:
        close_document(doc)
