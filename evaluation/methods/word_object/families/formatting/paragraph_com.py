"""families / formatting / paragraph_com with explicit evidence dependencies."""

from typing import Dict, List, Optional


def extract_paragraph_outline_info_by_text(
    docx_path: str, target_text: str, *, close_document, open_document
) -> Optional[Dict]:
    """
    Extract paragraph style/outline metadata for the paragraph containing target_text.

    Returns:
        Dict with paragraph metadata, or None if target_text not found.
    """
    doc = None
    try:
        doc = open_document(docx_path)

        from ..references.fields import required_style_name
        from .characters import character_property
        from ...observations import ObservedProperties

        matches = []
        for index in range(1, doc.Paragraphs.Count + 1):
            candidate = doc.Paragraphs(index)
            text = (candidate.Range.Text or "").rstrip("\r\x07")
            if target_text.casefold() in text.casefold():
                matches.append((candidate, required_style_name(candidate.Style)))
        if not matches:
            return None
        para, style = next(
            (item for item in matches if not item[1].strip().casefold().startswith("toc ")),
            matches[0],
        )
        state = ObservedProperties(target_text=target_text, paragraph_style=style)
        state.read("paragraph_text", lambda: para.Range.Text.rstrip("\r\x07"))
        state.read("outline_level", lambda: para.OutlineLevel)
        state.read("alignment", lambda: para.Alignment)
        state.read("font_size", lambda: character_property(
            doc, para.Range, lambda r: r.Font.Size, normalize=lambda v: round(float(v), 2)))
        return state
    finally:
        close_document(doc)


def extract_paragraph_borders(para) -> Dict[str, Optional[int]]:
    """Extract paragraph border line styles for the four outer edges."""
    border_types = {
        "top": -1,
        "left": -2,
        "bottom": -3,
        "right": -4,
    }
    from ..tables.border_xml import border_style
    result: Dict[str, Optional[int]] = {}
    for name, border_type in border_types.items():
        line_style = border_style(para, -border_type, name, scope="paragraph")
        if line_style == 9999999:
            raise ValueError(f"Mixed paragraph border: {name}")
        result[name] = line_style or None

    return result


def extract_paragraph_dropcap(para) -> Dict[str, Optional[int]]:
    """Extract drop cap properties from a paragraph."""
    result: Dict[str, Optional[int]] = {
        "position": None,
        "lines_to_drop": None,
    }
    try:
        dc = para.DropCap
        position = dc.Position
        lines_to_drop = dc.LinesToDrop
        result["position"] = position if position != 9999999 else None
        result["lines_to_drop"] = lines_to_drop if lines_to_drop != 9999999 else None
    except:
        pass
    return result


def extract_all_paragraphs_formatting(
    docx_path: str, *, close_document, open_document
) -> Optional[List[Dict]]:
    """
    Extract formatting properties of all paragraphs in document.

    Args:
        docx_path: Path to document

    Returns:
        List of dicts with paragraph formatting properties, or None on error
    """
    doc = None
    try:
        doc = open_document(docx_path)
        paragraphs = []

        for i in range(1, doc.Paragraphs.Count + 1):
            para = doc.Paragraphs(i)
            para_fmt = para.Format
            para_rng = para.Range

            # Get paragraph text (strip trailing paragraph mark)
            para_text = para_rng.Text
            if para_text.endswith("\r"):
                para_text = para_text[:-1]

            dropcap = extract_paragraph_dropcap(para)
            para_info = {
                "index": i,
                "text": para_text[:50] + ("..." if len(para_text) > 50 else ""),
                # Paragraph formatting
                "alignment": para_fmt.Alignment if para_fmt.Alignment != 9999999 else None,
                "left_indent": (
                    round(para_fmt.LeftIndent, 2) if para_fmt.LeftIndent != 9999999 else None
                ),
                "right_indent": (
                    round(para_fmt.RightIndent, 2) if para_fmt.RightIndent != 9999999 else None
                ),
                "first_line_indent": (
                    round(para_fmt.FirstLineIndent, 2)
                    if para_fmt.FirstLineIndent != 9999999
                    else None
                ),
                "line_spacing": para_fmt.LineSpacing if para_fmt.LineSpacing != 9999999 else None,
                "line_spacing_rule": (
                    para_fmt.LineSpacingRule if para_fmt.LineSpacingRule != 9999999 else None
                ),
                "space_before": para_fmt.SpaceBefore if para_fmt.SpaceBefore != 9999999 else None,
                "space_after": para_fmt.SpaceAfter if para_fmt.SpaceAfter != 9999999 else None,
                "paragraph_borders": extract_paragraph_borders(para),
                "dropcap_position": dropcap.get("position"),
                "dropcap_lines_to_drop": dropcap.get("lines_to_drop"),
            }

            # Extract list formatting if present
            try:
                list_fmt = para_rng.ListFormat
                para_info["list_type"] = list_fmt.ListType
                para_info["list_level"] = list_fmt.ListLevelNumber
                para_info["list_string"] = list_fmt.ListString
                # Extract bullet/number font from list template
                try:
                    level_num = list_fmt.ListLevelNumber
                    lt = list_fmt.ListTemplate
                    ll = lt.ListLevels(level_num)  # already 1-based
                    para_info["list_font"] = ll.Font.Name
                except:
                    para_info["list_font"] = None
            except:
                para_info["list_type"] = None
                para_info["list_level"] = None
                para_info["list_string"] = None
                para_info["list_font"] = None

            paragraphs.append(para_info)

        return paragraphs

    finally:
        close_document(doc)


def extract_paragraph_formatting(
    docx_path: str, paragraph_index: int, *, close_document, open_document
) -> Optional[Dict]:
    """
    Extract formatting properties of a specific paragraph.

    Args:
        docx_path: Path to document
        paragraph_index: 1-based paragraph index

    Returns:
        Dict with paragraph formatting properties, or None if not found
    """
    doc = None
    try:
        doc = open_document(docx_path)

        if paragraph_index < 1 or paragraph_index > doc.Paragraphs.Count:
            return None

        para = doc.Paragraphs(paragraph_index)
        para_fmt = para.Format
        para_rng = para.Range

        # Get paragraph text (strip trailing paragraph mark)
        para_text = para_rng.Text
        if para_text.endswith("\r"):
            para_text = para_text[:-1]

        dropcap = extract_paragraph_dropcap(para)
        para_info = {
            "index": paragraph_index,
            "text": para_text[:50] + ("..." if len(para_text) > 50 else ""),
            # Paragraph formatting
            "alignment": para_fmt.Alignment if para_fmt.Alignment != 9999999 else None,
            "left_indent": (
                round(para_fmt.LeftIndent, 2) if para_fmt.LeftIndent != 9999999 else None
            ),
            "right_indent": (
                round(para_fmt.RightIndent, 2) if para_fmt.RightIndent != 9999999 else None
            ),
            "first_line_indent": (
                round(para_fmt.FirstLineIndent, 2) if para_fmt.FirstLineIndent != 9999999 else None
            ),
            "line_spacing": para_fmt.LineSpacing if para_fmt.LineSpacing != 9999999 else None,
            "line_spacing_rule": (
                para_fmt.LineSpacingRule if para_fmt.LineSpacingRule != 9999999 else None
            ),
            "space_before": para_fmt.SpaceBefore if para_fmt.SpaceBefore != 9999999 else None,
            "space_after": para_fmt.SpaceAfter if para_fmt.SpaceAfter != 9999999 else None,
            "paragraph_borders": extract_paragraph_borders(para),
            "dropcap_position": dropcap.get("position"),
            "dropcap_lines_to_drop": dropcap.get("lines_to_drop"),
        }

        # Extract list formatting if present
        try:
            list_fmt = para_rng.ListFormat
            para_info["list_type"] = list_fmt.ListType
            para_info["list_level"] = list_fmt.ListLevelNumber
            para_info["list_string"] = list_fmt.ListString
        except:
            para_info["list_type"] = None
            para_info["list_level"] = None
            para_info["list_string"] = None

        return para_info

    finally:
        close_document(doc)


def extract_dropcap_properties(docx_path: str, *, close_document, open_document) -> Dict:
    """Extract drop cap properties from the first paragraph."""
    doc = None
    try:
        doc = open_document(docx_path)

        result = {
            "position": 0,
            "lines_to_drop": 0,
            "distance_from_text": 0.0,
            "font_name": "",
            "paragraph_index": 1,
            "paragraph_start": None,
            "paragraph_text": "",
            "target_character": "",
        }

        para = doc.Paragraphs(1)
        dc = para.DropCap
        result["position"] = dc.Position
        result["lines_to_drop"] = dc.LinesToDrop
        try:
            result["distance_from_text"] = round(float(dc.DistanceFromText), 2)
        except Exception:
            result["distance_from_text"] = None
        result["font_name"] = dc.FontName or ""
        paragraph_text = (para.Range.Text or "").rstrip("\r\x07")
        result["paragraph_start"] = int(para.Range.Start)
        result["paragraph_text"] = paragraph_text
        result["target_character"] = (
            paragraph_text.lstrip()[0] if paragraph_text.lstrip() else ""
        )

        return result

    finally:
        close_document(doc)
