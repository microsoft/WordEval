"""families / hyperlinks / extract_com with explicit evidence dependencies."""

from typing import Dict


def extract_hyperlink_info(docx_path: str, *, close_document, open_document) -> Dict:
    """
    Extract hyperlink properties from document via COM.

    Returns:
        Dict with:
        - hyperlink_count: Number of hyperlinks
        - hyperlinks: List of hyperlink details (index, address, sub_address,
          text_to_display, screen_tip, type)
        - document_text: Full document text (for verifying display text preserved)
    """
    doc = None
    try:
        doc = open_document(docx_path)

        result = {
            "hyperlink_count": 0,
            "hyperlinks": [],
            "document_text": "",
        }

        try:
            result["document_text"] = doc.Content.Text.strip()
        except:
            pass

        try:
            hyperlinks = doc.Hyperlinks
            result["hyperlink_count"] = hyperlinks.Count

            for i in range(1, hyperlinks.Count + 1):
                hl = hyperlinks(i)
                # Extract character style from the hyperlink range
                char_style = ""
                font_color = None
                font_underline = None
                try:
                    rng = hl.Range
                    if rng and rng.CharacterStyle:
                        char_style = str(rng.CharacterStyle.NameLocal or "")
                    if rng and rng.Font:
                        font_color = rng.Font.Color
                        font_underline = rng.Font.Underline
                except:
                    pass
                result["hyperlinks"].append(
                    {
                        "index": i,
                        "address": (hl.Address or "").strip(),
                        "sub_address": (hl.SubAddress or "").strip(),
                        "text_to_display": (hl.TextToDisplay or "").strip(),
                        "screen_tip": (hl.ScreenTip or "").strip(),
                        "type": hl.Type,  # 0=msoHyperlinkRange, 1=msoHyperlinkShape
                        "character_style": char_style,
                        "font_color": font_color,
                        "font_underline": font_underline,
                    }
                )

            # Check for residual Hyperlink character style (when hyperlinks removed
            # but styling left behind)
            if hyperlinks.Count == 0:
                result["has_residual_hyperlink_style"] = False
                try:
                    # Search for text with "Hyperlink" character style
                    content = doc.Content
                    find = content.Find
                    find.ClearFormatting()
                    find.Style = doc.Styles("Hyperlink")
                    found = find.Execute(FindText="", Forward=True, Wrap=0)
                    if found:
                        result["has_residual_hyperlink_style"] = True
                except:
                    pass  # Style may not exist in document
        except Exception as e:
            result["error"] = str(e)

        return result

    finally:
        close_document(doc)
