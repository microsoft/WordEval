"""Literal run signatures and font-aware bullet equivalence."""

import xml.etree.ElementTree as ET
from typing import Any, Dict, List, Optional

# Bullet character equivalence mapping for BulletsGalleryWord.
# Word's built-in bullet gallery uses legacy font-specific characters (Wingdings/Symbol PUA),
# but models may produce standard Unicode equivalents that render as the same bullet style.
# Map all known equivalent representations to a canonical form for comparison.
BULLET_CHAR_EQUIVALENTS = {
    # Checkmark: Wingdings 0xFC (PUA) <=> Unicode CHECK MARK
    "\uf0fc": "\u2713",  # Wingdings checkmark -> Unicode checkmark
    "\u2714": "\u2713",  # Heavy check mark -> Unicode checkmark
    # Solid circle: Symbol 0xB7 (PUA) <=> Unicode BLACK CIRCLE / BULLET
    "\uf0b7": "\u25cf",  # Symbol bullet -> Unicode black circle
    "\u2022": "\u25cf",  # Bullet character -> Unicode black circle
    # Solid square: Wingdings 0xA7 (PUA) / Wingdings 'n' (0x6E) <=> Unicode BLACK SQUARE
    "\uf0a7": "\u25a0",  # Wingdings small square -> Unicode black square
    "n": "\u25a0",  # Wingdings 'n' (large square variant) -> Unicode black square
    # Hollow circle: Courier New 'o' <=> Unicode WHITE CIRCLE
    "o": "\u25cb",  # lowercase 'o' -> Unicode white circle
}


def _normalize_bullet_char(char_str: Optional[str]) -> Optional[str]:
    """Normalize a bullet list_string to a canonical character for equivalence comparison."""
    if not char_str:
        return char_str
    c = char_str.strip()
    return BULLET_CHAR_EQUIVALENTS.get(c, c)


# Standard Unicode bullet characters that have specific geometric meanings.
# When used with legacy fonts (Symbol, Wingdings), these codepoints map to
# completely different glyphs (e.g., U+25CF in Symbol renders as "not element of").
_UNICODE_BULLET_CHARS = {
    "\u25cf",  # BLACK CIRCLE (●)
    "\u25cb",  # WHITE CIRCLE (○)
    "\u25a0",  # BLACK SQUARE (■)
    "\u2713",  # CHECK MARK (✓)
    "\u2714",  # HEAVY CHECK MARK (✔)
    "\u2022",  # BULLET (•)
}

_LEGACY_FONTS = {"Symbol", "Wingdings", "Wingdings 2", "Wingdings 3"}

# These ASCII code points are only bullet glyphs under the named fonts. In a
# normal text font they remain the literal letters "n" and "o".
_FONT_DEPENDENT_BULLET_REQUIREMENTS = {
    "n": {"wingdings"},
    "o": {"courier new"},
    "\uf0a7": {"wingdings"},
    "\uf0b7": {"symbol"},
    "\uf0fc": {"wingdings"},
}

_LEGACY_FONT_KEYS = {font.casefold() for font in _LEGACY_FONTS}


def _bullet_chars_equivalent(
    gold_char: Optional[str],
    pred_char: Optional[str],
    gold_font: Optional[str],
    pred_font: Optional[str],
) -> bool:
    """
    Check if two bullet characters are equivalent, considering their fonts.

    Rejects standard Unicode bullet chars used with legacy fonts (Symbol/Wingdings),
    because these fonts map Unicode codepoints to entirely different glyphs
    (e.g., U+25CF in Symbol renders as a "not element of" symbol, not a circle).
    """
    pred_c = pred_char.strip() if pred_char else pred_char
    pred_font_key = (pred_font or "").strip().casefold()
    required_fonts = _FONT_DEPENDENT_BULLET_REQUIREMENTS.get(pred_c)
    if required_fonts is not None and pred_font_key not in required_fonts:
        return False
    if gold_char == pred_char:
        return True
    # Reject standard Unicode bullet chars rendered in legacy fonts - wrong glyphs
    if pred_c in _UNICODE_BULLET_CHARS and pred_font_key in _LEGACY_FONT_KEYS:
        return False
    # Apply normalization for all other cases
    return _normalize_bullet_char(gold_char) == _normalize_bullet_char(pred_char)


def _adjacent_text_run_groups(
    root: ET.Element,
    word_namespace: str,
) -> List[List[ET.Element]]:
    """Return adjacent text runs without crossing their XML carrier."""
    body = root.find(f".//{{{word_namespace}}}body")
    if body is None:
        return []
    run_tag = f"{{{word_namespace}}}r"
    text_tag = f"{{{word_namespace}}}t"
    groups: List[List[ET.Element]] = []
    for parent in body.iter():
        adjacent: List[ET.Element] = []
        for child in list(parent):
            text = "".join(node.text or "" for node in child.findall(f".//{text_tag}"))
            if child.tag == run_tag and text:
                adjacent.append(child)
                continue
            if adjacent:
                groups.append(adjacent)
                adjacent = []
        if adjacent:
            groups.append(adjacent)
    return groups


def _xml_element_signature(element: Optional[ET.Element]) -> Any:
    """Return a prefix-independent structural signature for run properties."""
    if element is None:
        return None
    text = element.text if element.text and element.text.strip() else None
    return (
        element.tag,
        tuple(sorted(element.attrib.items())),
        text,
        tuple(_xml_element_signature(child) for child in element),
    )


def _matching_adjacent_run_signatures(
    runs: List[Dict[str, Any]],
    target_text: str,
    *,
    strip_text: bool,
) -> List[Dict[str, Any]]:
    """Match a target split across adjacent, format-identical saved runs."""
    normalize = (
        (lambda value: value.strip().casefold()) if strip_text else (lambda value: value.casefold())
    )
    target_key = normalize(target_text)
    if not target_key:
        return []

    matches = []
    for start, first in enumerate(runs):
        merge_formatting = {key: value for key, value in first.items() if key != "text"}
        result_formatting = {
            key: value
            for key, value in first.items()
            if key not in {"text", "_merge_run_properties"}
        }
        combined = ""
        for run in runs[start:]:
            run_formatting = {key: value for key, value in run.items() if key != "text"}
            if run_formatting != merge_formatting:
                break
            combined += run.get("text") or ""
            combined_key = normalize(combined)
            if combined_key == target_key:
                matches.append(
                    {
                        "text": combined.strip() if strip_text else combined,
                        **result_formatting,
                    }
                )
                break
            if not target_key.startswith(combined_key):
                break
    return matches
