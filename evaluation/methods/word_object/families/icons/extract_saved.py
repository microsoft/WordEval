"""families / icons / extract_saved with explicit evidence dependencies."""

from typing import Dict, List, Any
import re
import zipfile


def extract_svg_icon_signatures(docx_path: str) -> List[Dict[str, Any]]:
    """Extract coarse semantic signatures for embedded SVG icons.

    Word exposes inserted SVG icons through COM as inline or floating picture
    objects. The icon's star/circle geometry lives in word/media/*.svg, so
    this narrow OOXML supplement captures only the shape identity COM does
    not expose.
    """
    signatures: List[Dict[str, Any]] = []
    try:
        with zipfile.ZipFile(docx_path) as zf:
            svg_names = sorted(
                name
                for name in zf.namelist()
                if name.startswith("word/media/") and name.lower().endswith(".svg")
            )
            for name in svg_names:
                text = zf.read(name).decode("utf-8", errors="replace")
                signatures.append(
                    {
                        "part": name,
                        "shape": classify_svg_icon_shape(text),
                    }
                )
    except Exception as e:
        signatures.append({"part": None, "shape": "error", "error": str(e)})
    return signatures


def classify_svg_icon_shape(svg_text: str) -> str:
    """Classify the simple generated icon SVGs as star/circle/unknown."""
    text = svg_text.lower()
    if "<circle" in text:
        return "circle"

    if "<polygon" in text:
        points_match = re.search(r'<polygon\b[^>]*\bpoints=["\']([^"\']+)', text)
        if points_match:
            nums = re.findall(r"-?\d+(?:\.\d+)?", points_match.group(1))
            if len(nums) >= 16:
                return "star"

    path_match = re.search(r'<path\b[^>]*\bd=["\']([^"\']+)', text)
    if path_match:
        d = path_match.group(1)
        # Word normalizes our star polygon to a straight-line path with
        # many coordinate pairs. Circles/curves generally use arc/curve
        # commands, so keep this intentionally narrow.
        if not re.search(r"[acqs]", d, flags=re.IGNORECASE):
            nums = re.findall(r"-?\d+(?:\.\d+)?", d)
            if len(nums) >= 16:
                return "star"

    return "unknown"
