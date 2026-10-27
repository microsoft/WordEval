"""Shared XML names and diagnostic records."""
import xml.etree.ElementTree as ET
from typing import Any, Dict, Optional

def invalid_schema(
    part: str,
    element: str,
    validation_error: str,
) -> Dict[str, Any]:
    return {
        "reason": "invalid_output",
        "detected_format": "schema_invalid_ooxml",
        "validation_error": validation_error,
        "invalid_part": part,
        "invalid_element": element,
    }

def invalid_child_order(
    part: str,
    parent: ET.Element,
    parent_label: str,
    ordered_groups: tuple[tuple[str, ...], ...],
    namespace: str,
) -> Optional[Dict[str, Any]]:
    """Reject demonstrated strict-sequence inversions.

    OOXML complex types use ordered child sequences.  ElementTree
    parses out-of-order XML successfully, but Word can refuse the
    entire document.  This deliberately validates only a small set
    of sequences observed to make predictions unopenable.
    Unknown/extension children remain outside this preflight.
    """

    ranks = {
        f"{{{namespace}}}{local_name}": rank
        for rank, group in enumerate(ordered_groups)
        for local_name in group
    }
    element_prefix = parent_label.split(":", 1)[0]
    previous_rank = -1
    previous_name = ""
    for child in parent:
        rank = ranks.get(child.tag)
        if rank is None:
            continue
        local_name = child.tag.rsplit("}", 1)[-1]
        if rank < previous_rank:
            return invalid_schema(
                part,
                f"{parent_label}/{element_prefix}:{local_name}",
                (
                    f"pred.docx has invalid child order in "
                    f"{parent_label}: {element_prefix}:{local_name} "
                    f"appears after {element_prefix}:{previous_name}."
                ),
            )
        previous_rank = rank
        previous_name = local_name
    return None

w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
r_ns = (
    "http://schemas.openxmlformats.org/officeDocument/2006/"
    "relationships"
)
m_ns = "http://schemas.openxmlformats.org/officeDocument/2006/math"
c_ns = "http://schemas.openxmlformats.org/drawingml/2006/chart"
a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
cs_ns = (
    "http://schemas.microsoft.com/office/drawing/2012/"
    "chartStyle"
)
wp_ns = (
    "http://schemas.openxmlformats.org/drawingml/2006/"
    "wordprocessingDrawing"
)
wps_ns = (
    "http://schemas.microsoft.com/office/word/2010/"
    "wordprocessingShape"
)
a14_ns = "http://schemas.microsoft.com/office/drawing/2010/main"
mc_ns = (
    "http://schemas.openxmlformats.org/markup-compatibility/2006"
)
