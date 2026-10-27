"""families / icons / extract_com with explicit evidence dependencies."""

from evaluation.methods.word_object.families.icons import (
    extract_saved as families_icons_extract_saved,
)
from typing import Dict, Any


def extract_svg_icon_properties(docx_path: str, *, open_document) -> Dict[str, Any]:
    """Extract one benchmark SVG icon through saved state plus Word COM.

    The instruction does not require a particular text-wrapping mode, so a
    valid SVG may be exposed as either an InlineShape or a floating Shape.
    Saved SVG media establishes icon identity; WOM supplies the displayed
    object count, dimensions, and effective fill color.
    """
    signatures = families_icons_extract_saved.extract_svg_icon_signatures(docx_path)
    result: Dict[str, Any] = {
        "has_icon": False,
        "svg_signatures": signatures,
        "svg_shapes": [row.get("shape") for row in signatures],
    }
    doc = None
    converted_shape = None
    try:
        doc = open_document(docx_path, read_only=False)
        inline_count = doc.InlineShapes.Count
        floating_count = doc.Shapes.Count
        result.update(
            {
                "inline_shape_count": inline_count,
                "floating_shape_count": floating_count,
                "visual_object_count": inline_count + floating_count,
            }
        )
        if not signatures or inline_count + floating_count < 1:
            return result

        result["has_icon"] = True
        result["count"] = len(signatures)
        if inline_count:
            icon = doc.InlineShapes(1)
            result["object_kind"] = "inline"
            result["width"] = round(icon.Width, 2)
            result["height"] = round(icon.Height, 2)
            try:
                converted_shape = icon.ConvertToShape()
                result["fill_color_rgb"] = converted_shape.Fill.ForeColor.RGB
            except Exception:
                result["fill_color_rgb"] = None
        else:
            icon = doc.Shapes(1)
            result["object_kind"] = "floating"
            result["width"] = round(icon.Width, 2)
            result["height"] = round(icon.Height, 2)
            try:
                result["fill_color_rgb"] = icon.Fill.ForeColor.RGB
            except Exception:
                result["fill_color_rgb"] = None
        return result
    finally:
        if converted_shape is not None:
            try:
                converted_shape.ConvertToInlineShape()
            except Exception:
                pass
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass


def extract_inserted_svg_icon_state(
    docx_path: str, *, close_document, open_document
) -> Dict[str, Any]:
    """Read-only WOM display state for an SVG icon insertion.

    SVG identity and color live in the saved relationship graph and SVG
    payload, so this extractor intentionally owns only the visible object
    count, wrapping representation, and displayed dimensions.
    """
    doc = None
    try:
        doc = open_document(docx_path)
        inline_count = int(doc.InlineShapes.Count)
        floating_count = int(doc.Shapes.Count)
        result: Dict[str, Any] = {
            "visual_object_count": inline_count + floating_count,
            "inline_shape_count": inline_count,
            "floating_shape_count": floating_count,
        }
        if inline_count + floating_count != 1:
            return result
        if inline_count == 1:
            icon = doc.InlineShapes(1)
            object_kind = "inline"
        else:
            icon = doc.Shapes(1)
            object_kind = "floating"
        result.update(
            {
                "object_kind": object_kind,
                "width": round(float(icon.Width), 2),
                "height": round(float(icon.Height), 2),
            }
        )
        return result
    finally:
        close_document(doc)
