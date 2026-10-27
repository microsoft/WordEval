"""families / shapes / extract_com with explicit evidence dependencies."""

from .colors import _has_visible_gradient_progression
from ...observations import ObservedProperties
from typing import Dict, List, Optional, Any


def extract_shape_properties(docx_path: str, *, close_document, open_document) -> List[Dict]:
    """
    Extract properties of all shapes in document.

    Returns list of dicts with shape properties:
    - name, left, top, width, height
    - center_h, center_v (calculated)
    - z_order, rotation
    - is_group, group_items (for groups)
    - flip_horizontal, flip_vertical (for rotation/flip)
    - wrap_type (for text wrapping)
    - relative_h_position, relative_v_position (for positioning)
    """
    doc = None
    try:
        doc = open_document(docx_path)
        shapes = []

        # Get page dimensions and margins once
        page_setup = doc.PageSetup
        page_width = round(page_setup.PageWidth, 2)
        page_height = round(page_setup.PageHeight, 2)
        left_margin = round(page_setup.LeftMargin, 2)
        top_margin = round(page_setup.TopMargin, 2)
        right_margin = round(page_setup.RightMargin, 2)
        bottom_margin = round(page_setup.BottomMargin, 2)

        for i in range(1, doc.Shapes.Count + 1):
            shape = doc.Shapes(i)

            shape_info = ObservedProperties({
                "name": shape.Name,
                "left": round(shape.Left, 2),
                "top": round(shape.Top, 2),
                "width": round(shape.Width, 2),
                "height": round(shape.Height, 2),
                "center_h": round(shape.Left + shape.Width / 2, 2),
                "center_v": round(shape.Top + shape.Height / 2, 2),
                "z_order": shape.ZOrderPosition,
                "rotation": round(shape.Rotation, 2),
                "type": shape.Type,
                # Page dimensions for position calculations
                "page_width": page_width,
                "page_height": page_height,
                "left_margin": left_margin,
                "top_margin": top_margin,
                "right_margin": right_margin,
                "bottom_margin": bottom_margin,
            })

            shape_info.read('visible', lambda: shape.Visible, bool)

            # Extract flip state
            shape_info.read('flip_horizontal', lambda: shape.HorizontalFlip, bool)
            shape_info.read('flip_vertical', lambda: shape.VerticalFlip, bool)

            # Extract text wrap format
            shape_info.read('wrap_type', lambda: shape.WrapFormat.Type)

            # Extract relative position settings
            shape_info.read('relative_h_position', lambda: shape.RelativeHorizontalPosition)
            shape_info.read('relative_v_position', lambda: shape.RelativeVerticalPosition)

            # Extract fill transparency (for PictureTransparencyGallery)
            shape_info.read('fill_transparency', lambda: shape.Fill.Transparency, lambda value: round(value, 4))

            # Extract shadow properties (for ShadowNudge*, ShadowOnOrOff*, etc.)
            shape_info.read('shadow_visible', lambda: shape.Shadow.Visible, bool)
            shape_info.read('shadow_offset_x', lambda: shape.Shadow.OffsetX, lambda value: round(value, 2))
            shape_info.read('shadow_offset_y', lambda: shape.Shadow.OffsetY, lambda value: round(value, 2))
            shape_info.read('shadow_type', lambda: shape.Shadow.Type)

            # Extract text alignment for textbox shapes (for TextAlign* commands)
            # wdAlignParagraph values:
            # 0=Left, 1=Center, 2=Right, 3=Justify, 4=JustifyLow, 5=Distribute
            try:
                if shape.TextFrame.HasText:
                    # Get alignment from first paragraph in text frame
                    text_range = shape.TextFrame.TextRange
                    text_value = text_range.Text or ""
                    if text_value.endswith("\r"):
                        text_value = text_value[:-1]
                    shape_info["text"] = text_value
                    alignments = set()
                    for i in range(1, text_range.Paragraphs.Count + 1):
                        para = text_range.Paragraphs(i)
                        alignments.add(para.Alignment)
                    # Store the alignment (if all paragraphs same, store that; else store set)
                    if len(alignments) == 1:
                        shape_info["text_alignment"] = alignments.pop()
                    else:
                        shape_info["text_alignment"] = list(alignments)
                else:
                    shape_info["text"] = None
                    shape_info["text_alignment"] = None
            except Exception as exc:
                shape_info.record_unavailable("text", exc)
                shape_info.record_unavailable("text_alignment", exc)

            # Extract vertical alignment inside the textbox (TextAlignGallery).
            shape_info.read('text_vertical_anchor', lambda: shape.TextFrame.VerticalAnchor)

            # Extract textbox font color/effect color. Textbox prerequisite
            # compositions explicitly ask for text fill/color changes, which
            # are not represented by the shape fill itself.
            try:
                if shape.TextFrame.HasText:
                    font = shape.TextFrame.TextRange.Font
                    shape_info.read("text_font_color", lambda: font.Color)
                    shape_info.read('text_fill_type', lambda: font.Fill.Type)
                    shape_info.read('text_fill_color', lambda: font.Fill.ForeColor.RGB)
                else:
                    shape_info["text_font_color"] = None
                    shape_info["text_fill_type"] = None
                    shape_info["text_fill_color"] = None
            except Exception as exc:
                for field in ("text_font_color", "text_fill_type", "text_fill_color"):
                    shape_info.record_unavailable(field, exc)

            # Extract WordArt TextEffect properties
            shape_info.read('wordart_normalized_height', lambda: shape.TextEffect.NormalizedHeight)
            shape_info.read('wordart_kerned_pairs', lambda: shape.TextEffect.KernedPairs)
            shape_info.read('wordart_tracking', lambda: shape.TextEffect.Tracking, lambda value: round(value, 2))
            shape_info.read('wordart_rotated_chars', lambda: shape.TextEffect.RotatedChars)

            # Extract text frame orientation for vertical text (TextDirection command)
            # NOTE: TextFrame2.Orientation fails for textboxes, use TextFrame.TextRange.Orientation
            # wdTextOrientation values: 0=Horizontal,
            # 1=VerticalFarEast, 2=Upward, 3=Downward,
            # 4=HorizontalRotatedFarEast, 5=Vertical.
            shape_info.read("text_frame_orientation", lambda: shape.TextFrame.TextRange.Orientation)
            if "text_frame_orientation" in shape_info.unavailable:
                # Fallback to TextFrame2.Orientation for other shape types
                primary_error = shape_info.unavailable["text_frame_orientation"]
                shape_info.read('text_frame_orientation', lambda: shape.TextFrame2.Orientation)
                if "text_frame_orientation" in shape_info.unavailable:
                    fallback_error = shape_info.unavailable["text_frame_orientation"]
                    shape_info.record_unavailable("text_frame_orientation",
                        f"TextFrame.TextRange: {primary_error}; TextFrame2: {fallback_error}")

            # Extract 3D effect properties
            shape_info.read('threed_visible', lambda: shape.ThreeD.Visible)
            shape_info.read('threed_rotation_x', lambda: shape.ThreeD.RotationX, lambda value: round(value, 1))
            shape_info.read('threed_rotation_y', lambda: shape.ThreeD.RotationY, lambda value: round(value, 1))
            shape_info.read('threed_depth', lambda: shape.ThreeD.Depth, lambda value: round(value, 1))
            shape_info.read('threed_perspective', lambda: shape.ThreeD.Perspective)

            # Extract fill properties (for ShapeFillColorPicker, GradientGallery, etc.)
            shape_info.read('fill_visible', lambda: shape.Fill.Visible, bool)
            shape_info.read('fill_type', lambda: shape.Fill.Type)
            shape_info.read('fill_fore_color_rgb', lambda: shape.Fill.ForeColor.RGB)
            shape_info.read('fill_gradient_angle', lambda: shape.Fill.GradientAngle, lambda value: round(value, 1))
            shape_info.read('fill_gradient_style', lambda: shape.Fill.GradientStyle, int)
            shape_info.read('fill_gradient_variant', lambda: shape.Fill.GradientVariant, int)
            shape_info.read('fill_gradient_degree', lambda: shape.Fill.GradientDegree, lambda value: round(float(value), 3))
            try:
                gradient_stops = []
                stops = shape.Fill.GradientStops
                for stop_index in range(1, stops.Count + 1):
                    stop = stops.Item(stop_index)
                    gradient_stops.append(
                        {
                            "rgb": int(stop.Color.RGB),
                            "position": round(float(stop.Position), 4),
                            "transparency": round(float(stop.Transparency), 4),
                        }
                    )
                shape_info["fill_gradient_stops"] = gradient_stops
                shape_info["fill_gradient_has_visible_progression"] = (
                    _has_visible_gradient_progression(gradient_stops)
                )
            except Exception as exc:
                shape_info["fill_gradient_stops"] = []
                shape_info["fill_gradient_has_visible_progression"] = None
                shape_info.setdefault("_extraction_errors", []).append(
                    {
                        "property": "fill_gradient_stops",
                        "derived_properties": [
                            "fill_gradient_has_visible_progression",
                        ],
                        "error": str(exc),
                    }
                )

            # Extract line/outline properties (for ShapeOutlineColorPicker, OutlineDashesGallery, etc.)
            shape_info.read('line_visible', lambda: shape.Line.Visible, bool)
            shape_info.read('line_fore_color_rgb', lambda: shape.Line.ForeColor.RGB)
            shape_info.read('line_dash_style', lambda: shape.Line.DashStyle)
            shape_info.read('line_weight', lambda: shape.Line.Weight, lambda value: round(value, 2))
            shape_info.read('line_begin_arrowhead_style', lambda: shape.Line.BeginArrowheadStyle)
            shape_info.read('line_end_arrowhead_style', lambda: shape.Line.EndArrowheadStyle)

            # Extract glow properties (for GlowColorPicker, ObjectEffectGlowGallery)
            shape_info.read('glow_radius', lambda: shape.Glow.Radius, lambda value: round(value, 2))
            shape_info.read('glow_color_rgb', lambda: shape.Glow.Color.RGB)

            # Extract soft edge (for ObjectEffectSoftEdgesGallery)
            shape_info.read('soft_edge_radius', lambda: shape.SoftEdge.Radius, lambda value: round(value, 2))

            # Extract reflection (for ReflectionGallery)
            shape_info.read('reflection_type', lambda: shape.Reflection.Type)

            # Extract auto shape type (for ShapeChangeShapeGallery)
            shape_info.read('auto_shape_type', lambda: shape.AutoShapeType)

            # Extract 3D bevel (for BevelShapeGallery)
            shape_info.read('threed_bevel_top_type', lambda: shape.ThreeD.BevelTopType)
            shape_info.read('threed_bevel_top_depth', lambda: shape.ThreeD.BevelTopDepth, lambda value: round(value, 1))
            shape_info.read('threed_bevel_top_inset', lambda: shape.ThreeD.BevelTopInset, lambda value: round(value, 1))

            # Extract 3D rotation preset (for _3DRotationGallery)
            shape_info.read('threed_preset_camera', lambda: shape.ThreeD.PresetCamera)

            shape_info["is_group"] = shape_info["type"] == 6
            if shape_info["is_group"]:
                shape_info.read("group_item_count", lambda: shape.GroupItems.Count)

            shapes.append(shape_info)

        return shapes

    finally:
        close_document(doc)


def extract_shape_text_styles(
    docx_path: str,
    include_header_footer: bool = False,
    include_group_items: bool = False,
    *,
    close_document,
    open_document,
) -> List[Dict]:
    """Extract text/style signatures from shapes.

    By default this keeps the historical body-shape behavior used by TextBox
    galleries. Margin page-number galleries can request header/footer story
    shapes because Word stores those margin numbers in header/footer ranges.
    """
    doc = None
    try:
        doc = open_document(docx_path)
        if include_header_footer:
            from .margin_styles import extract_margin_shape_styles
            return extract_margin_shape_styles(doc, include_group_items=True)
        shapes = []

        def _shape_text_info(shape, index: str, story: str) -> Dict:
            def normalized(value: Any) -> Any:
                if value in (None, 9999999, -9999999):
                    return None
                return value

            def rounded(value: Any) -> Any:
                value = normalized(value)
                if value is None:
                    return None
                try:
                    return round(float(value), 2)
                except (TypeError, ValueError):
                    return value

            def safe(object_value, attribute: str, default=None):
                try:
                    return getattr(object_value, attribute)
                except Exception:
                    return default

            def font_signature(rng) -> Dict[str, Any]:
                font = rng.Font
                color = normalized(font.Color)
                if color in (-16777216, 0):
                    color = 0
                bold = normalized(font.Bold)
                italic = normalized(font.Italic)
                return {
                    "name": normalized(font.Name),
                    "size": rounded(font.Size),
                    "bold": None if bold is None else bool(bold),
                    "italic": None if italic is None else bool(italic),
                    "color": color,
                    "color_index": normalized(font.ColorIndex),
                    "spacing": rounded(font.Spacing),
                }

            fill = safe(shape, "Fill")
            line = safe(shape, "Line")
            text_frame = safe(shape, "TextFrame")
            wrap = safe(shape, "WrapFormat")
            shape_info = {
                "index": index,
                "story": story,
                "type": getattr(shape, "Type", None),
                "auto_shape_type": safe(shape, "AutoShapeType"),
                "relative_horizontal_position": safe(shape, "RelativeHorizontalPosition"),
                "relative_vertical_position": safe(shape, "RelativeVerticalPosition"),
                "left": rounded(safe(shape, "Left")),
                "top": rounded(safe(shape, "Top")),
                "width": rounded(safe(shape, "Width")),
                "height": rounded(safe(shape, "Height")),
                "rotation": rounded(safe(shape, "Rotation")),
                "wrap_type": safe(wrap, "Type"),
                "text_orientation": safe(text_frame, "Orientation"),
                "text_margin_left": rounded(safe(text_frame, "MarginLeft")),
                "text_margin_right": rounded(safe(text_frame, "MarginRight")),
                "text_margin_top": rounded(safe(text_frame, "MarginTop")),
                "text_margin_bottom": rounded(safe(text_frame, "MarginBottom")),
                "fill_visible": safe(fill, "Visible"),
                "fill_type": safe(fill, "Type"),
                "fill_color": safe(safe(fill, "ForeColor"), "RGB"),
                "fill_theme_color": safe(safe(fill, "ForeColor"), "ObjectThemeColor"),
                "fill_transparency": rounded(safe(fill, "Transparency")),
                "line_visible": safe(line, "Visible"),
                "line_style": safe(line, "Style"),
                "line_dash_style": safe(line, "DashStyle"),
                "line_weight": rounded(safe(line, "Weight")),
                "line_color": safe(safe(line, "ForeColor"), "RGB"),
                "line_theme_color": safe(safe(line, "ForeColor"), "ObjectThemeColor"),
                "line_transparency": rounded(safe(line, "Transparency")),
                "paragraph_styles": [],
                "paragraph_texts": [],
                "paragraph_alignments": [],
                "paragraph_fonts": [],
                "field_codes": [],
                "field_keywords": [],
                "field_fonts": [],
                "page_field_count": 0,
                "num_pages_field_count": 0,
            }
            try:
                if shape.TextFrame.HasText:
                    tr = shape.TextFrame.TextRange
                    for j in range(1, tr.Paragraphs.Count + 1):
                        try:
                            para = tr.Paragraphs(j)
                            para_text = para.Range.Text if para.Range else ""
                            if para_text.endswith("\r"):
                                para_text = para_text[:-1]
                            para_text = para_text.strip()
                            try:
                                style = para.Style
                                if hasattr(style, "NameLocal"):
                                    style_name = style.NameLocal if style else ""
                                else:
                                    style_name = str(style) if style else ""
                            except:
                                style_name = ""

                            shape_info["paragraph_styles"].append(style_name)
                            shape_info["paragraph_texts"].append(para_text)
                            shape_info["paragraph_alignments"].append(
                                normalized(para.Format.Alignment)
                            )
                            shape_info["paragraph_fonts"].append(font_signature(para.Range))
                        except:
                            pass
                    try:
                        fields = tr.Fields
                        for j in range(1, min(fields.Count + 1, 30)):
                            try:
                                field = fields(j)
                                code_text = ""
                                try:
                                    code_text = field.Code.Text if field.Code else ""
                                except:
                                    code_text = ""
                                code_text = (code_text or "").strip()
                                if code_text:
                                    normalized_code = " ".join(code_text.upper().split())
                                    shape_info["field_codes"].append(code_text[:120])
                                    shape_info["field_fonts"].append(font_signature(field.Result))
                                    if "NUMPAGES" in normalized_code:
                                        shape_info["field_keywords"].append("NUMPAGES")
                                        shape_info["num_pages_field_count"] += 1
                                    elif "PAGE" in normalized_code:
                                        shape_info["field_keywords"].append("PAGE")
                                        shape_info["page_field_count"] += 1
                                    else:
                                        shape_info["field_keywords"].append(
                                            normalized_code.split(" ")[0]
                                        )
                            except:
                                pass
                    except:
                        pass
            except:
                pass
            state = ObservedProperties(shape_info)
            return state

        def append_shape_with_group_items(shape, index: str, story: str) -> None:
            shapes.append(_shape_text_info(shape, index, story))
            if not include_group_items:
                return
            try:
                if shape.Type != 6:
                    return
                group_items = shape.GroupItems
                for group_index in range(1, group_items.Count + 1):
                    append_shape_with_group_items(
                        group_items(group_index),
                        f"{index}.{group_index}",
                        story,
                    )
            except Exception:
                return

        for i in range(1, doc.Shapes.Count + 1):
            try:
                append_shape_with_group_items(doc.Shapes(i), str(i), "body")
            except:
                pass

        return shapes
    finally:
        close_document(doc)


def extract_shape_placement(
    docx_path: str,
    heading_text: str,
    auto_shape_type: Optional[int] = None,
    *,
    close_document,
    open_document,
) -> Dict[str, Any]:
    """Locate a floating AutoShape relative to an instruction heading."""
    doc = None
    try:
        doc = open_document(docx_path)
        heading = doc.Content.Duplicate
        heading.Find.ClearFormatting()
        heading.Find.Text = heading_text
        if not heading.Find.Execute():
            return {
                "found": False,
                "error": f"Heading not found: {heading_text}",
            }
        heading_top = float(heading.Information(6))
        result = {
            "found": False,
            "heading_text": heading_text,
            "heading_start": int(heading.Start),
            "heading_top": round(heading_top, 2),
        }
        for index in range(1, doc.Shapes.Count + 1):
            shape = doc.Shapes(index)
            try:
                if auto_shape_type is not None and int(shape.AutoShapeType) != int(auto_shape_type):
                    continue
                relative_vertical = int(shape.RelativeVerticalPosition)
                raw_top = float(shape.Top)
                anchor_top = float(shape.Anchor.Information(6))
                if relative_vertical == 1:  # page
                    object_top = raw_top
                elif relative_vertical == 0:  # margin
                    object_top = float(doc.PageSetup.TopMargin) + raw_top
                else:
                    object_top = anchor_top + raw_top
                result.update(
                    {
                        "found": True,
                        "object_kind": "floating_shape",
                        "object_start": int(shape.Anchor.Start),
                        "object_top": round(object_top, 2),
                        "relative_vertical_position": relative_vertical,
                        "below_heading": object_top > heading_top + 1.0,
                    }
                )
                return result
            except Exception:
                continue
        return result
    finally:
        close_document(doc)


def find_shapes_by_type(docx_path: str, shape_type: int, *, close_document, open_document) -> Dict:
    """
    Find shapes of a specific type in document.

    Args:
        docx_path: Path to document
        shape_type: msoShapeType value (e.g., 20 for msoCanvas)

    Returns:
        Dict with:
        - found: Whether any shape of the type was found
        - count: Number of shapes of the type
        - shapes: List of matching shape details
    """
    doc = None
    try:
        doc = open_document(docx_path)

        result = {"found": False, "count": 0, "shapes": []}

        try:
            for i in range(1, doc.Shapes.Count + 1):
                shape = doc.Shapes(i)
                if shape.Type == shape_type:
                    result["found"] = True
                    result["count"] += 1
                    result["shapes"].append(
                        {
                            "index": i,
                            "name": shape.Name,
                            "type": shape.Type,
                            "left": round(shape.Left, 2),
                            "top": round(shape.Top, 2),
                            "width": round(shape.Width, 2),
                            "height": round(shape.Height, 2),
                        }
                    )
        except:
            pass

        return result

    finally:
        close_document(doc)
