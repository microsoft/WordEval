"""families / pictures / extract_com with explicit evidence dependencies."""

from ...observations import ObservedProperties
from .camera_xml import read_inline_camera, recover_camera_properties
from typing import Dict


def extract_picture_crop_properties(
    docx_path: str, shape_index: int = 1, *, close_document, open_document
) -> Dict:
    """
    Extract crop properties from an inline picture.

    Used for PictureFillCrop, PictureFitCrop, and other picture crop commands.

    Args:
        docx_path: Path to document
        shape_index: 1-based index of the inline shape (default: 1)

    Returns:
        Dict with crop properties (CropLeft, CropRight, CropTop, CropBottom)
        and dimensions (width, height), or error dict if extraction fails.
    """
    doc = None
    try:
        doc = open_document(docx_path)

        if doc.InlineShapes.Count < shape_index:
            return {"error": f"No inline shape at index {shape_index}"}

        shape = doc.InlineShapes(shape_index)

        # Check if it's a picture (Type 3 = wdInlineShapePicture)
        if shape.Type != 3:
            return {"error": f"Inline shape is not a picture (type={shape.Type})"}

        try:
            pf = shape.PictureFormat
            return {
                "crop_left": round(pf.CropLeft, 2),
                "crop_right": round(pf.CropRight, 2),
                "crop_top": round(pf.CropTop, 2),
                "crop_bottom": round(pf.CropBottom, 2),
                "width": round(shape.Width, 2),
                "height": round(shape.Height, 2),
                "type": shape.Type,
            }
        except Exception as e:
            return {"error": f"Failed to get PictureFormat: {e}"}

    finally:
        close_document(doc)


def extract_picture_format_properties(
    docx_path: str, shape_index: int = 1, *, close_document, open_document
) -> Dict:
    """Extract picture format properties (brightness, contrast, crop, dimensions)."""
    doc = None
    try:
        doc = open_document(docx_path)

        if doc.InlineShapes.Count < shape_index:
            return {"error": f"No inline shape at index {shape_index}"}

        shape = doc.InlineShapes(shape_index)
        if shape.Type != 3:
            return {"error": f"Inline shape is not a picture (type={shape.Type})"}

        try:
            pf = shape.PictureFormat
            return {
                "brightness": round(pf.Brightness, 4),
                "contrast": round(pf.Contrast, 4),
                "crop_left": round(pf.CropLeft, 2),
                "crop_right": round(pf.CropRight, 2),
                "crop_top": round(pf.CropTop, 2),
                "crop_bottom": round(pf.CropBottom, 2),
                "width": round(shape.Width, 2),
                "height": round(shape.Height, 2),
            }
        except Exception as e:
            return {"error": f"Failed to get PictureFormat: {e}"}

    finally:
        close_document(doc)


def extract_picture_effect_properties(
    docx_path: str, shape_index: int = 1, allow_floating: bool = False, *, open_document
) -> Dict:
    """
    Extract picture effect properties from a picture object.

    Uses InlineShape properties and ConvertToShape for additional
    effect properties not available on InlineShape directly.

    Returns dict with picture effect properties:
    - shadow_type, reflection_type, soft_edge_radius
    - glow_radius, glow_color_rgb
    - threed_bevel_top_type, threed_bevel_top_depth
    - threed_rotation_x, threed_rotation_y, threed_preset_camera
    - auto_shape_type (for PictureShapeGallery)
    - brightness, contrast (for PictureSharpenSoftenGallery)
    - color_type, color_temperature (for PictureRecolorGallery, PictureColorTempertatureGallery)
    - picture_saturation (for PictureSaturationGallery)
    """
    doc = None
    try:
        doc = open_document(docx_path, read_only=False)

        ils = None
        shape = None
        restore_inline = False
        if doc.InlineShapes.Count >= shape_index:
            ils = doc.InlineShapes(shape_index)
            visual_object = ils
            object_kind = "inline"
        elif allow_floating and doc.Shapes.Count >= shape_index:
            shape = doc.Shapes(shape_index)
            visual_object = shape
            object_kind = "floating"
        else:
            expected = "inline or floating picture" if allow_floating else "inline shape"
            return {"error": f"No {expected} at index {shape_index}"}

        result = ObservedProperties({
            "type": visual_object.Type,
            "object_kind": object_kind,
            "width": round(visual_object.Width, 2),
            "height": round(visual_object.Height, 2),
        })

        # Extract PictureFormat properties (brightness, contrast, color, crop)
        result.read('brightness', lambda: visual_object.PictureFormat.Brightness, lambda value: round(value, 4))
        result.read('contrast', lambda: visual_object.PictureFormat.Contrast, lambda value: round(value, 4))
        result.read('color_type', lambda: visual_object.PictureFormat.ColorType)
        result.read('crop_left', lambda: visual_object.PictureFormat.CropLeft, lambda value: round(value, 2))
        result.read('crop_right', lambda: visual_object.PictureFormat.CropRight, lambda value: round(value, 2))
        result.read('crop_top', lambda: visual_object.PictureFormat.CropTop, lambda value: round(value, 2))
        result.read('crop_bottom', lambda: visual_object.PictureFormat.CropBottom, lambda value: round(value, 2))

        common = {
            "shadow_type": (lambda obj: obj.Shadow.Type, lambda v: v),
            "reflection_type": (lambda obj: obj.Reflection.Type, lambda v: v),
            "soft_edge_radius": (lambda obj: obj.SoftEdge.Radius, lambda v: round(v, 2)),
            "glow_radius": (lambda obj: obj.Glow.Radius, lambda v: round(v, 2)),
            "glow_color_rgb": (lambda obj: obj.Glow.Color.RGB, lambda v: v),
            "line_visible": (lambda obj: obj.Line.Visible, bool),
            "line_fore_color_rgb": (lambda obj: obj.Line.ForeColor.RGB, lambda v: v),
            "line_dash_style": (lambda obj: obj.Line.DashStyle, lambda v: v),
            "line_weight": (lambda obj: obj.Line.Weight, lambda v: round(v, 2)),
            "fill_transparency": (lambda obj: obj.Fill.Transparency, lambda v: round(v, 4)),
        }
        shape_only = {
            "threed_bevel_top_type": (lambda obj: obj.ThreeD.BevelTopType, lambda v: v),
            "threed_bevel_top_depth": (lambda obj: obj.ThreeD.BevelTopDepth, lambda v: round(v, 1)),
            "threed_rotation_x": (lambda obj: obj.ThreeD.RotationX, lambda v: round(v, 1)),
            "threed_rotation_y": (lambda obj: obj.ThreeD.RotationY, lambda v: round(v, 1)),
            "threed_preset_camera": (lambda obj: obj.ThreeD.PresetCamera, lambda v: v),
            "auto_shape_type": (lambda obj: obj.AutoShapeType, lambda v: v),
        }

        def read_properties(obj, properties, *, retry_unavailable=False):
            for name, (getter, normalize) in properties.items():
                if not retry_unavailable or name in result.unavailable:
                    result.read(name, lambda getter=getter: getter(obj), normalize)

        def picture_effects(obj):
            effects = []
            collection = obj.Fill.PictureEffects
            for index in range(1, collection.Count + 1):
                effect = collection.Item(index)
                parameters = effect.EffectParameters
                effects.append({"type": effect.Type,
                    "params": [parameters.Item(i).Value for i in range(1, parameters.Count + 1)]})
            return effects

        # InlineShape exposes these effects directly. Read them before a
        # conversion can invalidate the inline proxy or fail on its container.
        read_properties(visual_object, common)
        result.read("picture_effects", lambda: picture_effects(visual_object))
        camera, camera_error = None, "camera fallback requires an original inline picture"
        if ils is not None:
            try:
                camera = read_inline_camera(ils)
            except Exception as exc:
                camera_error = exc
        try:
            if shape is None:
                shape = ils.ConvertToShape()
                restore_inline = True
        except Exception as exc:
            result["convert_error"] = str(exc)
            for name in shape_only:
                result.record_unavailable(name, exc)
        else:
            read_properties(shape, common, retry_unavailable=True)
            if "picture_effects" in result.unavailable:
                result.read("picture_effects", lambda: picture_effects(shape))
            read_properties(shape, shape_only)
        recover_camera_properties(result, camera, camera_error)
        if "picture_effects" in result.unavailable:
            result["picture_effects"] = []
            result.setdefault("_extraction_errors", []).append({
                "property": "picture_effects", "error": result.unavailable["picture_effects"]})

        if restore_inline:
            try:
                shape.ConvertToInlineShape()
            except:
                pass


        return result

    finally:
        # Don't save changes — we converted shape back already
        if doc is not None:
            try:
                doc.Close(False)
            except:
                pass


def extract_inline_shapes(docx_path: str, *, close_document, open_document) -> Dict:
    """Extract inline shapes (including horizontal lines, images)."""
    doc = None
    try:
        doc = open_document(docx_path)

        inline_count = doc.InlineShapes.Count
        shape_count = doc.Shapes.Count

        inline_shapes = []
        for i in range(1, min(inline_count + 1, 20)):
            try:
                shape = doc.InlineShapes(i)
                info = {
                    "index": i,
                    "type": shape.Type,
                    "width": round(shape.Width, 2),
                    "height": round(shape.Height, 2),
                }
                try:
                    info["range_start"] = shape.Range.Start
                    info["range_end"] = shape.Range.End
                    info["paragraph_text"] = shape.Range.Paragraphs(1).Range.Text
                except:
                    info["range_start"] = None
                    info["range_end"] = None
                    info["paragraph_text"] = ""
                try:
                    info["alternative_text"] = shape.AlternativeText or ""
                except:
                    info["alternative_text"] = ""
                try:
                    info["title"] = shape.Title or ""
                except:
                    info["title"] = ""
                try:
                    info["has_smartart"] = bool(shape.HasSmartArt)
                except:
                    info["has_smartart"] = False
                if info["has_smartart"]:
                    try:
                        info["smartart_layout_name"] = shape.SmartArt.Layout.Name or ""
                    except:
                        info["smartart_layout_name"] = ""
                    try:
                        info["smartart_color_name"] = shape.SmartArt.Color.Name or ""
                    except:
                        info["smartart_color_name"] = ""
                    try:
                        info["smartart_quick_style_name"] = shape.SmartArt.QuickStyle.Name or ""
                    except:
                        info["smartart_quick_style_name"] = ""
                    try:
                        info["smartart_node_count"] = shape.SmartArt.Nodes.Count
                    except:
                        info["smartart_node_count"] = None
                else:
                    info["smartart_layout_name"] = ""
                    info["smartart_color_name"] = ""
                    info["smartart_quick_style_name"] = ""
                    info["smartart_node_count"] = None
                inline_shapes.append(info)
            except:
                pass

        return {
            "inline_shape_count": inline_count,
            "shape_count": shape_count,
            "total_shapes": inline_count + shape_count,
            "inline_shapes": inline_shapes,
        }

    finally:
        close_document(doc)
