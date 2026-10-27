"""families / smartart / extract_com with explicit evidence dependencies."""

from typing import Dict, Any
from .placement import read_page_bounds
from ...observations import ObservedProperties


def extract_inserted_smartart_state(
    docx_path: str, *, close_document, open_document
) -> Dict[str, Any]:
    """Strict WOM state for the atomic SmartArt-insertion contract."""
    doc = None
    try:
        doc = open_document(docx_path)
        inline_count = int(doc.InlineShapes.Count)
        floating_count = int(doc.Shapes.Count)
        matches = []
        for collection_name, object_kind in (
            ("InlineShapes", "inline"),
            ("Shapes", "floating"),
        ):
            collection = getattr(doc, collection_name)
            for index in range(1, int(collection.Count) + 1):
                candidate = collection(index)
                if bool(candidate.HasSmartArt):
                    matches.append((candidate, object_kind))

        result: Dict[str, Any] = {
            "smartart_count": len(matches),
            "visual_object_count": inline_count + floating_count,
        }
        if len(matches) != 1:
            return result

        shape, source = matches[0]
        smartart = shape.SmartArt
        all_nodes = smartart.AllNodes
        node_texts = []
        node_levels = []
        for index in range(1, int(all_nodes.Count) + 1):
            node = all_nodes(index)
            node_texts.append(str(node.TextFrame2.TextRange.Text or "").strip())
            node_levels.append(int(node.Level))
        result.update(
            {
                "source": source,
                "layout_name": str(smartart.Layout.Name or "").strip(),
                "all_node_texts": node_texts,
                "all_node_levels": node_levels,
                "page_bounds": read_page_bounds(doc, shape, source),
            }
        )
        return result
    finally:
        close_document(doc)


def _find_smartart(doc, *, inline_first=False):
    """Return an observed object or absence; let discovery failures propagate."""
    collections = [("Shapes", "floating"), ("InlineShapes", "inline")]
    if inline_first:
        collections.reverse()
    for collection_name, source in collections:
        collection = getattr(doc, collection_name)
        for index in range(1, int(collection.Count) + 1):
            shape = collection(index)
            has_smartart = shape.HasSmartArt
            if has_smartart not in (0, -1, True):
                raise ValueError(f"Unavailable SmartArt identity: {has_smartart!r}")
            if has_smartart:
                return shape, source
    return None, None


def _read_gallery_properties(result, smartart):
    for field, getter in (
        ("layout_name", lambda: smartart.Layout.Name),
        ("color_name", lambda: smartart.Color.Name),
        ("quick_style_name", lambda: smartart.QuickStyle.Name),
    ):
        result.read(field, getter)


def extract_smartart_properties(docx_path: str, *, close_document, open_document) -> Dict:
    """Read gallery identities independently from the first observed SmartArt."""
    doc = None
    try:
        doc = open_document(docx_path)
        shape, source = _find_smartart(doc, inline_first=True)
        result = ObservedProperties(found=shape is not None)
        if shape is not None:
            smartart = shape.SmartArt
            result["source"] = source
            _read_gallery_properties(result, smartart)
            result.read("node_count", lambda: smartart.Nodes.Count)
        return result
    finally:
        close_document(doc)


def extract_smartart_subnode_properties(docx_path: str, *, close_document, open_document) -> Dict:
    """Read independent properties of the uniquely named Target at any depth.

    An observed missing target is different from failed discovery. Never use
    another node as a substitute for the node named by the instruction.
    """
    doc = None
    try:
        doc = open_document(docx_path)
        shape, source = _find_smartart(doc)
        result = ObservedProperties(found=False)
        if shape is None:
            return result
        nodes = shape.SmartArt.AllNodes
        matches = []
        for index in range(1, int(nodes.Count) + 1):
            node = nodes(index)
            text = node.TextFrame2.TextRange.Text
            if text is None:
                raise ValueError("Unavailable SmartArt node text")
            if text.strip() == "Target":
                matches.append(node)
        if not matches:
            return result
        if len(matches) != 1:
            raise ValueError("Ambiguous SmartArt Target: multiple nodes have this name")
        target = matches[0]
        result.update(found=True, source=source, target_text="Target")
        for field, getter, normalize in (
            ("line_dash_style", lambda: target.Shapes(1).Line.DashStyle, int),
            ("line_weight", lambda: target.Shapes(1).Line.Weight, lambda v: round(v, 2)),
            ("line_begin_arrow", lambda: target.Shapes(1).Line.BeginArrowheadStyle, int),
            ("line_end_arrow", lambda: target.Shapes(1).Line.EndArrowheadStyle, int),
            ("text_fill_rgb", lambda: target.TextFrame2.TextRange.Font.Fill.ForeColor.RGB, int),
            ("warp_format", lambda: target.TextFrame2.WarpFormat, int),
            ("bevel_top_type", lambda: target.TextFrame2.ThreeD.BevelTopType, int),
            ("bevel_top_depth", lambda: target.TextFrame2.ThreeD.BevelTopDepth, lambda v: round(v, 1)),
            ("threed_preset_camera", lambda: target.TextFrame2.ThreeD.PresetCamera, int),
            ("threed_rotation_x", lambda: target.TextFrame2.ThreeD.RotationX, lambda v: round(v, 1)),
            ("threed_rotation_y", lambda: target.TextFrame2.ThreeD.RotationY, lambda v: round(v, 1)),
        ):
            result.read(field, getter, normalize)
        # Connector discovery only belongs to the arrow contract. Its failure
        # must not erase independently observable text or node-outline state.
        try:
            connector = None
            shapes = target.Shapes
            for index in range(1, int(shapes.Count) + 1):
                candidate = shapes(index)
                if candidate.Type == 9:
                    connector = candidate
                    break
            for field, attribute in (("connector_begin_arrow", "BeginArrowheadStyle"),
                                     ("connector_end_arrow", "EndArrowheadStyle")):
                if connector is None:
                    result.record_unavailable(field, "No connector exposed for Target")
                else:
                    result.read(field, lambda attr=attribute: getattr(connector.Line, attr), int)
        except Exception as exc:
            for field in ("connector_begin_arrow", "connector_end_arrow"):
                result.record_unavailable(field, exc)
        return result
    finally:
        close_document(doc)


def extract_smartart_placement(
    docx_path: str, heading_text: str, *, close_document, open_document
) -> Dict:
    """Locate the first SmartArt relative to a required heading via WOM.

    Floating SmartArt can be anchored to the heading paragraph while
    rendering below it, so document anchor order alone is insufficient.
    Resolve a page-relative visual top for both inline and floating objects.
    """
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

        for index in range(1, doc.InlineShapes.Count + 1):
            shape = doc.InlineShapes(index)
            if not shape.HasSmartArt:
                continue
            object_top = float(shape.Range.Information(6))
            result.update(
                {
                    "found": True,
                    "object_kind": "inline",
                    "object_start": int(shape.Range.Start),
                    "object_top": round(object_top, 2),
                    "below_heading": object_top > heading_top + 1.0,
                }
            )
            return result

        for index in range(1, doc.Shapes.Count + 1):
            shape = doc.Shapes(index)
            if not shape.HasSmartArt:
                continue
            relative_vertical = int(shape.RelativeVerticalPosition)
            raw_top = float(shape.Top)
            anchor_top = float(shape.Anchor.Information(6))
            if relative_vertical == 1:  # wdRelativeVerticalPositionPage
                object_top = raw_top
            elif relative_vertical == 0:  # margin
                object_top = float(doc.PageSetup.TopMargin) + raw_top
            else:
                # Paragraph/line-relative positions use the anchor line
                # as their rendered origin. This is how prerequisite
                # SmartArt references are positioned below the heading.
                object_top = anchor_top + raw_top
            result.update(
                {
                    "found": True,
                    "object_kind": "floating",
                    "object_start": int(shape.Anchor.Start),
                    "object_top": round(object_top, 2),
                    "relative_vertical_position": relative_vertical,
                    "below_heading": object_top > heading_top + 1.0,
                }
            )
            return result

        return result
    finally:
        close_document(doc)


def extract_smartart_node_details(docx_path: str, *, close_document, open_document) -> Dict:
    """Read complete per-property node sequences, retaining collection failures.

    A failed traversal is unavailable, never a zero count or a partial list.
    Independent columns let each structural command select its own evidence.
    """
    doc = None
    try:
        doc = open_document(docx_path)
        shape, source = _find_smartart(doc)
        result = ObservedProperties(found=shape is not None)
        if shape is None:
            return result
        smartart = shape.SmartArt
        result["source"] = source
        _read_gallery_properties(result, smartart)
        result.read("reverse", lambda: smartart.Reverse)
        for collection_name, prefix in (("Nodes", "node"), ("AllNodes", "all_node")):
            try:
                collection = getattr(smartart, collection_name)
                count = int(collection.Count)
                nodes = [collection(index) for index in range(1, count + 1)]
                result[f"{prefix}_count"] = count
            except Exception as exc:
                for suffix in ("count", "texts", "levels", "types", "sizes"):
                    result.record_unavailable(f"{prefix}_{suffix}", exc)
                continue
            def read_column(suffix, getter):
                values = []
                for node in nodes:
                    scalar = ObservedProperties()
                    scalar.read("value", lambda: getter(node))
                    if scalar.unavailable or scalar.mixed or scalar["value"] is None:
                        raise ValueError(f"Unavailable SmartArt {prefix}_{suffix}: {scalar.unavailable}")
                    values.append(scalar["value"])
                return values
            for suffix, getter in (
                ("texts", lambda node: node.TextFrame2.TextRange.Text.strip()),
                ("levels", lambda node: node.Level),
                ("types", lambda node: node.Type),
            ):
                result.read(f"{prefix}_{suffix}", lambda s=suffix, g=getter: read_column(s, g))
            if prefix == "node":
                def sizes():
                    widths = read_column("widths", lambda node: node.Shapes(1).Width)
                    heights = read_column("heights", lambda node: node.Shapes(1).Height)
                    return [{"width": round(w, 1), "height": round(h, 1)} for w, h in zip(widths, heights)]
                result.read("node_sizes", sizes)
        return result
    finally:
        close_document(doc)


def extract_smartart_reset_state(
    docx_path: str, *, close_document, open_document
) -> Dict[str, Any]:
    """Extract target-scoped effective node paint/effect state for Reset.

    ``Document.Content.Find`` does not traverse SmartArt node text.  Read
    each node through ``SmartArt.AllNodes`` and its first rendered shape so
    Reset can require effective blue fill, white glyph paint, and removal
    of residual line/shadow/text-outline state without global color or
    screenshot normalization.
    """

    doc = None

    def value(obj: Any, name: str, default: Any = None) -> Any:
        try:
            return getattr(obj, name)
        except Exception:
            return default

    def paint(fill: Any) -> Dict[str, Any]:
        foreground = value(fill, "ForeColor")
        return {
            "visible": value(fill, "Visible"),
            "type": value(fill, "Type"),
            "transparency": value(fill, "Transparency"),
            "rgb": value(foreground, "RGB"),
            "theme_color": value(foreground, "ObjectThemeColor"),
            "tint_and_shade": value(foreground, "TintAndShade"),
        }

    try:
        doc = open_document(docx_path)
        shape, source = _find_smartart(doc)
        smartart = shape.SmartArt if shape is not None else None
        if smartart is None:
            return {"found": False}

        result: Dict[str, Any] = {
            "found": True,
            "source": source,
            "layout_name": value(smartart.Layout, "Name", ""),
            "color_name": value(smartart.Color, "Name", ""),
            "quick_style_name": value(smartart.QuickStyle, "Name", ""),
            "nodes": [],
        }
        all_nodes = smartart.AllNodes
        for index in range(1, int(all_nodes.Count) + 1):
            node = all_nodes(index)
            text_range = node.TextFrame2.TextRange
            font = text_range.Font
            node_state: Dict[str, Any] = {
                "text": str(value(text_range, "Text", "") or "").strip(),
                "level": value(node, "Level"),
                "node_type": value(node, "Type"),
            }
            if int(node.Shapes.Count) == 0:
                node_state["shape_found"] = False
                result["nodes"].append(node_state)
                continue
            shape = node.Shapes(1)
            node_state["shape_found"] = True
            node_state["fill"] = paint(shape.Fill)
            node_state["text_fill"] = paint(font.Fill)
            node_state["line"] = {
                **paint(shape.Line),
                "weight": value(shape.Line, "Weight"),
                "dash_style": value(shape.Line, "DashStyle"),
            }
            node_state["text_outline"] = {
                **paint(font.Line),
                "weight": value(font.Line, "Weight"),
                "dash_style": value(font.Line, "DashStyle"),
            }
            node_state["shadow"] = {
                "visible": value(shape.Shadow, "Visible"),
                "type": value(shape.Shadow, "Type"),
                "transparency": value(shape.Shadow, "Transparency"),
                "blur": value(shape.Shadow, "Blur"),
                "offset_x": value(shape.Shadow, "OffsetX"),
                "offset_y": value(shape.Shadow, "OffsetY"),
            }
            result["nodes"].append(node_state)
        result["node_count"] = len(result["nodes"])
        return result
    finally:
        close_document(doc)
