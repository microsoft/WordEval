"""families / equations / insert_saved with explicit evidence dependencies."""

from pathlib import Path
from typing import Any, Dict, List, Optional
import re
import unicodedata
import xml.etree.ElementTree as ET
import zipfile


def normalize_equation_text(text: str) -> str:
    """Normalize visually equivalent math alphabet and superscript text."""
    superscripts = str.maketrans(
        {
            "⁰": "^(0)",
            "¹": "^(1)",
            "²": "^(2)",
            "³": "^(3)",
            "⁴": "^(4)",
            "⁵": "^(5)",
            "⁶": "^(6)",
            "⁷": "^(7)",
            "⁸": "^(8)",
            "⁹": "^(9)",
            "⁺": "^(+)",
            "⁻": "^(-)",
        }
    )
    translated = (text or "").translate(superscripts)
    normalized = unicodedata.normalize("NFKC", translated)
    return re.sub(r"\s+", "", normalized)


def saved_equation_expression(equation: ET.Element) -> str:
    """Project OMML into a formatting-insensitive semantic expression.

    The projection intentionally distinguishes a real ``m:sSup`` structure
    from baseline text such as ``r2``.  Run splits, Cambria Math defaults,
    control properties, and mathematical-italic Unicode codepoints are not
    semantic differences for these insertion cases.
    """
    m = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    ns = {"m": m}

    def local_name(element: ET.Element) -> str:
        return element.tag.rsplit("}", 1)[-1]

    def child(parent: ET.Element, name: str) -> Optional[ET.Element]:
        return parent.find(f"m:{name}", ns)

    def visit(element: Optional[ET.Element]) -> str:
        if element is None:
            return ""
        local = local_name(element)
        if local == "t":
            return normalize_equation_text(element.text or "")
        if local.endswith("Pr") or local in {"ctrlPr", "argPr"}:
            return ""
        if local == "sSup":
            return f'{visit(child(element, "e"))}^({visit(child(element, "sup"))})'
        if local == "sSub":
            return f'{visit(child(element, "e"))}_({visit(child(element, "sub"))})'
        if local == "sSubSup":
            return (
                f'{visit(child(element, "e"))}'
                f'_({visit(child(element, "sub"))})'
                f'^({visit(child(element, "sup"))})'
            )
        if local == "f":
            return f'({visit(child(element, "num"))})/' f'({visit(child(element, "den"))})'
        return "".join(visit(item) for item in element)

    return visit(equation)


def extract_saved_equation_state(
    docx_path: Path,
    *,
    allow_equation_gallery_wrapper: bool = False,
    allow_native_equation_wrapper: bool = False,
    extract_saved_insertion_context,
    normalize_word_default_layout_twips,
) -> Dict[str, Any]:
    """Extract persistent equation semantics, anchor, and preservation state."""
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    m = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    ns = {"w": w, "m": m}
    with zipfile.ZipFile(docx_path) as package:
        root = ET.fromstring(package.read("word/document.xml"))

    parent_map = {child: parent for parent in root.iter() for child in parent}
    body = root.find("w:body", ns)
    body_children = list(body) if body is not None else []

    def ancestor(element: ET.Element, tag: str) -> Optional[ET.Element]:
        current = element
        while current in parent_map:
            current = parent_map[current]
            if current.tag == tag:
                return current
        return None

    def body_child(element: ET.Element) -> Optional[ET.Element]:
        current = element
        while current in parent_map and parent_map[current] is not body:
            current = parent_map[current]
        return current if body is not None and parent_map.get(current) is body else None

    def paragraph_sequence(
        paragraph: Optional[ET.Element],
        target: ET.Element,
    ) -> tuple[str, str]:
        if paragraph is None:
            return "", ""
        sequence: List[tuple[str, str]] = []

        def walk(element: ET.Element) -> None:
            if element is target:
                sequence.append(("equation", ""))
                return
            if element.tag == f"{{{w}}}t":
                sequence.append(("text", element.text or ""))
                return
            for item in element:
                walk(item)

        walk(paragraph)
        marker = next(
            (index for index, item in enumerate(sequence) if item[0] == "equation"),
            -1,
        )
        if marker < 0:
            return "", ""
        before = "".join(value for kind, value in sequence[:marker] if kind == "text")
        after = "".join(value for kind, value in sequence[marker + 1 :] if kind == "text")
        return before, after

    def has_only_whitespace(element: ET.Element) -> bool:
        return not (element.text or "").strip() and all(
            not (child.tail or "").strip() for child in element
        )

    def empty_element(
        element: ET.Element,
        *,
        allowed_attributes: Optional[set[str]] = None,
    ) -> bool:
        return (
            not list(element)
            and has_only_whitespace(element)
            and set(element.attrib).issubset(allowed_attributes or set())
        )

    def transparent_native_equation_wrapper(
        sdt: Optional[ET.Element],
        equation: ET.Element,
    ) -> bool:
        """Recognize the exact native EquationInsertNew placeholder SDT.

        Merely declaring ``w:equation`` does not make an arbitrary content
        control part of the requested target.  Only Word's inert native
        placeholder wrapper may be removed from the surrounding-content
        control count, and only for the EquationInsertNew route.
        """
        if not allow_native_equation_wrapper or sdt is None:
            return False
        sdt_children = list(sdt)
        property_nodes = [child for child in sdt_children if child.tag == f"{{{w}}}sdtPr"]
        content_nodes = [child for child in sdt_children if child.tag == f"{{{w}}}sdtContent"]
        end_nodes = [child for child in sdt_children if child.tag == f"{{{w}}}sdtEndPr"]
        if (
            len(property_nodes) != 1
            or len(content_nodes) != 1
            or len(end_nodes) > 1
            or len(sdt_children) != len(property_nodes) + len(content_nodes) + len(end_nodes)
            or sdt.attrib
            or not has_only_whitespace(sdt)
        ):
            return False

        properties = property_nodes[0]
        content = content_nodes[0]
        property_children = list(properties)
        expected_property_tags = {
            f"{{{w}}}rPr",
            f"{{{w}}}id",
            f"{{{w}}}placeholder",
            f"{{{w}}}temporary",
            f"{{{w}}}showingPlcHdr",
            f"{{{w}}}equation",
        }
        if (
            len(property_children) != len(expected_property_tags)
            or {child.tag for child in property_children} != expected_property_tags
            or properties.attrib
            or not has_only_whitespace(properties)
        ):
            return False
        property_by_tag = {child.tag: child for child in property_children}

        run_properties = property_by_tag[f"{{{w}}}rPr"]
        run_property_children = list(run_properties)
        if (
            len(run_property_children) != 2
            or {child.tag for child in run_property_children} != {f"{{{w}}}rFonts", f"{{{w}}}i"}
            or run_properties.attrib
            or not has_only_whitespace(run_properties)
        ):
            return False
        run_font = next(child for child in run_property_children if child.tag == f"{{{w}}}rFonts")
        italic = next(child for child in run_property_children if child.tag == f"{{{w}}}i")
        if (
            not empty_element(
                run_font,
                allowed_attributes={f"{{{w}}}ascii", f"{{{w}}}hAnsi"},
            )
            or set(run_font.attrib) != {f"{{{w}}}ascii", f"{{{w}}}hAnsi"}
            or run_font.get(f"{{{w}}}ascii") != "Cambria Math"
            or run_font.get(f"{{{w}}}hAnsi") != "Cambria Math"
            or not empty_element(
                italic,
                allowed_attributes={f"{{{w}}}val"},
            )
            or (
                italic.get(f"{{{w}}}val") is not None
                and italic.get(f"{{{w}}}val", "").casefold() not in {"1", "true", "on"}
            )
        ):
            return False

        control_id = property_by_tag[f"{{{w}}}id"]
        control_id_value = control_id.get(f"{{{w}}}val")
        if (
            not empty_element(
                control_id,
                allowed_attributes={f"{{{w}}}val"},
            )
            or set(control_id.attrib) != {f"{{{w}}}val"}
            or not control_id_value
            or re.fullmatch(r"-?\d+", control_id_value) is None
        ):
            return False

        placeholder = property_by_tag[f"{{{w}}}placeholder"]
        placeholder_children = list(placeholder)
        if (
            len(placeholder_children) != 1
            or placeholder_children[0].tag != f"{{{w}}}docPart"
            or placeholder.attrib
            or not has_only_whitespace(placeholder)
        ):
            return False
        placeholder_part = placeholder_children[0]
        placeholder_value = placeholder_part.get(f"{{{w}}}val")
        if (
            not empty_element(
                placeholder_part,
                allowed_attributes={f"{{{w}}}val"},
            )
            or set(placeholder_part.attrib) != {f"{{{w}}}val"}
            or not placeholder_value
            or re.fullmatch(r"DefaultPlaceholder(?:_\d+)?", placeholder_value) is None
        ):
            return False

        temporary = property_by_tag[f"{{{w}}}temporary"]
        showing_placeholder = property_by_tag[f"{{{w}}}showingPlcHdr"]
        equation_property = property_by_tag[f"{{{w}}}equation"]
        for on_off_property in (temporary, showing_placeholder):
            if not empty_element(
                on_off_property,
                allowed_attributes={f"{{{w}}}val"},
            ) or (
                on_off_property.get(f"{{{w}}}val") is not None
                and on_off_property.get(f"{{{w}}}val", "").casefold() not in {"1", "true", "on"}
            ):
                return False
        if not empty_element(equation_property):
            return False

        end_properties = end_nodes[0] if end_nodes else None
        if end_properties is not None and not empty_element(end_properties):
            return False
        if content.attrib or not has_only_whitespace(content) or list(content) != [equation]:
            return False
        return content.findall(".//m:oMath", ns) == [equation]

    def transparent_equation_gallery_wrapper(
        sdt: Optional[ET.Element],
        equation: ET.Element,
    ) -> bool:
        """Recognize target-only native Equations-gallery chrome.

        Word can persist an inserted gallery equation inside a docPart SDT
        even when the reference stores the same OMML directly.  Treat only
        the exact, inert gallery wrapper as transparent; any placeholder,
        lock, extra equation, visible text, object, paragraph, or nested
        content control remains preservation-relevant.
        """
        if not allow_equation_gallery_wrapper or sdt is None:
            return False
        sdt_children = list(sdt)
        property_nodes = [child for child in sdt_children if child.tag == f"{{{w}}}sdtPr"]
        content_nodes = [child for child in sdt_children if child.tag == f"{{{w}}}sdtContent"]
        end_nodes = [child for child in sdt_children if child.tag == f"{{{w}}}sdtEndPr"]
        if (
            len(property_nodes) != 1
            or len(content_nodes) != 1
            or len(end_nodes) > 1
            or len(sdt_children) != len(property_nodes) + len(content_nodes) + len(end_nodes)
        ):
            return False
        properties = property_nodes[0]
        content = content_nodes[0]
        if properties is None or content is None:
            return False
        property_children = list(properties)
        doc_parts = [child for child in property_children if child.tag == f"{{{w}}}docPartObj"]
        ids = [child for child in property_children if child.tag == f"{{{w}}}id"]
        if (
            len(doc_parts) != 1
            or len(ids) > 1
            or len(property_children) != len(doc_parts) + len(ids)
        ):
            return False
        doc_part = doc_parts[0]
        gallery = doc_part.find("w:docPartGallery", ns) if doc_part is not None else None
        gallery_value = gallery.get(f"{{{w}}}val") if gallery is not None else None
        unique = doc_part.find("w:docPartUnique", ns) if doc_part is not None else None
        unique_value = unique.get(f"{{{w}}}val") if unique is not None else None
        if (
            not gallery_value
            or gallery_value.casefold() != "equations"
            or unique is None
            or (unique_value is not None and unique_value.casefold() not in {"1", "true", "on"})
        ):
            return False
        if {child.tag for child in doc_part} != {
            f"{{{w}}}docPartGallery",
            f"{{{w}}}docPartUnique",
        } or len(list(doc_part)) != 2:
            return False
        if any(
            properties.find(tag, ns) is not None
            for tag in (
                "w:placeholder",
                "w:showingPlcHdr",
                "w:lock",
                "w:dataBinding",
            )
        ):
            return False
        end_properties = end_nodes[0] if end_nodes else None
        if end_properties is not None and (
            list(end_properties) or end_properties.attrib or (end_properties.text or "").strip()
        ):
            return False
        content_children = list(content)
        if len(content_children) != 1 or content_children[0].tag != f"{{{w}}}p":
            return False
        paragraph = content_children[0]
        paragraph_children = list(paragraph)
        if len(paragraph_children) != 1 or paragraph_children[0].tag != f"{{{m}}}oMathPara":
            return False
        math_paragraph = paragraph_children[0]
        if list(math_paragraph) != [equation]:
            return False
        return content.findall(".//m:oMath", ns) == [equation]

    records = []
    equation_sdts: set[ET.Element] = set()
    transparent_native_equation_sdts: set[ET.Element] = set()
    transparent_gallery_sdts: set[ET.Element] = set()
    for index, equation in enumerate(root.findall(".//m:oMath", ns), start=1):
        paragraph = ancestor(equation, f"{{{w}}}p")
        math_paragraph = ancestor(equation, f"{{{m}}}oMathPara")
        sdt = ancestor(equation, f"{{{w}}}sdt")
        equation_control = False
        showing_placeholder = False
        if sdt is not None:
            properties = sdt.find("w:sdtPr", ns)
            equation_control = (
                properties is not None and properties.find("w:equation", ns) is not None
            )
            showing_placeholder = (
                properties is not None and properties.find("w:showingPlcHdr", ns) is not None
            )
            if equation_control:
                equation_sdts.add(sdt)
                if transparent_native_equation_wrapper(sdt, equation):
                    transparent_native_equation_sdts.add(sdt)
            elif transparent_equation_gallery_wrapper(sdt, equation):
                transparent_gallery_sdts.add(sdt)
        before, after = paragraph_sequence(paragraph, equation)
        top_level = body_child(equation)
        records.append(
            {
                "index": index,
                "expression": saved_equation_expression(equation),
                "raw_text": "".join(node.text or "" for node in equation.findall(".//m:t", ns)),
                "superscript_count": len(equation.findall(".//m:sSup", ns)),
                "equation_control": equation_control,
                "showing_placeholder": showing_placeholder,
                "container": (
                    "display_math"
                    if math_paragraph is not None
                    else "inline_paragraph" if paragraph is not None else "other"
                ),
                "body_index": (
                    body_children.index(top_level) + 1
                    if top_level is not None and top_level in body_children
                    else None
                ),
                "preceding_text": before,
                "following_text": after,
            }
        )

    context = extract_saved_insertion_context(docx_path)
    ignored_target_sdts = transparent_native_equation_sdts | transparent_gallery_sdts
    context["content_control_count"] = max(
        0,
        int(context.get("content_control_count") or 0) - len(ignored_target_sdts),
    )
    context["section_layout"] = normalize_word_default_layout_twips(
        context.get("section_layout") or []
    )
    return {
        "equation_count": len(records),
        "equations": records,
        "equation_control_count": len(equation_sdts),
        "transparent_native_equation_wrapper_count": len(transparent_native_equation_sdts),
        "transparent_gallery_wrapper_count": len(transparent_gallery_sdts),
        "context": context,
    }
