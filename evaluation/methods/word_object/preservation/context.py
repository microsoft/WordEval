"""preservation / context with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.families.pages import (
    extract_saved as families_pages_extract_saved,
)
from evaluation.methods.word_object.ooxml import canonical as ooxml_canonical
from pathlib import Path
from typing import Any, Dict, Optional
import re
import xml.etree.ElementTree as ET
import zipfile


def extract_saved_semantic_context(
    docx_path: Path,
    *,
    strip_note_references: bool = False,
    strip_note_properties: Optional[str] = None,
    strip_tracked_insertions: bool = False,
    strip_manual_toc: bool = False,
) -> Dict[str, Any]:
    """Extract unrelated visible/structural state for note/revision routes."""
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ns = {"w": w}
    with zipfile.ZipFile(docx_path) as package:
        names = set(package.namelist())
        original = ET.fromstring(package.read("word/document.xml"))
        root = ET.fromstring(ET.tostring(original, encoding="utf-8"))
        parent_map = {child: parent for parent in root.iter() for child in parent}

        if strip_manual_toc:
            for control in list(root.findall(".//w:sdt", ns)):
                gallery = control.find("w:sdtPr/w:docPartObj/w:docPartGallery", ns)
                if gallery is not None and "table of contents" in gallery.get(f"{{{w}}}val", "").casefold():
                    parent_map[control].remove(control)

        # Canonicalize Word's explicit endnote-position default before
        # comparing unrelated section layout.  ``docEnd`` and an omitted
        # ``w:pos`` are the same effective setting.
        for properties in list(root.iter(f"{{{w}}}endnotePr")):
            position = properties.find(f"{{{w}}}pos")
            if position is not None and position.get(f"{{{w}}}val") == "docEnd":
                properties.remove(position)
            if (
                not list(properties)
                and not properties.attrib
                and not (properties.text or "").strip()
            ):
                parent = parent_map.get(properties)
                if parent is not None and properties in list(parent):
                    parent.remove(properties)

        if strip_note_references:
            # A legal note reference does not have to occupy its own run.
            # Some Word writers place the reference marker after ordinary
            # text in the same ``w:r``. Removing that whole run would make
            # the preservation projection falsely report preceding text as
            # deleted. Remove only the target-owned marker, then discard a
            # run only when no content remains besides rPr.
            reference_tags = {
                f"{{{w}}}footnoteReference",
                f"{{{w}}}endnoteReference",
            }
            # A custom-mark note declares ``w:customMarkFollows`` and then
            # writes the mark itself as run text right after the reference.
            # That text is part of the note reference, not body content, so
            # it must be stripped with the reference; otherwise the mark
            # leaks into the preservation projection and makes a correct
            # gold look like it changed unrelated text.  Only nodes after
            # the reference are removed, so ordinary text that shares the
            # run and precedes the marker is preserved.
            mark_tags = {f"{{{w}}}t", f"{{{w}}}sym"}
            for run in list(root.findall(".//w:r", ns)):
                removed_reference = False
                for node in list(run.iter()):
                    if node.tag not in reference_tags:
                        continue
                    parent = parent_map.get(node)
                    if parent is None or node not in list(parent):
                        continue
                    custom_mark_follows = node.get(f"{{{w}}}customMarkFollows") in {
                        "1",
                        "true",
                        "on",
                    }
                    siblings = list(parent)
                    index = siblings.index(node)
                    parent.remove(node)
                    removed_reference = True
                    if not custom_mark_follows:
                        continue
                    for following in siblings[index + 1 :]:
                        if following.tag not in mark_tags:
                            continue
                        if following in list(parent):
                            parent.remove(following)
                if not removed_reference:
                    continue
                remaining_content = any(child.tag != f"{{{w}}}rPr" for child in list(run)) or bool(
                    (run.text or "").strip()
                )
                if remaining_content:
                    continue
                parent = parent_map.get(run)
                if parent is not None and run in list(parent):
                    parent.remove(run)

        if strip_note_properties:
            tag = f"{{{w}}}{strip_note_properties}Pr"
            for properties in list(root.iter(tag)):
                parent = parent_map.get(properties)
                if parent is not None and properties in list(parent):
                    parent.remove(properties)

        if strip_tracked_insertions:
            revision_tags = {
                f"{{{w}}}ins",
                f"{{{w}}}moveTo",
                f"{{{w}}}rPrChange",
                f"{{{w}}}pPrChange",
                f"{{{w}}}tblPrChange",
                f"{{{w}}}trPrChange",
                f"{{{w}}}tcPrChange",
                f"{{{w}}}sectPrChange",
            }
            for revision in list(root.iter()):
                if revision.tag not in revision_tags:
                    continue
                parent = parent_map.get(revision)
                if parent is not None and revision in list(parent):
                    parent.remove(revision)

        header_footer = []
        for name in sorted(names):
            if not re.fullmatch(r"word/(header|footer)\d+\.xml", name):
                continue
            part = ET.fromstring(package.read(name))
            text_value = "".join(node.text or "" for node in part.findall(".//w:t", ns))
            object_count = sum(
                1
                for node in part.iter()
                if node.tag
                in {
                    f"{{{w}}}drawing",
                    f"{{{w}}}pict",
                    f"{{{w}}}object",
                }
            )
            if text_value.strip() or object_count:
                header_footer.append(
                    {
                        "part_kind": "header" if "/header" in name else "footer",
                        "text": text_value,
                        "object_count": object_count,
                        "visible": ooxml_canonical.canonical_visible_element(part),
                    }
                )

        pictures = ooxml_canonical.saved_picture_signatures(original, package)

    body = root.find("w:body", ns)
    visible_objects = families_pages_extract_saved.cover_page_ooxml_layout_signature(docx_path)
    # Target-owned note/revision runs appear in this sub-signature.  Their
    # semantics are checked separately; retain object/layout signals only.
    visible_objects.pop("paragraph_formatting", None)

    def section_layout(section: ET.Element) -> Any:
        clone = ET.fromstring(ET.tostring(section, encoding="utf-8"))
        for tag in ("headerReference", "footerReference"):
            for reference in list(clone.findall(f"w:{tag}", ns)):
                clone.remove(reference)
        return ooxml_canonical.canonical_visible_element(clone)

    context = {
        "body_text": (
            "".join(node.text or "" for node in body.findall(".//w:t", ns))
            if body is not None
            else ""
        ),
        "body_visible_format": ooxml_canonical.saved_body_visible_format_signature(root),
        "section_layout": [section_layout(section) for section in root.findall(".//w:sectPr", ns)],
        "table_count": len(root.findall(".//w:body//w:tbl", ns)),
        "content_control_count": len(root.findall(".//w:body//w:sdt", ns)),
        "header_footer": header_footer,
        "pictures": pictures,
        "visible_objects": visible_objects,
        "document_background": ooxml_canonical.canonical_visible_element(
            root.find(f"{{{w}}}background")
        ),
    }
    context["section_layout"] = ooxml_canonical.normalize_word_default_layout_twips(
        context["section_layout"]
    )
    context["visible_objects"] = ooxml_canonical.normalize_word_default_layout_twips(
        context["visible_objects"]
    )
    return context


def extract_saved_insertion_context(docx_path: Path, *, target_kind="drawing") -> Dict[str, Any]:
    """Extract visible non-target content/layout for insertion routes.

    Drawings and legacy pictures are removed before the body signature is
    built.  The route-specific verifier owns those inserted targets; this
    signature owns surrounding text, effective formatting, sections,
    tables, content controls, and visible header/footer content.
    """
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ns = {"w": w}
    with zipfile.ZipFile(docx_path) as package:
        names = set(package.namelist())
        root = ET.fromstring(package.read("word/document.xml"))
        stripped = ET.fromstring(ET.tostring(root, encoding="utf-8"))

        parent_map = {child: parent for parent in stripped.iter() for child in parent}
        removable = {
            f"{{{w}}}drawing",
            f"{{{w}}}pict",
            f"{{{w}}}object",
        }
        if target_kind == "table":
            # The table verifier owns every table's structure/content. Keep
            # drawings and all text/formatting outside those tables visible.
            removable = {f"{{{w}}}tbl"}
        elif target_kind != "drawing":
            raise ValueError(f"Unknown insertion target: {target_kind}")
        for node in list(stripped.iter()):
            if node.tag not in removable:
                continue
            if target_kind == "table":
                ancestor = parent_map.get(node)
                while ancestor is not None and ancestor.tag != f"{{{w}}}txbxContent":
                    ancestor = parent_map.get(ancestor)
                if ancestor is not None:
                    continue  # Main-story table edits do not own text-box tables.
            parent = parent_map.get(node)
            if parent is not None and node in list(parent):
                parent.remove(node)

        header_footer = []
        for name in sorted(names):
            if not re.fullmatch(r"word/(header|footer)\d+\.xml", name):
                continue
            part = ET.fromstring(package.read(name))
            text = "".join(node.text or "" for node in part.findall(".//w:t", ns))
            object_count = sum(
                1
                for node in part.iter()
                if node.tag in {f"{{{w}}}drawing", f"{{{w}}}pict", f"{{{w}}}object"}
            )
            if text.strip() or object_count:
                header_footer.append(
                    {
                        "part_kind": "header" if "/header" in name else "footer",
                        "text": text,
                        "object_count": object_count,
                        "visible": ooxml_canonical.canonical_visible_element(part),
                    }
                )

    body = stripped.find("w:body", ns)
    return {
        "body_text": (
            "".join(node.text or "" for node in body.findall(".//w:t", ns))
            if body is not None
            else ""
        ),
        "body_visible_format": ooxml_canonical.saved_body_visible_format_signature(stripped),
        "section_layout": [
            ooxml_canonical.canonical_visible_element(section)
            for section in stripped.findall(".//w:sectPr", ns)
        ],
        "table_count": len(stripped.findall(".//w:body//w:tbl", ns)),
        "content_control_count": len(stripped.findall(".//w:body//w:sdt", ns)),
        "header_footer": header_footer,
    }


def verify_saved_insertion_context(
    pred_path: Path, gold_path: Path, *, ignore_body_visible_format: bool = False,
    target_kind="drawing",
) -> Optional[VerificationResult]:
    gold_context = extract_saved_insertion_context(gold_path, target_kind=target_kind)
    pred_context = extract_saved_insertion_context(pred_path, target_kind=target_kind)
    if ignore_body_visible_format:
        gold_context.pop("body_visible_format", None)
        pred_context.pop("body_visible_format", None)
    if gold_context == pred_context:
        return None
    return VerificationResult(
        passed=False,
        message="FAIL: Surrounding document content/layout changed",
        details={
            "match_mode": "saved_ooxml_insertion_context",
            "gold_context": gold_context,
            "pred_context": pred_context,
        },
    )
