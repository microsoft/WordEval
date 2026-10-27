"""families / document_compositions / extract_saved with explicit evidence dependencies."""

from pathlib import Path
from typing import Any, Dict, List
import re
import xml.etree.ElementTree as ET
import zipfile


def extract_saved_enrichment_signature(docx_path: Path) -> Dict[str, Any]:
    """Read enrichment semantics that are fully persisted in a DOCX.

    This deliberately ignores unrelated body text and page-number
    alignment.  The composition asks for real note references, PAGE
    fields, a citation tied to the existing source tag, and a real
    bibliography field; each of those has an unambiguous saved-package
    representation.
    """
    w_uri = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    rel_uri = "http://schemas.openxmlformats.org/package/2006/relationships"
    office_rel_base = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    w = f"{{{w_uri}}}"
    rel = f"{{{rel_uri}}}"
    ns = {"w": w_uri}

    with zipfile.ZipFile(docx_path) as package:
        names = set(package.namelist())
        document = ET.fromstring(package.read("word/document.xml"))
        relationships = (
            ET.fromstring(package.read("word/_rels/document.xml.rels"))
            if "word/_rels/document.xml.rels" in names
            else None
        )

        def note_signature(kind: str) -> Dict[str, Any]:
            part_name = f"word/{kind}s.xml"
            if part_name not in names:
                return {
                    "references": [],
                    "orphan_notes": [],
                    "relationship_count": 0,
                }
            notes_root = ET.fromstring(package.read(part_name))
            note_texts = {}
            for note in notes_root.findall(f"w:{kind}", ns):
                note_id = note.get(f"{w}id")
                note_type = note.get(f"{w}type", "")
                try:
                    numeric_id = int(note_id) if note_id is not None else -1
                except ValueError:
                    numeric_id = -1
                if numeric_id <= 0:
                    continue
                if note_type in {
                    "separator",
                    "continuationSeparator",
                    "continuationNotice",
                }:
                    continue
                note_texts[note_id] = " ".join(
                    "".join(node.text or "" for node in note.findall(".//w:t", ns)).split()
                )

            references = []
            referenced_ids = set()
            reference_tag = f"{w}{kind}Reference"
            for paragraph_index, paragraph in enumerate(
                document.findall(".//w:body/w:p", ns), start=1
            ):
                offset = 0
                for node in paragraph.iter():
                    if node.tag == f"{w}t":
                        offset += len(node.text or "")
                    elif node.tag == reference_tag:
                        note_id = node.get(f"{w}id")
                        if note_id is not None:
                            referenced_ids.add(note_id)
                        references.append(
                            {
                                "paragraph_index": paragraph_index,
                                "character_offset": offset,
                                "note_text": note_texts.get(note_id),
                                "valid_id": (note_id is not None and note_id in note_texts),
                            }
                        )

            relationship_count = 0
            if relationships is not None:
                relationship_count = sum(
                    1
                    for relationship in relationships.findall(f"{rel}Relationship")
                    if relationship.get("Type") == f"{office_rel_base}{kind}s"
                )
            return {
                "references": references,
                "orphan_notes": sorted(
                    text for note_id, text in note_texts.items() if note_id not in referenced_ids
                ),
                "relationship_count": relationship_count,
            }

        def field_instructions(root: ET.Element) -> List[str]:
            instructions = [node.get(f"{w}instr", "") for node in root.iter(f"{w}fldSimple")]
            instructions.extend(node.text or "" for node in root.iter(f"{w}instrText"))
            return [" ".join(value.split()) for value in instructions if value]

        story_page_fields = {"header": 0, "footer": 0}
        for story_type in story_page_fields:
            for name in names:
                if not re.fullmatch(rf"word/{story_type}\d+\.xml", name):
                    continue
                root = ET.fromstring(package.read(name))
                story_page_fields[story_type] += sum(
                    1
                    for instruction in field_instructions(root)
                    if re.search(r"\bPAGE\b", instruction, re.IGNORECASE)
                )

        footnote_signature = note_signature("footnote")
        endnote_signature = note_signature("endnote")
        document_fields = field_instructions(document)

    citation_tags = sorted(
        match.group(1)
        for instruction in document_fields
        for match in [re.search(r"\bCITATION\s+([^\s\\]+)", instruction, re.IGNORECASE)]
        if match
    )
    return {
        "footnote": footnote_signature,
        "endnote": endnote_signature,
        "header_page_field_count": story_page_fields["header"],
        "footer_page_field_count": story_page_fields["footer"],
        "citation_tags": citation_tags,
        "bibliography_field_count": sum(
            1
            for instruction in document_fields
            if re.search(r"\bBIBLIOGRAPHY\b", instruction, re.IGNORECASE)
        ),
    }


def extract_saved_cleanup_signature(docx_path: Path) -> Dict[str, Any]:
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    v = "urn:schemas-microsoft-com:vml"
    ns = {"w": w, "v": v}
    revision_tags = {
        f"{{{w}}}ins",
        f"{{{w}}}del",
        f"{{{w}}}moveFrom",
        f"{{{w}}}moveTo",
    }
    comment_tags = {
        f"{{{w}}}commentRangeStart",
        f"{{{w}}}commentRangeEnd",
        f"{{{w}}}commentReference",
    }
    with zipfile.ZipFile(docx_path) as package:
        names = set(package.namelist())
        document = ET.fromstring(package.read("word/document.xml"))
        settings = (
            ET.fromstring(package.read("word/settings.xml"))
            if "word/settings.xml" in names
            else None
        )
        comment_count = 0
        if "word/comments.xml" in names:
            comments = ET.fromstring(package.read("word/comments.xml"))
            comment_count = len(comments.findall("w:comment", ns))

        watermark_count = 0
        header_payloads = []
        footer_payloads = []
        for name in names:
            if re.fullmatch(r"word/header\d+\.xml", name):
                root = ET.fromstring(package.read(name))
                text = "".join(node.text or "" for node in root.findall(".//w:t", ns))
                drawings = sum(
                    1
                    for node in root.iter()
                    if node.tag
                    in {
                        f"{{{w}}}drawing",
                        f"{{{w}}}pict",
                        f"{{{w}}}object",
                    }
                )
                fields = sum(
                    1
                    for node in root.iter()
                    if node.tag == f"{{{w}}}fldSimple"
                    or (
                        node.tag == f"{{{w}}}fldChar" and node.get(f"{{{w}}}fldCharType") == "begin"
                    )
                )
                if text.strip() or drawings or fields:
                    header_payloads.append(
                        {
                            "text": text.strip(),
                            "drawings": drawings,
                            "fields": fields,
                        }
                    )
                for shape in root.findall(".//v:shape", ns):
                    shape_id = shape.get("id", "")
                    shape_type = shape.get("type", "")
                    if (
                        "WaterMark" in shape_id
                        or shape_type == "#_x0000_t136"
                        or shape.find("v:textpath", ns) is not None
                    ):
                        watermark_count += 1
            elif re.fullmatch(r"word/footer\d+\.xml", name):
                root = ET.fromstring(package.read(name))
                text = "".join(node.text or "" for node in root.findall(".//w:t", ns))
                drawings = sum(
                    1
                    for node in root.iter()
                    if node.tag
                    in {
                        f"{{{w}}}drawing",
                        f"{{{w}}}pict",
                        f"{{{w}}}object",
                    }
                )
                fields = sum(
                    1
                    for node in root.iter()
                    if node.tag == f"{{{w}}}fldSimple"
                    or (
                        node.tag == f"{{{w}}}fldChar" and node.get(f"{{{w}}}fldCharType") == "begin"
                    )
                )
                if text.strip() or drawings or fields:
                    footer_payloads.append(
                        {
                            "text": text.strip(),
                            "drawings": drawings,
                            "fields": fields,
                        }
                    )

    return {
        "revision_count": sum(1 for node in document.iter() if node.tag in revision_tags),
        "track_revisions": (
            settings is not None and settings.find("w:trackRevisions", ns) is not None
        ),
        "body_text": "".join(node.text or "" for node in document.findall(".//w:body//w:t", ns)),
        "comment_marker_count": sum(1 for node in document.iter() if node.tag in comment_tags),
        "comment_count": comment_count,
        "watermark_count": watermark_count,
        "header_payloads": header_payloads,
        "footer_payloads": footer_payloads,
    }
