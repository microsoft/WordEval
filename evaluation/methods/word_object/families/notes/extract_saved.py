"""families / notes / extract_saved with explicit evidence dependencies."""

from collections import Counter
from pathlib import Path
from typing import Any, Dict, Optional
import xml.etree.ElementTree as ET
import zipfile


def note_property_tokens(properties: Optional[ET.Element], w_uri: str) -> Dict[str, Optional[str]]:
    if properties is None:
        return {}
    result = {}
    for child in properties:
        name = child.tag.rsplit("}", 1)[-1]
        if name in {"numFmt", "numRestart", "numStart", "pos"}:
            result[name] = child.get(f"{{{w_uri}}}val")
    return result


def extract_saved_note_signature(docx_path: Path, kind: str) -> Dict[str, Any]:
    """Extract durable note objects without comparing raw note IDs."""
    if kind not in {"footnote", "endnote"}:
        raise ValueError(f"Unsupported note kind: {kind}")
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    rel = "http://schemas.openxmlformats.org/package/2006/relationships"
    office_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" f"{kind}s"
    ns = {"w": w}
    with zipfile.ZipFile(docx_path) as package:
        names = set(package.namelist())
        document = ET.fromstring(package.read("word/document.xml"))
        part_name = f"word/{kind}s.xml"
        notes = ET.fromstring(package.read(part_name)) if part_name in names else None
        relationships = (
            ET.fromstring(package.read("word/_rels/document.xml.rels"))
            if "word/_rels/document.xml.rels" in names
            else None
        )
        settings = (
            ET.fromstring(package.read("word/settings.xml"))
            if "word/settings.xml" in names
            else None
        )

    note_texts: Dict[str, str] = {}
    special_note_definitions = []
    duplicate_note_ids = []
    seen_note_ids = set()
    for note in notes.findall(f"w:{kind}", ns) if notes is not None else []:
        note_id = note.get(f"{{{w}}}id")
        note_type = note.get(f"{{{w}}}type", "")
        if note_id in seen_note_ids:
            duplicate_note_ids.append(note_id)
        elif note_id is not None:
            seen_note_ids.add(note_id)
        try:
            numeric_id = int(note_id) if note_id is not None else -1
        except ValueError:
            numeric_id = -1
        if note_type in {
            "separator",
            "continuationSeparator",
            "continuationNotice",
        }:
            special_note_definitions.append(
                {
                    "id": note_id,
                    "type": note_type,
                }
            )
            continue
        if numeric_id <= 0:
            continue
        note_texts[note_id] = " ".join(
            "".join(node.text or "" for node in note.findall(".//w:t", ns)).split()
        )

    references = []
    referenced_ids = set()
    body_text = []
    section_index = 1
    for paragraph_index, paragraph in enumerate(document.findall(".//w:body//w:p", ns), start=1):
        logical_runs = []
        offset = 0
        paragraph_references = []
        for run in paragraph.findall(".//w:r", ns):
            custom_mark_reference = None
            for node in run.iter():
                if node.tag == f"{{{w}}}{kind}Reference":
                    note_id = node.get(f"{{{w}}}id")
                    if note_id is not None:
                        referenced_ids.add(note_id)
                    custom_mark_follows = node.get(f"{{{w}}}customMarkFollows") in {
                        "1",
                        "true",
                        "on",
                    }
                    reference = {
                        "note_id": note_id,
                        "paragraph_index": paragraph_index,
                        "character_offset": offset,
                        "section_index": section_index,
                        "custom_mark_follows": False,
                        "custom_mark": "",
                    }
                    paragraph_references.append(reference)
                    custom_mark_reference = reference if custom_mark_follows else None
                elif node.tag == f"{{{w}}}t":
                    text = node.text or ""
                    if custom_mark_reference is not None:
                        custom_mark_reference["custom_mark"] += text
                        custom_mark_reference["custom_mark_follows"] = bool(
                            custom_mark_reference["custom_mark"]
                        )
                    else:
                        logical_runs.append(text)
                        offset += len(text)
        paragraph_text = "".join(logical_runs)
        body_text.append(paragraph_text)
        for reference in paragraph_references:
            note_id = reference.pop("note_id")
            full_note_text = note_texts.get(note_id)
            custom_mark = reference.get("custom_mark") or ""
            note_text = full_note_text
            if custom_mark and (note_text or "").startswith(custom_mark):
                note_text = (note_text or "")[len(custom_mark) :].lstrip()
            reference.update(
                {
                    "paragraph_text": paragraph_text,
                    "note_text": note_text,
                    "valid_id": note_id is not None and note_id in note_texts,
                }
            )
            references.append(reference)
        if paragraph.find("./w:pPr/w:sectPr", ns) is not None:
            section_index += 1

    section_properties = []
    for index, section in enumerate(document.findall(".//w:sectPr", ns), start=1):
        properties = section.find(f"w:{kind}Pr", ns)
        section_properties.append(
            {
                "section_index": index,
                "properties": note_property_tokens(properties, w),
            }
        )
    settings_properties = note_property_tokens(
        settings.find(f"w:{kind}Pr", ns) if settings is not None else None,
        w,
    )

    relationship_count = (
        sum(
            1
            for relationship in relationships.findall(f"{{{rel}}}Relationship")
            if relationship.get("Type") == office_rel
        )
        if relationships is not None
        else 0
    )
    orphan_notes = sorted(
        text for note_id, text in note_texts.items() if note_id not in referenced_ids
    )
    raw_reference_ids = [
        node.get(f"{{{w}}}id") for node in document.findall(f".//w:{kind}Reference", ns)
    ]
    duplicate_reference_ids = sorted(
        note_id
        for note_id, count in Counter(raw_reference_ids).items()
        if note_id is not None and count > 1
    )
    return {
        "kind": kind,
        f"{kind}s_part_present": notes is not None,
        "body_paragraphs": body_text,
        "references": references,
        "orphan_notes": orphan_notes,
        "special_note_definitions": special_note_definitions,
        "duplicate_note_ids": sorted(duplicate_note_ids),
        "duplicate_reference_ids": duplicate_reference_ids,
        f"{kind}s_relationship_count": relationship_count,
        "settings_properties": settings_properties,
        "section_properties": section_properties,
        "section_count": len(section_properties),
    }


def extract_saved_footnote_signature(docx_path: Path) -> Dict[str, Any]:
    return extract_saved_note_signature(docx_path, "footnote")
