"""OOXML package relationships and content-type diagnostics."""
import posixpath
from pathlib import PurePosixPath
from collections import Counter
from typing import Optional
from .common import w_ns

def diagnose(xml_roots, names, include_word_open_failure_schema=False):
    document_root = xml_roots.get("word/document.xml")
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    content_types_ns = (
        "http://schemas.openxmlformats.org/package/2006/content-types"
    )
    office_document_rel = (
        "http://schemas.openxmlformats.org/officeDocument/2006/"
        "relationships/officeDocument"
    )
    singleton_main_relationship_types = {
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
        + suffix
        for suffix in (
            "styles", "settings", "numbering", "fontTable", "footnotes",
            "endnotes", "comments", "theme", "webSettings",
        )
    }

    def relationship_source(rels_name: str) -> Optional[str]:
        if rels_name == "_rels/.rels":
            return ""
        rels_path = PurePosixPath(rels_name)
        if rels_path.parent.name != "_rels" or not rels_path.name.endswith(".rels"):
            return None
        return str(rels_path.parent.parent / rels_path.name[:-5])

    def resolve_target(source_part: str, target: str) -> str:
        if target.startswith("/"):
            return target.lstrip("/")
        return posixpath.normpath(
            posixpath.join(posixpath.dirname(source_part), target)
        ).lstrip("./")

    root_rels = xml_roots["_rels/.rels"]
    office_targets = []
    for relationship in root_rels.findall(f"{{{rel_ns}}}Relationship"):
        if relationship.get("Type") == office_document_rel:
            office_targets.append(
                resolve_target("", relationship.get("Target", ""))
            )
    if len(office_targets) != 1 or office_targets[0] not in names:
        return {
            "reason": "invalid_output",
            "detected_format": "invalid_office_document_relationship",
            "validation_error": (
                "pred.docx must contain one valid package-level "
                f"officeDocument relationship; found {office_targets}."
            ),
            "office_document_targets": office_targets,
        }

    for rels_name, root in xml_roots.items():
        if not rels_name.endswith(".rels"):
            continue
        source_part = relationship_source(rels_name)
        if source_part is None:
            continue
        relationships = root.findall(f"{{{rel_ns}}}Relationship")
        if source_part == "word/document.xml":
            singleton_counts = Counter(
                relationship.get("Type")
                for relationship in relationships
                if relationship.get("Type") in singleton_main_relationship_types
            )
            duplicate_types = sorted(
                rel_type for rel_type, count in singleton_counts.items()
                if count > 1
            )
            if duplicate_types:
                return {
                    "reason": "invalid_output",
                    "detected_format": "duplicate_singleton_relationship",
                    "validation_error": (
                        "pred.docx has multiple main-document relationships "
                        "for singleton part type(s): "
                        + ", ".join(duplicate_types)
                    ),
                    "relationship_part": rels_name,
                    "duplicate_relationship_types": duplicate_types,
                }
        for relationship in relationships:
            if (relationship.get("TargetMode") or "").casefold() == "external":
                continue
            target_raw = relationship.get("Target", "")
            if not target_raw or target_raw.startswith("#"):
                continue
            target = resolve_target(source_part, target_raw)
            if target not in names:
                return {
                    "reason": "invalid_output",
                    "detected_format": "missing_relationship_target",
                    "validation_error": (
                        f"pred.docx relationship {rels_name}:"
                        f"{relationship.get('Id', '')} targets missing part {target}."
                    ),
                    "relationship_part": rels_name,
                    "relationship_id": relationship.get("Id", ""),
                    "missing_part": target,
                }

    # A stale content-type override is package metadata, not itself a
    # visible-document defect.  Word can open and render such files
    # normally (for example after a model correctly removes a header
    # part but leaves its override).  Use it only as corroborating
    # invalid-output evidence after Word has refused both normal and
    # repair-mode opens.
    if include_word_open_failure_schema:
        # Word can hang instead of returning a useful repair error when
        # the main document contains a live footnote/endnote reference
        # but its singleton relationship is missing.  Treat this as
        # prediction-owned invalid OOXML only after normal and repair
        # opens have already failed; usable Word state remains primary.
        document_root = xml_roots["word/document.xml"]
        document_rels = xml_roots.get(
            "word/_rels/document.xml.rels"
        )
        relationship_types = {
            relationship.get("Type")
            for relationship in (
                document_rels.findall(
                    f"{{{rel_ns}}}Relationship"
                )
                if document_rels is not None
                else []
            )
        }
        for note_type in ("footnote", "endnote"):
            reference_tag = f"{{{w_ns}}}{note_type}Reference"
            reference_ids = sorted({
                reference.get(f"{{{w_ns}}}id", "")
                for reference in document_root.iter(reference_tag)
            })
            relationship_type = (
                "http://schemas.openxmlformats.org/officeDocument/2006/"
                f"relationships/{note_type}s"
            )
            if reference_ids and relationship_type not in relationship_types:
                return {
                    "reason": "invalid_output",
                    "detected_format": "missing_note_relationship",
                    "validation_error": (
                        f"pred.docx contains live w:{note_type}Reference "
                        f"markers but has no main-document {note_type}s "
                        "relationship."
                    ),
                    "invalid_part": "word/document.xml",
                    "invalid_element": f"w:{note_type}Reference",
                    "note_type": note_type,
                    "reference_ids": reference_ids,
                }

        content_types = xml_roots["[Content_Types].xml"]
        for override in content_types.findall(
            f"{{{content_types_ns}}}Override"
        ):
            part = (override.get("PartName") or "").lstrip("/")
            if part and part not in names:
                return {
                    "reason": "invalid_output",
                    "detected_format": "missing_content_type_part",
                    "validation_error": (
                        "pred.docx content-type override targets "
                        f"missing part {part}."
                    ),
                    "missing_part": part,
                }

    return None
