"""families / comments / extract_saved with explicit evidence dependencies."""

from collections import Counter, defaultdict
from evaluation.methods.word_object.families.pages import (
    extract_saved as families_pages_extract_saved,
)
from evaluation.methods.word_object.ooxml import canonical as ooxml_canonical
from evaluation.methods.word_object.preservation import context as preservation_context
from pathlib import Path
from typing import Any, Dict, List
import xml.etree.ElementTree as ET
import zipfile


def extract_saved_comment_context(docx_path: Path) -> Dict[str, Any]:
    context = preservation_context.extract_saved_insertion_context(docx_path)
    context["visible_objects"] = families_pages_extract_saved.cover_page_ooxml_layout_signature(
        docx_path
    )
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    with zipfile.ZipFile(docx_path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    context["document_background"] = ooxml_canonical.canonical_visible_element(
        root.find(f"{{{w}}}background")
    )
    return context


def extract_comment_info_ooxml(docx_path: Path) -> Dict[str, Any]:
    """Read saved comment threads, anchors, parts, and relationships."""
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    w14 = "http://schemas.microsoft.com/office/word/2010/wordml"
    w15 = "http://schemas.microsoft.com/office/word/2012/wordml"
    rel = "http://schemas.openxmlformats.org/package/2006/relationships"
    namespaces = {"w": w, "w14": w14, "w15": w15, "pr": rel}
    comment_part_names = {
        "word/comments.xml",
        "word/commentsExtended.xml",
        "word/commentsIds.xml",
        "word/commentsExtensible.xml",
    }
    with zipfile.ZipFile(docx_path) as package:
        names = set(package.namelist())
        document_root = ET.fromstring(package.read("word/document.xml"))
        comments_root = (
            ET.fromstring(package.read("word/comments.xml"))
            if "word/comments.xml" in names
            else None
        )
        extended = {}
        if "word/commentsExtended.xml" in names:
            extended_root = ET.fromstring(package.read("word/commentsExtended.xml"))
            for comment_ex in extended_root.findall("w15:commentEx", namespaces):
                para_id = comment_ex.get(f"{{{w15}}}paraId", "")
                if para_id:
                    extended[para_id] = comment_ex
        relationship_types = []
        if "word/_rels/document.xml.rels" in names:
            relationships = ET.fromstring(package.read("word/_rels/document.xml.rels"))
            relationship_types = [
                relationship.get("Type", "")
                for relationship in relationships.findall("pr:Relationship", namespaces)
                if "comment" in relationship.get("Type", "").casefold()
            ]

    marker_counts = {
        "start": Counter(),
        "end": Counter(),
        "reference": Counter(),
    }
    scope_buffers = defaultdict(list)
    active_comment_ids: List[str] = []

    def walk(node: ET.Element):
        local = node.tag.rsplit("}", 1)[-1]
        # Some normally-openable packages serialize Word attributes under
        # the element's default namespace and leave ``id`` unqualified.
        comment_id = node.get(f"{{{w}}}id") or node.get("id", "")
        if local == "commentRangeStart":
            marker_counts["start"][comment_id] += 1
            active_comment_ids.append(comment_id)
            return
        if local == "commentRangeEnd":
            marker_counts["end"][comment_id] += 1
            if comment_id in active_comment_ids:
                active_comment_ids.remove(comment_id)
            return
        if local == "commentReference":
            marker_counts["reference"][comment_id] += 1
        if local == "t" and node.text:
            for active_id in active_comment_ids:
                scope_buffers[active_id].append(node.text)
        for child in node:
            walk(child)

    body = document_root.find("w:body", namespaces)
    if body is not None:
        walk(body)

    comments = []
    errors = []
    comment_elements = (
        comments_root.findall("w:comment", namespaces) if comments_root is not None else []
    )
    definition_ids = [comment.get(f"{{{w}}}id", "") for comment in comment_elements]
    if len(definition_ids) != len(set(definition_ids)):
        errors.append("duplicate comment definition IDs")

    for index, comment in enumerate(comment_elements, start=1):
        paragraphs = comment.findall(".//w:p", namespaces)
        paragraph = paragraphs[-1] if paragraphs else None
        para_id = (
            (paragraph.get(f"{{{w14}}}paraId", "") or paragraph.get(f"{{{w15}}}paraId", ""))
            if paragraph is not None
            else ""
        )
        comment_ex = extended.get(para_id)
        parent_id = comment_ex.get(f"{{{w15}}}paraIdParent", "") if comment_ex is not None else ""
        is_reply = bool(parent_id)
        is_resolved = bool(comment_ex is not None and comment_ex.get(f"{{{w15}}}done", "0") == "1")
        comments.append(
            {
                "index": index,
                "id": comment.get(f"{{{w}}}id", ""),
                "text": "".join(
                    node.text or "" for node in comment.findall(".//w:t", namespaces)
                ).strip(),
                "author": comment.get(f"{{{w}}}author", ""),
                "para_id": para_id,
                "parent_para_id": parent_id,
                "is_reply": is_reply,
                "is_resolved": is_resolved,
                "scope_text": "".join(scope_buffers.get(comment.get(f"{{{w}}}id", ""), [])),
            }
        )

    by_para_id = {comment["para_id"]: comment for comment in comments if comment.get("para_id")}
    for comment in comments:
        if not comment["is_reply"]:
            continue
        parent = by_para_id.get(comment["parent_para_id"])
        if parent is None:
            errors.append(f"reply comment {comment['id']} has a missing parent")
            continue
        if not comment["scope_text"]:
            comment["scope_text"] = parent.get("scope_text", "")

    known_ids = set(definition_ids)
    marker_ids = set().union(*(set(counter) for counter in marker_counts.values()))
    unknown_marker_ids = sorted(marker_ids - known_ids)
    if unknown_marker_ids:
        errors.append(f"comment markers reference missing definitions: {unknown_marker_ids}")

    part_names = sorted(comment_part_names & names)
    has_comment_relationship = any(
        relationship_type.rstrip("/").endswith("/comments")
        for relationship_type in relationship_types
    )
    has_extended_relationship = any(
        relationship_type.rstrip("/").endswith("/commentsExtended")
        for relationship_type in relationship_types
    )
    linked_reply_parent_para_ids = {
        comment["parent_para_id"]
        for comment in comments
        if comment["is_reply"] and comment.get("parent_para_id")
    }
    effective_comments = []
    orphan_comment_residue = []
    for comment in comments:
        comment_id = comment["id"]
        counts = tuple(marker_counts[kind][comment_id] for kind in ("start", "end", "reference"))
        has_extended_metadata = bool(comment.get("para_id") and comment["para_id"] in extended)
        is_cleanable_top_level_orphan = bool(
            not comment["is_reply"]
            and comment_id
            and counts == (0, 0, 0)
            and comment.get("para_id") not in linked_reply_parent_para_ids
            and has_comment_relationship
            and (not has_extended_metadata or has_extended_relationship)
        )
        if is_cleanable_top_level_orphan:
            # A top-level definition with no start/end/reference anchor is
            # not reachable through Word's comment model.  Word removes
            # this exact residue on SaveAs.  Preserve it as package
            # evidence, but do not treat it as an effective comment or as
            # an invalid anchor structure.  Linked reply parents and any
            # partial marker triplet remain outside this narrow exception.
            orphan_comment_residue.append(
                {
                    "id": comment_id,
                    "text": comment["text"],
                    "author": comment["author"],
                    "para_id": comment["para_id"],
                    "is_resolved": comment["is_resolved"],
                    "marker_counts": {
                        kind: marker_counts[kind][comment_id]
                        for kind in ("start", "end", "reference")
                    },
                }
            )
            continue
        if comment["is_reply"]:
            # Word may either give a reply its own marker triplet or render
            # it through its parent thread marker. Both are valid.
            if counts not in {(0, 0, 0), (1, 1, 1)}:
                errors.append(f"reply comment {comment_id} has inconsistent markers {counts}")
        elif counts != (1, 1, 1):
            errors.append(f"top-level comment {comment_id} has inconsistent markers {counts}")
        effective_comments.append(comment)

    comments = effective_comments
    reply_count = sum(int(comment["is_reply"]) for comment in comments)
    resolved_count = sum(int(comment["is_resolved"]) for comment in comments)
    if comments:
        if "word/comments.xml" not in names or not has_comment_relationship:
            errors.append("comment definitions part/relationship is missing")
        if (reply_count or resolved_count) and (
            "word/commentsExtended.xml" not in names or not has_extended_relationship
        ):
            errors.append("threaded/resolved comment metadata part/relationship is missing")
    elif marker_ids:
        errors.append("comment markers remain without comment definitions")

    return {
        "comment_count": len(comments),
        "comment_definition_count": len(comment_elements),
        "top_level_count": len(comments) - reply_count,
        "comments": comments,
        "resolved_count": resolved_count,
        "reply_count": reply_count,
        "orphan_comment_residue_count": len(orphan_comment_residue),
        "orphan_comment_residue": orphan_comment_residue,
        "marker_counts": {
            kind: dict(sorted(counter.items())) for kind, counter in marker_counts.items()
        },
        "comment_parts": part_names,
        "comment_relationship_types": sorted(relationship_types),
        "errors": errors,
    }
