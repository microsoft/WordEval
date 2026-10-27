"""families / document_compositions / preservation with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.composition import composition_command_ids
from evaluation.methods.word_object.ooxml import canonical as ooxml_canonical
from evaluation.methods.word_object.policies import (
    content_operations as policies_content_operations,
)
from evaluation.methods.word_object.preservation import context as preservation_context
from pathlib import Path
from typing import Any, Dict, Optional
import re
import xml.etree.ElementTree as ET
import zipfile


def verify_document_enrichment_preservation(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, doc_in_path
) -> VerificationResult:
    """Verify aggregate preservation not owned by an enrichment target.

    Bibliography variants use the bibliography verifier's saved
    outside-field text/format projection.  This aggregate check still owns
    body structures that no atomic target may change.  Compare those
    structures after excluding native citation/bibliography target
    wrappers.  A gallery-only variant has no body-field operation, so it
    retains the full note-stripped body/format/object comparison.
    """
    commands = set(composition_command_ids(spec))
    if doc_in_path is None:
        return policies_content_operations.error_with_details(
            "Input document is required for enrichment preservation verification"
        )

    has_bibliography = bool({"BibliographyInsert", "BibliographyGallery"} & commands)
    if has_bibliography:
        keys = (
            "table_count",
            "content_control_count",
            "pictures",
            "visible_objects",
            "section_layout",
            "document_background",
        )
    else:
        keys = (
            "body_text",
            "body_visible_format",
            "table_count",
            "content_control_count",
            "pictures",
            "visible_objects",
            "document_background",
        )

    def projection(path: Path) -> Dict[str, Any]:
        context = preservation_context.extract_saved_semantic_context(
            path, strip_note_references=True
        )
        if has_bibliography:
            w = "http://schemas.openxmlformats.org/" "wordprocessingml/2006/main"
            with zipfile.ZipFile(path) as package:
                root = ET.fromstring(package.read("word/document.xml"))

            section_layout = []
            for section in root.findall(f".//{{{w}}}sectPr"):
                clone = ET.fromstring(ET.tostring(section, encoding="utf-8"))
                # Exclude a section property only when an atomic target in
                # this composition owns it.  Bibliography-only variants
                # must still preserve existing notes and header/footer
                # references.
                target_owned_section_names = set()
                if "FootnoteInsert" in commands:
                    target_owned_section_names.add("footnotePr")
                if "EndnoteInsertWord" in commands:
                    target_owned_section_names.add("endnotePr")
                if {
                    "HeaderInsertGallery",
                    "PageNumbersInHeaderInsertGallery",
                } & commands:
                    target_owned_section_names.add("headerReference")
                if {
                    "FooterInsertGallery",
                    "PageNambersInFooterInsertGallery",
                } & commands:
                    target_owned_section_names.add("footerReference")
                for local_name in target_owned_section_names:
                    for target in list(clone.findall(f"./{{{w}}}{local_name}")):
                        clone.remove(target)
                section_layout.append(ooxml_canonical.canonical_visible_element(clone))
            context["section_layout"] = ooxml_canonical.normalize_word_default_layout_twips(
                section_layout
            )

            target_control_count = 0
            for control in root.findall(f".//{{{w}}}body//{{{w}}}sdt"):
                properties = control.find(f"./{{{w}}}sdtPr")
                instructions = [
                    field.get(f"{{{w}}}instr", "")
                    for field in control.findall(f".//{{{w}}}fldSimple")
                ]
                instructions.extend(
                    node.text or "" for node in control.findall(f".//{{{w}}}instrText")
                )
                compact_instructions = re.sub(r"\s+", "", "".join(instructions)).upper()
                owns_citation = (
                    "CitationInsert" in commands
                    and properties is not None
                    and properties.find(f"./{{{w}}}citation") is not None
                    and "CITATION" in compact_instructions
                )
                gallery = (
                    properties.find(f"./{{{w}}}docPartObj/{{{w}}}docPartGallery")
                    if properties is not None
                    else None
                )
                gallery_name = gallery.get(f"{{{w}}}val", "") if gallery is not None else ""
                native_bibliography = (
                    properties.find(f"./{{{w}}}bibliography") if properties is not None else None
                )
                owns_bibliography = (
                    bool({"BibliographyInsert", "BibliographyGallery"} & commands)
                    and (
                        bool(re.search(r"bibliograph", gallery_name, re.IGNORECASE))
                        or native_bibliography is not None
                    )
                    and "BIBLIOGRAPHY" in compact_instructions
                )
                if owns_citation or owns_bibliography:
                    target_control_count += 1

            context["content_control_count"] = max(
                0,
                int(context.get("content_control_count") or 0) - target_control_count,
            )
            visible_objects = dict(context.get("visible_objects") or {})
            # Bibliography text/paragraph roles remain target-owned and are
            # already checked by _verify_bibliography's outside projection.
            visible_objects.pop("paragraph_roles", None)
            visible_objects.pop("paragraph_formatting", None)
            context["visible_objects"] = visible_objects
        return {key: context.get(key) for key in keys}

    try:
        input_projection = projection(doc_in_path)
        gold_projection = projection(gold_path)
        pred_projection = projection(pred_path)
    except Exception as exc:
        return policies_content_operations.error_with_details(
            "Could not extract aggregate enrichment preservation evidence",
            {"exception": str(exc)},
        )
    if input_projection != gold_projection:
        return policies_content_operations.error_with_details(
            "Gold changes state outside the requested enrichment targets",
            {
                "match_mode": "saved_ooxml_enrichment_preservation",
                "input": input_projection,
                "gold": gold_projection,
            },
        )
    if pred_projection != gold_projection:
        return VerificationResult(
            passed=False,
            message="FAIL: Enrichment composition changed unrelated body/layout",
            details={
                "match_mode": "saved_ooxml_enrichment_preservation",
                "gold": gold_projection,
                "pred": pred_projection,
            },
        )
    return VerificationResult(
        passed=True,
        message="PASS: Enrichment body/layout preservation matches",
        details={
            "match_mode": "saved_ooxml_enrichment_preservation",
            "checked_properties": list(keys),
        },
    )
