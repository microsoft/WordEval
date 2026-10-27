"""families / references / bibliography with explicit evidence dependencies."""

from evaluation.methods.base import VerificationResult
from evaluation.methods.word_object.core import DocumentOpenError
from evaluation.methods.word_object.ooxml import canonical as ooxml_canonical
from pathlib import Path
from typing import Any, Dict, List, Optional
import re
from .fields import require_reference_sections
import xml.etree.ElementTree as ET
import zipfile


def verify_bibliography(
    pred_path: Path, gold_path: Path, spec: Optional[Dict] = None, *, read_bibliography_info
) -> VerificationResult:
    """Verify bibliography content (with style-tolerant mode for insert/gallery commands)."""
    idMso = spec.get("idMso", "") if spec else ""
    parameter = spec.get("parameter", {}) if spec else {}
    configured_gallery_heading = (
        str(parameter.get("name", "")).strip() if isinstance(parameter, dict) else ""
    )

    def saved_bibliography_state(path: Path) -> Dict[str, Any]:
        w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        with zipfile.ZipFile(path) as package:
            document_xml = package.read("word/document.xml")

        def collect_fields(
            root: ET.Element,
        ) -> tuple[List[Dict[str, str]], set[int]]:
            fields: List[Dict[str, str]] = []
            excluded_text_ids: set[int] = set()
            stack: List[Dict[str, Any]] = []
            for node in root.iter():
                if node.tag == f"{{{w}}}fldSimple":
                    instruction = node.get(f"{{{w}}}instr", "")
                    if re.search(r"\bBIBLIOGRAPHY\b", instruction, re.IGNORECASE):
                        text_nodes = list(node.iter(f"{{{w}}}t"))
                        excluded_text_ids.update(id(text) for text in text_nodes)
                        result = " ".join(
                            (text.text or "").strip()
                            for text in text_nodes
                            if (text.text or "").strip()
                        )
                        fields.append(
                            {
                                "instruction": " ".join(instruction.split()),
                                "result": " ".join(result.split()),
                            }
                        )
                    continue

                if node.tag == f"{{{w}}}fldChar":
                    kind = node.get(f"{{{w}}}fldCharType", "")
                    if kind == "begin":
                        stack.append(
                            {
                                "instruction_parts": [],
                                "result_parts": [],
                                "text_nodes": [],
                                "in_result": False,
                            }
                        )
                    elif kind == "separate" and stack:
                        stack[-1]["in_result"] = True
                    elif kind == "end" and stack:
                        field = stack.pop()
                        instruction = " ".join(field["instruction_parts"]).strip()
                        if re.search(r"\bBIBLIOGRAPHY\b", instruction, re.IGNORECASE):
                            excluded_text_ids.update(id(text) for text in field["text_nodes"])
                            fields.append(
                                {
                                    "instruction": " ".join(instruction.split()),
                                    "result": " ".join(" ".join(field["result_parts"]).split()),
                                }
                            )
                    continue

                if not stack:
                    continue
                if node.tag == f"{{{w}}}instrText":
                    stack[-1]["instruction_parts"].append(node.text or "")
                elif node.tag == f"{{{w}}}t":
                    for field in stack:
                        field["text_nodes"].append(node)
                        if field["in_result"]:
                            field["result_parts"].append(node.text or "")
            return fields, excluded_text_ids

        def bibliography_region(
            root: ET.Element,
        ) -> tuple[
            List[Dict[str, str]],
            set[int],
            set[int],
            List[str],
        ]:
            """Locate field-result text and an adjacent gallery heading.

            Word may save a bibliography heading inside the outer gallery
            content control while the BIBLIOGRAPHY field itself contains
            only the source entries.  A direct field insertion may omit
            the heading entirely.  Treat an exact, recognized heading as
            part of the bibliography region only when it is the field
            paragraph itself or immediately adjacent in the visible body
            flow.  This preserves unrelated headings elsewhere.
            """
            fields, field_text_ids = collect_fields(root)
            visible_paragraphs = []
            field_paragraph_ids = set()
            for paragraph in root.findall(f".//{{{w}}}body//{{{w}}}p"):
                if paragraph.findall(f".//{{{w}}}p"):
                    continue
                text_nodes = list(paragraph.findall(f".//{{{w}}}t"))
                text = " ".join("".join(node.text or "" for node in text_nodes).split())
                if not text:
                    continue
                if any(id(node) in field_text_ids for node in text_nodes):
                    field_paragraph_ids.add(id(paragraph))
                visible_paragraphs.append((paragraph, text, text_nodes))

            heading_text_ids: set[int] = set()
            headings: List[str] = []
            for index, (paragraph, text, text_nodes) in enumerate(visible_paragraphs):
                if not re.fullmatch(
                    r"(bibliography|references|works\s+cited)",
                    text,
                    flags=re.IGNORECASE,
                ):
                    continue
                adjacent_to_field = id(paragraph) in field_paragraph_ids
                for neighbor_index in (index - 1, index + 1):
                    if 0 <= neighbor_index < len(visible_paragraphs):
                        neighbor = visible_paragraphs[neighbor_index][0]
                        adjacent_to_field = adjacent_to_field or id(neighbor) in field_paragraph_ids
                if not adjacent_to_field:
                    continue
                headings.append(" ".join(text.split()))
                heading_text_ids.update(id(node) for node in text_nodes)
            return fields, field_text_ids, heading_text_ids, headings

        def paragraph_roles(root: ET.Element) -> List[str]:
            roles = []
            for paragraph in root.findall(f".//{{{w}}}body//{{{w}}}p"):
                if paragraph.findall(f".//{{{w}}}p"):
                    continue
                text = "".join(node.text or "" for node in paragraph.findall(f".//{{{w}}}t"))
                if not text.strip():
                    continue
                style = paragraph.find(f"./{{{w}}}pPr/{{{w}}}pStyle")
                roles.append(style.get(f"{{{w}}}val", "") if style is not None else "")
            return roles

        outside_root = ET.fromstring(document_xml)
        (
            fields,
            outside_field_text_ids,
            outside_heading_text_ids,
            headings,
        ) = bibliography_region(outside_root)
        outside_excluded_ids = outside_field_text_ids | outside_heading_text_ids
        for text in outside_root.iter(f"{{{w}}}t"):
            if id(text) in outside_excluded_ids:
                text.text = ""
        outside_text = " ".join(
            "".join(
                node.text or "" for node in outside_root.findall(f".//{{{w}}}body//{{{w}}}t")
            ).split()
        )
        outside_signature = {
            "text": outside_text,
            "paragraph_roles": paragraph_roles(outside_root),
            "visible_formatting": (
                ooxml_canonical.saved_body_visible_format_signature(outside_root)
            ),
        }

        topology_root = ET.fromstring(document_xml)
        (
            _,
            topology_field_text_ids,
            topology_heading_text_ids,
            _,
        ) = bibliography_region(topology_root)
        topology_region_ids = topology_field_text_ids | topology_heading_text_ids
        body_topology = []
        for paragraph in topology_root.findall(f".//{{{w}}}body//{{{w}}}p"):
            if paragraph.findall(f".//{{{w}}}p"):
                continue
            text_nodes = list(paragraph.findall(f".//{{{w}}}t"))
            if any(id(node) in topology_region_ids for node in text_nodes):
                body_topology.append("bibliography")
            elif "".join(node.text or "" for node in text_nodes).strip():
                body_topology.append("outside")
            else:
                body_topology.append("empty")

        field_root = ET.fromstring(document_xml)
        _, field_text_ids, _, _ = bibliography_region(field_root)
        for text in field_root.iter(f"{{{w}}}t"):
            if id(text) not in field_text_ids:
                text.text = ""

        # Word sometimes persists explicit values that are equivalent to
        # the rendering defaults.  They must not make two visually
        # identical bibliography entries compare differently.
        for parent in field_root.iter():
            for child in list(parent):
                local_name = child.tag.rsplit("}", 1)[-1]
                value = next(
                    (
                        attribute_value
                        for attribute_name, attribute_value in child.attrib.items()
                        if attribute_name.rsplit("}", 1)[-1] == "val"
                    ),
                    "",
                ).lower()
                if local_name == "rFonts":
                    for attribute_name in list(child.attrib):
                        if (
                            attribute_name.rsplit("}", 1)[-1] == "hint"
                            and child.attrib[attribute_name].lower() == "eastasia"
                        ):
                            del child.attrib[attribute_name]
                    if not child.attrib and not list(child) and not (child.text or "").strip():
                        parent.remove(child)
                        continue
                if (local_name == "kern" and value == "0") or (
                    local_name == "ligatures" and value == "none"
                ):
                    parent.remove(child)
        raw_field_formatting = ooxml_canonical.saved_body_visible_format_signature(field_root)

        # BibliographyInsert intentionally tolerates the citation style
        # selected by Word/the active template.  Still retain conspicuous
        # direct visual effects so additions such as red text, highlight,
        # bold, strikeout, or superscript cannot hide behind that
        # tolerance.
        visible_effect_tags = {
            "b",
            "bCs",
            "caps",
            "color",
            "dstrike",
            "emboss",
            "highlight",
            "imprint",
            "outline",
            "shadow",
            "shd",
            "smallCaps",
            "strike",
            "textEffect",
            "textFill",
            "textOutline",
            "u",
            "vertAlign",
        }
        field_visible_effects = set()
        for paragraph in field_root.findall(f".//{{{w}}}body//{{{w}}}p"):
            if paragraph.findall(f".//{{{w}}}p"):
                continue
            if not "".join(node.text or "" for node in paragraph.findall(f".//{{{w}}}t")).strip():
                continue
            property_roots = []
            paragraph_properties = paragraph.find(f"./{{{w}}}pPr")
            if paragraph_properties is not None:
                property_roots.append(paragraph_properties)
            for run in paragraph.findall(f".//{{{w}}}r"):
                if not "".join(node.text or "" for node in run.findall(f".//{{{w}}}t")):
                    continue
                run_properties = run.find(f"./{{{w}}}rPr")
                if run_properties is not None:
                    property_roots.append(run_properties)
            for properties in property_roots:
                for element in properties.iter():
                    local_name = element.tag.rsplit("}", 1)[-1]
                    if local_name not in visible_effect_tags:
                        continue
                    attributes = tuple(
                        sorted(
                            (
                                name.rsplit("}", 1)[-1],
                                value,
                            )
                            for name, value in element.attrib.items()
                        )
                    )
                    field_visible_effects.add((local_name, attributes))

        heading_root = ET.fromstring(document_xml)
        _, _, heading_text_ids, _ = bibliography_region(heading_root)
        for text in heading_root.iter(f"{{{w}}}t"):
            if id(text) not in heading_text_ids:
                text.text = ""
        raw_heading_formatting = ooxml_canonical.saved_body_visible_format_signature(heading_root)

        placement_root = ET.fromstring(document_xml)
        (
            _,
            placement_field_text_ids,
            placement_heading_text_ids,
            _,
        ) = bibliography_region(placement_root)
        placement_region_ids = placement_field_text_ids | placement_heading_text_ids
        body_text_nodes = list(placement_root.findall(f".//{{{w}}}body//{{{w}}}t"))
        region_indexes = [
            index for index, node in enumerate(body_text_nodes) if id(node) in placement_region_ids
        ]
        if region_indexes:
            first_region_index = min(region_indexes)
            last_region_index = max(region_indexes)
            before_text = " ".join(
                "".join(
                    node.text or ""
                    for node in body_text_nodes[:first_region_index]
                    if id(node) not in placement_region_ids
                ).split()
            )
            after_text = " ".join(
                "".join(
                    node.text or ""
                    for node in body_text_nodes[last_region_index + 1 :]
                    if id(node) not in placement_region_ids
                ).split()
            )
        else:
            before_text = ""
            after_text = ""

        def unique(values: List[Any]) -> List[Any]:
            result = []
            for value in values:
                if value not in result:
                    result.append(value)
            return sorted(result, key=repr)

        field_formatting = {
            "paragraph_roles": unique(paragraph_roles(field_root)),
            "paragraph_formats": unique(
                [paragraph.get("paragraph_format") for paragraph in raw_field_formatting]
            ),
            "run_formats": unique(
                [
                    run.get("format")
                    for paragraph in raw_field_formatting
                    for run in paragraph.get("runs", [])
                ]
            ),
        }
        heading_formatting = {
            "paragraph_roles": unique(paragraph_roles(heading_root)),
            "paragraph_formats": unique(
                [paragraph.get("paragraph_format") for paragraph in raw_heading_formatting]
            ),
            "run_formats": unique(
                [
                    run.get("format")
                    for paragraph in raw_heading_formatting
                    for run in paragraph.get("runs", [])
                ]
            ),
        }
        return {
            "fields": fields,
            "outside": outside_signature,
            "field_formatting": field_formatting,
            "field_visible_effects": sorted(field_visible_effects, key=repr),
            "headings": headings,
            "heading_formatting": heading_formatting,
            "body_topology": body_topology,
            "placement": {
                "before_text": before_text,
                "after_text": after_text,
            },
        }

    # Read semantic bibliography state through Word first. Saved field
    # codes/results below are persistence guardrails (for plain-text
    # lookalikes and stale caches), not the primary evaluator.
    com_error = None
    try:
        gold_bib = read_bibliography_info(str(gold_path))
        pred_bib = read_bibliography_info(str(pred_path))
        required = {"refresh", "document_text", "source_count", "sources", "bibliography_fields"}
        if idMso == "BibliographyStyle":
            required.add("bibliography_style")
        error = require_reference_sections({"gold": gold_bib, "pred": pred_bib}, required)
        if error is not None:
            return error
    except DocumentOpenError:
        raise
    except Exception as exc:
        com_error = str(exc)
        gold_bib = {}
        pred_bib = {}

    gold_saved = saved_bibliography_state(gold_path)
    pred_saved = saved_bibliography_state(pred_path)
    gold_saved_fields = gold_saved["fields"]
    pred_saved_fields = pred_saved["fields"]
    if not gold_saved_fields or not all(
        field.get("result", "").strip() for field in gold_saved_fields
    ):
        return VerificationResult.error_result(
            "Gold bibliography field/cache evidence is unavailable"
        )
    if len(pred_saved_fields) < len(gold_saved_fields):
        return VerificationResult(
            passed=False,
            message="FAIL: No saved BIBLIOGRAPHY field found (plain text is not accepted)",
            details={
                "match_mode": "saved_ooxml_bibliography_field",
                "gold_field_count": len(gold_saved_fields),
                "pred_field_count": len(pred_saved_fields),
                "verification_source": "word_com+saved_ooxml_guardrail",
                "word_com_primary": {
                    "gold": gold_bib,
                    "pred": pred_bib,
                },
            },
        )
    if len(pred_saved_fields) != len(gold_saved_fields):
        return VerificationResult(
            passed=False,
            message="FAIL: Saved BIBLIOGRAPHY field count differs",
            details={
                "match_mode": "saved_ooxml_bibliography_field",
                "gold_field_count": len(gold_saved_fields),
                "pred_field_count": len(pred_saved_fields),
                "verification_source": "word_com+saved_ooxml_guardrail",
            },
        )
    if not all(field.get("result", "").strip() for field in pred_saved_fields):
        return VerificationResult(
            passed=False,
            message="FAIL: Saved bibliography field result is empty",
            details={
                "match_mode": "saved_ooxml_bibliography_cache",
                "pred_fields": pred_saved_fields,
                "verification_source": "word_com+saved_ooxml_guardrail",
            },
        )
    gold_headings = gold_saved["headings"]
    pred_headings = pred_saved["headings"]
    if idMso == "BibliographyGallery":
        expected_heading = configured_gallery_heading or (
            gold_headings[0] if len(gold_headings) == 1 else ""
        )
        if not expected_heading:
            return VerificationResult.error_result(
                "Gold bibliography gallery heading evidence is unavailable"
            )
        if len(pred_headings) != 1 or pred_headings[0].casefold() != expected_heading.casefold():
            return VerificationResult(
                passed=False,
                message="FAIL: Bibliography gallery heading differs",
                details={
                    "match_mode": "saved_ooxml_bibliography_heading",
                    "expected_heading": expected_heading,
                    "gold_headings": gold_headings,
                    "pred_headings": pred_headings,
                    "verification_source": ("word_com+saved_ooxml_guardrail"),
                },
            )
    elif idMso == "BibliographyStyle":
        if pred_headings != gold_headings:
            return VerificationResult(
                passed=False,
                message="FAIL: Bibliography heading changed",
                details={
                    "match_mode": "saved_ooxml_bibliography_heading",
                    "gold_headings": gold_headings,
                    "pred_headings": pred_headings,
                    "verification_source": ("word_com+saved_ooxml_guardrail"),
                },
            )
    elif len(pred_headings) > 1:
        return VerificationResult(
            passed=False,
            message="FAIL: Multiple bibliography headings found",
            details={
                "match_mode": "saved_ooxml_bibliography_heading",
                "pred_headings": pred_headings,
                "verification_source": "word_com+saved_ooxml_guardrail",
            },
        )
    if pred_saved["placement"] != gold_saved["placement"]:
        return VerificationResult(
            passed=False,
            message="FAIL: Bibliography placement differs",
            details={
                "match_mode": "saved_ooxml_bibliography_placement",
                "gold_placement": gold_saved["placement"],
                "pred_placement": pred_saved["placement"],
                "verification_source": "word_com+saved_ooxml_guardrail",
            },
        )

    def stable_gallery_topology(topology: List[str]) -> List[str]:
        """Ignore bounded empty gallery-host shells.

        Word/gallery hosts may add or remove trailing empty paragraphs and
        one empty paragraph immediately before a leading bibliography
        region. Multiple leading blanks, or blanks separating preserved
        outside content from the bibliography, remain semantic evidence.
        """
        stable = list(topology)
        while stable and stable[-1] == "empty":
            stable.pop()
        if len(stable) >= 2 and stable[0] == "empty" and stable[1] == "bibliography":
            stable.pop(0)
        return stable

    gold_stable_topology = stable_gallery_topology(gold_saved["body_topology"])
    pred_stable_topology = stable_gallery_topology(pred_saved["body_topology"])
    if idMso == "BibliographyGallery" and pred_stable_topology != gold_stable_topology:
        return VerificationResult(
            passed=False,
            message="FAIL: Bibliography gallery changed body paragraph topology",
            details={
                "match_mode": "saved_ooxml_bibliography_topology",
                "gold_topology": gold_saved["body_topology"],
                "pred_topology": pred_saved["body_topology"],
                "gold_stable_topology": gold_stable_topology,
                "pred_stable_topology": pred_stable_topology,
                "verification_source": "word_com+saved_ooxml_guardrail",
            },
        )
    if pred_saved["outside"] != gold_saved["outside"]:
        return VerificationResult(
            passed=False,
            message="FAIL: Visible content/formatting outside bibliography differs",
            details={
                "match_mode": "saved_ooxml_bibliography_preservation",
                "gold_outside": gold_saved["outside"],
                "pred_outside": pred_saved["outside"],
                "verification_source": "word_com+saved_ooxml_guardrail",
            },
        )
    if (
        idMso == "BibliographyStyle"
        and pred_saved["heading_formatting"] != gold_saved["heading_formatting"]
    ):
        return VerificationResult(
            passed=False,
            message="FAIL: Visible bibliography heading formatting differs",
            details={
                "match_mode": "saved_ooxml_bibliography_heading_formatting",
                "gold_formatting": gold_saved["heading_formatting"],
                "pred_formatting": pred_saved["heading_formatting"],
                "verification_source": "word_com+saved_ooxml_guardrail",
            },
        )
    if idMso in {
        "BibliographyInsert",
        "BibliographyGallery",
        "BibliographyStyle",
    }:
        unexpected_effects = sorted(
            set(pred_saved["field_visible_effects"]) - set(gold_saved["field_visible_effects"]),
            key=repr,
        )
        if unexpected_effects:
            return VerificationResult(
                passed=False,
                message="FAIL: Unexpected visible bibliography formatting",
                details={
                    "match_mode": ("saved_ooxml_bibliography_visible_effects"),
                    "gold_visible_effects": (gold_saved["field_visible_effects"]),
                    "pred_visible_effects": (pred_saved["field_visible_effects"]),
                    "unexpected_effects": unexpected_effects,
                    "verification_source": ("word_com+saved_ooxml_guardrail"),
                },
            )
    elif pred_saved["field_formatting"] != gold_saved["field_formatting"]:
        return VerificationResult(
            passed=False,
            message="FAIL: Visible bibliography formatting differs",
            details={
                "match_mode": "saved_ooxml_bibliography_formatting",
                "gold_formatting": gold_saved["field_formatting"],
                "pred_formatting": pred_saved["field_formatting"],
                "verification_source": "word_com+saved_ooxml_guardrail",
            },
        )

    gold_cache = [field["result"] for field in gold_saved_fields]
    pred_cache = [field["result"] for field in pred_saved_fields]

    def normalize_saved_cache(text: str) -> str:
        """Normalize Word's non-semantic punctuation spacing."""
        normalized = " ".join((text or "").replace("\u00a0", " ").split())
        return re.sub(r"\s+([,.;:!?])", r"\1", normalized)

    normalized_gold_cache = [normalize_saved_cache(result) for result in gold_cache]
    normalized_pred_cache = [normalize_saved_cache(result) for result in pred_cache]

    # Source identity comes from the current gold document, never from a
    # corpus-specific literal.  The saved field result remains the
    # persistence guardrail because it is what a user actually sees.
    gold_sources = gold_bib.get("sources", []) or []
    gold_source_titles = [normalize_saved_cache(source.get("title", "")) for source in gold_sources]
    gold_source_count = gold_bib.get("source_count", 0)
    if com_error is None and (
        not isinstance(gold_source_count, int)
        or gold_source_count <= 0
        or len(gold_sources) != gold_source_count
        or any(not title for title in gold_source_titles)
    ):
        result = VerificationResult.error_result(
            "Gold bibliography source-title evidence is unavailable"
        )
        result.details = {
            "match_mode": "word_com_bibliography_source_identity",
            "error_owner": "gold_fixture",
            "gold_source_count": gold_source_count,
            "observed_gold_source_count": len(gold_sources),
            "gold_sources": gold_sources,
        }
        return result
    required_source_titles = sorted(set(gold_source_titles))

    def significant_tokens(text: str) -> set[str]:
        heading_tokens = {
            "bibliography",
            "references",
            "works",
            "cited",
        }
        return {
            token
            for token in re.findall(r"[A-Za-z0-9]+", text.lower())
            if len(token) >= 3 and token not in heading_tokens
        }

    gold_cache_text = " ".join(normalized_gold_cache)
    pred_cache_text = " ".join(normalized_pred_cache)
    gold_cache_tokens = significant_tokens(gold_cache_text)
    pred_cache_tokens = significant_tokens(pred_cache_text)
    overlap = gold_cache_tokens & pred_cache_tokens
    overlap_ratio = len(overlap) / len(gold_cache_tokens) if gold_cache_tokens else 0.0
    missing_source_titles = [
        title
        for title in required_source_titles
        if title.casefold() not in pred_cache_text.casefold()
    ]
    if missing_source_titles or overlap_ratio < 0.6:
        return VerificationResult(
            passed=False,
            message="FAIL: Saved bibliography field result content differs",
            details={
                "match_mode": "saved_ooxml_bibliography_cache",
                "gold_cache": gold_cache,
                "pred_cache": pred_cache,
                "overlap_ratio": round(overlap_ratio, 3),
                "required_source_titles": required_source_titles,
                "missing_source_titles": missing_source_titles,
                "verification_source": ("word_com+saved_ooxml_guardrail"),
            },
        )

    if com_error is not None:
        result = VerificationResult.error_result(
            "Word COM bibliography source identity is unavailable; "
            "matching saved field caches are not a complete fallback"
        )
        result.details = {
            "match_mode": "saved_ooxml_bibliography_incomplete_fallback",
            "fallback_contract": "incomplete",
            "gold_saved": gold_saved,
            "pred_saved": pred_saved,
            "word_com_error": com_error,
        }
        return result
    gold_bib_fields = gold_bib.get("bibliography_fields") or []
    pred_bib_fields = pred_bib.get("bibliography_fields") or []

    if gold_bib_fields and not pred_bib_fields:
        return VerificationResult(
            passed=False,
            message="FAIL: No BIBLIOGRAPHY field found in predicted document (plain text is not accepted)",
            details={"gold_bibliography": gold_bib, "pred_bibliography": pred_bib},
        )

    if gold_bib.get("source_count", 0) and pred_bib.get("source_count", 0) < gold_bib.get(
        "source_count", 0
    ):
        return VerificationResult(
            passed=False,
            message="FAIL: Bibliography source count is lower than gold",
            details={"gold_bibliography": gold_bib, "pred_bibliography": pred_bib},
        )

    gold_text = (gold_bib.get("document_text") or "").strip()
    pred_text = (pred_bib.get("document_text") or "").strip()

    def _normalized_bibliography_lines(text: str) -> List[str]:
        normalized = (text or "").replace("\r\n", "\n").replace("\r", "\n")
        return [line.strip() for line in normalized.split("\n") if line.strip()]

    def _split_bibliography_heading(text: str) -> tuple[str, str]:
        lines = _normalized_bibliography_lines(text)
        heading = ""
        if lines and re.fullmatch(
            r"(bibliography|references|works\s+cited)", lines[0], flags=re.IGNORECASE
        ):
            heading = lines[0]
            lines = lines[1:]
        return heading, "\n".join(lines).strip()

    if idMso == "BibliographyStyle":
        gold_style = (gold_bib.get("bibliography_style") or "").strip()
        pred_style = (pred_bib.get("bibliography_style") or "").strip()
        if gold_style and pred_style != gold_style:
            return VerificationResult(
                passed=False,
                message=f"FAIL: Bibliography style mismatch (pred={pred_style}, gold={gold_style})",
                details={"gold_bibliography": gold_bib, "pred_bibliography": pred_bib},
            )

    # Exact match remains preferred after the explicit style and saved-state
    # contracts have been checked.
    if pred_text == gold_text:
        return VerificationResult(
            passed=True,
            message="PASS: Bibliography content matches",
            details={
                "text_length": len(gold_text),
                "match_mode": "exact_text",
                "verification_source": "word_com+saved_ooxml_guardrail",
            },
        )

    if idMso == "BibliographyStyle":

        gold_heading, gold_body = _split_bibliography_heading(gold_text)
        pred_heading, pred_body = _split_bibliography_heading(pred_text)

        if pred_body == gold_body:
            return VerificationResult(
                passed=True,
                message="PASS: Bibliography content matches (heading-tolerant)",
                details={
                    "match_mode": "heading_tolerant_exact_body",
                    "gold_heading": gold_heading,
                    "pred_heading": pred_heading,
                    "body_length": len(gold_body),
                },
            )

        return VerificationResult(
            passed=False,
            message="FAIL: Bibliography body content differs",
            details={
                "match_mode": "heading_tolerant_exact_body",
                "gold_heading": gold_heading,
                "pred_heading": pred_heading,
                "gold_body": gold_body[:200] + "..." if len(gold_body) > 200 else gold_body,
                "pred_body": pred_body[:200] + "..." if len(pred_body) > 200 else pred_body,
            },
        )

    if idMso in {"BibliographyInsert", "BibliographyGallery"}:
        # Word may emit different bibliography/citation styles (APA/MLA/etc.)
        # for the same command depending on environment/template behavior.
        # For word_object, accept body style differences if the source
        # content appears semantically consistent with gold. Do not require
        # the first non-empty document line to match: composition cases have
        # body text before the bibliography region.
        def _first_nonempty_line(text: str) -> str:
            for line in text.splitlines():
                s = line.strip()
                if s:
                    return s
            return ""

        gold_heading = _first_nonempty_line(gold_text)
        pred_heading = _first_nonempty_line(pred_text)

        def _tokens(text: str) -> List[str]:
            return re.findall(r"[A-Za-z0-9]+", text.lower())

        gold_tokens = _tokens(gold_text)
        pred_tokens = _tokens(pred_text)
        gold_set = set(gold_tokens)
        pred_set = set(pred_tokens)

        # Ignore heading-only tokens that vary by style/gallery label.
        heading_tokens = {"bibliography", "references", "works", "cited"}
        gold_sig = {t for t in gold_set if len(t) >= 3 and t not in heading_tokens}
        pred_sig = {t for t in pred_set if len(t) >= 3 and t not in heading_tokens}

        if gold_sig:
            overlap = gold_sig & pred_sig
            overlap_ratio = len(overlap) / len(gold_sig)
        else:
            overlap = set()
            overlap_ratio = 0.0

        # Require non-empty pred text and substantial overlap with gold source tokens.
        if pred_text and gold_text and overlap_ratio >= 0.6:
            return VerificationResult(
                passed=True,
                message="PASS: Bibliography content matches (style-tolerant)",
                details={
                    "match_mode": "style_tolerant_token_overlap",
                    "overlap_ratio": round(overlap_ratio, 3),
                    "gold_significant_token_count": len(gold_sig),
                    "pred_significant_token_count": len(pred_sig),
                    "overlap_tokens_preview": sorted(list(overlap))[:20],
                    "gold_bibliography": gold_bib,
                    "pred_bibliography": pred_bib,
                },
            )

    if pred_text != gold_text:
        return VerificationResult(
            passed=False,
            message="FAIL: Document text content differs (bibliography mismatch)",
            details={
                "pred_length": len(pred_text),
                "gold_length": len(gold_text),
                "pred_text": pred_text[:200] + "..." if len(pred_text) > 200 else pred_text,
                "gold_text": gold_text[:200] + "..." if len(gold_text) > 200 else gold_text,
                "match_mode": "exact_text_required",
            },
        )
