"""preservation / visual with explicit evidence dependencies."""

from evaluation.methods.word_object.effect_commands import (
    MARKUP_COMPATIBILITY_NS,
    OFFICE_MATH_NS,
    PICTURE_NS,
    WORDPROCESSINGML_NS,
    WORDPROCESSING_GROUP_NS,
    WORDPROCESSING_SHAPE_NS,
)
from evaluation.methods.base import VerificationResult
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple
import io
import re
import xml.etree.ElementTree as ET
import zipfile


def saved_visual_preservation_signature(
    docx_path: Path, excluded_target_kind: Optional[str] = None
) -> Dict[str, Any]:
    """Project logical visible objects, paragraph topology, and layout.

    The projection deliberately separates a floating object's logical
    identity from the paragraph that happens to carry its anchor.  Word
    can move an equivalent floating TextBox anchor between adjacent
    paragraphs while preserving the same visible body topology.  Markup
    Compatibility ``AlternateContent`` contributes exactly one supported
    Choice (or its Fallback), so a DrawingML/VML fallback pair is one
    logical object rather than two visible objects.

    Run segmentation, relationship IDs, package order, and target drawing
    geometry remain route-owned.  Empty main-story paragraphs, tabs,
    breaks, and pagination-relevant paragraph properties are retained so
    representation normalization cannot hide collateral topology/layout
    changes.
    """
    w = f"{{{WORDPROCESSINGML_NS}}}"
    m = f"{{{OFFICE_MATH_NS}}}"
    mc = f"{{{MARKUP_COMPATIBILITY_NS}}}"
    wps = f"{{{WORDPROCESSING_SHAPE_NS}}}"
    pic = f"{{{PICTURE_NS}}}"
    v = "{urn:schemas-microsoft-com:vml}"
    dgm = "{http://schemas.openxmlformats.org/drawingml/2006/diagram}"
    chart = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"

    def namespace_map(xml_bytes: bytes) -> Dict[str, str]:
        return {
            prefix or "": uri
            for _, (prefix, uri) in ET.iterparse(io.BytesIO(xml_bytes), events=("start-ns",))
        }

    def supported_choice_namespace(uri: Optional[str]) -> bool:
        if not uri:
            return False
        return (
            uri
            in {
                WORDPROCESSING_SHAPE_NS,
                WORDPROCESSING_GROUP_NS,
            }
            or uri.startswith("http://schemas.microsoft.com/office/word/20")
            or uri.startswith("http://schemas.microsoft.com/office/drawing/20")
        )

    def selected_children(
        node: ET.Element,
        prefixes: Dict[str, str],
    ) -> List[ET.Element]:
        if node.tag != f"{mc}AlternateContent":
            return list(node)
        fallback: Optional[ET.Element] = None
        for branch in node:
            if branch.tag == f"{mc}Fallback":
                fallback = branch
                continue
            if branch.tag != f"{mc}Choice":
                continue
            requires = branch.get("Requires", "").split()
            if requires and all(
                supported_choice_namespace(prefixes.get(prefix)) for prefix in requires
            ):
                return list(branch)
        return list(fallback) if fallback is not None else []

    def selected_iter(
        node: ET.Element,
        prefixes: Dict[str, str],
    ):
        yield node
        for child in selected_children(node, prefixes):
            yield from selected_iter(child, prefixes)

    def paragraph_text(
        paragraph: ET.Element,
        prefixes: Dict[str, str],
    ) -> Tuple[str, bool]:
        tokens: List[str] = []
        has_non_text_control = False

        def walk(node: ET.Element, root: bool = False) -> None:
            nonlocal has_non_text_control
            for child in selected_children(node, prefixes):
                if not root and child.tag == f"{w}p":
                    continue
                if child.tag == f"{w}t":
                    tokens.append(child.text or "")
                elif child.tag == f"{w}tab":
                    tokens.append("\t")
                    has_non_text_control = True
                elif child.tag == f"{w}br":
                    tokens.append(f"\n<{child.get(f'{w}type', 'textWrapping')}>")
                    has_non_text_control = True
                else:
                    walk(child)

        walk(paragraph, root=True)
        return "".join(tokens), has_non_text_control

    def canonical_attributes(node: Optional[ET.Element]) -> Any:
        if node is None:
            return None
        if node.tag == f"{w}docGrid" and node.get(f"{w}type", "default") == "default":
            return None
        return {key.rsplit("}", 1)[-1]: value for key, value in sorted(node.attrib.items())}

    def paragraph_layout(paragraph: ET.Element) -> List[Dict[str, Any]]:
        ppr = paragraph.find(f"{w}pPr")
        if ppr is None:
            return []
        layout_tags = {
            "pStyle",
            "keepNext",
            "keepLines",
            "pageBreakBefore",
            "widowControl",
            "spacing",
            "framePr",
            "textDirection",
            "contextualSpacing",
            "suppressLineNumbers",
        }
        return [
            {
                "name": child.tag.rsplit("}", 1)[-1],
                "attributes": canonical_attributes(child),
            }
            for child in ppr
            if child.tag.rsplit("}", 1)[-1] in layout_tags
        ]

    def logical_object_kind(
        node: ET.Element,
        prefixes: Dict[str, str],
    ) -> Optional[str]:
        if node.tag == f"{wps}wsp":
            tags = {child.tag for child in selected_iter(node, prefixes)}
            if f"{wps}txbx" in tags:
                return "textbox"
            return "shape"
        if node.tag == f"{v}shape":
            tags = {child.tag for child in selected_iter(node, prefixes)}
            if f"{v}textbox" in tags:
                return "textbox"
            return "shape"
        if node.tag == f"{pic}pic":
            return "picture"
        if node.tag == f"{dgm}relIds":
            return "smartart"
        if node.tag == f"{chart}chart":
            return "chart"
        return None

    def logical_objects(
        root: ET.Element,
        prefixes: Dict[str, str],
    ) -> List[Dict[str, Any]]:
        objects: List[Dict[str, Any]] = []
        for node in selected_iter(root, prefixes):
            kind = logical_object_kind(node, prefixes)
            if kind is None:
                continue
            paragraphs: List[str] = []
            if kind == "textbox":
                for descendant in selected_iter(node, prefixes):
                    if descendant.tag != f"{w}txbxContent":
                        continue
                    for paragraph in selected_iter(descendant, prefixes):
                        if paragraph.tag != f"{w}p":
                            continue
                        text, _ = paragraph_text(paragraph, prefixes)
                        paragraphs.append(text)
            objects.append(
                {
                    "kind": kind,
                    "paragraph_text": paragraphs,
                }
            )
        return objects

    with zipfile.ZipFile(docx_path) as package:
        names = sorted(
            name
            for name in package.namelist()
            if (
                name == "word/document.xml"
                or re.fullmatch(r"word/(header|footer)\d+\.xml", name)
                or name
                in {
                    "word/footnotes.xml",
                    "word/endnotes.xml",
                }
            )
        )
        parts = []
        main_root = None
        for name in names:
            xml_bytes = package.read(name)
            prefixes = namespace_map(xml_bytes)
            root = ET.fromstring(xml_bytes)
            if name == "word/document.xml":
                main_root = root
            parents: Dict[ET.Element, ET.Element] = {}
            selected_nodes: List[ET.Element] = []

            def index_selected(node: ET.Element) -> None:
                selected_nodes.append(node)
                for child in selected_children(node, prefixes):
                    parents[child] = node
                    index_selected(child)

            index_selected(root)

            def is_inside_table(node: ET.Element) -> bool:
                parent = parents.get(node)
                while parent is not None:
                    if parent.tag == f"{w}tbl":
                        return True
                    parent = parents.get(parent)
                return False

            def is_story_paragraph(node: ET.Element) -> bool:
                parent = parents.get(node)
                while parent is not None:
                    if parent.tag in {
                        f"{w}p",
                        f"{w}tbl",
                        f"{w}txbxContent",
                    }:
                        return False
                    parent = parents.get(parent)
                return True

            def structured_hosts(paragraph: ET.Element) -> List[str]:
                hosts: List[str] = []
                nodes = list(selected_iter(paragraph, prefixes))
                if any(node.tag == f"{m}oMath" for node in nodes):
                    hosts.append("equation")
                hosts.extend(
                    kind
                    for node in nodes
                    for kind in [logical_object_kind(node, prefixes)]
                    if kind is not None
                )
                return hosts

            texts = []
            for paragraph in selected_nodes:
                if paragraph.tag != f"{w}p":
                    continue
                if excluded_target_kind == "table" and is_inside_table(paragraph):
                    continue
                value, _ = paragraph_text(paragraph, prefixes)
                if value:
                    texts.append(value)

            topology = []
            for paragraph in selected_nodes:
                if paragraph.tag != f"{w}p" or not is_story_paragraph(paragraph):
                    continue
                value, has_control = paragraph_text(paragraph, prefixes)
                nodes = list(selected_iter(paragraph, prefixes))
                topology.append(
                    {
                        "text": value,
                        "has_non_text_control": has_control,
                        "layout": paragraph_layout(paragraph),
                        "structured_hosts": structured_hosts(paragraph),
                        # A chart-insertion composition may leave one terminal
                        # paragraph sentinel.  Keep a deliberately stricter
                        # signal than the general empty-paragraph projection:
                        # only an empty paragraph with no properties and no
                        # child markup (apart from empty runs) is inert.
                        "inert_terminal_empty": (
                            value == ""
                            and not has_control
                            and paragraph.find(f"{w}pPr") is None
                            and all(child.tag == f"{w}r" and len(child) == 0 for child in paragraph)
                        ),
                        "equation_container": (
                            "block"
                            if any(node.tag == f"{m}oMathPara" for node in nodes)
                            else (
                                "inline" if any(node.tag == f"{m}oMath" for node in nodes) else None
                            )
                        ),
                    }
                )
            tables = []
            for table in (
                []
                if excluded_target_kind == "table"
                else (node for node in selected_nodes if node.tag == f"{w}tbl")
            ):
                rows = []
                for row in table.findall(f"{w}tr"):
                    cells = []
                    for cell in row.findall(f"{w}tc"):
                        cells.append(
                            [
                                paragraph_text(paragraph, prefixes)[0]
                                for paragraph in selected_iter(cell, prefixes)
                                if paragraph.tag == f"{w}p"
                            ]
                        )
                    rows.append(cells)
                tables.append(rows)
            visible_objects = logical_objects(root, prefixes)
            # Word may materialize otherwise empty footnote/endnote story
            # parts during an ordinary SaveAs.  With no visible text or
            # tables those optional parts are packaging noise, not a
            # collateral document change.
            semantic_topology = any(
                record["text"] or record["has_non_text_control"] or record["structured_hosts"]
                for record in topology
            )
            if (
                name == "word/document.xml"
                or texts
                or tables
                or visible_objects
                or semantic_topology
            ):
                parts.append(
                    {
                        "part": name,
                        "nonempty_paragraph_text": texts,
                        "paragraph_topology": topology,
                        "logical_visible_objects": visible_objects,
                        "tables": tables,
                    }
                )

    if main_root is None:
        raise ValueError("word/document.xml is missing")
    sections = []
    for section in main_root.iter(f"{w}sectPr"):
        sections.append(
            {
                "page_size": canonical_attributes(section.find(f"{w}pgSz")),
                "page_margins": canonical_attributes(section.find(f"{w}pgMar")),
                "columns": canonical_attributes(section.find(f"{w}cols")),
                "title_page": section.find(f"{w}titlePg") is not None,
                "text_direction": canonical_attributes(section.find(f"{w}textDirection")),
                "document_grid": canonical_attributes(section.find(f"{w}docGrid")),
            }
        )
    return {"parts": parts, "section_layout": sections}


def compare_saved_visual_preservation_signatures(
    gold: Dict[str, Any],
    pred: Dict[str, Any],
    *,
    allow_textbox_anchor_relocation: bool = False,
    allow_chart_terminal_empty: bool = False,
) -> Dict[str, Any]:
    """Compare logical-object and paragraph-topology projections.

    Whitespace-only run text is normalized only after the complete ordered
    structured-host sequence matches, and only on a paragraph carrying a
    matched host.  Tabs and line/page breaks are explicit controls and are
    never normalized.  A plain terminal paragraph may pair with Word's
    block-equation versus inline-equation serialization only when a later
    Word page-count check confirms that it is a non-rendering sentinel.
    """

    def host_sequence(part: Dict[str, Any]) -> List[str]:
        return [
            host
            for record in part.get("paragraph_topology", [])
            for host in record.get("structured_hosts", [])
        ]

    def canonical_record(record: Dict[str, Any]) -> Dict[str, Any]:
        text = record.get("text", "")
        if (
            record.get("structured_hosts")
            and text
            and text.isspace()
            and not record.get("has_non_text_control")
        ):
            text = ""
        return {
            "text": text,
            "has_non_text_control": bool(record.get("has_non_text_control")),
            "layout": record.get("layout", []),
        }

    def is_plain_empty(record: Dict[str, Any]) -> bool:
        canonical = canonical_record(record)
        return (
            canonical["text"] == ""
            and not canonical["has_non_text_control"]
            and not canonical["layout"]
            and not record.get("structured_hosts")
        )

    def terminal_equation_sentinel_equivalent(
        longer: List[Dict[str, Any]],
        shorter: List[Dict[str, Any]],
    ) -> bool:
        if len(longer) != len(shorter) + 1 or len(longer) < 2:
            return False
        if not is_plain_empty(longer[-1]):
            return False
        if (
            longer[-2].get("equation_container") != "inline"
            or "equation" not in longer[-2].get("structured_hosts", [])
            or shorter[-1].get("equation_container") != "block"
            or "equation" not in shorter[-1].get("structured_hosts", [])
        ):
            return False
        return [canonical_record(item) for item in longer[:-1]] == [
            canonical_record(item) for item in shorter
        ]

    def terminal_empty_after_equation_equivalent(
        longer: List[Dict[str, Any]],
        shorter: List[Dict[str, Any]],
    ) -> bool:
        """Allow one non-rendering empty paragraph after an equation."""
        if len(longer) != len(shorter) + 1 or not shorter:
            return False
        if not is_plain_empty(longer[-1]):
            return False
        if "equation" not in shorter[-1].get("structured_hosts", []):
            return False
        return [canonical_record(item) for item in longer[:-1]] == [
            canonical_record(item) for item in shorter
        ]

    def terminal_empty_after_inserted_chart_equivalent(
        longer: List[Dict[str, Any]],
        shorter: List[Dict[str, Any]],
    ) -> bool:
        """Allow one proven-inert final paragraph after an inserted chart."""
        if (
            not allow_chart_terminal_empty
            or len(longer) != len(shorter) + 1
            or not shorter
            or "chart" not in shorter[-1].get("structured_hosts", [])
            or "chart" not in longer[-2].get("structured_hosts", [])
            or not longer[-1].get("inert_terminal_empty")
        ):
            return False
        return [record.get("structured_hosts", []) for record in longer[:-1]] == [
            record.get("structured_hosts", []) for record in shorter
        ] and [canonical_record(item) for item in longer[:-1]] == [
            canonical_record(item) for item in shorter
        ]

    def inserted_textbox_anchor_relocation_equivalent(
        longer: List[Dict[str, Any]],
        shorter: List[Dict[str, Any]],
    ) -> bool:
        """Match one alternate floating-textbox anchor serialization.

        Word can serialize a newly inserted floating textbox in the text
        paragraph itself or in a new, adjacent empty anchor paragraph. In
        the latter form the pre-existing terminal empty paragraph remains,
        so the topology gains exactly one host-only paragraph. This is
        enabled only by the textbox-insertion composition route and still
        requires an equal Word page count in the caller.
        """
        if (
            not allow_textbox_anchor_relocation
            or len(longer) != len(shorter) + 1
            or host_sequence({"paragraph_topology": longer}) != ["textbox"]
            or host_sequence({"paragraph_topology": shorter}) != ["textbox"]
        ):
            return False
        candidates = [
            index
            for index, record in enumerate(longer)
            if record.get("structured_hosts") == ["textbox"]
            and canonical_record(record)
            == {
                "text": "",
                "has_non_text_control": False,
                "layout": [],
            }
        ]
        if len(candidates) != 1:
            return False
        host_index = candidates[0]
        if host_index < 1:
            return False
        shorter_host_indices = [
            index
            for index, record in enumerate(shorter)
            if record.get("structured_hosts") == ["textbox"]
        ]
        if shorter_host_indices != [host_index - 1]:
            return False
        reduced = longer[:host_index] + longer[host_index + 1 :]
        return [canonical_record(item) for item in reduced] == [
            canonical_record(item) for item in shorter
        ]

    gold_parts = gold.get("parts", [])
    pred_parts = pred.get("parts", [])
    if [part.get("part") for part in gold_parts] != [part.get("part") for part in pred_parts]:
        return {
            "equivalent": False,
            "reason": "story_part_topology",
            "requires_page_count_check": False,
            "normalizations": [],
        }

    normalizations: List[str] = []
    requires_page_count_check = False
    for gold_part, pred_part in zip(gold_parts, pred_parts):
        if gold_part.get("tables", []) != pred_part.get("tables", []):
            return {
                "equivalent": False,
                "reason": "table_topology",
                "requires_page_count_check": False,
                "normalizations": normalizations,
            }
        if gold_part.get("logical_visible_objects", []) != pred_part.get(
            "logical_visible_objects", []
        ):
            return {
                "equivalent": False,
                "reason": "logical_visible_objects",
                "requires_page_count_check": False,
                "normalizations": normalizations,
            }
        if host_sequence(gold_part) != host_sequence(pred_part):
            return {
                "equivalent": False,
                "reason": "structured_object_hosts",
                "requires_page_count_check": False,
                "normalizations": normalizations,
            }

        gold_topology = gold_part.get("paragraph_topology", [])
        pred_topology = pred_part.get("paragraph_topology", [])
        canonical_gold = [canonical_record(record) for record in gold_topology]
        canonical_pred = [canonical_record(record) for record in pred_topology]
        if canonical_gold == canonical_pred:
            if any(
                record.get("structured_hosts")
                and record.get("text", "")
                and record.get("text", "").isspace()
                and not record.get("has_non_text_control")
                for record in gold_topology + pred_topology
            ):
                normalizations.append("matched_host_run_whitespace")
            continue
        if terminal_equation_sentinel_equivalent(
            pred_topology, gold_topology
        ) or terminal_equation_sentinel_equivalent(gold_topology, pred_topology):
            normalizations.append("block_inline_equation_terminal_sentinel")
            requires_page_count_check = True
            continue
        if terminal_empty_after_equation_equivalent(
            pred_topology, gold_topology
        ) or terminal_empty_after_equation_equivalent(gold_topology, pred_topology):
            normalizations.append("equation_terminal_empty_paragraph")
            requires_page_count_check = True
            continue
        if terminal_empty_after_inserted_chart_equivalent(
            pred_topology, gold_topology
        ) or terminal_empty_after_inserted_chart_equivalent(gold_topology, pred_topology):
            normalizations.append("inserted_chart_terminal_inert_empty_paragraph")
            requires_page_count_check = True
            continue
        if inserted_textbox_anchor_relocation_equivalent(
            pred_topology, gold_topology
        ) or inserted_textbox_anchor_relocation_equivalent(gold_topology, pred_topology):
            normalizations.append("inserted_textbox_adjacent_empty_anchor_paragraph")
            requires_page_count_check = True
            continue
        return {
            "equivalent": False,
            "reason": "paragraph_topology",
            "requires_page_count_check": False,
            "normalizations": normalizations,
        }

    return {
        "equivalent": True,
        "reason": None,
        "requires_page_count_check": requires_page_count_check,
        "normalizations": normalizations,
    }


def verify_visual_object_preservation(
    pred_path: Path,
    gold_path: Path,
    *,
    allow_textbox_anchor_relocation: bool = False,
    page_info_reader: Optional[Callable[[str], Dict[str, Any]]] = None,
) -> VerificationResult:
    try:
        gold = saved_visual_preservation_signature(gold_path)
        pred = saved_visual_preservation_signature(pred_path)
    except Exception as exc:
        error = f"Visible object preservation evidence unavailable: {exc}"
        return VerificationResult(
            passed=False,
            message=f"ERROR: {error}",
            details={
                "error_owner": "evaluator",
                "match_mode": "saved_ooxml_visual_object_preservation",
            },
            error=error,
        )
    projection = compare_saved_visual_preservation_signatures(
        gold,
        pred,
        allow_textbox_anchor_relocation=(allow_textbox_anchor_relocation),
    )
    equivalent = projection["equivalent"] and gold.get("section_layout") == pred.get(
        "section_layout"
    )
    page_counts = None
    if equivalent and projection["requires_page_count_check"]:
        if page_info_reader is None:
            error = (
                "Word page-count evidence is required for the saved "
                "visual-preservation normalization"
            )
            return VerificationResult(
                passed=False,
                message=f"ERROR: {error}",
                details={
                    "error_owner": "evaluator",
                    "match_mode": ("saved_ooxml_visual_object_preservation"),
                    "projection_comparison": projection,
                },
                error=error,
            )
        try:
            gold_page_count = page_info_reader(str(gold_path)).get("page_count")
            pred_page_count = page_info_reader(str(pred_path)).get("page_count")
        except Exception as exc:
            error = f"Word page-count evidence unavailable: {exc}"
            return VerificationResult(
                passed=False,
                message=f"ERROR: {error}",
                details={
                    "error_owner": "evaluator",
                    "match_mode": ("saved_ooxml_visual_object_preservation"),
                    "projection_comparison": projection,
                },
                error=error,
            )
        page_counts = {
            "gold": gold_page_count,
            "pred": pred_page_count,
        }
        equivalent = (
            isinstance(gold_page_count, int)
            and gold_page_count >= 1
            and pred_page_count == gold_page_count
        )
    if not equivalent:
        return VerificationResult(
            passed=False,
            message="FAIL: Visible text/structure/page layout outside the requested object differs",
            details={
                "match_mode": "saved_ooxml_visual_object_preservation",
                "verification_source": "saved_ooxml_guardrail",
                "projection_comparison": projection,
                "page_counts": page_counts,
                "gold": gold,
                "pred": pred,
            },
        )
    return VerificationResult(
        passed=True,
        message="PASS: Visible text/structure/page layout is preserved",
        details={
            "match_mode": "saved_ooxml_visual_object_preservation",
            "verification_source": "saved_ooxml_guardrail",
            "projection_comparison": projection,
            "page_counts": page_counts,
        },
    )
