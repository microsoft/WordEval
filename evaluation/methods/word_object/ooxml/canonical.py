"""ooxml / canonical with explicit evidence dependencies."""

from typing import Any, Dict, List, Optional
import hashlib
import posixpath
import re
import xml.etree.ElementTree as ET
import zipfile


def canonical_saved_element(element: Optional[ET.Element]) -> Any:
    if element is None:
        return None
    return (
        element.tag,
        tuple(sorted(element.attrib.items())),
        (element.text or "").strip(),
        tuple(canonical_saved_element(child) for child in element),
    )


def canonical_visible_element(
    element: Optional[ET.Element],
    *,
    skipped_tags: Optional[set[str]] = None,
    skipped_attribute_names: Optional[set[str]] = None,
) -> Any:
    """Canonicalize visible OOXML properties while dropping edit metadata."""
    if element is None:
        return None
    skipped_tags = skipped_tags or set()
    skipped_attribute_names = skipped_attribute_names or set()
    if element.tag in skipped_tags:
        return None
    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    if element.tag == w + "docGrid" and element.get(w + "type", "default") == "default":
        # A disabled grid retains its settings for later re-enabling. Its
        # inactive pitch is not a visible layout difference from no grid.
        return None

    def local_name(name: str) -> str:
        return name.rsplit("}", 1)[-1]

    attribute_map = dict(element.attrib)
    local_attributes = {local_name(name): value for name, value in attribute_map.items()}
    theme_fallbacks = set()
    if "themeColor" in local_attributes:
        theme_fallbacks.add("val")
    if "themeFill" in local_attributes:
        theme_fallbacks.add("fill")
    if element.tag.rsplit("}", 1)[-1] == "rFonts":
        for direct_name, theme_name in (
            ("ascii", "asciiTheme"),
            ("hAnsi", "hAnsiTheme"),
            ("eastAsia", "eastAsiaTheme"),
            ("cs", "cstheme"),
        ):
            if theme_name in local_attributes:
                theme_fallbacks.add(direct_name)
    attributes = tuple(
        sorted(
            (name, value)
            for name, value in attribute_map.items()
            if local_name(name) not in skipped_attribute_names
            and local_name(name) not in theme_fallbacks
            and not local_name(name).startswith("rsid")
        )
    )
    children = tuple(
        child_value
        for child_value in (
            canonical_visible_element(
                child,
                skipped_tags=skipped_tags,
                skipped_attribute_names=skipped_attribute_names,
            )
            for child in element
        )
        if child_value is not None
    )
    if (
        not attributes
        and not children
        and not (element.text or "").strip()
        and element.tag.rsplit("}", 1)[-1] in {"pPr", "rPr"}
    ):
        return None
    return (
        element.tag,
        attributes,
        (element.text or "").strip(),
        children,
    )


def saved_body_visible_format_signature(
    root: ET.Element, *, exclude_hyperlink_runs: bool = False, hyperlink_extra_only: bool = False
) -> List[Dict[str, Any]]:
    """Return paragraph/run formatting bound to visible text.

    Run boundaries are collapsed when adjacent text has the same direct
    formatting, so harmless serialization splits do not affect the result.
    Language/proofing state is intentionally excluded because it is not a
    visible document-format contract.
    """
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ns = {"w": w}
    p_style = f"{{{w}}}pStyle"
    r_style = f"{{{w}}}rStyle"
    color = f"{{{w}}}color"
    underline = f"{{{w}}}u"
    lang = f"{{{w}}}lang"
    no_proof = f"{{{w}}}noProof"

    paragraphs = []
    for paragraph in root.findall(".//w:body//w:p", ns):
        # A body paragraph can contain a legacy VML textbox whose own
        # paragraphs are descendants. Only leaf paragraphs correspond to
        # actual visible text flows; counting the outer container would
        # duplicate all textbox text and make representation details look
        # like formatting differences.
        if paragraph.findall(".//w:p", ns):
            continue
        hyperlink_run_ids = {
            id(run)
            for hyperlink in paragraph.findall(".//w:hyperlink", ns)
            for run in hyperlink.findall(".//w:r", ns)
        }
        text = "".join(node.text or "" for node in paragraph.findall(".//w:t", ns))
        has_visible_object = (
            paragraph.find(".//w:drawing", ns) is not None
            or paragraph.find(".//w:pict", ns) is not None
        )
        if not text.strip() and not has_visible_object:
            continue

        script_font_attributes = set()
        # East-Asian script detection must include Hangul (syllables
        # U+AC00-D7AF, Jamo U+1100-11FF/U+A960-A97F/U+D7B0-D7FF): for
        # Korean-only text the eastAsia font attributes are render-visible
        # and must stay in the preservation signature.
        if not re.search(
            r"[\u1100-\u11ff\u2e80-\u9fff\ua960-\ua97f\uac00-\ud7ff" r"\uf900-\ufaff]",
            text,
        ):
            script_font_attributes.update({"eastAsia", "eastAsiaTheme"})
        if not re.search(r"[\u0590-\u08ff\ufb1d-\ufeff]", text):
            script_font_attributes.update({"cs", "cstheme"})

        ppr = canonical_visible_element(
            paragraph.find("./w:pPr", ns),
            skipped_tags={p_style, lang, no_proof},
            skipped_attribute_names={
                "paraId",
                "textId",
                *script_font_attributes,
            },
        )
        runs = []
        for run in paragraph.findall(".//w:r", ns):
            is_hyperlink_run = id(run) in hyperlink_run_ids
            if exclude_hyperlink_runs and is_hyperlink_run:
                continue
            if hyperlink_extra_only and not is_hyperlink_run:
                continue
            run_text = "".join(node.text or "" for node in run.findall(".//w:t", ns))
            if not run_text:
                continue
            skipped = {lang, no_proof}
            if hyperlink_extra_only:
                skipped.update({r_style, color, underline})
            rpr = canonical_visible_element(
                run.find("./w:rPr", ns),
                skipped_tags=skipped,
                skipped_attribute_names=script_font_attributes,
            )
            if runs and runs[-1]["format"] == rpr:
                runs[-1]["text"] += run_text
            else:
                runs.append({"text": run_text, "format": rpr})
        paragraphs.append({"paragraph_format": ppr, "runs": runs})
    return paragraphs


def saved_picture_signatures(root: ET.Element, package: zipfile.ZipFile) -> List[Dict[str, Any]]:
    """Extract referenced picture content and visible layout, excluding IDs."""
    wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    pic = "http://schemas.openxmlformats.org/drawingml/2006/picture"
    r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    rel = "http://schemas.openxmlformats.org/package/2006/relationships"
    ns = {"wp": wp, "a": a, "pic": pic}
    names = set(package.namelist())
    relationship_map: Dict[str, str] = {}
    rels_name = "word/_rels/document.xml.rels"
    if rels_name in names:
        relationships = ET.fromstring(package.read(rels_name))
        for relationship in relationships.findall(f"{{{rel}}}Relationship"):
            if relationship.get("TargetMode") == "External":
                continue
            relationship_id = relationship.get("Id")
            target = relationship.get("Target", "").replace("\\", "/")
            if relationship_id and target:
                relationship_map[relationship_id] = posixpath.normpath(
                    posixpath.join("word", target)
                )

    def digest_for_blip(blip: Optional[ET.Element]) -> Optional[str]:
        if blip is None:
            return None
        relationship_id = blip.get(f"{{{r}}}embed") or blip.get(f"{{{r}}}link")
        target = relationship_map.get(relationship_id or "")
        if not target or target not in names:
            return None
        return hashlib.sha256(package.read(target)).hexdigest()

    signatures = []
    for container_type in ("inline", "anchor"):
        for container in root.findall(f".//wp:{container_type}", ns):
            extent = container.find("wp:extent", ns)
            for picture in container.findall(".//pic:pic", ns):
                blip = picture.find("pic:blipFill/a:blip", ns)
                src_rect = picture.find("pic:blipFill/a:srcRect", ns)
                transform = picture.find("pic:spPr/a:xfrm", ns)
                geometry = picture.find("pic:spPr/a:prstGeom", ns)
                signatures.append(
                    {
                        "container": container_type,
                        "extent": (
                            extent.get("cx", "") if extent is not None else "",
                            extent.get("cy", "") if extent is not None else "",
                        ),
                        "media_sha256": digest_for_blip(blip),
                        "crop": canonical_visible_element(src_rect),
                        "transform": canonical_visible_element(transform),
                        "geometry": canonical_visible_element(geometry),
                    }
                )
    return signatures


def normalize_word_default_layout_twips(value: Any) -> Any:
    """Normalize Word's 35.4/36pt default header/column serialization.

    The original core-actionable input stores 720 twips while Word-created
    gold notes store 708 twips for the same environment-owned default.
    This narrow normalization does not cover page margins or page size.
    """
    if isinstance(value, dict):
        return {
            key: (
                "720"
                if key.rsplit("}", 1)[-1] in {"header", "footer", "space"}
                and item in {"708", "720"}
                else normalize_word_default_layout_twips(item)
            )
            for key, item in value.items()
        }
    if isinstance(value, (list, tuple)):
        if (
            len(value) == 2
            and isinstance(value[0], str)
            and value[0].rsplit("}", 1)[-1] in {"header", "footer", "space"}
            and value[1] in {"708", "720"}
        ):
            normalized = (value[0], "720")
            return normalized if isinstance(value, tuple) else list(normalized)
        normalized = [normalize_word_default_layout_twips(item) for item in value]
        return tuple(normalized) if isinstance(value, tuple) else normalized
    return value
