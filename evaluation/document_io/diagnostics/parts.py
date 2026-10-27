"""Structural diagnostics extracted without changing rule order."""
import io
import xml.etree.ElementTree as ET
from pathlib import PurePosixPath
from .common import (w_ns, r_ns, m_ns, c_ns, a_ns, cs_ns, wp_ns, wps_ns, a14_ns, mc_ns, invalid_schema, invalid_child_order)

def diagnose(xml_roots, include_word_open_failure_schema=False):
    document_root = xml_roots.get("word/document.xml")
    # Validate a narrow set of header/footer-part constructs that are
    # unequivocally non-schema and were independently observed to make
    # Word reject the prediction.  This remains intentionally smaller
    # than a full OOXML schema validator.
    for part_name, root in xml_roots.items():
        if not (
            part_name.startswith("word/")
            and part_name.endswith(".xml")
        ):
            continue
        basename = PurePosixPath(part_name).name
        if basename.startswith("header") and root.tag != f"{{{w_ns}}}hdr":
            return invalid_schema(
                part_name,
                "w:hdr",
                (
                    f"pred.docx header part {part_name} has root "
                    f"{root.tag.rsplit('}', 1)[-1]!r}; header parts "
                    "must use w:hdr."
                ),
            )
        if basename.startswith("footer") and root.tag != f"{{{w_ns}}}ftr":
            return invalid_schema(
                part_name,
                "w:ftr",
                (
                    f"pred.docx footer part {part_name} has root "
                    f"{root.tag.rsplit('}', 1)[-1]!r}; footer parts "
                    "must use w:ftr."
                ),
            )

        for tabs in root.iter(f"{{{w_ns}}}tabs"):
            for tab in tabs.findall(f"{{{w_ns}}}tab"):
                position = tab.get(f"{{{w_ns}}}pos")
                if position not in {
                    "left", "center", "right", "decimal", "bar",
                }:
                    continue
                return invalid_schema(
                    part_name,
                    "w:tabs/w:tab/@w:pos",
                    (
                        "pred.docx gives w:tab/@w:pos the non-numeric "
                        f"value {position!r}; tab position is stored in "
                        "twips, while alignment belongs in w:val."
                    ),
                )

        invalid_paragraph_children = {
            f"{{{w_ns}}}fldChar": "w:p/w:fldChar",
            f"{{{w_ns}}}instrText": "w:p/w:instrText",
            f"{{{w_ns}}}tab": "w:p/w:tab",
            f"{{{w_ns}}}pBdr": "w:p/w:pBdr",
        }
        for paragraph in root.iter(f"{{{w_ns}}}p"):
            for child in paragraph:
                invalid_label = invalid_paragraph_children.get(child.tag)
                if invalid_label is None:
                    continue
                return invalid_schema(
                    part_name,
                    invalid_label,
                    (
                        f"pred.docx places {invalid_label.rsplit('/', 1)[-1]} "
                        "directly inside w:p; field/tab content must be "
                        "inside w:r and paragraph borders belong in w:pPr."
                    ),
                )

        for table in root.iter(f"{{{w_ns}}}tbl"):
            if table.find(f"{{{w_ns}}}grid") is not None:
                return invalid_schema(
                    part_name,
                    "w:tbl/w:grid",
                    (
                        "pred.docx uses non-schema w:grid inside w:tbl; "
                        "the table grid element is w:tblGrid."
                    ),
                )
        for cell_properties in root.iter(f"{{{w_ns}}}tcPr"):
            if cell_properties.find(f"{{{w_ns}}}gridCols") is not None:
                return invalid_schema(
                    part_name,
                    "w:tcPr/w:gridCols",
                    (
                        "pred.docx uses non-schema w:gridCols inside "
                        "w:tcPr; column spanning is stored in w:gridSpan."
                    ),
                )

        for preset_geometry in root.iter(f"{{{a_ns}}}prstGeom"):
            if preset_geometry.get("prst") != "oval":
                continue
            return invalid_schema(
                part_name,
                "a:prstGeom/@prst",
                (
                    "pred.docx uses non-schema DrawingML preset "
                    "geometry 'oval'; the supported ellipse preset is "
                    "named 'ellipse'."
                ),
            )

        for radical_properties in document_root.iter(f"{{{m_ns}}}radPr"):
            if radical_properties.find(f"{{{m_ns}}}deg") is not None:
                return invalid_schema(
                    "word/document.xml",
                    "m:rad/m:radPr/m:deg",
                    (
                        "pred.docx places m:deg inside m:radPr; the degree "
                        "must be a sibling expression before m:e."
                    ),
                )
        for radical in document_root.iter(f"{{{m_ns}}}rad"):
            child_tags = [child.tag for child in radical]
            degree_tag = f"{{{m_ns}}}deg"
            expression_tag = f"{{{m_ns}}}e"
            if (
                degree_tag not in child_tags
                or expression_tag not in child_tags
                or child_tags.index(degree_tag) > child_tags.index(expression_tag)
            ):
                return invalid_schema(
                    "word/document.xml",
                    "m:rad",
                    (
                        "pred.docx has an invalid radical child sequence; "
                        "m:deg must precede m:e."
                    ),
                )

    return None
