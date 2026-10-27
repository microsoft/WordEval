"""Structural diagnostics extracted without changing rule order."""
import io
import xml.etree.ElementTree as ET
from pathlib import PurePosixPath
from .common import (w_ns, r_ns, m_ns, c_ns, a_ns, cs_ns, wp_ns, wps_ns, a14_ns, mc_ns, invalid_schema, invalid_child_order)

def diagnose(xml_roots, xml_payloads, include_word_open_failure_schema=False):
    if include_word_open_failure_schema:
        ignorable_attribute = f"{{{mc_ns}}}Ignorable"
        for part_name, root in xml_roots.items():
            declared_prefixes = {
                prefix
                for _, (prefix, _) in ET.iterparse(
                    io.BytesIO(xml_payloads[part_name]),
                    events=("start-ns",),
                )
                if prefix
            }
            for element in root.iter():
                ignorable_prefixes = set(
                    (element.get(ignorable_attribute) or "").split()
                )
                missing_prefixes = sorted(
                    ignorable_prefixes - declared_prefixes
                )
                if missing_prefixes:
                    return invalid_schema(
                        part_name,
                        "mc:Ignorable",
                        (
                            "pred.docx has an invalid mc:Ignorable "
                            "attribute that references undeclared "
                            "namespace prefix(es): "
                            + ", ".join(missing_prefixes)
                            + "."
                        ),
                    )
    document_root = xml_roots.get("word/document.xml")
    if document_root is not None:
        # Header/footer references are a strict leading sequence in
        # sectPr.  Several model outputs placed them after page/grid
        # settings (or invented similarly named elements), which Word
        # refuses to open rather than merely normalizing.
        for section_properties in document_root.iter(
            f"{{{w_ns}}}sectPr"
        ):
            invalid_section_children = {
                f"{{{w_ns}}}header": "w:sectPr/w:header",
                f"{{{w_ns}}}footer": "w:sectPr/w:footer",
                f"{{{w_ns}}}differentOddAndEvenPagesHeaderFooter": (
                    "w:sectPr/w:differentOddAndEvenPagesHeaderFooter"
                ),
            }
            for child in section_properties:
                invalid_label = invalid_section_children.get(child.tag)
                if invalid_label is not None:
                    return invalid_schema(
                        "word/document.xml",
                        invalid_label,
                        (
                            f"pred.docx places invalid {invalid_label} "
                            "inside w:sectPr; header/footer links must "
                            "use w:headerReference or w:footerReference, "
                            "and even/odd header mode belongs in settings."
                        ),
                    )
            if include_word_open_failure_schema:
                order_error = invalid_child_order(
                    "word/document.xml",
                    section_properties,
                    "w:sectPr",
                    (
                        ("headerReference", "footerReference"),
                        ("footnotePr",), ("endnotePr",), ("type",),
                        ("pgSz",), ("pgMar",), ("paperSrc",),
                        ("pgBorders",), ("lnNumType",), ("pgNumType",),
                        ("cols",), ("formProt",), ("vAlign",),
                        ("noEndnote",), ("titlePg",),
                        ("textDirection",), ("bidi",), ("rtlGutter",),
                        ("docGrid",), ("printerSettings",),
                        ("sectPrChange",),
                    ),
                    w_ns,
                )
                if order_error:
                    return order_error
            relationship_id = f"{{{r_ns}}}id"
            for reference_name in ("headerReference", "footerReference"):
                for reference in section_properties.findall(
                    f"{{{w_ns}}}{reference_name}"
                ):
                    if reference.get(relationship_id):
                        continue
                    return invalid_schema(
                        "word/document.xml",
                        f"w:sectPr/w:{reference_name}/@r:id",
                        (
                            f"pred.docx has w:{reference_name} without "
                            "the required r:id relationship attribute."
                        ),
                    )

        if include_word_open_failure_schema:
            if document_root.find(f".//{{{w_ns}}}numbering") is not None:
                return invalid_schema(
                    "word/document.xml",
                    "w:numbering",
                    (
                        "pred.docx embeds a w:numbering part inside "
                        "word/document.xml; numbering definitions must "
                        "be stored in a separate word/numbering.xml part."
                    ),
                )
            if document_root.find(f".//{{{w_ns}}}oMath") is not None:
                return invalid_schema(
                    "word/document.xml",
                    "w:oMath",
                    (
                        "pred.docx places an equation in the "
                        "WordprocessingML w:oMath namespace; Office Math "
                        "objects must use m:oMath."
                    ),
                )
            for parent in document_root.iter():
                for child in parent:
                    if child.tag != f"{{{wps_ns}}}wsp":
                        continue
                    if parent.tag == f"{{{a_ns}}}graphicData":
                        continue
                    parent_namespace = (
                        parent.tag[1:].split("}", 1)[0]
                        if parent.tag.startswith("{") else ""
                    )
                    parent_prefix = (
                        "w" if parent_namespace == w_ns
                        else "a" if parent_namespace == a_ns
                        else "unknown"
                    )
                    parent_name = parent.tag.rsplit("}", 1)[-1]
                    return invalid_schema(
                        "word/document.xml",
                        "wps:wsp",
                        (
                            "pred.docx places wps:wsp directly inside "
                            f"{parent_prefix}:{parent_name}; a Word 2010 "
                            "wordprocessing shape must be carried by "
                            "a:graphicData inside w:drawing."
                        ),
                    )
            for run in document_root.iter(f"{{{w_ns}}}r"):
                for child in run:
                    if child.tag in {
                        f"{{{wp_ns}}}inline",
                        f"{{{wp_ns}}}anchor",
                    }:
                        local_name = child.tag.rsplit("}", 1)[-1]
                        return invalid_schema(
                            "word/document.xml",
                            f"w:r/wp:{local_name}",
                            (
                                "pred.docx places wp:"
                                f"{local_name} directly inside w:r; "
                                "wordprocessing drawings must be wrapped "
                                "by w:drawing."
                            ),
                        )
                run_properties = [
                    index
                    for index, child in enumerate(run)
                    if child.tag == f"{{{w_ns}}}rPr"
                ]
                if (
                    len(run_properties) > 1
                    or (
                        bool(run_properties)
                        and run_properties[0] != 0
                    )
                ):
                    return invalid_schema(
                        "word/document.xml",
                        "w:r/w:rPr",
                        (
                            "pred.docx has invalid run-property order; "
                            "w:rPr may occur once and must be the first "
                            "child of w:r."
                        ),
                    )

            for run_properties in document_root.iter(
                f"{{{w_ns}}}rPr"
            ):
                for child in run_properties:
                    local_name = child.tag.rsplit("}", 1)[-1]
                    child_namespace = (
                        child.tag[1:].split("}", 1)[0]
                        if child.tag.startswith("{") else ""
                    )
                    invalid_text_effect = (
                        child.tag == f"{{{a_ns}}}reflect"
                        or child.tag == f"{{{a14_ns}}}textEffects"
                        or (
                            local_name == "effectLst"
                            and child_namespace in {w_ns, a_ns}
                        )
                        or (
                            child_namespace == w_ns
                            and local_name in {
                                "effectExtent", "textEffect",
                            }
                        )
                    )
                    if invalid_text_effect:
                        child_prefix = {
                            w_ns: "w",
                            a_ns: "a",
                            a14_ns: "a14",
                        }.get(child_namespace, "unknown")
                        return invalid_schema(
                            "word/document.xml",
                            (
                                "w:rPr/"
                                f"{child_prefix}:{local_name}"
                            ),
                            (
                                "pred.docx places an invalid text-effect "
                                "element directly inside w:rPr; Word "
                                "requires a supported WordprocessingML "
                                "text-effect element."
                            ),
                        )

            for paragraph_properties in document_root.iter(
                f"{{{w_ns}}}pPr"
            ):
                if paragraph_properties.find(
                    f"{{{w_ns}}}dropCap"
                ) is not None:
                    return invalid_schema(
                        "word/document.xml",
                        "w:pPr/w:dropCap",
                        (
                            "pred.docx places invalid w:dropCap directly "
                            "inside w:pPr; drop-cap settings belong on "
                            "w:framePr."
                        ),
                    )

            allowed_math_control_properties = {
                f"{{{w_ns}}}rPr",
                f"{{{w_ns}}}ins",
                f"{{{w_ns}}}del",
            }
            for control_properties in document_root.iter(
                f"{{{m_ns}}}ctrlPr"
            ):
                for child in control_properties:
                    if child.tag in allowed_math_control_properties:
                        continue
                    child_namespace = (
                        child.tag[1:].split("}", 1)[0]
                        if child.tag.startswith("{") else ""
                    )
                    child_prefix = (
                        "m" if child_namespace == m_ns
                        else "w" if child_namespace == w_ns
                        else "unknown"
                    )
                    local_name = child.tag.rsplit("}", 1)[-1]
                    return invalid_schema(
                        "word/document.xml",
                        f"m:ctrlPr/{child_prefix}:{local_name}",
                        (
                            "pred.docx places an invalid child inside "
                            "m:ctrlPr; math control properties accept "
                            "only w:rPr, w:ins, or w:del."
                        ),
                    )

            allowed_group_character_children = {
                f"{{{m_ns}}}groupChrPr",
                f"{{{m_ns}}}e",
            }
            for group_character in document_root.iter(
                f"{{{m_ns}}}groupChr"
            ):
                for child in group_character:
                    if child.tag in allowed_group_character_children:
                        continue
                    child_namespace = (
                        child.tag[1:].split("}", 1)[0]
                        if child.tag.startswith("{") else ""
                    )
                    child_prefix = (
                        "m" if child_namespace == m_ns
                        else "w" if child_namespace == w_ns
                        else "unknown"
                    )
                    local_name = child.tag.rsplit("}", 1)[-1]
                    return invalid_schema(
                        "word/document.xml",
                        f"m:groupChr/{child_prefix}:{local_name}",
                        (
                            "pred.docx places an invalid child inside "
                            "m:groupChr; a group-character object accepts "
                            "only m:groupChrPr and m:e. Delimiters with "
                            "multiple separated arguments must use m:d."
                        ),
                    )

            if document_root.find(
                f".//{{{m_ns}}}dIntegral"
            ) is not None:
                return invalid_schema(
                    "word/document.xml",
                    "m:dIntegral",
                    (
                        "pred.docx uses the non-schema m:dIntegral "
                        "element for an integral; Office Math integrals "
                        "must be represented by an m:nary object."
                    ),
                )

            allowed_nary_children = {
                f"{{{m_ns}}}naryPr",
                f"{{{m_ns}}}sub",
                f"{{{m_ns}}}sup",
                f"{{{m_ns}}}e",
            }
            for nary in document_root.iter(f"{{{m_ns}}}nary"):
                invalid_math_attributes = sorted(
                    name.rsplit("}", 1)[-1]
                    for name in nary.attrib
                    if name.startswith(f"{{{m_ns}}}")
                )
                if invalid_math_attributes:
                    attribute_name = invalid_math_attributes[0]
                    return invalid_schema(
                        "word/document.xml",
                        f"m:nary/@m:{attribute_name}",
                        (
                            "pred.docx places Office Math properties "
                            "directly on m:nary; properties such as "
                            "m:chr and m:limLoc must be children of "
                            "m:naryPr."
                        ),
                    )
                for child in nary:
                    if child.tag in allowed_nary_children:
                        continue
                    child_namespace = (
                        child.tag[1:].split("}", 1)[0]
                        if child.tag.startswith("{") else ""
                    )
                    child_prefix = (
                        "m" if child_namespace == m_ns
                        else "w" if child_namespace == w_ns
                        else "unknown"
                    )
                    local_name = child.tag.rsplit("}", 1)[-1]
                    return invalid_schema(
                        "word/document.xml",
                        f"m:nary/{child_prefix}:{local_name}",
                        (
                            "pred.docx places an invalid child inside "
                            "m:nary; an n-ary operator accepts only "
                            "m:naryPr, m:sub, m:sup, and m:e."
                        ),
                    )

            highlight_values = {
                "black", "blue", "cyan", "darkBlue", "darkCyan",
                "darkGray", "darkGreen", "darkMagenta", "darkRed",
                "darkYellow", "green", "lightGray", "magenta", "none",
                "red", "white", "yellow",
            }
            on_off_values = {"true", "false", "on", "off", "0", "1"}
            val_attribute = f"{{{w_ns}}}val"
            for highlight in document_root.iter(f"{{{w_ns}}}highlight"):
                invalid_attributes = sorted(
                    name.rsplit("}", 1)[-1]
                    for name in highlight.attrib
                    if name.startswith(f"{{{w_ns}}}")
                    and name != val_attribute
                )
                highlight_value = highlight.get(val_attribute)
                if (
                    highlight_value not in highlight_values
                    or invalid_attributes
                ):
                    return invalid_schema(
                        "word/document.xml",
                        "w:rPr/w:highlight",
                        (
                            "pred.docx has invalid w:highlight syntax; "
                            "w:val must be a highlight color and no other "
                            "WordprocessingML attributes are allowed."
                        ),
                    )
            for outline in document_root.iter(f"{{{w_ns}}}outline"):
                invalid_attributes = sorted(
                    name.rsplit("}", 1)[-1]
                    for name in outline.attrib
                    if name.startswith(f"{{{w_ns}}}")
                    and name != val_attribute
                )
                outline_value = outline.get(val_attribute)
                if (
                    outline_value is not None
                    and outline_value not in on_off_values
                ) or invalid_attributes:
                    return invalid_schema(
                        "word/document.xml",
                        "w:rPr/w:outline",
                        (
                            "pred.docx has invalid w:outline syntax; "
                            "w:outline is an on/off property and accepts "
                            "only the optional w:val attribute."
                        ),
                    )

            for text_direction in document_root.iter(
                f"{{{w_ns}}}textDirection"
            ):
                invalid_attributes = sorted(
                    name.rsplit("}", 1)[-1]
                    for name in text_direction.attrib
                    if name.startswith(f"{{{w_ns}}}")
                    and name != val_attribute
                )
                if (
                    text_direction.get(val_attribute) is None
                    or invalid_attributes
                ):
                    return invalid_schema(
                        "word/document.xml",
                        "w:tcPr/w:textDirection",
                        (
                            "pred.docx has invalid w:textDirection syntax; "
                            "the required direction is stored in w:val."
                        ),
                    )

            for table_properties in document_root.iter(
                f"{{{w_ns}}}tblPr"
            ):
                order_error = invalid_child_order(
                    "word/document.xml",
                    table_properties,
                    "w:tblPr",
                    (("tblStyle",), ("tblW",), ("tblLook",)),
                    w_ns,
                )
                if order_error:
                    return order_error
            for cell_properties in document_root.iter(
                f"{{{w_ns}}}tcPr"
            ):
                order_error = invalid_child_order(
                    "word/document.xml",
                    cell_properties,
                    "w:tcPr",
                    (("tcW",), ("tcBorders",), ("textDirection",)),
                    w_ns,
                )
                if order_error:
                    return order_error

        for sdt_properties in document_root.iter(f"{{{w_ns}}}sdtPr"):
            if sdt_properties.find(f"{{{w_ns}}}docPart") is not None:
                return invalid_schema(
                    "word/document.xml",
                    "w:sdtPr/w:docPart",
                    (
                        "pred.docx places invalid w:docPart inside w:sdtPr; "
                        "Word requires a valid content-control type such as "
                        "w:docPartObj or w:docPartList."
                    ),
                )

    return None
