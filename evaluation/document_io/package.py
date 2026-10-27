"""Check DOCX packaging independently of any evaluation method."""
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from pathlib import Path
from typing import Any, Dict, Optional
from .diagnostics import document, parts, charts, relationships

def validate_docx_output(
    path: Path,
    *,
    include_word_open_failure_schema: bool = False,
) -> Optional[Dict[str, Any]]:
    """Return invalid-output details for an unusable OOXML Word file.

    Some strict OOXML schema violations are tolerated and normalized by Word.
    They are checked only after Word has independently refused both a normal
    and repair-mode open; otherwise Word's effective state remains the primary
    benchmark contract.
    """
    try:
        with open(path, "rb") as f:
            header = f.read(8)
    except OSError:
        raise

    if header.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        return {
            "reason": "invalid_output",
            "detected_format": "ole_compound_document",
            "validation_error": "pred.docx is an OLE/legacy/protected Word document, not an OOXML .docx zip package.",
        }

    if not header.startswith(b"PK"):
        return {
            "reason": "invalid_output",
            "detected_format": "not_zip",
            "validation_error": "pred.docx is not an OOXML .docx zip package.",
        }

    try:
        with zipfile.ZipFile(path) as zf:
            name_list = zf.namelist()
            duplicate_parts = sorted(
                name for name, count in Counter(name_list).items() if count > 1
            )
            if duplicate_parts:
                return {
                    "reason": "invalid_output",
                    "detected_format": "ooxml_with_duplicate_parts",
                    "validation_error": (
                        "pred.docx contains duplicate OOXML part name(s): "
                        + ", ".join(duplicate_parts)
                    ),
                    "duplicate_parts": duplicate_parts,
                }

            bad_crc = zf.testzip()
            if bad_crc:
                return {
                    "reason": "invalid_output",
                    "detected_format": "zip_crc_failure",
                    "validation_error": f"pred.docx has a CRC failure in {bad_crc}.",
                    "invalid_part": bad_crc,
                }

            names = set(name_list)
            required_parts = {
                "[Content_Types].xml",
                "_rels/.rels",
                "word/document.xml",
            }
            missing_parts = sorted(required_parts - names)
            if missing_parts:
                return {
                    "reason": "invalid_output",
                    "detected_format": "zip_without_required_docx_parts",
                    "validation_error": f"pred.docx is missing required OOXML part(s): {', '.join(missing_parts)}.",
                    "missing_parts": missing_parts,
                }

            xml_payloads = {}
            xml_roots = {}
            for name in name_list:
                if not (name.endswith((".xml", ".rels")) or name == "[Content_Types].xml"):
                    continue
                try:
                    xml_payloads[name] = zf.read(name)
                    xml_roots[name] = ET.fromstring(xml_payloads[name])
                except ET.ParseError as exc:
                    return {
                        "reason": "invalid_output",
                        "detected_format": "malformed_ooxml",
                        "validation_error": f"pred.docx contains malformed XML in {name}: {exc}",
                        "invalid_part": name,
                    }

            for check, args in (
                (document.diagnose, (xml_roots, xml_payloads)),
                (parts.diagnose, (xml_roots,)),
                (charts.diagnose, (xml_roots,)),
                (relationships.diagnose, (xml_roots, names)),
            ):
                issue = check(*args, include_word_open_failure_schema)
                if issue is not None:
                    return issue
    except zipfile.BadZipFile as e:
        return {
            "reason": "invalid_output",
            "detected_format": "invalid_zip",
            "validation_error": str(e),
        }

    return None
