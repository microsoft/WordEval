"""families / equations / extract_saved with explicit evidence dependencies."""

from evaluation.methods.word_object.effect_commands import OFFICE_MATH_NS, WORDPROCESSINGML_NS
from pathlib import Path
from typing import Any, Dict, List, Optional
import unicodedata
import xml.etree.ElementTree as ET
import zipfile


def extract_saved_equation_operator_signature(docx_path: Path) -> List[Any]:
    m = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    w = WORDPROCESSINGML_NS
    with zipfile.ZipFile(docx_path) as package:
        root = ET.fromstring(package.read("word/document.xml"))

    ignored = {f"{{{m}}}rPr", f"{{{w}}}rPr", f"{{{w}}}rFonts"}
    # xml:space (and other xml-namespace attributes) are non-semantic
    # serialization markers.  The operator text below is already
    # whitespace-stripped, so xml:space="preserve" can never change operator
    # identity; ignore it so two byte-different-but-identical serializations
    # (e.g. one m:t tagged xml:space="preserve") compare equal.
    xml_ns = "http://www.w3.org/XML/1998/namespace"

    def canonical(node: ET.Element) -> Any:
        children = tuple(canonical(child) for child in node if child.tag not in ignored)
        attribs = tuple(
            sorted(
                (key, value)
                for key, value in node.attrib.items()
                if not key.startswith(f"{{{xml_ns}}}")
            )
        )
        return (
            node.tag,
            attribs,
            (node.text or "").strip(),
            children,
        )

    return [canonical(node) for node in root.iter(f"{{{m}}}oMath")]


def extract_saved_logarithm_with_base_signature(docx_path: Path) -> List[Dict[str, str]]:
    """Project the native ``log_2 x`` OMML function structure.

    Word's Object Model can expose a limit object inside the function name
    as the same high-level function type as a real subscript.  The named
    ``logarithm_with_empty_base`` gallery item is narrower: its saved
    function name must contain ``m:sSub``.  Run/control formatting and
    ``xml:space`` are intentionally outside this structural projection.
    """
    m = f"{{{OFFICE_MATH_NS}}}"
    with zipfile.ZipFile(docx_path) as package:
        root = ET.fromstring(package.read("word/document.xml"))

    def text(node: Optional[ET.Element]) -> str:
        if node is None:
            return ""
        value = "".join(descendant.text or "" for descendant in node.iter(f"{m}t"))
        return "".join(unicodedata.normalize("NFKC", value).split())

    signatures: List[Dict[str, str]] = []
    for function in root.iter(f"{m}func"):
        function_name = function.find(f"{m}fName")
        argument = function.find(f"{m}e")
        subscript = function_name.find(f"{m}sSub") if function_name is not None else None
        if subscript is None:
            continue
        signatures.append(
            {
                "function_name": text(subscript.find(f"{m}e")),
                "subscript": text(subscript.find(f"{m}sub")),
                "argument": text(argument),
            }
        )
    return signatures


def extract_saved_three_conditions_signature(docx_path: Path) -> List[Dict[str, Any]]:
    """Project the material layout of the named three-condition case.

    The gallery item is a compact ``m:eqArr`` inside a left-brace
    delimiter.  Formatting-only ``ctrlPr``/``rPr`` descendants are not
    projected.  Material row/distribution properties are normalized to
    their OMML semantic defaults; formatting-only properties remain out of
    scope.
    """
    m = f"{{{OFFICE_MATH_NS}}}"
    with zipfile.ZipFile(docx_path) as package:
        root = ET.fromstring(package.read("word/document.xml"))

    def value(properties: Optional[ET.Element], name: str) -> Optional[str]:
        child = properties.find(f"{m}{name}") if properties is not None else None
        return child.get(f"{m}val") if child is not None else None

    def text(node: ET.Element) -> str:
        raw = "".join(descendant.text or "" for descendant in node.iter(f"{m}t"))
        return "".join(unicodedata.normalize("NFKC", raw).split())

    prefix_by_delimiter: Dict[int, str] = {}
    for equation in root.iter(f"{m}oMath"):
        prefix_parts = []
        for child in equation:
            if child.tag == f"{m}d":
                for delimiter in child.iter(f"{m}d"):
                    prefix_by_delimiter[id(delimiter)] = "".join(prefix_parts)
                break
            if child.tag == f"{m}r":
                prefix_parts.append(text(child))

    def integer_or_raw(raw: Optional[str], default: int) -> Any:
        if raw is None or raw == "":
            return default
        try:
            return int(raw)
        except ValueError:
            return raw

    def boolean_or_raw(raw: Optional[str], default: bool) -> Any:
        if raw is None or raw == "":
            return default
        normalized = str(raw).strip().casefold()
        if normalized in {"true", "on"}:
            return True
        if normalized in {"false", "off"}:
            return False
        try:
            return int(normalized) != 0
        except ValueError:
            return raw

    def on_off_value(
        properties: Optional[ET.Element],
        name: str,
        default: bool = False,
    ) -> Any:
        child = properties.find(f"{m}{name}") if properties is not None else None
        if child is None:
            return default
        raw = child.get(f"{m}val")
        # OMML OnOffType follows the usual Open XML convention: a present
        # element with no val attribute means on; an absent element uses
        # the route default.
        if raw is None or raw == "":
            return True
        return boolean_or_raw(raw, True)

    signatures: List[Dict[str, Any]] = []
    for delimiter in root.iter(f"{m}d"):
        delimiter_properties = delimiter.find(f"{m}dPr")
        begin = value(delimiter_properties, "begChr")
        end = value(delimiter_properties, "endChr")
        element = delimiter.find(f"{m}e")
        if element is None:
            continue
        container = next(
            (child for child in element if child.tag in {f"{m}eqArr", f"{m}m"}),
            None,
        )
        if container is None:
            continue
        kind = container.tag.rsplit("}", 1)[-1]
        properties = container.find(f"{m}eqArrPr" if kind == "eqArr" else f"{m}mPr")
        if kind == "eqArr":
            rows = [text(row) for row in container.findall(f"{m}e")]
        else:
            rows = [
                [text(cell) for cell in row.findall(f"{m}e")] for row in container.findall(f"{m}mr")
            ]
        base_justification = (value(properties, "baseJc") or "center").casefold()
        signatures.append(
            {
                "prefix_text": prefix_by_delimiter.get(id(delimiter), ""),
                "container_kind": kind,
                "delimiter_begin": begin if begin is not None else "(",
                "delimiter_end": end if end is not None else ")",
                "rows": rows,
                "base_justification": base_justification,
                "row_spacing": integer_or_raw(value(properties, "rSp"), 0),
                "row_spacing_rule": integer_or_raw(value(properties, "rSpRule"), 0),
                "maximum_distance": on_off_value(properties, "maxDist"),
                "object_distance": on_off_value(properties, "objDist"),
            }
        )
    return signatures


def extract_delimiter_matrix_signature(docx_path: Path) -> Dict[str, Any]:
    """Read the semantic OMML delimiter/matrix shape, ignoring ctrlPr."""
    m = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    namespaces = {"m": m}
    with zipfile.ZipFile(docx_path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    equations = root.findall(".//m:oMath", namespaces)
    if not equations:
        return {"equation_count": 0}
    equation = equations[0]
    delimiter = equation.find("./m:d", namespaces)
    if delimiter is None:
        return {
            "equation_count": len(equations),
            "prefix_text": "".join(
                node.text or "" for node in equation.findall("./m:r/m:t", namespaces)
            ),
            "delimiter": None,
            "matrix": None,
        }
    delimiter_properties = delimiter.find("./m:dPr", namespaces)
    begin = (
        delimiter_properties.find("./m:begChr", namespaces)
        if delimiter_properties is not None
        else None
    )
    end = (
        delimiter_properties.find("./m:endChr", namespaces)
        if delimiter_properties is not None
        else None
    )
    matrix = delimiter.find("./m:e/m:m", namespaces)
    rows = []
    if matrix is not None:
        for row in matrix.findall("./m:mr", namespaces):
            rows.append(
                [
                    "".join(node.text or "" for node in cell.findall(".//m:t", namespaces))
                    for cell in row.findall("./m:e", namespaces)
                ]
            )
    return {
        "equation_count": len(equations),
        "prefix_text": "".join(
            node.text or "" for node in equation.findall("./m:r/m:t", namespaces)
        ),
        "delimiter": {
            "begin": begin.get(f"{{{m}}}val", "(") if begin is not None else "(",
            "end": end.get(f"{{{m}}}val", ")") if end is not None else ")",
        },
        "matrix": (
            {
                "row_count": len(rows),
                "column_counts": [len(row) for row in rows],
                "values": rows,
            }
            if matrix is not None
            else None
        ),
    }
