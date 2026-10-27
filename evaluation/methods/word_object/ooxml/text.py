"""Prefix-independent XML text extraction."""

import xml.etree.ElementTree as ET


def _xml_text_by_local_name(xml_text: str, name: str) -> str:
    """Return normalized element text without depending on an XML prefix."""
    root = ET.fromstring(xml_text)
    expected = name.casefold()
    for element in root.iter():
        if not isinstance(element.tag, str):
            continue
        local_name = element.tag.rsplit("}", 1)[-1]
        if local_name.casefold() == expected:
            return " ".join("".join(element.itertext()).split())
    return ""
