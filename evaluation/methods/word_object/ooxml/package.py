"""Resolve saved OOXML relationships without interpreting command semantics."""

import posixpath
import zipfile
import xml.etree.ElementTree as ET
from typing import Any, Dict


def part_relationships(
    package: zipfile.ZipFile,
    source_part: str,
) -> Dict[str, Dict[str, Any]]:
    """Resolve a package part's relationships to normalized part names."""
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    source_dir = posixpath.dirname(source_part)
    rels_part = posixpath.join(
        source_dir,
        "_rels",
        posixpath.basename(source_part) + ".rels",
    )
    if rels_part not in package.namelist():
        return {}
    root = ET.fromstring(package.read(rels_part))
    relationships = {}
    for relationship in root.findall(f"{{{rel_ns}}}Relationship"):
        target = relationship.get("Target", "")
        external = relationship.get("TargetMode") == "External"
        relationships[relationship.get("Id", "")] = {
            "type": relationship.get("Type", ""),
            "target": target,
            # A package-absolute target starts with '/', while ZIP member
            # names do not. Relative targets still resolve against the source.
            "part": (
                None
                if external
                else posixpath.normpath(posixpath.join(source_dir, target)).lstrip("/")
            ),
            "external": external,
        }
    return relationships
