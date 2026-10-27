"""families / revisions / extract_saved with explicit evidence dependencies."""

from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any, Dict
import xml.etree.ElementTree as ET
import zipfile


def render_saved_revision_text(element: ET.Element, w_uri: str, *, accept: bool) -> str:
    """Render the accepted or rejected text view of revision markup."""
    local = element.tag.rsplit("}", 1)[-1]
    if local in {"ins", "moveTo"} and not accept:
        return ""
    if local in {"del", "moveFrom"} and accept:
        return ""
    if local == "t":
        return element.text or ""
    if local == "delText":
        return element.text or ""
    if local == "tab":
        return "\t"
    if local in {"br", "cr"}:
        return "\n"
    return "".join(render_saved_revision_text(child, w_uri, accept=accept) for child in element)


def extract_saved_revision_signature(docx_path: Path) -> Dict[str, Any]:
    """Extract persistent tracking state and every live revision class."""
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ns = {"w": w}
    revision_names = {
        "ins",
        "del",
        "moveFrom",
        "moveTo",
        "rPrChange",
        "pPrChange",
        "tblPrChange",
        "trPrChange",
        "tcPrChange",
        "sectPrChange",
    }
    with zipfile.ZipFile(docx_path) as package:
        names = set(package.namelist())
        document = ET.fromstring(package.read("word/document.xml"))
        settings = (
            ET.fromstring(package.read("word/settings.xml"))
            if "word/settings.xml" in names
            else None
        )
        part_counts = Counter()
        for name in sorted(names):
            if not (
                name.startswith("word/")
                and name.endswith(".xml")
                and not name.startswith("word/theme/")
                and not name.startswith("word/glossary/")
            ):
                continue
            try:
                part = ET.fromstring(package.read(name))
            except ET.ParseError:
                continue
            for node in part.iter():
                local = node.tag.rsplit("}", 1)[-1]
                if local in revision_names:
                    part_counts[(name, local)] += 1

    body = document.find("w:body", ns)
    accepted_text = render_saved_revision_text(
        body if body is not None else document, w, accept=True
    )
    rejected_text = render_saved_revision_text(
        body if body is not None else document, w, accept=False
    )
    revisions = []

    def revision_record(node: ET.Element) -> Dict[str, Any]:
        local = node.tag.rsplit("}", 1)[-1]
        text_value = "".join(
            child.text or ""
            for child in node.iter()
            if child.tag in {f"{{{w}}}t", f"{{{w}}}delText"}
        )
        metadata = tuple(
            sorted(
                (name, value)
                for name, value in node.attrib.items()
                if name.rsplit("}", 1)[-1]
                not in {
                    "id",
                    "date",
                    "dateUtc",
                }
            )
        )
        return {
            "type": local,
            "text": text_value,
            "_metadata": metadata,
            "_date": node.get(f"{{{w}}}date"),
        }

    def compatible_insertion_metadata(
        previous: Dict[str, Any],
        current: Dict[str, Any],
    ) -> bool:
        if previous["_metadata"] != current["_metadata"]:
            return False
        previous_date = previous.get("_date")
        current_date = current.get("_date")
        if previous_date is None or current_date is None:
            return previous_date == current_date
        try:
            first = datetime.fromisoformat(previous_date.replace("Z", "+00:00"))
            second = datetime.fromisoformat(current_date.replace("Z", "+00:00"))
            return abs((second - first).total_seconds()) <= 300
        except (TypeError, ValueError):
            return previous_date == current_date

    def collect_revisions(parent: ET.Element) -> None:
        previous_revision_index = None
        for child_index, child in enumerate(list(parent)):
            local = child.tag.rsplit("}", 1)[-1]
            if local in revision_names:
                current = revision_record(child)
                can_coalesce = (
                    local == "ins"
                    and previous_revision_index == child_index - 1
                    and revisions
                    and revisions[-1]["type"] == "ins"
                    and compatible_insertion_metadata(revisions[-1], current)
                )
                if can_coalesce:
                    revisions[-1]["text"] += current["text"]
                else:
                    revisions.append(current)
                previous_revision_index = child_index
                continue
            previous_revision_index = None
            collect_revisions(child)

    collect_revisions(document)
    for revision in revisions:
        text_value = revision["text"]
        revision["at_body_end"] = bool(
            text_value and accepted_text.rstrip().endswith(text_value.rstrip())
        )
        del revision["_metadata"]
        del revision["_date"]

    normalized_type_counts = Counter(revision["type"] for revision in revisions)
    normalized_part_counts = Counter(
        {("word/document.xml", local): count for local, count in normalized_type_counts.items()}
    )
    for (name, local), count in part_counts.items():
        if name != "word/document.xml":
            normalized_part_counts[(name, local)] += count
            normalized_type_counts[local] += count
    return {
        "track_revisions": (
            settings is not None and settings.find("w:trackRevisions", ns) is not None
        ),
        "revision_count": sum(normalized_part_counts.values()),
        "revision_type_counts": dict(sorted(normalized_type_counts.items())),
        "revision_part_counts": {
            f"{name}:{local}": count
            for (name, local), count in sorted(normalized_part_counts.items())
        },
        "raw_revision_count": sum(part_counts.values()),
        "raw_revision_part_counts": {
            f"{name}:{local}": count for (name, local), count in sorted(part_counts.items())
        },
        "revisions": revisions,
        "accepted_text": accepted_text,
        "rejected_text": rejected_text,
    }
