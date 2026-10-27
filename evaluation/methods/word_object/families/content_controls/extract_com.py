"""Complete native control discovery, deduplicated by physical Word identity."""
from typing import Dict


def extract_content_controls(docx_path: str, *, close_document, open_document) -> Dict:
    doc = None
    try:
        doc = open_document(docx_path)
        controls, seen = [], set()

        def scan_range(rng):
            collection = rng.ContentControls
            for index in range(1, collection.Count + 1):
                control = collection(index)
                identity = control.ID
                if identity in seen:
                    continue
                seen.add(identity)
                controls.append({"index": len(controls) + 1, "type": control.Type,
                    "title": control.Title, "tag": control.Tag, "text": control.Range.Text.strip()})

        def scan_shapes(shapes):
            for index in range(1, shapes.Count + 1):
                shape = shapes(index)
                if shape.Type == 6:
                    scan_shapes(shape.GroupItems)
                elif shape.Type in {1, 5, 17} and shape.TextFrame.HasText:
                    scan_range(shape.TextFrame.TextRange)

        scan_range(doc)
        scan_range(doc.Content)
        scan_shapes(doc.Shapes)
        for index in range(1, doc.Sections.Count + 1):
            section = doc.Sections(index)
            for collection in (section.Headers, section.Footers):
                for kind in (1, 2, 3):
                    part = collection(kind)
                    if part.Exists:
                        scan_range(part.Range)
                        scan_shapes(part.Shapes)
        return {"content_control_count": len(controls), "controls": controls}
    finally:
        close_document(doc)
