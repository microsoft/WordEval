"""Persisted definitions belonging to a named SmartArt gallery operation."""
import xml.etree.ElementTree as ET
import zipfile
from ...ooxml.package import part_relationships
from ....base import VerificationResult


def gallery_definitions(path, command):
    d = 'http://schemas.openxmlformats.org/drawingml/2006/diagram'
    r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    role, expected_type = ('qs','diagramQuickStyle') if command == 'SmartArtStylesGallery' else ('cs','diagramColors')
    ignored_tags = {f'{{{d}}}{name}' for name in ('title','desc','catLst')}
    def canonical(node):
        # Package-local identity and localized gallery descriptions do not
        # define effects. Style-label names and every effect/color child do.
        attrs = tuple(sorted((k,v) for k,v in node.attrib.items() if k != 'uniqueId'))
        children_nodes = [child for child in node if child.tag not in ignored_tags]
        if node.tag in {f'{{{d}}}styleDef', f'{{{d}}}colorsDef'}:
            # Style labels are referenced by name; their serialization order is not paint order.
            labels = [child for child in children_nodes if child.tag == f'{{{d}}}styleLbl']
            names = [child.get('name') for child in labels]
            if None in names or len(set(names)) != len(names):
                raise ValueError('SmartArt gallery style labels are unnamed or duplicated')
            children_nodes = [child for child in children_nodes if child.tag != f'{{{d}}}styleLbl']
            children_nodes += sorted(labels,key=lambda child:child.get('name'))
        children = tuple(canonical(child) for child in children_nodes)
        return (node.tag, attrs, (node.text or '').strip(), children)
    with zipfile.ZipFile(path) as package:
        root = ET.fromstring(package.read('word/document.xml'))
        rels = part_relationships(package, 'word/document.xml')
        values = []
        for links in root.iter(f'{{{d}}}relIds'):
            relationship = rels.get(links.get(f'{{{r}}}{role}'))
            if (not relationship or relationship['external']
                    or not relationship['type'].endswith('/'+expected_type)):
                raise ValueError(f'SmartArt {role} definition relationship is invalid')
            definition = ET.fromstring(package.read(relationship['part']))
            if not definition.findall(f'{{{d}}}styleLbl'):
                raise ValueError(f'SmartArt {role} definition has no style labels')
            values.append(canonical(definition))
        if not values:raise ValueError('No referenced SmartArt gallery definition')
        return values


def verify_gallery_definitions(pred, gold, command):
    try:
        gold_value = gallery_definitions(gold, command)
    except Exception as exc:
        return VerificationResult.error_result(f'Gold SmartArt gallery definition unavailable: {exc}')
    try:
        pred_value = gallery_definitions(pred, command)
    except (ValueError, KeyError, ET.ParseError, zipfile.BadZipFile) as exc:
        return VerificationResult(passed=False,message=f'FAIL: SmartArt gallery definition is invalid ({exc})',
                                  details={'verification_source':'saved_ooxml_gallery_definition'})
    except OSError as exc:
        return VerificationResult.error_result(f'Cannot read prediction SmartArt definitions: {exc}')
    return VerificationResult(passed=gold_value == pred_value,
        message=('PASS: SmartArt gallery definitions match' if gold_value == pred_value
                 else 'FAIL: Named SmartArt gallery definition was changed'),
        details={'verification_source':'saved_ooxml_gallery_definition','idMso':command,
                 'gold':gold_value,'pred':pred_value})
