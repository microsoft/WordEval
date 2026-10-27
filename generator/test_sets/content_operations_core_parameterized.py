"""
Test case generator for content_operations Ã— core_parameterized.

This UNIFIED generator handles parameterized content-operation commands:
- INSERT: Add or place content using gallery variants (symbols, tables, text boxes, etc.)
- STYLE/CHOICE: Apply a selected variant from a gallery (header/footer style, TOC style, etc.)

All test cases are generated from the same flow:
- Create a command-specific precondition document
- Apply one parameter variant
- Save result as gold + emit spec.json

Key Design:
- Configuration-driven approach (CommandConfig dataclass)
- Shared precondition templates (default, with_sources, with_headings, with_bibliography)
- Command-specific primary verification overrides via config
- Variant definitions loaded from instructions/content_operations_x_core_parameterized_variants.json
"""

import json
import os
import shutil
import tempfile
import time
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, List, Optional
from xml.etree import ElementTree as ET

from ..config import (
    COMMAND_SETS_DIR, BENCHMARK_OUTPUT_DIR,
    ONE_WAY_VARIATIONS, FILL_VALUES,
)
from ..execution.setup import ensure_word, finalize_word
from generator.test_sets.spec_utils import write_spec_if_changed

# Path to variant file
CONTENT_OPERATIONS_GALLERY_VARIANTS_PATH = Path(__file__).parent.parent.parent / 'instructions' / 'content_operations_x_core_parameterized_variants.json'

# Default input configuration for this generator
DEFAULT_INPUT = {
    'full_text': 'Alpha TARGET Omega.',
    'cursor_after': 'TARGET',
    'filename': 'doc_in.docx',
}

NS_CORE_CP = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
NS_DC = 'http://purl.org/dc/elements/1.1/'
NS_DCTERMS = 'http://purl.org/dc/terms/'
NS_XSI = 'http://www.w3.org/2001/XMLSchema-instance'
NS_APP = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
NS_COVER_PAGE_PROPS = 'http://schemas.microsoft.com/office/2006/coverPageProps'
# Well-known storeItemID for CoverPageProperties custom XML data store
COVER_PAGE_PROPS_STORE_ID = '{55AF091B-3C7A-41E3-B477-F2FDAA23CFDA}'
# Mapping from field_mapping keys to CoverPageProperties element names
COVER_PAGE_PROPS_FIELD_MAP = {
    'Abstract': 'Abstract',
    'Address': 'CompanyAddress',
    'Email': 'CompanyEmail',
    'Publish Date': 'PublishDate',
}

@dataclass
class CommandConfig:
    """Configuration for command-specific spec generation."""
    idMso: str
    precondition_type: str = 'default'
    primary_method: str = 'structure_verification'
    primary_params: Optional[Dict] = None
    description: str = 'Word COM document structure comparison'
    skip_reason: Optional[str] = None
    needs_cursor_positioning: bool = True


def get_command_configs() -> Dict[str, CommandConfig]:
    """
    Build explicit command configurations (core_actionable style).

    Each command declares:
    - precondition template
    - cursor handling override
    - Word COM primary verification method
    """
    return {
        # ---------------------------------------------------------------------
        # Building blocks / gallery insertion
        # ---------------------------------------------------------------------
        'AutoTextGallery': CommandConfig(
            idMso='AutoTextGallery',
            skip_reason="AutoText entry 'BOLD_TARGET,' not present in default Normal template",
        ),
        'CoverPageInsertGallery': CommandConfig(
            idMso='CoverPageInsertGallery',
            precondition_type='with_sources',
            primary_method='cover_page_verification',
            description='Word COM cover page comparison',
            needs_cursor_positioning=False,
        ),
        'QuickPartsInsertGallery': CommandConfig(
            idMso='QuickPartsInsertGallery',
            skip_reason='GUI output is not reliably reproducible',
            needs_cursor_positioning=False,
        ),
        'QuickTablesInsertGallery': CommandConfig(
            idMso='QuickTablesInsertGallery',
            precondition_type='with_sources',
            primary_method='table_verification',
            description='Word COM table structure comparison',
            needs_cursor_positioning=False,
        ),
        'InsertBuildingBlocksEquationsGallery': CommandConfig(
            idMso='InsertBuildingBlocksEquationsGallery',
            precondition_type='with_sources',
            primary_method='equation_verification',
            description='Word COM equation comparison',
            needs_cursor_positioning=False,
        ),

        # ---------------------------------------------------------------------
        # Bibliography / TOC
        # ---------------------------------------------------------------------
        'CitationInsert': CommandConfig(
            idMso='CitationInsert',
            precondition_type='with_sources',
            primary_method='citation_verification',
            primary_params={},
            description='Citation field inserted referencing source',
            needs_cursor_positioning=False,
        ),
        'BibliographyGallery': CommandConfig(
            idMso='BibliographyGallery',
            precondition_type='with_sources',
            primary_method='bibliography_verification',
            description='Word COM bibliography text and heading comparison',
            needs_cursor_positioning=False,
        ),
        'BibliographyStyle': CommandConfig(
            idMso='BibliographyStyle',
            precondition_type='with_bibliography',
            primary_method='bibliography_verification',
            description='Word COM bibliography text comparison after style change',
            needs_cursor_positioning=False,
        ),
        'TableOfContentsGallery': CommandConfig(
            idMso='TableOfContentsGallery',
            precondition_type='with_headings',
            primary_method='toc_verification',
            description='Word COM TOC content comparison',
            needs_cursor_positioning=False,
        ),
        'TableOfContentsAddTextGallery': CommandConfig(
            idMso='TableOfContentsAddTextGallery',
            precondition_type='with_existing_toc',
            primary_method='toc_text_level_verification',
            description='Word COM TOC text level and updated TOC content comparison',
            needs_cursor_positioning=False,
        ),

        # ---------------------------------------------------------------------
        # Header / Footer family
        # ---------------------------------------------------------------------
        'HeaderInsertGallery': CommandConfig(
            idMso='HeaderInsertGallery',
            precondition_type='with_sources',
            primary_method='header_footer_content_check',
            primary_params={'part_type': 'header', 'condition': 'style_match'},
            description='Header style matches gold',
            needs_cursor_positioning=False,
        ),
        'FooterInsertGallery': CommandConfig(
            idMso='FooterInsertGallery',
            precondition_type='with_sources',
            primary_method='header_footer_content_check',
            primary_params={'part_type': 'footer', 'condition': 'style_match'},
            description='Footer style matches gold',
            needs_cursor_positioning=False,
        ),
        'PageNumbersInHeaderInsertGallery': CommandConfig(
            idMso='PageNumbersInHeaderInsertGallery',
            precondition_type='with_sources',
            primary_method='header_footer_content_check',
            primary_params={'part_type': 'header', 'condition': 'style_match'},
            description='Header page number style matches gold',
            needs_cursor_positioning=False,
        ),
        'PageNambersInFooterInsertGallery': CommandConfig(
            idMso='PageNambersInFooterInsertGallery',
            precondition_type='with_sources',
            primary_method='header_footer_content_check',
            primary_params={'part_type': 'footer', 'condition': 'style_match'},
            description='Footer page number style matches gold',
            needs_cursor_positioning=False,
        ),

        # ---------------------------------------------------------------------
        # Shapes / text boxes / wordart
        # ---------------------------------------------------------------------
        'TextBoxInsertGallery': CommandConfig(
            idMso='TextBoxInsertGallery',
            precondition_type='with_sources',
            primary_method='textbox_gallery_verification',
            description='Word COM textbox shape comparison',
            needs_cursor_positioning=False,
        ),
        'GalleryAllShapesAndCanvas': CommandConfig(
            idMso='GalleryAllShapesAndCanvas',
            precondition_type='with_sources',
            primary_method='shape_canvas_verification',
            description='Word COM shape/canvas property comparison',
            needs_cursor_positioning=False,
        ),
        'TextBoxInsert': CommandConfig(
            idMso='TextBoxInsert',
            precondition_type='with_sources',
            primary_method='textbox_insert_verification',
            description='Word COM textbox insertion comparison',
            needs_cursor_positioning=False,
        ),
        'TextBoxInsertVerticalWord': CommandConfig(
            idMso='TextBoxInsertVerticalWord',
            precondition_type='with_sources',
            primary_method='textbox_insert_verification',
            description='Word COM textbox insertion comparison',
            needs_cursor_positioning=False,
        ),
        'PageNambersInMarginsInsertGallery': CommandConfig(
            idMso='PageNambersInMarginsInsertGallery',
            precondition_type='with_sources',
            primary_method='margin_page_numbers_verification',
            description='Word COM margin page numbers comparison',
            needs_cursor_positioning=False,
        ),
        'WordArtInsertGallery': CommandConfig(
            idMso='WordArtInsertGallery',
            skip_reason="Variant/style mismatch with Word's current WordArt gallery",
        ),
        'WordArtInsertGalleryClassic': CommandConfig(
            idMso='WordArtInsertGalleryClassic',
            skip_reason="Variant/style mismatch with Word's current WordArt gallery",
        ),

        # ---------------------------------------------------------------------
        # Other insertions
        # ---------------------------------------------------------------------
        # DropCapInsertGallery moved to content_formatting_x_core_parameterized:
        # it modifies Paragraphs(1).DropCap.Position, i.e. paragraph formatting,
        # not content insertion.
        'SymbolInsertGallery': CommandConfig(
            idMso='SymbolInsertGallery',
            primary_method='text_content_verification',
            description='Word COM text content comparison',
        ),
        'TableInsertGallery': CommandConfig(
            idMso='TableInsertGallery',
            precondition_type='with_sources',
            primary_method='table_verification',
            description='Word COM table structure comparison',
            needs_cursor_positioning=False,
        ),
        'PageNumberFieldInsertGallery': CommandConfig(
            idMso='PageNumberFieldInsertGallery',
            precondition_type='with_sources',
            primary_method='page_number_field_verification',
            description='Word COM page number field comparison',
            needs_cursor_positioning=False,
        ),

        # ---------------------------------------------------------------------
        # Skip-only commands
        # ---------------------------------------------------------------------
        'ScreenshotInsertGallery': CommandConfig(
            idMso='ScreenshotInsertGallery',
            skip_reason='requires_screenshot_selection',
        ),
        'SignatureLineInsertMenu': CommandConfig(
            idMso='SignatureLineInsertMenu',
            skip_reason='requires_certificate_setup',
        ),
    }


def _inject_cover_page_properties(zin: zipfile.ZipFile, zout: zipfile.ZipFile, cpp_values: Dict[str, str]):
    """
    Inject or update a CoverPageProperties custom XML part in the docx package.

    This allows Word to auto-populate data-bound content controls
    (Abstract, CompanyAddress, CompanyEmail, PublishDate) when a cover page
    building block is inserted.

    Args:
        zin: Source zip (for reading existing custom XML items)
        zout: Destination zip (already has all items written; we add new ones)
        cpp_values: Dict mapping CoverPageProperties element names to values
                    e.g. {'Abstract': 'text', 'CompanyAddress': '123 Main St'}
    """
    if not cpp_values:
        return

    # Find existing CoverPageProperties part or determine next item number
    existing_items = [n for n in zin.namelist() if n.startswith('customXml/item') and n.endswith('.xml') and 'Props' not in n]
    next_idx = len(existing_items) + 1

    # Check if a CoverPageProperties part already exists
    NS_DS = 'http://schemas.openxmlformats.org/officeDocument/2006/customXml'
    for item_name in existing_items:
        props_name = item_name.replace('.xml', 'Props.xml').replace('item', 'itemProps')
        # Actually the naming is customXml/itemProps{N}.xml
        idx = item_name.replace('customXml/item', '').replace('.xml', '')
        props_name = f'customXml/itemProps{idx}.xml'
        try:
            props_data = zin.read(props_name).decode('utf-8', errors='replace')
            if COVER_PAGE_PROPS_STORE_ID.lower() in props_data.lower():
                # CoverPageProperties part already exists; update it
                existing_data = zin.read(item_name)
                root = ET.fromstring(existing_data)
                for el_name, el_value in cpp_values.items():
                    el = root.find(f'{{{NS_COVER_PAGE_PROPS}}}{el_name}')
                    if el is None:
                        el = ET.SubElement(root, f'{{{NS_COVER_PAGE_PROPS}}}{el_name}')
                    el.text = el_value
                # Overwrite in zout (already written; we re-add)
                updated_xml = ET.tostring(root, encoding='unicode', xml_declaration=True)
                zout.writestr(item_name, updated_xml)
                return
        except (KeyError, ET.ParseError):
            continue

    # No existing CoverPageProperties part — create new one
    item_name = f'customXml/item{next_idx}.xml'
    props_name = f'customXml/itemProps{next_idx}.xml'
    rels_name = f'customXml/_rels/item{next_idx}.xml.rels'

    # Build CoverPageProperties XML
    ET.register_namespace('', NS_COVER_PAGE_PROPS)
    root = ET.Element(f'{{{NS_COVER_PAGE_PROPS}}}CoverPageProperties')
    for el_name, el_value in cpp_values.items():
        el = ET.SubElement(root, f'{{{NS_COVER_PAGE_PROPS}}}{el_name}')
        el.text = el_value
    item_xml = ET.tostring(root, encoding='unicode', xml_declaration=True)
    zout.writestr(item_name, item_xml)

    # Build itemProps XML with the well-known storeItemID
    props_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="no"?>'
        f'<ds:datastoreItem ds:itemID="{COVER_PAGE_PROPS_STORE_ID}"'
        f' xmlns:ds="{NS_DS}">'
        '<ds:schemaRefs>'
        f'<ds:schemaRef ds:uri="{NS_COVER_PAGE_PROPS}"/>'
        '</ds:schemaRefs>'
        '</ds:datastoreItem>'
    )
    zout.writestr(props_name, props_xml)

    # Build rels for this custom XML item
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1"'
        ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"'
        f' Target="itemProps{next_idx}.xml"/>'
        '</Relationships>'
    )
    zout.writestr(rels_name, rels_xml)


def normalize_docx_package_identity(docx_path: Path, field_mapping: Optional[Dict] = None):
    """
    Normalize docProps metadata directly in the docx package.

    Args:
        docx_path: Path to the .docx file to normalize.
        field_mapping: Optional dict of document property values to set.
            Recognized keys and where they are written:
              core.xml  — Title (dc:title), Subject (dc:subject),
                          Author/creator (dc:creator)
              app.xml   — Company
              CoverPageProperties custom XML —
                          Abstract, CompanyAddress, CompanyEmail, PublishDate
            This allows Word to auto-populate content controls bound to
            built-in document properties when building blocks are inserted.
    """
    if not docx_path.exists():
        return

    fd, tmp_name = tempfile.mkstemp(suffix='.docx', prefix='meta_fix_')
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        # Keep canonical OpenXML prefixes so xsi:type="dcterms:..." remains valid.
        ET.register_namespace('cp', NS_CORE_CP)
        ET.register_namespace('dc', NS_DC)
        ET.register_namespace('dcterms', NS_DCTERMS)
        ET.register_namespace('xsi', NS_XSI)
        ET.register_namespace('', NS_APP)
        with zipfile.ZipFile(docx_path, 'r') as zin:
            # Read all entries upfront so zin.read() isn't affected by zout.writestr()
            # (writestr mutates the ZipInfo.header_offset of the original object)
            entries = {}
            for item in zin.infolist():
                entries[item.filename] = zin.read(item.filename)

        with zipfile.ZipFile(docx_path, 'r') as zin, zipfile.ZipFile(tmp_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for filename, data in entries.items():

                if filename == 'docProps/core.xml':
                    try:
                        root = ET.fromstring(data)
                        creator = root.find(f'{{{NS_DC}}}creator')
                        if creator is None:
                            creator = ET.SubElement(root, f'{{{NS_DC}}}creator')
                        creator.text = FILL_VALUES['Author']

                        last_modified_by = root.find(f'{{{NS_CORE_CP}}}lastModifiedBy')
                        if last_modified_by is None:
                            last_modified_by = ET.SubElement(root, f'{{{NS_CORE_CP}}}lastModifiedBy')
                        last_modified_by.text = FILL_VALUES['Author']

                        # Set document properties from field_mapping
                        if field_mapping:
                            if 'Title' in field_mapping:
                                el = root.find(f'{{{NS_DC}}}title')
                                if el is None:
                                    el = ET.SubElement(root, f'{{{NS_DC}}}title')
                                el.text = field_mapping['Title']
                            if 'Subtitle' in field_mapping:
                                el = root.find(f'{{{NS_DC}}}subject')
                                if el is None:
                                    el = ET.SubElement(root, f'{{{NS_DC}}}subject')
                                el.text = field_mapping['Subtitle']
                            if 'Author' in field_mapping:
                                creator.text = field_mapping['Author']

                        data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                    except Exception:
                        pass

                elif filename == 'docProps/app.xml':
                    try:
                        root = ET.fromstring(data)
                        company = root.find(f'{{{NS_APP}}}Company')
                        if company is None:
                            company = ET.SubElement(root, f'{{{NS_APP}}}Company')
                        company.text = FILL_VALUES['Company']

                        manager = root.find(f'{{{NS_APP}}}Manager')
                        if manager is None:
                            manager = ET.SubElement(root, f'{{{NS_APP}}}Manager')
                        manager.text = FILL_VALUES['Author']

                        # Set company from field_mapping
                        if field_mapping and 'Company' in field_mapping:
                            company.text = field_mapping['Company']

                        data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                    except Exception:
                        pass

                zout.writestr(filename, data)

            # Inject CoverPageProperties custom XML part if field_mapping has
            # any CoverPageProperties keys (Abstract, Address, Email, etc.)
            if field_mapping:
                cpp_values = {}
                for fm_key, cpp_el in COVER_PAGE_PROPS_FIELD_MAP.items():
                    if fm_key in field_mapping:
                        cpp_values[cpp_el] = field_mapping[fm_key]
                if cpp_values:
                    _inject_cover_page_properties(zin, zout, cpp_values)

        os.replace(str(tmp_path), str(docx_path))
    finally:
        try:
            if tmp_path.exists():
                tmp_path.unlink()
        except Exception:
            pass


def load_content_operations_gallery_variants() -> Dict:
    """Load variant definitions from content operations gallery variant file."""
    all_variants = {}

    # Load content operations gallery variants (auto-discovered)
    if CONTENT_OPERATIONS_GALLERY_VARIANTS_PATH.exists():
        with open(CONTENT_OPERATIONS_GALLERY_VARIANTS_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
            all_variants.update(data.get('commands', {}))
    else:
        print(f"[WARNING] Content operations gallery variants file not found: {CONTENT_OPERATIONS_GALLERY_VARIANTS_PATH}")

    return all_variants


def extract_parameters(wd, idMso: str) -> List[Dict]:
    """
    Extract all parameters for a command from variant definitions.

    Commands are defined in:
    - content_operations_gallery_variants.json: Auto-discovered content operations gallery variants

    Returns:
        List of parameter dictionaries, each with 'label' and command-specific fields.
    """
    # Load variant definitions
    gallery_variants = load_content_operations_gallery_variants()

    # Check if this command has defined variants
    if idMso not in gallery_variants:
        print(f"    [WARNING] Command '{idMso}' not found in variant definitions")
        return []

    variants = gallery_variants[idMso].get('variants', [])
    # Filter to only variants marked for generation
    variants = [v for v in variants if v.get('generate', True)]

    if not variants:
        print(f"    [WARNING] No variants to generate for '{idMso}'")
        return []

    print(f"    [Variants] {idMso}: {len(variants)} variants")
    parameters = []
    for variant in variants:
        param = {
            'label': variant['label'],
            'instruction': variant.get('instruction'),
        }
        # Add all variant-specific fields dynamically
        for key in ['char', 'code', 'name', 'style_name', 'position', 'description',
                    'rows', 'cols', 'type', 'index', 'field',
                    'level', 'shape_type', 'category',
                    'field_mapping', 'layout_description', 'source_tag',
                    'positional_fill',
                    'orientation', 'left', 'top', 'width', 'height',
                    'diagonal_type']:
            if key in variant:
                param[key] = variant[key]
        parameters.append(param)

    return parameters


def _get_normal_template(wd):
    return wd.Application.Templates(1)


def _find_building_block_entry(template, name: str, expected_type: Optional[int] = None):
    name_lower = name.strip().lower()
    for i in range(1, template.BuildingBlockEntries.Count + 1):
        bb = template.BuildingBlockEntries.Item(i)
        if bb.Name.strip().lower() != name_lower:
            continue
        if expected_type is None or bb.Type.Index == expected_type:
            return bb
    return None


def _find_building_block_entry_across_templates(wd, name: str, expected_type: Optional[int] = None):
    for t_idx in range(1, wd.Application.Templates.Count + 1):
        template = wd.Application.Templates(t_idx)
        entry = _find_building_block_entry(template, name, expected_type)
        if entry:
            return entry
    return None


def _apply_symbol_insert_gallery(wd, param: Dict):
    sel = wd.Selection
    sel.InsertSymbol(CharacterNumber=param['code'], Font="", Unicode=True)


def _apply_table_insert_gallery(wd, param: Dict):
    sel = wd.Selection
    tbl = wd.ActiveDocument.Tables.Add(Range=sel.Range, NumRows=param['rows'], NumColumns=param['cols'])
    tbl.Borders.Enable = 1


def _apply_named_building_block_from_normal_template(wd, param: Dict):
    sel = wd.Selection
    bb_entry = _find_building_block_entry_across_templates(wd, param['name'])
    if not bb_entry:
        # Fallback: try Normal template direct lookup (case-sensitive)
        template = _get_normal_template(wd)
        bb_entry = template.BuildingBlockEntries.Item(param['name'])
    bb_entry.Insert(Where=sel.Range, RichText=True)


def _fill_positional_content_controls(target_range, values: list):
    """Fill title-less ContentControls by position order.

    Re-queries the collection each time because setting Range.Text on a
    content control can destroy it, shifting all subsequent indices.
    """
    try:
        for val in values:
            controls = target_range.ContentControls
            filled = False
            for i in range(1, controls.Count + 1):
                try:
                    cc = controls(i)
                    title = (cc.Title or '').strip()
                    tag = (cc.Tag or '').strip()
                    if not title and not tag:
                        cc.Range.Text = str(val)
                        filled = True
                        break
                except Exception as e:
                    print(f"    [WARNING] _fill_positional_content_controls[{i}]: {e}")
            if not filled:
                print(f"    [WARNING] No unnamed content control found for value: {val}")
    except Exception:
        pass


def _fill_content_controls_in_range(target_range, field_mapping: Dict):
    """Fill ContentControls in a single range by matching Title/Tag to field_mapping keys."""
    try:
        controls = target_range.ContentControls
        for i in range(1, controls.Count + 1):
            try:
                cc = controls(i)
                title = (cc.Title or '').strip()
                tag = (cc.Tag or '').strip()
                value = field_mapping.get(title) or field_mapping.get(tag)
                if value is not None:
                    cc.Range.Text = str(value)
            except Exception as e:
                print(f"    [WARNING] _fill_content_controls_in_range[{i}]: {e}")
    except Exception:
        pass


def _fill_content_controls(doc_or_range, field_mapping: Dict, *, include_shapes: bool = False):
    """Fill ContentControls by matching Title/Tag to field_mapping keys.

    Args:
        doc_or_range: A Word Range object, or the Document object when include_shapes=True.
        field_mapping: Dict mapping CC Title/Tag to fill values.
        include_shapes: If True, also scan Shapes' TextFrame ContentControls
                        (needed for cover pages where CCs live inside Shapes).
    """
    if not field_mapping:
        return

    if not include_shapes:
        _fill_content_controls_in_range(doc_or_range, field_mapping)
        return

    # When include_shapes=True, doc_or_range should be the Document object.
    doc = doc_or_range

    # 1. Scan doc.Content (body)
    try:
        _fill_content_controls_in_range(doc.Content, field_mapping)
    except Exception:
        pass

    # 2. All header/footer stories (Primary=1, FirstPage=2, EvenPages=3)
    for hf_idx in [1, 2, 3]:
        for story_fn in [
            lambda idx=hf_idx: doc.Sections(1).Headers(idx).Range,
            lambda idx=hf_idx: doc.Sections(1).Footers(idx).Range,
        ]:
            try:
                _fill_content_controls_in_range(story_fn(), field_mapping)
            except Exception:
                pass

    # 3. Shapes with TextFrame, including GroupItems
    def _fill_shape(shape):
        try:
            if shape.TextFrame.HasText:
                _fill_content_controls_in_range(shape.TextFrame.TextRange, field_mapping)
        except Exception:
            pass
        try:
            for j in range(1, shape.GroupItems.Count + 1):
                try:
                    _fill_shape(shape.GroupItems(j))
                except Exception:
                    pass
        except Exception:
            pass

    try:
        for i in range(1, doc.Shapes.Count + 1):
            try:
                _fill_shape(doc.Shapes(i))
            except Exception:
                pass
    except Exception as e:
        print(f"    [WARNING] _fill_content_controls shapes: {e}")

    # 4. Shapes inside headers/footers (e.g. Facet, Motion, ViewMaster designs)
    try:
        section = doc.Sections(1)
        for hf_idx in [1, 2, 3]:
            for hf_fn in [
                lambda idx=hf_idx: section.Headers(idx),
                lambda idx=hf_idx: section.Footers(idx),
            ]:
                try:
                    hf = hf_fn()
                    for i in range(1, hf.Shapes.Count + 1):
                        try:
                            _fill_shape(hf.Shapes(i))
                        except Exception:
                            pass
                except Exception:
                    pass
    except Exception:
        pass


def _apply_header_insert_gallery(wd, param: Dict):
    expected_type = param.get('type', 5)  # wdTypeHeaders
    try:
        bb_entry = _find_building_block_entry_across_templates(wd, param['name'], expected_type)
        if not bb_entry:
            raise Exception(f"Header building block '{param['name']}' (Type {expected_type}) not found")
        section = wd.ActiveDocument.Sections(1)
        bb_entry.Insert(Where=section.Headers(1).Range, RichText=True)
        field_mapping = param.get('field_mapping')
        if field_mapping:
            _fill_content_controls(wd.ActiveDocument, field_mapping, include_shapes=True)
        positional_fill = param.get('positional_fill')
        if positional_fill:
            _fill_positional_content_controls(section.Headers(1).Range, positional_fill)
    except Exception as e:
        print(f"    [ERROR] Failed to insert header building block: {e}")
        raise


def _apply_footer_insert_gallery(wd, param: Dict):
    expected_type = param.get('type', 4)  # wdTypeFooters
    try:
        bb_entry = _find_building_block_entry_across_templates(wd, param['name'], expected_type)
        if not bb_entry:
            raise Exception(f"Footer building block '{param['name']}' (Type {expected_type}) not found")
        section = wd.ActiveDocument.Sections(1)
        # Even-page variants insert into the EvenPages footer slot (3); the
        # paired doc_in must already have DifferentOddEven enabled.
        slot = 3 if 'even_page' in param.get('label', '') else 1
        bb_entry.Insert(Where=section.Footers(slot).Range, RichText=True)
        field_mapping = param.get('field_mapping')
        if field_mapping:
            _fill_content_controls(wd.ActiveDocument, field_mapping, include_shapes=True)
        positional_fill = param.get('positional_fill')
        if positional_fill:
            _fill_positional_content_controls(section.Footers(slot).Range, positional_fill)
    except Exception as e:
        print(f"    [ERROR] Failed to insert footer building block: {e}")
        raise


def _apply_coverpage_insert_gallery(wd, param: Dict):
    expected_type = param.get('type', 2)  # wdTypeCoverPage
    bb_entry = _find_building_block_entry_across_templates(wd, param['name'], expected_type)
    if not bb_entry:
        raise Exception(f"Cover page building block '{param['name']}' (Type {expected_type}) not found")
    wd.ActiveWindow.View.SeekView = 0  # wdSeekMainDocument
    bb_entry.Insert(Where=wd.Selection.Range, RichText=True)
    time.sleep(1.0)  # Wait for cover page to render
    field_mapping = param.get('field_mapping')
    if field_mapping:
        _fill_content_controls(wd.ActiveDocument, field_mapping, include_shapes=True)


def _apply_wordart_by_index(wd, param: Dict):
    wd.ActiveDocument.Shapes.AddTextEffect(
        PresetTextEffect=param['index'],
        Text="Your Text Here",
        FontName="Arial",
        FontSize=36,
        FontBold=False,
        FontItalic=False,
        Left=100,
        Top=100
    )


def _apply_bibliography_gallery(wd, param: Dict):
    sel = wd.Selection
    try:
        bb_entry = _find_building_block_entry_across_templates(wd, param['name'])
        if not bb_entry:
            raise Exception(f"Bibliography building block '{param['name']}' not found")
        bb_entry.Insert(Where=sel.Range, RichText=True)
    except Exception:
        sel.Fields.Add(Range=sel.Range, Type=8)  # wdFieldBibliography = 8


def _apply_bibliography_style(wd, param: Dict):
    # doc_in already has a rendered bibliography in Chicago. Just change the
    # style and re-render the existing bibliography field.
    doc = wd.ActiveDocument
    doc.Bibliography.BibliographyStyle = param.get('style_name', param['name'])
    doc.Fields.Update()


def _apply_quick_parts_insert_gallery(wd, param: Dict):
    sel = wd.Selection
    template = _get_normal_template(wd)
    name_lower = param['name'].lower()
    # Case-insensitive search through BuildingBlocks collection
    bbs = template.BuildingBlockTypes(1).Categories("Built-In").BuildingBlocks
    bb = None
    for i in range(1, bbs.Count + 1):
        if bbs.Item(i).Name.lower() == name_lower:
            bb = bbs.Item(i)
            break
    if not bb:
        # Fallback: direct lookup
        bb = bbs(param['name'])
    bb.Insert(sel.Range, True)


def _apply_table_of_contents_gallery(wd, param: Dict):
    sel = wd.Selection
    try:
        bb_entry = _find_building_block_entry_across_templates(wd, param['name'])
        if not bb_entry:
            raise Exception(f"TOC building block '{param['name']}' not found")
        bb_entry.Insert(Where=sel.Range, RichText=True)
    except Exception:
        wd.ActiveDocument.TablesOfContents.Add(
            Range=sel.Range,
            UseHeadingStyles=True,
            UpperHeadingLevel=1,
            LowerHeadingLevel=3,
        )


def _update_document_fields_for_save(doc) -> None:
    """Refresh fields in body text and floating text frames before saving."""
    try:
        doc.Repaginate()
    except Exception:
        pass

    try:
        doc.Fields.Update()
    except Exception:
        pass

    def update_shape_fields(shapes) -> None:
        try:
            count = shapes.Count
        except Exception:
            return
        for idx in range(1, count + 1):
            try:
                shape = shapes(idx)
            except Exception:
                continue
            try:
                text_range = shape.TextFrame.TextRange
                text_range.Fields.Update()
            except Exception:
                pass
            try:
                group_items = shape.GroupItems
                for group_idx in range(1, group_items.Count + 1):
                    try:
                        group_shape = group_items(group_idx)
                        group_shape.TextFrame.TextRange.Fields.Update()
                    except Exception:
                        pass
            except Exception:
                pass

    update_shape_fields(doc.Shapes)

    try:
        sections_count = doc.Sections.Count
    except Exception:
        sections_count = 0
    for section_idx in range(1, sections_count + 1):
        try:
            section = doc.Sections(section_idx)
        except Exception:
            continue
        for collection_name in ("Headers", "Footers"):
            try:
                collection = getattr(section, collection_name)
            except Exception:
                continue
            for story_idx in range(1, 4):
                try:
                    story = collection(story_idx)
                    story.Range.Fields.Update()
                    update_shape_fields(story.Shapes)
                except Exception:
                    pass


def _apply_page_numbers_header_gallery(wd, param: Dict):
    try:
        bb_entry = _find_building_block_entry_across_templates(wd, param['name'], 11)  # Top of page
        if not bb_entry:
            raise Exception(f"Building block '{param['name']}' not found for headers")
        section = wd.ActiveDocument.Sections(1)
        bb_entry.Insert(Where=section.Headers(1).Range, RichText=True)
    except Exception:
        section = wd.ActiveDocument.Sections(1)
        header_range = section.Headers(1).Range
        header_range.Fields.Add(Range=header_range, Type=33)  # wdFieldPage


def _apply_page_numbers_footer_gallery(wd, param: Dict):
    try:
        bb_entry = _find_building_block_entry_across_templates(wd, param['name'], 12)  # Bottom of page
        if not bb_entry:
            raise Exception(f"Building block '{param['name']}' not found for footers")
        section = wd.ActiveDocument.Sections(1)
        bb_entry.Insert(Where=section.Footers(1).Range, RichText=True)
    except Exception as e:
        print(f"    [ERROR] Failed to insert footer building block: {e}")
        section = wd.ActiveDocument.Sections(1)
        footer_range = section.Footers(1).Range
        footer_range.Fields.Add(Range=footer_range, Type=33)  # wdFieldPage


def _apply_page_numbers_margins_gallery(wd, param: Dict):
    expected_type = param.get('type', 13)  # Page Numbers (Page Margins)
    try:
        bb_entry = _find_building_block_entry_across_templates(wd, param['name'], expected_type)
        if not bb_entry:
            raise Exception(f"Building block '{param['name']}' (Type {expected_type}) not found")
        bb_entry.Insert(Where=wd.Selection.Range, RichText=True)
    except Exception as e:
        print(f"    [ERROR] Failed to insert margin page number: {e}")
        raise


def _apply_page_number_field_gallery(wd, param: Dict):
    expected_type = param.get('type', 6)  # Current position page numbers are text boxes
    try:
        bb_entry = _find_building_block_entry_across_templates(wd, param['name'], expected_type)
        if not bb_entry:
            raise Exception(f"Building block '{param['name']}' (Type {expected_type}) not found")
        bb_entry.Insert(Where=wd.Selection.Range, RichText=True)
    except Exception as e:
        print(f"    [ERROR] Failed to insert page number field: {e}")
        raise


def _apply_toc_add_text_gallery(wd, param: Dict):
    doc = wd.ActiveDocument
    wd.Selection.HomeKey(Unit=6)  # wdStory
    wd.Selection.Find.ClearFormatting()
    found = wd.Selection.Find.Execute(
        FindText="Alpha TARGET Omega.",
        MatchCase=False,
        Forward=True,
        Wrap=0,  # wdFindStop
    )
    if found:
        level = param.get('level', 1)
        wd.Selection.Style = doc.Styles(f"Heading {level}")
        wd.Selection.Font.Size = 11
        try:
            doc.Fields.Update()
        except Exception:
            pass
        try:
            for idx in range(1, doc.TablesOfContents.Count + 1):
                doc.TablesOfContents(idx).Update()
        except Exception:
            pass


def _apply_gallery_all_shapes_and_canvas(wd, param: Dict):
    shape_type = param.get('shape_type', 1)  # Rectangle by default
    wd.ActiveDocument.Shapes.AddShape(
        Type=shape_type,
        Left=100,
        Top=100,
        Width=100,
        Height=80
    )


def _apply_citation_insert(wd, param: Dict):
    """Insert an in-text citation for an existing source."""
    source_tag = param.get('source_tag', 'Smith2020')
    sel = wd.Selection
    # No \l switch: pinning a locale would fix the rendered citation to one
    # language's conventions, which a solver cannot reproduce elsewhere.
    sel.Fields.Add(Range=sel.Range, Type=-1, Text=f'CITATION {source_tag}', PreserveFormatting=True)


def _apply_textbox_insert(wd, param: Dict):
    """Insert text box via Shapes.AddTextbox() COM API."""
    doc = wd.ActiveDocument
    doc.Shapes.AddTextbox(
        Orientation=param['orientation'],
        Left=param['left'],
        Top=param['top'],
        Width=param['width'],
        Height=param['height'],
    )


def _get_apply_handlers() -> Dict[str, Callable]:
    return {
        'SymbolInsertGallery': _apply_symbol_insert_gallery,
        'TableInsertGallery': _apply_table_insert_gallery,
        'TextBoxInsertGallery': _apply_named_building_block_from_normal_template,
        'HeaderInsertGallery': _apply_header_insert_gallery,
        'FooterInsertGallery': _apply_footer_insert_gallery,
        'CoverPageInsertGallery': _apply_coverpage_insert_gallery,
        'WordArtInsertGallery': _apply_wordart_by_index,
        'QuickTablesInsertGallery': _apply_named_building_block_from_normal_template,
        'AutoTextGallery': _apply_named_building_block_from_normal_template,
        'CitationInsert': _apply_citation_insert,
        'BibliographyGallery': _apply_bibliography_gallery,
        'BibliographyStyle': _apply_bibliography_style,
        'InsertBuildingBlocksEquationsGallery': _apply_named_building_block_from_normal_template,
        'QuickPartsInsertGallery': _apply_quick_parts_insert_gallery,
        'TableOfContentsGallery': _apply_table_of_contents_gallery,
        'WordArtInsertGalleryClassic': _apply_wordart_by_index,
        'PageNumbersInHeaderInsertGallery': _apply_page_numbers_header_gallery,
        'PageNambersInFooterInsertGallery': _apply_page_numbers_footer_gallery,
        'PageNambersInMarginsInsertGallery': _apply_page_numbers_margins_gallery,
        'PageNumberFieldInsertGallery': _apply_page_number_field_gallery,
        'TableOfContentsAddTextGallery': _apply_toc_add_text_gallery,
        'GalleryAllShapesAndCanvas': _apply_gallery_all_shapes_and_canvas,
        'TextBoxInsert': _apply_textbox_insert,
        'TextBoxInsertVerticalWord': _apply_textbox_insert,
    }


APPLY_HANDLERS = _get_apply_handlers()


def apply_parameter(wd, idMso: str, param: Dict) -> bool:
    """Apply parameter to insert content at cursor position."""
    try:
        try:
            wd.Templates.LoadBuildingBlocks()
        except Exception:
            pass  # Already loaded or unavailable

        handler = APPLY_HANDLERS.get(idMso)
        if handler is None:
            print(f"    [WARNING] Unknown command: {idMso}")
            return False

        handler(wd, param)
        return True

    except Exception as e:
        print(f"    [ERROR] Failed to apply parameter: {e}")
        import traceback
        traceback.print_exc()
        return False


def position_cursor_after_text(wd, text: str) -> bool:
    """
    Position cursor immediately after specified text.

    Args:
        wd: Word application instance
        text: Text to find and position cursor after

    Returns:
        True if cursor positioned successfully
    """
    # Go to document start
    wd.Selection.HomeKey(Unit=6)  # wdStory = 6

    # Find the text
    wd.Selection.Find.ClearFormatting()
    found = wd.Selection.Find.Execute(
        FindText=text,
        MatchCase=False,
        MatchWholeWord=False,
        Forward=True
    )

    if found:
        # Collapse selection to end (cursor after text)
        wd.Selection.Collapse(Direction=0)  # wdCollapseEnd = 0
        return True

    return False


class ContentOperationsCoreParameterizedGenerator:
    """
    Unified generator for content_operations Ã— core_parameterized test cases.

    Behavior:
    - content_operations
    Scope:
    - core_parameterized

    Strategy:
    - One test per generated variant label
    - Reuse precondition inputs by precondition type
    - Execute command with variant parameter and save gold/spec artifacts
    """

    def __init__(self, output_dir: Path = None):
        self.output_dir = output_dir or (BENCHMARK_OUTPUT_DIR / 'content_operations_x_core_parameterized')
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.inputs_dir = self.output_dir / 'inputs'
        self.input_paths = {}  # Cache for generated input files
        self.command_configs = get_command_configs()

    # Per-variation precondition override. Most variants use the command's default
    # precondition_type; entries here override for specific (idMso, variation) pairs.
    VARIATION_PRECONDITION_OVERRIDES = {
        ('FooterInsertGallery', 'facet_even_page'): 'with_sources_oddeven_2pages',
        ('FooterInsertGallery', 'facet_odd_page'): 'with_sources_oddeven_2pages',
    }

    def _effective_precondition_type(self, idMso: str, variation: Optional[str] = None) -> str:
        config = self.command_configs.get(idMso, CommandConfig(idMso=idMso))
        precondition_type = config.precondition_type
        if variation is not None:
            override = self.VARIATION_PRECONDITION_OVERRIDES.get((idMso, variation))
            if override is not None:
                precondition_type = override
        return precondition_type

    def generate_input_for_command(self, idMso: str, variation: Optional[str] = None) -> Path:
        """
        Generate input document with appropriate precondition for a command.

        Args:
            idMso: Command identifier
            variation: Optional variation label; allows per-variant precondition overrides
                       via VARIATION_PRECONDITION_OVERRIDES.

        Returns:
            Path to the generated input document
        """
        precondition_type = self._effective_precondition_type(idMso, variation)
        input_filename = f'doc_in_{precondition_type}.docx'
        input_path = self.inputs_dir / input_filename

        # Return cached path if already generated
        if precondition_type in self.input_paths:
            return self.input_paths[precondition_type]

        if input_path.exists():
            print(f"  Reusing existing input: {input_filename}")
            self.input_paths[precondition_type] = input_path
            return input_path

        print(f"  Generating input: {input_filename}")

        self.inputs_dir.mkdir(parents=True, exist_ok=True)
        wd = ensure_word()
        doc = None
        try:
            doc = wd.Documents.Add()

            if precondition_type == 'default':
                # Default: Simple text with cursor position
                wd.Selection.TypeText(DEFAULT_INPUT['full_text'])
                wd.Selection.WholeStory()
                wd.Selection.Font.Size = 11

            elif precondition_type == 'with_sources':
                # For bibliography-related commands: create a document with one citation source
                # and also include an in-text citation so style differences are visible.
                sources = doc.Bibliography.Sources

                # Add a single sample book source
                source_xml = '''<b:Source xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography">
                    <b:Tag>Smith2020</b:Tag>
                    <b:SourceType>Book</b:SourceType>
                    <b:Author>
                        <b:Author>
                            <b:NameList>
                                <b:Person>
                                    <b:Last>Smith</b:Last>
                                    <b:First>John</b:First>
                                </b:Person>
                            </b:NameList>
                        </b:Author>
                    </b:Author>
                    <b:Title>Sample Book Title</b:Title>
                    <b:Year>2020</b:Year>
                    <b:City>New York</b:City>
                    <b:Publisher>Sample Publisher</b:Publisher>
                </b:Source>'''
                sources.Add(source_xml)

                # Keep the document empty (except for citation sources).
                # Legacy behavior does not add body text or in-text citations for this precondition.

            elif precondition_type == 'with_sources_oddeven_2pages':
                # Same citation source as with_sources, plus a page break (so the
                # document spans two pages) and DifferentOddEven enabled — required
                # for header/footer gallery items meant for book layouts (Facet
                # Odd/Even Page, Motion Odd/Even Page, etc.).
                sources = doc.Bibliography.Sources
                source_xml = '''<b:Source xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography">
                    <b:Tag>Smith2020</b:Tag>
                    <b:SourceType>Book</b:SourceType>
                    <b:Author>
                        <b:Author>
                            <b:NameList>
                                <b:Person>
                                    <b:Last>Smith</b:Last>
                                    <b:First>John</b:First>
                                </b:Person>
                            </b:NameList>
                        </b:Author>
                    </b:Author>
                    <b:Title>Sample Book Title</b:Title>
                    <b:Year>2020</b:Year>
                    <b:City>New York</b:City>
                    <b:Publisher>Sample Publisher</b:Publisher>
                </b:Source>'''
                sources.Add(source_xml)
                wd.Selection.EndKey(Unit=6, Extend=0)  # wdStory
                wd.Selection.InsertBreak(Type=7)  # wdPageBreak
                doc.Sections(1).PageSetup.OddAndEvenPagesHeaderFooter = True

            elif precondition_type == 'with_bibliography':
                # For BibliographyStyle: doc has a source AND a rendered bibliography
                # already inserted in Chicago format. Test variants change the style to
                # something other than Chicago.
                #
                # Two-phase setup: building block lookup and Fields.Update only behave
                # correctly after the doc has been saved+reopened (so templates load).
                # Phase 1: add sources, set Chicago, save+close.
                # Phase 2: reopen, insert Bibliography building block, update fields, save.
                sources = doc.Bibliography.Sources
                source_xml = '''<b:Source xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography">
                    <b:Tag>Smith2020</b:Tag>
                    <b:SourceType>Book</b:SourceType>
                    <b:Author>
                        <b:Author>
                            <b:NameList>
                                <b:Person>
                                    <b:Last>Smith</b:Last>
                                    <b:First>John</b:First>
                                </b:Person>
                            </b:NameList>
                        </b:Author>
                    </b:Author>
                    <b:Title>Sample Book Title</b:Title>
                    <b:Year>2020</b:Year>
                    <b:City>New York</b:City>
                    <b:Publisher>Sample Publisher</b:Publisher>
                </b:Source>'''
                sources.Add(source_xml)
                doc.Bibliography.BibliographyStyle = 'Chicago'
                doc.SaveAs(str(input_path.absolute()), FileFormat=16)
                time.sleep(0.5)
                doc.Close(False)
                time.sleep(0.3)

                # Phase 2: reopen, insert bibliography, save again.
                doc = wd.Documents.Open(str(input_path.absolute()))
                time.sleep(0.5)
                bb_entry = _find_building_block_entry_across_templates(wd, 'Bibliography')
                if not bb_entry:
                    raise Exception('Bibliography building block not found in templates')
                bb_entry.Insert(Where=wd.Selection.Range, RichText=True)
                time.sleep(0.3)
                doc.Fields.Update()
                time.sleep(0.3)

            elif precondition_type in ('with_headings', 'with_existing_toc'):
                # For TableOfContentsGallery:
                # - Keep the first page for TOC insertion.
                # - Put headings on the second page so the TOC is non-empty after insertion.
                # For TableOfContentsAddTextGallery with_existing_toc:
                # - Pre-create a TOC on the first page.
                # - Keep "Alpha TARGET Omega." as normal text initially.
                # - The command marks that existing paragraph for TOC inclusion.

                # Create an empty first page.
                wd.Selection.TypeParagraph()
                wd.Selection.InsertBreak(7)  # wdPageBreak = 7

                # Add headings on page 2.
                wd.Selection.Style = wd.ActiveDocument.Styles("Heading 1")
                wd.Selection.TypeText("Section 1")
                wd.Selection.TypeParagraph()

                wd.Selection.Style = wd.ActiveDocument.Styles("Heading 2")
                wd.Selection.TypeText("Subsection 1.1")
                wd.Selection.TypeParagraph()

                wd.Selection.Style = wd.ActiveDocument.Styles("Normal")
                wd.Selection.TypeText("Alpha TARGET Omega.")
                wd.Selection.TypeParagraph()

                wd.Selection.Style = wd.ActiveDocument.Styles("Heading 1")
                wd.Selection.TypeText("Section 2")
                wd.Selection.TypeParagraph()

                wd.Selection.Style = wd.ActiveDocument.Styles("Normal")
                wd.Selection.TypeText("More body text.")
                wd.Selection.TypeParagraph()

                # Move cursor back to page 1 so TOC gets inserted there.
                wd.Selection.HomeKey(Unit=6)  # wdStory = 6

                wd.Selection.WholeStory()
                wd.Selection.Font.Size = 11

                if precondition_type == 'with_existing_toc':
                    wd.Selection.HomeKey(Unit=6)  # wdStory = 6
                    wd.ActiveDocument.TablesOfContents.Add(
                        Range=wd.Selection.Range,
                        UseHeadingStyles=True,
                        UpperHeadingLevel=1,
                        LowerHeadingLevel=3,
                    )
                    try:
                        wd.ActiveDocument.Fields.Update()
                    except Exception:
                        pass
                    wd.Selection.HomeKey(Unit=6)  # wdStory = 6

            doc.SaveAs(str(input_path.absolute()), FileFormat=16)
            time.sleep(1.0)
            #doc.Close(False)
            #normalize_docx_package_identity(input_path)
            #time.sleep(0.5)

            self.input_paths[precondition_type] = input_path
            print(f"    Created: {input_filename}")
            return input_path

        except Exception as e:
            print(f"    Error creating input: {e}")
            if doc:
                try:
                    doc.Close(False)
                except:
                    pass
            raise

    def load_commands(self) -> List[Dict]:
        """Load commands classified as content_operations Ã— core_parameterized."""
        commands_file = COMMAND_SETS_DIR / 'commands_by_policy_id.json'

        if not commands_file.exists():
            print(f"Warning: {commands_file} not found. Using empty command list.")
            return []

        with open(commands_file, 'r', encoding='utf-8') as f:
            all_commands = json.load(f)

        # Filter for our test set
        filtered = [
            cmd for cmd in all_commands
            if cmd.get('behavior') == 'content_operations'
            and cmd.get('scope') == 'core_parameterized'
        ]

        return filtered

    def _write_spec_file(
        self,
        spec_path: Path,
        idMso: str,
        param: Dict,
        input_path: Path,
        config: CommandConfig,
        skipped: bool = False,
        skip_reason: Optional[str] = None,
        failed: bool = False,
        failure_reason: Optional[str] = None,
    ) -> Dict:
        """Write spec.json with optional skipped/failed annotations."""
        spec = self._generate_spec(idMso, param, input_path, config)
        if skipped:
            spec['skipped'] = True
            if skip_reason:
                spec['skip_reason'] = skip_reason
        if failed:
            spec['failed'] = True
            if failure_reason:
                spec['failure_reason'] = failure_reason
        write_spec_if_changed(spec_path, spec, sort_keys=False)
        return spec

    def generate_parameter_test(self, idMso: str, param: Dict, input_path: Path, config: CommandConfig) -> Dict:
        """
        Generate test for a specific parameter variant.

        - doc_in: input with appropriate precondition
        - gold: result after applying parameter
        """
        test_dir = self.output_dir / idMso / param['label']
        test_dir.mkdir(parents=True, exist_ok=True)

        gold_path = test_dir / 'gold.docx'
        spec_path = test_dir / 'spec.json'

        result = {
            'idMso': idMso,
            'parameter': param['label'],
            'success': False,
            'message': '',
            'paths': {
                'doc_in': str(input_path),
                'gold': str(gold_path),
                'spec': str(spec_path),
            }
        }

        # Check for skip commands
        skip_reason = config.skip_reason
        if skip_reason:
            result['message'] = f'Skipped: {skip_reason}'
            result['skipped'] = True
            # Still generate spec for documentation
            self._write_spec_file(
                spec_path=spec_path,
                idMso=idMso,
                param=param,
                input_path=input_path,
                config=config,
                skipped=True,
                skip_reason=skip_reason,
            )
            return result

        # If field_mapping exists, prepare a temp copy of doc_in with docProps pre-filled
        # so Word auto-populates content controls bound to document properties.
        field_mapping = param.get('field_mapping')
        actual_input_path = input_path
        tmp_input_path = None
        if field_mapping:
            tmp_input_path = test_dir / '_doc_in_tmp.docx'
            shutil.copy2(input_path, tmp_input_path)
            normalize_docx_package_identity(tmp_input_path, field_mapping=field_mapping)
            actual_input_path = tmp_input_path

        wd = ensure_word()
        try:
            # Open input (with docProps pre-filled if field_mapping exists)
            doc = wd.Documents.Open(str(actual_input_path.absolute()))

            # Position cursor after target text only for commands that need generic cursor setup.
            if config.needs_cursor_positioning:
                cursor_text = DEFAULT_INPUT.get('cursor_after', 'TARGET')
                if not position_cursor_after_text(wd, cursor_text):
                    result['message'] = f"Failed to position cursor after: {cursor_text}"
                    result['failed'] = True
                    doc.Close(False)

                    # Generate spec.json even for cursor positioning failures
                    self._write_spec_file(
                        spec_path=spec_path,
                        idMso=idMso,
                        param=param,
                        input_path=input_path,
                        config=config,
                        failed=True,
                        failure_reason=result['message'],
                    )

                    return result

            # Apply the parameter
            if not apply_parameter(wd, idMso, param):
                result['message'] = f'Failed to apply parameter: {param["label"]}'
                result['failed'] = True
                doc.Close(False)

                # Generate spec.json even for failed cases
                self._write_spec_file(
                    spec_path=spec_path,
                    idMso=idMso,
                    param=param,
                    input_path=input_path,
                    config=config,
                    failed=True,
                    failure_reason=result['message'],
                )

                return result

            # Wait for operation to complete
            time.sleep(0.3)

            # Refresh fields before saving (prevents stale TOC/page numbers)
            _update_document_fields_for_save(doc)

            # Save as gold
            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
            doc.Close(False)
            time.sleep(0.2)

            # Generate spec.json
            self._write_spec_file(
                spec_path=spec_path,
                idMso=idMso,
                param=param,
                input_path=input_path,
                config=config,
            )

            result['success'] = True
            result['message'] = 'Test case generated successfully'

        except Exception as e:
            result['message'] = f'Error: {str(e)}'
            result['failed'] = True

            # Generate spec.json even for exceptions
            try:
                self._write_spec_file(
                    spec_path=spec_path,
                    idMso=idMso,
                    param=param,
                    input_path=input_path,
                    config=config,
                    failed=True,
                    failure_reason=result['message'],
                )
            except:
                pass  # If spec generation also fails, just skip it

            import traceback
            traceback.print_exc()

        finally:
            # Clean up temp input copy
            if tmp_input_path and tmp_input_path.exists():
                try:
                    tmp_input_path.unlink()
                except Exception:
                    pass

        return result

    def _generate_spec(self, idMso: str, param: Dict, input_path: Path = None, config: Optional[CommandConfig] = None) -> Dict:
        """Generate the spec.json for a test case."""

        cursor_text = DEFAULT_INPUT.get('cursor_after', 'TARGET')
        config = config or self.command_configs.get(idMso, CommandConfig(idMso=idMso))
        precondition_type = self._effective_precondition_type(idMso, param.get('label'))

        # Build precondition based on type
        if precondition_type in {'with_sources', 'with_sources_oddeven_2pages'}:
            is_oddeven_source_doc = precondition_type == 'with_sources_oddeven_2pages'
            precondition = {
                'type': precondition_type,
                'template': (
                    'empty_with_sources_oddeven_2pages'
                    if is_oddeven_source_doc
                    else 'empty_with_sources'
                ),
                'state': 'empty',
                'sources_count': 1,
                'source_info': 'Smith, J. (2020). Sample Book Title. New York: Sample Publisher.',
                'description': (
                    'Two-page empty document with odd/even headers and footers enabled, '
                    'plus one citation source (Smith 2020 book)'
                    if is_oddeven_source_doc
                    else 'Empty document with one citation source (Smith 2020 book)'
                ),
            }
            if is_oddeven_source_doc:
                precondition.update({
                    'page_count': 2,
                    'odd_and_even_pages_header_footer': True,
                })
        elif precondition_type == 'with_bibliography':
            precondition = {
                'type': precondition_type,
                'template': 'with_bibliography',
                'state': 'has_bibliography',
                'starting_style': 'Chicago',
                'sources_count': 1,
                'source_info': 'Smith, J. (2020). Sample Book Title. New York: Sample Publisher.',
                'description': 'Document with one citation source and a Bibliography rendered in Chicago format',
            }
        else:
            precondition = {
                'type': precondition_type,
                'template': 'simple_text',
                'full_text': DEFAULT_INPUT['full_text'],
                'cursor_after': cursor_text,
                'state': 'empty',
            }

        primary_spec = {
            'method': config.primary_method,
            'description': config.description,
        }
        if config.primary_params:
            primary_spec.update(config.primary_params)
        verification_spec = {
            'primary': primary_spec,
        }

        spec = {
            'idMso': idMso,
            'behavior': 'content_operations',
            'scope': 'core_parameterized',
            'variation': f"apply_{param['label']}",
            'parameter': param,
            'generated_at': datetime.now().isoformat(),
            'precondition': precondition,

            'verification': verification_spec,

            'files': {
                'doc_in': f'../../inputs/{input_path.name}' if input_path else '../../inputs/doc_in.docx',
                'gold': 'gold.docx',
            }
        }
        preview_path = self.output_dir / idMso / param['label'] / 'preview.png'
        if preview_path.exists():
            spec['files']['preview'] = 'preview.png'

        # Variant-specific verification overrides
        # Manual TOC uses Content Controls (not a TOC field), so toc_verification won't detect it.
        if idMso == 'TableOfContentsGallery' and param.get('label') == 'manual_table':
            spec['verification']['primary'] = {
                'method': 'content_control_verification',
                'description': 'Word COM content control verification (manual TOC is a Content Control)',
            }

        return spec

    def generate_for_command(self, command: Dict, input_path: Path) -> List[Dict]:
        """
        Generate all parameter test variations for a single command.

        Args:
            command: Command dict with 'idMso' and 'control_type'
            input_path: Path to input document with appropriate precondition

        Returns:
            List of generation results
        """
        idMso = command['idMso']
        config = self.command_configs.get(idMso, CommandConfig(idMso=idMso))

        # Extract parameters from variant definitions
        wd = ensure_word()
        parameters = extract_parameters(wd, idMso)

        if not parameters:
            print(f"  No parameters found for {idMso}")
            return []

        results = []
        for param in parameters:
            # Per-variation input override: most variants reuse the command's default
            # precondition, but a few (e.g. Footer odd/even page) need a custom doc_in.
            variant_input_path = self.generate_input_for_command(idMso, param.get('label', ''))
            result = self.generate_parameter_test(idMso, param, variant_input_path, config)
            results.append(result)

            if result.get('skipped'):
                status = 'SKIP'
            elif result['success']:
                status = 'OK'
            else:
                status = 'FAIL'
            print(f"  {param['label']}: {status} - {result['message']}")

        return results

    def generate_all(self, commands: List[Dict] = None, limit: int = None) -> Dict:
        """
        Generate test cases for all commands in this test set.

        Args:
            commands: List of command dicts (loads from file if None)
            limit: Optional limit on number of commands to process

        Returns:
            Summary dict with statistics
        """
        if commands is None:
            commands = self.load_commands()

        if limit:
            commands = commands[:limit]

        print(f"\n=== Generating test cases for {len(commands)} commands ===")
        print(f"Test set: content_operations core_parameterized")
        print(f"Output: {self.output_dir}\n")

        summary = {
            'test_set': 'content_operations_x_core_parameterized',
            'generated_at': datetime.now().isoformat(),
            'total_commands': len(commands),
            'total_test_cases': 0,
            'successful': 0,
            'failed': 0,
            'skipped': 0,
            'results': [],
        }

        wd = ensure_word()

        # Generate tests for each command
        print("Generating test cases...")
        for cmd in commands:
            idMso = cmd['idMso']
            control_type = cmd.get('control_type', 'gallery')

            print(f"\n{idMso} ({control_type})...")

            # Generate input document for this command
            input_path = self.generate_input_for_command(idMso)

            results = self.generate_for_command(cmd, input_path)

            for r in results:
                summary['total_test_cases'] += 1
                if r.get('skipped'):
                    summary['skipped'] += 1
                elif r['success']:
                    summary['successful'] += 1
                else:
                    summary['failed'] += 1
                summary['results'].append(r)

        # Save summary
        summary_path = self.output_dir / 'generation_summary.json'
        with open(summary_path, 'w', encoding='utf-8') as f:
            json.dump(summary, f, indent=2, ensure_ascii=False)

        print(f"\n=== Generation Complete ===")
        print(f"Total commands: {summary['total_commands']}")
        print(f"Total test cases: {summary['total_test_cases']}")
        print(f"Successful: {summary['successful']}")
        print(f"Skipped: {summary['skipped']}")
        print(f"Failed: {summary['failed']}")
        print(f"Summary saved to: {summary_path}")

        return summary


def main():
    """CLI entry point for generating content_operations Ã— core_parameterized tests."""
    import argparse

    parser = argparse.ArgumentParser(
        description='Generate test cases for content_operations Ã— core_parameterized'
    )
    parser.add_argument(
        '--command', '-c',
        help='Generate for a specific command (e.g., SymbolInsertGallery)'
    )
    parser.add_argument(
        '--limit', '-l',
        type=int,
        help='Limit number of commands to process'
    )
    parser.add_argument(
        '--output', '-o',
        help='Output directory'
    )

    args = parser.parse_args()

    output_dir = Path(args.output) if args.output else None
    generator = ContentOperationsCoreParameterizedGenerator(output_dir=output_dir)

    # Generate using generate_all (handles both single command and all commands)
    if args.command:
        # Filter to single command
        commands = generator.load_commands()
        cmd = next((c for c in commands if c['idMso'] == args.command), None)
        if cmd:
            generator.generate_all(commands=[cmd])
        else:
            # Command not in filtered set, assume it's a gallery
            print(f"Command {args.command} not found in content_operations Ã— core_parameterized set.")
            print(f"Generating as gallery...")
            cmd = {'idMso': args.command, 'control_type': 'gallery'}
            generator.generate_all(commands=[cmd])
    else:
        # Generate for all commands
        generator.generate_all(limit=args.limit)


if __name__ == '__main__':
    main()
