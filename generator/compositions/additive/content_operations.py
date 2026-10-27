"""
Additive composition generator for content_operations behavior.

Footnote + Endnote composition: inserts both on same document
Header + Footer composition: sets both text header and footer

Usage:
    python -m generator.compositions.additive.content_operations
"""

import json
import shutil
import time
from datetime import datetime
from pathlib import Path
from typing import Dict

from generator.config import BENCHMARK_OUTPUT_DIR
from generator.compositions.atomic_intents import build_content_operations_atomic_intents
from generator.compositions.interaction.shared import write_instructions_json
from generator.execution.setup import ensure_word
from generator.test_sets.spec_utils import write_spec_if_changed
from generator.test_sets.content_operations_core_parameterized import (
    apply_parameter,
    extract_parameters,
)

TEST_SET = 'additive_composition_x_content_operations'


# ============================================================================
# Variant definitions
# ============================================================================

# Legacy 2-op compositions removed — subsumed by AdditiveDocumentEnrichmentComposition V3
# which covers footnote + endnote + header + footer (4 ops) in a more demanding context.
FOOTNOTE_ENDNOTE_VARIANTS = {}
HEADER_FOOTER_VARIANTS = {}


# -- Document Enrichment composition (§2) --

DOCUMENT_ENRICHMENT_VARIANTS = {
    'variant_1': {
        'description': 'Footnote + endnote + header + footer + citation + bibliography (6 ops)',
        'instruction': "Insert a footnote after the word 'TARGET' with the text 'This is a footnote.' and an endnote with the text 'This is an endnote.' at the same position, add a header in the ViewMaster layout with Title = \"Quarterly Business Review\", add a footer in the Ion (Light) layout with Title = \"Quarterly Business Review\" and Author = \"Megan Bowen\", insert a citation for the Smith2020 source in a new paragraph at the very beginning of the document body, and add a bibliography section at the end of the document using the existing citation sources.",
        'footnote_text': 'This is a footnote.',
        'endnote_text': 'This is an endnote.',
        'header_type': 'gallery',
        'header_gallery': 'viewmaster',
        'footer_type': 'gallery',
        'footer_gallery': 'ion_light',
        'source_tag': 'Smith2020',
    },
    'variant_2': {
        'description': 'Footnote + endnote + page num header/footer + citation + bib (6 ops)',
        'instruction': "Insert a footnote after the word 'TARGET' with the text 'This is a footnote.' and an endnote with 'This is an endnote.', add a page number in the header and a page number in the footer both using the 'Plain Number 1' style (\"Number with no formatting or accents\"), insert a citation for the Smith2020 source in a new paragraph at the very beginning of the document body, and add a bibliography section at the end of the document using the existing citation sources.",
        'footnote_text': 'This is a footnote.',
        'endnote_text': 'This is an endnote.',
        'header_type': 'page_number',
        'header_page_number_gallery': 'plain_number_1',
        'footer_type': 'page_number',
        'footer_page_number_gallery': 'plain_number_1',
        'source_tag': 'Smith2020',
    },
    'variant_3': {
        'description': 'Footnote + endnote + header + footer (4 ops)',
        'instruction': "Insert a footnote after the word 'TARGET' with the text 'This is a footnote.' and an endnote with 'This is an endnote.', add a header in the ViewMaster layout with Title = \"Quarterly Business Review\", and add a footer in the Ion (Light) layout with Title = \"Quarterly Business Review\" and Author = \"Megan Bowen\".",
        'footnote_text': 'This is a footnote.',
        'endnote_text': 'This is an endnote.',
        'header_type': 'gallery',
        'header_gallery': 'viewmaster',
        'footer_type': 'gallery',
        'footer_gallery': 'ion_light',
    },
}

BIBLIOGRAPHY_SOURCE_XML = '''<b:Source xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography">
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

COMMENT_AUTHOR = 'TestUser'

CORE_PARAMETER_KEYS = [
    'char', 'code', 'name', 'style_name', 'position', 'description',
    'rows', 'cols', 'type', 'index', 'field', 'level', 'shape_type',
    'category', 'field_mapping', 'layout_description', 'source_tag',
    'positional_fill', 'orientation', 'left', 'top', 'width', 'height',
    'diagonal_type',
]


def _set_built_in_document_property(doc, name, value):
    """Set a built-in property across early- and late-bound COM dispatch."""
    properties = doc.BuiltInDocumentProperties
    try:
        prop = properties.Item(name)
    except (AttributeError, TypeError):
        try:
            prop = properties(name)
        except TypeError:
            # Some Office 16 type libraries expose DocumentProperties without
            # the default Item/Value members in the generated Python wrapper.
            # Invoke DISPID 0 directly for both the collection lookup and the
            # DocumentProperty's default Value setter.
            raw_prop = properties._oleobj_.Invoke(0, 0, 2, 1, name)
            raw_prop.Invoke(0, 0, 4, 0, value)
            return
    prop.Value = value

# -- Document Cleanup composition (§3) --

DOCUMENT_CLEANUP_VARIANTS = {
    'variant_1': {
        'description': 'Accept changes + delete comments + remove watermark + remove footer (4 ops)',
        'instruction': "Accept all tracked changes in the document, delete all comments, remove the watermark, and clear the footer.",
        'precondition': 'track_changes_comments_watermark_footer',
        'operations': ['accept_changes', 'delete_comments', 'remove_watermark', 'remove_footer'],
    },
    'variant_2': {
        'description': 'Reject changes + delete resolved comments + remove header + remove footer (4 ops)',
        'instruction': "Reject all tracked changes in the document, delete all resolved comments, and clear both the header and footer.",
        'precondition': 'track_changes_resolved_comments_header_footer',
        'operations': ['reject_changes', 'delete_resolved_comments', 'remove_header', 'remove_footer'],
    },
}


def _core_parameter(idMso: str, label: str) -> Dict:
    """Return a core-parameterized variant payload matching atomic generation."""
    for param in extract_parameters(None, idMso):
        if param.get('label') == label:
            return {
                key: value
                for key, value in param.items()
                if key == 'label' or key in CORE_PARAMETER_KEYS
            }
    raise KeyError(f"Core parameterized variant not found: {idMso}/{label}")


def _apply_core_parameter(wd, idMso: str, label: str):
    """Apply the same gallery/action helper used by the atomic benchmark."""
    param = _core_parameter(idMso, label)
    if not apply_parameter(wd, idMso, param):
        raise RuntimeError(f"Failed to apply {idMso}/{label}")


def _enrichment_field_mapping(config):
    """Return the document-property fields the chosen galleries fill.

    The atomic parameterized generator records this on its own specs, which is
    what tells the export sanitizer that a creator/company value is semantic
    rather than incidental account metadata. Compositions insert the same
    building blocks, so they need to declare it too.
    """
    mapping = {}
    for idMso, label in (
        ('HeaderInsertGallery', config.get('header_gallery')),
        ('FooterInsertGallery', config.get('footer_gallery')),
    ):
        if not label:
            continue
        mapping.update(_core_parameter(idMso, label).get('field_mapping') or {})
    return mapping


def _record_note_text(intents, config):
    """Record which sentence each note intent asked for.

    FootnoteInsert and EndnoteInsertWord both type "See reference." in their
    atomic form, while these variants ask for "This is a footnote." and "This
    is an endnote.". The sentence is a parameter of the operation rather than
    part of it, and distinct sentences are what let the reference catch a
    solver that swaps the two notes. Recording the binding keeps the intent
    honest instead of implying the atomic's wording carried over unchanged.
    """
    payloads = {'footnote': config.get('footnote_text'),
                'endnote': config.get('endnote_text')}
    for intent in intents:
        text = payloads.get(intent.get('operation'))
        if text:
            intent['parameter'] = {'note_text': text}
    return intents


def _enrichment_atomic_operations(config):
    """Return ordered atomic operation descriptors for document enrichment."""
    operations = [
        ('footnote', 'default'),
        ('endnote', 'default'),
    ]

    header_type = config.get('header_type')
    if header_type == 'gallery':
        operations.append(('header_gallery', config.get('header_gallery')))
    elif header_type == 'page_number':
        operations.append(('page_number_header', config.get('header_page_number_gallery')))

    footer_type = config.get('footer_type')
    if footer_type == 'gallery':
        operations.append(('footer_gallery', config.get('footer_gallery')))
    elif footer_type == 'page_number':
        operations.append(('page_number_footer', config.get('footer_page_number_gallery')))

    if config.get('source_tag'):
        operations.append(('citation', config['source_tag']))
        operations.append(('bibliography', 'insert'))

    return operations


def _cleanup_atomic_operations(config):
    """Return ordered atomic operation descriptors for document cleanup."""
    return [(operation, 'apply') for operation in config.get('operations', [])]


# ============================================================================
# Generator class
# ============================================================================

class AdditiveContentOperationsGenerator:
    """Generator for content_operations x additive_composition test cases."""

    def __init__(self, output_dir: Path = None, hidden: bool = False):
        self.output_dir = output_dir or (BENCHMARK_OUTPUT_DIR / TEST_SET)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.inputs_dir = self.output_dir / 'inputs'
        self.inputs_dir.mkdir(parents=True, exist_ok=True)
        self.hidden = hidden

    def _get_shared_input(self, input_name):
        """Get path to shared input file."""
        dest = self.inputs_dir / f'{input_name}.docx'
        if dest.exists():
            return dest

        search_dirs = [
            BENCHMARK_OUTPUT_DIR / 'content_operations_x_core_actionable' / 'inputs',
            BENCHMARK_OUTPUT_DIR / 'content_operations_x_core_parameterized' / 'inputs',
        ]
        for src_dir in search_dirs:
            src = src_dir / f'{input_name}.docx'
            if src.exists():
                shutil.copy2(src, dest)
                print(f"  Copied input: {input_name}.docx")
                return dest

        raise FileNotFoundError(
            f"Shared input '{input_name}.docx' not found in any benchmark inputs directory"
        )

    def _generate_footnote_endnote_variant(self, wd, case_name, variant_name, config):
        """Generate footnote + endnote composition variant."""
        variant_dir = self.output_dir / case_name / variant_name
        variant_dir.mkdir(parents=True, exist_ok=True)
        gold_path = variant_dir / 'gold.docx'
        spec_path = variant_dir / 'spec.json'

        shared_input = self._get_shared_input('doc_in')

        doc = wd.Documents.Open(str(shared_input.absolute()))
        try:
            # Find "TARGET" and position after it
            rng = doc.Content
            rng.Find.ClearFormatting()
            rng.Find.Execute(FindText='TARGET', MatchCase=False, Forward=True)
            if not rng.Find.Found:
                raise RuntimeError("TARGET not found in document")

            # Insert footnote at end of TARGET (two-step: add, then set text)
            fn_range = doc.Range(rng.End, rng.End)
            fn = doc.Footnotes.Add(Range=fn_range)
            time.sleep(0.3)
            fn.Range.Text = config['footnote_text']
            time.sleep(0.2)

            # Navigate back to body before inserting endnote
            wd.Selection.HomeKey(Unit=6)  # wdStory
            time.sleep(0.2)

            # Re-find TARGET (footnote insertion shifted positions)
            rng2 = doc.Content
            rng2.Find.ClearFormatting()
            rng2.Find.Execute(FindText='TARGET', MatchCase=False, Forward=True)
            if not rng2.Find.Found:
                raise RuntimeError("TARGET not found after footnote insertion")

            # Insert endnote at end of TARGET
            en_range = doc.Range(rng2.End, rng2.End)
            en = doc.Endnotes.Add(Range=en_range)
            time.sleep(0.3)
            en.Range.Text = config['endnote_text']
            time.sleep(0.2)

            # Navigate back to body
            wd.Selection.HomeKey(Unit=6)

            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
        finally:
            doc.Close(False)
        time.sleep(0.3)

        atomic_intents = build_content_operations_atomic_intents([
            ('footnote', 'default'),
            ('endnote', 'default'),
        ])
        spec = {
            'behavior': 'content_operations',
            'scope': 'additive_composition',
            'variation': variant_name,
            'generated_at': datetime.now().isoformat(),
            'composition': {
                'type': 'additive',
                'atomic_intents': atomic_intents,
                'operation_count': 2,
            },
            'precondition': {
                'input_state': 'blank',
                'cursor_handling': 'position_after_text',
                'full_text': 'Alpha TARGET Omega.',
                'cursor_after': 'TARGET',
            },
            'verification': {
                'primary': {
                    'method': 'footnote_endnote_composition_verification',
                    'description': 'Verify both footnote and endnote present and correct',
                },
            },
            'files': {
                'doc_in': '../../inputs/doc_in.docx',
                'gold': 'gold.docx',
            },
        }
        write_spec_if_changed(spec_path, spec)
        print(f"  [{case_name}/{variant_name}] Generated (2 operations)")
        return True

    def _generate_header_footer_variant(self, wd, case_name, variant_name, config):
        """Generate header + footer composition variant."""
        variant_dir = self.output_dir / case_name / variant_name
        variant_dir.mkdir(parents=True, exist_ok=True)
        gold_path = variant_dir / 'gold.docx'
        spec_path = variant_dir / 'spec.json'

        shared_input = self._get_shared_input('doc_in')

        doc = wd.Documents.Open(str(shared_input.absolute()))
        try:
            section = doc.Sections(1)

            # Set header text (wdHeaderFooterPrimary = 1)
            header = section.Headers(1)
            header.Range.Text = config['header_text']

            # Set footer text (wdHeaderFooterPrimary = 1)
            footer = section.Footers(1)
            footer.Range.Text = config['footer_text']

            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
        finally:
            doc.Close(False)
        time.sleep(0.3)

        atomic_intents = build_content_operations_atomic_intents([
            ('header_gallery', 'viewmaster'),
            ('footer_gallery', 'ion_light'),
        ])
        spec = {
            'behavior': 'content_operations',
            'scope': 'additive_composition',
            'variation': variant_name,
            'generated_at': datetime.now().isoformat(),
            'composition': {
                'type': 'additive',
                'atomic_intents': atomic_intents,
                'operation_count': 2,
            },
            'precondition': {
                'input_state': 'blank',
                'full_text': 'Alpha TARGET Omega.',
            },
            'verification': {
                'primary': {
                    'method': 'header_footer_composition_verification',
                    'description': 'Verify both header and footer content match gold',
                },
            },
            'files': {
                'doc_in': '../../inputs/doc_in.docx',
                'gold': 'gold.docx',
            },
        }
        write_spec_if_changed(spec_path, spec)
        print(f"  [{case_name}/{variant_name}] Generated (2 operations)")
        return True

    # ==== New composition generators (additive extension plan) ====

    def _create_enrichment_doc_in(self, wd, doc_in_path, config):
        """Create custom doc_in for document enrichment: body text + bibliography source."""
        doc = wd.Documents.Add()
        try:
            doc.Content.Text = 'Alpha TARGET Omega.'
            # Add bibliography source if needed
            if config.get('source_tag'):
                doc.Bibliography.Sources.Add(BIBLIOGRAPHY_SOURCE_XML)
            # The footer Author control these variants ask for is data-bound to
            # the Author document property, so the input must not arrive with a
            # meaningful author: that would satisfy the requested edit before
            # the solver does anything. It also keeps the generating account's
            # real name out of the input.
            _set_built_in_document_property(
                doc, 'Author', 'Anonymous Author')
            _set_built_in_document_property(doc, 'Title', '')
            doc.SaveAs(str(doc_in_path.absolute()), FileFormat=16)
        finally:
            doc.Close(False)
        time.sleep(0.3)

    def _generate_document_enrichment_variant(self, wd, case_name, variant_name, config):
        """Generate a document enrichment composition variant."""
        variant_dir = self.output_dir / case_name / variant_name
        variant_dir.mkdir(parents=True, exist_ok=True)
        gold_path = variant_dir / 'gold.docx'
        spec_path = variant_dir / 'spec.json'
        doc_in_path = variant_dir / 'doc_in.docx'

        # Create custom doc_in
        self._create_enrichment_doc_in(wd, doc_in_path, config)

        # Open doc_in and apply all operations to create gold
        doc = wd.Documents.Open(str(doc_in_path.absolute()))
        try:
            # 1. Insert footnote after TARGET
            rng = doc.Content
            rng.Find.ClearFormatting()
            rng.Find.Execute(FindText='TARGET', MatchCase=False, Forward=True)
            if not rng.Find.Found:
                raise RuntimeError("TARGET not found")
            fn_range = doc.Range(rng.End, rng.End)
            fn = doc.Footnotes.Add(Range=fn_range)
            time.sleep(0.3)
            fn.Range.Text = config['footnote_text']
            time.sleep(0.2)

            # 2. Insert endnote after TARGET
            wd.Selection.HomeKey(Unit=6)
            time.sleep(0.2)
            rng2 = doc.Content
            rng2.Find.ClearFormatting()
            rng2.Find.Execute(FindText='TARGET', MatchCase=False, Forward=True)
            if not rng2.Find.Found:
                raise RuntimeError("TARGET not found after footnote insertion")
            en_range = doc.Range(rng2.End, rng2.End)
            en = doc.Endnotes.Add(Range=en_range)
            time.sleep(0.3)
            en.Range.Text = config['endnote_text']
            time.sleep(0.2)
            # Set endnote numbering to Roman numerals (i, ii, iii) to distinguish
            # from footnotes (1, 2, 3) — both appear at page bottom in 1-page docs
            doc.Endnotes.NumberStyle = 2  # wdNoteNumberStyleLowercaseRoman
            wd.Selection.HomeKey(Unit=6)

            # 3. Set header
            if config.get('header_type') == 'gallery':
                _apply_core_parameter(
                    wd, 'HeaderInsertGallery', config['header_gallery']
                )
            elif config.get('header_type') == 'page_number':
                _apply_core_parameter(
                    wd,
                    'PageNumbersInHeaderInsertGallery',
                    config['header_page_number_gallery'],
                )

            # 4. Set footer
            if config.get('footer_type') == 'gallery':
                _apply_core_parameter(
                    wd, 'FooterInsertGallery', config['footer_gallery']
                )
            elif config.get('footer_type') == 'page_number':
                _apply_core_parameter(
                    wd,
                    'PageNambersInFooterInsertGallery',
                    config['footer_page_number_gallery'],
                )

            # 5. Insert citation (if source_tag present)
            #
            # The citation goes in its own paragraph at the start of the body,
            # matching the linked CitationInsert atomic contract. Anchoring it
            # to the start rather than the end keeps it clear of the footnote
            # and endnote reference marks after 'TARGET' and of the
            # bibliography at the end, so the composition stays order
            # independent as its additive classification requires.
            if config.get('source_tag'):
                doc.Paragraphs(1).Range.InsertParagraphBefore()
                time.sleep(0.2)
                citation_range = doc.Paragraphs(1).Range
                citation_range.Collapse(1)  # wdCollapseStart
                # No \l switch: pinning a locale would fix the rendered
                # citation to one language's conventions, which a solver
                # cannot reproduce elsewhere.
                doc.Fields.Add(
                    Range=citation_range, Type=-1,
                    Text=f'CITATION {config["source_tag"]}',
                    PreserveFormatting=True,
                )
                time.sleep(0.3)

            # 6. Insert bibliography (if source_tag present)
            if config.get('source_tag'):
                wd.Selection.EndKey(Unit=6)
                sel = wd.Selection
                sel.TypeParagraph()
                # Insert BIBLIOGRAPHY field (Type=-1 = wdFieldEmpty with explicit code)
                # Note: Type=8 is wdFieldIndex, NOT bibliography
                sel.Fields.Add(Range=sel.Range, Type=-1,
                               Text='BIBLIOGRAPHY', PreserveFormatting=True)
                time.sleep(0.5)
                # Update fields so bibliography populates from sources
                doc.Fields.Update()
                time.sleep(0.5)

            wd.Selection.HomeKey(Unit=6)
            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
        finally:
            doc.Close(False)
        time.sleep(0.3)

        atomic_operations = _enrichment_atomic_operations(config)
        op_count = len(atomic_operations)
        atomic_intents = _record_note_text(
            build_content_operations_atomic_intents(atomic_operations), config
        )
        spec = {
            'behavior': 'content_operations',
            'scope': 'additive_composition',
            'variation': variant_name,
            'generated_at': datetime.now().isoformat(),
            'composition': {
                'type': 'additive',
                'atomic_intents': atomic_intents,
                'operation_count': op_count,
            },
            'precondition': {
                'input_state': 'with_sources' if config.get('source_tag') else 'blank',
                'full_text': 'Alpha TARGET Omega.',
                'cursor_after': 'TARGET',
            },
            'verification': {
                'primary': {
                    'method': 'document_enrichment_composition_verification',
                    'description': 'Verify footnote, endnote, header, footer, citation, bibliography',
                },
            },
            'files': {
                'doc_in': 'doc_in.docx',
                'gold': 'gold.docx',
            },
        }
        field_mapping = _enrichment_field_mapping(config)
        if field_mapping:
            spec['field_mapping'] = field_mapping
        write_spec_if_changed(spec_path, spec)
        print(f"  [{case_name}/{variant_name}] Generated ({op_count} operations)")
        return True

    def _create_cleanup_doc_in(self, wd, doc_in_path, precondition_type):
        """Create custom doc_in for document cleanup with various preconditions."""
        doc = wd.Documents.Add()
        try:
            # Base text
            doc.TrackRevisions = False
            doc.Content.Text = 'Alpha TARGET Omega.'
            time.sleep(0.2)

            # Add tracked changes
            doc.TrackRevisions = True
            wd.Selection.HomeKey(Unit=6)
            wd.Selection.Find.ClearFormatting()
            wd.Selection.Find.Execute(FindText='TARGET', Forward=True)
            wd.Selection.TypeText('MODIFIED')
            time.sleep(0.2)

            if precondition_type == 'track_changes_comments_watermark_footer':
                # Add comments (normal)
                wd.Selection.HomeKey(Unit=6)
                doc.TrackRevisions = False
                wd.Selection.Find.ClearFormatting()
                if wd.Selection.Find.Execute(FindText='Alpha', Forward=True):
                    c = doc.Comments.Add(wd.Selection.Range, 'Normal comment.')
                    c.Author = COMMENT_AUTHOR
                wd.Selection.HomeKey(Unit=6)
                if wd.Selection.Find.Execute(FindText='Omega', Forward=True):
                    c2 = doc.Comments.Add(wd.Selection.Range, 'Another comment.')
                    c2.Author = COMMENT_AUTHOR

                # Add watermark
                try:
                    wd.Templates.LoadBuildingBlocks()
                    time.sleep(0.3)
                    template = wd.Application.Templates(1)
                    bb = template.BuildingBlockTypes(8).Categories("Disclaimers").BuildingBlocks("DRAFT 1")
                    section = doc.Sections(1)
                    bb.Insert(section.Headers(1).Range, True)
                except Exception as e:
                    print(f"    Warning: Could not add watermark: {e}")

                # Add footer text
                doc.Sections(1).Footers(1).Range.Text = 'Page Footer'

            elif precondition_type == 'track_changes_resolved_comments_header_footer':
                # Add resolved comments
                wd.Selection.HomeKey(Unit=6)
                doc.TrackRevisions = False
                wd.Selection.Find.ClearFormatting()
                if wd.Selection.Find.Execute(FindText='Alpha', Forward=True):
                    c = doc.Comments.Add(wd.Selection.Range, 'Resolved comment.')
                    c.Author = COMMENT_AUTHOR
                    c.Done = True
                wd.Selection.HomeKey(Unit=6)
                if wd.Selection.Find.Execute(FindText='Omega', Forward=True):
                    c2 = doc.Comments.Add(wd.Selection.Range, 'Normal comment.')
                    c2.Author = COMMENT_AUTHOR

                # Add header and footer text
                section = doc.Sections(1)
                section.Headers(1).Range.Text = 'Document Header'
                section.Footers(1).Range.Text = 'Page Footer'

            doc.SaveAs(str(doc_in_path.absolute()), FileFormat=16)
        finally:
            doc.Close(False)
        time.sleep(0.3)

    def _generate_document_cleanup_variant(self, wd, case_name, variant_name, config):
        """Generate a document cleanup composition variant."""
        variant_dir = self.output_dir / case_name / variant_name
        variant_dir.mkdir(parents=True, exist_ok=True)
        gold_path = variant_dir / 'gold.docx'
        spec_path = variant_dir / 'spec.json'
        doc_in_path = variant_dir / 'doc_in.docx'

        # Create custom doc_in with preconditions
        self._create_cleanup_doc_in(wd, doc_in_path, config['precondition'])

        # Open doc_in and apply cleanup operations to create gold
        doc = wd.Documents.Open(str(doc_in_path.absolute()))
        try:
            for op in config['operations']:
                if op == 'accept_changes':
                    doc.AcceptAllRevisions()
                    doc.TrackRevisions = False
                elif op == 'reject_changes':
                    doc.RejectAllRevisions()
                    doc.TrackRevisions = False
                elif op == 'delete_comments':
                    while doc.Comments.Count > 0:
                        doc.Comments(1).Delete()
                elif op == 'delete_resolved_comments':
                    # Delete resolved comments only (iterate backwards)
                    for i in range(doc.Comments.Count, 0, -1):
                        try:
                            if doc.Comments(i).Done:
                                doc.Comments(i).Delete()
                        except Exception:
                            pass
                elif op == 'remove_watermark':
                    # Remove shapes from primary header (watermarks are shapes in header)
                    header = doc.Sections(1).Headers(1)
                    for i in range(header.Shapes.Count, 0, -1):
                        try:
                            header.Shapes(i).Delete()
                        except Exception:
                            pass
                elif op == 'remove_header':
                    doc.Sections(1).Headers(1).Range.Delete()
                elif op == 'remove_footer':
                    doc.Sections(1).Footers(1).Range.Delete()
                time.sleep(0.2)

            wd.Selection.HomeKey(Unit=6)
            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
        finally:
            doc.Close(False)
        time.sleep(0.3)

        atomic_operations = _cleanup_atomic_operations(config)
        atomic_intents = build_content_operations_atomic_intents(atomic_operations)
        spec = {
            'behavior': 'content_operations',
            'scope': 'additive_composition',
            'variation': variant_name,
            'generated_at': datetime.now().isoformat(),
            'composition': {
                'type': 'additive',
                'atomic_intents': atomic_intents,
                'operation_count': len(atomic_operations),
            },
            'precondition': {
                'input_state': config['precondition'],
                'full_text': 'Alpha TARGET Omega.',
            },
            'verification': {
                'primary': {
                    'method': 'document_cleanup_composition_verification',
                    'description': 'Verify tracked changes, comments, header/footer, watermark cleanup',
                    'params': {'operations': config['operations']},
                },
            },
            'files': {
                'doc_in': 'doc_in.docx',
                'gold': 'gold.docx',
            },
        }
        write_spec_if_changed(spec_path, spec)
        print(f"  [{case_name}/{variant_name}] Generated ({len(atomic_operations)} operations)")
        return True

    def generate_all(self, command: str = None, variant_names=None):
        """Generate all content operations composition test cases."""
        requested_variants = set(variant_names or [])
        generated_variants = set()
        wd = ensure_word(visible=not self.hidden)
        try:
            results = {}

            if command is None or command == 'AdditiveFootnoteEndnoteComposition':
                print(f"\n=== Footnote + Endnote Compositions ===")
                for vn, cfg in FOOTNOTE_ENDNOTE_VARIANTS.items():
                    if requested_variants and vn not in requested_variants:
                        continue
                    try:
                        ok = self._generate_footnote_endnote_variant(
                            wd, 'AdditiveFootnoteEndnoteComposition', vn, cfg
                        )
                        results[f'AdditiveFootnoteEndnoteComposition/{vn}'] = ok
                        generated_variants.add(vn)
                    except Exception as e:
                        print(f"  [ERROR] {vn}: {e}")
                        import traceback
                        traceback.print_exc()
                        results[f'AdditiveFootnoteEndnoteComposition/{vn}'] = False

            if command is None or command == 'AdditiveHeaderFooterComposition':
                print(f"\n=== Header + Footer Compositions ===")
                for vn, cfg in HEADER_FOOTER_VARIANTS.items():
                    if requested_variants and vn not in requested_variants:
                        continue
                    try:
                        ok = self._generate_header_footer_variant(
                            wd, 'AdditiveHeaderFooterComposition', vn, cfg
                        )
                        results[f'AdditiveHeaderFooterComposition/{vn}'] = ok
                        generated_variants.add(vn)
                    except Exception as e:
                        print(f"  [ERROR] {vn}: {e}")
                        import traceback
                        traceback.print_exc()
                        results[f'AdditiveHeaderFooterComposition/{vn}'] = False

            # ---- New composition types (additive extension plan) ----

            # Document Enrichment compositions (§2)
            if command is None or command == 'AdditiveDocumentEnrichmentComposition':
                print(f"\n=== Document Enrichment Compositions ===")
                for vn, cfg in DOCUMENT_ENRICHMENT_VARIANTS.items():
                    if requested_variants and vn not in requested_variants:
                        continue
                    try:
                        ok = self._generate_document_enrichment_variant(
                            wd, 'AdditiveDocumentEnrichmentComposition', vn, cfg
                        )
                        results[f'AdditiveDocumentEnrichmentComposition/{vn}'] = ok
                        generated_variants.add(vn)
                    except Exception as e:
                        print(f"  [ERROR] {vn}: {e}")
                        import traceback
                        traceback.print_exc()
                        results[f'AdditiveDocumentEnrichmentComposition/{vn}'] = False

            # Document Cleanup compositions (§3)
            if command is None or command == 'AdditiveDocumentCleanupComposition':
                print(f"\n=== Document Cleanup Compositions ===")
                for vn, cfg in DOCUMENT_CLEANUP_VARIANTS.items():
                    if requested_variants and vn not in requested_variants:
                        continue
                    try:
                        ok = self._generate_document_cleanup_variant(
                            wd, 'AdditiveDocumentCleanupComposition', vn, cfg
                        )
                        results[f'AdditiveDocumentCleanupComposition/{vn}'] = ok
                        generated_variants.add(vn)
                    except Exception as e:
                        print(f"  [ERROR] {vn}: {e}")
                        import traceback
                        traceback.print_exc()
                        results[f'AdditiveDocumentCleanupComposition/{vn}'] = False

            # Write instructions JSON (source of truth is in variant defs)
            commands = {}
            for cmd_name, variants in [
                ('AdditiveDocumentEnrichmentComposition', DOCUMENT_ENRICHMENT_VARIANTS),
                ('AdditiveDocumentCleanupComposition', DOCUMENT_CLEANUP_VARIANTS),
            ]:
                commands[cmd_name] = {
                    vname: vdef['instruction']
                    for vname, vdef in variants.items()
                }
            write_instructions_json(TEST_SET, commands)

            unknown_variants = requested_variants - generated_variants
            if unknown_variants:
                raise ValueError(
                    'Requested variant(s) not found under the selected command: '
                    + ', '.join(sorted(unknown_variants))
                )

            total = len(results)
            passed = sum(1 for v in results.values() if v)
            print(f"\n=== Summary: {passed}/{total} test cases generated ===")

            if command is None and not requested_variants:
                summary_path = self.output_dir / 'generation_summary.json'
                summary = {
                    'test_set': TEST_SET,
                    'generated_at': datetime.now().isoformat(),
                    'total': total,
                    'passed': passed,
                    'results': {
                        k: 'success' if v else 'failed'
                        for k, v in results.items()
                    },
                }
                with open(summary_path, 'w', encoding='utf-8') as f:
                    json.dump(summary, f, indent=2, ensure_ascii=False)
            return results
        finally:
            pass


def main():
    import argparse
    parser = argparse.ArgumentParser(
        description='Generate additive content operations composition test cases'
    )
    parser.add_argument('--command', '-c', help='Generate specific case only')
    parser.add_argument(
        '--variant', '-v', action='append',
        help='Generate only this variant (repeat for multiple variants)',
    )
    parser.add_argument('--output', '-o', help='Output directory')
    parser.add_argument('--hidden', action='store_true', help='Run Word hidden')
    args = parser.parse_args()

    output_dir = Path(args.output) if args.output else None
    gen = AdditiveContentOperationsGenerator(output_dir=output_dir, hidden=args.hidden)
    gen.generate_all(command=args.command, variant_names=args.variant)


if __name__ == '__main__':
    main()
