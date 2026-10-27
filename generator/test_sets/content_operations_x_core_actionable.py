"""
Test case generator for content_operations × core_actionable.

This UNIFIED generator handles commands that perform operations on content elements:
- INSERT: Add new content (footnotes, page breaks, equations, etc.)
- MODIFY: Remove or update existing content (headers, footers, TOC, index)
- STRUCTURE: Modify document structure (breaks, sections)

All verified using XPath-based checks with different conditions.

Key Design:
- Configuration-driven approach (CommandConfig dataclass)
- Different input states: BLANK (for INSERT) vs UNIFIED (for MODIFY)
- Different cursor handling: position_after_text vs select_element
- XPath-based verification using new verification classes
"""

import json
import time
from dataclasses import dataclass
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Dict, List, Optional

from ..config import (
    COMMAND_SETS_DIR, BENCHMARK_OUTPUT_DIR,
    NAMESPACES, ONE_WAY_VARIATIONS
)
from ..execution.setup import ensure_word, finalize_word
from ..execution.execution import DirectExecute
from generator.test_sets.spec_utils import write_spec_if_changed


# =============================================================================
# CONFIGURATION ENUMS
# =============================================================================

class InputState(Enum):
    """Input document state for test cases."""
    BLANK = "blank"      # Simple text, for INSERT commands
    UNIFIED = "unified"  # Contains header, footer, TOC, index for MODIFY commands


class CursorHandling(Enum):
    """How to position cursor before command execution."""
    POSITION_AFTER_TEXT = "position_after_text"  # For INSERT: position after "TARGET"
    SELECT_ELEMENT = "select_element"            # For MODIFY: select specific element


# =============================================================================
# COMMAND CONFIGURATION
# =============================================================================

@dataclass
class CommandConfig:
    """Configuration for a single command's test case generation."""
    idMso: str
    input_state: InputState
    cursor_handling: CursorHandling
    post_execution: Optional[str] = None  # 'type_content' or None
    post_content: Optional[str] = None    # Content to type after execution
    xpath: str = ""                        # XPath pattern for verification
    condition: str = "exists"              # XPath condition: exists, not_exists, contains:text, etc.
    description: str = ""                  # Human-readable description
    skip_reason: Optional[str] = None      # If set, command will be skipped
    # Primary verification override (default: structure_verification via Word COM)
    primary_method: Optional[str] = None   # e.g., 'header_footer_content_check'
    primary_params: Optional[Dict] = None  # Parameters for primary method
    # Additional files to compare
    additional_files: Optional[List[str]] = None  # e.g., ['word/footnotes.xml']


# =============================================================================
# COMMAND CONFIGURATIONS
# =============================================================================

def get_command_configs() -> Dict[str, CommandConfig]:
    """
    Get configuration for all commands in this test set.

    Returns dict mapping idMso -> CommandConfig.
    """
    return {
        # =====================================================================
        # INSERT COMMANDS (7 directly verifiable)
        # =====================================================================

        # Footnotes and Endnotes
        # Note: Must compare footnotes.xml/endnotes.xml in addition to document.xml
        # because the actual footnote/endnote content is stored there, not in document.xml.
        # Without this, models with different language settings would pass incorrectly.
        'FootnoteInsert': CommandConfig(
            idMso='FootnoteInsert',
            input_state=InputState.BLANK,
            cursor_handling=CursorHandling.POSITION_AFTER_TEXT,
            post_execution='type_content',
            post_content='See reference.',
            xpath='//w:footnoteReference',
            condition='exists',
            description='Footnote reference marker inserted in document',
            primary_method='footnote_verification',
            additional_files=['word/footnotes.xml'],
        ),

        'EndnoteInsertWord': CommandConfig(
            idMso='EndnoteInsertWord',
            input_state=InputState.BLANK,
            cursor_handling=CursorHandling.POSITION_AFTER_TEXT,
            post_execution='type_content',
            post_content='See reference.',
            xpath='//w:endnoteReference',
            condition='exists',
            description='Endnote reference marker inserted in document',
            primary_method='endnote_verification',
            additional_files=['word/endnotes.xml'],
        ),

        # Page Breaks
        'PageBreakInsertWord': CommandConfig(
            idMso='PageBreakInsertWord',
            input_state=InputState.BLANK,
            cursor_handling=CursorHandling.POSITION_AFTER_TEXT,
            xpath='//w:br[@w:type="page"]',
            condition='exists',
            description='Page break element inserted',
            primary_method='page_verification',
        ),

        'BlankPageInsert': CommandConfig(
            idMso='BlankPageInsert',
            input_state=InputState.BLANK,
            cursor_handling=CursorHandling.POSITION_AFTER_TEXT,
            xpath='//w:br[@w:type="page"]',
            condition='exists',
            description='Blank page (page breaks) inserted',
            primary_method='page_verification',
        ),

        # Horizontal Line
        'HorizontalLineInsert': CommandConfig(
            idMso='HorizontalLineInsert',
            input_state=InputState.BLANK,
            cursor_handling=CursorHandling.POSITION_AFTER_TEXT,
            xpath=(
                '//w:pict/*[local-name()="rect" '
                'and @*[local-name()="hr"]="t"]'
            ),
            condition='exists',
            description='Word horizontal-line VML object inserted',
            primary_method='inline_shape_verification',
        ),

        # Equation
        'EquationInsertNew': CommandConfig(
            idMso='EquationInsertNew',
            input_state=InputState.BLANK,
            cursor_handling=CursorHandling.POSITION_AFTER_TEXT,
            xpath='//m:oMath',
            condition='exists',
            description='Math equation element inserted',
            primary_method='equation_verification',
        ),

        # Bibliography
        'BibliographyInsert': CommandConfig(
            idMso='BibliographyInsert',
            input_state=InputState.BLANK,
            cursor_handling=CursorHandling.POSITION_AFTER_TEXT,
            xpath='//w:sdt[.//w:docPartGallery]',
            condition='exists',
            description='Bibliography content control inserted',
            primary_method='bibliography_verification',
        ),

        # =====================================================================
        # MODIFY COMMANDS (4 directly verifiable)
        # =====================================================================

        'HeaderFooterRemoveHeaderWord': CommandConfig(
            idMso='HeaderFooterRemoveHeaderWord',
            input_state=InputState.UNIFIED,
            cursor_handling=CursorHandling.SELECT_ELEMENT,
            description='Header text content removed',
            # Use semantic check as primary - verify no text in header files
            primary_method='header_footer_content_check',
            primary_params={'part_type': 'header', 'condition': 'not_exists'},
        ),

        'HeaderFooterRemoveFooterWord': CommandConfig(
            idMso='HeaderFooterRemoveFooterWord',
            input_state=InputState.UNIFIED,
            cursor_handling=CursorHandling.SELECT_ELEMENT,
            description='Footer text content removed',
            # Use semantic check as primary - verify no text in footer files
            primary_method='header_footer_content_check',
            primary_params={'part_type': 'footer', 'condition': 'not_exists'},
        ),

        # Use XPath semantic check as primary because XML comparison fails due to
        # volatile elements (different lengths but Word sees them as identical).
        # The semantic check verifies the TOC SDT element is gone, which is the
        # core requirement for this command.
        'TableOfContentsRemove': CommandConfig(
            idMso='TableOfContentsRemove',
            input_state=InputState.UNIFIED,
            cursor_handling=CursorHandling.SELECT_ELEMENT,
            xpath='//w:sdt[w:sdtPr/w:docPartObj/w:docPartGallery[@w:val="Table of Contents"]]',
            condition='not_exists',
            description='Table of contents removed',
            primary_method='toc_verification',
            primary_params={},
        ),

        'IndexUpdate': CommandConfig(
            idMso='IndexUpdate',
            input_state=InputState.UNIFIED,
            cursor_handling=CursorHandling.SELECT_ELEMENT,
            description='Index updated to reflect current document content',
            primary_method='index_verification',
        ),

        # =====================================================================
        # SKIPPED COMMANDS (require special handling within core_actionable)
        # =====================================================================
        # Note: Commands requiring dialogs, file pickers, drawing interaction, etc.
        # have been reclassified to core_dialog or external_effect and removed from
        # this generator. Only commands that belong in core_actionable but need
        # special handling remain here.

        # MODIFY - Requires manual template (Building Block structures)
        'PageNumbersRemove': CommandConfig(
            idMso='PageNumbersRemove',
            input_state=InputState.UNIFIED,
            cursor_handling=CursorHandling.SELECT_ELEMENT,
            skip_reason='requires_manual_template'
        ),

        'CoverPageRemove': CommandConfig(
            idMso='CoverPageRemove',
            input_state=InputState.UNIFIED,
            cursor_handling=CursorHandling.SELECT_ELEMENT,
            skip_reason='requires_manual_template'
        ),

        # InsertNewComment lives in content_operations_x_specialized_actionable
        # (its real classification is specialized_actionable, on the Review tab).
    }


# =============================================================================
# INPUT DOCUMENT CONFIGURATIONS
# =============================================================================

# For INSERT commands: simple text with cursor positioning
BLANK_INPUT_CONFIG = {
    'full_text': 'Alpha TARGET Omega.',
    'cursor_after': 'TARGET',
    'filename': 'doc_in.docx',
}


def setup_unified_input(doc, wd):
    """
    Setup unified input document containing elements for MODIFY commands.

    Contains:
    - Header and Footer
    - Table of Contents with headings
    - Index with marked entries

    This single template serves all MODIFY commands.
    """
    # 1. Setup Header and Footer
    section = doc.Sections(1)
    header = section.Headers(1)  # wdHeaderFooterPrimary
    header.Range.Text = "Document Header"

    footer = section.Footers(1)  # wdHeaderFooterPrimary
    footer.Range.Text = "Document Footer"

    # 2. Basic content
    doc.Content.Text = "Sample Document\n\n"

    # 3. Add Headings for Table of Contents (minimum 2 headings)
    doc.Content.InsertAfter("Chapter 1\n")
    rng_h1 = doc.Range(doc.Content.End - 10, doc.Content.End - 1)
    rng_h1.Style = "Heading 1"

    doc.Content.InsertAfter("Content for chapter 1.\n\n")

    doc.Content.InsertAfter("Chapter 2\n")
    rng_h1_2 = doc.Range(doc.Content.End - 10, doc.Content.End - 1)
    rng_h1_2.Style = "Heading 1"

    doc.Content.InsertAfter("Content for chapter 2.\n\n")

    # 4. Create Table of Contents
    doc.Content.InsertAfter("\n")
    rng_toc = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    toc = doc.TablesOfContents.Add(rng_toc)
    toc.Update()

    # 5. Add Index entries and create Index
    doc.Content.InsertAfter("\n\nApple is a fruit.\n")
    rng1 = doc.Range(doc.Content.End - 19, doc.Content.End - 14)
    doc.Indexes.MarkEntry(rng1, "Apple")

    doc.Content.InsertAfter("Banana is a fruit.\n")
    rng2 = doc.Range(doc.Content.End - 19, doc.Content.End - 13)
    doc.Indexes.MarkEntry(rng2, "Banana")

    # Create index
    doc.Content.InsertAfter("\n")
    rng_idx = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    idx = doc.Indexes.Add(rng_idx)
    idx.Update()

    # Add one more entry after index (so IndexUpdate has something to update)
    doc.Content.InsertAfter("\nCherry is a fruit.\n")
    rng3 = doc.Range(doc.Content.End - 19, doc.Content.End - 13)
    doc.Indexes.MarkEntry(rng3, "Cherry")


# =============================================================================
# GENERATOR CLASS
# =============================================================================

class ContentOperationsCoreActionableGenerator:
    """
    Unified generator for content_operations × core_actionable test cases.

    Handles INSERT, MODIFY, and STRUCTURE operations with configuration-driven approach.

    Test set: 13 commands total
    - 11 directly verifiable commands (7 INSERT + 4 MODIFY)
    - 2 skipped commands requiring special handling:
      * CoverPageRemove, PageNumbersRemove (need manual templates)
    - All button type: 1 test each (apply)
    - Uses XPath-based verification
    - Different input states and cursor handling based on operation type

    Note: Commands requiring dialogs, file pickers, or drawing have been
    reclassified to core_dialog or external_effect and moved to separate generators.
    """

    def __init__(self, output_dir: Path = None):
        self.execution = DirectExecute()
        self.output_dir = output_dir or (BENCHMARK_OUTPUT_DIR / 'content_operations_x_core_actionable')
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.inputs_dir = self.output_dir / 'inputs'
        self.inputs_dir.mkdir(parents=True, exist_ok=True)

        # Load command configurations
        self.command_configs = get_command_configs()

        # Cache for input documents
        self.blank_input_path = None
        self.unified_input_path = None

    def load_commands(self) -> List[Dict]:
        """Load commands classified as content_operations × core_actionable."""
        commands_file = COMMAND_SETS_DIR / 'commands_by_policy_id.json'

        if not commands_file.exists():
            print(f"Warning: {commands_file} not found.")
            return []

        with open(commands_file, 'r', encoding='utf-8') as f:
            all_commands = json.load(f)

        return [
            cmd for cmd in all_commands
            if cmd.get('behavior') == 'content_operations'
            and cmd.get('scope') == 'core_actionable'
        ]

    def generate_blank_input(self) -> Path:
        """Generate blank input document (for INSERT commands)."""
        if self.blank_input_path and self.blank_input_path.exists():
            return self.blank_input_path

        input_path = self.inputs_dir / BLANK_INPUT_CONFIG['filename']

        if input_path.exists():
            print(f"  Blank input already exists: {input_path}")
            self.blank_input_path = input_path
            return input_path

        print(f"  Generating blank input: {input_path}")

        wd = ensure_word()
        try:
            doc = wd.Documents.Add()
            wd.Selection.TypeText(BLANK_INPUT_CONFIG['full_text'])

            # Add a citation source so BibliographyInsert has content to work with
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
            doc.Bibliography.Sources.Add(source_xml)

            doc.SaveAs(str(input_path.absolute()), FileFormat=16)
            doc.Close(False)
            time.sleep(0.2)

            self.blank_input_path = input_path
            print(f"    Created: {input_path}")
            return input_path

        except Exception as e:
            print(f"    Error creating blank input: {e}")
            raise

    def generate_unified_input(self) -> Path:
        """Generate unified input document (for MODIFY commands)."""
        if self.unified_input_path and self.unified_input_path.exists():
            return self.unified_input_path

        input_path = self.inputs_dir / 'doc_in_unified.docx'

        if input_path.exists():
            print(f"  Unified input already exists: {input_path}")
            self.unified_input_path = input_path
            return input_path

        print(f"  Generating unified input: {input_path}")

        wd = ensure_word()
        try:
            doc = wd.Documents.Add()
            setup_unified_input(doc, wd)
            doc.SaveAs(str(input_path.absolute()), FileFormat=16)
            doc.Close(False)
            time.sleep(0.2)

            self.unified_input_path = input_path
            print(f"    Created: {input_path}")
            return input_path

        except Exception as e:
            print(f"    Error creating unified input: {e}")
            raise

    def position_cursor_after_text(self, wd, text: str) -> bool:
        """Position cursor immediately after specified text."""
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

    def select_element_for_modify(self, wd, idMso: str) -> bool:
        """Select appropriate element for MODIFY commands."""
        doc = wd.ActiveDocument

        if idMso == 'IndexUpdate':
            if doc.Indexes.Count > 0:
                doc.Indexes(1).Range.Select()
                time.sleep(0.1)
                return True
            return False

        elif idMso == 'TableOfFiguresUpdate':
            if doc.TablesOfFigures.Count > 0:
                doc.TablesOfFigures(1).Range.Select()
                time.sleep(0.1)
                return True
            return False

        # For removal commands, no selection needed (they operate on entire document)
        return True

    def generate_apply_test(self, idMso: str, config: CommandConfig, variation: Dict) -> Dict:
        """Generate 'apply' test case for a command."""
        test_dir = self.output_dir / idMso / variation['name']
        test_dir.mkdir(parents=True, exist_ok=True)

        gold_path = test_dir / 'gold.docx'
        spec_path = test_dir / 'spec.json'

        # Determine input path based on input state
        if config.input_state == InputState.BLANK:
            input_path = self.blank_input_path
            input_filename = BLANK_INPUT_CONFIG['filename']
        else:  # InputState.UNIFIED
            input_path = self.unified_input_path
            input_filename = 'doc_in_unified.docx'

        result = {
            'idMso': idMso,
            'variation': variation['name'],
            'success': False,
            'message': '',
            'paths': {
                'doc_in': str(input_path),
                'gold': str(gold_path),
                'spec': str(spec_path),
            }
        }

        # Check for skip reason
        if config.skip_reason:
            result['message'] = f'Skipped: {config.skip_reason}'
            result['skipped'] = True
            # Still generate spec for documentation
            spec = self._generate_spec(idMso, config, variation, input_filename)
            spec['skipped'] = True
            spec['skip_reason'] = config.skip_reason
            write_spec_if_changed(spec_path, spec)
            return result

        wd = ensure_word()
        try:
            # Open input document
            doc = wd.Documents.Open(str(input_path.absolute()))

            # Handle cursor positioning based on configuration
            if config.cursor_handling == CursorHandling.POSITION_AFTER_TEXT:
                cursor_text = BLANK_INPUT_CONFIG.get('cursor_after', 'TARGET')
                if not self.position_cursor_after_text(wd, cursor_text):
                    result['message'] = f"Failed to position cursor after: {cursor_text}"
                    doc.Close(False)
                    return result

            elif config.cursor_handling == CursorHandling.SELECT_ELEMENT:
                if not self.select_element_for_modify(wd, idMso):
                    result['message'] = f"Failed to select element for: {idMso}"
                    doc.Close(False)
                    return result

            # Execute the command
            if not self.execution.execute(wd, idMso):
                result['message'] = f'Failed to execute command: {idMso}'
                doc.Close(False)
                return result

            # Handle post-execution actions
            if config.post_execution == 'type_content' and config.post_content:
                time.sleep(0.1)  # Brief pause for cursor to move
                wd.Selection.TypeText(config.post_content)

            time.sleep(0.5)

            # Save as gold
            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
            doc.Close(False)
            time.sleep(0.2)

            # Generate spec.json
            spec = self._generate_spec(idMso, config, variation, input_filename)
            write_spec_if_changed(spec_path, spec)

            result['success'] = True
            result['message'] = 'Test case generated successfully'

        except Exception as e:
            result['message'] = f'Error: {str(e)}'
            import traceback
            traceback.print_exc()

        return result

    def _generate_spec(
        self,
        idMso: str,
        config: CommandConfig,
        variation: Dict,
        input_filename: str
    ) -> Dict:
        """Generate spec.json for a test case."""
        # Determine primary verification method
        if config.primary_method:
            primary_spec = {
                'method': config.primary_method,
                'description': config.description or f'{config.primary_method} verification',
            }
            # Add parameters if specified
            if config.primary_params:
                primary_spec.update(config.primary_params)
        else:
            primary_spec = {
                'method': 'structure_verification',
                'description': 'Word COM document structure comparison',
            }
            if config.additional_files:
                primary_spec['additional_files'] = config.additional_files

        spec = {
            'idMso': idMso,
            'behavior': 'content_operations',
            'scope': 'core_actionable',
            'variation': variation['name'],
            'generated_at': datetime.now().isoformat(),

            'precondition': {
                'input_state': config.input_state.value,
                'cursor_handling': config.cursor_handling.value,
            },

            'verification': {
                'primary': primary_spec,
            },

            'files': {
                'doc_in': f'../../inputs/{input_filename}',
                'gold': 'gold.docx',
            }
        }

        # Add precondition details for blank input
        if config.input_state == InputState.BLANK:
            spec['precondition']['full_text'] = BLANK_INPUT_CONFIG['full_text']
            spec['precondition']['cursor_after'] = BLANK_INPUT_CONFIG.get('cursor_after', 'TARGET')

        return spec

    def generate_for_command(self, command: Dict) -> List[Dict]:
        """Generate all test variations for a single command."""
        idMso = command['idMso']

        # Get configuration (or use default if not configured)
        if idMso not in self.command_configs:
            print(f"  Warning: No configuration for {idMso}, using default")
            config = CommandConfig(
                idMso=idMso,
                input_state=InputState.BLANK,
                cursor_handling=CursorHandling.POSITION_AFTER_TEXT,
                skip_reason='not_configured'
            )
        else:
            config = self.command_configs[idMso]

        # Generate input document if needed
        if config.input_state == InputState.BLANK:
            if not self.blank_input_path:
                self.generate_blank_input()
        else:
            if not self.unified_input_path:
                self.generate_unified_input()

        # Generate test (only 'apply' variation for content_operations)
        results = []
        for var in ONE_WAY_VARIATIONS:
            if var['name'] == 'apply':
                result = self.generate_apply_test(idMso, config, var)
                results.append(result)

                status = 'SKIP' if result.get('skipped') else ('OK' if result['success'] else 'FAIL')
                print(f"  {var['name']}: {status} - {result['message']}")

        return results

    def generate_all(self, commands: List[Dict] = None, limit: int = None) -> Dict:
        """Generate test cases for all commands in this test set."""
        if commands is None:
            commands = self.load_commands()

        if limit:
            commands = commands[:limit]

        print(f"\n=== Generating test cases for {len(commands)} commands ===")
        print(f"Test set: content_operations × core_actionable")
        print(f"Output: {self.output_dir}\n")

        # Step 1: Generate input documents
        print("Step 1: Generating input documents...")
        self.generate_blank_input()
        self.generate_unified_input()
        print()

        summary = {
            'test_set': 'content_operations_x_core_actionable',
            'generated_at': datetime.now().isoformat(),
            'total_commands': len(commands),
            'total_test_cases': 0,
            'successful': 0,
            'failed': 0,
            'skipped': 0,
            'results': [],
        }

        # Step 2: Generate tests for each command
        print("Step 2: Generating test cases...")
        for cmd in commands:
            idMso = cmd['idMso']
            control_type = cmd.get('control_type', 'button')

            print(f"\n{idMso} ({control_type})...")

            results = self.generate_for_command(cmd)

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
            json.dump(summary, f, indent=2, ensure_ascii=False, sort_keys=True)

        # Print results
        self._print_summary(summary, summary_path)

        return summary

    def _print_summary(self, summary: Dict, summary_path: Path):
        """Print generation summary."""
        print(f"\n=== Generation Complete ===")
        print(f"\nResults\n")
        print("| Status     | Count |")
        print("|-----------|-------|")
        print(f"| Successful| {summary['successful']:<5} |")
        print(f"| Skipped   | {summary.get('skipped', 0):<5} |")
        print(f"| Failed    | {summary['failed']:<5} |")

        # Print successful test cases
        successful_results = [r for r in summary['results'] if r['success']]
        skipped_results = [r for r in summary['results'] if r.get('skipped')]

        if successful_results:
            print(f"\nSuccessful Test Cases ({len(successful_results)})\n")
            for r in successful_results:
                config = self.command_configs.get(r['idMso'])
                if config:
                    input_type = "INSERT" if config.input_state == InputState.BLANK else "MODIFY"
                    print(f"  {r['idMso']} ({input_type})")

        if skipped_results:
            print(f"\nSkipped Test Cases ({len(skipped_results)})\n")
            for r in skipped_results:
                config = self.command_configs.get(r['idMso'])
                skip_reason = config.skip_reason if config else 'unknown'
                print(f"  {r['idMso']} - {skip_reason}")

        print(f"\nTotal commands: {summary['total_commands']}")
        print(f"Total test cases: {summary['total_test_cases']}")
        print(f"Summary saved to: {summary_path}")


# =============================================================================
# CLI
# =============================================================================

def main():
    """CLI entry point for generating content_operations × core_actionable tests."""
    import argparse

    parser = argparse.ArgumentParser(
        description='Generate test cases for content_operations × core_actionable (unified INSERT/MODIFY)'
    )
    parser.add_argument(
        '--command', '-c',
        help='Generate for a specific command (e.g., FootnoteInsert, HeaderFooterRemoveHeaderWord)'
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
    generator = ContentOperationsCoreActionableGenerator(output_dir=output_dir)

    if args.command:
        # Generate for a single command
        commands = generator.load_commands()
        cmd = next((c for c in commands if c['idMso'] == args.command), None)

        if cmd:
            print(f"Generating test cases for {args.command}...")
            results = generator.generate_for_command(cmd)
            success = sum(1 for r in results if r['success'])
            print(f"\nDone: {success}/{len(results)} successful")
        else:
            # Command not in filtered set, try with default config
            print(f"Command {args.command} not found in content_operations × core_actionable set.")
            print(f"Attempting generation with default configuration...")
            cmd = {'idMso': args.command, 'control_type': 'button'}
            results = generator.generate_for_command(cmd)
            success = sum(1 for r in results if r['success'])
            print(f"\nDone: {success}/{len(results)} successful")
    else:
        # Generate for all commands
        generator.generate_all(limit=args.limit)


if __name__ == '__main__':
    main()
