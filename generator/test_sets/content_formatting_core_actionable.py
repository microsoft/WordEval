"""
Test case generator for content_formatting × core_actionable.

This is the Phase 1 implementation target.
"""

import json
import os
import shutil
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from ..config import (
    COMMAND_SETS_DIR, COMMAND_SPECS_DIR, BENCHMARK_OUTPUT_DIR,
    TOGGLE_VARIATIONS, ONE_WAY_VARIATIONS, NAMESPACES
)
from ..execution.setup import SimpleTextSetup, ensure_word, finalize_word
from ..execution.execution import DirectExecute
from generator.test_sets.spec_utils import write_spec_if_changed

# Default input configuration for this generator (shared with parameterized)
DEFAULT_INPUT = {
    'full_text': 'Alpha Target Omega.',
    'select_text': 'Target',
    'filename': 'doc_in_single_line.docx',
}

# Commands to skip (COM API errors, popup cancellations, or not available in Word version)
SKIP_COMMANDS = {
    'TextDirectionLeftToRight': 'COM error - may require popup interaction or not available',
    'TextDirectionRightToLeft': 'COM error - may require popup interaction or not available',
    'AlignJustifyWithMixedLanguages': 'COM error - may require popup interaction or not available',
    'AlignJustifyLow': 'COM error - may require popup interaction or not available',
    'AlignJustifyMedium': 'COM error - may require popup interaction or not available',
    'AlignJustifyHigh': 'COM error - may require popup interaction or not available',
    'AlignJustifyThai': 'COM error - may require popup interaction or not available',
    'ParagraphDistributed': 'NO_XML_EFFECT_EXECUTION - command has no visible effect in current Word version or requires specific context',
    'QuickStylesSaveSelectionAsNew': 'NO_XML_EFFECT_EXECUTION - command saves style but does not modify document content (wrong behavior classification)',
    'RejectAllChangesInDocAndStopTracking':'COM error'
}

# Commands that need local doc_in for 'on' variation (not shared input)
# Most 'on' tests use shared unformatted input
# Some need specific preconditions that are unique to that test
NEEDS_LOCAL_DOC_IN_ON = {
    'AlignLeft',  # Needs right-aligned input (unique to this test)
}

# Commands that need local doc_in for 'apply' variation (button commands with special preconditions)
NEEDS_LOCAL_DOC_IN_APPLY = {
    'IndentDecreaseWord',  # Needs indented input to decrease
    'ClearFormatting',  # Needs formatted input to clear
}


# Map command names to property names for property-based verification
COMMAND_TO_PROPERTY = {
    # Character formatting properties
    'Bold': 'bold',
    'Italic': 'italic',
    'Underline': 'underline',
    'Strikethrough': 'strikethrough',
    'DoubleStrikethrough': 'double_strikethrough',
    'Subscript': 'subscript',
    'Superscript': 'superscript',
    'CharacterBorder': 'border',
    'CharacterShading': 'shading',
    'FontSizeIncrease': 'font_size',
    'FontSizeDecrease': 'font_size',
    'FontSizeIncreaseWord': 'font_size',
    'FontSizeDecreaseWord': 'font_size',
    'ClearFormatting': 'clear_formatting',  # Special multi-property check

    # Paragraph formatting properties
    'AlignLeft': 'alignment',
    'AlignCenter': 'alignment',
    'AlignRight': 'alignment',
    'AlignJustify': 'alignment',
    'ParagraphDistributed': 'alignment',
    'IndentIncreaseWord': 'left_indent',
    'IndentDecreaseWord': 'left_indent',
}


class ContentFormattingCoreActionableGenerator:
    """
    Generator for content_formatting × core_actionable test cases.

    This test set includes:
    - Character formatting commands (Bold, Italic, Underline, etc.)
    - Paragraph formatting commands (AlignLeft, AlignCenter, etc.)

    All commands are:
    - BEHAVIOR: content_formatting (verify via w:rPr/w:pPr XML)
    - SCOPE: core_actionable (setup: text + select, execute: direct)

    Test strategy:
    - toggleButton: 2 tests (apply, toggle_off)
    - button: 1 test (default)

    Shared input: All commands in this set use the same unformatted doc_in.
    """

    def __init__(self, output_dir: Path = None):
        self.execution = DirectExecute()
        self.output_dir = output_dir or (BENCHMARK_OUTPUT_DIR / 'content_formatting_x_core_actionable')
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.inputs_dir = self.output_dir / 'inputs'
        self.shared_input_path = None

    def generate_shared_input(self) -> Path:
        """
        Generate the shared input document used by most commands in this set.

        Returns:
            Path to the shared input document
        """
        self.inputs_dir.mkdir(parents=True, exist_ok=True)
        input_path = self.inputs_dir / DEFAULT_INPUT['filename']

        if input_path.exists():
            print(f"Shared input already exists: {input_path}")
            self.shared_input_path = input_path
            return input_path

        print(f"Generating shared input: {input_path}")

        wd = ensure_word()
        try:
            doc = wd.Documents.Add()
            wd.Selection.TypeText(DEFAULT_INPUT['full_text'])
            # Set font size to 12pt for FontSize commands
            wd.Selection.WholeStory()
            wd.Selection.Font.Size = 12
            doc.SaveAs(str(input_path.absolute()), FileFormat=16)
            doc.Close(False)
            time.sleep(0.2)

            self.shared_input_path = input_path
            print(f"  Created: {input_path}")
            return input_path

        except Exception as e:
            print(f"  Error creating shared input: {e}")
            raise

    def load_commands(self) -> List[Dict]:
        """Load commands classified as content_formatting × core_actionable."""
        commands_file = COMMAND_SETS_DIR / 'commands_by_policy_id.json'

        if not commands_file.exists():
            print(f"Warning: {commands_file} not found. Using empty command list.")
            return []

        with open(commands_file, 'r', encoding='utf-8') as f:
            all_commands = json.load(f)

        # Filter for our test set
        filtered = [
            cmd for cmd in all_commands
            if cmd.get('behavior') == 'content_formatting'
            and cmd.get('scope') == 'core_actionable'
        ]

        # Separate skipped commands
        skipped_commands = [cmd for cmd in filtered if cmd['idMso'] in SKIP_COMMANDS]
        filtered = [cmd for cmd in filtered if cmd['idMso'] not in SKIP_COMMANDS]

        if skipped_commands:
            print(f"Skipping {len(skipped_commands)} commands:")
            for cmd in skipped_commands:
                reason = SKIP_COMMANDS[cmd['idMso']]
                print(f"  - {cmd['idMso']}: {reason}")

        return filtered

    def get_variations_for_command(self, command: Dict) -> List[Dict]:
        """
        Get appropriate test variations based on control type.

        Args:
            command: Command dict with 'control_type' field

        Returns:
            List of variation dicts
        """
        control_type = command.get('control_type', 'button')

        if control_type == 'toggleButton':
            return TOGGLE_VARIATIONS
        else:
            return ONE_WAY_VARIATIONS

    def generate_on_test(self, idMso: str, variation: Dict) -> Dict:
        """
        Generate 'on' test: turn formatting on (unformatted → formatted).

        - doc_in: shared unformatted input (or local doc_in for special cases)
        - gold: result after applying command
        """
        test_dir = self.output_dir / idMso / variation['name']
        test_dir.mkdir(parents=True, exist_ok=True)

        doc_in_path = test_dir / 'doc_in.docx'
        gold_path = test_dir / 'gold.docx'
        spec_path = test_dir / 'spec.json'

        # Check if this command needs a local doc_in (not shared)
        needs_local_doc_in = idMso in NEEDS_LOCAL_DOC_IN_ON

        result = {
            'idMso': idMso,
            'variation': variation['name'],
            'success': False,
            'message': '',
            'paths': {
                'doc_in': str(doc_in_path) if needs_local_doc_in else str(self.shared_input_path),
                'gold': str(gold_path),
                'spec': str(spec_path),
            }
        }

        wd = ensure_word()
        try:
            # Step 1: Create doc_in if needed (for commands with special preconditions)
            if needs_local_doc_in:
                # Create custom doc_in with specific precondition
                doc = wd.Documents.Open(str(self.shared_input_path.absolute()))

                if idMso == 'AlignLeft':
                    # Apply right alignment as precondition
                    if not self.execution.select(wd, DEFAULT_INPUT['select_text']):
                        result['message'] = f"Failed to select text: {DEFAULT_INPUT['select_text']}"
                        doc.Close(False)
                        return result
                    if not self.execution.execute(wd, 'AlignRight'):
                        result['message'] = f'Failed to apply AlignRight for precondition'
                        doc.Close(False)
                        return result

                doc.SaveAs(str(doc_in_path.absolute()), FileFormat=16)
                doc.Close(False)
                time.sleep(0.2)

                # Open the custom doc_in
                doc = wd.Documents.Open(str(doc_in_path.absolute()))
            else:
                # Open shared input
                doc = wd.Documents.Open(str(self.shared_input_path.absolute()))

            # Step 2: Select the target text
            if not self.execution.select(wd, DEFAULT_INPUT['select_text']):
                result['message'] = f"Failed to select text: {DEFAULT_INPUT['select_text']}"
                doc.Close(False)
                return result

            # Step 3: Execute the command to create gold
            if not self.execution.execute(wd, idMso):
                result['message'] = f'Failed to execute command: {idMso}'
                doc.Close(False)
                return result

            # Step 4: Save as gold
            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
            doc.Close(False)
            time.sleep(0.2)

            # Generate spec.json
            spec = self._generate_spec(idMso, variation, is_off_test=False, needs_local_doc_in=needs_local_doc_in)
            write_spec_if_changed(spec_path, spec)

            result['success'] = True
            result['message'] = 'Test case generated successfully'

        except Exception as e:
            result['message'] = f'Error: {str(e)}'
            import traceback
            traceback.print_exc()

        return result

    def generate_off_test(self, idMso: str, variation: Dict) -> Dict:
        """
        Generate 'off' test: turn formatting off (formatted → unformatted).

        - doc_in: text with formatting already applied (command-specific)
        - gold: unformatted text (same as shared input)
        """
        test_dir = self.output_dir / idMso / variation['name']
        test_dir.mkdir(parents=True, exist_ok=True)

        doc_in_path = test_dir / 'doc_in.docx'
        gold_path = test_dir / 'gold.docx'
        spec_path = test_dir / 'spec.json'

        result = {
            'idMso': idMso,
            'variation': variation['name'],
            'success': False,
            'message': '',
            'paths': {
                'doc_in': str(doc_in_path),
                'gold': str(gold_path),
                'spec': str(spec_path),
            }
        }

        wd = ensure_word()
        try:
            # Step 1: Create doc_in (formatted version)
            # Open shared input, apply command, save as doc_in
            doc = wd.Documents.Open(str(self.shared_input_path.absolute()))

            if not self.execution.select(wd, DEFAULT_INPUT['select_text']):
                result['message'] = f"Failed to select text: {DEFAULT_INPUT['select_text']}"
                doc.Close(False)
                return result

            if not self.execution.execute(wd, idMso):
                result['message'] = f'Failed to execute command: {idMso}'
                doc.Close(False)
                return result

            doc.SaveAs(str(doc_in_path.absolute()), FileFormat=16)
            doc.Close(False)
            time.sleep(0.2)

            # Step 2: gold = unformatted = copy of shared input
            shutil.copy(self.shared_input_path, gold_path)

            # Generate spec.json
            spec = self._generate_spec(idMso, variation, is_off_test=True)
            write_spec_if_changed(spec_path, spec)

            result['success'] = True
            result['message'] = 'Test case generated successfully'

        except Exception as e:
            result['message'] = f'Error: {str(e)}'
            import traceback
            traceback.print_exc()

        return result

    def generate_apply_test(self, idMso: str, variation: Dict) -> Dict:
        """
        Generate 'apply' test for one-way (button) commands.

        Most commands use shared input (same as 'on' test).
        Some commands need special preconditions (local doc_in).
        """
        # Check if this command needs a local doc_in with special precondition
        needs_local_doc_in = idMso in NEEDS_LOCAL_DOC_IN_APPLY

        if not needs_local_doc_in:
            # Use standard logic (shared input)
            return self.generate_on_test(idMso, variation)

        # Generate test with local doc_in (special precondition)
        test_dir = self.output_dir / idMso / variation['name']
        test_dir.mkdir(parents=True, exist_ok=True)

        doc_in_path = test_dir / 'doc_in.docx'
        gold_path = test_dir / 'gold.docx'
        spec_path = test_dir / 'spec.json'

        result = {
            'idMso': idMso,
            'variation': variation['name'],
            'success': False,
            'message': '',
            'paths': {
                'doc_in': str(doc_in_path),
                'gold': str(gold_path),
                'spec': str(spec_path),
            }
        }

        wd = ensure_word()
        try:
            # Step 1: Create doc_in with special precondition
            doc = wd.Documents.Open(str(self.shared_input_path.absolute()))

            if idMso == 'IndentDecreaseWord':
                # Apply indent as precondition
                if not self.execution.select(wd, DEFAULT_INPUT['select_text']):
                    result['message'] = f"Failed to select text: {DEFAULT_INPUT['select_text']}"
                    doc.Close(False)
                    return result
                if not self.execution.execute(wd, 'IndentIncreaseWord'):
                    result['message'] = f'Failed to apply IndentIncreaseWord for precondition'
                    doc.Close(False)
                    return result

            elif idMso == 'ClearFormatting':
                # Apply multiple formatting as precondition
                if not self.execution.select(wd, DEFAULT_INPUT['select_text']):
                    result['message'] = f"Failed to select text: {DEFAULT_INPUT['select_text']}"
                    doc.Close(False)
                    return result
                # Apply bold, italic, and underline
                for cmd in ['Bold', 'Italic', 'Underline']:
                    if not self.execution.execute(wd, cmd):
                        result['message'] = f'Failed to apply {cmd} for precondition'
                        doc.Close(False)
                        return result

            doc.SaveAs(str(doc_in_path.absolute()), FileFormat=16)
            doc.Close(False)
            time.sleep(0.2)

            # Step 2: Open doc_in and execute command to create gold
            doc = wd.Documents.Open(str(doc_in_path.absolute()))

            if not self.execution.select(wd, DEFAULT_INPUT['select_text']):
                result['message'] = f"Failed to select text: {DEFAULT_INPUT['select_text']}"
                doc.Close(False)
                return result

            if not self.execution.execute(wd, idMso):
                result['message'] = f'Failed to execute command: {idMso}'
                doc.Close(False)
                return result

            # Step 3: Save as gold
            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
            doc.Close(False)
            time.sleep(0.2)

            # Generate spec.json
            spec = self._generate_spec(idMso, variation, is_off_test=False, needs_local_doc_in=True)
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
        variation: Dict,
        is_off_test: bool = False,
        needs_local_doc_in: bool = False
    ) -> Dict:
        """Generate the spec.json for a test case."""

        select_text = DEFAULT_INPUT['select_text']
        variation_name = variation['name']

        # Calculate relative path to input
        if is_off_test or needs_local_doc_in:
            doc_in_ref = 'doc_in.docx'  # Local file
        else:
            doc_in_ref = '../../inputs/doc_in_single_line.docx'  # Reference to shared input

        spec = {
            'idMso': idMso,
            'behavior': 'content_formatting',
            'scope': 'core_actionable',
            'variation': variation_name,
            'generated_at': datetime.now().isoformat(),

            'precondition': {
                'template': 'simple_text',
                'full_text': DEFAULT_INPUT['full_text'],
                'select_text': select_text,
                'state': variation.get('precondition', 'unformatted'),
            },

            'verification': {},

            'files': {
                'doc_in': doc_in_ref,
                'gold': 'gold.docx',
            }
        }

        # Use property-based verification as primary method if property mapping exists
        property_name = COMMAND_TO_PROPERTY.get(idMso)
        if property_name:
            spec['verification']['primary'] = {
                'method': 'property_based',
                'property_name': property_name,
                'target_text': select_text,
                'description': f'Verify {property_name} property on target text',
            }
            spec['verification']['invariants'] = [
                {
                    'method': 'text_preserved',
                    'description': 'Plain text content unchanged',
                }
            ]
            # Add surrounding text unchanged invariant for character formatting
            char_properties = ['bold', 'italic', 'underline', 'strikethrough',
                             'subscript', 'superscript', 'font_size', 'border', 'shading']
            if property_name in char_properties:
                spec['verification']['invariants'].append({
                    'method': 'surrounding_text_unchanged',
                    'target_text': select_text,
                    'property_name': property_name,
                    'description': f'Non-target text {property_name} unchanged',
                })
        else:
            # Fallback to property-based comparison for unmapped commands
            spec['verification']['primary'] = {
                'method': 'property_based',
                'description': 'Word COM property comparison with gold',
            }

        return spec

    def generate_for_command(self, command: Dict) -> List[Dict]:
        """
        Generate all test variations for a single command.

        Args:
            command: Command dict with 'idMso' and 'control_type'

        Returns:
            List of generation results
        """
        idMso = command['idMso']
        variations = self.get_variations_for_command(command)

        results = []
        for var in variations:
            if var['name'] == 'on':
                result = self.generate_on_test(idMso, var)
            elif var['name'] == 'off':
                result = self.generate_off_test(idMso, var)
            elif var['name'] == 'apply':
                result = self.generate_apply_test(idMso, var)
            else:
                continue

            results.append(result)
            status = 'OK' if result['success'] else 'FAIL'
            print(f"  {var['name']}: {status} - {result['message']}")

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
        print(f"Test set: content_formatting × core_actionable")
        print(f"Output: {self.output_dir}\n")

        # Step 1: Generate shared input
        print("Step 1: Generating shared input...")
        self.generate_shared_input()
        print()

        # Count skipped commands
        all_commands = self.load_commands.__wrapped__(self) if hasattr(self.load_commands, '__wrapped__') else []
        if not all_commands:
            commands_file = COMMAND_SETS_DIR / 'commands_by_policy_id.json'
            if commands_file.exists():
                with open(commands_file, 'r', encoding='utf-8') as f:
                    all_commands_data = json.load(f)
                    all_commands = [
                        cmd for cmd in all_commands_data
                        if cmd.get('behavior') == 'content_formatting'
                        and cmd.get('scope') == 'core_actionable'
                    ]

        skipped_count = sum(1 for cmd in all_commands if cmd['idMso'] in SKIP_COMMANDS)

        summary = {
            'test_set': 'content_formatting_x_core_actionable',
            'generated_at': datetime.now().isoformat(),
            'total_commands': len(commands) + skipped_count,
            'skipped': skipped_count,
            'skipped_commands': list(SKIP_COMMANDS.keys()),
            'total_test_cases': 0,
            'successful': 0,
            'failed': 0,
            'toggle_commands': 0,
            'button_commands': 0,
            'results': [],
        }

        # Step 2: Generate tests for each command
        print("Step 2: Generating test cases...")
        for cmd in commands:
            idMso = cmd['idMso']
            control_type = cmd.get('control_type', 'button')

            if control_type == 'toggleButton':
                summary['toggle_commands'] += 1
            else:
                summary['button_commands'] += 1

            print(f"\n{idMso} ({control_type})...")

            results = self.generate_for_command(cmd)

            for r in results:
                summary['total_test_cases'] += 1
                if r['success']:
                    summary['successful'] += 1
                else:
                    summary['failed'] += 1
                summary['results'].append(r)

        # Print summary
        print(f"\n=== Generation Complete ===")
        print(f"Total commands: {summary['total_commands']}")
        print(f"  - Skipped: {summary['skipped']}")
        print(f"  - Processed: {len(commands)}")
        print(f"    - toggleButton: {summary['toggle_commands']} (2 tests each)")
        print(f"    - button: {summary['button_commands']} (1 test each)")
        print(f"Total test cases: {summary['total_test_cases']}")
        print(f"Successful: {summary['successful']}")
        print(f"Failed: {summary['failed']}")
        print(f"\nNext step: Run summarize_test_cases.py to categorize test cases")

        return summary


def main():
    """CLI entry point for generating content_formatting × core_actionable tests."""
    import argparse

    parser = argparse.ArgumentParser(
        description='Generate test cases for content_formatting × core_actionable'
    )
    parser.add_argument(
        '--command', '-c',
        help='Generate for a specific command (e.g., Bold)'
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
    generator = ContentFormattingCoreActionableGenerator(output_dir=output_dir)

    if args.command:
        # Generate shared input first for single command mode
        generator.generate_shared_input()
        # Generate for a single command - need to look up control_type
        commands = generator.load_commands()
        cmd = next((c for c in commands if c['idMso'] == args.command), None)

        if cmd:
            print(f"Generating test cases for {args.command} ({cmd.get('control_type', 'button')})...")
            results = generator.generate_for_command(cmd)
            success = sum(1 for r in results if r['success'])
            print(f"\nDone: {success}/{len(results)} successful")
        else:
            # Command not in filtered set, assume it's a button
            print(f"Command {args.command} not found in content_formatting × core_actionable set.")
            print(f"Generating as button (1 test)...")
            cmd = {'idMso': args.command, 'control_type': 'button'}
            results = generator.generate_for_command(cmd)
            success = sum(1 for r in results if r['success'])
            print(f"\nDone: {success}/{len(results)} successful")
    else:
        # Generate for all commands
        generator.generate_all(limit=args.limit)


if __name__ == '__main__':
    main()
