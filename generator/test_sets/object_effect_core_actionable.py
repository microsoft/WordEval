"""
Test case generator for object_effect × core_actionable.

This handles commands that apply effects to objects (shapes, images, etc.)
or manipulate object properties like grouping, alignment, and z-order.
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
    NAMESPACES, TOGGLE_VARIATIONS, ONE_WAY_VARIATIONS
)
from ..execution.setup import ensure_word, finalize_word
from ..execution.execution import DirectExecute
from generator.test_sets.spec_utils import write_spec_if_changed

# Z-order commands (require 3 shapes: back, middle, front)
# Always select the middle shape to show clear z-order changes
SHAPE_ZORDER_COMMANDS = {
    'ObjectBringToFront', 'ObjectSendToBack',
    'ObjectBringForward', 'ObjectSendBackward',
}

# Single shape commands (text wrapping, layout)
SHAPE_COMMANDS = {
    'ObjectBringInFrontOfText', 'ObjectSendBehindText',
}

# Multi-select shape commands (grouping, alignment, distribution)
SHAPE_MULTI_SELECT_COMMANDS = {
    'ObjectsGroup',
    'ObjectsAlignLeftSmart', 'ObjectsAlignRightSmart',
    'ObjectsAlignTopSmart', 'ObjectsAlignBottomSmart',
    'ObjectsAlignCenterHorizontalSmart', 'ObjectsAlignMiddleVerticalSmart',
    'AlignDistributeHorizontally', 'AlignDistributeVertically',
}

# Alignment commands that need post-execution position correction
# Word has a bug where shapes loaded from saved files don't align to exact positions
# We manually correct positions after alignment to create semantically correct gold files
#
# For edge alignments (left/right/top/bottom), we align to the original extreme edge
# before Word's command was executed. This requires reading doc_in positions.
ALIGNMENT_COMMANDS_NEED_CORRECTION = {
    'ObjectsAlignLeftSmart': 'left_edge',       # All left edges to min original left
    'ObjectsAlignRightSmart': 'right_edge',     # All right edges to max original right
    'ObjectsAlignTopSmart': 'top_edge',         # All top edges to min original top
    'ObjectsAlignBottomSmart': 'bottom_edge',   # All bottom edges to max original bottom
    'ObjectsAlignCenterHorizontalSmart': 'center_h',  # All centers to average original center
    'ObjectsAlignMiddleVerticalSmart': 'center_v',    # All centers to average original center
}


def get_original_shape_positions(doc) -> list:
    """
    Get shape positions from an open document.
    Call this BEFORE executing alignment command to capture original positions.
    """
    shapes = []
    for i in range(1, doc.Shapes.Count + 1):
        s = doc.Shapes(i)
        shapes.append({
            'left': s.Left,
            'top': s.Top,
            'width': s.Width,
            'height': s.Height,
            'right': s.Left + s.Width,
            'bottom': s.Top + s.Height,
            'center_h': s.Left + s.Width / 2,
            'center_v': s.Top + s.Height / 2,
        })
    return shapes


def correct_alignment_positions(doc, idMso: str, original_shapes: list = None) -> bool:
    """
    Correct shape positions after alignment command to fix Word's precision bug.

    Word has a bug where shapes loaded from saved files don't align to exact same
    position - they end up with ~0.25-0.5pt variations. This function uses the
    original positions (captured before alignment) to set shapes to the correct position.

    Args:
        doc: Word document object (after alignment command executed)
        idMso: Command identifier
        original_shapes: List of original shape positions (from get_original_shape_positions)

    Returns:
        True if correction was applied, False if not needed
    """
    if idMso not in ALIGNMENT_COMMANDS_NEED_CORRECTION:
        return False

    align_type = ALIGNMENT_COMMANDS_NEED_CORRECTION[idMso]

    if doc.Shapes.Count < 2:
        return False

    if not original_shapes:
        return False

    # Calculate target position based on alignment type
    if align_type == 'left_edge':
        # Align left edges to the minimum (leftmost) original left
        target_left = min(s['left'] for s in original_shapes)
        for i in range(1, doc.Shapes.Count + 1):
            doc.Shapes(i).Left = target_left

    elif align_type == 'right_edge':
        # Align right edges to the maximum (rightmost) original right
        target_right = max(s['right'] for s in original_shapes)
        for i in range(1, doc.Shapes.Count + 1):
            doc.Shapes(i).Left = target_right - doc.Shapes(i).Width

    elif align_type == 'top_edge':
        # Align top edges to the minimum (topmost) original top
        target_top = min(s['top'] for s in original_shapes)
        for i in range(1, doc.Shapes.Count + 1):
            doc.Shapes(i).Top = target_top

    elif align_type == 'bottom_edge':
        # Align bottom edges to the maximum (bottommost) original bottom
        target_bottom = max(s['bottom'] for s in original_shapes)
        for i in range(1, doc.Shapes.Count + 1):
            doc.Shapes(i).Top = target_bottom - doc.Shapes(i).Height

    elif align_type == 'center_h':
        # Align horizontal centers to the average original center
        target_center = sum(s['center_h'] for s in original_shapes) / len(original_shapes)
        for i in range(1, doc.Shapes.Count + 1):
            doc.Shapes(i).Left = target_center - doc.Shapes(i).Width / 2

    elif align_type == 'center_v':
        # Align vertical centers to the average original center
        target_center = sum(s['center_v'] for s in original_shapes) / len(original_shapes)
        for i in range(1, doc.Shapes.Count + 1):
            doc.Shapes(i).Top = target_center - doc.Shapes(i).Height / 2

    return True

# Grouped shape commands
GROUPED_SHAPE_COMMANDS = {
    'ObjectsUngroup',  # Needs grouped shape
}

# Uneven shape commands
UNEVEN_SHAPE_COMMANDS = {
    'AlignDistributeHorizontally',  # Needs grouped shape
    'AlignDistributeVertically',
}


def get_precondition_type(idMso: str) -> str:
    """
    Determine the precondition type for a command.

    Returns:
        One of: 'three_shapes', 'grouped_shape'
    """
    if idMso in GROUPED_SHAPE_COMMANDS:
        return 'grouped_shape'  # Needs pre-grouped shapes
    elif idMso in UNEVEN_SHAPE_COMMANDS:
        return 'three_shapes_uneven'
    else:
        return 'three_shapes'  # Default: 3 shapes covers all other cases


def setup_precondition(doc, wd, precondition_type: str, full_text: str = "Alpha TARGET Omega."):
    """
    Setup document based on precondition type.

    Args:
        doc: Word document object
        wd: Word application object
        precondition_type: Type of precondition to setup
        full_text: Unused (kept for backward compatibility)
    """
    if precondition_type == 'three_shapes':
        # Insert three shapes (default for all commands)
        # Back, Middle, Front - overlapping to show clear z-order changes
        doc.Shapes.AddShape(1, 50, 50, 100, 100)   # Rectangle 1 (back)
        doc.Shapes.AddShape(1, 100, 100, 100, 100) # Rectangle 2 (middle) - overlapping
        doc.Shapes.AddShape(1, 150, 150, 100, 100) # Rectangle 3 (front)
    
    if precondition_type == 'three_shapes_uneven':
        # Insert three shapes (default for all commands)
        # Back, Middle, Front - overlapping to show clear z-order changes
        doc.Shapes.AddShape(1, 50, 50, 100, 100)   # Rectangle 1 (back)
        doc.Shapes.AddShape(1, 70, 70, 100, 100) # Rectangle 2 (middle) - overlapping
        doc.Shapes.AddShape(1, 150, 150, 100, 100) # Rectangle 3 (front)

    elif precondition_type == 'grouped_shape':
        # Insert and group two shapes (for ObjectsUngroup)
        doc.Shapes.AddShape(1, 50, 50, 100, 100)
        doc.Shapes.AddShape(1, 200, 50, 100, 100)
        shape_range = doc.Shapes.Range([1, 2])
        shape_range.Group()


def select_for_command(doc, wd, execution, precondition_type: str, select_text: str = "TARGET", idMso: str = None) -> bool:
    """
    Select appropriate object(s) based on precondition type and command.

    Args:
        doc: Word document object
        wd: Word application object
        execution: Execution helper object
        precondition_type: Type of precondition
        select_text: Unused (kept for backward compatibility)
        idMso: Command identifier (used to determine selection strategy)

    Returns:
        True if selection successful
    """
    if precondition_type == 'three_shapes':
        if doc.Shapes.Count < 3:
            return False

        # Z-order commands: select middle shape (Shape 2)
        if idMso in SHAPE_ZORDER_COMMANDS:
            doc.Shapes(2).Select()
            return True

        # Multi-select commands (grouping, alignment): select all 3 shapes
        elif idMso in SHAPE_MULTI_SELECT_COMMANDS:
            doc.Shapes(1).Select()
            doc.Shapes(2).Select(False)  # Add to selection
            doc.Shapes(3).Select(False)  # Add to selection
            return True

        # Single shape commands (text wrapping, layout): select Shape 2
        elif idMso in SHAPE_COMMANDS:
            doc.Shapes(2).Select()
            return True

        # Default: select middle shape
        else:
            doc.Shapes(2).Select()
            return True

    elif precondition_type == 'three_shapes_uneven':
        # Multi-select commands (grouping, alignment): select all 3 shapes
        if idMso in SHAPE_MULTI_SELECT_COMMANDS:
            doc.Shapes(1).Select()
            doc.Shapes(2).Select(False)  # Add to selection
            doc.Shapes(3).Select(False)  # Add to selection
            return True

    elif precondition_type == 'grouped_shape':
        # Select the grouped shape (for ObjectsUngroup)
        if doc.Shapes.Count > 0:
            doc.Shapes(1).Select()
            return True
        return False

    return False


# Commands that need special handling or are not fully automatable
# Based on execution results, some commands fail due to:
# - Not supported in current Word version
# - Require specific language/region settings
# - Require special document state
SPECIAL_HANDLING_COMMANDS = {
    # ObjectsRegroup requires ungroup history which is lost on document save
    # This command cannot be automated in a save/load workflow
    'ObjectsRegroup': 'requires_ungroup_history_in_session',

    # Text wrapping commands require text content to show visible effect
    # Without text, wrapping change has no visible impact (0 revisions)
    'ObjectBringInFrontOfText': 'requires_text_content_for_visible_effect',
    'ObjectSendBehindText': 'requires_text_content_for_visible_effect',
}


# Semantic verification methods for commands
# These use the new semantic verification classes instead of strict XML comparison
SEMANTIC_VERIFICATION_METHODS = {
    # Alignment commands - check shape positions are aligned
    'ObjectsAlignLeftSmart': {
        'method': 'alignment_verification',
        'params': {'alignment_type': 'left'},
        'description': 'Shapes aligned by left edge',
    },
    'ObjectsAlignRightSmart': {
        'method': 'alignment_verification',
        'params': {'alignment_type': 'right'},
        'description': 'Shapes aligned by right edge',
    },
    'ObjectsAlignTopSmart': {
        'method': 'alignment_verification',
        'params': {'alignment_type': 'top'},
        'description': 'Shapes aligned by top edge',
    },
    'ObjectsAlignBottomSmart': {
        'method': 'alignment_verification',
        'params': {'alignment_type': 'bottom'},
        'description': 'Shapes aligned by bottom edge',
    },
    'ObjectsAlignCenterHorizontalSmart': {
        'method': 'alignment_verification',
        'params': {'alignment_type': 'center_h'},
        'description': 'Shapes aligned by horizontal center',
    },
    'ObjectsAlignMiddleVerticalSmart': {
        'method': 'alignment_verification',
        'params': {'alignment_type': 'center_v'},
        'description': 'Shapes aligned by vertical center',
    },

    # Distribute commands - check even spacing between shapes
    'AlignDistributeHorizontally': {
        'method': 'distribute_verification',
        'params': {'direction': 'horizontal'},
        'description': 'Shapes evenly distributed horizontally',
    },
    'AlignDistributeVertically': {
        'method': 'distribute_verification',
        'params': {'direction': 'vertical'},
        'description': 'Shapes evenly distributed vertically',
    },

    # Group commands - check group structure
    'ObjectsGroup': {
        'method': 'group_structure_verification',
        'params': {},
        'description': 'Group structure matches (count and shapes)',
    },

    # Ungroup command - check shapes ungrouped
    'ObjectsUngroup': {
        'method': 'ungroup_verification',
        'params': {},
        'description': 'Shapes ungrouped successfully',
    },
}


class ObjectEffectCoreActionableGenerator:
    """
    Generator for object_effect × core_actionable test cases.

    After classification fix (2025-12-19), this test set includes 20 commands:
    - Shape grouping commands (ObjectsGroup, ObjectsUngroup, ObjectsRegroup)
    - Shape z-order commands (ObjectBringToFront, ObjectSendToBack, etc.)
    - Shape alignment commands (ObjectsAlignLeftSmart, etc.)
    - Shape distribution commands (AlignDistributeHorizontally, AlignDistributeVertically)
    - Shape text wrapping (ObjectBringInFrontOfText, ObjectSendBehindText)

    Removed commands (reclassified):
    - SetDefaultObjectLayout → settings_effect × core_actionable (sets defaults, not object manipulation)

    All commands are:
    - BEHAVIOR: object_effect (verify via object property changes in XML)
    - SCOPE: core_actionable (setup: precondition + select, execute: direct)

    Test strategy:
    - toggleButton: 2 tests (on, off)
    - button: 1 test (apply)

    Key characteristics:
    - Requires specific preconditions (shapes, grouped shapes, etc.)
    - Each command may need different input document setup
    - No text content manipulation (those are content_formatting)
    """

    def __init__(self, output_dir: Path = None):
        self.execution = DirectExecute()
        self.output_dir = output_dir or (BENCHMARK_OUTPUT_DIR / 'object_effect_x_core_actionable')
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.inputs_dir = self.output_dir / 'inputs'
        self.inputs_dir.mkdir(parents=True, exist_ok=True)

    def generate_input_for_command(self, idMso: str) -> Path:
        """
        Generate input document with appropriate precondition for a command.

        Args:
            idMso: Command identifier

        Returns:
            Path to the generated input document
        """
        precondition_type = get_precondition_type(idMso)
        input_filename = f'doc_in_{precondition_type}.docx'
        input_path = self.inputs_dir / input_filename

        if input_path.exists():
            return input_path

        print(f"  Generating input: {input_filename}")

        wd = ensure_word()
        try:
            doc = wd.Documents.Add()

            # Setup precondition
            full_text = "Alpha TARGET Omega."
            setup_precondition(doc, wd, precondition_type, full_text)

            doc.SaveAs(str(input_path.absolute()), FileFormat=16)
            doc.Close(False)
            time.sleep(0.2)

            return input_path

        except Exception as e:
            print(f"  Error creating input: {e}")
            raise

    def load_commands(self) -> List[Dict]:
        """Load commands classified as object_effect × core_actionable."""
        commands_file = COMMAND_SETS_DIR / 'commands_by_policy_id.json'

        if not commands_file.exists():
            print(f"Warning: {commands_file} not found. Using empty command list.")
            return []

        with open(commands_file, 'r', encoding='utf-8') as f:
            all_commands = json.load(f)

        # Filter for our test set
        filtered = [
            cmd for cmd in all_commands
            if cmd.get('behavior') == 'object_effect'
            and cmd.get('scope') == 'core_actionable'
        ]

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

    def generate_on_test(self, idMso: str, variation: Dict, input_path: Path) -> Dict:
        """
        Generate 'on' test: apply effect to object.

        - doc_in: input with precondition setup
        - gold: result after applying command
        """
        test_dir = self.output_dir / idMso / variation['name']
        test_dir.mkdir(parents=True, exist_ok=True)

        gold_path = test_dir / 'gold.docx'
        spec_path = test_dir / 'spec.json'

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

        # Skip check is now handled in generate_for_command()
        # No need to check again here

        wd = ensure_word()
        try:
            # Open input, apply command, save as gold
            doc = wd.Documents.Open(str(input_path.absolute()))

            # Capture original shape positions BEFORE alignment (needed for correction)
            original_positions = get_original_shape_positions(doc) if idMso in ALIGNMENT_COMMANDS_NEED_CORRECTION else None

            # Select the appropriate object(s)
            precondition_type = get_precondition_type(idMso)
            if not select_for_command(doc, wd, self.execution, precondition_type, "TARGET", idMso):
                result['message'] = f"Failed to select object(s) for: {idMso}"
                doc.Close(False)
                return result

            # Execute the command
            if not self.execution.execute(wd, idMso):
                result['message'] = f'Failed to execute command: {idMso}'
                doc.Close(False)
                return result

            # Apply post-command corrections for alignment commands
            # This fixes Word's precision bug where shapes don't align to exact positions
            if correct_alignment_positions(doc, idMso, original_positions):
                result['message'] = 'Test case generated with alignment correction'

            # Save as gold
            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
            doc.Close(False)
            time.sleep(0.2)

            # Generate spec.json
            spec = self._generate_spec(idMso, variation, input_path, is_off_test=False)
            write_spec_if_changed(spec_path, spec)

            result['success'] = True
            result['message'] = 'Test case generated successfully'

        except Exception as e:
            result['message'] = f'Error: {str(e)}'
            import traceback
            traceback.print_exc()

        return result

    def generate_off_test(self, idMso: str, variation: Dict, input_path: Path) -> Dict:
        """
        Generate 'off' test: remove effect from object.

        - doc_in: input with effect already applied
        - gold: result after removing effect (back to precondition state)
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

        # Skip check is now handled in generate_for_command()
        # No need to check again here

        wd = ensure_word()
        try:
            # Step 1: Create doc_in (with effect applied)
            doc = wd.Documents.Open(str(input_path.absolute()))

            precondition_type = get_precondition_type(idMso)
            if not select_for_command(doc, wd, self.execution, precondition_type, "TARGET", idMso):
                result['message'] = f"Failed to select object(s) for: {idMso}"
                doc.Close(False)
                return result

            if not self.execution.execute(wd, idMso):
                result['message'] = f'Failed to execute command: {idMso}'
                doc.Close(False)
                return result

            doc.SaveAs(str(doc_in_path.absolute()), FileFormat=16)
            doc.Close(False)
            time.sleep(0.2)

            # Step 2: gold = precondition state (copy of original input)
            shutil.copy(input_path, gold_path)

            # Generate spec.json
            spec = self._generate_spec(idMso, variation, input_path, is_off_test=True)
            write_spec_if_changed(spec_path, spec)

            result['success'] = True
            result['message'] = 'Test case generated successfully'

        except Exception as e:
            result['message'] = f'Error: {str(e)}'
            import traceback
            traceback.print_exc()

        return result

    def generate_apply_test(self, idMso: str, variation: Dict, input_path: Path) -> Dict:
        """
        Generate 'apply' test for one-way (button) commands.
        Same logic as 'on' test.
        """
        return self.generate_on_test(idMso, variation, input_path)

    def _generate_spec(
        self,
        idMso: str,
        variation: Dict,
        input_path: Path,
        is_off_test: bool = False
    ) -> Dict:
        """Generate the spec.json for a test case."""

        semantic_info = SEMANTIC_VERIFICATION_METHODS.get(idMso, {})
        variation_name = variation['name']
        precondition_type = get_precondition_type(idMso)

        # Calculate relative path to input
        if is_off_test:
            doc_in_ref = 'doc_in.docx'  # Local file for 'off' test
        else:
            doc_in_ref = f'../../inputs/doc_in_{precondition_type}.docx'

        spec = {
            'idMso': idMso,
            'behavior': 'object_effect',
            'scope': 'core_actionable',
            'variation': variation_name,
            'generated_at': datetime.now().isoformat(),

            'precondition': {
                'type': precondition_type,
                'state': variation.get('precondition', 'base'),
            },

            'files': {
                'doc_in': doc_in_ref,
                'gold': 'gold.docx',
            }
        }

        # Use semantic verification if available, otherwise fall back to shape comparison
        if semantic_info and not is_off_test:
            spec['verification'] = {
                'primary': {
                    'method': semantic_info['method'],
                    'params': semantic_info.get('params', {}),
                    'description': semantic_info.get('description', ''),
                },
                'invariants': [
                    {
                        'method': 'structure_preserved',
                        'description': 'Document structure preserved (object properties may change)',
                    }
                ],
            }
        else:
            # Fall back to shape property comparison for z-order commands and off tests
            spec['verification'] = {
                'primary': {
                    'method': 'shape_property_comparison',
                    'description': 'Word COM shape property comparison',
                },
                'invariants': [
                    {
                        'method': 'structure_preserved',
                        'description': 'Document structure preserved (object properties may change)',
                    }
                ],
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

        # Check if command should be skipped (before generating input)
        if idMso in SPECIAL_HANDLING_COMMANDS:
            # Kill Word completely to ensure clean state for next command
            try:
                finalize_word()
            except:
                pass

            # Kill any remaining Word processes
            import subprocess
            try:
                subprocess.run(['taskkill', '/F', '/IM', 'WINWORD.EXE'],
                             capture_output=True, timeout=5)
                time.sleep(1)
            except:
                pass

            # Generate skip results without input file
            results = []
            for var in variations:
                result = {
                    'idMso': idMso,
                    'variation': var['name'],
                    'success': False,
                    'message': f"Skipped: {SPECIAL_HANDLING_COMMANDS[idMso]} - requires special handling",
                    'skipped': True,
                    'paths': {}
                }
                results.append(result)
                print(f"  {var['name']}: SKIP - {SPECIAL_HANDLING_COMMANDS[idMso]}")
            return results

        # Generate input document for this command
        input_path = self.generate_input_for_command(idMso)

        results = []
        for var in variations:
            if var['name'] == 'on':
                result = self.generate_on_test(idMso, var, input_path)
            elif var['name'] == 'off':
                result = self.generate_off_test(idMso, var, input_path)
            elif var['name'] == 'apply':
                result = self.generate_apply_test(idMso, var, input_path)
            else:
                continue

            results.append(result)
            if result.get('skipped'):
                status = 'SKIP'
            elif result['success']:
                status = 'OK'
            else:
                status = 'FAIL'
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
        print(f"Test set: object_effect × core_actionable")
        print(f"Output: {self.output_dir}\n")

        summary = {
            'test_set': 'object_effect_x_core_actionable',
            'generated_at': datetime.now().isoformat(),
            'total_commands': len(commands),
            'total_test_cases': 0,
            'successful': 0,
            'failed': 0,
            'skipped': 0,
            'toggle_commands': 0,
            'button_commands': 0,
            'results': [],
        }

        # Generate tests for each command
        print("Generating test cases...")
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

        # Print detailed results table
        print(f"\n=== Generation Complete ===")
        print(f"\nResults\n")
        print("| Status     | Count |")
        print("|-----------|-------|")
        print(f"| Successful| {summary['successful']:<5} |")
        print(f"| Skipped   | {summary.get('skipped', 0):<5} |")
        print(f"| Failed    | {summary['failed']:<5} |")

        # Print generated test cases table (including skipped)
        successful_results = [r for r in summary['results'] if r['success']]
        skipped_results = [r for r in summary['results'] if r.get('skipped')]
        all_generated = successful_results + skipped_results

        if all_generated:
            print(f"\nGenerated Test Cases ({len(all_generated)})\n")
            print("| Command              | Variation |")
            print("|----------------------|-----------|")
            for r in all_generated:
                idMso = r['idMso']
                variation = r['variation']
                skipped = "[SKIPPED]" if r.get('skipped') else ""
                print(f"| {idMso:<20} | {variation:<9} {skipped}")

        # Print output structure
        print(f"\nOutput Structure\n")
        print(f"{self.output_dir.name}/")

        # Get unique input files
        input_files = set()
        for r in summary['results']:
            precondition_type = get_precondition_type(r['idMso'])
            input_files.add(f"doc_in_{precondition_type}.docx")

        for input_file in sorted(input_files):
            print(f"  inputs/{input_file}")

        for r in successful_results:
            idMso = r['idMso']
            variation = r['variation']
            print(f"  {idMso}/{variation}/{{gold.docx, spec.json}}")

        if skipped_results:
            for r in skipped_results:
                idMso = r['idMso']
                variation = r['variation']
                print(f"  {idMso}/{variation}/{{spec.json}} [SKIPPED]")

        print(f"\nTotal commands: {summary['total_commands']}")
        print(f"  - toggleButton: {summary['toggle_commands']} (2 tests each)")
        print(f"  - button: {summary['button_commands']} (1 test each)")
        print(f"Total test cases: {summary['total_test_cases']}")
        print(f"Summary saved to: {summary_path}")

        return summary


def main():
    """CLI entry point for generating object_effect × core_actionable tests."""
    import argparse

    parser = argparse.ArgumentParser(
        description='Generate test cases for object_effect × core_actionable'
    )
    parser.add_argument(
        '--command', '-c',
        help='Generate for a specific command (e.g., ObjectsGroup)'
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
    generator = ObjectEffectCoreActionableGenerator(output_dir=output_dir)

    if args.command:
        # Generate for a single command
        commands = generator.load_commands()
        cmd = next((c for c in commands if c['idMso'] == args.command), None)

        if cmd:
            print(f"Generating test cases for {args.command} ({cmd.get('control_type', 'button')})...")
            results = generator.generate_for_command(cmd)
            success = sum(1 for r in results if r['success'])
            print(f"\nDone: {success}/{len(results)} successful")
        else:
            # Command not in filtered set, assume it's a button
            print(f"Command {args.command} not found in object_effect × core_actionable set.")
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
