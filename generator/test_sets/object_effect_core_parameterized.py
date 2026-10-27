"""
Test case generator for object_effect × core_parameterized.

Generates test cases for Word commands that apply effects to objects with parameter selection:
- Gallery controls (rotation, position, text wrapping)
- Split buttons with parameter menus

Parameters are loaded from variant definition file:
- object_effect_x_core_parameterized_variants.json
"""

import json
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional
from win32com.client import constants as c

from ..config import (
    COMMAND_SETS_DIR, BENCHMARK_OUTPUT_DIR,
)
from ..execution.setup import ensure_word, finalize_word
from ..execution.execution import DirectExecute
from generator.test_sets.spec_utils import write_spec_if_changed

# Path to variant file
OBJECT_EFFECT_VARIANTS_PATH = Path(__file__).parent.parent.parent / 'instructions' / 'object_effect_x_core_parameterized_variants.json'

# Default input configuration for this generator
DEFAULT_INPUT = {
    'shape_count': 1,  # Default: single shape
}

# Commands to skip (COM API limitations or special requirements)
SKIP_COMMANDS = {
}

# Precondition type mapping (determines which input template to use)
PRECONDITION_TYPE_MAP = {
    # Cross shape with text wrapping options (deep 90-degree concave corners for tight vs through)
    'TextWrapGallery': 'cross_with_text',

    # Default: single shape (rotation, position)
    # ObjectRotateGallery, PicturePositionGallery
}


def get_precondition_type(idMso: str) -> str:
    """
    Determine the precondition type for a command.

    Returns:
        One of: 'single_shape' (default), 'moon_with_text'
    """
    return PRECONDITION_TYPE_MAP.get(idMso, 'single_shape')


def load_object_effect_variants() -> Dict:
    """Load variant definitions from object effect gallery variant file."""
    if not OBJECT_EFFECT_VARIANTS_PATH.exists():
        print(f"[WARNING] Object effect variants file not found: {OBJECT_EFFECT_VARIANTS_PATH}")
        return {}

    with open(OBJECT_EFFECT_VARIANTS_PATH, 'r', encoding='utf-8') as f:
        data = json.load(f)
        return data.get('commands', {})


def extract_parameters(wd, idMso: str) -> List[Dict]:
    """
    Extract all parameters for a command from variant definitions.

    Returns:
        List of parameter dictionaries, each with 'label' and command-specific fields.
    """
    # Load variant definitions
    variants_dict = load_object_effect_variants()

    # Check if this command has defined variants
    if idMso not in variants_dict:
        print(f"    [WARNING] Command '{idMso}' not found in variant definitions")
        return []

    variants = variants_dict[idMso].get('variants', [])
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
        for key in ['angle', 'position', 'horizontal', 'vertical', 'wrap_type', 'flip_cmd', 'direction', 'type',
                    'h_align', 'v_align', 'position_value', 'precondition_type']:
            if key in variant:
                param[key] = variant[key]
        parameters.append(param)

    return parameters


def apply_parameter(wd, idMso: str, param: Dict, shape=None) -> bool:
    """Apply parameter to selected shape."""
    try:
        # If shape is provided directly, use it
        if shape is None:
            sel = wd.Selection

            # Try to get shape from selection
            try:
                shape = sel.ShapeRange(1)
            except:
                print(f"    [ERROR] No shape in selection (Selection.Type={sel.Type})")
                return False

        if idMso == 'ObjectRotateGallery':
            # Handle both rotation and flip operations
            if 'angle' in param:
                # Rotation operation using IncrementRotation (matches Word command behavior)
                angle = param['angle']
                shape.IncrementRotation(angle)
            elif 'flip_cmd' in param:
                # Flip operation
                flip_cmd = param['flip_cmd']
                shape.Flip(flip_cmd)
            else:
                print(f"    [ERROR] No angle or flip_cmd in parameter: {param}")
                return False

        elif idMso == 'PicturePositionGallery':
            # Position shape using alignment-based positioning (matches Word UI gallery)
            # Uses WdShapePosition constants (e.g., wdShapeLeft, wdShapeRight, wdShapeTop, wdShapeBottom)
            # This matches OOXML structure: <wp:positionH relativeFrom="margin"><wp:align>left/center/right</wp:align>

            # Get alignment values from parameter (WdShapePosition constants)
            h_align = param.get('h_align', c.wdShapeCenter)  # Default: center
            v_align = param.get('v_align', c.wdShapeCenter)  # Default: center

            # Set relative positioning base to Margin (matches Word UI gallery behavior)
            shape.RelativeHorizontalPosition = c.wdRelativeHorizontalPositionMargin
            shape.RelativeVerticalPosition = c.wdRelativeVerticalPositionMargin

            # Set alignment using WdShapePosition constants
            # These special constants (e.g., wdShapeLeft=-999998) tell Word to use alignment
            shape.Left = h_align
            shape.Top = v_align

        elif idMso == 'TextWrapGallery':
            # Text wrapping type
            wrap_type = param['wrap_type']
            shape.WrapFormat.Type = wrap_type

        else:
            print(f"    [SKIP] No apply method for {idMso}")
            return False

        return True

    except Exception as e:
        print(f"    [ERROR] Failed to apply parameter: {e}")
        import traceback
        traceback.print_exc()
        return False


class ObjectEffectCoreParameterizedGenerator:
    """Generator for object_effect × core_parameterized test cases."""

    def __init__(self, output_dir: Path = None):
        self.execution = DirectExecute()
        self.output_dir = output_dir or (BENCHMARK_OUTPUT_DIR / 'object_effect_x_core_parameterized')
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.inputs_dir = self.output_dir / 'inputs'
        self.inputs_dir.mkdir(parents=True, exist_ok=True)
        self.input_paths = {}  # Cache for generated input files

    def generate_input_for_command(self, idMso: str) -> Path:
        """Generate input document with appropriate precondition for a command."""
        precondition_type = get_precondition_type(idMso)
        return self.get_or_create_input(ensure_word(), precondition_type)

    def get_or_create_input(self, wd, precondition_type: str) -> Path:
        """Get or create an input document for a given precondition type."""
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

        wd = ensure_word()
        doc = None
        try:
            doc = wd.Documents.Add()
            time.sleep(1.0)  # Wait for document to be ready

            if precondition_type == 'single_shape':
                # Default: Single shape (teardrop) for rotation, flip and positioning
                doc.Shapes.AddShape(160, 100, 100, 100, 100)  # Teardrop
                time.sleep(0.2)

            elif precondition_type == 'cross_with_text':
                # Cross shape with surrounding text for text wrapping
                # Use cross (11) to show visible differences
                # between square/tight/through wrapping modes
                # Cross has large 90-degree concave corners where text can flow through

                # Add dense, continuous text for text wrapping demonstration
                # Need enough text to flow around all sides of the shape
                lorem = (
                    "Lorem ipsum dolor sit amet, consectetur adipiscing elit. "
                    "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. "
                    "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. "
                )
                # Multiple repetitions to ensure text flows around shape on all sides
                for _ in range(5):
                    wd.Selection.TypeText(lorem)
                time.sleep(0.2)

                # Insert cross shape - positioned in center of text area
                # Position (200, 150) to have text on all sides for square wrapping visibility
                # Size 150x150 for good text flow visibility
                shape = doc.Shapes.AddShape(11, 200, 150, 150, 150)  # Cross, centered

                # Make cross arms thinner (default 0.25, higher = thinner arms)
                # Use 0.4 for thinner arms and deeper concave corners
                shape.Adjustments.SetItem(1, 0.4)

                # Set initial wrap to "in front of text" so all wrap modes show visible change
                # wdWrapNone (3) = In front of text - shape floats over text
                # This ensures square/tight/through/behind all produce different results
                shape.WrapFormat.Type = 3  # wdWrapNone (in front of text)
                time.sleep(0.2)

            elif precondition_type == 'moon_with_text':
                # Moon/crescent shape with text — visually distinct from the cross
                # used by tight wrapping, so through and tight test cases differ
                lorem = (
                    "Lorem ipsum dolor sit amet, consectetur adipiscing elit. "
                    "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. "
                    "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. "
                )
                for _ in range(5):
                    wd.Selection.TypeText(lorem)
                time.sleep(0.2)

                # msoShapeMoon = 24
                shape = doc.Shapes.AddShape(24, 180, 120, 180, 180)

                # Adjustment controls crescent thickness (0.5 = half-moon)
                shape.Adjustments.SetItem(1, 0.5)

                # Start with wdWrapNone so applying through is a visible change
                shape.WrapFormat.Type = 3  # wdWrapNone
                time.sleep(0.2)

            elif precondition_type == 'cross_with_text_behind_wrap':
                # Cross shape with behind-text wrapping — so applying "in front of text"
                # (wdWrapNone) is bidirectional with apply_behind (which goes from no-wrap to behind)
                lorem = (
                    "Lorem ipsum dolor sit amet, consectetur adipiscing elit. "
                    "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. "
                    "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. "
                )
                for _ in range(5):
                    wd.Selection.TypeText(lorem)
                time.sleep(0.2)

                shape = doc.Shapes.AddShape(11, 200, 150, 150, 150)  # Cross
                shape.Adjustments.SetItem(1, 0.4)

                # Start with wdWrapBehind so switching to wdWrapNone is bidirectional with apply_behind
                shape.WrapFormat.Type = 5  # wdWrapBehind
                time.sleep(0.2)

            doc.SaveAs(str(input_path.absolute()), FileFormat=16)
            time.sleep(1.0)
            doc.Close(False)
            time.sleep(0.5)

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
        """Load commands classified as object_effect × core_parameterized."""
        commands_file = COMMAND_SETS_DIR / 'commands_by_policy_id.json'

        if not commands_file.exists():
            print(f"Warning: {commands_file} not found.")
            return []

        with open(commands_file, 'r', encoding='utf-8') as f:
            all_commands = json.load(f)

        return [
            cmd for cmd in all_commands
            if cmd.get('behavior') == 'object_effect'
            and cmd.get('scope') == 'core_parameterized'
        ]

    def generate_test_case(self, wd, idMso: str, param: Dict, input_path: Path) -> Dict:
        """Generate a single test case."""
        result = {
            'idMso': idMso,
            'variation': f"apply_{param['label']}",
            'parameter': param['label'],
            'success': False,
            'message': '',
            'paths': {}
        }

        try:
            # Open appropriate input for this command
            doc = wd.Documents.Open(str(input_path.absolute()))
            time.sleep(0.5)  # Wait for document to fully load

            # Debug: Check document state
            print(f"    [DEBUG] After open: Shapes.Count={doc.Shapes.Count}, InlineShapes.Count={doc.InlineShapes.Count}")

            # Check if shape exists
            if doc.Shapes.Count == 0:
                result['message'] = 'No shapes in document'
                doc.Close(False)
                return result

            # Get first shape directly (don't rely on selection)
            shape = doc.Shapes(1)
            time.sleep(0.2)

            # Apply parameter directly to shape
            if not apply_parameter(wd, idMso, param, shape=shape):
                result['message'] = f'Failed to apply parameter'
                doc.Close(False)
                return result

            time.sleep(0.2)

            # Save gold output
            test_dir = self.output_dir / idMso / f"apply_{param['label']}"
            test_dir.mkdir(parents=True, exist_ok=True)

            gold_path = test_dir / 'gold.docx'
            spec_path = test_dir / 'spec.json'

            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
            doc.Close(False)
            time.sleep(0.2)

            # Calculate relative path to input file
            input_filename = input_path.name
            doc_in_ref = f'../../inputs/{input_filename}'

            # Determine precondition type (per-variant override or command default)
            precondition_type = param.get('precondition_type') or get_precondition_type(idMso)

            # Generate spec
            spec = {
                'idMso': idMso,
                'behavior': 'object_effect',
                'scope': 'core_parameterized',
                'variation': f"apply_{param['label']}",
                'generated_at': datetime.now().isoformat(),
                'parameter': param,
                'precondition': {
                    'type': precondition_type,
                    'select_object': 'shape_1',
                },
                'verification': {
                    'primary': {
                        'method': 'shape_property_comparison',
                        'description': 'Word COM shape property comparison',
                    },
                    'invariants': [
                        {
                            'method': 'structure_preserved',
                            'description': 'Document structure preserved (object properties change)',
                        }
                    ],
                },
                'files': {
                    'doc_in': doc_in_ref,
                    'gold': 'gold.docx',
                }
            }

            write_spec_if_changed(spec_path, spec)

            result['success'] = True
            result['paths'] = {
                'doc_in': str(input_path),
                'gold': str(gold_path),
                'spec': str(spec_path)
            }

        except Exception as e:
            result['message'] = f'Error: {str(e)}'
            try:
                doc.Close(False)
            except:
                pass

        return result

    def generate_all(self, limit: int = None, command_filter: str = None) -> Dict:
        """
        Generate test cases for all commands or a specific command.

        Args:
            limit: Limit number of commands to process
            command_filter: If specified, only process this command (idMso)
        """
        commands = self.load_commands()

        # Filter out skipped commands
        skipped_commands = [cmd for cmd in commands if cmd['idMso'] in SKIP_COMMANDS]
        commands = [cmd for cmd in commands if cmd['idMso'] not in SKIP_COMMANDS]

        # Filter by specific command if requested
        if command_filter:
            commands = [cmd for cmd in commands if cmd['idMso'] == command_filter]
            if not commands:
                print(f"[ERROR] Command '{command_filter}' not found in command set")
                return None

        if limit:
            commands = commands[:limit]

        print(f"\n=== Generating test cases ===")
        if command_filter:
            print(f"Command filter: {command_filter}")
        print(f"Total commands loaded: {len(commands) + len(skipped_commands)}")
        print(f"Skipped commands: {len(skipped_commands)}")
        print(f"Processing commands: {len(commands)}")
        print(f"Output: {self.output_dir}\n")

        summary = {
            'test_set': 'object_effect_x_core_parameterized',
            'generated_at': datetime.now().isoformat(),
            'total_commands': len(commands) + len(skipped_commands),
            'skipped': len(skipped_commands),
            'skipped_commands': [cmd['idMso'] for cmd in skipped_commands],
            'successful': 0,
            'failed': 0,
            'results': []
        }

        wd = ensure_word()

        for cmd in commands:
            idMso = cmd['idMso']
            print(f"\n{idMso}...")

            # Generate input document for this command
            input_path = self.generate_input_for_command(idMso)

            # Extract parameters
            params = extract_parameters(wd, idMso)

            if not params:
                print(f"  [SKIP] No parameters extracted")
                summary['failed'] += 1
                summary['results'].append({
                    'idMso': idMso,
                    'success': False,
                    'message': 'No parameters extracted'
                })
                continue

            # Generate test case for each parameter variant
            for param in params:
                # Use per-variant precondition if specified, else command default
                variant_precondition = param.get('precondition_type')
                if variant_precondition:
                    variant_input = self.get_or_create_input(wd, variant_precondition)
                else:
                    variant_input = input_path
                result = self.generate_test_case(wd, idMso, param, variant_input)

                summary['results'].append(result)
                if result['success']:
                    summary['successful'] += 1
                    print(f"  {result['variation']}: OK")
                else:
                    summary['failed'] += 1
                    print(f"  {result['variation']}: FAIL - {result['message']}")

                time.sleep(1.0)

        # Save summary
        summary_path = self.output_dir / 'generation_summary.json'
        with open(summary_path, 'w', encoding='utf-8') as f:
            json.dump(summary, f, indent=2, ensure_ascii=False, sort_keys=True)

        print(f"\n=== Generation Complete ===")
        print(f"Successful: {summary['successful']}")
        print(f"Failed: {summary['failed']}")
        print(f"Skipped: {summary['skipped']}")

        return summary


def main():
    """CLI entry point."""
    import argparse

    parser = argparse.ArgumentParser(
        description='Generate test cases for object_effect × core_parameterized'
    )
    parser.add_argument('--limit', '-l', type=int, help='Limit number of commands')
    parser.add_argument('--output', '-o', help='Output directory')
    parser.add_argument('--command', '-c', help='Generate only for specific command (idMso), e.g., ObjectRotateGallery')

    args = parser.parse_args()

    output_dir = Path(args.output) if args.output else None
    generator = ObjectEffectCoreParameterizedGenerator(output_dir=output_dir)
    generator.generate_all(limit=args.limit, command_filter=args.command)


if __name__ == '__main__':
    main()
