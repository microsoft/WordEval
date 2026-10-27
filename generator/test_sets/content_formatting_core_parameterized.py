"""
Test case generator for content_formatting × core_parameterized.

Generates test cases for Word commands that require parameter selection:
- Gallery controls (color pickers, style galleries, text effects)
- ComboBox controls (Font, FontSize)

Parameters are loaded from variant definition files:
- List galleries: Load from content_formatting_x_core_parameterized_list_variants.json (auto-discovered)
- Other galleries: Load from content_formatting_x_core_parameterized_text_variants.json (auto-discovered + manual)
"""

import json
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from ..config import (
    COMMAND_SETS_DIR, BENCHMARK_OUTPUT_DIR,
    TOGGLE_VARIATIONS, ONE_WAY_VARIATIONS
)
from ..execution.setup import ensure_word, finalize_word
from ..execution.execution import DirectExecute
from generator.test_sets.spec_utils import write_spec_if_changed

# Paths to variant files
LIST_GALLERY_VARIANTS_PATH = Path(__file__).parent.parent.parent / 'instructions' / 'content_formatting_x_core_parameterized_list_variants.json'
GALLERY_VARIANTS_PATH = Path(__file__).parent.parent.parent / 'instructions' / 'content_formatting_x_core_parameterized_text_variants.json'

# Default input configuration for this generator
DEFAULT_INPUT = {
    'full_text': 'Alpha Target Omega.',
    'select_text': 'Target',
    'filename': 'doc_in_single_line.docx',
}


# Commands that need full document selection
COMMANDS_NEED_FULL_SELECTION = {
    'BulletsGalleryWord',
    'NumberingGalleryWord',
    'MultilevelListGallery',
    'LineSpacingGallery',
    # Note: ListLevelGallery is NOT here - it selects TARGET paragraph specifically,
    'LigatureGalleryWord',
    'StylisticSetsMenuWord',
    'QuickStylesGallery',  # Heading/Title styles typically apply to entire line
    'ParagraphSpacingBefore',  # Spinner control — applies to all paragraphs
    'ParagraphSpacingAfter',   # Spinner control — applies to all paragraphs
}

# Commands to skip (COM API limitations)
SKIP_COMMANDS = {
    'TextEffectShadowGallery': 'COM API limitation - Shadow returns int',
    'TextEffectGlowGallery': 'COM API limitation - GlowFormat.Preset missing',
    'TextGlowColorPicker': 'Duplicate of TextEffectGlowGallery',
    'GalleryAllShapesAndCanvas': 'Shape gallery (not text formatting)',
    'TextEffectsGallery': 'Combined effect presets',
    'DictationMenu': 'Dictation menu (not formatting)',
    'NumberStyleGalleryWord': 'COM API limitation - Font.NumberStyle not available',
    # Spinner controls (control_type=control) — reclassified from generic_control;
    # not yet implemented (need COM property setting, not ExecuteMso)
    'ParagraphIndentLeft': 'spinner_control — set ParagraphFormat.LeftIndent via COM',
    'ParagraphIndentRight': 'spinner_control — set ParagraphFormat.RightIndent via COM',
}

# Precondition type mapping (determines which input template to use)
PRECONDITION_TYPE_MAP = {
    # List commands need multiple paragraphs to show list formatting effect
    'BulletsGalleryWord': 'paragraphs',
    'NumberingGalleryWord': 'paragraphs',

    # Line spacing needs multiple paragraphs to show spacing effect
    'LineSpacingGallery': 'paragraphs',

    # Multilevel list command needs hierarchical paragraph structure
    'MultilevelListGallery': 'multilevel_list',

    # List level command needs existing multilevel list
    'ListLevelGallery': 'list_with_levels',

    # OpenType font features need specific setup for visible effects
    'LigatureGalleryWord': 'opentype_font_ligatures',  # Ligatures disabled in doc_in
    'StylisticSetsMenuWord': 'opentype_font_stylistic',  # Text with Q/R/K/J for flourishes

    # QuickStyles (Heading/Title) typically apply to entire line/paragraph
    'QuickStylesGallery': 'heading_line',

    # Underline color only has visible meaning when the target already has an underline.
    'UnderlineColorPicker': 'underlined_text',

    # Paragraph spacing needs multiple paragraphs with preset spacing to show change
    'ParagraphSpacingBefore': 'paragraph_spacing_before',
    'ParagraphSpacingAfter': 'paragraph_spacing_after',

    # Drop cap reformats the first letter of the first paragraph; needs a single
    # long-enough paragraph for the styled letter to be visually distinct from body text.
    'DropCapInsertGallery': 'with_paragraph',

    # Default: single line with Target (doc_in_single_line.docx)
    # Text formatting commands: Font, FontSize, colors, effects, borders, etc.
    # All commands including ChangeCaseGallery select "Target" (mixed case word)
}

# Special variation-specific preconditions (for cases where specific variants need different input)
# Key format: "idMso/variation" -> precondition_type
VARIATION_PRECONDITION_MAP = {
    # Capitalize Each Word: needs multiple words + select ALL to show "each word" behavior
    # "alpha target omega" → "Alpha Target Omega" (every word capitalized)
    'ChangeCaseGallery/apply_Capitalize_Each_Word': 'lowercase_text_select_all',
    # Sentence case: select ALL text to show sentence case effect
    # "alpha target omega" → "Alpha target omega" (only first word of sentence capitalized)
    'ChangeCaseGallery/apply_Sentence_case': 'lowercase_text_select_all',
    # Toggle case: Input "Alpha Target Omega.", select "Target" → becomes "tARGET"
    # Use single_line (default) since we only toggle one word
    # Calibri is Word's default font - need different starting font to show change
    'Font/apply_Calibri': 'times_new_roman_text',
}

# Custom selection text for specific variants (overrides DEFAULT_INPUT['select_text'])
# Key format: "idMso/variation" -> select_text
VARIATION_SELECT_TEXT_MAP = {
    # Words underline only underlines words, not spaces between them
    # Single underline underlines everything including spaces
    # Select "Target Omega" (two words with space) to show the difference
    'UnderlineGallery/apply_Words': 'Target Omega',
    'UnderlineGallery/apply_Single': 'Target Omega',
}


def get_precondition_type(idMso: str, variation: str = None) -> str:
    """
    Determine the precondition type for a command (and optionally a specific variation).

    Args:
        idMso: Command identifier
        variation: Optional variation name (e.g., 'apply_Capitalize_Each_Word')

    Returns:
        One of: 'single_line' (default), 'paragraphs', 'multilevel_list', 'list_with_levels', 'opentype_font', 'lowercase_text'
    """
    # Check for variation-specific precondition first
    if variation:
        variation_key = f'{idMso}/{variation}'
        if variation_key in VARIATION_PRECONDITION_MAP:
            return VARIATION_PRECONDITION_MAP[variation_key]

    return PRECONDITION_TYPE_MAP.get(idMso, 'single_line')


def load_gallery_variants() -> Dict:
    """Load variant definitions from both list gallery and general gallery variant files."""
    all_variants = {}

    # Load list gallery variants (auto-discovered)
    if LIST_GALLERY_VARIANTS_PATH.exists():
        with open(LIST_GALLERY_VARIANTS_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
            all_variants.update(data.get('commands', {}))
    else:
        print(f"[WARNING] List gallery variants file not found: {LIST_GALLERY_VARIANTS_PATH}")

    # Load other gallery variants (manually-defined)
    if GALLERY_VARIANTS_PATH.exists():
        with open(GALLERY_VARIANTS_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
            all_variants.update(data.get('commands', {}))
    else:
        print(f"[WARNING] Gallery variants file not found: {GALLERY_VARIANTS_PATH}")

    return all_variants


def extract_parameters(wd, idMso: str) -> List[Dict]:
    """
    Extract all parameters for a command from variant definitions.

    Commands are defined in two variant files:
    - content_formatting_x_core_parameterized_list_variants.json: Auto-discovered list gallery variants
    - content_formatting_x_core_parameterized_text_variants.json: Auto-discovered + manual gallery variants

    Returns:
        List of parameter dictionaries, each with 'label' and command-specific fields.
    """
    # Load variant definitions
    gallery_variants = load_gallery_variants()

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
        for key in ['template_index', 'level', 'target_paragraph', 'value', 'style', 'weight', 'index', 'type', 'name', 'fore_color', 'back_color', 'position']:
            if key in variant:
                param[key] = variant[key]
        parameters.append(param)

    return parameters


def apply_parameter(wd, idMso: str, param: Dict) -> bool:
    """Apply parameter to selection."""
    try:
        sel = wd.Selection

        if idMso == 'Font':
            sel.Font.Name = param['name']
        elif idMso == 'FontSize':
            sel.Font.Size = param['value']
        elif idMso == 'QuickStylesGallery':
            sel.Style = param['style']
        elif idMso == 'BulletsGalleryWord':
            template = wd.ListGalleries(1).ListTemplates(param['template_index'])
            sel.Range.ListFormat.ApplyListTemplateWithLevel(
                ListTemplate=template,
                ContinuePreviousList=False,
                ApplyTo=0,  # Apply to selection
                DefaultListBehavior=1
            )
        elif idMso == 'NumberingGalleryWord':
            template = wd.ListGalleries(2).ListTemplates(param['template_index'])
            sel.Range.ListFormat.ApplyListTemplateWithLevel(
                ListTemplate=template,
                ContinuePreviousList=False,
                ApplyTo=0,  # Apply to selection
                DefaultListBehavior=1
            )
        elif idMso == 'MultilevelListGallery':
            template = wd.ListGalleries(3).ListTemplates(param['template_index'])
            sel.Range.ListFormat.ApplyListTemplateWithLevel(
                ListTemplate=template,
                ContinuePreviousList=False,
                ApplyTo=0,  # Apply to selection
                DefaultListBehavior=0
            )
        elif idMso == 'ChangeCaseGallery':
            sel.Range.Case = param['value']
        elif idMso == 'FontColorPicker':
            sel.Font.Color = param['value']
        elif idMso == 'TextHighlightColorPicker':
            sel.Range.HighlightColorIndex = param['index']
        elif idMso == 'ShadingColorPicker':
            sel.Shading.BackgroundPatternColor = param['value']
        elif idMso == 'UnderlineColorPicker':
            sel.Font.UnderlineColor = param['value']
        elif idMso == 'UnderlineGallery':
            sel.Font.Underline = param['value']
        elif idMso == 'BordersSelectionGallery':
            # Word border indices: -1=Top, -2=Left, -3=Bottom, -4=Right
            # Variant index maps: 0=top_border→-1, 1=bottom_border→-3
            border_index_map = {0: -1, 1: -3}  # top, bottom
            border_idx = border_index_map.get(param.get('index'), -1)
            # Use Paragraphs(1).Borders to apply PARAGRAPH border (top/bottom line)
            # NOT sel.Borders which applies CHARACTER border (box around text)
            sel.Paragraphs(1).Borders(border_idx).LineStyle = 1
        elif idMso == 'ListLevelGallery':
            if sel.Range.ListFormat.ListType == 0:
                template = wd.ListGalleries(3).ListTemplates(1)
                sel.Range.ListFormat.ApplyListTemplateWithLevel(
                    ListTemplate=template,
                    ContinuePreviousList=False
                )
            sel.Range.ListFormat.ListLevelNumber = param['level']
        elif idMso == 'LineSpacingGallery':
            sel.ParagraphFormat.LineSpacingRule = 0
            sel.ParagraphFormat.LineSpacing = param['value'] * 12
        elif idMso == 'ParagraphSpacingBefore':
            doc = wd.ActiveDocument
            for i in range(doc.Paragraphs.Count):
                doc.Paragraphs(i + 1).Format.SpaceBefore = param['value']
        elif idMso == 'ParagraphSpacingAfter':
            doc = wd.ActiveDocument
            for i in range(doc.Paragraphs.Count):
                doc.Paragraphs(i + 1).Format.SpaceAfter = param['value']
        elif idMso == 'LigatureGalleryWord':
            sel.Font.Ligatures = param['type']
        elif idMso == 'StylisticSetsMenuWord':
            sel.Font.StylisticSet = param['type']
        elif idMso == 'NumberStyleGalleryWord':
            sel.Font.NumberStyle = param['type']
        elif idMso == 'TextFillGradientGallery':
            sel.Font.Fill.Visible = True
            sel.Font.Fill.ForeColor.RGB = param.get('fore_color', 0x0000FF)  # Default: Red
            sel.Font.Fill.BackColor.RGB = param.get('back_color', 0xFF0000)  # Default: Blue
            sel.Font.Fill.TwoColorGradient(param['style'], 1)  # Apply gradient after setting colors
        elif idMso == 'TextOutlineDashesGallery':
            sel.Font.Line.Visible = True
            sel.Font.Line.Weight = 1.0  # Set weight so outline is visible
            sel.Font.Line.DashStyle = param['style']
        elif idMso == 'TextOutlineWeightGallery':
            sel.Font.Line.Visible = True
            sel.Font.Line.Weight = param['weight']
        elif idMso == 'TextOutlineColorPickerAlternate':
            sel.Font.Line.Visible = True
            sel.Font.Line.Weight = 1.0  # Set weight so outline is visible
            sel.Font.Line.ForeColor.RGB = param['value']
        elif idMso == 'TextReflectionGallery':
            sel.Font.Reflection.Type = param['type']
        elif idMso == 'DropCapInsertGallery':
            wd.Selection.HomeKey(Unit=6)  # wdStory
            wd.ActiveDocument.Paragraphs(1).DropCap.Position = param['position']
        else:
            print(f"    [SKIP] No apply method for {idMso}")
            return False

        return True

    except Exception as e:
        print(f"    [ERROR] Failed to apply parameter: {e}")
        import traceback
        traceback.print_exc()
        return False


class ContentFormattingCoreParameterizedGenerator:
    """Generator for content_formatting × core_parameterized test cases."""

    def __init__(self, output_dir: Path = None, hidden: bool = False):
        self.execution = DirectExecute()
        self.output_dir = output_dir or (BENCHMARK_OUTPUT_DIR / 'content_formatting_x_core_parameterized')
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.inputs_dir = self.output_dir / 'inputs'
        self.inputs_dir.mkdir(parents=True, exist_ok=True)
        self.input_paths = {}  # Cache for generated input files
        self.hidden = hidden  # Run Word hidden to reduce distraction

    def generate_input_for_command(self, idMso: str, variation: str = None) -> Path:
        """
        Generate input document with appropriate precondition for a command.

        Args:
            idMso: Command identifier
            variation: Optional variation name (e.g., 'apply_Capitalize_Each_Word')

        Returns:
            Path to the generated input document
        """
        precondition_type = get_precondition_type(idMso, variation)

        # Some precondition types share the same input file but differ in selection
        # Map them to the base input file
        input_file_type = precondition_type
        if precondition_type == 'lowercase_text_select_all':
            input_file_type = 'lowercase_text'  # Same file, different selection

        input_filename = f'doc_in_{input_file_type}.docx'
        input_path = self.inputs_dir / input_filename

        # Return cached path if already generated
        if input_file_type in self.input_paths:
            return self.input_paths[input_file_type]

        if input_path.exists():
            print(f"  Reusing existing input: {input_filename}")
            self.input_paths[input_file_type] = input_path
            return input_path

        print(f"  Generating input: {input_filename}")

        wd = ensure_word(visible=not self.hidden)
        doc = None
        try:
            doc = wd.Documents.Add()

            if precondition_type == 'single_line':
                # Default: Single line with Target marker (mixed case for ChangeCaseGallery)
                # Use 24pt font for better visibility of formatting effects
                wd.Selection.TypeText(DEFAULT_INPUT['full_text'])
                wd.Selection.WholeStory()
                wd.Selection.Font.Size = 36

            elif precondition_type == 'paragraphs':
                # Multiple separate paragraphs for list formatting and line spacing
                wd.Selection.TypeText("First paragraph text")
                wd.Selection.TypeParagraph()
                wd.Selection.TypeText("Second paragraph text")
                wd.Selection.TypeParagraph()
                wd.Selection.TypeText("Third paragraph text")
                wd.Selection.WholeStory()
                wd.Selection.Font.Size = 36  # Consistent with other preconditions

            elif precondition_type == 'multilevel_list':
                # Hierarchical paragraphs with different indentation levels
                # Level 1 paragraphs
                wd.Selection.TypeText("First main point")
                wd.Selection.TypeParagraph()

                # Level 2 paragraphs (indented)
                wd.Selection.ParagraphFormat.LeftIndent = 36  # 0.5 inch
                wd.Selection.TypeText("First sub-point")
                wd.Selection.TypeParagraph()

                wd.Selection.TypeText("Second sub-point")
                wd.Selection.TypeParagraph()

                # Level 3 paragraphs (more indented)
                wd.Selection.ParagraphFormat.LeftIndent = 72  # 1 inch
                wd.Selection.TypeText("Detailed item one")
                wd.Selection.TypeParagraph()

                wd.Selection.TypeText("Detailed item two")
                wd.Selection.TypeParagraph()

                # Back to Level 2
                wd.Selection.ParagraphFormat.LeftIndent = 36
                wd.Selection.TypeText("Third sub-point")
                wd.Selection.TypeParagraph()

                # Back to Level 1
                wd.Selection.ParagraphFormat.LeftIndent = 0
                wd.Selection.TypeText("Second main point")
                wd.Selection.TypeParagraph()

                # Level 2 again
                wd.Selection.ParagraphFormat.LeftIndent = 36
                wd.Selection.TypeText("Fourth sub-point")

            elif precondition_type == 'list_with_levels':
                # Document with existing multilevel list structure
                # Apply multilevel list template first, then add content
                template = wd.ListGalleries(3).ListTemplates(1)  # First multilevel template

                # Level 1
                wd.Selection.TypeText("First main item")
                wd.Selection.Range.ListFormat.ApplyListTemplateWithLevel(
                    ListTemplate=template,
                    ContinuePreviousList=False,
                    DefaultListBehavior=1
                )
                wd.Selection.Range.ListFormat.ListLevelNumber = 1
                wd.Selection.TypeParagraph()

                # Level 2
                wd.Selection.TypeText("First sub-item")
                wd.Selection.Range.ListFormat.ListLevelNumber = 2
                wd.Selection.TypeParagraph()

                # Level 3 (this is the target paragraph we'll modify - starts at level 3)
                wd.Selection.TypeText("Second sub-item")
                wd.Selection.Range.ListFormat.ListLevelNumber = 3
                wd.Selection.TypeParagraph()

                # Level 3
                wd.Selection.TypeText("Detailed point")
                wd.Selection.Range.ListFormat.ListLevelNumber = 3
                wd.Selection.TypeParagraph()

                # Back to Level 1
                wd.Selection.TypeText("Second main item")
                wd.Selection.Range.ListFormat.ListLevelNumber = 1
                wd.Selection.TypeParagraph()

                # Level 2
                wd.Selection.TypeText("Third sub-item")
                wd.Selection.Range.ListFormat.ListLevelNumber = 2

            elif precondition_type == 'opentype_font_ligatures':
                # Text with ligature-eligible pairs (fi, fl, ff, ffi, ffl) for ligature testing
                # Use Calibri (Word's default font) which supports ligatures
                # IMPORTANT: Explicitly disable ligatures so applying them shows a visible change
                wd.Selection.TypeText("office efficient shuffle")
                wd.Selection.WholeStory()
                wd.Selection.Font.Name = "Calibri"  # Word's default font, supports ligatures
                wd.Selection.Font.Size = 36  # Large size to make ligatures clearly visible
                wd.Selection.Font.Ligatures = 0  # Disable ligatures in doc_in

            elif precondition_type == 'opentype_font_stylistic':
                # Text with characters that show dramatic stylistic set differences (Q, R, K, J)
                # Gabriola's stylistic sets 6-7 show flourishes on these characters
                wd.Selection.TypeText("Query Jockey Kraft")
                wd.Selection.WholeStory()
                wd.Selection.Font.Name = "Gabriola"
                wd.Selection.Font.Size = 36  # Large size to see flourishes clearly
                wd.Selection.Font.StylisticSet = 0  # No stylistic set in doc_in

            elif precondition_type == 'opentype_font':
                # Legacy: Single line with OpenType font (fallback)
                wd.Selection.TypeText("Alpha office shuffle Target Omega.")
                wd.Selection.WholeStory()
                wd.Selection.Font.Name = "Gabriola"
                wd.Selection.Font.Size = 36  # Consistent with other preconditions

            elif precondition_type in ('lowercase_text', 'lowercase_text_select_all'):
                # All lowercase text for Capitalize Each Word / Sentence case to show visible effect
                wd.Selection.TypeText("alpha target omega.")
                wd.Selection.WholeStory()
                wd.Selection.Font.Size = 36

            elif precondition_type == 'times_new_roman_text':
                # Text in Times New Roman for Font/apply_Calibri to show visible change
                # (Word default is Calibri, so applying Calibri to Calibri shows no change)
                # Use 24pt font for better visibility
                wd.Selection.TypeText(DEFAULT_INPUT['full_text'])
                wd.Selection.WholeStory()
                wd.Selection.Font.Name = "Times New Roman"
                wd.Selection.Font.Size = 36

            elif precondition_type == 'underlined_text':
                # Underline color tests should change color, not create underline.
                wd.Selection.TypeText(DEFAULT_INPUT['full_text'])
                wd.Selection.WholeStory()
                wd.Selection.Font.Size = 36
                rng = doc.Content
                rng.Find.ClearFormatting()
                rng.Find.Execute(FindText=DEFAULT_INPUT['select_text'])
                if rng.Find.Found:
                    rng.Font.Underline = 1

            elif precondition_type == 'uppercase_text':
                # All uppercase text for Sentence case to show visible effect
                # Sentence case will convert "ALPHA TARGET OMEGA." -> "Alpha target omega."
                # Use 24pt font for better visibility
                wd.Selection.TypeText("ALPHA TARGET OMEGA.")
                wd.Selection.WholeStory()
                wd.Selection.Font.Size = 36

            elif precondition_type == 'heading_line':
                # Single line for heading/title styles (realistic use case)
                wd.Selection.TypeText("Project Overview")
                wd.Selection.WholeStory()
                wd.Selection.Font.Size = 36  # Consistent with other preconditions

            elif precondition_type == 'paragraph_spacing_before':
                # Multiple paragraphs with 36pt space before (to show change when set to 0 or 24)
                paragraphs = [
                    'The quick brown fox jumps over the lazy dog.',
                    'Pack my box with five dozen liquor jugs.',
                    'How vexingly quick daft zebras jump.',
                ]
                for i, text in enumerate(paragraphs):
                    wd.Selection.TypeText(text)
                    if i < len(paragraphs) - 1:
                        wd.Selection.TypeParagraph()
                for i in range(doc.Paragraphs.Count):
                    doc.Paragraphs(i + 1).Format.SpaceBefore = 36
                    doc.Paragraphs(i + 1).Format.SpaceAfter = 0

            elif precondition_type == 'paragraph_spacing_after':
                # Multiple paragraphs with 36pt space after (to show change when set to 0 or 24)
                paragraphs = [
                    'The quick brown fox jumps over the lazy dog.',
                    'Pack my box with five dozen liquor jugs.',
                    'How vexingly quick daft zebras jump.',
                ]
                for i, text in enumerate(paragraphs):
                    wd.Selection.TypeText(text)
                    if i < len(paragraphs) - 1:
                        wd.Selection.TypeParagraph()
                for i in range(doc.Paragraphs.Count):
                    doc.Paragraphs(i + 1).Format.SpaceBefore = 0
                    doc.Paragraphs(i + 1).Format.SpaceAfter = 36

            elif precondition_type == 'with_paragraph':
                # DropCap reformats the first letter of the first paragraph; the
                # paragraph needs to be long enough for the styled letter to be
                # visually distinct from body text.
                dropcap_text = (
                    'Alpha TARGET Omega. '
                    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. '
                    'Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. '
                    'Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. '
                    'Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. '
                    'Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.'
                )
                wd.Selection.TypeText(dropcap_text)
                wd.Selection.WholeStory()
                wd.Selection.Font.Size = 11

            # Note: lowercase_text_select_all uses lowercase_text input (handled above)
            # Note: multi_word_text removed - use single_line with custom selection instead

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
        """Load commands classified as content_formatting × core_parameterized."""
        commands_file = COMMAND_SETS_DIR / 'commands_by_policy_id.json'

        if not commands_file.exists():
            print(f"Warning: {commands_file} not found.")
            return []

        with open(commands_file, 'r', encoding='utf-8') as f:
            all_commands = json.load(f)

        return [
            cmd for cmd in all_commands
            if cmd.get('behavior') == 'content_formatting'
            and cmd.get('scope') == 'core_parameterized'
        ]

    def generate_test_case(self, wd, idMso: str, param: Dict, input_path: Path) -> Dict:
        """Generate a single test case."""
        variation = f"apply_{param['label']}"
        result = {
            'idMso': idMso,
            'variation': variation,
            'parameter': param['label'],
            'success': False,
            'message': '',
            'paths': {}
        }

        try:
            # Open appropriate input for this command
            doc = wd.Documents.Open(str(input_path.absolute()))

            # Determine selection strategy based on precondition type (with variation-specific handling)
            precondition_type = get_precondition_type(idMso, variation)

            # Select text
            if idMso in COMMANDS_NEED_FULL_SELECTION:
                wd.Selection.WholeStory()
                print(f"    [SELECTION] Selected entire document")
            elif precondition_type == 'list_with_levels':
                # For ListLevelGallery, select specific paragraph by number
                target_para = param.get('target_paragraph', 3)
                try:
                    # Move to the target paragraph
                    wd.Selection.HomeKey(Unit=6)  # Go to start of document
                    for i in range(target_para - 1):
                        wd.Selection.MoveDown(Unit=4, Count=1)  # Move down by paragraph
                    wd.Selection.MoveDown(Unit=4, Count=1, Extend=1)  # Select the paragraph
                    print(f"    [SELECTION] Selected paragraph {target_para}")
                except Exception as e:
                    result['message'] = f'Failed to select paragraph {target_para}: {e}'
                    doc.Close(False)
                    return result
            elif precondition_type == 'lowercase_text':
                # For lowercase_text precondition, select "target" (lowercase)
                if not self.execution.select(wd, 'target'):
                    result['message'] = f'Failed to select text'
                    doc.Close(False)
                    return result
                print(f"    [SELECTION] Selected 'target' (lowercase)")
            elif precondition_type == 'uppercase_text':
                # For uppercase_text precondition, select "TARGET" (uppercase)
                if not self.execution.select(wd, 'TARGET'):
                    result['message'] = f'Failed to select text'
                    doc.Close(False)
                    return result
                print(f"    [SELECTION] Selected 'TARGET' (uppercase)")
            elif precondition_type == 'lowercase_text_select_all':
                # For sentence case, select all text (not just one word)
                wd.Selection.WholeStory()
                print(f"    [SELECTION] Selected entire document (for sentence case)")
            else:
                # Check for custom selection text (e.g., "Target Omega" for underline variants)
                variation_key = f"{idMso}/{variation}"
                custom_select = VARIATION_SELECT_TEXT_MAP.get(variation_key)
                select_text = custom_select if custom_select else DEFAULT_INPUT['select_text']

                if not self.execution.select(wd, select_text):
                    result['message'] = f'Failed to select text'
                    doc.Close(False)
                    return result
                print(f"    [SELECTION] Selected '{select_text}'")

            # Apply parameter
            if not apply_parameter(wd, idMso, param):
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

            # Determine select_text for spec based on precondition or custom selection
            variation_key = f"{idMso}/{variation}"
            custom_select = VARIATION_SELECT_TEXT_MAP.get(variation_key)

            if idMso in COMMANDS_NEED_FULL_SELECTION:
                select_text = 'all'
            elif precondition_type == 'list_with_levels':
                target_para = param.get('target_paragraph', 3)
                select_text = f'paragraph_{target_para}'
            elif precondition_type == 'lowercase_text':
                select_text = 'target'  # lowercase
            elif precondition_type == 'uppercase_text':
                select_text = 'TARGET'  # uppercase
            elif precondition_type == 'lowercase_text_select_all':
                select_text = 'all'  # select entire document
            elif custom_select:
                select_text = custom_select  # Custom selection (e.g., "Target Omega")
            else:
                select_text = DEFAULT_INPUT['select_text']

            # Determine full_text for spec based on precondition
            if precondition_type == 'single_line':
                full_text = DEFAULT_INPUT['full_text']
            elif precondition_type == 'paragraphs':
                full_text = "Three paragraphs"
            elif precondition_type == 'lowercase_text':
                full_text = "alpha target omega."
            elif precondition_type == 'uppercase_text':
                full_text = "ALPHA TARGET OMEGA."
            elif precondition_type == 'lowercase_text_select_all':
                full_text = "alpha target omega."
            elif precondition_type == 'heading_line':
                full_text = "Project Overview"
            elif precondition_type == 'underlined_text':
                full_text = DEFAULT_INPUT['full_text']
            elif precondition_type in ('paragraph_spacing_before', 'paragraph_spacing_after'):
                full_text = None  # Multi-paragraph structure
            else:
                # For other types (multilevel_list, list_with_levels, opentype_font), describe the structure rather than exact text
                full_text = None

            # Determine verification method based on command type
            is_list_command = idMso in ['BulletsGalleryWord', 'NumberingGalleryWord', 'MultilevelListGallery', 'ListLevelGallery']
            is_paragraph_spacing = idMso in ('ParagraphSpacingBefore', 'ParagraphSpacingAfter')

            if is_list_command:
                verification_spec = {
                    'primary': {
                        'method': 'numbering_property_comparison',
                        'description': 'Word COM list numbering property comparison',
                    },
                }
            elif is_paragraph_spacing:
                prop_name = param.get('property_name', 'SpaceBefore')
                verification_spec = {
                    'primary': {
                        'method': 'paragraph_spacing_verification',
                        'description': f'Check ParagraphFormat.{prop_name} == {param["value"]} for all paragraphs',
                        'params': {
                            'property_name': prop_name,
                            'value': param['value'],
                        },
                    },
                }
            else:
                verification_spec = {
                    'primary': {
                        'method': 'property_based',
                        'description': 'Word COM property comparison with gold',
                    },
                }

            # Generate spec
            spec = {
                'idMso': idMso,
                'behavior': 'content_formatting',
                'scope': 'core_parameterized',
                'variation': f"apply_{param['label']}",
                'generated_at': datetime.now().isoformat(),
                'parameter': param,
                'precondition': {
                    'type': precondition_type,
                    'select_text': select_text,
                },
                'verification': verification_spec,
                'files': {
                    'doc_in': doc_in_ref,
                    'gold': 'gold.docx',
                }
            }

            # Add full_text for default precondition type
            if full_text:
                spec['precondition']['full_text'] = full_text

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
            'test_set': 'content_formatting_x_core_parameterized',
            'generated_at': datetime.now().isoformat(),
            'total_commands': len(commands) + len(skipped_commands),
            'skipped': len(skipped_commands),
            'skipped_commands': [cmd['idMso'] for cmd in skipped_commands],
            'successful': 0,
            'failed': 0,
            'results': []
        }

        wd = ensure_word(visible=not self.hidden)

        for cmd in commands:
            idMso = cmd['idMso']
            print(f"\n{idMso}...")

            # Extract parameters (may return multiple variants for list galleries)
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
                # Generate input document for this command/variation (may be variation-specific)
                variation = f"apply_{param['label']}"
                input_path = self.generate_input_for_command(idMso, variation)

                result = self.generate_test_case(wd, idMso, param, input_path)

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
        description='Generate test cases for content_formatting × core_parameterized'
    )
    parser.add_argument('--limit', '-l', type=int, help='Limit number of commands')
    parser.add_argument('--output', '-o', help='Output directory')
    parser.add_argument('--command', '-c', help='Generate only for specific command (idMso), e.g., BulletsGalleryWord')
    parser.add_argument('--hidden', action='store_true', help='Run Word hidden (no visible window)')

    args = parser.parse_args()

    output_dir = Path(args.output) if args.output else None
    generator = ContentFormattingCoreParameterizedGenerator(output_dir=output_dir, hidden=args.hidden)
    generator.generate_all(limit=args.limit, command_filter=args.command)


if __name__ == '__main__':
    main()
