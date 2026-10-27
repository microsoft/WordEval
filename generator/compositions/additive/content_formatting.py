"""
Additive composition generator for content_formatting behavior.

Track 1A: Single-target font compositions (4 variants)
  - Composes multiple character/font operations on "Target" text in doc_in_single_line.docx
  - Verifier: _verify_property_based (full property comparison, zero changes)

Track 1B: Paragraph compositions (6 variants)
  - Composes paragraph operations on multi-paragraph doc with preset 36pt spacing
  - Variants 1-3: alignment, line spacing, indent, border
  - Variants 4-6: alignment + paragraph spacing + drop cap
  - Verifier: _verify_property_based (full property comparison, zero changes)

Track 2A: Multi-paragraph list + formatting compositions (3 variants)
  - Composes list + paragraph operations on doc_in_paragraphs.docx
  - Verifier: _verify_all_paragraphs with composition dispatch

Usage:
    python -m generator.compositions.additive.content_formatting
    python -m generator.compositions.additive.content_formatting --command AdditiveFontComposition
"""

import json
import shutil
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from generator.config import BENCHMARK_OUTPUT_DIR
from generator.compositions.atomic_intents import build_content_formatting_atomic_intents
from generator.compositions.interaction.shared import write_instructions_json
from generator.execution.setup import ensure_word, finalize_word
from generator.test_sets.spec_utils import write_spec_if_changed

TEST_SET = 'additive_composition_x_content_formatting'

# ============================================================================
# Slot value → COM property mappings
# ============================================================================

def _apply_slot_value(wd, rng, slot_id, value):
    """Apply a single slot value to a Range (or via Selection for ExecuteMso).

    Args:
        wd: Word application instance
        rng: Range object positioned on target text
        slot_id: Slot identifier (C1, C2, ..., P1, P2, ...)
        value: Slot value string
    """
    sel = wd.Selection

    if slot_id == 'bold':  # bold
        rng.Font.Bold = (value == 'on')

    elif slot_id == 'italic':  # italic
        rng.Font.Italic = (value == 'on')

    elif slot_id == 'underline_style':  # underline
        underline_map = {'Single': 1, 'Double': 3, 'Dash': 7, 'Dotted': 4}
        rng.Font.Underline = underline_map[value]

    elif slot_id == 'underline_color':  # underline color
        color_map = {'Red': 255, 'Yellow': 65535}
        rng.Font.UnderlineColor = color_map[value]

    elif slot_id == 'strikethrough':  # strikethrough
        rng.Font.StrikeThrough = (value == 'on')

    elif slot_id == 'sub_super':  # subscript / superscript
        if value == 'subscript':
            rng.Font.Superscript = False
            rng.Font.Subscript = True
        elif value == 'superscript':
            rng.Font.Subscript = False
            rng.Font.Superscript = True
        elif value == 'off':
            rng.Font.Subscript = False
            rng.Font.Superscript = False

    elif slot_id == 'char_border':  # character border (toggle via ExecuteMso)
        rng.Select()
        time.sleep(0.1)
        wd.CommandBars.ExecuteMso('CharacterBorder')
        time.sleep(0.1)

    elif slot_id == 'char_shading':  # character shading (toggle via ExecuteMso)
        rng.Select()
        time.sleep(0.1)
        wd.CommandBars.ExecuteMso('CharacterShading')
        time.sleep(0.1)

    elif slot_id == 'font_name':  # font name
        font_map = {'Arial': 'Arial', 'Comic_Sans': 'Comic Sans MS'}
        rng.Font.Name = font_map[value]

    elif slot_id == 'font_size':  # font size
        if value == '18pt':
            rng.Font.Size = 18
        elif value == '32pt':
            rng.Font.Size = 32
        elif value == '40pt':
            rng.Font.Size = 40
        elif value == '8pt':
            rng.Font.Size = 8
        elif value == 'increase':
            rng.Select()
            time.sleep(0.1)
            wd.CommandBars.ExecuteMso('FontSizeIncreaseWord')
            time.sleep(0.1)
        elif value == 'decrease':
            rng.Select()
            time.sleep(0.1)
            wd.CommandBars.ExecuteMso('FontSizeDecreaseWord')
            time.sleep(0.1)

    elif slot_id == 'font_color':  # font color
        color_map = {'Red': 255, 'Yellow': 65535}
        rng.Font.Color = color_map[value]

    elif slot_id == 'highlight':  # highlight color
        index_map = {'Blue': 2, 'Yellow': 7}
        rng.HighlightColorIndex = index_map[value]

    elif slot_id == 'text_case':  # case
        case_map = {'UPPERCASE': 1, 'lowercase': 0, 'tOGGLE': 5}
        rng.Case = case_map[value]

    elif slot_id == 'outline_color':  # text outline color
        color_map = {'Blue': 16711680, 'Yellow': 65535}
        rng.Font.Line.Visible = True
        rng.Font.Line.Weight = 1.0  # ensure visible
        rng.Font.Line.ForeColor.RGB = color_map[value]

    elif slot_id == 'outline_weight':  # text outline weight
        weight_map = {'1pt': 1.0, '2pt': 2.0, '3pt': 3.0}
        rng.Font.Line.Visible = True
        rng.Font.Line.Weight = weight_map[value]

    elif slot_id == 'outline_dash':  # text outline dashes
        style_map = {'Dash': 4, 'Round_Dot': 3}
        rng.Font.Line.Visible = True
        rng.Font.Line.Weight = 1.0  # ensure visible
        rng.Font.Line.DashStyle = style_map[value]

    elif slot_id == 'reflection':  # text reflection
        type_map = {'type_1': 1, 'type_3': 3, 'type_5': 5, 'type_7': 7}
        rng.Font.Reflection.Type = type_map[value]

    elif slot_id == 'fill_gradient':  # text fill gradient
        style_map = {
            'horizontal': 2, 'vertical': 1,
            'diag_L': 4, 'diag_R': 3,
        }
        rng.Font.Fill.Visible = True
        rng.Font.Fill.ForeColor.RGB = 255       # Red
        rng.Font.Fill.BackColor.RGB = 16711680   # Blue
        rng.Font.Fill.TwoColorGradient(style_map[value], 1)

    elif slot_id == 'alignment':  # alignment
        align_map = {'center': 1, 'right': 2, 'justify': 3, 'distributed': 5}
        rng.ParagraphFormat.Alignment = align_map[value]

    elif slot_id == 'indent':  # left indent (increase/decrease via ExecuteMso)
        rng.Select()
        time.sleep(0.1)
        if value == 'increase':
            wd.CommandBars.ExecuteMso('IndentIncreaseWord')
        elif value == 'decrease':
            wd.CommandBars.ExecuteMso('IndentDecreaseWord')
        time.sleep(0.1)

    elif slot_id == 'line_spacing':  # line spacing
        spacing_map = {'1.15': 1.15, '1.5': 1.5, 'Double': 2.0, 'Triple': 3.0}
        rng.ParagraphFormat.LineSpacingRule = 0  # wdLineSpaceMultiple
        rng.ParagraphFormat.LineSpacing = spacing_map[value] * 12

    elif slot_id == 'para_border':  # paragraph border
        index_map = {'bottom': -3, 'top': -1}
        rng.Paragraphs(1).Borders(index_map[value]).LineStyle = 1

    elif slot_id == 'spacing_before':  # paragraph spacing before (all paragraphs)
        doc = wd.ActiveDocument
        spacing_map = {'0pt': 0, '24pt': 24}
        for i in range(1, doc.Paragraphs.Count + 1):
            doc.Paragraphs(i).Format.SpaceBefore = spacing_map[value]

    elif slot_id == 'spacing_after':  # paragraph spacing after (all paragraphs)
        doc = wd.ActiveDocument
        spacing_map = {'0pt': 0, '24pt': 24}
        for i in range(1, doc.Paragraphs.Count + 1):
            doc.Paragraphs(i).Format.SpaceAfter = spacing_map[value]

    elif slot_id == 'drop_cap':  # drop cap (first paragraph)
        doc = wd.ActiveDocument
        position_map = {'dropped': 1, 'in_margin': 2}
        doc.Paragraphs(1).DropCap.Position = position_map[value]


def _apply_precondition(wd, rng, slot_id, value):
    """Apply precondition to doc_in range for 'off'/'decrease' values.

    Sets the property to the state that the instruction will ask to remove.
    """
    if slot_id == 'bold' and value == 'off':
        rng.Font.Bold = True
    elif slot_id == 'italic' and value == 'off':
        rng.Font.Italic = True
    elif slot_id == 'strikethrough' and value == 'off':
        rng.Font.StrikeThrough = True
    elif slot_id == 'sub_super' and value == 'off':
        # Pre-apply superscript so instruction can turn it off
        rng.Font.Superscript = True
    elif slot_id == 'char_border' and value == 'off':
        # Character border on → toggle via ExecuteMso
        rng.Select()
        time.sleep(0.1)
        wd.CommandBars.ExecuteMso('CharacterBorder')
        time.sleep(0.1)
    elif slot_id == 'char_shading' and value == 'off':
        # Character shading on → toggle via ExecuteMso
        rng.Select()
        time.sleep(0.1)
        wd.CommandBars.ExecuteMso('CharacterShading')
        time.sleep(0.1)
    elif slot_id == 'indent' and value == 'decrease':
        # Indent increase so the instruction can decrease it
        rng.Select()
        time.sleep(0.1)
        wd.CommandBars.ExecuteMso('IndentIncreaseWord')
        time.sleep(0.1)


# ============================================================================
# Track 1: Single-target font / paragraph composition variants
# ============================================================================

# Values that require preconditions in doc_in
NEEDS_PRECONDITION = {
    ('bold', 'off'), ('italic', 'off'), ('strikethrough', 'off'), ('sub_super', 'off'),
    ('char_border', 'off'), ('char_shading', 'off'), ('indent', 'decrease'),
}

PARAGRAPH_TEXTS = [
    "Alpha Target Omega. This paragraph provides the primary text for "
    "composition testing, with enough length to wrap across multiple lines "
    "so that alignment, indentation, and border changes remain clearly "
    "visible in the document layout.",

    "The second paragraph adds depth to the document structure and shows "
    "that paragraph spacing changes affect all paragraphs consistently. "
    "Line spacing and indentation properties can be compared across "
    "paragraphs to verify correct application scope.",

    "A third paragraph rounds out the document with additional body text. "
    "Paragraph properties like spacing before and after are clearly "
    "distinguishable between each block of text, and drop cap styling "
    "on the first paragraph is visible against this multi-paragraph context.",
]


FONT_VARIANTS = {
    'variant_full_bundle': {
        'description': 'Large font bundle (14 operations)',
        'instruction': "Make the word 'Target' bold and italic with a double red underline (RGB: 255, 0, 0), add strikethrough and subscript, add a character border, change the font to Arial at 32 points, set the font color to red (RGB: 255, 0, 0) with yellow highlighting, add a blue text outline (RGB: 0, 0, 255) at 2pt weight, and add type 1 text reflection (tight reflection, touching).",
        'slots': {
            'bold': 'on', 'italic': 'on', 'underline_style': 'Double', 'underline_color': 'Red',
            'strikethrough': 'on', 'sub_super': 'subscript', 'char_border': 'on',
            'font_name': 'Arial', 'font_size': '32pt', 'font_color': 'Red', 'highlight': 'Yellow',
            'outline_color': 'Blue', 'outline_weight': '2pt', 'reflection': 'type_1',
        },
    },
    'variant_mixed_effects': {
        'description': 'Mixed font bundle (13 operations)',
        'instruction': "Make the word 'Target' bold and italic with a single yellow underline, add strikethrough and superscript, apply character shading, change the font to Comic Sans MS at 40 points, apply blue highlighting, set the text outline to dashed style, set the text outline weight to 1pt, and add a horizontal text fill gradient.",
        'slots': {
            'bold': 'on', 'italic': 'on', 'underline_style': 'Single', 'underline_color': 'Yellow',
            'strikethrough': 'on', 'sub_super': 'superscript', 'char_shading': 'on', 'font_name': 'Comic_Sans',
            'font_size': '40pt', 'highlight': 'Blue',
            'outline_dash': 'Dash', 'outline_weight': '1pt', 'fill_gradient': 'horizontal',
        },
    },
    'variant_toggle_mixed': {
        'description': 'Mixed on/off font bundle (11 operations)',
        'instruction': "Remove bold and italic from the word 'Target', add a dashed underline, add a character border, remove character shading, change the font to Arial, increase the font size from 36pt to 48pt, set the font color to red (RGB: 255, 0, 0), convert to UPPERCASE, set the text outline weight to 3pt, and add type 5 text reflection (half reflection, 4pt offset).",
        'slots': {
            'bold': 'off', 'italic': 'off', 'underline_style': 'Dash', 'char_border': 'on',
            'char_shading': 'off', 'font_name': 'Arial', 'font_size': 'increase',
            'font_color': 'Red', 'text_case': 'UPPERCASE', 'outline_weight': '3pt',
            'reflection': 'type_5',
        },
    },
    'variant_off_heavy': {
        'description': 'Off-heavy font bundle (9 operations)',
        'instruction': "Make the word 'Target' italic, remove strikethrough, remove superscript, remove the character border, decrease the font size from 36pt to 28pt, apply yellow highlighting, add a blue text outline (RGB: 0, 0, 255) with round dot style, and set the text outline weight to 1pt.",
        'slots': {
            'italic': 'on', 'strikethrough': 'off', 'sub_super': 'off', 'char_border': 'off',
            'font_size': 'decrease', 'highlight': 'Yellow',
            'outline_color': 'Blue', 'outline_dash': 'Round_Dot', 'outline_weight': '1pt',
        },
    },
}


PARAGRAPH_VARIANTS = {
    'variant_center_spacing_border': {
        'description': 'Alignment + spacing + border (3 operations)',
        'instruction': "Center the paragraph containing the word 'Target', set its line spacing to 1.5 lines, and add a top border to the paragraph containing 'Target'.",
        'slots': {
            'alignment': 'center', 'line_spacing': '1.5', 'para_border': 'top',
        },
    },
    'variant_right_indent_double': {
        'description': 'Alignment + indent + spacing + border (4 operations)',
        'instruction': "Right-align the paragraph containing the word 'Target', increase its left indentation to 1.27 cm, set double line spacing, and add a bottom border to the paragraph containing 'Target'.",
        'verification_comparisons': {
            'indent': {'left_indent': 'exact'},
        },
        'slots': {
            'alignment': 'right', 'indent': 'increase', 'line_spacing': 'Double', 'para_border': 'bottom',
        },
    },
    'variant_justify_unindent_triple': {
        'description': 'Alignment + indent decrease + spacing + border (4 operations)',
        'instruction': "Justify the paragraph containing the word 'Target', decrease its left indentation from 1.27 cm to 0 cm (no indent), set triple line spacing, and add a top border to the paragraph containing 'Target'.",
        'verification_comparisons': {
            'indent': {'left_indent': 'exact'},
        },
        'slots': {
            'alignment': 'justify', 'indent': 'decrease', 'line_spacing': 'Triple', 'para_border': 'top',
        },
    },
    'variant_center_dropcap': {
        'description': 'Alignment + spacing before/after + drop cap (4 operations)',
        'instruction': "Center the paragraph containing the word 'Target', set paragraph spacing to 0 points before and 0 points after for all paragraphs, and apply a dropped drop cap to the first paragraph.",
        'slots': {
            'alignment': 'center', 'spacing_before': '0pt', 'spacing_after': '0pt', 'drop_cap': 'dropped',
        },
    },
    'variant_right_double_dropcap': {
        'description': 'Alignment + line spacing + spacing before + drop cap (4 operations)',
        'instruction': "Right-align the paragraph containing the word 'Target', set its line spacing to double, set paragraph spacing before to 24 points for all paragraphs, and apply an in-margin drop cap to the first paragraph.",
        'slots': {
            'alignment': 'right', 'line_spacing': 'Double', 'spacing_before': '24pt', 'drop_cap': 'in_margin',
        },
    },
    'variant_justify_indent_dropcap': {
        'description': 'Alignment + indent + spacing + drop cap (5 operations)',
        'instruction': "Justify the paragraph containing the word 'Target', increase its left indentation, set paragraph spacing to 0 points before and 24 points after for all paragraphs, and apply a dropped drop cap to the first paragraph.",
        'slots': {
            'alignment': 'justify', 'indent': 'increase', 'spacing_before': '0pt', 'spacing_after': '24pt', 'drop_cap': 'dropped',
        },
    },
}


# ============================================================================
# Track 2A: Multi-paragraph list + formatting compositions
# ============================================================================

LIST_VARIANTS = {
    'variant_bullets_spacing': {
        'description': 'Bullets + line spacing (2 operations)',
        'instruction': "Add solid circle bullet points to all paragraphs and set the line spacing to 1.5 lines.",
        'slots': {
            'list_type': ('BulletsGalleryWord', 1),  # solid circle = template 1
            'list_spacing': '1.5',
        },
    },
    'variant_numbered_center': {
        'description': 'Numbering + line spacing + alignment (3 operations)',
        'instruction': "Add a numbered list (1. 2. 3.) to all paragraphs, set double line spacing, and center all text.",
        'slots': {
            'list_type': ('NumberingGalleryWord', 1),  # arabic period = template 1
            'list_spacing': 'Double',
            'list_alignment': 'center',
        },
    },
    'variant_checkmarks_indent': {
        'description': 'Bullets + indent + line spacing (3 operations)',
        'instruction': "Add checkmark bullet points to all paragraphs, increase the indentation level of all list items, and set triple line spacing.",
        'slots': {
            'list_type': ('BulletsGalleryWord', 7),  # checkmark = template 7
            'list_indent': 'increase',
            'list_spacing': 'Triple',
        },
    },
}


def _apply_list_slot_value(wd, slot_id, value):
    """Apply a list composition slot value. Operates on current Selection."""
    sel = wd.Selection

    if slot_id == 'list_type':
        idMso, template_index = value
        if idMso == 'BulletsGalleryWord':
            gallery_type = 1
        else:  # NumberingGalleryWord
            gallery_type = 2
        template = wd.ListGalleries(gallery_type).ListTemplates(template_index)
        sel.Range.ListFormat.ApplyListTemplateWithLevel(
            ListTemplate=template,
            ContinuePreviousList=False,
            ApplyTo=0,
            DefaultListBehavior=1
        )

    elif slot_id == 'list_spacing':  # line spacing
        spacing_map = {'1.5': 1.5, 'Double': 2.0, 'Triple': 3.0}
        sel.ParagraphFormat.LineSpacingRule = 0  # wdLineSpaceMultiple
        sel.ParagraphFormat.LineSpacing = spacing_map[value] * 12

    elif slot_id == 'list_alignment':  # alignment
        align_map = {'center': 1, 'right': 2}
        sel.ParagraphFormat.Alignment = align_map[value]

    elif slot_id == 'list_indent':  # indent
        if value == 'increase':
            wd.CommandBars.ExecuteMso('IndentIncreaseWord')
            time.sleep(0.1)




# ============================================================================
# Generator class
# ============================================================================

class AdditiveContentFormattingGenerator:
    """Generator for content_formatting x additive_composition test cases."""

    def __init__(self, output_dir: Path = None, hidden: bool = False):
        self.output_dir = output_dir or (BENCHMARK_OUTPUT_DIR / TEST_SET)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.inputs_dir = self.output_dir / 'inputs'
        self.inputs_dir.mkdir(parents=True, exist_ok=True)
        self.hidden = hidden

    def _find_range(self, doc, target_text):
        """Find target text in document and return a Range object."""
        rng = doc.Content
        rng.Find.ClearFormatting()
        rng.Find.Execute(FindText=target_text, MatchCase=False, Forward=True)
        if not rng.Find.Found:
            raise RuntimeError(f"Target text '{target_text}' not found in document")
        return rng

    def _select_all(self, wd, doc):
        """Select all document content."""
        wd.Selection.HomeKey(Unit=6)  # wdStory
        wd.Selection.WholeStory()

    def _get_shared_input(self, input_name):
        """Get path to shared input file, copying from existing benchmark if needed."""
        dest = self.inputs_dir / f'{input_name}.docx'
        if dest.exists():
            return dest

        # Find source from existing benchmark test sets
        search_dirs = [
            BENCHMARK_OUTPUT_DIR / 'content_formatting_x_core_parameterized' / 'inputs',
            BENCHMARK_OUTPUT_DIR / 'content_formatting_x_core_actionable' / 'inputs',
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

    # ------------------------------------------------------------------
    # Track 1A/1B: Font and paragraph composition generation
    # ------------------------------------------------------------------

    def _get_paragraph_input(self, wd):
        """Create a shared multi-paragraph input with preset spacing.

        Three longer paragraphs at 18pt font with 36pt before/after spacing.
        First paragraph contains 'Target' for range-based operations.
        """
        dest = self.inputs_dir / 'doc_in_paragraphs_with_spacing.docx'
        if dest.exists():
            return dest

        doc = wd.Documents.Add()
        try:
            for i, text in enumerate(PARAGRAPH_TEXTS):
                if i > 0:
                    wd.Selection.TypeParagraph()
                wd.Selection.TypeText(text)
            for i in range(1, doc.Paragraphs.Count + 1):
                doc.Paragraphs(i).Range.Font.Size = 18
                doc.Paragraphs(i).Format.SpaceBefore = 36
                doc.Paragraphs(i).Format.SpaceAfter = 36
            doc.SaveAs(str(dest.absolute()), FileFormat=16)
        finally:
            doc.Close(False)
        time.sleep(0.3)
        print("  Created input: doc_in_paragraphs_with_spacing.docx")
        return dest

    def _generate_target_text_variant(
        self,
        wd,
        case_name,
        variant_name,
        config,
        *,
        shared_input,
        shared_input_name,
        template_name,
        full_text,
    ):
        """Generate a single single-target font or paragraph composition variant."""
        variant_dir = self.output_dir / case_name / variant_name
        variant_dir.mkdir(parents=True, exist_ok=True)

        doc_in_path = variant_dir / 'doc_in.docx'
        gold_path = variant_dir / 'gold.docx'
        spec_path = variant_dir / 'spec.json'

        slots = config['slots']

        # Determine if we need a custom doc_in (has preconditions)
        has_preconditions = any(
            (sid, val) in NEEDS_PRECONDITION for sid, val in slots.items()
        )

        # Step 1: Build doc_in
        if has_preconditions:
            doc = wd.Documents.Open(str(shared_input.absolute()))
            try:
                rng = self._find_range(doc, 'Target')
                for slot_id, value in slots.items():
                    if (slot_id, value) in NEEDS_PRECONDITION:
                        _apply_precondition(wd, rng, slot_id, value)
                doc.SaveAs(str(doc_in_path.absolute()), FileFormat=16)
            finally:
                doc.Close(False)
            time.sleep(0.3)
            doc_in_ref = 'doc_in.docx'
        else:
            doc_in_ref = f'../../inputs/{shared_input_name}.docx'

        # Step 2: Build gold from doc_in
        source = doc_in_path if has_preconditions else shared_input
        doc = wd.Documents.Open(str(source.absolute()))
        try:
            rng = self._find_range(doc, 'Target')
            # Apply all slot values — direct COM first, then ExecuteMso-based
            # Split to ensure direct COM properties are set before ExecuteMso
            # operations that might re-select
            executemso_slots = {'char_border', 'char_shading'}
            relative_slots = set()
            for sid, val in slots.items():
                if sid == 'font_size' and val in ('increase', 'decrease'):
                    relative_slots.add(sid)
                elif sid == 'indent':
                    relative_slots.add(sid)

            # Phase 1: Direct COM property settings
            for slot_id, value in slots.items():
                if slot_id not in executemso_slots and slot_id not in relative_slots:
                    _apply_slot_value(wd, rng, slot_id, value)

            # Phase 2: ExecuteMso toggles (need re-select after direct COM)
            for slot_id in executemso_slots:
                if slot_id in slots:
                    rng = self._find_range(doc, 'Target')
                    _apply_slot_value(wd, rng, slot_id, slots[slot_id])

            # Phase 3: Relative operations (font size increase/decrease, indent)
            for slot_id in relative_slots:
                if slot_id in slots:
                    rng = self._find_range(doc, 'Target')
                    _apply_slot_value(wd, rng, slot_id, slots[slot_id])

            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
        finally:
            doc.Close(False)
        time.sleep(0.3)

        # Step 3: Build spec
        atomic_intents = build_content_formatting_atomic_intents(
            slots,
            target_comparisons=config.get('verification_comparisons'),
        )

        spec = {
            'behavior': 'content_formatting',
            'scope': 'additive_composition',
            'variation': variant_name,
            'generated_at': datetime.now().isoformat(),
            'composition': {
                'type': 'additive',
                'atomic_intents': atomic_intents,
                'operation_count': len(slots),
            },
            'precondition': {
                'template': template_name,
                'full_text': full_text,
                'select_text': 'Target',
            },
            'verification': {
                'primary': {
                    'method': 'property_based',
                    'description': 'Word COM property comparison with gold (all properties)',
                },
                'invariants': [
                    {
                        'method': 'text_preserved',
                        'description': 'Plain text content unchanged',
                    },
                ],
            },
            'files': {
                'doc_in': doc_in_ref,
                'gold': 'gold.docx',
            },
        }
        write_spec_if_changed(spec_path, spec)

        print(f"  [{case_name}/{variant_name}] Generated ({len(slots)} operations)")
        return True

    # ------------------------------------------------------------------
    # Track 2A: List composition generation
    # ------------------------------------------------------------------

    def _generate_list_variant(self, wd, case_name, variant_name, config):
        """Generate a single Track 2A list composition variant."""
        variant_dir = self.output_dir / case_name / variant_name
        variant_dir.mkdir(parents=True, exist_ok=True)

        gold_path = variant_dir / 'gold.docx'
        spec_path = variant_dir / 'spec.json'

        slots = config['slots']
        shared_input = self._get_shared_input('doc_in_paragraphs')

        # Build gold from shared input
        doc = wd.Documents.Open(str(shared_input.absolute()))
        try:
            self._select_all(wd, doc)
            for slot_id, value in slots.items():
                _apply_list_slot_value(wd, slot_id, value)
            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
        finally:
            doc.Close(False)
        time.sleep(0.3)

        atomic_intents = build_content_formatting_atomic_intents(
            slots,
            target_comparisons=config.get('verification_comparisons'),
        )

        spec = {
            'behavior': 'content_formatting',
            'scope': 'additive_composition',
            'variation': variant_name,
            'generated_at': datetime.now().isoformat(),
            'composition': {
                'type': 'additive',
                'atomic_intents': atomic_intents,
                'operation_count': len(slots),
            },
            'precondition': {
                'template': 'paragraphs',
                'select_text': 'all',
            },
            'verification': {
                'primary': {
                    'method': 'property_based',
                    'description': 'Word COM paragraph property comparison with gold',
                },
                'invariants': [
                    {
                        'method': 'text_preserved',
                        'description': 'Plain text content unchanged',
                    },
                ],
            },
            'files': {
                'doc_in': '../../inputs/doc_in_paragraphs.docx',
                'gold': 'gold.docx',
            },
        }
        write_spec_if_changed(spec_path, spec)

        print(f"  [{case_name}/{variant_name}] Generated ({len(slots)} operations)")
        return True

    # ------------------------------------------------------------------
    # Main orchestration
    # ------------------------------------------------------------------

    def generate_all(self, command: str = None):
        """Generate all composition test cases.

        Args:
            command: If set, only generate this specific case name
                     (e.g., 'AdditiveFontComposition' or 'AdditiveListComposition')
        """
        wd = ensure_word(visible=not self.hidden)
        try:
            results = {}

            # Track 1A: Font compositions
            if command is None or command == 'AdditiveFontComposition':
                print(f"\n=== Track 1A: Font Compositions ===")
                shared_input = self._get_shared_input('doc_in_single_line')
                for variant_name, config in FONT_VARIANTS.items():
                    try:
                        ok = self._generate_target_text_variant(
                            wd, 'AdditiveFontComposition', variant_name, config,
                            shared_input=shared_input,
                            shared_input_name='doc_in_single_line',
                            template_name='simple_text',
                            full_text='Alpha Target Omega.',
                        )
                        results[f'AdditiveFontComposition/{variant_name}'] = ok
                    except Exception as e:
                        print(f"  [ERROR] {variant_name}: {e}")
                        import traceback
                        traceback.print_exc()
                        results[f'AdditiveFontComposition/{variant_name}'] = False

            # Track 1B: Paragraph compositions
            if command is None or command == 'AdditiveParagraphComposition':
                print(f"\n=== Track 1B: Paragraph Compositions ===")
                shared_input = self._get_paragraph_input(wd)
                for variant_name, config in PARAGRAPH_VARIANTS.items():
                    try:
                        ok = self._generate_target_text_variant(
                            wd, 'AdditiveParagraphComposition', variant_name, config,
                            shared_input=shared_input,
                            shared_input_name='doc_in_paragraphs_with_spacing',
                            template_name='paragraphs_with_spacing',
                            full_text=None,
                        )
                        results[f'AdditiveParagraphComposition/{variant_name}'] = ok
                    except Exception as e:
                        print(f"  [ERROR] {variant_name}: {e}")
                        import traceback
                        traceback.print_exc()
                        results[f'AdditiveParagraphComposition/{variant_name}'] = False

            # Track 2A: List compositions
            if command is None or command == 'AdditiveListComposition':
                print(f"\n=== Track 2A: List Compositions ===")
                for variant_name, config in LIST_VARIANTS.items():
                    try:
                        ok = self._generate_list_variant(
                            wd, 'AdditiveListComposition', variant_name, config
                        )
                        results[f'AdditiveListComposition/{variant_name}'] = ok
                    except Exception as e:
                        print(f"  [ERROR] {variant_name}: {e}")
                        import traceback
                        traceback.print_exc()
                        results[f'AdditiveListComposition/{variant_name}'] = False

            # Write instructions JSON (source of truth is in variant defs)
            commands = {}
            for cmd_name, variants in [
                ('AdditiveFontComposition', FONT_VARIANTS),
                ('AdditiveParagraphComposition', PARAGRAPH_VARIANTS),
                ('AdditiveListComposition', LIST_VARIANTS),
            ]:
                commands[cmd_name] = {
                    vname: vdef['instruction']
                    for vname, vdef in variants.items()
                }
            write_instructions_json(TEST_SET, commands)

            # Summary
            total = len(results)
            passed = sum(1 for v in results.values() if v)
            print(f"\n=== Summary: {passed}/{total} test cases generated ===")

            # Write generation summary
            summary_path = self.output_dir / 'generation_summary.json'
            summary = {
                'test_set': TEST_SET,
                'generated_at': datetime.now().isoformat(),
                'total': total,
                'passed': passed,
                'results': {k: 'success' if v else 'failed' for k, v in results.items()},
            }
            with open(summary_path, 'w', encoding='utf-8') as f:
                json.dump(summary, f, indent=2, ensure_ascii=False)

            return results
        finally:
            # Don't quit Word — may be reused
            pass


def main():
    import argparse
    parser = argparse.ArgumentParser(
        description='Generate additive content formatting composition test cases'
    )
    parser.add_argument('--command', '-c', help='Generate specific case only')
    parser.add_argument('--output', '-o', help='Output directory')
    parser.add_argument('--hidden', action='store_true', help='Run Word hidden')
    args = parser.parse_args()

    output_dir = Path(args.output) if args.output else None
    gen = AdditiveContentFormattingGenerator(output_dir=output_dir, hidden=args.hidden)
    gen.generate_all(command=args.command)


if __name__ == '__main__':
    main()
