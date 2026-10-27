"""Scope-changing interaction composition generator.

Each case: Step A changes document structure (e.g., insert row, merge cells,
add SmartArt node), Step B operates on elements that only exist after A.
Also includes multi-step chains (3 steps).

Usage:
    python -m generator.compositions.run --test-set interaction_composition_x_scope_changing
"""

import json
import time
from datetime import datetime
from pathlib import Path

from generator.config import BENCHMARK_OUTPUT_DIR
from generator.compositions.atomic_intents import (
    build_scope_changing_atomic_intents,
)
from generator.execution.setup import ensure_word

from .shared import (
    build_spec,
    close_doc,
    create_equation,
    create_smartart,
    create_table,
    get_shared_input,
    open_doc,
    save_as_gold,
    write_instructions_json,
    write_spec,
)

TEST_SET = 'interaction_composition_x_scope_changing'


# ═════════════════════════════════════════════════════════════════════════
# SmartArt helpers
# ═════════════════════════════════════════════════════════════════════════

def _find_node_by_text(sa, text):
    """Find a SmartArt node by its text content."""
    for i in range(1, sa.AllNodes.Count + 1):
        try:
            if sa.AllNodes(i).TextFrame2.TextRange.Text.strip() == text:
                return sa.AllNodes(i)
        except Exception:
            pass
    return None


def _find_smartart(doc):
    """Find the first SmartArt shape in the document."""
    for i in range(1, doc.Shapes.Count + 1):
        if doc.Shapes(i).HasSmartArt:
            return doc.Shapes(i), doc.Shapes(i).SmartArt
    return None, None


# ═════════════════════════════════════════════════════════════════════════
# Case definitions — grouped by command/directory name
# ═════════════════════════════════════════════════════════════════════════

CASE_GROUPS = [
    # ── Table Scope-Changing (2 variants) ──────────────────────────
    {
        'group': 'ScopeChangingTable',
        'target_object_type': 'table',
        'behavior': 'object_effect',
        'verification_method': 'composition_dispatch',
        'doc_in_name': 'doc_in_table',
        'variants': {
            'variant_insert_split_distribute_scopes': {
                'steps': [
                    {'operation': 'insert_row',
                     'params': {'position': 'below_first', 'after_row': 1}},
                    {'operation': 'insert_column',
                     'params': {'position': 'right_of_first', 'after_col': 1}},
                    {'operation': 'split_table',
                     'params': {'at_row': 2}},
                    {'operation': 'distribute_columns',
                     'params': {'table_index': 1}},
                    {'operation': 'distribute_rows',
                     'params': {'table_index': 2}},
                ],
                'interaction_class':
                    'insert_row_column_split_then_distribute_scopes',
                'verification_description':
                    'Document has two tables; the upper table has evenly '
                    'distributed columns, and the lower table has evenly '
                    'distributed rows',
                'instruction':
                    "Insert a new row below the first row and a new column "
                    "to the right of the first column, split the table at "
                    "the inserted row, then distribute the columns in the "
                    "upper table and distribute the rows in the lower table.",
                'doc_in_name': 'doc_in_table_uneven',
            },
            'variant_split_cells_align_block': {
                'steps': [
                    {'operation': 'split_cells',
                     'params': {'row': 1, 'col': 1, 'num_rows': 2,
                                'num_cols': 2}},
                    {'operation': 'align_cell',
                     'params': {'row': 1, 'col': 1, 'row_span': 2,
                                'col_span': 2, 'vertical': 1,
                                'horizontal': 1}},
                ],
                'interaction_class': 'split_cells_then_align_block',
                'verification_description':
                    'Merged cell split into a 2x2 block, and the newly split '
                    'top-left block is aligned middle-center',
                'instruction':
                    "Split the merged cell containing 'Merged 2x2' into 2 "
                    "rows and 2 columns, then set the newly split top-left "
                    "2x2 block to middle-center alignment.",
                'doc_in_name': 'doc_in_table_merged',
            },
        },
    },

    # ── SmartArt Scope-Changing (6 variants) ───────────────────────
    {
        'group': 'ScopeChangingSmartArt',
        'target_object_type': None,  # varies per variant
        'behavior': 'object_effect',
        'verification_method': 'composition_dispatch',
        'variants': {
            'variant_add_node_promote': {
                'steps': [
                    {'operation': 'add_smartart_node',
                     'params': {'after_text': 'Member1', 'new_text': 'New',
                                'position': 2}},
                    {'operation': 'promote_node',
                     'params': {'node_text': 'New'}},
                ],
                'target_object_type': 'smartart_hierarchy',
                'interaction_class': 'add_node_then_promote',
                'verification_description':
                    "SmartArt has 'New' node promoted one level in hierarchy",
                'instruction':
                    "Add a new node labeled 'New' after 'Member1' in the "
                    "SmartArt hierarchy, then promote it one level (it "
                    "becomes a sibling of Team1/Team2, and Member2 becomes "
                    "its subordinate).",
                'doc_in_name': 'doc_in_smartart_hierarchy',
            },
            'variant_add_node_reorder_up': {
                'steps': [
                    {'operation': 'add_smartart_node',
                     'params': {'after_text': 'Echo', 'new_text': 'New',
                                'position': 2}},
                    {'operation': 'reorder_node_up',
                     'params': {'node_text': 'New'}},
                ],
                'target_object_type': 'smartart_list',
                'interaction_class': 'add_node_then_reorder_up',
                'verification_description':
                    "SmartArt has 'New' node moved up one position",
                'instruction':
                    "Add a new node labeled 'New' after the last node 'Echo' "
                    "in the SmartArt, then move it up one position.",
                'doc_in_name': 'doc_in_smartart_list',
            },
            'variant_add_child_demote': {
                'steps': [
                    {'operation': 'add_smartart_child',
                     'params': {'parent_text': 'Team1', 'new_text': 'New',
                                'position': 5}},
                    {'operation': 'demote_node',
                     'params': {'node_text': 'New'}},
                ],
                'target_object_type': 'smartart_hierarchy',
                'interaction_class': 'add_child_then_demote',
                'verification_description':
                    "SmartArt has 'New' node demoted one level below Team1's "
                    "subordinates",
                'instruction':
                    "Add a subordinate node labeled 'New' under 'Team1', "
                    "after 'Member2', then demote it one level further (it "
                    "becomes a subordinate of 'Member2').",
                'doc_in_name': 'doc_in_smartart_hierarchy',
            },
            'variant_promote_reorder': {
                'steps': [
                    {'operation': 'promote_node',
                     'params': {'node_text': 'Member1'}},
                    {'operation': 'reorder_node_down',
                     'params': {'node_text': 'Member1'}},
                ],
                'target_object_type': 'smartart_hierarchy',
                'interaction_class': 'promote_then_reorder',
                'verification_description':
                    "'Member1' promoted to sibling of Team1/Team2, then "
                    "reordered down one position",
                'instruction':
                    "Promote 'Member1' one level so it becomes a sibling of "
                    "'Team1' and 'Team2' (Member2 becomes its subordinate), "
                    "then move it one position to the right among its new "
                    "siblings.",
                'doc_in_name': 'doc_in_smartart_hierarchy',
            },
            'variant_add_sibling_add_child': {
                'steps': [
                    {'operation': 'add_smartart_node',
                     'params': {'after_text': 'Team2', 'new_text': 'New',
                                'position': 2}},
                    {'operation': 'add_smartart_child',
                     'params': {'parent_text': 'New', 'new_text': 'SubNew',
                                'position': 5}},
                ],
                'target_object_type': 'smartart_hierarchy',
                'interaction_class': 'add_sibling_then_add_child',
                'verification_description':
                    "SmartArt has 'New' node after 'Team2' with 'SubNew' "
                    "subordinate below it",
                'instruction':
                    "Add a new node labeled 'New' after 'Team2' in the "
                    "SmartArt hierarchy, then add a subordinate node labeled "
                    "'SubNew' below it.",
                'doc_in_name': 'doc_in_smartart_hierarchy',
            },
            'variant_add_promote_reorder': {
                'steps': [
                    {'operation': 'add_smartart_node',
                     'params': {'after_text': 'Member1', 'new_text': 'New',
                                'position': 2}},
                    {'operation': 'promote_node',
                     'params': {'node_text': 'New'}},
                    {'operation': 'reorder_node_down',
                     'params': {'node_text': 'New'}},
                ],
                'target_object_type': 'smartart_hierarchy',
                'interaction_class': 'add_promote_then_reorder',
                'verification_description':
                    "'New' node added after 'Member1', promoted one level, "
                    "then reordered down one position",
                'instruction':
                    "Add a new node labeled 'New' after 'Member1' in the "
                    "SmartArt hierarchy, promote it one level (it becomes a "
                    "sibling of Team1/Team2, and Member2 becomes its "
                    "subordinate), then move it one position to the right "
                    "among its new siblings.",
                'doc_in_name': 'doc_in_smartart_hierarchy',
            },
        },
    },

    # ── Equation Scope-Changing (2 variants) ───────────────────────
    {
        'group': 'ScopeChangingEquation',
        'target_object_type': 'equation',
        'behavior': 'object_effect',
        'verification_method': 'composition_dispatch',
        'doc_in_name': 'doc_in_equation',
        'variants': {
            'variant_fraction_integral': {
                'steps': [
                    {'operation': 'insert_fraction',
                     'params': {'frac_type': 0, 'num': 'a', 'den': 'b'}},
                    {'operation': 'insert_integral_in_numerator',
                     'params': {'char': 8747, 'sub': '0', 'sup': '1',
                                'e': 'x'}},
                ],
                'interaction_class': 'insert_fraction_then_integral',
                'verification_description':
                    'Equation has fraction with integral in numerator',
                'instruction':
                    "Insert a stacked fraction with denominator 'b' after "
                    "'y=' in the equation, then insert an integral of 'x' "
                    "from '0' to '1' as the numerator.",
            },
            'variant_delimiter_matrix': {
                'steps': [
                    {'operation': 'insert_delimiter',
                     'params': {'beg_char': 91, 'end_char': 93,
                                'placeholder': 'a'}},
                    {'operation': 'insert_matrix_in_delimiter',
                     'params': {'rows': 2, 'cols': 2,
                                'values': ['a', 'b', 'c', 'd']}},
                ],
                'interaction_class': 'insert_delimiter_then_matrix',
                'verification_description':
                    'Equation has square brackets containing a 2x2 matrix',
                'instruction':
                    "Insert square brackets after 'y=' in the equation, then "
                    "insert a 2x2 matrix with entries a, b, c, d (row by "
                    "row) inside the brackets.",
            },
        },
    },

    # ── List Scope-Changing (2 variants) ───────────────────────────
    {
        'group': 'ScopeChangingList',
        'target_object_type': 'list',
        'behavior': 'content_formatting',
        'verification_method': 'property_based',
        'doc_in_name': 'doc_in_paragraphs',
        'variants': {
            'variant_bullets_level2': {
                'steps': [
                    {'operation': 'apply_bullets',
                     'params': {'gallery': 1, 'template': 1}},
                    {'operation': 'set_list_level',
                     'params': {'paragraph': 2, 'level': 2}},
                ],
                'interaction_class': 'apply_bullets_then_indent',
                'verification_description':
                    'All paragraphs have bullet list, paragraph 2 is at level 2',
                'instruction':
                    "Apply filled round bullet list formatting to all "
                    "paragraphs, then indent the second paragraph to list "
                    "level 2 (hollow circle bullet).",
            },
            'variant_numbering_level2': {
                'steps': [
                    {'operation': 'apply_numbering',
                     'params': {'gallery': 2, 'template': 1}},
                    {'operation': 'set_list_level_range',
                     'params': {'para_start': 2, 'para_end': 3, 'level': 2}},
                ],
                'interaction_class': 'apply_numbering_then_indent',
                'verification_description':
                    'All paragraphs have numbered list, paragraphs 2-3 are at '
                    'level 2',
                'instruction':
                    "Apply numbered list (1. 2. 3.) to all paragraphs, then "
                    "indent paragraphs 2 and 3 to list level 2 (a. b. c.).",
            },
        },
    },
]


# ═════════════════════════════════════════════════════════════════════════
# Generator class
# ═════════════════════════════════════════════════════════════════════════

class ScopeChangingGenerator:
    """Generator for scope-changing interaction test cases."""

    def __init__(self, output_dir: Path = None, hidden: bool = False):
        self.output_dir = output_dir or (BENCHMARK_OUTPUT_DIR / TEST_SET)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.inputs_dir = self.output_dir / 'inputs'
        self.inputs_dir.mkdir(parents=True, exist_ok=True)
        self.hidden = hidden

    def generate_all(self, command: str = None):
        """Generate all scope-changing test cases."""
        wd = ensure_word(visible=not self.hidden)
        results = {}

        for group_def in CASE_GROUPS:
            group_name = group_def['group']
            if command and command != group_name:
                continue
            for variant_name, variant_def in group_def['variants'].items():
                case_key = f'{group_name}/{variant_name}'
                # Merge group-level and variant-level fields
                case_def = self._merge_case_def(group_def, variant_def, variant_name)
                try:
                    ok = self._generate_case(wd, group_name, variant_name, case_def)
                    results[case_key] = ok
                    status = 'OK' if ok else 'FAILED'
                    print(f"  [{case_key}] {status}")
                except Exception as e:
                    results[case_key] = False
                    print(f"  [{case_key}] ERROR: {e}")

        # Write instructions JSON (source of truth is in case defs)
        commands = {}
        for group_def in CASE_GROUPS:
            group_name = group_def['group']
            variants = {}
            for variant_name, variant_def in group_def['variants'].items():
                variants[variant_name] = variant_def['instruction']
            commands[group_name] = variants
        write_instructions_json(TEST_SET, commands)

        # Write generation summary
        total = len(results)
        passed = sum(1 for v in results.values() if v)
        summary = {
            'test_set': TEST_SET,
            'generated_at': datetime.now().isoformat(),
            'total': total,
            'passed': passed,
            'results': {k: 'success' if v else 'failed'
                        for k, v in results.items()},
        }
        summary_path = self.output_dir / 'generation_summary.json'
        with open(summary_path, 'w', encoding='utf-8') as f:
            json.dump(summary, f, indent=2, ensure_ascii=False)

        print(f"\nSummary: {passed}/{total} cases generated successfully")
        return results

    # ─────────────────────────────────────────────────────────────────
    # Merge helper
    # ─────────────────────────────────────────────────────────────────

    def _merge_case_def(self, group_def, variant_def, variant_name):
        """Merge group-level defaults with variant-level overrides."""
        case_def = {
            'group': group_def['group'],
            'variant': variant_name,
            'target_object_type': variant_def.get('target_object_type') or group_def.get('target_object_type'),
            'behavior': variant_def.get('behavior') or group_def.get('behavior'),
            'verification_method': variant_def.get('verification_method') or group_def.get('verification_method'),
            'steps': variant_def['steps'],
            'interaction_class': variant_def['interaction_class'],
            'verification_description': variant_def['verification_description'],
            'instruction': variant_def['instruction'],
        }
        # Optional fields from variant or group
        for key in ('doc_in_name',):
            val = variant_def.get(key) or group_def.get(key)
            if val:
                case_def[key] = val
        return case_def

    # ─────────────────────────────────────────────────────────────────
    # Per-case generation
    # ─────────────────────────────────────────────────────────────────

    def _generate_case(self, wd, group_name, variant_name, case_def):
        """Generate a single test case variant."""
        variant_dir = self.output_dir / group_name / variant_name
        variant_dir.mkdir(parents=True, exist_ok=True)
        gold_path = variant_dir / 'gold.docx'
        spec_path = variant_dir / 'spec.json'

        # Resolve doc_in
        doc_in_name = case_def['doc_in_name']
        self._ensure_scope_input(wd, doc_in_name)
        doc_in_path = get_shared_input(self.inputs_dir, doc_in_name)
        doc_in_ref = f'../../inputs/{doc_in_name}.docx'

        # Generate gold via COM
        self._generate_gold(wd, case_def, doc_in_path, gold_path)

        # Build precondition
        precondition = self._build_precondition(case_def)
        atomic_intents = build_scope_changing_atomic_intents(case_def['steps'])

        # Build and write spec
        spec = build_spec(
            behavior=case_def['behavior'],
            composition_type='scope_changing',
            interaction_class=case_def['interaction_class'],
            target_object_type=case_def['target_object_type'],
            steps=case_def['steps'],
            precondition=precondition,
            verification_method=case_def['verification_method'],
            verification_description=case_def['verification_description'],
            doc_in_ref=doc_in_ref,
            variant_name=variant_name,
            atomic_intents=atomic_intents,
        )
        write_spec(spec_path, spec)
        return True

    # ─────────────────────────────────────────────────────────────────
    # Precondition builder
    # ─────────────────────────────────────────────────────────────────

    def _ensure_scope_input(self, wd, doc_in_name):
        """Create scope-specific inputs that are not shared by atomic sets."""
        dest = self.inputs_dir / f'{doc_in_name}.docx'
        if dest.exists():
            return

        if doc_in_name == 'doc_in_table_uneven':
            doc = wd.Documents.Add()
            try:
                table = create_table(doc, 3, 3)
                self._apply_visible_table_borders(table)
                table.AutoFitBehavior(0)  # wdAutoFitFixed
                try:
                    table.AllowAutoFit = False
                except Exception:
                    pass
                # Match the atomic intent setup: fixed, uneven sizes so
                # distribute has a visible effect, but narrow enough that
                # inserting a row/column does not push the table past margins.
                for idx, width in enumerate([45, 90, 135], 1):
                    table.Columns(idx).Width = width
                for idx, height in enumerate([30, 54, 78], 1):
                    row = table.Rows(idx)
                    row.Height = height
                    row.HeightRule = 1  # wdRowHeightExactly
                for row_idx in range(1, 4):
                    for col_idx in range(1, 4):
                        cell = table.Cell(row_idx, col_idx)
                        cell.VerticalAlignment = 1
                        cell.Range.ParagraphFormat.Alignment = 1
                save_as_gold(doc, dest)
            finally:
                close_doc(doc)
            return

    def _apply_visible_table_borders(self, table):
        """Give generated table inputs visible grid borders for previews."""
        try:
            table.Borders.Enable = True
        except Exception:
            pass
        for border_index in range(1, 7):
            try:
                border = table.Borders(border_index)
                border.LineStyle = 1  # wdLineStyleSingle
                border.LineWidth = 4  # wdLineWidth050pt
            except Exception:
                pass

    def _build_precondition(self, case_def):
        """Build the precondition dict for the spec."""
        doc_in_name = case_def['doc_in_name']

        if doc_in_name == 'doc_in_table':
            precondition = {
                'type': 'table',
                'description': '3x3 table with cells labeled R1C1..R3C3',
            }
        elif doc_in_name == 'doc_in_table_uneven':
            precondition = {
                'type': 'table',
                'description':
                    '3x3 table with uneven fixed row heights and column widths',
            }
        elif doc_in_name == 'doc_in_table_merged':
            precondition = {
                'type': 'table',
                'description': '3x3 table with R1C1:R2C2 merged',
            }
        elif doc_in_name == 'doc_in_smartart_list':
            precondition = {
                'type': 'smartart',
                'description':
                    'Basic Block List SmartArt with nodes: '
                    'Alpha, Beta, Charlie, Delta, Echo',
            }
        elif doc_in_name == 'doc_in_smartart_hierarchy':
            precondition = {
                'type': 'smartart',
                'description':
                    'Hierarchy SmartArt: Boss -> Team1/Team2 -> Members',
            }
        elif doc_in_name == 'doc_in_chart':
            precondition = {
                'type': 'chart',
                'description': 'Document with an inline chart',
            }
        elif doc_in_name == 'doc_in_equation':
            precondition = {
                'type': 'equation',
                'description': "Document with equation starting with 'y='",
            }
        elif doc_in_name == 'doc_in_paragraphs':
            precondition = {
                'type': 'paragraphs',
                'select_text': 'all',
                'description': '3+ paragraphs of plain text',
            }
        else:
            precondition = {
                'type': 'unknown',
                'description': f'Input document: {doc_in_name}',
            }

        return precondition

    # ─────────────────────────────────────────────────────────────────
    # Gold generation — dispatches by object type
    # ─────────────────────────────────────────────────────────────────

    def _generate_gold(self, wd, case_def, doc_in_path, gold_path):
        """Open doc_in, execute all steps via COM, save as gold."""
        target = case_def['target_object_type']

        if target == 'table':
            self._gold_table(wd, case_def, doc_in_path, gold_path)
        elif target in ('smartart_list', 'smartart_hierarchy'):
            self._gold_smartart(wd, case_def, doc_in_path, gold_path)
        elif target == 'chart':
            self._gold_chart(wd, case_def, doc_in_path, gold_path)
        elif target == 'equation':
            self._gold_equation(wd, case_def, doc_in_path, gold_path)
        elif target == 'list':
            self._gold_list(wd, case_def, doc_in_path, gold_path)
        else:
            raise ValueError(f'Unknown target_object_type: {target}')

    # ── Table gold generation ───────────────────────────────────────

    def _gold_table(self, wd, case_def, doc_in_path, gold_path):
        """Generate gold for table scope-changing cases."""
        variant = case_def['variant']
        doc = open_doc(wd, doc_in_path)
        try:
            table = doc.Tables(1)

            if variant == 'variant_insert_split_distribute_scopes':
                # Step A: Insert row below the first row.
                table.Cell(1, 1).Range.Select()
                time.sleep(0.2)
                wd.CommandBars.ExecuteMso('TableRowsInsertBelowWord')
                time.sleep(0.5)
                table = doc.Tables(1)

                # Step B: Insert column to the right of the first column.
                table.Columns(1).Select()
                time.sleep(0.2)
                wd.CommandBars.ExecuteMso('TableColumnsInsertRight')
                time.sleep(0.5)
                table = doc.Tables(1)

                # Step C: Split at the inserted second row.
                table.Cell(2, 1).Range.Select()
                wd.Selection.SplitTable()
                time.sleep(0.5)

                # Step D/E: Operate on the two scopes created by split.
                doc.Tables(1).Columns.DistributeWidth()
                doc.Tables(2).Rows.DistributeHeight()
                time.sleep(0.3)

            elif variant == 'variant_split_cells_align_block':
                # Step A: Split merged cell into a 2x2 block.
                table.Cell(1, 1).Range.Select()
                wd.Selection.Cells.Split(NumRows=2, NumColumns=2)
                time.sleep(0.5)
                table = doc.Tables(1)

                # Step B: Align the newly split top-left 2x2 region.
                for row_idx in range(1, 3):
                    for col_idx in range(1, 3):
                        cell = table.Cell(row_idx, col_idx)
                        cell.VerticalAlignment = 1  # wdCellAlignVerticalCenter
                        cell.Range.ParagraphFormat.Alignment = 1
                time.sleep(0.3)

            save_as_gold(doc, gold_path)
        finally:
            close_doc(doc)

    # ── SmartArt gold generation ────────────────────────────────────

    def _gold_smartart(self, wd, case_def, doc_in_path, gold_path):
        """Generate gold for SmartArt scope-changing cases."""
        variant = case_def['variant']
        doc = open_doc(wd, doc_in_path)
        try:
            sa_shape, sa = _find_smartart(doc)

            if variant == 'variant_add_node_promote':
                # Step A: Add node after Member1 (child of Team1 in hierarchy)
                target = _find_node_by_text(sa, 'Member1')
                new_node = target.AddNode(2)  # msoSmartArtNodeAfter
                new_node.TextFrame2.TextRange.Text = 'New'
                time.sleep(1.5)
                # Step B: Promote the new node (from level 3 to level 2)
                new_node = _find_node_by_text(sa, 'New')
                new_node.Promote()
                time.sleep(1.0)

            elif variant == 'variant_add_node_reorder_up':
                # Step A: Add node after Echo
                target = _find_node_by_text(sa, 'Echo')
                new_node = target.AddNode(2)  # msoSmartArtNodeAfter
                new_node.TextFrame2.TextRange.Text = 'New'
                time.sleep(1.0)
                # Step B: Reorder up
                new_node = _find_node_by_text(sa, 'New')
                new_node.ReorderUp()
                time.sleep(1.0)

            elif variant == 'variant_add_child_demote':
                # Step A: Add child below Team1
                target = _find_node_by_text(sa, 'Team1')
                new_node = target.AddNode(5)  # msoSmartArtNodeBelow
                new_node.TextFrame2.TextRange.Text = 'New'
                time.sleep(1.0)
                # Step B: Demote the new node
                new_node = _find_node_by_text(sa, 'New')
                new_node.Demote()
                time.sleep(1.0)

            elif variant == 'variant_promote_reorder':
                # Step A: Promote Member1
                target = _find_node_by_text(sa, 'Member1')
                target.Promote()
                time.sleep(1.0)
                # Step B: Reorder down
                target = _find_node_by_text(sa, 'Member1')
                target.ReorderDown()
                time.sleep(1.0)

            elif variant == 'variant_add_sibling_add_child':
                # Step A: Add sibling after Team2 (in hierarchy)
                target = _find_node_by_text(sa, 'Team2')
                new_node = target.AddNode(2)  # msoSmartArtNodeAfter
                new_node.TextFrame2.TextRange.Text = 'New'
                time.sleep(1.5)
                # Step B: Add child of new sibling
                new_node = _find_node_by_text(sa, 'New')
                child = new_node.AddNode(5)  # msoSmartArtNodeBelow
                child.TextFrame2.TextRange.Text = 'SubNew'
                time.sleep(1.0)

            elif variant == 'variant_add_promote_reorder':
                # Step A: Add node after Member1 (in hierarchy)
                target = _find_node_by_text(sa, 'Member1')
                new_node = target.AddNode(2)  # msoSmartArtNodeAfter
                new_node.TextFrame2.TextRange.Text = 'New'
                time.sleep(1.5)
                # Step B: Promote the new node
                new_node = _find_node_by_text(sa, 'New')
                new_node.Promote()
                time.sleep(1.0)
                # Step C: Reorder down
                new_node = _find_node_by_text(sa, 'New')
                new_node.ReorderDown()
                time.sleep(1.0)

            save_as_gold(doc, gold_path)
        finally:
            close_doc(doc)

    # ── Chart gold generation ───────────────────────────────────────

    def _gold_chart(self, wd, case_def, doc_in_path, gold_path):
        """Generate gold for chart scope-changing cases."""
        variant = case_def['variant']
        doc = open_doc(wd, doc_in_path)
        try:
            chart = doc.InlineShapes(1).Chart

            if variant == 'variant_change_type_droplines':
                # Step A: Add data labels outside each data point
                for i in range(1, chart.SeriesCollection().Count + 1):
                    chart.SeriesCollection(i).HasDataLabels = True
                    chart.SeriesCollection(i).DataLabels().Position = 2
                time.sleep(0.3)
                # Step B: Add data table with legend keys
                chart.HasDataTable = True
                chart.DataTable.ShowLegendKey = True
                time.sleep(0.3)

            elif variant == 'variant_add_title_format':
                # Step A: Add chart title above the chart
                chart.HasTitle = True
                chart.ChartTitle.Text = 'Chart Title'
                chart.ChartTitle.IncludeInLayout = True
                time.sleep(0.3)
                # Step B: Move legend to the right
                chart.HasLegend = True
                chart.Legend.Position = -4152  # xlLegendPositionRight
                time.sleep(0.3)

            save_as_gold(doc, gold_path)
        finally:
            close_doc(doc)

    # ── Equation gold generation ────────────────────────────────────

    def _gold_equation(self, wd, case_def, doc_in_path, gold_path):
        """Generate gold for equation scope-changing cases."""
        variant = case_def['variant']
        doc = open_doc(wd, doc_in_path)
        try:
            omath = doc.OMaths(1)

            if variant == 'variant_fraction_integral':
                # Step A: Insert fraction at end
                rng = omath.Range
                rng.Collapse(0)  # wdCollapseEnd
                frac = omath.Functions.Add(rng, 7)  # omMathFuncFrac
                frac.Frac.Type = 0  # stacked
                frac.Frac.Num.Range.Text = 'a'
                frac.Frac.Den.Range.Text = 'b'
                time.sleep(0.3)
                # Step B: Insert integral in numerator
                num_rng = frac.Frac.Num.Range
                num_rng.Text = ''  # Clear placeholder
                nary = omath.Functions.Add(num_rng, 13)  # omMathFuncNary
                nary.Nary.Char = 8747  # integral sign
                nary.Nary.SubSupLim = False
                nary.Nary.E.Range.Text = 'x'
                nary.Nary.Sub.Range.Text = '0'
                nary.Nary.Sup.Range.Text = '1'
                time.sleep(0.3)

            elif variant == 'variant_delimiter_matrix':
                # Step A: Insert delimiter (square brackets)
                rng = omath.Range
                rng.Collapse(0)  # wdCollapseEnd
                delim = omath.Functions.Add(rng, 5)  # omMathFuncDelim
                delim.Delim.BegChar = 91  # [
                delim.Delim.EndChar = 93  # ]
                delim.Delim.E(1).Range.Text = 'a'
                time.sleep(0.3)
                # Step B: Insert 2x2 matrix inside brackets
                inner_rng = delim.Delim.E(1).Range
                inner_rng.Text = ''  # Clear placeholder
                mat = omath.Functions.Add(inner_rng, 12)  # omMathFuncMat
                mat.Mat.Cols.Add()  # Add second column (starts with 1)
                mat.Mat.Rows.Add()  # Add second row
                # Fill 2x2 cells
                vals = iter(['a', 'b', 'c', 'd'])
                for r in range(1, 3):
                    for c in range(1, 3):
                        try:
                            mat.Mat.Cell(r, c).Range.Text = next(vals)
                        except Exception:
                            pass
                time.sleep(0.3)

            save_as_gold(doc, gold_path)
        finally:
            close_doc(doc)

    # ── List gold generation ────────────────────────────────────────

    def _gold_list(self, wd, case_def, doc_in_path, gold_path):
        """Generate gold for list scope-changing cases."""
        variant = case_def['variant']
        doc = open_doc(wd, doc_in_path)
        try:
            if variant == 'variant_bullets_level2':
                # Step A: Apply bullets to all paragraphs
                doc.Content.Select()
                template = wd.ListGalleries(1).ListTemplates(1)  # filled_round_bullet
                wd.Selection.Range.ListFormat.ApplyListTemplateWithLevel(
                    ListTemplate=template, ContinuePreviousList=False,
                    ApplyTo=0, DefaultListBehavior=1)
                time.sleep(0.3)
                # Step B: Set paragraph 2 to level 2
                doc.Paragraphs(2).Range.Select()
                wd.Selection.Range.ListFormat.ListLevelNumber = 2
                time.sleep(0.3)

            elif variant == 'variant_numbering_level2':
                # Step A: Apply numbering to all paragraphs
                doc.Content.Select()
                template = wd.ListGalleries(2).ListTemplates(1)  # arabic_period
                wd.Selection.Range.ListFormat.ApplyListTemplateWithLevel(
                    ListTemplate=template, ContinuePreviousList=False,
                    ApplyTo=0, DefaultListBehavior=1)
                time.sleep(0.3)
                # Step B: Set paragraphs 2-3 to level 2
                start = doc.Paragraphs(2).Range.Start
                end = doc.Paragraphs(3).Range.End
                doc.Range(start, end).Select()
                wd.Selection.Range.ListFormat.ListLevelNumber = 2
                time.sleep(0.3)

            save_as_gold(doc, gold_path)
        finally:
            close_doc(doc)
